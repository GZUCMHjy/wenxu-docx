param(
    [ValidateSet('discover', 'metadata', 'pdf', 'docx', 'inspect', 'edit-doc')][string]$Action = 'discover',
    [string]$InputPath,
    [string]$OutputPath,
    [string]$PlanPath
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
# Use the local WPS automation server, never a cloud document service.
if ($Action -eq 'discover') {
    @{available = [bool]([Type]::GetTypeFromProgID('KWPS.Application')); temp = [IO.Path]::GetTempPath()} | ConvertTo-Json -Compress
    exit 0
}
$mutex = [Threading.Mutex]::new($false, 'Local\WenxuOfficeBridge')
$locked = $false; $office = $null; $document = $null; $owned = $false
$security = $null; $links = $null; $alerts = $null; $visible = $null
function Write-Snapshot($doc, [string]$prefix) {
    # Word exposes table row terminators through Paragraphs; they are not OOXML
    # paragraphs. Use the actual engine flag, never guess from repeated text.
    $rows = @()
    foreach ($paragraph in $doc.Paragraphs) {
        $range = $paragraph.Range
        try {
            if (-not $range.Information(31)) {
                $row = @{start = $range.Start; end = $range.End; text = $range.Text}
                $expectedLength = $range.Text.Length
                # End-of-cell CR/BEL is two UTF-16 units but one native position.
                if ($range.Text.EndsWith("`r`a")) { $expectedLength-- }
                if ($range.End - $range.Start -ne $expectedLength) {
                    # An OLE/field marker can occupy many native positions.
                    # Ask the editor for ranges instead of inferring offsets.
                    $row.characters = @(
                        foreach ($character in $range.Characters) {
                            try {
                                if ($character.End - $character.Start -ne $character.Text.Length -and $character.Text -ne [string][char]1 -and $character.Text -ne "`r`a") {
                                    # Hidden anchor positions may precede ordinary
                                    # text. Preserve them as gaps, never targets.
                                    for ($cp = $character.Start; $cp -lt $character.End; $cp++) {
                                        $unit = $doc.Range($cp, $cp + 1)
                                        try { @{start = $unit.Start; end = $unit.End; text = $unit.Text} }
                                        finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($unit) }
                                    }
                                } else { @{start = $character.Start; end = $character.End; text = $character.Text} }
                            }
                            finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($character) }
                        }
                    )
                }
                $rows += $row
            }
        } finally {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($range)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($paragraph)
        }
    }
    # WordOpenXML is a read-only structural snapshot. It is never a working
    # artifact and is never loaded back into the DOC being edited.
    [IO.File]::WriteAllText(($prefix + '.xml'), $doc.WordOpenXML, [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText(($prefix + '.json'), (ConvertTo-Json -InputObject $rows -Depth 8 -Compress), [Text.UTF8Encoding]::new($false))
}
function Apply-Assignment($doc, $assignment) {
    $property = [string]$assignment.property; $value = $assignment.value
    if ($property -in @('top_margin','bottom_margin','left_margin','right_margin')) {
        switch ($property) {
            top_margin { $doc.PageSetup.TopMargin = [single]$value }
            bottom_margin { $doc.PageSetup.BottomMargin = [single]$value }
            left_margin { $doc.PageSetup.LeftMargin = [single]$value }
            right_margin { $doc.PageSetup.RightMargin = [single]$value }
        }
        return
    }
    $range = $doc.Range([int]$assignment.start, [int]$assignment.end)
    try {
        if ($range.Text -cne [string]$assignment.text) { throw 'RANGE_MISMATCH' }
        switch ($property) {
            font_east_asia { $range.Font.NameFarEast = [string]$value }
            font_western { $range.Font.NameAscii = [string]$value; $range.Font.NameOther = [string]$value; $range.Font.NameBi = [string]$value }
            font_size { $range.Font.Size = [single]$value; $range.Font.SizeBi = [single]$value }
            bold { $range.Font.Bold = $(if ($value) { -1 } else { 0 }) }
            italic { $range.Font.Italic = $(if ($value) { -1 } else { 0 }) }
            underline { $range.Font.Underline = $(if ($value) { 1 } else { 0 }) }
            color {
                $rgb = [Convert]::ToInt32(([string]$value).TrimStart('#'), 16)
                $range.Font.Color = (($rgb -shr 16) -band 255) + ($rgb -band 65280) + (($rgb -band 255) -shl 16)
            }
            alignment { $range.ParagraphFormat.Alignment = @{left=0;center=1;right=2;justify=3}[[string]$value] }
            first_line_indent { $range.ParagraphFormat.FirstLineIndent = [single]$value }
            left_indent { $range.ParagraphFormat.LeftIndent = [single]$value }
            right_indent { $range.ParagraphFormat.RightIndent = [single]$value }
            space_before { $range.ParagraphFormat.SpaceBeforeAuto = 0; $range.ParagraphFormat.SpaceBefore = [single]$value }
            space_after { $range.ParagraphFormat.SpaceAfterAuto = 0; $range.ParagraphFormat.SpaceAfter = [single]$value }
            keep_with_next { $range.ParagraphFormat.KeepWithNext = $(if ($value) { -1 } else { 0 }) }
            line_spacing {
                if ($value.mode -eq 'exact') { $range.ParagraphFormat.LineSpacingRule = 4; $range.ParagraphFormat.LineSpacing = [single]$value.value }
                elseif ($value.mode -eq 'multiple') { $range.ParagraphFormat.LineSpacingRule = 5; $range.ParagraphFormat.LineSpacing = [single]$value.value * 12 }
                else { throw 'UNSUPPORTED_PROPERTY' }
            }
            default { throw 'UNSUPPORTED_PROPERTY' }
        }
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($range) }
}
try {
    try { $locked = $mutex.WaitOne(90000) } catch [Threading.AbandonedMutexException] { $locked = $true }
    if (-not $locked) { throw 'OFFICE_BUSY' }
    $office = New-Object -ComObject KWPS.Application
    # Some COM servers attach to an existing desktop instance. Do not change,
    # hide or close an instance containing the user's documents.
    if ($office.Documents.Count -ne 0) { throw 'OFFICE_IN_USE' }
    $owned = $true
    $security = $office.AutomationSecurity; $links = $office.Options.UpdateLinksAtOpen
    $alerts = $office.DisplayAlerts; $visible = $office.Visible
    $office.AutomationSecurity = 3; $office.Options.UpdateLinksAtOpen = $false
    $office.DisplayAlerts = 0; $office.Visible = $false
    if ($Action -eq 'metadata') {
        $fonts = @(); foreach ($font in $office.FontNames) { $fonts += [string]$font }
        $binary = Join-Path $office.Path 'wps.exe'
        $version = if (Test-Path -LiteralPath $binary) { [Diagnostics.FileVersionInfo]::GetVersionInfo($binary).FileVersion } else { 'COM ' + [string]$office.Version }
        @{renderer = 'WPS local automation'; version = $version; font_families = $fonts} | ConvertTo-Json -Compress
    } else {
        if (-not [IO.Path]::IsPathRooted($InputPath) -or -not [IO.Path]::IsPathRooted($OutputPath)) { throw 'INVALID_PATH' }
        # All paths refer to our disposable local job directory. Original bytes
        # are never opened for writing; no file is added to Office recent files.
        $document = $office.Documents.Open($InputPath, $false, $true, $false)
        if ($Action -eq 'inspect') {
            Write-Snapshot $document $OutputPath
        } elseif ($Action -eq 'edit-doc') {
            $request = Get-Content -LiteralPath $PlanPath -Raw -Encoding UTF8 | ConvertFrom-Json
            $hasher = [Security.Cryptography.SHA256]::Create()
            try { $inputHash = [BitConverter]::ToString($hasher.ComputeHash([IO.File]::ReadAllBytes($InputPath))).Replace('-', '') }
            finally { $hasher.Dispose() }
            if ($inputHash -ine $request.input_hash) { throw 'INPUT_MISMATCH' }
            foreach ($assignment in $request.assignments) { Apply-Assignment $document $assignment }
            $document.SaveAs2($OutputPath, 0, $false, '', $false)
            $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document); $document = $null
            # Verify the persisted artifact, not the mutable in-memory editor.
            $document = $office.Documents.Open($OutputPath, $false, $true, $false)
            Write-Snapshot $document $OutputPath
        } elseif ($Action -eq 'pdf') {
            $document.ExportAsFixedFormat($OutputPath, 17)
        } else {
            # Omitting CompatibilityMode preserves the input's compatibility
            # settings. Do not call Convert() or re-save via LibreOffice.
            $document.SaveAs2($OutputPath, 12, $false, '', $false)
        }
        if ($Action -ne 'inspect' -and -not (Test-Path -LiteralPath $OutputPath -PathType Leaf)) { throw 'NO_OUTPUT' }
        @{ok = $true} | ConvertTo-Json -Compress
    }
} catch {
    # Avoid leaking document contents or local filenames into application logs.
    $code = if ($_.Exception.Message -in @('OFFICE_BUSY', 'OFFICE_IN_USE')) { $_.Exception.Message } else { 'OFFICE_FAILED' }
    [Console]::Error.WriteLine($code)
    exit 1
} finally {
    if ($document) {
        try { $document.Close(0) } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    }
    if ($office) {
        if ($owned) {
            if ($null -ne $security) { $office.AutomationSecurity = $security }
            if ($null -ne $links) { $office.Options.UpdateLinksAtOpen = $links }
            if ($null -ne $alerts) { $office.DisplayAlerts = $alerts }
            # Never close documents opened by a user while this job was running.
            if ($office.Documents.Count -eq 0) {
                if ($null -ne $visible) { $office.Visible = $visible }
                $office.Quit()
            }
        }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($office)
    }
    if ($locked) { $mutex.ReleaseMutex() }; $mutex.Dispose()
}
