"""Local WPS adapter, callable from WSL without changing the app runtime."""
from functools import lru_cache
import json
import ntpath
import os
from pathlib import Path
import platform
import subprocess
import tempfile

from docx_formatting import FormatError


def _powershell():
    if platform.system() == 'Windows':
        root = Path(os.environ.get('SystemRoot', r'C:\Windows'))
    elif 'microsoft' in platform.release().lower():
        root = Path('/mnt/c/Windows')
    else:
        return None
    # WPS registers its local COM server in the 32-bit registry view.
    executable = root / 'SysWOW64/WindowsPowerShell/v1.0/powershell.exe'
    return str(executable) if executable.is_file() else None


def _path(path, windows=True):
    if platform.system() == 'Windows':
        return str(path)
    try:
        return subprocess.check_output(['wslpath', '-w' if windows else '-u', str(path)], text=True, timeout=10).strip()
    except (OSError, subprocess.SubprocessError) as exc:
        raise FormatError('无法访问本机文档渲染目录。') from exc


class LocalOffice:
    def __init__(self, executable, temp):
        self.executable = executable
        self.temp = Path(_path(temp, windows=False))
        self.script = _path(Path(__file__).with_name('office_bridge.ps1').resolve())

    def call(self, action, *args):
        try:
            result = subprocess.run([self.executable, '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                                     '-File', self.script, '-Action', action, *args], capture_output=True, timeout=180)
        except (OSError, subprocess.TimeoutExpired) as exc:
            raise FormatError('本机 WPS 转换超时或启动失败；原稿保留。') from exc
        if result.returncode:
            code = result.stderr.decode('utf-8', errors='replace')
            if 'OFFICE_IN_USE' in code or 'OFFICE_BUSY' in code:
                raise FormatError('本机 WPS 正在处理其他文档，本次未转换；请关闭占用的 WPS 文档后重试。')
            raise FormatError('本机 WPS 未生成可验证结果；原稿保留，未切换到其他渲染器。')
        try:
            return json.loads(result.stdout.decode('utf-8-sig'))
        except (ValueError, UnicodeError) as exc:
            raise FormatError('本机 WPS 返回了无法验证的结果。') from exc

    def convert(self, data, extension, target):
        target = target.split(':', 1)[0]
        if extension not in {'doc', 'docx'} or target not in {'pdf', 'docx'}:
            raise FormatError('本机转换仅支持 DOC/DOCX 输入及 PDF/DOCX 输出。')
        with tempfile.TemporaryDirectory(prefix='wenxu-office-', dir=self.temp) as temp:
            source = Path(temp) / ('input.' + extension)
            output = Path(temp) / ('output.' + target)
            source.write_bytes(data)
            # wslpath on Windows mounts requires the target to exist. Translate
            # the existing directory, then append the new output filename.
            windows_temp = _path(Path(temp))
            self.call(target, '-InputPath', ntpath.join(windows_temp, source.name),
                      '-OutputPath', ntpath.join(windows_temp, output.name))
            if not output.is_file() or not output.stat().st_size:
                raise FormatError('本机 WPS 未产生输出。')
            return output.read_bytes()

    def inspect_doc(self, data, request=None):
        """Read a DOC snapshot, or edit that DOC and inspect its saved bytes."""
        with tempfile.TemporaryDirectory(prefix='wenxu-native-', dir=self.temp) as temp:
            root = Path(temp); source = root / 'input.doc'; output = root / 'output.doc'
            source.write_bytes(data)
            windows_temp = _path(root)
            args = ['-InputPath', ntpath.join(windows_temp, source.name), '-OutputPath', ntpath.join(windows_temp, output.name)]
            if request is not None:
                (root / 'plan.json').write_text(json.dumps(request, ensure_ascii=False), encoding='utf-8')
                args += ['-PlanPath', ntpath.join(windows_temp, 'plan.json')]
            self.call('edit-doc' if request is not None else 'inspect', *args)
            return (output.read_bytes() if request is not None else data,
                    Path(str(output) + '.xml').read_bytes(),
                    json.loads(Path(str(output) + '.json').read_text(encoding='utf-8-sig')))

    @lru_cache(maxsize=1)
    def environment(self):
        result = self.call('metadata')
        result['renderer'] += ' ' + result.pop('version')
        result.update(platform=platform.platform(), python=platform.python_version(),
                      parameters='120dpi/max50/macro-disabled/no-link-update/compatibility-preserved')
        return result


@lru_cache(maxsize=1)
def native_office():
    """Select once per process; never silently change engines after a failure."""
    mode = os.environ.get('WENXU_RENDERER', 'auto')
    if mode not in {'auto', 'wps', 'libreoffice'}:
        raise FormatError('WENXU_RENDERER 仅支持 auto、wps、libreoffice。')
    if mode == 'libreoffice':
        return None
    executable = _powershell()
    if executable:
        script = _path(Path(__file__).with_name('office_bridge.ps1').resolve())
        try:
            result = subprocess.run([executable, '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                                     '-File', script, '-Action', 'discover'], capture_output=True, timeout=20)
            info = json.loads(result.stdout.decode('utf-8-sig')) if result.returncode == 0 else {}
        except (OSError, subprocess.SubprocessError, ValueError):
            info = {}
        if info.get('available'):
            return LocalOffice(executable, info['temp'])
    if mode == 'wps':
        raise FormatError('未发现可用的本机 WPS 自动化接口。')
    return None
