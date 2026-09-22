# 文序 Windows 桌面版

当前版本：1.0.0-preview.1，LTS 开发预览版。

## 使用

1. 将 ZIP 完整解压到本机目录。保留 `Wenxu.exe` 与 `_internal` 目录的相对位置，双击 `Wenxu.exe`。
2. 点击“打开文档”，选择 DOC 或 DOCX。需要本机安装兼容的 WPS；不需要安装 Python、Git、WSL，也不需要浏览器。
3. 选择修改范围，指定字体或勾选字号等属性，然后点击“应用修改”。其他属性保持原样。
4. 查看预览，可打开“对照原稿”，通过页码逐页检查。核对后勾选“已核对当前版本的预览”。
5. 点击“另存副本”，保存到新文件。DOC 仍保存为 DOC，DOCX 仍保存为 DOCX。原稿不能被覆盖。

顶部可以撤销或切换本次会话内的版本。关闭程序后不会保留会话，请及时另存结果。
预览仅显示页面；通过右侧范围选择具体内容。“查找指定文字”须在搜索结果中选中目标，可按 Ctrl 多选。

## 运行条件与边界

- Windows x64；当前包需在目标 Windows / WPS 组合上验证后再作为正式稳定版使用。
- 本机 WPS 提供 KWPS 自动化接口，目标字体已安装。Microsoft Word 单独安装暂不能替代此接口。
- 若提示 WPS 正在处理其他文档，先保存并关闭 WPS 中打开的文件，再重试。程序不会强制关闭您的 WPS 文档。
- 继承当前核心的范围：10 MB 以内、单节、渲染后不超过 50 页，拒绝加密或含宏等未经验证的文档。
- 缺少原文字体时会提示；即使字节不变，也不能保证跨不同字体环境的版式一致。
- 文档由本机引擎处理，程序不提供上传、账号或遥测功能。用户文档不随分发包提供。
- 这是免安装预览包，尚未提供自动更新、安装器、数字签名或干净机器的兼容承诺。删除解压目录即可移除程序，不影响另存的文档。

## 开发与验收

维护分支 `lts/windows-desktop`；稳定维护规则见仓库 `docs/specs/windows-desktop-lts.md`。

在 Windows Python 3.12 x64 中：

```powershell
python -m venv .venv-desktop
.\.venv-desktop\Scripts\python.exe -m pip install -r requirements-desktop-build.txt
.\.venv-desktop\Scripts\python.exe -m unittest tests.test_desktop
.\.venv-desktop\Scripts\python.exe desktop_app.py
.\.venv-desktop\Scripts\python.exe scripts/build_desktop.py
```

可执行包的合成文档自检：`Wenxu.exe --self-test D:\WenxuChecks`。自检只生成新样例和结果，不读取用户私有文档；输出包括验收 JSON 和窗口截图，依赖实际 WPS。
正式分发前需完成干净 Windows 机器验收，并处理第三方再分发许可及签名。包内 `THIRD-PARTY` 保存依赖许可与元数据，`build-info.json` 记录版本。
