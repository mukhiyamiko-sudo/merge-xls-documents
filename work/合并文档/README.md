# 合并文档

**最终需求以 [PRD-V5.md](PRD-V5.md) 为准。**V1—V4 只是修复与完善过程，不是需要分别实现的版本；旧《需求文档》只供历史参考。

V5 可拖入同文件夹的多个 `.xls`，或选择/拖入一个文件夹并只扫描第一层。载入后先显示候选文件数量，用户点击“开始合并”才读取表格、校验并输出。结果在原十列前增加文本列 `Invoice number`，删除每份源表底部的五项汇总；Windows 输出 `.xls`，macOS 输出 `.xlsx`。界面提供进度、取消、运行日志和在文件管理器中选中输出文件。完整边界条件与 macOS 目标请看上方 PRD。

## 开发运行

    python -m venv .venv
    .\.venv\Scripts\python.exe -m pip install -r requirements.txt
    .\.venv\Scripts\python.exe .\work\合并文档\app.py

## 测试

    .\.venv\Scripts\python.exe -m unittest discover -s .\work\合并文档\tests -v

## 打包

    .\.venv\Scripts\pyinstaller.exe --noconfirm --clean --onefile --windowed --name 合并文档V5 --icon .\work\合并文档\assets\app_icon.ico --add-data ".\work\合并文档\assets\app_icon.ico;assets" --additional-hooks-dir .\work\合并文档 .\work\合并文档\app.py

## macOS Intel 候选包

仓库的 GitHub Actions 工作流 `Build V5 macOS Intel candidate` 在 Intel Mac 运行器上打包 `.app`，运行测试与合成 `.xls` 输入、`.xlsx` 输出自检，再将 `.app` 保存在 ZIP 中。下载 ZIP 后在 Mac 上解压，把 `合并文档V5.app` 放进“应用程序”即可尝试打开；它未使用 Apple Developer ID 公证，首次打开可能需要按 macOS 安全提示手动允许。请只从本项目可信的 GitHub 页面下载，不要对来源不明或显示“已损坏”的应用绕过系统警告。

工作流使用较新的 macOS 构建机，因此生成物在目标 macOS Big Sur 11.7.9 Intel 机器上完成启动、拖拽、合并与 Finder 定位测试前，只能称为**候选包**，不能宣称已验证兼容。业务源表不会上传到仓库或构建服务；构建自检使用程序生成的假数据。
