# 合并文档

**最终需求以 [PRD-V5.md](PRD-V5.md) 为准。**V1—V4 只是修复与完善过程，不是需要分别实现的版本；旧《需求文档》只供历史参考。

Windows V5 是本仓库当前最终效果：可拖入同文件夹的多个 `.xls`，或选择/拖入一个文件夹并只扫描第一层。载入后先显示候选文件数量，用户点击“开始合并”才读取表格、校验并输出。结果在原十列前增加文本列 `Invoice number`，删除每份源表底部的五项汇总；界面提供进度、取消、运行日志和在资源管理器中选中输出文件。完整边界条件与 macOS 目标请看上方 PRD。

## 开发运行

    python -m venv .venv
    .\.venv\Scripts\python.exe -m pip install -r requirements.txt
    .\.venv\Scripts\python.exe .\work\合并文档\app.py

## 测试

    .\.venv\Scripts\python.exe -m unittest discover -s .\work\合并文档\tests -v

## 打包

    .\.venv\Scripts\pyinstaller.exe --noconfirm --clean --onefile --windowed --name 合并文档V5 --icon .\work\合并文档\assets\app_icon.ico --add-data ".\work\合并文档\assets\app_icon.ico;assets" --additional-hooks-dir .\work\合并文档 .\work\合并文档\app.py
