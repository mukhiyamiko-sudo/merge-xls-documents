# 合并文档

本地 Windows `.xls` 批量合并工具。V2 在拖入或选择文件夹后只扫描并显示数量，用户点击“开始合并”后才执行；同时支持结构校验、进度显示、取消、运行日志和打开输出位置。

## 开发运行

    python -m venv .venv
    .\.venv\Scripts\python.exe -m pip install -r requirements.txt
    .\.venv\Scripts\python.exe .\work\合并文档\app.py

## 测试

    .\.venv\Scripts\python.exe -m unittest discover -s .\work\合并文档\tests -v

## 打包

    .\.venv\Scripts\pyinstaller.exe --noconfirm --clean --onefile --windowed --name 合并文档 --additional-hooks-dir .\work\合并文档 .\work\合并文档\app.py
