# 合并文档

本地 Windows `.xls` 批量合并工具。拖入或选择文件夹后只扫描并显示数量，用户点击“开始合并”后才执行；同时支持结构校验、进度显示、取消、运行日志和打开输出位置。

V3 在输出表最左侧新增 `Invoice number`，原有 10 列保持原样。每个源文件名中从左向右找到连续五个 `0`，把这五个 `0` 后直到文件扩展名前的数字作为文本填入该文件的每条明细，保留剩余的前导零。例如 `claim_details_000000400699589.xls` 得到 `0400699589`，`claim_details_000005301106261.xls` 得到 `5301106261`。文件名不符合规则时跳过并记入运行日志。

V4 可将同一文件夹中的一个或多个 `.xls` 文件直接拖入窗口，只合并拖入的文件；“选择文件夹”按钮仍扫描该文件夹第一层。两种方式都先显示文件数量，点击“开始合并”后才处理。完成后，“打开所在文件夹”会在资源管理器中选中刚生成的结果文件。

## 开发运行

    python -m venv .venv
    .\.venv\Scripts\python.exe -m pip install -r requirements.txt
    .\.venv\Scripts\python.exe .\work\合并文档\app.py

## 测试

    .\.venv\Scripts\python.exe -m unittest discover -s .\work\合并文档\tests -v

## 打包

    .\.venv\Scripts\pyinstaller.exe --noconfirm --clean --onefile --windowed --name 合并文档V4 --additional-hooks-dir .\work\合并文档 .\work\合并文档\app.py
