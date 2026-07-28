# wheelhouse_server.exe 打包说明

本文档说明如何将 `wheel_trim\server\server.py` 打包为本地可执行服务 `wheelhouse_server.exe`，并供前端页面调用。

## 1. 打包前准备

先进入车轮罩项目目录：

```powershell
cd C:\Users\Administrator\Desktop\zhiji_project\IM_Wsp\wheel_trim
```

在 Windows 环境中确认已安装 Python 依赖：

```powershell
pip install -r server\requirements.txt
pip install pyinstaller
```

项目中已提供 PyInstaller 配置文件：

```text
server\wheelhouse_server.spec
```

该配置会把以下内容一起打进 exe：

- 后端入口：`server\server.py`
- 前端页面：`web`
- 子进程入口模块：`server\run_algorithm_worker.py`
- 算法主文件：`server\pycatia_wheelhouse_regulation_verification\main.py`
- 截面导出工具：`section_curve_export_tool.py`
- 截图工具：`catia_picture_capture.py`
- 标注工具：`catia_annotation_tools.py`
- 算法目录下的 `.CATPart` 示例文件
- `uvicorn`、`win32com`、`pythoncom`、`python-docx`、`python-multipart` 等运行依赖

## 2. 默认打包命令

在车轮罩项目目录执行：

```powershell
cd C:\Users\Administrator\Desktop\zhiji_project\IM_Wsp\wheel_trim
python -m PyInstaller --clean --noconfirm --workpath build server\wheelhouse_server.spec
```

生成的 exe 默认位于：

```text
dist\wheelhouse_server.exe
```

完整路径示例：

```text
C:\Users\Administrator\Desktop\zhiji_project\IM_Wsp\wheel_trim\dist\wheelhouse_server.exe
```

## 3. 指定输出目录打包

如果旧 exe 正在运行，或者希望保留不同版本，可以指定输出目录：

```powershell
cd C:\Users\Administrator\Desktop\zhiji_project\IM_Wsp\wheel_trim
python -m PyInstaller --clean --noconfirm --workpath build --distpath dist_wheelhouse server\wheelhouse_server.spec
```

生成位置：

```text
dist_wheelhouse\wheelhouse_server.exe
```

`--distpath dist_wheelhouse` 表示把最终 exe 输出到 `dist_wheelhouse` 目录，避免覆盖默认的 `dist` 目录。

`--workpath build` 表示把 PyInstaller 的临时构建文件输出到当前目录下的 `build` 目录。

注意：指定 `--distpath dist_wheelhouse` 后，旧的 `dist\wheelhouse_server.exe` 不会被自动删除。如果之前已经打包过默认 `dist`，该文件仍会留在那里；应以控制台最后一行 `The results are available in: ...` 或 `dist_wheelhouse\wheelhouse_server.exe` 的修改时间判断本次输出位置。

## 4. 运行 exe

双击运行：

```text
wheelhouse_server.exe
```

或在 PowerShell 中运行：

```powershell
.\dist\wheelhouse_server.exe
```

服务启动后访问：

```text
http://localhost:8010
```

如果 `8010` 端口被占用，程序会自动尝试后续端口，例如 `8011`、`8012` 等，并在控制台输出实际访问地址。

也可以手动指定端口：

```powershell
$env:WHEELHOUSE_SERVER_PORT="8020"
.\dist\wheelhouse_server.exe
```

然后访问：

```text
http://localhost:8020
```

## 5. 输出文件位置

打包后的 exe 运行时，会在 exe 同级目录生成运行时文件。

示例：

```text
dist\
  wheelhouse_server.exe
  uploads\
  session.json
  output\
```

其中：

- `uploads\`：前端上传的轮罩 `.CATPart` 和车轮装配文件
- `session.json`：服务会话状态
- `output\`：算法生成的结果文件

检测完成后，常见输出文件包括：

```text
output\
  output-yyyyMMdd_HHmmss\
    result-yyyymmdd_HHmmss\
      wheelhouse_regulation_verification_yyyyMMdd_HHmmss.CATProduct
      Wheelhouse_Regulation_Axis_Lines_yyyyMMdd_HHmmss.CATPart
      Wheelhouse_Regulation_Annotations.CATPart
      wheelhouse_regulation_verification_result_yyyyMMdd_HHmmss.json
      wheelhouse_regulation_report_yyyyMMdd_HHmmss.docx
      screenshots\
```

## 6. 修改前端或算法后是否需要重新打包

需要。

当前前端 `web\index.html`、`web\app.js`、`web\style.css` 会被打进 exe。算法文件和通用工具文件也会被打进 exe。修改后如果继续运行旧 exe，页面和算法仍然是旧版本。

修改后重新执行：

```powershell
cd C:\Users\Administrator\Desktop\zhiji_project\IM_Wsp\wheel_trim
python -m PyInstaller --clean --noconfirm --workpath build server\wheelhouse_server.spec
```

浏览器如果仍显示旧页面，可按：

```text
Ctrl + F5
```

强制刷新缓存。

## 7. exe 文件名如何修改

exe 文件名由 spec 文件控制：

```text
server\wheelhouse_server.spec
```

其中：

```python
name="wheelhouse_server"
```

决定最终生成：

```text
wheelhouse_server.exe
```

如果想改成：

```text
WheelhouseRegulationServer.exe
```

可修改为：

```python
name="WheelhouseRegulationServer"
```

然后重新打包。

## 8. 常见问题

### 8.1 提示旧 exe 被占用

如果打包时提示无法覆盖 `wheelhouse_server.exe`，通常是旧服务正在运行。

处理方式：

1. 在任务管理器中结束 `wheelhouse_server.exe`
2. 或者使用 `--distpath` 输出到新目录

示例：

```powershell
cd C:\Users\Administrator\Desktop\zhiji_project\IM_Wsp\wheel_trim
python -m PyInstaller --clean --noconfirm --workpath build --distpath dist_new server\wheelhouse_server.spec
```

### 8.2 检测失败：找不到算法文件或工具文件

应确认使用的是当前 `server\wheelhouse_server.spec` 重新打包后的 exe。

本项目运行时会通过文件路径动态加载以下工具：

```text
pycatia_wheelhouse_regulation_verification\main.py
pycatia_wheelhouse_regulation_verification\section_curve_export_tool.py
pycatia_wheelhouse_regulation_verification\catia_picture_capture.py
pycatia_wheelhouse_regulation_verification\catia_annotation_tools.py
```

这些文件必须在 spec 的 `datas` 中一起打包。

### 8.3 output 没有生成到 exe 同级目录

打包后的服务会把输出目录设置为：

```text
<exe所在目录>\output
```

如果没有看到输出文件，请先确认：

- exe 运行目录有写入权限
- 前端检测流程确实完成
- 控制台没有 CATIA 或文件保存异常

### 8.4 CATIA 无法连接

确认：

- 本机已安装 CATIA
- CATIA COM 自动化可用
- CATIA 已打开，或前端点击“打开 CATIA”
- 当前用户权限可以访问 CATIA COM
- 上传的是有效 `.CATPart` / `.CATProduct` 文件

### 8.5 Word 报告没有截图

确认：

- 已安装 `python-docx`
- CATIA 截图过程中窗口没有被异常关闭
- 输出目录下存在 `screenshots` 文件夹和对应图片
- 打包时没有遗漏 `catia_picture_capture.py`

## 9. 推荐发布方式

建议将 exe 放在固定、有写权限的目录，例如：

```text
D:\WheelhouseServer\wheelhouse_server.exe
```

不建议放在：

```text
C:\Program Files\
C:\Windows\
只读目录
网络盘路径
```

原因是服务运行时需要写入：

- `uploads`
- `session.json`
- `output`
