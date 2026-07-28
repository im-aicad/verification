# rearview_server.exe 打包说明

本文档说明如何将 `rearview_mirror\server\server.py` 打包为本地可执行服务 `rearview_server.exe`，并供前端页面调用。

## 1. 打包前准备

先进入后视镜项目目录：

```powershell
cd C:\Users\Administrator\Desktop\zhiji_project\IM_Wsp\rearview_mirror
```

在 Windows 环境中确认已安装 Python 依赖：

```powershell
pip install -r server\requirements.txt
pip install pyinstaller
```

项目中已提供 PyInstaller 配置文件：

```text
server\rearview_server.spec
```

该配置会把以下内容一起打进 exe：

- 后端入口：`server\server.py`
- 前端页面：`web`
- 算法文件：`server\pycatia_regulation_reflection_point_detection\main.py`
- 算法资源：`resources`
- 算法目录下的 `.CATPart` 示例文件
- `uvicorn`、`win32com`、`pythoncom`、`python-multipart` 等运行依赖

## 2. 默认打包命令

在后视镜项目目录执行：

```powershell
cd C:\Users\Administrator\Desktop\zhiji_project\IM_Wsp\rearview_mirror
python -m PyInstaller --clean --noconfirm server\rearview_server.spec --workpath build
```

生成的 exe 默认位于：

```text
dist\rearview_server.exe
```

完整路径示例：

```text
C:\Users\Administrator\Desktop\zhiji_project\IM_Wsp\rearview_mirror\dist\rearview_server.exe
```

## 3. 指定输出目录打包

如果旧 exe 正在运行，或者希望保留不同版本，可以指定输出目录：

```powershell
cd C:\Users\Administrator\Desktop\zhiji_project\IM_Wsp\rearview_mirror
python -m PyInstaller --clean --noconfirm server\rearview_server.spec --workpath build --distpath dist_rearview
```

生成位置：

```text
dist_rearview\rearview_server.exe
```

`--distpath dist_rearview` 表示把最终 exe 输出到 `dist_rearview` 目录，避免覆盖默认的 `dist` 目录。

`--workpath build` 表示把 PyInstaller 的临时构建文件输出到当前目录下的 `build` 目录。

由于当前命令已经先进入 `rearview_mirror` 目录，因此会生成：

```text
rearview_mirror\
  build\
  dist_rearview\
    rearview_server.exe
```

## 4. 运行 exe

双击运行：

```text
rearview_server.exe
```

或在 PowerShell 中运行：

```powershell
.\dist\rearview_server.exe
```

服务启动后访问：

```text
http://localhost:8000
```

如果 `8000` 端口被占用，程序会自动尝试后续端口，例如 `8001`、`8002` 等，并在控制台输出实际访问地址。

也可以手动指定端口：

```powershell
$env:REARVIEW_SERVER_PORT="8010"
.\dist\rearview_server.exe
```

然后访问：

```text
http://localhost:8010
```

## 5. 输出文件位置

打包后的 exe 运行时，会在 exe 同级目录生成运行时文件。

示例：

```text
dist\
  rearview_server.exe
  uploads\
  session.json
  output\
```

其中：

- `uploads\`：前端上传的原始 `.CATPart` 文件
- `session.json`：服务会话状态
- `output\`：算法生成的结果文件

检测完成后，常见输出文件包括：

```text
output\
  xxx_rearview_result_....CATPart
  label_....CATProduct
  Rearview_Distance_Annotations_....CATPart
  法规视野截图_....jpg
  左后视镜截图_....jpg
  右后视镜截图_....jpg
  外后视镜视野校核报告_....docx
```

## 6. 修改前端后是否需要重新打包

需要。

当前前端 `web\index.html` 会被打进 exe。修改前端后，如果继续运行旧 exe，页面仍然是旧版本。

修改前端后重新执行：

```powershell
cd C:\Users\Administrator\Desktop\zhiji_project\IM_Wsp\rearview_mirror
python -m PyInstaller --clean --noconfirm server\rearview_server.spec --workpath build
```

或输出到新目录：

```powershell
cd C:\Users\Administrator\Desktop\zhiji_project\IM_Wsp\rearview_mirror
python -m PyInstaller --clean --noconfirm server\rearview_server.spec --workpath build --distpath dist_rearview
```

浏览器如果仍显示旧页面，可按：

```text
Ctrl + F5
```

强制刷新缓存。

## 7. exe 文件名如何修改

exe 文件名由 spec 文件控制：

```text
server\rearview_server.spec
```

其中：

```python
name="rearview_server"
```

决定最终生成：

```text
rearview_server.exe
```

如果想改成：

```text
RearviewMirrorService.exe
```

可修改为：

```python
name="RearviewMirrorService"
```

然后重新打包。

## 8. 常见问题

### 8.1 提示旧 exe 被占用

如果打包时提示无法覆盖 `rearview_server.exe`，通常是旧服务正在运行。

处理方式：

1. 在任务管理器中结束 `rearview_server.exe`
2. 或者使用 `--distpath` 输出到新目录

示例：

```powershell
cd C:\Users\Administrator\Desktop\zhiji_project\IM_Wsp\rearview_mirror
python -m PyInstaller --clean --noconfirm server\rearview_server.spec --workpath build --distpath dist_new
```

### 8.2 检测失败：找不到 run_algorithm_worker.py

应使用当前修复后的 spec 和源码重新打包。修复后的 exe 使用同一个 exe 的 `--rearview-worker` 模式启动检测子进程，不再依赖临时目录中的 `run_algorithm_worker.py` 文件。

### 8.3 output 没有生成到 exe 同级目录

应确认使用的是修复后的新版 exe。新版服务会向算法子进程传入：

```text
REARVIEW_OUTPUT_DIR=<exe所在目录>\output
```

因此输出应生成在 exe 同级的 `output` 目录。

### 8.4 CATIA 无法连接

确认：

- 本机已安装 CATIA
- CATIA COM 自动化可用
- CATIA 已打开，或前端点击“打开 CATIA”
- 当前用户权限可以访问 CATIA COM
- 上传的是有效 `.CATPart` 文件

## 9. 推荐发布方式

建议将 exe 放在固定、有写权限的目录，例如：

```text
D:\RearviewServer\rearview_server.exe
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
