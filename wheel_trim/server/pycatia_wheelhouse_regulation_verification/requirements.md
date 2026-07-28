# 车轮罩法规检测程序需求

## 1. 项目目标

车轮罩法规检测程序用于在本机 CATIA 环境中自动完成车轮罩与轮胎的匹配、法规轴线和截面构造、轮罩包围盒测量、轮罩到轴距离测量、截面轮廓高度测量、3D 标注、截图和 Word 报告输出。

程序需要支持 Web 工作台调用，也需要保留算法脚本直接运行能力。当前实现允许一次传入 1-4 个轮罩零件：

- 左前轮罩：`Left_Front_Wheelhouse`
- 右前轮罩：`Right_Front_Wheelhouse`
- 左后轮罩：`Left_Rear_Wheelhouse`
- 右后轮罩：`Right_Rear_Wheelhouse`

车轮或轮胎装配仍作为一个完整文件夹上传，并必须提供根 `CATProduct`。

## 2. 程序入口

### 2.1 Web 服务入口

后端入口：

```text
wheel_trim/server/server.py
```

前端入口：

```text
wheel_trim/web/index.html
```

子进程入口：

```text
wheel_trim/server/run_algorithm_worker.py
```

### 2.2 算法入口

算法主文件：

```text
wheel_trim/server/pycatia_wheelhouse_regulation_verification/main.py
```

主函数：

```python
main(
    front_wheelhouse_part_path=None,
    rear_wheelhouse_part_path=None,
    wheel_assembly_path=...,
    left_front_wheelhouse_part_path=None,
    right_front_wheelhouse_part_path=None,
    left_rear_wheelhouse_part_path=None,
    right_rear_wheelhouse_part_path=None,
)
```

兼容旧模式的 `front_wheelhouse_part_path`、`rear_wheelhouse_part_path` 仍保留，但 Web 当前使用四轮罩参数模式。

## 3. 输入规格

### 3.1 轮罩输入

轮罩输入为 `.CATPart` 文件，至少提供一个，最多提供四个。未提供的轮罩不参与计算，不应在检测报告中生成测量项或法规对比结果。

| 输入项 | 是否必填 | 文件类型 | 测量前缀 |
|---|---|---|---|
| 左前轮罩 | 可选 | `.CATPart` | `Left-Front` |
| 右前轮罩 | 可选 | `.CATPart` | `Right-Front` |
| 左后轮罩 | 可选 | `.CATPart` | `Left-Rear` |
| 右后轮罩 | 可选 | `.CATPart` | `Right-Rear` |

### 3.2 车轮装配输入

车轮装配以文件夹形式上传，必须指定根 `CATProduct`。后端会保存上传文件夹结构，并根据用户选择的根节点定位装配入口。

### 3.3 报告参数输入

Web 前端需要传入：

| 参数 | 说明 |
|---|---|
| `tire_radius_r` | 轮胎半径 r，用于 `c < 2r` 法规对比。 |
| `tire_width_y_b` | 轮胎 Y 向宽度 b，用于 `q > b` 法规对比。 |

## 4. 输出规格

每次运行会在 `output/output-时间戳/result-时间戳/` 下输出结果文件。

| 输出项 | 说明 |
|---|---|
| 校核 CATProduct | 根校核总成，命名为 `wheelhouse_regulation_verification_时间戳.CATProduct`。 |
| 法规轴线过程 CATPart | `Wheelhouse_Regulation_Axis_Lines_时间戳.CATPart`，保存轴线、切割面、包围盒线框等过程几何。 |
| 标注 CATPart | `Wheelhouse_Regulation_Annotations.CATPart`，保存法规测量点线标注几何。 |
| JSON 结果 | `wheelhouse_regulation_verification_result_时间戳.json`。 |
| Word 报告 | `wheelhouse_regulation_report_时间戳.docx`。 |
| 截图 | `screenshots/` 下按测量名输出 `.png` 图片。 |

## 5. 命名规范

CATIA `PartNumber` 和程序生成的文件名必须使用英文、数字和下划线，不使用中文、空格和特殊符号。

推荐命名：

| 对象 | 命名 |
|---|---|
| 根校核 CATProduct | `Wheelhouse_Regulation_Verification` |
| 车轮装配组件 | `Wheel_Assembly` |
| 左前轮罩组件 | `Left_Front_Wheelhouse` |
| 右前轮罩组件 | `Right_Front_Wheelhouse` |
| 左后轮罩组件 | `Left_Rear_Wheelhouse` |
| 右后轮罩组件 | `Right_Rear_Wheelhouse` |
| 法规轴线过程 Part | `Wheelhouse_Regulation_Axis_Lines_时间戳` |
| 法规标注 Part | `Wheelhouse_Regulation_Annotations` |

## 6. 坐标系统一要求

所有重心、轴线、包围盒、截面极值点、标注点和截图视点均以新建校核 CATProduct 的装配世界坐标系为基准。

要求：

- 叶子零件局部轴线方向必须通过装配位姿矩阵转换到世界坐标。
- 方向向量只使用旋转部分转换，不受平移影响。
- 重心和点坐标需要转换到根校核装配世界坐标。
- 多级 Product 嵌套时必须累乘从叶子实例到根校核 CATProduct 的所有位姿。
- 正反方向视为同轴，但生成几何时使用归一化后的稳定方向。

## 7. 核心检测流程

### 7.1 日志助手流程

前端日志助手当前展示 11 步流程：

1. 连接 CATIA
2. 打开输入文件
3. 筛选轮胎候选
4. 匹配轮罩和车轮
5. 创建轴线和切割面
6. 获取轮罩包围盒
7. 截面导出与极值点线
8. 法规距离测量
9. 法规标注
10. 法规截图
11. 保存校核结果

### 7.2 装配输入文件

程序创建新的校核 CATProduct，将本次提供的 1-4 个轮罩和车轮装配根节点装配进去。装配完成后会对根节点执行居中操作，便于后续查看。

### 7.3 轮胎候选筛选

程序遍历车轮装配下显示状态的叶子零件。隐藏零件和隐藏子装配不参与重心、近邻分组和包围盒测量。

轮胎候选识别逻辑：

- 优先使用圆线拓扑推导轮胎轴线。
- 显式轴特征作为增强信息和校验信息。
- 无显式轴但圆线拓扑满足条件的零件仍可作为候选。
- 同一空间车轮位置按轴线方向、轴线距离和重心距离聚类。
- 同一位置组内优先选择包围盒更大、拓扑评分更高、名称更像轮胎的零件作为 Tire 代表件。

### 7.4 轮罩和 Tire 匹配

程序先获取本次输入轮罩的重心，再将车轮装配内显示零件按距离轮罩重心 400mm 内进行近邻筛选。筛选后的零件才继续进入包围盒和轮胎代表件判断。

轮罩和 Tire 匹配根据装配世界坐标下的轮罩重心、Tire 重心、Tire 轴线距离和综合评分决定。仅本次输入的轮罩参与匹配。

### 7.5 法规轴线和切割面

对每个匹配成功的轮罩创建一组法规轴线几何：

- 轴线段中心为 Tire 重心投影到 Tire 轴线上的点。
- 轴线段长度为 `2 * REGULATION_AXIS_HALF_LENGTH`，默认 1000mm。
- 创建 0deg 截面平面和 -30deg 截面平面。
- 截面平面名称和测量名按轮罩前缀生成，例如 `Left-Front-p`、`Left-Front-p30`。

### 7.6 轮罩包围盒

轮罩包围盒使用远平面极值法测量，只测可见主体。包围盒结果用于法规 `q` 值，当前 `q` 取轮罩包围盒 Y 方向宽度。

包围盒标注点取 Y 向边两个端点，并优先选择 XYZ 坐标较大的候选边。包围盒标注偏移方向为 `(1, 0, 0)`。

### 7.7 截面导出和极值点

截面流程要求：

- 主流程先将轮罩工作文件复制到过程目录。
- 在单独打开的截面 CATPart 中创建截面曲线。
- 导出截面 CATPart 后，主程序单独 `Documents.Open(exported_path)`。
- 在独立 Part 文档中通过几何集名称和截面特征名称精确找到 SectionResult 曲线。
- 使用远平面 + 曲线到平面最小距离点方式取方向极值。
- 在截面 Part 中创建最高点、最低点和辅助连线；辅助连线默认隐藏，避免与绿色标注线重叠。
- 隐藏 `__SECTION_EXTREME_FAR_PLANES__` 几何图形集。
- `part.Update()`、`document.Save()`、`document.Close()` 后，再将保存好的 CATPart 添加到 Product。

### 7.8 法规距离测量

当前法规测量项：

| 后缀 | 含义 | 测量值 |
|---|---|---|
| `q` | 轮罩包围盒 Y 向宽度 | `bbox_y_size` |
| `c` | 轮罩可见主体到 Tire 轴线的最小距离 | Body 到轴线最小距离 |
| `p` | 0deg 截面轮廓高度 | 两个截面极值点世界坐标 Z 值差 |
| `p30` | 30deg 截面轮廓高度 | 两个截面极值点世界坐标 Z 值差 |

报告法规对比：

- `q > b`
- `c < 2r`
- `p > 30mm`
- `p30 > 30mm`

报告仅展示本次实际输入轮罩对应的测量项，不显示未输入轮罩的 0 值或空结果。

### 7.9 标注和截图

标注和截图在测量过程中执行，不在最后统一显隐筛选。

每得到一组测量点后：

1. 在 `Wheelhouse_Regulation_Annotations.CATPart` 中创建点、线和 3D 文本。
2. 调整视角并截图。
3. 截图后隐藏本条标注点、线和文本。
4. 后续全部截图完成后，将轮罩、过程零件、标注点线和 3D 文本恢复为显示状态。

标注参数：

| 参数 | 值 |
|---|---|
| 颜色 | 绿色 `(0, 255, 0)` |
| 线宽 | 2 |
| 文本字号 | 8 |
| 文本偏移距离 | 150mm |
| 标注几何集 | `Wheelhouse Regulation Distance Annotations` |

截图参数：

| 类型 | 视线方向 | 上方向 | 视距 |
|---|---|---|---:|
| 包围盒 q | `(0, 0, -1)` | `(-1, 0, 0)` | 5000 |
| 轮罩到轴 c | `(0, -1, 0)` | `(0, 0, 1)` | 5000 |
| 截面 p/p30 | `(-1, 0, 0)` | `(0, 0, 1)` | 1000 |

### 7.10 Word 报告

后端根据 JSON 结果生成 Word 报告。报告包括：

- 报告信息。
- 轮胎半径 r 和轮胎 Y 向宽度 b。
- 车轮罩法规距离测量表。
- 法规对比结果表。
- 判定规则。
- 对应测量截图。

截图写入报告时要求紧凑排列，不强制一页一图。

## 8. Web 服务要求

### 8.1 API

| 接口 | 方法 | 说明 |
|---|---|---|
| `/` | GET | 返回前端页面。 |
| `/app.js` | GET | 返回前端脚本。 |
| `/style.css` | GET | 返回样式文件。 |
| `/verify` | POST | 上传轮罩、车轮装配文件夹和报告参数，启动检测。 |
| `/api/status` | GET | 查询当前会话状态。 |
| `/api/catia-status` | GET | 检查 CATIA 是否可连接。 |
| `/api/open-catia` | POST | 尝试启动或连接 CATIA。 |
| `/api/result` | GET | 返回最近一次 JSON 结果。 |
| `/api/download-result` | GET | 下载结果 CATProduct。 |
| `/api/download-process` | GET | 下载过程 CATPart。 |
| `/api/download-report` | GET | 下载 JSON 结果文件。 |
| `/api/download-docx-report` | GET | 下载 Word 法规校核报告。 |
| `/api/reset` | POST | 重置后端会话。 |
| `/ws` | WebSocket | 推送日志和状态。 |

### 8.2 运行目录

源码运行时输出在：

```text
wheel_trim/server/output
```

打包 exe 运行时输出在：

```text
<exe所在目录>/output
```

## 9. 打包要求

PyInstaller 配置文件：

```text
wheel_trim/server/wheelhouse_server.spec
```

打包命令：

```powershell
cd C:\Users\Administrator\Desktop\zhiji_project\IM_Wsp\wheel_trim
python -m PyInstaller --clean --noconfirm --workpath build --distpath dist_wheelhouse server\wheelhouse_server.spec
```

打包时需要包含：

- `web`
- `server.py`
- `run_algorithm_worker.py`
- `main.py`
- `section_curve_export_tool.py`
- `catia_picture_capture.py`
- `catia_annotation_tools.py`

## 10. 异常处理要求

- 未上传任何轮罩时，后端返回错误。
- 车轮装配根节点不是 `.CATProduct` 时，后端返回错误。
- 输入文件不存在或类型错误时，直接终止并提示具体路径或文件名。
- 车轮装配中未找到可分析零件时，终止程序。
- 车轮装配中未找到 Tire 代表件时，终止程序。
- 某条标注、截图或截面处理失败时，应记录错误并尽量继续处理其他测量项。
- worker 子进程异常退出时，后端返回最后日志和 traceback。
- CATProduct 保存失败时，仍返回已有 JSON、过程文件和警告信息。

## 11. 当前实现状态

| 功能 | 状态 |
|---|---|
| 1-4 个轮罩输入 | 已完成 |
| 车轮装配文件夹上传 | 已完成 |
| CATIA 状态检测和启动 | 已完成 |
| worker 子进程隔离 | 已完成 |
| 可见零件过滤 | 已完成 |
| Tire 候选筛选 | 已完成 |
| 轮罩与 Tire 匹配 | 已完成 |
| 法规轴线和切割面创建 | 已完成 |
| 可见主体包围盒 | 已完成 |
| 截面导出和极值点 | 已完成 |
| 法规距离测量 | 已完成 |
| 过程内标注和截图 | 已完成 |
| Word 报告 | 已完成 |
| 打包配置和说明 | 已完成 |
| 截图后完整恢复用户原始视角 | 待优化 |
