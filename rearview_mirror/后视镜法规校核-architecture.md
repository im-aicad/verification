%% 📌 后视镜法规校核 智能体架构图
%% 版本: v1.0 | 日期: 2026-07-28
%% 作者: Claude Code | 状态: ✅ 已发布
%% 技术栈: 本地 Web + FastAPI + WebSocket + Python Worker + CATIA V5 COM
%% 分层: 7+1（含技能层）

```mermaid
graph BT

  classDef appNode fill:#DBEAFE,stroke:#4A90D9,stroke-width:1.5px,color:#1E3A5F
  classDef orcNode fill:#CCFBF1,stroke:#2EA8B8,stroke-width:1.5px,color:#134E4A
  classDef ageNode fill:#EDE9FE,stroke:#8B5CF6,stroke-width:1.5px,color:#4C1D95
  classDef sklNode fill:#FCE7F3,stroke:#EC4899,stroke-width:1.5px,color:#831843
  classDef modNode fill:#FEF3C7,stroke:#F59E0B,stroke-width:1.5px,color:#92400E
  classDef tooNode fill:#FEE2E2,stroke:#EF4444,stroke-width:1.5px,color:#991B1B
  classDef datNode fill:#F1F5F9,stroke:#78716C,stroke-width:1.5px,color:#44403C
  classDef cross fill:#F8FAFC,stroke:#94A3B8,stroke-width:2px,stroke-dasharray:6 3,color:#475569
  classDef external fill:#FFFBEB,stroke:#D97706,stroke-width:1.5px,stroke-dasharray:4 2,color:#92400E

  subgraph Cross["⚡ 横切关注层 Cross-cutting"]
    direction TB
    CR0["🔍 实时日志\n流程状态推送"]:::cross
    CR1["🔒 输入防护\n文件与特征校验"]:::cross
    CR2["💬 异常处理\n运行锁与错误"]:::cross
  end

  subgraph L7["L7: 应用与交互层 Application"]
    direction LR
    L7_0["🖥️ 本地工作台\n上传与查看"]:::appNode
    L7_1["🖥️ 浏览器\nChrome/Edge"]:::external
  end

  subgraph L6["L6: 编排与路由层 Orchestration"]
    direction LR
    L6_0["🚪 FastAPI服务\n接口与静态页"]:::orcNode
    L6_1["📤 WebSocket\n实时日志"]:::orcNode
    L6_2["🎯 Worker进程\n隔离CATIA"]:::orcNode
  end


  subgraph L4["L4: 技能层 Skill"]
    direction LR
    L4_0["🎯 输入检查\n校验特征名"]:::sklNode
    L4_1["🎯 法规线构造\n生成法规点线"]:::sklNode
    L4_2["🎯 镜片搜索\n角度组合搜索"]:::sklNode
    L4_3["🎯 反射取点\n生成反射点"]:::sklNode
    L4_4["🎯 间隙校核\n边界距离判定"]:::sklNode
    L4_5["🎯 截图报告\n截图与docx"]:::sklNode
  end


  subgraph L2["L2: 工具与协议层 Tool & Protocol"]
    direction LR
    L2_0["🔌 文件IO\n读写本地文件"]:::tooNode
    L2_1["🔗 CATIA V5\n几何建模环境"]:::external
    L2_2["🔌 WindowsCOM\n驱动CATIA"]:::tooNode
    L2_3["🌐 屏幕捕获\n预览CATIA"]:::tooNode
  end

  subgraph L1["L1: 数据与记忆层 Data & Memory"]
    direction LR
    L1_0["💾 session文件\n后端会话"]:::datNode
    L1_1["🗄️ IndexedDB\n前端版本库"]:::datNode
    L1_2["📦 上传目录\nCATPart输入"]:::datNode
    L1_3["📚 输出文件\n结果与报告"]:::datNode
  end

  CR0 -.->|贯穿| CR1
  CR1 -.->|贯穿| CR2

  L7_0 -->|HTTP/gRPC| L6_0
  L7_1 -->|HTTP/gRPC| L6_0
  L2_0 -->|CRUD/检索| L1_0
  L2_0 -->|CRUD/检索| L1_1
  L2_0 -->|CRUD/检索| L1_2
  L2_0 -->|CRUD/检索| L1_3
  L6_0 -->|日志推送| L6_1
  L6_0 -->|启动检测| L6_2
  L6_2 -->|执行流程| L4_0
  L4_0 -->|通过后| L4_1
  L4_1 -->|构造后| L4_2
  L4_2 -->|最优角度| L4_3
  L4_3 -->|测量点| L4_4
  L4_4 -->|校核结果| L4_5
  L4_1 -->|COM自动化| L2_2
  L2_2 -->|控制| L2_1
  L4_5 -->|写结果| L2_0
  L7_0 -->|版本库| L1_1

```

## 架构图文字预览
```
  L7: 应用与交互层 Application    🖥️本地工作台 / 🖥️浏览器
  L6: 编排与路由层 Orchestration    🚪FastAPI服务 / 📤WebSocket / 🎯Worker进程
  L5: 智能体层 Agent    (此层为空)
  L4: 技能层 Skill    🎯输入检查 / 🎯法规线构造 / 🎯镜片搜索 / 🎯反射取点 / 🎯间隙校核 / 🎯截图报告
  L3: 模型层 Model / LLM    (此层为空)
  L2: 工具与协议层 Tool & Protocol    🔌文件IO / 🔗CATIA V5 / 🔌WindowsCOM / 🌐屏幕捕获
  L1: 数据与记忆层 Data & Memory    💾session文件 / 🗄️IndexedDB / 📦上传目录 / 📚输出文件
  ⚡ 横切: 🔍实时日志 / 🔒输入防护 / 💬异常处理
```
