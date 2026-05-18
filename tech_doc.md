# 断面数据处理系统 - 技术文档

## 1. 项目概述

### 1.1 项目简介
断面数据处理系统是一款用于水利工程断面测量数据处理的专业工具，提供CSV数据转换、成果表生成、库容计算等功能。

### 1.2 技术栈

| 分类 | 技术 | 版本 |
|------|------|------|
| 语言 | Python | 3.12+ |
| UI框架 | PySide6 | 6.6.0+ |
| Excel处理 | openpyxl | 3.1.2+ |
| 数值计算 | numpy | 1.26.3+ |

### 1.3 架构设计

```
┌─────────────────────────────────────────────────────────────────┐
│                        应用层 (UI)                            │
│  MainWindow | Views | Widgets | Styles                        │
├─────────────────────────────────────────────────────────────────┤
│                        业务层 (Core)                           │
│  Services | Models | Events                                   │
├─────────────────────────────────────────────────────────────────┤
│                        数据层 (Modules)                        │
│  CSV转换 | 高程转换 | 库容计算 | 数据检查                       │
├─────────────────────────────────────────────────────────────────┤
│                        基础层                                  │
│  Config | Utils | Logging                                     │
└─────────────────────────────────────────────────────────────────┘
```

## 2. 目录结构

```
src/
├── app/
│   ├── core/                  # 核心业务层
│   │   ├── __init__.py
│   │   ├── models.py          # 数据模型
│   │   ├── services.py        # 业务服务
│   │   └── events.py          # 事件系统
│   ├── ui/                    # 用户界面层
│   │   ├── __init__.py
│   │   ├── main_window.py     # 主窗口
│   │   ├── styles.py          # 样式管理
│   │   ├── widgets.py         # 自定义组件
│   │   └── views/             # 视图模块
│   │       ├── __init__.py
│   │       ├── base_view.py
│   │       ├── settings_view.py
│   │       ├── conversion_view.py
│   │       ├── processing_view.py
│   │       └── check_view.py
│   ├── modules/                # 数据处理模块
│   │   ├── csv_to_report.py
│   │   ├── report_to_85.py
│   │   ├── header_fill.py
│   │   ├── auto_plot.py
│   │   ├── folder_integration.py
│   │   ├── storage_calculation.py
│   │   ├── data_check.py
│   │   └── plot_summary.py
│   ├── utils/                  # 工具模块
│   ├── __init__.py
│   ├── config.py               # 配置管理
│   └── utils.py                # 通用工具函数
├── tests/                     # 测试模块
├── logs/                      # 日志目录
└── main_new.py                # 新架构入口
```

## 3. 核心模块说明

### 3.1 UI层

#### 3.1.1 MainWindow
- **职责**：主窗口管理，负责视图切换和整体布局
- **路径**：`src/app/ui/main_window.py`

#### 3.1.2 Views
- **SettingsView**：路径设置视图
- **ConversionView**：CSV转换视图  
- **ProcessingView**：成果处理视图
- **CheckView**：数据检查视图

#### 3.1.3 Widgets
- **NavButton**：导航按钮组件
- **ActionButton**：操作按钮组件
- **CardFrame**：卡片容器组件
- **LogWidget**：日志显示组件
- **ProgressWidget**：进度显示组件
- **StatsCard**：统计卡片组件

### 3.2 Core层

#### 3.2.1 Models
| 模型 | 说明 |
|------|------|
| Point | 测点数据模型 |
| Section | 断面数据模型 |
| BridgeData | 桥数据模型 |
| ProcessingResult | 处理结果模型 |
| ProjectConfig | 项目配置模型 |

#### 3.2.2 Services
| 服务 | 说明 |
|------|------|
| StorageCalculationService | 库容计算服务 |
| ConversionService | 数据转换服务 |
| DataCheckService | 数据检查服务 |

#### 3.2.3 Events
- **EventBus**：事件总线（单例模式）
- **ProgressTracker**：进度跟踪器
- **EventType**：事件类型枚举

### 3.3 Modules层

| 模块 | 功能 |
|------|------|
| csv_to_report.py | CSV转成果表 |
| report_to_85.py | 高程转85基准 |
| header_fill.py | 填写表头信息 |
| auto_plot.py | 自动生成图表 |
| folder_integration.py | 文件夹整合合并 |
| storage_calculation.py | 库容计算 |
| data_check.py | 数据检查 |
| plot_summary.py | 导出TXT |

## 4. 功能清单

### 4.1 路径设置
- [x] 根目录选择
- [x] 数据目录自动检测
- [x] 模板目录自动检测
- [x] 输出目录自动检测
- [x] 路径验证
- [x] 配置保存

### 4.2 CSV转成果表
- [x] 批量CSV文件加载
- [x] 横断面/纵断面分类统计
- [x] CSV转Excel转换
- [x] 高斯投影反算
- [x] 深泓点标记
- [x] 堤顶识别

### 4.3 成果处理
- [x] 转85高程
- [x] 填写表头
- [x] 自动成图
- [x] 导出TXT
- [x] 整合合并
- [x] 库容计算
- [x] 停止计算

### 4.4 数据检查
- [x] 特征点检查（左堤顶、右堤顶、深泓点）
- [x] 糙率检查
- [x] 起点距检查
- [x] 报告生成

## 5. 关键算法

### 5.1 高斯投影反算
```
用途：将平面坐标转换为经纬度
输入：x, y, central_meridian(中央子午线)
输出：纬度(B), 经度(L)
```

### 5.2 断面面积计算
```
用途：计算断面面积
方法：梯形法累加
输入：测点列表, 顶点高程
输出：断面面积
```

### 5.3 Haversine公式
```
用途：根据经纬度计算两点距离
输入：lon1, lat1, lon2, lat2
输出：距离(米)
```

### 5.4 断面排序
```
用途：按深泓点高程排序（下游→上游）
规则：深泓点高程升序排列
```

## 6. 数据流

### 6.1 CSV转换流程
```
CSV文件 → 读取数据 → 查找ZJ/YJ点 → 投影计算 → 排序 → 写入模板 → 成果表
```

### 6.2 库容计算流程
```
成果表 → 读取断面数据 → 提取深泓点 → 计算面积 → 排序断面 → 计算棱柱体体积 → 累加库容
```

## 7. 事件机制

### 7.1 事件类型
| 事件 | 说明 |
|------|------|
| FILE_LOADED | 文件加载完成 |
| PROCESS_STARTED | 处理开始 |
| PROCESS_COMPLETED | 处理完成 |
| PROCESS_PROGRESS | 处理进度更新 |
| ERROR_OCCURRED | 错误发生 |
| WARNING_OCCURRED | 警告发生 |
| STATUS_CHANGED | 状态变更 |

### 7.2 使用方式
```python
from app.core.events import EventBus, Event, EventType

event_bus = EventBus()

# 订阅事件
event_bus.subscribe(EventType.PROCESS_PROGRESS, handle_progress)

# 发布事件
event_bus.publish(Event(EventType.PROCESS_PROGRESS, data))
```

## 8. 配置管理

### 8.1 默认路径
```python
DEFAULT_DATA_DIR = os.path.join(ROOT_DIR, '断面')
DEFAULT_TEMPLATE_DIR = os.path.join(ROOT_DIR, '模板')
DEFAULT_OUTPUT_DIR = os.path.join(ROOT_DIR, '成果')
```

### 8.2 配置文件格式
```ini
ROOT_DIR=/path/to/project
DATA_DIR=/path/to/project/断面
OUTPUT_DIR=/path/to/project/成果
TEMPLATE_DIR=/path/to/project/模板
```

## 9. 日志系统

### 9.1 日志级别
- INFO：常规操作信息
- WARNING：警告信息
- ERROR：错误信息

### 9.2 日志路径
- 运行日志：`logs/app.log`
- 错误日志：`logs/error.log`

## 10. 测试覆盖

### 10.1 单元测试
- 断面面积计算测试
- 距离计算测试
- 数据模型验证测试

### 10.2 集成测试
- CSV转换完整流程测试
- 库容计算完整流程测试

### 10.3 UI测试
- 界面交互测试
- 路径设置测试

## 11. 打包部署

### 11.1 依赖安装
```bash
pip install -r requirements.txt
```

### 11.2 运行方式
```bash
python src/main_new.py
```

### 11.3 打包命令
```bash
pyinstaller --onefile --windowed --name "断面数据处理系统" src/main_new.py
```

## 12. 版本历史

| 版本 | 日期 | 变更 |
|------|------|------|
| v1.0.0 | 2024-XX-XX | 初始版本 |
| v2.0.0 | 2024-XX-XX | 现代化重构 |