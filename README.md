# 抖音渠道分析看板 - GitHub Pages 自动化方案

## 项目简介

本项目实现：用户上传Excel到GitHub仓库 → GitHub Actions自动处理数据 → 生成HTML看板 → 部署到GitHub Pages固定网址。

## 目录结构

```
├── generate_report.py    # 核心数据处理脚本
├── template.html         # HTML看板模板
├── docs/                 # GitHub Pages部署目录
│   └── index.html        # 生成的看板页面
├── .github/
│   └── workflows/
│       └── deploy.yml    # GitHub Actions工作流
├── data/                 # 数据目录（存放Excel文件）
│   └── .gitkeep
├── README.md
└── requirements.txt
```

## 使用步骤

### 1. 创建GitHub仓库

1. 登录 GitHub，点击右上角 "+" → "New repository"
2. 仓库名称建议：`douyin-dashboard` 或 `channel-analysis`
3. 选择 Private（私有）或 Public（公开）均可
4. 点击 "Create repository"

### 2. 启用GitHub Pages

1. 进入仓库设置（Settings）
2. 左侧菜单找到 "Pages"
3. Source 部分选择：
   - Branch: `gh-pages`
   - Folder: `/ (root)`
4. 点击 "Save"
5. 等待1-2分钟，页面部署完成

### 3. 配置仓库

1. 将本项目所有文件上传到仓库根目录
2. 特别是 `.github/workflows/deploy.yml` 和 `generate_report.py`

### 4. 上传数据并自动发布

1. 将新的Excel数据文件上传到仓库的 `data/` 目录
2. GitHub Actions会自动触发构建
3. 等待约1-2分钟后，看板即可访问

访问地址格式：`https://[用户名].github.io/[仓库名]/`

例如：`https://yourusername.github.io/douyin-dashboard/`

### 5. 手动触发更新

如果GitHub Actions没有自动触发，可以：
1. 进入仓库 "Actions" 页面
2. 点击 "Generate and Deploy Dashboard"
3. 点击 "Run workflow" 手动执行

## 数据要求

### Excel文件要求

- 文件放在 `data/` 目录下
- 支持 .xlsx 格式
- 脚本会自动查找最新的 .xlsx 文件

### 关键数据列（列索引，0-based）

| 列索引 | 字段名 | 说明 |
|--------|--------|------|
| 0 | 日期 | 原始日期 |
| 1 | 中心 | 中心名称 |
| 2 | 主管 | 主管/团队名称 |
| 3 | 顾问 | 顾问姓名 |
| 5 | 家庭单回溯后订单进线时间 | 业务日期 |
| 7 | 是否成交 | 1=成交 |
| 17 | 例子id | 用于去重 |
| 18 | 家庭单回溯后订单号 | 订单维度去重 |
| 19 | 是否删除好友 | 是/否 |
| 20 | 例子是否有效 | 1=有效 |
| 23 | 渠道名称-回溯家庭单后 | 渠道来源 |
| 27 | 总规模保费 | 保费金额 |
| 52 | 对话回合数 | 沟通轮次 |
| 60 | 覆盖天次 | 覆盖天数 |
| 61 | 回复天次 | 回复天数 |
| 65 | 第一次电话 | 一通标记 |
| 66 | 信息收集 | 信息收集标记 |
| 112 | 是否预约一通 | 预约标记 |
| 114 | 是否二通 | 二通标记 |

### 渠道映射规则

根据实际数据，脚本会匹配以下渠道：

| 原始渠道值（渠道名称-回溯家庭单后列） | 映射后 |
|-----------------------------------|--------|
| 2.抖音投放 | 抖音投放 |
| 3.2抖音自然流 | 抖音自然流 |
| 1.小红书投放 | 小红书基准 |
| 3.1小红书自然流 | 小红书基准 |

> 注意：如果实际数据中的渠道名称不同，请修改 `generate_report.py` 中的 `CHANNEL_MAP` 字典。

## KPI指标说明

本看板计算以下8个核心指标：

### 均值类指标
1. **平均覆盖天次** - 覆盖天次的平均值
2. **平均回复天次** - 回复天次的平均值
3. **平均沟通回合数** - 对话回合数之和 / 进线量

### 比率类指标
4. **微信删除率** - 删除好友数量 / 总进线量 × 100%
5. **电话预约率** - 预约一通数量 / 进线量 × 100%
6. **一通率** - 第一次电话数量 / 进线量 × 100%
7. **二通率** - 二通数量 / 进线量 × 100%
8. **成交率** - 成交订单数 / 总订单数 × 100%（按家庭单回溯后订单号去重）

## 技术栈

- **数据处理**: Python 3.x + pandas + openpyxl
- **图表**: ECharts 5.5.0
- **托管**: GitHub Pages
- **自动化**: GitHub Actions

## 本地运行

如需本地调试：

```bash
# 安装依赖
pip install -r requirements.txt

# 运行脚本
python generate_report.py

# 查看输出
open docs/index.html
```

## 注意事项

1. **数据去重**: 按例子id去重，只保留"例子是否有效=1"的记录
2. **订单维度**: 成交率按"家庭单回溯后订单号"去重计算
3. **除零保护**: 所有比率计算都有除零保护
4. **数值精度**: 所有数值保留2位小数
5. **渠道匹配**: 确保Excel中的渠道名称在脚本的CHANNEL_MAP中有对应的映射

## 许可证

MIT License
