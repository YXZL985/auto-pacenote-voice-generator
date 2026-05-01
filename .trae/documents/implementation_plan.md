# 自动化角色路书语音包生成器 - 实施计划

## 项目概述

根据 plan.md 编写一个 Python 3.11 Windows 桌面自动化程序，使用 Playwright 操作 Chrome 浏览器，批量生成并下载音频文件。

## 实施步骤

### 第一步：项目初始化与依赖配置

1. 检查工作目录 `d:\devfiles\auto_media_pack_acg` 结构
2. 创建主程序文件 `main.py`
3. 添加依赖说明（用户需手动执行 pip install）

### 第二步：编写主程序代码

按照 plan.md 的四个步骤实现：

#### 2.1 环境准备与初始化
- 导入必要库：`playwright.sync_api`, `openpyxl`, `os`, `shutil`, `time`, `pathlib`
- 添加交互式输入：LOCALHOST_PORT、OUTPUT_DIRECTORY
- 初始化 Playwright，连接到本地端口

#### 2.2 一次性手动操作
- 选择第一个下拉控件 (`#component-5 > div.svelte-vomtxz.container > div > div.wrap-inner.svelte-vomtxz`)，选择 `GPT_weights_v2ProPlus/Fugue-e12.ckpt`
- 选择第二个下拉控件 (`#component-6 > div.svelte-vomtxz.container > div > div.wrap-inner.svelte-vomtxz`)，选择 `SoVITS_weights_v2ProPlus/Fugue_e8_s384.pth`
- 上传第一个音频文件 (`E:\intomedia\fugue\c41d585d86093789287d57fd257d0080_5386240350757193136.wav`)
- 在文本框 (`#component-15 > label > textarea`) 输入指定文本
- 上传多个音频文件（5个文件路径列表）
- 调整控件 (`#component-40 > div.wrap.svelte-pc1gm4 > div > input`) 值为 37

#### 2.3 批量循环操作
- 读取 Excel 文件 (`D:\devfiles\auto_media_pack_acg\pacenote_view.xlsx`)，工作表 `pacenote_view_202412300958`
- 读取 C 列（文件名）和 D 列（输入文本），第 2-347 行
- 循环处理每一行：
  - 在输入框 (`#component-28 > label > textarea`) 输入 D 列文本
  - 点击生成按钮 (`#component-47`)
  - 等待下载按钮 (`#component-48 > div.icon-buttons.svelte-rvdo70 > a > button`) 出现并点击
  - 保存下载文件到 OUTPUT_DIRECTORY，使用 C 列文件名

#### 2.4 验证
- 检查输出目录文件数量是否为 346 个
- 验证文件名匹配

### 第三步：代码优化

1. 添加异常处理和日志输出
2. 添加适当的 `time.sleep()` 等待逻辑
3. 添加进度显示
4. 添加错误截图功能

### 第四步：初始化 Git 仓库并提交

1. 初始化 Git 仓库
2. 创建 `.gitignore` 文件（排除输出目录、临时文件等）
3. 添加并提交所有文件

## 文件结构

```
d:\devfiles\auto_media_pack_acg\
├── .trae\
│   └── documents\
│       └── implementation_plan.md  # 本计划文件
├── .gitignore                      # Git 忽略文件
├── main.py                         # 主程序
├── plan.md                         # 原始需求文档
└── pacenote_view.xlsx              # Excel 源数据（已存在）
```

## 关键配置项

| 配置项 | 值 |
|--------|-----|
| LOCALHOST_PORT | 用户输入 |
| OUTPUT_DIRECTORY | 用户输入 |
| Excel 路径 | `D:\devfiles\auto_media_pack_acg\pacenote_view.xlsx` |
| 工作表名 | `pacenote_view_202412300958` |
| 音频文件路径 | `E:\intomedia\fugue\` 下的 5 个文件 |

## 依赖安装命令（用户需手动执行）

```bash
pip install playwright
playwright install chromium
pip install openpyxl
```

## 实施完成标准

- [ ] main.py 文件完整实现所有功能
- [ ] 代码包含必要的注释和异常处理
- [ ] .gitignore 文件正确配置
- [ ] Git 仓库初始化并提交
