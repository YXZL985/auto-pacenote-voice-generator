<div align="center">

# 🎙️ 自动化角色路书语音包生成器

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/)
[![Playwright](https://img.shields.io/badge/Playwright-✓-green.svg)](https://playwright.dev/)
[![GPT-SoVITS](https://img.shields.io/badge/GPT--SoVITS-✓-orange.svg)](https://github.com/RVC-Boss/GPT-SoVITS)
[![License](https://img.shields.io/badge/License-GPLv3-red.svg)](LICENSE)

一个基于 Playwright 的自动化工具，用于批量生成并保存角色路书语音包音频文件。

🎮 适用于：[ZTMZ路书工具](https://gitee.com/ztmz/ztmz_pacenote)

</div>

---

## 功能特性

- **自动化语音合成**：使用 Playwright 控制浏览器自动操作 `GPT-SoVITS 推理 WebUI `，生成语音音频
- **批量处理**：从 Excel 文件读取文本和文件名配置，批量生成音频
- **智能下载**：自动监听下载事件并整理文件到指定目录
- **元数据生成**：自动生成语音包 info.json 元数据文件
- **特殊文件处理**：支持特殊文件名（如 `end_stage`、`start_stage` 等）自动创建子目录

## 环境要求

- Python 3.7+
- Playwright
- openpyxl
- 本机运行的的 GPT-SoVITS WebUI

## 配置环境
建议使用Miniconda部署环境。
创建环境：
```bash
conda create -n auto_media_pack_acg python=3.7
```
激活环境：
```bash
conda activate auto_media_pack_acg
```
安装依赖：
```bash
pip install playwright openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple/
playwright install chromium
```

## 使用说明

### 1. 配置准备

在运行脚本前，请确保已正确配置以下参数（位于 `main.py` 文件开头）：

| 配置项 | 说明 | 默认值 |
|--------|------|--------|
| `EXCEL_FILE_PATH` | Excel 文件路径 | `D:\devfiles\auto_media_pack_acg\pacenote_view.xlsx` |
| `SHEET_NAME` | Excel 工作表名称 | `pacenote_view_202412300958` |
| `AUDIO_FILES` | 参考音频文件路径列表 | 需自行配置 |
| `DROPDOWN_OPTION_1` | GPT 模型选择 | `GPT_weights_v2ProPlus/Fugue-e12.ckpt` |
| `DROPDOWN_OPTION_2` | SoVITS 模型选择 | `SoVITS_weights_v2ProPlus/Fugue_e8_s384.pth` |
| `INITIAL_INPUT_TEXT` | 参考音频对应的文本 | 需自行配置 |
| `CONTROL_TARGET_VALUE` | top_k 参数值 | `37` |

### 2. 运行程序
在运行开始之前，需要训练好 GPT-SoVITS 模型，确保模型可在下拉列表被选择，确保`TTS推理WebUI`已经可以被打开

至于如何训练GPT-SoVITS模型，可以参考[官方文档](https://www.yuque.com/baicaigongchang1145haoyuangong/ib3g1e)。

在运行开始之前，需要确保TTS推理WebUI已经可以被打开，端口号默认为9872。

在运行开始之前，需要确保参考音频文件路径列表中包含所有需要合成的音频文件。

在运行开始之前，需要确保参考文本对应的文本文件中包含所有需要合成的文本。

```bash
python main.py
```

### 3. 交互式输入

程序启动后会提示输入以下信息：

1. **本地端口号**：`GPT-SoVITS 推理 WebUI `的端口号（例如：`9872`）
2. **输出目录路径**：语音包保存的基础目录（例如：`D:\output`）
3. **角色名字**：语音包角色名称（例如：`Elysia（爱莉希雅）`）
4. **作者信息**：语音包作者（例如：`by_燕戏竹林`）

### 4. Excel 文件格式

Excel 文件需包含以下列：

- **C 列**：文件名（多个文件名可用逗号分隔，取第一个）
- **D 列**：要合成的文本内容

数据范围：第 2 行至第 347 行

仓库中存在，可直接使用。
## 工作流程

1. **初始设置**：
   - 选择 GPT 和 SoVITS 模型
   - 上传参考音频文件
   - 输入参考文本
   - 设置 top_k 参数

2. **批量处理**：
   - 读取 Excel 中的文本和文件名
   - 逐个输入文本并点击生成
   - 等待语音合成完成
   - 自动下载并重命名文件

3. **输出整理**：
   - 按角色名创建输出目录
   - 保存所有音频文件
   - 生成 info.json 元数据文件

## 特殊文件名处理

以下文件名会自动创建子目录并将音频放入其中：

- `end_stage`
- `system_end_stage`
- `start_stage`
- `system_start_stage`

## 目录结构

```
输出目录/
└── 角色名/
    ├── info.json
    ├── audio1.wav
    ├── audio2.wav
    ├── end_stage/
    │   └── end_stage.wav
    └── start_stage/
        └── start_stage.wav
```

## 注意事项

1. **第一个音频文件**：不要超过 10 秒，否则可能导致生成失败
2. **网络连接**：确保 GPT-SoVITS WebUI 已启动并可访问
3. **浏览器窗口**：程序以非无头模式运行，便于调试和观察进度
4. **超时设置**：语音合成长时间未完成会自动跳过并记录错误

## 依赖项目

- [GPT-SoVITS](https://github.com/RVC-Boss/GPT-SoVITS) - 语音合成模型
- [Playwright](https://playwright.dev/) - 用于浏览器自动化操作
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/) - 用于读取 Excel 文件
- [ZTMZ路书工具](https://gitee.com/ztmz/ztmz_pacenote) - 语音包的应用场景

## 鸣谢

感谢以下开源项目和工具的支持：

| 项目 | 说明 |
|------|------|
| [GPT-SoVITS](https://github.com/RVC-Boss/GPT-SoVITS) | 强大的少样本语音转换与语音合成工具 |
| [Playwright](https://playwright.dev/) | 微软出品的现代 Web 自动化测试框架 |
| [openpyxl](https://openpyxl.readthedocs.io/en/stable/) | Python 读写 Excel 文件的库 |
| [ZTMZ路书工具](https://gitee.com/ztmz/ztmz_pacenote) | 优秀的赛车游戏路书语音工具 |

## ⚠️ 免责声明

> **本工具仅供学习和研究使用。使用本工具即表示您同意以下条款：**

1. **知识产权**：用户使用本工具生成的语音内容，应确保拥有原始音频素材的合法使用权。禁止未经授权使用他人受版权保护的声音进行训练或生成内容。

2. **合法使用**：用户承诺不将本工具用于任何违法、侵权或欺诈目的，包括但不限于：
   - 冒充他人身份进行诈骗或误导
   - 生成虚假信息或深度伪造内容用于非法目的
   - 侵犯他人肖像权、声音权或其他合法权益

3. **责任限制**：本工具按"原样"提供，作者不对因使用本工具而产生的任何直接或间接损失承担责任，包括但不限于：
   - 因使用生成内容导致的法律纠纷
   - 因模型训练或生成过程中的技术问题导致的数据丢失
   - 因违反当地法律法规而产生的后果

4. **合规性**：用户应自行确保使用本工具的行为符合所在国家/地区的法律法规。如因使用本工具违反相关法律，用户应自行承担全部法律责任。

5. **第三方服务**：本工具依赖 GPT-SoVITS 等第三方开源项目，用户应同时遵守这些项目的使用条款和许可协议。

## License

本项目采用 [GPLv3](LICENSE) 开源许可证。

> **注意**：GPLv3 许可证仅适用于本工具的代码本身，不适用于用户使用本工具生成的内容。用户对其生成的内容负有完全责任。
