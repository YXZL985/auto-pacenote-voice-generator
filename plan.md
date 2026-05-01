# 自动化角色路书语音包生成器编写计划

## 项目概述

**目标**：编写一个 Python 3.11 Windows 桌面自动化程序，用于操作 Chrome 浏览器（本机端口 `9872` 上的网页），批量生成并下载音频文件。

**技术栈**：
- **语言**：Python 3.11
- **浏览器自动化框架**：Playwright（同步 API）
- **Excel/表格读取**：openpyxl
- **文件与目录操作**：`pathlib`、`shutil`、`os`
- **目标浏览器**：Chromium（Playwright 自带二进制，无需额外安装 ChromeDriver）
- **运行环境**：Windows

**代码风格要求**：
- 使用 Playwright 同步 API（`from playwright.sync_api import sync_playwright`），**不使用异步**
- 脚本应为一个独立的 `.py` 文件，所有逻辑在同一文件内完成（可包含辅助函数）
- 添加必要的注释、日志打印与异常处理
- 需要的地方添加适当的 `time.sleep()` 或等待逻辑，确保页面加载完成后再操作


## 第一步：环境准备与初始化

### 1.1 依赖安装

在使用 Playwright 前，必须先安装所需的依赖。请在终端执行以下命令：

```bash
pip install playwright
playwright install chromium
pip install openpyxl
```

> **说明**：`playwright install chromium` 会下载一个独立的 Chromium 浏览器二进制文件，不依赖系统已安装的 Chrome。
> 这部分在编写时不需要操作。请交给人类。

### 1.2 初始化浏览器，连接到目标网页（LOCALHOST_PORT）

```python
from playwright.sync_api import sync_playwright
import openpyxl
import os
import shutil
import time
from pathlib import Path

LOCALHOST_PORT = _____  # 这是一个交互式输入的变量
BASE_URL = f"http://localhost:{LOCALHOST_PORT}" 

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)  # headless=False 用于调试，正式可改为 True
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()
    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
```

> **占位符**：`LOCALHOST_PORT` 需填入实际端口号。
> **注意**：`accept_downloads=True` 是启用文件下载功能的前提。


## 第二步：执行一次性手动操作

> **以下步骤仅执行一次，用于完成页面的初始设置。所有 CSS 选择器/XPath 均需根据实际页面结构调整。**

### 2.1 选择第一个下拉控件，选择指定项

**UI 元素**：第一个下拉菜单（CSS 选择器待填入）
**操作**：从下拉菜单中，按可见文本选择指定项

**Playwright 代码模式**：
```python
# 方式 A：如果是标准 <select> 元素，按可见文本选择
page.select_option("CSS_SELECTOR_1", label="TARGET_TEXT_1")

# 方式 B：如果非标准下拉（如 Ant Design / Element UI 组件），需先点击展开，再点击选项
page.click("CSS_SELECTOR_1")
page.wait_for_selector("CSS_SELECTOR_1_OPTION")
page.click("CSS_SELECTOR_1_OPTION")
```

**占位符**：
- `CSS_SELECTOR_1`：实际为`#component-5 > div.svelte-vomtxz.container > div > div.wrap-inner.svelte-vomtxz`这个选择器，请在代码中直接使用。
- `TARGET_TEXT_1`：实际为`GPT_weights_v2ProPlus/Fugue-e12.ckpt`

> **提示**：Playwright 选择下拉项有三种方式：按 `label`（可见文本）、按 `value`、按 `index`。优先按 label 选择。

### 2.2 选择第二个下拉控件，选择指定项

**UI 元素**：第二个下拉菜单（CSS 选择器待填入）
**操作**：与 2.1 类似

**Playwright 代码模式**：
```python
page.select_option("CSS_SELECTOR_2", label="TARGET_TEXT_2")
```

**占位符**：
- `CSS_SELECTOR_2`：实际为`#component-6 > div.svelte-vomtxz.container > div > div.wrap-inner.svelte-vomtxz`这个选择器，请在代码中直接使用。
- `TARGET_TEXT_2`：实际为`SoVITS_weights_v2ProPlus/Fugue_e8_s384.pth`

### 2.3 点击上传按钮，选择指定的音频文件上传

**UI 元素**：上传按钮（按钮或 input[type=file]，选择器待填入）
**操作**：点击上传按钮并选择本地音频文件

**Playwright 代码模式**：
```python
# 如果是标准的 input[type=file] 元素，直接使用 set_input_files
page.set_input_files("UPLOAD_INPUT_SELECTOR", ["AUDIO_FILE_PATH_1"])

# 如果上传按钮非标准 input，需使用 file_chooser 事件
with page.expect_file_chooser() as fc_info:
    page.click("UPLOAD_BUTTON_SELECTOR")
file_chooser = fc_info.value
file_chooser.set_files(["AUDIO_FILE_PATH_1"])
```

对于多文件上传，可以传入文件路径列表：
```python
page.locator("UPLOAD_INPUT_SELECTOR").set_input_files(["file1.mp3", "file2.mp3"])
```

**占位符**：
- `UPLOAD_INPUT_SELECTOR` 或 `UPLOAD_BUTTON_SELECTOR`：上传元素的选择器，实际为`#component-11 > div.audio-container.svelte-cbyffp > button > div`这个选择器，请在代码中直接使用。
- `AUDIO_FILE_PATH_1`：第一个音频文件的绝对路径或相对路径，实则为`"E:\intomedia\fugue\c41d585d86093789287d57fd257d0080_5386240350757193136.wav"`

### 2.4 等待指定文本框出现，输入指定的文本

**UI 元素**：文本输入框（CSS 选择器待填入）
**操作**：等待文本框可见后，输入指定文本

**Playwright 代码模式**：
```python
# 等待文本框出现
page.wait_for_selector("TEXT_INPUT_SELECTOR", state="visible", timeout=10000)
# 清空已有内容并输入新文本
page.fill("TEXT_INPUT_SELECTOR", "INPUT_TEXT_1")
```

**占位符**：
- `TEXT_INPUT_SELECTOR`：文本框的 CSS 选择器,实际为`#component-15 > label > textarea`这个选择器，请在代码中直接使用。
- `INPUT_TEXT_1`：要输入的文本内容，实则为`旅途可还顺利？若是得闲，我这又来了一批好茶，等着列位恩公登门品鉴。`

> **说明**：`page.fill()` 会自动清空现有内容后输入。如果需要追加，请改用 `page.type()`. `wait_for_selector()` 默认等待元素可见。如遇到“element is not visible”错误，可以增加 `timeout` 参数。

### 2.5 再次点击上传按钮，上传多个音频文件

**操作**：复用 2.3 的逻辑，传入多个音频文件路径列表

**Playwright 代码模式**：
```python
audio_files = [
    "AUDIO_FILE_PATH_2",
    "AUDIO_FILE_PATH_3",
    # ... 更多文件
]
page.locator("UPLOAD_INPUT_SELECTOR").set_input_files(audio_files) # 这一步的选择器是`#component-20 > button > div`
```

**占位符**：`AUDIO_FILE_PATH_2`、`AUDIO_FILE_PATH_3` 等需替换为实际路径。如`"E:\intomedia\fugue\7d813ee0ec4ef7377392344d0d4bc4d5_5409923193007825143.wav"``"E:\intomedia\fugue\cb1d0458aadd7b2ec6a4059c11f64d22_2930146964451322386.wav"``e:\intomedia\fugue\d3d7b0d0aca94c83dd0388d38aa2c77c_2965525889137205769.wav``"E:\intomedia\fugue\fb5bce40d8c1bb3bf7770020844c45da_1091382152274400042.wav"`

### 2.6 调整指定控件为指定数值

**UI 元素**：滑块/数值输入框/下拉等控件（CSS 选择器待填入）
**操作**：将控件值设置为指定数值

**Playwright 代码模式**：
```python
# 如果是指定数值的输入框
page.fill("CONTROL_SELECTOR", "TARGET_VALUE")

# 如果是滑块（slider），可能需要 JS 直接设置 value
page.evaluate("CONTROL_SELECTOR => { document.querySelector('CONTROL_SELECTOR').value = TARGET_VALUE; }")
```

**占位符**：
- `CONTROL_SELECTOR`：控件的 CSS 选择器,实际为`#component-40 > div.wrap.svelte-pc1gm4 > div > input`这个选择器，请在代码中直接使用。
- `TARGET_VALUE`：目标数值，实则为`37`


## 第三步：批量循环操作（核心重复流程）

### 3.1 读取 Excel 文件

**Excel 结构**：
- 文件路径：`EXCEL_FILE_PATH`
- 工作表名：`SHEET_NAME`
- D 列（第4列），第 2~347 行：每个音频的输入文本（D2 到 D347，共 346 行）
- C 列（第3列），第 2~347 行：每个音频的目标文件名（C2 到 C347，共 346 行）
- 起始行：2，结束行：347

**Playwright 代码模式**：
```python
workbook = openpyxl.load_workbook("EXCEL_FILE_PATH")
sheet = workbook["SHEET_NAME"]

input_texts = []   # 从 D 列读取的文本
file_names = []    # 从 C 列读取的文件名

for row in range(2, 348):  # D2 ~ D347 和 C2 ~ C347
    input_text = sheet.cell(row=row, column=4).value  # D 列
    file_name = sheet.cell(row=row, column=3).value   # C 列
    if input_text is None or file_name is None:
        continue
    input_texts.append(str(input_text))
    file_names.append(str(file_name))

workbook.close()
```

**占位符**：
- `EXCEL_FILE_PATH`：Excel 表格文件的路径，`"D:\devfiles\auto_media_pack_acg\pacenote_view.xlsx"`
- `SHEET_NAME`：工作表名称,实际为`pacenote_view_202412300958`

### 3.2 批量生成音频（循环）

**操作**：对每一行（共 346 行），依次执行：

1. **在指定控件输入指定文本**（文本取自 D 列）
2. **点击指定按钮，生成音频**
3. **等待下载按钮出现，点击下载按钮**
4. **将下载文件命名成 C 列指定的文件名**
5. **将文件转移到指定目标目录**

**完整循环代码模式**：

```python
DEST_DIR = Path("OUTPUT_DIRECTORY")  # 最终输出目录
os.makedirs(DEST_DIR, exist_ok=True)

DOWNLOAD_DIR = Path("TEMP_DOWNLOAD_DIR")  # 临时下载目录
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

for i, (input_text, output_filename) in enumerate(zip(input_texts, file_names)):
    print(f"[{i+1}/{len(input_texts)}] 正在处理: {output_filename}")

    # Step 1：在输入控件中输入 D 列的文本
    page.fill("INPUT_CONTROL_SELECTOR", input_text)

    # Step 2：点击生成音频的按钮
    page.click("GENERATE_BUTTON_SELECTOR")

    # Step 3：等待下载按钮出现，然后监听下载
    with page.expect_download() as download_info:
        page.wait_for_selector("DOWNLOAD_BUTTON_SELECTOR", state="visible", timeout=30000)
        page.click("DOWNLOAD_BUTTON_SELECTOR")

    download = download_info.value
    # 等待下载完成并获取路径
    downloaded_path = download.path()

    # Step 4：重命名并转移到输出目录
    # 注意：output_filename 可能来自 Excel，需要清理非法字符（如有必要）
    safe_filename = "".join(c for c in output_filename if c not in r'<>:"/\|?*')
    final_path = DEST_DIR / safe_filename
    shutil.move(downloaded_path, final_path)
    print(f"   已保存: {final_path}")

    # 短暂等待，避免操作过快
    time.sleep(0.5)

print(f"处理完成。所有音频文件已保存到 {DEST_DIR}")
```

**关键 API 说明**：

- **下载处理**：`page.expect_download()` 必须在触发下载的 **click 之前** 设置。`download_info.value` 获取 `Download` 对象。调用 `download.path()` 会等待下载完成后返回临时文件路径。
- **文件重命名与保存**：`download.path()` 返回的是浏览器生成的临时 GUID 文件名。`download.suggested_filename` 可获取服务器返回的建议文件名。
- **元素等待**：`page.wait_for_selector("selector", state="visible")` 确保下载按钮已出现且可见。
- **网络空闲等待**：如果页面在点击后需要等待 API 响应，可在点击生成按钮后添加：
- **关于下载按钮出现时机**：通常会在点击生成按钮后等待一些时间出现，不要刷新网页。
- **关于输出文件名**：C例中的文件名通常是一格内有多个并列，使用`,`分隔，任选其一即可。

```python
page.wait_for_load_state("networkidle")
```

**占位符**：
- `INPUT_CONTROL_SELECTOR`：输入框的 CSS 选择器,实际为`#component-28 > label > textarea`这个选择器，请在代码中直接使用。
- `GENERATE_BUTTON_SELECTOR`：生成按钮的 CSS 选择器,实际为`#component-47`这个选择器，请在代码中直接使用。
- `DOWNLOAD_BUTTON_SELECTOR`：下载按钮的 CSS 选择器,实际为`#component-48 > div.icon-buttons.svelte-rvdo70 > a > button`这个选择器，请在代码中直接使用。
- `OUTPUT_DIRECTORY`：最终输出目录的路径，请设计为交互式由用户输入
- `TEMP_DOWNLOAD_DIR`：临时下载目录的路径（可选，Playwright 自带临时目录管理机制），或者使用脚本所在目录


## 第四步：验证

循环结束后：
1. 检查 `OUTPUT_DIRECTORY` 中是否包含 346 个文件
2. 检查各个文件名是否与表格 C 列中的值匹配
3. 检查文件是否可以正常播放

**验证代码模式**：
```python
output_files = list(DEST_DIR.iterdir())
print(f"输出目录实际文件数: {len(output_files)}")
assert len(output_files) == 346, f"期望 346 个文件，实际 {len(output_files)}"
print("验证通过")
```


## 注意事项与最佳实践

### 文件下载
- 浏览器上下文必须启用 `accept_downloads=True`。
- `expect_download()` 必须在触发下载动作之前注册。
- 下载的文件在浏览器上下文关闭后会被删除，因此必须在关闭前保存到目标目录。
- 如果页面在点击下载按钮后，并非直接下载而是在新标签页中进行，Playwright 将识别为弹窗并可能被 Chrome 过滤拦截。建议优先采用 `expect_download` 方案。

### 选择器与稳定性
- Playwright 的 **自动等待** 机制已内置：`click()`、`fill()`、`type()` 等操作会自动等待元素可见、可用且稳定。大多数情况下不需手动显式等待。
- 但在动态加载场景下，仍建议在关键节点使用 `page.wait_for_selector()` 或 `page.wait_for_load_state("networkidle")`。
- **推荐选择器优先级**：`getByRole` > `getByLabel` / `getByPlaceholder` > `getByTestId`（如有） > CSS 选择器 > XPath。

### 异常处理
- 为关键操作添加 `try...except`：
  ```python
  try:
      page.click("SELECTOR", timeout=10000)
  except Exception as e:
      print(f"操作失败: {e}")
      # 可选：截图以定位问题
      page.screenshot(path=f"error_{time.time()}.png")
  ```
- 如果操作频繁触发“stale element”错误，尝试在每次循环时重新获取元素。

### 性能
- 当 `headless=True` 时，脚本可后台运行，处理速度更快。建议开发和调试阶段使用 `headless=False`，最终稳定后切换为 True。
- 如果页面无复杂视觉交互，开启无头模式不会影响功能。

### 目录结构建议

```
project/
├── main.py                 # 主程序脚本
├── excel/
│   └── input.xlsx          # Excel 源数据表格
├── audio_source/
│   └── *.mp3               # 待上传的初始音频文件
├── temp_download/
│   └── (临时下载文件)       # Playwright 临时下载目录（可选）
└── output/
    └── (346 个最终音频文件)
```


## 占位符汇总清单

| 编号 | 占位符 | 用途 | 示例值 |
|------|--------|------|--------|
| 1 | `LOCALHOST_PORT` | 本机网页端口号 | `9872` |
| 2 | `CSS_SELECTOR_1` | 第 1 个下拉控件选择器 | `#dropdown1` |
| 3 | `TARGET_TEXT_1` | 第 1 个下拉目标项文本 | `选项 A` |
| 4 | `CSS_SELECTOR_2` | 第 2 个下拉控件选择器 | `#dropdown2` |
| 5 | `TARGET_TEXT_2` | 第 2 个下拉目标项文本 | `选项 B` |
| 6 | `UPLOAD_INPUT_SELECTOR` | 上传控件的选择器 | `input[type="file"]` |
| 7 | `AUDIO_FILE_PATH_1` | 上传的音频文件路径 | `./audio_source/audio1.mp3` |
| 8 | `TEXT_INPUT_SELECTOR` | 文本输入框选择器 | `#text-input` |
| 9 | `INPUT_TEXT_1` | 输入到文本框的内容 | `示例文本` |
| 10 | `AUDIO_FILE_PATH_2~N` | 多个音频文件路径 | `["./f1.mp3", "./f2.mp3"]` |
| 11 | `CONTROL_SELECTOR` | 数值控件选择器 | `#value-slider` |
| 12 | `TARGET_VALUE` | 目标数值 | `50` |
| 13 | `EXCEL_FILE_PATH` | Excel 表格文件路径 | `./excel/input.xlsx` |
| 14 | `SHEET_NAME` | 工作表名称 | `Sheet1` |
| 15 | `INPUT_CONTROL_SELECTOR` | 循环中输入框选择器 | `#text-input` |
| 16 | `GENERATE_BUTTON_SELECTOR` | 生成按钮选择器 | `#generate-btn` |
| 17 | `DOWNLOAD_BUTTON_SELECTOR` | 下载按钮选择器 | `#download-btn` |
| 18 | `OUTPUT_DIRECTORY` | 最终输出目录路径 | `./output` |
| 19 | `TEMP_DOWNLOAD_DIR` | 临时下载目录路径 | `./temp_download` |


