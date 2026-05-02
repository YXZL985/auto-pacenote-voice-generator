# -*- coding: utf-8 -*-
"""
自动化角色路书语音包生成器
使用 Playwright 批量生成并下载音频文件
"""

from playwright.sync_api import sync_playwright
import openpyxl
import os
import shutil
import time
import json
import uuid
from pathlib import Path


# ==================== 配置常量 ====================

# 角色名字和作者信息将在main()中通过交互式输入获取
CHARACTER_NAME = ""
BY_AUTHOR = ""

# Excel 文件配置
EXCEL_FILE_PATH = r"D:\devfiles\auto_media_pack_acg\pacenote_view.xlsx"
SHEET_NAME = "pacenote_view_202412300958"

# 音频文件路径列表
AUDIO_FILES = [
    r"E:\intomedia\fugue\c41d585d86093789287d57fd257d0080_5386240350757193136.wav",
    r"E:\intomedia\fugue\7d813ee0ec4ef7377392344d0d4bc4d5_5409923193007825143.wav",
    r"E:\intomedia\fugue\cb1d0458aadd7b2ec6a4059c11f64d22_2930146964451322386.wav",
    r"E:\intomedia\fugue\d3d7b0d0aca94c83dd0388d38aa2c77c_2965525889137205769.wav",
    r"E:\intomedia\fugue\fb5bce40d8c1bb3bf7770020844c45da_1091382152274400042.wav",
]

# 第一个音频文件（单独上传）
FIRST_AUDIO_FILE = AUDIO_FILES[0]

# 初始文本输入内容
INITIAL_INPUT_TEXT = "旅途可还顺利？若是得闲，我这又来了一批好茶，等着列位恩公登门品鉴。"

# CSS 选择器
SELECTOR_DROPDOWN_1 = "#component-5 > div.svelte-vomtxz.container > div > div.wrap-inner.svelte-vomtxz"
SELECTOR_DROPDOWN_2 = "#component-6 > div.svelte-vomtxz.container > div > div.wrap-inner.svelte-vomtxz"
SELECTOR_UPLOAD_1 = "#component-11 > div.audio-container.svelte-cbyffp > button > div"
SELECTOR_TEXT_INPUT_1 = "#component-15 > label > textarea"
SELECTOR_UPLOAD_2 = "#component-20 > button > div"
SELECTOR_CONTROL = "#component-40 > div.wrap.svelte-pc1gm4 > div > input"
SELECTOR_INPUT_CONTROL = "#component-28 > label > textarea"
SELECTOR_GENERATE_BUTTON = "#component-47"
SELECTOR_DOWNLOAD_BUTTON = "#component-48 > div.icon-buttons.svelte-rvdo70 > a > button"

# 下拉选项文本
DROPDOWN_OPTION_1 = "GPT_weights_v2ProPlus/Fugue-e12.ckpt"
DROPDOWN_OPTION_2 = "SoVITS_weights_v2ProPlus/Fugue_e8_s384.pth"

# 控件目标值
CONTROL_TARGET_VALUE = "37"


# ==================== 辅助函数 ====================

def clean_filename(filename: str) -> str:
    """
    清理文件名中的非法字符
    """
    invalid_chars = r'<>:"/\|?*'
    return "".join(c for c in filename if c not in invalid_chars)


def read_excel_data(excel_path: str, sheet_name: str):
    """
    从 Excel 文件读取数据
    返回: (input_texts, file_names) 元组
    """
    print(f"正在读取 Excel 文件: {excel_path}")
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook[sheet_name]

    input_texts = []   # D 列文本
    file_names = []    # C 列文件名

    for row in range(2, 348):  # D2 ~ D347 和 C2 ~ C347
        input_text = sheet.cell(row=row, column=4).value  # D 列
        file_name = sheet.cell(row=row, column=3).value   # C 列

        if input_text is None or file_name is None:
            continue

        # C 列可能包含多个文件名（逗号分隔），取第一个
        file_name_str = str(file_name)
        if ',' in file_name_str:
            file_name_str = file_name_str.split(',')[0].strip()

        input_texts.append(str(input_text))
        file_names.append(file_name_str)

    workbook.close()
    print(f"成功读取 {len(input_texts)} 条数据")
    return input_texts, file_names


def select_dropdown_option(page, selector: str, option_text: str):
    """
    选择下拉菜单选项（适用于非标准下拉组件）
    """
    print(f"选择下拉选项: {option_text}")
    try:
        # 先点击下拉框展开
        page.click(selector)
        time.sleep(0.5)

        # 等待选项出现并点击
        option_selector = f"text={option_text}"
        page.wait_for_selector(option_selector, state="visible", timeout=10000)
        page.click(option_selector)
        time.sleep(0.5)
        print(f"  成功选择: {option_text}")
    except Exception as e:
        print(f"  选择下拉选项失败: {e}")
        raise


def upload_audio_files(page, selector: str, file_paths: list):
    """
    上传音频文件
    """
    print(f"上传音频文件: {len(file_paths)} 个")
    try:
        # 使用 file_chooser 事件处理上传
        with page.expect_file_chooser() as fc_info:
            page.click(selector)
        file_chooser = fc_info.value
        file_chooser.set_files(file_paths)
        time.sleep(1)
        print(f"  成功上传 {len(file_paths)} 个文件")
    except Exception as e:
        print(f"  上传文件失败: {e}")
        raise


def perform_initial_setup(page):
    """
    执行一次性初始设置
    """
    print("\n=== 开始初始设置 ===")

    # 2.1 选择第一个下拉控件
    print("\n步骤 2.1: 选择第一个下拉控件")
    select_dropdown_option(page, SELECTOR_DROPDOWN_1, DROPDOWN_OPTION_1)

    # 2.2 选择第二个下拉控件
    print("\n步骤 2.2: 选择第二个下拉控件")
    select_dropdown_option(page, SELECTOR_DROPDOWN_2, DROPDOWN_OPTION_2)

    # 2.3 上传第一个音频文件
    print("\n步骤 2.3: 上传第一个音频文件")
    upload_audio_files(page, SELECTOR_UPLOAD_1, [FIRST_AUDIO_FILE])

    # 2.4 在文本框输入指定文本
    print("\n步骤 2.4: 输入初始文本")
    page.wait_for_selector(SELECTOR_TEXT_INPUT_1, state="visible", timeout=10000)
    page.fill(SELECTOR_TEXT_INPUT_1, INITIAL_INPUT_TEXT)
    print(f"  已输入文本: {INITIAL_INPUT_TEXT}")
    time.sleep(0.5)

    # 2.5 上传多个音频文件
    print("\n步骤 2.5: 上传多个音频文件")
    # 上传除第一个外的其他音频文件
    remaining_files = AUDIO_FILES[1:]
    if remaining_files:
        upload_audio_files(page, SELECTOR_UPLOAD_2, remaining_files)

    # 2.6 调整控件数值
    print("\n步骤 2.6: 调整控件数值")
    page.fill(SELECTOR_CONTROL, CONTROL_TARGET_VALUE)
    print(f"  已设置控件值为: {CONTROL_TARGET_VALUE}")
    time.sleep(0.5)

    print("\n=== 初始设置完成 ===")


def process_batch(page, input_texts: list, file_names: list, output_dir: Path):
    """
    批量处理音频生成
    """
    print(f"\n=== 开始批量处理，共 {len(input_texts)} 条数据 ===\n")

    for i, (input_text, output_filename) in enumerate(zip(input_texts, file_names)):
        try:
            print(f"[{i+1}/{len(input_texts)}] 正在处理: {output_filename}")
            print(f"   输入文本: {input_text[:50]}{'...' if len(input_text) > 50 else ''}")

            # Step 1: 清空并输入新文本，确保状态干净
            page.fill(SELECTOR_INPUT_CONTROL, "")
            time.sleep(0.2)
            page.fill(SELECTOR_INPUT_CONTROL, input_text)
            time.sleep(0.5)
            print(f"   已输入文本，等待浏览器处理...")

            # Step 2: 点击生成按钮
            page.click(SELECTOR_GENERATE_BUTTON)
            print(f"   已点击生成按钮，等待语音合成完成...")

            # 等待网络空闲
            page.wait_for_load_state("networkidle")

            # Step 3: 等待下载按钮变为可用状态（生成完成的信号）
            # 使用更长的超时时间，因为语音合成可能需要较长时间
            page.wait_for_selector(SELECTOR_DOWNLOAD_BUTTON, state="visible", timeout=60000)
            print(f"   下载按钮已可用，语音合成完成")

            # 额外等待2秒，确保音频文件完全生成并稳定
            time.sleep(2)

            # Step 4: 点击下载按钮并监听下载
            with page.expect_download(timeout=30000) as download_info:
                page.click(SELECTOR_DOWNLOAD_BUTTON)

            download = download_info.value

            # 等待下载完成并获取路径
            downloaded_path = download.path()
            print(f"   下载完成: {downloaded_path}")

            # 验证下载成功
            if not downloaded_path or not os.path.exists(downloaded_path):
                raise Exception(f"下载文件不存在: {downloaded_path}")

            # Step 5: 重命名并转移到输出目录
            safe_filename = clean_filename(output_filename)
            # 确保文件名有 .wav 扩展名
            if not safe_filename.endswith('.wav'):
                safe_filename += '.wav'

            final_path = output_dir / safe_filename

            # 如果文件已存在，先删除
            if final_path.exists():
                final_path.unlink()

            shutil.move(str(downloaded_path), str(final_path))
            print(f"   已保存: {final_path}")

            # Step 6: 处理间隔，避免浏览器队列堆积
            time.sleep(1)

        except Exception as e:
            print(f"   处理失败 [{i+1}/{len(input_texts)}]: {e}")
            print(f"   目标文件名: {output_filename}")
            print(f"   输入文本: {input_text[:100]}{'...' if len(input_text) > 100 else ''}")
            # 截图保存以便调试
            error_screenshot = f"error_{i+1}_{int(time.time())}.png"
            try:
                page.screenshot(path=error_screenshot)
                print(f"   已保存错误截图: {error_screenshot}")
            except Exception as screenshot_error:
                print(f"   截图保存失败: {screenshot_error}")
            continue

    print(f"\n=== 批量处理完成 ===")


def verify_output(output_dir: Path, expected_count: int = 346):
    """
    验证输出结果
    """
    print(f"\n=== 验证输出结果 ===")
    output_files = list(output_dir.iterdir())
    actual_count = len(output_files)

    print(f"输出目录: {output_dir}")
    print(f"期望文件数: {expected_count}")
    print(f"实际文件数: {actual_count}")

    if actual_count == expected_count:
        print("✓ 验证通过：文件数量正确")
    else:
        print(f"✗ 验证失败：期望 {expected_count} 个文件，实际 {actual_count} 个")

    return actual_count == expected_count


def create_info_json(output_dir: Path, character_name: str, by_author: str):
    """
    创建语言包元数据文件 info.json
    """
    info_data = {
        "id": str(uuid.uuid4()),
        "name": character_name,
        "description": f"AI合成，{by_author}",
        "gender": "F",
        "language": "普通话",
        "homepage": "",
        "version": "1.0.0"
    }

    info_path = output_dir / "info.json"
    with open(info_path, 'w', encoding='utf-8') as f:
        json.dump(info_data, f, ensure_ascii=False, indent=4)

    print(f"  已创建 info.json: {info_path}")


def main():
    """
    主程序入口
    """
    global CHARACTER_NAME, BY_AUTHOR

    print("=" * 60)
    print("自动化角色路书语音包生成器")
    print("=" * 60)

    # 交互式输入配置
    print("\n请输入配置信息：")
    localhost_port = input("本地端口号 (例如: 9872): ").strip()
    output_base_directory = input("输出目录路径 (例如: D:\\output): ").strip()

    # 交互式输入角色信息
    print("\n请输入角色信息：")
    CHARACTER_NAME = input("角色名字 (例如: Elysia（爱莉希雅）): ").strip()
    BY_AUTHOR = input("作者信息 (例如: 燕戏竹林): ").strip()

    if not localhost_port or not output_base_directory:
        print("错误：端口号和输出目录不能为空")
        return

    if not CHARACTER_NAME:
        print("错误：角色名字不能为空")
        return

    if not BY_AUTHOR:
        print("错误：作者信息不能为空")
        return

    base_url = f"http://localhost:{localhost_port}"

    # 创建以角色名命名的输出文件夹
    output_dir = Path(output_base_directory) / CHARACTER_NAME

    # 创建输出目录
    os.makedirs(output_dir, exist_ok=True)
    print(f"\n输出目录: {output_dir}")

    # 读取 Excel 数据
    try:
        input_texts, file_names = read_excel_data(EXCEL_FILE_PATH, SHEET_NAME)
    except Exception as e:
        print(f"读取 Excel 文件失败: {e}")
        return

    # 启动 Playwright
    print(f"\n正在连接到: {base_url}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # headless=False 用于调试
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            # 访问目标网页
            page.goto(base_url)
            page.wait_for_load_state("networkidle")
            print("页面加载完成")
            time.sleep(2)

            # 执行初始设置
            perform_initial_setup(page)

            # 执行批量处理
            process_batch(page, input_texts, file_names, output_dir)

            # 验证结果
            verify_output(output_dir)

            # 创建 info.json 元数据文件
            print("\n=== 创建语言包元数据文件 ===")
            create_info_json(output_dir, CHARACTER_NAME, BY_AUTHOR)

            print(f"\n处理完成。所有音频文件已保存到: {output_dir}")

        except Exception as e:
            print(f"\n程序执行出错: {e}")
            # 保存错误截图
            error_screenshot = f"error_main_{int(time.time())}.png"
            try:
                page.screenshot(path=error_screenshot)
                print(f"已保存错误截图: {error_screenshot}")
            except:
                pass

        finally:
            # 关闭浏览器
            print("\n正在关闭浏览器...")
            context.close()
            browser.close()

    print("\n程序结束")


if __name__ == "__main__":
    main()
