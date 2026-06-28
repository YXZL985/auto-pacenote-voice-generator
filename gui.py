# -*- coding: utf-8 -*-
"""
自动角色路书语音包生成器 - PySide6 图形界面

提供完整的 GUI 操作界面，替代原有的 CLI 交互方式。
在后台 QThread 中运行 Playwright 自动化流程，保持界面响应。

SPDX-License-Identifier: GPL-3.0-or-later
"""

import json
import os
import shutil
import sys
import threading
import time
import traceback
from pathlib import Path
from datetime import datetime

# PySide6
from PySide6.QtCore import (
    Qt, QThread, QObject, Signal, Slot, QSettings,
)
from PySide6.QtGui import (
    QAction, QFont, QTextCursor, QColor,
)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QFormLayout, QGroupBox, QScrollArea, QTabWidget, QTableWidget,
    QTableWidgetItem, QHeaderView, QLineEdit, QPushButton, QLabel,
    QProgressBar, QTextEdit, QListWidget, QListWidgetItem,
    QCheckBox, QMessageBox, QFileDialog, QDialog,
    QStatusBar, QMenuBar, QMenu, QSizePolicy, QAbstractItemView,
    QSplitter, QFrame, QTreeView,
)
from PySide6.QtGui import QStandardItemModel, QStandardItem

# 复用 main.py 中的纯函数和常量
from main import (
    clean_filename,
    read_excel_data,
    create_info_json,
    create_legal_notice_file,
    verify_output,
    SKIP_VOICE_GENERATION,
    LEGAL_CONFIRMATION_TEXT,
    SELECTOR_DROPDOWN_1, SELECTOR_DROPDOWN_2,
    SELECTOR_UPLOAD_1, SELECTOR_TEXT_INPUT_1,
    SELECTOR_UPLOAD_2, SELECTOR_CONTROL,
    SELECTOR_INPUT_CONTROL, SELECTOR_GENERATE_BUTTON,
    SELECTOR_DOWNLOAD_BUTTON,
    LEGAL_NOTICE_FILE,
)


# ==================== 常量 ====================

APP_ORG = "auto_media_pack_acg"
APP_NAME = "VoicePackGenerator"
SETTINGS_LEGAL_ACCEPTED = "legal/accepted"
SETTINGS_WINDOW_GEOMETRY = "window/geometry"
SETTINGS_WINDOW_STATE = "window/state"

# 设置步骤总数
SETUP_STEP_COUNT = 6


# ==================== 法律声明对话框 ====================

class LegalDisclaimerDialog(QDialog):
    """模态法律声明确认对话框，仅在首次运行时显示。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("重要法律声明")
        self.setMinimumSize(700, 520)
        self.setModal(True)
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)

        # 标题
        title = QLabel("重 要 法 律 声 明")
        title.setAlignment(Qt.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setStyleSheet("color: #c0392b;")
        layout.addWidget(title)

        # 法律文本（只读）
        self._text_edit = QTextEdit()
        self._text_edit.setReadOnly(True)
        self._text_edit.setFont(QFont("Consolas", 10))
        self._text_edit.setPlainText(
            "╔══════════════════════════════════════════════════════════════════════════════╗\n"
            "║                              重 要 法 律 声 明                                ║\n"
            "╠══════════════════════════════════════════════════════════════════════════════╣\n"
            "║                                                                              ║\n"
            "║  1. 【声音合法来源】您必须确保使用的参考音频已取得权利人合法授权               ║\n"
            "║                                                                              ║\n"
            "║  2. 【合规法规】本工具遵守《生成式人工智能服务管理暂行办法》                  ║\n"
            "║                                                                              ║\n"
            "║  3. 【使用责任】您对本工具生成的内容承担完全法律责任                          ║\n"
            "║                                                                              ║\n"
            "╠══════════════════════════════════════════════════════════════════════════════╣\n"
            "║  依据《民法典》第1023条，自然人的声音受人格权保护                             ║\n"
            "║  未经授权使用他人声音进行AI合成可能构成侵权，面临法律责任                     ║\n"
            "║  参考案例：北京互联网法院（2023）京0491民初12142号民事判决书（AI声音侵权案）  ║\n"
            "╚══════════════════════════════════════════════════════════════════════════════╝\n"
            "\n"
            "本工具仅供学习和研究使用。详细法律声明请参见 DISCLAIMER.md 文件。\n"
            "\n"
            '请仔细阅读上述法律声明。如果您已阅读并同意遵守上述条款，请勾选下方复选框并点击"接受"。\n'
        )
        layout.addWidget(self._text_edit)

        # 复选框
        self._checkbox = QCheckBox("我已阅读并同意遵守上述所有条款")
        self._checkbox.toggled.connect(self._on_checkbox_toggled)
        layout.addWidget(self._checkbox)

        # 按钮
        btn_layout = QHBoxLayout()
        self._accept_btn = QPushButton("接受并继续")
        self._accept_btn.setEnabled(False)
        self._accept_btn.clicked.connect(self.accept)
        self._accept_btn.setStyleSheet(
            "QPushButton { background-color: #27ae60; color: white; padding: 8px 24px; "
            "font-size: 14px; } QPushButton:disabled { background-color: #bdc3c7; }"
        )

        self._decline_btn = QPushButton("不同意，退出程序")
        self._decline_btn.clicked.connect(self._on_decline)

        btn_layout.addStretch()
        btn_layout.addWidget(self._accept_btn)
        btn_layout.addSpacing(16)
        btn_layout.addWidget(self._decline_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

    def _on_checkbox_toggled(self, checked: bool):
        self._accept_btn.setEnabled(checked)

    def _on_decline(self):
        QApplication.quit()
        sys.exit(0)


# ==================== Playwright 工作线程 ====================

class PlaywrightWorker(QObject):
    """在后台线程中运行所有 Playwright 自动化操作。"""

    # 日志信号
    log_message = Signal(str)
    log_error = Signal(str)

    # 初始设置进度 (step_number, description)
    setup_progress = Signal(int, str)

    # 批次进度
    batch_total = Signal(int)
    batch_progress = Signal(int, str)   # (index, filename)
    item_success = Signal(int, str)     # (index, filename)
    item_error = Signal(int, str, str)  # (index, filename, error_msg)
    item_skipped = Signal(int, str)     # (index, filename)

    # 生命周期
    browser_opened = Signal()
    browser_closed = Signal()
    error_occurred = Signal(str)

    # 完成信号
    finished = Signal(bool, str, int, int, int)  # (success?, output_dir, ok_count, fail_count, skip_count)
    verification_done = Signal(bool, int, int)    # (passed, expected, actual)

    def __init__(self, config: dict, parent=None):
        super().__init__(parent)
        self.config = config
        self._pause_event = threading.Event()
        self._pause_event.set()      # set = 未暂停
        self._cancel_event = threading.Event()
        self._cancel_event.clear()
        self._page = None
        self._browser = None
        self._context = None

    # ---- 暂停/恢复/停止 ----

    def pause(self):
        self._pause_event.clear()

    def resume(self):
        self._pause_event.set()

    def stop(self):
        self._cancel_event.set()
        self._pause_event.set()  # 解除暂停阻塞，使取消检查能执行

    # ---- 主入口 ----

    @Slot()
    def do_work(self):
        """工作线程入口——包含所有浏览器自动化逻辑。"""
        success_count = 0
        fail_count = 0
        skip_count = 0
        output_dir = None

        try:
            cfg = self.config
            self.log_message.emit("正在读取 Excel 文件...")
            input_texts, file_names = read_excel_data(cfg["excel_path"], cfg["sheet_name"])
            self.batch_total.emit(len(input_texts))
            self.log_message.emit(f"成功读取 {len(input_texts)} 条数据")

            # 创建输出目录
            output_dir = Path(cfg["output_directory"]) / cfg["character_name"]
            os.makedirs(output_dir, exist_ok=True)
            create_legal_notice_file(output_dir)
            self.log_message.emit(f"输出目录: {output_dir}")

            # 连接浏览器
            base_url = f"http://localhost:{cfg['port']}"
            self.log_message.emit(f"正在连接: {base_url}")

            # Playwright 上下文管理器（确保浏览器正确关闭）
            success_count, fail_count, skip_count = self._run_browser(
                input_texts, file_names, output_dir, base_url
            )

        except Exception as e:
            trace = traceback.format_exc()
            self.log_error.emit(f"发生严重错误: {e}")
            self.log_message.emit(trace)
            self.error_occurred.emit(str(e))
            if self._page:
                try:
                    self._page.screenshot(path=f"error_gui_{int(time.time())}.png")
                    self.log_message.emit("已保存错误截图")
                except Exception:
                    pass
        finally:
            self._cleanup_browser()
            self.finished.emit(
                not self._cancel_event.is_set(),
                str(output_dir) if output_dir else "",
                success_count, fail_count, skip_count,
            )

    # ---- 浏览器逻辑 ----

    def _run_browser(self, input_texts, file_names, output_dir, base_url):
        """创建并操作浏览器（在 Playwright 上下文内）。
        返回: (success_count, fail_count, skip_count) 元组。"""
        from playwright.sync_api import sync_playwright

        with sync_playwright() as p:
            self._browser = p.chromium.launch(headless=False)
            self._context = self._browser.new_context(accept_downloads=True)
            self._page = self._context.new_page()
            self.browser_opened.emit()

            # 导航
            self._page.goto(base_url)
            self._page.wait_for_load_state("networkidle")
            time.sleep(2)
            self.log_message.emit("页面加载完成")

            if self._cancel_event.is_set():
                self.log_message.emit("已在初始设置前取消")
                return 0, 0, 0

            # 初始设置
            self._perform_initial_setup(self._page)

            if self._cancel_event.is_set():
                self.log_message.emit("已在初始设置后取消")
                return 0, 0, 0

            # 批量处理
            success_count, fail_count, skip_count = self._process_batch(
                self._page, input_texts, file_names, output_dir
            )

            # 验证
            expected = len(input_texts)
            passed = verify_output(output_dir, expected)
            actual = len(list(output_dir.iterdir()))
            self.verification_done.emit(passed, expected, actual)

            # 元数据
            create_info_json(output_dir, self.config["character_name"], self.config["author"])
            self.log_message.emit("已创建 info.json 元数据文件")

            self.browser_closed.emit()

            return success_count, fail_count, skip_count

    def _cleanup_browser(self):
        """确保浏览器被关闭。"""
        if self._context:
            try:
                self._context.close()
            except Exception:
                pass
        if self._browser:
            try:
                self._browser.close()
            except Exception:
                pass
        self.log_message.emit("浏览器已关闭")

    def _perform_initial_setup(self, page):
        """执行一次性初始设置（6 步）。"""
        self.log_message.emit("=== 开始初始设置 ===")
        cfg = self.config

        # 步骤 1: 选择 GPT 模型
        self.setup_progress.emit(1, "选择 GPT 模型...")
        step_label = "步骤 1: 选择 GPT 模型"
        self.log_message.emit(step_label)
        self._select_dropdown_option(page, SELECTOR_DROPDOWN_1, cfg["dropdown_1"])

        if self._cancel_event.is_set():
            return

        # 步骤 2: 选择 SoVITS 模型
        self.setup_progress.emit(2, "选择 SoVITS 模型...")
        self.log_message.emit("步骤 2: 选择 SoVITS 模型")
        self._select_dropdown_option(page, SELECTOR_DROPDOWN_2, cfg["dropdown_2"])

        if self._cancel_event.is_set():
            return

        # 步骤 3: 上传第一个音频文件
        self.setup_progress.emit(3, "上传参考音频...")
        self.log_message.emit("步骤 3: 上传第一个音频文件")
        self._upload_audio_files(page, SELECTOR_UPLOAD_1, [cfg["audio_files"][0]])

        if self._cancel_event.is_set():
            return

        # 步骤 4: 输入参考文本
        self.setup_progress.emit(4, "输入参考文本...")
        self.log_message.emit("步骤 4: 输入参考文本")
        page.wait_for_selector(SELECTOR_TEXT_INPUT_1, state="visible", timeout=10000)
        page.fill(SELECTOR_TEXT_INPUT_1, cfg["initial_text"])
        self.log_message.emit(f"  已输入文本: {cfg['initial_text']}")
        time.sleep(0.5)

        if self._cancel_event.is_set():
            return

        # 步骤 5: 检查并上传多个音频文件
        self.setup_progress.emit(5, "上传附加音频...")
        self.log_message.emit("步骤 5: 检查并上传多个音频文件")
        remaining_files = cfg["audio_files"][1:]
        if remaining_files:
            try:
                page.wait_for_selector(SELECTOR_UPLOAD_2, state="visible", timeout=2000)
                self._upload_audio_files(page, SELECTOR_UPLOAD_2, remaining_files)
            except Exception:
                self.log_message.emit("  检测到当前音源模型只需要第一个音频，跳过第二次上传")
        else:
            self.log_message.emit("  没有额外的音频文件需要上传")

        if self._cancel_event.is_set():
            return

        # 步骤 6: 调整参数
        self.setup_progress.emit(6, "设置生成参数...")
        self.log_message.emit("步骤 6: 调整参数")
        page.fill(SELECTOR_CONTROL, cfg.get("control_value", ""))
        self.log_message.emit(f"  已设置参数值: {cfg.get('control_value', '')}")
        time.sleep(0.5)

        self.log_message.emit("=== 初始设置完成 ===")

    def _process_batch(self, page, input_texts, file_names, output_dir):
        """批量处理音频生成。"""
        success_count = 0
        fail_count = 0
        skip_count = 0

        self.log_message.emit(f"\n=== 开始批量处理，共 {len(input_texts)} 条数据 ===")

        for i, (input_text, output_filename) in enumerate(zip(input_texts, file_names)):
            # ---- 暂停检查 ----
            self._pause_event.wait()

            # ---- 取消检查 ----
            if self._cancel_event.is_set():
                self.log_message.emit("用户取消了批量处理")
                break

            self.batch_progress.emit(i, output_filename)

            try:
                self.log_message.emit(
                    f"[{i+1}/{len(input_texts)}] 正在处理: {output_filename}"
                )
                self.log_message.emit(
                    f"   输入文本: {input_text[:50]}{'...' if len(input_text) > 50 else ''}"
                )

                # 特殊文件名处理（创建子目录）
                is_special = output_filename in SKIP_VOICE_GENERATION
                if is_special:
                    dir_path = output_dir / output_filename
                    os.makedirs(dir_path, exist_ok=True)
                    self.log_message.emit(f"   特殊处理：创建子目录 {dir_path}")

                # 清空并输入文本
                page.fill(SELECTOR_INPUT_CONTROL, "")
                time.sleep(0.2)
                page.fill(SELECTOR_INPUT_CONTROL, input_text)
                time.sleep(0.5)

                # 点击生成
                page.click(SELECTOR_GENERATE_BUTTON)
                self.log_message.emit("   已点击生成按钮，等待语音合成...")

                # 等待网络和下载按钮
                page.wait_for_load_state("networkidle")
                page.wait_for_selector(SELECTOR_DOWNLOAD_BUTTON, state="visible", timeout=60000)
                self.log_message.emit("   语音合成完成")
                time.sleep(2)

                # 下载
                with page.expect_download(timeout=30000) as download_info:
                    page.click(SELECTOR_DOWNLOAD_BUTTON)
                download = download_info.value
                downloaded_path = download.path()

                if not downloaded_path or not os.path.exists(downloaded_path):
                    raise FileNotFoundError(f"下载文件不存在: {downloaded_path}")

                # 重命名并移动
                safe_filename = clean_filename(output_filename)
                if not safe_filename.endswith(".wav"):
                    safe_filename += ".wav"

                if is_special:
                    final_path = output_dir / output_filename / safe_filename
                else:
                    final_path = output_dir / safe_filename

                if final_path.exists():
                    final_path.unlink()

                shutil.move(str(downloaded_path), str(final_path))
                self.log_message.emit(f"   已保存: {final_path}")

                self.item_success.emit(i, output_filename)
                success_count += 1
                time.sleep(1)

            except Exception as e:
                self.item_error.emit(i, output_filename, str(e))
                self.log_error.emit(
                    f"[{i+1}/{len(input_texts)}] 失败: {output_filename} -- {e}"
                )
                try:
                    err_shot = f"error_{i+1}_{int(time.time())}.png"
                    page.screenshot(path=err_shot)
                    self.log_message.emit(f"   已保存错误截图: {err_shot}")
                except Exception:
                    pass
                fail_count += 1
                continue

        self.log_message.emit(f"\n=== 批量处理完成 ===")
        return success_count, fail_count, skip_count

    def _select_dropdown_option(self, page, selector: str, option_text: str):
        """选择下拉菜单选项。"""
        if not option_text:
            self.log_message.emit(f"  跳过空选项: {selector}")
            return
        self.log_message.emit(f"  选择下拉选项: {option_text}")
        try:
            page.click(selector)
            time.sleep(0.5)
            option_selector = f"text={option_text}"
            page.wait_for_selector(option_selector, state="visible", timeout=10000)
            page.click(option_selector)
            time.sleep(0.5)
            self.log_message.emit(f"  ✓ 成功选择: {option_text}")
        except Exception as e:
            self.log_error.emit(f"  选择下拉选项失败: {e}")
            raise

    def _upload_audio_files(self, page, selector: str, file_paths: list):
        """上传音频文件。"""
        self.log_message.emit(f"  上传音频文件: {len(file_paths)} 个")
        try:
            with page.expect_file_chooser() as fc_info:
                page.click(selector)
            file_chooser = fc_info.value
            file_chooser.set_files(file_paths)
            time.sleep(1)
            self.log_message.emit(f"  ✓ 成功上传 {len(file_paths)} 个文件")
        except Exception as e:
            self.log_error.emit(f"  上传文件失败: {e}")
            raise


# ==================== 配置面板 ====================

class ConfigWidget(QWidget):
    """配置页面——所有用户可编辑的输入字段。"""

    config_changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._settings = QSettings(APP_ORG, APP_NAME)
        self._init_ui()
        self._load_settings()

    def _init_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)

        container = QWidget()
        layout = QVBoxLayout(container)

        # ---- 服务器连接 ----
        grp_server = QGroupBox("服务器连接")
        form_server = QFormLayout(grp_server)
        self._port_input = QLineEdit()
        self._port_input.setPlaceholderText("9872")
        self._port_input.setText("9872")
        form_server.addRow("端口号:", self._port_input)
        layout.addWidget(grp_server)

        # ---- 输出设置 ----
        grp_output = QGroupBox("输出设置")
        form_output = QFormLayout(grp_output)
        self._output_dir_input = QLineEdit()
        self._output_dir_input.setPlaceholderText("例如: D:\\output")
        btn_browse_output = QPushButton("浏览...")
        btn_browse_output.clicked.connect(self._browse_output_dir)

        h_out = QHBoxLayout()
        h_out.addWidget(self._output_dir_input)
        h_out.addWidget(btn_browse_output)
        form_output.addRow("输出目录:", h_out)

        self._char_name_input = QLineEdit()
        self._char_name_input.setPlaceholderText("例如: Elysia（爱莉希雅）")
        form_output.addRow("角色名字:", self._char_name_input)

        self._author_input = QLineEdit()
        self._author_input.setPlaceholderText("例如: 燕戏竹林")
        form_output.addRow("作者信息:", self._author_input)
        layout.addWidget(grp_output)

        # ---- Excel 数据源 ----
        grp_excel = QGroupBox("Excel 数据源")
        form_excel = QFormLayout(grp_excel)
        self._excel_path_input = QLineEdit()
        self._excel_path_input.setPlaceholderText("选择 .xlsx 文件...")
        btn_browse_excel = QPushButton("浏览...")
        btn_browse_excel.clicked.connect(self._browse_excel)

        h_excel = QHBoxLayout()
        h_excel.addWidget(self._excel_path_input)
        h_excel.addWidget(btn_browse_excel)
        form_excel.addRow("Excel 路径:", h_excel)

        self._sheet_name_input = QLineEdit()
        self._sheet_name_input.setPlaceholderText("pacenote_view_202412300958")
        form_excel.addRow("工作表名:", self._sheet_name_input)

        self._btn_load_excel = QPushButton("加载 Excel 并预览数据")
        self._btn_load_excel.clicked.connect(self._on_load_excel)
        form_excel.addRow("", self._btn_load_excel)
        layout.addWidget(grp_excel)

        # ---- 模型配置 ----
        grp_model = QGroupBox("模型配置")
        form_model = QFormLayout(grp_model)
        self._dropdown_1_input = QLineEdit()
        self._dropdown_1_input.setPlaceholderText("GPT_weights_v2ProPlus/...")
        form_model.addRow("GPT 模型:", self._dropdown_1_input)

        self._dropdown_2_input = QLineEdit()
        self._dropdown_2_input.setPlaceholderText("SoVITS_weights_v2ProPlus/...")
        form_model.addRow("SoVITS 模型:", self._dropdown_2_input)

        self._control_value_input = QLineEdit()
        self._control_value_input.setPlaceholderText("37")
        form_model.addRow("top_k 参数:", self._control_value_input)

        self._initial_text_input = QLineEdit()
        self._initial_text_input.setPlaceholderText("参考音频文本内容")
        form_model.addRow("参考文本:", self._initial_text_input)
        layout.addWidget(grp_model)

        # ---- 参考音频 ----
        grp_audio = QGroupBox("参考音频文件")
        v_audio = QVBoxLayout(grp_audio)
        note_label = QLabel("第一个音频文件不要超过 10 秒")
        note_label.setStyleSheet("color: #888; font-size: 11px;")
        v_audio.addWidget(note_label)

        self._audio_list = QListWidget()
        self._audio_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        v_audio.addWidget(self._audio_list)

        h_audio_btn = QHBoxLayout()
        btn_add_audio = QPushButton("添加文件...")
        btn_add_audio.clicked.connect(self._add_audio_files)
        btn_remove_audio = QPushButton("移除选中")
        btn_remove_audio.clicked.connect(self._remove_audio_files)
        h_audio_btn.addWidget(btn_add_audio)
        h_audio_btn.addWidget(btn_remove_audio)
        h_audio_btn.addStretch()
        v_audio.addLayout(h_audio_btn)
        layout.addWidget(grp_audio)

        # 弹簧
        layout.addStretch()

        # 容器设为滚动区域内容
        scroll.setWidget(container)

        # 最外层布局
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll)

    # ---- 槽函数 ----

    def _browse_output_dir(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if d:
            self._output_dir_input.setText(d)

    def _browse_excel(self):
        f, _ = QFileDialog.getOpenFileName(
            self, "选择 Excel 文件", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        if f:
            self._excel_path_input.setText(f)

    def _add_audio_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择音频文件", "", "Audio Files (*.wav *.mp3 *.flac *.ogg);;All Files (*)"
        )
        for f in files:
            self._audio_list.addItem(QListWidgetItem(f))

    def _remove_audio_files(self):
        for item in self._audio_list.selectedItems():
            self._audio_list.takeItem(self._audio_list.row(item))

    def _on_load_excel(self):
        """加载 Excel 并通知 MainWindow 更新数据预览。"""
        path = self._excel_path_input.text().strip()
        sheet = self._sheet_name_input.text().strip()
        if not path or not os.path.isfile(path):
            QMessageBox.warning(self, "错误", "Excel 文件路径无效")
            return
        if not sheet:
            QMessageBox.warning(self, "错误", "请输入工作表名称")
            return
        self.config_changed.emit()

    # ---- 获取/保存配置 ----

    def get_config(self) -> dict:
        """从所有输入字段收集配置。"""
        audio_files = []
        for i in range(self._audio_list.count()):
            audio_files.append(self._audio_list.item(i).text())
        return {
            "port": self._port_input.text().strip(),
            "output_directory": self._output_dir_input.text().strip(),
            "character_name": self._char_name_input.text().strip(),
            "author": self._author_input.text().strip(),
            "excel_path": self._excel_path_input.text().strip(),
            "sheet_name": self._sheet_name_input.text().strip(),
            "dropdown_1": self._dropdown_1_input.text().strip(),
            "dropdown_2": self._dropdown_2_input.text().strip(),
            "control_value": self._control_value_input.text().strip(),
            "initial_text": self._initial_text_input.text().strip(),
            "audio_files": audio_files,
        }

    def validate(self) -> str:
        """验证配置是否完整。返回空字符串表示通过，否则返回错误信息。"""
        cfg = self.get_config()
        if not cfg["port"]:
            return "端口号不能为空"
        if not cfg["output_directory"]:
            return "输出目录不能为空"
        if not cfg["character_name"]:
            return "角色名字不能为空"
        if not cfg["author"]:
            return "作者信息不能为空"
        if not cfg["excel_path"] or not os.path.isfile(cfg["excel_path"]):
            return "Excel 文件路径无效或不存在"
        if not cfg["sheet_name"]:
            return "工作表名称不能为空"
        if not cfg["audio_files"]:
            return "请至少添加一个参考音频文件"
        if not cfg["dropdown_1"]:
            return "请填写 GPT 模型名称"
        if not cfg["dropdown_2"]:
            return "请填写 SoVITS 模型名称"
        return ""

    def load_excel_data(self):
        """由 MainWindow 调用以加载 Excel 数据并返回。"""
        path = self._excel_path_input.text().strip()
        sheet = self._sheet_name_input.text().strip()
        if path and os.path.isfile(path) and sheet:
            return read_excel_data(path, sheet)
        return [], []

    def _load_settings(self):
        """从 QSettings 恢复字段值。"""
        s = self._settings
        self._port_input.setText(s.value("config/port", "9872"))
        self._output_dir_input.setText(s.value("config/output_dir", ""))
        self._char_name_input.setText(s.value("config/character_name", ""))
        self._author_input.setText(s.value("config/author", ""))
        self._excel_path_input.setText(s.value("config/excel_path", ""))
        self._sheet_name_input.setText(s.value("config/sheet_name", "pacenote_view_202412300958"))
        self._dropdown_1_input.setText(s.value("config/dropdown_1", ""))
        self._dropdown_2_input.setText(s.value("config/dropdown_2", ""))
        self._control_value_input.setText(s.value("config/control_value", ""))
        self._initial_text_input.setText(s.value("config/initial_text", ""))

        # 恢复音频文件列表
        audio_json = s.value("config/audio_files", "[]")
        try:
            audio_list = json.loads(audio_json)
            for f in audio_list:
                if f:
                    self._audio_list.addItem(QListWidgetItem(f))
        except (json.JSONDecodeError, TypeError):
            pass

    def save_settings(self):
        """将所有输入字段保存到 QSettings。"""
        s = self._settings
        cfg = self.get_config()
        s.setValue("config/port", cfg["port"])
        s.setValue("config/output_dir", cfg["output_directory"])
        s.setValue("config/character_name", cfg["character_name"])
        s.setValue("config/author", cfg["author"])
        s.setValue("config/excel_path", cfg["excel_path"])
        s.setValue("config/sheet_name", cfg["sheet_name"])
        s.setValue("config/dropdown_1", cfg["dropdown_1"])
        s.setValue("config/dropdown_2", cfg["dropdown_2"])
        s.setValue("config/control_value", cfg["control_value"])
        s.setValue("config/initial_text", cfg["initial_text"])
        s.setValue("config/audio_files", json.dumps(cfg["audio_files"]))

    def set_readonly(self, readonly: bool):
        """启用/禁用所有输入字段。"""
        self._port_input.setReadOnly(readonly)
        self._output_dir_input.setReadOnly(readonly)
        self._char_name_input.setReadOnly(readonly)
        self._author_input.setReadOnly(readonly)
        self._excel_path_input.setReadOnly(readonly)
        self._sheet_name_input.setReadOnly(readonly)
        self._dropdown_1_input.setReadOnly(readonly)
        self._dropdown_2_input.setReadOnly(readonly)
        self._control_value_input.setReadOnly(readonly)
        self._initial_text_input.setReadOnly(readonly)
        self._btn_load_excel.setEnabled(not readonly)


# ==================== 数据预览 ====================

class DataPreviewWidget(QWidget):
    """Excel 数据预览表格。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)

        self._status_label = QLabel("未加载数据")
        layout.addWidget(self._status_label)

        self._table = QTableWidget()
        self._table.setColumnCount(3)
        self._table.setHorizontalHeaderLabels(["#", "文件名", "输入文本"])
        self._table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._table.setAlternatingRowColors(True)
        self._table.setSortingEnabled(True)
        self._table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self._table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Interactive)
        self._table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self._table.horizontalHeader().resizeSection(1, 200)
        layout.addWidget(self._table)

    def set_data(self, file_names: list, input_texts: list):
        """用 Excel 数据填充表格。"""
        n = len(file_names)
        self._table.setRowCount(n)
        for i, (fname, text) in enumerate(zip(file_names, input_texts)):
            self._table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
            item_name = QTableWidgetItem(fname)
            item_name.setToolTip(fname)
            self._table.setItem(i, 1, item_name)
            item_text = QTableWidgetItem(text)
            item_text.setToolTip(text)
            self._table.setItem(i, 2, item_text)

        self._status_label.setText(f"已加载 {n} 条数据")


# ==================== 运行控制面板 ====================

class RunControlWidget(QWidget):
    """运行控制页：进度、日志、控制按钮。"""

    # 按钮信号
    start_clicked = Signal()
    pause_clicked = Signal()
    resume_clicked = Signal()
    stop_clicked = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._is_paused = False
        self._success_count = 0
        self._fail_count = 0
        self._skip_count = 0
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)

        # ---- 进度条区域 ----
        progress_group = QGroupBox("处理进度")
        pg_layout = QVBoxLayout(progress_group)

        # 总进度
        progress_row1 = QHBoxLayout()
        progress_row1.addWidget(QLabel("总体进度:"))
        self._progress_bar = QProgressBar()
        self._progress_bar.setRange(0, 100)
        self._progress_bar.setValue(0)
        progress_row1.addWidget(self._progress_bar)
        pg_layout.addLayout(progress_row1)

        # 统计
        stats_row = QHBoxLayout()
        self._lbl_progress = QLabel("进度: 0 / 0")
        self._lbl_stats = QLabel("成功: 0 | 失败: 0 | 跳过: 0")
        stats_row.addWidget(self._lbl_progress)
        stats_row.addStretch()
        stats_row.addWidget(self._lbl_stats)
        pg_layout.addLayout(stats_row)

        # 设置进度
        setup_row = QHBoxLayout()
        setup_row.addWidget(QLabel("设置步骤:"))
        self._setup_bar = QProgressBar()
        self._setup_bar.setRange(0, SETUP_STEP_COUNT)
        self._setup_bar.setValue(0)
        self._setup_bar.setVisible(False)
        setup_row.addWidget(self._setup_bar)

        self._lbl_setup = QLabel("")
        setup_row.addWidget(self._lbl_setup)
        pg_layout.addLayout(setup_row)

        layout.addWidget(progress_group)

        # ---- 日志区域 ----
        log_label = QLabel("运行日志:")
        layout.addWidget(log_label)

        self._log_output = QTextEdit()
        self._log_output.setReadOnly(True)
        monospace = QFont("Consolas", 10)
        self._log_output.setFont(monospace)
        self._log_output.document().setMaximumBlockCount(10000)
        self._log_output.setLineWrapMode(QTextEdit.NoWrap)
        layout.addWidget(self._log_output)

        # ---- 控制按钮 ----
        btn_layout = QHBoxLayout()
        self._btn_start = QPushButton("开始处理")
        self._btn_start.setMinimumHeight(36)
        self._btn_start.setStyleSheet(
            "QPushButton { background-color: #27ae60; color: white; "
            "font-size: 14px; padding: 4px 20px; }"
        )
        self._btn_start.clicked.connect(self.start_clicked.emit)

        self._btn_pause = QPushButton("暂停")
        self._btn_pause.setMinimumHeight(36)
        self._btn_pause.setEnabled(False)
        self._btn_pause.clicked.connect(self._on_pause_clicked)

        self._btn_stop = QPushButton("停止")
        self._btn_stop.setMinimumHeight(36)
        self._btn_stop.setEnabled(False)
        self._btn_stop.clicked.connect(self.stop_clicked.emit)

        btn_clear = QPushButton("清除日志")
        btn_clear.clicked.connect(self._clear_log)

        btn_save_log = QPushButton("保存日志...")
        btn_save_log.clicked.connect(self._save_log)

        btn_layout.addWidget(self._btn_start)
        btn_layout.addWidget(self._btn_pause)
        btn_layout.addWidget(self._btn_stop)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_clear)
        btn_layout.addWidget(btn_save_log)
        layout.addLayout(btn_layout)

    # ---- 日志追加 ----

    def append_log(self, text: str):
        """附加普通日志（黑色）。"""
        self._log_output.append(text)
        self._auto_scroll()

    def append_error_log(self, text: str):
        """附加错误日志（红色）。"""
        self._log_output.setTextColor(QColor("#e74c3c"))
        self._log_output.append(text)
        self._log_output.setTextColor(QColor("#000000"))
        self._auto_scroll()

    def _auto_scroll(self):
        cursor = self._log_output.textCursor()
        cursor.movePosition(QTextCursor.End)
        self._log_output.setTextCursor(cursor)

    def _clear_log(self):
        self._log_output.clear()

    def _save_log(self):
        f, _ = QFileDialog.getSaveFileName(
            self, "保存日志", f"log_{datetime.now():%Y%m%d_%H%M%S}.txt",
            "Text Files (*.txt);;All Files (*)"
        )
        if f:
            with open(f, "w", encoding="utf-8") as fp:
                fp.write(self._log_output.toPlainText())
            self.append_log(f"日志已保存: {f}")

    # ---- 按钮状态管理 ----

    def _on_pause_clicked(self):
        if self._is_paused:
            self._is_paused = False
            self._btn_pause.setText("暂停")
            self.resume_clicked.emit()
        else:
            self._is_paused = True
            self._btn_pause.setText("继续")
            self.pause_clicked.emit()

    def set_state_idle(self):
        """切换到空闲状态。"""
        self._btn_start.setEnabled(True)
        self._btn_pause.setEnabled(False)
        self._btn_pause.setText("暂停")
        self._btn_stop.setEnabled(False)
        self._is_paused = False

    def set_state_running(self):
        """切换到运行中状态。"""
        self._btn_start.setEnabled(False)
        self._btn_pause.setEnabled(True)
        self._btn_pause.setText("暂停")
        self._btn_stop.setEnabled(True)
        self._is_paused = False

    def set_state_finished(self):
        """切换到完成状态。"""
        self.set_state_idle()

    # ---- 进度更新 ----

    def set_batch_total(self, total: int):
        self._progress_bar.setRange(0, total)
        self._progress_bar.setValue(0)
        self._lbl_progress.setText(f"进度: 0 / {total}")
        self._success_count = 0
        self._fail_count = 0
        self._skip_count = 0
        self._update_stats()

    def update_batch_progress(self, index: int, _filename: str):
        self._progress_bar.setValue(index + 1)
        total = self._progress_bar.maximum()
        self._lbl_progress.setText(f"进度: {index + 1} / {total}")

    def update_setup_progress(self, step: int, desc: str):
        self._setup_bar.setVisible(True)
        self._setup_bar.setValue(step)
        self._lbl_setup.setText(f"[{step}/{SETUP_STEP_COUNT}] {desc}")

    def hide_setup_progress(self):
        self._setup_bar.setVisible(False)
        self._lbl_setup.setText("")

    def on_item_success(self, _index: int, _filename: str):
        self._success_count += 1
        self._update_stats()

    def on_item_error(self, _index: int, _filename: str, _error: str):
        self._fail_count += 1
        self._update_stats()

    def on_item_skipped(self, _index: int, _filename: str):
        self._skip_count += 1
        self._update_stats()

    def show_summary(self, success_msg: str = "处理完成"):
        """在日志中显示汇总信息。"""
        total = self._success_count + self._fail_count + self._skip_count
        summary = (
            f"\n{'='*50}\n"
            f"{success_msg}\n"
            f"总计: {total} | 成功: {self._success_count} | "
            f"失败: {self._fail_count} | 跳过: {self._skip_count}\n"
            f"{'='*50}"
        )
        self.append_log(summary)

    def _update_stats(self):
        self._lbl_stats.setText(
            f"成功: {self._success_count} | 失败: {self._fail_count} | "
            f"跳过: {self._skip_count}"
        )


# ==================== 输出浏览面板 ====================

class OutputBrowserWidget(QWidget):
    """浏览生成好的输出文件。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._output_dir = None
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)

        self._lbl_path = QLabel("输出目录: (尚未生成)")
        layout.addWidget(self._lbl_path)

        # 模型
        self._model = QStandardItemModel()
        self._model.setHorizontalHeaderLabels(["名称", "大小", "类型"])
        self._tree = QTreeView()
        self._tree.setModel(self._model)
        self._tree.setAlternatingRowColors(True)
        self._tree.setSortingEnabled(True)
        layout.addWidget(self._tree)

        # 按钮
        btn_layout = QHBoxLayout()
        btn_refresh = QPushButton("刷新")
        btn_refresh.clicked.connect(self.refresh)
        btn_open = QPushButton("在资源管理器中打开")
        btn_open.clicked.connect(self._open_in_explorer)
        btn_layout.addWidget(btn_refresh)
        btn_layout.addWidget(btn_open)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

    def set_output_dir(self, path: str):
        """设置输出目录路径。"""
        self._output_dir = path
        self._lbl_path.setText(f"输出目录: {path}")

    def refresh(self):
        """刷新文件树。"""
        self._model.clear()
        self._model.setHorizontalHeaderLabels(["名称", "大小", "类型"])
        if not self._output_dir or not os.path.isdir(self._output_dir):
            return
        self._populate_tree(self._model.invisibleRootItem(), self._output_dir)
        self._tree.expandAll()

    def _populate_tree(self, parent_item, dir_path):
        """递归填充文件树。"""
        try:
            for entry in sorted(os.listdir(dir_path)):
                full_path = os.path.join(dir_path, entry)
                is_dir = os.path.isdir(full_path)
                name_item = QStandardItem(entry)
                if is_dir:
                    name_item.setIcon(self.style().standardIcon(48))  # SP_DirIcon
                    size_item = QStandardItem("")
                    type_item = QStandardItem("文件夹")
                    parent_item.appendRow([name_item, size_item, type_item])
                    self._populate_tree(name_item, full_path)
                else:
                    try:
                        size = os.path.getsize(full_path)
                        size_str = f"{size / 1024:.1f} KB" if size < 1024 * 1024 else f"{size / (1024*1024):.1f} MB"
                    except Exception:
                        size_str = ""
                    ext = os.path.splitext(entry)[1].upper() or "文件"
                    size_item = QStandardItem(size_str)
                    type_item = QStandardItem(ext)
                    parent_item.appendRow([name_item, size_item, type_item])
        except PermissionError:
            pass

    def _open_in_explorer(self):
        if self._output_dir and os.path.isdir(self._output_dir):
            os.startfile(self._output_dir)


# ==================== 主窗口 ====================

class MainWindow(QMainWindow):
    """应用程序主窗口。"""

    def __init__(self):
        super().__init__()
        self._settings = QSettings(APP_ORG, APP_NAME)
        self._thread = None
        self._worker = None
        self._init_ui()
        self._check_legal()
        self._restore_geometry()

    def _init_ui(self):
        self.setWindowTitle("Auto Media Pack ACG - 语音包生成器")
        self.setMinimumSize(1000, 700)

        # ---- 菜单栏 ----
        menubar = self.menuBar()

        file_menu = menubar.addMenu("文件")
        save_act = QAction("保存设置", self)
        save_act.triggered.connect(self._on_save_settings)
        file_menu.addAction(save_act)
        file_menu.addSeparator()
        exit_act = QAction("退出", self)
        exit_act.setShortcut("Ctrl+Q")
        exit_act.triggered.connect(self.close)
        file_menu.addAction(exit_act)

        help_menu = menubar.addMenu("帮助")
        legal_act = QAction("查看法律声明", self)
        legal_act.triggered.connect(self._show_legal_dialog)
        help_menu.addAction(legal_act)
        help_menu.addSeparator()
        about_act = QAction("关于", self)
        about_act.triggered.connect(self._show_about)
        help_menu.addAction(about_act)

        # ---- 选项卡 ----
        self._tabs = QTabWidget()

        self._config_widget = ConfigWidget()
        self._config_widget.config_changed.connect(self._on_config_changed)
        self._tabs.addTab(self._config_widget, "配置")

        self._preview_widget = DataPreviewWidget()
        self._tabs.addTab(self._preview_widget, "数据预览")

        self._run_widget = RunControlWidget()
        self._run_widget.start_clicked.connect(self._on_start)
        self._run_widget.pause_clicked.connect(self._on_pause)
        self._run_widget.resume_clicked.connect(self._on_resume)
        self._run_widget.stop_clicked.connect(self._on_stop)
        self._tabs.addTab(self._run_widget, "运行")

        self._output_widget = OutputBrowserWidget()
        self._tabs.addTab(self._output_widget, "输出")

        self.setCentralWidget(self._tabs)

        # ---- 状态栏 ----
        self._status_bar = QStatusBar()
        self.setStatusBar(self._status_bar)
        self._lbl_worker_status = QLabel("状态: 空闲")
        self._lbl_port_display = QLabel("端口: -")
        self._status_bar.addPermanentWidget(self._lbl_worker_status)
        self._status_bar.addPermanentWidget(self._lbl_port_display)

    # ---- 法律声明检查 ----

    def _check_legal(self):
        """检查用户是否已接受法律声明，若未接受则显示对话框。"""
        accepted = self._settings.value(SETTINGS_LEGAL_ACCEPTED, False, type=bool)
        if not accepted:
            dlg = LegalDisclaimerDialog(self)
            if dlg.exec() == QDialog.Accepted:
                self._settings.setValue(SETTINGS_LEGAL_ACCEPTED, True)
                self._settings.sync()

    def _show_legal_dialog(self):
        dlg = LegalDisclaimerDialog(self)
        dlg.exec()

    def _show_about(self):
        QMessageBox.about(
            self, "关于",
            "Auto Media Pack ACG - 语音包生成器\n\n"
            "基于 Playwright 的 GPT-SoVITS TTS 批量自动化工具\n\n"
            "许可证: GPL-3.0-or-later\n"
            "依赖: Playwright (Apache-2.0), openpyxl (MIT)\n"
            "GPT-SoVITS (MIT): https://github.com/RVC-Boss/GPT-SoVITS"
        )

    # ---- 配置槽函数 ----

    def _on_config_changed(self):
        """配置面板中的「加载 Excel」按钮被点击。"""
        file_names, input_texts = self._config_widget.load_excel_data()
        if file_names:
            self._preview_widget.set_data(file_names, input_texts)
            self._tabs.setCurrentIndex(1)  # 切换到数据预览
        else:
            QMessageBox.warning(self, "加载失败", "无法加载 Excel 数据，请检查路径和工作表名。")

    def _on_save_settings(self):
        self._config_widget.save_settings()
        self._status_bar.showMessage("配置已保存", 3000)

    # ---- 运行控制 ----

    def _on_start(self):
        """开始处理。"""
        error = self._config_widget.validate()
        if error:
            QMessageBox.warning(self, "配置不完整", error)
            return

        # 保存当前配置
        self._config_widget.save_settings()
        config = self._config_widget.get_config()

        # 更新状态栏
        self._lbl_worker_status.setText("状态: 运行中")
        self._lbl_port_display.setText(f"端口: {config['port']}")

        # 重置 UI
        self._run_widget.set_state_running()
        self._run_widget.hide_setup_progress()
        self._tabs.setCurrentIndex(2)  # 切换到运行页
        self._config_widget.set_readonly(True)

        # 创建线程和工作器
        self._thread = QThread(self)
        self._worker = PlaywrightWorker(config)
        self._worker.moveToThread(self._thread)

        # 连接信号
        self._worker.log_message.connect(self._run_widget.append_log)
        self._worker.log_error.connect(self._run_widget.append_error_log)
        self._worker.setup_progress.connect(self._run_widget.update_setup_progress)
        self._worker.batch_total.connect(self._run_widget.set_batch_total)
        self._worker.batch_progress.connect(self._run_widget.update_batch_progress)
        self._worker.item_success.connect(self._run_widget.on_item_success)
        self._worker.item_error.connect(self._run_widget.on_item_error)
        self._worker.item_skipped.connect(self._run_widget.on_item_skipped)
        self._worker.browser_opened.connect(self._on_browser_opened)
        self._worker.browser_closed.connect(self._on_browser_closed)
        self._worker.error_occurred.connect(self._on_worker_error)
        self._worker.verification_done.connect(self._on_verification_done)
        self._worker.finished.connect(self._on_worker_finished)

        # 线程生命周期
        self._thread.started.connect(self._worker.do_work)
        self._worker.finished.connect(self._thread.quit)
        self._thread.finished.connect(self._worker.deleteLater)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.finished.connect(self._on_thread_finished)

        self._thread.start()
        self._run_widget.append_log("=== 开始处理 ===")

    def _on_pause(self):
        """暂停处理。"""
        if self._worker is None:
            return
        self._worker.pause()
        self._run_widget.append_log("=== 已暂停 ===")

    def _on_resume(self):
        """继续处理。"""
        if self._worker is None:
            return
        self._worker.resume()
        self._run_widget.append_log("=== 继续处理 ===")

    def _on_stop(self):
        """停止处理。"""
        if self._worker is None:
            return
        self._run_widget.append_log("=== 正在停止... ===")
        self._worker.stop()

    # ---- 工作器信号槽 ----

    def _on_browser_opened(self):
        self._status_bar.showMessage("浏览器已打开", 3000)

    def _on_browser_closed(self):
        self._status_bar.showMessage("浏览器已关闭", 3000)

    def _on_worker_error(self, error_msg: str):
        self._run_widget.append_error_log(f"严重错误: {error_msg}")

    def _on_verification_done(self, passed: bool, expected: int, actual: int):
        if passed:
            self._run_widget.append_log(f"✓ 验证通过：文件数量正确 ({expected})")
        else:
            self._run_widget.append_error_log(f"✗ 验证失败：期望 {expected} 个文件，实际 {actual} 个")

    def _on_worker_finished(self, success: bool, out_dir: str, ok: int, fail: int, skip: int):
        if success:
            self._run_widget.show_summary("处理完成")
        else:
            self._run_widget.show_summary("处理已取消/中断")

        # 更新输出浏览器
        if out_dir:
            self._output_widget.set_output_dir(out_dir)
            self._output_widget.refresh()

    def _on_thread_finished(self):
        """线程退出后清理并恢复 UI。"""
        self._thread = None
        self._worker = None
        self._run_widget.set_state_idle()
        self._config_widget.set_readonly(False)
        self._lbl_worker_status.setText("状态: 空闲")
        self._status_bar.showMessage("处理结束", 5000)
        # 切换到输出页
        if self._output_widget._output_dir:
            self._tabs.setCurrentIndex(3)

    # ---- 窗口生命周期 ----

    def _restore_geometry(self):
        geo = self._settings.value(SETTINGS_WINDOW_GEOMETRY)
        if geo:
            self.restoreGeometry(geo)
        state = self._settings.value(SETTINGS_WINDOW_STATE)
        if state:
            self.restoreState(state)

    def closeEvent(self, event):
        """关闭窗口时保存设置并确保线程退出。"""
        if self._thread and self._thread.isRunning():
            reply = QMessageBox.question(
                self, "处理进行中",
                "当前正在处理批量任务。\n\n"
                "停止处理并关闭程序？\n"
                "(已下载的文件将被保留。)",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply == QMessageBox.No:
                event.ignore()
                return
            # 停止工作器
            if self._worker:
                self._worker.stop()
            self._thread.quit()
            if not self._thread.wait(10000):
                self._thread.terminate()
                self._thread.wait(3000)

        # 保存设置
        self._config_widget.save_settings()
        self._settings.setValue(SETTINGS_WINDOW_GEOMETRY, self.saveGeometry())
        self._settings.setValue(SETTINGS_WINDOW_STATE, self.saveState())
        self._settings.sync()

        event.accept()


# ==================== 入口 ====================

def main():
    app = QApplication(sys.argv)
    app.setOrganizationName(APP_ORG)
    app.setApplicationName(APP_NAME)

    # Fusion 风格：跨平台一致
    app.setStyle("Fusion")

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
