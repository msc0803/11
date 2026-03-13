"""
视频素材批量查询 & 加入工作台 - PyQt6 GUI
"""
import sys
import os
import time
import math
import pandas as pd
from datetime import datetime

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QComboBox, QTableWidget,
    QTableWidgetItem, QTextEdit, QProgressBar, QGroupBox, QSpinBox,
    QSplitter, QHeaderView, QMessageBox, QStatusBar, QLineEdit, QCheckBox, QMenu
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSortFilterProxyModel
from PyQt6.QtGui import QColor, QFont

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import api_core

RESULTS_DIR = os.path.join(os.path.dirname(__file__), "results")
os.makedirs(RESULTS_DIR, exist_ok=True)


# ─────────────────────────── 后台线程 ────────────────────────────

class TokenWorker(QThread):
    success = pyqtSignal(str)
    failed  = pyqtSignal(str)
    log     = pyqtSignal(str)

    def __init__(self, ask_login=False):
        super().__init__()
        self.ask_login = ask_login

    def run(self):
        try:
            if not self.ask_login:
                # 先试缓存
                token = api_core.load_cached_token()
                if token and api_core.is_token_valid(token):
                    self.success.emit(token)
                    return
            # 启动浏览器
            self.log.emit("正在启动浏览器...")
            api_core._ensure_chrome_debug()
            self.log.emit("浏览器已就绪，请登录 sucaiwang.zhishangsoft.com 后点「已登录」")
            token = api_core.get_token_from_browser()
            api_core.save_token(token)
            self.success.emit(token)
        except Exception as e:
            self.failed.emit(str(e))
    progress     = pyqtSignal(int, int)
    row_done     = pyqtSignal(int, str, str, str, str, str)  # idx,name,status,ids,names,remark
    log          = pyqtSignal(str)
    finished_signal = pyqtSignal(int, int, int, int)
    token_refreshed = pyqtSignal(str)

    def __init__(self, df, name_col, start_row, token):
        super().__init__()
        self.df        = df
        self.name_col  = name_col
        self.start_row = start_row
        self.token     = token
        self._stop     = False

    def stop(self): self._stop = True

    def run(self):
        found = not_found = error = skip = 0
        total = len(self.df)
        done  = 0

        for idx, row in self.df.iterrows():
            if self._stop:
                self.log.emit("用户停止查询")
                break

            # 跳过逻辑
            if idx < self.start_row:
                skip += 1; done += 1
                self.progress.emit(done, total); continue

            name = row[self.name_col]
            if pd.isna(name) or str(name).strip() == "":
                skip += 1; done += 1
                self.progress.emit(done, total); continue
            name = str(name).strip()

            if str(row.get("查询状态", "")).strip() == "找到":
                skip += 1; done += 1
                self.progress.emit(done, total); continue

            # 查询 + 限流重试
            self.log.emit(f"查询 [{done+1}/{total}]: {name}")
            ok, vid, extra = api_core.search_video(self.token, name)
            for _ in range(3):
                if extra.get("error", "").startswith("当前访问人数"):
                    time.sleep(3)
                    ok, vid, extra = api_core.search_video(self.token, name)
                else:
                    break

            # token 过期
            if extra.get("token_expired"):
                self.log.emit("Token 失效，自动刷新...")
                try:
                    self.token = api_core.get_token_from_browser()
                    api_core.save_token(self.token)
                    self.token_refreshed.emit(self.token)
                    ok, vid, extra = api_core.search_video(self.token, name)
                except Exception as e:
                    self.log.emit(f"Token 刷新失败: {e}")

            if ok:
                cnt    = extra.get("match_count", 1)
                raws   = extra.get("all_raws", [])
                names  = "；".join(v.get("name", "") for v in raws)
                remark = f"匹配{cnt}条 | state={extra.get('videoState','')} cost={extra.get('sumStatCost',0)} roi={extra.get('sumRoi',0)}"
                self.log.emit(f"  ✅ 找到 {cnt} 个视频: {vid}")
                self.row_done.emit(idx, name, "找到", vid, names, remark)
                found += 1
            elif "error" in extra:
                self.log.emit(f"  ❌ 错误: {extra['error']}")
                self.row_done.emit(idx, name, "错误", "", "", extra["error"])
                error += 1
            else:
                self.log.emit(f"  — 未找到")
                self.row_done.emit(idx, name, "未找到", "", "", extra.get("info", ""))
                not_found += 1

            done += 1
            self.progress.emit(done, total)
            if done % 50 == 0:
                self.log.emit(f"进度 {done}/{total}  找到:{found} 未找到:{not_found} 错误:{error}")
            time.sleep(0.5)

        self.finished_signal.emit(found, not_found, error, skip)


class WorkbenchWorker(QThread):
    log        = pyqtSignal(str)
    batch_done = pyqtSignal(int, int, int)   # batch_idx, total_batches, count
    all_done   = pyqtSignal()

    def __init__(self, objects, batch_size=200):
        super().__init__()
        self.objects    = objects
        self.batch_size = batch_size
        self._stop      = False
        self._go        = False   # 等待继续信号

    def stop(self):
        self._stop = True
        self._go   = True

    def continue_next(self):
        self._go = True

    def run(self):
        total_b = math.ceil(len(self.objects) / self.batch_size)
        self.log.emit(f"共 {len(self.objects)} 个，分 {total_b} 批（每批 {self.batch_size}）")

        for bi in range(total_b):
            if self._stop: break
            batch = self.objects[bi * self.batch_size:(bi + 1) * self.batch_size]
            self.log.emit(f"\n--- 第 {bi+1}/{total_b} 批（{len(batch)} 个）---")
            try:
                cnt = api_core.set_workbench_via_cdp(batch)
                self.log.emit(f"写入工作台: {cnt} 个")
                self.batch_done.emit(bi + 1, total_b, len(batch))
            except Exception as e:
                self.log.emit(f"写入失败: {e}")
                continue

            if bi < total_b - 1:
                self._go = False
                while not self._go and not self._stop:
                    time.sleep(0.2)

        if not self._stop:
            self.all_done.emit()


# ─────────────────────────── 主窗口 ──────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("视频素材批量查询工具")
        self.setMinimumSize(1200, 800)

        self.df             = None
        self.df_filtered    = None
        self.excel_path     = ""
        self.token          = ""
        self.sheet_names    = []
        self.current_sheet  = None
        self.search_worker  = None
        self.wb_worker      = None
        self._wb_pushed     = 0
        self._wb_total_objs = 0
        self.results_dir    = RESULTS_DIR

        self._build_ui()
        self._connect_token()

    # ──────────── UI 构建 ────────────

    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        vbox = QVBoxLayout(root)
        vbox.setContentsMargins(8, 8, 8, 8)
        vbox.setSpacing(6)

        # === 文件区 ===
        fg = QGroupBox("Excel 文件")
        fl = QHBoxLayout(fg)
        self.file_label  = QLabel("未选择文件")
        self.file_label.setMinimumWidth(300)
        btn_open = QPushButton("打开 Excel")
        btn_open.setFixedWidth(90)
        btn_open.clicked.connect(self._open_file)
        self.sheet_combo = QComboBox(); self.sheet_combo.setMinimumWidth(110)
        self.sheet_combo.setPlaceholderText("工作表")
        self.sheet_combo.currentIndexChanged.connect(self._on_sheet_changed)
        self.col_combo   = QComboBox(); self.col_combo.setMinimumWidth(140)
        self.col_combo.setPlaceholderText("搜索列")
        self.start_spin  = QSpinBox()
        self.start_spin.setPrefix("起始行: "); self.start_spin.setMaximum(99999)
        self.chk_header = QCheckBox("首行为表头")
        self.chk_header.setChecked(True)
        self.chk_header.setToolTip("勾选：Excel 第一行是列名；取消：第一行是数据")
        fl.addWidget(btn_open)
        fl.addWidget(self.file_label, 1)
        fl.addWidget(QLabel("工作表:")); fl.addWidget(self.sheet_combo)
        fl.addWidget(QLabel("搜索列:")); fl.addWidget(self.col_combo)
        fl.addWidget(self.start_spin)
        fl.addWidget(self.chk_header)
        vbox.addWidget(fg)

        # === 结果目录 ===
        dir_row = QHBoxLayout()
        dir_row.addWidget(QLabel("结果保存目录:"))
        self.dir_label = QLabel(RESULTS_DIR)
        self.dir_label.setStyleSheet("color:#1565C0; font-size:12px;")
        self.dir_label.setMinimumWidth(300)
        btn_dir = QPushButton("更改目录")
        btn_dir.setFixedWidth(80)
        btn_dir.clicked.connect(self._choose_results_dir)
        dir_row.addWidget(self.dir_label, 1)
        dir_row.addWidget(btn_dir)
        vbox.addLayout(dir_row)

        # === 筛选栏 ===
        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("筛选状态:"))
        self.filter_combo = QComboBox()
        self.filter_combo.addItems(["全部", "找到", "未找到", "错误", "未查询"])
        self.filter_combo.currentIndexChanged.connect(self._apply_filter)
        self.filter_combo.setFixedWidth(100)
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("关键词搜索（素材名称）")
        self.search_box.textChanged.connect(self._apply_filter)
        self.search_box.setFixedWidth(220)
        self.row_count_label = QLabel("显示: 0 / 0 行")
        filter_row.addWidget(self.filter_combo)
        filter_row.addWidget(QLabel("  关键词:"))
        filter_row.addWidget(self.search_box)
        filter_row.addWidget(self.row_count_label)
        filter_row.addStretch()
        vbox.addLayout(filter_row)

        # === 分割：表格 + 日志 ===
        splitter = QSplitter(Qt.Orientation.Vertical)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectItems)
        self.table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        self.table.setSortingEnabled(True)
        self.table.keyPressEvent = self._table_key_press
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._table_context_menu)
        splitter.addWidget(self.table)

        log_g = QGroupBox("运行日志")
        log_l = QVBoxLayout(log_g)
        log_l.setContentsMargins(4, 4, 4, 4)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Menlo", 11))
        self.log_text.setMaximumHeight(180)
        log_l.addWidget(self.log_text)
        splitter.addWidget(log_g)

        splitter.setStretchFactor(0, 4)
        splitter.setStretchFactor(1, 1)
        vbox.addWidget(splitter, 1)

        # === 按钮行 ===
        btn_row = QHBoxLayout()
        def _btn(text, color, hover, slot):
            b = QPushButton(text)
            b.setFixedHeight(34)
            b.setStyleSheet(
                f"QPushButton{{background:{color};color:white;font-size:13px;"
                f"border-radius:4px;padding:0 14px}}"
                f"QPushButton:hover{{background:{hover}}}"
                f"QPushButton:disabled{{background:#bbb}}"
            )
            b.clicked.connect(slot)
            return b

        self.btn_search    = _btn("开始查询",    "#4CAF50", "#388E3C", self._start_search)
        self.btn_stop      = _btn("停止",        "#f44336", "#c62828", self._stop_search)
        self.btn_stop.setEnabled(False)
        self.btn_workbench = _btn("加入工作台",  "#2196F3", "#1565C0", self._add_to_workbench)
        self.btn_save      = _btn("保存 Excel",  "#FF9800", "#E65100", self._save_excel)
        self.btn_token     = _btn("刷新 Token",  "#607D8B", "#37474F", self._refresh_token)
        self.btn_requery_errors = _btn("重查错误行", "#9C27B0", "#6A1B9A", self._requery_errors)
        self.btn_reset_all      = _btn("重置全部重查", "#795548", "#4E342E", self._reset_all)

        for b in [self.btn_search, self.btn_stop, self.btn_workbench,
                  self.btn_requery_errors, self.btn_reset_all, self.btn_save, self.btn_token]:
            btn_row.addWidget(b)
        btn_row.addStretch()
        vbox.addLayout(btn_row)

        # === 工作台统计 + 进度 ===
        bottom = QHBoxLayout()
        self.wb_stats = QLabel("工作台: 总计 0 | 已推 0 批 | 剩余 0")
        self.wb_stats.setStyleSheet("color:#1565C0; font-weight:bold")
        self.stats_label = QLabel("找到:0 | 未找到:0 | 错误:0 | 跳过:0")
        self.progress = QProgressBar()
        self.progress.setFixedHeight(22)
        bottom.addWidget(self.wb_stats)
        bottom.addStretch()
        bottom.addWidget(self.stats_label)
        vbox.addLayout(bottom)
        vbox.addWidget(self.progress)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

    # ──────────── Token ────────────

    def _choose_results_dir(self):
        d = QFileDialog.getExistingDirectory(
            self, "选择结果保存目录", self.results_dir)
        if d:
            self.results_dir = d
            self.dir_label.setText(d)
            self._log(f"结果目录已更改: {d}")

    def _connect_token(self):
        self._log("正在获取 Token...")
        self.status_bar.showMessage("正在获取 Token...")
        self._token_worker = TokenWorker(ask_login=False)
        self._token_worker.log.connect(self._log)
        self._token_worker.success.connect(self._on_token_ok)
        self._token_worker.failed.connect(self._on_token_fail_silent)
        self._token_worker.start()

    def _refresh_token(self):
        self._log("正在关闭浏览器并以调试模式重新启动...")
        self.btn_token.setEnabled(False)
        self._token_worker = TokenWorker(ask_login=True)
        self._token_worker.log.connect(self._log)
        self._token_worker.success.connect(self._on_token_ok)
        self._token_worker.failed.connect(self._on_token_fail)
        self._token_worker.start()

    def _on_token_ok(self, token):
        self.token = token
        self._log("Token 获取成功")
        self.status_bar.showMessage("Token 就绪")
        self.btn_token.setEnabled(True)

    def _on_token_fail_silent(self, err):
        self._log(f"自动获取失败，请点「刷新 Token」: {err}")
        self.status_bar.showMessage("请点「刷新 Token」并在浏览器登录")
        self.btn_token.setEnabled(True)

    def _on_token_fail(self, err):
        self._log(f"Token 获取失败: {err}")
        self.status_bar.showMessage("Token 获取失败")
        self.btn_token.setEnabled(True)
        QMessageBox.warning(self, "失败", f"获取失败，请确认浏览器已打开并登录\n\n{err}")

    # ──────────── 日志 ────────────

    def _table_context_menu(self, pos):
        menu = QMenu(self)
        act_copy_cell  = menu.addAction("复制单元格")
        act_copy_row   = menu.addAction("复制整行")
        act_copy_all   = menu.addAction("复制选中区域")
        action = menu.exec(self.table.viewport().mapToGlobal(pos))
        if action == act_copy_cell:
            item = self.table.itemAt(pos)
            if item:
                QApplication.clipboard().setText(item.text())
        elif action == act_copy_row:
            row = self.table.rowAt(pos.y())
            if row >= 0:
                parts = []
                for ci in range(self.table.columnCount()):
                    it = self.table.item(row, ci)
                    parts.append(it.text() if it else "")
                QApplication.clipboard().setText("\t".join(parts))
        elif action == act_copy_all:
            self._copy_selection()

    def _copy_selection(self):
        selected = self.table.selectedRanges()
        if not selected:
            return
        sel_rows, sel_cols = set(), set()
        for r in selected:
            for ri in range(r.topRow(), r.bottomRow() + 1):
                sel_rows.add(ri)
            for ci in range(r.leftColumn(), r.rightColumn() + 1):
                sel_cols.add(ci)
        sel_rows, sel_cols = sorted(sel_rows), sorted(sel_cols)
        if len(sel_rows) == 1 and len(sel_cols) == 1:
            item = self.table.item(sel_rows[0], sel_cols[0])
            QApplication.clipboard().setText(item.text() if item else "")
            return
        headers = [self.table.horizontalHeaderItem(c).text()
                   if self.table.horizontalHeaderItem(c) else str(c)
                   for c in sel_cols]
        lines = ["\t".join(headers)]
        for ri in sel_rows:
            lines.append("\t".join(
                self.table.item(ri, ci).text() if self.table.item(ri, ci) else ""
                for ci in sel_cols))
        QApplication.clipboard().setText("\n".join(lines))

    def _table_key_press(self, event):
        is_copy = (event.key() == Qt.Key.Key_C and (
            event.modifiers() & Qt.KeyboardModifier.ControlModifier or
            event.modifiers() & Qt.KeyboardModifier.MetaModifier
        ))
        if is_copy:
            self._copy_selection()
        else:
            QTableWidget.keyPressEvent(self.table, event)

    def _log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{ts}] {msg}")
        sb = self.log_text.verticalScrollBar()
        sb.setValue(sb.maximum())
        QApplication.processEvents()

    # ──────────── 文件 / Sheet ────────────

    def _open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "打开 Excel", "",
                                              "Excel Files (*.xlsx *.xls)")
        if not path: return
        self.excel_path = path
        self.file_label.setText(os.path.basename(path))
        try:
            xls = pd.ExcelFile(path, engine="openpyxl")
            self.sheet_names = xls.sheet_names
            self.sheet_combo.blockSignals(True)
            self.sheet_combo.clear()
            self.sheet_combo.addItems(self.sheet_names)
            self.sheet_combo.blockSignals(False)
            # 自动检测是否有表头
            self._auto_detect_header(path, self.sheet_names[0])
            self._load_sheet(self.sheet_names[0])
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))

    def _auto_detect_header(self, path, sheet):
        """检测第一行是否包含表头关键词，自动设置复选框"""
        try:
            df_raw = pd.read_excel(path, sheet_name=sheet, engine="openpyxl", header=None, nrows=1)
            first_row_vals = [str(v).strip() for v in df_raw.iloc[0].tolist()]
            header_keywords = {"素材名称", "查询状态", "视频ID", "视频id", "名称", "状态"}
            has_header = bool(header_keywords & set(first_row_vals))
            self.chk_header.blockSignals(True)
            self.chk_header.setChecked(has_header)
            self.chk_header.blockSignals(False)
            self._log(f"自动检测: {'有' if has_header else '无'}表头")
        except Exception:
            pass

    def _on_sheet_changed(self, idx):
        if 0 <= idx < len(self.sheet_names):
            self._load_sheet(self.sheet_names[idx])

    def _load_sheet(self, sheet):
        try:
            has_header = self.chk_header.isChecked()
            header_arg = 0 if has_header else None
            self.df = pd.read_excel(self.excel_path, sheet_name=sheet,
                                    engine="openpyxl", header=header_arg)
            if not has_header:
                # 生成列名，第一列命名为素材名称
                cols = [f"列{i+1}" for i in range(len(self.df.columns))]
                cols[0] = "素材名称"
                self.df.columns = cols
                # 如果某列的第一行值是"查询状态"等关键词，说明是之前程序追加的列，删掉重建
                drop_cols = []
                for c in self.df.columns:
                    if str(self.df[c].iloc[0]).strip() in {"查询状态", "视频ID", "查询备注", "查询时间", "匹配名称"}:
                        drop_cols.append(c)
                if drop_cols:
                    self.df.drop(columns=drop_cols, inplace=True)
                    self._log(f"检测到旧结果列 {drop_cols}，已移除并重建")

            self.current_sheet = sheet
            for col in ["查询状态", "视频ID", "匹配名称", "查询备注", "查询时间"]:
                if col not in self.df.columns:
                    self.df[col] = ""
                else:
                    # 有表头时，保留已有状态（断点续查）
                    pass

            self._log(f"加载 [{sheet}]: {len(self.df)} 行  表头={'有' if has_header else '无'}")
            self.start_spin.setMaximum(max(0, len(self.df) - 1))
            self.col_combo.clear()
            self.col_combo.addItems([str(c) for c in self.df.columns])
            if "素材名称" in self.df.columns:
                self.col_combo.setCurrentText("素材名称")
            else:
                self.col_combo.setCurrentIndex(0)
            self._update_stats_label()
            self._apply_filter()
            self.status_bar.showMessage(f"[{sheet}] {len(self.df)} 行")
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))

    # ──────────── 筛选 ────────────

    def _apply_filter(self):
        if self.df is None:
            return
        status_filter = self.filter_combo.currentText()
        keyword = self.search_box.text().strip().lower()
        name_col = self.col_combo.currentText()

        mask = pd.Series([True] * len(self.df))

        if status_filter == "未查询":
            mask &= self.df["查询状态"].astype(str).str.strip() == ""
        elif status_filter != "全部":
            mask &= self.df["查询状态"].astype(str).str.strip() == status_filter

        if keyword and name_col in self.df.columns:
            mask &= self.df[name_col].astype(str).str.lower().str.contains(keyword, na=False)

        self.df_filtered = self.df[mask].copy()
        self._render_table(self.df_filtered)
        self.row_count_label.setText(f"显示: {len(self.df_filtered)} / {len(self.df)} 行")

    def _render_table(self, df):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels([str(c) for c in df.columns])

        name_col = self.col_combo.currentText()
        key_cols = {name_col, "查询状态", "视频ID", "匹配名称", "查询备注", "查询时间"}

        color_map = {"找到": "#C8E6C9", "未找到": "#FFF9C4", "错误": "#FFCDD2"}

        for ri, (_, row) in enumerate(df.iterrows()):
            status = str(row.get("查询状态", "")).strip()
            row_color = color_map.get(status)
            for ci, col in enumerate(df.columns):
                val = row[col]
                text = "" if pd.isna(val) else str(val)
                if str(col) not in key_cols and len(text) > 25:
                    text = text[:25] + "…"
                item = QTableWidgetItem(text)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                if str(col) == "查询状态" and row_color:
                    item.setBackground(QColor(row_color))
                elif row_color:
                    item.setBackground(QColor(row_color).lighter(120))
                self.table.setItem(ri, ci, item)

        self.table.resizeColumnsToContents()
        self.table.setSortingEnabled(True)

    def _update_row(self, orig_idx, name, status, ids, matched_names, remark):
        if self.df is None: return
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.df.at[orig_idx, "查询状态"]  = status
        self.df.at[orig_idx, "视频ID"]    = ids
        self.df.at[orig_idx, "匹配名称"]  = matched_names
        self.df.at[orig_idx, "查询备注"]  = remark
        self.df.at[orig_idx, "查询时间"]  = now
        # 实时更新统计标签
        self._update_stats_label()
        # 在表格中找到对应行并原地更新（不重渲染整表）
        self._patch_table_row(orig_idx, status, ids, matched_names, remark, now)

    def _patch_table_row(self, orig_idx, status, ids, matched_names, remark, now):
        """在当前表格中找到 orig_idx 对应的行，原地更新状态/ID/备注"""
        if self.df_filtered is None:
            return
        color_map = {"找到": "#C8E6C9", "未找到": "#FFF9C4", "错误": "#FFCDD2"}
        row_color = color_map.get(status)
        cols = list(self.df_filtered.columns) if self.df_filtered is not None else list(self.df.columns)
        update_vals = {
            "查询状态": status, "视频ID": ids,
            "匹配名称": matched_names, "查询备注": remark, "查询时间": now,
        }
        # 找到表格中 orig_idx 对应的显示行（df_filtered 的 index）
        if self.df_filtered is not None and orig_idx in self.df_filtered.index:
            display_row = list(self.df_filtered.index).index(orig_idx)
            for ci, col in enumerate(cols):
                col_s = str(col)
                if col_s in update_vals:
                    val = update_vals[col_s]
                    item = self.table.item(display_row, ci)
                    if item is None:
                        item = QTableWidgetItem()
                        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                        self.table.setItem(display_row, ci, item)
                    item.setText(str(val) if val else "")
                    if col_s == "查询状态" and row_color:
                        item.setBackground(QColor(row_color))
                    elif row_color:
                        item.setBackground(QColor(row_color).lighter(120))
        else:
            pass  # 该行不在当前筛选视图中，不需要更新表格

    def _update_stats_label(self):
        if self.df is None: return
        vc = self.df["查询状态"].value_counts()
        found    = vc.get("找到", 0)
        notfound = vc.get("未找到", 0)
        error    = vc.get("错误", 0)
        pending  = (self.df["查询状态"].astype(str).str.strip() == "").sum()
        # 计算视频总数（视频ID列逗号分隔）
        total_videos = 0
        if "视频ID" in self.df.columns:
            for v in self.df["视频ID"].dropna():
                s = str(v).strip()
                if s and s != "nan":
                    total_videos += len([x for x in s.split(",") if x.strip()])
        self.stats_label.setText(
            f"✅ 找到:{found}行/{total_videos}个视频  ❌ 未找到:{notfound}  ⚠️ 错误:{error}  ⏳ 未查:{pending}"
        )
        self.stats_label.setStyleSheet(
            "font-weight:bold; font-size:13px;"
            f"color:{'#1B5E20' if found > 0 else '#333'}"
        )

    # ──────────── 查询 ────────────

    def _start_search(self):
        if self.df is None:
            QMessageBox.warning(self, "提示", "请先导入 Excel"); return
        name_col = self.col_combo.currentText()
        if not name_col or name_col not in self.df.columns:
            QMessageBox.warning(self, "提示", "请选择搜索列"); return
        if not self.token:
            self._refresh_token()
            if not self.token: return

        self.btn_search.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.progress.setMaximum(len(self.df))
        self.progress.setValue(0)
        self._log(f"开始查询 | 列:{name_col} | 起始行:{self.start_spin.value()}")

        self.search_worker = SearchWorker(
            self.df, name_col, self.start_spin.value(), self.token)
        self.search_worker.progress.connect(
            lambda c, t: self.progress.setValue(c))
        self.search_worker.row_done.connect(self._update_row)
        self.search_worker.log.connect(self._log)
        self.search_worker.token_refreshed.connect(
            lambda t: setattr(self, "token", t))
        self.search_worker.finished_signal.connect(self._search_done)
        self.search_worker.start()

    def _stop_search(self):
        if self.search_worker:
            self.search_worker.stop()
            self._log("正在停止...")

    def _search_done(self, found, not_found, error, skip):
        self.btn_search.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self._update_stats_label()
        self._log(f"查询完成 | 找到:{found} 未找到:{not_found} 错误:{error} 跳过:{skip}")
        self.status_bar.showMessage("查询完成")
        self._auto_save()

    def _requery_errors(self):
        if self.df is None: return
        mask = self.df["查询状态"].astype(str).str.strip() == "错误"
        count = mask.sum()
        if count == 0:
            QMessageBox.information(self, "提示", "没有错误行"); return
        self.df.loc[mask, "查询状态"] = ""
        self.df.loc[mask, "视频ID"]   = ""
        self.df.loc[mask, "查询备注"] = ""
        self._log(f"已重置 {count} 条错误行，开始重查...")
        self._apply_filter()
        self._start_search()

    def _reset_all(self):
        if self.df is None: return
        ret = QMessageBox.question(
            self, "确认",
            f"将清空所有 {len(self.df)} 行的查询状态，重新查询全部行。\n确认吗？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if ret != QMessageBox.StandardButton.Yes: return
        self.df["查询状态"] = ""
        self.df["视频ID"]   = ""
        self.df["匹配名称"] = ""
        self.df["查询备注"] = ""
        self.df["查询时间"] = ""
        self._log(f"已重置全部 {len(self.df)} 行，开始重查...")
        self._apply_filter()
        self._start_search()

    def _save_sheet_to_file(self, path, sheet_name, df):
        """保存当前 sheet 到文件，保留其他 sheet 不动"""
        from openpyxl import load_workbook
        from openpyxl.utils.dataframe import dataframe_to_rows
        import openpyxl

        try:
            wb = load_workbook(path)
        except Exception:
            wb = openpyxl.Workbook()
            if "Sheet" in wb.sheetnames:
                del wb["Sheet"]

        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
        ws = wb.create_sheet(title=sheet_name)

        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        wb.save(path)

    def _auto_save(self):
        if not self.excel_path or not self.current_sheet:
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = os.path.splitext(os.path.basename(self.excel_path))[0]
        out_path = os.path.join(self.results_dir, f"{base}_结果_{ts}.xlsx")
        try:
            self.df.to_excel(out_path, index=False, engine="openpyxl")
            self._log(f"结果已保存: {out_path}")
        except Exception as e:
            self._log(f"自动保存失败: {e}")

    def _save_excel(self):
        if self.df is None: return
        base = os.path.splitext(os.path.basename(self.excel_path))[0] if self.excel_path else "结果"
        default_path = os.path.join(self.results_dir, f"{base}_结果.xlsx")
        path, _ = QFileDialog.getSaveFileName(self, "保存", default_path, "Excel (*.xlsx)")
        if not path: return
        try:
            self._save_sheet_to_file(path, self.current_sheet or "Sheet1", self.df)
            self._log(f"已保存: {path}")
        except Exception as e:
            QMessageBox.critical(self, "保存失败", str(e))

    # ──────────── 工作台 ────────────

    def _add_to_workbench(self):
        if self.df is None:
            QMessageBox.warning(self, "提示", "请先导入并查询"); return

        all_ids = []
        for v in self.df[self.df["查询状态"] == "找到"]["视频ID"].dropna():
            for sid in str(v).split(","):
                sid = sid.strip()
                if sid and sid not in ("nan", ""):
                    try:
                        all_ids.append(str(int(float(sid))))
                    except (ValueError, OverflowError):
                        pass
        all_ids = list(dict.fromkeys(all_ids))

        if not all_ids:
            QMessageBox.information(self, "提示", "没有可添加的视频 ID"); return

        self._log(f"收集到 {len(all_ids)} 个唯一视频 ID，获取完整数据...")
        self.btn_workbench.setEnabled(False)
        QApplication.processEvents()

        try:
            objects = api_core.fetch_video_objects_by_ids(self.token, all_ids)
        except Exception as e:
            self._log(f"获取失败: {e}")
            self.btn_workbench.setEnabled(True)
            return

        if not objects:
            QMessageBox.information(self, "提示", "未获取到视频数据")
            self.btn_workbench.setEnabled(True)
            return

        self._wb_total_objs = len(objects)
        self._wb_pushed     = 0
        self._update_wb_stats(0, math.ceil(len(objects) / 200))
        self._log(f"获取到 {len(objects)} 个视频对象，开始分批写入...")

        self.wb_worker = WorkbenchWorker(objects, batch_size=200)
        self.wb_worker.log.connect(self._log)
        self.wb_worker.batch_done.connect(self._wb_batch_done)
        self.wb_worker.all_done.connect(self._wb_all_done)
        self.progress.setMaximum(math.ceil(len(objects) / 200))
        self.progress.setValue(0)
        self.wb_worker.start()

    def _wb_batch_done(self, bi, total_b, count):
        self._wb_pushed = bi
        remain_batches  = total_b - bi
        remain_videos   = remain_batches * 200  # 估算
        self.progress.setValue(bi)
        self._update_wb_stats(bi, total_b)
        self._log(f"第 {bi}/{total_b} 批（{count}个）写入完成")

        if bi < total_b:
            ret = QMessageBox.question(
                self, "工作台",
                f"✅ 第 {bi} 批（{count} 个）已加入工作台！\n\n"
                f"请在浏览器处理推送后，点 Yes 继续第 {bi+1} 批\n"
                f"剩余 {total_b - bi} 批",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if ret == QMessageBox.StandardButton.Yes:
                if self.wb_worker: self.wb_worker.continue_next()
            else:
                if self.wb_worker: self.wb_worker.stop()
                self.btn_workbench.setEnabled(True)
                self._log("已停止")

    def _wb_all_done(self):
        self.btn_workbench.setEnabled(True)
        total_b = math.ceil(self._wb_total_objs / 200)
        self._update_wb_stats(total_b, total_b)
        self.progress.setValue(self.progress.maximum())
        self._log("全部批次加入工作台完成！")
        QMessageBox.information(self, "完成", "全部视频已加入工作台！")

    def _update_wb_stats(self, pushed_batches, total_batches):
        pushed_videos  = min(pushed_batches * 200, self._wb_total_objs)
        remain_videos  = self._wb_total_objs - pushed_videos
        self.wb_stats.setText(
            f"工作台: 总计 {self._wb_total_objs} 个 | "
            f"已推 {pushed_batches}/{total_batches} 批 ({pushed_videos}个) | "
            f"剩余 {remain_videos} 个"
        )


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
