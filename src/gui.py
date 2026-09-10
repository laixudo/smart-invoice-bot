"""V0.3 GUI：极简秩序风（大圆角卡片 / 灵动岛 / 深浅主题）。"""

from __future__ import annotations

import logging
import os
import threading
from pathlib import Path
from tkinter import filedialog

import customtkinter as ctk

from .config_loader import load_config
from .excel_to_pdf import list_office_choices
from .local_settings import load_prefs, save_prefs
from .mailer import SMTP_PRESETS
from .paths import app_root
from .pipeline import run_pipeline
from .ui_motion import SoftDropdown


logger = logging.getLogger(__name__)

PROVIDER_LABELS = [(v["label"], k) for k, v in SMTP_PRESETS.items()]
PROVIDER_BY_LABEL = {lb: k for lb, k in PROVIDER_LABELS}
PROVIDER_BY_KEY = {k: lb for lb, k in PROVIDER_LABELS}

# —— 设计令牌（克制、统一）——
R = 18          # 大卡片圆角
R_CTL = 14      # 控件圆角
R_PILL = 20     # 灵动岛 / 胶囊
ACCENT = "#3478F6"
ACCENT_HOVER = "#2A66D9"
MUTED_L = "#8E8E93"
MUTED_D = "#98989F"
LINE_L = "#E5E5EA"
LINE_D = "#3A3A3C"
CARD_L = "#FFFFFF"
CARD_D = "#1C1C1E"
BG_L = "#F2F2F7"
BG_D = "#000000"
FONT = "Microsoft YaHei UI"


def apply_theme(_mode: str = "Light") -> None:
    ctk.set_appearance_mode("Light")
    ctk.set_default_color_theme("blue")


class SoftDialog(ctk.CTkToplevel):
    """轻量确认弹窗：圆角卡片 + 柔和按钮。"""

    def __init__(self, master, title: str, message: str, kind: str = "info"):
        super().__init__(master)
        self.result = False
        self.title(title or "提示")
        self.resizable(False, False)
        self.transient(master)
        self.configure(fg_color=(BG_L, BG_D))
        self.protocol("WM_DELETE_WINDOW", self._no)

        w, h = 380, 200
        self.geometry(f"{w}x{h}")
        self.update_idletasks()
        try:
            x = master.winfo_rootx() + max(0, (master.winfo_width() - w) // 2)
            y = master.winfo_rooty() + max(0, (master.winfo_height() - h) // 2)
            self.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            pass
        self.lift()
        self.focus_force()
        self.after(1, self._arm_modal)

        card = ctk.CTkFrame(
            self,
            corner_radius=R,
            fg_color=(CARD_L, CARD_D),
            border_width=1,
            border_color=(LINE_L, LINE_D),
        )
        card.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(
            card,
            text=title,
            font=ctk.CTkFont(family=FONT, size=16, weight="bold"),
            anchor="w",
        ).pack(fill="x", padx=20, pady=(18, 6))
        ctk.CTkLabel(
            card,
            text=message,
            font=ctk.CTkFont(family=FONT, size=13),
            text_color=(MUTED_L, MUTED_D),
            wraplength=300,
            justify="left",
            anchor="w",
        ).pack(fill="x", padx=20, pady=(0, 16))

        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(fill="x", padx=20, pady=(0, 18))
        if kind == "confirm":
            ctk.CTkButton(
                row,
                text="取消",
                width=100,
                height=36,
                corner_radius=R_CTL,
                fg_color=("gray92", "gray25"),
                hover_color=("gray86", "gray32"),
                text_color=("gray20", "gray90"),
                command=self._no,
            ).pack(side="right", padx=(8, 0))
            ctk.CTkButton(
                row,
                text="确定",
                width=100,
                height=36,
                corner_radius=R_CTL,
                fg_color=ACCENT,
                hover_color=ACCENT_HOVER,
                command=self._yes,
            ).pack(side="right")
        else:
            ctk.CTkButton(
                row,
                text="好",
                width=100,
                height=36,
                corner_radius=R_CTL,
                fg_color=ACCENT,
                hover_color=ACCENT_HOVER,
                command=self._yes,
            ).pack(side="right")
        self.bind("<Escape>", lambda e: self._no())
        self.bind("<Return>", lambda e: self._yes())

    def _arm_modal(self) -> None:
        try:
            self.grab_set()
            self.lift()
            self.focus_force()
        except Exception:
            pass

    def _yes(self) -> None:
        self.result = True
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

    def _no(self) -> None:
        self.result = False
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()


def ask_ok(master, title: str, message: str) -> None:
    d = SoftDialog(master, title, message, kind="info")
    master.wait_window(d)


def ask_warn(master, title: str, message: str) -> None:
    d = SoftDialog(master, title, message, kind="info")
    master.wait_window(d)


def ask_confirm(master, title: str, message: str) -> bool:
    d = SoftDialog(master, title, message, kind="confirm")
    master.wait_window(d)
    return bool(d.result)


class App(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("SMART 单票 Invoice 助手")
        self.geometry("920x680")
        self.minsize(720, 520)
        # 低配机先亮窗，避免长时间黑屏
        self.configure(fg_color=(BG_L, BG_D))
        self.update_idletasks()
        try:
            self.deiconify()
            self.lift()
            self.update()
        except Exception:
            pass

        self.root_dir = app_root()
        self.cfg = load_config()
        self.prefs = load_prefs()
        apply_theme("Light")
        self.configure(fg_color=(BG_L, BG_D))

        provider = self.prefs.get("provider", "qq")
        if provider not in SMTP_PRESETS:
            provider = "qq"
        self.provider_key = ctk.StringVar(value=provider)
        self.username = ctk.StringVar(value="")
        self.auth_code = ctk.StringVar(value="")
        self.show_auth = ctk.BooleanVar(value=False)
        self.office_engine = ctk.StringVar(value=self.prefs.get("office_engine", "auto"))

        # 默认不预填路径、不自动建文件夹，由用户点「浏览」选择
        self.excel_dir = ctk.StringVar(value=self.prefs.get("excel_dir", ""))
        self.result_dir = ctk.StringVar(value=self.prefs.get("result_dir", ""))
        self.status_text = ctk.StringVar(value="就绪")
        self._running = False
        self._abort = threading.Event()

        # 根布局：内容 + 底栏
        self.body = ctk.CTkFrame(self, fg_color="transparent")
        self.body.pack(fill="both", expand=True, padx=32, pady=(20, 8))

        self.statusbar = ctk.CTkFrame(
            self,
            height=36,
            corner_radius=0,
            fg_color=("gray96", "gray12"),
            border_width=0,
        )
        self.statusbar.pack(fill="x", side="bottom")
        self.statusbar.pack_propagate(False)
        ctk.CTkLabel(
            self.statusbar,
            textvariable=self.status_text,
            font=ctk.CTkFont(family=FONT, size=11),
            text_color=(MUTED_L, MUTED_D),
            anchor="w",
        ).pack(side="left", padx=20)
        ctk.CTkLabel(
            self.statusbar,
            text="v0.3  ·  不发邮件",
            font=ctk.CTkFont(family=FONT, size=11),
            text_color=(MUTED_L, MUTED_D),
            anchor="e",
        ).pack(side="right", padx=20)

        self.show_login()

    def _clear(self) -> None:
        for w in self.body.winfo_children():
            w.destroy()

    def _island(self, parent, text: str) -> ctk.CTkFrame:
        bar = ctk.CTkFrame(
            parent,
            height=34,
            corner_radius=R_PILL,
            fg_color=("gray90", "gray18"),
        )
        bar.pack_propagate(False)
        ctk.CTkLabel(
            bar,
            text=text,
            font=ctk.CTkFont(family=FONT, size=12),
            text_color=("gray35", "gray75"),
        ).pack(expand=True, padx=18)
        return bar

    def _card(self, parent, **kw) -> ctk.CTkFrame:
        return ctk.CTkFrame(
            parent,
            corner_radius=R,
            fg_color=(CARD_L, CARD_D),
            border_width=1,
            border_color=(LINE_L, LINE_D),
            **kw,
        )

    def _label_sm(self, parent, text: str) -> None:
        ctk.CTkLabel(
            parent,
            text=text,
            font=ctk.CTkFont(family=FONT, size=12),
            text_color=(MUTED_L, MUTED_D),
            anchor="w",
        ).pack(fill="x", padx=24, pady=(14, 6))

    def _icon_dot(self, parent, color: str) -> None:
        """圆润小图标点。"""
        ctk.CTkFrame(parent, width=10, height=10, corner_radius=5, fg_color=color).pack(
            side="left", padx=(0, 8)
        )

    # ---------- Login ----------
    def show_login(self) -> None:
        self._clear()
        self._page = "login"
        try:
            self.grab_release()
        except Exception:
            pass
        self.status_text.set("请完善账号信息后进入")

        wrap = ctk.CTkFrame(self.body, fg_color="transparent")
        wrap.pack(expand=True)

        island = self._island(wrap, "SMART Invoice  ·  仅生成 PDF")
        island.configure(width=260)
        island.pack(pady=(8, 22))

        card = self._card(wrap, width=440)
        card.pack()

        # 标题区
        head = ctk.CTkFrame(card, fg_color="transparent")
        head.pack(fill="x", padx=28, pady=(28, 8))
        ctk.CTkLabel(
            head,
            text="欢迎使用",
            font=ctk.CTkFont(family=FONT, size=24, weight="bold"),
            anchor="w",
        ).pack(fill="x")
        ctk.CTkLabel(
            head,
            text="填写邮箱信息后进入。本期只生成 Invoice，不会发送邮件。",
            font=ctk.CTkFont(family=FONT, size=13),
            text_color=(MUTED_L, MUTED_D),
            wraplength=360,
            justify="left",
            anchor="w",
        ).pack(fill="x", pady=(6, 0))

        # 分隔
        ctk.CTkFrame(card, height=1, fg_color=(LINE_L, LINE_D)).pack(fill="x", padx=24, pady=12)

        self._label_sm(card, "邮箱类型")
        labels = [x[0] for x in PROVIDER_LABELS]
        current = PROVIDER_BY_KEY.get(self.provider_key.get(), labels[0])
        self.provider_var = ctk.StringVar(value=current)
        self.provider_menu = SoftDropdown(
            card,
            labels,
            variable=self.provider_var,
            height=42,
            corner_radius=R_CTL,
        )
        self.provider_menu.set(current)
        self.provider_menu.pack(fill="x", padx=24)

        self._label_sm(card, "邮箱账号")
        ctk.CTkEntry(
            card,
            textvariable=self.username,
            height=42,
            corner_radius=R_CTL,
            border_width=1,
            border_color=(LINE_L, LINE_D),
            fg_color=("gray98", "gray17"),
            placeholder_text="name@example.com",
            font=ctk.CTkFont(family=FONT, size=13),
        ).pack(fill="x", padx=24)

        self._label_sm(card, "授权码")
        auth_row = ctk.CTkFrame(card, fg_color="transparent")
        auth_row.pack(fill="x", padx=24)
        self.auth_entry = ctk.CTkEntry(
            auth_row,
            textvariable=self.auth_code,
            height=42,
            corner_radius=R_CTL,
            border_width=1,
            border_color=(LINE_L, LINE_D),
            fg_color=("gray98", "gray17"),
            show="•",
            placeholder_text="客户端授权码",
            font=ctk.CTkFont(family=FONT, size=13),
        )
        self.auth_entry.pack(side="left", fill="x", expand=True)
        self.btn_show_auth = ctk.CTkButton(
            auth_row,
            text="显示",
            width=64,
            height=42,
            corner_radius=R_CTL,
            fg_color=("gray92", "gray25"),
            hover_color=("gray86", "gray32"),
            text_color=("gray20", "gray90"),
            border_width=1,
            border_color=(LINE_L, LINE_D),
            font=ctk.CTkFont(family=FONT, size=13),
            command=self._toggle_auth,
        )
        self.btn_show_auth.pack(side="left", padx=(8, 0))

        # 登录主按钮用原生 CTkButton，保证点击必达
        ctk.CTkButton(
            card,
            text="进入系统",
            height=46,
            corner_radius=R_CTL,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(family=FONT, size=15, weight="bold"),
            command=self.on_login,
        ).pack(fill="x", padx=24, pady=(16, 10))

        ctk.CTkLabel(
            card,
            text="授权码 ≠ 登录密码。可在邮箱「设置 → 账户 → POP3/SMTP」生成。",
            font=ctk.CTkFont(family=FONT, size=11),
            text_color=(MUTED_L, MUTED_D),
            wraplength=360,
            justify="left",
            anchor="w",
        ).pack(fill="x", padx=28, pady=(0, 26))

    def _toggle_auth(self) -> None:
        self.show_auth.set(not self.show_auth.get())
        show = self.show_auth.get()
        self.auth_entry.configure(show="" if show else "•")
        if hasattr(self, "btn_show_auth"):
            self.btn_show_auth.configure(text="隐藏" if show else "显示")

    def _on_theme(self, value: str) -> None:
        # 已取消主题切换，保留空实现以防旧调用
        return

    def on_login(self) -> None:
        label = self.provider_menu.get().strip()
        key = PROVIDER_BY_LABEL.get(label, "qq")
        user = self.username.get().strip()
        code = self.auth_code.get().strip()
        if not label or not user or not code:
            self.status_text.set("请填写完整：邮箱类型、账号、授权码")
            ask_warn(self, "还差一点点", "请填写完整：邮箱类型、账号、授权码")
            return
        if "@" not in user:
            self.status_text.set("请填写完整邮箱地址（需包含 @）")
            ask_warn(self, "账号格式", "请填写完整邮箱地址（需包含 @）")
            return
        self.provider_key.set(key)
        prefs = load_prefs()
        prefs["provider"] = key
        # 不再记住账号/授权码
        prefs.pop("username", None)
        prefs.pop("auth_code", None)
        prefs.pop("remember_account", None)
        prefs.pop("remember_auth", None)
        save_prefs(prefs)
        self.prefs = prefs
        self.status_text.set("已进入")
        self.show_main()

    # ---------- Main ----------
    def show_main(self) -> None:
        self._clear()
        self._page = "main"
        try:
            self.grab_release()
        except Exception:
            pass
        account = self.username.get().strip() or self.prefs.get("username", "")
        self.status_text.set("选择文件夹后开始生成")
        self._log_expanded = True

        # 顶栏固定；下方可滚动，缩小窗口也能看到进度
        top = ctk.CTkFrame(self.body, fg_color="transparent")
        top.pack(fill="x", pady=(0, 12))

        right = ctk.CTkFrame(top, fg_color="transparent")
        right.pack(side="right")
        ctk.CTkButton(
            right,
            text="切换账号",
            width=86,
            height=32,
            corner_radius=R_CTL,
            fg_color=("gray92", "gray25"),
            hover_color=("gray86", "gray32"),
            text_color=("gray20", "gray90"),
            border_width=1,
            border_color=(LINE_L, LINE_D),
            font=ctk.CTkFont(family=FONT, size=12),
            command=self.on_switch_account,
        ).pack(side="left")

        left = ctk.CTkFrame(top, fg_color="transparent")
        left.pack(side="left", fill="x", expand=True)
        title_row = ctk.CTkFrame(left, fg_color="transparent")
        title_row.pack(anchor="w")
        self._icon_dot(title_row, ACCENT)
        ctk.CTkLabel(
            title_row,
            text="Invoice 助手",
            font=ctk.CTkFont(family=FONT, size=20, weight="bold"),
        ).pack(side="left")

        short = account if len(account) <= 28 else account[:25] + "…"
        island = self._island(left, f"{short}  ·  仅生成")
        island.configure(width=min(280, 28 + len(f"{short}  ·  仅生成") * 8))
        island.pack(anchor="w", pady=(8, 0))

        self.scroll = ctk.CTkScrollableFrame(
            self.body,
            fg_color="transparent",
            corner_radius=0,
            scrollbar_button_color=("gray80", "gray35"),
            scrollbar_button_hover_color=("gray70", "gray45"),
        )
        self.scroll.pack(fill="both", expand=True)

        # 模块 1：路径
        card1 = self._card(self.scroll)
        card1.pack(fill="x", pady=(0, 12))
        self._module_head(card1, "01", "文件夹")
        self._path_row(card1, "待处理 Excel", self.excel_dir, self.browse_excel)
        self._path_row(card1, "结果保存位置", self.result_dir, self.browse_result)

        self._label_sm(card1, "导出引擎（Excel / WPS）")
        office_choices = list_office_choices()
        if not office_choices:
            office_choices = [("未检测到 Excel/WPS（将用内置兜底）", "auto")]
        labels = [x[0] for x in office_choices]
        self._office_by_label = {lb: key for lb, key in office_choices}
        self._office_by_key = {key: lb for lb, key in office_choices}
        cur_key = self.office_engine.get()
        if cur_key not in self._office_by_key:
            cur_key = office_choices[0][1]
            self.office_engine.set(cur_key)
        cur_label = self._office_by_key[cur_key]
        self.office_menu = SoftDropdown(
            card1,
            labels,
            variable=ctk.StringVar(value=cur_label),
            height=42,
            corner_radius=R_CTL,
            command=self._on_office_pick,
        )
        self.office_menu.set(cur_label)
        self.office_menu.pack(fill="x", padx=24, pady=(0, 4))
        ctk.CTkFrame(card1, height=14, fg_color="transparent").pack()

        # 模块 2：操作
        card2 = self._card(self.scroll)
        card2.pack(fill="x", pady=(0, 12))
        self._module_head(card2, "02", "生成")

        actions = ctk.CTkFrame(card2, fg_color="transparent")
        actions.pack(fill="x", padx=24, pady=(4, 10))
        self.btn_gen = ctk.CTkButton(
            actions,
            text="开始生成 Invoice",
            height=44,
            corner_radius=R_CTL,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(family=FONT, size=14, weight="bold"),
            command=self.on_generate,
        )
        self.btn_gen.pack(side="left")
        self.btn_open = ctk.CTkButton(
            actions,
            text="打开结果",
            width=110,
            height=44,
            corner_radius=R_CTL,
            fg_color=("gray92", "gray25"),
            hover_color=("gray86", "gray32"),
            text_color=("gray20", "gray90"),
            border_width=1,
            border_color=(LINE_L, LINE_D),
            font=ctk.CTkFont(family=FONT, size=13),
            command=self.open_result,
        )
        self.btn_open.pack(side="left", padx=10)
        self.btn_abort = ctk.CTkButton(
            actions,
            text="中断任务",
            width=110,
            height=44,
            corner_radius=R_CTL,
            fg_color=("#FF3B30", "#C62828"),
            hover_color=("#E0342B", "#B71C1C"),
            font=ctk.CTkFont(family=FONT, size=13, weight="bold"),
            command=self.on_abort,
            state="disabled",
        )
        self.btn_abort.pack(side="left")

        self.progress = ctk.CTkProgressBar(
            card2,
            height=5,
            corner_radius=8,
            progress_color=ACCENT,
            fg_color=("gray90", "gray25"),
        )
        self.progress.pack(fill="x", padx=24, pady=(6, 16))
        self.progress.set(0)

        # 模块 3：进度（可折叠 + 内部滚动）
        card3 = self._card(self.scroll)
        card3.pack(fill="x", pady=(0, 12))
        head = ctk.CTkFrame(card3, fg_color="transparent", cursor="hand2")
        head.pack(fill="x", padx=20, pady=(16, 4))
        badge = ctk.CTkFrame(
            head, width=28, height=28, corner_radius=10, fg_color=("gray92", "gray25")
        )
        badge.pack(side="left")
        badge.pack_propagate(False)
        ctk.CTkLabel(
            badge,
            text="03",
            font=ctk.CTkFont(family=FONT, size=11, weight="bold"),
            text_color=(MUTED_L, MUTED_D),
        ).pack(expand=True)
        ctk.CTkLabel(
            head,
            text="进度",
            font=ctk.CTkFont(family=FONT, size=14, weight="bold"),
        ).pack(side="left", padx=10)
        self.log_toggle_btn = ctk.CTkButton(
            head,
            text="收起 ▴",
            width=72,
            height=28,
            corner_radius=10,
            fg_color=("gray92", "gray25"),
            hover_color=("gray86", "gray32"),
            text_color=("gray25", "gray85"),
            font=ctk.CTkFont(family=FONT, size=12),
            command=self._toggle_log,
        )
        self.log_toggle_btn.pack(side="right")
        for w in (head, badge):
            w.bind("<Button-1>", lambda e: self._toggle_log())

        self.log_wrap = ctk.CTkFrame(card3, fg_color="transparent")
        self.log_wrap.pack(fill="x", padx=16, pady=(0, 16))
        self.log = ctk.CTkTextbox(
            self.log_wrap,
            height=220,
            corner_radius=R_CTL,
            border_width=0,
            fg_color=("gray97", "gray14"),
            text_color=("gray20", "gray90"),
            font=ctk.CTkFont(family=FONT, size=12),
            wrap="word",
            activate_scrollbars=True,
        )
        self.log.pack(fill="both", expand=True)
        self._append("就绪。确认路径后点击「开始生成 Invoice」。窗口缩小时可滚动本页，或展开「进度」。")

    def _module_head(self, parent, index: str, title: str) -> None:
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=20, pady=(16, 4))
        badge = ctk.CTkFrame(
            row, width=28, height=28, corner_radius=10, fg_color=("gray92", "gray25")
        )
        badge.pack(side="left")
        badge.pack_propagate(False)
        ctk.CTkLabel(
            badge,
            text=index,
            font=ctk.CTkFont(family=FONT, size=11, weight="bold"),
            text_color=(MUTED_L, MUTED_D),
        ).pack(expand=True)
        ctk.CTkLabel(
            row,
            text=title,
            font=ctk.CTkFont(family=FONT, size=14, weight="bold"),
        ).pack(side="left", padx=10)

    def _path_row(self, parent, title: str, var: ctk.StringVar, cmd) -> None:
        self._label_sm(parent, title)
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=24, pady=(0, 2))
        ctk.CTkEntry(
            row,
            textvariable=var,
            height=42,
            corner_radius=R_CTL,
            border_width=1,
            border_color=(LINE_L, LINE_D),
            fg_color=("gray98", "gray17"),
            font=ctk.CTkFont(family=FONT, size=12),
        ).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(
            row,
            text="浏览",
            width=72,
            height=42,
            corner_radius=R_CTL,
            fg_color=("gray92", "gray25"),
            hover_color=("gray86", "gray32"),
            text_color=("gray20", "gray90"),
            border_width=1,
            border_color=(LINE_L, LINE_D),
            font=ctk.CTkFont(family=FONT, size=13),
            command=cmd,
        ).pack(side="left", padx=(8, 0))

    def _on_office_pick(self, label: str) -> None:
        key = self._office_by_label.get(label, "auto")
        self.office_engine.set(key)
        prefs = load_prefs()
        prefs["office_engine"] = key
        save_prefs(prefs)

    def on_switch_account(self) -> None:
        if self._running:
            if not ask_confirm(
                self,
                "确认中断并切换账号？",
                "当前正在生成 Invoice。切换账号将中断任务，已生成的 PDF 会保留。",
            ):
                return
            self._abort.set()
            self._append("已请求中断，正在停止…")
        self.show_login()

    def on_abort(self) -> None:
        if not self._running:
            ask_warn(self, "没有进行中的任务", "当前没有正在生成的任务。")
            return
        if not ask_confirm(
            self,
            "确认中断任务？",
            "中断后将停止继续生成。已经生成好的 PDF 会保留，可稍后继续处理剩余文件。",
        ):
            return
        self._abort.set()
        self._append("已确认中断，正在停止…")
        self.status_text.set("正在中断…")

    def browse_excel(self) -> None:
        path = filedialog.askdirectory(
            parent=self,
            title="选择待处理 Excel 文件夹",
            initialdir=self.excel_dir.get() or str(self.root_dir),
        )
        if path:
            self.excel_dir.set(path)
            self.status_text.set(f"已选 Excel 目录：{path}")

    def browse_result(self) -> None:
        path = filedialog.askdirectory(
            parent=self,
            title="选择结果保存文件夹",
            initialdir=self.result_dir.get() or str(self.root_dir),
        )
        if path:
            self.result_dir.set(path)
            self.status_text.set(f"已选结果目录：{path}")

    def open_result(self) -> None:
        path = Path(self.result_dir.get().strip())
        if not path.exists():
            self.status_text.set("结果文件夹还不存在，请先生成")
            ask_warn(self, "还没有结果", "结果文件夹还不存在，请先生成")
            return
        os.startfile(str(path))  # noqa: S606
        self.status_text.set(f"已打开：{path}")

    def _toggle_log(self) -> None:
        self._log_expanded = not getattr(self, "_log_expanded", True)
        if not hasattr(self, "log_wrap"):
            return
        if self._log_expanded:
            self.log_wrap.pack(fill="x", padx=16, pady=(0, 16))
            if hasattr(self, "log_toggle_btn"):
                self.log_toggle_btn.configure(text="收起 ▴")
        else:
            self.log_wrap.pack_forget()
            if hasattr(self, "log_toggle_btn"):
                self.log_toggle_btn.configure(text="展开 ▾")

    def _append(self, text: str) -> None:
        if not hasattr(self, "log"):
            self.status_text.set(text if len(text) < 48 else text[:45] + "…")
            return
        # 有新日志时自动展开进度区，方便缩小窗口后仍能看到
        if hasattr(self, "log_wrap") and not getattr(self, "_log_expanded", True):
            self._log_expanded = False
            self._toggle_log()
        self.log.insert("end", text + "\n")
        self.log.see("end")
        short = text if len(text) < 48 else text[:45] + "…"
        self.status_text.set(short)
        # 滚到页面底部进度区
        try:
            if hasattr(self, "scroll"):
                self.scroll._parent_canvas.yview_moveto(1.0)  # noqa: SLF001
        except Exception:
            pass

    def _save_paths(self) -> None:
        prefs = load_prefs()
        prefs["excel_dir"] = self.excel_dir.get().strip()
        prefs["result_dir"] = self.result_dir.get().strip()
        prefs["office_engine"] = self.office_engine.get().strip() or "auto"
        save_prefs(prefs)

    def _set_running_ui(self, running: bool) -> None:
        self._running = running
        if running:
            self.btn_gen.configure(state="disabled")
            self.btn_abort.configure(state="normal")
        else:
            self.btn_gen.configure(state="normal")
            if hasattr(self, "btn_abort"):
                self.btn_abort.configure(state="disabled")

    def on_generate(self) -> None:
        excel_dir = Path(self.excel_dir.get().strip())
        result_dir = Path(self.result_dir.get().strip())
        if not excel_dir.exists() or not excel_dir.is_dir():
            ask_warn(self, "路径无效", "请先选择有效的「待处理 Excel 文件夹」")
            return
        if not str(result_dir).strip():
            ask_warn(self, "路径无效", "请先选择「结果保存文件夹」")
            return
        files = sorted(
            p
            for p in excel_dir.iterdir()
            if p.is_file()
            and p.suffix.lower() in {".xlsx", ".xlsm"}
            and not p.name.startswith("~$")
        )
        if not files:
            ask_warn(self, "没有 Excel", "该文件夹里没有 .xlsx，请放入已确认账单。")
            return
        # 同步下拉当前值
        try:
            label = self.office_menu.get().strip()
            self.office_engine.set(self._office_by_label.get(label, "auto"))
        except Exception:
            pass
        self._save_paths()
        self._abort.clear()
        self._set_running_ui(True)
        self.progress.set(0.04)
        engine = self.office_engine.get() or "auto"
        self._append(f"开始处理，共 {len(files)} 个 Excel…（引擎：{engine}）")
        threading.Thread(
            target=self._generate_worker,
            args=(files, result_dir, engine),
            daemon=True,
        ).start()

    def _generate_worker(
        self, files: list[Path], result_dir: Path, office_engine: str
    ) -> None:
        cancelled = False
        try:
            result_dir.mkdir(parents=True, exist_ok=True)
            multi = len(files) > 1
            ok_n = fail_n = 0
            total = len(files)
            for i, f in enumerate(files, 1):
                if self._abort.is_set():
                    cancelled = True
                    break
                out_dir = result_dir / f.stem if multi else result_dir
                log_dir = (
                    result_dir / "_logs" / f.stem if multi else result_dir / "_logs"
                )
                self.after(0, lambda name=f.name: self._append(f"正在处理：{name}"))
                result = run_pipeline(
                    input_xlsx=f,
                    output_dir=out_dir,
                    logs_dir=log_dir,
                    office_engine=office_engine,
                    cancel_check=self._abort.is_set,
                )
                if result.get("cancelled"):
                    cancelled = True
                    self.after(
                        0,
                        lambda r=result: self._append(f"■ 已中断 · {r['message']}"),
                    )
                    break
                self.after(0, lambda v=i / total: self.progress.set(v))
                if result["batch_ok"]:
                    ok_n += 1
                    self.after(
                        0,
                        lambda r=result, o=out_dir: self._append(
                            f"✓ {r['bill_name']} · {r['container_count']} 个 PDF"
                        ),
                    )
                else:
                    fail_n += 1
                    self.after(
                        0,
                        lambda r=result: self._append(
                            f"✗ {r['bill_name']} · {r['message']}"
                        ),
                    )
            if cancelled:
                msg = f"任务已中断：成功 {ok_n}，失败 {fail_n}（已生成的文件保留）"
                self.after(0, lambda: self._append(msg))
                self.after(0, lambda: ask_ok(self, "已中断", msg))
            else:
                msg = f"完成：成功 {ok_n}，失败 {fail_n}"
                self.after(0, lambda: self.progress.set(1))
                self.after(0, lambda: self._append(msg))
                self.after(
                    0, lambda: ask_ok(self, "生成完成", msg + "\n可点击「打开结果」查看。")
                )
        except InterruptedError:
            self.after(0, lambda: self._append("任务已中断"))
            self.after(0, lambda: ask_ok(self, "已中断", "任务已中断，已生成的 PDF 会保留。"))
        except Exception as e:
            logger.exception("generate failed")
            self.after(0, lambda: ask_warn(self, "生成失败", str(e)))
            self.after(0, lambda: self._append(f"错误：{e}"))
        finally:
            self.after(0, lambda: self._set_running_ui(False))


def main() -> None:
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s"
    )
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
