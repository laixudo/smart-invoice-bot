"""按钮 / 下拉：柔和反馈。"""

from __future__ import annotations

import time
from typing import Callable

import customtkinter as ctk


def _blend(c1: str, c2: str, t: float) -> str:
    def rgb(h: str) -> tuple[int, int, int]:
        h = h.lstrip("#")
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

    r1, g1, b1 = rgb(c1)
    r2, g2, b2 = rgb(c2)
    r = int(r1 + (r2 - r1) * t)
    g = int(g1 + (g2 - g1) * t)
    b = int(b1 + (b2 - b1) * t)
    return f"#{r:02x}{g:02x}{b:02x}"


class SoftButton(ctk.CTkButton):
    """悬停轻微变色；不改尺寸，避免吃掉点击。"""

    def __init__(
        self,
        master,
        *,
        base_color: str = "#3478F6",
        hover_color: str = "#2A66D9",
        press_color: str = "#1E4FB8",
        animate_ms: int = 16,
        **kwargs,
    ):
        self._base = base_color
        self._hover = hover_color
        self._press = press_color
        self._animate_ms = animate_ms
        self._anim_id: str | None = None
        self._hovering = False
        self._user_command: Callable | None = kwargs.pop("command", None)

        kwargs.setdefault("fg_color", base_color)
        kwargs.setdefault("hover_color", hover_color)
        kwargs["command"] = self._on_click
        super().__init__(master, **kwargs)

        self.bind("<Enter>", self._on_enter, add="+")
        self.bind("<Leave>", self._on_leave, add="+")

    def _cancel_anim(self) -> None:
        if self._anim_id is not None:
            try:
                self.after_cancel(self._anim_id)
            except Exception:
                pass
            self._anim_id = None

    def _set_fg(self, color: str) -> None:
        try:
            if str(self.cget("state")) == "disabled":
                return
            self.configure(fg_color=color)
        except Exception:
            pass

    def _tween_to(self, target: str, steps: int = 4) -> None:
        self._cancel_anim()
        try:
            start = str(self.cget("fg_color"))
            if isinstance(start, (tuple, list)):
                start = start[0]
        except Exception:
            start = self._base
        if not isinstance(start, str) or not start.startswith("#"):
            start = self._base
        if start.lower() == target.lower():
            self._set_fg(target)
            return

        def step(i: int = 0, s: str = start) -> None:
            if i >= steps:
                self._set_fg(target)
                self._anim_id = None
                return
            t = (i + 1) / steps
            self._set_fg(_blend(s, target, t))
            self._anim_id = self.after(self._animate_ms, lambda: step(i + 1, s))

        step(0)

    def _on_enter(self, _e=None) -> None:
        if str(self.cget("state")) == "disabled":
            return
        self._hovering = True
        self._tween_to(self._hover)

    def _on_leave(self, _e=None) -> None:
        self._hovering = False
        self._tween_to(self._base)

    def _on_click(self) -> None:
        cmd = self._user_command
        self._set_fg(self._press)
        self.after(80, lambda: self._tween_to(self._hover if self._hovering else self._base))
        if cmd:
            cmd()


def primary_button(master, text: str, command, **kwargs) -> SoftButton:
    kwargs.setdefault("height", 46)
    kwargs.setdefault("corner_radius", 14)
    kwargs.setdefault(
        "font", ctk.CTkFont(family="Microsoft YaHei UI", size=15, weight="bold")
    )
    return SoftButton(
        master,
        text=text,
        command=command,
        base_color="#3478F6",
        hover_color="#2A66D9",
        press_color="#1E4FB8",
        **kwargs,
    )


def secondary_button(master, text: str, command, **kwargs) -> SoftButton:
    kwargs.setdefault("height", 42)
    kwargs.setdefault("corner_radius", 14)
    kwargs.setdefault("font", ctk.CTkFont(family="Microsoft YaHei UI", size=13))
    kwargs.setdefault("text_color", ("gray20", "gray90"))
    kwargs.setdefault("border_width", 1)
    kwargs.setdefault("border_color", ("#E5E5EA", "#3A3A3C"))
    mode = ctk.get_appearance_mode()
    if mode == "Dark":
        base, hover, press = "#2C2C2E", "#3A3A3C", "#48484A"
    else:
        base, hover, press = "#F2F2F7", "#E5E5EA", "#D1D1D6"
    return SoftButton(
        master,
        text=text,
        command=command,
        base_color=base,
        hover_color=hover,
        press_color=press,
        **kwargs,
    )


class SoftDropdown(ctk.CTkFrame):
    """
    自定义下拉。
    面板宽度与触发框完全对齐（处理 CTk 缩放，避免比登录框更宽）。
    """

    _GRACE_SEC = 0.28

    def __init__(
        self,
        master,
        values: list[str],
        *,
        variable: ctk.StringVar | None = None,
        height: int = 42,
        corner_radius: int = 14,
        command: Callable[[str], None] | None = None,
    ):
        super().__init__(master, fg_color="transparent")
        self._values = list(values)
        self._var = variable or ctk.StringVar(value=values[0] if values else "")
        self._command = command
        self._open = False
        self._panel: ctk.CTkToplevel | None = None
        self._anim_id: str | None = None
        self._opened_at = 0.0
        self._global_bound = False
        self._anchor_x = 0
        self._anchor_y = 0
        self._panel_w = 220
        self._panel_h = 100

        self._btn = ctk.CTkFrame(
            self,
            height=height,
            corner_radius=corner_radius,
            fg_color=("#F5F5F7", "#2C2C2E"),
            border_width=1,
            border_color=("#E5E5EA", "#3A3A3C"),
            cursor="hand2",
        )
        self._btn.pack(fill="x")
        self._btn.pack_propagate(False)

        self._label = ctk.CTkLabel(
            self._btn,
            textvariable=self._var,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13),
            anchor="w",
            cursor="hand2",
        )
        self._label.pack(side="left", fill="x", expand=True, padx=14)

        self._chev = ctk.CTkLabel(
            self._btn,
            text="▾",
            width=28,
            font=ctk.CTkFont(size=14),
            text_color=("#6C6C70", "#98989F"),
            cursor="hand2",
        )
        self._chev.pack(side="right", padx=10)

        for widget in (self._btn, self._label, self._chev):
            widget.bind("<ButtonRelease-1>", self._on_trigger)
            widget.bind("<Enter>", self._hover_on)
            widget.bind("<Leave>", self._hover_off)

        self.bind("<Destroy>", self._on_destroy)

    def get(self) -> str:
        return self._var.get()

    def set(self, value: str) -> None:
        self._var.set(value)

    def _on_destroy(self, _e=None) -> None:
        self.close()

    def _hover_on(self, _e=None) -> None:
        if not self._open:
            self._btn.configure(fg_color=("#EBEBF0", "#3A3A3C"))

    def _hover_off(self, _e=None) -> None:
        if not self._open:
            self._btn.configure(fg_color=("#F5F5F7", "#2C2C2E"))

    def _on_trigger(self, _e=None):
        now = time.monotonic()
        if now - getattr(self, "_last_toggle", 0) < 0.18:
            return "break"
        self._last_toggle = now
        if self._open:
            self.close()
        else:
            self.open()
        return "break"

    def close(self) -> None:
        was_open = self._open
        self._open = False
        try:
            self._chev.configure(text="▾")
            self._btn.configure(fg_color=("#F5F5F7", "#2C2C2E"))
        except Exception:
            pass
        if self._anim_id is not None:
            try:
                self.after_cancel(self._anim_id)
            except Exception:
                pass
            self._anim_id = None
        panel = self._panel
        self._panel = None
        if panel is not None:
            try:
                panel.destroy()
            except Exception:
                pass
        return was_open

    def _geo_size(self, screen_px: int) -> int:
        """CTk geometry 宽高会被再乘缩放，传入前先反缩放。"""
        try:
            scale = float(self._get_widget_scaling())
        except Exception:
            scale = 1.0
        if scale <= 0:
            scale = 1.0
        return max(1, int(round(screen_px / scale)))

    def open(self) -> None:
        self.close()
        self._open = True
        self._opened_at = time.monotonic()
        self._chev.configure(text="▴")
        self._btn.configure(fg_color=("#E5E5EA", "#48484A"))
        self.update_idletasks()

        panel = ctk.CTkToplevel(self)
        panel.withdraw()
        panel.overrideredirect(True)
        try:
            panel.attributes("-topmost", True)
        except Exception:
            pass

        x = self._btn.winfo_rootx()
        y = self._btn.winfo_rooty() + self._btn.winfo_height() + 4
        # 宽度：winfo 是屏幕像素，geometry 要反缩放才能与触发框对齐
        btn_w = int(self._btn.winfo_width())
        if btn_w < 40:
            btn_w = int(self.winfo_width()) or 280
        row_h = 36
        max_rows = min(len(self._values), 8)
        # 高度：内部控件 height/pad 是「设计像素」，geometry 也会再乘缩放
        # 因此高度直接用设计像素，不要再 /scale，否则选项会被裁切
        design_h = 14 + max_rows * (row_h + 2)

        geo_w = self._geo_size(btn_w)
        geo_h = design_h
        self._anchor_x, self._anchor_y = x, y
        self._panel_w, self._panel_h = geo_w, geo_h
        panel.geometry(f"{geo_w}x{geo_h}+{x}+{y - 6}")

        card = ctk.CTkFrame(
            panel,
            corner_radius=14,
            fg_color=("#FFFFFF", "#1C1C1E"),
            border_width=1,
            border_color=("#E5E5EA", "#3A3A3C"),
        )
        card.pack(fill="both", expand=True)

        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=6, pady=6)

        for val in self._values:
            selected = val == self._var.get()
            row = ctk.CTkFrame(
                body,
                height=row_h - 2,
                corner_radius=10,
                fg_color=("#3478F6" if selected else "transparent"),
                cursor="hand2",
            )
            row.pack(fill="x", pady=1)
            row.pack_propagate(False)
            lab = ctk.CTkLabel(
                row,
                text=val,
                anchor="w",
                font=ctk.CTkFont(family="Microsoft YaHei UI", size=13),
                text_color=("white" if selected else ("gray15", "gray90")),
                cursor="hand2",
            )
            lab.pack(fill="both", expand=True, padx=12)

            def _enter(e, r=row, sel=selected):
                if not sel:
                    r.configure(fg_color=("#F2F2F7", "#2C2C2E"))

            def _leave(e, r=row, sel=selected):
                if not sel:
                    r.configure(fg_color="transparent")

            def _pick(e, v=val):
                self._pick(v)
                return "break"

            for widget in (row, lab):
                widget.bind("<Enter>", _enter)
                widget.bind("<Leave>", _leave)
                widget.bind("<ButtonRelease-1>", _pick)

        self._panel = panel
        panel.deiconify()
        panel.lift()
        # 再钉一次宽度，防止被内容撑开
        panel.geometry(f"{geo_w}x{geo_h}+{x}+{y - 6}")
        self._animate_open(panel, 0)
        panel.bind("<Escape>", lambda e: self.close())

        if not self._global_bound:
            self.winfo_toplevel().bind_all(
                "<ButtonRelease-1>", self._on_global_release, add="+"
            )
            self._global_bound = True

    def _hit(self, x_root: int, y_root: int) -> bool:
        try:
            bx, by = self._btn.winfo_rootx(), self._btn.winfo_rooty()
            bw, bh = self._btn.winfo_width(), self._btn.winfo_height()
            if bx <= x_root <= bx + bw and by <= y_root <= by + bh:
                return True
            if self._panel is not None:
                px, py = self._panel.winfo_rootx(), self._panel.winfo_rooty()
                pw, ph = self._panel.winfo_width(), self._panel.winfo_height()
                if px <= x_root <= px + pw and py <= y_root <= py + ph:
                    return True
        except Exception:
            pass
        return False

    def _on_global_release(self, event) -> None:
        try:
            if not self.winfo_exists() or not self._open:
                return
        except Exception:
            return
        if time.monotonic() - self._opened_at < self._GRACE_SEC:
            return
        if self._hit(event.x_root, event.y_root):
            return
        self.close()

    def _animate_open(self, panel: ctk.CTkToplevel, step: int) -> None:
        if self._panel is not panel:
            return
        steps = 6
        t = min(1.0, (step + 1) / steps)
        ease = 1 - (1 - t) * (1 - t)
        y = int(self._anchor_y - 6 + 6 * ease)
        try:
            panel.geometry(f"{self._panel_w}x{self._panel_h}+{self._anchor_x}+{y}")
        except Exception:
            return
        if t < 1.0:
            self._anim_id = self.after(14, lambda: self._animate_open(panel, step + 1))
        else:
            self._anim_id = None

    def _pick(self, value: str) -> None:
        self._var.set(value)
        self.close()
        if self._command:
            self._command(value)
