# -*- coding: utf-8 -*-
import os
import re
import struct
import math
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from dataclasses import dataclass
from typing import Optional, Tuple


from PIL import Image, ImageTk
import cv2


def clamp(v, lo, hi):
    return max(lo, min(hi, v))


def rgb888_to_rgb565(r, g, b) -> int:
    return ((r & 0xF8) << 8) | ((g & 0xFC) << 3) | (b >> 3)


def safe_name(stem: str) -> str:
    return re.sub(r"[^\w\-\.]+", "_", stem)


@dataclass
class ItemState:
    path: str
    base_orig: Image.Image
    orig: Image.Image
    img_x: int = 0
    img_y: int = 0
    crop_box: Optional[Tuple[int, int, int, int]] = None
    _placed_once: bool = False


@dataclass
class VideoState:
    path: str
    cap: any
    native_fps: float
    total_frames: int
    duration_s: float
    first_frame_rgb: Image.Image


class RGB565ToolApp(tk.Tk):
    SECTOR_SIZE = 512  # 扇区固定 512

    def __init__(self):
        super().__init__()
        self.title("图片/视频 转 RGB565 工具      @今晚早起吃午饭")
        self.geometry("1280x840")

        self.items = []
        self.current_index = -1

        # UI vars
        self.var_disp_w = tk.IntVar(value=240)
        self.var_disp_h = tk.IntVar(value=320)

        self.var_img_w = tk.IntVar(value=240)
        self.var_img_h = tk.IntVar(value=320)

        self.var_autofit = tk.BooleanVar(value=True)
        self.var_lock_ar = tk.BooleanVar(value=True)

        self.var_rot = tk.StringVar(value="0")
        self.var_ms = tk.IntVar(value=0)  # 图片 old-bin 用；视频帧我默认写 0 交给fps
        self.var_endian = tk.StringVar(value="big")
        self.var_snap = tk.BooleanVar(value=True)

        self.var_out_dir = tk.StringVar(value="")
        self.var_fname_preview = tk.StringVar(value="")

        # 默认视频帧率为10
        self.DEFAULT_VIDEO_FPS = 10.0
        self.var_video_name = tk.StringVar(value="media_1_video.bin")

        # previews
        self.tk_orig_preview = None
        self.tk_conv_img = None
        self.conv_img_pil = None

        # drag
        self._dragging = False
        self._drag_start = (0, 0)
        self._img_start = (0, 0)

        # crop
        self._crop_start = None
        self._crop_rect_id = None

        # lock ratio handling
        self._ratio_guard = False
        self._last_ratio = self.var_img_w.get() / max(1, self.var_img_h.get())

        # rotation swap handling
        self._prev_rot = 0

        # -------- BIN preview mode --------
        self.bin_preview_active = False
        self.bin_preview_img: Optional[Image.Image] = None
        self.bin_preview_dx = 0
        self.bin_preview_dy = 0
        self.tk_bin_preview = None
        self.bin_preview_path = ""
        # --------------------------------

        # -------- VIDEO mode --------
        self.video_state: Optional[VideoState] = None
        self.video_preview_active = False
        self.video_preview_base: Optional[Image.Image] = None
        self.video_preview_path = ""
        # ---------------------------

        self._build_ui()
        self._bind_events()

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(50, self._refresh_all)

    # ---------- UI ----------
    def _build_ui(self):
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        left = ttk.Frame(self, padding=8)
        left.grid(row=0, column=0, sticky="nsw")
        left.rowconfigure(2, weight=1)

        center = ttk.Frame(self, padding=8)
        center.grid(row=0, column=1, sticky="nsew")
        center.columnconfigure(1, weight=1)
        center.rowconfigure(2, weight=1)

        ttk.Label(left, text="视频/图片列表").grid(row=0, column=0, sticky="w")

        btns = ttk.Frame(left)
        btns.grid(row=1, column=0, sticky="ew", pady=(6, 6))
        ttk.Button(btns, text="导入视频/图片", command=self.on_import_mixed).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(btns, text="移除选中", command=self.on_remove).grid(row=0, column=1)

        self.listbox = tk.Listbox(left, width=42, height=28)
        self.listbox.grid(row=2, column=0, sticky="nsew")

        # 添加预览区域标题
        preview_title = ttk.Label(
            center,
            text="左：原图 / 视频首帧预览    右：预览（图片可拖动 | 右键裁切 | 双击退出 BIN 预览）",
            anchor="w"
        )
        preview_title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(4, 0))

        fname_row = ttk.Frame(center)
        fname_row.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 2))
        fname_row.columnconfigure(1, weight=1)
        ttk.Label(fname_row, text="输出/预览标识：").grid(row=0, column=0, sticky="w")
        ttk.Label(fname_row, textvariable=self.var_fname_preview).grid(row=0, column=1, sticky="w")

        # 固定预览区域高度，防止跳动
        previews = ttk.Frame(center, height=420)
        previews.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 8))
        previews.grid_propagate(False)
        previews.columnconfigure(0, weight=1)
        previews.columnconfigure(1, weight=1)
        previews.rowconfigure(0, weight=1)

        # 固定左预览标签尺寸
        self.lbl_orig = ttk.Label(previews, text="", anchor="center", width=60)
        self.lbl_orig.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        # 固定右预览画布尺寸
        self.canvas = tk.Canvas(
            previews,
            bg="#202020",
            width=520,
            height=360,
            highlightthickness=1,
            highlightbackground="#555"
        )
        self.canvas.grid(row=0, column=1, sticky="nsew")

        params = ttk.LabelFrame(center, text="参数设置", padding=10)
        params.grid(row=3, column=0, sticky="ew")
        params.columnconfigure(7, weight=1)

        ttk.Label(params, text="显示器宽").grid(row=0, column=0, sticky="w")
        ttk.Entry(params, width=7, textvariable=self.var_disp_w).grid(row=0, column=1, padx=(4, 14))
        ttk.Label(params, text="显示器高").grid(row=0, column=2, sticky="w")
        ttk.Entry(params, width=7, textvariable=self.var_disp_h).grid(row=0, column=3, padx=(4, 14))

        ttk.Checkbutton(params, text="吸附 Snap", variable=self.var_snap, command=self._refresh_conv_preview).grid(
            row=0, column=4, padx=(0, 14)
        )
        ttk.Checkbutton(params, text="自适应最大边", variable=self.var_autofit, command=self._on_autofit_toggled).grid(
            row=0, column=5, padx=(0, 14)
        )

        ttk.Label(params, text="目标图宽").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(params, width=7, textvariable=self.var_img_w).grid(row=1, column=1, padx=(4, 14), pady=(8, 0))
        ttk.Label(params, text="目标图高").grid(row=1, column=2, sticky="w", pady=(8, 0))
        ttk.Entry(params, width=7, textvariable=self.var_img_h).grid(row=1, column=3, padx=(4, 14), pady=(8, 0))

        ttk.Checkbutton(params, text="锁比例", variable=self.var_lock_ar).grid(row=1, column=4, padx=(0, 14), pady=(8, 0))

        ttk.Label(params, text="旋转").grid(row=1, column=5, sticky="w", pady=(8, 0))
        ttk.Combobox(
            params, width=6, textvariable=self.var_rot, values=["0", "90", "180", "270"], state="readonly"
        ).grid(row=1, column=6, padx=(4, 14), pady=(8, 0))

        ttk.Label(params, text="播放时间(ms)").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(params, width=10, textvariable=self.var_ms).grid(row=2, column=1, padx=(4, 14), pady=(8, 0))

        ttk.Label(params, text="RGB565字节序").grid(row=2, column=2, sticky="w", pady=(8, 0))
        ttk.Combobox(
            params, width=10, textvariable=self.var_endian, values=["little", "big"], state="readonly"
        ).grid(row=2, column=3, padx=(4, 14), pady=(8, 0))

        ttk.Button(params, text="重置", command=self.on_reset_current).grid(row=2, column=4, padx=(0, 14), pady=(8, 0))
        ttk.Button(params, text="自适应最大边框", command=self.on_fit_max_frame).grid(
            row=2, column=5, padx=(0, 14), pady=(8, 0), sticky="w"
        )

        # 修改标签文本
        ttk.Label(params, text="输出名").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(params, width=18, textvariable=self.var_video_name).grid(row=3, column=3, padx=(4, 14), pady=(8, 0))

        actions = ttk.Frame(center)
        actions.grid(row=3, column=1, sticky="ne", padx=(8, 0))

        ttk.Button(actions, text="导出当前媒体", command=self.on_export_current).grid(row=2, column=0, sticky="ew", pady=(0, 10))
        ttk.Button(actions, text="批量导出媒体", command=self.on_export_all).grid(row=3, column=0, sticky="ew", pady=(0, 10))

        ttk.Separator(actions, orient="horizontal").grid(row=4, column=0, sticky="ew", pady=(0, 10))

        ttk.Button(actions, text="预览bin", command=self.on_preview_bin).grid(row=5, column=0, sticky="ew", pady=(0, 6))

        outrow = ttk.Frame(center)
        outrow.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        outrow.columnconfigure(1, weight=1)
        ttk.Label(outrow, text="输出文件夹").grid(row=0, column=0, sticky="w")
        ttk.Entry(outrow, textvariable=self.var_out_dir).grid(row=0, column=1, sticky="ew", padx=(8, 8))
        ttk.Button(outrow, text="选择…", command=self.on_choose_out_dir).grid(row=0, column=2)

    def _bind_events(self):
        self.listbox.bind("<<ListboxSelect>>", self.on_select)

        for v in [self.var_disp_w, self.var_disp_h, self.var_ms]:
            v.trace_add("write", lambda *args: self._refresh_all_safe())

        self.var_img_w.trace_add("write", lambda *args: self._on_img_w_changed())
        self.var_img_h.trace_add("write", lambda *args: self._on_img_h_changed())
        self.var_rot.trace_add("write", lambda *args: self._on_rot_changed())

        self.var_video_name.trace_add("write", lambda *args: self._refresh_conv_preview())

        self.canvas.bind("<ButtonPress-1>", self.on_left_down)
        self.canvas.bind("<B1-Motion>", self.on_left_move)
        self.canvas.bind("<ButtonRelease-1>", self.on_left_up)

        self.canvas.bind("<ButtonPress-3>", self.on_right_down)
        self.canvas.bind("<B3-Motion>", self.on_right_move)
        self.canvas.bind("<ButtonRelease-3>", self.on_right_up)

        self.canvas.bind("<Double-Button-1>", self.exit_bin_preview)
        self.bind("<Configure>", lambda e: self._refresh_conv_preview())

    def on_close(self):
        try:
            if self.video_state and self.video_state.cap:
                self.video_state.cap.release()
        except Exception:
            pass
        self.destroy()

    # ---------- helpers ----------
    def _get_rotated_size(self, img: Image.Image) -> Tuple[int, int]:
        rot = int(self.var_rot.get())
        if rot % 180 == 0:
            return img.width, img.height
        return img.height, img.width

    def _compute_max_fit_wh_for_current(self) -> Optional[Tuple[int, int]]:
        if self.current_index < 0:
            return None
        it = self.items[self.current_index]
        src = it.orig
        sw, sh = self._get_rotated_size(src)
        if sw <= 0 or sh <= 0:
            return None

        disp_w = max(1, int(self.var_disp_w.get()))
        disp_h = max(1, int(self.var_disp_h.get()))

        ar = sw / sh
        ar_box = disp_w / disp_h
        if ar >= ar_box:
            new_w = disp_w
            new_h = max(1, int(round(disp_w / ar)))
        else:
            new_h = disp_h
            new_w = max(1, int(round(disp_h * ar)))
        return new_w, new_h

    # ---------- ratio lock ----------
    def _on_img_w_changed(self):
        if self._ratio_guard:
            return
        try:
            w = int(self.var_img_w.get())
        except Exception:
            return
        w = max(1, w)

        if not self.var_lock_ar.get():
            h = max(1, int(self.var_img_h.get()))
            self._last_ratio = w / h
            self._refresh_all_safe()
            return

        self._ratio_guard = True
        try:
            ratio = self._last_ratio if self._last_ratio > 0 else (w / max(1, int(self.var_img_h.get())))
            h_new = max(1, int(round(w / ratio)))
            self.var_img_h.set(h_new)
        finally:
            self._ratio_guard = False

        self._refresh_all_safe()

    def _on_img_h_changed(self):
        if self._ratio_guard:
            return
        try:
            h = int(self.var_img_h.get())
        except Exception:
            return
        h = max(1, h)

        if not self.var_lock_ar.get():
            w = max(1, int(self.var_img_w.get()))
            self._last_ratio = w / h
            self._refresh_all_safe()
            return

        self._ratio_guard = True
        try:
            ratio = self._last_ratio if self._last_ratio > 0 else (max(1, int(self.var_img_w.get())) / h)
            w_new = max(1, int(round(h * ratio)))
            self.var_img_w.set(w_new)
        finally:
            self._ratio_guard = False

        self._refresh_all_safe()

    # ---------- autofit toggle ----------
    def _on_autofit_toggled(self):
        if self.var_autofit.get():
            self._apply_target_to_actual_fit_size()
        self._refresh_all_safe()

    def _apply_target_to_actual_fit_size(self):
        wh = self._compute_max_fit_wh_for_current()
        if not wh:
            return
        w_fit, h_fit = wh
        self._ratio_guard = True
        try:
            self.var_img_w.set(w_fit)
            self.var_img_h.set(h_fit)
            self._last_ratio = w_fit / max(1, h_fit)
        finally:
            self._ratio_guard = False

    # ---------- rotation change ----------
    def _on_rot_changed(self):
        try:
            rot = int(self.var_rot.get())
        except Exception:
            rot = 0

        prev_odd = (self._prev_rot % 180) == 90
        now_odd = (rot % 180) == 90
        if prev_odd != now_odd:
            self._ratio_guard = True
            try:
                w = int(self.var_img_w.get())
                h = int(self.var_img_h.get())
                self.var_img_w.set(h)
                self.var_img_h.set(w)
                self._last_ratio = self.var_img_w.get() / max(1, self.var_img_h.get())
            finally:
                self._ratio_guard = False

        self._prev_rot = rot
        if self.var_autofit.get() and (self.current_index >= 0):
            self._apply_target_to_actual_fit_size()

        # 重新居中图片
        self._recenter_after_rotation()
        self._refresh_all_safe()

    def _recenter_after_rotation(self):
        """旋转后自动居中图片"""
        if self.current_index < 0:
            return
        
        it = self.items[self.current_index]
        
        # 重新计算旋转后的图片尺寸
        src = it.orig
        rot = int(self.var_rot.get())
        if rot != 0:
            src = src.rotate(rot, expand=True, resample=Image.Resampling.BICUBIC)

        box_w = max(1, int(self.var_img_w.get()))
        box_h = max(1, int(self.var_img_h.get()))
        
        # 调整图片尺寸
        resized_img = src.resize((box_w, box_h), Image.Resampling.LANCZOS)
        
        # 计算居中位置
        disp_w = int(self.var_disp_w.get())
        disp_h = int(self.var_disp_h.get())
        
        it.img_x = (disp_w - box_w) // 2
        it.img_y = (disp_h - box_h) // 2

    def on_fit_max_frame(self):
        self.var_autofit.set(True)
        self._apply_target_to_actual_fit_size()
        self._refresh_all_safe()

    # ---------- import/list ----------
    def on_import_mixed(self):
        paths = filedialog.askopenfilenames(
            title="选择视频或图片",
            filetypes=[
                ("视频/图片", "*.mp4;*.avi;*.mov;*.mkv;*.wmv;*.webm;*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp"),
                ("视频", "*.mp4;*.avi;*.mov;*.mkv;*.wmv;*.webm"),
                ("图片", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp"),
                ("All", "*.*")
            ]
        )
        if not paths:
            return

        for p in paths:
            ext = os.path.splitext(p)[1].lower()
            if ext in ['.mp4', '.avi', '.mov', '.mkv', '.wmv', '.webm']:
                # 导入视频
                try:
                    cap = cv2.VideoCapture(p)
                    if not cap.isOpened():
                        raise RuntimeError("无法打开视频。")
                    
                    native_fps = float(cap.get(cv2.CAP_PROP_FPS) or 0.0)
                    total = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)
                    duration_s = 0.0
                    if total > 0 and native_fps > 0:
                        duration_s = total / native_fps
                    
                    cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
                    ok, frame = cap.read()
                    if not ok or frame is None:
                        raise RuntimeError("读取首帧失败。")
                    
                    frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    pil = Image.fromarray(frame).convert("RGB")
                    
                    # 创建一个ItemState来保存视频信息
                    st = ItemState(path=p, base_orig=pil.copy(), orig=pil.copy())
                    self.items.append(st)
                    self.listbox.insert(tk.END, f"[视频] {os.path.basename(p)}")
                    
                    cap.release()
                except Exception as e:
                    messagebox.showwarning("导入视频失败", f"{p}\n\n{e}")
            else:
                # 导入图片
                try:
                    img = Image.open(p).convert("RGB")
                    st = ItemState(path=p, base_orig=img.copy(), orig=img.copy())
                    self.items.append(st)
                    self.listbox.insert(tk.END, f"[图片] {os.path.basename(p)}")
                except Exception as e:
                    messagebox.showwarning("导入图片失败", f"{p}\n\n{e}")

        if self.current_index < 0 and self.items:
            self.current_index = 0
            self.listbox.selection_set(0)

        if self.var_autofit.get():
            self._apply_target_to_actual_fit_size()

        self.exit_bin_preview()
        self._exit_video_preview()
        self._refresh_all_safe()

    def on_remove(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        self.listbox.delete(idx)
        self.items.pop(idx)
        if not self.items:
            self.current_index = -1
            self.lbl_orig.configure(image="", text="")
            self.canvas.delete("all")
            self.var_fname_preview.set("")
            return
        self.current_index = min(idx, len(self.items) - 1)
        self.listbox.selection_set(self.current_index)

        if self.var_autofit.get():
            self._apply_target_to_actual_fit_size()

        self.exit_bin_preview()
        self._exit_video_preview()
        self._refresh_all_safe()
        
    def _update_output_name_by_current(self):
        if self.current_index < 0:
            return

        it = self.items[self.current_index]
        ext = os.path.splitext(it.path)[1].lower()
        idx = self.current_index + 1

        if ext in ['.mp4', '.avi', '.mov', '.mkv', '.wmv', '.webm']:
            name = f"media_{idx}_video.bin"
        else:
            name = f"media_{idx}_image.bin"

        self.var_video_name.set(name)
        
    def on_select(self, _evt=None):
        sel = self.listbox.curselection()
        if not sel:
            return
        self.current_index = sel[0]

        if self.var_autofit.get():
            self._apply_target_to_actual_fit_size()

        self.exit_bin_preview()
        self._exit_video_preview()
        self._refresh_all_safe()
        self._update_output_name_by_current()
    # ---------- out dir ----------
    def on_choose_out_dir(self):
        d = filedialog.askdirectory(title="选择导出目录")
        if d:
            self.var_out_dir.set(d)

    # ---------- reset ----------
    def on_reset_current(self):
        if self.current_index < 0:
            return
        it = self.items[self.current_index]
        it.orig = it.base_orig.copy()
        it.crop_box = None
        it.img_x = 0
        it.img_y = 0
        it._placed_once = False

        self._ratio_guard = True
        try:
            self.var_rot.set("0")
            self._prev_rot = 0
            self.var_autofit.set(True)
        finally:
            self._ratio_guard = False

        self.exit_bin_preview()
        self._exit_video_preview()

        self._apply_target_to_actual_fit_size()
        self._refresh_all_safe()

    # ---------- transforms ----------
    def _compute_transformed_image(self, it: ItemState) -> Image.Image:
        """
        ✅ 图片：FIT（等比缩放，不拉伸）
        - 先旋转
        - 再等比缩放塞进目标框 box_w x box_h
        - 贴到黑底居中（避免拉伸）
        """
        src = it.orig
        rot = int(self.var_rot.get())
        if rot != 0:
            src = src.rotate(rot, expand=True, resample=Image.Resampling.BICUBIC)
    
        box_w = max(1, int(self.var_img_w.get()))
        box_h = max(1, int(self.var_img_h.get()))

        sw, sh = src.width, src.height
        if sw <= 0 or sh <= 0:
            return Image.new("RGB", (box_w, box_h), (0, 0, 0))
    
        # 等比缩放比例：塞进目标框
        scale = min(box_w / sw, box_h / sh)
        new_w = max(1, int(round(sw * scale)))
        new_h = max(1, int(round(sh * scale)))

        resized_img = src.resize((new_w, new_h), Image.Resampling.LANCZOS)

        # 黑底目标框
        result = Image.new("RGB", (box_w, box_h), (0, 0, 0))
        paste_x = (box_w - new_w) // 2
        paste_y = (box_h - new_h) // 2
        result.paste(resized_img, (paste_x, paste_y))

        return result

    def _compute_transformed_image_video_fit(self, pil_rgb: Image.Image) -> Image.Image:
        """
        ✅ 视频：FIT（完整显示，不裁切，不补黑）
        - 先旋转
        - 再等比缩放"塞进目标框 box_w x box_h"
        - 不补黑：返回缩放后的真实画面尺寸
        """
        src = pil_rgb
        rot = int(self.var_rot.get())
        if rot != 0:
            src = src.rotate(rot, expand=True, resample=Image.Resampling.BICUBIC)

        box_w = max(1, int(self.var_img_w.get()))
        box_h = max(1, int(self.var_img_h.get()))

        sw, sh = src.width, src.height
        if sw <= 0 or sh <= 0:
            return Image.new("RGB", (1, 1), (0, 0, 0))

        # 计算缩放比例
        scale = min(box_w / sw, box_h / sh)
        new_w = max(1, int(round(sw * scale)))
        new_h = max(1, int(round(sh * scale)))
        
        # 缩放图像
        resized_img = src.resize((new_w, new_h), Image.Resampling.LANCZOS)
        
        # 创建目标尺寸的黑色背景
        result = Image.new("RGB", (box_w, box_h), (0, 0, 0))
        
        # 计算居中位置
        paste_x = (box_w - new_w) // 2
        paste_y = (box_h - new_h) // 2
        
        # 将缩放后的图像粘贴到中心
        result.paste(resized_img, (paste_x, paste_y))
        return result

    # ---------- refresh ----------
    def _refresh_all_safe(self):
        try:
            self._refresh_all()
        except Exception:
            pass

    def _refresh_all(self):
        self._refresh_orig_preview()
        self._refresh_conv_preview()

    def _refresh_orig_preview(self):
        if self.video_preview_active and self.video_preview_base is not None:
            img = self.video_preview_base
        elif self.current_index >= 0 and self.current_index < len(self.items):
            img = self.items[self.current_index].orig
        else:
            self.lbl_orig.configure(image="", text="")
            return

        w = max(1, self.lbl_orig.winfo_width() or 520)
        h = max(1, self.lbl_orig.winfo_height() or 520)
        scale = min(w / img.width, h / img.height, 1.0)
        pv = img.resize((max(1, int(img.width * scale)), max(1, int(img.height * scale))), Image.Resampling.LANCZOS)
        self.tk_orig_preview = ImageTk.PhotoImage(pv)
        self.lbl_orig.configure(image=self.tk_orig_preview, text="")

    def _frame_rect_on_canvas(self):
        cw = max(1, self.canvas.winfo_width())
        ch = max(1, self.canvas.winfo_height())
        disp_w = max(1, int(self.var_disp_w.get()))
        disp_h = max(1, int(self.var_disp_h.get()))
        margin = 30
        scale = min((cw - 2 * margin) / disp_w, (ch - 2 * margin) / disp_h)
        scale = max(0.5, min(scale, 3.0))
        fw = int(disp_w * scale)
        fh = int(disp_h * scale)
        fx0 = (cw - fw) // 2
        fy0 = (ch - fh) // 2
        fx1 = fx0 + fw
        fy1 = fy0 + fh
        return fx0, fy0, fx1, fy1, scale

    def _compute_export_params(self, it_like):
        """
        严格裁剪规则：
       - 只导出【屏幕内可见区域】
        - 完全在屏外 → 输出 0x0
        - dx / dy / cw / ch 永远合法
        """

        disp_w = int(self.var_disp_w.get())
        disp_h = int(self.var_disp_h.get())

        if self.conv_img_pil is None:
            return 0, 0, 0, 0, 0, 0

        iw = self.conv_img_pil.width
        ih = self.conv_img_pil.height

        # 图片在“显示屏坐标系”中的矩形
        img_x0 = it_like.img_x
        img_y0 = it_like.img_y
        img_x1 = img_x0 + iw
        img_y1 = img_y0 + ih

        # 显示屏矩形
        scr_x0 = 0
        scr_y0 = 0
        scr_x1 = disp_w
        scr_y1 = disp_h

        # 求交集
        vis_x0 = max(img_x0, scr_x0)
        vis_y0 = max(img_y0, scr_y0)
        vis_x1 = min(img_x1, scr_x1)
        vis_y1 = min(img_y1, scr_y1)

        # ❌ 完全不可见
        if vis_x1 <= vis_x0 or vis_y1 <= vis_y0:
            return 0, 0, 0, 0, 0, 0

        # 输出区域尺寸
        cw = vis_x1 - vis_x0
        ch = vis_y1 - vis_y0

        # 在显示屏内的起点
        dx = vis_x0
        dy = vis_y0

        # 在图片内的裁剪起点
        sx = vis_x0 - img_x0
        sy = vis_y0 - img_y0

        # 最终安全 clamp（双保险）
        dx = max(0, min(dx, disp_w))
        dy = max(0, min(dy, disp_h))

        sx = max(0, min(sx, iw))
        sy = max(0, min(sy, ih))

        cw = max(0, min(cw, iw - sx, disp_w - dx))
        ch = max(0, min(ch, ih - sy, disp_h - dy))

        return int(dx), int(dy), int(sx), int(sy), int(cw), int(ch)


    def _make_preview_with_fade(self, pil_img_rgb: Image.Image, it_like, fx0, fy0, fx1, fy1, scale) -> Image.Image:
        prev_w = max(1, int(pil_img_rgb.width * scale))
        prev_h = max(1, int(pil_img_rgb.height * scale))
        img_rgba = pil_img_rgb.convert("RGBA").resize((prev_w, prev_h), Image.Resampling.NEAREST)

        img_cx = fx0 + int(it_like.img_x * scale)
        img_cy = fy0 + int(it_like.img_y * scale)

        img_x0 = img_cx
        img_y0 = img_cy
        img_x1 = img_cx + prev_w
        img_y1 = img_cy + prev_h

        ov_x0 = max(img_x0, fx0)
        ov_y0 = max(img_y0, fy0)
        ov_x1 = min(img_x1, fx1)
        ov_y1 = min(img_y1, fy1)

        fade_alpha = 30
        mask = Image.new("L", (prev_w, prev_h), color=fade_alpha)

        loc_x0 = int(round(ov_x0 - img_x0))
        loc_y0 = int(round(ov_y0 - img_y0))
        loc_x1 = int(round(ov_x1 - img_x0))
        loc_y1 = int(round(ov_y1 - img_y0))

        loc_x0 = clamp(loc_x0, 0, prev_w)
        loc_y0 = clamp(loc_y0, 0, prev_h)
        loc_x1 = clamp(loc_x1, 0, prev_w)
        loc_y1 = clamp(loc_y1, 0, prev_h)

        if loc_x1 > loc_x0 and loc_y1 > loc_y0:
            opaque = Image.new("L", (loc_x1 - loc_x0, loc_y1 - loc_y0), color=255)
            mask.paste(opaque, (loc_x0, loc_y0))

        r, g, b, _a = img_rgba.split()
        return Image.merge("RGBA", (r, g, b, mask))

    def _refresh_conv_preview(self):
        self.canvas.delete("all")

        fx0, fy0, fx1, fy1, scale = self._frame_rect_on_canvas()
        self.canvas.create_rectangle(fx0, fy0, fx1, fy1, outline="#E6E6E6", width=2)
        self.canvas.create_text(
            fx0 + 4, fy0 - 12,
            text=f"{self.var_disp_w.get()}x{self.var_disp_h.get()}",
            fill="#E6E6E6", anchor="w", font=("Segoe UI", 10)
        )

        # ✅ 视频首帧预览：FIT（不裁切、不补黑）
        if self.video_preview_active and self.video_preview_base is not None:
            preview = self._compute_transformed_image_video_fit(self.video_preview_base)

            disp_w = int(self.var_disp_w.get())
            disp_h = int(self.var_disp_h.get())
            x = (disp_w - preview.width) // 2
            y = (disp_h - preview.height) // 2

            class _Tmp:
                pass
            t = _Tmp()
            t.img_x = x
            t.img_y = y

            preview_rgba = self._make_preview_with_fade(preview, t, fx0, fy0, fx1, fy1, scale)
            self.tk_bin_preview = ImageTk.PhotoImage(preview_rgba)
            self.canvas.create_image(fx0 + int(t.img_x * scale), fy0 + int(t.img_y * scale),
                                     image=self.tk_bin_preview, anchor="nw")

            self.conv_img_pil = preview
            dx, dy, sx, sy, cw, ch = self._compute_export_params(t)

            # 计算总帧数
            fps_out = self.DEFAULT_VIDEO_FPS
            fps_out = max(0.001, fps_out)

            dur_s = self.video_state.duration_s if self.video_state else 0.0
            frame_count = max(1, int(math.ceil(dur_s * fps_out)))

            # 统一设置为输出文件名
            self.var_fname_preview.set(self.var_video_name.get())
            
            # 显示正确的宽高信息
            box_w = max(1, int(self.var_img_w.get()))
            box_h = max(1, int(self.var_img_h.get()))
            self.canvas.create_text(
                fx0 + 6, fy1 + 14,
                text=f"[视频首帧] FIT完整显示 | 目标框={box_w}x{box_h} | 起始点({dx},{dy}) | 输出={cw}x{ch} | 总帧将按时长*fps计算",
                fill="#CFCFCF", anchor="w", font=("Segoe UI", 9)
            )
            return

        # ✅ BIN 预览（旧格式）
        if self.bin_preview_active and self.bin_preview_img is not None:
            class _Tmp:
                pass
            tmp = _Tmp()
            tmp.img_x = int(self.bin_preview_dx)
            tmp.img_y = int(self.bin_preview_dy)

            preview_rgba = self._make_preview_with_fade(self.bin_preview_img, tmp, fx0, fy0, fx1, fy1, scale)
            self.tk_bin_preview = ImageTk.PhotoImage(preview_rgba)
            self.canvas.create_image(fx0 + int(tmp.img_x * scale), fy0 + int(tmp.img_y * scale),
                                     image=self.tk_bin_preview, anchor="nw")

            self.canvas.create_text(
                fx0 + 6, fy1 + 14,
                text=f"[BIN预览] 起始点({tmp.img_x},{tmp.img_y}) | 尺寸={self.bin_preview_img.width}x{self.bin_preview_img.height} | 双击退出",
                fill="#CFCFCF", anchor="w", font=("Segoe UI", 9)
            )
            return

        # 图片预览（保持原逻辑）
        if self.current_index < 0 or self.current_index >= len(self.items):
            # 统一设置为输出文件名
            self.var_fname_preview.set(self.var_video_name.get())
            return

        it = self.items[self.current_index]
        self.conv_img_pil = self._compute_transformed_image(it)

        if not it._placed_once:
            it._placed_once = True
            it.img_x = (int(self.var_disp_w.get()) - self.conv_img_pil.width) // 2
            it.img_y = (int(self.var_disp_h.get()) - self.conv_img_pil.height) // 2

        if self.var_snap.get():
            it.img_x, it.img_y = self._apply_snap_image(it)

        preview_rgba = self._make_preview_with_fade(self.conv_img_pil, it, fx0, fy0, fx1, fy1, scale)
        self.tk_conv_img = ImageTk.PhotoImage(preview_rgba)
        self.canvas.create_image(fx0 + int(it.img_x * scale), fy0 + int(it.img_y * scale),
                                 image=self.tk_conv_img, anchor="nw")

        dx, dy, sx, sy, cw, ch = self._compute_export_params(it)
        # 统一设置为输出文件名
        self.var_fname_preview.set(self.var_video_name.get())

        # 显示正确的宽高信息
        box_w = max(1, int(self.var_img_w.get()))
        box_h = max(1, int(self.var_img_h.get()))
        self.canvas.create_text(
            fx0 + 6, fy1 + 14,
            text=f"[图片] 目标框={box_w}x{box_h} | 起始点({dx},{dy}) | 输出={cw}x{ch}",
            fill="#CFCFCF", anchor="w", font=("Segoe UI", 9)
        )

    def _apply_snap_image(self, it: ItemState):
        disp_w = int(self.var_disp_w.get())
        disp_h = int(self.var_disp_h.get())
        iw = self.conv_img_pil.width if self.conv_img_pil else 0
        ih = self.conv_img_pil.height if self.conv_img_pil else 0
        x, y = it.img_x, it.img_y
        thresh = 6
        if abs(x - 0) <= thresh:
            x = 0
        if abs(y - 0) <= thresh:
            y = 0
        if abs((x + iw) - disp_w) <= thresh:
            x = disp_w - iw
        if abs((y + ih) - disp_h) <= thresh:
            y = disp_h - ih
        it.img_x, it.img_y = x, y
        return x, y

    # ---------- mouse ----------
    def on_left_down(self, e):
        if self.video_preview_active or self.bin_preview_active:
            return
        if self.current_index < 0:
            return
        it = self.items[self.current_index]
        fx0, fy0, fx1, fy1, _scale = self._frame_rect_on_canvas()
        if not (fx0 <= e.x <= fx1 and fy0 <= e.y <= fy1):
            return
        self._dragging = True
        self._drag_start = (e.x, e.y)
        self._img_start = (it.img_x, it.img_y)

    def on_left_move(self, e):
        if self.video_preview_active or self.bin_preview_active:
            return
        if not self._dragging or self.current_index < 0:
            return
        it = self.items[self.current_index]
        fx0, fy0, fx1, fy1, scale = self._frame_rect_on_canvas()
        dx = int(round((e.x - self._drag_start[0]) / scale))
        dy = int(round((e.y - self._drag_start[1]) / scale))
        it.img_x = self._img_start[0] + dx
        it.img_y = self._img_start[1] + dy
        self._refresh_conv_preview()

    def on_left_up(self, _e):
        self._dragging = False

    def on_right_down(self, e):
        if self.video_preview_active or self.bin_preview_active:
            return
        if self.current_index < 0:
            return
        fx0, fy0, fx1, fy1, _scale = self._frame_rect_on_canvas()
        if not (fx0 <= e.x <= fx1 and fy0 <= e.y <= fy1):
            return
        self._crop_start = (e.x, e.y)
        self._clear_crop_rect()
        self._crop_rect_id = self.canvas.create_rectangle(e.x, e.y, e.x, e.y, outline="#00FFAA", width=2)

    def on_right_move(self, e):
        if self.video_preview_active or self.bin_preview_active:
            return
        if not self._crop_start or self._crop_rect_id is None:
            return
        x0, y0 = self._crop_start
        self.canvas.coords(self._crop_rect_id, x0, y0, e.x, e.y)

    def on_right_up(self, e):
        if self.video_preview_active or self.bin_preview_active:
            return
        if self._crop_start:
            x0, y0 = self._crop_start
            x1, y1 = e.x, e.y
            
            # 计算实际裁剪区域
            fx0, fy0, fx1, fy1, scale = self._frame_rect_on_canvas()
            it = self.items[self.current_index]
            
            # 计算相对于图片的位置
            img_x_start = int((x0 - fx0) / scale - it.img_x)
            img_y_start = int((y0 - fy0) / scale - it.img_y)
            img_x_end = int((x1 - fx0) / scale - it.img_x)
            img_y_end = int((y1 - fy0) / scale - it.img_y)
            
            # 确保坐标顺序正确
            if img_x_start > img_x_end:
                img_x_start, img_x_end = img_x_end, img_x_start
            if img_y_start > img_y_end:
                img_y_start, img_y_end = img_y_end, img_y_start
            
            # 限制在图片范围内
            img_width = self.conv_img_pil.width
            img_height = self.conv_img_pil.height
            img_x_start = max(0, min(img_x_start, img_width))
            img_y_start = max(0, min(img_y_start, img_height))
            img_x_end = max(0, min(img_x_end, img_width))
            img_y_end = max(0, min(img_y_end, img_height))
            
            # 设置裁剪框
            it.crop_box = (img_x_start, img_y_start, img_x_end, img_y_end)
            
            # 更新原图
            original_size = it.base_orig.size
            scale_factor_x = original_size[0] / img_width
            scale_factor_y = original_size[1] / img_height
            
            orig_crop_box = (
                int(img_x_start * scale_factor_x),
                int(img_y_start * scale_factor_y),
                int(img_x_end * scale_factor_x),
                int(img_y_end * scale_factor_y)
            )
            
            # 裁剪原始图片
            it.orig = it.base_orig.crop(orig_crop_box)

            # ✅ 裁切后：按裁切后的比例自适应最大边框（不拉伸）
            if self.var_autofit.get():
                self._apply_target_to_actual_fit_size()

            # 重置位置（让它刷新后重新居中）
            it.img_x = 0
            it.img_y = 0
            it._placed_once = False  # 可选：强制走一次"首次放置"逻辑

            # 清除裁剪框
            self._clear_crop_rect()

            # 更新预览
            self._refresh_all_safe()

        else:
            self._clear_crop_rect()
            self._refresh_all_safe()

    def _clear_crop_rect(self):
        if self._crop_rect_id is not None:
            try:
                self.canvas.delete(self._crop_rect_id)
            except Exception:
                pass
        self._crop_rect_id = None

    # ---------- rgb565 ----------
    def _build_rgb565_bytes(self, pil_img: Image.Image) -> bytes:
        endian = self.var_endian.get()
        pix = pil_img.convert("RGB").load()
        out = bytearray()
        for y in range(pil_img.height):
            for x in range(pil_img.width):
                r, g, b = pix[x, y]
                v = rgb888_to_rgb565(r, g, b)
                if endian == "little":
                    out.append(v & 0xFF)
                    out.append((v >> 8) & 0xFF)
                else:
                    out.append((v >> 8) & 0xFF)
                    out.append(v & 0xFF)
        return bytes(out)

    def _decode_rgb565_to_image(self, data: bytes, w: int, h: int) -> Image.Image:
        endian = self.var_endian.get()
        img = Image.new("RGB", (w, h))
        px = img.load()
        i = 0
        for y in range(h):
            for x in range(w):
                if i + 1 >= len(data):
                    break
                if endian == "little":
                    v = data[i] | (data[i + 1] << 8)
                else:
                    v = (data[i] << 8) | data[i + 1]
                i += 2
                r = (v >> 11) & 0x1F
                g = (v >> 5) & 0x3F
                b = v & 0x1F
                rr = (r << 3) | (r >> 2)
                gg = (g << 2) | (g >> 4)
                bb = (b << 3) | (b >> 2)
                px[x, y] = (rr, gg, bb)
        return img

    # ---------- export (image old-format) ----------
    def _build_output_filename(self, it: ItemState, dx, dy, cw, ch):
        dur = clamp(int(self.var_ms.get()), 0, 65535)
        base = safe_name(os.path.splitext(os.path.basename(it.path))[0])
        return f"{dx}_{dy}_{cw}_{ch}_{dur}_{base}.bin"

    def _get_current_export_blob(self):
        if self.current_index < 0:
            raise RuntimeError("没有选择图片/视频。")

        it = self.items[self.current_index]
        ext = os.path.splitext(it.path)[1].lower()

        # 视频仍走你原来的 V565（不动）
        if ext in ['.mp4', '.avi', '.mov', '.mkv', '.wmv', '.webm']:
            raise RuntimeError("当前是视频，请用视频导出逻辑。")

        # 图片：先得到目标框图（box_w/box_h）
        img_box = self._compute_transformed_image(it)
        self.conv_img_pil = img_box

        # 把目标框图按拖动位置贴到整屏（disp_w/disp_h）
        full = self._compose_fullscreen(it, img_box)

        # 生成头：注意 MCU 端会强制全屏，但 play_ms 仍会用到
        play_ms = clamp(int(self.var_ms.get()), 0, 65535)
        hdr512 = self._build_image_header_512(
            0, 0,
            int(self.var_disp_w.get()),
            int(self.var_disp_h.get()),
            play_ms
        )

        data = self._build_rgb565_bytes(full)
        fname = self.var_video_name.get().strip() or f"media_{self.current_index + 1}_image.bin"
        if not fname.lower().endswith(".bin"):
            fname += ".bin"
        return {"fname": safe_name(fname), "data": hdr512 + data}


    def _ensure_out_dir(self) -> str:
        d = (self.var_out_dir.get() or "").strip()
        if d and os.path.isdir(d):
            return d
        d = filedialog.askdirectory(title="选择导出目录")
        if not d:
            raise RuntimeError("未选择导出目录。")
        self.var_out_dir.set(d)
        return d

    def on_export_current(self):
        try:
            # 根据当前项目类型决定是导出图片还是视频
            if self.current_index < 0:
                raise RuntimeError("没有选择项目。")
                
            it = self.items[self.current_index]
            ext = os.path.splitext(it.path)[1].lower()
            
            if ext in ['.mp4', '.avi', '.mov', '.mkv', '.wmv', '.webm']:
                # 导出视频 - 从items中获取视频信息
                self._export_current_video(it)
            else:
                # 导出图片
                blob = self._get_current_export_blob()
                out_dir = self._ensure_out_dir()
                
                # 修改图片导出文件名为 media_数字_image.bin
                item_index = self.items.index(it)
                image_name = f"media_{item_index + 1}_image.bin"
                fpath = os.path.join(out_dir, image_name)
                
                with open(fpath, "wb") as f:
                    f.write(blob["data"])
                messagebox.showinfo("导出成功", f"已导出：\n{fpath}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def _calculate_adaptive_fps(self, width: int, height: int) -> float:
        """根据分辨率计算自适应FPS"""
        # 基准分辨率: 180x320 -> 10fps
        baseline_width = 180
        baseline_height = 320
        baseline_fps = 10.0
        
        # 计算相对面积变化
        baseline_area = baseline_width * baseline_height
        current_area = width * height
        
        # FPS与面积成反比，但限制在合理范围内
        adaptive_fps = baseline_fps * (baseline_area / current_area)
        
        # 限制在合理范围内 (1-60fps)
        adaptive_fps = max(1.0, min(60.0, adaptive_fps))
        
        return adaptive_fps

    def _export_current_video(self, video_item: ItemState):
        """导出当前视频项"""
        try:
            out_dir = self._ensure_out_dir()
            
            # 获取视频信息
            cap = cv2.VideoCapture(video_item.path)
            if not cap.isOpened():
                raise RuntimeError("无法打开视频。")
            
            native_fps = float(cap.get(cv2.CAP_PROP_FPS) or 0.0)
            total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)
            duration_s = 0.0
            if total_frames > 0 and native_fps > 0:
                duration_s = total_frames / native_fps
            
            # 获取首帧用于计算尺寸
            cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
            ok_frame, frame = cap.read()
            if not ok_frame or frame is None:
                raise RuntimeError("读取首帧失败。")
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            pil_first = Image.fromarray(frame).convert("RGB")
            
            # 计算视频处理后的尺寸
            conv = self._compute_transformed_image_video_fit(pil_first)
            
            # 使用自适应FPS
            target_width = int(self.var_img_w.get())
            target_height = int(self.var_img_h.get())
            fps_out = self._calculate_adaptive_fps(target_width, target_height)
            fps_out = max(0.001, fps_out)
            frame_count = max(1, int(math.ceil(duration_s * fps_out)))
            
            # 构建视频文件名
            item_index = self.items.index(video_item)
            video_num = item_index + 1
            video_name = (self.var_video_name.get() or f"media_{video_num}_video.bin").strip()
            if not video_name.lower().endswith(".bin"):
                video_name += ".bin"
            out_path = os.path.join(out_dir, safe_name(video_name))
            
            # 构建视频头
            with open(out_path, "wb") as f:
                f.write(self._build_video_header_sector(fps_out=fps_out, frame_count=frame_count))
                
                # 逐帧处理
                for i in range(frame_count):
                    t_s = i / fps_out
                    self._seek_frame_by_time(cap, t_s, native_fps)
                    
                    ok_frame, frame = cap.read()
                    if not ok_frame or frame is None:
                        # 空帧
                        empty = bytearray()
                        empty += self._pack_u16(0) + self._pack_u16(0)
                        empty += self._pack_u16(0) + self._pack_u16(0)
                        empty += self._pack_u16(0)  # dur_ms = 0
                        empty += b"\x00\x00"
                        empty += self._pack_u32(0)
                        f.write(self._pad_to_sector(bytes(empty)))
                        continue

                    frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    pil = Image.fromarray(frame).convert("RGB")
                    f.write(self._build_one_frame_record_fit(pil, dur_ms=0))
            
            cap.release()
            messagebox.showinfo("导出成功", f"已导出视频：\n{out_path}\nFPS: {fps_out:.3f}, 总帧数: {frame_count}, 时长: {duration_s:.2f}s")
        except Exception as e:
            raise e

    def _build_one_frame_record_fit(self, pil_rgb: Image.Image, dur_ms: int) -> bytes:
        """
        每帧 record：
          16字节帧头：
            u16 cw,ch
            u16 dx,dy
            u16 dur_ms
            u16 reserved
            u32 data_len
          + data_len 字节RGB565
          + pad到512对齐
        """
        conv = self._compute_transformed_image_video_fit(pil_rgb)

        disp_w = int(self.var_disp_w.get())
        disp_h = int(self.var_disp_h.get())
        img_x = (disp_w - conv.width) // 2
        img_y = (disp_h - conv.height) // 2

        class _Tmp:
            pass
        it_like = _Tmp()
        it_like.img_x = img_x
        it_like.img_y = img_y

        self.conv_img_pil = conv
        dx, dy, sx, sy, cw, ch = self._compute_export_params(it_like)

        if cw <= 0 or ch <= 0:
            rgb565 = b""
            cw = ch = 0
            dx = dy = 0
        else:
            crop = conv.crop((sx, sy, sx + cw, sy + ch))
            rgb565 = self._build_rgb565_bytes(crop)

        dur = clamp(int(dur_ms), 0, 65535)

        rec = bytearray()
        rec += self._pack_u16(cw) + self._pack_u16(ch)
        rec += self._pack_u16(dx) + self._pack_u16(dy)
        rec += self._pack_u16(dur)
        rec += b"\x00\x00"
        rec += self._pack_u32(len(rgb565))
        rec += rgb565

        return self._pad_to_sector(bytes(rec))

    def on_export_all(self):
        if not self.items:
            return
        try:
            out_dir = self._ensure_out_dir()
        except Exception as e:
            messagebox.showerror("导出失败", str(e))
            return

        ok = 0
        fails = []
        for idx, it in enumerate(self.items):
            try:
                ext = os.path.splitext(it.path)[1].lower()
                if ext in ['.mp4', '.avi', '.mov', '.mkv', '.wmv', '.webm']:
                    # 处理视频
                    video_cap = cv2.VideoCapture(it.path)
                    if not video_cap.isOpened():
                        raise RuntimeError("无法打开视频。")
                    
                    native_fps = float(video_cap.get(cv2.CAP_PROP_FPS) or 0.0)
                    total_frames = int(video_cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)
                    duration_s = 0.0
                    if total_frames > 0 and native_fps > 0:
                        duration_s = total_frames / native_fps
                    
                    # 获取视频首帧用于获取尺寸信息
                    video_cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
                    ok_frame, frame = video_cap.read()
                    if not ok_frame or frame is None:
                        raise RuntimeError("读取首帧失败。")
                    frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    pil_first = Image.fromarray(frame).convert("RGB")
                    
                    # 计算视频处理后的尺寸
                    conv = self._compute_transformed_image_video_fit(pil_first)
                    
                    # 使用自适应FPS
                    target_width = int(self.var_img_w.get())
                    target_height = int(self.var_img_h.get())
                    fps_out = self._calculate_adaptive_fps(target_width, target_height)
                    fps_out = max(0.001, fps_out)
                    frame_count = max(1, int(math.ceil(duration_s * fps_out)))
                    
                    # 构建视频文件名
                    video_num = idx + 1
                    video_name = f"media_{video_num}_video.bin"
                    out_path = os.path.join(out_dir, safe_name(video_name))
                    
                    # 构建视频头
                    with open(out_path, "wb") as f:
                        f.write(self._build_video_header_sector(fps_out=fps_out, frame_count=frame_count))
                        
                        # 逐帧处理
                        for i in range(frame_count):
                            t_s = i / fps_out
                            self._seek_frame_by_time(video_cap, t_s, native_fps)
                            
                            ok_frame, frame = video_cap.read()
                            if not ok_frame or frame is None:
                                # 空帧
                                empty = bytearray()
                                empty += self._pack_u16(0) + self._pack_u16(0)
                                empty += self._pack_u16(0) + self._pack_u16(0)
                                empty += self._pack_u16(0)  # dur_ms = 0
                                empty += b"\x00\x00"
                                empty += self._pack_u32(0)
                                f.write(self._pad_to_sector(bytes(empty)))
                                continue

                            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                            pil = Image.fromarray(frame).convert("RGB")
                            f.write(self._build_one_frame_record_fit(pil, dur_ms=0))
                    
                    video_cap.release()
                    ok += 1
                else:
                    # 处理图片：导出 512头 + 整屏
                    img_box = self._compute_transformed_image(it)
                    self.conv_img_pil = img_box
                    full = self._compose_fullscreen(it, img_box)

                    play_ms = clamp(int(self.var_ms.get()), 0, 65535)
                    hdr512 = self._build_image_header_512(
                        0, 0,
                        int(self.var_disp_w.get()),
                        int(self.var_disp_h.get()),
                        play_ms
                    )
                    data = hdr512 + self._build_rgb565_bytes(full)

                    image_name = f"media_{idx + 1}_image.bin"
                    with open(os.path.join(out_dir, image_name), "wb") as f:
                        f.write(data)
                    ok += 1

            except Exception as e:
                fails.append((os.path.basename(it.path), str(e)))

        msg = f"完成：成功 {ok} / {len(self.items)}"
        if fails:
            msg += "\n\n失败列表：\n" + "\n".join([f"- {n}: {err}" for n, err in fails[:12]])
            if len(fails) > 12:
                msg += f"\n... 还有 {len(fails)-12} 条"
        messagebox.showinfo("批量导出结果", msg)

    # ---------- preview bin (old-format) ----------
    def on_preview_bin(self):
        path = filedialog.askopenfilename(
            title="选择bin文件",
            filetypes=[("BIN", "*.bin"), ("All", "*.*")]
        )
        if not path:
            return

        fname = os.path.basename(path)
        try:
            with open(path, "rb") as f:
                header = f.read(512)
                if len(header) == 512:
                    x = struct.unpack("<H", header[0:2])[0]
                    y = struct.unpack("<H", header[2:4])[0]
                    w = struct.unpack("<H", header[4:6])[0]
                    h = struct.unpack("<H", header[6:8])[0]
                    play_ms = struct.unpack("<H", header[8:10])[0]

                    # 合理性判断：w/h 必须 >0 且后面数据足够
                    need = w * h * 2
                    rest = f.read(need)
                    if w > 0 and h > 0 and len(rest) == need:
                        img = self._decode_rgb565_to_image(rest, w, h)
                        self.bin_preview_active = True
                        self.bin_preview_img = img
                        self.bin_preview_dx = 0
                        self.bin_preview_dy = 0
                        self.bin_preview_path = path
                        self._exit_video_preview()
                        self.var_fname_preview.set(f"[512头BIN] {fname} | {w}x{h} | play_ms={play_ms}")
                        self._refresh_all_safe()
                        return
        except Exception:
            pass
        m = re.match(r"^(\d+)_(\d+)_(\d+)_(\d+)_(\d+)_.*\.bin$", fname, re.IGNORECASE)
        if not m:
            # 尝试识别是否是V565格式
            try:
                with open(path, "rb") as f:
                    header_data = f.read(512)
                    if len(header_data) < 512:
                        messagebox.showerror("解析失败", "bin文件太小")
                        return
                    magic = header_data[0:4].decode('ascii', errors='ignore')
                    if magic == "V565":
                        # 这是一个视频bin文件，提取第一帧
                        fps_q8 = struct.unpack("<H", header_data[8:10])[0] / 256.0
                        frame_count = struct.unpack("<H", header_data[10:12])[0]
                        box_w = struct.unpack("<H", header_data[16:18])[0]
                        box_h = struct.unpack("<H", header_data[18:20])[0]
                        
                        # 读取第一帧数据
                        frame_offset = 512
                        f.seek(frame_offset)
                        frame_header = f.read(16)
                        if len(frame_header) < 16:
                            messagebox.showerror("解析失败", "无法读取第一帧数据")
                            return
                        
                        cw = struct.unpack("<H", frame_header[0:2])[0]
                        ch = struct.unpack("<H", frame_header[2:4])[0]
                        data_len = struct.unpack("<I", frame_header[12:16])[0]
                        
                        if cw > 0 and ch > 0:
                            rgb565_data = f.read(data_len)
                            if len(rgb565_data) < data_len:
                                messagebox.showerror("解析失败", "RGB565数据不完整")
                                return
                            img = self._decode_rgb565_to_image(rgb565_data, cw, ch)
                            
                            # 居中放置到目标框中
                            full_img = Image.new("RGB", (box_w, box_h), (0, 0, 0))
                            paste_x = (box_w - cw) // 2
                            paste_y = (box_h - ch) // 2
                            full_img.paste(img, (paste_x, paste_y))
                            
                            self.bin_preview_active = True
                            self.bin_preview_img = full_img
                            self.bin_preview_dx = (box_w - cw) // 2
                            self.bin_preview_dy = (box_h - ch) // 2
                            self.bin_preview_path = path
                            
                            self._exit_video_preview()
                            self.var_fname_preview.set(f"[视频BIN预览] {fname} | FPS:{fps_q8:.2f} | 帧数:{frame_count}")
                            self._refresh_all_safe()
                        else:
                            messagebox.showerror("解析失败", "视频BIN文件的第一帧为空")
                        return
                    else:
                        messagebox.showerror("解析失败", "bin文件不符合：dx_dy_w_h_time_xxx.bin（至少要有5段数字）或V565格式")
                        return
            except Exception as e:
                messagebox.showerror("解析失败", f"无法解析BIN文件: {str(e)}")
                return

        dx = int(m.group(1))
        dy = int(m.group(2))
        w = int(m.group(3))
        h = int(m.group(4))
        if w <= 0 or h <= 0:
            messagebox.showerror("解析失败", f"宽高非法：{w}x{h}")
            return

        try:
            with open(path, "rb") as f:
                data = f.read()
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return

        need = w * h * 2
        if len(data) < need:
            messagebox.showerror("bin长度不对", f"需要至少 {need} 字节，但文件只有 {len(data)} 字节")
            return

        img = self._decode_rgb565_to_image(data[:need], w, h)

        self.bin_preview_active = True
        self.bin_preview_img = img
        self.bin_preview_dx = dx
        self.bin_preview_dy = dy
        self.bin_preview_path = path

        self._exit_video_preview()
        self.var_fname_preview.set(fname)
        self._refresh_all_safe()

    def exit_bin_preview(self, _evt=None):
        if not self.bin_preview_active:
            return
        self.bin_preview_active = False
        self.bin_preview_img = None
        self.bin_preview_dx = 0
        self.bin_preview_dy = 0
        self.bin_preview_path = ""
        self.tk_bin_preview = None
        self._refresh_all_safe()

    # =========================
    # VIDEO import / preview / export (V565 container bin)
    # =========================
    def on_import_video(self):
        path = filedialog.askopenfilename(
            title="选择视频文件",
            filetypes=[("Video", "*.mp4;*.avi;*.mov;*.mkv;*.wmv;*.webm"), ("All", "*.*")]
        )
        if not path:
            return

        try:
            try:
                if self.video_state and self.video_state.cap:
                    self.video_state.cap.release()
            except Exception:
                pass

            cap = cv2.VideoCapture(path)
            if not cap.isOpened():
                raise RuntimeError("无法打开视频。")

            native_fps = float(cap.get(cv2.CAP_PROP_FPS) or 0.0)
            total = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)

            duration_s = 0.0
            if total > 0 and native_fps > 0:
                duration_s = total / native_fps

            cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
            ok, frame = cap.read()
            if not ok or frame is None:
                raise RuntimeError("读取首帧失败。")

            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            pil = Image.fromarray(frame).convert("RGB")

            self.video_state = VideoState(
                path=path,
                cap=cap,
                native_fps=native_fps,
                total_frames=total,
                duration_s=duration_s,
                first_frame_rgb=pil
            )

            self.exit_bin_preview()
            self._exit_video_preview()
            self.on_preview_video_first()

            messagebox.showinfo(
                "导入视频成功",
                f"文件：{os.path.basename(path)}\nFPS(原)：{native_fps:.3f}\n总帧(原)：{total if total > 0 else '未知'}\n"
                f"时长：{duration_s:.3f}s\n模式：FIT完整显示（不裁切、不补黑；bin里不写黑边）"
            )
        except Exception as e:
            messagebox.showerror("导入视频失败", str(e))

    def on_preview_video_first(self):
        if not self.video_state:
            messagebox.showwarning("提示", "请先导入视频。")
            return

        self.exit_bin_preview()
        self.video_preview_active = True
        self.video_preview_base = self.video_state.first_frame_rgb
        self.video_preview_path = self.video_state.path
        self._refresh_all_safe()

    def _exit_video_preview(self):
        self.video_preview_active = False
        self.video_preview_base = None
        self.video_preview_path = ""

    # --------- V565 bin packing ----------
    def _pack_u16(self, v: int) -> bytes:
        return struct.pack("<H", int(v) & 0xFFFF)

    def _pack_u32(self, v: int) -> bytes:
        return struct.pack("<I", int(v) & 0xFFFFFFFF)
    def _build_image_header_512(self, x: int, y: int, w: int, h: int, play_ms: int) -> bytes:
        """
        对应 MCU: parse_header_512()
        offset:
          0..1  x (le16)
          2..3  y (le16)
          4..5  w (le16)
          6..7  h (le16)
          8..9  play_ms (le16)
          其余填0到512
        """
        x = clamp(int(x), 0, 65535)
        y = clamp(int(y), 0, 65535)
        w = clamp(int(w), 0, 65535)
        h = clamp(int(h), 0, 65535)
        play_ms = clamp(int(play_ms), 0, 65535)

        hdr = struct.pack("<HHHHH", x, y, w, h, play_ms)
        if len(hdr) > self.SECTOR_SIZE:
            raise RuntimeError("header pack overflow")
        hdr += b"\x00" * (self.SECTOR_SIZE - len(hdr))
        return hdr

    def _compose_fullscreen(self, it_like, img_box: Image.Image) -> Image.Image:
        """
        把 box 图(目标框大小 var_img_w/h) 按 it_like.img_x/img_y 贴到整屏 disp_w/h 上。
        允许负坐标、超出屏幕：自动裁剪。
        """
        disp_w = max(1, int(self.var_disp_w.get()))
        disp_h = max(1, int(self.var_disp_h.get()))
        screen = Image.new("RGB", (disp_w, disp_h), (0, 0, 0))

        sx0 = 0
        sy0 = 0
        dx0 = int(it_like.img_x)
        dy0 = int(it_like.img_y)
        dx1 = dx0 + img_box.width
        dy1 = dy0 + img_box.height

        # 与屏幕求交
        vis_x0 = max(0, dx0)
        vis_y0 = max(0, dy0)
        vis_x1 = min(disp_w, dx1)
        vis_y1 = min(disp_h, dy1)
        if vis_x1 <= vis_x0 or vis_y1 <= vis_y0:
            return screen  # 完全在屏外 -> 全黑屏

        # 对应到源图裁剪
        sx0 = vis_x0 - dx0
        sy0 = vis_y0 - dy0
        sx1 = sx0 + (vis_x1 - vis_x0)
        sy1 = sy0 + (vis_y1 - vis_y0)

        patch = img_box.crop((sx0, sy0, sx1, sy1))
        screen.paste(patch, (vis_x0, vis_y0))
        return screen
    def _pad_to_sector(self, b: bytes) -> bytes:
        pad = (-len(b)) % self.SECTOR_SIZE
        if pad:
            b += b"\x00" * pad
        return b

    def _build_video_header_sector(self, fps_out: float, frame_count: int) -> bytes:
        """
        V565 头协议（512字节）：
        0..3   "V565"
        4..5   version=1
        6      endian_flag (0 little, 1 big) -> RGB565 payload字节序
        7      reserved
        8..9   fps_q8_8  (fps_out * 256)
        10..11 frame_count
        12..13 disp_w
        14..15 disp_h
        16..17 box_w
        18..19 box_h
        20..21 rot
        其余填0到512
        """
        fps_q8 = int(round(max(0.0, float(fps_out)) * 256.0))
        fps_q8 = clamp(fps_q8, 1, 65535)  # fps不能为0

        endian_flag = 0 if self.var_endian.get() == "little" else 1
        rot = int(self.var_rot.get())
        disp_w = int(self.var_disp_w.get())
        disp_h = int(self.var_disp_h.get())
        box_w = int(self.var_img_w.get())
        box_h = int(self.var_img_h.get())

        out = bytearray()
        out += b"V565"
        out += self._pack_u16(1)         # version
        out += bytes([endian_flag])
        out += b"\x00"
        out += self._pack_u16(fps_q8)    # fps Q8.8
        out += self._pack_u16(frame_count)
        out += self._pack_u16(disp_w) + self._pack_u16(disp_h)
        out += self._pack_u16(box_w) + self._pack_u16(box_h)
        out += self._pack_u16(rot)

        if len(out) > self.SECTOR_SIZE:
            raise RuntimeError(f"头扇区字段超过 512：{len(out)}")
        out += b"\x00" * (self.SECTOR_SIZE - len(out))
        return bytes(out)

    def _build_one_frame_record_fit(self, pil_rgb: Image.Image, dur_ms: int) -> bytes:
        """
        每帧 record：
          16字节帧头：
            u16 cw,ch
            u16 dx,dy
            u16 dur_ms
            u16 reserved
            u32 data_len
          + data_len 字节RGB565
          + pad到512对齐
        """
        conv = self._compute_transformed_image_video_fit(pil_rgb)

        disp_w = int(self.var_disp_w.get())
        disp_h = int(self.var_disp_h.get())
        img_x = (disp_w - conv.width) // 2
        img_y = (disp_h - conv.height) // 2

        class _Tmp:
            pass
        it_like = _Tmp()
        it_like.img_x = img_x
        it_like.img_y = img_y

        self.conv_img_pil = conv
        dx, dy, sx, sy, cw, ch = self._compute_export_params(it_like)

        if cw <= 0 or ch <= 0:
            rgb565 = b""
            cw = ch = 0
            dx = dy = 0
        else:
            crop = conv.crop((sx, sy, sx + cw, sy + ch))
            rgb565 = self._build_rgb565_bytes(crop)

        dur = clamp(int(dur_ms), 0, 65535)

        rec = bytearray()
        rec += self._pack_u16(cw) + self._pack_u16(ch)
        rec += self._pack_u16(dx) + self._pack_u16(dy)
        rec += self._pack_u16(dur)
        rec += b"\x00\x00"
        rec += self._pack_u32(len(rgb565))
        rec += rgb565

        return self._pad_to_sector(bytes(rec))

    def _seek_frame_by_time(self, cap: cv2.VideoCapture, t_s: float, native_fps: float):
        """
        尽量精确到时间取帧：
        - 若native_fps有效：用 frame_index = round(t * native_fps)
        - 否则用 CAP_PROP_POS_MSEC
        """
        if native_fps > 0.0:
            idx = int(round(t_s * native_fps))
            if idx < 0:
                idx = 0
            cap.set(cv2.CAP_PROP_POS_FRAMES, idx)
        else:
            cap.set(cv2.CAP_PROP_POS_MSEC, float(t_s) * 1000.0)

    def on_export_video_bin(self):
        if not self.video_state:
            messagebox.showwarning("提示", "请先导入视频。")
            return

        try:
            out_dir = self._ensure_out_dir()

            # 使用自适应FPS
            target_width = int(self.var_img_w.get())
            target_height = int(self.var_img_h.get())
            fps_out = self._calculate_adaptive_fps(target_width, target_height)
            if fps_out <= 0:
                raise RuntimeError("视频每秒帧数必须 > 0")

            dur_s = float(self.video_state.duration_s or 0.0)
            if dur_s <= 0.0:
                # 时长未知：尽力用 total_frames/native_fps 推断，否则按1秒算
                if self.video_state.total_frames > 0 and self.video_state.native_fps > 0:
                    dur_s = self.video_state.total_frames / self.video_state.native_fps
                else:
                    dur_s = 1.0

            frame_count = max(1, int(math.ceil(dur_s * fps_out)))

            out_name = (self.var_video_name.get() or "media_1_video.bin").strip()
            if not out_name.lower().endswith(".bin"):
                out_name += ".bin"
            out_path = os.path.join(out_dir, safe_name(out_name))

            cap = self.video_state.cap
            native_fps = float(self.video_state.native_fps or 0.0)

            # ✅ 建议 dur_ms=0：由播放端用 fps 来延时（你的 Show_Veo_From_Bin 已支持）
            dur_ms = 0

            with open(out_path, "wb") as f:
                f.write(self._build_video_header_sector(fps_out=fps_out, frame_count=frame_count))

                # i=0..frame_count-1 取 t=i/fps_out
                for i in range(frame_count):
                    t_s = i / fps_out
                    self._seek_frame_by_time(cap, t_s, native_fps)

                    ok, frame = cap.read()
                    if not ok or frame is None:
                        # 空帧
                        empty = bytearray()
                        empty += self._pack_u16(0) + self._pack_u16(0)
                        empty += self._pack_u16(0) + self._pack_u16(0)
                        empty += self._pack_u16(dur_ms)
                        empty += b"\x00\x00"
                        empty += self._pack_u32(0)
                        f.write(self._pad_to_sector(bytes(empty)))
                        continue

                    frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    pil = Image.fromarray(frame).convert("RGB")
                    f.write(self._build_one_frame_record_fit(pil, dur_ms=dur_ms))

            messagebox.showinfo(
                "导出成功",
                f"已导出：\n{out_path}\n"
                f"扇区=512 | fps_out={fps_out:.3f} | 时长≈{dur_s:.3f}s | 总帧={frame_count} | 模式=FIT完整显示(不裁/不补黑)"
            )
        except Exception as e:
            messagebox.showerror("导出失败", str(e))


if __name__ == "__main__":
    app = RGB565ToolApp()
    app.mainloop()
