import datetime
import json
import logging
import math
import os
import subprocess
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox
import winsound
import ctypes
import customtkinter as ctk

try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "timer_config.json")
ICON_FILE = os.path.join(BASE_DIR, "OMtimer.ico")
LOG_FILE = os.path.join(BASE_DIR, "timer.log")

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.ERROR,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8"
)

CREATE_NO_WINDOW = 0x08000000


def get_system_voices():
    """WindowsシステムにインストールされているTTSボイス一覧を取得"""
    voices = []
    if HAS_WIN32COM:
        try:
            import pythoncom
            pythoncom.CoInitialize()
            tts = win32com.client.Dispatch("SAPI.SpVoice")
            for v in tts.GetVoices():
                voices.append(v.GetDescription())
            pythoncom.CoUninitialize()
        except Exception as e:
            logging.error(f"Voice一覧取得エラー: {e}")
    if not voices:
        voices = ["標準音声"]
    return voices


class PomodoroTimerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("時報 ＆ パターン対応ポモドーロタイマー")
        
        # 設定・状態変数
        self.presets = [5, 10, 15, 30]
        self.saved_patterns = {
            "標準ポモドーロ": {"sequence": "4, 1, 4, 1, 4, 3", "snapshot_presets": [5, 10, 15, 30]},
            "ショート": {"sequence": "1, 2", "snapshot_presets": [5, 10, 15, 30]}
        }
        self.selected_preset_idx = 0
        self.selected_seconds = self.remaining_seconds = 5 * 60
        self.is_running = self.is_paused = False
        self.start_timestamp = self.target_seconds = 0.0
        self.last_checked_hour = self.last_checked_half_hour = -1
        
        self.long_press_job = None
        self.long_pressed_flag = False

        self.sound_mode = "default"
        self.custom_audio_path = ""
        self.chime_enabled = True
        self.half_chime_mode = "default"
        self.custom_half_chime_path = ""
        self.hourly_chime_mode = "tts"

        # 追加設定：透過率と時報の予告音・声質
        self.min_opacity = 0.3            # 最小モード時の非ホバー不透明度 (0.1 ~ 1.0)
        self.pre_chime_enabled = True      # 時報前の予告音
        self.selected_voice = ""          # 選択された声質名称

        self.view_mode = 0  # 0:大, 1:中, 2:小
        self.pattern_sequence = []
        self.current_step_idx = 0
        
        self._drag_start_x = self._drag_start_y = 0
        self.window_x = self.window_y = None

        self.load_config()
        self._create_widgets()
        self._setup_hover_events()
        self._apply_initial_geometry()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self._main_loop()

    # --- 設定保存・読み込み ---
    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    p = data.get("presets")
                    if isinstance(p, list) and len(p) == 4 and all(isinstance(x, (int, float)) and x > 0 for x in p):
                        self.presets = [int(x) for x in p]

                    self.sound_mode = data.get("sound_mode", self.sound_mode)
                    self.custom_audio_path = data.get("custom_audio_path", data.get("custom_wav_path", self.custom_audio_path))
                    self.chime_enabled = data.get("chime_enabled", self.chime_enabled)
                    self.half_chime_mode = data.get("half_chime_mode", self.half_chime_mode)
                    self.custom_half_chime_path = data.get("custom_half_chime_path", self.custom_half_chime_path)
                    self.hourly_chime_mode = data.get("hourly_chime_mode", data.get("midnight_chime_mode", self.hourly_chime_mode))

                    self.min_opacity = float(data.get("min_opacity", self.min_opacity))
                    self.pre_chime_enabled = data.get("pre_chime_enabled", self.pre_chime_enabled)
                    self.selected_voice = data.get("selected_voice", self.selected_voice)

                    self.window_x = data.get("window_x")
                    self.window_y = data.get("window_y")
                    self.view_mode = data.get("view_mode", 0)

                    if isinstance(data.get("patterns"), dict):
                        self.saved_patterns = {
                            k: ({"sequence": v, "snapshot_presets": list(self.presets)} if isinstance(v, str) else v)
                            for k, v in data["patterns"].items()
                        }
            except Exception as e:
                logging.error(f"設定読み込みエラー: {e}")
                
        self.selected_seconds = self.remaining_seconds = self.presets[0] * 60

    def save_config(self):
        try:
            self._update_current_pos()
            data = {
                "presets": self.presets,
                "patterns": self.saved_patterns,
                "sound_mode": self.sound_mode,
                "custom_audio_path": self.custom_audio_path,
                "chime_enabled": self.chime_enabled,
                "half_chime_mode": self.half_chime_mode,
                "custom_half_chime_path": self.custom_half_chime_path,
                "hourly_chime_mode": self.hourly_chime_mode,
                "min_opacity": self.min_opacity,
                "pre_chime_enabled": self.pre_chime_enabled,
                "selected_voice": self.selected_voice,
                "window_x": self.window_x,
                "window_y": self.window_y,
                "view_mode": self.view_mode
            }
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.error(f"設定保存エラー: {e}")

    def _update_current_pos(self):
        if self.root.winfo_exists():
            x, y = self.root.winfo_x(), self.root.winfo_y()
            if x > -5000 and y > -5000:
                self.window_x, self.window_y = x, y

    def _apply_initial_geometry(self):
        m = self.view_mode
        if m == 1:
            self._show_frame(self.frame_compact)
            w, h = 250, 105
            self.root.wm_attributes("-alpha", 1.0)
        elif m == 2:
            self._show_frame(self.frame_micro)
            self.root.attributes("-topmost", True)
            self.root.overrideredirect(True)
            w, h = 180, 44
            self.root.wm_attributes("-alpha", self.min_opacity)
        else:
            self._show_frame([self.frame_top, self.frame_bottom_container])
            w, h = 460, 440
            self.root.wm_attributes("-alpha", 1.0)

        if self.window_x is not None and self.window_y is not None:
            self.root.geometry(f"{w}x{h}+{self.window_x}+{self.window_y}")
        else:
            self.root.geometry(f"{w}x{h}")

    def on_closing(self):
        self.save_config()
        self.root.destroy()

    # --- UIパーツ作成ヘルパー ---
    def _create_mode_button(self, parent, mode, text, width=30, height=28, font_size=14):
        return ctk.CTkButton(
            parent, text=text, width=width, height=height,
            font=("MeiryoUI", font_size, "bold"),
            command=lambda: self.switch_view_mode(mode)
        )

    def _bind_drag(self, widget):
        widget.bind("<ButtonPress-1>", self._on_drag_start)
        widget.bind("<B1-Motion>", self._on_drag_motion)

    # --- マウスホバー(透過率)設定 ---
    def _setup_hover_events(self):
        """最小モード時の透過率動的コントロール"""
        def _on_enter(e):
            if self.view_mode == 2:
                self.root.wm_attributes("-alpha", 1.0)

        def _on_leave(e):
            if self.view_mode == 2:
                self.root.wm_attributes("-alpha", self.min_opacity)

        self.root.bind("<Enter>", _on_enter)
        self.root.bind("<Leave>", _on_leave)

    # --- UI描画 ---
    def _create_widgets(self):
        # 1. 常時/標準(大)エリア
        self.frame_top = ctk.CTkFrame(self.root, fg_color="transparent")
        
        f_status = ctk.CTkFrame(self.frame_top, fg_color="transparent")
        f_status.pack(fill="x", padx=15, pady=(2, 0))
        self.label_status = ctk.CTkLabel(f_status, text="停止中", text_color="gray", font=("MeiryoUI", 13, "bold"))
        self.label_status.pack(side="left")
        self.label_active_preset = ctk.CTkLabel(f_status, text="", font=("MeiryoUI", 13, "bold"))
        self.label_active_preset.pack(side="right")

        f_clock = ctk.CTkFrame(self.frame_top, fg_color="transparent")
        f_clock.pack(fill="x", padx=15, pady=0)
        self.label_time = ctk.CTkLabel(f_clock, text="05:00", font=("Helvetica", 54, "bold"), cursor="fleur")
        self.label_time.pack(side="left", expand=True, anchor="c")
        self._bind_drag(self.label_time)

        f_btn_stack = ctk.CTkFrame(f_clock, fg_color="transparent")
        f_btn_stack.pack(side="right", padx=(0, 5))
        self._create_mode_button(f_btn_stack, 1, "◼", 34, 30, 18).pack(pady=(0, 4))
        self._create_mode_button(f_btn_stack, 2, "▪", 34, 30, 18).pack()

        # 2. 折りたたみ(Compact/中)エリア
        self.frame_compact = ctk.CTkFrame(self.root, fg_color="transparent")
        f_cmp_top = ctk.CTkFrame(self.frame_compact, fg_color="transparent")
        f_cmp_top.pack(fill="x", padx=8, pady=(2, 0))
        self.label_cmp_status = ctk.CTkLabel(f_cmp_top, text="停止中", text_color="gray", font=("MeiryoUI", 11, "bold"))
        self.label_cmp_status.pack(side="left")

        f_cmp_modes = ctk.CTkFrame(f_cmp_top, fg_color="transparent")
        f_cmp_modes.pack(side="right")
        self._create_mode_button(f_cmp_modes, 0, "⬛", 26, 22, 14).pack(side="left", padx=1)
        self._create_mode_button(f_cmp_modes, 2, "▪", 26, 22, 14).pack(side="left", padx=1)

        f_cmp_mid = ctk.CTkFrame(self.frame_compact, fg_color="transparent")
        f_cmp_mid.pack(fill="x", padx=8, pady=4)
        
        f_cmp_ctrl = ctk.CTkFrame(f_cmp_mid, fg_color="transparent")
        f_cmp_ctrl.pack(side="left", padx=(0, 4))
        self.btn_cmp_start = ctk.CTkButton(f_cmp_ctrl, text="▶", width=36, height=34, font=("MeiryoUI", 18, "bold"), fg_color="#2E7D32", hover_color="#1B5E20", command=self.toggle_timer)
        self.btn_cmp_start.pack(side="left", padx=1)
        self.btn_cmp_reset = ctk.CTkButton(f_cmp_ctrl, text="⏹", width=36, height=34, font=("MeiryoUI", 18, "bold"), fg_color="#555555", hover_color="#333333", command=self.reset_timer)
        self.btn_cmp_reset.pack(side="left", padx=1)

        self.label_cmp_time = ctk.CTkLabel(f_cmp_mid, text="05:00", font=("Helvetica", 32, "bold"), cursor="fleur")
        self.label_cmp_time.pack(side="left", expand=True)
        self._bind_drag(self.label_cmp_time)

        # 3. 極小(Micro/小)エリア
        self.frame_micro = ctk.CTkFrame(self.root, fg_color="transparent")
        self.btn_micro_start = ctk.CTkButton(self.frame_micro, text="▶", width=32, height=28, font=("MeiryoUI", 15, "bold"), fg_color="#2E7D32", hover_color="#1B5E20", command=self.toggle_timer)
        self.btn_micro_start.pack(side="left", padx=(4, 2))

        self.label_micro_time = ctk.CTkLabel(self.frame_micro, text="05:00", font=("Helvetica", 22, "bold"), cursor="fleur")
        self.label_micro_time.pack(side="left", expand=True)
        self._bind_drag(self.label_micro_time)

        self._create_mode_button(self.frame_micro, 0, "⬛", 30, 28, 14).pack(side="right", padx=(2, 4))

        # 4. メイン操作部（下部）
        self.frame_bottom_container = ctk.CTkFrame(self.root, fg_color="transparent")

        f_presets = ctk.CTkFrame(self.frame_bottom_container, fg_color="transparent")
        f_presets.pack(fill="x", pady=2)
        self.btn_presets = []
        for i in range(4):
            btn = ctk.CTkButton(f_presets, text="", width=85, height=36, font=("MeiryoUI", 14, "bold"), command=lambda idx=i: self._on_preset_click(idx))
            btn.bind("<ButtonPress-1>", lambda e, idx=i: self._on_preset_press(idx), add="+")
            btn.bind("<ButtonRelease-1>", lambda e: self._on_preset_release(), add="+")
            btn.bind("<Leave>", lambda e: self._cancel_long_press(), add="+")
            btn.pack(side="left", expand=True, padx=1)
            self.btn_presets.append(btn)
        self._update_preset_buttons_text()

        f_slider = ctk.CTkFrame(self.frame_bottom_container, fg_color="transparent")
        f_slider.pack(fill="x", pady=2)
        self.slider = ctk.CTkSlider(f_slider, from_=1, to=60, number_of_steps=59, command=self._on_slider_move)
        self.slider.set(self.presets[0])
        self.slider.pack(side="left", fill="x", expand=True, padx=(5, 8))
        self.btn_save_preset = ctk.CTkButton(f_slider, text="💾セット", width=75, height=30, font=("MeiryoUI", 12, "bold"), command=self.save_current_to_preset)
        self.btn_save_preset.pack(side="right")

        f_pat = ctk.CTkFrame(self.frame_bottom_container)
        f_pat.pack(fill="x", pady=2, padx=2, ipady=3)
        ctk.CTkLabel(f_pat, text=" 再生モード設定 ", font=("MeiryoUI", 13, "bold")).pack(anchor="w", padx=8, pady=(2, 0))
        
        f_chk = ctk.CTkFrame(f_pat, fg_color="transparent")
        f_chk.pack(fill="x", padx=8, pady=1)
        self.var_enable_pattern = tk.BooleanVar(value=False)
        self.chk_pattern = ctk.CTkCheckBox(f_chk, text="パターン再生を有効化", font=("MeiryoUI", 13, "bold"), checkbox_width=18, checkbox_height=18, variable=self.var_enable_pattern, command=self._toggle_pattern_ui_state)
        self.chk_pattern.pack(side="left")
        self.var_loop_pattern = tk.BooleanVar(value=True)
        self.chk_loop = ctk.CTkCheckBox(f_chk, text="無限ループ", font=("MeiryoUI", 13, "bold"), checkbox_width=18, checkbox_height=18, variable=self.var_loop_pattern)
        self.chk_loop.pack(side="right")

        f_entry = ctk.CTkFrame(f_pat, fg_color="transparent")
        f_entry.pack(fill="x", padx=8, pady=2)
        ctk.CTkLabel(f_entry, text="順序:", font=("MeiryoUI", 14, "bold")).pack(side="left")
        self.entry_pattern = ctk.CTkEntry(f_entry, height=34, font=("MeiryoUI", 15, "bold"))
        self.entry_pattern.insert(0, "4, 1, 4, 1, 4, 3")
        self.entry_pattern.pack(side="left", fill="x", expand=True, padx=5)

        f_combo = ctk.CTkFrame(f_pat, fg_color="transparent")
        f_combo.pack(fill="x", padx=8, pady=2)
        ctk.CTkLabel(f_combo, text="保存済み:", font=("MeiryoUI", 13, "bold")).pack(side="left")
        self.combo_patterns = ctk.CTkOptionMenu(f_combo, height=32, values=[], command=self._on_pattern_selected, font=("MeiryoUI", 14, "bold"))
        self.combo_patterns.pack(side="left", padx=5)
        self._update_pattern_combo()
        
        ctk.CTkButton(f_combo, text="🗑️ 削除", width=55, height=30, font=("MeiryoUI", 12, "bold"), fg_color="#D32F2F", hover_color="#9A0007", command=self.delete_pattern_preset).pack(side="left", padx=2)
        ctk.CTkButton(f_combo, text="保存/上書き", width=90, height=30, font=("MeiryoUI", 12, "bold"), command=self.save_pattern_preset).pack(side="right")

        f_ctrl = ctk.CTkFrame(self.frame_bottom_container, fg_color="transparent")
        f_ctrl.pack(pady=(4, 0))
        self.btn_start = ctk.CTkButton(f_ctrl, text=" ▶ スタート ", font=("MeiryoUI", 16, "bold"), width=125, height=38, fg_color="#2E7D32", hover_color="#1B5E20", command=self.toggle_timer)
        self.btn_start.pack(side="left", padx=4)
        self.btn_reset = ctk.CTkButton(f_ctrl, text=" ⏹ リセット ", font=("MeiryoUI", 16, "bold"), width=125, height=38, fg_color="#555555", hover_color="#333333", command=self.reset_timer)
        self.btn_reset.pack(side="left", padx=4)
        self.btn_settings = ctk.CTkButton(f_ctrl, text="⚙ 設定", font=("MeiryoUI", 16, "bold"), width=95, height=38, fg_color="#455A64", hover_color="#263238", command=self.open_settings_window)
        self.btn_settings.pack(side="left", padx=4)

        self._toggle_pattern_ui_state()
        self._update_time_display()
        self._update_active_preset_label()

    # --- 画面表示切り替え ---
    def _show_frame(self, target_frames):
        for f in [self.frame_top, self.frame_compact, self.frame_micro, self.frame_bottom_container]:
            f.pack_forget()
        frames = target_frames if isinstance(target_frames, list) else [target_frames]
        for f in frames:
            if f == self.frame_top:
                f.pack(fill="x", pady=(2, 0))
            elif f == self.frame_bottom_container:
                f.pack(fill="x", padx=10, pady=(0, 2))
            elif f == self.frame_compact:
                f.pack(fill="both", expand=True, padx=4, pady=2)
            elif f == self.frame_micro:
                f.pack(fill="both", expand=True, padx=4, pady=4)

    def switch_view_mode(self, mode):
        if self.view_mode == mode:
            return
        self._update_current_pos()
        self.view_mode = mode

        if mode == 0:
            self.root.overrideredirect(False)
            self.root.attributes("-topmost", False)
            self.root.wm_attributes("-alpha", 1.0)
            self._show_frame([self.frame_top, self.frame_bottom_container])
            self._resize_and_set_pos(460, 440)
        elif mode == 1:
            self.root.overrideredirect(False)
            self.root.attributes("-topmost", False)
            self.root.wm_attributes("-alpha", 1.0)
            self._show_frame(self.frame_compact)
            self._resize_and_set_pos(250, 105)
        elif mode == 2:
            self._show_frame(self.frame_micro)
            self.root.attributes("-topmost", True)
            self.root.overrideredirect(True)
            self.root.wm_attributes("-alpha", self.min_opacity)
            self._resize_and_set_pos(180, 44)

        self.save_config()

    def _resize_and_set_pos(self, width, height):
        if self.window_x is not None and self.window_y is not None:
            self.root.geometry(f"{width}x{height}+{self.window_x}+{self.window_y}")
        else:
            self.root.geometry(f"{width}x{height}")

    # --- 設定サブウィンドウ ---
    def open_settings_window(self):
        win = ctk.CTkToplevel(self.root)
        win.title("設定 - 音声・時報・透過率")
        win.geometry("440x530")
        win.resizable(False, False)
        win.attributes("-topmost", True)
        win.after(100, win.grab_set)

        container = ctk.CTkFrame(win, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=15, pady=15)

        # 透過率（不透明度）設定
        f_opacity = ctk.CTkFrame(container)
        f_opacity.pack(fill="x", pady=(0, 8), ipady=4)
        ctk.CTkLabel(f_opacity, text=" 最小化モード時の外観 ", font=("MeiryoUI", 13, "bold")).pack(anchor="w", padx=8, pady=(4, 2))

        f_op_inner = ctk.CTkFrame(f_opacity, fg_color="transparent")
        f_op_inner.pack(fill="x", padx=12, pady=2)
        ctk.CTkLabel(f_op_inner, text="非ホバー時不透明度:", font=("MeiryoUI", 12)).pack(side="left")
        
        lbl_op_val = ctk.CTkLabel(f_op_inner, text=f"{int(self.min_opacity * 100)}%", font=("MeiryoUI", 12, "bold"), width=40)
        lbl_op_val.pack(side="right")

        def _on_op_change(val):
            self.min_opacity = round(float(val), 2)
            lbl_op_val.configure(text=f"{int(self.min_opacity * 100)}%")
            if self.view_mode == 2:
                self.root.wm_attributes("-alpha", self.min_opacity)
            self.save_config()

        slider_op = ctk.CTkSlider(f_opacity, from_=0.1, to=1.0, number_of_steps=18, command=_on_op_change)
        slider_op.set(self.min_opacity)
        slider_op.pack(fill="x", padx=12, pady=(0, 4))

        # 音声設定
        f_sound = ctk.CTkFrame(container)
        f_sound.pack(fill="x", pady=(0, 8), ipady=4)
        ctk.CTkLabel(f_sound, text=" タイマー終了音設定 ", font=("MeiryoUI", 13, "bold")).pack(anchor="w", padx=8, pady=(4, 2))
        
        var_sound_mode = tk.StringVar(value=self.sound_mode)
        def _save_s(): self.sound_mode = var_sound_mode.get(); self.save_config()

        ctk.CTkRadioButton(f_sound, text="標準チャイム音", value="default", variable=var_sound_mode, font=("MeiryoUI", 12), radiobutton_width=16, radiobutton_height=16, command=_save_s).pack(anchor="w", padx=12, pady=2)
        f_audio = ctk.CTkFrame(f_sound, fg_color="transparent")
        f_audio.pack(fill="x", padx=12, pady=3)
        ctk.CTkRadioButton(f_audio, text="音声ファイル:", value="file", variable=var_sound_mode, font=("MeiryoUI", 12), radiobutton_width=16, radiobutton_height=16, command=_save_s).pack(side="left")
        
        entry_audio = ctk.CTkEntry(f_audio, height=32, font=("MeiryoUI", 14))
        entry_audio.insert(0, self.custom_audio_path)
        entry_audio.pack(side="left", fill="x", expand=True, padx=5)

        def _update_audio(e=None):
            self.custom_audio_path = entry_audio.get().strip()
            self.save_config()
        entry_audio.bind("<FocusOut>", _update_audio)
        entry_audio.bind("<Return>", _update_audio)

        def _browse_audio():
            path = filedialog.askopenfilename(title="タイマー終了音を選択", filetypes=[("音声ファイル (*.wav, *.mp3)", "*.wav;*.mp3")], parent=win)
            if path:
                self.custom_audio_path = path
                entry_audio.delete(0, tk.END)
                entry_audio.insert(0, path)
                var_sound_mode.set("file")
                _save_s()

        ctk.CTkButton(f_audio, text="参照...", width=55, height=28, font=("MeiryoUI", 11, "bold"), command=_browse_audio).pack(side="right")

        # 時報設定
        f_chime = ctk.CTkFrame(container)
        f_chime.pack(fill="x", pady=0, ipady=4)
        ctk.CTkLabel(f_chime, text=" 時報機能設定 ", font=("MeiryoUI", 13, "bold")).pack(anchor="w", padx=8, pady=(4, 2))

        var_chime = tk.BooleanVar(value=self.chime_enabled)
        ctk.CTkCheckBox(f_chime, text="時報機能を有効にする", font=("MeiryoUI", 12, "bold"), checkbox_width=18, checkbox_height=18, variable=var_chime, command=lambda: setattr(self, 'chime_enabled', var_chime.get()) or self.save_config()).pack(anchor="w", padx=12, pady=2)

        ctk.CTkLabel(f_chime, text="【30分時報音】", font=("MeiryoUI", 11, "bold")).pack(anchor="w", padx=12, pady=(4, 0))
        var_half_mode = tk.StringVar(value=self.half_chime_mode)
        def _save_h(): self.half_chime_mode = var_half_mode.get(); self.save_config()

        ctk.CTkRadioButton(f_chime, text="標準効果音（ポーン）", value="default", variable=var_half_mode, font=("MeiryoUI", 12), radiobutton_width=16, radiobutton_height=16, command=_save_h).pack(anchor="w", padx=16, pady=1)
        f_half_file = ctk.CTkFrame(f_chime, fg_color="transparent")
        f_half_file.pack(fill="x", padx=16, pady=3)
        ctk.CTkRadioButton(f_half_file, text="音声ファイル:", value="file", variable=var_half_mode, font=("MeiryoUI", 12), radiobutton_width=16, radiobutton_height=16, command=_save_h).pack(side="left")
        
        entry_half = ctk.CTkEntry(f_half_file, height=32, font=("MeiryoUI", 14))
        entry_half.insert(0, self.custom_half_chime_path)
        entry_half.pack(side="left", fill="x", expand=True, padx=5)

        def _update_half(e=None):
            self.custom_half_chime_path = entry_half.get().strip()
            self.save_config()
        entry_half.bind("<FocusOut>", _update_half)
        entry_half.bind("<Return>", _update_half)

        def _browse_half():
            path = filedialog.askopenfilename(title="30分時報音を選択", filetypes=[("音声ファイル (*.wav, *.mp3)", "*.wav;*.mp3")], parent=win)
            if path:
                self.custom_half_chime_path = path
                entry_half.delete(0, tk.END)
                entry_half.insert(0, path)
                var_half_mode.set("file")
                _save_h()

        ctk.CTkButton(f_half_file, text="参照...", width=55, height=28, font=("MeiryoUI", 11, "bold"), command=_browse_half).pack(side="right")

        ctk.CTkLabel(f_chime, text="【毎時0分の挙動 ＆ 声質】", font=("MeiryoUI", 11, "bold")).pack(anchor="w", padx=12, pady=(4, 0))
        var_hourly = tk.StringVar(value=self.hourly_chime_mode)
        def _save_hr(): self.hourly_chime_mode = var_hourly.get(); self.save_config()

        f_hourly = ctk.CTkFrame(f_chime, fg_color="transparent")
        f_hourly.pack(fill="x", padx=16, pady=2)
        ctk.CTkRadioButton(f_hourly, text="時刻読み上げ", value="tts", variable=var_hourly, font=("MeiryoUI", 12), radiobutton_width=16, radiobutton_height=16, command=_save_hr).pack(side="left", padx=(0, 10))
        ctk.CTkRadioButton(f_hourly, text="30分時報と同じ音", value="half", variable=var_hourly, font=("MeiryoUI", 12), radiobutton_width=16, radiobutton_height=16, command=_save_hr).pack(side="left")

        # 予告音（予鈴）チェックボックス
        var_pre_chime = tk.BooleanVar(value=self.pre_chime_enabled)
        ctk.CTkCheckBox(f_chime, text="読み上げ前に予告チャイムを鳴らす", font=("MeiryoUI", 11, "bold"), checkbox_width=16, checkbox_height=16, variable=var_pre_chime, command=lambda: setattr(self, 'pre_chime_enabled', var_pre_chime.get()) or self.save_config()).pack(anchor="w", padx=16, pady=3)

        # 声質（ボイス）選択
        f_voice = ctk.CTkFrame(f_chime, fg_color="transparent")
        f_voice.pack(fill="x", padx=16, pady=2)
        ctk.CTkLabel(f_voice, text="声質選択:", font=("MeiryoUI", 11, "bold")).pack(side="left")
        
        voices = get_system_voices()
        if not self.selected_voice and voices:
            self.selected_voice = voices[0]

        def _on_voice_change(choice):
            self.selected_voice = choice
            self.save_config()

        combo_voice = ctk.CTkOptionMenu(f_voice, values=voices, command=_on_voice_change, font=("MeiryoUI", 11), height=28)
        combo_voice.set(self.selected_voice if self.selected_voice in voices else (voices[0] if voices else ""))
        combo_voice.pack(side="left", fill="x", expand=True, padx=5)

        ctk.CTkButton(container, text=" 閉じる ", font=("MeiryoUI", 12, "bold"), width=100, height=30, command=win.destroy).pack(pady=(12, 0))

    # --- ドラッグ処理 ---
    def _on_drag_start(self, event):
        self._drag_start_x, self._drag_start_y = event.x, event.y

    def _on_drag_motion(self, event):
        x = self.root.winfo_x() + (event.x - self._drag_start_x)
        y = self.root.winfo_y() + (event.y - self._drag_start_y)
        self.window_x, self.window_y = x, y
        self.root.geometry(f"+{x}+{y}")

    # --- 補助・更新処理 ---
    def _set_timer_minutes(self, minutes):
        self.slider.set(minutes)
        self.selected_seconds = self.remaining_seconds = minutes * 60
        self._update_time_display()

    def _update_preset_buttons_text(self):
        symbols = ["①", "②", "③", "④"]
        for idx, btn in enumerate(self.btn_presets):
            btn.configure(text=f"{symbols[idx]} {self.presets[idx]}分")

    def _update_active_preset_label(self):
        symbols = ["①", "②", "③", "④"]
        state_txt, color = ("⚙ 実行中", "#2E7D32") if self.is_running else ("▶ 選択中", "#1E88E5")
        self.label_active_preset.configure(
            text=f"{state_txt}: ボタン{symbols[self.selected_preset_idx]} ({self.presets[self.selected_preset_idx]}分)",
            text_color=color
        )

    def _update_pattern_combo(self):
        keys = list(self.saved_patterns.keys())
        self.combo_patterns.configure(values=keys if keys else ["なし"])
        self.combo_patterns.set(keys[0] if keys else "なし")

    def _toggle_pattern_ui_state(self):
        st = "normal" if self.var_enable_pattern.get() else "disabled"
        self.entry_pattern.configure(state=st)
        self.combo_patterns.configure(state=st)

    def _update_time_display(self):
        mins, secs = divmod(self.remaining_seconds, 60)
        timer_str = f"{mins:02d}:{secs:02d}"
        self.label_time.configure(text=timer_str)
        self.label_cmp_time.configure(text=timer_str)

        if not self.is_running and not self.is_paused:
            self.label_micro_time.configure(text=datetime.datetime.now().strftime("%H:%M"))
        else:
            self.label_micro_time.configure(text=timer_str)

    # --- プリセット＆パターン操作 ---
    def _on_preset_press(self, index):
        if self.is_running: return
        self._cancel_long_press()
        self.long_pressed_flag = False
        self.long_press_job = self.root.after(450, self._on_preset_long_press, index)

    def _on_preset_release(self):
        self._cancel_long_press()

    def _cancel_long_press(self):
        if self.long_press_job:
            try: self.root.after_cancel(self.long_press_job)
            except tk.TclError: pass
            self.long_press_job = None

    def _on_preset_long_press(self, index):
        self.long_press_job = None
        self.long_pressed_flag = True
        if self.is_running: return
        self.presets[index] = int(float(self.slider.get()))
        self.selected_preset_idx = index
        self._update_preset_buttons_text()
        self._update_active_preset_label()
        self.save_config()
        winsound.Beep(1800, 100)

    def _on_preset_click(self, index):
        if self.long_pressed_flag or self.is_running:
            self.long_pressed_flag = False
            return

        if self.var_enable_pattern.get():
            st = self.entry_pattern.cget("state")
            self.entry_pattern.configure(state="normal")
            cur = self.entry_pattern.get().strip()
            self.entry_pattern.delete(0, tk.END)
            self.entry_pattern.insert(0, f"{cur}, {index + 1}" if cur else str(index + 1))
            self.entry_pattern.configure(state=st)
        else:
            self.selected_preset_idx = index
            self._set_timer_minutes(self.presets[index])
            self._update_active_preset_label()

    def _on_pattern_selected(self, choice):
        if choice in self.saved_patterns:
            data = self.saved_patterns[choice]
            st = self.entry_pattern.cget("state")
            self.entry_pattern.configure(state="normal")
            self.entry_pattern.delete(0, tk.END)
            self.entry_pattern.insert(0, data.get("sequence", ""))
            self.entry_pattern.configure(state=st)

            snapshot = data.get("snapshot_presets")
            if snapshot and len(snapshot) == 4:
                self.presets = list(snapshot)
                self._update_preset_buttons_text()
                self._set_timer_minutes(self.presets[self.selected_preset_idx])
                self._update_active_preset_label()
                self.save_config()

    def save_pattern_preset(self):
        pattern_str = self.entry_pattern.get().strip()
        if not pattern_str: return messagebox.showwarning("警告", "パターンが入力されていません。")

        win = ctk.CTkToplevel(self.root)
        win.title("パターンの保存")
        win.geometry("300x130")
        win.resizable(False, False)
        win.attributes("-topmost", True)
        
        ctk.CTkLabel(win, text="パターン名を入力してください:", font=("MeiryoUI", 11, "bold")).pack(pady=4)
        entry_name = ctk.CTkEntry(win, width=200, height=32, font=("MeiryoUI", 13, "bold"))
        entry_name.insert(0, self.combo_patterns.get())
        entry_name.pack(pady=4)

        def _do_save():
            name = entry_name.get().strip()
            if name:
                if name in self.saved_patterns and not messagebox.askyesno("確認", f"「{name}」を上書きしますか？", parent=win):
                    return
                self.saved_patterns[name] = {"sequence": pattern_str, "snapshot_presets": list(self.presets)}
                self.save_config()
                self._update_pattern_combo()
                self.combo_patterns.set(name)
                win.destroy()
                messagebox.showinfo("完了", f"パターン「{name}」を保存しました。")
        
        ctk.CTkButton(win, text="保存", width=80, height=30, font=("MeiryoUI", 11, "bold"), command=_do_save).pack(pady=4)

    def delete_pattern_preset(self):
        name = self.combo_patterns.get()
        if not name or name not in self.saved_patterns:
            return messagebox.showwarning("警告", "削除するパターンを選択してください。")
        if messagebox.askyesno("削除確認", f"パターン「{name}」を削除しますか？"):
            del self.saved_patterns[name]
            self.save_config()
            self._update_pattern_combo()
            st = self.entry_pattern.cget("state")
            self.entry_pattern.configure(state="normal")
            self.entry_pattern.delete(0, tk.END)
            self.entry_pattern.configure(state=st)

    def _on_slider_move(self, value):
        if not self.is_running:
            self.selected_seconds = self.remaining_seconds = int(float(value)) * 60
            self._update_time_display()

    def save_current_to_preset(self):
        if not self.is_running:
            self.presets[self.selected_preset_idx] = int(float(self.slider.get()))
            self._update_preset_buttons_text()
            self._update_active_preset_label()
            self.save_config()

    def _parse_pattern_string(self, raw_str):
        trans = str.maketrans("０１２３４５６７８９①②③④、，", "01234567891234,,")
        return [int(x) - 1 for x in raw_str.translate(trans).replace(",", " ").split() if x.isdecimal() and 1 <= int(x) <= 4]

    # --- タイマーコントロール ---
    def _set_btn_states(self, start_text, color, hover_color, status_text, status_color):
        for btn in [self.btn_start, self.btn_cmp_start, self.btn_micro_start]:
            btn.configure(text=start_text if btn == self.btn_start else start_text.strip().split()[0], fg_color=color, hover_color=hover_color)
        self.label_status.configure(text=status_text, text_color=status_color)
        self.label_cmp_status.configure(text=status_text, text_color=status_color)

    def toggle_timer(self):
        if self.is_running:
            self.is_running = False
            self.is_paused = True
            self._set_btn_states(" ▶ スタート ", "#2E7D32", "#1B5E20", "一時停止中", "orange")
            self.set_ui_state(is_active=True)
        else:
            resuming = self.is_paused and self.remaining_seconds > 0
            self.is_paused = False

            if not resuming:
                if self.var_enable_pattern.get():
                    seq = self._parse_pattern_string(self.entry_pattern.get())
                    if not seq: return messagebox.showwarning("エラー", "有効なパターンが指定されていません。(1〜4の数字)")
                    self.pattern_sequence = seq
                    self.current_step_idx = 0
                    self._load_current_pattern_step()
                else:
                    if self.remaining_seconds <= 0:
                        self.remaining_seconds = self.selected_seconds
                    self._set_btn_states(" ❚❚ 一時停止 ", "#E65100", "#B71C1C", "タイマー実行中", "#2E7D32")

            if self.remaining_seconds > 0:
                self.is_running = True
                self.target_seconds = self.remaining_seconds
                self.start_timestamp = time.monotonic()
                self._set_btn_states(" ❚❚ 一時停止 ", "#E65100", "#B71C1C", "タイマー実行中", "#2E7D32")
                self.set_ui_state(is_active=False)

        self._update_active_preset_label()
        self._update_time_display()

    def _load_current_pattern_step(self):
        idx = self.pattern_sequence[self.current_step_idx]
        self.selected_preset_idx = idx
        self.selected_seconds = self.remaining_seconds = self.target_seconds = self.presets[idx] * 60
        self.start_timestamp = time.monotonic()
        self._update_time_display()
        self._update_active_preset_label()
        txt = f"パターン進行中 [{self.current_step_idx + 1}/{len(self.pattern_sequence)}]"
        self.label_status.configure(text=txt, text_color="gray")
        self.label_cmp_status.configure(text=txt, text_color="gray")

    def _advance_pattern_step(self):
        self.current_step_idx += 1
        if self.current_step_idx >= len(self.pattern_sequence):
            if self.var_loop_pattern.get():
                self.current_step_idx = 0
                self._load_current_pattern_step()
                self.is_running = True
            else:
                self.is_running = self.is_paused = False
                self._set_btn_states(" ▶ スタート ", "#2E7D32", "#1B5E20", "パターン再生完了", "#1E88E5")
                self._update_active_preset_label()
                self.set_ui_state(is_active=True)
        else:
            self._load_current_pattern_step()
            self.is_running = True

    def reset_timer(self):
        self.is_running = self.is_paused = False
        self._set_btn_states(" ▶ スタート ", "#2E7D32", "#1B5E20", "停止中", "gray")
        self.remaining_seconds = self.selected_seconds
        self._update_time_display()
        self._update_active_preset_label()
        self.set_ui_state(is_active=True)

    def set_ui_state(self, is_active):
        st = "normal" if is_active else "disabled"
        for btn in self.btn_presets: btn.configure(state=st)
        self.slider.configure(state=st)
        self.btn_save_preset.configure(state=st)
        self.chk_pattern.configure(state=st)
        self.chk_loop.configure(state=st)
        if is_active:
            self._toggle_pattern_ui_state()
        else:
            self.entry_pattern.configure(state="disabled")
            self.combo_patterns.configure(state="disabled")

    # --- メイン監視 ＆ 音声再生 ---
    def _main_loop(self):
        now = datetime.datetime.now()
        if self.chime_enabled:
            if now.minute == 0 and self.last_checked_hour != now.hour:
                self.last_checked_hour = now.hour
                self.play_time_signal(now.hour)
            if now.minute == 30 and self.last_checked_half_hour != now.hour:
                self.last_checked_half_hour = now.hour
                self.play_half_hour_chime()

        if self.is_running:
            calc_rem = int(math.ceil(self.target_seconds - (time.monotonic() - self.start_timestamp)))
            if calc_rem > 0:
                if calc_rem != self.remaining_seconds:
                    self.remaining_seconds = calc_rem
                    self._update_time_display()
            else:
                self.remaining_seconds = 0
                self._update_time_display()
                self.play_alarm()
                self.is_paused = False
                if self.var_enable_pattern.get():
                    self._advance_pattern_step()
                else:
                    self.is_running = False
                    self._set_btn_states(" ▶ スタート ", "#2E7D32", "#1B5E20", "タイマー終了", "#1E88E5")
                    self._update_active_preset_label()
                    self.set_ui_state(is_active=True)
        else:
            if self.view_mode == 2:
                self._update_time_display()

        self.root.after(200, self._main_loop)

    def play_time_signal(self, current_hour):
        if self.hourly_chime_mode == "half":
            self.play_half_hour_chime()
            return

        def _speak():
            # 1. 予告チャイムを再生
            if self.pre_chime_enabled:
                self.play_half_hour_chime()
                time.sleep(1.2)  # チャイムが鳴り終わるまでの間隔

            # 2. 音声読み上げ
            text = f"{current_hour}時をお知らせします。"
            success = False

            if HAS_WIN32COM:
                try:
                    import pythoncom
                    pythoncom.CoInitialize()
                    tts = win32com.client.Dispatch("SAPI.SpVoice")
                    tts.Rate = -2     # 少しゆっくり聞き取りやすく
                    tts.Volume = 100

                    # 選択されたボイスを適用
                    if self.selected_voice:
                        for v in tts.GetVoices():
                            if v.GetDescription() == self.selected_voice:
                                tts.Voice = v
                                break

                    tts.Speak(text)
                    pythoncom.CoUninitialize()
                    success = True
                except Exception as e:
                    logging.error(f"SAPI5エラー: {e}")

            if not success:
                try:
                    ps_cmd = f"Add-Type -AssemblyName System.Speech; $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer; $synth.Rate = -2; $synth.Speak('{text}');"
                    r = subprocess.run(["powershell", "-NoProfile", "-Command", ps_cmd], creationflags=CREATE_NO_WINDOW, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    if r.returncode == 0: success = True
                except Exception as e:
                    logging.error(f"PowerShell TTSエラー: {e}")

            if not success:
                winsound.Beep(1568, 200)
                winsound.Beep(2000, 500)

        threading.Thread(target=_speak, daemon=True).start()

    def play_half_hour_chime(self):
        def _chime():
            if self.half_chime_mode == "file" and self.custom_half_chime_path and os.path.exists(self.custom_half_chime_path):
                try:
                    mci = ctypes.windll.winmm.mciSendStringW
                    mci("close half_chime", None, 0, 0)
                    if mci(f'open "{self.custom_half_chime_path}" type mpegvideo alias half_chime', None, 0, 0) == 0:
                        mci("play half_chime", None, 0, 0)
                    else:
                        winsound.PlaySound(self.custom_half_chime_path, winsound.SND_FILENAME | winsound.SND_ASYNC)
                except Exception as e:
                    logging.error(f"30分時報エラー: {e}")
                    winsound.Beep(2093, 1200)
            else:
                winsound.Beep(2093, 1200)
        threading.Thread(target=_chime, daemon=True).start()

    def play_alarm(self):
        def _alarm():
            if self.sound_mode == "file" and self.custom_audio_path and os.path.exists(self.custom_audio_path):
                try:
                    mci = ctypes.windll.winmm.mciSendStringW
                    mci("close timer_alarm", None, 0, 0)
                    if mci(f'open "{self.custom_audio_path}" type mpegvideo alias timer_alarm', None, 0, 0) == 0:
                        mci("play timer_alarm", None, 0, 0)
                    else:
                        winsound.PlaySound(self.custom_audio_path, winsound.SND_FILENAME | winsound.SND_ASYNC)
                except Exception as e:
                    logging.error(f"アラーム再生エラー: {e}")
                    self._play_default_chime()
            else:
                self._play_default_chime()
        threading.Thread(target=_alarm, daemon=True).start()

    def _play_default_chime(self):
        for freq, dur in [(1047, 150), (1318, 150), (1568, 150), (2093, 400)]:
            winsound.Beep(freq, dur)


if __name__ == "__main__":
    root = ctk.CTk()
    if os.path.exists(ICON_FILE):
        try: root.iconbitmap(default=ICON_FILE)
        except Exception as e: logging.error(f"アイコンエラー: {e}")
    app = PomodoroTimerApp(root)
    root.mainloop()