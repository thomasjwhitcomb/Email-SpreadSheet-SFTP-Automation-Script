"""
gui.py  -  AHA Registration Bot  -  Desktop Application
=========================================================
A customtkinter GUI that wraps aha_registration_bot.py.

Run:
    python gui.py

Pages
-----
  Home          Dashboard stats, activity log, auto-mode controls
  Outlook       Email scanning status and configuration
  Google Sheets Sheet statistics for both AHA and Acuity sheets
  SFTP          Upload status, next window, manual trigger
  Settings      Edit all .env values with save support
"""

import json
import os
import queue
import threading
import logging
import tkinter as tk
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Callable, cast

import customtkinter as ctk
from dotenv import set_key, dotenv_values

import aha_registration_bot as bot
from aha_registration_bot import (
    get_analytics,
    run_scan,
    sftp_upload_sheet,
    auto_mode,
    reauthenticate,
    OAuthExpiredError,
    # SFTP keyring
    sftp_keyring_configured,
    set_sftp_password,
    # Outlook keyring
    outlook_keyring_configured,
    set_outlook_password,
    get_outlook_password,
    # Atlas keyring
    atlas_keyring_configured,
    set_atlas_password,
    get_atlas_password,
    # Test-connection helpers
    test_outlook_connection,
    test_sheets_connection,
    test_sftp_connection,
    # SFTP host-key helpers
    check_sftp_host_key,
    add_sftp_host_key,
)

EnvDict = dict[str, str | None]

# --- Appearance ---
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

_ACCENT           = "#C8102E"          # AHA red
_SIDEBAR          = 215
_ENV              = Path(__file__).parent / ".env"
_ANALYTICS_CACHE  = Path(__file__).parent / ".analytics_cache.json"


def _env_text(env: EnvDict, key: str, default: str = "") -> str:
    """Return a non-None string from dotenv_values()."""
    value = env.get(key)
    return default if value is None else value

# --- Sidebar navigation items ---
_NAV = [
    ("Home",           "home"),
    ("Analytics",      "analytics"),
    ("Outlook",        "outlook"),
    ("Google Sheets",  "sheets"),
    ("RQI Upload",     "sftp"),
    ("Settings",       "settings"),
]

# --- Settings schema ---
# (group, label, env_key, kind)
# kind: "text" | "password" | "int" | "bool" | "choice:opt1,opt2,..."
_SCHEMA = [
    ("Outlook",       "Email",                 "OUTLOOK_EMAIL",            "text"),
    ("Outlook",       "Password",              "OUTLOOK_PASSWORD",         "keyring"),
    ("Atlas",         "Email",                 "ATLAS_EMAIL",              "text"),
    ("Atlas",         "Password",              "ATLAS_PASSWORD",           "keyring"),
    ("Atlas",         "Organization Name",     "ORGANIZATION_NAME",        "text"),
    ("Google Sheets", "AHA Sheet Name",        "GOOGLE_SHEET_NAME",        "text"),
    ("Google Sheets", "Acuity Sheet Name",     "ACUITY_GOOGLE_SHEET_NAME", "text"),
    ("Google Sheets", "RQI Delta Sheet Name",  "RQI_DELTA_SHEET_NAME",     "text"),
    ("SFTP",          "Host",                  "SFTP_HOST",                "text"),
    ("SFTP",          "Port",                  "SFTP_PORT",                "int"),
    ("SFTP",          "Username",              "SFTP_USERNAME",            "text"),
    ("SFTP",          "Remote Directory",      "SFTP_REMOTE_DIR",          "text"),
    ("SFTP",          "Local Directory",       "SFTP_LOCAL_DIR",           "text"),
    ("SFTP",          "Filename",              "SFTP_FILENAME",            "text"),
    ("SFTP",          "Keyring Service",       "SFTP_KEYRING_SERVICE",     "text"),
    ("Acuity",        "Sender Email",          "ACUITY_SENDER_EMAIL",      "text"),
    ("Automation",    "Scan Interval (sec)",   "SCAN_INTERVAL_SECONDS",    "int"),
    ("Automation",    "Email Lookback (days)", "EMAIL_LOOKBACK_DAYS",      "int"),
    ("Automation",    "Browser",               "BROWSER",                  "choice:chromium,firefox,webkit"),
    ("Automation",    "Headless Mode",         "HEADLESS",                 "bool"),
    ("Automation",    "Dry Run",               "DRY_RUN",                  "bool"),
]


# ---  ---
class _PasswordSetupDialog(ctk.CTkToplevel):
    """
    Generic modal dialog for storing any credential in the OS keychain.

    Parameters
    ----------
    parent       : App
        Main window (provides the callback after a successful save).
    title        : str
        Window title bar text.
    header       : str
        Large heading shown inside the dialog (may include emoji).
    description  : str
        Two-line explanation shown under the heading.
    pw_label     : str
        Label for the password entry field (e.g. "SFTP Password").
    info_rows    : list[tuple[str, str]]
        (label, value) pairs shown in the read-only info box.
    save_fn      : callable(str) -> bool
        Called with the entered password; returns True on success.
    on_success_cb: callable()
        Called on the *parent* after a successful keyring write.
    """

    def __init__(
        self,
        parent: "App",
        *,
        title: str,
        header: str,
        description: str,
        pw_label: str,
        info_rows: list[tuple[str, str]],
        save_fn: Callable[[str], bool],
        on_success_cb: Callable[[], None],
    ):
        super().__init__(parent)
        self.title(title)
        self.geometry("500x370")
        self.resizable(False, False)
        self.grab_set()
        self.lift()
        self.focus_force()

        self._save_fn      = save_fn
        self._on_success   = on_success_cb

        # --- Header ---
        ctk.CTkLabel(
            self, text=header,
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(padx=28, pady=(28, 4), anchor="w")

        ctk.CTkLabel(
            self, text=description,
            font=ctk.CTkFont(size=12), text_color="gray55", justify="left",
        ).pack(padx=28, pady=(0, 18), anchor="w")

        # --- Info box (read-only) ---
        if info_rows:
            info = ctk.CTkFrame(self, fg_color="#1a1a1a", corner_radius=8)
            info.pack(padx=28, fill="x")
            for lbl, val in info_rows:
                r = ctk.CTkFrame(info, fg_color="transparent")
                r.pack(fill="x", padx=14, pady=3)
                ctk.CTkLabel(r, text=lbl + ":", width=75, anchor="w",
                             font=ctk.CTkFont(size=11), text_color="gray50").pack(side="left")
                ctk.CTkLabel(r, text=val, anchor="w",
                             font=ctk.CTkFont(size=11, weight="bold")).pack(side="left")

        # --- Password entry ---
        ctk.CTkLabel(
            self, text=pw_label, anchor="w",
            font=ctk.CTkFont(size=12),
        ).pack(padx=28, pady=(20, 4), anchor="w")

        pw_row = ctk.CTkFrame(self, fg_color="transparent")
        pw_row.pack(padx=28, fill="x")
        self._pw_var   = ctk.StringVar()
        self._pw_entry = ctk.CTkEntry(
            pw_row, textvariable=self._pw_var,
            show="*", height=36, font=ctk.CTkFont(size=13),
        )
        self._pw_entry.pack(side="left", fill="x", expand=True)
        self._pw_entry.bind("<Return>", lambda _: self._save())
        self._pw_entry.focus()

        self._shown = False
        ctk.CTkButton(
            pw_row, text="Show", width=72, height=36,
            fg_color="#2b2b2b", hover_color="#3a3a3a",
            command=self._toggle_show,
        ).pack(side="left", padx=(10, 0))

        # --- Status label ---
        self._status_var = ctk.StringVar(value="")
        ctk.CTkLabel(
            self, textvariable=self._status_var,
            font=ctk.CTkFont(size=11), text_color="#ff6b6b",
        ).pack(padx=28, pady=(8, 0), anchor="w")

        # --- Buttons ---
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(padx=28, pady=(16, 28), anchor="e")

        ctk.CTkButton(
            btn_row, text="Skip for Now",
            fg_color="transparent", border_width=1, border_color="gray35",
            width=120, height=36,
            command=self.destroy,
        ).pack(side="left", padx=(0, 12))

        self._btn_save = ctk.CTkButton(
            btn_row, text="Save Password",
            fg_color=_ACCENT, hover_color="#a00d24",
            width=148, height=36,
            command=self._save,
        )
        self._btn_save.pack(side="left")

    def _toggle_show(self):
        self._shown = not self._shown
        self._pw_entry.configure(show="" if self._shown else "*")

    def _save(self):
        password = self._pw_var.get()
        if not password:
            self._status_var.set("WARN  Password cannot be empty.")
            return
        self._btn_save.configure(state="disabled", text="Saving...")
        if self._save_fn(password):
            self._on_success()
            self.destroy()
        else:
            self._status_var.set(
                "ERROR  Could not save to keychain - see Activity Log for details."
            )
            self._btn_save.configure(state="normal", text="Save Password")


# ---  ---
class _HostKeyDialog(ctk.CTkToplevel):
    """
    Modal dialog shown at startup when the SFTP host key is not yet in
    known_hosts.  Displays the key fingerprint and asks the user to
    approve adding it before any upload is attempted.
    """

    def __init__(self, parent: "App", key_type: str, fingerprint: str):
        super().__init__(parent)
        self.title("SFTP Host Key Verification")
        self.geometry("520x310")
        self.resizable(False, False)
        self.grab_set()
        self.lift()
        self.focus_force()

        self._parent = parent

        ctk.CTkLabel(
            self, text="New SFTP Host Key Detected",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(padx=28, pady=(28, 4), anchor="w")

        ctk.CTkLabel(
            self,
            text=(
                "The SFTP server's host key is not yet in your known_hosts file.\n"
                "Verify the fingerprint below matches your server before accepting."
            ),
            font=ctk.CTkFont(size=12), text_color="gray55", justify="left",
        ).pack(padx=28, pady=(0, 16), anchor="w")

        # Key info box
        info = ctk.CTkFrame(self, fg_color="#1a1a1a", corner_radius=8)
        info.pack(padx=28, fill="x")
        for label, value in [("Host", f"{bot.SFTP_HOST}:{bot.SFTP_PORT}"),
                              ("Key type", key_type or "unknown"),
                              ("Fingerprint", fingerprint)]:
            row = ctk.CTkFrame(info, fg_color="transparent")
            row.pack(fill="x", padx=12, pady=3)
            ctk.CTkLabel(row, text=f"{label}:", width=90,
                         font=ctk.CTkFont(size=12), text_color="gray60",
                         anchor="w").pack(side="left")
            ctk.CTkLabel(row, text=value,
                         font=ctk.CTkFont(size=12, family="Courier"),
                         anchor="w").pack(side="left", fill="x", expand=True)

        # Buttons
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(padx=28, pady=20, fill="x")

        self._status = ctk.CTkLabel(btn_row, text="", text_color="gray55",
                                    font=ctk.CTkFont(size=11))
        self._status.pack(side="left", fill="x", expand=True)

        ctk.CTkButton(
            btn_row, text="Decline", width=100, fg_color="#333",
            hover_color="#444", command=self._decline,
        ).pack(side="right", padx=(8, 0))

        ctk.CTkButton(
            btn_row, text="Trust & Add", width=120,
            command=self._accept,
        ).pack(side="right")

    def _accept(self):
        self._status.configure(text="Adding key...")
        self.update_idletasks()
        ok, msg = add_sftp_host_key()
        if ok:
            self._parent.add_log(f"OK  SFTP host key added to known_hosts. {msg}")
            self.destroy()
        else:
            self._status.configure(text=f"ERROR  {msg}")

    def _decline(self):
        self._parent.add_log(
            "WARN  SFTP host key not added - uploads will fail until the key is trusted."
        )
        self.destroy()


# ---  ---
class App(ctk.CTk):
    """Main application window."""

    # --- Initialisation ---
    def __init__(self):
        super().__init__()
        self.title("CPR Lifeline Automation")
        self.geometry("1160x730")
        self.minsize(920, 600)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # Window / taskbar icon
        _icon = Path(__file__).parent / "icon.ico"
        if _icon.exists():
            self.iconbitmap(str(_icon))

        # Threading state
        self._stop   = threading.Event()
        self._thread: threading.Thread | None = None
        self._busy   = False
        self._analytics_running = False   # guard: only one analytics thread at a time

        # Async result queues
        self._aq: queue.Queue[tuple[str, Any]] = queue.Queue()   # analytics
        self._lq: queue.Queue[str] = queue.Queue()   # log lines

        # Dynamic label variables  {key: StringVar}
        self._sv: dict[str, ctk.StringVar] = {}

        # Settings entry variables  {env_key: tk.Variable}
        self._svar: dict[str, tk.Variable] = {}

        # Page frames and nav buttons
        self._pages:   dict[str, ctk.CTkFrame]  = {}
        self._navbtns: dict[str, ctk.CTkButton] = {}

        # Widget refs for enable/disable
        self._btn_auto:        ctk.CTkButton  | None = None
        self._btn_run_once:    ctk.CTkButton  | None = None
        self._btn_scan_now:    ctk.CTkButton  | None = None
        self._btn_sftp_now:    ctk.CTkButton  | None = None
        self._btn_reauth:      ctk.CTkButton  | None = None
        self._oauth_banner:    ctk.CTkFrame   | None = None
        self._scan_fail_banner: ctk.CTkFrame  | None = None
        self._log_box:         ctk.CTkTextbox | None = None

        self._tick_count = 0
        self._last_scan_ts: datetime | None = None   # tracks the latest completed scan timestamp

        # Analytics baseline tracking for daily/weekly delta views
        self._analytics_daily_baseline: dict[str, Any] | None = None
        self._analytics_weekly_baseline: dict[str, Any] | None = None
        self._analytics_baseline_date:   str = ""   # YYYY-MM-DD of current daily baseline
        self._analytics_baseline_week:   str = ""   # YYYY-MM-DD of current weekly (Monday) baseline
        self._load_analytics_baselines()

        self._build()
        self._hook_logger()
        self._show("home")
        self._tick()
        self.after(5000, self._refresh_analytics)
        self.after(1200, self._check_all_credentials)   # first-run keyring checks

    # --- Layout ---
    def _build(self):
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # --- Sidebar ---
        sb = ctk.CTkFrame(self, width=_SIDEBAR, corner_radius=0, fg_color="#161616")
        sb.grid(row=0, column=0, sticky="nsew")
        sb.grid_propagate(False)
        sb.grid_rowconfigure(7, weight=1)
        sb.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            sb, text="CPR Lifeline",
            font=ctk.CTkFont(size=21, weight="bold"),
            text_color=_ACCENT,
        ).grid(row=0, column=0, padx=16, pady=(26, 0))
        ctk.CTkLabel(
            sb, text="Automation",
            font=ctk.CTkFont(size=10), text_color="gray40",
        ).grid(row=1, column=0, padx=16, pady=(2, 22))

        for i, (label, key) in enumerate(_NAV, start=2):
            btn = ctk.CTkButton(
                sb, text=label, anchor="w",
                fg_color="transparent", hover_color="#242424",
                font=ctk.CTkFont(size=13), height=42, corner_radius=6,
                command=lambda k=key: self._show(k),
            )
            btn.grid(row=i, column=0, padx=10, pady=2, sticky="ew")
            self._navbtns[key] = btn

        ctk.CTkLabel(
            sb, text="CSC131  -  Team 08",
            font=ctk.CTkFont(size=9), text_color="gray30",
        ).grid(row=8, column=0, padx=16, pady=14, sticky="s")

        # --- Content area ---
        self._content = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self._content.grid(row=0, column=1, sticky="nsew")
        self._content.grid_rowconfigure(0, weight=1)
        self._content.grid_columnconfigure(0, weight=1)

        # --- OAuth expired banner (row=1, hidden until needed) ---
        banner = ctk.CTkFrame(self, height=46, corner_radius=0, fg_color="#4a1500")
        banner.grid(row=1, column=0, columnspan=2, sticky="ew")
        banner.grid_remove()   # hidden by default
        self._oauth_banner = banner

        ctk.CTkLabel(
            banner,
            text="Google authentication expired - Sheets access is unavailable.",
            font=ctk.CTkFont(size=12),
            text_color="#ffbb80",
        ).pack(side="left", padx=16)

        btn_reauth = ctk.CTkButton(
            banner, text="Re-authenticate",
            fg_color="#c85000", hover_color="#a03c00",
            width=148, height=30,
            command=self._do_reauth,
        )
        btn_reauth.pack(side="right", padx=14, pady=8)
        self._btn_reauth = btn_reauth

        # --- Scan failure banner (row=2, hidden until a scan fails) ---
        self._sv["fail_msg"] = ctk.StringVar(value="")
        fail_banner = ctk.CTkFrame(self, height=46, corner_radius=0, fg_color="#3a0a0a")
        fail_banner.grid(row=2, column=0, columnspan=2, sticky="ew")
        fail_banner.grid_remove()   # hidden by default
        self._scan_fail_banner = fail_banner

        ctk.CTkLabel(
            fail_banner,
            text="WARN",
            font=ctk.CTkFont(size=16),
            text_color="#ff8080",
        ).pack(side="left", padx=(14, 4))
        ctk.CTkLabel(
            fail_banner,
            textvariable=self._sv["fail_msg"],
            font=ctk.CTkFont(size=12),
            text_color="#ff9090",
        ).pack(side="left", padx=(0, 16))
        ctk.CTkButton(
            fail_banner, text="Dismiss",
            fg_color="#5a1515", hover_color="#7a2020",
            width=90, height=30,
            command=self._dismiss_fail_banner,
        ).pack(side="right", padx=14, pady=8)

        # --- Status bar (row=3) ---
        bar = ctk.CTkFrame(self, height=30, corner_radius=0, fg_color="#0f0f0f")
        bar.grid(row=3, column=0, columnspan=2, sticky="ew")

        for key, val in [("dot", "Idle"), ("mode", "Idle"), ("scan", "Last scan: Never"), ("err", "")]:
            self._sv[key] = ctk.StringVar(value=val)

        ctk.CTkLabel(bar, textvariable=self._sv["dot"],
                     font=ctk.CTkFont(size=11)).pack(side="left", padx=(10, 3), pady=5)
        ctk.CTkLabel(bar, textvariable=self._sv["mode"],
                     font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 18), pady=5)
        ctk.CTkLabel(bar, textvariable=self._sv["scan"],
                     font=ctk.CTkFont(size=11), text_color="gray55").pack(side="left", pady=5)
        ctk.CTkLabel(bar, textvariable=self._sv["err"],
                     font=ctk.CTkFont(size=11), text_color="#ff6b6b").pack(side="right", padx=12, pady=5)

        # --- Build and stack all pages ---
        for name, fn in [
            ("home",      self._page_home),
            ("analytics", self._page_analytics),
            ("outlook",   self._page_outlook),
            ("sheets",    self._page_sheets),
            ("sftp",      self._page_sftp),
            ("settings",  self._page_settings),
        ]:
            frame = fn()
            frame.grid(row=0, column=0, sticky="nsew", in_=self._content)
            self._pages[name] = frame

    # --- Navigation ---
    def _show(self, name: str):
        self._pages[name].tkraise()
        for key, btn in self._navbtns.items():
            active = key == name
            btn.configure(
                fg_color="#282828" if active else "transparent",
                text_color="white"  if active else "gray60",
            )

    # --- Page: Home ---
    def _page_home(self) -> ctk.CTkFrame:
        p = ctk.CTkFrame(self._content, fg_color="transparent")
        p.grid_columnconfigure(0, weight=1)
        p.grid_rowconfigure(3, weight=1)   # row 3 = activity log

        ctk.CTkLabel(
            p, text="Dashboard",
            font=ctk.CTkFont(size=22, weight="bold"),
        ).grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

        # --- Controls ---
        ctrl = ctk.CTkFrame(p, fg_color="transparent")
        ctrl.grid(row=1, column=0, padx=20, pady=(0, 14), sticky="w")

        self._sv["auto_label"] = ctk.StringVar(value="Start Auto Mode")
        btn_auto = ctk.CTkButton(
            ctrl, textvariable=self._sv["auto_label"],
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=_ACCENT, hover_color="#a00d24",
            width=190, height=38,
            command=self._toggle_auto,
        )
        btn_auto.pack(side="left", padx=(0, 10))
        self._btn_auto = btn_auto

        btn_run_once = ctk.CTkButton(
            ctrl, text="Run Once",
            fg_color="#2b2b2b", hover_color="#3a3a3a",
            width=130, height=38,
            command=self._run_once,
        )
        btn_run_once.pack(side="left", padx=(0, 10))
        self._btn_run_once = btn_run_once

        ctk.CTkButton(
            ctrl, text="Refresh Stats",
            fg_color="transparent", border_width=1, border_color="gray35",
            width=145, height=38,
            command=self._refresh_analytics,
        ).pack(side="left")

        # --- Last Scan Result card ---
        for key, default in [
            ("sr_status",   "No scan yet"),
            ("sr_time",     "-"),
            ("sr_aha",      "-"),
            ("sr_students", "-"),
            ("sr_acuity",   "-"),
            ("sr_duration", "-"),
        ]:
            self._sv[key] = ctk.StringVar(value=default)

        sr_card = ctk.CTkFrame(p, corner_radius=10, fg_color="#1a1a1a")
        sr_card.grid(row=2, column=0, padx=20, pady=(0, 14), sticky="ew")

        sr_hdr = ctk.CTkFrame(sr_card, fg_color="transparent")
        sr_hdr.pack(fill="x", padx=16, pady=(12, 4))
        ctk.CTkLabel(
            sr_hdr, text="Last Scan Result",
            font=ctk.CTkFont(size=13, weight="bold"), text_color="gray65",
        ).pack(side="left")
        ctk.CTkLabel(
            sr_hdr, textvariable=self._sv["sr_status"],
            font=ctk.CTkFont(size=12), anchor="w",
        ).pack(side="left", padx=(12, 0))
        ctk.CTkLabel(
            sr_hdr, textvariable=self._sv["sr_time"],
            font=ctk.CTkFont(size=11), text_color="gray50", anchor="e",
        ).pack(side="right")

        ctk.CTkFrame(sr_card, height=1, fg_color="#2e2e2e").pack(fill="x", padx=16, pady=(0, 8))

        sr_metrics = ctk.CTkFrame(sr_card, fg_color="transparent")
        sr_metrics.pack(fill="x", padx=16, pady=(0, 14))

        for col_i, (lbl, sv_key) in enumerate([
            ("AHA Emails",    "sr_aha"),
            ("Students Reg.", "sr_students"),
            ("Acuity Appts",  "sr_acuity"),
            ("Duration",      "sr_duration"),
        ]):
            cell = ctk.CTkFrame(sr_metrics, fg_color="transparent")
            cell.grid(row=0, column=col_i, padx=(0, 36))
            ctk.CTkLabel(cell, textvariable=self._sv[sv_key],
                         font=ctk.CTkFont(size=28, weight="bold")).pack()
            ctk.CTkLabel(cell, text=lbl,
                         font=ctk.CTkFont(size=10), text_color="gray45").pack()

        # --- Activity log ---
        log_wrap = ctk.CTkFrame(p, fg_color="transparent")
        log_wrap.grid(row=3, column=0, padx=20, pady=(0, 16), sticky="nsew")
        log_wrap.grid_rowconfigure(1, weight=1)
        log_wrap.grid_columnconfigure(0, weight=1)

        log_hdr = ctk.CTkFrame(log_wrap, fg_color="transparent")
        log_hdr.grid(row=0, column=0, sticky="ew", pady=(0, 5))

        ctk.CTkLabel(
            log_hdr, text="Activity Log",
            font=ctk.CTkFont(size=13, weight="bold"), text_color="gray60",
        ).pack(side="left")
        ctk.CTkButton(
            log_hdr, text="File  Open Log File",
            fg_color="transparent", border_width=1, border_color="gray35",
            width=138, height=26,
            font=ctk.CTkFont(size=11),
            command=self._open_log_file,
        ).pack(side="right")

        log_box = ctk.CTkTextbox(
            log_wrap, state="disabled", wrap="word",
            font=ctk.CTkFont(family="Courier New", size=11),
            fg_color="#0d0d0d", text_color="#c8c8c8",
        )
        log_box.grid(row=1, column=0, sticky="nsew")
        self._log_box = log_box

        return p

    # --- Page: Analytics ---
    def _page_analytics(self) -> ctk.CTkFrame:
        p = ctk.CTkFrame(self._content, fg_color="transparent")
        p.grid_columnconfigure(0, weight=1)
        p.grid_rowconfigure(1, weight=1)

        hdr = ctk.CTkFrame(p, fg_color="transparent")
        hdr.grid(row=0, column=0, padx=20, pady=(20, 12), sticky="ew")
        ctk.CTkLabel(hdr, text="Analytics",
                     font=ctk.CTkFont(size=22, weight="bold")).pack(side="left")
        ctk.CTkButton(
            hdr, text="Refresh",
            fg_color="transparent", border_width=1, border_color="gray35",
            width=110, height=32,
            command=self._refresh_analytics,
        ).pack(side="right")

        tv = ctk.CTkTabview(
            p,
            fg_color="#161616",
            segmented_button_fg_color="#111111",
            segmented_button_selected_color=_ACCENT,
            segmented_button_selected_hover_color="#a00d24",
            segmented_button_unselected_color="#111111",
            segmented_button_unselected_hover_color="#1e1e1e",
            text_color="gray70",
        )
        tv.grid(row=1, column=0, padx=20, pady=(0, 16), sticky="nsew")

        for tab_name in ("Total", "Daily", "Weekly"):
            tv.add(tab_name)

        self._build_total_tab(tv.tab("Total"))
        self._build_period_tab(tv.tab("Daily"),  "d_")
        self._build_period_tab(tv.tab("Weekly"), "w_")

        return p

    # --- Tab builders ---
    def _make_table_row(self, parent, label: str, sv_key: str, row_idx: int):
        """Create one alternating-color row in an analytics table."""
        self._sv[sv_key] = ctk.StringVar(value="-")
        bg = "#1e1e1e" if row_idx % 2 == 0 else "#252525"
        f = ctk.CTkFrame(parent, fg_color=bg, corner_radius=0)
        f.pack(fill="x")
        ctk.CTkLabel(
            f, text=label, anchor="w",
            font=ctk.CTkFont(size=12), text_color="gray65", width=230,
        ).pack(side="left", padx=(14, 0), pady=5)
        ctk.CTkLabel(
            f, textvariable=self._sv[sv_key], anchor="w",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(side="left", padx=(10, 14))

    @staticmethod
    def _table_header(parent, col1: str, col2: str):
        hdr = ctk.CTkFrame(parent, fg_color="#111111", corner_radius=0)
        hdr.pack(fill="x", pady=(0, 2))
        ctk.CTkLabel(hdr, text=col1, anchor="w",
                     font=ctk.CTkFont(size=11, weight="bold"), text_color="gray45",
                     width=230).pack(side="left", padx=(14, 0), pady=4)
        ctk.CTkLabel(hdr, text=col2, anchor="w",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="gray45").pack(side="left", padx=(10, 14))

    def _build_total_tab(self, parent: ctk.CTkFrame):
        self._table_header(parent, "Metric", "All-Time Total")
        rows = [
            ("AHA Students (Master Sheet)",  "tab_t_aha_students"),
            ("RQI Students",                 "tab_t_rqi_students"),
            ("Pending - New",                "tab_t_pending_new"),
            ("Pending - Changed",            "tab_t_pending_chg"),
            ("Uploaded to RQI",              "tab_t_uploaded"),
            ("Upcoming Appointments (7 d)",  "tab_t_upcoming"),
            ("Cross-Registered Students",    "tab_t_cross_reg"),
            ("Scans Run (session)",          "tab_t_scans"),
            ("Students Found (session)",     "tab_t_students_found"),
            ("Reminders Sent (session)",     "tab_t_reminders"),
            ("Top Course",                   "tab_t_top_course"),
        ]
        for i, (lbl, key) in enumerate(rows):
            self._make_table_row(parent, lbl, key, i)

    def _build_period_tab(self, parent: ctk.CTkFrame, prefix: str):
        period = "Today" if prefix == "d_" else "This Week"
        self._table_header(parent, "Metric", f"New {period}")
        rows = [
            (f"AHA Students (new {period.lower()})",  f"tab_{prefix}aha_students"),
            (f"RQI Students (new {period.lower()})",  f"tab_{prefix}rqi_students"),
            ("Pending New (change)",                  f"tab_{prefix}pending_new"),
            (f"Uploaded to RQI (new {period.lower()})", f"tab_{prefix}uploaded"),
            ("Scans Run (session)",                   f"tab_{prefix}scans"),
            ("Students Found (session)",              f"tab_{prefix}students_found"),
            ("Reminders Sent (session)",              f"tab_{prefix}reminders"),
        ]
        for i, (lbl, key) in enumerate(rows):
            self._make_table_row(parent, lbl, key, i)

    # --- Page: Outlook ---
    def _page_outlook(self) -> ctk.CTkFrame:
        p = ctk.CTkFrame(self._content, fg_color="transparent")
        p.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            p, text="Outlook Parser",
            font=ctk.CTkFont(size=22, weight="bold"),
        ).grid(row=0, column=0, padx=20, pady=(20, 16), sticky="w")

        # Scan status
        stat = self._make_section(p, "Scan Status", row=1)
        for key, label, default in [
            ("ol_last_scan",  "Last Scan",              "-"),
            ("ol_students",   "Students Found (total)",  "-"),
            ("ol_acuity",     "Acuity Appointments",     "-"),
            ("ol_scans",      "Total Scans (session)",   "0"),
            ("ol_errors",     "Consecutive Errors",      "0"),
        ]:
            self._sv[key] = ctk.StringVar(value=default)
            self._make_kv(stat, label, self._sv[key])

        # Actions
        act = self._make_section(p, "Actions", row=2)
        btn_row_ol = ctk.CTkFrame(act, fg_color="transparent")
        btn_row_ol.pack(anchor="w", padx=16, pady=(10, 4), fill="x")

        btn_scan_now = ctk.CTkButton(
            btn_row_ol, text="Scan Outlook Now",
            fg_color=_ACCENT, hover_color="#a00d24",
            width=190, height=36,
            command=self._run_scan_now,
        )
        btn_scan_now.pack(side="left", padx=(0, 10))
        self._btn_scan_now = btn_scan_now

        self._sv["ol_test_result"] = ctk.StringVar(value="")
        ctk.CTkButton(
            btn_row_ol, text="Test Connection",
            fg_color="#2b2b2b", hover_color="#3a3a3a",
            width=170, height=36,
            command=self._test_outlook,
        ).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(
            btn_row_ol, textvariable=self._sv["ol_test_result"],
            font=ctk.CTkFont(size=11), anchor="w",
        ).pack(side="left", fill="x", expand=True)

        ctk.CTkFrame(act, height=8, fg_color="transparent").pack()

        # Configuration reference
        env = cast(EnvDict, dotenv_values(str(_ENV)))
        cfg = self._make_section(p, "Current Configuration  (edit in Settings)", row=3)
        for label, key, default in [
            ("Outlook Email",    "OUTLOOK_EMAIL",         "-"),
            ("Acuity Sender",    "ACUITY_SENDER_EMAIL",   "-"),
            ("Scan Interval",    "SCAN_INTERVAL_SECONDS", "120 s"),
            ("Email Lookback",   "EMAIL_LOOKBACK_DAYS",   "7 days"),
        ]:
            self._make_kv(cfg, label, None, static=_env_text(env, key, default))

        return p

    # --- Page: Google Sheets ---
    def _page_sheets(self) -> ctk.CTkFrame:
        p = ctk.CTkFrame(self._content, fg_color="transparent")
        p.grid_columnconfigure(0, weight=1)

        hdr = ctk.CTkFrame(p, fg_color="transparent")
        hdr.grid(row=0, column=0, padx=20, pady=(20, 16), sticky="ew")
        ctk.CTkLabel(hdr, text="Google Sheets",
                     font=ctk.CTkFont(size=22, weight="bold")).pack(side="left")
        self._sv["sh_test_result"] = ctk.StringVar(value="")
        ctk.CTkLabel(
            hdr, textvariable=self._sv["sh_test_result"],
            font=ctk.CTkFont(size=11), anchor="w",
        ).pack(side="left", padx=(16, 0))
        ctk.CTkButton(hdr, text="Test Access", width=128, height=32,
                      fg_color="#2b2b2b", hover_color="#3a3a3a",
                      command=self._test_sheets).pack(side="right", padx=(8, 0))
        ctk.CTkButton(hdr, text="Refresh", width=115, height=32,
                      fg_color="transparent", border_width=1, border_color="gray35",
                      command=self._refresh_analytics).pack(side="right")

        # AHA Registration sheet
        aha = self._make_section(p, "AHA Registration Sheet", row=1)
        for key, label, default in [
            ("sh_total",     "Total Unique Students",   "-"),
            ("sh_uploaded",  "Uploaded to RQI",         "-"),
            ("sh_pend_new",  "Pending - New",            "-"),
            ("sh_pend_chg",  "Pending - Data Changed",   "-"),
            ("sh_recent",    "Most Recent Registration", "-"),
            ("sh_last_rqi",  "Last RQI Upload",          "Never"),
            ("sh_top",       "Top Course",               "-"),
            ("sh_courses",   "Students per Course",      "-"),
        ]:
            self._sv[key] = ctk.StringVar(value=default)
            self._make_kv(aha, label, self._sv[key])

        # Acuity sheet
        acuity = self._make_section(p, "Acuity Appointments Sheet (RQI)", row=2)
        for key, label, default in [
            ("ac_total",    "Total RQI Students",        "-"),
            ("ac_upcoming", "Upcoming (next 7 days)",    "-"),
            ("ac_cross",    "Cross-Registered Students", "-"),
            ("ac_r3d",      "3-Day Reminders Sent",      "-"),
            ("ac_r1d",      "1-Day Reminders Sent",      "-"),
            ("ac_r2hr",     "2-Hour Reminders Sent",     "-"),
        ]:
            self._sv[key] = ctk.StringVar(value=default)
            self._make_kv(acuity, label, self._sv[key])

        return p

    # --- Page: SFTP ---
    def _page_sftp(self) -> ctk.CTkFrame:
        p = ctk.CTkFrame(self._content, fg_color="transparent")
        p.grid_rowconfigure(1, weight=1)
        p.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            p, text="RQI Upload",
            font=ctk.CTkFont(size=22, weight="bold"),
        ).grid(row=0, column=0, padx=20, pady=(20, 16), sticky="w")

        body = ctk.CTkScrollableFrame(p, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=0, pady=(0, 16))
        body.grid_columnconfigure(0, weight=1)

        # Static connection info from .env
        env = cast(EnvDict, dotenv_values(str(_ENV)))
        conn = self._make_section(body, "Server Connection", row=1)
        for label, key, default in [
            ("Host",        "SFTP_HOST",       "-"),
            ("Port",        "SFTP_PORT",        "22"),
            ("Username",    "SFTP_USERNAME",    "-"),
            ("Remote Dir",  "SFTP_REMOTE_DIR",  "-"),
            ("Local Dir",   "SFTP_LOCAL_DIR",   "-"),
            ("Filename",    "SFTP_FILENAME",    "-"),
        ]:
            self._make_kv(conn, label, None, static=_env_text(env, key, default))

        # Dynamic upload status
        status = self._make_section(body, "Upload Status", row=2)
        for key, label, default in [
            ("sftp_last",    "Last Upload",            "Never"),
            ("sftp_count",   "Records in Last Delta",  "-"),
            ("sftp_pending", "Records Pending",        "-"),
            ("sftp_next",    "Next Upload Window",     "-"),
        ]:
            self._sv[key] = ctk.StringVar(value=default)
            self._make_kv(status, label, self._sv[key])

        # Keychain password status
        pw_section = self._make_section(body, "Keychain Password", row=3)
        pw_row = ctk.CTkFrame(pw_section, fg_color="transparent")
        pw_row.pack(fill="x", padx=16, pady=(6, 12))

        self._sv["sftp_keyring"] = ctk.StringVar(value="Checking...")
        self._sv["sftp_keyring_color"] = ctk.StringVar(value="gray55")
        ctk.CTkLabel(
            pw_row, text="Status:", width=80, anchor="w",
            font=ctk.CTkFont(size=12), text_color="gray55",
        ).pack(side="left")
        ctk.CTkLabel(
            pw_row, textvariable=self._sv["sftp_keyring"],
            font=ctk.CTkFont(size=12),
        ).pack(side="left", padx=(0, 20))
        ctk.CTkButton(
            pw_row, text="Set Up Password",
            fg_color="#2b2b2b", hover_color="#3a3a3a",
            width=148, height=30,
            command=self._open_sftp_setup,
        ).pack(side="left")

        # Actions
        act = self._make_section(body, "Actions", row=4)
        btn_row_sftp = ctk.CTkFrame(act, fg_color="transparent")
        btn_row_sftp.pack(anchor="w", padx=16, pady=(10, 4), fill="x")

        btn_sftp_now = ctk.CTkButton(
            btn_row_sftp, text="Upload Now",
            fg_color=_ACCENT, hover_color="#a00d24",
            width=160, height=36,
            command=self._run_sftp_now,
        )
        btn_sftp_now.pack(side="left", padx=(0, 10))
        self._btn_sftp_now = btn_sftp_now

        self._sv["sftp_test_result"] = ctk.StringVar(value="")
        ctk.CTkButton(
            btn_row_sftp, text="Test Connection",
            fg_color="#2b2b2b", hover_color="#3a3a3a",
            width=170, height=36,
            command=self._test_sftp,
        ).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(
            btn_row_sftp, textvariable=self._sv["sftp_test_result"],
            font=ctk.CTkFont(size=11), anchor="w",
        ).pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(
            act,
            text="  Uploads the current delta (new / changed records only) to the RQI SFTP server.",
            font=ctk.CTkFont(size=11), text_color="gray50",
        ).pack(anchor="w", padx=16, pady=(0, 12))

        return p

    # --- Page: Settings ---
    def _page_settings(self) -> ctk.CTkFrame:
        p = ctk.CTkFrame(self._content, fg_color="transparent")
        p.grid_rowconfigure(1, weight=1)
        p.grid_columnconfigure(0, weight=1)

        # Header + save button
        hdr = ctk.CTkFrame(p, fg_color="transparent")
        hdr.grid(row=0, column=0, padx=20, pady=(20, 8), sticky="ew")
        ctk.CTkLabel(hdr, text="Settings",
                     font=ctk.CTkFont(size=22, weight="bold")).pack(side="left")
        ctk.CTkButton(
            hdr, text="Save", width=110, height=36,
            fg_color=_ACCENT, hover_color="#a00d24",
            command=self._save_settings,
        ).pack(side="right")

        # Scrollable body
        body = ctk.CTkScrollableFrame(p, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 16))
        body.grid_columnconfigure(1, weight=1)

        env = cast(EnvDict, dotenv_values(str(_ENV)))
        cur_group = None
        row = 0

        for group, label, key, kind in _SCHEMA:
            # Group header
            if group != cur_group:
                cur_group = group
                ctk.CTkLabel(
                    body, text=group,
                    font=ctk.CTkFont(size=13, weight="bold"),
                    text_color=_ACCENT,
                ).grid(row=row, column=0, columnspan=3,
                       padx=0, pady=(18 if row else 6, 5), sticky="w")
                row += 1

            ctk.CTkLabel(
                body, text=label, anchor="w",
                font=ctk.CTkFont(size=12), text_color="gray65", width=185,
            ).grid(row=row, column=0, padx=(0, 12), pady=3, sticky="w")

            current = _env_text(env, key)

            if kind == "bool":
                var: tk.Variable = tk.BooleanVar(
                    value=current.strip().lower() in ("true", "1", "yes"))
                ctk.CTkSwitch(body, text="", variable=var,
                              onvalue=True, offvalue=False
                              ).grid(row=row, column=1, pady=3, sticky="w")

            elif kind.startswith("choice:"):
                choices = kind.split(":", 1)[1].split(",")
                var = ctk.StringVar(value=current if current in choices else choices[0])
                ctk.CTkOptionMenu(body, values=choices, variable=var, width=200
                                  ).grid(row=row, column=1, pady=3, sticky="w")

            elif kind == "keyring":
                # Keyring-backed password: show status + Set Password button.
                # NOT added to _svar so _save_settings never writes it to .env.
                sv_key = f"keyring_{key}"
                if sv_key not in self._sv:
                    self._sv[sv_key] = ctk.StringVar(value="Checking...")
                kr_row = ctk.CTkFrame(body, fg_color="transparent")
                kr_row.grid(row=row, column=1, columnspan=2, pady=3, sticky="ew")
                ctk.CTkLabel(
                    kr_row, textvariable=self._sv[sv_key],
                    font=ctk.CTkFont(size=12), anchor="w",
                ).pack(side="left", padx=(0, 14))
                ctk.CTkButton(
                    kr_row, text="Set Password", width=120, height=28,
                    fg_color="#2b2b2b", hover_color="#3a3a3a",
                    command=lambda k=key: self._open_keyring_dialog(k),
                ).pack(side="left")
                row += 1
                continue

            elif kind == "password":
                var = ctk.StringVar(value=current)
                entry = ctk.CTkEntry(body, textvariable=var, show="*",
                                     width=300, font=ctk.CTkFont(size=12))
                entry.grid(row=row, column=1, pady=3, sticky="ew")
                shown = False

                def _toggle_show(e=entry):
                    nonlocal shown
                    shown = not shown
                    e.configure(show="" if shown else "*")

                ctk.CTkButton(body, text="Show", width=60, height=28,
                              fg_color="#2b2b2b", hover_color="#3a3a3a",
                              command=_toggle_show,
                              ).grid(row=row, column=2, padx=(8, 0), pady=3)

            else:  # "text" or "int"
                var = ctk.StringVar(value=current)
                ctk.CTkEntry(body, textvariable=var,
                             width=300, font=ctk.CTkFont(size=12)
                             ).grid(row=row, column=1, pady=3, sticky="ew")

            self._svar[key] = var
            row += 1

        # Notes at the bottom
        ctk.CTkLabel(
            body,
            text=(
                "INFO   Outlook, Atlas, and SFTP passwords are stored in the OS keychain,\n"
                "    not in .env.  Use the 'Set Password' buttons above to configure them."
            ),
            font=ctk.CTkFont(size=11), text_color="gray50", justify="left",
        ).grid(row=row, column=0, columnspan=3, pady=(18, 4), sticky="w")
        row += 1

        ctk.CTkLabel(
            body,
            text="WARN   Restart the application after saving for all changes to take effect.",
            font=ctk.CTkFont(size=11), text_color="#e0a020", justify="left",
        ).grid(row=row, column=0, columnspan=3, pady=(0, 20), sticky="w")

        return p

    # --- UI helpers ---
    @staticmethod
    def _make_stat_card(parent, title: str, var: ctk.StringVar, sub: str) -> ctk.CTkFrame:
        f = ctk.CTkFrame(parent, corner_radius=10, fg_color="#1a1a1a")
        ctk.CTkLabel(f, text=title,
                     font=ctk.CTkFont(size=11), text_color="gray55").pack(pady=(16, 3))
        ctk.CTkLabel(f, textvariable=var,
                     font=ctk.CTkFont(size=34, weight="bold")).pack()
        ctk.CTkLabel(f, text=sub,
                     font=ctk.CTkFont(size=10), text_color="gray45").pack(pady=(2, 16))
        return f

    @staticmethod
    def _make_section(parent, title: str, row: int) -> ctk.CTkFrame:
        """Titled card that children can pack into."""
        card = ctk.CTkFrame(parent, corner_radius=10, fg_color="#1a1a1a")
        card.grid(row=row, column=0, padx=20, pady=(0, 12), sticky="ew")
        ctk.CTkLabel(card, text=title,
                     font=ctk.CTkFont(size=13, weight="bold"), text_color="gray65",
                     ).pack(anchor="w", padx=16, pady=(13, 5))
        ctk.CTkFrame(card, height=1, fg_color="#2e2e2e").pack(fill="x", padx=16, pady=(0, 6))
        return card

    @staticmethod
    def _make_kv(parent, label: str, var: ctk.StringVar | None, *, static: str = ""):
        """Key-value row packed inside a section card."""
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=16, pady=3)
        ctk.CTkLabel(row, text=label, anchor="w",
                     font=ctk.CTkFont(size=12), text_color="gray55",
                     width=200).pack(side="left")
        if var is not None:
            ctk.CTkLabel(row, textvariable=var, anchor="w",
                         font=ctk.CTkFont(size=12)).pack(side="left", fill="x", expand=True)
        else:
            ctk.CTkLabel(row, text=static, anchor="w",
                         font=ctk.CTkFont(size=12), text_color="gray80",
                         ).pack(side="left", fill="x", expand=True)

    # --- Logger hook ---
    def _hook_logger(self):
        q = self._lq

        class _GUIHandler(logging.Handler):
            def emit(self, record):
                q.put(self.format(record))

        h = _GUIHandler()
        h.setFormatter(logging.Formatter(
            "%(asctime)s  [%(levelname)-8s]  %(message)s", datefmt="%H:%M:%S"))
        logging.getLogger().addHandler(h)

    def _flush_logs(self):
        """Drain the log queue into the textbox (main-thread only)."""
        MAX_LINES = 600
        while True:
            try:
                msg = self._lq.get_nowait()
            except queue.Empty:
                break
            box = self._log_box
            if box is None:
                break
            box.configure(state="normal")
            box.insert("end", msg + "\n")
            total = int(box.index("end-1c").split(".")[0])
            if total > MAX_LINES:
                box.delete("1.0", f"{total - MAX_LINES}.0")
            box.see("end")
            box.configure(state="disabled")

    # --- Analytics ---
    def _refresh_analytics(self):
        """Kick off a background thread to fetch analytics (no-op if one is already running)."""
        if self._analytics_running:
            return
        self._analytics_running = True

        def _work():
            try:
                self._aq.put(("ok", get_analytics()))
            except OAuthExpiredError as exc:
                self._aq.put(("auth", str(exc)))
            except Exception as exc:
                self._aq.put(("err", str(exc)))
            finally:
                self._analytics_running = False

        threading.Thread(target=_work, daemon=True).start()

    def _apply_analytics(self, a: dict):
        now = datetime.now()
        today_str = now.date().isoformat()
        week_str  = (now.date() - timedelta(days=now.weekday())).isoformat()

        # Rotate daily baseline at midnight
        if self._analytics_baseline_date != today_str:
            self._analytics_daily_baseline = a.copy()
            self._analytics_baseline_date  = today_str

        # Rotate weekly baseline at start of week (Monday)
        if self._analytics_baseline_week != week_str:
            self._analytics_weekly_baseline = a.copy()
            self._analytics_baseline_week   = week_str

        # Initialise baselines on first call
        if self._analytics_daily_baseline is None:
            self._analytics_daily_baseline = a.copy()
            self._analytics_baseline_date  = today_str
        if self._analytics_weekly_baseline is None:
            self._analytics_weekly_baseline = a.copy()
            self._analytics_baseline_week   = week_str

        assert self._analytics_daily_baseline is not None
        assert self._analytics_weekly_baseline is not None
        self._save_analytics_baselines()

        def _delta(key: str, baseline: dict) -> str:
            curr = a.get(key, 0)
            base = baseline.get(key, curr)
            try:
                d = int(curr) - int(base)
                return f"+{d}" if d > 0 else str(d)
            except (TypeError, ValueError):
                return str(curr)

        db = self._analytics_daily_baseline
        wb = self._analytics_weekly_baseline

        # --- Total tab ---
        self._sv["tab_t_aha_students"].set(str(a["total_unique_students"]))
        self._sv["tab_t_rqi_students"].set(str(a["total_acuity_appointments"]))
        self._sv["tab_t_pending_new"].set(str(a["students_pending_new"]))
        self._sv["tab_t_pending_chg"].set(str(a["students_pending_changed"]))
        self._sv["tab_t_uploaded"].set(str(a["students_uploaded_rqi"]))
        self._sv["tab_t_upcoming"].set(str(a["upcoming_appointments_7d"]))
        self._sv["tab_t_cross_reg"].set(str(a["cross_registered"]))
        self._sv["tab_t_scans"].set(str(a["total_scans"]))
        self._sv["tab_t_students_found"].set(str(a["total_students_found"]))
        self._sv["tab_t_reminders"].set(str(a["total_reminders_sent"]))
        self._sv["tab_t_top_course"].set(a.get("top_course") or "-")

        # --- Daily tab (deltas since midnight) ---
        self._sv["tab_d_aha_students"].set(_delta("total_unique_students",    db))
        self._sv["tab_d_rqi_students"].set(_delta("total_acuity_appointments", db))
        self._sv["tab_d_pending_new"].set(_delta("students_pending_new",       db))
        self._sv["tab_d_uploaded"].set(_delta("students_uploaded_rqi",         db))
        self._sv["tab_d_scans"].set(str(a["total_scans"]))
        self._sv["tab_d_students_found"].set(str(a["total_students_found"]))
        self._sv["tab_d_reminders"].set(str(a["total_reminders_sent"]))

        # --- Weekly tab (deltas since Monday) ---
        self._sv["tab_w_aha_students"].set(_delta("total_unique_students",    wb))
        self._sv["tab_w_rqi_students"].set(_delta("total_acuity_appointments", wb))
        self._sv["tab_w_pending_new"].set(_delta("students_pending_new",       wb))
        self._sv["tab_w_uploaded"].set(_delta("students_uploaded_rqi",         wb))
        self._sv["tab_w_scans"].set(str(a["total_scans"]))
        self._sv["tab_w_students_found"].set(str(a["total_students_found"]))
        self._sv["tab_w_reminders"].set(str(a["total_reminders_sent"]))

        # --- Outlook page ---
        self._sv["ol_last_scan"].set(a["last_scan_time"])
        self._sv["ol_students"].set(str(a["total_students_found"]))
        self._sv["ol_acuity"].set(str(a["total_acuity_appointments"]))
        self._sv["ol_scans"].set(str(a["total_scans"]))
        self._sv["ol_errors"].set(str(a["consecutive_errors"]))

        # --- Sheets - AHA ---
        self._sv["sh_total"].set(str(a["total_unique_students"]))
        self._sv["sh_uploaded"].set(str(a["students_uploaded_rqi"]))
        self._sv["sh_pend_new"].set(str(a["students_pending_new"]))
        self._sv["sh_pend_chg"].set(str(a["students_pending_changed"]))
        self._sv["sh_recent"].set(a["most_recent_registration"] or "-")
        self._sv["sh_last_rqi"].set(a["last_sftp_upload"] or "Never")
        self._sv["sh_top"].set(a["top_course"] or "-")

        courses = a.get("students_per_course", {})
        if courses:
            top5 = sorted(courses.items(), key=lambda x: -x[1])[:5]
            self._sv["sh_courses"].set("   ".join(f"{c}: {n}" for c, n in top5))
        else:
            self._sv["sh_courses"].set("-")

        # --- Sheets - Acuity / RQI ---
        self._sv["ac_total"].set(str(a["total_acuity_appointments"]))
        self._sv["ac_upcoming"].set(str(a["upcoming_appointments_7d"]))
        self._sv["ac_cross"].set(str(a["cross_registered"]))
        self._sv["ac_r3d"].set(str(a["reminders_3d_sent"]))
        self._sv["ac_r1d"].set(str(a["reminders_1d_sent"]))
        self._sv["ac_r2hr"].set(str(a["reminders_2hr_sent"]))

        # --- SFTP ---
        self._sv["sftp_last"].set(a["last_sftp_upload_time"])
        self._sv["sftp_count"].set(str(a["last_delta_count"]))
        self._sv["sftp_pending"].set(str(a["students_pending_total"]))

        # --- Status bar ---
        self._sv["scan"].set(f"Last scan: {a['last_scan_time']}")
        errs = a["consecutive_errors"]
        self._sv["err"].set(f"WARN  {errs} consecutive error(s)" if errs > 0 else "")

    # --- Analytics baseline persistence ---
    def _load_analytics_baselines(self):
        try:
            with open(_ANALYTICS_CACHE) as f:
                cache = json.load(f)
            today_str = datetime.now().date().isoformat()
            week_str  = (datetime.now().date() - timedelta(days=datetime.now().weekday())).isoformat()
            if cache.get("daily_date") == today_str:
                self._analytics_daily_baseline = cache.get("daily_baseline")
                self._analytics_baseline_date  = today_str
            if cache.get("weekly_date") == week_str:
                self._analytics_weekly_baseline = cache.get("weekly_baseline")
                self._analytics_baseline_week   = week_str
        except (FileNotFoundError, json.JSONDecodeError, KeyError):
            pass

    def _save_analytics_baselines(self):
        try:
            with open(_ANALYTICS_CACHE, "w") as f:
                json.dump({
                    "daily_date":      self._analytics_baseline_date,
                    "daily_baseline":  self._analytics_daily_baseline,
                    "weekly_date":     self._analytics_baseline_week,
                    "weekly_baseline": self._analytics_weekly_baseline,
                }, f)
        except Exception:
            pass

    # --- Periodic tick (1 s) ---
    def _tick(self):
        self._flush_logs()

        # Drain analytics queue
        try:
            kind, payload = self._aq.get_nowait()
            if kind == "ok":
                self._apply_analytics(payload)
            elif kind == "auth":
                self._lq.put(
                    "Google OAuth token expired - click 'Re-authenticate' "
                    "in the banner to open a fresh browser sign-in."
                )
                self._show_auth_banner()
            elif kind == "err":
                self._lq.put(f"WARN  Analytics error: {payload}")
        except queue.Empty:
            pass

        # Update SFTP next-window every tick (it changes each minute)
        if "sftp_next" in self._sv:
            self._sv["sftp_next"].set(self._next_sftp_window())

        # Derive UI state
        auto_running = self._thread is not None and self._thread.is_alive()
        can_act      = not auto_running and not self._busy

        # Live step progress or countdown
        step = bot.get_scan_step()
        if step:
            # Actively scanning - show current step
            self._sv["scan"].set(step)
        next_scan_time = bot.get_next_scan_time()
        if not step and auto_running and next_scan_time is not None:
            # Between scans - show countdown
            remaining = (next_scan_time - datetime.now()).total_seconds()
            if remaining > 0:
                mins, secs = divmod(int(remaining), 60)
                self._sv["scan"].set(f"Next scan in {mins}m {secs:02d}s")
            else:
                self._sv["scan"].set("Scanning...")

        # Sync scan result summary card whenever bot finishes a new cycle
        last_scan_result = bot.get_last_scan_result()
        bot_ts = last_scan_result.get("ts")
        if bot_ts is not None and bot_ts != self._last_scan_ts:
            self._last_scan_ts = bot_ts
            self._apply_scan_summary(last_scan_result)

        # Status bar dot + mode
        if auto_running:
            self._sv["dot"].set("RUN")
            self._sv["mode"].set("Auto Mode Running")
        elif self._busy:
            self._sv["dot"].set("BUSY")
            self._sv["mode"].set("Running...")
        else:
            self._sv["dot"].set("IDLE")
            self._sv["mode"].set("Idle")

        # Auto-mode button label
        if auto_running:
            self._sv["auto_label"].set("Stop Auto Mode")
            if self._btn_auto:
                self._btn_auto.configure(fg_color="#3a3a3a", hover_color="#4a4a4a")
        else:
            self._sv["auto_label"].set("Start Auto Mode")
            if self._btn_auto:
                self._btn_auto.configure(fg_color=_ACCENT, hover_color="#a00d24")

        # Enable / disable manual trigger buttons
        state = "normal" if can_act else "disabled"
        for btn in (self._btn_run_once, self._btn_scan_now, self._btn_sftp_now):
            if btn:
                btn.configure(state=state)

        # Refresh analytics every 30 s
        self._tick_count += 1
        if self._tick_count % 30 == 0:
            self._refresh_analytics()

        self.after(1000, self._tick)

    def add_log(self, message: str) -> None:
        """Queue a message for the Activity Log."""
        self._lq.put(message)

    # --- Bot controls ---
    def _toggle_auto(self):
        if self._thread and self._thread.is_alive():
            self._stop.set()
        else:
            self._stop.clear()
            interval = int(os.getenv("SCAN_INTERVAL_SECONDS", "120"))
            thread = threading.Thread(
                target=auto_mode,
                kwargs={"scan_interval_seconds": interval, "stop_event": self._stop},
                daemon=True,
            )
            self._thread = thread
            thread.start()

    def _run_once(self):
        if self._busy or (self._thread and self._thread.is_alive()):
            return
        self._busy = True

        def _work():
            bot.run_once(stop_event=self._stop)
            self.after(0, self._on_done)

        threading.Thread(target=_work, daemon=True).start()

    def _run_scan_now(self):
        if self._busy or (self._thread and self._thread.is_alive()):
            return
        self._busy = True

        def _work():
            run_scan(stop_event=self._stop)
            self.after(0, self._on_done)

        threading.Thread(target=_work, daemon=True).start()

    def _run_sftp_now(self):
        if self._busy or (self._thread and self._thread.is_alive()):
            return
        self._busy = True

        def _work():
            sftp_upload_sheet()
            self.after(0, self._on_done)

        threading.Thread(target=_work, daemon=True).start()

    def _on_done(self):
        self._busy = False
        self._refresh_analytics()

    # --- Test-connection handlers ---
    def _run_test(self, test_fn, sv_key: str):
        """
        Run *test_fn* in a background thread; write OK/ERROR result to *sv_key*.
        Updates the StringVar with "Testing..." immediately, then the result.
        """
        if sv_key in self._sv:
            self._sv[sv_key].set("WAIT  Testing...")

        def _work():
            ok, msg = test_fn()
            icon = "OK" if ok else "ERROR"
            result = f"{icon}  {msg}"
            def _set_result() -> None:
                if sv_key in self._sv:
                    self._sv[sv_key].set(result)

            self.after(0, _set_result)
            self._lq.put(f"{'OK' if ok else 'ERROR'}  Test [{sv_key.split('_test')[0]}]: {msg}")

        threading.Thread(target=_work, daemon=True).start()

    def _test_outlook(self):
        self._run_test(test_outlook_connection, "ol_test_result")

    def _test_sheets(self):
        self._run_test(test_sheets_connection, "sh_test_result")

    def _test_sftp(self):
        self._run_test(test_sftp_connection, "sftp_test_result")

    # --- Scan result summary ---
    def _apply_scan_summary(self, result: dict):
        """Update the Last Scan Result card from a scan-result dict."""
        ok  = result.get("ok", False)
        ts = result.get("ts")
        ts_dt = ts if isinstance(ts, datetime) else None
        dur = result.get("duration_s", 0)

        if ok:
            self._sv["sr_status"].set("OK  Completed")
            self._dismiss_fail_banner()          # clear any previous failure banner
        else:
            err = result.get("error", "Unknown error")
            # Truncate long error text to keep the card readable
            short = err[:60] + "..." if len(err) > 60 else err
            self._sv["sr_status"].set(f"ERROR  Failed - {short}")
            # Show the failure banner with a plain-English message
            self._sv["fail_msg"].set(
                f"Last scan failed - {short}  -  Check the Activity Log for details."
            )
            if self._scan_fail_banner:
                self._scan_fail_banner.grid()

        self._sv["sr_time"].set(
            ts_dt.strftime("%H:%M:%S  %b %d") if ts_dt else "-"
        )
        self._sv["sr_aha"].set(str(result.get("aha_emails", "-")))
        self._sv["sr_students"].set(str(result.get("students", "-")))
        self._sv["sr_acuity"].set(str(result.get("acuity", "-")))

        if dur >= 60:
            mins, secs = divmod(int(dur), 60)
            self._sv["sr_duration"].set(f"{mins}m {secs:02d}s")
        else:
            self._sv["sr_duration"].set(f"{dur:.1f}s")

    def _dismiss_fail_banner(self):
        """Hide the scan failure banner."""
        if self._scan_fail_banner:
            self._scan_fail_banner.grid_remove()

    # --- Log file ---
    def _open_log_file(self):
        """Open aha_bot.log in the default OS text viewer."""
        log_path = bot.get_log_file()
        if os.path.exists(log_path):
            os.startfile(log_path)
        else:
            self._lq.put(f"WARN  Log file not found: {log_path}")

    # --- Credential / keyring helpers ---
    def _check_all_credentials(self):
        """
        Called once ~1.2 s after startup.
        Updates all keyring status labels and auto-opens the SFTP dialog if its
        password is missing (SFTP is silent-fail; Outlook/Atlas fail loudly).
        """
        # SFTP password
        sftp_ok = sftp_keyring_configured()
        self._update_keyring_sv("sftp_keyring", sftp_ok, "SFTP upload will be skipped")
        if not sftp_ok:
            self._lq.put(
                "WARN  SFTP password not found in OS keychain - "
                "opening setup dialog (or go to the SFTP page to set it)."
            )
            self._open_sftp_setup()

        # SFTP host key
        def _hk_check():
            known, key_type, fingerprint = check_sftp_host_key()
            if not known:
                self._lq.put(
                    f"WARN  SFTP host key for {bot.SFTP_HOST}:{bot.SFTP_PORT} not in "
                    "known_hosts - opening verification dialog."
                )
                self.after(0, lambda: _HostKeyDialog(self, key_type, fingerprint))

        threading.Thread(target=_hk_check, daemon=True).start()

        # Outlook
        ol_ok = outlook_keyring_configured()
        self._update_keyring_sv("keyring_OUTLOOK_PASSWORD", ol_ok, "not set")
        if not ol_ok:
            self._lq.put("WARN  Outlook password not configured - scans will fail at login.")

        # Atlas
        at_ok = atlas_keyring_configured()
        self._update_keyring_sv("keyring_ATLAS_PASSWORD", at_ok, "not set")
        if not at_ok:
            self._lq.put("WARN  Atlas password not configured - scans will fail at login.")

    def _update_keyring_sv(self, sv_key: str, configured: bool, warn_suffix: str):
        """Update a keyring status StringVar if it exists."""
        if sv_key in self._sv:
            if configured:
                self._sv[sv_key].set("OK  Configured (keychain)")
            else:
                self._sv[sv_key].set(f"WARN  Not in keychain - {warn_suffix}")

    def _open_keyring_dialog(self, key: str):
        """Open the password setup dialog for *key* (OUTLOOK_PASSWORD or ATLAS_PASSWORD)."""
        env = cast(EnvDict, dotenv_values(str(_ENV)))
        if key == "OUTLOOK_PASSWORD":
            _PasswordSetupDialog(
                self,
                title="Outlook Password Setup",
                header="Outlook Keychain Setup",
                description=(
                    "Your Outlook password will be stored in the OS keychain.\n"
                    "It is never written to .env or any file."
                ),
                pw_label="Outlook Password",
                info_rows=[
                    ("Service",  "aha-outlook"),
                    ("Email",    _env_text(env, "OUTLOOK_EMAIL", "-")),
                ],
                save_fn=set_outlook_password,
                on_success_cb=self._on_outlook_password_set,
            )
        elif key == "ATLAS_PASSWORD":
            _PasswordSetupDialog(
                self,
                title="Atlas Password Setup",
                header="Atlas Keychain Setup",
                description=(
                    "Your Atlas password will be stored in the OS keychain.\n"
                    "It is never written to .env or any file."
                ),
                pw_label="Atlas Password",
                info_rows=[
                    ("Service",  "aha-atlas"),
                    ("Email",    _env_text(env, "ATLAS_EMAIL", "-")),
                ],
                save_fn=set_atlas_password,
                on_success_cb=self._on_atlas_password_set,
            )

    def _open_sftp_setup(self):
        """Open the SFTP password setup dialog."""
        env = cast(EnvDict, dotenv_values(str(_ENV)))
        _PasswordSetupDialog(
            self,
            title="SFTP Password Setup",
            header="SFTP Keychain Setup",
            description=(
                "The SFTP upload needs a password stored in the OS keychain.\n"
                "Enter it once here - it is never written to .env or any file."
            ),
            pw_label="SFTP Password",
            info_rows=[
                ("Service",  _env_text(env, "SFTP_KEYRING_SERVICE", "rqi-sftp")),
                ("Username", _env_text(env, "SFTP_USERNAME", "-")),
                ("Host",     _env_text(env, "SFTP_HOST", "-")),
            ],
            save_fn=set_sftp_password,
            on_success_cb=self._on_sftp_password_set,
        )

    def _on_sftp_password_set(self):
        self._update_keyring_sv("sftp_keyring", True, "")
        self._lq.put("OK  SFTP password saved to OS keychain - upload is now enabled.")

    def _on_outlook_password_set(self):
        self._update_keyring_sv("keyring_OUTLOOK_PASSWORD", True, "")
        self._lq.put("OK  Outlook password saved to OS keychain.")

    def _on_atlas_password_set(self):
        self._update_keyring_sv("keyring_ATLAS_PASSWORD", True, "")
        self._lq.put("OK  Atlas password saved to OS keychain.")

    # --- OAuth re-authentication ---
    def _show_auth_banner(self):
        """Make the OAuth expired banner visible."""
        if self._oauth_banner:
            self._oauth_banner.grid()

    def _hide_auth_banner(self):
        """Hide the OAuth expired banner."""
        if self._oauth_banner:
            self._oauth_banner.grid_remove()

    def _do_reauth(self):
        """
        Start a fresh OAuth browser flow in a background thread.
        Disables the button while in-progress, re-enables on completion.
        """
        if self._btn_reauth:
            self._btn_reauth.configure(state="disabled", text="Authenticating...")
        self._lq.put("Opening browser for Google re-authentication ...")

        def _work():
            success = reauthenticate()
            self.after(0, self._on_reauth_done, success)

        threading.Thread(target=_work, daemon=True).start()

    def _on_reauth_done(self, success: bool):
        if self._btn_reauth:
            self._btn_reauth.configure(state="normal", text="Re-authenticate")
        if success:
            self._hide_auth_banner()
            self._lq.put("OK  Re-authentication successful - Google Sheets access restored.")
            self._refresh_analytics()
        else:
            self._lq.put(
                "ERROR  Re-authentication failed - ensure the browser sign-in completed "
                "and try again, or check that credentials.json is valid."
            )

    # --- Settings save ---
    def _save_settings(self):
        from tkinter import messagebox
        try:
            for key, var in self._svar.items():
                raw = var.get()
                value = ("true" if raw else "false") if isinstance(raw, bool) else str(raw)
                set_key(str(_ENV), key, value)
            messagebox.showinfo(
                "Settings Saved",
                "Settings written to .env.\n\nRestart the application for all changes to take effect.",
            )
        except PermissionError:
            messagebox.showerror(
                "Save Failed",
                "Could not write to .env - the file may be locked by OneDrive sync.\n\n"
                "Wait a moment for OneDrive to finish syncing, then try saving again.",
            )
        except Exception as exc:
            messagebox.showerror("Save Failed", f"Unexpected error writing .env:\n{exc}")

    # --- SFTP window calculation ---
    @staticmethod
    def _next_sftp_window() -> str:
        """Return HH:MM of the next :12/:27/:42/:57 upload target."""
        now  = datetime.now()
        m    = now.minute % 15
        wait = (12 - m) if m < 12 else (15 - m + 12)
        if wait == 0 and now.second > 30:
            wait = 15
        nxt  = now + timedelta(minutes=wait, seconds=-now.second)
        return nxt.strftime("%H:%M")

    # --- Window close ---
    def _on_close(self):
        self._stop.set()
        self.destroy()


# --- Entry point ---
if __name__ == "__main__":
    app = App()
    app.mainloop()


