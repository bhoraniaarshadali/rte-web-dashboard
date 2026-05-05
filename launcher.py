import sys
import os
import time
import shutil
import threading
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox

# ── Paths ────────────────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    # Running as .exe — executable location
    BASE_DIR = os.path.dirname(sys.executable)
    # PyInstaller extracts bundled files to _MEIPASS temp folder
    BUNDLE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    BUNDLE_DIR = BASE_DIR

os.chdir(BASE_DIR)
sys.path.insert(0, BASE_DIR)

# ── Extract bundled HTML files next to .exe on first run ─────────
def extract_assets():
    for fname in ["dashboard.html", "analysis.html"]:
        src = os.path.join(BUNDLE_DIR, fname)
        dst = os.path.join(BASE_DIR, fname)
        if os.path.exists(src) and not os.path.exists(dst):
            shutil.copy2(src, dst)

extract_assets()

DASHBOARD_PATH = os.path.join(BASE_DIR, "dashboard.html")
SYNC_PORT = 5001

# ── GUI Launcher Window ──────────────────────────────────────────
class LauncherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RTE Live Dashboard")
        self.root.geometry("420x280")
        self.root.resizable(False, False)
        self.root.configure(bg="#1e3a5f")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Try to center window
        root.update_idletasks()
        x = (root.winfo_screenwidth() - 420) // 2
        y = (root.winfo_screenheight() - 280) // 2
        root.geometry(f"420x280+{x}+{y}")

        self._build_ui()
        self.server_thread = None
        self.running = False

        # Auto-start on launch
        self.root.after(500, self.start)

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self.root, bg="#1e3a5f")
        hdr.pack(fill="x", pady=(18, 0))

        tk.Label(hdr, text="🏫 RTE Live Dashboard",
                 font=("Segoe UI", 16, "bold"),
                 fg="white", bg="#1e3a5f").pack()
        tk.Label(hdr, text="Shaikh ul Islam Trust, Surat",
                 font=("Segoe UI", 10),
                 fg="#93c5fd", bg="#1e3a5f").pack(pady=(2, 0))

        # Status
        mid = tk.Frame(self.root, bg="#1e3a5f")
        mid.pack(fill="x", padx=28, pady=18)

        self.status_var = tk.StringVar(value="⏳ Starting...")
        self.status_lbl = tk.Label(mid, textvariable=self.status_var,
                                   font=("Segoe UI", 10),
                                   fg="#fbbf24", bg="#1e3a5f",
                                   wraplength=360, justify="left")
        self.status_lbl.pack(anchor="w")

        # Progress bar
        style = ttk.Style()
        style.theme_use("default")
        style.configure("blue.Horizontal.TProgressbar",
                        troughcolor="#0f2744", background="#3b82f6",
                        thickness=8)
        self.pbar = ttk.Progressbar(mid, style="blue.Horizontal.TProgressbar",
                                    mode="indeterminate", length=364)
        self.pbar.pack(pady=(10, 0))

        # Buttons
        btn_frame = tk.Frame(self.root, bg="#1e3a5f")
        btn_frame.pack(fill="x", padx=28, pady=8)

        self.open_btn = tk.Button(btn_frame, text="🌐 Open Dashboard",
                                  font=("Segoe UI", 10, "bold"),
                                  bg="#2563eb", fg="white",
                                  relief="flat", padx=14, pady=8,
                                  cursor="hand2",
                                  command=self.open_browser)
        self.open_btn.pack(side="left")

        tk.Button(btn_frame, text="✕ Stop & Exit",
                  font=("Segoe UI", 10),
                  bg="#475569", fg="white",
                  relief="flat", padx=14, pady=8,
                  cursor="hand2",
                  command=self.on_close).pack(side="right")

        # Footer
        tk.Label(self.root,
                 text=f"Server: http://localhost:{SYNC_PORT}",
                 font=("Segoe UI", 8),
                 fg="#64748b", bg="#1e3a5f").pack(pady=(0, 8))

    def set_status(self, msg, color="#fbbf24"):
        self.status_var.set(msg)
        self.status_lbl.config(fg=color)

    def start(self):
        self.pbar.start(12)
        self.set_status("⏳ Starting sync server...", "#fbbf24")
        self.server_thread = threading.Thread(target=self._run_server, daemon=True)
        self.server_thread.start()

    def _run_server(self):
        try:
            # Import and run rte_checker main (which starts server)
            import rte_checker
            # Patch to not block on main loop – just start server + data
            self.root.after(0, self._on_server_ready)
            rte_checker.main()
        except Exception as e:
            self.root.after(0, lambda: self.set_status(f"❌ Error: {e}", "#ef4444"))
            self.root.after(0, self.pbar.stop)

    def _on_server_ready(self):
        self.pbar.stop()
        self.pbar.config(mode="determinate", value=100)
        self.set_status("✅ Server running! Dashboard is open in browser.", "#4ade80")
        # Open browser
        self.root.after(1500, self.open_browser)

    def open_browser(self):
        if os.path.exists(DASHBOARD_PATH):
            webbrowser.open(f"file:///{DASHBOARD_PATH.replace(os.sep, '/')}")
        else:
            messagebox.showerror("Error", f"dashboard.html not found at:\n{DASHBOARD_PATH}")

    def on_close(self):
        if messagebox.askyesno("Exit", "Server band karna chahte hain?"):
            self.root.destroy()
            os._exit(0)


def main():
    root = tk.Tk()
    app = LauncherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
