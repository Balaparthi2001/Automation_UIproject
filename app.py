from flask import Flask, request, jsonify, send_from_directory, Response, stream_with_context
from pathlib import Path
import os
import time
import subprocess
import signal
import sys
import socket
import webbrowser
import threading
import shutil
import ctypes
from ctypes import wintypes

# --------------------------------------------------------------------------------------
# Paths
# --------------------------------------------------------------------------------------
BASE_DIR   = Path(__file__).parent.resolve()
AUTO_DIR   = BASE_DIR / "automations"
UPLOAD_DIR = BASE_DIR / "uploads"
LOG_DIR    = BASE_DIR / "logs"
STATIC_DIR = BASE_DIR / "static"
TMP_DIR    = BASE_DIR / "tmp_runs"   # ephemeral edited copies

# Ensure folders exist
for d in (AUTO_DIR, UPLOAD_DIR, LOG_DIR, TMP_DIR):
    d.mkdir(parents=True, exist_ok=True)

# --------------------------------------------------------------------------------------
# Optional: pdf2image / pytesseract runtime tools on PATH (if bundled)
# --------------------------------------------------------------------------------------
POPPLER_BIN = BASE_DIR / "poppler-25.12.0" / "Library" / "bin"
if POPPLER_BIN.exists():
    os.environ["PATH"] = str(POPPLER_BIN.resolve()) + os.pathsep + os.environ.get("PATH", "")

TESSERACT_DIR = BASE_DIR / "Tesseract-OCR"
if TESSERACT_DIR.exists():
    os.environ["PATH"] = str(TESSERACT_DIR.resolve()) + os.pathsep + os.environ.get("PATH", "")

# --------------------------------------------------------------------------------------
# Flask app
# --------------------------------------------------------------------------------------
app = Flask(__name__, static_folder=str(STATIC_DIR), static_url_path="/static")

# Single-user local state
STATE = {
    "proc":  None,    # subprocess.Popen
    "file":  None,    # str (relative path under automations/ or tmp_runs/)
    "start": None,    # float (epoch seconds)
    "scope": None,    # "auto" or "tmp"
}

# --------------------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------------------
def _is_within(child: Path, parent: Path) -> bool:
    """Return True if 'child' is inside 'parent' (after resolving)."""
    try:
        child.resolve().relative_to(parent.resolve())
        return True
    except Exception:
        return False

# ---- Windows-only: bring a process window (e.g., Tkinter) to front/topmost ----------
def bring_to_front_windows(pid: int, retries: int = 20, delay: float = 0.3):
    if os.name != "nt":
        return
    user32 = ctypes.windll.user32

    EnumWindows = user32.EnumWindows
    EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    GetWindowThreadProcessId = user32.GetWindowThreadProcessId
    IsWindowVisible = user32.IsWindowVisible
    SetForegroundWindow = user32.SetForegroundWindow
    ShowWindow = user32.ShowWindow
    SetWindowPos = user32.SetWindowPos

    SW_SHOWNORMAL = 1
    HWND_TOPMOST = -1
    HWND_NOTOPMOST = -2
    SWP_NOSIZE = 0x0001
    SWP_NOMOVE = 0x0002

    target_hwnds = []

    def callback(hwnd, _lparam):
        if IsWindowVisible(hwnd):
            lpdwProcessId = wintypes.DWORD()
            GetWindowThreadProcessId(hwnd, ctypes.byref(lpdwProcessId))
            if lpdwProcessId.value == pid:
                target_hwnds.append(hwnd)
        return True

    for _ in range(retries):
        target_hwnds.clear()
        EnumWindows(EnumWindowsProc(callback), 0)
        if target_hwnds:
            try:
                hwnd = target_hwnds[0]
                ShowWindow(hwnd, SW_SHOWNORMAL)
                SetForegroundWindow(hwnd)
                # Toggle topmost to force z-order bump
                SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE)
                time.sleep(0.05)
                SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE)
            except Exception:
                pass
            return
        time.sleep(delay)

def reap_old_tmp(seconds: int = 3600):
    """Remove temp edited .py files older than N seconds."""
    now = time.time()
    for p in TMP_DIR.glob("*.py"):
        try:
            if now - p.stat().st_mtime > seconds:
                p.unlink(missing_ok=True)
        except Exception:
            pass

def pick_python(gui: bool = False) -> str:
    """
    Return an interpreter to run child .py files.

    - Dev (not frozen): use the current interpreter (sys.executable).
    - Frozen (PyInstaller): prefer embedded/portable Python in BASE_DIR/py/.
      WinPython layout: BASE_DIR/py/python/python.exe (and pythonw.exe)
      Flat layout   : BASE_DIR/py/python.exe (and pythonw.exe)
    - If embedded not found, try any system python on PATH.
    - Last resort, return sys.executable.
    """
    if getattr(sys, "frozen", False):  # PyInstaller EXE
        py_dir   = BASE_DIR / "py"
        py_home  = py_dir / "python"   # WinPython layout

        if os.name == "nt":
            # Prefer GUI-less when capturing stdout (gui=False); else pythonw.exe
            if gui:
                cand = py_home / "pythonw.exe"
                if cand.exists():
                    return str(cand)
            cand = py_home / "python.exe"
            if cand.exists():
                return str(cand)

            # flat layout fallback
            if gui:
                cand = py_dir / "pythonw.exe"
                if cand.exists():
                    return str(cand)
            cand = py_dir / "python.exe"
            if cand.exists():
                return str(cand)
        else:
            cand = py_dir / "bin" / "python3"
            if cand.exists():
                return str(cand)

        # Fallback: any python on PATH
        for name in ("python.exe", "python3.exe", "python3", "python"):
            p = shutil.which(name)
            if p:
                return p

        return sys.executable

    # Not frozen: use the current venv/interpreter
    return sys.executable

# --------------------------------------------------------------------------------------
# Routes
# --------------------------------------------------------------------------------------
@app.route("/")
def home():
    return send_from_directory(STATIC_DIR, "index.html")

# ---------- Files: flat list (relative paths) ----------
@app.get("/files")
def list_files_flat():
    files = []
    # root (uncategorized)
    for f in AUTO_DIR.glob("*.py"):
        files.append(f.name)
    # subfolders
    for f in AUTO_DIR.rglob("*.py"):
        if f.parent == AUTO_DIR:
            continue
        rel = f.relative_to(AUTO_DIR).as_posix()
        files.append(rel)
    files = sorted(set(files))
    return jsonify(files)

# ---------- Files: tree by category ----------
@app.get("/files/tree")
def list_files_tree():
    tree = {}
    # Uncategorized
    for f in AUTO_DIR.glob("*.py"):
        tree.setdefault("Uncategorized", []).append(f.name)
    # Subfolders
    for f in AUTO_DIR.rglob("*.py"):
        if f.parent == AUTO_DIR:
            continue
        parts = f.relative_to(AUTO_DIR).parts
        cat = parts[0] if len(parts) > 1 else "Uncategorized"
        tree.setdefault(cat, []).append("/".join(parts))
    # Sort result
    out = [{"category": cat, "files": sorted(items)} for cat, items in tree.items()]
    out.sort(key=lambda x: x["category"].lower())
    return jsonify(out)

# ---------- File content (for editor) ----------
@app.get("/files/content")
def file_content():
    rel = (request.args.get("path") or "").strip()
    if not rel:
        return jsonify({"ok": False, "error": "path required"}), 400
    target = (AUTO_DIR / rel).resolve()
    if not _is_within(target, AUTO_DIR) or not target.exists():
        return jsonify({"ok": False, "error": "not found"}), 404
    try:
        text = target.read_text(encoding="utf-8", errors="replace")
        return jsonify({"ok": True, "content": text})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

# ---------- Upload new (optionally into category) ----------
@app.post("/upload")
def upload():
    if "file" not in request.files or "name" not in request.form:
        return jsonify({"ok": False, "error": "file and name required"}), 400

    f        = request.files["file"]
    name     = (request.form["name"] or "").strip()
    category = (request.form.get("category") or "").strip()

    if not name.endswith(".py"):
        return jsonify({"ok": False, "error": "name must end with .py"}), 400

    # If name already includes a path, that wins; else, use category
    if "/" in name or "\\" in name:
        dest = (AUTO_DIR / name).resolve()
    else:
        dest = ((AUTO_DIR / category / name).resolve() if category else (AUTO_DIR / name).resolve())

    if not _is_within(dest, AUTO_DIR):
        return jsonify({"ok": False, "error": "invalid destination"}), 400
    if dest.exists():
        return jsonify({"ok": False, "error": "file exists. use update"}), 409

    dest.parent.mkdir(parents=True, exist_ok=True)
    f.save(dest)
    return jsonify({"ok": True, "name": dest.relative_to(AUTO_DIR).as_posix()})

# ---------- Update existing (target path relative to automations/) ----------
@app.post("/update")
def update():
    if "file" not in request.files or "target" not in request.form:
        return jsonify({"ok": False, "error": "file and target required"}), 400

    f          = request.files["file"]
    target_rel = (request.form["target"] or "").strip()
    dest       = (AUTO_DIR / target_rel).resolve()

    if not _is_within(dest, AUTO_DIR) or not dest.exists():
        return jsonify({"ok": False, "error": "target not found"}), 404

    f.save(dest)
    return jsonify({"ok": True, "target": target_rel})

# ---------- Delete (via query param to safely carry slashes) ----------
@app.delete("/files")
def delete_file():
    rel = (request.args.get("path") or "").strip()
    if not rel:
        return jsonify({"ok": False, "error": "path required"}), 400

    dest = (AUTO_DIR / rel).resolve()
    if not _is_within(dest, AUTO_DIR) or not dest.exists():
        return jsonify({"ok": False, "error": "not found"}), 404

    # Prevent deleting a currently-running file
    if STATE["file"] == rel and STATE["proc"] and STATE["proc"].poll() is None:
        return jsonify({"ok": False, "error": "file is running"}), 409

    dest.unlink()
    return jsonify({"ok": True})

# ---------- Run (SSE stream) ----------
@app.get("/run/stream")
def run_stream():
    """
    Query params:
      - file  : relative path under automations/ or tmp_runs/
      - scope : 'auto' (default) or 'tmp' indicating base directory
    """
    rel   = (request.args.get("file") or "").strip()
    scope = (request.args.get("scope") or "auto").strip() or "auto"

    if not rel:
        return jsonify({"error": "file param required"}), 400

    base   = AUTO_DIR if scope == "auto" else TMP_DIR
    script = (base / rel).resolve()

    if not _is_within(script, base) or not script.exists():
        return jsonify({"error": f"script not found: {rel}"}), 404

    _terminate_proc()
    log_path = LOG_DIR / f"{Path(rel).stem}_{int(time.time())}.log"

    def generate():
        # Force unbuffered Python for .py files
        py_path = pick_python(gui=False)
        cmd = [py_path]
        if script.suffix.lower() == ".py":
            cmd.append("-u")
        cmd.append(str(script))

        # Child environment: keep stdout unbuffered and UTF-8
        child_env = os.environ.copy()
        child_env.setdefault("PYTHONUNBUFFERED", "1")
        child_env.setdefault("PYTHONIOENCODING", "utf-8")

        # WinPython: help Tkinter resolve in edge cases
        if getattr(sys, "frozen", False):
            py_home = BASE_DIR / "py" / "python"    # WinPython layout
            if py_home.exists():
                child_env.setdefault("PYTHONHOME", str(py_home.resolve()))
                tcl_dir = py_home / "tcl"
                if tcl_dir.exists():
                    try:
                        tcl_lib = next((tcl_dir / p for p in os.listdir(tcl_dir) if p.lower().startswith("tcl")), None)
                        tk_lib  = next((tcl_dir / p for p in os.listdir(tcl_dir) if p.lower().startswith("tk")),  None)
                        if tcl_lib and tcl_lib.exists():
                            child_env.setdefault("TCL_LIBRARY", str(tcl_lib.resolve()))
                        if tk_lib and tk_lib.exists():
                            child_env.setdefault("TK_LIBRARY",  str(tk_lib.resolve()))
                    except Exception:
                        pass

        creationflags = 0
        if os.name == "nt" and hasattr(subprocess, "CREATE_NEW_PROCESS_GROUP"):
            creationflags = subprocess.CREATE_NEW_PROCESS_GROUP

        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                encoding="utf-8",
                errors="replace",
                creationflags=creationflags,
                env=child_env,
            )
        except Exception as e:
            yield f"data: __ERROR__:{str(e)}\n\n"
            yield "data: __EXIT_CODE__:127\n\n"
            return

        # Try to bring new GUI (e.g., Tkinter) to front on Windows
        threading.Thread(target=bring_to_front_windows, args=(proc.pid,), daemon=True).start()

        STATE["proc"]  = proc
        STATE["file"]  = rel
        STATE["scope"] = scope
        STATE["start"] = time.time()

        with open(log_path, "w", encoding="utf-8", errors="replace") as lf:
            try:
                assert proc.stdout is not None
                for line in proc.stdout:
                    line = line.rstrip("\r\n")
                    lf.write(line + "\n")
                    yield f"data: {line}\n\n"
            finally:
                rc = proc.wait()
                yield f"data: __EXIT_CODE__:{rc}\n\n"
                STATE["proc"]  = None
                STATE["file"]  = None
                STATE["scope"] = None
                STATE["start"] = None
                # Cleanup temp edited file after run ends
                if scope == "tmp":
                    try:
                        script.unlink(missing_ok=True)
                    except Exception:
                        pass

    headers = {
        "Content-Type": "text/event-stream",
        "Cache-Control": "no-cache",
        "X-Accel-Buffering": "no",
    }
    return Response(stream_with_context(generate()), headers=headers)

# ---------- Start a preview run: saves code to tmp_runs/, returns temp filename ----------
@app.post("/run/preview/start")
def run_preview_start():
    data = request.get_json(silent=True) or {}
    name = (data.get("name") or "").strip()  # original relative path (for display only)
    code = data.get("code")
    if not code:
        return jsonify({"ok": False, "error": "code required"}), 400
    try:
        reap_old_tmp()
        ts       = int(time.time())
        safe     = Path(name).stem if name else "edit"
        tmp_name = f"{safe}_{ts}.py"
        tmp_path = (TMP_DIR / tmp_name).resolve()
        tmp_path.write_text(code, encoding="utf-8", errors="ignore")
        return jsonify({"ok": True, "temp": tmp_name})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

# ---------- Stop (terminate current process) ----------
@app.post("/run/stop")
def stop_run():
    ok, msg = _terminate_proc()
    return jsonify({"ok": ok, "message": msg})

def _terminate_proc():
    proc = STATE.get("proc")
    if not proc or proc.poll() is not None:
        return False, "no running process"
    try:
        if os.name == "nt":
            try:
                proc.send_signal(signal.CTRL_BREAK_EVENT)
            except Exception:
                pass
            try:
                proc.wait(timeout=2)
            except Exception:
                pass
        if proc.poll() is None:
            proc.kill()
        return True, "terminated"
    except Exception as e:
        return False, f"terminate failed: {e}"

# --------------------------------------------------------------------------------------
# Dev server with robust auto-open
# --------------------------------------------------------------------------------------
if __name__ == "__main__":
    HOST = "0.0.0.0"
    PORT = 10000
    URL  = f"http://{HOST}:{PORT}#CPSC"  # change to '#run' if your front-end anchors expect that

    def wait_and_open(url: str, host: str, port: int, timeout: float = 10.0):
        deadline = time.time() + timeout
        while time.time() < deadline:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(0.5)
                try:
                    if s.connect_ex((host, port)) == 0:
                        webbrowser.open(url)
                        return
                except Exception:
                    pass
            time.sleep(0.25)
        # Fallback: try anyway
        try:
            webbrowser.open(url)
        except Exception:
            pass

    threading.Thread(target=wait_and_open, args=(URL, HOST, PORT), daemon=True).start()
    app.run(host=HOST, port=PORT, debug=False)
