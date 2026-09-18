import collections
import hashlib
import json
import msvcrt
import os
import re
from pathlib import Path
import shutil
import subprocess
import sys
import time
import zipfile
from contextlib import contextmanager
from xml.etree import ElementTree

# ---- platform layer (Windows) ----
# Cross-platform swap points: CACHE dir (LOCALAPPDATA→XDG_CACHE_HOME), cache_lock
# (msvcrt→fcntl.flock), PIPE_NAME/mutex (named pipe→Unix socket), and the
# DETACHED_PROCESS/CREATE_NO_WINDOW subprocess flags.
CACHE = Path(os.environ["LOCALAPPDATA"]) / "yazi/docx-pages"
LIMIT = 64 * 1024 * 1024
MAX_AGE = 3 * 86400
TIMEOUT = 30
PROBE_TIMEOUT = 0.8
CONNECT_TIMEOUT = 20
SERVER_IDLE = 600
PIPE_NAME = r"\\.\pipe\yazi-docx-svc"
PDFTOPPM = r"C:\software\CLI\poppler\Library\bin\pdftoppm.exe"
PDFINFO = r"C:\software\CLI\poppler\Library\bin\pdfinfo.exe"
YA = r"C:\software\CLI\yazi\ya.exe"


class TransportError(Exception):
    pass


class ServerBusy(TransportError):
    """The server answered the connection but is saturated (queued behind a
    tail or a long export). Distinct from a dead server: a busy server must
    NOT trigger the --export fallback, which would spawn a competing COM
    instance and snowball the congestion."""
    pass


class ProbeMiss(Exception):
    pass


@contextmanager
def cache_lock(timeout=TIMEOUT, path=None):
    CACHE.mkdir(parents=True, exist_ok=True)
    with (path or (CACHE / ".lock")).open("a+b") as handle:
        if handle.tell() == 0:
            handle.write(b"0")
            handle.flush()
        deadline = time.monotonic() + timeout
        while True:
            handle.seek(0)
            try:
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
                break
            except OSError:
                if time.monotonic() >= deadline:
                    raise ProbeMiss("Preview conversion is still running")
                time.sleep(0.1)
        try:
            yield
        finally:
            handle.seek(0)
            msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)


def prune(current):
    entries = []
    for entry in CACHE.iterdir():
        if not entry.is_dir() or entry.is_symlink() or len(entry.name) != 64:
            continue
        if any(c not in "0123456789abcdef" for c in entry.name):
            continue
        age = time.time() - entry.stat().st_mtime
        if entry != current and age > MAX_AGE:
            shutil.rmtree(entry)
            continue
        size = sum(f.stat().st_size for f in entry.iterdir() if f.is_file())
        entries.append((entry.stat().st_mtime, size, entry))
    total = sum(size for _, size, _ in entries)
    for _, size, entry in sorted(entries):
        if total <= LIMIT:
            break
        if entry != current:
            shutil.rmtree(entry)
            total -= size
    if total > LIMIT:
        raise ValueError("This document exceeds the 64 MiB preview cache limit")


def benign_template(target):
    if target.startswith(("\\\\", "//")):
        return False
    m = re.match(r"(?i)^file:///(.*)$", target)
    if m:
        path = m.group(1).replace("/", "\\")
        return not os.path.exists(path)
    if re.match(r"(?i)^[a-z][a-z0-9+.-]*:", target):
        return False
    return True


OFFICE_KIND = {
    ".docx": "word", ".doc": "word", ".docm": "word", ".dotx": "word",
    ".dotm": "word", ".dot": "word", ".rtf": "word",
    ".pptx": "ppt", ".pptm": "ppt", ".ppt": "ppt", ".ppsx": "ppt",
    ".ppsm": "ppt", ".pps": "ppt", ".potx": "ppt", ".potm": "ppt", ".pot": "ppt",
}


def kind_of(source):
    return OFFICE_KIND.get(source.suffix.lower())


PPT_OOXML = {ext for ext, kind in OFFICE_KIND.items()
             if kind == "ppt" and ext not in (".ppt", ".pps", ".pot")}
WORD_OOXML = {ext for ext, kind in OFFICE_KIND.items()
              if kind == "word" and ext not in (".doc", ".dot", ".rtf")}


def quick_pages(source):
    """Page/slide count without conversion, for the early status hint.
    PPTX slide count from the package is exact; DOCX app.xml <Pages> is
    whatever Word last saved (approximate but better than nothing)."""
    ext = source.suffix.lower()
    try:
        if ext == ".pdf":
            return pdf_pages(source)
        if ext in PPT_OOXML:
            with zipfile.ZipFile(source) as z:
                pres = z.read("ppt/presentation.xml").decode("utf-8", "replace")
                count = len(re.findall(r"<p:sldId[ >]", pres))
                if count:
                    return count
                return sum(1 for n in z.namelist()
                           if n.startswith("ppt/slides/slide") and n.endswith(".xml"))
        if ext in WORD_OOXML:
            with zipfile.ZipFile(source) as z:
                xml = z.read("docProps/app.xml").decode("utf-8", "replace")
            m = re.search(r"<Pages>(\d+)</Pages>", xml)
            return int(m.group(1)) if m else None
    except Exception:
        return None


def validate_doc(source):
    if source.stat().st_size > 64 * 1024 * 1024:
        raise ValueError("Document exceeds the 64 MiB automatic preview limit")
    with source.open("rb") as handle:
        magic = handle.read(8)
    if magic.startswith(b"{\\rtf") or magic.startswith(b"\xd0\xcf\x11\xe0"):
        return
    if not magic.startswith(b"PK"):
        raise ValueError("File is not an Office document")
    with zipfile.ZipFile(source) as archive:
        for info in archive.infolist():
            name = info.filename.lower()
            if "vbaproject" in name or "/activex/" in name:
                raise ValueError("Active content is not supported by automatic preview")
            if name.endswith(".rels"):
                if info.file_size > 1024 * 1024:
                    raise ValueError("Oversized document relationships")
                for rel in ElementTree.fromstring(archive.read(info)):
                    if rel.get("TargetMode") != "External":
                        continue
                    rtype = rel.get("Type", "")
                    if rtype.endswith("/hyperlink"):
                        continue
                    if rtype.endswith("/attachedTemplate") and benign_template(rel.get("Target", "")):
                        continue
                    raise ValueError("External linked content is not loaded for preview")
            if name.startswith("word/") and name.endswith(".xml"):
                if info.file_size > 16 * 1024 * 1024:
                    raise ValueError("Oversized document XML")
                tree = ElementTree.fromstring(archive.read(info))
                instructions = "".join(node.text or "" for node in tree.iter() if node.tag.endswith("}instrText"))
                instructions += " ".join(value for node in tree.iter() if node.tag.endswith("}fldSimple") for attr, value in node.attrib.items() if attr.endswith("}instr"))
                if re.search(r"\b(DDEAUTO|DDE|INCLUDETEXT|INCLUDEPICTURE|DATABASE|LINK)\b", instructions, re.IGNORECASE):
                    raise ValueError("External or dynamic data fields are not evaluated for preview")


def start_word():
    import psutil
    import pythoncom
    import win32com.client
    import win32process

    previous = set(psutil.pids())
    app = win32com.client.DispatchEx("Word.Application")
    try:
        document = app.Documents.Add(Visible=False)
        _, pid = win32process.GetWindowThreadProcessId(document.Windows.Item(1).Hwnd)
        process = psutil.Process(pid)
        if pid in previous or process.name().lower() != "winword.exe":
            raise RuntimeError("Word did not provide an isolated process; existing Word sessions were left untouched")
        created = process.create_time()
        register_owned("winword", pid, created)
        app.Visible = False
        app.DisplayAlerts = 0
        app.AutomationSecurity = 3
        try:
            app.Options.UpdateLinksAtOpen = False
        except Exception:
            pass
        document.Close(SaveChanges=0)
    except Exception:
        try:
            app.Quit(SaveChanges=0)
        except Exception:
            pass
        raise
    return app, pid, created


def start_powerpoint():
    """Dispatch PowerPoint COM. If the user already has PowerPoint running we
    attach to it (read-only, invisible open) — pid is then None so the caller
    knows not to quit it."""
    import psutil
    import win32com.client

    previous = {p.pid for p in psutil.process_iter(["name"])
                if (p.info["name"] or "").lower() == "powerpnt.exe"}
    app = win32com.client.DispatchEx("PowerPoint.Application")
    pid = created = None
    time.sleep(0.5)
    for p in psutil.process_iter(["name", "create_time"]):
        if (p.info["name"] or "").lower() == "powerpnt.exe" and p.pid not in previous:
            pid, created = p.pid, p.info["create_time"]
    register_owned("powerpnt", pid, created)
    try:
        app.AutomationSecurity = 3
    except Exception:
        pass
    return app, pid, created


PPT_PX = 1600  # long-edge pixels for direct slide PNG export


def export_via_ppt(app, source, entry):
    """Slides → page-N.png. First slides go out individually so page 0 is
    usable immediately; the rest are bulk-exported in tail() which runs
    after the reply is sent. Falls back to PDF export on failure."""
    pres = app.Presentations.Open(str(source), ReadOnly=-1, Untitled=0, WithWindow=0)
    try:
        count = pres.Slides.Count
        w_pt, h_pt = float(pres.PageSetup.SlideWidth), float(pres.PageSetup.SlideHeight)
        if w_pt >= h_pt:
            sw, sh = PPT_PX, max(1, round(PPT_PX * h_pt / w_pt))
        else:
            sh, sw = PPT_PX, max(1, round(PPT_PX * w_pt / h_pt))
        try:
            fast = min(3, count)
            for i in range(1, fast + 1):
                pres.Slides(i).Export(str(entry / f"page-{i - 1}.png"), "PNG", sw, sh)
            if count <= fast:
                pres.Close()
                return {"pages": count, "png": True}, None
        except Exception:
            pass

        def tail():
            # Generator: one slide per yield so the serve loop can interleave
            # queued requests between slides instead of blocking on a bulk
            # export. gen.close() on shutdown lands in this finally.
            failed = False
            try:
                for i in range(1, count + 1):
                    target = entry / f"page-{i - 1}.png"
                    try:
                        if not target.exists():
                            pres.Slides(i).Export(str(target), "PNG", sw, sh)
                    except Exception:
                        failed = True
                    yield
                if failed:
                    try:
                        pres.SaveAs(str(entry / "document.pdf"), 32)  # ppSaveAsPDF fallback
                    except Exception:
                        pass
            finally:
                pres.Close()

        return {"pages": count, "png": True}, tail()
    except Exception:
        try:
            pres.Close()
        except Exception:
            pass
        raise


def export_via_word(app, source, entry):
    """DOCX → PDF in two stages: ComputeStatistics forces pagination, then
    pages 1-3 go to head.pdf so page 0 is usable immediately; tail() exports
    the whole document to document.pdf after the reply is sent. Documents
    of <=3 pages skip the head stage."""
    document = app.Documents.Open(
        FileName=str(source), ConfirmConversions=False, ReadOnly=True,
        AddToRecentFiles=False, PasswordDocument="", WritePasswordDocument="",
        Visible=False, OpenAndRepair=False, NoEncodingDialog=True,
    )
    try:
        pages = document.ComputeStatistics(2)  # wdStatisticPages
        if pages <= 3:
            document.ExportAsFixedFormat(
                OutputFileName=str(entry / "document.pdf"), ExportFormat=17,
                OpenAfterExport=False, OptimizeFor=1, IncludeDocProps=False,
            )
            document.Close(SaveChanges=0)
            return {"pages": pages}, None
        document.ExportAsFixedFormat(
            OutputFileName=str(entry / "head.pdf"), ExportFormat=17,
            OpenAfterExport=False, OptimizeFor=1, IncludeDocProps=False,
            Range=3, From=1, To=3,  # wdExportFromTo
        )

        def tail():
            # Single-step generator (uniform with the PPT tail interface).
            try:
                document.ExportAsFixedFormat(
                    OutputFileName=str(entry / "document.pdf"), ExportFormat=17,
                    OpenAfterExport=False, OptimizeFor=1, IncludeDocProps=False,
                )
            finally:
                document.Close(SaveChanges=0)
            yield

        return {"pages": pages, "head": 3}, tail()
    except Exception:
        try:
            document.Close(SaveChanges=0)
        except Exception:
            pass
        raise


def export_text(app, source):
    """Whole-document text via Word COM: Content.Text is just the character
    stream - no pagination or export needed, so it is much faster than the
    PDF path (~1-2s on a warm instance). Serves legacy OLE formats (.doc,
    .dot, .rtf) that have no OOXML fast path."""
    document = app.Documents.Open(
        FileName=str(source), ConfirmConversions=False, ReadOnly=True,
        AddToRecentFiles=False, PasswordDocument="", WritePasswordDocument="",
        Visible=False, OpenAndRepair=False, NoEncodingDialog=True,
    )
    try:
        text = document.Content.Text
    finally:
        document.Close(SaveChanges=0)
    return text[:2 * 1024 * 1024].replace("\r", "\n").replace("\x07", "\t")


def export_office(source, entry):
    import pythoncom

    pythoncom.CoInitialize()
    kind = kind_of(source) or "word"
    app = None
    owned = True
    try:
        if kind == "ppt":
            app, pid, created = start_powerpoint()
            owned = pid is not None
            metadata, tail = export_via_ppt(app, source, entry)
        else:
            app, pid, created = start_word()
            metadata, tail = export_via_word(app, source, entry)
        (entry / "owner.json").write_text(json.dumps({"pid": pid or 0, "created": created or 0}), encoding="utf-8")
        if tail is not None:
            collections.deque(tail, maxlen=0)  # tails are generators now
        print(json.dumps(metadata), flush=True)
    finally:
        try:
            if app is not None and owned:
                if kind == "word":
                    app.Quit(SaveChanges=0)
                else:
                    app.Quit()
        finally:
            pythoncom.CoUninitialize()


def cleanup_word(entry):
    import psutil

    owner = entry / "owner.json"
    if not owner.exists():
        return
    identity = json.loads(owner.read_text(encoding="utf-8"))
    try:
        process = psutil.Process(identity["pid"])
        if process.create_time() == identity["created"] and process.name().lower() in ("winword.exe", "powerpnt.exe"):
            try:
                process.wait(timeout=0.5)
            except psutil.TimeoutExpired:
                process.kill()
                process.wait(timeout=5)
    except psutil.NoSuchProcess:
        pass
    owner.unlink(missing_ok=True)


def register_owned(name, pid, created):
    """Ledger entry for an Office process we spawned: .owned-{name}-{pid}-{created}
    holding the spawning python pid. sweep_stale_servers() kills any such app
    whose owner python is gone - covers hard-killed servers/exports where the
    graceful Quit in finally never ran."""
    if not pid:
        return
    try:
        import psutil

        marker = CACHE / f".owned-{name}-{pid}-{created}"
        tmp = marker.with_name(marker.name + ".tmp")
        tmp.write_text(json.dumps({"owner": os.getpid(),
                                   "owner_created": psutil.Process().create_time()}), encoding="utf-8")
        os.replace(tmp, marker)
    except Exception:
        pass


def kill_server_processes(identity):
    import psutil

    for pid_key, create_key, name in (("server", "server_created", None), ("word", "word_created", "winword.exe"),
                                      ("ppt", "ppt_created", "powerpnt.exe")):
        try:
            process = psutil.Process(identity[pid_key])
            if process.create_time() == identity[create_key] and (not name or process.name().lower() == name):
                process.kill()
        except (psutil.NoSuchProcess, KeyError):
            pass


def sweep_stale_servers():
    import psutil

    server_pids = set()
    global_path = CACHE / ".server.json"
    if global_path.exists():
        try:
            identity = json.loads(global_path.read_text(encoding="utf-8"))
            process = psutil.Process(identity["server"])
            if process.create_time() != identity["server_created"]:
                raise psutil.NoSuchProcess(identity["server"])
            for field in ("word", "ppt", "server"):
                if identity.get(field):
                    server_pids.add(int(identity[field]))
        except (psutil.NoSuchProcess, Exception):
            try:
                kill_server_processes(identity)
            except Exception:
                pass
            global_path.unlink(missing_ok=True)
    for legacy in CACHE.glob(".server-*.json"):
        try:
            if not psutil.pid_exists(int(legacy.stem.rsplit("-", 1)[1])):
                kill_server_processes(json.loads(legacy.read_text(encoding="utf-8")))
                legacy.unlink(missing_ok=True)
        except Exception:
            pass
    for marker in CACHE.glob(".owned-*"):
        try:
            name, pid_s, created_s = marker.name[7:].rsplit("-", 2)
            pid, created = int(pid_s), float(created_s)
            owner = json.loads(marker.read_text(encoding="utf-8"))
            owner_alive = psutil.pid_exists(owner["owner"]) and \
                psutil.Process(owner["owner"]).create_time() == owner["owner_created"]
            try:
                proc = psutil.Process(pid)
                alive = proc.create_time() == created and proc.name().lower() == name + ".exe"
            except psutil.NoSuchProcess:
                alive = False
            if alive and not owner_alive:
                proc.kill()
                alive = False
            if not alive:
                marker.unlink(missing_ok=True)
        except Exception:
            pass
    # Untracked automation instances: a python owner killed between the COM
    # spawn and its .owned-* marker write leaves an orphan no ledger knows.
    # "-Embedding" only appears on automation instances - never on a user's
    # interactive Word/PowerPoint - so cmdline matching is safe. Procs younger
    # than 10s may be mid-registration and are skipped.
    try:
        tracked = set(server_pids)
        for marker in CACHE.glob(".owned-*"):
            try:
                tracked.add(int(marker.name[7:].rsplit("-", 2)[1]))
            except Exception:
                pass
        now = time.time()
        for proc in psutil.process_iter(["pid", "name", "cmdline", "create_time"]):
            try:
                if proc.info["pid"] in tracked:
                    continue
                if (proc.info["name"] or "").lower() not in ("winword.exe", "powerpnt.exe"):
                    continue
                if now - (proc.info["create_time"] or 0) < 10:
                    continue
                cmd = " ".join(proc.info["cmdline"] or []).lower()
                if "-embedding" in cmd:
                    proc.kill()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
    except Exception:
        pass


def server_request(payload, timeout, spawn=True):
    """One request on the preview pipe. TransportError = server unreachable
    (a fresh --serve may be spawned once); ServerBusy = connected but the
    reply did not arrive within timeout (the server is alive and queued)."""
    from multiprocessing.connection import Client

    deadline = time.monotonic() + CONNECT_TIMEOUT
    conn, spawned = None, False
    while conn is None and time.monotonic() < deadline:
        try:
            conn = Client(PIPE_NAME, family="AF_PIPE")
        except (FileNotFoundError, OSError):
            if not spawned and spawn:
                subprocess.Popen(
                    [sys.executable, "-X", "utf8", __file__, "--serve"],
                    stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                    creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NO_WINDOW,
                    close_fds=True,
                )
                spawned = True
            time.sleep(0.3)
    if conn is None:
        raise TransportError("Word preview server did not start")
    try:
        conn.send(payload)
        if not conn.poll(timeout):
            raise ServerBusy("Word preview server is busy")
        reply = conn.recv()
    except (EOFError, OSError) as exc:
        raise TransportError(str(exc))
    finally:
        conn.close()
    if not reply.get("ok"):
        raise RuntimeError(reply.get("error", "Word preview server error"))
    return reply


def export_via_server(source, entry, timeout, kind="word"):
    reply = server_request(
        {"source": str(source), "entry": str(entry), "kind": kind}, timeout)
    return {"pages": reply["pages"], "png": reply.get("png"), "head": reply.get("head")}


def serve():
    import queue
    import threading

    import psutil
    import pythoncom
    import win32api
    import win32event
    from multiprocessing.connection import Listener

    mutex = win32event.CreateMutex(None, True, "Global\\yazi-docx-preview-svc")
    if win32api.GetLastError() == 183:
        return
    pythoncom.CoInitialize()
    apps, owners = {}, {}
    try:
        apps["word"], pid, created = start_word()
        owners["word"] = (pid, created, True)
    except Exception:
        pythoncom.CoUninitialize()
        sys.exit(1)
    identity_path = CACHE / ".server.json"
    identity_path.write_text(json.dumps(
        {"server": os.getpid(), "server_created": psutil.Process().create_time(),
         "word": owners["word"][0], "word_created": owners["word"][1]}), encoding="utf-8")

    kind_locks = {"word": threading.Lock(), "ppt": threading.Lock()}

    def app_for(kind):
        if kind not in apps:
            with kind_locks[kind]:
                if kind not in apps:
                    if kind == "ppt":
                        app, pid, created = start_powerpoint()
                    else:
                        app, pid, created = start_word()
                    apps[kind] = app
                    owners[kind] = (pid, created, kind != "ppt" or pid is not None)
                    try:
                        data = json.loads(identity_path.read_text(encoding="utf-8"))
                        data[kind] = pid or 0
                        data[kind + "_created"] = created or 0
                        identity_path.write_text(json.dumps(data), encoding="utf-8")
                    except Exception:
                        pass
        return apps[kind]
    stop = threading.Event()
    last_request = [time.time()]
    listener = Listener(PIPE_NAME, family="AF_PIPE")
    warmed = {"ppt": False}

    def watchdog():
        from multiprocessing.connection import Client
        while not stop.wait(5):
            idle = time.time() - last_request[0]
            if idle > SERVER_IDLE:
                break
            # Pre-warm PowerPoint during idle gaps so the first PPT hover
            # doesn't eat its ~5s COM boot. One-shot; failures stay lazy.
            if not warmed["ppt"] and idle > 3 and "ppt" not in apps:
                warmed["ppt"] = True
                try:
                    c = Client(PIPE_NAME, family="AF_PIPE")
                    c.send({"kind": "ppt", "warm": True})
                    c.close()
                except Exception:
                    pass
        stop.set()
        try:
            Client(PIPE_NAME, family="AF_PIPE").close()
        except Exception:
            pass

    threading.Thread(target=watchdog, daemon=True).start()
    # Requests queue up so the main loop can interleave them with pending
    # tail generators: a new request is served between slide exports instead
    # of waiting for the previous deck's tail to finish.
    requests = queue.Queue()

    def acceptor():
        while not stop.is_set():
            try:
                conn = listener.accept()
            except OSError:
                break
            requests.put(conn)

    threading.Thread(target=acceptor, daemon=True).start()
    pending_tails = collections.deque()
    try:
        while not stop.is_set():
            try:
                conn = requests.get(timeout=0.5)
            except queue.Empty:
                conn = None
            if conn is None:
                # Idle slice: advance every pending tail one step, then block
                # for the next request again.
                for _ in range(len(pending_tails)):
                    gen = pending_tails.popleft()
                    try:
                        next(gen)
                        pending_tails.append(gen)
                    except StopIteration:
                        pass
                    except Exception:
                        pass
                continue
            kind = None
            tail = None
            try:
                request = conn.recv()
                last_request[0] = time.time()
                kind = request.get("kind", "word")
                if request.get("warm"):
                    app_for(kind)
                    conn.send({"ok": True, "pages": 0})
                    continue
                if request.get("text"):
                    conn.send({"ok": True, "text": export_text(
                        app_for("word"), Path(request["source"]))})
                    continue
                for attempt in (0, 1):
                    app = app_for(kind)
                    try:
                        if kind == "ppt":
                            meta, tail = export_via_ppt(app, Path(request["source"]), Path(request["entry"]))
                        else:
                            meta, tail = export_via_word(app, Path(request["source"]), Path(request["entry"]))
                        break
                    except Exception:
                        # Dead COM handle (killed/crashed app) is unrecoverable:
                        # drop it once and let app_for respawn. Document-level
                        # errors on a live app are raised straight through.
                        pid = owners.get(kind, (None,))[0]
                        if attempt or (pid is not None and psutil.pid_exists(pid)):
                            raise
                        apps.pop(kind, None)
                        owners.pop(kind, None)
                conn.send({"ok": True, **meta})
            except Exception as exc:
                try:
                    conn.send({"ok": False, "error": str(exc)[-800:]})
                except Exception:
                    pass
            finally:
                conn.close()
            if tail is not None:
                pending_tails.append(tail)
    finally:
        for gen in pending_tails:
            try:
                gen.close()  # lands in the tail's finally -> pres/document.Close
            except Exception:
                pass
        for kind, app in apps.items():
            try:
                if owners.get(kind, (None, None, False))[2]:
                    if kind == "word":
                        app.Quit(SaveChanges=0)
                    else:
                        app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
        identity_path.unlink(missing_ok=True)


def pdf_pages(source):
    result = subprocess.run(
        [PDFINFO, str(source)],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=15,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    m = re.search(r"^Pages:\s+(\d+)", result.stdout, re.MULTILINE)
    if result.returncode or not m:
        raise RuntimeError("Could not read PDF page count: " + result.stderr.strip()[-300:])
    return int(m.group(1))


def render_all_pages(entry, edge, pdf, pages):
    missing = [p for p in range(pages)
               if not (entry / f"page-{p}-{edge}.jpg").exists()]
    if not missing:
        return
    prefix = entry / f"batch-{os.getpid()}"
    result = subprocess.run(
        [PDFTOPPM, "-scale-to", str(edge), "-jpeg", "-jpegopt", "quality=75",
         str(pdf), str(prefix)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=max(60, pages * 3), creationflags=subprocess.CREATE_NO_WINDOW,
    )
    produced = list(entry.glob(prefix.name + "-*.jpg"))
    for f in produced:
        try:
            n = int(f.stem.rsplit("-", 1)[1])
        except ValueError:
            f.unlink(missing_ok=True)
            continue
        if 1 <= n <= pages:
            os.replace(f, entry / f"page-{n - 1}-{edge}.jpg")
        else:
            f.unlink(missing_ok=True)
    if result.returncode and not produced:
        raise RuntimeError("PDF batch rendering failed: " + result.stderr.strip()[-500:])


def render_page(entry, page, edge, pdf):
    image = entry / f"page-{page}-{edge}.jpg"
    if image.exists():
        return image
    tmp = image.with_name(image.name + f".{os.getpid()}.tmp")
    result = subprocess.run(
        [PDFTOPPM, "-f", str(page + 1), "-l", str(page + 1), "-singlefile",
         "-scale-to", str(edge), "-jpeg", "-jpegopt", "quality=75",
         str(pdf), str(tmp.with_suffix(""))],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=15,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    if not result.returncode and tmp.with_suffix(".jpg").exists():
        os.replace(tmp.with_suffix(".jpg"), image)
        return image
    tmp.with_suffix(".jpg").unlink(missing_ok=True)
    if result.returncode or not image.exists():
        raise RuntimeError("PDF page rendering failed: " + result.stderr.strip()[-500:])
    return image


def convert_to_pdf(source, entry, stat):
    """Office document → PDF seam. Windows backend: persistent Word COM server,
    with a detached one-shot --export fallback. A LibreOffice backend
    (soffice --headless --convert-to pdf) slots in here for other platforms.
    Returns the {"pages": n} metadata dict."""
    validate_doc(source)
    kind = kind_of(source) or "word"
    timeout = max(30, min(120, stat.st_size // (2 * 1024 * 1024) + 15))
    try:
        return export_via_server(source, entry, timeout, kind)
    except ServerBusy:
        # The server is alive but queued behind another export/tail; spawning
        # a competing COM instance via --export would only snowball the delay.
        raise
    except TransportError:
        try:
            return export_via_server(source, entry, timeout, kind)
        except ServerBusy:
            raise
        except TransportError:
            try:
                result = subprocess.run(
                    [sys.executable, "-X", "utf8", __file__, "--export", str(source), str(entry)],
                    capture_output=True, text=True, encoding="utf-8", timeout=TIMEOUT,
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
            finally:
                cleanup_word(entry)
            if result.returncode:
                raise RuntimeError("Word conversion failed: " + result.stderr.strip()[-800:])
            metadata = json.loads(result.stdout)
    latest = source.stat()
    if (latest.st_mtime_ns, latest.st_size) != (stat.st_mtime_ns, stat.st_size):
        raise RuntimeError("Document changed during conversion; please preview again")
    return metadata


def embedded_thumb(source, entry):
    """docProps/thumbnail.* inside the OOXML package -> JPEG in entry, or None.
    Blank/template thumbnails (generator defaults) are rejected."""
    import io

    try:
        from PIL import Image
        with zipfile.ZipFile(source) as z:
            name = next((n for n in z.namelist()
                         if n.lower().startswith("docprops/thumbnail.")), None)
            if not name:
                return None
            image = Image.open(io.BytesIO(z.read(name)))
            image.load()
        if image.width < 48 or image.height < 48:
            return None
        lo, hi = image.convert("L").getextrema()
        if hi - lo < 24:
            return None
        entry.mkdir(parents=True, exist_ok=True)
        out = entry / "thumb.jpg"
        image.convert("RGB").save(out, "JPEG", quality=80)
        return out
    except Exception:
        return None


def write_pages_record(source, pages):
    """Persist the total page count for the status bar. A single shared
    .current.json races whenever another file's render/preload finishes -
    each file therefore gets its own record keyed by a djb2 hash of the
    (case-folded) url, mirroring the hash init.lua can recompute."""
    url = str(source)
    h = 5381
    for b in url.encode("utf-8"):
        o = b + 32 if 65 <= b <= 90 else b  # ASCII lower, matches Lua string.lower on bytes
        h = ((h * 33) + o) & 0xFFFFFFFF
    try:
        (CACHE / f".pages-{h:08x}.json").write_text(json.dumps({"pages": pages}), encoding="utf-8")
        (CACHE / ".current.json").write_text(json.dumps({"url": url, "pages": pages}), encoding="utf-8")
    except Exception:
        pass


def thumb_for(source, entry):
    """Fast preview image while conversion runs: embedded OOXML thumbnail
    first (real document content), then the Explorer shell cache (covers
    legacy OLE formats). Best effort; None means the caller falls back."""
    thumb = embedded_thumb(source, entry)
    if thumb:
        return thumb
    try:
        import shell_thumb
        return shell_thumb.get(source, entry)
    except Exception:
        return None


def render(source, page, edge, probe=False, lock_timeout=None):
    source = source.resolve(strict=True)
    stat = source.stat()
    identity = f"{str(source).casefold()}\0{stat.st_size}\0{stat.st_mtime_ns}"
    key = hashlib.sha256(identity.encode("utf-8")).hexdigest()
    entry = CACHE / key
    edge = max(200, min(3200, edge))
    if probe and not (entry / "metadata.json").exists():
        # Fast path outside the cache lock: a conversion holds it for the
        # whole COM export, so probes must not wait on it just to answer
        # "not ready yet". Serve a thumbnail instead when one exists.
        pages_hint = quick_pages(source)
        if pages_hint:
            write_pages_record(source, pages_hint)
        if (entry / "failed").exists():
            raise ProbeMiss("PREVFAILED: conversion previously failed")
        thumb = thumb_for(source, entry)
        if thumb:
            return {"thumb": str(thumb), "dir": str(entry), "page": 0,
                    "pages": pages_hint or 0, "probe_thumb": True}
        raise ProbeMiss("Document is not converted yet")
    wait = PROBE_TIMEOUT if probe else (lock_timeout if lock_timeout is not None else TIMEOUT)
    with cache_lock(wait):
        prune(entry)
        sweep_stale_servers()
    # Per-entry lock: conversion serializes only against same-file workers;
    # other files proceed in parallel. Page waits and rasterization below run
    # entirely outside any lock - a slow slide export must not stall them.
    with cache_lock(wait, CACHE / f"{key}.lock"):
        manifest = entry / "metadata.json"
        converted = not manifest.exists()
        try:
            if converted and probe:
                if (entry / "failed").exists():
                    raise ProbeMiss("PREVFAILED: conversion previously failed")
                raise ProbeMiss("Document is not converted yet")
            is_pdf = source.suffix.lower() == ".pdf"
            if converted:
                entry.mkdir(exist_ok=True)
                if is_pdf:
                    metadata = {"pages": pdf_pages(source)}
                else:
                    metadata = convert_to_pdf(source, entry, stat)
                manifest.write_text(json.dumps(metadata), encoding="utf-8")
                (entry / "failed").unlink(missing_ok=True)
            metadata = json.loads(manifest.read_text(encoding="utf-8"))
            write_pages_record(source, metadata["pages"])
            page = max(0, min(page, metadata["pages"] - 1))
        except Exception as exc:
            if not probe:
                shutil.rmtree(entry, ignore_errors=True)
                try:
                    entry.mkdir(exist_ok=True)
                    (entry / "failed").write_text(str(exc)[-400:], encoding="utf-8")
                except Exception:
                    pass
            raise
    if metadata.get("png"):
        # PPT path: slide 1..3 arrive first; the rest are still being
        # exported in the server tail - poll briefly before giving up.
        image = entry / f"page-{page}.png"
        deadline = time.monotonic() + 15
        while not image.exists() and time.monotonic() < deadline:
            time.sleep(0.2)
        if not image.exists():
            raise RuntimeError(f"slide image missing: {image.name}")
        os.utime(image, None)
    else:
        # DOCX path: while the server tail is still writing the full
        # document.pdf, head.pdf covers pages 0..head-1; deeper pages
        # briefly poll for the full export before giving up.
        pdf = source if is_pdf else entry / "document.pdf"
        if not is_pdf and not pdf.exists():
            head = metadata.get("head") or 0
            if page >= head:
                deadline = time.monotonic() + 15
                while not pdf.exists() and time.monotonic() < deadline:
                    time.sleep(0.2)
            if not pdf.exists():
                pdf = entry / "head.pdf"
        targets = {page, page + 1} | ({0, 1, 2} if converted else set())
        for target in sorted(t for t in targets if 0 <= t < metadata["pages"] and t != page):
            try:
                render_page(entry, target, edge, pdf)
            except Exception:
                pass
        image = render_page(entry, page, edge, pdf)
        os.utime(image, None)
    pages = sorted(entry.glob("page-*.jpg"), key=lambda f: f.stat().st_mtime, reverse=True)
    for obsolete in pages[50:]:
        obsolete.unlink()
    os.utime(entry, None)
    return {"image": str(image), "dir": str(entry), "edge": edge,
            "page": page, "pages": metadata["pages"], "png": bool(metadata.get("png")),
            "cached": not converted, "converted": converted}


def notify(url, page):
    try:
        subprocess.run([YA, "emit", "plugin", "omni-previewer", f"refresh|{url}|{page}"],
                       stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                       timeout=10, creationflags=subprocess.CREATE_NO_WINDOW)
    except Exception:
        pass


if __name__ == "__main__":
    try:
        args = sys.argv[1:]
        if args[0] == "--export":
            export_office(Path(args[1]), Path(args[2]))
        elif args[0] == "--serve":
            serve()
        else:
            probe = "--probe" in args
            url = args[args.index("--notify") + 1] if "--notify" in args else None
            # Preloads must not queue behind a running conversion: the in-flight
            # worker will emit its own refresh when it finishes.
            result = render(Path(args[0]), int(args[1]), int(args[2]), probe=probe,
                            lock_timeout=2 if url else None)
            print(json.dumps(result), flush=True)
            if url:
                notify(url, result["page"])
                # Fill remaining pages in the background WITHOUT holding the
                # cache lock, so probes/flips are never blocked by the batch.
                try:
                    src = Path(args[0])
                    pdf = src if src.suffix.lower() == ".pdf" else Path(result["dir"]) / "document.pdf"
                    if not result.get("png") and pdf.exists() and result.get("pages", 0) <= 60:
                        render_all_pages(Path(result["dir"]), result["edge"], pdf, result["pages"])
                except Exception:
                    pass
    except ProbeMiss as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(2)
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)
