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

CACHE = Path(os.environ["LOCALAPPDATA"]) / "yazi/docx-pages"
LIMIT = 128 * 1024 * 1024
MAX_AGE = 3 * 86400
TIMEOUT = 30
PROBE_TIMEOUT = 0.8
CONNECT_TIMEOUT = 20
SERVER_IDLE = 600
PIPE_NAME = r"\\.\pipe\yazi-docx-svc"
PDFTOPPM = r"C:\software\CLI\poppler\Library\bin\pdftoppm.exe"
YA = r"C:\software\CLI\yazi\ya.exe"


class TransportError(Exception):
    pass


class ProbeMiss(Exception):
    pass


@contextmanager
def cache_lock(timeout=TIMEOUT):
    CACHE.mkdir(parents=True, exist_ok=True)
    with (CACHE / ".lock").open("a+b") as handle:
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
        raise ValueError("This document exceeds the 128 MiB preview cache limit")


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


def validate_doc(source):
    if source.stat().st_size > 64 * 1024 * 1024:
        raise ValueError("Document exceeds the 64 MiB automatic preview limit")
    with source.open("rb") as handle:
        magic = handle.read(8)
    if magic.startswith(b"{\\rtf") or magic.startswith(b"\xd0\xcf\x11\xe0"):
        return
    if not magic.startswith(b"PK"):
        raise ValueError("File is not a Word document")
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


def export_via_word(app, source, entry):
    document = app.Documents.Open(
        FileName=str(source), ConfirmConversions=False, ReadOnly=True,
        AddToRecentFiles=False, PasswordDocument="", WritePasswordDocument="",
        Visible=False, OpenAndRepair=False, NoEncodingDialog=True,
    )
    try:
        document.ExportAsFixedFormat(
            OutputFileName=str(entry / "document.pdf"), ExportFormat=17,
            OpenAfterExport=False, OptimizeFor=1, IncludeDocProps=False,
        )
        return document.ComputeStatistics(2)
    finally:
        document.Close(SaveChanges=0)


def export_word(source, entry):
    import pythoncom

    pythoncom.CoInitialize()
    app = None
    try:
        app, pid, created = start_word()
        (entry / "owner.json").write_text(json.dumps({"pid": pid, "created": created}), encoding="utf-8")
        print(json.dumps({"pages": export_via_word(app, source, entry)}), flush=True)
    finally:
        try:
            if app is not None:
                app.Quit(SaveChanges=0)
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
        if process.create_time() == identity["created"] and process.name().lower() == "winword.exe":
            try:
                process.wait(timeout=0.5)
            except psutil.TimeoutExpired:
                process.kill()
                process.wait(timeout=5)
    except psutil.NoSuchProcess:
        pass
    owner.unlink(missing_ok=True)


def kill_server_processes(identity):
    import psutil

    for pid_key, create_key, name in (("server", "server_created", None), ("word", "created", "winword.exe")):
        try:
            process = psutil.Process(identity[pid_key])
            if process.create_time() == identity[create_key] and (not name or process.name().lower() == name):
                process.kill()
        except (psutil.NoSuchProcess, KeyError):
            pass


def sweep_stale_servers():
    import psutil

    global_path = CACHE / ".server.json"
    if global_path.exists():
        try:
            identity = json.loads(global_path.read_text(encoding="utf-8"))
            process = psutil.Process(identity["server"])
            if process.create_time() != identity["server_created"]:
                raise psutil.NoSuchProcess(identity["server"])
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


def export_via_server(source, entry, timeout):
    from multiprocessing.connection import Client

    deadline = time.monotonic() + CONNECT_TIMEOUT
    conn, spawned = None, False
    while conn is None and time.monotonic() < deadline:
        try:
            conn = Client(PIPE_NAME, family="AF_PIPE")
        except (FileNotFoundError, OSError):
            if not spawned:
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
        conn.send({"source": str(source), "entry": str(entry)})
        if not conn.poll(timeout):
            raise TransportError("Word preview server did not respond")
        reply = conn.recv()
    except (EOFError, OSError) as exc:
        raise TransportError(str(exc))
    finally:
        conn.close()
    if not reply.get("ok"):
        raise RuntimeError(reply.get("error", "Word conversion failed"))
    return {"pages": reply["pages"]}


def serve():
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
    try:
        app, word_pid, created = start_word()
    except Exception:
        pythoncom.CoUninitialize()
        sys.exit(1)
    identity_path = CACHE / ".server.json"
    identity_path.write_text(json.dumps(
        {"server": os.getpid(), "server_created": psutil.Process().create_time(),
         "word": word_pid, "created": created}), encoding="utf-8")
    stop = threading.Event()
    last_request = [time.time()]
    listener = Listener(PIPE_NAME, family="AF_PIPE")

    def watchdog():
        while not stop.wait(5):
            if time.time() - last_request[0] > SERVER_IDLE:
                break
        stop.set()
        try:
            from multiprocessing.connection import Client
            Client(PIPE_NAME, family="AF_PIPE").close()
        except Exception:
            pass

    threading.Thread(target=watchdog, daemon=True).start()
    try:
        while not stop.is_set():
            try:
                conn = listener.accept()
            except OSError:
                break
            try:
                request = conn.recv()
                last_request[0] = time.time()
                pages = export_via_word(app, Path(request["source"]), Path(request["entry"]))
                conn.send({"ok": True, "pages": pages})
            except Exception as exc:
                try:
                    conn.send({"ok": False, "error": str(exc)[-800:]})
                except Exception:
                    pass
            finally:
                conn.close()
    finally:
        try:
            app.Quit(SaveChanges=0)
        except Exception:
            pass
        pythoncom.CoUninitialize()
        identity_path.unlink(missing_ok=True)


def render_page(entry, page, edge):
    image = entry / f"page-{page}-{edge}.jpg"
    if image.exists():
        return image
    tmp = image.with_name(image.name + f".{os.getpid()}.tmp")
    result = subprocess.run(
        [PDFTOPPM, "-f", str(page + 1), "-l", str(page + 1), "-singlefile",
         "-scale-to", str(edge), "-jpeg", "-jpegopt", "quality=75",
         str(entry / "document.pdf"), str(tmp.with_suffix(""))],
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


def render(source, page, edge, probe=False):
    source = source.resolve(strict=True)
    stat = source.stat()
    identity = f"{str(source).casefold()}\0{stat.st_size}\0{stat.st_mtime_ns}"
    key = hashlib.sha256(identity.encode("utf-8")).hexdigest()
    entry = CACHE / key
    edge = max(200, min(3200, edge))
    with cache_lock(PROBE_TIMEOUT if probe else TIMEOUT):
        prune(entry)
        sweep_stale_servers()
        try:
            manifest = entry / "metadata.json"
            converted = not manifest.exists()
            if converted and probe:
                if (entry / "failed").exists():
                    raise ProbeMiss("PREVFAILED: conversion previously failed")
                raise ProbeMiss("Document is not converted yet")
            if converted:
                entry.mkdir(exist_ok=True)
                validate_doc(source)
                timeout = max(30, min(120, stat.st_size // (2 * 1024 * 1024) + 15))
                try:
                    metadata = export_via_server(source, entry, timeout)
                except TransportError:
                    try:
                        metadata = export_via_server(source, entry, timeout)
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
                manifest.write_text(json.dumps(metadata), encoding="utf-8")
                (entry / "failed").unlink(missing_ok=True)
            metadata = json.loads(manifest.read_text(encoding="utf-8"))
            try:
                (CACHE / ".current.json").write_text(
                    json.dumps({"url": str(source), "pages": metadata["pages"]}), encoding="utf-8")
            except Exception:
                pass
            page = max(0, min(page, metadata["pages"] - 1))
            targets = {page, page + 1} | ({0, 1, 2} if converted else set())
            for target in sorted(t for t in targets if 0 <= t < metadata["pages"] and t != page):
                try:
                    render_page(entry, target, edge)
                except Exception:
                    pass
            image = render_page(entry, page, edge)
            os.utime(image, None)
            pages = sorted(entry.glob("page-*.jpg"), key=lambda f: f.stat().st_mtime, reverse=True)
            for obsolete in pages[3:]:
                obsolete.unlink()
            os.utime(entry, None)
            prune(entry)
            return {"image": str(image), "dir": str(entry), "edge": edge,
                    "page": page, "pages": metadata["pages"], "cached": not converted, "converted": converted}
        except Exception as exc:
            if not probe:
                shutil.rmtree(entry, ignore_errors=True)
                try:
                    entry.mkdir(exist_ok=True)
                    (entry / "failed").write_text(str(exc)[-400:], encoding="utf-8")
                except Exception:
                    pass
            raise


def notify(url, page):
    try:
        subprocess.run([YA, "emit", "plugin", "docx-pages", f"refresh|{url}|{page}"],
                       stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                       timeout=10, creationflags=subprocess.CREATE_NO_WINDOW)
    except Exception:
        pass


if __name__ == "__main__":
    try:
        args = sys.argv[1:]
        if args[0] == "--export":
            export_word(Path(args[1]), Path(args[2]))
        elif args[0] == "--serve":
            serve()
        else:
            probe = "--probe" in args
            url = args[args.index("--notify") + 1] if "--notify" in args else None
            result = render(Path(args[0]), int(args[1]), int(args[2]), probe=probe)
            print(json.dumps(result), flush=True)
            if url:
                notify(url, result["page"])
    except ProbeMiss as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(2)
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)
