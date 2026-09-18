"""Windows Shell thumbnail fast path via ctypes.

IThumbnailCache::GetThumbnail with WTS_INCACHEONLY: returns the file's
Explorer-cached thumbnail in milliseconds and never triggers the (potentially
expensive) thumbnail provider. A miss - or a blank/tiny result - yields None
so the caller falls back to the next preview tier.

Chain (all raw ctypes COM):
  CoCreateInstance(LocalThumbnailCache) -> IThumbnailCache
  SHCreateItemFromParsingName(IID_IShellItem) -> IShellItem
  GetThumbnail -> ISharedBitmap -> HBITMAP -> PIL -> JPEG

Note: IShellItemImageFactory is not implemented by shell items on every
system (E_NOINTERFACE observed), so the documented IThumbnailCache API is
used instead.
"""
import ctypes
from ctypes import wintypes
from pathlib import Path

CLSID_LocalThumbnailCache = "{50ef4544-ac9f-4a8e-b21b-8a26180db13f}"
IID_IThumbnailCache = "{f676c15d-596a-4ce2-8234-33996f445db1}"
IID_IShellItem = "{43826d1e-e718-42ee-bc55-a1e261c37bfe}"

CLSCTX_INPROC_SERVER = 0x1
WTS_INCACHEONLY = 0x1
WTS_SCALETOREQUESTEDSIZE = 0x40

ole32 = ctypes.windll.ole32
shell32 = ctypes.windll.shell32
gdi32 = ctypes.windll.gdi32
user32 = ctypes.windll.user32


class GUID(ctypes.Structure):
    _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD),
                ("Data3", wintypes.WORD), ("Data4", ctypes.c_ubyte * 8)]


class BITMAP(ctypes.Structure):
    _fields_ = [("bmType", wintypes.LONG), ("bmWidth", wintypes.LONG),
                ("bmHeight", wintypes.LONG), ("bmWidthBytes", wintypes.LONG),
                ("bmPlanes", wintypes.WORD), ("bmBitsPixel", wintypes.WORD),
                ("bmBits", wintypes.LPVOID)]


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG),
                ("biHeight", wintypes.LONG), ("biPlanes", wintypes.WORD),
                ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
                ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD),
                ("biClrImportant", wintypes.DWORD)]


class BITMAPINFO(ctypes.Structure):
    _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", wintypes.DWORD * 3)]


def _guid(text):
    g = GUID()
    ole32.CLSIDFromString(text, ctypes.byref(g))
    return g


def _vtbl(obj):
    return ctypes.cast(ctypes.cast(obj, ctypes.POINTER(ctypes.c_void_p))[0],
                       ctypes.POINTER(ctypes.c_void_p))


def _release(item):
    proto = ctypes.WINFUNCTYPE(wintypes.ULONG, ctypes.c_void_p)
    proto(_vtbl(item)[2])(item)


def _thumb_cache():
    """CoCreateInstance(LocalThumbnailCache) -> IThumbnailCache."""
    proto = ctypes.WINFUNCTYPE(
        ctypes.HRESULT, ctypes.POINTER(GUID), ctypes.c_void_p, wintypes.DWORD,
        ctypes.POINTER(GUID), ctypes.POINTER(ctypes.c_void_p))
    create = proto(("CoCreateInstance", ole32))
    tc = ctypes.c_void_p()
    hr = create(ctypes.byref(_guid(CLSID_LocalThumbnailCache)), None,
                CLSCTX_INPROC_SERVER, ctypes.byref(_guid(IID_IThumbnailCache)),
                ctypes.byref(tc))
    return tc if hr == 0 else None


def _shell_item(source):
    proto = ctypes.WINFUNCTYPE(
        ctypes.HRESULT, wintypes.LPCWSTR, ctypes.c_void_p,
        ctypes.POINTER(GUID), ctypes.POINTER(ctypes.c_void_p))
    create = proto(("SHCreateItemFromParsingName", shell32))
    item = ctypes.c_void_p()
    hr = create(str(source), None, ctypes.byref(_guid(IID_IShellItem)), ctypes.byref(item))
    return item if hr == 0 else None


def _get_thumbnail(tc, item, px):
    """IThumbnailCache::GetThumbnail at slot 3 -> ISharedBitmap."""
    class THUMBNAILID(ctypes.Structure):
        _fields_ = [("rgbKey", ctypes.c_ubyte * 16)]

    proto = ctypes.WINFUNCTYPE(
        ctypes.HRESULT, ctypes.c_void_p, ctypes.c_void_p, wintypes.UINT,
        wintypes.DWORD, ctypes.POINTER(ctypes.c_void_p),
        ctypes.POINTER(wintypes.DWORD), ctypes.POINTER(THUMBNAILID))
    get_thumb = proto(_vtbl(tc)[3])
    shared = ctypes.c_void_p()
    cache_flags = wintypes.DWORD()
    tid = THUMBNAILID()
    hr = get_thumb(tc, item, px, WTS_INCACHEONLY | WTS_SCALETOREQUESTEDSIZE,
                   ctypes.byref(shared), ctypes.byref(cache_flags), ctypes.byref(tid))
    return shared if hr == 0 else None


def _shared_bitmap(shared):
    """ISharedBitmap::GetSharedBitmap at slot 3 -> HBITMAP."""
    proto = ctypes.WINFUNCTYPE(
        ctypes.HRESULT, ctypes.c_void_p, ctypes.POINTER(wintypes.HBITMAP))
    get_bmp = proto(_vtbl(shared)[3])
    hbmp = wintypes.HBITMAP()
    hr = get_bmp(shared, ctypes.byref(hbmp))
    return hbmp if hr == 0 else None


def _hbmp_to_pil(hbmp):
    """DDB bitmap -> PIL Image via GetDIBits (top-down 32bpp BGRA)."""
    from PIL import Image

    bm = BITMAP()
    if not gdi32.GetObjectW(hbmp, ctypes.sizeof(bm), ctypes.byref(bm)):
        return None
    w, h = bm.bmWidth, bm.bmHeight
    if w <= 0 or h <= 0:
        return None
    bmi = BITMAPINFO()
    hdr = bmi.bmiHeader
    hdr.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    hdr.biWidth, hdr.biHeight = w, -h  # negative = top-down rows
    hdr.biPlanes, hdr.biBitCount = 1, 32
    buf = (ctypes.c_ubyte * (w * h * 4))()
    hdc = user32.GetDC(None)
    try:
        rows = gdi32.GetDIBits(hdc, hbmp, 0, h, buf, ctypes.byref(bmi), 0)
    finally:
        user32.ReleaseDC(None, hdc)
    if rows != h:
        return None
    return Image.frombuffer("RGBA", (w, h), bytes(buf), "raw", "BGRA", 0, 1)


def _usable(image):
    """Reject blank/template thumbnails (white slides, flat icons)."""
    if image is None:
        return False
    if image.width < 48 or image.height < 48:
        return False
    lo, hi = image.convert("L").getextrema()
    return hi - lo >= 24


def get(source, entry, px=512):
    """Shell-cached thumbnail for source -> JPEG in entry, or None."""
    source, entry = Path(source), Path(entry)
    ole32.CoInitialize(None)
    tc = item = shared = hbmp = None
    try:
        tc = _thumb_cache()
        item = _shell_item(source)
        if tc and item:
            shared = _get_thumbnail(tc, item, px)
        if shared:
            hbmp = _shared_bitmap(shared)
        image = _hbmp_to_pil(hbmp) if hbmp else None
        if not _usable(image):
            return None
        entry.mkdir(parents=True, exist_ok=True)
        out = entry / "thumb.jpg"
        image.convert("RGB").save(out, "JPEG", quality=80)
        return out
    except Exception:
        return None
    finally:
        if hbmp:
            gdi32.DeleteObject(hbmp)
        for com in (shared, item, tc):
            if com:
                _release(com)
        ole32.CoUninitialize()
