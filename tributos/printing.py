"""Native Windows print dialog and PDF page output."""

from ctypes import POINTER, Structure, byref, pointer, sizeof, wintypes

HANDLE = wintypes.HANDLE
DWORD = wintypes.DWORD


class PageRange(Structure):
    _fields_ = [("start", DWORD), ("end", DWORD)]


class PrintDialog(Structure):
    _fields_ = [
        ("size", DWORD), ("owner", HANDLE), ("devmode", HANDLE), ("devnames", HANDLE),
        ("dc", HANDLE), ("flags", DWORD), ("flags2", DWORD), ("exclusions", DWORD),
        ("range_count", DWORD), ("max_ranges", DWORD), ("ranges", POINTER(PageRange)),
        ("minimum", DWORD), ("maximum", DWORD), ("copies", DWORD), ("instance", HANDLE),
        ("template", wintypes.LPCWSTR), ("callback", HANDLE), ("property_count", DWORD),
        ("properties", HANDLE), ("start_page", DWORD), ("action", DWORD),
    ]


class DocumentInfo(Structure):
    _fields_ = [("size", wintypes.INT), ("name", wintypes.LPCWSTR),
                ("output", wintypes.LPCWSTR), ("datatype", wintypes.LPCWSTR), ("flags", DWORD)]


def print_pdf(pdf, owner, title):
    from ctypes import windll
    from PIL import ImageWin

    dialogs, gdi, kernel = windll.comdlg32, windll.gdi32, windll.kernel32
    root_window = windll.user32.GetAncestor
    root_window.argtypes = [HANDLE, wintypes.UINT]
    root_window.restype = HANDLE
    dialogs.PrintDlgExW.argtypes = [POINTER(PrintDialog)]
    gdi.GetDeviceCaps.argtypes = [HANDLE, wintypes.INT]
    gdi.StartDocW.argtypes = [HANDLE, POINTER(DocumentInfo)]
    for name in ("StartPage", "EndPage", "EndDoc", "AbortDoc", "DeleteDC"):
        getattr(gdi, name).argtypes = [HANDLE]
    kernel.GlobalFree.argtypes = [HANDLE]
    kernel.GlobalFree.restype = HANDLE
    selection = PageRange(1, len(pdf))
    options = PrintDialog(size=sizeof(PrintDialog), owner=root_window(owner, 2) or owner,
                          flags=0x100 | 0x80000 | 0x40000 | 0x4,
                          max_ranges=1, ranges=pointer(selection),
                          minimum=1, maximum=len(pdf), copies=1, start_page=0xffffffff)
    try:
        if dialogs.PrintDlgExW(byref(options)) != 0:
            raise OSError("Não foi possível abrir a caixa de impressão do Windows.")
        if options.action != 1:
            return False
        first, last = (selection.start, selection.end) if options.range_count else (1, len(pdf))
        dc = options.dc
        width, height, xdpi, ydpi = (gdi.GetDeviceCaps(dc, index) for index in (8, 10, 88, 90))
        if min(width, height, xdpi, ydpi) <= 0:
            raise OSError("Não foi possível consultar a área de impressão.")
        info = DocumentInfo(sizeof(DocumentInfo), title)
        if gdi.StartDocW(dc, byref(info)) <= 0:
            raise OSError("A impressora não iniciou o trabalho.")
        try:
            for index in range(first - 1, last):
                if gdi.StartPage(dc) <= 0:
                    raise OSError("Não foi possível iniciar uma página na impressora.")
                page = pdf[index]
                try:
                    w, h = page.get_size()
                    scale = min(width / (w * xdpi), height / (h * ydpi))
                    dw, dh = round(w * xdpi * scale), round(h * ydpi * scale)
                    x, y = (width - dw) // 2, (height - dh) // 2
                    bitmap = page.render(scale=min(200 / 72, 2500 / max(w, h)))
                    try:
                        ImageWin.Dib(bitmap.to_pil().convert("RGB")).draw(dc, (x, y, x + dw, y + dh))
                    finally:
                        bitmap.close()
                finally:
                    page.close()
                if gdi.EndPage(dc) <= 0:
                    raise OSError("Não foi possível enviar uma página à impressora.")
            if gdi.EndDoc(dc) <= 0:
                raise OSError("Não foi possível finalizar o trabalho de impressão.")
        except Exception:
            gdi.AbortDoc(dc)
            raise
        return True
    finally:
        if options.dc:
            gdi.DeleteDC(options.dc)
        for handle in (options.devmode, options.devnames):
            if handle:
                kernel.GlobalFree(handle)
