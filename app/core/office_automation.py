from __future__ import annotations

import re

try:
    import winreg
except Exception:
    winreg = None


OFFICE_PROG_IDS = ("Ket.Application", "WPS.Application", "Excel.Application")
LOCAL_VIRTUAL_PRINTERS = (
    "Microsoft Print to PDF",
    "Microsoft XPS Document Writer",
    "导出为WPS PDF",
    "OneNote (Desktop)",
)
_ACTIVE_PRINTER_CONNECTOR_RE = re.compile(r"^.+?(\s+\S+\s+)[^:\s]+:?\s*$")


def disable_office_print_communication(app) -> None:
    """Disable Office printer communication to avoid blocking on offline/network default printers."""
    if app is None:
        return
    try:
        app.PrintCommunication = False
    except Exception:
        pass


def _iter_local_virtual_printer_candidates(current_active_printer: str = "") -> list[str]:
    if winreg is None:
        return list(LOCAL_VIRTUAL_PRINTERS)

    devices: dict[str, str] = {}
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows NT\CurrentVersion\Devices",
        ) as key:
            idx = 0
            while True:
                try:
                    name, value, _ = winreg.EnumValue(key, idx)
                    if isinstance(name, str) and isinstance(value, str):
                        devices[name] = value
                    idx += 1
                except OSError:
                    break
    except Exception:
        return list(LOCAL_VIRTUAL_PRINTERS)

    connector = None
    if isinstance(current_active_printer, str):
        match = _ACTIVE_PRINTER_CONNECTOR_RE.match(current_active_printer)
        if match:
            connector = match.group(1)

    candidates: list[str] = []
    for printer_name in LOCAL_VIRTUAL_PRINTERS:
        device_val = devices.get(printer_name)
        if not device_val:
            continue
        port = device_val.split(",")[-1].strip()
        if port:
            if connector:
                candidates.append(f"{printer_name}{connector}{port}")
            for default_conn in (" 在 ", " on "):
                cand = f"{printer_name}{default_conn}{port}"
                if cand not in candidates:
                    candidates.append(cand)
        if printer_name not in candidates:
            candidates.append(printer_name)
    return candidates


def configure_office_local_printer(app) -> None:
    """Switch Office ActivePrinter away from offline/network/remote printers to a local virtual printer."""
    if app is None:
        return

    disable_office_print_communication(app)
    try:
        current_printer = getattr(app, "ActivePrinter", "")
    except Exception:
        current_printer = ""

    if isinstance(current_printer, str) and any(
        local_name in current_printer for local_name in LOCAL_VIRTUAL_PRINTERS
    ):
        disable_office_print_communication(app)
        return

    for candidate in _iter_local_virtual_printer_candidates(
        current_printer if isinstance(current_printer, str) else ""
    ):
        try:
            app.ActivePrinter = candidate
            break
        except Exception:
            continue

    disable_office_print_communication(app)


def start_isolated_office_application(win32):
    """Start a dedicated Office process without attaching to a user's open workbook."""
    dispatch_ex = getattr(win32, "DispatchEx", None)
    if not callable(dispatch_ex):
        return None

    for prog_id in OFFICE_PROG_IDS:
        try:
            app = dispatch_ex(prog_id)
            if app is not None:
                try:
                    app.Visible = False
                    app.DisplayAlerts = False
                except Exception:
                    pass
                disable_office_print_communication(app)
                return app
        except Exception:
            continue
    return None

