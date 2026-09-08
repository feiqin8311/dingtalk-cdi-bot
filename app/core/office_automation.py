from __future__ import annotations


OFFICE_PROG_IDS = ("Ket.Application", "WPS.Application", "Excel.Application")


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
                return app
        except Exception:
            continue
    return None
