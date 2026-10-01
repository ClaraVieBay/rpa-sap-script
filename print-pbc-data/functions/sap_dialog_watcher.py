"""
Monitor em thread para fechar diálogos modais do SAP GUI que bloqueiam o SAP Scripting.
Usa Win32 (EnumWindows / botões) para não depender do objeto COM do SAP durante o travamento.
"""

from __future__ import annotations

import ctypes
import threading
import time
from typing import Callable, Dict, List, Optional, Set, Tuple

import win32con
import win32gui

# Intervalo entre varreduras (segundos)
DEFAULT_POLL_INTERVAL = 0.35

_watcher_lock = threading.Lock()
_watcher_thread: Optional[threading.Thread] = None
_stop_event: Optional[threading.Event] = None
_on_dismiss: Optional[Callable[[str, str], None]] = None
_last_dismiss_at: Dict[int, float] = {}
DISMISS_COOLDOWN_SEC = 0.65


def _normalize_button_label(text: str) -> str:
    return text.replace("&", "").strip().lower()


def _wm_gettext_unicode(hwnd: int) -> str:
    """Texto via WM_GETTEXT (útil quando GetWindowText vem vazio em Static/Edit)."""
    try:
        ln = int(win32gui.SendMessage(hwnd, win32con.WM_GETTEXTLENGTH, 0, 0))
        if ln <= 0:
            return ""
        buf = ctypes.create_unicode_buffer(ln + 1)
        win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, ln + 1, buf)
        return (buf.value or "").strip()
    except Exception:
        return ""


def _text_from_control(hwnd: int) -> str:
    t = win32gui.GetWindowText(hwnd)
    if t:
        return t.strip()
    try:
        cls = win32gui.GetClassName(hwnd)
        if cls in ("Static", "Edit", "Button", "ComboBox"):
            return _wm_gettext_unicode(hwnd)
    except Exception:
        pass
    return ""


def _collect_dialog_message(hwnd: int) -> str:
    """
    Junta título e textos dos controles visíveis (recursivo).
    Em diálogos Win32 padrão isso costuma reproduzir a mensagem quase por completo.
    """
    lines: List[str] = []
    seen: Set[str] = set()

    def add_line(s: str) -> None:
        s = (s or "").strip()
        if not s or s in seen:
            return
        seen.add(s)
        lines.append(s)

    try:
        add_line(win32gui.GetWindowText(hwnd))
    except Exception:
        pass

    def visit(parent: int, depth: int) -> None:
        if depth > 14:
            return

        def enum_child(child: int, _: object) -> bool:
            try:
                if win32gui.IsWindowVisible(child):
                    add_line(_text_from_control(child))
                visit(child, depth + 1)
            except Exception:
                pass
            return True

        try:
            win32gui.EnumChildWindows(parent, enum_child, None)
        except Exception:
            pass

    visit(hwnd, 0)
    return "\n".join(lines)


def _try_click_button_labels(hwnd: int, labels: List[str]) -> bool:
    targets = {_normalize_button_label(x) for x in labels}

    clicked = False

    def enum_child(child: int, _: object) -> bool:
        nonlocal clicked
        if clicked:
            return False
        try:
            if win32gui.GetClassName(child) != "Button":
                return True
            if not win32gui.IsWindowVisible(child):
                return True
            raw = win32gui.GetWindowText(child)
            if _normalize_button_label(raw) in targets:
                win32gui.PostMessage(child, win32con.BM_CLICK, 0, 0)
                clicked = True
                return False
        except Exception:
            pass
        return True

    try:
        win32gui.EnumChildWindows(hwnd, enum_child, None)
    except Exception:
        pass
    return clicked


def _try_post_dialog_command(hwnd: int, command_id: int) -> None:
    try:
        win32gui.PostMessage(hwnd, win32con.WM_COMMAND, command_id, 0)
    except Exception:
        pass


def _try_close(hwnd: int) -> None:
    try:
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    except Exception:
        pass


def _match_action(hwnd: int) -> Optional[Tuple[str, str]]:
    """
    Retorna (rótulo_lógico, texto_completo_para_log) se a janela deve ser fechada; caso contrário None.
    """
    if not win32gui.IsWindowVisible(hwnd):
        return None
    try:
        title = win32gui.GetWindowText(hwnd)
    except Exception:
        return None

    combined = _collect_dialog_message(hwnd)
    low = combined.lower()
    title_l = title.lower()

    # 1) SAP GUI for Windows — erros de rede / balanceamento (imagens com Yes/No)
    if "sap gui for windows" in title_l:
        if any(
            s in low
            for s in (
                "wsaeconnreset",
                "connection reset by peer",
                "logon balancing error",
                "pbc: connection",
                "cannot connect with message server",
                "do you want to see the detailed error",
                "do you want to see detailed error",
            )
        ):
            return (
                "sap_gui_windows_modal",
                combined or (title or "SAP GUI for Windows"),
            )

    # 2) SAP Frontend Server — dados do servidor
    if "sap frontend server" in title_l:
        if "errors in data received" in low or "application may not work" in low:
            return (
                "sap_frontend_server",
                combined or (title or "SAP Frontend Server"),
            )

    # 3) Diálogo interno "Error" / "Erro" — texto característico (evita fechar outros "Error")
    if title in ("Error", "Erro"):
        if (
            "function is not possible" in low
            or "esta função não é possível" in low
            or "esta funcao nao e possivel" in low
        ):
            return ("sap_internal_error", combined or title)

    return None


def _dismiss_window(hwnd: int, kind: str) -> bool:
    if kind == "sap_gui_windows_modal":
        if _try_click_button_labels(hwnd, ["No", "Não", "Nao", "Nein", "Non"]):
            return True
        _try_post_dialog_command(hwnd, win32con.IDCANCEL)
        _try_close(hwnd)
        return True

    if kind == "sap_frontend_server":
        if _try_click_button_labels(hwnd, ["OK", "Ok"]):
            return True
        _try_post_dialog_command(hwnd, win32con.IDOK)
        _try_close(hwnd)
        return True

    if kind == "sap_internal_error":
        if _try_click_button_labels(hwnd, ["OK", "Ok"]):
            return True
        _try_post_dialog_command(hwnd, win32con.IDOK)
        _try_close(hwnd)
        return True

    return False


def _enum_top_level(hwnd: int, dismissed_recent: Set[int]) -> bool:
    if hwnd in dismissed_recent:
        return True
    try:
        if not win32gui.IsWindowVisible(hwnd):
            return True
        m = _match_action(hwnd)
        if not m:
            return True
        now = time.monotonic()
        last = _last_dismiss_at.get(hwnd, 0.0)
        if now - last < DISMISS_COOLDOWN_SEC:
            return True
        kind, summary = m
        if _dismiss_window(hwnd, kind):
            _last_dismiss_at[hwnd] = now
            dismissed_recent.add(hwnd)
            if _on_dismiss:
                try:
                    _on_dismiss(kind, summary)
                except Exception:
                    pass
        return True
    except Exception:
        return True


def _watcher_loop(stop_event: threading.Event, interval: float) -> None:
    dismissed_recent: Set[int] = set()
    while not stop_event.wait(interval):
        try:
            dismissed_recent.clear()

            def _enum_cb(h: int, _: object) -> bool:
                return _enum_top_level(h, dismissed_recent)

            win32gui.EnumWindows(_enum_cb, None)
        except Exception:
            pass
        if len(_last_dismiss_at) > 500:
            cutoff = time.monotonic() - 120.0
            dead = [k for k, t in _last_dismiss_at.items() if t < cutoff]
            for k in dead:
                del _last_dismiss_at[k]


def start_sap_error_dialog_watcher(
    poll_interval: float = DEFAULT_POLL_INTERVAL,
    on_dismiss: Optional[Callable[[str, str], None]] = None,
) -> None:
    """
    Inicia o monitor em thread daemon. Seguro chamar várias vezes: mantém uma única thread ativa.

    on_dismiss: callback (tipo_dialogo, texto_mensagem_completo) após tentar fechar o popup.
    texto_mensagem_completo vem de _collect_dialog_message (Win32); UIA pode capturar mais em telas não padrão.
    """
    global _watcher_thread, _stop_event, _on_dismiss
    with _watcher_lock:
        if _watcher_thread is not None and _watcher_thread.is_alive():
            _on_dismiss = on_dismiss
            return
        _on_dismiss = on_dismiss
        _stop_event = threading.Event()
        _watcher_thread = threading.Thread(
            target=_watcher_loop,
            name="sap-error-dialog-watcher",
            args=(_stop_event, poll_interval),
            daemon=True,
        )
        _watcher_thread.start()


def stop_sap_error_dialog_watcher() -> None:
    global _watcher_thread, _stop_event, _on_dismiss
    with _watcher_lock:
        if _stop_event is not None:
            _stop_event.set()
        if _watcher_thread is not None:
            _watcher_thread.join(timeout=3.0)
        _watcher_thread = None
        _stop_event = None
        _on_dismiss = None
