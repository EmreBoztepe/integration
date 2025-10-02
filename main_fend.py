# -*- coding: utf-8 -*-
import time, subprocess, winreg as wr
import pythoncom, win32com.client
from pywinauto import Application as UIApp

def _get_localserver_exe(progid="Vision.Application"):
    # HKCR\Vision.Application\CLSID -> {GUID} -> HKCR\CLSID\{GUID}\LocalServer32 = "...\VISION.exe"
    with wr.OpenKey(wr.HKEY_CLASSES_ROOT, progid + r"\CLSID") as k:
        clsid = wr.QueryValueEx(k, "")[0]
    with wr.OpenKey(wr.HKEY_CLASSES_ROOT, r"CLSID\%s\LocalServer32" % clsid) as k:
        exe = wr.QueryValueEx(k, "")[0]
    return exe.strip('"')

def ensure_vision_gui_running():
    exe = _get_localserver_exe()   # Vision.EXE yolunu registry'den al
    try:
        # VISION zaten açıksa connect denemesi başarılı olur
        ui = UIApp(backend="win32").connect(title_re=r"^VISION .*")
        return ui
    except Exception:
        pass
    # Açıksa connect edilemedi, EXE'yi başlat
    subprocess.Popen([exe], shell=False)
    # Ana pencere gelene kadar bekle
    for _ in range(40):
        try:
            ui = UIApp(backend="win32").connect(title_re=r"^VISION .*")
            return ui
        except Exception:
            time.sleep(0.25)
    raise RuntimeError("VISION GUI bulunamadı/başlatılamadı.")

def open_can1_properties_and_go_settings(prj):
    # 1) GUI ayakta olsun
    ui = ensure_vision_gui_running()

    # 2) CANChannel1'i bul ve Properties aç
    can1 = prj.FindDevice("CANChannel1")
    if not can1:
        raise RuntimeError("CANChannel1 bulunamadı (önce USB→CANary→kanal adımlarını hazırla).")

    can1.EditProperties()  # parametresiz; sadece dialogu açar

    # 3) Dialogu yakala ve Settings sekmesine geç
    dlg = None
    for pat in (r"CANary Properties.*CANChannel1", r"Properties.*CANChannel1"):
        try:
            dlg = ui.window(title_re=pat)
            dlg.wait("exists enabled visible ready", timeout=10)
            break
        except Exception:
            dlg = None
    if dlg is None:
        raise RuntimeError("CANChannel1 Properties penceresi bulunamadı.")

    # Sekmeyi seç (TR/EN olasılıkları)
    try:
        dlg.child_window(control_type="TabItem", title_re=r"Settings|Ayarlar").select()
    except Exception:
        pass
    return dlg  # istersen burada checkbox/bitrate'e de tıklatabilirsin

# ---- kullanım örneği (senin main.py içinde) ----
pythoncom.CoInitialize()
try:
    prj = win32com.client.gencache.EnsureDispatch("Vision.ProjectInterface")
    # ... (proje aç, cihaz ağacını kur, vs.)
    dlg = open_can1_properties_and_go_settings(prj)
    print("✅ CANChannel1 Properties açıldı; Settings sekmesine geçildi.")
finally:
    pythoncom.CoUninitialize()
