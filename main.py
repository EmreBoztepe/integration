# make_vst.py — A2L -> VST (UI'siz, doğrudan StrategyFileInterface)
import os
import pythoncom
import win32com.client
import time
from win32com.client import VARIANT
# >>> BURAYI DÜZENLE
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
A2L_PATH   = os.path.join(SCRIPT_DIR, "example.a2l")     # kendi .a2l dosyan
VST_OUT    = os.path.join(SCRIPT_DIR, "out", "MyECU.vst")  # çıkış .vst
CAL_OUT    = os.path.join(SCRIPT_DIR, "out", "MyECU.cal")  # çıkış .cal
PRJ_OUT    = os.path.join(SCRIPT_DIR, "base", "base.vpj")  # çıkış .cal
S19_PATH   = os.path.join(SCRIPT_DIR, "example.s19") #s19 dosyası.
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
os.makedirs(OUT_DIR, exist_ok=True)
UPLOADED_VST = os.path.join(OUT_DIR, "MyECU.vst")
# Enum değerleri (dokümandaki VISION_DEVICE_TYPES)
VISION_DEVICE_VIRTUALPCM = 36   # Virtual PCM
VISION_DEVICE_APIPORT    = 37   # VISION API Port
VISION_DEVICE_USBPORT    = 1
def ensure_dir(p):
    d = os.path.dirname(p)
    if d and not os.path.exists(d):
        os.makedirs(d, exist_ok=True)

def import_a2l(strat, a2l_path):
    strat.SetASAP2ImportProperties2(
                "",  # StrategyPreset
                True,   # ImportFunctions
                False,  # SwapAxes
                False,  # IgnoreMemoryRegions
                False,  # ExtendLimits
                True,   # EnforceLimits
                True,   # DeleteExistingItems
                False,   # ReplaceExistingItems
                True,   # ClearDeviceSettings
                True,   # AllowBrackets
                True,   # OrganizeDataItemInGroups
                False,  # UseDisplayIdentifiers
                1,      # StructureNameOption
                '_',    # GroupSeparator
                0       # CharacterSet
            )
    if hasattr(strat, "Import"):
        try:
            strat.Import(a2l_path)
            print("A2L import ✅")
            return True
        except Exception as e:
            print("   -> Import() da başarisiz:", e)
    return False

def import_s19(strat, s19_path):
    
    strat.SetSRecordImportProperties(
        1,               # DisableRangeChecking
        0,              # EnableLimits
        0,                  # StartLimit (EnableLimits=False iken yok sayılır)
        0,                  # EndLimit   (EnableLimits=False iken yok sayılır)
        [],          # Regions (boş bırak → A2L memory regions)
        0               # CreateRegionsFromData
    )
    if hasattr(strat, "Import"):
        try:
            strat.Import(s19_path)
            print("s19 import ✅")
            return True
        except Exception as e:
            print("s19 import basarisiz", e)
    return False

def save_vst(strat, out_path):
    ensure_dir(out_path)
    # SaveAs ilk tercih
    if hasattr(strat, "SaveAs"):
        strat.SaveAs(out_path)
        return True
    # Yedek: Save() varsa önce dosyayı set eden bir metot gerekebilir; genelde SaveAs var.
    if hasattr(strat, "Save"):
        strat.Save()
        return os.path.exists(out_path)
    return False

def export_calib(strat, out_path):
    import os, pythoncom
    out_path = os.path.abspath(str(out_path))
    if not out_path.lower().endswith(".cal"):
        out_path += ".cal"
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    strat.ExportCalibration(
        ExportFileName=out_path,
        FilterFileName="",            # filtre yok
        ModificationSource="",        # kaynak filtresi yok
        ModificationFromDateTime=pythoncom.Missing,  # tarihleri tamamen atla
        ModificationToDateTime=pythoncom.Missing
    )
    return True

def open_device_properties_with_logs(dev):
    import time, win32com.client

    app = win32com.client.gencache.EnsureDispatch("Vision.Application")
    # 1) Pencerenin görünmesini garanti et
    try:
        print("Before -> Visible:", getattr(app, "Visible", None),
              "FullyAutomatedMode:", getattr(app, "FullyAutomatedMode", None))
        app.Visible = True                 # UI görünür olsun
        app.FullyAutomatedMode = False     # dialog bastırmayı KAPAT
        print("After  -> Visible:", app.Visible, "FullyAutomatedMode:", app.FullyAutomatedMode)
    except Exception as e:
        print("App visibility/automation ayarlanamadı:", e)

    # 2) (İsteğe bağlı) bu cihazı ‘current’ yap
    try:
        dev.SetCurrentDevice(dev)
        print("SetCurrentDevice OK")
    except Exception as e:
        print("SetCurrentDevice yok/başarısız:", e)

    # 3) Cihaz temel bilgilerini yazdır
    try:
        props = dev.Properties   # özelliktir, çağırılmaz
        print("Device:", props.Name, "| Enabled:", props.Enabled, "| Comments:", props.Comments)
    except Exception as e:
        print("Properties okunamadı:", e)

    # 4) Diyaloğu aç (non-blocking)
    print("EditProperties çağrılıyor…")
    dev.EditProperties()
    try:
        app.BringToFront()  # bazen arka planda açılır
    except Exception:
        pass
    time.sleep(0.5)         # kullanıcı görebilsin diye ufak bekleme
    print("EditProperties çağrısı döndü (pencere modal ama çağrı parametresiz ve geri değer yok).")


def create_project(prj, PRJ_PATH,vst):
    import os, time
    # Kaydet & açık olduğundan emin ol
    os.makedirs(os.path.dirname(PRJ_PATH), exist_ok=True)
    rc = prj.SaveAs(PRJ_PATH)
    # Proje açık değilse bir kez daha açmayı dene
    if not prj.IsOpen:
        prj.Open(PRJ_PATH)

    if not prj.IsOpen:
        raise RuntimeError("Project not open after SaveAs/Open")

    # ---- Device ağacı: Computer (RootDevice) -> USB Port -> (auto) VID ----
    root = prj.RootDevice  # Bilgisayar düğümü (device tree kökü) :contentReference[oaicite:1]{index=1}
    #dump_tree(root)
    VISION_DEVICE_USBPORT = 1   # USB Port device :contentReference[oaicite:2]{index=2}
    VISION_DEVICE_VID     = 96  # VID device (CANary) :contentReference[oaicite:3]{index=3}

    # 1) USB Port ekle
    usb = root.AddDevice(VISION_DEVICE_USBPORT)
    usb.QueryForSubDevices()
    can1 = prj.FindDevice("CANChannel1")

    open_device_properties_with_logs(can1)

    can1.AddDevice(60)
    pcm = prj.FindDevice("PCM")
    pcm.AddStrategy(vst)

    return True

def open_base_project(prj, PRJ_PATH):


    prj.Open(PRJ_PATH)
    

def main():
    if not os.path.exists(A2L_PATH):
        raise FileNotFoundError(f"A2L bulunamadı: {A2L_PATH}")

    if not os.path.exists(S19_PATH):
        raise FileNotFoundError(f"S19 bulunamadı: {S19_PATH}")
    
    pythoncom.CoInitialize()
    try:
        # Doğrudan StrategyFileInterface'e bağlan
        strat = win32com.client.DispatchEx("Vision.StrategyFileInterface")
        print("✅ StrategyFileInterface bağlı.")

        prj = win32com.client.gencache.EnsureDispatch("Vision.ProjectInterface")

        print("✅ ProjectInterface bağlı.")

        #calib = win32com.client.gencache.EnsureDispatch("Vision.CalibrationInterface")

        print("✅ DeviceInterface bağlı.")
        
            # 2) Cihazları ekle: önce API Port, sonra Virtual PCM
        #root.AddDevice(VISION_DEVICE_USBPORT)     # :contentReference[oaicite:8]{index=8}
        #root.AddDevice(VISION_DEVICE_VIRTUALPCM)  # :contentReference[oaicite:9]{index=9}

        # (İsteğe bağlı) ASAP2 import ayarları — çoğu durumda gerekmez, varsayılanlar kullanılır
        # Örn. eğer sende bu metotlar varsa ve kullanmak istersen:
        # if hasattr(strat, "SetASAP2ImportProperties2"):
        #     # parametre imzasını bilmiyorsan dokunma; defaults gayet çalışır.
        #     pass

        # A2L içe aktar

        if not import_a2l(strat, A2L_PATH):
            raise RuntimeError("A2L import edilemedi (Import başarisiz).")
        
        if not import_s19(strat, S19_PATH):
            raise RuntimeError("S19 import edilemedi (Import başarisiz).")
        
        # VST kaydet
        if not save_vst(strat, VST_OUT):
            raise RuntimeError("VST kaydedilemedi (SaveAs/Save başarisiz).")

        if not export_calib(strat, CAL_OUT):
           raise RuntimeError("VST kaydedilemedi (SaveAs/Save başarisiz).")
        
        open_base_project(prj,PRJ_OUT)

        pcm = prj.FindDevice("PCM")
        pcm.AddStrategy(strat)
        
        pcm.EnableAutoDownload = False
        pcm.DisableAutoSync = True
        prj.Online = True
        
        time.sleep(4)
        vst_path = os.path.abspath(VST_OUT)
        pcm.UploadActiveStrategy(vst_path)
        breakCount = 0
        
        while True:
            state = pcm.State  # VISION_DEVICE_STATE_CODES
            if state == 9:  # VISION_DEVICE_UPLOADING
                print("Upload devam ediyor...")
            elif state == 5:  # VISION_DEVICE_ONLINE
                print("Upload tamamlandı.")
                break
            else:
                print(f"Durum: {state}")
                breakCount+=1
                if breakCount == 5:
                    break
            time.sleep(1)

        save_vst(strat, VST_OUT)
        #create_project(prj, PRJ_OUT,strat)


        print(f"✅ Bitti.\n VST: {VST_OUT}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
