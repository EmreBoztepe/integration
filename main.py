# make_vst.py — A2L -> VST (UI'siz, doğrudan StrategyFileInterface)
import os
import pythoncom
import win32com.client
from win32com.client import VARIANT
# >>> BURAYI DÜZENLE
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
A2L_PATH   = os.path.join(SCRIPT_DIR, "example.a2l")     # kendi .a2l dosyan
VST_OUT    = os.path.join(SCRIPT_DIR, "out", "MyECU.vst")  # çıkış .vst
CAL_OUT    = os.path.join(SCRIPT_DIR, "out", "MyECU.cal")  # çıkış .cal
PRJ_OUT    = os.path.join(SCRIPT_DIR, "out", "MyECU.vpj")  # çıkış .cal
S19_PATH   = os.path.join(SCRIPT_DIR, "example.s19") #s19 dosyası.

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

def create_project(prj, PRJ_PATH):
    prj.SaveAs(PRJ_PATH)
    return True


def main():
    if not os.path.exists(A2L_PATH):
        raise FileNotFoundError(f"A2L bulunamadı: {A2L_PATH}")

    if not os.path.exists(S19_PATH):
        raise FileNotFoundError(f"S19 bulunamadı: {S19_PATH}")
    
    pythoncom.CoInitialize()
    try:
        # Doğrudan StrategyFileInterface'e bağlan
        strat = win32com.client.gencache.EnsureDispatch("Vision.StrategyFileInterface")
        print("✅ StrategyFileInterface bağlı.")

        prj = win32com.client.gencache.EnsureDispatch("Vision.ProjectInterface")
        print("✅ ProjectInterface bağlı.")


        if not create_project(prj,PRJ_OUT):
            raise RuntimeError("A2L import edilemedi (Import başarisiz).")
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
        
        print(f"✅ Bitti.\n VST: {VST_OUT}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
