# make_vst.py — A2L -> VST (UI'siz, doğrudan StrategyFileInterface)
import os
import pythoncom
import win32com.client

# >>> BURAYI DÜZENLE
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
A2L_PATH   = os.path.join(SCRIPT_DIR, "example.a2l")     # kendi .a2l dosyan
VST_OUT    = os.path.join(SCRIPT_DIR, "out", "MyECU.vst")  # çıkış .vst

def ensure_dir(p):
    d = os.path.dirname(p)
    if d and not os.path.exists(d):
        os.makedirs(d, exist_ok=True)

def import_a2l_try_all(strat, a2l_path):
    """
    Bazı kurulumlarda ImportAs(typeHint) istenir.
    Sırayla birkaç tür adı deneriz; olmazsa Import() çağırırız.
    """
    candidates = ["ASAP2", "A2L", "ASAP", "ASAP2_FILE"]
    # Önce ImportAs varsa onu dene
    if hasattr(strat, "ImportAs"):
        for hint in candidates:
            try:
                print(f"[A2L] ImportAs('{hint}', '{a2l_path}')")
                strat.ImportAs(a2l_path, hint)
                return True
            except Exception as e:
                print(f"   -> {hint} olmadı: {e}")
    # Son çare: Import()
    if hasattr(strat, "Import"):
        try:
            print(f"[A2L] Import('{a2l_path}')")
            strat.Import(a2l_path)
            return True
        except Exception as e:
            print("   -> Import() da başarısız:", e)
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

def main():
    if not os.path.exists(A2L_PATH):
        raise FileNotFoundError(f"A2L bulunamadı: {A2L_PATH}")

    pythoncom.CoInitialize()
    try:
        # Doğrudan StrategyFileInterface'e bağlan
        strat = win32com.client.gencache.EnsureDispatch("Vision.StrategyFileInterface")
        print("[OK] StrategyFileInterface bağlı.")

        # (İsteğe bağlı) ASAP2 import ayarları — çoğu durumda gerekmez, varsayılanlar kullanılır
        # Örn. eğer sende bu metotlar varsa ve kullanmak istersen:
        # if hasattr(strat, "SetASAP2ImportProperties2"):
        #     # parametre imzasını bilmiyorsan dokunma; defaults gayet çalışır.
        #     pass

        # A2L içe aktar
        if not import_a2l_try_all(strat, A2L_PATH):
            raise RuntimeError("A2L import edilemedi (ImportAs/Import başarısız).")
        
        # VST kaydet
        if not save_vst(strat, VST_OUT):
            raise RuntimeError("VST kaydedilemedi (SaveAs/Save başarısız).")

        print(f"✅ Bitti. VST: {VST_OUT}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
