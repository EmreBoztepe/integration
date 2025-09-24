# main.py — Vision.StrategyFileInterface envanteri (tam help + üyeler)
import os
import sys
import io
import pythoncom
import win32com.client

OUT_TXT = "Vision.StrategyFileInterface_inventory.txt"

def capture_help(obj) -> str:
    # help() çıktısını stringe yakala
    buf = io.StringIO()
    stdout_old = sys.stdout
    try:
        sys.stdout = buf
        help(obj)
    finally:
        sys.stdout = stdout_old
    return buf.getvalue()

def main():
    pythoncom.CoInitialize()
    try:
        # Makepy’li sınıfı garanti altına al
        strat = win32com.client.gencache.EnsureDispatch("Vision.StrategyFileInterface")
        print("[OK] Dispatch: Vision.StrategyFileInterface")

        # 1) Tam help çıktısı (metotlar, açıklamalar, CLSID vs.)
        help_text = capture_help(strat)

        # 2) Tüm üye adları (filtre yok)
        all_members = sorted(dir(strat))

        # 3) Property haritaları (varsa)
        prop_get = getattr(strat, "_prop_map_get_", {})
        prop_put = getattr(strat, "_prop_map_put_", {})

        # Konsola kısa özet
        print("\n--- MEMBERS (names) ---")
        for n in all_members:
            print(" ", n)

        # Dosyaya ayrıntılı dök
        with open(OUT_TXT, "w", encoding="utf-8") as f:
            f.write("=== help(strat) ===\n")
            f.write(help_text)
            f.write("\n\n=== dir(strat) — all members ===\n")
            for n in all_members:
                f.write(n + "\n")
            f.write("\n\n=== _prop_map_get_ (readable properties) ===\n")
            for k, v in prop_get.items():
                f.write(f"{k}: {v}\n")
            f.write("\n\n=== _prop_map_put_ (writable properties) ===\n")
            for k, v in prop_put.items():
                f.write(f"{k}: {v}\n")

        print(f"\n✅ Bitti. Ayrıntılı çıktı: {os.path.abspath(OUT_TXT)}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
