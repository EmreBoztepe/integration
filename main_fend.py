import winreg as wr
from contextlib import suppress

ROOTS = [wr.HKEY_CLASSES_ROOT]  # HKCR

def iter_subkeys(hkey, path):
    try:
        with wr.OpenKey(hkey, path) as k:
            i = 0
            while True:
                try:
                    name = wr.EnumKey(k, i)
                except OSError:
                    break
                yield name
                i += 1
    except FileNotFoundError:
        return

def read_str(hkey, path, value=""):
    with suppress(FileNotFoundError, OSError):
        with wr.OpenKey(hkey, path) as k:
            v, _ = wr.QueryValueEx(k, value)
            return v
    return None

def get_server_info(clsid):
    base = rf"CLSID\{clsid}"
    inproc = read_str(wr.HKEY_CLASSES_ROOT, base + r"\InprocServer32")
    local  = read_str(wr.HKEY_CLASSES_ROOT, base + r"\LocalServer32")
    threading = read_str(wr.HKEY_CLASSES_ROOT, base + r"\InprocServer32", "ThreadingModel")
    if inproc:
        return "InprocServer32 (DLL)", inproc, threading
    if local:
        return "LocalServer32 (EXE)", local, None
    return "Unknown", None, None

def main():
    results = []
    # HKCR altında Vision.* ProgID’lerini tara
    for name in iter_subkeys(wr.HKEY_CLASSES_ROOT, ""):
        if not name.lower().startswith("vision."):
            continue
        # ProgID altında CLSID var mı?
        clsid = read_str(wr.HKEY_CLASSES_ROOT, rf"{name}\CLSID")
        if not clsid:
            # Kimi ProgID’ler default alt değerde CLSID tutabilir, deneyelim
            clsid = read_str(wr.HKEY_CLASSES_ROOT, rf"{name}")
        if not clsid or not clsid.startswith("{"):
            continue
        server_type, server_path, threading = get_server_info(clsid)
        results.append((name, clsid, server_type, server_path or "-", threading or "-"))

    # Çıktıyı düzenli yazdır
    if not results:
        print("Vision.* altında ProgID bulunamadı.")
        return

    width = [max(len(str(col)) for col in colset) for colset in zip(
        ["ProgID","CLSID","ServerType","ServerPath","ThreadingModel"],
        *results
    )]
    header = ["ProgID","CLSID","ServerType","ServerPath","ThreadingModel"]
    print(
        f"{header[0]:<{width[0]}}  {header[1]:<{width[1]}}  {header[2]:<{width[2]}}  {header[3]:<{width[3]}}  {header[4]:<{width[4]}}"
    )
    print("-" * (sum(width) + 8))
    for r in results:
        print(f"{r[0]:<{width[0]}}  {r[1]:<{width[1]}}  {r[2]:<{width[2]}}  {r[3]:<{width[3]}}  {r[4]:<{width[4]}}")

if __name__ == "__main__":
    main()
