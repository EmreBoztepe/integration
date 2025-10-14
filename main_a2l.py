# a2l_addr_fill.py
from pathlib import Path
import re
import csv
import argparse
from elftools.elf.elffile import ELFFile
from elftools.elf.sections import SymbolTableSection

# Satırda bir HEX adres + @ECU_Address@<Name>@ yorumu varsa yakalar.
# Hem "ECU_ADDRESS 0x0000 ..." hem de "/* ECU Address */ 0x0000 ..." tiplerini kapsar.
LINE_RE = re.compile(
    r'^(?P<prefix>.*?\b)'                         # adres öncesi her şey
    r'(?P<addr>0x[0-9A-Fa-f]+)'                   # hex adres
    r'(?P<suffix>.*?/\*\s*@ECU_Address@'          # yorum başlangıcı + etiket
    r'(?P<name>[^@]+)@\s*\*/.*)$'                 # parametre adı + yorum bitişi
)

def build_symbol_map(elf_path: Path):
    sym = {}
    with elf_path.open("rb") as f:
        elf = ELFFile(f)
        for sec in elf.iter_sections():
            if isinstance(sec, SymbolTableSection):
                for s in sec.iter_symbols():
                    nm = s.name or ""
                    if nm:
                        # Son görüleni bırak (genelde .symtab > .dynsym)
                        sym[nm] = s.entry["st_value"]
    return sym

def resolve_address(symmap: dict, pname: str):
    # Önce mtlb_<name>, olmazsa <name>
    for key in (f"mtlb_{pname}", pname):
        if key in symmap:
            return symmap[key], key
    return None, None

def process(a2l_in: Path, a2l_out: Path, symmap: dict, csv_out: Path):
    lines = a2l_in.read_text(encoding="utf-8", errors="ignore").splitlines()
    resolved, missing, unchanged = [], [], []
    new_lines = []

    for ln in lines:
        m = LINE_RE.match(ln)
        if not m:
            new_lines.append(ln)
            continue

        current = m.group("addr")
        pname = m.group("name").strip()

        # Sadece 0x0000/0x0 ise değiştir (sende hepsi böyle)
        if current.lower() not in ("0x0000", "0x0"):
            unchanged.append((pname, current))
            new_lines.append(ln)
            continue

        addr_val, used_key = resolve_address(symmap, pname)
        if addr_val is not None:
            addr_hex = f"0x{addr_val:X}"
            new_ln = f"{m.group('prefix')}{addr_hex}{m.group('suffix')}"
            new_lines.append(new_ln)
            resolved.append((pname, addr_hex, used_key))
        else:
            new_lines.append(ln)  # bulamadık, olduğu gibi bırak
            missing.append(pname)

    a2l_out.write_text("\n".join(new_lines), encoding="utf-8")

    # CSV özet
    with open(csv_out, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["ParameterName", "Result", "AddressOrKey"])
        for n, a, k in resolved:
            w.writerow([n, "RESOLVED", f"{a} (via {k})"])
        for n in missing:
            w.writerow([n, "MISSING", "symbol not found (tried mtlb_<name> and <name>)"])
        for n, a in unchanged:
            w.writerow([n, "UNCHANGED_NONZERO", a])

    return len(lines), len(resolved), len(missing), len(unchanged)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--elf", required=True, help="ELF file path")
    ap.add_argument("--in", dest="a2l_in", required=True, help="Input A2L")
    ap.add_argument("--out", dest="a2l_out", required=True, help="Output A2L")
    ap.add_argument("--csv", dest="csv_out", default="a2l_address_resolution_summary.csv")
    args = ap.parse_args()

    elf_path = Path(args.elf)
    a2l_in  = Path(args.a2l_in)
    a2l_out = Path(args.a2l_out)
    csv_out = Path(args.csv_out)

    assert elf_path.exists(), f"ELF bulunamadı: {elf_path}"
    assert a2l_in.exists(), f"A2L bulunamadı: {a2l_in}"

    symmap = build_symbol_map(elf_path)
    total, nres, nmiss, nunch = process(a2l_in, a2l_out, symmap, csv_out)
    print(f"Processed lines: {total}")
    print(f"Resolved: {nres}  Missing: {nmiss}  Already non-zero: {nunch}")

if __name__ == "__main__":
    main()
