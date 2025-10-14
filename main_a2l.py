#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path
import re
import csv
import argparse
from typing import Optional, Tuple

from elftools.elf.elffile import ELFFile
from elftools.elf.sections import SymbolTableSection
from elftools.dwarf.descriptions import describe_form_class

# A2L satırı: ... 0x0000 /* @ECU_Address@ParamName@ */
LINE_RE = re.compile(
    r'^(?P<prefix>.*?\b)'
    r'(?P<addr>0x[0-9A-Fa-f]+)'
    r'(?P<suffix>.*?/\*\s*@ECU_Address@(?P<name>[^@]+)@\s*\*/.*)$'
)

# ---------- ELF sembolleri ----------

def build_symbol_map(elf: ELFFile) -> dict:
    sym = {}
    for sec in elf.iter_sections():
        if isinstance(sec, SymbolTableSection):
            for s in sec.iter_symbols():
                nm = s.name or ""
                if nm:
                    sym[nm] = s.entry["st_value"]
    return sym

def resolve_direct_symbol(symmap: dict, pname: str) -> Optional[Tuple[int, str]]:
    """Düz isim için: önce mtlb_<name>, yoksa <name>. Bulunursa (addr, kullanılan_anahtar) döner."""
    for key in (f"mtlb_{pname}", pname):
        if key in symmap:
            return symmap[key], key
    return None

# ---------- DWARF yardımcıları (tek seviyeli struct.member) ----------

def ref_to_die(dwarfinfo, die, attr_name):
    """DW_AT_type gibi referansları çözer: önce CU-relative, olmazsa global dener."""
    attr = die.attributes.get(attr_name)
    if not attr:
        return None
    val = attr.value
    # 1) CU-relative
    for off in (die.cu.cu_offset + val, val):
        try:
            d = dwarfinfo.get_DIE_from_refaddr(off)
            if d:
                return d
        except Exception:
            pass
    return None

def follow_type(die, dwarfinfo):
    """typedef/const/volatile/restrict sarmallarını indirip gerçek tipe ulaş."""
    t = die
    while True:
        nxt = ref_to_die(dwarfinfo, t, 'DW_AT_type')
        if nxt is None:
            return t
        if nxt.tag in ('DW_TAG_typedef','DW_TAG_const_type','DW_TAG_volatile_type','DW_TAG_restrict_type'):
            t = nxt
            continue
        return nxt

def parse_uleb128(data: bytes, idx=0):
    val = 0
    shift = 0
    i = idx
    while i < len(data):
        b = data[i]
        val |= (b & 0x7F) << shift
        i += 1
        if (b & 0x80) == 0:
            break
        shift += 7
    return val, i

def parse_member_location(loc_attr) -> Optional[int]:
    """
    DW_AT_data_member_location -> ofset (int)
    Desteklenenler:
      - constant
      - exprloc / block: DW_OP_plus_uconst(0x23), DW_OP_constu(0x10), DW_OP_lit0..31(0x30..0x4F)
    """
    if not loc_attr:
        return 0
    form = describe_form_class(loc_attr.form)

    if form == 'constant':
        return int(loc_attr.value)

    if form in ('exprloc', 'block'):
        expr = loc_attr.value or b""
        if not expr:
            return None
        i = 0
        offset = 0
        while i < len(expr):
            op = expr[i]; i += 1
            # DW_OP_lit0 .. DW_OP_lit31
            if 0x30 <= op <= 0x4F:
                offset = (op - 0x30)
                continue
            # DW_OP_constu
            if op == 0x10:
                val, i = parse_uleb128(expr, i)
                offset = val
                continue
            # DW_OP_plus_uconst
            if op == 0x23:
                val, i = parse_uleb128(expr, i)
                offset += val
                continue
            # Destek dışı başka op görürsek şimdilik vazgeç
            return None
        return offset

    # Diğer formlar: yok say
    return None

def find_global_var_die(dwarfinfo, name: str):
    """Top-level DW_TAG_variable && DW_AT_name == name olan DIE'i bul."""
    for cu in dwarfinfo.iter_CUs():
        top = cu.get_top_DIE()
        for d in top.iter_children():
            if d.tag == 'DW_TAG_variable':
                nm = d.attributes.get('DW_AT_name')
                if nm and nm.value.decode(errors='ignore') == name:
                    return d
    return None

def member_offset_in_struct(struct_die, member_name: str) -> Optional[int]:
    for child in struct_die.iter_children():
        if child.tag != 'DW_TAG_member':
            continue
        nm = child.attributes.get('DW_AT_name')
        if not nm:
            continue
        if nm.value.decode(errors='ignore') != member_name:
            continue
        loc = child.attributes.get('DW_AT_data_member_location')
        return parse_member_location(loc)
    return None

def resolve_struct_member_addr(elf: ELFFile, dwarfinfo, symmap: dict, dotted_name: str) -> Optional[Tuple[int, str]]:
    """
    'Base.member' için:
      - Base adresini yalnız TAM ADLA (mtlb_ yok) symmap'ten al
      - DWARF'ta Base global değişkenini bul (DW_TAG_variable)
      - Base tipi struct olmalı; member ofsetini al
      - adres = base + ofset
    Döner: (addr, "Base+DWARF(<off>)")
    """
    if '.' not in dotted_name or dwarfinfo is None:
        return None
    base, member = dotted_name.split('.', 1)

    base_addr = symmap.get(base)
    if base_addr is None:
        return None

    var_die = find_global_var_die(dwarfinfo, base)
    if not var_die:
        return None

    t_die = ref_to_die(dwarfinfo, var_die, 'DW_AT_type')
    if not t_die:
        return None
    t_die = follow_type(t_die, dwarfinfo)
    if t_die.tag != 'DW_TAG_structure_type':
        return None

    off = member_offset_in_struct(t_die, member)
    if off is None:
        return None

    return base_addr + off, f"{base}+DWARF({off})"

# ---------- A2L işleme ----------

def process_a2l(a2l_in: Path, a2l_out: Path, elf: ELFFile, symmap: dict, csv_out: Path):
    lines = a2l_in.read_text(encoding="utf-8", errors="ignore").splitlines()
    dwarfinfo = elf.get_dwarf_info()  # struct.member çözümü için
    resolved, missing, unchanged = [], [], []
    new_lines = []

    for ln in lines:
        m = LINE_RE.match(ln)
        if not m:
            new_lines.append(ln)
            continue

        current = m.group("addr")
        pname = m.group("name").strip()

        # Sadece 0x0000 ise değiştir
        if current.lower() not in ("0x0000", "0x0"):
            unchanged.append((pname, current))
            new_lines.append(ln)
            continue

        # 1) struct.member: mtlb_ deneme, doğrudan base + DWARF offset
        if '.' in pname:
            res = resolve_struct_member_addr(elf, dwarfinfo, symmap, pname)
            if res:
                addr, note = res
                new_lines.append(f"{m.group('prefix')}0x{addr:X}{m.group('suffix')}")
                resolved.append((pname, f"0x{addr:X}", note, "STRUCT_MEMBER"))
                continue

        # 2) düz isim: önce mtlb_<name>, yoksa <name>
        direct = resolve_direct_symbol(symmap, pname)
        if direct:
            addr, used = direct
            new_lines.append(f"{m.group('prefix')}0x{addr:X}{m.group('suffix')}")
            resolved.append((pname, f"0x{addr:X}", used, "DIRECT"))
            continue

        # 3) bulunamadı
        new_lines.append(ln)
        missing.append(pname)

    a2l_out.write_text("\n".join(new_lines), encoding="utf-8")

    with csv_out.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["ParameterName", "Result", "AddressOrNote", "Mode"])
        for n, a, note, mode in resolved:
            w.writerow([n, "RESOLVED", f"{a} ({note})", mode])
        for n in missing:
            w.writerow([n, "MISSING", "symbol not found (needs DWARF for struct members or missing symbol)", ""])
        for n, a in unchanged:
            w.writerow([n, "UNCHANGED_NONZERO", a, ""])

# ---------- main ----------

def main():
    ap = argparse.ArgumentParser(description="A2L ECU_ADDRESS doldurucu (pyelftools)")
    ap.add_argument("--elf", required=True, help="ELF dosyası (struct için DWARF gerekli)")
    ap.add_argument("--in",  dest="a2l_in",  required=True, help="Girdi A2L")
    ap.add_argument("--out", dest="a2l_out", required=True, help="Çıktı A2L")
    ap.add_argument("--csv", dest="csv_out", default="a2l_address_resolution_summary.csv", help="Özet CSV")
    args = ap.parse_args()

    elf_path = Path(args.elf)
    a2l_in  = Path(args.a2l_in)
    a2l_out = Path(args.a2l_out)
    csv_out = Path(args.csv_out)

    assert elf_path.exists(), f"ELF bulunamadı: {elf_path}"
    assert a2l_in.exists(),  f"A2L bulunamadı: {a2l_in}"

    with elf_path.open("rb") as f:
        elf = ELFFile(f)
        symmap = build_symbol_map(elf)
        process_a2l(a2l_in, a2l_out, elf, symmap, csv_out)

if __name__ == "__main__":
    main()
