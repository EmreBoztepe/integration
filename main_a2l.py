#!/usr/bin/env python3
from pathlib import Path
import re, csv, argparse
from typing import Optional, Tuple
from elftools.elf.elffile import ELFFile
from elftools.elf.sections import SymbolTableSection
from elftools.dwarf.descriptions import describe_form_class

LINE_RE = re.compile(r'^(?P<prefix>.*?\b)(?P<addr>0x[0-9A-Fa-f]+)(?P<suffix>.*?/\*\s*@ECU_Address@(?P<name>[^@]+)@\s*\*/.*)$')
BASE_INDEX_RE = re.compile(r'^(?P<base>[A-Za-z0-9_.$:+-]+)(\[(?P<idx>\d+)\])?$')

def build_symbol_map(elf: ELFFile) -> dict:
    sym = {}
    for sec in elf.iter_sections():
        if isinstance(sec, SymbolTableSection):
            for s in sec.iter_symbols():
                nm = s.name or ""
                if nm: sym[nm] = s.entry["st_value"]
    return sym

def resolve_direct_symbol(symmap: dict, pname: str) -> Optional[Tuple[int, str]]:
    for key in (f"mtlb_{pname}", pname):
        if key in symmap: return symmap[key], key
    return None

def ref_to_die(dwarfinfo, die, attr_name):
    attr = die.attributes.get(attr_name)
    if not attr: return None
    val = attr.value
    for off in (die.cu.cu_offset + val, val):
        try:
            d = dwarfinfo.get_DIE_from_refaddr(off)
            if d: return d
        except Exception: pass
    return None

def follow_type(die, dwarfinfo):
    t = die
    while True:
        nxt = ref_to_die(dwarfinfo, t, 'DW_AT_type')
        if nxt is None: return t
        if nxt.tag in ('DW_TAG_typedef','DW_TAG_const_type','DW_TAG_volatile_type','DW_TAG_restrict_type'):
            t = nxt; continue
        return nxt

def parse_uleb128(data: bytes, idx=0):
    val = 0; shift = 0; i = idx
    while i < len(data):
        b = data[i]; i += 1
        val |= (b & 0x7F) << shift
        if (b & 0x80) == 0: break
        shift += 7
    return val, i

def parse_member_location(loc_attr) -> Optional[int]:
    if not loc_attr: return 0
    form = describe_form_class(loc_attr.form)
    if form == 'constant': return int(loc_attr.value)
    if form in ('exprloc','block'):
        expr = loc_attr.value or b""
        i = 0; off = 0
        while i < len(expr):
            op = expr[i]; i += 1
            if 0x30 <= op <= 0x4F: off = (op - 0x30); continue         # DW_OP_lit0..31
            if op == 0x10: val, i = parse_uleb128(expr,i); off = val; continue  # DW_OP_constu
            if op == 0x23: val, i = parse_uleb128(expr,i); off += val; continue # DW_OP_plus_uconst
            return None
        return off
    return None

def find_global_var_die(dwarfinfo, name: str):
    for cu in dwarfinfo.iter_CUs():
        top = cu.get_top_DIE()
        for d in top.iter_children():
            if d.tag == 'DW_TAG_variable':
                nm = d.attributes.get('DW_AT_name')
                if nm and nm.value.decode(errors='ignore') == name: return d
    return None

def member_offset_in_struct(struct_die, member_name: str) -> Optional[int]:
    for child in struct_die.iter_children():
        if child.tag != 'DW_TAG_member': continue
        nm = child.attributes.get('DW_AT_name')
        if not nm: continue
        if nm.value.decode(errors='ignore') != member_name: continue
        return parse_member_location(child.attributes.get('DW_AT_data_member_location'))
    return None

def element_size_of_array(array_die, dwarfinfo) -> Optional[int]:
    arr = follow_type(array_die, dwarfinfo)
    if arr.tag != 'DW_TAG_array_type': return None
    elem_die = ref_to_die(dwarfinfo, arr, 'DW_AT_type')
    if not elem_die: return None
    elem_die = follow_type(elem_die, dwarfinfo)
    bs = elem_die.attributes.get('DW_AT_byte_size')
    if bs: return int(bs.value)
    bbs = elem_die.attributes.get('DW_AT_bit_size')
    if bbs: return (int(bbs.value) + 7)//8
    return None

def struct_size(struct_die) -> Optional[int]:
    bs = struct_die.attributes.get('DW_AT_byte_size')
    return int(bs.value) if bs else None

def resolve_struct_member_addr(elf: ELFFile, dwarfinfo, symmap: dict, dotted_name: str) -> Optional[Tuple[int, str]]:
    """Desteklenen: Base.member  ve  Base[idx].member  (idx >= 0)"""
    if '.' not in dotted_name or dwarfinfo is None: return None
    head, member = dotted_name.split('.', 1)
    m = BASE_INDEX_RE.match(head)
    if not m: return None
    base_name = m.group('base')
    idx = int(m.group('idx')) if m.group('idx') is not None else None

    base_addr = symmap.get(base_name)
    if base_addr is None: return None

    var_die = find_global_var_die(dwarfinfo, base_name)
    if not var_die: return None

    t_die = follow_type(ref_to_die(dwarfinfo, var_die, 'DW_AT_type') or var_die, dwarfinfo)

    base_ofs = 0
    if idx is None:
        if t_die.tag != 'DW_TAG_structure_type': return None
        struct_die = t_die
    else:
        if t_die.tag == 'DW_TAG_array_type':
            esize = element_size_of_array(t_die, dwarfinfo)
            if esize is None: return None
            base_ofs = idx * esize
            elem_die = follow_type(ref_to_die(dwarfinfo, t_die, 'DW_AT_type'), dwarfinfo)
            if not elem_die or elem_die.tag != 'DW_TAG_structure_type': return None
            struct_die = elem_die
        elif t_die.tag == 'DW_TAG_structure_type':
            # Stride fallback: struct boyutunu eleman adımı kabul et (yaygın yerleşim)
            sz = struct_size(t_die)
            if sz is None: return None
            base_ofs = idx * sz
            struct_die = t_die
        else:
            return None

    mem_off = member_offset_in_struct(struct_die, member)
    if mem_off is None: return None

    final_addr = base_addr + base_ofs + mem_off
    note = f"{base_name}"
    if idx is not None: note += f"[{idx}]"
    note += f"+DWARF({mem_off})"
    return final_addr, note

def process_a2l(a2l_in: Path, a2l_out: Path, elf: ELFFile, symmap: dict, csv_out: Path):
    lines = a2l_in.read_text(encoding="utf-8", errors="ignore").splitlines()
    dwarfinfo = elf.get_dwarf_info()
    resolved, missing, unchanged = [], [], []
    new_lines = []

    for ln in lines:
        m = LINE_RE.match(ln)
        if not m: new_lines.append(ln); continue
        cur = m.group("addr"); pname = m.group("name").strip()

        if cur.lower() not in ("0x0000","0x0"):
            unchanged.append((pname, cur)); new_lines.append(ln); continue

        if '.' in pname:
            r = resolve_struct_member_addr(elf, dwarfinfo, symmap, pname)
            if r:
                addr, note = r
                new_lines.append(f"{m.group('prefix')}0x{addr:X}{m.group('suffix')}")
                resolved.append((pname, f"0x{addr:X}", note, "STRUCT_MEMBER")); continue

        d = resolve_direct_symbol(symmap, pname)
        if d:
            addr, used = d
            new_lines.append(f"{m.group('prefix')}0x{addr:X}{m.group('suffix')}")
            resolved.append((pname, f"0x{addr:X}", used, "DIRECT")); continue

        new_lines.append(ln); missing.append(pname)

    a2l_out.write_text("\n".join(new_lines), encoding="utf-8")
    with csv_out.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["ParameterName","Result","AddressOrNote","Mode"])
        for n,a,note,mode in resolved: w.writerow([n,"RESOLVED",f"{a} ({note})",mode])
        for n in missing: w.writerow([n,"MISSING","symbol not found (needs DWARF or missing symbol)",""])
        for n,a in unchanged: w.writerow([n,"UNCHANGED_NONZERO",a,""])

def main():
    ap = argparse.ArgumentParser(description="A2L ECU_ADDRESS doldurucu (pyelftools, struct & array destekli)")
    ap.add_argument("--elf", required=True)
    ap.add_argument("--in", dest="a2l_in", required=True)
    ap.add_argument("--out", dest="a2l_out", required=True)
    ap.add_argument("--csv", dest="csv_out", default="a2l_address_resolution_summary.csv")
    args = ap.parse_args()
    elf_path, a2l_in, a2l_out, csv_out = Path(args.elf), Path(args.a2l_in), Path(args.a2l_out), Path(args.csv_out)
    assert elf_path.exists(), f"ELF bulunamadı: {elf_path}"
    assert a2l_in.exists(), f"A2L bulunamadı: {a2l_in}"
    with elf_path.open("rb") as f:
        elf = ELFFile(f)
        symmap = build_symbol_map(elf)
        process_a2l(a2l_in, a2l_out, elf, symmap, csv_out)

if __name__ == "__main__":
    main()
