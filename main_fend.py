# dump_any_api.py
import pythoncom, win32com.client

def dump(obj, title):
    print(f"\n=== {title} :: {obj.__class__.__name__} ===")
    ti = obj._oleobj_.GetTypeInfo()
    ta = ti.GetTypeAttr()
    for i in range(ta.cFuncs):
        fd = ti.GetFuncDesc(i)
        names = ti.GetNames(fd.memid)
        kind = {1:"method",2:"prop_get",4:"prop_put",8:"prop_putref"}.get(fd.invkind, fd.invkind)
        print(f"- {kind:10} {names[0]}({', '.join(names[1:])})")
    for mname in ("_prop_map_get_","_prop_map_put_"):
        if hasattr(obj, mname):
            print(f"\n{mname}:")
            for k in sorted(getattr(obj, mname)): print("  ", k)

pythoncom.CoInitialize()
try:
    prj  = win32com.client.gencache.EnsureDispatch("Vision.ProjectInterface")
    can1 = prj.FindDevice("CANChannel1")
    dump(can1.Properties, "CANChannel1.Properties")
    dump(can1.StrategySettings, "CANChannel1.StrategySettings")
finally:
    pythoncom.CoUninitialize()
