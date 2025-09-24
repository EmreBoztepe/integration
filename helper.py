import pythoncom, win32com.client
pythoncom.CoInitialize()
try:
    strat = win32com.client.gencache.EnsureDispatch("Vision.StrategyFileInterface")
    print(strat.ImportAs.__doc__)
finally:
    pythoncom.CoUninitialize()