' Virtual PCM örneği (temizlenmiş)
Option Explicit

Const VISION_DATAITEM_STRING   = 12
Const VISION_DEVICE_VIRTUALPCM = 2
Const VISION_DEVICE_APIPORT    = 1

Dim objStdOut : Set objStdOut = WScript.StdOut

Dim Project : Set Project = CreateObject("Vision.ProjectInterface")
Dim V_App   : Set V_App   = CreateObject("Vision.Application")
Dim StrategyFileIntf : Set StrategyFileIntf = CreateObject("Vision.StrategyFileInterface")

If Project.IsOpen Then
  Dim Device : Set Device = Project.FindDevice("PCM")
  Dim Virtual_API_Port : Set Virtual_API_Port = Project.FindDevice("VISION API Port")
  Dim Virtual_Device, ActiveStrategy, Virtual_ActiveStrategy
  Dim FullStrategyFileName, BaseVirtualFile, VirtualFileName, FullVirtualStrategyFileName

  ' Eğer Virtual API Port yoksa ekle ve altına Virtual PCM ekle
  If Virtual_API_Port Is Nothing Then
    objStdOut.WriteLine "No Virtual API Port exists...adding one."
    Dim CurrentProject : Set CurrentProject = V_App.GetCurrentProjectInterface
    Dim MyRootDevice   : Set MyRootDevice   = CurrentProject.RootDevice

    Set Virtual_API_Port = MyRootDevice.AddDevice(VISION_DEVICE_APIPORT)
    If Virtual_API_Port Is Nothing Then
      objStdOut.WriteLine "Unable to add a Virtual API Port - exiting!!"
      WScript.Quit 1
    End If

    Set Virtual_Device = Virtual_API_Port.AddDevice(VISION_DEVICE_VIRTUALPCM)
    If Virtual_Device Is Nothing Then
      objStdOut.WriteLine "Unable to add a Virtual PCM - exiting!!!"
      WScript.Quit 1
    End If
  Else
    ' Varsa mevcut Virtual PCM’i bulmaya çalış
    Set Virtual_Device = Project.FindDevice("Virtual PCM")
    If Virtual_Device Is Nothing Then
      Set Virtual_Device = Virtual_API_Port.AddDevice(VISION_DEVICE_VIRTUALPCM)
    End If
  End If

  ' Aktif stratejileri al
  Set ActiveStrategy         = Device.ActiveStrategy
  Set Virtual_ActiveStrategy = Virtual_Device.ActiveStrategy

  ' Sanal tarafta aktif strateji yoksa: gerçek .vst’den sanal kopya üret
  If Virtual_ActiveStrategy Is Nothing Then
    FullStrategyFileName = Project.Directory & "\" & Device.Name & "\" & ActiveStrategy.Name & ".vst"
    Virtual_Device.AddStrategy FullStrategyFileName

    BaseVirtualFile = ActiveStrategy.Name & "_virtual"
    VirtualFileName = BaseVirtualFile & ".vst"
    FullVirtualStrategyFileName = Project.Directory & "\" & Virtual_Device.Name & "\" & VirtualFileName

    Virtual_Device.ActiveStrategySaveAs FullVirtualStrategyFileName
    Virtual_Device.RemoveStrategy FullVirtualStrategyFileName   ' refresh için kapat-aç
    Virtual_Device.AddStrategy FullVirtualStrategyFileName
  Else
    ' Zaten sanal strateji var; dosya yolunu yeniden kur
    VirtualFileName = Virtual_ActiveStrategy.Name & ".vst"
    FullVirtualStrategyFileName = Project.Directory & "\" & Virtual_Device.Name & "\" & VirtualFileName
  End If

  ' Sanal stratejiyi aç ve metin item’ları oluştur
  StrategyFileIntf.Open FullVirtualStrategyFileName

  Dim BaseGroup : Set BaseGroup = StrategyFileIntf.GroupDataItem
  If BaseGroup Is Nothing Then
    objStdOut.WriteLine "Measurement Group Not found"
    WScript.Quit 1
  Else
    If BaseGroup.FindDataItem("Active_Strategy") Is Nothing Then
      objStdOut.WriteLine "Creating a Virtual Active Strategy data item"
      BaseGroup.CreateItem VISION_DATAITEM_STRING, "Active_Strategy"
    End If
    If BaseGroup.FindDataItem("Active_Calibration") Is Nothing Then
      objStdOut.WriteLine "Creating a Virtual Active Calibration data item"
      BaseGroup.CreateItem VISION_DATAITEM_STRING, "Active_Calibration"
    End If
    StrategyFileIntf.Save
    Virtual_Device.ActiveStrategy.Reload
  End If

  ' Metin item’larını linkle ve değerleri güncelle
  Dim Virtual_PCM_ActiveStrategy : Set Virtual_PCM_ActiveStrategy = Virtual_Device.FindTextString("Active_Strategy")
  Dim Virtual_PCM_ActiveCalibration : Set Virtual_PCM_ActiveCalibration = Virtual_Device.FindTextString("Active_Calibration")

  objStdOut.WriteLine "Linking Virtual items..."
  Virtual_PCM_ActiveStrategy.ActualValue    = ActiveStrategy.Name
  Virtual_PCM_ActiveCalibration.ActualValue = ActiveStrategy.ActiveCalibration.Name

  Project.Online = True
  objStdOut.WriteLine "Updating the Virtual PCM Text String Data items every second"
  Do While Project.Online = True
    WScript.Sleep 1000
    Virtual_PCM_ActiveCalibration.ActualValue = Device.ActiveStrategy.ActiveCalibration.Name
    Virtual_PCM_ActiveStrategy.ActualValue    = Device.ActiveStrategy.Name
  Loop
  objStdOut.WriteLine "Done looping for now!!"

Else
  objStdOut.WriteLine "No project open - please remedy"
End If
