'https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-processor

Sub getCPUInformation()
  
  Dim objWMIService As Object
  Set objWMIService = GetObject("winmgmts:")
  Dim cpu As Object
  For Each cpu In objWMIService.instancesof("Win32_Processor")
        MsgBox cpu.Name
        MsgBox cpu.CurrentClockSpeed & " Mhz"
  Next cpu
  
End Sub
