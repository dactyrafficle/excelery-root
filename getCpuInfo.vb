Sub getCPUInformation()
  
  Dim objWMIService As Object
  Set objWMIService = GetObject("winmgmts:")
  Dim cpu As Object
  For Each cpu In objWMIService.instancesof("Win32_Processor")
        MsgBox cpu.Name
        MsgBox cpu.CurrentClockSpeed & " Mhz"
  Next cpu
  
End Sub
