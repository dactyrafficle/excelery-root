'https://docs.microsoft.com/en-us/office365/troubleshoot/compile-error-editing-vba-macro

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

'this thing will run every 250ms
Sub test()
Dim i As Long

For i = 1 To 5

    'Debug.Print Now()
    Range("a1").Offset(i - 1, 0).Value = Rnd()
    Sleep 250    'wait 0.25 seconds
    
Next i
End Sub
