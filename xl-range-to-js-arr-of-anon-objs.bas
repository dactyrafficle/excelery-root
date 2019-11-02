Attribute VB_Name = "Module1"

'this will convert the contiguous block of cells from a1 into a js array of objects

Sub abc()

'first go to tools > references > microsoft scripting runtime

Dim fso As Scripting.FileSystemObject
Set fso = New Scripting.FileSystemObject

Dim f As Scripting.TextStream
Set f = fso.CreateTextFile(ThisWorkbook.Path & "\" & Date & "-" & Hour(Now) & "-" & Minute(Now) & "-" & Second(Now) & ".json", 1, 0)

'oddly, both key and data will be 2d arrays

'get the key names
Dim key As Variant
key = Range("a1", Range("a1").End(xlToRight)).Value

'get the values (named data to avoid overlapping names and confusion)
Dim data As Variant
data = Range("a2", Range("a2").End(xlToRight).End(xlDown)).Value

Dim arr_str As String
arr_str = "["

Dim x, y As Long
For y = 1 To UBound(data, 1)

    Dim str As String
    str = "{"

    For x = 1 To UBound(data, 2)
    
        str = str & Chr(34) & key(1, x) & Chr(34) & ":" & Chr(34) & data(y, x) & Chr(34)
        
        If x <> UBound(data, 2) Then
        
            str = str & ","
        
        End If
        
    
    Next x
    
    str = str & "}"

    Debug.Print str
    
    If y <> UBound(data, 1) Then
    
        str = str & ","
    
    End If
    
    arr_str = arr_str & str
    

Next y

arr_str = arr_str & "]"

f.WriteLine arr_str

End Sub
