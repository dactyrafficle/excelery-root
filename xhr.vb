Sub abc()

    'Dim xhr As Object
    'Set xhr = CreateObject("MSXML2.serverXMLHTTP")
    
    'or tools > references > microsoft xml etc etc. (just like fso)
    Dim xhr As New MSXML2.XMLHTTP60
    
    Dim url As String
    url = "http://api.openweathermap.org/data/2.5/weather?q=London,uk&appid=0c5b40099f42275292567c7af6d887b9"
    xhr.Open "GET", url, False
    xhr.send
    
    MsgBox xhr.responseText

End Sub
