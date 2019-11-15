Sub abc()

    'Dim xhr As Object
    'Set xhr = CreateObject("MSXML2.serverXMLHTTP")
    
    'or tools > references > microsoft xml etc etc. (just like fso)
    Dim xhr As New MSXML2.XMLHTTP60
    
    Dim url As String
    'url = "http://api.openweathermap.org/data/2.5/weather?q=London,uk&appid=0c5b40099f42275292567c7af6d887b9"
    url = "http://shakespeare.mit.edu/romeo_juliet/romeo_juliet.1.0.html"
    xhr.Open "GET", url, False
    xhr.send
    
    'tools > references > microsoft html doc
    'lots to figure out here
    Dim html As New HTMLDocument
    html.body.innerhtml = xhr.responseText 'does nothing here atm
    
    'in this case, the html document comes back as a big string - so i have to parse it
    Debug.Print xhr.responseText

End Sub
