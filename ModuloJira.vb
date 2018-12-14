Sub AtualizarJira()
   
    Dim hReq As Object
    Dim var As Variant
    Dim jira As Variant
    
    Dim element As Object
    Dim elements As Object
    Dim e As Variant
    
    Dim ws As Worksheet
    Set ws = Worksheets("JIRA")
        ws.Activate
    
    Dim i As Long
    i = 2
    Do While Not IsEmpty(ws.Range("B" & i))
    
        ws.Range("O" & i).Select
        
        jira = ws.Range("B" & i).value
    
        'create our URL string and pass the user entered information to it
        Dim strUrl As String
            strUrl = "http://jira.vivo.com.br/si/jira.issueviews:issue-xml/" & jira & "/" & jira & ".xml"
        
        Set hReq = CreateObject("MSXML2.XMLHTTP")
            With hReq
                .Open "GET", strUrl, False
                .Send
            End With
          
        'use the LoadXML method to load a known XML string
        Set xmlDoc = CreateObject("MSXML2.DOMDocument")
            xmlDoc.LoadXML hReq.ResponseText
          
        ' status
        Set elements = xmlDoc.getElementsByTagName("status")
        For Each e In elements
            Debug.Print jira & " (status): " & e.Text
            ws.Range("F" & i).value = UCase(e.Text)
            'color
            If UCase(ws.Range("F" & i).value) <> UCase(ws.Range("E" & i).value) Then
                ws.Range("F" & i).Interior.ColorIndex = 6
                'ws.Range("F" & i).Font.Color = 2
            Else
                ws.Range("F" & i).Interior.ColorIndex = 0
                'ws.Range("F" & i).Font.Color = 1
            End If
        Next
        
        ' title
        Set elements = xmlDoc.getElementsByTagName("title")
        For Each e In elements
            ws.Range("N" & i).value = e.Text
            Debug.Print jira & " (title): " & e.Text
        Next
        
        ' pontos
        Set element = xmlDoc.SelectSingleNode("//*[@id = 'customfield_11110']")
        If Not (element Is Nothing) Then
            Set elements = element.getElementsByTagName("customfieldvalue")
            For Each e In elements
                Debug.Print jira & " (pontos): " & e.Text
                ws.Range("H" & i).value = e.Text
                'color
                If ws.Range("H" & i).value <> ws.Range("G" & i).value Then
                    ws.Range("H" & i).Interior.ColorIndex = 6
                    'ws.Range("H" & i).Font.Color = 2
                Else
                    ws.Range("H" & i).Interior.ColorIndex = 0
                    'ws.Range("H" & i).Font.Color = 1
                End If
            Next
        Else
            Debug.Print jira & " (pontos): <Nothing>"
        End If
            
        i = i + 1
    Loop
      
    'clear
    Set var = Nothing
    Set hReq = Nothing
    
End Sub
