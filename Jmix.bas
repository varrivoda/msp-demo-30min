Attribute VB_Name = "Jmix"
Type Config
    client As String
    secret As String
    urlToken As String
'    urlJmixApplication As String
End Type

Public Type JmixToken
    token As String
    expirationDate As Date
End Type



Sub SendRequestGetUsers()
    Dim token As JmixToken
    Dim ccDate As Date
    token = LoadTokenFromDocumentProperties
    
    If token.token = "" Or CDate(token.expirationDate) <= Now Then
        MsgBox "getting new Token"
        token = GetNewAccessToken
    End If

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    http.Open "GET", "http://localhost:8080/rest/entities/User", False
    http.setRequestHeader "Authorization", "Bearer " & token.token
    
    On Error Resume Next
    http.send
    
    If Err.Number <> 0 Then
        MsgBox "Ошибка выполнения запроса: " & Err.Description
    Else
        MsgBox "Ответ: " & http.responseText
    End If
    On Error GoTo 0
    
    Set http = Nothing
End Sub

Sub SaveTokenToDocumentProperties(newToken As JmixToken)
    On Error Resume Next
    ActiveProject.CustomDocumentProperties("jmixToken").Delete
    ActiveProject.CustomDocumentProperties("jmixTokenExpirationDate").Delete
    On Error GoTo 0
    
    ActiveProject.CustomDocumentProperties.Add Name:="jmixToken", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=newToken.token
    ActiveProject.CustomDocumentProperties.Add Name:="jmixTokenExpirationDate", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=CStr(newToken.expirationDate)

    Exit Sub

ErrorHandler:
    Debug.Print "Error " & Err.Number & ": " & Err.Description
End Sub

Function LoadTokenFromDocumentProperties() As JmixToken
    On Error Resume Next
    Dim loadJmixToken As JmixToken
    
    loadJmixToken.token = ActiveProject.CustomDocumentProperties("jmixToken").Value
    loadJmixToken.expirationDate = ActiveProject.CustomDocumentProperties("jmixTokenExpirationDate").Value ' ("jmixToken")
    LoadTokenFromDocumentProperties = loadJmixToken
    On Error GoTo 0
End Function



Function GetNewAccessToken() As JmixToken

    Dim expDate As String
    expDate = ActiveProject.CustomDocumentProperties("jmixTokenExpirationDate").Value
    MsgBox "getting new Token because ols exp date is " & expDate

    Dim newJmixToken As JmixToken
    Dim httpRequest As Object
    
    Dim url As String
    Dim client As String
    Dim secret As String
    
    Dim requestBody As String
    Dim response As String
    Dim jsonResponse As Object
    Dim expiresIn As Long
    
    Dim accessToken As String
    Dim expirationDate As Date

    Dim myConfig As Config
    Call ReadConfigOrPrompt(myConfig)

    ' Данные для запроса
    url = myConfig.urlToken
    client = myConfig.client
    secret = myConfig.secret
    
    ' requestBody = "grant_type=client_credentials" ХЗ почему не работает?
    
    credentials = Base64Encode(client & ":" & secret)
    ' Создаем HTTP-запрос
    Set httpRequest = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    httpRequest.Open "POST", url, False
        httpRequest.setRequestHeader "Authorization", "Basic " & credentials
        httpRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        httpRequest.send "grant_type=client_credentials"
    
    ' Получаем и парсим ответ
    response = httpRequest.responseText
    Set jsonResponse = JsonConverter.ParseJson(response)
    
    ' Извлекаем токен доступа и время жизни
    accessToken = jsonResponse("access_token")
    expiresIn = jsonResponse("expires_in")
    
    MsgBox response
    
    ' Рассчитываем время истечения токена
    expirationDate = Now + TimeSerial(0, expiresIn \ 60, expiresIn Mod 60)
    
    newJmixToken.token = accessToken
    newJmixToken.expirationDate = expirationDate
    
    Call SaveTokenToDocumentProperties(newJmixToken)
    GetNewAccessToken = newJmixToken
    
End Function

Function Base64Encode(inData As String) As String
    Dim arrData() As Byte
    Dim objXML As Object
    Dim objNode As Object
    
    arrData = StrConv(inData, vbFromUnicode)
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    Base64Encode = objNode.Text
    
    Set objNode = Nothing
    Set objXML = Nothing
End Function

Sub ReadConfigOrPrompt(ByRef myConfig As Config)
    On Error GoTo ErrorHandler
    
    Dim client As String
    Dim secret As String
    Dim urlToken As String
    
    Dim filePath As String
    Dim fileNum As Integer
    Dim line As String
    Dim configFound As Boolean
    
    ' Используем относительный путь к конфигурационному файлу
    filePath = ActiveProject.Path & "\ConfigFile.txt"
    configFound = False

    If Not Dir(filePath) = vbNullString Then
        fileNum = FreeFile
        Open filePath For Input As #fileNum
      
        ' Читаем строки из файла
        Do Until EOF(fileNum)
            Line Input #fileNum, line ' Чтение строки
            If InStr(line, "client=") = 1 Then
                client = Mid(line, Len("client=") + 1)
            ElseIf InStr(line, "secret=") = 1 Then
                secret = Mid(line, Len("secret=") + 1)
            ElseIf InStr(line, "urlToken=") = 1 Then
                urlToken = Mid(line, Len("urlToken=") + 1)
            End If
        Loop
        
        Close #fileNum
        configFound = True
    End If

    ' Если файл не найден или не удалось прочитать данные, показываем диалог
    If Not configFound Or client = "" Or secret = "" Or urlToken = "" Then
        ShowLoginDialog client, secret, urlToken
    End If
    
    myConfig.client = client
    myConfig.secret = secret
    myConfig.urlToken = urlToken

ErrorHandler:
    If Err.Number <> 0 And Err.Number <> 20 Then
        MsgBox "Error " & Err.Number & " - " & Err.Description & " - " & Err.Source
    End If
    Resume Next

End Sub

Function ShowLoginDialog(client, secret, urlToken)
    Dim loginForm As New UserForm1 '
    loginForm.Show vbModal ' открываем форму модально
    
    If loginForm.SubmissionConfirmed Then
        client = loginForm.UserName
        secret = loginForm.UserPassword
        urlToken = "http://localhost:8080/oauth2/token"
    End If
    
    Unload loginForm ' закрываем форму
End Function

