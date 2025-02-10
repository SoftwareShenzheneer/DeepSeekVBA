Function CallDeepSeekAPI(api_key As String, inputText As String) As String
    Dim API As String
    Dim SendTxt As String
    Dim Http As Object
    Dim status_code As Integer
    Dim response As String
    
    API = "https://api.deepseek.com/chat/completions"
    SendTxt = "{""model"": ""deepseek-reasoner"", ""messages"": [{""role"":""system"", ""content"":""You are a Word assistant""}, {""role"":""user"", ""content"":""" & inputText & """}], ""stream"": false}"
    
    Set Http = CreateObject("MSXML2.XMLHTTP")
    With Http
        .Open "POST", API, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & api_key
        .send SendTxt
        status_code = .Status
        response = .responseText
    End With
    
    ' 弹出窗口显示 API 响应（调试用）
    MsgBox "API Response: " & response, vbInformation, "Debug Info"
    
    If status_code = 200 Then
        CallDeepSeekAPI = response
    Else
        CallDeepSeekAPI = "Error: " & status_code & " - " & response
    End If
    
    Set Http = Nothing
End Function

Sub DeepSeekV3()
    Dim api_key As String
    Dim inputText As String
    Dim response As String
    Dim regex As Object
    Dim reasoningRegex As Object
    Dim contentRegex As Object
    Dim matches As Object
    Dim reasoningMatches As Object
    Dim originalSelection As Object
    Dim reasoningContent As String
    Dim finalContent As String
    
    api_key = "替换为你的api key"
    If api_key = "" Then
        MsgBox "Please enter the API key."
        Exit Sub
    ElseIf Selection.Type <> wdSelectionNormal Then
        MsgBox "Please select text."
        Exit Sub
    End If
    
    ' 保存原始选中的文本
    Set originalSelection = Selection.Range.Duplicate
    
    inputText = Replace(Replace(Replace(Replace(Replace(Selection.text, "\", "\\"), vbCrLf, ""), vbCr, ""), vbLf, ""), Chr(34), "\""")
    response = CallDeepSeekAPI(api_key, inputText)
    
    If Left(response, 5) <> "Error" Then
        ' 创建正则表达式对象来分别匹配推理内容和最终回答
        Set reasoningRegex = CreateObject("VBScript.RegExp")
        With reasoningRegex
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = """reasoning_content"":""(.*?)"""
        End With
        
        Set contentRegex = CreateObject("VBScript.RegExp")
        With contentRegex
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = """content"":""(.*?)"""
        End With
        
        ' 提取推理内容
        Set reasoningMatches = reasoningRegex.Execute(response)
        If reasoningMatches.Count > 0 Then
            reasoningContent = reasoningMatches(0).SubMatches(0)
            reasoningContent = Replace(reasoningContent, "\n\n", vbNewLine)
            reasoningContent = Replace(reasoningContent, "\n", vbNewLine)
            reasoningContent = Replace(Replace(reasoningContent, """", Chr(34)), """", Chr(34))
        End If
        
        ' 提取最终回答
        Set matches = contentRegex.Execute(response)
        If matches.Count > 0 Then
            finalContent = matches(0).SubMatches(0)
            finalContent = Replace(finalContent, "\n\n", vbNewLine)
            finalContent = Replace(finalContent, "\n", vbNewLine)
            finalContent = Replace(Replace(finalContent, """", Chr(34)), """", Chr(34))
            
            ' 插入推理过程（如果存在）
            If Len(reasoningContent) > 0 Then
                Selection.TypeParagraph
                Selection.TypeText "推理过程: "
                Selection.TypeParagraph
                Selection.TypeText reasoningContent
                Selection.TypeParagraph
                Selection.TypeText "最终回答: "
                Selection.TypeParagraph
            End If
            
            ' 插入最终回答
            Selection.TypeText finalContent
            
            ' 将光标移回原始文本的末尾
            originalSelection.Select
        Else
            MsgBox "Failed to parse API response.", vbExclamation
        End If
    Else
        MsgBox response, vbCritical
    End If
End Sub