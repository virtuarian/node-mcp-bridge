Attribute VB_Name = "OLD"
' Gemini API���g�p����ReAct�G�[�W�F���g
' �K�v�ȃ��t�@�����X:
' - Microsoft Scripting Runtime
' - Microsoft WinHTTP Services

Option Explicit


' �c�[����`
Private Type ToolParameter
    name As String
    description As String
    required As Boolean
    paramType As String ' string, number, boolean �Ȃ�
End Type

Private Type Tool
    name As String
    description As String
    parameters() As ToolParameter
End Type

' �O���[�o���ϐ�
Private mTools() As Tool
Private mToolCount As Integer
Private mConversationHistory As Collection

Private mToolsStr As String

' ���C���G�[�W�F���g���s�֐�
Public Sub RunGeminiAgent()
    ' ������
    InitializeAgent
    
    ' UI�̐ݒ�i�C�ӂ̃��[�N�V�[�g�Ɍ��ʕ\���j
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Agent")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' �V�[�g�����݂��Ȃ��ꍇ�͍쐬
        Set ws = ThisWorkbook.Sheets.Add
        ws.name = "Agent"
    End If
    
    'ws.Cells.Clear
    ws.Range("A1").value = "Gemini ReAct �G�[�W�F���g"
    ws.Range("A2").value = "����:"
    ws.Range("A3").value = "����:"
    
    ' ���[�U�[���͂��擾
    Dim userInput As String
    userInput = Range("prompt")
    
    If userInput = "" Then Exit Sub
    
    ' �G�[�W�F���g�Ɏw�������s������
    Dim result As String
    result = ExecuteAgentTask(userInput)
    
    ' ���ʂ�\��
    ws.Range("B3").value = result
    
    ' �������N���A
    Set mConversationHistory = New Collection
End Sub

' �G�[�W�F���g�̏�����
Private Sub InitializeAgent()
    ' ��b�����̏�����
    Set mConversationHistory = New Collection
    
    ' �c�[���̏�����
    InitializeTools
End Sub

' �c�[���̏�����
Private Sub InitializeTools()
    ' �ŏ���MCP�u���b�W����c�[����`���擾
    FetchMCPBridgeTools
    
    ' VBA�֐��Ăяo���c�[����ǉ�
    mToolCount = mToolCount + 1
    ReDim Preserve mTools(1 To mToolCount)
    
    ' VBA�֐��c�[���ݒ�
    Dim vbaToolIndex As Integer
    vbaToolIndex = mToolCount
    
    mTools(vbaToolIndex).name = "callVBAFunction"
    mTools(vbaToolIndex).description = "�o�^�ς݂�VBA�֐����Ăяo���܂�"
    ReDim mTools(vbaToolIndex).parameters(1 To 2)
    
    mTools(vbaToolIndex).parameters(1).name = "functionName"
    mTools(vbaToolIndex).parameters(1).description = "�Ăяo��VBA�֐���"
    mTools(vbaToolIndex).parameters(1).required = True
    mTools(vbaToolIndex).parameters(1).paramType = "string"
    
    mTools(vbaToolIndex).parameters(2).name = "params"
    mTools(vbaToolIndex).parameters(2).description = "�֐��ɓn���p�����[�^�iJSON�z��`���j"
    mTools(vbaToolIndex).parameters(2).required = False
    mTools(vbaToolIndex).parameters(2).paramType = "string"
End Sub

' MCP�u���b�W����c�[����`���擾����
Private Sub FetchMCPBridgeTools()
    On Error GoTo ErrorHandler
    
    ' HTTP GET���N�G�X�g
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' function-calling�G���h�|�C���g���Ăяo��
    Dim apiEndpoint As String
    apiEndpoint = MCP_BRIDGE_URL & "/tools/function-calling?provider=gemini"
    
    http.Open "GET", apiEndpoint, False
    http.Send
    
    ' ���X�|���X�̎擾
    Dim response As String
    response = http.responseText
    
    ' ���K�\�����g���Ē���JSON����K�v�ȏ��𒊏o
    mToolsStr = ExtractToolsArray(response)
    Exit Sub
    
ErrorHandler:
    ' �G���[�̏ꍇ�͋�̃c�[���z��ŏ�����
    Debug.Print "�c�[���擾�G���[: " & Err.description
    mToolCount = 0
    ReDim mTools(1 To 1)
End Sub

Private Function ExtractToolsArray(jsonStr As String) As String
    Dim startPos As Long
    Dim endPos As Long
    
    ' "tools": ������T��
    startPos = InStr(jsonStr, """tools"":")
    If startPos > 0 Then
        ' [ �̈ʒu��T��
        startPos = InStr(startPos, jsonStr, "[")
        If startPos > 0 Then
            ' �Ή���������ʂ�������
            Dim bracketCount As Integer
            bracketCount = 1
            endPos = startPos
            
            Do While bracketCount > 0 And endPos < Len(jsonStr)
                endPos = endPos + 1
                If Mid(jsonStr, endPos, 1) = "[" Then
                    bracketCount = bracketCount + 1
                ElseIf Mid(jsonStr, endPos, 1) = "]" Then
                    bracketCount = bracketCount - 1
                End If
            Loop
            
            ' ���S�Ȕz����擾
            ExtractToolsArray = Mid(jsonStr, startPos, endPos - startPos + 1)
            Exit Function
        End If
    End If
    
    ' ���s�����ꍇ�͋�̔z���Ԃ�
    ExtractToolsArray = "[]"
End Function

' JSON���X�|���X����c�[����`�𒊏o���Ē���mTools�ɐݒ�
Private Sub ExtractToolsFromJsonString(jsonStr As String)
    On Error GoTo ErrorHandler
    
    ' ���K�\���I�u�W�F�N�g���쐬
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' "type":"function","function":{"name":"TOOLNAME" �̃p�^�[��������
    regex.Pattern = """type""\s*:\s*""function""\s*,\s*""function""\s*:\s*\{\s*""name""\s*:\s*""([^""]+)"""
    regex.Global = True
    
    Dim matches As Object
    Set matches = regex.Execute(jsonStr)
    
    ' �c�[�����Ɋ�Â��Ĕz���������
    mToolCount = matches.Count
    If mToolCount > 0 Then
        ReDim mTools(1 To mToolCount)
    Else
        ReDim mTools(1 To 1)  ' �ŏ��T�C�Y
        mToolCount = 0
        Exit Sub
    End If
    
    Dim i As Integer
    For i = 1 To matches.Count
        ' �C���f�b�N�X�̒����iVBA��1�n�܂�j
        Dim index As Integer
        index = i
        
        Dim toolName As String
        toolName = matches(i - 1).SubMatches(0)
        
        ' �����𒊏o
        Dim descPattern As String
        descPattern = """name""\s*:\s*""" & Replace(toolName, "_", "\\_") & """\s*,\s*""description""\s*:\s*""([^""]+)"""
        
        Dim descRegex As Object
        Set descRegex = CreateObject("VBScript.RegExp")
        descRegex.Pattern = descPattern
        
        Dim descMatches As Object
        Set descMatches = descRegex.Execute(jsonStr)
        
        Dim description As String
        If descMatches.Count > 0 Then
            description = descMatches(0).SubMatches(0)
        Else
            description = "�c�[���̐���������܂���"
        End If
        
        ' �c�[������ݒ�
        mTools(index).name = toolName
        mTools(index).description = description
        
        ' �p�����[�^�̒��o
        ' �܂� properties �u���b�N��T��
        Dim propertiesBlockPattern As String
        propertiesBlockPattern = """properties""\s*:\s*\{([^\}]*)\}"
        
        Dim propertiesRegex As Object
        Set propertiesRegex = CreateObject("VBScript.RegExp")
        propertiesRegex.Pattern = propertiesBlockPattern
        
        Dim propertiesMatches As Object
        Set propertiesMatches = propertiesRegex.Execute(jsonStr)
        
        ' �p�����[�^�����J�E���g
        Dim paramCount As Integer
        paramCount = 1 ' �f�t�H���g��1��
        
        If propertiesMatches.Count > 0 Then
            Dim propertiesBlock As String
            propertiesBlock = propertiesMatches(0).SubMatches(0)
            
            ' �v���p�e�B���𒊏o�i��: "url"�j
            Dim paramNameRegex As Object
            Set paramNameRegex = CreateObject("VBScript.RegExp")
            paramNameRegex.Pattern = """([^""]+)""\s*:"
            paramNameRegex.Global = True
            
            Dim paramNameMatches As Object
            Set paramNameMatches = paramNameRegex.Execute(propertiesBlock)
            
            ' �p�����[�^�����X�V
            paramCount = paramNameMatches.Count
            If paramCount = 0 Then paramCount = 1
            
            ' �p�����[�^�z���������
            ReDim mTools(index).parameters(1 To paramCount)
            
            ' �e�p�����[�^�̏ڍׂ𒊏o
            Dim j As Integer
            For j = 1 To paramNameMatches.Count
                Dim paramName As String
                paramName = paramNameMatches(j - 1).SubMatches(0)
                
                ' �p�����[�^�^�C�v�𒊏o
                Dim typePattern As String
                typePattern = """" & paramName & """\s*:\s*\{[^}]*""type""\s*:\s*""([^""]+)"""
                
                Dim typeRegex As Object
                Set typeRegex = CreateObject("VBScript.RegExp")
                typeRegex.Pattern = typePattern
                
                Dim typeMatches As Object
                Set typeMatches = typeRegex.Execute(jsonStr)
                
                Dim paramType As String
                If typeMatches.Count > 0 Then
                    paramType = typeMatches(0).SubMatches(0)
                Else
                    paramType = "object" ' �f�t�H���g
                End If
                
                ' �p�����[�^�����𒊏o
                Dim descParamPattern As String
                descParamPattern = """" & paramName & """\s*:\s*\{[^}]*""description""\s*:\s*""([^""]+)"""
                
                Dim descParamRegex As Object
                Set descParamRegex = CreateObject("VBScript.RegExp")
                descParamRegex.Pattern = descParamPattern
                
                Dim descParamMatches As Object
                Set descParamMatches = descParamRegex.Execute(jsonStr)
                
                Dim paramDesc As String
                If descParamMatches.Count > 0 Then
                    paramDesc = descParamMatches(0).SubMatches(0)
                Else
                    paramDesc = paramName & "�p�����[�^"
                End If
                
                ' �p�����[�^����ݒ�
                mTools(index).parameters(j).name = paramName
                mTools(index).parameters(j).description = paramDesc
                mTools(index).parameters(j).paramType = paramType
                
                ' �K�{�p�����[�^�̊m�F
                Dim requiredPattern As String
                requiredPattern = """required""\s*:\s*\[(.*?)\]"
                
                Dim reqRegex As Object
                Set reqRegex = CreateObject("VBScript.RegExp")
                reqRegex.Pattern = requiredPattern
                
                Dim reqMatches As Object
                Set reqMatches = reqRegex.Execute(jsonStr)
                
                ' �f�t�H���g��false�ɐݒ�
                mTools(index).parameters(j).required = False
                
                ' �K�{�p�����[�^�����������ꍇ
                If reqMatches.Count > 0 Then
                    Dim requiredStr As String
                    requiredStr = reqMatches(0).SubMatches(0)
                    
                    ' �p�����[�^�����K�{���X�g�Ɋ܂܂�Ă��邩�m�F
                    If InStr(requiredStr, """" & paramName & """") > 0 Then
                        mTools(index).parameters(j).required = True
                    End If
                End If
            Next j
        Else
            ' �v���p�e�B��������Ȃ��ꍇ�̓f�t�H���g�p�����[�^��ݒ�
            ReDim mTools(index).parameters(1 To 1)
            mTools(index).parameters(1).name = "args"
            mTools(index).parameters(1).description = "�c�[���ɓn�������iJSON�`���j"
            mTools(index).parameters(1).paramType = "object"
            mTools(index).parameters(1).required = True
        End If
    Next i
    Exit Sub
    
ErrorHandler:
    Debug.Print "JSON��̓G���[: " & Err.description
    mToolCount = 0
    ReDim mTools(1 To 1)
End Sub

' �G�[�W�F���g�^�X�N���s
Private Function ExecuteAgentTask(prompt As String) As String
    Dim maxIterations As Integer
    maxIterations = 5 ' �ő唽����
    
    Dim currentIteration As Integer
    currentIteration = 0
    
    Dim finalResult As String
    finalResult = ""
    
    ' ���[�U�[���͂𗚗��ɒǉ�
    AddToHistory "user", prompt
    
    ' ReAct���[�v - �v�l�ƍs�����J��Ԃ�
    Do While currentIteration < maxIterations
        currentIteration = currentIteration + 1
        
        ' LLM�ɐ��_�Ǝ��̃A�N�V���������߂�i�c�[����`�t���j
        Dim llmResponse As String
        llmResponse = CallGeminiAPIWithTools(GetPromptWithHistory())
        
        ' Gemini API�̉����`�������
        Dim functionName As String
        Dim functionArgs As String
        Dim responseText As String
        Dim isFunctionCall As Boolean
        
        ParseGeminiResponse llmResponse, functionName, functionArgs, responseText, isFunctionCall
        
        ' ���X�|���X�𗚗��ɒǉ�
        AddToHistory "assistant", responseText
        
        ' �֐��Ăяo��������ꍇ�A�c�[�������s
        If isFunctionCall Then
            ' �c�[�������s
            Dim toolResult As String
            toolResult = ExecuteTool(functionName, functionArgs)
            
            ' �c�[�����s���ʂ𗚗��ɒǉ�
            AddToHistory "tool", toolResult
        Else
            ' �֐��Ăяo�����Ȃ��ꍇ�͍ŏI�񓚂Ƃ݂Ȃ�
            finalResult = responseText
            Exit Do
        End If
    Loop
    
    ' �ő唽���񐔂ɒB�����ꍇ
    If finalResult = "" Then
        finalResult = "�^�X�N�������ł��܂���ł����B�ڍ׏�񂪕K�v�ł��B"
    End If
    
    ExecuteAgentTask = finalResult
End Function

' Gemini API�̉�������́i���K�\���Łj
Private Sub ParseGeminiResponse(jsonResponse As String, ByRef functionName As String, ByRef functionArgs As String, ByRef responseText As String, ByRef isFunctionCall As Boolean)
    On Error GoTo ErrorHandler
    
    ' �f�t�H���g�l��ݒ�
    functionName = ""
    functionArgs = "{}"
    responseText = ""
    isFunctionCall = False
    
    ' functionCall�����邩�ǂ����𐳋K�\���Ŋm�F
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' functionCall������
    regex.Pattern = """functionCall""\s*:\s*\{[^\}]+\}"
    regex.Global = True
    
    Dim matches As Object
    Set matches = regex.Execute(jsonResponse)
    
    If matches.Count > 0 Then
        ' functionCall����������
        isFunctionCall = True
        
        ' �ŏ���functionCall�������i���������Ă�1�������j
        ' �c�[�����𒊏o
        Dim nameRegex As Object
        Set nameRegex = CreateObject("VBScript.RegExp")
        nameRegex.Pattern = """name""\s*:\s*""([^""]+)"""
        
        Dim nameMatches As Object
        Set nameMatches = nameRegex.Execute(matches(0))
        
        If nameMatches.Count > 0 Then
            functionName = nameMatches(0).SubMatches(0)
        Else
            functionName = "unknown_tool"
        End If
        
        ' �����𒊏o
        Dim argsRegex As Object
        Set argsRegex = CreateObject("VBScript.RegExp")
        argsRegex.Pattern = """args""\s*:\s*(\{[^\}]*\})"
        
        Dim argsMatches As Object
        Set argsMatches = argsRegex.Execute(matches(0))
        
        If argsMatches.Count > 0 Then
            functionArgs = argsMatches(0).SubMatches(0)
        Else
            functionArgs = "{}"
        End If
    End If
    
    ' �e�L�X�g�����𒊏o
    Dim textRegex As Object
    Set textRegex = CreateObject("VBScript.RegExp")
    textRegex.Pattern = """text""\s*:\s*""([^""]+)"""
    
    Dim textMatches As Object
    Set textMatches = textRegex.Execute(jsonResponse)
    
    If textMatches.Count > 0 Then
        responseText = Replace(textMatches(0).SubMatches(0), "\\", "\")
        responseText = Replace(responseText, "\""", """")
    Else
        ' �e�L�X�g��������Ȃ��ꍇ
        responseText = "�����e�L�X�g������܂���"
    End If
    
    Exit Sub
    
ErrorHandler:
    responseText = "�G���[: Gemini API�̃��X�|���X��͂Ɏ��s���܂����B" & Err.description
    isFunctionCall = False
End Sub

' �c�[�����s
Private Function ExecuteTool(toolName As String, argsJsonStr As String) As String
    On Error GoTo ErrorHandler
    
    ' VBA�֐��Ăяo���̏ꍇ
    If toolName = "callVBAFunction" Then
        ExecuteTool = CallVBAFunction(argsJsonStr)
        Exit Function
    End If
    
    ' MCP�c�[���Ăяo���̏ꍇ
    ' MCP�u���b�W�̃��N�G�X�g�{�����쐬
    Dim requestJson As String
    requestJson = "{""name"":""" & toolName & """,""arguments"":" & argsJsonStr & "}"
    
    ' HTTP POST���N�G�X�g
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' MCP�u���b�W��function-call�G���h�|�C���g
    Dim apiEndpoint As String
    apiEndpoint = MCP_BRIDGE_URL & "/tools/function-call"
    
    http.Open "POST", apiEndpoint, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send requestJson
    
    ' ���X�|���X��Ԃ�
    ExecuteTool = http.responseText
    Exit Function
    
ErrorHandler:
    ExecuteTool = "�c�[�����s�G���[: " & Err.description
End Function

' VBA�֐��Ăяo��
Private Function CallVBAFunction(inputJson As String) As String
    On Error GoTo ErrorHandler
    
    ' JSON����́i�P���Ȑ��K�\���p�[�X�ő�p�j
    Dim functionName As String
    Dim paramsArray As String
    
    ' �֐����𒊏o
    Dim nameRegex As Object
    Set nameRegex = CreateObject("VBScript.RegExp")
    nameRegex.Pattern = """functionName""\s*:\s*""([^""]+)"""
    
    Dim nameMatches As Object
    Set nameMatches = nameRegex.Execute(inputJson)
    
    If nameMatches.Count > 0 Then
        functionName = nameMatches(0).SubMatches(0)
    Else
        CallVBAFunction = "�G���[: �֐�����������܂���"
        Exit Function
    End If
    
    ' �p�����[�^�𒊏o
    Dim paramsRegex As Object
    Set paramsRegex = CreateObject("VBScript.RegExp")
    paramsRegex.Pattern = """params""\s*:\s*(\[[^\]]*\])"
    
    Dim paramsMatches As Object
    Set paramsMatches = paramsRegex.Execute(inputJson)
    
    If paramsMatches.Count > 0 Then
        paramsArray = paramsMatches(0).SubMatches(0)
    Else
        paramsArray = "[]"
    End If
    
    ' �֐����ɉ���������
    Select Case functionName
        Case "GetCellValue"
            ' �p�����[�^��: ["Sheet1", "A1"]
            ' �P���Ȑ��K�\���Ŕz��v�f�𒊏o
            Dim elemRegex As Object
            Set elemRegex = CreateObject("VBScript.RegExp")
            elemRegex.Pattern = """([^""]+)"""
            elemRegex.Global = True
            
            Dim elemMatches As Object
            Set elemMatches = elemRegex.Execute(paramsArray)
            
            If elemMatches.Count >= 2 Then
                Dim sheet As String, cell As String
                sheet = elemMatches(0).SubMatches(0)
                cell = elemMatches(1).SubMatches(0)
                
                On Error Resume Next
                Dim cellValue As Variant
                cellValue = ThisWorkbook.Sheets(sheet).Range(cell).value
                
                If Err.Number = 0 Then
                    CallVBAFunction = CStr(cellValue)
                Else
                    CallVBAFunction = "�G���[: �Z���l�̎擾�Ɏ��s���܂��� - " & Err.description
                End If
                On Error GoTo ErrorHandler
            Else
                CallVBAFunction = "�G���[: �p�����[�^�s��"
            End If
            
        Case "CalculateSum"
            ' �p�����[�^��: ["Sheet1", "A1:A10"]
            Dim elemRegex2 As Object
            Set elemRegex2 = CreateObject("VBScript.RegExp")
            elemRegex2.Pattern = """([^""]+)"""
            elemRegex2.Global = True
            
            Dim elemMatches2 As Object
            Set elemMatches2 = elemRegex2.Execute(paramsArray)
            
            If elemMatches2.Count >= 2 Then
                Dim sumSheet As String, sumRange As String
                sumSheet = elemMatches2(0).SubMatches(0)
                sumRange = elemMatches2(1).SubMatches(0)
                
                On Error Resume Next
                Dim sumResult As Double
                sumResult = Application.Sum(ThisWorkbook.Sheets(sumSheet).Range(sumRange))
                
                If Err.Number = 0 Then
                    CallVBAFunction = CStr(sumResult)
                Else
                    CallVBAFunction = "�G���[: ���v�̌v�Z�Ɏ��s���܂��� - " & Err.description
                End If
                On Error GoTo ErrorHandler
            Else
                CallVBAFunction = "�G���[: �p�����[�^�s��"
            End If
            
        Case "ShowMessage"
            ' �p�����[�^��: ["����ɂ���"]
            Dim msgRegex As Object
            Set msgRegex = CreateObject("VBScript.RegExp")
            msgRegex.Pattern = """([^""]+)"""
            
            Dim msgMatches As Object
            Set msgMatches = msgRegex.Execute(paramsArray)
            
            If msgMatches.Count >= 1 Then
                Dim msg As String
                msg = msgMatches(0).SubMatches(0)
                MsgBox msg, vbInformation, "���b�Z�[�W"
                CallVBAFunction = "���b�Z�[�W��\�����܂���: " & msg
            Else
                CallVBAFunction = "�G���[: ���b�Z�[�W���w�肳��Ă��܂���"
            End If
            
        Case Else
            CallVBAFunction = "�G���[: ���m�̊֐� '" & functionName & "'"
    End Select
    
    Exit Function
    
ErrorHandler:
    CallVBAFunction = "�G���[: " & Err.description
End Function

' �c�[����`���܂߂�Gemini API���Ăяo��
Private Function CallGeminiAPIWithTools(prompt As String) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' API�G���h�|�C���g�Ƀ��N�G�X�g
    http.Open "POST", GEMINI_API_URL & "?key=" & GEMINI_API_KEY, False
    http.SetRequestHeader "Content-Type", "application/json"
    
    ' ���N�G�X�g�{�f�B�̍\�z�i�c�[����`���܂ށj
    Dim requestBody As String
    requestBody = BuildGeminiRequestWithTools(prompt)
    
    ' ���N�G�X�g���M
    http.Send requestBody
    
    ' ���X�|���X�擾
    Dim response As String
    response = http.responseText
    
    CallGeminiAPIWithTools = response
    Exit Function
    
ErrorHandler:
    CallGeminiAPIWithTools = "{""error"": """ & Err.description & """}"
End Function

' �c�[����`���܂�Gemini���N�G�X�g�̍\�z
Private Function BuildGeminiRequestWithTools(prompt As String) As String

On Error GoTo Err
    ' JSON������𒼐ڍ\�z
    Dim jsonStr As String
    
    jsonStr = "{"
    jsonStr = jsonStr & """contents"": ["
    jsonStr = jsonStr & "  {"
    jsonStr = jsonStr & "    ""parts"": ["
    jsonStr = jsonStr & "      {"
    jsonStr = jsonStr & "        ""text"": """ & EscapeJsonString(prompt) & """"
    jsonStr = jsonStr & "      }"
    jsonStr = jsonStr & "    ]"
    jsonStr = jsonStr & "  }"
    jsonStr = jsonStr & "],"
    
    ' �c�[����`��ǉ�
   jsonStr = jsonStr & """tools"": ["
    jsonStr = jsonStr & mToolsStr
    jsonStr = jsonStr & "],"
    
    ' �����ݒ��ǉ�
    jsonStr = jsonStr & """generationConfig"": {"
    jsonStr = jsonStr & "  ""temperature"": 0.2,"
    jsonStr = jsonStr & "  ""topP"": 0.95,"
    jsonStr = jsonStr & "  ""topK"": 40"
    jsonStr = jsonStr & "}"
    
    jsonStr = jsonStr & "}"
    
    BuildGeminiRequestWithTools = jsonStr
    
    Exit Function
    
Err:
    Debug.Print Err.description
    
End Function

' �c�[����`��JSON�\�z
Private Function BuildToolsDefinitionJson() As String
    Dim result As String
    result = ""
    
    Dim i As Integer, j As Integer
    
    For i = 1 To mToolCount
        If i > 1 Then result = result & ","
        
        result = result & "{"
        result = result & """function_declarations"": ["
        result = result & "  {"
        result = result & "    ""name"": """ & mTools(i).name & ""","
        result = result & "    ""description"": """ & EscapeJsonString(mTools(i).description) & ""","
        result = result & "    ""parameters"": {"
        result = result & "      ""type"": ""object"","
        result = result & "      ""properties"": {"
        
        For j = 1 To UBound(mTools(i).parameters)
            If j > 1 Then result = result & ","
            
            result = result & "        """ & mTools(i).parameters(j).name & """: {"
            result = result & "          ""type"": """ & mTools(i).parameters(j).paramType & ""","
            result = result & "          ""description"": """ & EscapeJsonString(mTools(i).parameters(j).description) & """"
            result = result & "        }"
        Next j
        
        result = result & "      },"
        
        ' �K�{�p�����[�^�̐ݒ�
        result = result & "      ""required"": ["
        
        Dim firstRequired As Boolean
        firstRequired = True
        
        For j = 1 To UBound(mTools(i).parameters)
            If mTools(i).parameters(j).required Then
                If Not firstRequired Then result = result & ","
                result = result & "        """ & mTools(i).parameters(j).name & """"
                firstRequired = False
            End If
        Next j
        
        result = result & "      ]"
        result = result & "    }"
        result = result & "  }"
        result = result & "]"
        result = result & "}"
    Next i
    
    BuildToolsDefinitionJson = result
End Function

' ��b�����ɒǉ�
Private Sub AddToHistory(role As String, content As String)
    Dim historyItem As Object
    Set historyItem = CreateObject("Scripting.Dictionary")
    
    historyItem.Add "role", role
    historyItem.Add "content", content
    
    mConversationHistory.Add historyItem
End Sub

' ��b�������܂ރv�����v�g�𐶐�
Private Function GetPromptWithHistory() As String
    Dim prompt As String
    prompt = "���Ȃ��͗l�X�ȃc�[�����g����AI�A�V�X�^���g�ł��B���[�U�[�̎w���Ɋ�Â��āA�ȉ��̃c�[�����g�p���ă^�X�N�����s���܂��B" & vbCrLf & vbCrLf
    
    ' �c�[���̐�����ǉ�
    prompt = prompt & "���p�\�ȃc�[��:" & vbCrLf
    
    Dim i As Integer, j As Integer
    For i = 1 To mToolCount
        prompt = prompt & "- " & mTools(i).name & ": " & mTools(i).description & vbCrLf
        prompt = prompt & "  �p�����[�^:" & vbCrLf
        
        For j = 1 To UBound(mTools(i).parameters)
            prompt = prompt & "  - " & mTools(i).parameters(j).name & ": " & mTools(i).parameters(j).description
            If mTools(i).parameters(j).required Then
                prompt = prompt & " (�K�{)"
            End If
            prompt = prompt & vbCrLf
        Next j
        
        prompt = prompt & vbCrLf
    Next i
    
    prompt = prompt & "���[�U�[�̎w���ɏ]���A�K�v�ȃc�[�����g�p���ĖړI��B�����Ă��������B" & vbCrLf & vbCrLf
    
    ' ��b������ǉ�
    prompt = prompt & "����܂ł̉�b:" & vbCrLf
    
    Dim item As Variant
    For Each item In mConversationHistory
        Dim role As String, content As String
        role = item("role")
        content = item("content")
        
        Select Case role
            Case "user"
                prompt = prompt & "���[�U�[: " & content & vbCrLf
            Case "assistant"
                prompt = prompt & "�A�V�X�^���g: " & content & vbCrLf
            Case "tool"
                prompt = prompt & "�c�[������: " & content & vbCrLf
        End Select
    Next item
    
    prompt = prompt & vbCrLf & "���̉����𐶐����Ă��������B"
    
    GetPromptWithHistory = prompt
End Function

' JSON������̃G�X�P�[�v
Private Function EscapeJsonString(text As String) As String
    On Error Resume Next
    
    Dim result As String
    result = text
    
    ' ���ꕶ���̃G�X�P�[�v
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\r")
    result = Replace(result, vbTab, "\t")
    
    EscapeJsonString = result
End Function

