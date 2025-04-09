' Gemini APIを使用したReActエージェント
' 必要なリファレンス:
' - Microsoft Scripting Runtime
' - Microsoft WinHTTP Services
'

Option Explicit

' グローバル変数
Private mTools() As Tool
Private mToolCount As Integer
Private mConversationHistory As Collection
Private mToolsStr As String
Private mRealTimeDisplay As Boolean
Private mCurrentDisplayRow As Integer
Private mDisplayWorksheet As Worksheet

' AIアシスタントの見た目と出力を改善
Public Sub RunGeminiReActAgent()
    ' 初期化
    InitializeAgent
    InitializeTools
    ' UIの設定
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ReActAgent")
    On Error GoTo 0
    
    ' ユーザー入力を取得
    Dim userInput As String
    userInput = Range("prompt").value
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.name = "ReActAgent"
    End If
    
    ws.Cells.Clear
    ws.Range("A1").value = "Gemini ReAct AI エージェント"
    ws.Range("A2").value = "指示:"
    ws.Range("A3").value = userInput
    ws.Range("A5").value = "実行プロセス:"
    
    ws.Range("A:A").ColumnWidth = 100
    ws.Range("A:A").WrapText = True
    
    If userInput = "" Then
        ws.Range("A6").value = "指示が入力されていません。「prompt」セルに指示を入力してください。"
        Exit Sub
    End If
    
    ' 処理中表示を行うセルを用意
    ws.Range("A6").value = "思考中..."
    
    ' リアルタイム表示用の行番号を初期化
    Dim currentRow As Integer
    currentRow = 6
    
    ' リアルタイム処理フラグをセット
    mRealTimeDisplay = True
    mCurrentDisplayRow = currentRow
    Set mDisplayWorksheet = ws
    
    ' エージェントを実行
    Dim result As String
    Dim startTime As Double
    startTime = Timer
    
    result = ExecuteAgentTask(userInput)
    
    ' 処理時間を計算
    Dim processingTime As Double
    processingTime = Timer - startTime
    
    ' 処理時間表示
    ws.Cells(mCurrentDisplayRow, 1).value = "処理時間: " & Format(processingTime, "0.00") & " 秒"
    
    ' 履歴をクリア
    Set mConversationHistory = New Collection
End Sub

' エージェントの初期化
Private Sub InitializeAgent()
    Debug.Print "エージェントを初期化しています..."
    
    ' 会話履歴の初期化
    Set mConversationHistory = New Collection
    
    ' ツールの初期化
    InitializeTools
End Sub

' ツールの初期化
Private Sub InitializeTools()
    Debug.Print "ツールを初期化しています..."
    
    ' MCPブリッジからツール定義を取得
    FetchMCPBridgeTools
    
    ' VBA関数呼び出しツールを追加
    mToolCount = mToolCount + 1
    ReDim Preserve mTools(1 To mToolCount)

    ' VBA関数ツール設定
    Dim vbaToolIndex As Integer
    vbaToolIndex = mToolCount

    mTools(vbaToolIndex).name = "excel_function"
    mTools(vbaToolIndex).description = "登録済みのVBA関数を呼び出します"
    ReDim mTools(vbaToolIndex).parameters(1 To 2)

    mTools(vbaToolIndex).parameters(1).name = "functionName"
    mTools(vbaToolIndex).parameters(1).description = "呼び出すVBA関数名"
    mTools(vbaToolIndex).parameters(1).required = True
    mTools(vbaToolIndex).parameters(1).paramType = "string"

    mTools(vbaToolIndex).parameters(2).name = "params"
    mTools(vbaToolIndex).parameters(2).description = "関数に渡すパラメータ（JSON配列形式）"
    mTools(vbaToolIndex).parameters(2).required = False
    mTools(vbaToolIndex).parameters(2).paramType = "string"
    
    Debug.Print "ツール初期化完了。合計: " & mToolCount & " ツール"
End Sub

' MCPブリッジからツール定義を取得する
Private Sub FetchMCPBridgeTools()
    On Error GoTo ErrorHandler
    
    Debug.Print "MCPブリッジからツール情報を取得しています..."
    
    ' HTTP GETリクエスト
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' function-callingエンドポイントを呼び出し
    Dim apiEndpoint As String
    apiEndpoint = MCP_BRIDGE_URL & "/tools/function-calling?provider=gemini"
    
    http.Open "GET", apiEndpoint, False
    http.Send
    
    ' レスポンスの取得
    Dim response As String
    response = http.responseText
    
    ' 正規表現を使って直接JSONから必要な情報を抽出
    mToolsStr = ExtractToolsArray(response)
    ExtractToolsFromJsonString response
    
    Debug.Print "ツール情報の取得に成功しました。ツール数: " & mToolCount
    Exit Sub
    
ErrorHandler:
    ' エラーの場合は空のツール配列で初期化
    Debug.Print "ツール取得エラー: " & Err.description
    mToolCount = 0
    ReDim mTools(1 To 1)
    mToolsStr = "[]"
End Sub

' AIエージェントの実行部分
Private Function ExecuteAgentTask(prompt As String) As String
    
    Dim currentIteration As Integer
    currentIteration = 0
    
    Dim finalResult As String
    finalResult = ""
    
    Debug.Print "ReActエージェント実行開始: " & prompt
    
    ' ユーザー入力を履歴に追加
    AddToHistory "user", prompt
    
    ' 最初の思考を生成
    Dim initialThought As String
    initialThought = "ユーザーの要求を分析し、対応方法を考えています。"
    
    ' ReActループ - 思考と行動を繰り返す
    Do While currentIteration < MAX_ITERATIONS
        currentIteration = currentIteration + 1
        Debug.Print "--- ReAct実行サイクル #" & currentIteration & " ---"
        
        ' STEP 1: REASONING/THOUGHT - AIモデルに思考を求める
        Dim thoughtPrompt As String
        thoughtPrompt = GetPromptWithReActFormat(prompt)
        
        ' Gemini APIを呼び出して思考または行動を取得
        Dim apiResponse As String
        apiResponse = CallGeminiAPIWithTools(thoughtPrompt)
        
        ' レスポンスを解析
        Dim responseType As String
        Dim responseContent As String
        ParseReActResponse apiResponse, responseType, responseContent
        
        ' レスポンスタイプに基づいて処理
        Select Case responseType
        Case "Thought"
            ' 思考を履歴に追加（ParseReActResponse内で既に処理済み）
            Debug.Print "思考: " & responseContent
            
        Case "Action"
            ' ツール名と引数を抽出
            Dim toolName As String
            Dim toolArgs As String
            ExtractToolCall responseContent, toolName, toolArgs
            
            ' ツール呼び出しを履歴に追加
            Debug.Print "ツール実行: " & toolName & " " & toolArgs
            AddToHistory "action", "Action: " & responseContent
            
            ' リアルタイム表示
            DisplayRealtimeOutput "action", "Action: " & toolName & " " & toolArgs
            
            ' STEP 2: ACTING - ツールを実行
            Dim toolResult As String
            toolResult = ExecuteTool(toolName, toolArgs)
            
            ' STEP 3: OBSERVATION - 結果を記録
            Debug.Print "観察: " & left(toolResult, 100) & IIf(Len(toolResult) > 100, "...", "")
            AddToHistory "observation", "Observation: " & toolResult
            
            ' リアルタイム表示
            DisplayRealtimeOutput "observation", "Observation: " & ShortenContent(toolResult, 100)
            
        Case "Answer"
            ' 最終回答
            Debug.Print "最終回答: " & responseContent
            finalResult = responseContent
            
            ' リアルタイム表示
            DisplayRealtimeOutput "answer", "Answer: " & responseContent
            Exit Do
                
            Case Else
                ' 予期しない応答形式の場合
                Debug.Print "予期しない応答: " & responseType & ": " & responseContent
                AddToHistory "error", "エラー: 応答形式が不正です: " & responseType
                
                ' 5秒間待機処理を追加
                Dim waitTime As Date
                waitTime = Now + TimeSerial(0, 0, 5)
                Debug.Print "予期しない応答のため5秒間待機します..." & Format(Now, "hh:mm:ss")
                Application.Wait waitTime
                Debug.Print "待機終了: " & Format(Now, "hh:mm:ss")
                
                ' リアルタイム表示
                DisplayRealtimeOutput "error", "予期しない応答のため5秒間待機しました: " & responseType
                    
        End Select
    Loop
    
    ' 最大反復回数に達した場合
    If finalResult = "" Then
        Debug.Print "最大反復回数に達しました"
        finalResult = "タスクを完了できませんでした。複雑すぎるか、追加情報が必要です。"
    End If
    
    ' 最終結果をフォーマット
    ExecuteAgentTask = finalResult
End Function

' ReActフォーマットの応答を解析（Thoughtとアクションを両方処理）
Private Sub ParseReActResponse(apiResponse As String, ByRef responseType As String, ByRef responseContent As String)
    On Error GoTo ErrorHandler
    
    ' デフォルト値
    responseType = ""
    responseContent = ""
    
    ' Gemini APIの応答を解析
    Dim jsonObject As Object
    Set jsonObject = JsonConverter.ParseJson(apiResponse)
    
    Dim fullText As String
    fullText = ""
    
    ' 新しいGemini APIのレスポンス構造を解析
    On Error Resume Next
    If jsonObject.Exists("candidates") Then
        If jsonObject("candidates").Count > 0 Then
            If TypeOf jsonObject("candidates")(1) Is Dictionary Then
                Dim candidate As Dictionary
                Set candidate = jsonObject("candidates")(1)
                
                If candidate.Exists("content") Then
                    Dim content As Dictionary
                    Set content = candidate("content")
                    
                    If content.Exists("parts") Then
                        If content("parts").Count > 0 Then
                            If TypeOf content("parts")(1) Is Dictionary Then
                                If content("parts")(1).Exists("text") Then
                                    fullText = CStr(content("parts")(1)("text"))
                                    Debug.Print "抽出されたテキスト: " & fullText
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    On Error GoTo ErrorHandler
    
    ' エラーレスポンスをチェック
    If fullText = "" Then
        If jsonObject.Exists("error") Then
            fullText = "APIエラー: "
            If TypeOf jsonObject("error") Is Dictionary Then
                If jsonObject("error").Exists("message") Then
                    fullText = fullText & jsonObject("error")("message")
                End If
            End If
        Else
            fullText = "応答テキストがありません"
        End If
        responseType = "error"
        responseContent = fullText
        Exit Sub
    End If
    
    ' より包括的なエスケープ処理
    fullText = Replace(fullText, "\n", vbCrLf)
    fullText = Replace(fullText, "\r", vbCr)
    fullText = Replace(fullText, "\t", vbTab)
    
    ' ReActフォーマットのパターンを検索
    Dim reactRegex As Object
    Set reactRegex = CreateObject("VBScript.RegExp")
    reactRegex.Pattern = "(Thought|Action|Observation|Answer)\s*:\s*([\s\S]+?)(?=Thought:|Action:|Observation:|Answer:|$)"
    reactRegex.Global = True
    reactRegex.IgnoreCase = True
    
    Dim reactMatches As Object
    Set reactMatches = reactRegex.Execute(fullText)
    
    ' マッチしたReActパターンを確認
    If reactMatches.Count > 0 Then
        ' Thought情報の抽出（常にデバッグ出力してリアルタイム表示）
        Dim i As Integer
        Dim thoughtText As String
        thoughtText = ""
        
        ' まずThoughtを探して記録（デバッグ用）
        For i = 0 To reactMatches.Count - 1
            If UCase(Trim(reactMatches(i).SubMatches(0))) = "THOUGHT" Then
                thoughtText = Trim(reactMatches(i).SubMatches(1))
                Debug.Print "思考内容: " & thoughtText
                
                ' リアルタイム表示を行う
                DisplayRealtimeOutput "thought", thoughtText
                
                ' 思考内容を履歴に追加
                On Error Resume Next
                AddToHistory "thought", "Thought: " & thoughtText
                On Error GoTo ErrorHandler
                
                Exit For
            End If
        Next i
        
        ' 優先度順に処理するアクションを探す
        Dim actionPriorities As Variant
        actionPriorities = Array("ACTION", "ANSWER", "OBSERVATION")
        Dim priority As Variant
        Dim foundAction As Boolean
        foundAction = False
        
        ' 優先度順にアクションを探す
        For Each priority In actionPriorities
            For i = 0 To reactMatches.Count - 1
                If UCase(Trim(reactMatches(i).SubMatches(0))) = priority Then
                    responseType = reactMatches(i).SubMatches(0)
                    responseContent = Trim(reactMatches(i).SubMatches(1))
                    foundAction = True
                    Exit For
                End If
            Next i
            
            If foundAction Then Exit For
        Next priority
        
        ' アクションが見つからなければ、Thoughtを最終結果として使用
        If Not foundAction Then
            ' Thoughtが見つかっていれば使用
            If thoughtText <> "" Then
                responseType = "Thought"
                responseContent = thoughtText
            Else
                ' Thoughtも見つからなければ最初のパターンを使用
                responseType = Trim(reactMatches(0).SubMatches(0))
                responseContent = Trim(reactMatches(0).SubMatches(1))
            End If
        End If
        
        ' 応答内容から不要な改行や重複を除去
        responseContent = CleanupResponseContent(responseContent)
    Else
        ' ReActパターンが見つからない場合は、思考として扱う
        responseType = "Thought"
        responseContent = CleanupResponseContent(fullText)
    End If
    
    Exit Sub
    
ErrorHandler:
    responseType = "error"
    responseContent = "エラー: レスポンスの解析に失敗しました。" & Err.description
End Sub

' 応答内容の整形（重複と余分な改行を削除）
Private Function CleanupResponseContent(content As String) As String
    ' 連続する改行を1つにまとめる
    Dim cleanText As String
    cleanText = content
    
    ' 繰り返しの置換（改行の連続を1つにする）
    Do While InStr(cleanText, vbCrLf & vbCrLf) > 0
        cleanText = Replace(cleanText, vbCrLf & vbCrLf, vbCrLf)
    Loop
    
    ' 行末の空白を削除
    Dim lines() As String
    lines = Split(cleanText, vbCrLf)
    
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        lines(i) = Trim(lines(i))
    Next i
    
    cleanText = Join(lines, vbCrLf)
    
    ' 重複する内容の検出と削除
    Dim fullLength As Long
    fullLength = Len(cleanText)
    
    ' 50文字以上の場合のみ重複チェックを行う（短いテキストでは誤検出の可能性）
    If fullLength > 50 Then
        Dim halfLength As Long
        halfLength = fullLength \ 2
        
        ' テキストの後半が前半の繰り返しになっていないか確認
        If Right(cleanText, halfLength) = left(cleanText, halfLength) Then
            cleanText = left(cleanText, fullLength - halfLength)
        End If
    End If
    
    CleanupResponseContent = Trim(cleanText)
End Function

' ツール呼び出し情報を抽出
Private Sub ExtractToolCall(actionText As String, ByRef toolName As String, ByRef toolArgs As String)
    On Error GoTo ErrorHandler
    
    ' デフォルト値
    toolName = ""
    toolArgs = "{}"
    
    ' パターン: ツール名 { "param": "value" }
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "([a-zA-Z0-9_]+)\s*(\{[\s\S]*\})"
    
    Dim matches As Object
    Set matches = regex.Execute(actionText)
    
    If matches.Count > 0 Then
        toolName = Trim(matches(0).SubMatches(0))
        toolArgs = Trim(matches(0).SubMatches(1))
    Else
        ' 代替パターン: "name": "ツール名", "arguments": { ... }
        Dim altRegex As Object
        Set altRegex = CreateObject("VBScript.RegExp")
        altRegex.Pattern = """name""[\s:]+""([^""]+)""|'name'[\s:]+""([^""]+)"""
        
        Dim nameMatches As Object
        Set nameMatches = altRegex.Execute(actionText)
        
        If nameMatches.Count > 0 Then
            If nameMatches(0).SubMatches(0) <> "" Then
                toolName = nameMatches(0).SubMatches(0)
            Else
                toolName = nameMatches(0).SubMatches(1)
            End If
            
            ' 引数を抽出
            Dim argsRegex As Object
            Set argsRegex = CreateObject("VBScript.RegExp")
            argsRegex.Pattern = """arguments""[\s:]+(\{[\s\S]*\})"
            
            Dim argsMatches As Object
            Set argsMatches = argsRegex.Execute(actionText)
            
            If argsMatches.Count > 0 Then
                toolArgs = argsMatches(0).SubMatches(0)
            End If
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    toolName = "error"
    toolArgs = "{""error"": """ & Err.description & """}"
End Sub

' ReActフォーマットのプロンプト生成
Private Function GetPromptWithReActFormat(userMessage As String) As String
    Dim prompt As String
    
    prompt = "あなたは問題解決ができるAIアシスタントです。" & vbCrLf & vbCrLf
    
    prompt = prompt & "基本ルール:" & vbCrLf
    prompt = prompt & "- 挨拶や一般的な質問には直接回答し、ツールは使わないでください" & vbCrLf
    prompt = prompt & "- 複雑な問題解決が必要な場合のみツールを使ってください" & vbCrLf
    prompt = prompt & "- あなたの思考は必ず日本語で行ってください" & vbCrLf & vbCrLf
    
    prompt = prompt & "ユーザの要求: " & userMessage & vbCrLf & vbCrLf
    
    prompt = prompt & "あなたの行動パターン:" & vbCrLf
    prompt = prompt & "1. Thought: ユーザの要求を理解し、どうすべきか考えます。また今までの行動結果をもとに次の行動を考えます。" & vbCrLf
    prompt = prompt & "2. Action: 必要に応じてツールを使います。" & vbCrLf
    prompt = prompt & "    存在するツール名とパラメータのみをJSON形式で記述してください。" & vbCrLf
    prompt = prompt & "    例： Action: ツール名 { ""param1"": ""value1"", ""param2"": ""value2"" }" & vbCrLf
    prompt = prompt & "    注意：" & vbCrLf
    prompt = prompt & "    　・必ず存在するツール名を指定すること" & vbCrLf
    prompt = prompt & "    　・１度に１つのツールを指定すること" & vbCrLf
    prompt = prompt & "3. Observation: ツールの結果を確認します。" & vbCrLf
    prompt = prompt & "4. Thought→Action→Objservation→Action...を順番に１回ずつ繰り返し、ユーザの要求に対して最終結果が出たら「Answer:」で回答します" & vbCrLf & vbCrLf
    
    prompt = prompt & "応答形式:" & vbCrLf
    prompt = prompt & "※注意事項：必ず「Thought:」+「Action:」「Observation:」「Answer:」のいずれか１つを出力してください。" & vbCrLf
    prompt = prompt & "- ""Thought: [あなたの思考]""" & vbCrLf
    prompt = prompt & "- ""Action:  [使用するツール名 { ""パラメータ"": ""値"", ... }]""" & vbCrLf
    prompt = prompt & "- ""Observation: [ツールの結果]""" & vbCrLf
    prompt = prompt & "- ""Answer: [最終回答]""" & vbCrLf & vbCrLf
    
    ' 利用可能なMCPツールを追加
    prompt = prompt & "利用可能なMCPツール:" & vbCrLf
    
    Dim i As Integer, j As Integer
    For i = 1 To mToolCount
        prompt = prompt & "- " & mTools(i).name & ": " & mTools(i).description & vbCrLf
        prompt = prompt & "  パラメータ:" & vbCrLf
        
        For j = 1 To UBound(mTools(i).parameters)
            prompt = prompt & "  - " & mTools(i).parameters(j).name & ": " & mTools(i).parameters(j).description
            If mTools(i).parameters(j).required Then
                prompt = prompt & " (必須)"
            End If
            prompt = prompt & vbCrLf
        Next j
        
        prompt = prompt & vbCrLf
    Next i

    ' 利用可能なExcelツールを追加
    prompt = prompt & GetProcedureInfo & vbCrLf
    
    ' 会話履歴を追加
    prompt = prompt & "これまでの会話:" & vbCrLf
    
    Dim item As Variant
    For Each item In mConversationHistory
        Dim role As String, content As String
        role = item("role")
        content = item("content")
        
        Select Case role
            Case "user"
                prompt = prompt & "ユーザー: " & content & vbCrLf
            Case "assistant"
                prompt = prompt & "アシスタント: " & content & vbCrLf
            Case "thought"
                prompt = prompt & content & vbCrLf
            Case "action"
                prompt = prompt & content & vbCrLf
            Case "observation"
                prompt = prompt & content & vbCrLf
            Case "tool"
                prompt = prompt & "Observation: " & content & vbCrLf
            Case "error"
                prompt = prompt & content & vbCrLf
        End Select
    Next item
    
    prompt = prompt & vbCrLf & "次のステップを選択してください。"
    
    GetPromptWithReActFormat = prompt
End Function

' ツール実行
Private Function ExecuteTool(toolName As String, argsJsonStr As String) As String
    On Error GoTo ErrorHandler
    
    Debug.Print "ツール実行: " & toolName & " 引数: " & argsJsonStr
    
    ' VBA関数呼び出しの場合
    If toolName = "excel_function" Then
        ExecuteTool = CallVBAFunction(argsJsonStr)
        Exit Function
    End If
    
    ' MCPツール呼び出しの場合
    ' MCPブリッジのリクエスト本文を作成
    Dim requestJson As String
    requestJson = "{""name"":""" & toolName & """,""arguments"":" & argsJsonStr & "}"
    
    ' HTTP POSTリクエスト
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' MCPブリッジのfunction-callエンドポイント
    Dim apiEndpoint As String
    apiEndpoint = MCP_BRIDGE_URL & "/tools/function-call"
    
    http.Open "POST", apiEndpoint, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send requestJson
    
    ' レスポンスを返す
    ExecuteTool = http.responseText
    Exit Function
    
ErrorHandler:
    Dim errorMsg As String
    errorMsg = "ツール実行エラー: " & Err.description
    Debug.Print errorMsg
    ExecuteTool = errorMsg
End Function

' ツール定義を含めてGemini APIを呼び出す
Private Function CallGeminiAPIWithTools(prompt As String) As String
    On Error GoTo ErrorHandler
    
    Debug.Print "Gemini API呼び出し..."
    
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' APIエンドポイントにリクエスト
    http.Open "POST", GEMINI_API_URL & "?key=" & GEMINI_API_KEY, False
    http.SetRequestHeader "Content-Type", "application/json"
    
    ' リクエストボディの構築（ツール定義を含む）
    Dim requestBody As String
    requestBody = BuildGeminiRequestWithTools(prompt)
    
    ' リクエスト送信
    http.Send requestBody
    
    ' レスポンス取得
    Dim response As String
    response = http.responseText
    
    Debug.Print "Gemini API応答を受信しました"
    CallGeminiAPIWithTools = response
    Exit Function
    
ErrorHandler:
    Dim errorMsg As String
    errorMsg = "{""error"": """ & Err.description & """}"
    Debug.Print "Gemini API呼び出しエラー: " & Err.description
    CallGeminiAPIWithTools = errorMsg
End Function

' ツール定義を含むGeminiリクエストの構築
Private Function BuildGeminiRequestWithTools(prompt As String) As String
    On Error GoTo Err
    ' JSON文字列を直接構築
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
    
    ' ツール定義を追加
    jsonStr = jsonStr & """tools"": ["
    jsonStr = jsonStr & mToolsStr
    jsonStr = jsonStr & "],"
    
    ' 生成設定を追加 - ReAct用に最適化
    jsonStr = jsonStr & """generationConfig"": {"
    jsonStr = jsonStr & "  ""temperature"": 0.1,"  ' より決定的な応答
    jsonStr = jsonStr & "  ""topP"": 0.8,"
    jsonStr = jsonStr & "  ""topK"": 40,"
    jsonStr = jsonStr & "  ""maxOutputTokens"": 1024,"  ' 十分な出力長
    jsonStr = jsonStr & "},"
    jsonStr = jsonStr & "}"
    
    BuildGeminiRequestWithTools = jsonStr
    
    Exit Function
    
Err:
    Debug.Print "リクエスト構築エラー: " & Err.description
    BuildGeminiRequestWithTools = "{}"
End Function

' 会話履歴に追加
Private Sub AddToHistory(role As String, content As String)
    Dim historyItem As Object
    Set historyItem = CreateObject("Scripting.Dictionary")
    
    historyItem.Add "role", role
    historyItem.Add "content", content
    
    mConversationHistory.Add historyItem
End Sub

' 会話履歴を含むプロンプトを生成
Private Function GetPromptWithHistory() As String
    Dim prompt As String
    prompt = "あなたは様々なツールが使えるReactタイプのAIアシスタントです。ユーザーの指示に基づいてツールを使用してタスクを実行します。" & vbCrLf & vbCrLf
    
    ' ツールの説明を追加
'    prompt = prompt & "利用可能なツール:" & vbCrLf
'
'    Dim i As Integer, j As Integer
'    For i = 1 To mToolCount
'        prompt = prompt & "- " & mTools(i).name & ": " & mTools(i).description & vbCrLf
'        prompt = prompt & "  パラメータ:" & vbCrLf
'
'        For j = 1 To UBound(mTools(i).parameters)
'            prompt = prompt & "  - " & mTools(i).parameters(j).name & ": " & mTools(i).parameters(j).description
'            If mTools(i).parameters(j).required Then
'                prompt = prompt & " (必須)"
'            End If
'            prompt = prompt & vbCrLf
'        Next j
'
'        prompt = prompt & vbCrLf
'    Next i
    
    prompt = prompt & "ユーザーの指示に従い、必要なツールを使用して目的を達成してください。" & vbCrLf & vbCrLf
    
    ' 会話履歴を追加
    prompt = prompt & "これまでの会話:" & vbCrLf
    
    Dim item As Variant
    For Each item In mConversationHistory
        Dim role As String, content As String
        role = item("role")
        content = item("content")
        
        Select Case role
            Case "user"
                prompt = prompt & "ユーザー: " & content & vbCrLf
            Case "assistant"
                prompt = prompt & "アシスタント: " & content & vbCrLf
            Case "tool"
                prompt = prompt & "ツール結果: " & content & vbCrLf
        End Select
    Next item
    
    prompt = prompt & vbCrLf & "次の応答を生成してください。"
    
    GetPromptWithHistory = prompt
End Function

' JSON文字列のエスケープ
Private Function EscapeJsonString(text As String) As String
    On Error Resume Next
    
    Dim result As String
    result = text
    
    ' 特殊文字のエスケープ
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\r")
    result = Replace(result, vbTab, "\t")
    
    EscapeJsonString = result
End Function

' JSONレスポンスからツール定義を抽出して直接mToolsに設定
Private Sub ExtractToolsFromJsonString(jsonStr As String)
    On Error GoTo ErrorHandler
    
    Debug.Print "JSONからツール定義を抽出しています..."
    
    ' 変数の明示的宣言
    Dim jsonObject As Object
    Dim toolsArray As Collection
    Dim toolCount As Integer
    Dim toolEntry As Variant
    Dim functionDecl As Variant
    Dim i As Integer
    Dim index As Integer
    Dim toolInfo As Dictionary
    Dim toolName As String
    Dim description As String
    Dim parameters As Dictionary
    Dim properties As Dictionary
    Dim requiredParams As Collection
    Dim paramCount As Integer
    Dim j As Integer
    Dim k As Integer
    Dim paramNames As Variant
    Dim paramName As String
    Dim paramInfo As Dictionary
    Dim paramType As String
    Dim paramDesc As String
    Dim isRequired As Boolean
    
    ' JsonConverterを使ってJSONを解析
    Set jsonObject = JsonConverter.ParseJson(jsonStr)
    
    ' ツール配列を取得
    Set toolsArray = New Collection
    toolCount = 0
    
    If jsonObject.Exists("tools") Then
        ' 各ツールエントリを処理
        For Each toolEntry In jsonObject("tools")
            If TypeOf toolEntry Is Dictionary Then
                If toolEntry.Exists("function_declarations") Then
                    For Each functionDecl In toolEntry("function_declarations")
                        toolsArray.Add functionDecl
                    Next functionDecl
                End If
            End If
        Next toolEntry
        
        toolCount = toolsArray.Count
    End If
    
    ' ツール数に基づいて配列を初期化
    mToolCount = toolCount
    If mToolCount > 0 Then
        ReDim mTools(1 To mToolCount)
    Else
        ReDim mTools(1 To 1)  ' 最小サイズ
        mToolCount = 0
        Debug.Print "ツールが見つかりませんでした"
        Exit Sub
    End If
    
    Debug.Print toolCount & "個のツールが見つかりました。解析を開始します..."
    
    ' 各ツールの情報を抽出
    For i = 1 To toolsArray.Count
        ' インデックスの調整（VBAは1始まり）
        index = i
        
        Set toolInfo = toolsArray(i)
        
        ' ツール名と説明を取得
        toolName = ""
        description = ""
        
        If toolInfo.Exists("name") Then
            toolName = CStr(toolInfo("name"))
        Else
            toolName = "unnamed_tool_" & i
        End If
        
        ' 説明があれば取得
        If toolInfo.Exists("description") Then
            description = CStr(toolInfo("description"))
        Else
            description = "ツールの説明がありません"
        End If
        
        ' ツール情報を設定
        mTools(index).name = toolName
        mTools(index).description = description
        
        Debug.Print "ツール[" & index & "]: " & toolName & " - " & left(description, 30) & IIf(Len(description) > 30, "...", "")
        
        ' パラメータ情報を抽出
        If toolInfo.Exists("parameters") Then
            Set parameters = toolInfo("parameters")
            
            ' プロパティ情報を取得
            Set properties = New Dictionary
            Set requiredParams = New Collection
            paramCount = 0
            
            ' パラメータ数をカウント
            If parameters.Exists("properties") Then
                Set properties = parameters("properties")
                paramCount = properties.Count
            End If
            
            ' 必須パラメータのリストを取得
            If parameters.Exists("required") Then
                If TypeOf parameters("required") Is Collection Then
                    Set requiredParams = parameters("required")
                End If
            End If
            
            ' パラメータが存在する場合
            If paramCount > 0 Then
                Debug.Print "  パラメータ数: " & paramCount
                
                ' パラメータ配列を初期化
                ReDim mTools(index).parameters(1 To paramCount)
                
                ' 各パラメータの詳細を抽出
                paramNames = properties.Keys
                
                For j = 1 To paramCount
                    paramName = CStr(paramNames(j - 1))
                    Set paramInfo = properties(paramName)
                    
                    ' タイプを取得（大文字を小文字に変換）
                    paramType = "object"  ' デフォルト値
                    If paramInfo.Exists("type") Then
                        paramType = LCase(CStr(paramInfo("type")))
                    End If
                    
                    ' 説明を取得
                    paramDesc = paramName & "パラメータ"  ' デフォルト値
                    If paramInfo.Exists("description") Then
                        paramDesc = CStr(paramInfo("description"))
                    End If
                    
                    ' 必須かどうかをチェック
                    isRequired = False
                    On Error Resume Next
                    For k = 1 To requiredParams.Count
                        If CStr(requiredParams(k)) = paramName Then
                            isRequired = True
                            Exit For
                        End If
                    Next k
                    On Error GoTo ErrorHandler
                    
                    ' パラメータ情報を設定
                    mTools(index).parameters(j).name = paramName
                    mTools(index).parameters(j).description = paramDesc
                    mTools(index).parameters(j).paramType = paramType
                    mTools(index).parameters(j).required = isRequired
                    
                    Debug.Print "    パラメータ[" & j & "]: " & paramName & " (" & paramType & ")" & IIf(isRequired, " [必須]", "")
                Next j
            Else
                ' パラメータがない場合
                Debug.Print "  パラメータなし"
                ReDim mTools(index).parameters(1 To 1)
                mTools(index).parameters(1).name = "なし"
                mTools(index).parameters(1).description = "このツールにはパラメータが必要ありません"
                mTools(index).parameters(1).paramType = "object"
                mTools(index).parameters(1).required = False
            End If
        Else
            ' パラメータブロックがない場合
            Debug.Print "  パラメータブロックが見つかりません。デフォルト設定を使用します。"
            ReDim mTools(index).parameters(1 To 1)
            mTools(index).parameters(1).name = "args"
            mTools(index).parameters(1).description = "ツールに渡す引数（JSON形式）"
            mTools(index).parameters(1).paramType = "object"
            mTools(index).parameters(1).required = True
        End If
    Next i
    
    Debug.Print "ツール定義の抽出が完了しました。"
    Exit Sub
    
ErrorHandler:
    ' エラー情報を詳細に表示
    Dim errorMsg As String
    errorMsg = "JSON解析エラー: " & Err.Number & " - " & Err.description
    If Erl > 0 Then
        errorMsg = errorMsg & " (行: " & Erl & ")"
    End If
    Debug.Print errorMsg
    
    ' ツール配列を初期化
    mToolCount = 0
    ReDim mTools(1 To 1)
    mTools(1).name = "error"
    mTools(1).description = "ツール情報の取得中にエラーが発生しました"
    ReDim mTools(1).parameters(1 To 1)
    mTools(1).parameters(1).name = "error"
    mTools(1).parameters(1).description = errorMsg
    mTools(1).parameters(1).paramType = "string"
    mTools(1).parameters(1).required = False
End Sub

' Gemini形式のJSONからツール配列を抽出する
Private Function ExtractToolsArray(jsonStr As String) As String
    On Error GoTo ErrorHandler
    
    ' JsonConverterを使ってJSONを解析
    Dim jsonObject As Object
    Set jsonObject = JsonConverter.ParseJson(jsonStr)
    
    ' toolsプロパティの値をJSON文字列として返す
    If jsonObject.Exists("tools") Then
        ExtractToolsArray = JsonConverter.ConvertToJson(jsonObject("tools"))
    Else
        ExtractToolsArray = "[]"
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "ツール配列抽出エラー: " & Err.description
    ExtractToolsArray = "[]"
End Function

' リアルタイム表示用の関数
Private Sub DisplayRealtimeOutput(outputType As String, outputContent As String)
    If Not mRealTimeDisplay Then Exit Sub
    
    Dim isNext As Boolean
    isNext = True
    
    ' 既存の出力を更新
    With mDisplayWorksheet.Cells(mCurrentDisplayRow, 1)
        ' 出力タイプに応じた表示形式
        Dim prefix As String
        Dim cellColor As Long
        Dim outMessage As String
        
        outMessage = ShortenContent(outputContent, 100)
        Select Case outputType
            Case "thought"
                prefix = "・ "
                cellColor = RGB(255, 255, 220)
            Case "action"
                prefix = "・ "
                cellColor = RGB(230, 245, 230)
            Case "observation"
                prefix = "・ "
                cellColor = RGB(245, 230, 245)
            Case Else
                prefix = "● "
                cellColor = RGB(240, 240, 240)
                isNext = False
                outMessage = ShortenContent(outputContent, 500)
        End Select
        
        ' 内容を表示（長すぎる場合は切り詰める）
        .value = prefix & outMessage
        .Interior.Color = cellColor
        .WrapText = True
        .Rows.AutoFit
    End With
    
    ' 次の行へ
    mCurrentDisplayRow = mCurrentDisplayRow + 1
    If isNext Then
        mDisplayWorksheet.Cells(mCurrentDisplayRow, 1).value = "思考中..."
    End If
    
    ' 表示を更新
    DoEvents
End Sub

' 長いコンテンツを短く表示するためのヘルパー関数
Private Function ShortenContent(content As String, maxLength As Integer) As String
    If Len(content) <= maxLength Then
        ShortenContent = content
    Else
        ShortenContent = left(content, maxLength - 3) & "..."
    End If
End Function

