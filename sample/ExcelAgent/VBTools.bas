' VBA関数呼び出し
Public Function CallVBAFunction(inputJson As String) As String
    On Error GoTo ErrorHandler
    
    Debug.Print "VBA関数呼び出し: " & inputJson
    
    ' JSONを解析（単純な正規表現パースで代用）
    Dim functionName As String
    Dim paramsArray As String
    
    ' 関数名を抽出
    Dim nameRegex As Object
    Set nameRegex = CreateObject("VBScript.RegExp")
    nameRegex.Pattern = """functionName""\s*:\s*""([^""]+)"""
    
    Dim nameMatches As Object
    Set nameMatches = nameRegex.Execute(inputJson)
    
    If nameMatches.Count > 0 Then
        functionName = nameMatches(0).SubMatches(0)
    Else
        CallVBAFunction = "エラー: 関数名が見つかりません"
        Exit Function
    End If
    
    ' パラメータを抽出
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
    
    ' 関数名に応じた処理
    Select Case functionName
        Case "GetCellValue"
            ' パラメータ例: ["Sheet1", "A1"]
            ' 単純な正規表現で配列要素を抽出
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
                    CallVBAFunction = "エラー: セル値の取得に失敗しました - " & Err.description
                End If
                On Error GoTo ErrorHandler
            Else
                CallVBAFunction = "エラー: パラメータ不足"
            End If
            
        Case "CalculateSum"
            ' パラメータ例: ["Sheet1", "A1:A10"]
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
                    CallVBAFunction = "エラー: 合計の計算に失敗しました - " & Err.description
                End If
                On Error GoTo ErrorHandler
            Else
                CallVBAFunction = "エラー: パラメータ不足"
            End If
            
        Case "ShowMessage"
            ' パラメータ例: ["こんにちは"]
            Dim msgRegex As Object
            Set msgRegex = CreateObject("VBScript.RegExp")
            msgRegex.Pattern = """([^""]+)"""
            
            Dim msgMatches As Object
            Set msgMatches = msgRegex.Execute(paramsArray)
            
            If msgMatches.Count >= 1 Then
                Dim msg As String
                msg = msgMatches(0).SubMatches(0)
                MsgBox msg, vbInformation, "メッセージ"
                CallVBAFunction = "メッセージを表示しました: " & msg
            Else
                CallVBAFunction = "エラー: メッセージが指定されていません"
            End If
            
        Case Else
            CallVBAFunction = "エラー: 未知の関数 '" & functionName & "'"
    End Select
    
    Debug.Print "VBA関数実行結果: " & functionName & " - " & Left(CallVBAFunction, 50) & IIf(Len(CallVBAFunction) > 50, "...", "")
    Exit Function
    
ErrorHandler:
    CallVBAFunction = "エラー: " & Err.description
    Debug.Print "VBA関数実行エラー: " & Err.description
End Function

