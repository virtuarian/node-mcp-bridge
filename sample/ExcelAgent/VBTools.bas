
' 関数の説明
Public Function GetProcedureInfo()
    Dim info As String
    info = "Excel操作ツールの使用方法:" & vbCrLf _
        & "- CreateWorksheet: 新しいワークシートを作成します" & vbCrLf _
        & "  使用例: excel_function { ""functionName"": ""CreateWorksheet"", ""params"": [""新しいシート"", ""last""] }" & vbCrLf _
        & "- CreateChart: グラフを作成します" & vbCrLf _
        & "  使用例: excel_function { ""functionName"": ""CreateChart"", ""params"": [""Sheet1"", ""A1:B10"", ""column"", ""売上グラフ""] }" & vbCrLf _
        & "- CreateTable: Excelテーブルを作成します" & vbCrLf _
        & "  使用例: excel_function { ""functionName"": ""CreateTable"", ""params"": [""Sheet1"", ""A1:D10"", ""MyTable""] }" & vbCrLf _
        & "- CreatePivotTable: ピボットテーブルを作成します" & vbCrLf _
        & "  使用例: excel_function { ""functionName"": ""CreatePivotTable"", ""params"": [""データシート"", ""A1:D100"", ""ピボットシート"", ""A3"", ""部署,社員"", ""年度"", ""売上""] }" & vbCrLf _
        & "- SortData: データを並べ替えます" & vbCrLf _
        & "  使用例: excel_function { ""functionName"": ""SortData"", ""params"": [""Sheet1"", ""A1:D10"", ""A"", ""ascending"", true] }" & vbCrLf _
        & "- ApplyFilter: データにフィルターを適用します" & vbCrLf _
        & "  使用例: excel_function { ""functionName"": ""ApplyFilter"", ""params"": [""Sheet1"", ""A1:D10"", ""A"", ""値"", ""equals""] }" & vbCrLf _
        & "- ApplyConditionalFormat: 条件付き書式を適用します" & vbCrLf _
        & "  使用例: excel_function { ""functionName"": ""ApplyConditionalFormat"", ""params"": [""Sheet1"", ""A1:D10"", ""cellvalue"", ""greater than,100"", ""lightred""] }" & vbCrLf _
        & "- InsertFormula: 数式を挿入します" & vbCrLf _
        & "  使用例: excel_function { ""functionName"": ""InsertFormula"", ""params"": [""Sheet1"", ""A1"", ""SUM(B1:B10)""] }" & vbCrLf _
        
        
    info = info & "- GetMultipleCellValues: 複数のセル値を一度に取得します" & vbCrLf _
        & "  使用例: excel_function { ""functionName"": ""GetMultipleCellValues"", ""params"": [""Sheet1"", ""A1,B2,C3""] }" & vbCrLf _
        & "  JSON形式: excel_function { ""functionName"": ""GetMultipleCellValues"", ""params"": [""Sheet1"", ""[""A1"",""B2"",""C3""]""] }" & vbCrLf _
        & "  カスタムキー: excel_function { ""functionName"": ""GetMultipleCellValues"", ""params"": [""Sheet1"", ""{""""name"":""A1"",""value"":""B2""""}""] }" & vbCrLf _
        & "- SetMultipleCellValues: 複数のセル値を一度に設定します" & vbCrLf _
        & "  使用例: excel_function { ""functionName"": ""SetMultipleCellValues"", ""params"": [""Sheet1"", ""{""""A1"":""値1"",""B2"":""値2""""}""] }" & vbCrLf

    info = info & "- FormatCells: セルの書式を設定します" & vbCrLf _
        & "  使用例: excel_function { ""functionName"": ""FormatCells"", ""params"": [""Sheet1"", ""A1:B10"", ""{""font"":{""name"":""Arial"",...}}""] }" & vbCrLf _
        & "- FormatBorders: セルの罫線を設定します" & vbCrLf _
        & "  使用例: excel_function { ""functionName"": ""FormatBorders"", ""params"": [""Sheet1"", ""A1:D10"", ""{""position"":""outline"",...}""] }" & vbCrLf _
        & "- FormatNumberStyle: 数値書式を設定します" & vbCrLf _
        & "  使用例: excel_function { ""functionName"": ""FormatNumberStyle"", ""params"": [""Sheet1"", ""A1:A10"", ""currency""] }" & vbCrLf

    GetProcedureInfo = info

End Function

' VBA関数呼び出し
Public Function CallVBAFunction(inputJson As String) As String
    On Error GoTo ErrorHandler
    
    Debug.Print "VBA関数呼び出し: " & inputJson
    
    ' JSONを解析
    Dim jsonObject As Object
    Set jsonObject = JsonConverter.ParseJson(inputJson)
    
    ' 関数名を取得
    Dim functionName As String
    If jsonObject.Exists("functionName") Then
        functionName = jsonObject("functionName")
    Else
        CallVBAFunction = "エラー: 関数名が指定されていません"
        Exit Function
    End If
    
    ' パラメータを取得
    Dim params As Collection
    If jsonObject.Exists("params") Then
        Set params = jsonObject("params")
    Else
        Set params = New Collection
    End If
    
    ' 関数名に応じた処理
    Select Case functionName
        Case "GetCellValue"
            ' パラメータ例: ["Sheet1", "A1"]
            If params.Count >= 2 Then
                CallVBAFunction = GetCellValue(params(1), params(2))
            Else
                CallVBAFunction = "エラー: GetCellValueには2つのパラメータが必要です"
            End If
        
        Case "SetCellValue"
            ' パラメータ例: ["Sheet1", "A1", "新しい値"]
            If params.Count >= 3 Then
                CallVBAFunction = SetCellValue(params(1), params(2), params(3))
            Else
                CallVBAFunction = "エラー: SetCellValueには3つのパラメータが必要です"
            End If
            
        Case "CreateWorksheet"
            ' パラメータ例: ["新しいシート"]
            ' または ["新しいシート", "last"]
            If params.Count >= 1 Then
                If params.Count >= 2 Then
                    CallVBAFunction = CreateWorksheet(params(1), params(2))
                Else
                    CallVBAFunction = CreateWorksheet(params(1))
                End If
            Else
                CallVBAFunction = "エラー: CreateWorksheetには少なくとも1つのパラメータが必要です"
            End If
            
        Case "CreateChart"
            ' パラメータ例: ["Sheet1", "A1:B10", "column", "売上グラフ", 100, 100, 350, 250]
            If params.Count >= 3 Then
                ' デバッグ出力を追加
                Debug.Print "CreateChart - パラメータ数: " & params.Count
                For i = 1 To params.Count
                    Debug.Print "params(" & i & "): " & params(i)
                Next i
                
                Dim title As String, left As Long, top As Long, width As Long, height As Long
                
                ' シート名、範囲、グラフ種類を変数に格納
                Dim sheetName As String, dataRange As String, chartType As String
                sheetName = params(1)
                dataRange = params(2)
                chartType = params(3)
                
                ' オプションパラメータを安全に取得
                title = ""
                left = 100
                top = 100 
                width = 350
                height = 250
                
                ' パラメータの数に応じてオプション引数を設定
                If params.Count >= 4 Then title = params(4)
                If params.Count >= 5 Then left = CLng(params(5))
                If params.Count >= 6 Then top = CLng(params(6))
                If params.Count >= 7 Then width = CLng(params(7))
                If params.Count >= 8 Then height = CLng(params(8))
                
                CallVBAFunction = CreateChart(sheetName, dataRange, chartType, title, left, top, width, height)
            Else
                CallVBAFunction = "エラー: CreateChartには少なくとも3つのパラメータが必要です"
            End If
            
        Case "CreateTable"
            ' パラメータ例: ["Sheet1", "A1:D10", "MyTable", "TableStyleMedium2"]
            If params.Count >= 2 Then
                Dim tableName As String, tableStyle As String
                
                ' デバッグ出力を追加
                Debug.Print "CreateTable - パラメータ数: " & params.Count
                For i = 1 To params.Count
                    Debug.Print "params(" & i & "): " & params(i)
                Next i
                
                ' 修正：安全にパラメータを取得
                If params.Count >= 3 Then
                    tableName = params(3)
                Else
                    tableName = ""
                End If
                
                If params.Count >= 4 Then
                    tableStyle = params(4)
                Else
                    tableStyle = "TableStyleMedium2"
                End If
                
                CallVBAFunction = CreateTable(params(1), params(2), tableName, tableStyle)
            Else
                CallVBAFunction = "エラー: CreateTableには少なくとも2つのパラメータが必要です"
            End If
            
        Case "CreatePivotTable"
            ' パラメータ例: ["データシート", "A1:D100", "ピボットシート", "A3", "部署,社員", "年度", "売上"]
             If params.Count >= 4 Then
                Dim rowFields As String, columnFields As String, dataFields As String
                
                ' 安全にパラメータを取得
                rowFields = ""
                columnFields = ""
                dataFields = ""
                
                If params.Count >= 5 Then rowFields = params(5)
                If params.Count >= 6 Then columnFields = params(6)
                If params.Count >= 7 Then dataFields = params(7)
                
                CallVBAFunction = CreatePivotTable(params(1), params(2), params(3), params(4), _
                                                rowFields, columnFields, dataFields)
            Else
                CallVBAFunction = "エラー: CreatePivotTableには少なくとも4つのパラメータが必要です"
            End If

        Case "SortData"
            ' パラメータ例: ["Sheet1", "A1:D10", "A", "ascending", true]
             If params.Count >= 3 Then
                Dim sortOrder As String, hasHeader As Boolean
                
                ' 安全にパラメータを取得（IIf使用しない）
                sortOrder = "ascending"  ' デフォルト値
                hasHeader = True         ' デフォルト値
                
                If params.Count >= 4 Then
                    sortOrder = params(4)
                End If
                
                If params.Count >= 5 Then
                    hasHeader = CBool(params(5))
                End If
                
                CallVBAFunction = SortData(params(1), params(2), params(3), sortOrder, hasHeader)
            Else
                CallVBAFunction = "エラー: SortDataには少なくとも3つのパラメータが必要です"
            End If

        Case "ApplyFilter"
            ' パラメータ例: ["Sheet1", "A1:D10", "A", "値", "equals"]
            If params.Count >= 4 Then
                Dim operator As String
                
                ' 安全にパラメータを取得
                operator = "equals"  ' デフォルト値
                
                If params.Count >= 5 Then
                    operator = params(5)
                End If
                
                CallVBAFunction = ApplyFilter(params(1), params(2), params(3), params(4), operator)
            Else
                CallVBAFunction = "エラー: ApplyFilterには少なくとも4つのパラメータが必要です"
            End If

        Case "ApplyConditionalFormat"
            ' パラメータ例: ["Sheet1", "A1:D10", "cellvalue", "greater than,100", "lightred"]
            If params.Count >= 4 Then
                Dim formatStyle As String
                
                ' 安全にパラメータを取得
                formatStyle = "default"  ' デフォルト値
                
                If params.Count >= 5 Then
                    formatStyle = params(5)
                End If
                
                CallVBAFunction = ApplyConditionalFormat(params(1), params(2), params(3), params(4), formatStyle)
            Else
                CallVBAFunction = "エラー: ApplyConditionalFormatには少なくとも4つのパラメータが必要です"
            End If

        Case "InsertFormula"
            ' パラメータ例: ["Sheet1", "A1", "SUM(B1:B10)"]
            If params.Count >= 3 Then
                ' params(1) = シート名
                ' params(2) = セル参照
                ' params(3) = 数式
                CallVBAFunction = InsertFormula(params(1), params(2), params(3))
            Else
                CallVBAFunction = "エラー: InsertFormulaには3つのパラメータが必要です"
            End If

       
        Case "GetMultipleCellValues"
            ' パラメータ例: ["Sheet1", "A1,B2,C3"] または ["Sheet1", "[\"A1\",\"B2\",\"C3\"]"] または ["Sheet1", "{\"name\":\"A1\",\"value\":\"B2\"}"]
            If params.Count >= 2 Then
                CallVBAFunction = GetMultipleCellValues(params(1), params(2))
            Else
                CallVBAFunction = "エラー: GetMultipleCellValuesには少なくとも2つのパラメータが必要です"
            End If

        Case "SetMultipleCellValues"
            ' パラメータ例: ["Sheet1", "{\"A1\":\"新しい値1\",\"B2\":\"新しい値2\"}"]
            If params.Count >= 2 Then
                CallVBAFunction = SetMultipleCellValues(params(1), params(2))
            Else
                CallVBAFunction = "エラー: SetMultipleCellValuesには少なくとも2つのパラメータが必要です"
            End If

        Case "FormatCells"
            ' パラメータ例: ["Sheet1", "A1:B10", "{\"font\":{\"name\":\"Arial\",\"size\":12,\"bold\":true,\"color\":\"blue\"},\"fill\":{\"color\":\"lightyellow\"},\"alignment\":{\"horizontal\":\"center\"}}"]
            If params.Count >= 2 Then
                Dim formatSettings As String
                formatSettings = ""
                
                If params.Count >= 3 Then
                    formatSettings = params(3)
                End If
                
                CallVBAFunction = FormatCells(params(1), params(2), formatSettings)
            Else
                CallVBAFunction = "エラー: FormatCellsには少なくとも2つのパラメータが必要です"
            End If

        Case "FormatBorders"
            ' パラメータ例: ["Sheet1", "A1:D10", "{\"position\":\"outline\",\"weight\":\"medium\",\"color\":\"blue\",\"style\":\"continuous\"}"]
            If params.Count >= 3 Then
                CallVBAFunction = FormatBorders(params(1), params(2), params(3))
            Else
                CallVBAFunction = "エラー: FormatBordersには少なくとも3つのパラメータが必要です"
            End If

        Case "FormatNumberStyle"
            ' パラメータ例: ["Sheet1", "A1:A10", "currency"] または ["Sheet1", "A1:A10", "#,##0.00"]
            If params.Count >= 3 Then
                CallVBAFunction = FormatNumberStyle(params(1), params(2), params(3))
            Else
                CallVBAFunction = "エラー: FormatNumberStyleには少なくとも3つのパラメータが必要です"
            End If
             
        Case Else
            CallVBAFunction = "エラー: 未知の関数 '" & functionName & "'"
            
    End Select
    
    'Debug.Print "VBA関数実行結果: " & functionName & " - " & Left(CallVBAFunction, 50) & IIf(Len(CallVBAFunction) > 50, "...", "")
    Exit Function
    
ErrorHandler:
    CallVBAFunction = "エラー: " & Err.description
    Debug.Print "VBA関数実行エラー: " & Err.description
End Function

' GetCellValue関数
Public Function GetCellValue(sheetName As String, cellReference As String) As String
    On Error GoTo ErrorHandler
    
    Dim result As String
    result = CStr(ThisWorkbook.Sheets(sheetName).Range(cellReference).value)
    GetCellValue = result
    Exit Function
    
ErrorHandler:
    GetCellValue = "エラー: セル値の取得に失敗しました - " & Err.description
End Function

' SetCellValue関数
Public Function SetCellValue(sheetName As String, cellReference As String, newValue As String) As String
    On Error GoTo ErrorHandler
    
    ThisWorkbook.Sheets(sheetName).Range(cellReference).value = newValue
    SetCellValue = "セル " & sheetName & "!" & cellReference & " に値 """ & newValue & """ を設定しました"
    Exit Function
    
ErrorHandler:
    SetCellValue = "エラー: セル値の設定に失敗しました - " & Err.description
End Function

' ワークシートの作成
Public Function CreateWorksheet(sheetName As String, Optional position As String = "last") As String
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    
    ' 既存のシート名をチェック
    Dim sheetExists As Boolean
    sheetExists = False
    
    For Each sheet In ThisWorkbook.Sheets
        If sheet.name = sheetName Then
            sheetExists = True
            Exit For
        End If
    Next sheet
    
    If sheetExists Then
        CreateWorksheet = "シート '" & sheetName & "' は既に存在します"
        Exit Function
    End If
    
    ' 位置パラメータに基づいてシートを追加
    Select Case LCase(position)
        Case "first"
            Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        Case "last"
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        Case Else
            ' 数値として解釈可能な場合は、その位置に挿入
            On Error Resume Next
            Dim posNum As Integer
            posNum = CInt(position)
            
            If Err.Number = 0 And posNum > 0 And posNum <= ThisWorkbook.Worksheets.Count Then
                Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(posNum))
            Else
                ' デフォルトは最後
                Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            End If
            On Error GoTo ErrorHandler
    End Select
    
    ws.name = sheetName
    
    CreateWorksheet = "シート '" & sheetName & "' を作成しました"
    Exit Function
    
ErrorHandler:
    CreateWorksheet = "エラー: ワークシートの作成に失敗しました - " & Err.description
End Function

' 簡易グラフの作成
Public Function CreateChart(sheetName As String, dataRange As String, chartType As String, _
                            Optional title As String = "", Optional left As Long = 100, _
                            Optional top As Long = 100, Optional width As Long = 350, _
                            Optional height As Long = 250) As String
    On Error GoTo ErrorHandler
    
    ' シートの存在確認
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        CreateChart = "エラー: シート '" & sheetName & "' が見つかりません"
        Exit Function
    End If
    
    ' チャートタイプの解析
    Dim chartTypeEnum As XlChartType
    Select Case LCase(chartType)
        Case "column", "縦棒", "棒グラフ"
            chartTypeEnum = xlColumnClustered
        Case "bar", "横棒"
            chartTypeEnum = xlBarClustered
        Case "line", "折れ線"
            chartTypeEnum = xlLine
        Case "pie", "円", "円グラフ"
            chartTypeEnum = xlPie
        Case "area", "面", "面グラフ"
            chartTypeEnum = xlAreaStacked
        Case "scatter", "散布図"
            chartTypeEnum = xlXYScatter
        Case Else
            chartTypeEnum = xlColumnClustered ' デフォルトは縦棒グラフ
    End Select
    
    ' グラフ作成
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(left, top, width, height)
    
    ' データ範囲設定
    chartObj.Chart.SetSourceData Source:=ws.Range(dataRange)
    chartObj.Chart.chartType = chartTypeEnum
    
    ' タイトル設定
    If title <> "" Then
        chartObj.Chart.HasTitle = True
        chartObj.Chart.ChartTitle.text = title
    End If
    
    CreateChart = "シート '" & sheetName & "' にグラフを作成しました"
    Exit Function
    
ErrorHandler:
    CreateChart = "エラー: グラフの作成に失敗しました - " & Err.description
End Function

' Excelテーブルの作成
Public Function CreateTable(sheetName As String, dataRange As String, _
                            Optional tableName As String = "", _
                            Optional tableStyle As String = "TableStyleMedium2") As String
    On Error GoTo ErrorHandler
    
    ' シートの存在確認
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        CreateTable = "エラー: シート '" & sheetName & "' が見つかりません"
        Exit Function
    End If
    
    ' テーブル名が指定されていない場合は自動生成
    If tableName = "" Then
        tableName = "Table" & (ThisWorkbook.TableStyles.Count + 1)
    End If
    
    ' テーブル作成
    Dim tableObj As ListObject
    Set tableObj = ws.ListObjects.Add(xlSrcRange, ws.Range(dataRange), , xlYes)
    
    ' テーブル名と書式の設定
    On Error Resume Next
    tableObj.name = tableName
    tableObj.tableStyle = tableStyle
    On Error GoTo ErrorHandler
    
    CreateTable = "シート '" & sheetName & "' に '" & tableName & "' テーブルを作成しました"
    Exit Function
    
ErrorHandler:
    CreateTable = "エラー: テーブルの作成に失敗しました - " & Err.description
End Function

' ピボットテーブルの作成
Public Function CreatePivotTable(sourceSheet As String, sourceRange As String, _
                                destinationSheet As String, pivotLocation As String, _
                                Optional rowFields As String = "", _
                                Optional columnFields As String = "", _
                                Optional dataFields As String = "") As String
    On Error GoTo ErrorHandler
    
    ' ソースシートの存在確認
    Dim srcWs As Worksheet
    On Error Resume Next
    Set srcWs = ThisWorkbook.Sheets(sourceSheet)
    On Error GoTo ErrorHandler
    
    If srcWs Is Nothing Then
        CreatePivotTable = "エラー: ソースシート '" & sourceSheet & "' が見つかりません"
        Exit Function
    End If
    
    ' 出力先シートの存在確認
    Dim destWs As Worksheet
    On Error Resume Next
    Set destWs = ThisWorkbook.Sheets(destinationSheet)
    On Error GoTo ErrorHandler
    
    If destWs Is Nothing Then
        CreatePivotTable = "エラー: 出力先シート '" & destinationSheet & "' が見つかりません"
        Exit Function
    End If
    
    ' ピボットキャッシュを作成
    Dim pvtCache As PivotCache
    Set pvtCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcWs.Range(sourceRange))
    
    ' ピボットテーブルを作成
    Dim pvtTable As PivotTable
    Set pvtTable = pvtCache.CreatePivotTable(TableDestination:=destWs.Range(pivotLocation))
    
    ' 行フィールドを追加
    If rowFields <> "" Then
        Dim rowFieldsArray() As String
        rowFieldsArray = Split(rowFields, ",")
        
        Dim rowField As Variant
        For Each rowField In rowFieldsArray
            pvtTable.PivotFields(Trim(rowField)).Orientation = xlRowField
        Next rowField
    End If
    
    ' 列フィールドを追加
    If columnFields <> "" Then
        Dim colFieldsArray() As String
        colFieldsArray = Split(columnFields, ",")
        
        Dim colField As Variant
        For Each colField In colFieldsArray
            pvtTable.PivotFields(Trim(colField)).Orientation = xlColumnField
        Next colField
    End If
    
    ' データフィールドを追加
    If dataFields <> "" Then
        Dim dataFieldsArray() As String
        dataFieldsArray = Split(dataFields, ",")
        
        Dim dataField As Variant
        For Each dataField In dataFieldsArray
            Dim fieldName As String
            fieldName = Trim(dataField)
            
            ' 集計方法の指定があるか確認（例: "売上:Sum"）
            Dim summaryFunction As XlConsolidationFunction
            summaryFunction = xlSum  ' デフォルトはSum
            
            If InStr(fieldName, ":") > 0 Then
                Dim parts() As String
                parts = Split(fieldName, ":")
                fieldName = Trim(parts(0))
                
                Select Case LCase(Trim(parts(1)))
                    Case "sum", "合計"
                        summaryFunction = xlSum
                    Case "count", "個数"
                        summaryFunction = xlCount
                    Case "average", "平均"
                        summaryFunction = xlAverage
                    Case "max", "最大"
                        summaryFunction = xlMax
                    Case "min", "最小"
                        summaryFunction = xlMin
                End Select
            End If
            
            pvtTable.AddDataField pvtTable.PivotFields(fieldName), fieldName & " の " & _
                Choose(summaryFunction, "合計", "個数", "平均", "最大", "最小"), summaryFunction
        Next dataField
    End If
    
    CreatePivotTable = "ピボットテーブルを作成しました: ソース=[" & sourceSheet & "]" & sourceRange & _
                       ", 出力先=[" & destinationSheet & "]" & pivotLocation
    Exit Function
    
ErrorHandler:
    CreatePivotTable = "エラー: ピボットテーブルの作成に失敗しました - " & Err.description
End Function

' データの並べ替え
Public Function SortData(sheetName As String, dataRange As String, _
                         sortField As String, Optional sortOrder As String = "ascending", _
                         Optional hasHeader As Boolean = True) As String
    On Error GoTo ErrorHandler
    
    ' シートの存在確認
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        SortData = "エラー: シート '" & sheetName & "' が見つかりません"
        Exit Function
    End If
    
    ' ソート順の設定
    Dim orderValue As XlSortOrder
    Select Case LCase(sortOrder)
        Case "asc", "ascending", "昇順"
            orderValue = xlAscending
        Case "desc", "descending", "降順"
            orderValue = xlDescending
        Case Else
            orderValue = xlAscending
    End Select
    
    ' ヘッダーの設定
    Dim headerValue As XlYesNoGuess
    If hasHeader Then
        headerValue = xlYes
    Else
        headerValue = xlNo
    End If
    
    ' 並べ替え実行
    ws.Range(dataRange).Sort _
        Key1:=ws.Range(sortField), _
        Order1:=orderValue, _
        Header:=headerValue
    
    SortData = "データを並べ替えました: [" & sheetName & "]" & dataRange & _
               ", ソートフィールド=" & sortField & ", 順序=" & sortOrder
    Exit Function
    
ErrorHandler:
    SortData = "エラー: データの並べ替えに失敗しました - " & Err.description
End Function

' フィルター適用
Public Function ApplyFilter(sheetName As String, dataRange As String, _
                            filterColumn As String, filterCriteria As String, _
                            Optional operator As String = "equals") As String
    On Error GoTo ErrorHandler
    
    ' シートの存在確認
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        ApplyFilter = "エラー: シート '" & sheetName & "' が見つかりません"
        Exit Function
    End If
    
    ' 範囲にオートフィルターがなければ適用
    Dim rng As Range
    Set rng = ws.Range(dataRange)
    
    If Not rng.AutoFilter Then
        rng.AutoFilter
    End If
    
    ' フィルター条件の特定
    Dim columnIndex As Integer
    
    ' 列名が指定された場合、列インデックスを取得
    If IsNumeric(filterColumn) Then
        columnIndex = CInt(filterColumn)
    Else
        ' 列名から列インデックスを取得
        Dim firstRow As Range
        Set firstRow = rng.Rows(1)
        
        Dim cell As Range
        Dim columnFound As Boolean
        columnFound = False
        
        For Each cell In firstRow.Cells
            If cell.value = filterColumn Then
                columnIndex = cell.Column - rng.Cells(1, 1).Column + 1
                columnFound = True
                Exit For
            End If
        Next cell
        
        If Not columnFound Then
            ApplyFilter = "エラー: 列 '" & filterColumn & "' が見つかりません"
            Exit Function
        End If
    End If
    
    ' 演算子設定
    Dim xlOperator As XlAutoFilterOperator
    Select Case LCase(operator)
        Case "equals", "equal", "eq", "="
            xlOperator = xlFilterValues
            ws.Range(dataRange).AutoFilter Field:=columnIndex, Criteria1:=filterCriteria
        Case "contains", "like", "含む"
            xlOperator = xlFilterValues
            ws.Range(dataRange).AutoFilter Field:=columnIndex, Criteria1:="*" & filterCriteria & "*"
        Case "greater", "greaterthan", "gt", ">", "より大きい"
            xlOperator = xlFilterValues
            ws.Range(dataRange).AutoFilter Field:=columnIndex, Criteria1:=">" & filterCriteria
        Case "less", "lessthan", "lt", "<", "より小さい"
            xlOperator = xlFilterValues
            ws.Range(dataRange).AutoFilter Field:=columnIndex, Criteria1:="<" & filterCriteria
        Case "beginswith", "startswith", "始まる"
            xlOperator = xlFilterValues
            ws.Range(dataRange).AutoFilter Field:=columnIndex, Criteria1:=filterCriteria & "*"
        Case "endswith", "終わる"
            xlOperator = xlFilterValues
            ws.Range(dataRange).AutoFilter Field:=columnIndex, Criteria1:="*" & filterCriteria
        Case Else
            ' デフォルトは完全一致
            xlOperator = xlFilterValues
            ws.Range(dataRange).AutoFilter Field:=columnIndex, Criteria1:=filterCriteria
    End Select
    
    ApplyFilter = "フィルターを適用しました: [" & sheetName & "]" & dataRange & _
                 ", 列=" & filterColumn & ", 条件=" & operator & " " & filterCriteria
    Exit Function
    
ErrorHandler:
    ApplyFilter = "エラー: フィルター適用に失敗しました - " & Err.description
End Function

' 条件付き書式の適用
Public Function ApplyConditionalFormat(sheetName As String, dataRange As String, _
                                     formatType As String, condition As String, _
                                     Optional formatStyle As String = "default") As String
    On Error GoTo ErrorHandler
    
    ' シートの存在確認
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        ApplyConditionalFormat = "エラー: シート '" & sheetName & "' が見つかりません"
        Exit Function
    End If
    
    ' 条件付き書式のタイプを設定
    Dim formatRng As Range
    Set formatRng = ws.Range(dataRange)
    
    ' 書式をクリア
    formatRng.FormatConditions.Delete
    
    ' 書式タイプと条件に基づいて適用
    Select Case LCase(formatType)
        Case "cellvalue", "cellisvalue", "セル値"
            ' 条件を解析
            Dim operator As XlFormatConditionOperator
            Dim value1 As String, value2 As String
            
            ' 条件とその値を分割（例: "greater than,100"）
            Dim condParts() As String
            condParts = Split(condition, ",")
            
            If UBound(condParts) >= 0 Then
                value1 = condParts(1)
                If UBound(condParts) >= 2 Then
                    value2 = condParts(2)
                End If
                
                Select Case LCase(condParts(0))
                    Case "greater than", "greaterthan", ">"
                        operator = xlGreater
                        formatRng.FormatConditions.Add Type:=xlCellValue, operator:=operator, Formula1:=value1
                    Case "less than", "lessthan", "<"
                        operator = xlLess
                        formatRng.FormatConditions.Add Type:=xlCellValue, operator:=operator, Formula1:=value1
                    Case "equal to", "equalto", "="
                        operator = xlEqual
                        formatRng.FormatConditions.Add Type:=xlCellValue, operator:=operator, Formula1:=value1
                    Case "between", "両方の間"
                        operator = xlBetween
                        formatRng.FormatConditions.Add Type:=xlCellValue, operator:=operator, Formula1:=value1, Formula2:=value2
                    Case Else
                        ' デフォルトは「次の値に等しい」
                        operator = xlEqual
                        formatRng.FormatConditions.Add Type:=xlCellValue, operator:=operator, Formula1:=value1
                End Select
            Else
                ApplyConditionalFormat = "エラー: 条件の形式が不正です - " & condition
                Exit Function
            End If
            
        Case "colorscale", "カラースケール"
            ' 2色または3色のカラースケール
            Dim colorCount As Integer
            colorCount = CInt(condition)
            
            If colorCount = 2 Then
                formatRng.FormatConditions.AddColorScale ColorScaleType:=2
            ElseIf colorCount = 3 Then
                formatRng.FormatConditions.AddColorScale ColorScaleType:=3
            Else
                ' デフォルトは2色
                formatRng.FormatConditions.AddColorScale ColorScaleType:=2
            End If
            
        Case "databar", "データバー"
            ' データバーの色を指定（例: "blue"）
            Dim barColor As XlDataBarColor
            
            Select Case LCase(condition)
                Case "blue", "青"
                    barColor = xlDataBarColorBlue
                Case "green", "緑"
                    barColor = xlDataBarColorGreen
                Case "red", "赤"
                    barColor = xlDataBarColorRed
                Case "orange", "オレンジ"
                    barColor = xlDataBarColorOrange
                Case "lightblue", "lightblue", "水色"
                    barColor = xlDataBarColorLightBlue
                Case "purple", "紫"
                    barColor = xlDataBarColorPurple
                Case Else
                    ' デフォルトは青
                    barColor = xlDataBarColorBlue
            End Select
            
            formatRng.FormatConditions.AddDatabar
            formatRng.FormatConditions(1).ShowValue = True
            
        Case "iconset", "アイコンセット"
            ' アイコンセットの種類を指定（例: "3arrows"）
            Dim iconSetType As XlIconSet
            
            Select Case LCase(condition)
                Case "3arrows", "3矢印"
                    iconSetType = xl3Arrows
                Case "3trafficlights", "3信号"
                    iconSetType = xl3TrafficLights
                Case "3signs", "3記号"
                    iconSetType = xl3Signs
                Case "3symbols", "3シンボル"
                    iconSetType = xl3Symbols
                Case "4arrows", "4矢印"
                    iconSetType = xl4Arrows
                Case "4trafficlights", "4信号"
                    iconSetType = xl4TrafficLights
                Case "5arrows", "5矢印"
                    iconSetType = xl5Arrows
                Case "5ratings", "5評価"
                    iconSetType = xl5Ratings
                Case Else
                    ' デフォルトは3矢印
                    iconSetType = xl3Arrows
            End Select
            
            formatRng.FormatConditions.AddIconSetCondition
            formatRng.FormatConditions(1).IconSet = ThisWorkbook.IconSets(iconSetType)
            
        Case "formula", "数式"
            ' 数式に基づく条件付き書式
            formatRng.FormatConditions.Add Type:=xlExpression, Formula1:=condition
            
        Case Else
            ApplyConditionalFormat = "エラー: サポートされていない書式タイプです - " & formatType
            Exit Function
    End Select
    
    ' 書式スタイルの適用
    If formatRng.FormatConditions.Count > 0 Then
        Select Case LCase(formatStyle)
            Case "lightred", "薄い赤"
                With formatRng.FormatConditions(1).Interior
                    .Color = RGB(255, 199, 206)
                End With
            Case "lightgreen", "薄い緑"
                With formatRng.FormatConditions(1).Interior
                    .Color = RGB(198, 239, 206)
                End With
            Case "lightyellow", "薄い黄"
                With formatRng.FormatConditions(1).Interior
                    .Color = RGB(255, 235, 156)
                End With
            Case "bold", "太字"
                formatRng.FormatConditions(1).Font.Bold = True
            Case "italic", "斜体"
                formatRng.FormatConditions(1).Font.Italic = True
            Case "custom", "カスタム"
                ' カスタムスタイルはここでは適用しない
            Case "default", "デフォルト"
                ' デフォルトスタイルはそのまま
            Case Else
                ' 無効なスタイルの場合は警告
                Debug.Print "警告: 不明な書式スタイル - " & formatStyle
        End Select
    End If
    
    ApplyConditionalFormat = "条件付き書式を適用しました: [" & sheetName & "]" & dataRange & _
                           ", タイプ=" & formatType & ", 条件=" & condition
    Exit Function
    
ErrorHandler:
    ApplyConditionalFormat = "エラー: 条件付き書式の適用に失敗しました - " & Err.description
End Function

' 数式の挿入
Public Function InsertFormula(sheetName As String, cellReference As String, formula As String) As String
    On Error GoTo ErrorHandler
    
    ' シートの存在確認
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        InsertFormula = "エラー: シート '" & sheetName & "' が見つかりません"
        Exit Function
    End If
    
    ' 数式の先頭に=が含まれているか確認し、なければ追加
    Dim formulaText As String
    formulaText = Trim(formula)
    
    If left(formulaText, 1) <> "=" Then
        formulaText = "=" & formulaText
    End If
    
    ' 数式を挿入
    ws.Range(cellReference).formula = formulaText
    
    ' 結果を取得
    Dim result As Variant
    result = ws.Range(cellReference).value
    
    InsertFormula = "数式を挿入しました: [" & sheetName & "]" & cellReference & _
                   " に「" & formulaText & "」を挿入、結果=" & CStr(result)
    Exit Function
    
ErrorHandler:
    InsertFormula = "エラー: 数式の挿入に失敗しました - " & Err.description
End Function


' 複数のセル値を一度に取得 (JSON形式)
Public Function GetMultipleCellValues(sheetName As String, cellRanges As String) As String
    On Error GoTo ErrorHandler
    
    ' JSONオブジェクトを作成
    Dim resultObj As Object
    Set resultObj = CreateObject("Scripting.Dictionary")
    
    ' シートの存在確認
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        GetMultipleCellValues = "エラー: シート '" & sheetName & "' が見つかりません"
        Exit Function
    End If
    
    ' JSONとして渡された場合はパース
    If Left(cellRanges, 1) = "[" Or Left(cellRanges, 1) = "{" Then
        ' JSONとして解析
        Dim rangesCollection As Object
        
        If Left(cellRanges, 1) = "[" Then
            ' 配列形式: ["A1", "B2", "C3"]
            Set rangesCollection = JsonConverter.ParseJson(cellRanges)
            
            ' 各セル参照に対して値を取得
            Dim i As Long
            For i = 1 To rangesCollection.Count
                Dim cellRef As String
                cellRef = rangesCollection(i)
                
                On Error Resume Next
                Dim cellValue As Variant
                cellValue = ws.Range(cellRef).Value
                
                ' 結果に追加
                If Err.Number = 0 Then
                    resultObj.Add cellRef, CStr(cellValue)
                Else
                    resultObj.Add cellRef, "エラー: " & cellRef & " の取得に失敗"
                End If
                On Error GoTo ErrorHandler
            Next i
            
        ElseIf Left(cellRanges, 1) = "{" Then
            ' オブジェクト形式: {"name": "A1", "value": "B2"}
            Set rangesCollection = JsonConverter.ParseJson(cellRanges)
            
            ' 各キーに対して値を取得
            Dim key As Variant
            For Each key In rangesCollection.Keys
                cellRef = rangesCollection(key)
                
                On Error Resume Next
                cellValue = ws.Range(cellRef).Value
                
                ' 結果に追加
                If Err.Number = 0 Then
                    resultObj.Add CStr(key), CStr(cellValue)
                Else
                    resultObj.Add CStr(key), "エラー: " & cellRef & " の取得に失敗"
                End If
                On Error GoTo ErrorHandler
            Next key
        End If
    Else
        ' カンマ区切りの場合
        Dim rangesArray() As String
        rangesArray = Split(cellRanges, ",")
        
        ' 各セル参照に対して値を取得
        Dim j As Long
        For j = LBound(rangesArray) To UBound(rangesArray)
            cellRef = Trim(rangesArray(j))
            
            On Error Resume Next
            cellValue = ws.Range(cellRef).Value
            
            ' 結果に追加
            If Err.Number = 0 Then
                resultObj.Add cellRef, CStr(cellValue)
            Else
                resultObj.Add cellRef, "エラー: " & cellRef & " の取得に失敗"
            End If
            On Error GoTo ErrorHandler
        Next j
    End If
    
    ' 結果をJSON文字列に変換して返す
    GetMultipleCellValues = JsonConverter.ConvertToJson(resultObj)
    Exit Function
    
ErrorHandler:
    ' エラー発生時はエラーメッセージを返す
    GetMultipleCellValues = "{""error"": """ & Err.Description & """}"
End Function

' 複数のセル値を一度に設定 (JSON形式)
Public Function SetMultipleCellValues(sheetName As String, valuesJson As String) As String
    On Error GoTo ErrorHandler
    
    ' シートの存在確認
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        SetMultipleCellValues = "エラー: シート '" & sheetName & "' が見つかりません"
        Exit Function
    End If
    
    ' 結果オブジェクトを作成
    Dim resultObj As Object
    Set resultObj = CreateObject("Scripting.Dictionary")
    resultObj.Add "success", True
    resultObj.Add "updated", CreateObject("Scripting.Dictionary")
    resultObj.Add "errors", CreateObject("Scripting.Dictionary")
    
    ' JSONとして解析
    Dim valuesObj As Object
    Set valuesObj = JsonConverter.ParseJson(valuesJson)
    
    ' 各キーに対して値を設定
    Dim key As Variant
    Dim successCount As Long
    Dim errorCount As Long
    successCount = 0
    errorCount = 0
    
    For Each key In valuesObj.Keys
        Dim cellRef As String
        Dim newValue As String
        
        cellRef = CStr(key)
        newValue = CStr(valuesObj(key))
        
        On Error Resume Next
        ws.Range(cellRef).Value = newValue
        
        ' 結果を記録
        If Err.Number = 0 Then
            ' 成功
            successCount = successCount + 1
            resultObj("updated").Add cellRef, newValue
        Else
            ' 失敗
            errorCount = errorCount + 1
            resultObj("errors").Add cellRef, "エラー: " & Err.Description
        End If
        On Error GoTo ErrorHandler
    Next key
    
    ' 統計情報を追加
    resultObj.Add "summary", "合計: " & (successCount + errorCount) & " セル, 成功: " & successCount & ", 失敗: " & errorCount
    
    ' 結果をJSON文字列に変換して返す
    SetMultipleCellValues = JsonConverter.ConvertToJson(resultObj)
    Exit Function
    
ErrorHandler:
    ' エラー発生時はエラーメッセージを返す
    SetMultipleCellValues = "{""success"": false, ""error"": """ & Err.Description & """}"
End Function

' セルの書式設定（フォント、色、配置など）
Public Function FormatCells(sheetName As String, rangeAddress As String, Optional formatSettings As String = "") As String
    On Error GoTo ErrorHandler
    
    ' シートの存在確認
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        FormatCells = "エラー: シート '" & sheetName & "' が見つかりません"
        Exit Function
    End If
    
    ' 対象範囲
    Dim rng As Range
    Set rng = ws.Range(rangeAddress)
    
    ' 書式設定をJSONとして解析
    Dim formatObj As Object
    If formatSettings <> "" Then
        Set formatObj = JsonConverter.ParseJson(formatSettings)
    Else
        Set formatObj = CreateObject("Scripting.Dictionary")
    End If
    
    ' フォントの設定
    If formatObj.Exists("font") Then
        If formatObj("font").Exists("name") Then rng.Font.Name = formatObj("font")("name")
        If formatObj("font").Exists("size") Then rng.Font.Size = formatObj("font")("size")
        If formatObj("font").Exists("bold") Then rng.Font.Bold = formatObj("font")("bold")
        If formatObj("font").Exists("italic") Then rng.Font.Italic = formatObj("font")("italic")
        If formatObj("font").Exists("underline") Then
            Select Case LCase(formatObj("font")("underline"))
                Case "single", "true"
                    rng.Font.Underline = xlUnderlineStyleSingle
                Case "double"
                    rng.Font.Underline = xlUnderlineStyleDouble
                Case "none", "false"
                    rng.Font.Underline = xlUnderlineStyleNone
            End Select
        End If
        If formatObj("font").Exists("color") Then
            Dim fontColor As String
            fontColor = formatObj("font")("color")
            ' 色名またはRGB値で指定
            rng.Font.Color = ConvertColorValue(fontColor)
        End If
    End If
    
    ' 背景色
    If formatObj.Exists("fill") Then
        If formatObj("fill").Exists("color") Then
            Dim fillColor As String
            fillColor = formatObj("fill")("color")
            rng.Interior.Color = ConvertColorValue(fillColor)
        End If
        If formatObj("fill").Exists("pattern") Then
            Select Case LCase(formatObj("fill")("pattern"))
                Case "solid"
                    rng.Interior.Pattern = xlSolid
                Case "none", "transparent"
                    rng.Interior.Pattern = xlNone
                Case "lightgray", "lightgrey"
                    rng.Interior.Pattern = xlLightDown
                Case "gray", "grey"
                    rng.Interior.Pattern = xlGray16
                Case "darkgray", "darkgrey"
                    rng.Interior.Pattern = xlGray50
                Case "lightvertical"
                    rng.Interior.Pattern = xlLightVertical
                Case "lighthorizontal"
                    rng.Interior.Pattern = xlLightHorizontal
            End Select
        End If
    End If
    
    ' セルの配置
    If formatObj.Exists("alignment") Then
        ' 水平方向の配置
        If formatObj("alignment").Exists("horizontal") Then
            Select Case LCase(formatObj("alignment")("horizontal"))
                Case "left"
                    rng.HorizontalAlignment = xlLeft
                Case "center"
                    rng.HorizontalAlignment = xlCenter
                Case "right"
                    rng.HorizontalAlignment = xlRight
                Case "justify"
                    rng.HorizontalAlignment = xlJustify
                Case "distributed"
                    rng.HorizontalAlignment = xlDistributed
            End Select
        End If
        
        ' 垂直方向の配置
        If formatObj("alignment").Exists("vertical") Then
            Select Case LCase(formatObj("alignment")("vertical"))
                Case "top"
                    rng.VerticalAlignment = xlTop
                Case "center", "middle"
                    rng.VerticalAlignment = xlCenter
                Case "bottom"
                    rng.VerticalAlignment = xlBottom
                Case "justify"
                    rng.VerticalAlignment = xlJustify
                Case "distributed"
                    rng.VerticalAlignment = xlDistributed
            End Select
        End If
        
        ' 文字の折り返し
        If formatObj("alignment").Exists("wrapText") Then rng.WrapText = formatObj("alignment")("wrapText")
        
        ' 縮小して表示
        If formatObj("alignment").Exists("shrinkToFit") Then rng.ShrinkToFit = formatObj("alignment")("shrinkToFit")
        
        ' セルの結合
        If formatObj("alignment").Exists("merge") Then
            If formatObj("alignment")("merge") = True Then
                rng.Merge
            ElseIf formatObj("alignment")("merge") = False Then
                rng.UnMerge
            End If
        End If
        
        ' 文字の回転
        If formatObj("alignment").Exists("rotation") Then rng.Orientation = formatObj("alignment")("rotation")
    End If
    
    ' セル幅と行の高さ
    If formatObj.Exists("size") Then
        If formatObj("size").Exists("columnWidth") Then
            If IsNumeric(formatObj("size")("columnWidth")) Then
                rng.ColumnWidth = formatObj("size")("columnWidth")
            ElseIf formatObj("size")("columnWidth") = "autofit" Then
                rng.Columns.AutoFit
            End If
        End If
        
        If formatObj("size").Exists("rowHeight") Then
            If IsNumeric(formatObj("size")("rowHeight")) Then
                rng.RowHeight = formatObj("size")("rowHeight")
            ElseIf formatObj("size")("rowHeight") = "autofit" Then
                rng.Rows.AutoFit
            End If
        End If
    End If
    
    FormatCells = "セル [" & sheetName & "]" & rangeAddress & " の書式を設定しました"
    Exit Function
    
ErrorHandler:
    FormatCells = "エラー: セル書式の設定に失敗しました - " & Err.Description
End Function

' セルの罫線設定
Public Function FormatBorders(sheetName As String, rangeAddress As String, borderSettings As String) As String
    On Error GoTo ErrorHandler
    
    ' シートの存在確認
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        FormatBorders = "エラー: シート '" & sheetName & "' が見つかりません"
        Exit Function
    End If
    
    ' 対象範囲
    Dim rng As Range
    Set rng = ws.Range(rangeAddress)
    
    ' 罫線設定をJSONとして解析
    Dim borderObj As Object
    Set borderObj = JsonConverter.ParseJson(borderSettings)
    
    ' すべての罫線をクリア（オプション）
    If borderObj.Exists("clearAll") And borderObj("clearAll") = True Then
        rng.Borders.LineStyle = xlNone
    End If
    
    ' 罫線の太さを設定
    Dim lineWeight As XlBorderWeight
    lineWeight = xlThin ' デフォルト値
    
    If borderObj.Exists("weight") Then
        Select Case LCase(borderObj("weight"))
            Case "thin", "細い"
                lineWeight = xlThin
            Case "medium", "普通"
                lineWeight = xlMedium
            Case "thick", "太い"
                lineWeight = xlThick
            Case "hairline", "極細"
                lineWeight = xlHairline
        End Select
    End If
    
    ' typeパラメータがある場合はweightとして扱う
    If borderObj.Exists("type") Then
        Select Case LCase(borderObj("type"))
            Case "thin", "細い"
                lineWeight = xlThin
            Case "medium", "普通"
                lineWeight = xlMedium
            Case "thick", "太い"
                lineWeight = xlThick
            Case "hairline", "極細"
                lineWeight = xlHairline
        End Select
    End If
    
    ' 罫線の色を設定
    Dim borderColor As Long
    borderColor = RGB(0, 0, 0) ' デフォルト黒
    
    If borderObj.Exists("color") Then
        borderColor = ConvertColorValue(borderObj("color"))
    End If
    
    ' 罫線の種類を設定
    Dim lineStyle As XlLineStyle
    lineStyle = xlContinuous ' デフォルト実線
    
    If borderObj.Exists("style") Then
        Select Case LCase(borderObj("style"))
            Case "continuous", "solid", "実線"
                lineStyle = xlContinuous
            Case "dash", "dashed", "破線"
                lineStyle = xlDash
            Case "dot", "dotted", "点線"
                lineStyle = xlDot
            Case "dashdot", "破線・点線"
                lineStyle = xlDashDot
            Case "dashdotdot", "一点鎖線"
                lineStyle = xlDashDotDot
            Case "slantdashdot", "斜め破線"
                lineStyle = xlSlantDashDot
            Case "double", "二重線"
                lineStyle = xlDouble
            Case "none", "なし"
                lineStyle = xlNone
        End Select
    End If
    
    ' 罫線の位置を設定
    If borderObj.Exists("position") Then
        Select Case LCase(borderObj("position"))
            Case "all", "すべて"
                With rng.Borders
                    .LineStyle = lineStyle
                    .Color = borderColor
                    .Weight = lineWeight
                End With
            Case "outline", "外枠"
                With rng.BorderAround
                    .LineStyle = lineStyle
                    .Color = borderColor
                    .Weight = lineWeight
                End With
            Case "inside", "内側"
                With rng.Borders(xlInsideHorizontal)
                    .LineStyle = lineStyle
                    .Color = borderColor
                    .Weight = lineWeight
                End With
                With rng.Borders(xlInsideVertical)
                    .LineStyle = lineStyle
                    .Color = borderColor
                    .Weight = lineWeight
                End With
            Case "top", "上"
                With rng.Borders(xlEdgeTop)
                    .LineStyle = lineStyle
                    .Color = borderColor
                    .Weight = lineWeight
                End With
            Case "bottom", "下"
                With rng.Borders(xlEdgeBottom)
                    .LineStyle = lineStyle
                    .Color = borderColor
                    .Weight = lineWeight
                End With
            Case "left", "左"
                With rng.Borders(xlEdgeLeft)
                    .LineStyle = lineStyle
                    .Color = borderColor
                    .Weight = lineWeight
                End With
            Case "right", "右"
                With rng.Borders(xlEdgeRight)
                    .LineStyle = lineStyle
                    .Color = borderColor
                    .Weight = lineWeight
                End With
            Case "horizontalinside", "水平内側"
                With rng.Borders(xlInsideHorizontal)
                    .LineStyle = lineStyle
                    .Color = borderColor
                    .Weight = lineWeight
                End With
            Case "verticalinside", "垂直内側"
                With rng.Borders(xlInsideVertical)
                    .LineStyle = lineStyle
                    .Color = borderColor
                    .Weight = lineWeight
                End With
        End Select
    Else
        ' デフォルトは全ての罫線
        With rng.Borders
            .LineStyle = lineStyle
            .Color = borderColor
            .Weight = lineWeight
        End With
    End If
    
    FormatBorders = "セル [" & sheetName & "]" & rangeAddress & " の罫線を設定しました"
    Exit Function
    
ErrorHandler:
    FormatBorders = "エラー: 罫線の設定に失敗しました - " & Err.Description
End Function

' 数値書式の設定
Public Function FormatNumberStyle(sheetName As String, rangeAddress As String, formatCode As String) As String
    On Error GoTo ErrorHandler
    
    ' シートの存在確認
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        FormatNumberStyle = "エラー: シート '" & sheetName & "' が見つかりません"
        Exit Function
    End If
    
    ' 対象範囲
    Dim rng As Range
    Set rng = ws.Range(rangeAddress)
    
    ' 書式コードが直接指定されているか、一般的な名前で指定されているかをチェック
    Dim actualFormatCode As String
    
    Select Case LCase(formatCode)
        Case "general", "標準"
            actualFormatCode = "General"
        Case "number", "数値"
            actualFormatCode = "0.00"
        Case "currency", "通貨"
            actualFormatCode = "\#,##0.00"
        Case "accounting", "会計"
            actualFormatCode = "_\* #,##0.00_;_\* -#,##0.00_;_\* ""-""??_;_@_"
        Case "date", "日付"
            actualFormatCode = "yyyy/mm/dd"
        Case "time", "時刻"
            actualFormatCode = "h:mm:ss"
        Case "datetime", "日付時刻"
            actualFormatCode = "yyyy/mm/dd h:mm:ss"
        Case "percentage", "パーセント"
            actualFormatCode = "0.00%"
        Case "fraction", "分数"
            actualFormatCode = "# ?/?"
        Case "scientific", "指数"
            actualFormatCode = "0.00E+00"
        Case "text", "文字列"
            actualFormatCode = "@"
        Case "comma", "カンマ"
            actualFormatCode = "#,##0.00"
        Case "integer", "整数"
            actualFormatCode = "0"
        Case Else
            ' 直接書式コードとして使用
            actualFormatCode = formatCode
    End Select
    
    ' 書式を適用
    rng.NumberFormat = actualFormatCode
    
    FormatNumberStyle = "セル [" & sheetName & "]" & rangeAddress & " の数値書式を '" & actualFormatCode & "' に設定しました"
    Exit Function
    
ErrorHandler:
    FormatNumberStyle = "エラー: 数値書式の設定に失敗しました - " & Err.Description
End Function

' 色文字列をRGB値に変換するヘルパー関数
Private Function ConvertColorValue(colorValue As String) As Long
    ' RGB値が直接指定されている場合（例: "RGB(255,0,0)"）
    If Left(LCase(colorValue), 4) = "rgb(" Then
        Dim rgbParts As Variant
        rgbParts = Split(Mid(colorValue, 5, Len(colorValue) - 5), ",")
        
        If UBound(rgbParts) >= 2 Then
            Dim r As Integer, g As Integer, b As Integer
            r = CInt(Trim(rgbParts(0)))
            g = CInt(Trim(rgbParts(1)))
            b = CInt(Trim(rgbParts(2)))
            
            ConvertColorValue = RGB(r, g, b)
            Exit Function
        End If
    End If
    
    ' 色名で指定されている場合
    Select Case LCase(colorValue)
        Case "black", "黒"
            ConvertColorValue = RGB(0, 0, 0)
        Case "white", "白"
            ConvertColorValue = RGB(255, 255, 255)
        Case "red", "赤"
            ConvertColorValue = RGB(255, 0, 0)
        Case "green", "緑"
            ConvertColorValue = RGB(0, 128, 0)
        Case "blue", "青"
            ConvertColorValue = RGB(0, 0, 255)
        Case "yellow", "黄"
            ConvertColorValue = RGB(255, 255, 0)
        Case "magenta", "紫"
            ConvertColorValue = RGB(255, 0, 255)
        Case "cyan", "シアン"
            ConvertColorValue = RGB(0, 255, 255)
        Case "gray", "grey", "グレー"
            ConvertColorValue = RGB(128, 128, 128)
        Case "lightgray", "lightgrey", "薄いグレー"
            ConvertColorValue = RGB(192, 192, 192)
        Case "darkgray", "darkgrey", "濃いグレー"
            ConvertColorValue = RGB(64, 64, 64)
        Case "orange", "オレンジ"
            ConvertColorValue = RGB(255, 165, 0)
        Case "pink", "ピンク"
            ConvertColorValue = RGB(255, 192, 203)
        Case "lightblue", "薄い青"
            ConvertColorValue = RGB(173, 216, 230)
        Case "lightgreen", "薄い緑"
            ConvertColorValue = RGB(144, 238, 144)
        Case "lightyellow", "薄い黄"
            ConvertColorValue = RGB(255, 255, 224)
        Case "brown", "茶色"
            ConvertColorValue = RGB(165, 42, 42)
        Case "purple", "紫"
            ConvertColorValue = RGB(128, 0, 128)
        Case "navy", "紺"
            ConvertColorValue = RGB(0, 0, 128)
        Case "teal", "ティール"
            ConvertColorValue = RGB(0, 128, 128)
        Case "maroon", "栗色"
            ConvertColorValue = RGB(128, 0, 0)
        Case "olive", "オリーブ"
            ConvertColorValue = RGB(128, 128, 0)
        Case Else
            ' デフォルトは黒
            ConvertColorValue = RGB(0, 0, 0)
    End Select
End Function