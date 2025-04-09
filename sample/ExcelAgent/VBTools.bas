
' �֐��̐���
Public Function GetProcedureInfo()
    Dim info As String
    info = "Excel����c�[���̎g�p���@:" & vbCrLf _
        & "- CreateWorksheet: �V�������[�N�V�[�g���쐬���܂�" & vbCrLf _
        & "  �g�p��: excel_function { ""functionName"": ""CreateWorksheet"", ""params"": [""�V�����V�[�g"", ""last""] }" & vbCrLf _
        & "- CreateChart: �O���t���쐬���܂�" & vbCrLf _
        & "  �g�p��: excel_function { ""functionName"": ""CreateChart"", ""params"": [""Sheet1"", ""A1:B10"", ""column"", ""����O���t""] }" & vbCrLf _
        & "- CreateTable: Excel�e�[�u�����쐬���܂�" & vbCrLf _
        & "  �g�p��: excel_function { ""functionName"": ""CreateTable"", ""params"": [""Sheet1"", ""A1:D10"", ""MyTable""] }" & vbCrLf _
        & "- CreatePivotTable: �s�{�b�g�e�[�u�����쐬���܂�" & vbCrLf _
        & "  �g�p��: excel_function { ""functionName"": ""CreatePivotTable"", ""params"": [""�f�[�^�V�[�g"", ""A1:D100"", ""�s�{�b�g�V�[�g"", ""A3"", ""����,�Ј�"", ""�N�x"", ""����""] }" & vbCrLf _
        & "- SortData: �f�[�^����בւ��܂�" & vbCrLf _
        & "  �g�p��: excel_function { ""functionName"": ""SortData"", ""params"": [""Sheet1"", ""A1:D10"", ""A"", ""ascending"", true] }" & vbCrLf _
        & "- ApplyFilter: �f�[�^�Ƀt�B���^�[��K�p���܂�" & vbCrLf _
        & "  �g�p��: excel_function { ""functionName"": ""ApplyFilter"", ""params"": [""Sheet1"", ""A1:D10"", ""A"", ""�l"", ""equals""] }" & vbCrLf _
        & "- ApplyConditionalFormat: �����t��������K�p���܂�" & vbCrLf _
        & "  �g�p��: excel_function { ""functionName"": ""ApplyConditionalFormat"", ""params"": [""Sheet1"", ""A1:D10"", ""cellvalue"", ""greater than,100"", ""lightred""] }" & vbCrLf _
        & "- InsertFormula: ������}�����܂�" & vbCrLf _
        & "  �g�p��: excel_function { ""functionName"": ""InsertFormula"", ""params"": [""Sheet1"", ""A1"", ""SUM(B1:B10)""] }" & vbCrLf _
        
        
    info = info & "- GetMultipleCellValues: �����̃Z���l����x�Ɏ擾���܂�" & vbCrLf _
        & "  �g�p��: excel_function { ""functionName"": ""GetMultipleCellValues"", ""params"": [""Sheet1"", ""A1,B2,C3""] }" & vbCrLf _
        & "  JSON�`��: excel_function { ""functionName"": ""GetMultipleCellValues"", ""params"": [""Sheet1"", ""[""A1"",""B2"",""C3""]""] }" & vbCrLf _
        & "  �J�X�^���L�[: excel_function { ""functionName"": ""GetMultipleCellValues"", ""params"": [""Sheet1"", ""{""""name"":""A1"",""value"":""B2""""}""] }" & vbCrLf _
        & "- SetMultipleCellValues: �����̃Z���l����x�ɐݒ肵�܂�" & vbCrLf _
        & "  �g�p��: excel_function { ""functionName"": ""SetMultipleCellValues"", ""params"": [""Sheet1"", ""{""""A1"":""�l1"",""B2"":""�l2""""}""] }" & vbCrLf

    info = info & "- FormatCells: �Z���̏�����ݒ肵�܂�" & vbCrLf _
        & "  �g�p��: excel_function { ""functionName"": ""FormatCells"", ""params"": [""Sheet1"", ""A1:B10"", ""{""font"":{""name"":""Arial"",...}}""] }" & vbCrLf _
        & "- FormatBorders: �Z���̌r����ݒ肵�܂�" & vbCrLf _
        & "  �g�p��: excel_function { ""functionName"": ""FormatBorders"", ""params"": [""Sheet1"", ""A1:D10"", ""{""position"":""outline"",...}""] }" & vbCrLf _
        & "- FormatNumberStyle: ���l������ݒ肵�܂�" & vbCrLf _
        & "  �g�p��: excel_function { ""functionName"": ""FormatNumberStyle"", ""params"": [""Sheet1"", ""A1:A10"", ""currency""] }" & vbCrLf

    GetProcedureInfo = info

End Function

' VBA�֐��Ăяo��
Public Function CallVBAFunction(inputJson As String) As String
    On Error GoTo ErrorHandler
    
    Debug.Print "VBA�֐��Ăяo��: " & inputJson
    
    ' JSON�����
    Dim jsonObject As Object
    Set jsonObject = JsonConverter.ParseJson(inputJson)
    
    ' �֐������擾
    Dim functionName As String
    If jsonObject.Exists("functionName") Then
        functionName = jsonObject("functionName")
    Else
        CallVBAFunction = "�G���[: �֐������w�肳��Ă��܂���"
        Exit Function
    End If
    
    ' �p�����[�^���擾
    Dim params As Collection
    If jsonObject.Exists("params") Then
        Set params = jsonObject("params")
    Else
        Set params = New Collection
    End If
    
    ' �֐����ɉ���������
    Select Case functionName
        Case "GetCellValue"
            ' �p�����[�^��: ["Sheet1", "A1"]
            If params.Count >= 2 Then
                CallVBAFunction = GetCellValue(params(1), params(2))
            Else
                CallVBAFunction = "�G���[: GetCellValue�ɂ�2�̃p�����[�^���K�v�ł�"
            End If
        
        Case "SetCellValue"
            ' �p�����[�^��: ["Sheet1", "A1", "�V�����l"]
            If params.Count >= 3 Then
                CallVBAFunction = SetCellValue(params(1), params(2), params(3))
            Else
                CallVBAFunction = "�G���[: SetCellValue�ɂ�3�̃p�����[�^���K�v�ł�"
            End If
            
        Case "CreateWorksheet"
            ' �p�����[�^��: ["�V�����V�[�g"]
            ' �܂��� ["�V�����V�[�g", "last"]
            If params.Count >= 1 Then
                If params.Count >= 2 Then
                    CallVBAFunction = CreateWorksheet(params(1), params(2))
                Else
                    CallVBAFunction = CreateWorksheet(params(1))
                End If
            Else
                CallVBAFunction = "�G���[: CreateWorksheet�ɂ͏��Ȃ��Ƃ�1�̃p�����[�^���K�v�ł�"
            End If
            
        Case "CreateChart"
            ' �p�����[�^��: ["Sheet1", "A1:B10", "column", "����O���t", 100, 100, 350, 250]
            If params.Count >= 3 Then
                ' �f�o�b�O�o�͂�ǉ�
                Debug.Print "CreateChart - �p�����[�^��: " & params.Count
                For i = 1 To params.Count
                    Debug.Print "params(" & i & "): " & params(i)
                Next i
                
                Dim title As String, left As Long, top As Long, width As Long, height As Long
                
                ' �V�[�g���A�͈́A�O���t��ނ�ϐ��Ɋi�[
                Dim sheetName As String, dataRange As String, chartType As String
                sheetName = params(1)
                dataRange = params(2)
                chartType = params(3)
                
                ' �I�v�V�����p�����[�^�����S�Ɏ擾
                title = ""
                left = 100
                top = 100 
                width = 350
                height = 250
                
                ' �p�����[�^�̐��ɉ����ăI�v�V����������ݒ�
                If params.Count >= 4 Then title = params(4)
                If params.Count >= 5 Then left = CLng(params(5))
                If params.Count >= 6 Then top = CLng(params(6))
                If params.Count >= 7 Then width = CLng(params(7))
                If params.Count >= 8 Then height = CLng(params(8))
                
                CallVBAFunction = CreateChart(sheetName, dataRange, chartType, title, left, top, width, height)
            Else
                CallVBAFunction = "�G���[: CreateChart�ɂ͏��Ȃ��Ƃ�3�̃p�����[�^���K�v�ł�"
            End If
            
        Case "CreateTable"
            ' �p�����[�^��: ["Sheet1", "A1:D10", "MyTable", "TableStyleMedium2"]
            If params.Count >= 2 Then
                Dim tableName As String, tableStyle As String
                
                ' �f�o�b�O�o�͂�ǉ�
                Debug.Print "CreateTable - �p�����[�^��: " & params.Count
                For i = 1 To params.Count
                    Debug.Print "params(" & i & "): " & params(i)
                Next i
                
                ' �C���F���S�Ƀp�����[�^���擾
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
                CallVBAFunction = "�G���[: CreateTable�ɂ͏��Ȃ��Ƃ�2�̃p�����[�^���K�v�ł�"
            End If
            
        Case "CreatePivotTable"
            ' �p�����[�^��: ["�f�[�^�V�[�g", "A1:D100", "�s�{�b�g�V�[�g", "A3", "����,�Ј�", "�N�x", "����"]
             If params.Count >= 4 Then
                Dim rowFields As String, columnFields As String, dataFields As String
                
                ' ���S�Ƀp�����[�^���擾
                rowFields = ""
                columnFields = ""
                dataFields = ""
                
                If params.Count >= 5 Then rowFields = params(5)
                If params.Count >= 6 Then columnFields = params(6)
                If params.Count >= 7 Then dataFields = params(7)
                
                CallVBAFunction = CreatePivotTable(params(1), params(2), params(3), params(4), _
                                                rowFields, columnFields, dataFields)
            Else
                CallVBAFunction = "�G���[: CreatePivotTable�ɂ͏��Ȃ��Ƃ�4�̃p�����[�^���K�v�ł�"
            End If

        Case "SortData"
            ' �p�����[�^��: ["Sheet1", "A1:D10", "A", "ascending", true]
             If params.Count >= 3 Then
                Dim sortOrder As String, hasHeader As Boolean
                
                ' ���S�Ƀp�����[�^���擾�iIIf�g�p���Ȃ��j
                sortOrder = "ascending"  ' �f�t�H���g�l
                hasHeader = True         ' �f�t�H���g�l
                
                If params.Count >= 4 Then
                    sortOrder = params(4)
                End If
                
                If params.Count >= 5 Then
                    hasHeader = CBool(params(5))
                End If
                
                CallVBAFunction = SortData(params(1), params(2), params(3), sortOrder, hasHeader)
            Else
                CallVBAFunction = "�G���[: SortData�ɂ͏��Ȃ��Ƃ�3�̃p�����[�^���K�v�ł�"
            End If

        Case "ApplyFilter"
            ' �p�����[�^��: ["Sheet1", "A1:D10", "A", "�l", "equals"]
            If params.Count >= 4 Then
                Dim operator As String
                
                ' ���S�Ƀp�����[�^���擾
                operator = "equals"  ' �f�t�H���g�l
                
                If params.Count >= 5 Then
                    operator = params(5)
                End If
                
                CallVBAFunction = ApplyFilter(params(1), params(2), params(3), params(4), operator)
            Else
                CallVBAFunction = "�G���[: ApplyFilter�ɂ͏��Ȃ��Ƃ�4�̃p�����[�^���K�v�ł�"
            End If

        Case "ApplyConditionalFormat"
            ' �p�����[�^��: ["Sheet1", "A1:D10", "cellvalue", "greater than,100", "lightred"]
            If params.Count >= 4 Then
                Dim formatStyle As String
                
                ' ���S�Ƀp�����[�^���擾
                formatStyle = "default"  ' �f�t�H���g�l
                
                If params.Count >= 5 Then
                    formatStyle = params(5)
                End If
                
                CallVBAFunction = ApplyConditionalFormat(params(1), params(2), params(3), params(4), formatStyle)
            Else
                CallVBAFunction = "�G���[: ApplyConditionalFormat�ɂ͏��Ȃ��Ƃ�4�̃p�����[�^���K�v�ł�"
            End If

        Case "InsertFormula"
            ' �p�����[�^��: ["Sheet1", "A1", "SUM(B1:B10)"]
            If params.Count >= 3 Then
                ' params(1) = �V�[�g��
                ' params(2) = �Z���Q��
                ' params(3) = ����
                CallVBAFunction = InsertFormula(params(1), params(2), params(3))
            Else
                CallVBAFunction = "�G���[: InsertFormula�ɂ�3�̃p�����[�^���K�v�ł�"
            End If

       
        Case "GetMultipleCellValues"
            ' �p�����[�^��: ["Sheet1", "A1,B2,C3"] �܂��� ["Sheet1", "[\"A1\",\"B2\",\"C3\"]"] �܂��� ["Sheet1", "{\"name\":\"A1\",\"value\":\"B2\"}"]
            If params.Count >= 2 Then
                CallVBAFunction = GetMultipleCellValues(params(1), params(2))
            Else
                CallVBAFunction = "�G���[: GetMultipleCellValues�ɂ͏��Ȃ��Ƃ�2�̃p�����[�^���K�v�ł�"
            End If

        Case "SetMultipleCellValues"
            ' �p�����[�^��: ["Sheet1", "{\"A1\":\"�V�����l1\",\"B2\":\"�V�����l2\"}"]
            If params.Count >= 2 Then
                CallVBAFunction = SetMultipleCellValues(params(1), params(2))
            Else
                CallVBAFunction = "�G���[: SetMultipleCellValues�ɂ͏��Ȃ��Ƃ�2�̃p�����[�^���K�v�ł�"
            End If

        Case "FormatCells"
            ' �p�����[�^��: ["Sheet1", "A1:B10", "{\"font\":{\"name\":\"Arial\",\"size\":12,\"bold\":true,\"color\":\"blue\"},\"fill\":{\"color\":\"lightyellow\"},\"alignment\":{\"horizontal\":\"center\"}}"]
            If params.Count >= 2 Then
                Dim formatSettings As String
                formatSettings = ""
                
                If params.Count >= 3 Then
                    formatSettings = params(3)
                End If
                
                CallVBAFunction = FormatCells(params(1), params(2), formatSettings)
            Else
                CallVBAFunction = "�G���[: FormatCells�ɂ͏��Ȃ��Ƃ�2�̃p�����[�^���K�v�ł�"
            End If

        Case "FormatBorders"
            ' �p�����[�^��: ["Sheet1", "A1:D10", "{\"position\":\"outline\",\"weight\":\"medium\",\"color\":\"blue\",\"style\":\"continuous\"}"]
            If params.Count >= 3 Then
                CallVBAFunction = FormatBorders(params(1), params(2), params(3))
            Else
                CallVBAFunction = "�G���[: FormatBorders�ɂ͏��Ȃ��Ƃ�3�̃p�����[�^���K�v�ł�"
            End If

        Case "FormatNumberStyle"
            ' �p�����[�^��: ["Sheet1", "A1:A10", "currency"] �܂��� ["Sheet1", "A1:A10", "#,##0.00"]
            If params.Count >= 3 Then
                CallVBAFunction = FormatNumberStyle(params(1), params(2), params(3))
            Else
                CallVBAFunction = "�G���[: FormatNumberStyle�ɂ͏��Ȃ��Ƃ�3�̃p�����[�^���K�v�ł�"
            End If
             
        Case Else
            CallVBAFunction = "�G���[: ���m�̊֐� '" & functionName & "'"
            
    End Select
    
    'Debug.Print "VBA�֐����s����: " & functionName & " - " & Left(CallVBAFunction, 50) & IIf(Len(CallVBAFunction) > 50, "...", "")
    Exit Function
    
ErrorHandler:
    CallVBAFunction = "�G���[: " & Err.description
    Debug.Print "VBA�֐����s�G���[: " & Err.description
End Function

' GetCellValue�֐�
Public Function GetCellValue(sheetName As String, cellReference As String) As String
    On Error GoTo ErrorHandler
    
    Dim result As String
    result = CStr(ThisWorkbook.Sheets(sheetName).Range(cellReference).value)
    GetCellValue = result
    Exit Function
    
ErrorHandler:
    GetCellValue = "�G���[: �Z���l�̎擾�Ɏ��s���܂��� - " & Err.description
End Function

' SetCellValue�֐�
Public Function SetCellValue(sheetName As String, cellReference As String, newValue As String) As String
    On Error GoTo ErrorHandler
    
    ThisWorkbook.Sheets(sheetName).Range(cellReference).value = newValue
    SetCellValue = "�Z�� " & sheetName & "!" & cellReference & " �ɒl """ & newValue & """ ��ݒ肵�܂���"
    Exit Function
    
ErrorHandler:
    SetCellValue = "�G���[: �Z���l�̐ݒ�Ɏ��s���܂��� - " & Err.description
End Function

' ���[�N�V�[�g�̍쐬
Public Function CreateWorksheet(sheetName As String, Optional position As String = "last") As String
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    
    ' �����̃V�[�g�����`�F�b�N
    Dim sheetExists As Boolean
    sheetExists = False
    
    For Each sheet In ThisWorkbook.Sheets
        If sheet.name = sheetName Then
            sheetExists = True
            Exit For
        End If
    Next sheet
    
    If sheetExists Then
        CreateWorksheet = "�V�[�g '" & sheetName & "' �͊��ɑ��݂��܂�"
        Exit Function
    End If
    
    ' �ʒu�p�����[�^�Ɋ�Â��ăV�[�g��ǉ�
    Select Case LCase(position)
        Case "first"
            Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        Case "last"
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        Case Else
            ' ���l�Ƃ��ĉ��߉\�ȏꍇ�́A���̈ʒu�ɑ}��
            On Error Resume Next
            Dim posNum As Integer
            posNum = CInt(position)
            
            If Err.Number = 0 And posNum > 0 And posNum <= ThisWorkbook.Worksheets.Count Then
                Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(posNum))
            Else
                ' �f�t�H���g�͍Ō�
                Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            End If
            On Error GoTo ErrorHandler
    End Select
    
    ws.name = sheetName
    
    CreateWorksheet = "�V�[�g '" & sheetName & "' ���쐬���܂���"
    Exit Function
    
ErrorHandler:
    CreateWorksheet = "�G���[: ���[�N�V�[�g�̍쐬�Ɏ��s���܂��� - " & Err.description
End Function

' �ȈՃO���t�̍쐬
Public Function CreateChart(sheetName As String, dataRange As String, chartType As String, _
                            Optional title As String = "", Optional left As Long = 100, _
                            Optional top As Long = 100, Optional width As Long = 350, _
                            Optional height As Long = 250) As String
    On Error GoTo ErrorHandler
    
    ' �V�[�g�̑��݊m�F
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        CreateChart = "�G���[: �V�[�g '" & sheetName & "' ��������܂���"
        Exit Function
    End If
    
    ' �`���[�g�^�C�v�̉��
    Dim chartTypeEnum As XlChartType
    Select Case LCase(chartType)
        Case "column", "�c�_", "�_�O���t"
            chartTypeEnum = xlColumnClustered
        Case "bar", "���_"
            chartTypeEnum = xlBarClustered
        Case "line", "�܂��"
            chartTypeEnum = xlLine
        Case "pie", "�~", "�~�O���t"
            chartTypeEnum = xlPie
        Case "area", "��", "�ʃO���t"
            chartTypeEnum = xlAreaStacked
        Case "scatter", "�U�z�}"
            chartTypeEnum = xlXYScatter
        Case Else
            chartTypeEnum = xlColumnClustered ' �f�t�H���g�͏c�_�O���t
    End Select
    
    ' �O���t�쐬
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(left, top, width, height)
    
    ' �f�[�^�͈͐ݒ�
    chartObj.Chart.SetSourceData Source:=ws.Range(dataRange)
    chartObj.Chart.chartType = chartTypeEnum
    
    ' �^�C�g���ݒ�
    If title <> "" Then
        chartObj.Chart.HasTitle = True
        chartObj.Chart.ChartTitle.text = title
    End If
    
    CreateChart = "�V�[�g '" & sheetName & "' �ɃO���t���쐬���܂���"
    Exit Function
    
ErrorHandler:
    CreateChart = "�G���[: �O���t�̍쐬�Ɏ��s���܂��� - " & Err.description
End Function

' Excel�e�[�u���̍쐬
Public Function CreateTable(sheetName As String, dataRange As String, _
                            Optional tableName As String = "", _
                            Optional tableStyle As String = "TableStyleMedium2") As String
    On Error GoTo ErrorHandler
    
    ' �V�[�g�̑��݊m�F
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        CreateTable = "�G���[: �V�[�g '" & sheetName & "' ��������܂���"
        Exit Function
    End If
    
    ' �e�[�u�������w�肳��Ă��Ȃ��ꍇ�͎�������
    If tableName = "" Then
        tableName = "Table" & (ThisWorkbook.TableStyles.Count + 1)
    End If
    
    ' �e�[�u���쐬
    Dim tableObj As ListObject
    Set tableObj = ws.ListObjects.Add(xlSrcRange, ws.Range(dataRange), , xlYes)
    
    ' �e�[�u�����Ə����̐ݒ�
    On Error Resume Next
    tableObj.name = tableName
    tableObj.tableStyle = tableStyle
    On Error GoTo ErrorHandler
    
    CreateTable = "�V�[�g '" & sheetName & "' �� '" & tableName & "' �e�[�u�����쐬���܂���"
    Exit Function
    
ErrorHandler:
    CreateTable = "�G���[: �e�[�u���̍쐬�Ɏ��s���܂��� - " & Err.description
End Function

' �s�{�b�g�e�[�u���̍쐬
Public Function CreatePivotTable(sourceSheet As String, sourceRange As String, _
                                destinationSheet As String, pivotLocation As String, _
                                Optional rowFields As String = "", _
                                Optional columnFields As String = "", _
                                Optional dataFields As String = "") As String
    On Error GoTo ErrorHandler
    
    ' �\�[�X�V�[�g�̑��݊m�F
    Dim srcWs As Worksheet
    On Error Resume Next
    Set srcWs = ThisWorkbook.Sheets(sourceSheet)
    On Error GoTo ErrorHandler
    
    If srcWs Is Nothing Then
        CreatePivotTable = "�G���[: �\�[�X�V�[�g '" & sourceSheet & "' ��������܂���"
        Exit Function
    End If
    
    ' �o�͐�V�[�g�̑��݊m�F
    Dim destWs As Worksheet
    On Error Resume Next
    Set destWs = ThisWorkbook.Sheets(destinationSheet)
    On Error GoTo ErrorHandler
    
    If destWs Is Nothing Then
        CreatePivotTable = "�G���[: �o�͐�V�[�g '" & destinationSheet & "' ��������܂���"
        Exit Function
    End If
    
    ' �s�{�b�g�L���b�V�����쐬
    Dim pvtCache As PivotCache
    Set pvtCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcWs.Range(sourceRange))
    
    ' �s�{�b�g�e�[�u�����쐬
    Dim pvtTable As PivotTable
    Set pvtTable = pvtCache.CreatePivotTable(TableDestination:=destWs.Range(pivotLocation))
    
    ' �s�t�B�[���h��ǉ�
    If rowFields <> "" Then
        Dim rowFieldsArray() As String
        rowFieldsArray = Split(rowFields, ",")
        
        Dim rowField As Variant
        For Each rowField In rowFieldsArray
            pvtTable.PivotFields(Trim(rowField)).Orientation = xlRowField
        Next rowField
    End If
    
    ' ��t�B�[���h��ǉ�
    If columnFields <> "" Then
        Dim colFieldsArray() As String
        colFieldsArray = Split(columnFields, ",")
        
        Dim colField As Variant
        For Each colField In colFieldsArray
            pvtTable.PivotFields(Trim(colField)).Orientation = xlColumnField
        Next colField
    End If
    
    ' �f�[�^�t�B�[���h��ǉ�
    If dataFields <> "" Then
        Dim dataFieldsArray() As String
        dataFieldsArray = Split(dataFields, ",")
        
        Dim dataField As Variant
        For Each dataField In dataFieldsArray
            Dim fieldName As String
            fieldName = Trim(dataField)
            
            ' �W�v���@�̎w�肪���邩�m�F�i��: "����:Sum"�j
            Dim summaryFunction As XlConsolidationFunction
            summaryFunction = xlSum  ' �f�t�H���g��Sum
            
            If InStr(fieldName, ":") > 0 Then
                Dim parts() As String
                parts = Split(fieldName, ":")
                fieldName = Trim(parts(0))
                
                Select Case LCase(Trim(parts(1)))
                    Case "sum", "���v"
                        summaryFunction = xlSum
                    Case "count", "��"
                        summaryFunction = xlCount
                    Case "average", "����"
                        summaryFunction = xlAverage
                    Case "max", "�ő�"
                        summaryFunction = xlMax
                    Case "min", "�ŏ�"
                        summaryFunction = xlMin
                End Select
            End If
            
            pvtTable.AddDataField pvtTable.PivotFields(fieldName), fieldName & " �� " & _
                Choose(summaryFunction, "���v", "��", "����", "�ő�", "�ŏ�"), summaryFunction
        Next dataField
    End If
    
    CreatePivotTable = "�s�{�b�g�e�[�u�����쐬���܂���: �\�[�X=[" & sourceSheet & "]" & sourceRange & _
                       ", �o�͐�=[" & destinationSheet & "]" & pivotLocation
    Exit Function
    
ErrorHandler:
    CreatePivotTable = "�G���[: �s�{�b�g�e�[�u���̍쐬�Ɏ��s���܂��� - " & Err.description
End Function

' �f�[�^�̕��בւ�
Public Function SortData(sheetName As String, dataRange As String, _
                         sortField As String, Optional sortOrder As String = "ascending", _
                         Optional hasHeader As Boolean = True) As String
    On Error GoTo ErrorHandler
    
    ' �V�[�g�̑��݊m�F
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        SortData = "�G���[: �V�[�g '" & sheetName & "' ��������܂���"
        Exit Function
    End If
    
    ' �\�[�g���̐ݒ�
    Dim orderValue As XlSortOrder
    Select Case LCase(sortOrder)
        Case "asc", "ascending", "����"
            orderValue = xlAscending
        Case "desc", "descending", "�~��"
            orderValue = xlDescending
        Case Else
            orderValue = xlAscending
    End Select
    
    ' �w�b�_�[�̐ݒ�
    Dim headerValue As XlYesNoGuess
    If hasHeader Then
        headerValue = xlYes
    Else
        headerValue = xlNo
    End If
    
    ' ���בւ����s
    ws.Range(dataRange).Sort _
        Key1:=ws.Range(sortField), _
        Order1:=orderValue, _
        Header:=headerValue
    
    SortData = "�f�[�^����בւ��܂���: [" & sheetName & "]" & dataRange & _
               ", �\�[�g�t�B�[���h=" & sortField & ", ����=" & sortOrder
    Exit Function
    
ErrorHandler:
    SortData = "�G���[: �f�[�^�̕��בւ��Ɏ��s���܂��� - " & Err.description
End Function

' �t�B���^�[�K�p
Public Function ApplyFilter(sheetName As String, dataRange As String, _
                            filterColumn As String, filterCriteria As String, _
                            Optional operator As String = "equals") As String
    On Error GoTo ErrorHandler
    
    ' �V�[�g�̑��݊m�F
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        ApplyFilter = "�G���[: �V�[�g '" & sheetName & "' ��������܂���"
        Exit Function
    End If
    
    ' �͈͂ɃI�[�g�t�B���^�[���Ȃ���ΓK�p
    Dim rng As Range
    Set rng = ws.Range(dataRange)
    
    If Not rng.AutoFilter Then
        rng.AutoFilter
    End If
    
    ' �t�B���^�[�����̓���
    Dim columnIndex As Integer
    
    ' �񖼂��w�肳�ꂽ�ꍇ�A��C���f�b�N�X���擾
    If IsNumeric(filterColumn) Then
        columnIndex = CInt(filterColumn)
    Else
        ' �񖼂����C���f�b�N�X���擾
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
            ApplyFilter = "�G���[: �� '" & filterColumn & "' ��������܂���"
            Exit Function
        End If
    End If
    
    ' ���Z�q�ݒ�
    Dim xlOperator As XlAutoFilterOperator
    Select Case LCase(operator)
        Case "equals", "equal", "eq", "="
            xlOperator = xlFilterValues
            ws.Range(dataRange).AutoFilter Field:=columnIndex, Criteria1:=filterCriteria
        Case "contains", "like", "�܂�"
            xlOperator = xlFilterValues
            ws.Range(dataRange).AutoFilter Field:=columnIndex, Criteria1:="*" & filterCriteria & "*"
        Case "greater", "greaterthan", "gt", ">", "���傫��"
            xlOperator = xlFilterValues
            ws.Range(dataRange).AutoFilter Field:=columnIndex, Criteria1:=">" & filterCriteria
        Case "less", "lessthan", "lt", "<", "��菬����"
            xlOperator = xlFilterValues
            ws.Range(dataRange).AutoFilter Field:=columnIndex, Criteria1:="<" & filterCriteria
        Case "beginswith", "startswith", "�n�܂�"
            xlOperator = xlFilterValues
            ws.Range(dataRange).AutoFilter Field:=columnIndex, Criteria1:=filterCriteria & "*"
        Case "endswith", "�I���"
            xlOperator = xlFilterValues
            ws.Range(dataRange).AutoFilter Field:=columnIndex, Criteria1:="*" & filterCriteria
        Case Else
            ' �f�t�H���g�͊��S��v
            xlOperator = xlFilterValues
            ws.Range(dataRange).AutoFilter Field:=columnIndex, Criteria1:=filterCriteria
    End Select
    
    ApplyFilter = "�t�B���^�[��K�p���܂���: [" & sheetName & "]" & dataRange & _
                 ", ��=" & filterColumn & ", ����=" & operator & " " & filterCriteria
    Exit Function
    
ErrorHandler:
    ApplyFilter = "�G���[: �t�B���^�[�K�p�Ɏ��s���܂��� - " & Err.description
End Function

' �����t�������̓K�p
Public Function ApplyConditionalFormat(sheetName As String, dataRange As String, _
                                     formatType As String, condition As String, _
                                     Optional formatStyle As String = "default") As String
    On Error GoTo ErrorHandler
    
    ' �V�[�g�̑��݊m�F
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        ApplyConditionalFormat = "�G���[: �V�[�g '" & sheetName & "' ��������܂���"
        Exit Function
    End If
    
    ' �����t�������̃^�C�v��ݒ�
    Dim formatRng As Range
    Set formatRng = ws.Range(dataRange)
    
    ' �������N���A
    formatRng.FormatConditions.Delete
    
    ' �����^�C�v�Ə����Ɋ�Â��ēK�p
    Select Case LCase(formatType)
        Case "cellvalue", "cellisvalue", "�Z���l"
            ' ���������
            Dim operator As XlFormatConditionOperator
            Dim value1 As String, value2 As String
            
            ' �����Ƃ��̒l�𕪊��i��: "greater than,100"�j
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
                    Case "between", "�����̊�"
                        operator = xlBetween
                        formatRng.FormatConditions.Add Type:=xlCellValue, operator:=operator, Formula1:=value1, Formula2:=value2
                    Case Else
                        ' �f�t�H���g�́u���̒l�ɓ������v
                        operator = xlEqual
                        formatRng.FormatConditions.Add Type:=xlCellValue, operator:=operator, Formula1:=value1
                End Select
            Else
                ApplyConditionalFormat = "�G���[: �����̌`�����s���ł� - " & condition
                Exit Function
            End If
            
        Case "colorscale", "�J���[�X�P�[��"
            ' 2�F�܂���3�F�̃J���[�X�P�[��
            Dim colorCount As Integer
            colorCount = CInt(condition)
            
            If colorCount = 2 Then
                formatRng.FormatConditions.AddColorScale ColorScaleType:=2
            ElseIf colorCount = 3 Then
                formatRng.FormatConditions.AddColorScale ColorScaleType:=3
            Else
                ' �f�t�H���g��2�F
                formatRng.FormatConditions.AddColorScale ColorScaleType:=2
            End If
            
        Case "databar", "�f�[�^�o�["
            ' �f�[�^�o�[�̐F���w��i��: "blue"�j
            Dim barColor As XlDataBarColor
            
            Select Case LCase(condition)
                Case "blue", "��"
                    barColor = xlDataBarColorBlue
                Case "green", "��"
                    barColor = xlDataBarColorGreen
                Case "red", "��"
                    barColor = xlDataBarColorRed
                Case "orange", "�I�����W"
                    barColor = xlDataBarColorOrange
                Case "lightblue", "lightblue", "���F"
                    barColor = xlDataBarColorLightBlue
                Case "purple", "��"
                    barColor = xlDataBarColorPurple
                Case Else
                    ' �f�t�H���g�͐�
                    barColor = xlDataBarColorBlue
            End Select
            
            formatRng.FormatConditions.AddDatabar
            formatRng.FormatConditions(1).ShowValue = True
            
        Case "iconset", "�A�C�R���Z�b�g"
            ' �A�C�R���Z�b�g�̎�ނ��w��i��: "3arrows"�j
            Dim iconSetType As XlIconSet
            
            Select Case LCase(condition)
                Case "3arrows", "3���"
                    iconSetType = xl3Arrows
                Case "3trafficlights", "3�M��"
                    iconSetType = xl3TrafficLights
                Case "3signs", "3�L��"
                    iconSetType = xl3Signs
                Case "3symbols", "3�V���{��"
                    iconSetType = xl3Symbols
                Case "4arrows", "4���"
                    iconSetType = xl4Arrows
                Case "4trafficlights", "4�M��"
                    iconSetType = xl4TrafficLights
                Case "5arrows", "5���"
                    iconSetType = xl5Arrows
                Case "5ratings", "5�]��"
                    iconSetType = xl5Ratings
                Case Else
                    ' �f�t�H���g��3���
                    iconSetType = xl3Arrows
            End Select
            
            formatRng.FormatConditions.AddIconSetCondition
            formatRng.FormatConditions(1).IconSet = ThisWorkbook.IconSets(iconSetType)
            
        Case "formula", "����"
            ' �����Ɋ�Â������t������
            formatRng.FormatConditions.Add Type:=xlExpression, Formula1:=condition
            
        Case Else
            ApplyConditionalFormat = "�G���[: �T�|�[�g����Ă��Ȃ������^�C�v�ł� - " & formatType
            Exit Function
    End Select
    
    ' �����X�^�C���̓K�p
    If formatRng.FormatConditions.Count > 0 Then
        Select Case LCase(formatStyle)
            Case "lightred", "������"
                With formatRng.FormatConditions(1).Interior
                    .Color = RGB(255, 199, 206)
                End With
            Case "lightgreen", "������"
                With formatRng.FormatConditions(1).Interior
                    .Color = RGB(198, 239, 206)
                End With
            Case "lightyellow", "������"
                With formatRng.FormatConditions(1).Interior
                    .Color = RGB(255, 235, 156)
                End With
            Case "bold", "����"
                formatRng.FormatConditions(1).Font.Bold = True
            Case "italic", "�Α�"
                formatRng.FormatConditions(1).Font.Italic = True
            Case "custom", "�J�X�^��"
                ' �J�X�^���X�^�C���͂����ł͓K�p���Ȃ�
            Case "default", "�f�t�H���g"
                ' �f�t�H���g�X�^�C���͂��̂܂�
            Case Else
                ' �����ȃX�^�C���̏ꍇ�͌x��
                Debug.Print "�x��: �s���ȏ����X�^�C�� - " & formatStyle
        End Select
    End If
    
    ApplyConditionalFormat = "�����t��������K�p���܂���: [" & sheetName & "]" & dataRange & _
                           ", �^�C�v=" & formatType & ", ����=" & condition
    Exit Function
    
ErrorHandler:
    ApplyConditionalFormat = "�G���[: �����t�������̓K�p�Ɏ��s���܂��� - " & Err.description
End Function

' �����̑}��
Public Function InsertFormula(sheetName As String, cellReference As String, formula As String) As String
    On Error GoTo ErrorHandler
    
    ' �V�[�g�̑��݊m�F
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        InsertFormula = "�G���[: �V�[�g '" & sheetName & "' ��������܂���"
        Exit Function
    End If
    
    ' �����̐擪��=���܂܂�Ă��邩�m�F���A�Ȃ���Βǉ�
    Dim formulaText As String
    formulaText = Trim(formula)
    
    If left(formulaText, 1) <> "=" Then
        formulaText = "=" & formulaText
    End If
    
    ' ������}��
    ws.Range(cellReference).formula = formulaText
    
    ' ���ʂ��擾
    Dim result As Variant
    result = ws.Range(cellReference).value
    
    InsertFormula = "������}�����܂���: [" & sheetName & "]" & cellReference & _
                   " �Ɂu" & formulaText & "�v��}���A����=" & CStr(result)
    Exit Function
    
ErrorHandler:
    InsertFormula = "�G���[: �����̑}���Ɏ��s���܂��� - " & Err.description
End Function


' �����̃Z���l����x�Ɏ擾 (JSON�`��)
Public Function GetMultipleCellValues(sheetName As String, cellRanges As String) As String
    On Error GoTo ErrorHandler
    
    ' JSON�I�u�W�F�N�g���쐬
    Dim resultObj As Object
    Set resultObj = CreateObject("Scripting.Dictionary")
    
    ' �V�[�g�̑��݊m�F
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        GetMultipleCellValues = "�G���[: �V�[�g '" & sheetName & "' ��������܂���"
        Exit Function
    End If
    
    ' JSON�Ƃ��ēn���ꂽ�ꍇ�̓p�[�X
    If Left(cellRanges, 1) = "[" Or Left(cellRanges, 1) = "{" Then
        ' JSON�Ƃ��ĉ��
        Dim rangesCollection As Object
        
        If Left(cellRanges, 1) = "[" Then
            ' �z��`��: ["A1", "B2", "C3"]
            Set rangesCollection = JsonConverter.ParseJson(cellRanges)
            
            ' �e�Z���Q�Ƃɑ΂��Ēl���擾
            Dim i As Long
            For i = 1 To rangesCollection.Count
                Dim cellRef As String
                cellRef = rangesCollection(i)
                
                On Error Resume Next
                Dim cellValue As Variant
                cellValue = ws.Range(cellRef).Value
                
                ' ���ʂɒǉ�
                If Err.Number = 0 Then
                    resultObj.Add cellRef, CStr(cellValue)
                Else
                    resultObj.Add cellRef, "�G���[: " & cellRef & " �̎擾�Ɏ��s"
                End If
                On Error GoTo ErrorHandler
            Next i
            
        ElseIf Left(cellRanges, 1) = "{" Then
            ' �I�u�W�F�N�g�`��: {"name": "A1", "value": "B2"}
            Set rangesCollection = JsonConverter.ParseJson(cellRanges)
            
            ' �e�L�[�ɑ΂��Ēl���擾
            Dim key As Variant
            For Each key In rangesCollection.Keys
                cellRef = rangesCollection(key)
                
                On Error Resume Next
                cellValue = ws.Range(cellRef).Value
                
                ' ���ʂɒǉ�
                If Err.Number = 0 Then
                    resultObj.Add CStr(key), CStr(cellValue)
                Else
                    resultObj.Add CStr(key), "�G���[: " & cellRef & " �̎擾�Ɏ��s"
                End If
                On Error GoTo ErrorHandler
            Next key
        End If
    Else
        ' �J���}��؂�̏ꍇ
        Dim rangesArray() As String
        rangesArray = Split(cellRanges, ",")
        
        ' �e�Z���Q�Ƃɑ΂��Ēl���擾
        Dim j As Long
        For j = LBound(rangesArray) To UBound(rangesArray)
            cellRef = Trim(rangesArray(j))
            
            On Error Resume Next
            cellValue = ws.Range(cellRef).Value
            
            ' ���ʂɒǉ�
            If Err.Number = 0 Then
                resultObj.Add cellRef, CStr(cellValue)
            Else
                resultObj.Add cellRef, "�G���[: " & cellRef & " �̎擾�Ɏ��s"
            End If
            On Error GoTo ErrorHandler
        Next j
    End If
    
    ' ���ʂ�JSON������ɕϊ����ĕԂ�
    GetMultipleCellValues = JsonConverter.ConvertToJson(resultObj)
    Exit Function
    
ErrorHandler:
    ' �G���[�������̓G���[���b�Z�[�W��Ԃ�
    GetMultipleCellValues = "{""error"": """ & Err.Description & """}"
End Function

' �����̃Z���l����x�ɐݒ� (JSON�`��)
Public Function SetMultipleCellValues(sheetName As String, valuesJson As String) As String
    On Error GoTo ErrorHandler
    
    ' �V�[�g�̑��݊m�F
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        SetMultipleCellValues = "�G���[: �V�[�g '" & sheetName & "' ��������܂���"
        Exit Function
    End If
    
    ' ���ʃI�u�W�F�N�g���쐬
    Dim resultObj As Object
    Set resultObj = CreateObject("Scripting.Dictionary")
    resultObj.Add "success", True
    resultObj.Add "updated", CreateObject("Scripting.Dictionary")
    resultObj.Add "errors", CreateObject("Scripting.Dictionary")
    
    ' JSON�Ƃ��ĉ��
    Dim valuesObj As Object
    Set valuesObj = JsonConverter.ParseJson(valuesJson)
    
    ' �e�L�[�ɑ΂��Ēl��ݒ�
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
        
        ' ���ʂ��L�^
        If Err.Number = 0 Then
            ' ����
            successCount = successCount + 1
            resultObj("updated").Add cellRef, newValue
        Else
            ' ���s
            errorCount = errorCount + 1
            resultObj("errors").Add cellRef, "�G���[: " & Err.Description
        End If
        On Error GoTo ErrorHandler
    Next key
    
    ' ���v����ǉ�
    resultObj.Add "summary", "���v: " & (successCount + errorCount) & " �Z��, ����: " & successCount & ", ���s: " & errorCount
    
    ' ���ʂ�JSON������ɕϊ����ĕԂ�
    SetMultipleCellValues = JsonConverter.ConvertToJson(resultObj)
    Exit Function
    
ErrorHandler:
    ' �G���[�������̓G���[���b�Z�[�W��Ԃ�
    SetMultipleCellValues = "{""success"": false, ""error"": """ & Err.Description & """}"
End Function

' �Z���̏����ݒ�i�t�H���g�A�F�A�z�u�Ȃǁj
Public Function FormatCells(sheetName As String, rangeAddress As String, Optional formatSettings As String = "") As String
    On Error GoTo ErrorHandler
    
    ' �V�[�g�̑��݊m�F
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        FormatCells = "�G���[: �V�[�g '" & sheetName & "' ��������܂���"
        Exit Function
    End If
    
    ' �Ώ۔͈�
    Dim rng As Range
    Set rng = ws.Range(rangeAddress)
    
    ' �����ݒ��JSON�Ƃ��ĉ��
    Dim formatObj As Object
    If formatSettings <> "" Then
        Set formatObj = JsonConverter.ParseJson(formatSettings)
    Else
        Set formatObj = CreateObject("Scripting.Dictionary")
    End If
    
    ' �t�H���g�̐ݒ�
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
            ' �F���܂���RGB�l�Ŏw��
            rng.Font.Color = ConvertColorValue(fontColor)
        End If
    End If
    
    ' �w�i�F
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
    
    ' �Z���̔z�u
    If formatObj.Exists("alignment") Then
        ' ���������̔z�u
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
        
        ' ���������̔z�u
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
        
        ' �����̐܂�Ԃ�
        If formatObj("alignment").Exists("wrapText") Then rng.WrapText = formatObj("alignment")("wrapText")
        
        ' �k�����ĕ\��
        If formatObj("alignment").Exists("shrinkToFit") Then rng.ShrinkToFit = formatObj("alignment")("shrinkToFit")
        
        ' �Z���̌���
        If formatObj("alignment").Exists("merge") Then
            If formatObj("alignment")("merge") = True Then
                rng.Merge
            ElseIf formatObj("alignment")("merge") = False Then
                rng.UnMerge
            End If
        End If
        
        ' �����̉�]
        If formatObj("alignment").Exists("rotation") Then rng.Orientation = formatObj("alignment")("rotation")
    End If
    
    ' �Z�����ƍs�̍���
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
    
    FormatCells = "�Z�� [" & sheetName & "]" & rangeAddress & " �̏�����ݒ肵�܂���"
    Exit Function
    
ErrorHandler:
    FormatCells = "�G���[: �Z�������̐ݒ�Ɏ��s���܂��� - " & Err.Description
End Function

' �Z���̌r���ݒ�
Public Function FormatBorders(sheetName As String, rangeAddress As String, borderSettings As String) As String
    On Error GoTo ErrorHandler
    
    ' �V�[�g�̑��݊m�F
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        FormatBorders = "�G���[: �V�[�g '" & sheetName & "' ��������܂���"
        Exit Function
    End If
    
    ' �Ώ۔͈�
    Dim rng As Range
    Set rng = ws.Range(rangeAddress)
    
    ' �r���ݒ��JSON�Ƃ��ĉ��
    Dim borderObj As Object
    Set borderObj = JsonConverter.ParseJson(borderSettings)
    
    ' ���ׂĂ̌r�����N���A�i�I�v�V�����j
    If borderObj.Exists("clearAll") And borderObj("clearAll") = True Then
        rng.Borders.LineStyle = xlNone
    End If
    
    ' �r���̑�����ݒ�
    Dim lineWeight As XlBorderWeight
    lineWeight = xlThin ' �f�t�H���g�l
    
    If borderObj.Exists("weight") Then
        Select Case LCase(borderObj("weight"))
            Case "thin", "�ׂ�"
                lineWeight = xlThin
            Case "medium", "����"
                lineWeight = xlMedium
            Case "thick", "����"
                lineWeight = xlThick
            Case "hairline", "�ɍ�"
                lineWeight = xlHairline
        End Select
    End If
    
    ' type�p�����[�^������ꍇ��weight�Ƃ��Ĉ���
    If borderObj.Exists("type") Then
        Select Case LCase(borderObj("type"))
            Case "thin", "�ׂ�"
                lineWeight = xlThin
            Case "medium", "����"
                lineWeight = xlMedium
            Case "thick", "����"
                lineWeight = xlThick
            Case "hairline", "�ɍ�"
                lineWeight = xlHairline
        End Select
    End If
    
    ' �r���̐F��ݒ�
    Dim borderColor As Long
    borderColor = RGB(0, 0, 0) ' �f�t�H���g��
    
    If borderObj.Exists("color") Then
        borderColor = ConvertColorValue(borderObj("color"))
    End If
    
    ' �r���̎�ނ�ݒ�
    Dim lineStyle As XlLineStyle
    lineStyle = xlContinuous ' �f�t�H���g����
    
    If borderObj.Exists("style") Then
        Select Case LCase(borderObj("style"))
            Case "continuous", "solid", "����"
                lineStyle = xlContinuous
            Case "dash", "dashed", "�j��"
                lineStyle = xlDash
            Case "dot", "dotted", "�_��"
                lineStyle = xlDot
            Case "dashdot", "�j���E�_��"
                lineStyle = xlDashDot
            Case "dashdotdot", "��_����"
                lineStyle = xlDashDotDot
            Case "slantdashdot", "�΂ߔj��"
                lineStyle = xlSlantDashDot
            Case "double", "��d��"
                lineStyle = xlDouble
            Case "none", "�Ȃ�"
                lineStyle = xlNone
        End Select
    End If
    
    ' �r���̈ʒu��ݒ�
    If borderObj.Exists("position") Then
        Select Case LCase(borderObj("position"))
            Case "all", "���ׂ�"
                With rng.Borders
                    .LineStyle = lineStyle
                    .Color = borderColor
                    .Weight = lineWeight
                End With
            Case "outline", "�O�g"
                With rng.BorderAround
                    .LineStyle = lineStyle
                    .Color = borderColor
                    .Weight = lineWeight
                End With
            Case "inside", "����"
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
            Case "top", "��"
                With rng.Borders(xlEdgeTop)
                    .LineStyle = lineStyle
                    .Color = borderColor
                    .Weight = lineWeight
                End With
            Case "bottom", "��"
                With rng.Borders(xlEdgeBottom)
                    .LineStyle = lineStyle
                    .Color = borderColor
                    .Weight = lineWeight
                End With
            Case "left", "��"
                With rng.Borders(xlEdgeLeft)
                    .LineStyle = lineStyle
                    .Color = borderColor
                    .Weight = lineWeight
                End With
            Case "right", "�E"
                With rng.Borders(xlEdgeRight)
                    .LineStyle = lineStyle
                    .Color = borderColor
                    .Weight = lineWeight
                End With
            Case "horizontalinside", "��������"
                With rng.Borders(xlInsideHorizontal)
                    .LineStyle = lineStyle
                    .Color = borderColor
                    .Weight = lineWeight
                End With
            Case "verticalinside", "��������"
                With rng.Borders(xlInsideVertical)
                    .LineStyle = lineStyle
                    .Color = borderColor
                    .Weight = lineWeight
                End With
        End Select
    Else
        ' �f�t�H���g�͑S�Ă̌r��
        With rng.Borders
            .LineStyle = lineStyle
            .Color = borderColor
            .Weight = lineWeight
        End With
    End If
    
    FormatBorders = "�Z�� [" & sheetName & "]" & rangeAddress & " �̌r����ݒ肵�܂���"
    Exit Function
    
ErrorHandler:
    FormatBorders = "�G���[: �r���̐ݒ�Ɏ��s���܂��� - " & Err.Description
End Function

' ���l�����̐ݒ�
Public Function FormatNumberStyle(sheetName As String, rangeAddress As String, formatCode As String) As String
    On Error GoTo ErrorHandler
    
    ' �V�[�g�̑��݊m�F
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        FormatNumberStyle = "�G���[: �V�[�g '" & sheetName & "' ��������܂���"
        Exit Function
    End If
    
    ' �Ώ۔͈�
    Dim rng As Range
    Set rng = ws.Range(rangeAddress)
    
    ' �����R�[�h�����ڎw�肳��Ă��邩�A��ʓI�Ȗ��O�Ŏw�肳��Ă��邩���`�F�b�N
    Dim actualFormatCode As String
    
    Select Case LCase(formatCode)
        Case "general", "�W��"
            actualFormatCode = "General"
        Case "number", "���l"
            actualFormatCode = "0.00"
        Case "currency", "�ʉ�"
            actualFormatCode = "\#,##0.00"
        Case "accounting", "��v"
            actualFormatCode = "_\* #,##0.00_;_\* -#,##0.00_;_\* ""-""??_;_@_"
        Case "date", "���t"
            actualFormatCode = "yyyy/mm/dd"
        Case "time", "����"
            actualFormatCode = "h:mm:ss"
        Case "datetime", "���t����"
            actualFormatCode = "yyyy/mm/dd h:mm:ss"
        Case "percentage", "�p�[�Z���g"
            actualFormatCode = "0.00%"
        Case "fraction", "����"
            actualFormatCode = "# ?/?"
        Case "scientific", "�w��"
            actualFormatCode = "0.00E+00"
        Case "text", "������"
            actualFormatCode = "@"
        Case "comma", "�J���}"
            actualFormatCode = "#,##0.00"
        Case "integer", "����"
            actualFormatCode = "0"
        Case Else
            ' ���ڏ����R�[�h�Ƃ��Ďg�p
            actualFormatCode = formatCode
    End Select
    
    ' ������K�p
    rng.NumberFormat = actualFormatCode
    
    FormatNumberStyle = "�Z�� [" & sheetName & "]" & rangeAddress & " �̐��l������ '" & actualFormatCode & "' �ɐݒ肵�܂���"
    Exit Function
    
ErrorHandler:
    FormatNumberStyle = "�G���[: ���l�����̐ݒ�Ɏ��s���܂��� - " & Err.Description
End Function

' �F�������RGB�l�ɕϊ�����w���p�[�֐�
Private Function ConvertColorValue(colorValue As String) As Long
    ' RGB�l�����ڎw�肳��Ă���ꍇ�i��: "RGB(255,0,0)"�j
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
    
    ' �F���Ŏw�肳��Ă���ꍇ
    Select Case LCase(colorValue)
        Case "black", "��"
            ConvertColorValue = RGB(0, 0, 0)
        Case "white", "��"
            ConvertColorValue = RGB(255, 255, 255)
        Case "red", "��"
            ConvertColorValue = RGB(255, 0, 0)
        Case "green", "��"
            ConvertColorValue = RGB(0, 128, 0)
        Case "blue", "��"
            ConvertColorValue = RGB(0, 0, 255)
        Case "yellow", "��"
            ConvertColorValue = RGB(255, 255, 0)
        Case "magenta", "��"
            ConvertColorValue = RGB(255, 0, 255)
        Case "cyan", "�V�A��"
            ConvertColorValue = RGB(0, 255, 255)
        Case "gray", "grey", "�O���["
            ConvertColorValue = RGB(128, 128, 128)
        Case "lightgray", "lightgrey", "�����O���["
            ConvertColorValue = RGB(192, 192, 192)
        Case "darkgray", "darkgrey", "�Z���O���["
            ConvertColorValue = RGB(64, 64, 64)
        Case "orange", "�I�����W"
            ConvertColorValue = RGB(255, 165, 0)
        Case "pink", "�s���N"
            ConvertColorValue = RGB(255, 192, 203)
        Case "lightblue", "������"
            ConvertColorValue = RGB(173, 216, 230)
        Case "lightgreen", "������"
            ConvertColorValue = RGB(144, 238, 144)
        Case "lightyellow", "������"
            ConvertColorValue = RGB(255, 255, 224)
        Case "brown", "���F"
            ConvertColorValue = RGB(165, 42, 42)
        Case "purple", "��"
            ConvertColorValue = RGB(128, 0, 128)
        Case "navy", "��"
            ConvertColorValue = RGB(0, 0, 128)
        Case "teal", "�e�B�[��"
            ConvertColorValue = RGB(0, 128, 128)
        Case "maroon", "�I�F"
            ConvertColorValue = RGB(128, 0, 0)
        Case "olive", "�I���[�u"
            ConvertColorValue = RGB(128, 128, 0)
        Case Else
            ' �f�t�H���g�͍�
            ConvertColorValue = RGB(0, 0, 0)
    End Select
End Function