Attribute VB_Name = "Módulo1"
Sub Geral()

Application.ScreenUpdating = False

    'Tipo Var
    Dim valor As String
    
    valor = MsgBox("Processar todos os dados?", vbOKCancel, "VALIDAÇÃO DE ATIVAÇÃO DE MACROS")
     
    If valor = 1 Then
  
    Sheets("TDs").Select
    Range("D9:E9").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("C8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("D8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D7").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("E8").Select
    Selection.Copy
    Range("D8").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("N9:O9").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("M8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("N8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("N7").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("O8").Select
    Selection.Copy
    Range("N8").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("X9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("W8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("X8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("X7").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("Y8").Select
    Selection.Copy
    Range("X8").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("D8:E8").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlDown)).Select
    Range("N8:O8").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlDown)).Select
    Range("X8:Y8").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlDown)).Select
    Range("B8").Select
  
    Sheets("ID COBRANÇA").Select
    Range("B3:D3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("E4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("B3").Select
      
    Call Base_Limpa
    Call Base_Tratada_01
    Call Filtro_Chave
    Call ID_Cobranca
    Call Base_Tratada_02
    Call Base_Final
    
    Else
    End If
    
    Sheets("MACROS").Select
    Range("B7").Select

Application.ScreenUpdating = True

End Sub


Sub Base_Limpa()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim linhai As Double
    Dim linhaf As Double

    Sheets("BASE LIMPA").Select
    Range("B4").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C3").Value > 0 Then
        linhaf = linhai - Range("C3").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C3").Value < 0 Then
        linhaf = linhai + Range("C3").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B5").Select
    Sheets("BASE INICIAL").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B7").Select
    Sheets("BASE LIMPA").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B5").Select
    Range("AO2").Select
    Selection.Copy
    Range("AO5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B5").Select
    
    Application.ScreenUpdating = True

End Sub

Sub Base_Tratada_01()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim linhai As Double
    Dim linhaf As Double

    Sheets("BASE TRATADA").Select
    Range("B4").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C3").Value > 0 Then
        linhaf = linhai - Range("C3").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C3").Value < 0 Then
        linhaf = linhai + Range("C3").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B5").Select
    Sheets("BASE LIMPA").Select
    Range("AO4").Select
    ActiveSheet.Range("$B$4:$AO$350000").AutoFilter Field:=40, Criteria1:="=1" _
        , Operator:=xlAnd
    Range("B4:AN4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("BASE TRATADA").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("BASE LIMPA").Select
    Range("B4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Selection.AutoFilter
    Range("B5").Select
    Sheets("BASE TRATADA").Select
    Range("B5").Select
    Range("AO5").Select
    Selection.Copy
    Range("AO6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B5").Select
    
    Application.ScreenUpdating = True

End Sub

Sub Filtro_Chave()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim linhai As Double
    Dim linhaf As Double

    Sheets("FILTRO CHAVE").Select
    Range("B4").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C3").Value > 0 Then
        linhaf = linhai - Range("C3").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C3").Value < 0 Then
        linhaf = linhai + Range("C3").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B5").Select
    Sheets("BASE TRATADA").Select
    Range("E4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("FILTRO CHAVE").Select
    Range("B4").Select
    ActiveSheet.Paste
    Sheets("BASE TRATADA").Select
    Range("I4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("FILTRO CHAVE").Select
    Range("C4").Select
    ActiveSheet.Paste
    Sheets("BASE TRATADA").Select
    Range("P4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("FILTRO CHAVE").Select
    Range("D4").Select
    ActiveSheet.Paste
    Sheets("BASE TRATADA").Select
    Range("W4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("FILTRO CHAVE").Select
    Range("E4").Select
    ActiveSheet.Paste
    Sheets("BASE TRATADA").Select
    Columns("Z:Z").Select
    Range("Z4").Activate
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("FILTRO CHAVE").Select
    Columns("F:F").Select
    Range("F4").Activate
    ActiveSheet.Paste
    Sheets("BASE TRATADA").Select
    Range("AO4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("FILTRO CHAVE").Select
    Range("G4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B4").Select
    Sheets("BASE TRATADA").Select
    Application.CutCopyMode = False
    Range("B5").Select
    Sheets("FILTRO CHAVE").Select
    Range("B5").Select
    Range("E4").Select
    ActiveWorkbook.Worksheets("FILTRO CHAVE").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("FILTRO CHAVE").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("E4:E350000"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("FILTRO CHAVE").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("F4").Select
    ActiveWorkbook.Worksheets("FILTRO CHAVE").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("FILTRO CHAVE").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("F4:F350000"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("FILTRO CHAVE").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.Range("$B$4:$H$350000").RemoveDuplicates Columns:=Array(1, 2, 3), _
        Header:=xlYes
    ActiveWorkbook.RefreshAll
    Range("B5").Select
    Range("H3").Select
    Selection.Copy
    Range("H5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B5").Select

    Application.ScreenUpdating = True

End Sub

Sub ID_Cobranca()

Application.ScreenUpdating = False

    Sheets("FILTRO CHAVE").Select
    ActiveSheet.Range("$B$4:$H$300000").AutoFilter Field:=7, Criteria1:="<>1", _
        Operator:=xlAnd
    Range("B86").Select
    Sheets("FILTRO CHAVE").Select
    Range("G4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("ID COBRANÇA").Select
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("FILTRO CHAVE").Select
    Range("D4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("ID COBRANÇA").Select
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("FILTRO CHAVE").Select
    Range("H4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("ID COBRANÇA").Select
    Range("D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("D3").Select
    Sheets("FILTRO CHAVE").Select
    Range("H4").Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Selection.AutoFilter
    Range("B5").Select
    Sheets("ID COBRANÇA").Select
    Range("E3").Select
    Selection.Copy
    Range("D2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("B3").Select
    ActiveSheet.Range("$B$2:$E$50000").RemoveDuplicates Columns:=1, Header:= _
        xlYes

Application.ScreenUpdating = True

End Sub

Sub Base_Tratada_02()

Application.ScreenUpdating = False

    Sheets("BASE TRATADA").Select
    Range("AP5:AR5").Select
    Selection.Copy
    Range("AP6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B5").Select

Application.ScreenUpdating = True

End Sub

Sub Base_Final()

Application.ScreenUpdating = False

    'Tipo Var
    Dim linhai As Double
    Dim linhaf As Double

    Sheets("BASE FINAL").Select
    Range("B4").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C3").Value > 0 Then
        linhaf = linhai - Range("C3").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C3").Value < 0 Then
        linhaf = linhai + Range("C3").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B5").Select
    Sheets("BASE TRATADA").Select
    Range("AR4").Select
    ActiveSheet.Range("$B$4:$AS$350000").AutoFilter Field:=43, Criteria1:= _
        "=Não", Operator:=xlAnd
    Range("AR4").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("BASE FINAL").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B5").Select
    Sheets("BASE TRATADA").Select
    Application.CutCopyMode = False
    Range("AS4").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Selection.AutoFilter
    Selection.AutoFilter
    Range("B5").Select
    Sheets("BASE FINAL").Select
    Range("B5").Select
    Range("AS5:AY5").Select
    Selection.Copy
    Range("AS6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B5").Select
    Columns("AB:AJ").Select
    Range("AB4").Activate
    Selection.Replace What:=".", Replacement:=",", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("B5").Select
    Range("AB5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("AB5"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="-", FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Range("B5").Select
    
    ActiveWorkbook.RefreshAll
    
Application.ScreenUpdating = True

End Sub

Sub Arquivo_Leitura()

Application.ScreenUpdating = False
    'Tipo Var
    Dim valor As String
    
    valor = MsgBox("Gerar arquivo final processado?", vbOKCancel, "VALIDAÇÃO DE ATIVAÇÃO DE MACROS")
     
    If valor = 1 Then
    
    ActiveWorkbook.Save
    ChDir _
        ActiveWorkbook.Path
    ActiveWorkbook.SaveAs Filename:= _
        ActiveWorkbook.Path & "\Taxas ADQ - Tendência MS - " & Worksheets("MACROS").Range("C10").Value & ".xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    Sheets(Array("MACROS", "BASE INICIAL", "BASE LIMPA", "BASE TRATADA", "FILTRO CHAVE" _
        , "TD CHAVE", "ID COBRANÇA")).Select
    Sheets("ID COBRANÇA").Activate
    ActiveWindow.SelectedSheets.Delete
    Range("C2").Select
    Selection.Copy
    Range("C3").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("AS2").Select
    Selection.Copy
    Range("AS3:AX3").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("B5").Select
    ActiveWindow.DisplayHeadings = False
    Sheets("TDs").Select
    ActiveWindow.DisplayHeadings = False
    Sheets("BASE FINAL").Select
    ActiveWorkbook.Save

    Else
    End If

Application.ScreenUpdating = True

End Sub
