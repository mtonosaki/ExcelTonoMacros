Attribute VB_Name = "KeyMacros"
' Tono Excel Utilities (c) 2020 Manabu Tonosaki Licensed under the MIT license.

Sub MacroCSV()
Attribute MacroCSV.VB_Description = "Value Paste"
Attribute MacroCSV.VB_ProcData.VB_Invoke_Func = "V\n14"
'
' MacroCSV Macro
' Value Paste
'
' Keyboard Shortcut: Ctrl+Shift+V
'
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub

Sub MacroCSF()
Attribute MacroCSF.VB_Description = "Auto filter"
Attribute MacroCSF.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' MacroCSF Macro
' Auto filter
'
' Keyboard Shortcut: Ctrl+Shift+F
'
    On Error GoTo NOR
    Dim r As Range
    Set r = Selection
    If r.Rows.Count = 1 And r.Columns.Count >= 4 Then
        r.AutoFilter
        End
    End If
NOR:
    On Error GoTo NOR2
    Dim re As Range
    Set re = r.Rows(1).Columns(1).EntireRow
    Dim re1 As Range
    Set re1 = re.Rows(1).Columns(1)
    If re1 = "" Then
        re1 = "@@dummy__@@"
    End If
    re.AutoFilter
    If re1 = "@@dummy__@@" Then
        re1 = ""
    End If
NOR2:
    On Error GoTo 0
End Sub

Sub MacroCST()
Attribute MacroCST.VB_Description = "Formatting"
Attribute MacroCST.VB_ProcData.VB_Invoke_Func = "T\n14"
'
' MacroCST Macro
' Formatting
'
' Keyboard Shortcut: Ctrl+Shift+T
'
    Set nowcell = Selection
    Rows("1:1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .WrapText = False
        .Orientation = 90
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    Cells.Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .WrapText = False
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    Rows("1:1").Select
    Selection.AutoFilter
    Selection.HorizontalAlignment = xlLeft
    Selection.VerticalAlignment = xlTop
    Rows("1:1").EntireRow.AutoFit
    Cells.Select
    Cells.EntireColumn.AutoFit
    nowcell.Select
End Sub

Sub MacroCSW()
Attribute MacroCSW.VB_Description = "Set/Reset Freeze window position"
Attribute MacroCSW.VB_ProcData.VB_Invoke_Func = "W\n14"
'
    ' MacroCSW Macro
    ' Set/Reset Freeze window position
    '
    ' Keyboard Shortcut: Ctrl+Shift+W
    '
    ActiveWindow.FreezePanes = Not ActiveWindow.FreezePanes
End Sub

Sub MacroCSM()
Attribute MacroCSM.VB_Description = "Merge Cells"
Attribute MacroCSM.VB_ProcData.VB_Invoke_Func = "M\n14"
'
    ' MacroCSM Macro
    ' Merge Cells
    '
    ' Keyboard Shortcut: Ctrl+Shift+M
    '

    On Error GoTo ERM
    If Selection.MergeCells Then
        Selection.MergeCells = False
    Else
        Selection.MergeCells = True
    End If
ERM:
    On Error GoTo 0
End Sub

Sub MacroCSH()
Attribute MacroCSH.VB_Description = "Change format to yyyy/mm/dd HH:MM:SS"
Attribute MacroCSH.VB_ProcData.VB_Invoke_Func = "H\n14"
'
    ' MacroCSH Macro
    ' Change format to yyyy/mm/dd HH:MM:SS
    '
    ' Keyboard Shortcut: Ctrl+Shift+H
    '
    Selection.NumberFormat = "yyyy/mm/dd hh:mm:ss"
    Selection.EntireColumn.AutoFit
End Sub
