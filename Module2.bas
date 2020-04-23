' Tono Excel Utilities (c) 2019 Manabu Tonosaki Licensed under the MIT license.

Attribute VB_Name = "Module2"
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
