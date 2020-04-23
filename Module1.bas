' Tono Excel Utilities (c) 2019 Manabu Tonosaki Licensed under the MIT license.

Attribute VB_Name = "Module1"
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
    If r.Rows.Count = 1 And r.Columns.Count >= 255 Then
        r.AutoFilter
        End
    End If
NOR:
    On Error GoTo 0
    Cells.AutoFilter
End Sub
