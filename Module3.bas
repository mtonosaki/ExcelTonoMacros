' Tono Excel Utilities (c) 2019 Manabu Tonosaki Licensed under the MIT license.

Attribute VB_Name = "Module3"
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
