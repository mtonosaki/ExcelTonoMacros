' Tono Excel Utilities (c) 2020 Manabu Tonosaki Licensed under the MIT license.

Sub MacroCSV()
'
' MacroCSV Macro
' Value Paste
'
' Keyboard Shortcut: Ctrl+Shift+V
'
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub

Sub MacroCSF()
'
' MacroCSF Macro
' Auto filter
'
' Keyboard Shortcut: Ctrl+Shift+F
'
    On Error GoTo NOR
    Dim r As Range
    Set r = Selection
    If r.Rows.Count = 1 And r.Columns.Count > 1 Then
        r.AutoFilter
        GoTo NOR2
    End If
NOR:
    On Error GoTo NOR2
    RMAX = 24
    Dim re As Range
    For i = 1 To RMAX
        Set r = Range("A1").Offset(i - 1, 0)
        For st = 0 To 25
            If r.Offset(0, st).Value <> "" Then
                Exit For
            End If
        Next
        blankN = 0
        For ed = st To st + 26
            If r.Offset(0, ed).Value = "" Then
                blankN = blankN + 1
                If blankN > 4 Then
                    ed = ed - blankN
                    Exit For
                End If
            Else
                blankN = 0
            End If
        Next
        Set re0 = Range(r.Offset(0, st).Address, r.Offset(0, ed).Address)
        If re Is Nothing Then  ' find the first longest row
            Set re = re0
        End If
        If re0.Cells.Count > re.Cells.Count Then
            Set re = re0
        End If
    Next
    If re Is Nothing Then
        MsgBox "Cannot find table header row", vbOKOnly, "WARNING"
    Else
        re.AutoFilter
    End If

NOR2:
    On Error GoTo 0
End Sub

Sub MacroCST()
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
'
    ' MacroCSW Macro
    ' Set/Reset Freeze window position
    '
    ' Keyboard Shortcut: Ctrl+Shift+W
    '
    On Error GoTo ER
    ActiveWindow.FreezePanes = Not ActiveWindow.FreezePanes
    GoTo FIN
ER:
    MsgBox "Cannot get/set FreezePanes property. Requested 'Enable Editing' mode.", vbInformation Or vbOKOnly, "ERROR"
    On Error GoTo 0
FIN:

End Sub

Sub MacroCSM()
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
    '
    ' MacroCSH Macro
    ' Change format to yyyy/mm/dd HH:MM:SS
    '
    ' Keyboard Shortcut: Ctrl+Shift+H
    '
    
    Static PrevCshSelection As Range
    Static CshStoryCounter As Integer
    
    On Error GoTo ER0
    
    If IsSameRange(PrevCshSelection, Selection) = False Then
        CshStoryCounter = 0
    Else
        CshStoryCounter = CshStoryCounter + 1
    End If
    
    Select Case CshStoryCounter Mod 3
        Case 0
            Selection.NumberFormat = "yyyy/mm/dd hh:mm:ss"
        Case 1
            Selection.NumberFormat = "yyyy/mm/dd hh:mm"
        Case 2
            Selection.NumberFormat = "yyyy/mm/dd"
    End Select
    Selection.EntireColumn.AutoFit
    Set PrevCshSelection = Selection
ER0:
    On Error GoTo 0
End Sub

Private Function IsSameRange(a As Range, b As Range) As Boolean
    If a Is Nothing And b Is Nothing Then
        IsSameRange = True
        Exit Function
    End If
    If a Is Nothing Or b Is Nothing Then
        IsSameRange = False
        Exit Function
    End If
    If a.Address = b.Address Then
        If a.Rows.Count = b.Rows.Count Then
            If a.Columns.Count = b.Columns.Count Then
                IsSameRange = True
                Exit Function
            End If
        End If
    End If
    IsSameRange = False
End Function
