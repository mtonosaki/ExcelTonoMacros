' Tono Excel Utilities (c) 2019 Manabu Tonosaki Licensed under the MIT license.

Attribute VB_Name = "Funcs"

' Get Color Index of a range
Function ColorIndex(r As Range) As Integer
    ColorIndex = r.Interior.ColorIndex
End Function

' Make BoxCode
' CODE = "Inch" 3digits
Function MakeBoxCode(Lmm As Double, Wmm As Double, Hmm As Double, Pref As String) As String
    If Lmm < Wmm Then
        st = Lmm
        lg = Wmm
    Else
        st = Wmm
        lg = Lmm
    End If
    st = Round(st / 25.4, 0)
    lg = Round(lg / 25.4, 0)
    ht = Round(Hmm / 25.4, 0)
    MakeBoxCode = Pref & Right("000" & lg, 3) & Right("000" & st, 3) & Right("000" & ht, 3)
End Function

' Copy specified text to clipboard
Public Function CopyTextToClipboard(s)
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Navigate "about:blank"
    Do While IE.Busy Or IE.Document.ReadyState <> "complete"
        DoEvents
    Loop
    'CopyTextToClipboard = IE.Document.ParentWindow.ClipboardData.GetData("text")
    IE.Document.ParentWindow.ClipboardData.SetData "text", s
    IE.Quit
    Set IE = Nothing
End Function

' Get text from Clipboard
Public Function GetTextFromClipboard()
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Navigate "about:blank"
    Do While IE.Busy Or IE.Document.ReadyState <> "complete"
        DoEvents
    Loop
    GetTextFromClipboard = IE.Document.ParentWindow.ClipboardData.GetData("text")
    IE.Quit
    Set IE = Nothing
End Function

' Count Strings
Public Function CountStr(str As String, check As String) As Integer
    CountStr = -1
    n = 0
    If check = "" Then Exit Function
    c = Left(check, 1)
    For i = 1 To Len(str)
        If Mid(str, i, 1) = c Then
            If Mid(str, i, Len(check)) = check Then
                n = n + 1
            End If
        End If
    Next
    CountStr = n
End Function

' Get Round up value divisible by 24
Public Function Roundup24(n As Double) As Integer
    If n <= 1 Then Roundup24 = 1: Exit Function
    If n <= 2 Then Roundup24 = 2: Exit Function
    If n <= 3 Then Roundup24 = 3: Exit Function
    If n <= 4 Then Roundup24 = 4: Exit Function
    If n <= 6 Then Roundup24 = 6: Exit Function
    If n <= 8 Then Roundup24 = 8: Exit Function
    If n <= 12 Then Roundup24 = 12: Exit Function
    If n <= 24 Then Roundup24 = 24: Exit Function
    If n > 24 Then Roundup24 = 24: Exit Function
End Function

' Get Round down value divisible by 24
Public Function Rounddown24(n As Double) As Integer
    If n >= 24 Then Rounddown24 = 24: Exit Function
    If n >= 12 Then Rounddown24 = 12: Exit Function
    If n >= 8 Then Rounddown24 = 8: Exit Function
    If n >= 6 Then Rounddown24 = 6: Exit Function
    If n >= 4 Then Rounddown24 = 4: Exit Function
    If n >= 3 Then Rounddown24 = 3: Exit Function
    If n >= 2 Then Rounddown24 = 2: Exit Function
    Rounddown24 = 1
End Function

'' Get Round up value divisible by 36
Public Function Roundup36(n As Double) As Integer
    If n <= 1 Then Roundup36 = 1: Exit Function
    If n <= 2 Then Roundup36 = 2: Exit Function
    If n <= 3 Then Roundup36 = 3: Exit Function
    If n <= 4 Then Roundup36 = 4: Exit Function
    If n <= 6 Then Roundup36 = 6: Exit Function
    If n <= 9 Then Roundup36 = 9: Exit Function
    If n <= 12 Then Roundup36 = 12: Exit Function
    If n <= 18 Then Roundup36 = 18: Exit Function
    Roundup36 = 36
End Function

' Get Round down value divisible by 36
Public Function Rounddown36(n As Double) As Integer
    If n >= 36 Then Rounddown36 = 36: Exit Function
    If n >= 18 Then Rounddown36 = 18: Exit Function
    If n >= 12 Then Rounddown36 = 12: Exit Function
    If n >= 9 Then Rounddown36 = 9: Exit Function
    If n >= 6 Then Rounddown36 = 6: Exit Function
    If n >= 4 Then Rounddown36 = 4: Exit Function
    If n >= 3 Then Rounddown36 = 3: Exit Function
    If n >= 2 Then Rounddown36 = 2: Exit Function
    Rounddown36 = 1
End Function

' Get Round down value divisible by 1,2, 3, 4,and 8
Public Function Rounddown12348(n As Double) As Integer
    If n >= 8 Then Rounddown12348 = 8: Exit Function
    If n >= 4 Then Rounddown12348 = 4: Exit Function
    If n >= 3 Then Rounddown12348 = 3: Exit Function
    If n >= 2 Then Rounddown12348 = 2: Exit Function
    Rounddown12348 = 1
End Function

' Select first cell of number to duplicate each lines the number times
Sub LineDuplication()
    Dim org As Range
    Set org = Selection
    Application.ScreenUpdating = False
    For lp = 1 To 65535
On Error GoTo VALERR
        If IsEmpty(org.Value) Or IsNull(org.Value) Or org.Value = "" Then End
        n = Int(org.Value)
On Error GoTo 0
        If n > 1 Then
            For i = 1 To n - 1
                Rows(org.Row).Select
                Selection.Copy
                Selection.Insert shift:=xlDown
            Next
        End If
        Set org = org.Offset(1, 0)
    Next
    End
VALERR:
    On Error GoTo 0
    Application.ScreenUpdating = False
    MsgBox org.Address & " is not a value of duplication"
End Sub


Sub LineDuplicationCsv()
    Dim org As Range
    Set org = Selection
    SkipN = 0
    Application.ScreenUpdating = False
    For lp = 1 To 65535
On Error GoTo VALERR
        If IsEmpty(org.Value) Or IsNull(org.Value) Or org.Value = "" Then
            SkipN = SkipN + 1
            If SkipN > 16 Then
                End
            Else
                GoTo SKIP_BLANK
            End If
        Else
            SkipN = 0
        End If
        Dim csv() As String
        ReDim csv(8) As String
        csvn = 0
        s = CStr(org.Value)
        sn = Len(s)
        iip = 1
        For ii = 1 To sn
            c = Mid(s, ii, 1)
            If c = "," Then
                csvn = csvn + 1
                If csvn > UBound(csv) Then
                    ReDim Preserve csv(csvn + 8) As String
                End If
                css = Mid(s, iip, ii - iip)
                csv(csvn) = Trim(css)
                iip = ii + 1
            End If
        Next
        csvn = csvn + 1
        If csvn > UBound(csv) Then
            ReDim Preserve csv(csvn + 1) As String
        End If
        css = Mid(s, iip, ii - iip)
        csv(csvn) = Trim(css)
        n = csvn
On Error GoTo 0
        If n > 1 Then
            For i = 1 To n - 1
                Rows(org.Row).Select
                Selection.Copy
                Selection.Insert shift:=xlDown
            Next
            pos = 0
            For i = n To 1 Step -1
                org.Offset(pos, 0).Value = csv(i)
                pos = pos - 1
            Next
        End If
SKIP_BLANK:
        Set org = org.Offset(1, 0)
    Next
    End
VALERR:
    On Error GoTo 0
    Application.ScreenUpdating = False
    MsgBox org.Address & " is not a value of duplication"
End Sub

