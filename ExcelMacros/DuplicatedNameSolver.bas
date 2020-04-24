Attribute VB_Name = "DuplicatedNameSolver"
' Tono Excel Utilities (c) 2020 Manabu Tonosaki Licensed under the MIT license.

' key cache
Public keys As Scripting.Dictionary

' !! Execute this method in "Workbook_SheetCalculate" event.
Public Sub GetSoloName_Finalize()
    If keys Is Nothing = False Then
        Set keys = Nothing
    End If
End Sub

' Make unique text
' iName: Name for duplicated check
' iRange: Check range
' iLen: Maximum length of result string
Function GetSoloName(iName As String, iRange As Range, iLen As Integer) As String
    Dim nam As String
    nam = Trim(iName)
    
    If iLen < 5 Then
        GetSoloName = "Minimum iLen number is 6"
        Exit Function
    End If
    
    getSoloNameProc iName, iRange, iLen
    If keys.Exists(nam) = False Then
        GetSoloName_Finalize
        getSoloNameProc iName, iRange, iLen
    End If
    GetSoloName = keys(nam)(1)
End Function

' speedup caching process
Private Function getSoloNameProc(iName As String, iRange As Range, iLen As Integer) As String
    If keys Is Nothing Then
        Set keys = New Scripting.Dictionary
    
        Dim nam As String
        nam = Trim(iName)
            
        Dim rs As String
        Dim lst As Scripting.Dictionary
        Dim r As Range
        
        For Each r In iRange
            rs = Trim(r.Value)
            If rs <> "" Then
                If keys.Exists(rs) = False Then
                    Set lst = New Scripting.Dictionary
                    lst(lst.Count + 1) = rs
                    Set keys(rs) = lst
                End If
            End If
        Next

        Dim k As Variant
        For Each k In keys
            If Len(k) > iLen Then
                Dim k2 As String
                k2 = Left(k, iLen)
                If keys.Exists(k2) = False Then
                    Set lst = New Scripting.Dictionary
                    Set keys(k2) = lst
                Else
                    Set lst = keys(k2)
                End If
                Set lst(lst.Count + 1) = k
            End If
        Next
        Dim dups As Scripting.Dictionary
        Set dups = New Scripting.Dictionary
        Dim k3 As Variant
        
        For Each k In keys.keys
            Set lst = keys(k)
            If lst.Count = 1 Then
                keys(k)(1) = Left(k, iLen)
            End If
        Next
        
        For Each k In keys.keys
            Set lst = keys(k)
            If lst.Count > 1 Then
                j = 1
                Do
                    dups.RemoveAll
                    For i = 1 To lst.Count
                        k1 = lst(i)
                        k2 = Left(k1, iLen - j - 1) & "-" & Right(k1, j)
                        dups(k2) = k1
                    Next
                    If dups.Count = lst.Count Then
                        nextflag = False
                        For Each k3 In dups.keys
                            If keys.Exists(k3) Then
                                nextflag = True
                                Exit For
                            End If
                            If keys.Exists(dups(k3)) = False Then
                                Set keys(dups(k3)) = New Scripting.Dictionary
                            End If
                            keys(dups(k3))(1) = k3
                        Next
                        If nextflag = False Then
                            Exit Do
                        End If
                    End If
                    
                    j = j + 1
                    If j > iLen - 2 Then
                        For iii = 1 To lst.Count
                            keys(lst(iii))(1) = Left(lst(iii), iLen) & " - DUP ERROR"
                        Next
                        Exit Do
                    End If
                Loop
            End If
        Next
    End If
End Function

