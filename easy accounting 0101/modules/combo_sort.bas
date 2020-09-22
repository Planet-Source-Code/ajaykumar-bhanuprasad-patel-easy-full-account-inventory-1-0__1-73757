Attribute VB_Name = "combo_sort"
Public i%
Public j%
Public ChoosingFirst As Boolean, ChoosingLast As Boolean, TheNums()
Enum lOrder
    Ascending
    Descending
End Enum
Public Sub RandomList(ByVal intMax As Integer)
    Dim High(), Nums(), iMax As Integer
    iMax% = intMax% - 1
    ReDim High(iMax%)
    ReDim Nums(iMax%)
    For i% = 0 To iMax%
        High(i%) = i%
    Next
    For i% = iMax% To 0 Step -1
        Chosen% = Int(i% * Rnd)
        Nums(iMax% - i%) = High(Chosen%)
        High(Chosen%) = High(i%)
    Next
    TheNums = Nums
End Sub
Public Sub ScrambleList(ByVal lList As ListBox)
    Dim arrList() As String
    ReDim arrList(lList.ListCount - 1)
    RandomList lList.ListCount
    For i% = 0 To UBound(arrList)
        arrList(i%) = lList.List(TheNums(i%))
    Next
    For i% = 0 To UBound(arrList)
        lList.List(i%) = arrList(i%)
    Next
End Sub
Public Sub SortArray(arrList())
    Dim strTemp As String, InOrder As Boolean
    Dim c As Integer, l As Integer
    Dim t1 As String, t2 As String
    Dim c1 As String, c2 As String
    For i% = 0 To UBound(arrList()) - 1
        For j% = i% + 1 To UBound(arrList())
            t1$ = CStr(arrList(i%))
            t2$ = CStr(arrList(j%))
            If Len(t1$) > Len(t2$) Then
                l% = Len(t1$)
            Else
                l% = Len(t2$)
            End If

            For c% = 1 To l%
                
                c1$ = UCase$(Mid$(t1$, c%, 1))
                c2$ = UCase$(Mid$(t2$, c%, 1))
                If Len(c1$) = 0 Then
                    InOrder = True
                    GoTo skip
                ElseIf Len(c2$) = 0 Then
                    InOrder = False
                    GoTo skip
                End If
                If Asc(c1$) < Asc(c2$) Then
                    InOrder = True
                    GoTo skip
                ElseIf Asc(c2$) < Asc(c1$) Then
                    InOrder = False
                    GoTo skip
                End If
            Next
skip:
            If Not InOrder Then
                strTemp$ = arrList(j%)
                arrList(j%) = arrList(i%)
                arrList(i%) = strTemp$
            End If
        Next
    Next
End Sub

Public Sub SortList(ByVal lList As ComboBox, ByVal FirstIndex As Integer, ByVal LastIndex As Integer, ByVal Order As lOrder)
    Dim arrList(), NewFirstIndex As Integer, NewLastIndex As Integer, NewStep As Integer, OtherIndex As Integer
    If FirstIndex% >= LastIndex% Then
        Exit Sub
    End If
    If lList.ListCount < 2 Then
        MsgBox lList.Name & " does not contain enough list items."
        Exit Sub
    Else
        If lList.ListCount = 2 And FirstIndex% <> 0 Then
            MsgBox "FirstIndex must be equal to 0."
            Exit Sub
        ElseIf FirstIndex% < 0 Or FirstIndex% > lList.ListCount - 2 Then
            MsgBox "FirstIndex must be between or equal to 0 and " & lList.ListCount - 2 & "."
            Exit Sub
        End If
        If lList.ListCount = 2 And LastIndex% <> 1 Then
            MsgBox "LastIndex must be equal to 1."
            Exit Sub
        ElseIf LastIndex% < 1 Or LastIndex% > lList.ListCount - 1 Then
            MsgBox "LastIndex must be between or equal to 1 and " & lList.ListCount - 1 & "."
            Exit Sub
        End If
    End If
    ReDim arrList(LastIndex% - FirstIndex%)
    
    For i% = FirstIndex% To LastIndex%
        arrList(i% - FirstIndex%) = lList.List(i%)
    Next
    SortArray arrList()
    If Order = Descending Then
        NewFirstIndex% = LastIndex%
        NewLastIndex% = FirstIndex%
        NewStep% = -1
    Else
        NewFirstIndex% = FirstIndex%
        NewLastIndex% = LastIndex%
        NewStep% = 1
    End If
    OtherIndex% = 0
    For i = NewFirstIndex% To NewLastIndex% Step NewStep%
        lList.List(i) = arrList(OtherIndex%)
        OtherIndex% = OtherIndex% + 1
    Next
End Sub
