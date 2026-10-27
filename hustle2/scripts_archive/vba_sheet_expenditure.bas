Private Sub Worksheet_Activate()
    On Error GoTo SafeExit
    Dim wsDB As Worksheet
    Set wsDB = ThisWorkbook.Sheets("DATABASE_EXPENDITURE")

    Dim lastRow As Long
    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row

    Dim months() As String
    Dim cnt As Long: cnt = 0
    Dim i As Long, j As Long, v As String, isNew As Boolean

    For i = 2 To lastRow
        v = Trim(CStr(wsDB.Cells(i, 1).Value))
        If v <> "" And Left(v, 3) <> "---" Then
            isNew = True
            For j = 1 To cnt
                If months(j) = v Then isNew = False: Exit For
            Next j
            If isNew Then
                cnt = cnt + 1
                ReDim Preserve months(1 To cnt)
                months(cnt) = v
            End If
        End If
    Next i

    If cnt > 0 Then
        Dim listStr As String
        listStr = months(1)
        For j = 2 To cnt
            listStr = listStr & "," & months(j)
        Next j
        With Me.Range("C1").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:=listStr
            .ShowError = False ' Allow user to type new months without error
        End With
    End If

    Call UpdateMonthHeader
SafeExit:
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("C1")) Is Nothing Then
        Call UpdateMonthHeader
    End If
End Sub

Private Sub UpdateMonthHeader()
    Dim val As Variant
    val = Me.Range("C1").Value
    If IsEmpty(val) Or CStr(val) = "" Then Exit Sub
    
    Dim mName As String
    mName = ""
    
    If IsDate(val) Then
        mName = UCase(Format(CDate(val), "MMMM"))
    Else
        Dim lbl As String
        lbl = Trim(CStr(val))
        If InStr(lbl, "-") > 0 Then
            Dim parts() As String
            parts = Split(lbl, "-")
            mName = UCase(Trim(parts(0)))
        Else
            mName = UCase(lbl)
        End If
    End If
    
    If mName <> "" Then
        Me.Cells(2, 5).Value = mName
    End If
End Sub
