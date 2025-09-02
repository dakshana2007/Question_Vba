Sub ShuffleQuestionsAndOptions()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1) ' Update if your sheet is named differently

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Store all rows temporarily (A-J range)
    Dim data() As Variant
    data = ws.Range("A2:J" & lastRow).Value ' Now includes column J

    ' Shuffle questions (rows)
    Call ShuffleRows(data)

    ' Clear and rewrite shuffled rows
    ws.Range("A2:J" & lastRow).ClearContents
    
    ' Write data back using a loop to assign values correctly
    Dim i As Long, j As Long
    For i = 1 To UBound(data, 1)
        For j = 1 To UBound(data, 2)
            ws.Cells(i + 1, j).Value = data(i, j)
        Next j
    Next i

    ' Now shuffle options in each question (columns B to E are options)
    Dim k As Long
    For k = 1 To UBound(data)
        Dim options(1 To 4) As String
        Dim correctLetter As String
        correctLetter = UCase(Trim(data(k, 6))) ' Correct answer is in column F (6)

        ' Store options (columns B to E)
        For j = 1 To 4
            options(j) = data(k, j + 1)
        Next j

        ' Determine correct option index
        Dim correctIndex As Long
        correctIndex = Asc(correctLetter) - 64 ' A=1, B=2, etc.

        ' Identify correct option text (without "?")
        Dim correctOption As String
        correctOption = RemoveTick(options(correctIndex))

        ' Shuffle option indices
        Dim shuffledIndexes() As Integer
        ReDim shuffledIndexes(1 To 4)
        For j = 1 To 4: shuffledIndexes(j) = j: Next j
        Call ShuffleArray(shuffledIndexes)

        ' Reassign shuffled options
        Dim newOptions(1 To 4) As String
        For j = 1 To 4
            newOptions(j) = options(shuffledIndexes(j))
        Next j

        ' Find new correct index
        For j = 1 To 4
            If RemoveTick(newOptions(j)) = correctOption Then
                data(k, 6) = Chr(64 + j) ' Update correct answer (A-D)
                Exit For
            End If
        Next j

        ' Write back shuffled options (columns B to E)
        For j = 1 To 4
            data(k, j + 1) = newOptions(j)
        Next j
    Next k

    ' Output the final shuffled data (columns A to J)
    ws.Range("A2").Resize(UBound(data), UBound(data, 2)).Value = data

    MsgBox "Questions and options shuffled!", vbInformation

End Sub

Sub ShuffleArray(arr() As Integer)

    Dim i As Long, j As Long, temp As Integer

    Randomize

    For i = UBound(arr) To LBound(arr) + 1 Step -1
        j = Int((i - LBound(arr) + 1) * Rnd + LBound(arr))
        temp = arr(i)
        arr(i) = arr(j)
        arr(j) = temp
    Next i

End Sub

Sub ShuffleRows(ByRef arr As Variant)

    Dim i As Long, j As Long, k As Long
    Dim temp As Variant

    Randomize

    For i = UBound(arr, 1) To LBound(arr, 1) + 1 Step -1
        j = Int((i - LBound(arr, 1) + 1) * Rnd + LBound(arr, 1))

        ' Shuffle rows (entire row from column A to J = 10 columns)
        For k = 1 To 10
            temp = arr(i, k)
            arr(i, k) = arr(j, k)
            arr(j, k) = temp
        Next k
    Next i

End Sub

Function RemoveTick(opt As String) As String
    ' Removes ? mark if present
    opt = Trim(opt)
    If InStr(opt, "?") > 0 Then
        RemoveTick = Trim(Replace(opt, "?", ""))
    Else
        RemoveTick = opt
    End If
End Function

