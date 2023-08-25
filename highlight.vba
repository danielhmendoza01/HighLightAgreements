Sub HighlightAgreements()
    Dim wb As Workbook
    Dim ws, ws1 As Worksheet
    Dim totalYellow, totalRed, totalGreen As Long
    Dim dataDict As New Scripting.Dictionary

    Set ws1 = ActiveWorkbook.Worksheets("Sheet1")

    Set wb = Workbooks.Open(ActiveWorkbook.Path & "/" & Range("A2").Value)

    Set ws = wb.Sheets(1)

    FirstCol = Columns(ws1.Range("B2").Value).Column

    LastCol = Columns(ws1.Range("C2").Value).Column

    Columns("A:J").Hidden = True

    For k = FirstCol To LastCol
        FirstValue = ws.Cells(3, k).Value
        SecondValue = ws.Cells(4, k).Value

        Arr1 = Split(FirstValue, ",")
        Arr2 = Split(SecondValue, ",")
        
        ArrMerged = Join(Arr1, ",") & "," & Join(Arr2, ",")
        ArrMerged = Split(ArrMerged, ",")

        Matched = 0
        NotMatched = 0

        For i = LBound(Arr1) To UBound(Arr1)
            For j = LBound(Arr2) To UBound(Arr2)
                If Arr1(i) = Arr2(j) Then
                    Matched = Matched + 1
                Else
                    NotMatched = NotMatched + 1
                End If
            Next j
        Next i
        
        Dim areEqual As Boolean
        areEqual = StrComp(Join(Arr1, ","), Join(Arr2, ","), vbBinaryCompare) = 0
        If areEqual Then
            NotMatched = 0
        End If

        If ws.Cells(3, k).Value <> "" And ws.Cells(4, k) <> "" Then
            Dim arrKey As Variant
            For Each arrKey In ArrMerged
                arrKey = Trim(LCase(arrKey))
                
                ' Track response
                If Not dataDict.Exists(arrKey) Then
                    Set dataDict(arrKey) = CreateObject("Scripting.Dictionary")
                    dataDict(arrKey).Add "Total", 0
                    dataDict(arrKey).Add "Agree", 0
                    dataDict(arrKey).Add "Disagree", 0
                    dataDict(arrKey).Add "Partial", 0
                End If

                dataDict(arrKey)("Total") = dataDict(arrKey)("Total") + 1
                If Matched >= 1 And NotMatched >= 1 Then
                    ws.Cells(3, k).Interior.Color = vbYellow
                    ws.Cells(4, k).Interior.Color = vbYellow
                    ws.Cells(5, k).Value = 2
                    dataDict(arrKey)("Partial") = dataDict(arrKey)("Partial") + 1
                ElseIf Matched = 0 And NotMatched >= 1 Then
                    ws.Cells(3, k).Interior.Color = vbRed
                    ws.Cells(4, k).Interior.Color = vbRed
                    ws.Cells(5, k).Value = 3
                    dataDict(arrKey)("Disagree") = dataDict(arrKey)("Disagree") + 1
                Else
                    ws.Cells(3, k).Interior.Color = vbGreen
                    ws.Cells(4, k).Interior.Color = vbGreen
                    ws.Cells(5, k).Value = 1
                    dataDict(arrKey)("Agree") = dataDict(arrKey)("Agree") + 1
                End If
            Next arrKey
        End If

    Next k

    ' Write total results to K and L columns
    i = 10
    Dim key As Variant
    For Each key In dataDict.Keys
        ws.Cells(i, "K").Value = key
        ws.Cells(i, "L").Value = dataDict(key)("Total")
        ws.Cells(i + 1, "K").Value = "Agree"
        ws.Cells(i + 1, "L").Value = dataDict(key)("Agree")
        ws.Cells(i + 1, "M").Value = IIf(dataDict(key)("Agree") = 0, "0%", dataDict(key)("Agree") / dataDict(key)("Total") * 100 & "%")
        ws.Cells(i + 2, "K").Value = "Disagree"
        ws.Cells(i + 2, "L").Value = dataDict(key)("Disagree")
        ws.Cells(i + 2, "M").Value = IIf(dataDict(key)("Disagree") = 0, "0%", dataDict(key)("Disagree") / dataDict(key)("Total") * 100 & "%")
        ws.Cells(i + 3, "K").Value = "Partial"
        ws.Cells(i + 3, "L").Value = dataDict(key)("Partial")
        ws.Cells(i + 3, "M").Value = IIf(dataDict(key)("Partial") = 0, "0%", dataDict(key)("Partial") / dataDict(key)("Total") * 100 & "%")
        i = i + 5
    Next
    totalYellow = 0
    totalRed = 0
    totalGreen = 0

    ' Iterate through each cell in row 5
    For Each cell In ws.Range("5:5")
        ' Check cell's value and increment appropriate counter
        Select Case cell.Value
            Case 1
                totalGreen = totalGreen + 1
            Case 2
                totalYellow = totalYellow + 1
            Case 3
                totalRed = totalRed + 1
        End Select
    Next cell
    ws.Range("K6").Value = "Yellow"
    ws.Range("L6").Value = totalYellow
    ws.Range("K7").Value = "Red"
    ws.Range("L7").Value = totalRed
    ws.Range("K8").Value = "Green"
    ws.Range("L8").Value = totalGreen

End Sub

