
Sub Sub_production_planning()
'

'
    'orders (n) a products (m)

    'Application.ScreenUpdating = False
    
    
    Worksheets("List1").Activate
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    
    Dim n As Long, m As Long
    m = Selection.Rows.count - 1
    n = Selection.Columns.count - 2
        ' first cell Cells(vyr + 1, obj + 2).Select
        
   Dim wsPlan As Worksheet, List1 As Worksheet
   Dim i As Long, j As Long, day As Long, k As Long
    
    ' ——————————————————————————————————————————————
    Dim capacities() As Double
    ReDim capacities(1 To m)

    For i = 1 To m
        capacities(i) = CDbl(Cells(i + 1, 2).Value)
    Next i
    
    Dim NameProduct() As String
    ReDim NameProduct(1 To m)
    
    For i = 1 To m
        NameProduct(i) = Cells(i + 1, 1).Value
    Next i
    ' ——————————————————————————————————————————————
    Dim planMatrix() As Double
    Dim x As Long
    x = 1000000
    ReDim planMatrix(1 To m, 1 To x)

    Dim StartDay As Long
    StartDay = 45870

    Dim Products() As Long
    ReDim Products(1 To m)
    
    ' ——————————————————————————————————————————————
    For i = 1 To n
        For j = 1 To m
            Products(j) = Products(j) + Cells(j + 1, i + 2).Value
        Next j
    Next i

    
    ' ——————————————————————————————————————————————
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Plan2").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Dim wsPlan2 As Worksheet
    Set wsPlan2 = ThisWorkbook.Worksheets.Add
    wsPlan2.Name = "Plan2"

    
    ' Planmatrix
    Dim Largest, index, Days, DaysAll, count As Long
    Largest = 0
    DaysAll = 0
    For i = 1 To m
        wsPlan2.Cells(1 + i, 1).Value = NameProduct(i)
    Next i
      
    count = 0
    
    For i = 1 To m
        If Products(i) > 0 Then
            count = count + 1
        End If
    Next i
    
    
    Do While count > 0
        For j = 1 To m
            If Products(j) > Largest Then
                index = j
                Largest = Products(j)
            End If
        Next j
        If Products(index) > capacities(index) Then
            If (Int(Products(index) / capacities(index))) - (Products(index) / capacities(index)) = 0 Then
                Days = (Int(Products(index) / capacities(index)))
                For k = 1 To Days
                    DaysAll = DaysAll + 1
                    planMatrix(index, DaysAll) = capacities(index)
                    Cells(index + 1, DaysAll + 1) = planMatrix(index, DaysAll)
                Next k
            Else
                Days = (Int(Products(index) / capacities(index)))
                For k = 1 To Days
                    DaysAll = DaysAll + 1
                    planMatrix(index, DaysAll) = capacities(index)
                    Cells(index + 1, DaysAll + 1) = planMatrix(index, DaysAll)
                Next k
                DaysAll = DaysAll + 1
                planMatrix(index, DaysAll) = (Products(index) - capacities(index) * (Int(Products(index) / capacities(index))))
                Cells(index + 1, DaysAll + 1) = planMatrix(index, DaysAll)
            End If
        Else
            DaysAll = DaysAll + 1
            planMatrix(index, DaysAll) = Products(index)
            Cells(index + 1, DaysAll + 1) = planMatrix(index, DaysAll)
        End If
        Products(index) = 0
        Largest = 0
        count = count - 1
    Loop
    
    For i = 1 To DaysAll
        Cells(1, i + 1) = i
    Next i

    wsPlan2.Columns.AutoFit
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Plan3").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Dim wsPlan3 As Worksheet
    Set wsPlan3 = ThisWorkbook.Worksheets.Add
    wsPlan3.Name = "Plan3"

    Dim index2, rows3, columns3, weeknum As Long
    rows3 = 2
    columns3 = 0
    weeknum = 1
    
    For i = 1 To DaysAll
        For k = 1 To m
            If planMatrix(k, i) <> 0 Then
                index2 = k
            End If
        Next k
        
        If columns3 = 5 Or columns3 = 11 Or columns3 = 17 Then
            Range(Cells(rows3 - 1, columns3 - 4), Cells(rows3 - 1, columns3)).Merge
            Cells(rows3 - 1, columns3 - 4).HorizontalAlignment = xlCenter
            Cells(rows3 - 1, columns3 - 4).Value = "Week " & weeknum
            Cells(rows3 - 1, columns3 - 4).Font.Color = vbWhite
            Cells(rows3 - 1, columns3 - 4).Font.Bold = True
            Cells(rows3 - 1, columns3 - 4).Interior.Color = RGB(162, 0, 112)
            columns3 = columns3 + 1
            weeknum = weeknum + 1
        End If
        
        If columns3 = 18 Then
            rows3 = rows3 + 6
            columns3 = 0
        End If
        

        
        columns3 = columns3 + 1
        
        Cells(rows3, columns3) = i
        Cells(rows3, columns3).Interior.Color = RGB(204, 0, 102)
        Cells(rows3, columns3).Font.Color = vbWhite
        Cells(rows3, columns3).Font.Bold = True
        Cells(rows3, columns3).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Cells(rows3, columns3).Borders(xlEdgeBottom).Weight = xlThick
        Cells(rows3, columns3).Borders(xlEdgeBottom).Color = vbBlack
        
        Cells(rows3 + 1, columns3).NumberFormat = "# ##0"
        
        Cells(rows3 + 1, columns3) = planMatrix(index2, i)
        Cells(rows3 + 2, columns3) = NameProduct(index2)
        Cells(rows3 + 2, columns3).HorizontalAlignment = xlRight
    Next i
    
    wsPlan3.Columns.AutoFit

    MsgBox "Plan completed", vbInformation


'Application.ScreenUpdating = True

End Sub
