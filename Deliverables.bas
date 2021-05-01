Attribute VB_Name = "Deliverables"
Sub Button2_Click()
'User inputs amount of deliverables they will have to complete into C4
'and the code generates that many deliverables that can be commented on
    ActiveSheet.Range("B5:D1000").Clear
    ActiveSheet.Range("B5:B1000").Font.Color = RGB(255, 255, 255)
    Dim val As Integer
    Dim rng As Range
    j = 5
    i = 2
    k = 1
    val = Range("B4").Value

ActiveSheet.Cells.Interior.Color = RGB(255, 255, 255)
ActiveSheet.Cells(4, 3).Interior.Color = RGB(255, 192, 0)
'i = "B", i + 1 = "C", i + 2 = "D"
    For l = 1 To val
        m = j + 4
        Rows(j).RowHeight = 20
        'All for row 5
        ActiveSheet.Cells(j, i) = "Milestone " & k
        ActiveSheet.Cells(j, i).Font.Name = "Arial"
        ActiveSheet.Cells(j, i).Interior.Color = RGB(255, 192, 0)
        ActiveSheet.Cells(j, i + 1).Interior.Color = RGB(255, 255, 255)
        ActiveSheet.Cells(j, i + 2).Interior.Color = RGB(255, 255, 255)
        ActiveSheet.Cells(j, i + 1).Interior.Color = RGB(255, 192, 0)
        ActiveSheet.Cells(j, i + 2).Interior.Color = RGB(255, 192, 0)
        
        'For row 6
        ActiveSheet.Cells(j + 1, i + 1).Interior.Color = RGB(255, 255, 255)
        ActiveSheet.Cells(j + 1, i + 2).Interior.Color = RGB(255, 255, 255)

        ActiveSheet.Cells(j + 1, i) = "Feedback:"
        ActiveSheet.Cells(j + 1, i).Font.Name = "Arial"
        ActiveSheet.Cells(j + 1, i).Interior.Color = RGB(255, 192, 0)
        Range(Cells(j + 1, i + 1), Cells(j + 1, i + 2)).Merge
        
        'For row 7
        ActiveSheet.Cells(j + 2, i).Interior.Color = RGB(255, 192, 0)
        ActiveSheet.Cells(j + 2, i + 1).Interior.Color = RGB(255, 255, 255)
        ActiveSheet.Cells(j + 2, i + 2).Interior.Color = RGB(255, 255, 255)
        Range(Cells(j + 2, i + 1), Cells(j + 2, i + 2)).Merge
        
        'For row 8
        ActiveSheet.Cells(j + 3, i).Interior.Color = RGB(255, 192, 0)
        ActiveSheet.Cells(j + 3, i + 1).Interior.Color = RGB(255, 255, 255)
        ActiveSheet.Cells(j + 3, i + 2).Interior.Color = RGB(255, 255, 255)
        Range(Cells(j + 3, i + 1), Cells(j + 3, i + 2)).Merge
        k = k + 1
        
        'For row 9
        Rows(m).RowHeight = 47
        ActiveSheet.Cells(j + 4, i) = "Rubric:"
        ActiveSheet.Cells(j + 4, i).Font.Name = "Arial"
        ActiveSheet.Cells(j + 4, i).Interior.Color = RGB(255, 192, 0)
        ActiveSheet.Cells(j + 4, i + 2).Interior.Color = RGB(255, 255, 255)
        ActiveSheet.Cells(j + 4, i + 2).VerticalAlignment = xlTop
        ActiveSheet.Cells(j + 4, i + 1).Interior.Color = RGB(255, 255, 255)
            ActiveSheet.Cells(j + 4, i + 1) = ChrW("&h0032")
            ActiveSheet.Cells(j + 4, i + 1).Font.Name = "wingdings"
            ActiveSheet.Cells(j + 4, i + 1).Font.Size = 33
            ActiveSheet.Cells(j + 4, i + 1).HorizontalAlignment = xlCenter
            ActiveSheet.Cells(j + 4, i + 1).VerticalAlignment = xlCenter
            ActiveSheet.Cells(j + 4, i + 1).Font.Bold = True
        n = j + 5
        Rows(n).RowHeight = 20
        
        'For row 10
        ActiveSheet.Cells(j + 5, i) = "Grade:"
        ActiveSheet.Cells(j + 5, i).Font.Name = "Arial"
        ActiveSheet.Cells(j + 5, i).Interior.Color = RGB(255, 192, 0)
        ActiveSheet.Cells(j + 5, i + 1).Interior.Color = RGB(255, 255, 255)
        ActiveSheet.Cells(j + 5, i + 2).Interior.Color = RGB(255, 255, 255)
        Range(Cells(j + 5, i + 1), Cells(j + 5, i + 2)).Merge
        
        'Font styles of rows (where user inputs info)
        Range(Cells(j + 1, i + 1), Cells(j + 3, i + 2)).Font.Name = "Arial"
        ActiveSheet.Cells(j + 4, i + 2).Font.Name = "Arial"
        ActiveSheet.Cells(j + 5, i + 1).Font.Name = "Arial"
            
        'Borders applying for each Milestone
        With Range(Cells(j + 1, i + 1), Cells(j + 5, i + 2)).Borders
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlThin
        End With
        
        With Cells(j, i)
            .Borders(xlEdgeTop).Weight = xlThin
        End With
        
        With Range(Cells(j, i + 1), Cells(j, i + 2))
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeBottom).Weight = xlThin
        End With
        
        With Range(Cells(j, i), Cells(j + 5, i))
            .Borders(xlEdgeLeft).Weight = xlThin
        End With
        
        With Cells(j, i + 2)
            .Borders(xlEdgeRight).Weight = xlThin
        End With
        
        With Cells(j + 5, i)
            .Borders(xlEdgeBottom).Weight = xlThin
        End With
        
        j = j + 8
        
    Next l
    
End Sub
