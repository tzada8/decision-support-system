Attribute VB_Name = "DataMsgBox"
Sub FindData()

Profiles.Activate
check = Range("C4").Value

Dim CalDate As Long
Dim Day As Long
Day = Sheets("Profiles").Range("E3").Value

'Only certain code will run depending on which term is selected in "C4"

'Code uses date from specific search box and creates a message box for the user
'telling them if there is anything due / starting on the specifc date they inputted

If check = "Winter" Then
    'used if value being searched for is in January
    For i = 1 To 5
        For j = 2 To 8
            If Month1.Cells(i * 5, j) <> "" Then
                CalDate = CLng(Month1.Cells(i * 5, j))
                    If CalDate = Day Then
                        MsgBox Month1.Cells(i * 5 + 1, j) & VBA.Constants.vbNewLine & Month1.Cells(i * 5 + 2, j) & VBA.Constants.vbNewLine & Month1.Cells(i * 5 + 3, j) & VBA.Constants.vbNewLine & Month1.Cells(i * 5 + 4, j), , "Upcoming:"
                    End If
            End If
        Next j
    Next i

    'February
    For i = 1 To 5
        For j = 2 To 8
            If Month2.Cells(i * 5, j) <> "" Then
                CalDate = CLng(Month2.Cells(i * 5, j))
                    If CalDate = Day Then
                        MsgBox Month2.Cells(i * 5 + 1, j) & VBA.Constants.vbNewLine & Month2.Cells(i * 5 + 2, j) & VBA.Constants.vbNewLine & Month2.Cells(i * 5 + 3, j) & VBA.Constants.vbNewLine & Month2.Cells(i * 5 + 4, j), , "Upcoming:"
                    End If
            End If
        Next j
    Next i

    'March
    For i = 1 To 5
        For j = 2 To 8
            If Month3.Cells(i * 5, j) <> "" Then
                CalDate = CLng(Month3.Cells(i * 5, j))
                    If CalDate = Day Then
                        MsgBox Month3.Cells(i * 5 + 1, j) & VBA.Constants.vbNewLine & Month3.Cells(i * 5 + 2, j) & VBA.Constants.vbNewLine & Month3.Cells(i * 5 + 3, j) & VBA.Constants.vbNewLine & Month3.Cells(i * 5 + 4, j), , "Upcoming:"
                    End If
            End If
        Next j
    Next i

    'April
    For i = 1 To 5
        For j = 2 To 8
            If Month4.Cells(i * 5, j) <> "" Then
                CalDate = CLng(Month4.Cells(i * 5, j))
                    If CalDate = Day Then
                        MsgBox Month4.Cells(i * 5 + 1, j) & VBA.Constants.vbNewLine & Month4.Cells(i * 5 + 2, j) & VBA.Constants.vbNewLine & Month4.Cells(i * 5 + 3, j) & VBA.Constants.vbNewLine & Month4.Cells(i * 5 + 4, j), , "Upcoming:"
                    End If
            End If
        Next j
    Next i
End If

If check = "Spring" Then

    'May
    For i = 1 To 5
        For j = 2 To 8
            If Month5.Cells(i * 5, j) <> "" Then
                CalDate = CLng(Month5.Cells(i * 5, j))
                    If CalDate = Day Then
                        MsgBox Month5.Cells(i * 5 + 1, j) & VBA.Constants.vbNewLine & Month5.Cells(i * 5 + 2, j) & VBA.Constants.vbNewLine & Month5.Cells(i * 5 + 3, j) & VBA.Constants.vbNewLine & Month5.Cells(i * 5 + 4, j), , "Upcoming:"
                    End If
            End If
        Next j
    Next i

    'June
    For i = 1 To 5
        For j = 2 To 8
            If Month6.Cells(i * 5, j) <> "" Then
                CalDate = CLng(Month6.Cells(i * 5, j))
                    If CalDate = Day Then
                        MsgBox Month6.Cells(i * 5 + 1, j) & VBA.Constants.vbNewLine & Month6.Cells(i * 5 + 2, j) & VBA.Constants.vbNewLine & Month6.Cells(i * 5 + 3, j) & VBA.Constants.vbNewLine & Month6.Cells(i * 5 + 4, j), , "Upcoming:"
                    End If
            End If
        Next j
    Next i
    
    'July
    For i = 1 To 5
        For j = 2 To 8
            If Month7.Cells(i * 5, j) <> "" Then
                CalDate = CLng(Month7.Cells(i * 5, j))
                    If CalDate = Day Then
                        MsgBox Month7.Cells(i * 5 + 1, j) & VBA.Constants.vbNewLine & Month7.Cells(i * 5 + 2, j) & VBA.Constants.vbNewLine & Month7.Cells(i * 5 + 3, j) & VBA.Constants.vbNewLine & Month7.Cells(i * 5 + 4, j), , "Upcoming:"
                    End If
            End If
        Next j
    Next i
    
    'August
    For i = 1 To 5
        For j = 2 To 8
            If Month8.Cells(i * 5, j) <> "" Then
                CalDate = CLng(Month8.Cells(i * 5, j))
                    If CalDate = Day Then
                        MsgBox Month8.Cells(i * 5 + 1, j) & VBA.Constants.vbNewLine & Month8.Cells(i * 5 + 2, j) & VBA.Constants.vbNewLine & Month8.Cells(i * 5 + 3, j) & VBA.Constants.vbNewLine & Month8.Cells(i * 5 + 4, j), , "Upcoming:"
                    End If
            End If
        Next j
    Next i
End If

If check = "Fall" Then

    'September
    For i = 1 To 5
        For j = 2 To 8
            If Month9.Cells(i * 5, j) <> "" Then
                CalDate = CLng(Month9.Cells(i * 5, j))
                    If CalDate = Day Then
                        MsgBox Month9.Cells(i * 5 + 1, j) & VBA.Constants.vbNewLine & Month9.Cells(i * 5 + 2, j) & VBA.Constants.vbNewLine & Month9.Cells(i * 5 + 3, j) & VBA.Constants.vbNewLine & Month9.Cells(i * 5 + 4, j), , "Upcoming:"
                    End If
            End If
        Next j
    Next i

    'October
    For i = 1 To 5
        For j = 2 To 8
            If Month10.Cells(i * 5, j) <> "" Then
                CalDate = CLng(Month10.Cells(i * 5, j))
                    If CalDate = Day Then
                        MsgBox Month10.Cells(i * 5 + 1, j) & VBA.Constants.vbNewLine & Month10.Cells(i * 5 + 2, j) & VBA.Constants.vbNewLine & Month10.Cells(i * 5 + 3, j) & VBA.Constants.vbNewLine & Month10.Cells(i * 5 + 4, j), , "Upcoming:"
                    End If
            End If
        Next j
    Next i

    'November
    For i = 1 To 5
        For j = 2 To 8
            If Month2.Cells(i * 5, j) <> "" Then
                CalDate = CLng(Month11.Cells(i * 5, j))
                    If CalDate = Day Then
                        MsgBox Month11.Cells(i * 5 + 1, j) & VBA.Constants.vbNewLine & Month11.Cells(i * 5 + 2, j) & VBA.Constants.vbNewLine & Month11.Cells(i * 5 + 3, j) & VBA.Constants.vbNewLine & Month11.Cells(i * 5 + 4, j), , "Upcoming:"
                    End If
            End If
        Next j
    Next i

    'December
    For i = 1 To 5
        For j = 2 To 8
            If Month2.Cells(i * 5, j) <> "" Then
                CalDate = CLng(Month12.Cells(i * 5, j))
                    If CalDate = Day Then
                        MsgBox Month12.Cells(i * 5 + 1, j) & VBA.Constants.vbNewLine & Month12.Cells(i * 5 + 2, j) & VBA.Constants.vbNewLine & Month12.Cells(i * 5 + 3, j) & VBA.Constants.vbNewLine & Month12.Cells(i * 5 + 4, j), , "Upcoming:"
                    End If
            End If
        Next j
    Next i
End If

End Sub

