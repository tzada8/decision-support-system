Attribute VB_Name = "InfoToCalendar"
Sub DataTransfer()
    'Code is unique only for "Winter" term, then repeats for the other 11 months
    Dim mdte As Long
    Dim cdte As Long
    Dim check As String
    Dim MMon As Variant
    Dim ws As Object
    Dim i As Integer
'mdte = "Main Table" Date
'cdte = Calendar Date
'MMon = A Term (4 months)

check = Profiles.Range("C4").Value
'3 different "if" statements present relating to "C4" cell where whatever it says in that
'specific cell will run one of the 3 if statements (e.g. if C4 says "Winter" then this first
'code will run, where it will input information only on the January-April calendars)

'If Term "Winter" is selected in C4 then run following code:
    If check = "Winter" Then
        MMon = Array("January", "February", "March", "April")
        For a = 0 To 3
            Worksheets(MMon(a)).Activate


        'months of Jan-Apr
            For i = 1 To 5
                ActiveSheet.Range(Cells((5 * i) + 1, 2), Cells((5 * i) + 4, 8)) = ""
                ActiveSheet.Range(Cells((5 * i) + 1, 2), Cells((5 * i) + 4, 8)).Interior.Color = RGB(255, 255, 255)
            Next i
            
            ActiveSheet.Range(Cells(31, 2), Cells(34, 3)) = ""
            ActiveSheet.Range(Cells(31, 2), Cells(34, 3)).Interior.Color = RGB(255, 255, 255)
            
            'fills all calendar parts for their start dates
            k = 4
            Do While Sheet3.Cells(k, 3) <> ""
                mdte = CLng(Sheet3.Cells(k, 3))
                For i = 1 To 5
                    For j = 2 To 8
                        If ActiveSheet.Cells(i * 5, j) <> "" Then
                            cdte = CLng(ActiveSheet.Cells(5 * i, j))
                                If cdte = mdte Then
                                    If ActiveSheet.Cells(5 * i + 1, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 1, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(5 * i + 1, j).Interior.Color = RGB(0, 140, 0)
                                    ElseIf ActiveSheet.Cells(5 * i + 2, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 2, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(5 * i + 2, j).Interior.Color = RGB(0, 140, 0)
                                    ElseIf ActiveSheet.Cells(5 * i + 3, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 3, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(5 * i + 3, j).Interior.Color = RGB(0, 140, 0)
                                    ElseIf ActiveSheet.Cells(5 * i + 4, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 4, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(5 * i + 4, j).Interior.Color = RGB(0, 140, 0)
                                    End If
                                Else:
                                End If
                        Else:
                        End If
                    Next j
                Next i
                k = k + 1
            Loop
                
            'Due date for Winter Term
            k = 4
            Do While Sheet3.Cells(k, 4) <> ""
                mdte = CLng(Sheet3.Cells(k, 4))
                For i = 1 To 5
                    For j = 2 To 8
                        If ActiveSheet.Cells(i * 5, j) <> "" Then
                            cdte = CLng(ActiveSheet.Cells(5 * i, j))
                                If cdte = mdte Then
                                    If ActiveSheet.Cells(5 * i + 1, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 1, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(5 * i + 1, j).Interior.Color = RGB(190, 0, 0)
                                    ElseIf ActiveSheet.Cells(5 * i + 2, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 2, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(5 * i + 2, j).Interior.Color = RGB(190, 0, 0)
                                    ElseIf ActiveSheet.Cells(5 * i + 3, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 3, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(5 * i + 3, j).Interior.Color = RGB(190, 0, 0)
                                    ElseIf ActiveSheet.Cells(5 * i + 4, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 4, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(5 * i + 4, j).Interior.Color = RGB(190, 0, 0)
                                    End If
                                Else:
                                End If
                        Else:
                        End If
                    Next j
                Next i
                k = k + 1
            Loop
              
            'last two boxes of calendar (start dates)
            k = 4
            Do While Sheet3.Cells(k, 3) <> ""
                mdte = CLng(Sheet3.Cells(k, 3))
                    For j = 2 To 3
                        If ActiveSheet.Cells(30, j) <> "" Then
                            cdte = CLng(ActiveSheet.Cells(30, j))
                                If cdte = mdte Then
                                    If ActiveSheet.Cells(30 + 1, j) = "" Then
                                        ActiveSheet.Cells(30 + 1, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(30 + 1, j).Interior.Color = RGB(0, 140, 0)
                                    ElseIf ActiveSheet.Cells(30 + 2, j) = "" Then
                                        ActiveSheet.Cells(30 + 2, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(30 + 2, j).Interior.Color = RGB(0, 140, 0)
                                    ElseIf ActiveSheet.Cells(30 + 3, j) = "" Then
                                        ActiveSheet.Cells(30 + 3, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(30 + 3, j).Interior.Color = RGB(0, 140, 0)
                                    ElseIf ActiveSheet.Cells(30 + 4, j) = "" Then
                                        ActiveSheet.Cells(30 + 4, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(30 + 4, j).Interior.Color = RGB(0, 140, 0)
                                    End If
                                Else:
                                End If
                        Else:
                        End If
                    Next j
                k = k + 1
            Loop
                
            'last two boxes of calendar (end dates)
            k = 4
            Do While Sheet3.Cells(k, 4) <> ""
                mdte = CLng(Sheet3.Cells(k, 4))
                    For j = 2 To 3
                        If ActiveSheet.Cells(30, j) <> "" Then
                            cdte = CLng(ActiveSheet.Cells(30, j))
                                If cdte = mdte Then
                                    If ActiveSheet.Cells(30 + 1, j) = "" Then
                                        ActiveSheet.Cells(30 + 1, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(30 + 1, j).Interior.Color = RGB(190, 0, 0)
                                    ElseIf ActiveSheet.Cells(30 + 2, j) = "" Then
                                        ActiveSheet.Cells(30 + 2, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(30 + 2, j).Interior.Color = RGB(190, 0, 0)
                                    ElseIf ActiveSheet.Cells(30 + 3, j) = "" Then
                                        ActiveSheet.Cells(30 + 3, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(30 + 3, j).Interior.Color = RGB(190, 0, 0)
                                    ElseIf ActiveSheet.Cells(30 + 4, j) = "" Then
                                        ActiveSheet.Cells(30 + 4, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(30 + 4, j).Interior.Color = RGB(190, 0, 0)
                                    End If
                                Else:
                                End If
                        Else:
                        End If
                    Next j
                k = k + 1
            Loop
            
        Next a
        
'If Term "Spring" is selected in C4 then run following code:
    ElseIf check = "Spring" Then

    MMon = Array("May", "June", "July", "August")
    For a = 0 To 3
        Worksheets(MMon(a)).Activate


        'months of May-Aug
            For i = 1 To 5
                ActiveSheet.Range(Cells((5 * i) + 1, 2), Cells((5 * i) + 4, 8)) = ""
                ActiveSheet.Range(Cells((5 * i) + 1, 2), Cells((5 * i) + 4, 8)).Interior.Color = RGB(255, 255, 255)
            Next i
            
            ActiveSheet.Range(Cells(31, 2), Cells(34, 3)) = ""
            ActiveSheet.Range(Cells(31, 2), Cells(34, 3)).Interior.Color = RGB(255, 255, 255)
            
            'fills calendar for their start dates
            k = 4
            Do While Sheet3.Cells(k, 3) <> ""
                mdte = CLng(Sheet3.Cells(k, 3))
                For i = 1 To 5
                    For j = 2 To 8
                        If ActiveSheet.Cells(i * 5, j) <> "" Then
                            cdte = CLng(ActiveSheet.Cells(5 * i, j))
                                If cdte = mdte Then
                                    If ActiveSheet.Cells(5 * i + 1, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 1, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(5 * i + 1, j).Interior.Color = RGB(0, 140, 0)
                                    ElseIf ActiveSheet.Cells(5 * i + 2, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 2, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(5 * i + 2, j).Interior.Color = RGB(0, 140, 0)
                                    ElseIf ActiveSheet.Cells(5 * i + 3, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 3, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(5 * i + 3, j).Interior.Color = RGB(0, 140, 0)
                                    ElseIf ActiveSheet.Cells(5 * i + 4, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 4, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(5 * i + 4, j).Interior.Color = RGB(0, 140, 0)
                                    End If
                                Else:
                                End If
                        Else:
                        End If
                    Next j
                Next i
                k = k + 1
            Loop
                
            'Due date for Spring Term
            k = 4
            Do While Sheet3.Cells(k, 4) <> ""
                mdte = CLng(Sheet3.Cells(k, 4))
                For i = 1 To 5
                    For j = 2 To 8
                        If ActiveSheet.Cells(i * 5, j) <> "" Then
                            cdte = CLng(ActiveSheet.Cells(5 * i, j))
                                If cdte = mdte Then
                                    If ActiveSheet.Cells(5 * i + 1, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 1, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(5 * i + 1, j).Interior.Color = RGB(190, 0, 0)
                                    ElseIf ActiveSheet.Cells(5 * i + 2, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 2, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(5 * i + 2, j).Interior.Color = RGB(190, 0, 0)
                                    ElseIf ActiveSheet.Cells(5 * i + 3, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 3, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(5 * i + 3, j).Interior.Color = RGB(190, 0, 0)
                                    ElseIf ActiveSheet.Cells(5 * i + 4, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 4, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(5 * i + 4, j).Interior.Color = RGB(190, 0, 0)
                                    End If
                                Else:
                                End If
                        Else:
                        End If
                    Next j
                Next i
                k = k + 1
            Loop
             
            k = 4
            Do While Sheet3.Cells(k, 3) <> ""
                mdte = CLng(Sheet3.Cells(k, 3))
                    For j = 2 To 3
                        If ActiveSheet.Cells(30, j) <> "" Then
                            cdte = CLng(ActiveSheet.Cells(30, j))
                                If cdte = mdte Then
                                    If ActiveSheet.Cells(30 + 1, j) = "" Then
                                        ActiveSheet.Cells(30 + 1, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(30 + 1, j).Interior.Color = RGB(0, 140, 0)
                                    ElseIf ActiveSheet.Cells(30 + 2, j) = "" Then
                                        ActiveSheet.Cells(30 + 2, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(30 + 2, j).Interior.Color = RGB(0, 140, 0)
                                    ElseIf ActiveSheet.Cells(30 + 3, j) = "" Then
                                        ActiveSheet.Cells(30 + 3, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(30 + 3, j).Interior.Color = RGB(0, 140, 0)
                                    ElseIf ActiveSheet.Cells(30 + 4, j) = "" Then
                                        ActiveSheet.Cells(30 + 4, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(30 + 4, j).Interior.Color = RGB(0, 140, 0)
                                    End If
                                Else:
                                End If
                        Else:
                        End If
                    Next j
                k = k + 1
            Loop
                
            'last two boxes of calendar (end dates)
            k = 4
            Do While Sheet3.Cells(k, 4) <> ""
                mdte = CLng(Sheet3.Cells(k, 4))
                    For j = 2 To 3
                        If ActiveSheet.Cells(30, j) <> "" Then
                            cdte = CLng(ActiveSheet.Cells(30, j))
                                If cdte = mdte Then
                                    If ActiveSheet.Cells(30 + 1, j) = "" Then
                                        ActiveSheet.Cells(30 + 1, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(30 + 1, j).Interior.Color = RGB(190, 0, 0)
                                    ElseIf ActiveSheet.Cells(30 + 2, j) = "" Then
                                        ActiveSheet.Cells(30 + 2, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(30 + 2, j).Interior.Color = RGB(190, 0, 0)
                                    ElseIf ActiveSheet.Cells(30 + 3, j) = "" Then
                                        ActiveSheet.Cells(30 + 3, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(30 + 3, j).Interior.Color = RGB(190, 0, 0)
                                    ElseIf ActiveSheet.Cells(30 + 4, j) = "" Then
                                        ActiveSheet.Cells(30 + 4, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(30 + 4, j).Interior.Color = RGB(190, 0, 0)
                                    End If
                                Else:
                                End If
                        Else:
                        End If
                    Next j
                k = k + 1
            Loop
            
        Next a
    
'If Term "Fall" is selected in C4 then run following code:
    ElseIf check = "Fall" Then

    MMon = Array("September", "October", "November", "December")
    For a = 0 To 3
        Worksheets(MMon(a)).Activate


        'months of Sept-Dec
            For i = 1 To 5
                ActiveSheet.Range(Cells((5 * i) + 1, 2), Cells((5 * i) + 4, 8)) = ""
                ActiveSheet.Range(Cells((5 * i) + 1, 2), Cells((5 * i) + 4, 8)).Interior.Color = RGB(255, 255, 255)
            Next i
            
            ActiveSheet.Range(Cells(31, 2), Cells(34, 3)) = ""
            ActiveSheet.Range(Cells(31, 2), Cells(34, 3)).Interior.Color = RGB(255, 255, 255)
            
            'fills calendar for their start dates
            k = 4
            Do While Sheet3.Cells(k, 3) <> ""
                mdte = CLng(Sheet3.Cells(k, 3))
                For i = 1 To 5
                    For j = 2 To 8
                        If ActiveSheet.Cells(i * 5, j) <> "" Then
                            cdte = CLng(ActiveSheet.Cells(5 * i, j))
                                If cdte = mdte Then
                                    If ActiveSheet.Cells(5 * i + 1, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 1, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(5 * i + 1, j).Interior.Color = RGB(0, 140, 0)
                                    ElseIf ActiveSheet.Cells(5 * i + 2, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 2, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(5 * i + 2, j).Interior.Color = RGB(0, 140, 0)
                                    ElseIf ActiveSheet.Cells(5 * i + 3, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 3, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(5 * i + 3, j).Interior.Color = RGB(0, 140, 0)
                                    ElseIf ActiveSheet.Cells(5 * i + 4, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 4, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(5 * i + 4, j).Interior.Color = RGB(0, 140, 0)
                                    End If
                                Else:
                                End If
                        Else:
                        End If
                    Next j
                Next i
                k = k + 1
            Loop
                
            'Due date for Fall Term
            k = 4
            Do While Sheet3.Cells(k, 4) <> ""
                mdte = CLng(Sheet3.Cells(k, 4))
                For i = 1 To 5
                    For j = 2 To 8
                        If ActiveSheet.Cells(i * 5, j) <> "" Then
                            cdte = CLng(ActiveSheet.Cells(5 * i, j))
                                If cdte = mdte Then
                                    If ActiveSheet.Cells(5 * i + 1, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 1, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(5 * i + 1, j).Interior.Color = RGB(190, 0, 0)
                                    ElseIf ActiveSheet.Cells(5 * i + 2, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 2, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(5 * i + 2, j).Interior.Color = RGB(190, 0, 0)
                                    ElseIf ActiveSheet.Cells(5 * i + 3, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 3, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(5 * i + 3, j).Interior.Color = RGB(190, 0, 0)
                                    ElseIf ActiveSheet.Cells(5 * i + 4, j) = "" Then
                                        ActiveSheet.Cells(5 * i + 4, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(5 * i + 4, j).Interior.Color = RGB(190, 0, 0)
                                    End If
                                Else:
                                End If
                        Else:
                        End If
                    Next j
                Next i
                k = k + 1
            Loop
            
            k = 4
            Do While Sheet3.Cells(k, 3) <> ""
                mdte = CLng(Sheet3.Cells(k, 3))
                    For j = 2 To 3
                        If ActiveSheet.Cells(30, j) <> "" Then
                            cdte = CLng(ActiveSheet.Cells(30, j))
                                If cdte = mdte Then
                                    If ActiveSheet.Cells(30 + 1, j) = "" Then
                                        ActiveSheet.Cells(30 + 1, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(30 + 1, j).Interior.Color = RGB(0, 140, 0)
                                    ElseIf ActiveSheet.Cells(30 + 2, j) = "" Then
                                        ActiveSheet.Cells(30 + 2, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(30 + 2, j).Interior.Color = RGB(0, 140, 0)
                                    ElseIf ActiveSheet.Cells(30 + 3, j) = "" Then
                                        ActiveSheet.Cells(30 + 3, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(30 + 3, j).Interior.Color = RGB(0, 140, 0)
                                    ElseIf ActiveSheet.Cells(30 + 4, j) = "" Then
                                        ActiveSheet.Cells(30 + 4, j) = Sheet3.Cells(k, 5) & "- " & Sheet3.Cells(k, 7)
                                        ActiveSheet.Cells(30 + 4, j).Interior.Color = RGB(0, 140, 0)
                                    End If
                                Else:
                                End If
                        Else:
                        End If
                    Next j
                k = k + 1
            Loop
                
            'last two boxes of calendar (end dates)
            k = 4
            Do While Sheet3.Cells(k, 4) <> ""
                mdte = CLng(Sheet3.Cells(k, 4))
                    For j = 2 To 3
                        If ActiveSheet.Cells(30, j) <> "" Then
                            cdte = CLng(ActiveSheet.Cells(30, j))
                                If cdte = mdte Then
                                    If ActiveSheet.Cells(30 + 1, j) = "" Then
                                        ActiveSheet.Cells(30 + 1, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(30 + 1, j).Interior.Color = RGB(190, 0, 0)
                                    ElseIf ActiveSheet.Cells(30 + 2, j) = "" Then
                                        ActiveSheet.Cells(30 + 2, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(30 + 2, j).Interior.Color = RGB(190, 0, 0)
                                    ElseIf ActiveSheet.Cells(30 + 3, j) = "" Then
                                        ActiveSheet.Cells(30 + 3, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(30 + 3, j).Interior.Color = RGB(190, 0, 0)
                                    ElseIf ActiveSheet.Cells(30 + 4, j) = "" Then
                                        ActiveSheet.Cells(30 + 4, j) = Sheet3.Cells(k, 5) & " Due"
                                        ActiveSheet.Cells(30 + 4, j).Interior.Color = RGB(190, 0, 0)
                                    End If
                                Else:
                                End If
                        Else:
                        End If
                    Next j
                k = k + 1
            Loop
            
        Next a
    End If
    
Sheet3.Activate
    
End Sub

