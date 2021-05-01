VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub worksheet_change(ByVal target As Range)
    Dim checkTarget As Range
    Set checkTarget = Application.Intersect(target, Range("F4:F150"))
    If checkTarget Is Nothing Then
        Debug.Print "not active range, exit"
    Else
        Dim list_ppl As String
        Dim TempString As String
        TempString = ""
        
'looks at columns titles "Skill 1", "Skill 2", and "Skill 3", on "Profiles" sheet and relates
'name of person from group members table with the skills they believe to possess and
'narrows down the search option, so when assigning tasks only certain names will appear

        Set rng = Sheets("Profiles").ListObjects("Group").ListColumns("Skill 1").DataBodyRange
        For Each aCell In rng.Cells
            If aCell.Value = checkTarget.Value Then
                list_ppl = list_ppl & "," & Sheets("Profiles").Range("B" & aCell.Row).Value
            End If
        Next aCell
        
        Set rng = Sheets("Profiles").ListObjects("Group").ListColumns("Skill 2").DataBodyRange
        For Each aCell In rng.Cells
            If aCell.Value = checkTarget.Value Then
                list_ppl = list_ppl & "," & Sheets("Profiles").Range("B" & aCell.Row).Value
            End If
        Next aCell
        
        Set rng = Sheets("Profiles").ListObjects("Group").ListColumns("Skill 3").DataBodyRange
        For Each aCell In rng.Cells
            If aCell.Value = checkTarget.Value Then
                list_ppl = list_ppl & "," & Sheets("Profiles").Range("B" & aCell.Row).Value
            End If
        Next aCell
        
        If list_ppl <> "" Then
            Range("G" & checkTarget.Row).Value = ""
            TempString = Right(list_ppl, Len(list_ppl) - 1)
            With Range("G" & checkTarget.Row).Validation
            .Delete
            'replace "=A1:A6" with the range the data is in.
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:=TempString
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
            End With
        Else
            Range("G" & checkTarget.Row).Value = ""
            With Range("G" & checkTarget.Row).Validation
            .Delete
            End With
        End If
        
    End If

If target.Address = "$M$12" Then
    If ActiveSheet.Cells(12, "M").Value = "" Then
        ActiveSheet.Cells(12, "M").NumberFormat = general
    End If
End If

End Sub

