Attribute VB_Name = "Module1"
'Adds rows to increase amount of visible rows of group members in "Profiles"
Sub AddRow()

Dim ws As Worksheet
    
    Set ws = ActiveSheet
        ws.ListObjects("Group").ListRows.Add
End Sub

'Deletes rows to decrease amount of visible rows of group members in "Profiles"
Sub RemoveRow()

Dim ws As Worksheet
Dim lastrow As Long
    
    Set ws = ActiveSheet
        lastrow = ws.ListObjects("Group").Range.Rows.Count
        ws.ListObjects("Group").ListRows(lastrow - 1).Delete
    
End Sub


'Hides sheet "Skill_Set" which we do not want the user to see
Sub Hide_Sheet7()

Worksheets("Skill_Set").Visible = xlSheetVeryHidden

End Sub

'Shows sheet "Skill_Set" incase the DSS needs to be edited / changed
Sub Show_Sheet7()

Worksheets("Skill_Set").Visible = xlSheetVisible

End Sub

