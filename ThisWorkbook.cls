VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Const orange As Long = 49407 'orange
Const blue As Long = 12611584 'blue
Const lgreen As Long = 5296274 'light green
Const dred As Long = 192 'dark red

Sub Workbook_SheetChange(ByVal Sh As Object, ByVal target As Range)
Select Case Range("C4")

Dim ws As Object

'Each Case is the same. When a user clicks on the specific option,
'only the orange and that specific color's sheets will appear

Case "Winter"
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Tab.Color = blue Or ws.Tab.Color = orange Then
            ws.Visible = True
        End If
    Next ws
    
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Tab.Color = lgreen Or ws.Tab.Color = dred Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    
Case "Spring"
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Tab.Color = lgreen Or ws.Tab.Color = orange Then
            ws.Visible = True
        End If
    Next ws
    
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Tab.Color = blue Or ws.Tab.Color = dred Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

Case "Fall"
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Tab.Color = dred Or ws.Tab.Color = orange Then
            ws.Visible = True
        End If
    Next ws

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Tab.Color = blue Or ws.Tab.Color = lgreen Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

End Select

End Sub

'This code just hides all sheets except for the red ones.
    'It should only be used when editing sheets, and when the entire DSS
    'is complete; so the user will not see a term's months when they open
    'the DSS for their first time
Sub Hide_All_Sheets()
    
Dim ws As Object
    
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Tab.Color <> orange Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
        
End Sub

'Shows all worksheets; only should be used during the creation of the DSS
Sub Show_All_Sheets()

Dim ws As Object

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Tab.Color <> orange Then
            ws.Visible = xlSheetVisible
        End If
    Next ws
    
End Sub


'Color Property Value for Standard Colors
    '192 - Dark Red
    '255 - Red
    '65535 - Yellow
    '15773696 - Light Blue
    '12611584 - Blue
    '5296274 - Light Green
    '5287936 - Green
    '49407 - Orange
