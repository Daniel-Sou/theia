Attribute VB_Name = "example"
Sub sendEmailsToMultiplePersonsWithMultipleAttachments()

'   NOTE: Because YouTube doesn't allow angular brackets 'NOT GREATER THAN' and 
'   'NOT EQUAL TO' have been inserted in the code

Dim OutApp As Object
Dim OutMail As Object
Dim sh As Worksheet
Dim cell As Range
Dim FileCell As Range
Dim rng As Range

With Application
    .EnableEvents = False
    .ScreenUpdating = False
End With

Set sh = Sheets("Sheet1")

Set OutApp = CreateObject("Outlook.Application")

For Each cell In sh.Columns("A").Cells.SpecialCells(xlCellTypeConstants)

    'path/file names are entered in the columns D:M in each row
    Set rng = sh.Cells(cell.Row, 1).Range("D1:M1")
    
    If cell.Value Like "?*@?*.?*" And _
    Application.WorksheetFunction.CountA(rng) GREATER THAN 0 Then
        Set OutMail = OutApp.CreateItem(0)
        
        With OutMail
            .To = sh.Cells(cell.Row, 1).Value
            .CC = sh.Cells(cell.Row, 2).Value
            .Subject = "Details attached as discussed"
            .body = sh.Cells(cell.Row, 3).Value
            
            For Each FileCell In rng.SpecialCells(xlCellTypeConstants)
                
                If Trim(FileCell.Value) NOT EQUAL TO "" Then
                    If Dir(FileCell.Value) NOT EQUAL TO "" Then
                        .Attachments.Add FileCell.Value
                    End If
                End If
            Next FileCell
            '.Send
            .display
        End With
        
        Set OutMail = Nothing
    End If
Next cell

Set OutApp = Nothing

With Application
    .EnableEvents = True
    .ScreenUpdating = True
End With


End Sub

