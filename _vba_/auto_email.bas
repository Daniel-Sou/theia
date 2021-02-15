' Example 01
Attribute VB_Name = "Module1"

Sub Email_From_Excel_Basic()
    ' TeachExcel.com

    Dim emailApplication As Object
    Dim emailItem As Object

    Set emailApplication = CreateObject("Outlook.Application")
    Set emailItem = emailApplication.CreateItem(0)

    ' Now we build the email.

    emailItem.to = "email@test.com; email2@test.com; anotheremail@test.com"

    emailItem.Subject = "Subject line for the email."

    emailItem.Body = "The message for the email."

    ' Send the Email
    ' Use this OR .Display, but not both together.
    emailItem.Send

    ' Display the Email so the user can change it as desired before sending it.
    ' Use this OR .Send, but not both together.
    'emailItem.Display

    Set emailItem = Nothing
    Set emailApplication = Nothing

End Sub

Sub Email_From_Excel_More_Options()
    ' TeachExcel.com

    Dim emailApplication As Object
    Dim emailItem As Object

    Set emailApplication = CreateObject("Outlook.Application")
    Set emailItem = emailApplication.CreateItem(0)

    ' Now we build the email.

    emailItem.to = "email@test.com"

    emailItem.CC = "email2@test.com"

    emailItem.BCC = "email3@test.com"

    emailItem.Subject = "Subject line for the email."

    emailItem.Body = "The message for the email."

    ' Send the Email
    emailItem.Send

    Set emailItem = Nothing
    Set emailApplication = Nothing

End Sub

Sub Email_From_Excel_Attachments()
    ' TeachExcel.com

    Dim emailApplication As Object
    Dim emailItem As Object

    Set emailApplication = CreateObject("Outlook.Application")
    Set emailItem = emailApplication.CreateItem(0)

    ' Now we build the email.

    emailItem.to = "email@test.com"

    emailItem.Subject = "Subject line for the email."

    emailItem.Body = "The message for the email."

    ' Attach current Workbook
    emailItem.Attachments.Add ActiveWorkbook.FullName

    ' Attach any file from your computer.
    'emailItem.Attachments.Add ("C:\test\test.xlsx")

    ' Send the Email
    emailItem.Send

    Set emailItem = Nothing
    Set emailApplication = Nothing

End Sub

Sub Send_Email_With_Code_Hints()
    ' This is NOT recommended because it requires the addition of
    ' an object library reference.
    ' If you send this workbook to someone who does not also have
    ' the correct reference added, it will cause an error.
    '
    ' Steps to enable the required reference:
    ' Tools > References > Microsoft Outlook XX.0 Object Library (Make sure there is a check mark next to this.)
    '
    ' TeachExcel.com

    Dim emailApplication As Outlook.Application
    Dim emailItem As Outlook.MailItem

    Set emailApplication = New Outlook.Application
    Set emailItem = emailApplication.CreateItem(olMailItem)
    
    ' Now we build the email.

    emailItem.to = "email@test.com"

    emailItem.Subject = "Subject line for the email."

    emailItem.Body = "The message for the email."

    ' Send the Email
    emailItem.Send

    Set emailItem = Nothing
    Set emailApplication = Nothing

End Sub
