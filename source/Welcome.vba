Attribute VB_Name = "Welcome"

Private Sub StartButton()

    Worksheets("Facility XML").Activate

End Sub

Private Sub TutorialButton()

    TutForm.SecNum.Caption = "0"
    TutForm.SecDec.Caption = "0"
    TutForm.Show

End Sub

Private Sub WorkbookInfoButton()

    DialogueForm.DialogueBox.Text = _
 _
    "Welcome to The ShakeCast Excel Workbook!" & vbNewLine & vbNewLine & _
    "This workbook was designed to simplify the process of creating ShakeCast readable " & _
    "facility, notification group, and user files. The Start button will take you to " & _
    "the Facility XML spreadsheet where you can start entering your information. The " & _
    "Notification XML spreadsheet should be visited next, and the User XML sheet last. " & _
    "These sheets have information buttons as well, which should help you to get your information entered " & _
    "correctly. " & vbNewLine & vbNewLine & _
    "If this is your first time using this Workbook, you'll see that we've provided some example " & _
    "information for you. Follow the examples to input your own data, but if you leave " & _
    "the example information in the workbook, it will be exported and uploaded to ShakeCast as well!" & vbewnline & vbNewLine & _
    "If this is your first time checking out this worksheet, go ahead an click Take a Tutorial " & _
    "so we can walk you through the process!"
    
    
    DialogueForm.Show
    
End Sub
