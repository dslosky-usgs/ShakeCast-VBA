VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OptionsForm 
   Caption         =   "Spreadsheet Options"
   ClientHeight    =   1580
   ClientLeft      =   -1720
   ClientTop       =   -14340
   ClientWidth     =   8060
   OleObjectBlob   =   "OptionsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OptionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ComboBox1_Change()

End Sub

Private Sub GoButton_Click()

Unload Me

Application.EnableEvents = False
Application.ScreenUpdating = False

On Error GoTo ExitHandler

If Me.OptionCombo.value = "Create Facility Type" Then
    CreateFac.Show
ElseIf Me.OptionCombo.value = "Add/Update Fragility Model" Then
    FragForm.Show
ElseIf Me.OptionCombo.value = "Create an Attribute" Then
    AddAtt.Show
    
    MsgBox "You can add attributes to a facility in row AE!"
    
ElseIf Me.OptionCombo.value = "Turn Off Data Analysis" Then

    Application.EnableEvents = False
    
    MsgBox "Data analysis has been turned off"
    
ElseIf Me.OptionCombo.value = "Turn On Data Analysis" Then
    
    Application.EnableEvents = True
    
    MsgBox "Data analysis has been turned on"
    
ElseIf Me.OptionCombo.value = "Export XML" Then

    If ActiveSheet.Name = "Facility XML" Then
        Worksheets("Facility XML").Unprotect
        
        Worksheets("ShakeCast Ref Lookup Values").Range("Q2").value = "FacilityXML"
    
        ProgressForm.Show vbModeless
        
    ElseIf ActiveSheet.Name = "Notification XML" Then
        Worksheets("Notification XML").Unprotect
        
        Worksheets("ShakeCast Ref Lookup Values").Range("Q2").value = "GroupXML"
        ProgressForm.Show vbModeless
    ElseIf ActiveSheet.Name = "User XML" Then
        Worksheets("User XML").Unprotect
        
        Worksheets("ShakeCast Ref Lookup Values").Range("Q2").value = "UserXML"
        ProgressForm.Show vbModeless
    End If
    
ElseIf Me.OptionCombo.value = "Export JSON" Then

    If ActiveSheet.Name = "Facility XML" Then
        Worksheets("Facility XML").Unprotect
        
        Worksheets("ShakeCast Ref Lookup Values").Range("Q2").value = "exportFacilityJson"
    
        ProgressForm.Show vbModeless
        
    ElseIf ActiveSheet.Name = "Notification XML" Then
        Worksheets("Notification XML").Unprotect
        
        Worksheets("ShakeCast Ref Lookup Values").Range("Q2").value = "GroupXML"
        ProgressForm.Show vbModeless
    ElseIf ActiveSheet.Name = "User XML" Then
        Worksheets("User XML").Unprotect
        
        Worksheets("ShakeCast Ref Lookup Values").Range("Q2").value = "UserXML"
        ProgressForm.Show vbModeless
    End If
    
ElseIf Me.OptionCombo.value = "Export Master XML" Then

    Worksheets("Facility XML").Unprotect
    Worksheets("Notification XML").Unprotect
    Worksheets("User XML").Unprotect
    
    Worksheets("ShakeCast Ref Lookup Values").Range("Q2").value = "MasterXML"
    ProgressForm.Show vbModeless
    
ElseIf Me.OptionCombo.value = "Update Worksheet" Then

    If ActiveSheet.Name = "Facility XML" Then
        Worksheets("Facility XML").Unprotect
        
        Worksheets("ShakeCast Ref Lookup Values").Range("Q2").value = "FacUpdate"
        ProgressForm.Show vbModeless
        
        'Application.Run "UpdateFacButton"
    ElseIf ActiveSheet.Name = "Notification XML" Then
        Worksheets("Notification XML").Unprotect
        
    ElseIf ActiveSheet.Name = "User XML" Then
        Worksheets("User XML").Unprotect
        
        Worksheets("ShakeCast Ref Lookup Values").Range("Q2").value = "UserUpdate"
        ProgressForm.Show vbModeless
    End If
    
ElseIf Me.OptionCombo.value = "Unlock Data" Then

    If ActiveSheet.Name = "Facility XML" Then
        Worksheets("Facility XML").Unprotect
        Application.Run "facilityUnlock"
    ElseIf ActiveSheet.Name = "Notification XML" Then
        Worksheets("Notification XML").Unprotect
        Application.Run "groupUnlock"
    ElseIf ActiveSheet.Name = "User XML" Then
        Worksheets("User XML").Unprotect
        Application.Run "userUnlock"
    End If

ElseIf Me.OptionCombo.value = "Clear All Data" Then

    Application.Run "clearSheet"
    
    ' now update the worksheet
    If ActiveSheet.Name = "Facility XML" Then
        Worksheets("Facility XML").Unprotect
        Application.Run "UpdateFacButton"
    ElseIf ActiveSheet.Name = "Notification XML" Then
        Worksheets("Notification XML").Unprotect
        Application.Run "CheckGroups"
    ElseIf ActiveSheet.Name = "User XML" Then
        Worksheets("User XML").Unprotect
        Application.Run "UpdateGroupsButton"
    End If
ElseIf Me.OptionCombo.value = "Add multiple facility types" Then
    DialogueForm.DialogueBox.Text = "Just click in the Facility Type column where you would " & _
        "like to have multiple groups. From the advanced user worksheet, you will always " & _
        "have the option to add multiple facility types to a notification group."
    
    DialogueForm.Show
    

ElseIf Me.OptionCombo.value = "Import CSV" Then

        Worksheets("ShakeCast Ref Lookup Values").Range("Q2").value = "ImportCSV"
        ProgressForm.Show vbModeless
        
ElseIf Me.OptionCombo.value = "Import Conf" Then

        Worksheets("ShakeCast Ref Lookup Values").Range("Q2").value = "ImportConf"
        ProgressForm.Show vbModeless
    
ElseIf Me.OptionCombo.value = "Access General User Worksheet" Then
        
    If ActiveSheet.Name = "Facility XML" Then
            
        Dim AdvCaption As String
        AdvCaption = FacSheetForm.AdvUser.Caption
        ' hide HAZUS worksheet
        
        FacSheetForm.AdvUser.Caption = "Access Advanced User Worksheet"
        
        Sheets("HAZUS Facility Model Data").Visible = False
        
        ' hide Component and Component Class
        ' hide geometry info
        
        Range("D:E, I:I, AD:AD").EntireColumn.Hidden = True
        
        
        
        ' Change the color of the headers to regular
        
        ChangeColors "Good", Range("A1:AD1"), "Facility"
        ChangeColors "Good", Range("A2:AD2"), "Facility"
        
        ' Change Adv/Gen user caption
        
        Range("A4").Select
        Range("A2").value = "General User"
            
        With Range("A2").Font
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = -0.249977111117893
        End With
        
                
                
        FacSheetForm.DialogueBox.Text = vbTab & vbTab & vbTab & "                           " & _
            "The Facility Spreadsheet" & _
            vbNewLine & _
            vbTab & vbTab & vbTab & "                " & _
            "-------------------------------------------" & _
            vbNewLine & vbNewLine & _
            "This spreadsheet can be used to to convert information about your facilities into " & _
            "XML format, which is readable by the ShakeCast application." & _
            vbNewLine & vbNewLine & _
            "This spreadsheet was made to be completed from left to right. Please hit tab or enter in order to submit your " & _
            "information to the spreadsheet. Some of the fields to " & _
            "the right will automatically fill as you move along. You may feel free to change " & _
            "any values we've supplied for you. HAZUS values cannot be changed from this spreadsheet. They must be " & _
            "changed in the HAZUS spreadsheet which can be viewed from the advanced user spreadsheet." & vbNewLine & vbNewLine & _
            "If you are uncertain of the information you should be providing for a field, " & _
            "click on the ""More Info"" button at the top of that column. These hold information " & _
            "about our expectations for your input for each field." & vbNewLine & vbNewLine & _
            "When deleting a row, it is best to: select a couple cells in that row, highlight the delete drop-down menu, and select ""table rows""." & vbNewLine & vbNewLine & _
            "When you are finished updating your facility information, hit the ""Export XML"" " & _
            "button. You will be prompted to select a save location and name for the file. The default save " & _
            "location is the folder this workbook is currently running in! " & _
            "This file can then easily be uploaded to ShakeCast by dragging and dropping it into " & _
            "the upload page." & vbNewLine & vbNewLine & _
            "It is also possible to export all facility, group, and user information in a single XML file, by clicking " & _
            "the ""Export Master XML"" button. However, this function is not yet accepted by the ShakeCast application."
    
    
    ElseIf ActiveSheet.Name = "Notification XML" Then
    
        Set mySheet = Worksheets("Notification XML")
        
        ' Change Spreadsheet information button to say, "change to general user mode
        FacSheetForm.AdvUser.Caption = "Access Advanced User Worksheet"


        Range("I:M").EntireColumn.Hidden = True
    
        ' Header Color automatically changes
    
        ' Change Adv/Gen user caption
    
        mySheet.Range("A2").value = "General User"
        
        With mySheet.Range("A1", "P2").Interior
            .Color = RGB(196, 215, 155)
        End With
        
        With mySheet.Range("A2").Font
            .Color = RGB(83, 141, 243)
        End With
        
        mySheet.Range("A4").Select
        
        GroupGenInfo
    End If
ElseIf Me.OptionCombo.value = "Access Advanced User Worksheet" Then

    If ActiveSheet.Name = "Facility XML" Then
    
        ' Change Spreadsheet information button to say, "change to general user mode"
    
        FacSheetForm.AdvUser.Caption = "Access General User Worksheet"
    
        ' Unhide HAZUS worksheet
    
        Sheets("HAZUS Facility Model Data").Unprotect
        Sheets("HAZUS Facility Model Data").Visible = True
    
        ' Unhide Component and Component Class
        ' Unhide geometry info
    
        Range("D:E, I:I, AD:AD").EntireColumn.Hidden = False
    
        ' Change the color of the headers
    
        ChangeColors "Advanced", Range("A1", "AE2"), "Facility"
    
        ' Change Adv/Gen user caption
    
        Range("A4").Select
        Range("A2").value = "Advanced User"
        
        With Range("A2").Font
            .Color = RGB(31, 73, 152)
        End With
    
        FacSheetForm.DialogueBox.Text = vbTab & vbTab & vbTab & "               " & _
            "The Advanced Facility Spreadsheet" & _
            vbNewLine & _
            vbTab & vbTab & vbTab & "          " & _
            "---------------------------------------------------" & _
            vbNewLine & vbNewLine & _
            "The advanced user spreadsheet can be used to manually enter components, component " & _
            "classes, and fragility information. " & vbNewLine & vbNewLine & _
            "You can manually set component and component class " & _
            "information in the columns of this spreadsheet." & vbNewLine & vbNewLine & _
            "In order to edit fragility data, click over to the ""HAZUS Facility Model Data"" " & _
            "spreadsheet. From here you can edit the values we have input for specific facility models " & _
            "or create your own facility by adding your own row to the bottom of the sheet."

    ElseIf ActiveSheet.Name = "Notification XML" Then
    
    
        Set mySheet = Worksheets("Notification XML")
        
        
    
        ' Change Spreadsheet information button to say, "change to general user mode
        FacSheetForm.AdvUser.Caption = "Access General User Worksheet"


        Range("I:M").EntireColumn.Hidden = False
    
        ' Change the color of the headers
    
        ChangeColors "Advanced", Range("A1", "P2"), "Group"
    
        ' Change Adv/Gen user caption
    
        mySheet.Range("A2").value = "Advanced User"
        
        With mySheet.Range("A2").Font
            .Color = RGB(31, 73, 152)
        End With
        
        mySheet.Range("A4").Select
        
        GroupAdvInfo
    End If
End If


ExitHandler:
Application.Run "protectWorkbook"
Application.ScreenUpdating = True

If Me.OptionCombo.value <> "Turn Off Data Analysis" Then
    Application.EnableEvents = True
Else
    Application.EnableEvents = False
End If

ActiveSheet.ScrollArea = ""
End Sub

Private Sub Label1_Click()

End Sub

Private Sub OptionCombo_Change()

End Sub

Private Sub UserForm_Click()
    
End Sub

Private Sub UserForm_Initialize()

If ActiveSheet.Name = "Facility XML" Then
    Set mySheet = Worksheets("Facility XML")

    If mySheet.Range("A2").value = "Advanced User" Then
        Me.OptionCombo.AddItem "Create Facility Type"
        Me.OptionCombo.AddItem "Add/Update Fragility Model"
        Me.OptionCombo.AddItem "Create an Attribute"
        Me.OptionCombo.AddItem "Import CSV"
        If Application.EnableEvents = True Then
            Me.OptionCombo.AddItem "Turn Off Data Analysis"
        Else
            Me.OptionCombo.AddItem "Turn On Data Analysis"
        End If
        
        Me.OptionCombo.AddItem "Export XML"
        Me.OptionCombo.AddItem "Export JSON"
        Me.OptionCombo.AddItem "Export Master XML"
        Me.OptionCombo.AddItem "Update Worksheet"
        Me.OptionCombo.AddItem "Unlock Data"
        Me.OptionCombo.AddItem "Clear All Data"
        Me.OptionCombo.AddItem "Access General User Worksheet"
    Else
        Me.OptionCombo.AddItem "Export XML"
        Me.OptionCombo.AddItem "Export JSON"
        Me.OptionCombo.AddItem "Export Master XML"
        Me.OptionCombo.AddItem "Update Worksheet"
        Me.OptionCombo.AddItem "Unlock Data"
        Me.OptionCombo.AddItem "Clear All Data"
        Me.OptionCombo.AddItem "Import CSV"
        Me.OptionCombo.AddItem "Access Advanced User Worksheet"
    End If
    
ElseIf ActiveSheet.Name = "Notification XML" Then
    Set mySheet = Worksheets("Notification XML")
    
    If mySheet.Range("A2").value = "Advanced User" Then
        Me.OptionCombo.AddItem "Add multiple facility types"
        Me.OptionCombo.AddItem "Export XML"
        Me.OptionCombo.AddItem "Export Master XML"
        Me.OptionCombo.AddItem "Update Worksheet"
        Me.OptionCombo.AddItem "Unlock Data"
        Me.OptionCombo.AddItem "Clear All Data"
        Me.OptionCombo.AddItem "Import Conf"
        Me.OptionCombo.AddItem "Access General User Worksheet"
    Else
        Me.OptionCombo.AddItem "Export XML"
        Me.OptionCombo.AddItem "Export Master XML"
        Me.OptionCombo.AddItem "Update Worksheet"
        Me.OptionCombo.AddItem "Unlock Data"
        Me.OptionCombo.AddItem "Clear All Data"
        Me.OptionCombo.AddItem "Import Conf"
        Me.OptionCombo.AddItem "Access Advanced User Worksheet"
    End If
    
ElseIf ActiveSheet.Name = "User XML" Then
    Set mySheet = Worksheets("User XML")
    
    If mySheet.Range("A2").value = "Advanced User" Then
        Me.OptionCombo.AddItem "Export XML"
        Me.OptionCombo.AddItem "Export Master XML"
        Me.OptionCombo.AddItem "Update Worksheet"
        Me.OptionCombo.AddItem "Unlock Data"
        Me.OptionCombo.AddItem "Clear All Data"
        Me.OptionCombo.AddItem "Import CSV"
        Me.OptionCombo.AddItem "Access General User Worksheet"
    Else
        Me.OptionCombo.AddItem "Export XML"
        Me.OptionCombo.AddItem "Export Master XML"
        Me.OptionCombo.AddItem "Update Worksheet"
        Me.OptionCombo.AddItem "Unlock Data"
        Me.OptionCombo.AddItem "Import CSV"
        Me.OptionCombo.AddItem "Clear All Data"
    End If
End If

End Sub

Private Sub WhatsThisButton_Click()
    If Me.OptionCombo.value = "Export XML" Then
        MsgBox "This is the whole point of the ShakeCast Workbook! When you hit Go, we'll export the information you've entered into this worksheet into an XML file. " & _
            "This file can then be uploaded to the ShakeCast application. You can export XML information sheet by sheet with this button, or switch to ""Export Master XML"" to " & _
            "export the information from the entire workbook at once."
            
    ElseIf Me.OptionCombo.value = "Export Master XML" Then
        MsgBox "Export all the information from this workbook into a single file!"
        
    ElseIf Me.OptionCombo.value = "Update Worksheet" Then
        MsgBox "This option will run through all the information in your present worksheet. Use this option if you suspect that " & _
            "some data was not processed properly by the workbook. In the User XML worksheet, this option will also remove all " & _
            "group relationships that are not defined in the Notification XML worksheet."
            
    ElseIf Me.OptionCombo.value = "Unlock Data" Then
        MsgBox "Since this is a protected workbook, it is possible for a user to accidentally end up locking cells " & _
            "that they need access to. Hitting Go will unlock all of the cells that a user can enter data into in this worksheet"
    ElseIf Me.OptionCombo.value = "Clear All Data" Then
        MsgBox "Remove all the information in this worksheet. This option is not undo-able. If you want to test out this option, save " & _
            "the workbook, then hit go. This way, if you don't like the results, you can exit out of the current workbook without saving " & _
            "and open up the version you saved just before you cleared your data!"
            
    ElseIf Me.OptionCombo.value = "Access Advanced User Worksheet" Then
        MsgBox "Access the same worksheet with options that only advanced ShakeCast users need to take advantage of."
            
    ElseIf Me.OptionCombo.value = "Access General User Worksheet" Then
        MsgBox "Get back to the simple worksheet!"
        
    ElseIf Me.OptionCombo.value = "Add multiple facility types" Then
        MsgBox "This option will give you information on how to add multiple facility types to a notificaiton group. " & _
            "When you hit Go, another window will pop up with information, but no changes will be made to your worksheet."
            
    ElseIf Me.OptionCombo.value = "Create Facility Type" Then
        MsgBox "A pop-up window will allow you to create your own facility type!"
        
    ElseIf Me.OptionCombo.value = "Add/Update Fragility Model" Then
        MsgBox "A pop-up window will allow you to input fragility information for a specific model. If you enter the name of " & _
            "a model that already exists in this workbook, the fragility information you input will override the existing defining. If the " & _
            "model does not already exist, we create it for you!"
            
    ElseIf Me.OptionCombo.value = "Create an Attribute" Then
        MsgBox "A pop-up window will allow you to define a new facility attribute. These can be used to group and filter facilities as well as " & _
            "to input additional fragility information about a specific facility."
    
    ElseIf Me.OptionCombo.value = "Turn Off Data Analysis" Then
        MsgBox "If you have many rows of data that you wish to paste into this workbook and you believe the information is " & _
            "valid, you many wish to stop data analysis while you paste. This will save you some time, and you can always " & _
            "run ""Update Worksheet"" when all your data is entered to check the validity of your input."
    
    ElseIf Me.OptionCombo.value = "Turn On Data Analysis" Then
        MsgBox "If you would like your inputs to be checked in real time, turn on data analysis."
    End If
End Sub
