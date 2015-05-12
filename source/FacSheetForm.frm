VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FacSheetForm 
   Caption         =   "Facility Worksheet Information"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10350
   OleObjectBlob   =   "FacSheetForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FacSheetForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub CommandButton1_Click()

End Sub

Private Sub AdvUser_Click()

ActiveSheet.Unprotect

If ActiveSheet.Range("A1").value = "Facility Worksheet" Then

    Dim AdvCaption As String
    AdvCaption = FacSheetForm.AdvUser.Caption
    
    If AdvCaption = "Access Advanced User Worksheet" Then
    
    ' Change Spreadsheet information button to say, "change to general user mode"
    
        FacSheetForm.AdvUser.Caption = "Access General User Worksheet"
    
    ' Unhide HAZUS worksheet
    
        Sheets("HAZUS Facility Model Data").Unprotect
        Sheets("HAZUS Facility Model Data").Visible = True
    
    ' Unhide Component and Component Class
    ' Unhide geometry info
    
        Range("D:E, I:I, AE:AE").EntireColumn.Hidden = False
    
    ' Change the color of the headers
    
        ChangeColors "Advanced", Range("A1", "AE2"), "Facility"
    
    ' Change Adv/Gen user caption
    
        Range("A2").Select
        Selection.value = "Advanced User"
        
        With Selection.Font
            .Color = RGB(31, 73, 152)
        End With
    
        Application.Run "FacAdvInfo"
        
    
    Else
    ' Change Spreadsheet information button to say, "change to general user mode"
    
        FacSheetForm.AdvUser.Caption = "Access Advanced User Worksheet"
        
    ' hide HAZUS worksheet
    
        Sheets("HAZUS Facility Model Data").Visible = False
    
    ' hide Component and Component Class
    ' hide geometry info
    
        Range("D:E, I:I, AE:AE").EntireColumn.Hidden = True
    
    
    
    ' Change the color of the headers to regular
    
        ChangeColors "Good", Range("A1:AD1"), "Facility"
        ChangeColors "Good", Range("A2:AD2"), "Facility"
    
    ' Change Adv/Gen user caption
    
        Range("A2").Select
        Selection.value = "General User"
        
        With Selection.Font
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = -0.249977111117893
        End With
    
            
        Application.Run "FacGenInfo"

    End If

ElseIf ActiveSheet.Range("A1").value = "Notification Worksheet" Then

    Set mySheet = Worksheets("Notification XML")

    If Me.AdvUser.Caption = "Access Advanced User Worksheet" Then
        
        GroupAdvInfo
    
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
        
        

    Else
        Set mySheet = Worksheets("Notification XML")
    
        GroupGenInfo
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
        
        
        
    End If
        
End If

ActiveSheet.Range("A4").Select
ActiveSheet.Protect AllowFormattingCells:=True, AllowDeletingRows:=True, AllowInsertingRows:=True, UserInterfaceOnly:=True

End Sub

Private Sub CloseButton_Click()

    FacSheetForm.DialogueBox.SetFocus
    FacSheetForm.DialogueBox.CurLine = 0
    
    Unload Me

End Sub

Private Sub DialogueBox_Change()

End Sub

Private Sub DialogueBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)

End Sub


Private Sub UserForm_Initialize()

End Sub

Private Sub UserForm_Activate()
    FacSheetForm.DialogueBox.CurLine = 0
End Sub

Private Sub UserForm_Terminate()
    FacSheetForm.AdvUser.Visible = True
End Sub
