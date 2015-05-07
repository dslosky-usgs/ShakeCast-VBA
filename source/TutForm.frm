VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TutForm 
   Caption         =   "ShakeCase Workbook Tutorial"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11325
   OleObjectBlob   =   "TutForm.frx":0000
End
Attribute VB_Name = "TutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Back_Click()

    Dim tutNum As Integer
    Dim tutDec As Integer
    Dim tutInfo As Integer

    tutNum = Me.SecNum.Caption
    tutDec = Me.SecDec.Caption
    tutInfo = Me.InfoClick.Caption
    
    tutInfo = tutInfo - 1
    
    If tutInfo = -1 Then
        tutDec = tutDec - 1
        tutInfo = 0
        
        If tutDec = -1 Then
            tutNum = tutNum - 1
            tutDec = 0
            If tutNum = -1 Then
                tutNum = 0
            End If
        End If
    End If

    Me.InfoClick.Caption = tutInfo
    Me.SecNum.Caption = tutNum
    Me.SecDec.Caption = tutDec
    tutWindow

    tutCont


End Sub

Private Sub Continue_Click()
    
    Dim tutNum As Integer
    Dim tutDec As Integer
    Dim tutInfo As Integer

    tutNum = Me.SecNum.Caption
    tutDec = Me.SecDec.Caption
    tutInfo = Me.InfoClick.Caption
    
    tutInfo = tutInfo + 1
    
    
    If tutNum < 1 Or (tutNum = 1 And tutDec = 0) Or (tutNum = 1 And tutDec = 1 And tutInfo = 12) Or _
        (tutNum = 1 And tutDec = 2 And tutInfo = 2) Or (tutNum = 1 And tutDec = 3 And tutInfo = 3) Or _
        (tutNum = 2 And tutDec = 0) Or (tutNum = 2 And tutDec = 1 And tutInfo = 14) Or (tutNum = 2 And tutDec = 2 And tutInfo = 5) Or _
        (tutNum = 3 And tutDec = 0) Or (tutNum = 3 And tutDec = 1 And tutInfo = 8) Or (tutNum = 3 And tutDec = 2 And tutInfo = 3) Then
        
        ' next section
        tutSec
        ' make the tutorial window
        tutWindow
        
        tutInfo = 0
    End If
    
    Me.InfoClick.Caption = tutInfo
    
    tutCont
    
    Me.DialogueBox.SetFocus
    Me.DialogueBox.CurLine = 0
    
End Sub

Private Sub Done_Click()
    Unload Me
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub Label_0_Click()

Application.Run "copyFirstRows"

Me.SecNum.Caption = 1
Me.SecDec.Caption = 0
Me.InfoClick.Caption = 0

tutWindow

If ActiveSheet.Name <> "Facility XML" Then
    Worksheets("Facility XML").Activate
End If

tutCont



End Sub

Private Sub Label_1_Click()

Application.Run "copyFirstRows"

Me.SecNum.Caption = 1
Me.SecDec.Caption = 1
Me.InfoClick.Caption = 0

tutWindow

If ActiveSheet.Name <> "Facility XML" Then
    Worksheets("Facility XML").Activate
End If

tutCont

End Sub

Private Sub Label_2_Click()

Application.Run "copyFirstRows"

Me.SecNum.Caption = 1
Me.SecDec.Caption = 2
Me.InfoClick.Caption = 0

tutWindow

If ActiveSheet.Name <> "Facility XML" Then
    Worksheets("Facility XML").Activate
End If

tutCont

End Sub

Private Sub Label_3_Click()

Application.Run "copyFirstRows"

Me.SecNum.Caption = 1
Me.SecDec.Caption = 3
Me.InfoClick.Caption = 0

tutWindow

tutCont

End Sub

Private Sub Label_4_Click()

Application.Run "copyFirstRows"

Me.SecNum.Caption = 2
Me.SecDec.Caption = 0
Me.InfoClick.Caption = 0

tutWindow

If ActiveSheet.Name <> "Notification XML" Then
    Worksheets("Notification XML").Activate
End If

tutCont

End Sub

Private Sub Label_5_Click()

Application.Run "copyFirstRows"

Me.SecNum.Caption = 2
Me.SecDec.Caption = 1
Me.InfoClick.Caption = 0

tutWindow

If ActiveSheet.Name <> "Notification XML" Then
    Worksheets("Notification XML").Activate
End If

tutCont

End Sub

Private Sub Label_6_Click()

Application.Run "copyFirstRows"

Me.SecNum.Caption = 2
Me.SecDec.Caption = 2
Me.InfoClick.Caption = 0

tutWindow

If ActiveSheet.Name <> "User XML" Then
    Worksheets("User XML").Activate
End If

tutCont

End Sub

Private Sub Label_7_Click()

Application.Run "copyFirstRows"

Me.SecNum.Caption = 3
Me.SecDec.Caption = 0
Me.InfoClick.Caption = 0

tutWindow

If ActiveSheet.Name <> "User XML" Then
    Worksheets("User XML").Activate
End If

tutCont

End Sub

Private Sub Label_8_Click()

Application.Run "copyFirstRows"

Me.SecNum.Caption = 3
Me.SecDec.Caption = 1
Me.InfoClick.Caption = 0

tutWindow

If ActiveSheet.Name <> "User XML" Then
    Worksheets("User XML").Activate
End If

tutCont

End Sub


Private Sub Label_9_Click()

Application.Run "copyFirstRows"

Me.SecNum.Caption = 3
Me.SecDec.Caption = 2
Me.InfoClick.Caption = 0

tutWindow

If ActiveSheet.Name <> "User XML" Then
    Worksheets("User XML").Activate
End If

tutCont

End Sub

Private Sub SecFrame_Click()

End Sub

Private Sub Skip_Click()
    
If Me.SecNum.Caption < 1 Then
    Me.SecNum.Caption = 1
    Worksheets("Facility XML").Activate
ElseIf Me.SecNum.Caption < 2 Then
    Me.SecNum.Caption = 2
    Worksheets("Notification XML").Activate
ElseIf Me.SecNum.Caption < 3 Then
    Me.SecNum.Caption = 3
    Worksheets("User XML").Activate
End If

Me.InfoClick.Caption = 0
Me.SecDec.Caption = 0
    
tutWindow
tutCont
    
End Sub

Private Sub UserForm_Activate()
    
End Sub

Private Sub UserForm_Click()



End Sub

Private Sub UserForm_Deactivate()

End Sub

Private Sub UserForm_Initialize()


End Sub

Private Sub UserForm_Terminate()

Dim tutRow As Integer
Dim copyRow As Integer
Set copySheet = Worksheets("ShakeCast Ref Lookup Values")

If copySheet.Range("A99").Value = "yes" Then
    Set mySheet = Worksheets("Facility XML")
    tutRow = 4
    copyRow = 100
    mySheet.Range("A" & tutRow & ":" & "AF" & tutRow).Value = _
        copySheet.Range("A" & copyRow & ":" & "AF" & copyRow).Value
    
    copySheet.Rows(copyRow & ":" & copyRow).EntireRow.Clear
    
    copySheet.Range("A99").Value = "no"
End If

If copySheet.Range("A199").Value = "yes" Then
    Set mySheet = Worksheets("Notification XML")
    tutRow = 4
    copyRow = 200

    mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Value = _
        copySheet.Range("A" & copyRow & ":" & "Q" & copyRow + 9).Value
        
    copySheet.Rows(copyRow & ":" & copyRow + 9).EntireRow.Clear

    copySheet.Range("A199").Value = "no"
End If

If copySheet.Range("A299").Value = "yes" Then
    Set mySheet = Worksheets("User XML")
    tutRow = 4
    copyRow = 300

    mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 1).Value = _
        copySheet.Range("A" & copyRow & ":" & "Q" & copyRow + 1).Value
        
    copySheet.Rows(copyRow & ":" & copyRow + 9).EntireRow.Clear

    copySheet.Range("A299").Value = "no"
End If

Worksheets("Welcome").Activate

End Sub
