VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FragForm 
   Caption         =   "Define a Fragility Model"
   ClientHeight    =   3100
   ClientLeft      =   -40
   ClientTop       =   -2840
   ClientWidth     =   18140
   OleObjectBlob   =   "FragForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FragForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Private Sub ComboBox1_Change()

End Sub

Private Sub AddButton_Click()
    Set mySheet = Worksheets("HAZUS Facility Model Data")
    
    Dim startRow As Integer
    Dim lastRow As Integer
    
    startRow = 2
    lastRow = mySheet.Cells(Rows.count, "A").End(xlUp).row
    
    For i = startRow To lastRow
        If mySheet.Range("A" & i).value = Me.ModName.Text Then
        
            mySheet.Range("B" & i).value = Me.ModDesc.Text
            mySheet.Range("C" & i).value = ""
            mySheet.Range("D" & i).value = "SYSTEM"
            mySheet.Range("E" & i).value = "SYSTEM"
            mySheet.Range("F" & i).value = Me.MetricCombo.value
            mySheet.Range("G" & i).value = Me.GreenAlpha.Text
            mySheet.Range("H" & i).value = Me.GreenBeta.Text
            mySheet.Range("I" & i).value = Me.MetricCombo.value
            mySheet.Range("J" & i).value = Me.YellowAlpha.Text
            mySheet.Range("K" & i).value = Me.YellowBeta.Text
            mySheet.Range("L" & i).value = Me.MetricCombo.value
            mySheet.Range("M" & i).value = Me.OrangeAlpha.Text
            mySheet.Range("N" & i).value = Me.OrangeBeta.Text
            mySheet.Range("O" & i).value = Me.MetricCombo.value
            mySheet.Range("P" & i).value = Me.RedAlpha.Text
            mySheet.Range("Q" & i).value = Me.RedBeta.Text
            mySheet.Range("R" & i).value = Me.MetricCombo.value
            mySheet.Range("S" & i).value = Me.GreyAlpha.Text
            mySheet.Range("T" & i).value = Me.GreyBeta.Text
            
            MsgBox "You already defined this fragility model, so we just updated its information."
            
            Worksheets("Facility XML").Unprotect
            Application.Run "UpdateFacButton"
            
            Application.Run "protectWorkbook"
            
            Unload Me
            
            Exit Sub
        End If
    Next i

    mySheet.Range("A" & lastRow + 1).value = Me.ModName.Text
    mySheet.Range("B" & lastRow + 1).value = Me.ModDesc.Text
    mySheet.Range("C" & lastRow + 1).value = ""
    mySheet.Range("D" & lastRow + 1).value = "SYSTEM"
    mySheet.Range("E" & lastRow + 1).value = "SYSTEM"
    mySheet.Range("F" & lastRow + 1).value = Me.MetricCombo.value
    mySheet.Range("G" & lastRow + 1).value = Me.GreenAlpha.Text
    mySheet.Range("H" & lastRow + 1).value = Me.GreenBeta.Text
    mySheet.Range("I" & lastRow + 1).value = Me.MetricCombo.value
    mySheet.Range("J" & lastRow + 1).value = Me.YellowAlpha.Text
    mySheet.Range("K" & lastRow + 1).value = Me.YellowBeta.Text
    mySheet.Range("L" & lastRow + 1).value = Me.MetricCombo.value
    mySheet.Range("M" & lastRow + 1).value = Me.OrangeAlpha.Text
    mySheet.Range("N" & lastRow + 1).value = Me.OrangeBeta.Text
    mySheet.Range("O" & lastRow + 1).value = Me.MetricCombo.value
    mySheet.Range("P" & lastRow + 1).value = Me.RedAlpha.Text
    mySheet.Range("Q" & lastRow + 1).value = Me.RedBeta.Text
    mySheet.Range("R" & lastRow + 1).value = Me.MetricCombo.value
    mySheet.Range("S" & lastRow + 1).value = Me.GreyAlpha.Text
    mySheet.Range("T" & lastRow + 1).value = Me.GreyBeta.Text
    
    MsgBox "Your new fragility model, """ & Me.ModName.Text & """, has been created. Update your worksheet in order to see this fragility model in drop down menus."
    Unload Me

End Sub

Private Sub Label29_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub ModName_Change()



End Sub

Private Sub ModName_Click()

    Set mySheet = Worksheets("HAZUS Facility Model Data")
    For Each cell In mySheet.Range("A:A")
    
        If cell.value <> Me.ModName.value Then GoTo NextCell
        If IsEmpty(cell) Then GoTo TheEnd

        Dim rowNum As Integer
        rowNum = cell.row

        Me.ModDesc.Text = mySheet.Range("B" & rowNum).value
        Me.MetricCombo.value = mySheet.Range("F" & rowNum).value
        Me.GreenAlpha = mySheet.Range("G" & rowNum).value
        Me.GreenBeta = mySheet.Range("H" & rowNum).value
        Me.YellowAlpha = mySheet.Range("J" & rowNum).value
        Me.YellowBeta = mySheet.Range("K" & rowNum).value
        Me.OrangeAlpha = mySheet.Range("M" & rowNum).value
        Me.OrangeBeta = mySheet.Range("N" & rowNum).value
        Me.RedAlpha = mySheet.Range("P" & rowNum).value
        Me.RedBeta = mySheet.Range("Q" & rowNum).value
        Me.GreyAlpha = mySheet.Range("S" & rowNum).value
        Me.GreyBeta = mySheet.Range("T" & rowNum).value
        
        GoTo TheEnd
        
NextCell:
    Next cell
TheEnd:

End Sub

Private Sub ModName_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    


End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    
    Me.MetricCombo.AddItem "PGA"
    Me.MetricCombo.AddItem "MMI"
    Me.MetricCombo.AddItem "PGV"
    Me.MetricCombo.AddItem "PSA03"
    Me.MetricCombo.AddItem "PSA10"
    Me.MetricCombo.AddItem "PSA30"
    
    For Each cell In Worksheets("HAZUS Facility Model Data").Range("A:A")
    
        If cell.row = 1 Then GoTo NextCell
        If IsEmpty(cell) Then GoTo TheEnd

        Me.ModName.AddItem cell.value
    
NextCell:
    Next cell
TheEnd:
    
End Sub
