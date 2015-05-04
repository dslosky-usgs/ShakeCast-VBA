VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FragForm 
   Caption         =   "Define a Fragility Model"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18135
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
        If mySheet.Range("A" & i).Value = Me.ModName.Text Then
        
            mySheet.Range("B" & i).Value = Me.ModDesc.Text
            mySheet.Range("C" & i).Value = ""
            mySheet.Range("D" & i).Value = "SYSTEM"
            mySheet.Range("E" & i).Value = "SYSTEM"
            mySheet.Range("F" & i).Value = Me.MetricCombo.Value
            mySheet.Range("G" & i).Value = Me.GreenAlpha.Text
            mySheet.Range("H" & i).Value = Me.GreenBeta.Text
            mySheet.Range("I" & i).Value = Me.MetricCombo.Value
            mySheet.Range("J" & i).Value = Me.YellowAlpha.Text
            mySheet.Range("K" & i).Value = Me.YellowBeta.Text
            mySheet.Range("L" & i).Value = Me.MetricCombo.Value
            mySheet.Range("M" & i).Value = Me.OrangeAlpha.Text
            mySheet.Range("N" & i).Value = Me.OrangeBeta.Text
            mySheet.Range("O" & i).Value = Me.MetricCombo.Value
            mySheet.Range("P" & i).Value = Me.RedAlpha.Text
            mySheet.Range("Q" & i).Value = Me.RedBeta.Text
            mySheet.Range("R" & i).Value = Me.MetricCombo.Value
            mySheet.Range("S" & i).Value = Me.GreyAlpha.Text
            mySheet.Range("T" & i).Value = Me.GreyBeta.Text
            
            MsgBox "You already defined this fragility model, so we just updated its information."
            
            Worksheets("Facility XML").Unprotect
            Application.Run "UpdateFacButton"
            
            Application.Run "protectWorkbook"
            
            Unload Me
            
            Exit Sub
        End If
    Next i

    mySheet.Range("A" & lastRow + 1).Value = Me.ModName.Text
    mySheet.Range("B" & lastRow + 1).Value = Me.ModDesc.Text
    mySheet.Range("C" & lastRow + 1).Value = ""
    mySheet.Range("D" & lastRow + 1).Value = "SYSTEM"
    mySheet.Range("E" & lastRow + 1).Value = "SYSTEM"
    mySheet.Range("F" & lastRow + 1).Value = Me.MetricCombo.Value
    mySheet.Range("G" & lastRow + 1).Value = Me.GreenAlpha.Text
    mySheet.Range("H" & lastRow + 1).Value = Me.GreenBeta.Text
    mySheet.Range("I" & lastRow + 1).Value = Me.MetricCombo.Value
    mySheet.Range("J" & lastRow + 1).Value = Me.YellowAlpha.Text
    mySheet.Range("K" & lastRow + 1).Value = Me.YellowBeta.Text
    mySheet.Range("L" & lastRow + 1).Value = Me.MetricCombo.Value
    mySheet.Range("M" & lastRow + 1).Value = Me.OrangeAlpha.Text
    mySheet.Range("N" & lastRow + 1).Value = Me.OrangeBeta.Text
    mySheet.Range("O" & lastRow + 1).Value = Me.MetricCombo.Value
    mySheet.Range("P" & lastRow + 1).Value = Me.RedAlpha.Text
    mySheet.Range("Q" & lastRow + 1).Value = Me.RedBeta.Text
    mySheet.Range("R" & lastRow + 1).Value = Me.MetricCombo.Value
    mySheet.Range("S" & lastRow + 1).Value = Me.GreyAlpha.Text
    mySheet.Range("T" & lastRow + 1).Value = Me.GreyBeta.Text
    
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
    
        If cell.Value <> Me.ModName.Value Then GoTo NextCell
        If IsEmpty(cell) Then GoTo TheEnd

        Dim rowNum As Integer
        rowNum = cell.row

        Me.ModDesc.Text = mySheet.Range("B" & rowNum).Value
        Me.MetricCombo.Value = mySheet.Range("F" & rowNum).Value
        Me.GreenAlpha = mySheet.Range("G" & rowNum).Value
        Me.GreenBeta = mySheet.Range("H" & rowNum).Value
        Me.YellowAlpha = mySheet.Range("J" & rowNum).Value
        Me.YellowBeta = mySheet.Range("K" & rowNum).Value
        Me.OrangeAlpha = mySheet.Range("M" & rowNum).Value
        Me.OrangeBeta = mySheet.Range("N" & rowNum).Value
        Me.RedAlpha = mySheet.Range("P" & rowNum).Value
        Me.RedBeta = mySheet.Range("Q" & rowNum).Value
        Me.GreyAlpha = mySheet.Range("S" & rowNum).Value
        Me.GreyBeta = mySheet.Range("T" & rowNum).Value
        
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

        Me.ModName.AddItem cell.Value
    
NextCell:
    Next cell
TheEnd:
    
End Sub
