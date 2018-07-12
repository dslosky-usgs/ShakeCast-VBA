VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateFac 
   Caption         =   "Create a Facility Type"
   ClientHeight    =   2300
   ClientLeft      =   -360
   ClientTop       =   -4680
   ClientWidth     =   6740
   OleObjectBlob   =   "CreateFac.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False











Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub CreateButton_Click()
    Set mySheet = Worksheets("ShakeCast Ref Lookup Values")
    
    Dim startRow As Integer
    Dim lastRow As Integer
    
    startRow = 34
    lastRow = mySheet.Cells(Rows.count, "C").End(xlUp).row
    
    For i = startRow To lastRow
        If mySheet.Range("C" & i).value = FacName.Text Then
            mySheet.Range("D" & i).value = FacDesc.Text
            
            MsgBox "You already defined this facility type, so we just updated the facility description!"
            
            Unload Me
            
            Exit Sub
        End If
    Next i
    
    mySheet.Range("C" & lastRow + 1).value = FacName.Text
    mySheet.Range("D" & lastRow + 1).value = FacDesc.Text
    
    MsgBox "Your new facility type """ & FacName.Text & """ has been created. Update your worksheet to see this facility type in drop down menus."
    
    Unload Me
    
End Sub

Private Sub Label2_Click()

End Sub

Private Sub UserForm_Click()

End Sub
