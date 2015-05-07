VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Activate()
    
    If ActiveSheet.Name = "Welcome" Then
    
        Worksheets("Welcome").Unprotect
        Worksheets("Welcome").Rows("1:36").Select
        
        ActiveWindow.Zoom = True
        
        Application.Run "protectWorkbook"
        
        Worksheets("Welcome").Range("Z200").Activate
    End If

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    
    Application.ScreenUpdating = False
    
        Sheet1.Select
        Sheet1.Range("A4").Select
        Sheet2.Select
        Sheet2.Range("A4").Select
        Sheet3.Select
        Sheet3.Range("A4").Select
        
        Sheet9.Activate
        
    Application.ScreenUpdating = True
    
    ' we don't want to save on exit because this is really the only way to undo macro changes
    ' Me.Save
    
End Sub

Private Sub Workbook_Open()

    Set welcomeSheet = Worksheets("Welcome")
'    welcomeSheet.Activate
'
    welcomeSheet.ScrollArea = "A1"
'
'    welcomeSheet.Range("Z35").Activate
    
    For Each Sheet In Application.ThisWorkbook.Sheets
        Sheet.Unprotect
    Next Sheet
    
    Application.Run "protectWorkbook"
    
    Application.EnableEvents = True
    
End Sub





