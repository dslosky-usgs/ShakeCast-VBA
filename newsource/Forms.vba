Attribute VB_Name = "Forms"
Sub makeAttCheck()

Set mySheet = Worksheets("ShakeCast Ref Lookup Values")
Set attCell = mySheet.Range("P2")

Dim attList() As String
Dim attCount As Integer

attList = Split(attCell.value, "%")


' get list of attributes associated with this facility
Dim attStr As String
Dim attArr() As String
Dim justAttStr As String
Dim eachAtt() As String

attStr = ActiveCell.value
attArr = Split(attStr, "%")
justAttStr = ""

For Each cont In AttCheckBox.AttFrame.Controls
    If TypeOf cont Is MSForms.Label Then
    
            If justAttStr = "" Then
                justAttStr = "%" & cont.Caption & "%"
            Else
                justAttStr = justAttStr & cont.Caption & "%"
            End If
            
        
    End If
Next cont

' turn list into check boxes
Dim curColumn   As Long
Dim lastRow     As Long
Dim i           As Long
Dim chkBox      As MSForms.CheckBox

For i = 0 To UBound(attList)
    Set chkBox = AttCheckBox.AttFrame.Add("Forms.CheckBox.1", "CheckBox_" & i)
    chkBox.Caption = attList(i)
    chkBox.Left = 5
    chkBox.Top = 5 + (i * 28)
    chkBox.Font.Size = 12
    chkBox.Height = 22
    
    ' select the right checkboxes
    If InStr(justAttStr, "%" & attList(i) & "%") Then
        chkBox.value = True
    End If
    
Next i

' Keep the number of check boxes we've created to reference later in a hidden label

Dim totalHeight As Integer
totalHeight = 5

For Each cont In AttCheckBox.AttFrame.Controls

    If TypeOf cont Is MSForms.CheckBox Then
        totalHeight = totalHeight + 28
    End If

Next cont

If totalHeight > AttCheckBox.AttFrame.Height Then
    AttCheckBox.AttFrame.ScrollHeight = totalHeight
ElseIf totalHeight = 5 Then
    Set txtbox = AttCheckBox.AttFrame.Add("Forms.TextBox.1", "TextBox_1")
    txtbox.Left = 5
    txtbox.Top = 5
    txtbox.Font.Size = 12
    txtbox.Height = 230
    txtbox.Width = 190
    txtbox.WordWrap = True
    txtbox.MultiLine = True
    txtbox.Text = "It looks like you haven't defined any facility attributes! To define facility attributes, hit ""Cancel"" " & _
                    "and click ""Manage Attributes"". From here, you can create new attributes or delete ones you aren't " & _
                    "using anymore. You will have to return to this window in order to associate an attribute to a specific " & _
                    "facility."
    

End If

End Sub

Sub makeFacTypesForm()
Set mySheet = Worksheets("ShakeCast Ref Lookup Values")

Dim lastFac As Integer
lastFac = mySheet.Cells(Rows.count, "C").End(xlUp).row
Set FacTypeCells = Worksheets("ShakeCast Ref Lookup Values").Range("C1:C" & lastFac)

Dim facList As Variant
facList = FacTypeCells.value


' turn list into check boxes
Dim curColumn   As Long
Dim lastRow     As Long
Dim i           As Long
Dim chkBox      As MSForms.CheckBox

For i = 1 To UBound(facList)
    Set chkBox = AttCheckBox.AttFrame.Add("Forms.CheckBox.1", "CheckBox_" & i)
    chkBox.Caption = facList(i, 1)
    chkBox.Left = 5
    chkBox.Top = 5 + ((i - 1) * 28)
    chkBox.Font.Size = 12
    chkBox.Height = 22
    
    ' select the right checkboxes
    If InStr(ActiveCell.value, facList(i, 1)) Then
        chkBox.value = True
    End If
    
Next i

' Keep the number of check boxes we've created to reference later in a hidden label

Dim totalHeight As Integer
totalHeight = 5

For Each cont In AttCheckBox.AttFrame.Controls

    If TypeOf cont Is MSForms.CheckBox Then
        totalHeight = totalHeight + 28
    End If

Next cont

If totalHeight > AttCheckBox.AttFrame.Height Then
    AttCheckBox.AttFrame.ScrollHeight = totalHeight
End If

AttCheckBox.Caption = "Select Facility Types"

End Sub




