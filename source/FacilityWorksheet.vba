Attribute VB_Name = "FacilityWorksheet"
'' CheckFacilities
'' Daniel Slosky
'' Last Update: 3/3/2015
''
'' This sub will run each time the user spreadsheet is edited. It will check the active row
'' as well as the row above the active row. One of these will be the row in which the edit took place
'' unless the user entered a value and clicked somewhere else on the worksheet. The entered user info
'' will then be validated and default data will be supplied
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Sub CheckFacilities(ByVal target As Range)

' set our working spreadsheet to the User XML sheet
Set mySheet = Worksheets("Facility XML")

' keep the worksheet looking at the same cell while code executes
mySheet.ScrollArea = ActiveCell.Address

' make variable to track the active cell when we started this subroutine
Dim startActive As Range
Set startActive = ActiveCell

Dim checkRow As Long
checkRow = target.row

Dim lastRow As Long
lastRow = mySheet.Cells(Rows.count, "A").End(xlUp).row ' where we stop!

    ' skip the headers. Since the active row goes first, if the uprow is a header, the active row
    ' will still get evaluated
    
    If checkRow < 4 Then GoTo TheEnd

    ' for each row, check for user input
    ' check for no user input
    If WorksheetFunction.CountBlank(mySheet.Range("A" & checkRow, "C" & checkRow)) = 3 And _
            WorksheetFunction.CountBlank(mySheet.Range("F" & checkRow, "H" & checkRow)) = 3 And _
            WorksheetFunction.CountBlank(mySheet.Range("J" & checkRow, "M" & checkRow)) = 4 Then

    ' The user input for the row is empty, so lets just clear the row
        mySheet.Range("A" & checkRow, "AE" & checkRow).Clear
        mySheet.Range("A" & checkRow, "AE" & checkRow).Locked = False
        
        ChangeColors "Good", mySheet.Range("A" & checkRow, "AD" & checkRow), "Facility"
        
        GoTo NextRow
    
    ' check for missing required input
    ElseIf IsEmpty(mySheet.Range("A" & checkRow)) Or _
            IsEmpty(mySheet.Range("B" & checkRow)) Or _
            IsEmpty(mySheet.Range("F" & checkRow)) Or _
            IsEmpty(mySheet.Range("J" & checkRow)) Or _
            IsEmpty(mySheet.Range("K" & checkRow)) Then

        ChangeColors "Bad", mySheet.Range("A" & checkRow, "AD" & checkRow), "Facility"
        
        mySheet.Range("AE" & checkRow).value = "Bad"
        
    ' otherwise the row is good!
    Else
        ' all the required fields are filled
        ChangeColors "Good", mySheet.Range("A" & checkRow, "AE" & checkRow), "Facility"
        mySheet.Range("AE" & checkRow).value = "Good"
    End If


    ' change HAZUS info when column 14 is altered
    If target.Column = 13 Then
        Application.Run "fillHazus", target
    End If
    
    ' make drop down data validation when the facility ID is completed, but not above row 5000 to avoid breaking software
    If target.Column = 1 Then
'        Application.Run "facDropDowns", target
    End If
    
    ' if ANY user input, supply autodata
    FillFacilityInfo target
        

NextRow:


    If WorksheetFunction.CountBlank(mySheet.Range("A" & checkRow, "C" & checkRow)) = 3 And _
            WorksheetFunction.CountBlank(mySheet.Range("F" & checkRow, "H" & checkRow)) = 3 And _
            WorksheetFunction.CountBlank(mySheet.Range("J" & checkRow, "M" & checkRow)) = 4 Then


        ' The user input for the row is empty, so lets just clear the row
        mySheet.Range("A" & checkRow, "AE" & checkRow).Clear
        mySheet.Range("A" & checkRow, "AE" & checkRow).Locked = False

        ChangeColors "Good", mySheet.Range("A" & checkRow, "AE" & checkRow), "Facility"
    End If


TheEnd:

' set the scroll area back to the used range
'mySheet.ScrollArea = "A1:AE" & (LastRow + 50)
mySheet.ScrollArea = ""
End Sub


Sub FillFacilityInfo(ByVal target As Range)

Dim checkRow As Long
checkRow = target.row

Set mySheet = Worksheets("Facility XML")

' Fill info that user will not edit
mySheet.Range("C" & checkRow).value = FillFacility(mySheet.Range("B" & checkRow), mySheet.Range("C" & checkRow))
' mySheet.Range("J" & checkRow).Value = ManLatLong(mySheet.Range("K" & checkRow), mySheet.Range("L" & checkRow))
' mySheet.Range("I" & checkRow).Value = GeomType(mySheet.Range("J" & checkRow))
mySheet.Range("I" & checkRow).value = GeomType(ManLatLong(mySheet.Range("J" & checkRow), mySheet.Range("K" & checkRow)))

mySheet.Range("J" & checkRow).HorizontalAlignment = xlCenter
mySheet.Range("K" & checkRow).HorizontalAlignment = xlCenter

' fill component and component class if left empty
If IsEmpty(mySheet.Range("D" & checkRow)) And Not IsEmpty(mySheet.Range("A" & checkRow)) Then
    mySheet.Range("D" & checkRow).value = FillSystem(mySheet.Range("A1"))
ElseIf IsEmpty(mySheet.Range("A" & checkRow)) Then
    mySheet.Range("D" & checkRow).value = Empty
End If

If IsEmpty(mySheet.Range("E" & checkRow)) And Not IsEmpty(mySheet.Range("A" & checkRow)) Then
    mySheet.Range("E" & checkRow).value = FillSystem(mySheet.Range("A1"))
ElseIf IsEmpty(mySheet.Range("A" & checkRow)) Then
    mySheet.Range("E" & checkRow).value = Empty
End If



End Sub

Private Sub fillHazus(target As Range)

Set mySheet = Worksheets("Facility XML")

Dim checkRow As Long
checkRow = target.row

' fill HAZUS information
'If target.column = 14 Then
    
    If IsEmpty(target) Then
    
        mySheet.Range("N" & checkRow, "AC" & checkRow).ClearContents
        Exit Sub
    End If
    Set hazSheet = Worksheets("HAZUS Facility Model Data")
    
    Dim lastHazRow As Integer
    lastHazRow = hazSheet.Cells(Rows.count, "A").End(xlUp).row
    
    For Each hazMod In hazSheet.Range("A1:A" & lastHazRow)
    
        If target.value = hazMod.value Then
            mySheet.Range("N" & checkRow).value = hazSheet.Range("B" & hazMod.row).value
            mySheet.Range("O" & checkRow, "AC" & checkRow).value = hazSheet.Range("F" & hazMod.row, "T" & hazMod.row).value
        
            ' mySheet.Range("O" & checkRow).HorizontalAlignment = xlCenter
            mySheet.Range("N" & checkRow, "AC" & checkRow).HorizontalAlignment = xlCenter
        
            Exit Sub
        End If
    
    Next hazMod
    
    
'End If



End Sub

Private Sub facDropDowns(ByVal target As Range)

Dim checkRow As Long
checkRow = target.row

Set mySheet = Worksheets("Facility XML")

' create User Type and Group drop down menus
Set FacType = mySheet.Range("B" & checkRow)
Set ModType = mySheet.Range("M" & checkRow)

Dim FacTypes() As String
Dim ModTypes() As String


Set LookUpSheet = Worksheets("ShakeCast Ref Lookup Values")
Dim lastFac As Long
lastFac = LookUpSheet.Cells(Rows.count, "C").End(xlUp).row
Set FacTypeCells = Worksheets("ShakeCast Ref Lookup Values").Range("C1:C" & lastFac)

Set hazSheet = Worksheets("Hazus Facility Model Data")
Dim lastHaz As Long

lastHaz = hazSheet.Cells(Rows.count, "A").End(xlUp).row

Set ModTypeCells = hazSheet.Range("A2:A" & lastHaz)

If target.Column = 1 Then
    With FacType.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="='" & LookUpSheet.Name & "'!" & FacTypeCells.Address
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Facility Type"
        .ErrorTitle = ""
        .InputMessage = "Please select a facility type from the drop-down list"
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End If

If target.Column = 1 Then
    With ModType.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="='" & hazSheet.Name & "'!" & ModTypeCells.Address
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "HAZUS Model Type"
        .ErrorTitle = ""
        .InputMessage = "Please select a model type from the drop-down list"
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End If

End Sub


'' FacilityXML
'' Daniel Slosky
'' Last Updated: 2/18/2016
'' Creates an XML document that holds all the information the user input into the facility spreadsheet. Similar
'' programs will be written for the user and group spreadsheets.
''
'' This programs essentially looks at each piece of data within the spreadsheet, the looks backwards to determine
'' its parents. For each cell it thinks: Is this data? Okay, it is (Or nope, next row). Does it have a first level
'' parent? If so, does it have a second level parent? Then places the data surrounded by its header in the
'' appropriate place.
''
'' It looks like:
''
''  <Root>
''      <Header>DATA</Header>
''      <Parent1>
''          <Header>DATA</Header>
''      </Parent1>
''      <Parent1>
''          <Parent2>
''              <Header>DATA</Header>
''          </Parent2>
''      </Parent1>
''  </Root>
''
'' The same structure with the terminology I used in this program looks like:
''
''  <Root>
''      <Header>DATA</Header>
''      <Field>
''          <Header>DATA</Header>
''      </Field>
''      <Field>
''          <Subfield>
''              <Header>DATA</Header>
''          </Subfield>
''      </Field>
''  </Root>
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FacilityXML(master As String, _
                    Optional ByVal docCount As Integer = 0, _
                    Optional ByVal overFlowCount As Integer = 0, _
                    Optional ByVal docMax As Integer = 15000, _
                    Optional ByVal docStr As String = "")
                                        
'On Error GoTo XMLFinish
'On Error Resume Next

Dim docArr() As String

Dim getOS As String
getOS = Application.OperatingSystem

If master = "Master" Then
    docArr = Split(docStr, ",")
    GoTo MasterSkip1
End If




Application.EnableEvents = False
Application.ScreenUpdating = False
ActiveSheet.Unprotect

Close #2


' We now get XML info from the worksheet FacilityXMLexport, and we have to move all the info over there
Dim xmlInfo() As Variant
xmlInfo = Application.Run("FacXMLTable")                            ' This function populates the XML table in the HIDDEN
                                                                    ' spreadsheet FacilityXMLexport


'                          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                 OPEN XML FILE FOR WRITING, AND DETERMINE THE START AND END CELLS TO BE EXAMINED

' open file location
Dim dir As String
dir = Application.ActiveWorkbook.Path

Dim docNum As Double

docNum = xmlInfo(0) / docMax

If WorksheetFunction.Ceiling(docNum, 1) = docNum Then
    docMax = docMax - 1
    docNum = infoAcc / docMax
End If

If docNum < 1 Then
    docNum = 1
    docStr = "FacilityXML.xml"
Else
    docNum = Application.WorksheetFunction.Ceiling(docNum, 1)
    docStr = "FacilityXML1.xml"
    For i = 2 To docNum
        docStr = docStr & "," & "FacilityXML" & i & ".xml"
    Next i
End If

ExportXML.FileDest.Text = dir
ExportXML.FileName = docStr

DoEvents
ExportXML.Show
DoEvents

ProgressForm.ProcessName.Caption = "Exporting XML"
ProgressForm.ProgressLabel.Width = 0
ProgressForm.ProgressFrame.Caption = "0" & "%"

DoEvents

docArr = Split(ExportXML.FileName, ",")
overFlowCount = 0

If UBound(docArr) < 0 Then
    ReDim docArr(0 To 0)
    docArr(0) = ""
End If

If InStr(getOS, "Windows") = 0 Then

    XMLPath = ExportXML.FileDest.Text & ":" & docArr(docCount)       ' We save the .xml here for Macs
    
Else

    XMLPath = ExportXML.FileDest.Text & "\" & docArr(docCount)      ' Save XML for PC
End If


If XMLPath = "\" Or XMLPath = ":" Then
    GoTo XMLFinish
End If

Open XMLPath For Output As #2



MasterSkip1:

Set XMLSheet = Worksheets("FacilityXMLexport")

Dim startCol As Integer                      ' Figure out which rows and columns we want to sort through
Dim startRow As Integer
Dim endCol As Integer
Dim endRow As Long


startCol = 1                                 ' The first column that holds XML info
startRow = 4                                 ' The First Row with XML info


' The last column we will look at
endCol = XMLSheet.Cells(1, Columns.count).End(xlToLeft).Column
'EndCol = 30

' This will pull cells with formulas and no values, but we can filter those out!
endRow = XMLSheet.Cells(Rows.count, "B").End(xlUp).row + 1




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''' BEGIN XML '''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim printStr As String
If master = "Master" Then GoTo MasterSkip2
printStr = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbNewLine
MasterSkip2:

printStr = printStr & "<FacilityTable>"
           
Print #2, printStr


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''' Write XML '''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' This section is also broken up into multiple parts... some of which should probably be their own individual
' functions



Dim TestCell As Range         ' Used to decide if the row is empty or not
Dim CurCell As Range          ' The current cell we are iterating over
Dim Field As Range            ' An element with one layer of children that may also have another layer of children
Dim Subf As Range             ' The layer of children under the Field
Dim Head As Range             ' The title of the tags that actually surround the data
Dim oldField As String        ' keeping track of the last field and subfield allow us to determine when to open
Dim oldSubf As String         ' tags
Dim endField As String

Dim NextField As Range        ' Keeping track of the next field and subfield, we can determine when to close tags
Dim NextSubf As Range

oldField = ""
oldSubf = ""
endField = "No"

' when to progress the progress bar
' set up progress bar
Dim progressCount As Long
Dim progressWhen As Long
Dim pcntDone As Double

progressCount = 0
progressWhen = endRow * 0.01
pcntDone = 0
ProgressForm.ProcessName.Caption = "Exporting Facility XML"

' create variables to deal with attribute fields
Dim attStr As String
Dim attArr() As String
Dim eachAtt() As String

attStr = ""

' This loop will run through all the cells in the XML region of the document.
Dim row As Long
Dim indent As Integer
row = 4                   ' The first row that holds data
indent = 0                ' Keeps track of how many parents each tag has
Do While row < endRow


    If progressCount > progressWhen Then
        
        pcntDone = ((row - 4) / (endRow - 3))
        
        ProgressForm.ProgressLabel.Width = pcntDone * ProgressForm.MaxWidth.Caption
        ProgressForm.ProgressFrame.Caption = Round(pcntDone * 100, 0) & "%"
        
        DoEvents
        progressCount = 0
        
    End If

    Set TestCell = XMLSheet.Cells(row, 1) ' This is the first XML cell, which should be filled for any valid entry
    
    If TestCell.value = Empty Then         ' If it IS empty, then we just skip this row
        GoTo ContinueLoop
        
    Else
        
        ' Since there IS data for this row, we start a new entry in the XML page
        printStr = vbTab & "<FacilityRow>" & vbNewLine
        
        ' Now we can run through all the columns in the row. Row 32 is the first one to hold XML data
        For col = 1 To endCol
        
            ' Assign the field, subfield, and header for the current cell.
            Set Field = XMLSheet.Cells(1, col)
            Set Subf = XMLSheet.Cells(2, col)
            Set Head = XMLSheet.Cells(3, col)
            Set CurCell = XMLSheet.Cells(row, col)
            
            
            If Field.value = "END" Then GoTo NextColumn

            
'                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                REMOVE NUMBERS FROM FIELD AND SUBFIELD NAMES AND HEADERS

            ' Some of the column heads have numbers in them because their names are duplicates. This gets rid of
            ' those numbers
            Dim FieldValue As String
            Dim SubfValue As String
            Dim HeadValue As String
            Dim CellValue As String
            
            FieldValue = Field.value
            SubfValue = Subf.value
            HeadValue = Head.value
            CellValue = CurCell.value
            
            ' check if we should close the field
            Set NextField = XMLSheet.Cells(1, col + 1)
            Set NextCell = XMLSheet.Cells(row, col + 1)
            
            If IsEmpty(CurCell.value) Then GoTo NextColumn
            
            ' get all the numbers out of the repeated headers. stupid excel puts them in
            For i = 0 To 9

                FieldValue = replace(FieldValue, i, "")

                SubfValue = replace(SubfValue, i, "")

                HeadValue = replace(HeadValue, i, "")

                
             Next i
             
             ' replace XML and HTML reserved characters
             CellValue = replace(CellValue, "&", "&amp;")
             CellValue = replace(CellValue, "<", "&lt;")
             CellValue = replace(CellValue, ">", "&gt;")
             CellValue = replace(CellValue, "'", "&apos;")
             
             
            
'                      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            
            If Field.value = Empty Then            ' There can be no subfield without a field, this piece of data
                indent = 2                         ' only has its FacilityRow as a parent. We open no additional
                                                   ' XML tages for this data
                
                
                ' Save line into a string which can later be printed to XML
                printStr = printStr & vbTab & vbTab & "<" & HeadValue & _
                ">" & CellValue & "</" & Head & ">"
        
            ElseIf Field.value <> Empty And Subf.value = Empty Then
                indent = 2
                
                                                   ' In this case, the data resides inside of a field, but no sub-
                                                   ' field. This means we need to open the field tag if it hasn't
                                                   ' already been opened
                
                ' Check if the field has already been opened
                If Field.value <> oldField Then
                    
                    ' if the new field and old field are not the same, we need to open the tag in the XML for the
                    ' new field
                    printStr = printStr & vbTab & vbTab & "<" & FieldValue & ">" & vbNewLine
                    
                    endField = "Yes"
                    
                End If
                
                indent = indent + 1
                
                If FieldValue <> "ATTRIBUTE" Then
                    printStr = printStr & vbTab & vbTab & vbTab & _
                            "<" & HeadValue & ">" & CellValue & "</" & HeadValue & ">"
                
                Else
                    attStr = CellValue
                    attArr = Split(attStr, "%")
                    
                    For Each entry In attArr
                    
                        eachAtt = Split(entry, ":")
                    
                        HeadValue = eachAtt(0)
                        CellValue = eachAtt(1)
                    
                        If entry = attArr(0) Then
                            printStr = printStr & vbTab & vbTab & vbTab & _
                                "<" & HeadValue & ">" & CellValue & "</" & HeadValue & ">"
                        Else
                             printStr = printStr & vbNewLine & vbTab & vbTab & vbTab & _
                                "<" & HeadValue & ">" & CellValue & "</" & HeadValue & ">"
                        End If
                    
                    Next entry
                End If
                            
            Else                                   ' The data is inside both a field and subfield. We will have to
                indent = 2                         ' check if either of them is open already
                
                ' check if the field has already been opened
                If Field.value <> oldField Then
                    printStr = vbTab & vbTab & "<" & FieldValue & ">" & vbNewLine
                    
                    endField = "Yes"
                End If
                
                indent = indent + 1
                
                ' check if the subfield has already been opened
                If Subf.value <> oldSubf Then
                    printStr = printStr & vbTab & vbTab & vbTab & "<" & SubfValue & ">" & vbNewLine
                End If
                
                indent = indent + 1
                
                ' print the info for the data
                printStr = printStr & vbTab & vbTab & vbTab & vbTab & _
                        "<" & HeadValue & ">" & CellValue & "</" & HeadValue & ">"
                
            End If

            
            ' check if we should close the subfield
            Set NextSubf = XMLSheet.Cells(2, col + 1)
            If Subf.value <> NextSubf.value And Subf.value <> Empty Then
            
                printStr = printStr & vbNewLine & vbTab & vbTab & vbTab & "</" & SubfValue & ">"
                
            End If
            

            If Field.value <> NextField.value And Field.value <> Empty Then

                printStr = printStr & vbNewLine & vbTab & vbTab & "</" & FieldValue & ">"
                
            End If
            
            
            
            ' Print the string that we built!
            Print #2, printStr
            
            ' Sets the Field and subfield we just used as the old fields
            oldField = Field.value
            oldSubf = Subf.value
            
            ' Reset the print string for the next piece of data
            printStr = ""
            
NextColumn:
            If CellValue = "" And (Field.value <> NextField.value And Field.value <> Empty) And NextField <> "END" And endField = "Yes" Then

                printStr = printStr & vbTab & vbTab & "</" & FieldValue & ">" & vbNewLine
                
                endField = "No"
            End If

        Next col
    End If

' When finished going through the columns, we close the facility row
printStr = printStr & vbTab & "</FacilityRow>"

Print #2, printStr

' count how many entries rows are in the current XML doc
overFlowCount = overFlowCount + 1
' check if a new XML document should be started
If overFlowCount > docMax - 1 Then

    Print #2, "</FacilityTable>"
    
    If master = "Master" Then
        Print #2, "</Inventory>"
    End If
    
    Close #2

    docCount = docCount + 1
    
    If docCount > UBound(docArr) Then GoTo AfterXML
    
    If InStr(getOS, "Windows") = 0 Then

        XMLPath = ExportXML.FileDest.Text & ":" & docArr(docCount)       ' We save the .xml here for Macs
    
    Else

        XMLPath = ExportXML.FileDest.Text & "\" & docArr(docCount)      ' Save XML for PC
    End If


    If XMLPath = "\" Or XMLPath = ":" Then
        GoTo XMLFinish
    End If

    Open XMLPath For Output As #2

    overFlowCount = 0

    If row + 1 < endRow And master = "Master" Then
        ' start the new XML document
        printStr = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbNewLine & _
                    "<Inventory>" & vbNewLine & _
                    "<FacilityTable>"
           
        Print #2, printStr
        
    ElseIf row + 1 < endRow Then
    
        printStr = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbNewLine & _
                    "<FacilityTable>"
                    
        Print #2, printStr
        
    Else: GoTo AfterXML
    
    End If
    
    
End If








' The software skips here if the first value in the XML table is blank
ContinueLoop:

progressCount = progressCount + 1

' reset old field and old subfield
oldField = ""
oldSubf = ""

' iterate to the next row
row = row + 1
Loop

ProgressForm.ProgressLabel.Width = ProgressForm.MaxWidth.Caption
ProgressForm.ProgressFrame.Caption = "100" & "%"

' Close the root XML tag
Print #2, "</FacilityTable>"


AfterXML:

If master = "Master" Then GoTo MasterSkip3
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''' Make Error Statement ''''''''''''''''''''''''''''''''''''''''
Dim errStr As String
errStr = MakeErrStr(xmlInfo, "Facility")

Dim DiaStr As String

DiaStr = "Your XML worksheet has been created. It's file location is: " & _
    vbNewLine & vbNewLine
    
For Each doc In docArr
    If InStr(getOS, "Windows") = 0 Then
        DiaStr = DiaStr & ExportXML.FileDest.Text & ":" & doc & vbNewLine
    Else
        DiaStr = DiaStr & ExportXML.FileDest.Text & "\" & doc & vbNewLine
    End If
Next doc


DiaStr = DiaStr & vbNewLine & vbNewLine & _
    "Facilities Accepted: " & xmlInfo(0) & _
    vbNewLine & _
    "Facilities Declined: " & xmlInfo(1) & _
    vbNewLine & vbNewLine & _
    "--------------------------------------------------------------------------" & vbNewLine & _
    "--------------------------------------------------------------------------"
      
        
DiaStr = DiaStr & _
    vbNewLine & vbNewLine & _
    "Errors: " & xmlInfo(1)

If errStr <> "" Then

    DiaStr = DiaStr & _
        vbNewLine & vbNewLine & _
        "Any facilities that you attempted to include in your XML document that were rejected are highlighted " & _
        "in blue." & _
        vbNewLine & vbNewLine & _
        "The following cells contain invalid entries and are stopping some facilities from being included in the XML document: " & _
        vbNewLine & vbNewLine & _
        errStr

Else
    DiaStr = DiaStr & vbNewLine & vbNewLine & _
        "All facility information has been converted to XML."

End If

'Worksheets("Dialogue").Activate
'Set DialogBox = DiaSheet.Label21

'DialogBox.Caption = DiaStr


Unload ProgressForm

DialogueForm.DialogueBox.Text = DiaStr
DialogueForm.Show



' close the paths to the XML document and the text Dialogue
Close #2


' If the user exits the export box, we don't want to export the info!
XMLFinish:
If XMLPath = "\" Or XMLPath = ":" Or Err Then

    If Err Then
    
        MsgBox "Your XML document was not exported due to:" & _
            vbNewLine & _
            Err.Description
    Else
        MsgBox "Your XML document was not exported"
    End If
        
End If

Application.Run "protectWorkbook"
Application.ScreenUpdating = True
Application.EnableEvents = True
MasterSkip3:

End Sub

'' MakeXMLTable
'' Daniel Slosky
'' Last Updated: 2/12/2015
''
'' This software works in conjunction with FacilityXML to generate an XML document that holds
'' facility data for ShakeCast
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Function FacXMLTable()

'-------------------------------- Create Variables to Track Errors ---------------------------------

Dim xmlNums(0 To 1) As Variant ' this will be the return variable with two values,
                               ' the number of accepted and declined facilities
xmlNums(0) = 0
xmlNums(1) = 0

Dim ErrAdd() As Variant          ' ErrAdd is an array that holds the addresses of the cells with
                                       ' errors

'---------------------------------------------------------------------------------------------------




' create some handles to access the worksheets
Set FacSheet = Worksheets("Facility XML")
Set XMLSheet = Worksheets("FacilityXMLexport")

' Start with a fresh XML table
Dim lastRow As Long
Dim lastCol As Integer
lastRow = FacSheet.Cells(Rows.count, "F").End(xlUp).row


LastXML = XMLSheet.Cells(Rows.count, "A").End(xlUp).row

If LastXML < 4 Then
    LastXML = 4
End If

XMLSheet.Range("A4:" & "AD" & LastXML).Clear


' These values are used to determine starting and ending position for the loops used to read the
' facility worksheet

Dim startRow As Integer
Dim startCol As Integer
Dim endRow As Long
Dim endCol As Integer

startRow = 4
startCol = 1
endRow = FacSheet.Cells(Rows.count, "F").End(xlUp).row
endCol = 31

' Create variable to monitor how many rows of the XML table are filled. The integer keeps track of
' the NEXT ROW TO BE FILLED
Dim XMLcol As Integer
Dim XMLrow As Long
XMLcol = 1
XMLrow = 4

' set up progress bar
Dim progressCount As Long
Dim progressWhen As Long
Dim pcntDone As Double

progressCount = 0
progressWhen = endRow * 0.01
pcntDone = 0
ProgressForm.ProcessName.Caption = "Making Facility XML Table"

Dim RowRange1 As String
Dim RowRange2 As String

Dim FacRow As Long

FacRow = startRow
Do While FacRow < endRow + 1

    If progressCount > progressWhen Then
        
        pcntDone = ((FacRow - 4) / (endRow - 3))
        
        ProgressForm.ProgressLabel.Width = pcntDone * ProgressForm.MaxWidth.Caption
        ProgressForm.ProgressFrame.Caption = Round(pcntDone * 100, 0) & "%"
        
        DoEvents
        progressCount = 0
        
    End If

    RowRange1 = "A" & FacRow
    RowRange2 = "AE" & FacRow
    
    
'    ' check facilities on export
'    If FacRow > 3 And (Application.WorksheetFunction.CountBlank(FacSheet.Range(RowRange1, RowRange2)) > 0 Or _
'            FacSheet.Range("AE" & FacRow) = "Bad") Then
'        CheckFacilities FacRow
'    End If
    If Application.WorksheetFunction.CountBlank(Range(RowRange1, RowRange2)) = 31 Then
            
        ' ChangeColors "Good", Range(RowRange1, RowRange2), "Facility"
        GoTo ContinueLoop
        
    ElseIf IsEmpty(FacSheet.Range("A" & FacRow)) Or _
        IsEmpty(FacSheet.Range("B" & FacRow)) Or _
        IsEmpty(FacSheet.Range("F" & FacRow)) Or _
        IsEmpty(FacSheet.Range("J" & FacRow)) Or _
        IsEmpty(FacSheet.Range("K" & FacRow)) Then
            
            ' Count the number of facilities that will be rejected
            xmlNums(1) = xmlNums(1) + 1
            
            If FacSheet.Range("AE" & FacRow).value <> "bad" Then
                FacSheet.Range("AE" & FacRow).value = "bad"
                ChangeColors "Bad", Range(RowRange1, RowRange2), "Facility"
            End If
            
            ReDim Preserve ErrAdd(0 To xmlNums(1) - 1) As Variant
            ErrAdd(xmlNums(1) - 1) = FacRow
            GoTo ContinueLoop
        ' A row with zero attempted fields is unattempted, so we don't need to highlight it

'    ElseIf IsEmpty(FacSheet.Range("AE" & FacRow)) Then
'        GoTo ContinueLoop
    End If



    ''''''''''''''''''''''''''''''''''''''''''''''
    ' If this code is executing, the row is FULL '
    ''''''''''''''''''''''''''''''''''''''''''''''
    
    ' Count the number of facilities that will make it all the way to the XML
    xmlNums(0) = xmlNums(0) + 1
    

    ' copy data from Facility Sheet to the XML table
    XMLSheet.Range("A" & XMLrow, "B" & XMLrow).value = FacSheet.Range("A" & FacRow, "B" & FacRow).value
    XMLSheet.Range("C" & XMLrow, "H" & XMLrow).value = FacSheet.Range("D" & FacRow, "I" & FacRow).value
    XMLSheet.Range("I" & XMLrow).value = ManLatLong(FacSheet.Range("J" & FacRow), FacSheet.Range("K" & FacRow))
    XMLSheet.Range("J" & XMLrow, "K" & XMLrow).value = FacSheet.Range("L" & FacRow, "M" & FacRow).value
    XMLSheet.Range("L" & XMLrow, "AA" & XMLrow).value = FacSheet.Range("O" & FacRow, "AD" & FacRow).value
            
    
    XMLrow = XMLrow + 1
    
ContinueLoop:

progressCount = progressCount + 1
FacRow = FacRow + 1
    
Loop


Dim exportInfo(0 To 3) As Variant

ProgressForm.ProgressLabel.Width = ProgressForm.MaxWidth.Caption
ProgressForm.ProgressFrame.Caption = "100" & "%"
        
DoEvents

exportInfo(0) = xmlNums(0)
exportInfo(1) = xmlNums(1)
exportInfo(2) = ErrAdd

FacXMLTable = exportInfo
End Function



'' FillFacility
'' Daniel Slosky
'' Last Update: 2/4/2014
'' This function will check to see if the user has selected a building model before filling out a row before it
'' VLOOKs for the contents to fill the cell
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 Function FillFacility(Model As Range, _
                        cell As Range) As String
                        

If Model.value = Empty Then
    
    FillFacility = ""
    Exit Function
    
End If




ModVal = Model.value
    
Dim col As Integer
col = cell.Column

Dim Lookup As Integer
If col = 15 Then
    Lookup = 2
    


ElseIf col = 3 Then

    Lookup = col - 1
    
    FillFacility = Application.WorksheetFunction.VLookup(Model.value, Sheet5.Range("$C:$D"), Lookup, False)
    
    cell.HorizontalAlignment = xlLeft
    Exit Function

ElseIf col >= 16 Then
    Lookup = col - 10
    
End If
    
FillFacility = Application.WorksheetFunction.VLookup(Model.value, Sheet4.Range("$A:$T"), Lookup, False)


cell.HorizontalAlignment = xlCenter


End Function

Function FillSystem(cell As Range)

If cell.value = Empty Then
    FillSystem = ""
Else
    FillSystem = "SYSTEM"

End If
End Function

'' GeomType
'' Daniel Slosky
'' Last Update: 2/3/2015
'' This function looks at the processed values for the latitude and longitude as entered by the user to determine the
'' GEOM_TYPE for the XML file
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GeomType(LatLon As String)

' If there are no new lines in the input string, then there is only one specified point
If LatLon = "" Or InStr(LatLon, "ERROR") = 1 Then

    GeomType = ""
    Exit Function
    
ElseIf InStr(LatLon, " ") = 0 Then
    GeomType = "POINT"
    
    Exit Function
    
' Otherwise, we have to break up the string to find out how many points have been specified
Else
    Dim LatLonArr As Variant
    LatLonArr = Split(LatLon, " ")
    
    Dim Lats As Variant
    Dim Longs As Variant
    
    Lats = LatLonArr
    Longs = LatLonArr
    
    For i = 0 To UBound(LatLonArr)
    
        Point = Split(LatLonArr(i), ",")
        Lats(i) = Point(0)
        Longs(i) = Point(1)
       
    Next i
End If

' Since you can't make a shape with only 3 points (repeating the start), this has to be a line
If UBound(LatLonArr) > 0 And UBound(LatLonArr) <= 2 Then
    GeomType = "POLYLINE"
    
    Exit Function
' Check for the special case of a rectangle
ElseIf UBound(LatLonArr) = 4 And (StrComp(Lats(0), Lats(UBound(Lats))) = 0 And _
        StrComp(Longs(0), Longs(UBound(Longs)) = 0)) Then
    
    ' We can think of this as all the x-coordinates being in the Lats array and all the y-coordinates in the
    ' Longs array. We can move the origin of this system to the first point by subtracting the coordinates of the
    ' first point from the other points. If the coordinates truly make a rectangle, either the x or y coordinate for
    ' the second point should be zero and the opposite should be zero for the fourth point. Then when the second and
    ' fourth coordinates are subtracted from the third coordinate, the result should be ZERO.
    

    If ((StrComp(Lats(0), Lats(1)) = 0 And StrComp(Longs(0), Longs(3)) = 0) Or _
       (StrComp(Lats(0), Lats(3)) = 0 And StrComp(Longs(0), Longs(1)) = 0)) And _
       ((StrComp(Lats(2), Lats(1)) = 0 And StrComp(Longs(2), Longs(3)) = 0) Or _
       (StrComp(Lats(2), Lats(3)) = 0 And StrComp(Longs(2), Longs(1)) = 0)) Then
       
        ' don't support rectangle anymore
        ' GeomType = "RECTANGLE"
        GeomType = "POLYGON"
        
        Exit Function
    Else
        GeomType = "POLYGON"
        
        Exit Function
        
    End If
    
' If not a rectangle but the starting point is the same as the end point, we have a polyline!
ElseIf UBound(LatLonArr) > 2 And (StrComp(Lats(0), Lats(UBound(Lats))) = 0 And _
        StrComp(Longs(0), Longs(UBound(Longs)) = 0)) Then
    
    GeomType = "POLYGON"
    Exit Function
    
Else
    GeomType = "POLYLINE"
    Exit Function
    
End If

GeomType = ""

End Function


'' CopyIf
'' Daniel Slosky
'' Last Update: 2/5/2015
'' This function will check to see if all the required fields are filled by the customer before filling the XML row
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function CopyIf(cell As Range, _
                    CurCell As Range, _
                    RefCell As Range)

' Let's see where we are... for debugging
Dim col As Integer
Dim ro As Integer
col = CurCell.Column
ro = CurCell.row


' Column 32 is used as a reference for the XML cells. If it is changed, the other cells change too. This saves some computing power
If CurCell.Column = 32 Then

    For Each ref In RefCell
        
        If ref.value = Empty Then
            
            CopyIf = ""
            Exit Function
        End If
    Next ref
        
ElseIf RefCell.value = Empty Then
    
    CopyIf = ""
    Exit Function
    
End If


If cell.value = Empty Then
    CopyIf = ""
    Exit Function
End If

Dim cellRow As String
cellRow = cell.row

' Sum up all the required cells that are filled
Dim sumEmpty As Integer
Dim CellRange1 As String
Dim CellRange2 As String

CellRange1 = "A" & cellRow
CellRange2 = "AD" & cellRow
sumEmpty = Application.WorksheetFunction.CountBlank(Range(CellRange1, CellRange2))


If sumEmpty = 0 And InStr(cell.value, Chr(10)) <> 0 Then
    Dim strArr As Variant
    strArr = Split(cell.value, Chr(10))
    
    Dim printStr As String
    Dim i As Integer
    For i = 0 To UBound(strArr)
    
        If i = 0 Then
            printStr = strArr(i)
        Else
        printStr = printStr & vbNewLine & strArr(i)
        End If
    
    Next i
    

    
    CopyIf = printStr
ElseIf sumEmpty = 0 Then
    
    CopyIf = cell.value
    
Else
    CopyIf = ""

End If

End Function


'' ManLatLong
'' Daniel Slosky
'' Last Update: 2/11/2015
'' This macro will check that the same number of latitudinal and longitudinal coordinants were supplied by the user
'' then concatinate the values from "cell1: Lat cell2: long" to "OneCell: Lat,Long,elevation". The default value for
'' elevation will be 0, but this section can be manually changed by the user if they wish.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function ManLatLong(LatCell As Range, _
                    LongCell As Range)

 
Dim LatRows As Integer     ' The number of rows in the latitude column
Dim LongRows2 As Integer   ' Number of rows in the longitude column
Dim LatArray As Variant    ' Array of latitude rows
Dim LongArray As Variant   ' Array of longitude rows
Dim printStr As String
 
If IsEmpty(LatCell) Or IsEmpty(LongCell) Or IsError(LatCell) Or IsError(LongCell) Then
    ManLatLong = ""
    Exit Function
    
' Get Row lengths for latitude
ElseIf InStr(LatCell.value, ";") = 0 Then
    LatRows = 1
    
Else
    ' lets also check if the row number the user provided is not too big.
    LatArray = Split(LatCell.value, ";")
    LatRows = UBound(LatArray) + 1
End If


' Get Row lengths for longitude
If InStr(LongCell.value, ";") = 0 Then
    LongRows = 1

Else
    ' lets also check if the row number the user provided is not too big.
    LongArray = Split(LongCell.value, ";")
    LongRows = UBound(LongArray) + 1
    
End If


' Check that the latitude and longitude inputs have the same amount of rows
If LatRows > LongRows Or LatRows < LongRows Then
    ManLatLong = "ERROR: CHECK LAT LONG"
    MsgBox "Error: IT looks like your latitude and longitude in row " & LatCell.row & " doesn't make sense! " & _
        "Check to make sure you have the same number of latitude and longitude entries."
    Exit Function
    
ElseIf LatRows = 1 And LongRows = 1 Then
    printStr = printStr & LongCell.value & "," & LatCell.value & ",0"
Else
    printStr = ""
    
    Dim i As Integer
    For i = 0 To UBound(LatArray)
    
        If i = 0 Then
            printStr = printStr & LongArray(i) & "," & LatArray(i) & ",0"
        Else
        printStr = printStr & " " & LongArray(i) & "," & LatArray(i) & ",0"
        End If
    
    Next i
    
End If

ManLatLong = printStr

End Function


Private Sub RefreshButton()

Dim refreshInfo() As Variant
refreshInfo = Application.Run("RefreshFormulas")

Dim DiaStr As String

DiaStr = "Your Formulas have been refreshed!" & _
    vbNewLine & vbNewLine & _
    "--------------------------------------------------------------------------" & _
    "--------------------------------------------------------------------------" & _
    vbNewLine & vbNewLine & _
    "Formulas Refreshed: " & refreshInfo(0) & _
    vbNewLine & vbNewLine

If refreshInfo(1) <> "" Then
    DiaStr = DiaStr & _
        refreshInfo(1) & _
        vbNewLine & vbNewLine
        
End If
      
        
DiaStr = DiaStr & _
    "--------------------------------------------------------------------------" & _
    "--------------------------------------------------------------------------"


DialogueForm.DialogueBox.Text = DiaStr
DialogueForm.Show


End Sub

Function RefreshFormulas() As Variant

Set FacSheet = Worksheets("Facility XML")

Dim refreshStr As String
Dim rowCheck As Integer
Dim total As Integer
Dim tabSpace As String

refreshStr = ""  ' Used to build a string which describes the fields that were refreshed
rowCheck = 0     ' If rowCheck = 0, then a new cell header is created. Otherwise, it isn't!
total = 0        ' count the total amount of cells refreshed
tabSpace = "      "


startRow = 4
endRow = FacSheet.Cells(Rows.count, "B").End(xlUp).row + 1

For i = startRow To endRow

    If i > 9 And i < 100 Then
        tabSpace = "    "
                
    ElseIf i > 99 And i < 1000 Then
        tabSpace = "   "
        
    ElseIf i > 999 And i < 10000 Then
        tabSpace = "  "
        
    ElseIf i > 9999 And i < 100000 Then
    
        tabSpace = " "
    
    End If
    
    Dim row As Range
    Set row = FacSheet.Cells.Range("A" & i, "AD" & i)
    
    If Application.WorksheetFunction.CountBlank(row) < 30 Then
        
        For Each cell In row.Cells
            
            If (cell.Column = 4 Or cell.Column = 5) And (IsError(cell) _
                Or IsEmpty(cell)) Then
                
                cell.Formula = "=FillSystem(B" & cell.row & ")"
              
                cell.HorizontalAlignment = xlLeft
                
                ' add the cell address the the refreshStr which will be exported
                If rowCheck = 0 Then
                    refreshStr = refreshStr & "Row " & cell.row & ": " & tabSpace & cell.Address(False, False)
                Else
                    refreshStr = refreshStr & " :: " & cell.Address(False, False)
                End If
                    
                rowCheck = rowCheck + 1
                    
            ElseIf cell.Column = 9 And (IsError(cell) _
                Or IsEmpty(cell)) Then
                
                cell.Formula = "=GeomType(J" & cell.row & ")"
                
                ' add the cell address the the refreshStr which will be exported
                If rowCheck = 0 Then
                    refreshStr = refreshStr & "Row " & cell.row & ": " & tabSpace & cell.Address(False, False)
                Else
                    refreshStr = refreshStr & " :: " & cell.Address(False, False)
                End If
                
                rowCheck = rowCheck + 1
                
            ElseIf cell.Column = 10 And (IsError(cell) _
                Or IsEmpty(cell)) Then
                
                cell.Formula = "=ManLatLong(K" & cell.row & ", L" & cell.row & ")"
                
                
                ' add the cell address the the refreshStr which will be exported
                If rowCheck = 0 Then
                    refreshStr = refreshStr & "Row " & cell.row & ": " & tabSpace & cell.Address(False, False)
                Else
                    refreshStr = refreshStr & " :: " & cell.Address(False, False)
                End If
                
                rowCheck = rowCheck + 1
                
                
            ElseIf cell.Column > 14 And (IsError(cell) _
                Or IsEmpty(cell)) Then
                
                cell.Formula = "=FillFacility($N" & cell.row & "," & cell.Address(False, False) & ")"
                
                cell.HorizontalAlignment = xlCenter
                
                ' add the cell address the the refreshStr which will be exported
                If rowCheck = 0 Then
                    refreshStr = refreshStr & "Row " & cell.row & ": " & tabSpace & cell.Address(False, False)
                Else
                    refreshStr = refreshStr & " :: " & cell.Address(False, False)
                End If
                
                rowCheck = rowCheck + 1
                
                
                
            ElseIf cell.Column = 3 And (IsError(cell) _
                Or IsEmpty(cell)) Then
                
                cell.Formula = "=FillFacility($B" & cell.row & "," & cell.Address(False, False) & ")"
                
                cell.HorizontalAlignment = xlLeft
                
                ' add the cell address the the refreshStr which will be exported
                If rowCheck = 0 Then
                    refreshStr = refreshStr & "Row " & cell.row & ": " & tabSpace & cell.Address(False, False)
                Else
                    refreshStr = refreshStr & " :: " & cell.Address(False, False)
                End If
                
                rowCheck = rowCheck + 1
                
                
            End If
            
            If cell.row > 9 And cell.row < 100 Then
                tabSpace = "   "
                
            ElseIf cell.row > 99 And cell.row < 1000 Then
                tabSpace = "  "
            
            End If
            
            If rowCheck = 12 Or rowCheck = 24 Then

                If cell.row < 10 Then
                    refreshStr = refreshStr & vbNewLine & "                      "
                ElseIf cell.row > 9 And cell.row < 100 Then
                    refreshStr = refreshStr & vbNewLine & "                        "
                ElseIf cell.row > 99 And cell.row < 1000 Then
                    refreshStr = refreshStr & vbNewLine & "                           "
                ElseIf cell.row > 999 And cell.row < 10000 Then
                    refreshStr = refreshStr & vbNewLine & "                              "
                ElseIf cell.row > 9999 And cell.row < 100000 Then
                    refreshStr = refreshStr & vbNewLine & "                                 "
                    
                End If
                
                
            End If
            
        Next cell
    End If
    
    If rowCheck > 0 Then
        
        refreshStr = refreshStr & vbNewLine & vbNewLine
        
        total = total + rowCheck
        
        rowCheck = 0
        
    End If
Next i

Dim refInfo(0 To 1) As Variant

refInfo(0) = total
refInfo(1) = refreshStr

RefreshFormulas = refInfo

End Function

Sub UpdateFacButton()
    Application.EnableEvents = False
    
    ' create some handles to access the worksheets
    Set FacSheet = Worksheets("Facility XML")
    
    ' These values are used to determine starting and ending position for the loops used to read the
    ' facility worksheet
    
    Dim startRow As Integer
    Dim startCol As Integer
    Dim endRow As Long
    Dim endCol As Integer
    
    startRow = 4
    startCol = 1
    
    If FacSheet.Cells(Rows.count, "A").End(xlUp).row > FacSheet.Cells(Rows.count, "AD").End(xlUp).row Then
        endRow = FacSheet.Cells(Rows.count, "A").End(xlUp).row
    Else
        endRow = FacSheet.Cells(Rows.count, "AD").End(xlUp).row
    End If
    
    endCol = 30
    
    Dim progressCount As Long
    Dim progressWhen As Long
    Dim pcntDone As Double
    
    progressCount = 0
    progressWhen = endRow * 0.01
    prcntDone = 0
    ProgressForm.ProcessName.Caption = "Updating Worksheet"
    
    Dim RowRange1 As String
    Dim RowRange2 As String
    
    Dim FacRow As Long
    FacRow = startRow
    Do While FacRow < endRow + 1
    
        If progressCount > progressWhen Then
            
            pcntDone = ((FacRow - 4) / (endRow - 3))
            
            ProgressForm.ProgressLabel.Width = pcntDone * ProgressForm.MaxWidth.Caption
            ProgressForm.ProgressFrame.Caption = Round(pcntDone * 100, 0) & "%"
            
            DoEvents
            progressCount = 0
            
        End If
    
        If WorksheetFunction.CountBlank(FacSheet.Range("N" & FacRow, "AC" & FacRow)) = 16 Then
            CheckFacilities FacSheet.Range("M" & FacRow)
        Else
            CheckFacilities FacSheet.Range("B" & FacRow)
        End If
        
        
ContinueLoop:
    FacRow = FacRow + 1
    progressCount = progressCount + 1
    Loop
    
    ProgressForm.ProgressLabel.Width = ProgressForm.MaxWidth.Caption
    ProgressForm.ProgressFrame.Caption = "100" & "%"
    
    Unload ProgressForm
    
    MsgBox "Worksheet Updated"
    
    Application.EnableEvents = True
End Sub

Private Sub facilityUnlock()

    Set mySheet = Worksheets("Facility XML")

    ' figure out the used range of the workbook
    Dim startRow As Integer
    Dim endRow As Long
    Dim startCol As String
    Dim endCol As String
    
    startRow = 4
    startCol = "A"
    endCol = "AE"
    
    endRow = mySheet.Cells(Rows.count, "A").End(xlUp).row
    
    mySheet.Range(startCol & startRow, endCol & endRow).Locked = False
    

End Sub

Sub OptionsButton()

    OptionsForm.Show

End Sub


Private Sub FacSheetInfo()

Set mySheet = Worksheets("Facility XML")


If mySheet.Range("A2").value = "General User" Then

    Application.Run "FacGenInfo"

Else

    Application.Run "FacAdvInfo"

    
End If


End Sub

Private Sub FacIDinfo()

DialogueForm.DialogueBox.Text = "Facility ID Info:" & _
vbNewLine & vbNewLine & _
"This field is required." & _
vbNewLine & vbNewLine & _
"The facility ID can be any combination of numbers and letters, but should not have any spaces. It is used to keep track of your facilities in our database in conjunction " & _
"with the facility type." & vbNewLine & vbNewLine & _
"A facility ID can be the same for multiple facilities of different types (like a bridge and a dam), " & _
"but should not be the same for two facilities of the same type. For example, two bridges should not have " & _
"the facility ID ""112""." & _
vbNewLine & vbNewLine & _
"It turns out that advanced users can actually break this rule. If you wish to define multiple fragilities to a single " & _
"facility, this can be done by entering multiple facility rows with the same Facility ID and Facility Type, but with different " & _
"Components. Component names can be changed from the advanced user spreadsheet" & vbNewLine & vbNewLine & _
"This field is to be filled in by the user. This field is also mandatory for all facilites. If it is not completed, " & _
"this facility will not be uploaded to the ShakeCast system."

DialogueForm.Show

End Sub
Private Sub FacTypeInfo()
DialogueForm.DialogueBox.Text = "Facility Type Info:" & _
vbNewLine & vbNewLine & _
"This field is required" & vbNewLine & vbNewLine & _
"The facility type must be selected from the drop down menu. If you don't see a drop down menu in the cell, make sure " & _
"that you have already completed the Facility ID section"

DialogueForm.Show

End Sub

Private Sub FacTypeDesInfo()

DialogueForm.DialogueBox.Text = "Facility Type Description Info:" & _
vbNewLine & vbNewLine & _
"This field required, but should automatically populate. It exists to help describe your facility type description."

DialogueForm.Show

End Sub
Private Sub FacFullNameInfo()

DialogueForm.DialogueBox.Text = "Facility Full Name Info:" & vbNewLine & vbNewLine & _
"This field is required." & _
"This field is the name that will be displayed to you in the ShakeCast application."

DialogueForm.Show

End Sub
Private Sub FacDescInfo()

DialogueForm.DialogueBox.Text = "Facility Description Info:" & vbNewLine & vbNewLine & _
"This field is required." & vbNewLine & vbNewLine & _
"The information in this field will be shown in the ShakeCast application when the specific facility is selected."

DialogueForm.Show

End Sub

Private Sub ShortNameInfo()

DialogueForm.DialogueBox.Text = "Facility Short Name Info:" & vbNewLine & vbNewLine & _
"This field is required." & vbNewLine & vbNewLine & _
"The name in this field will be displayed when there is a limit to the number of characters ShakeCast can display."

DialogueForm.Show

End Sub

Private Sub LatLonInfo()

DialogueForm.DialogueBox.Text = "Latitude and Longitude Info:" & vbNewLine & vbNewLine & _
"These fields are required." & vbNewLine & vbNewLine & _
"Combined, the latitude and longitude fields describe the location and shape of your facility. If you only wish to " & _
"describe the location of your facility, enter a single coordinate for the latitude and a single coordinate for the " & _
"longitude. If you wish to use multiple points to describe the shape of your facility, seperate the values with a " & _
"semi-colon. This would look like: " & vbNewLine & vbNewLine & _
"Latitude: 37.7451193;36.571893;36.147219" & vbNewLine & _
"Longtitude: -122.1843401;-118.427016;-120.756117" & vbNewLine & vbNewLine & _
"This facility is now described by the coordinates: (37.7451193,-122.1843401), (36.571893,-118.427016), and " & _
"(36.147219, -120.756117)"

DialogueForm.Show

End Sub
Private Sub HTMLInfo()

DialogueForm.DialogueBox.Text = "HTML Snippet Info:" & vbNewLine & vbNewLine & _
"This field is not required." & vbNewLine & vbNewLine & _
"This field is for a HTML description of your facility. It will be displayed on the map in the ShakeCast application."

DialogueForm.Show

End Sub
Private Sub HAZUSInfo()

DialogueForm.DialogueBox.Text = "HAZUS Model Building Type Info:" & vbNewLine & vbNewLine & _
"This field is required." & vbNewLine & vbNewLine & _
"FEMA has created the HAZUS model system to determine potential loss in a disater situation. By selecting a model " & _
"building type that matches your facility, we can automatically supply fragility parameters for your facility. If you " & _
"wish to enter personalized fragilites, click the ""Spreadsheet Info"" button, then click ""Access Advanced User Spreadsheet"". " & _
"This will unhide a spreadsheet called ""HAZUS Facility Model Data"". Any value changed on this spreadsheet will also be " & _
"changed in the facility spreadsheet."

DialogueForm.Show

End Sub

Private Sub FacAdvInfo()

    FacSheetForm.SheetInfoName.Caption = "The Advanced Facility Spreadsheet"
    FacSheetForm.DialogueBox.Text = "The advanced user spreadsheet can be used to manually enter components, component " & _
        "classes, facility attributes, and fragility information. " & vbNewLine & vbNewLine & _
        "You can manually set component, component class, and attribute " & _
        "information in the columns of this spreadsheet." & vbNewLine & vbNewLine & _
        "In order to edit HAZUS or save your own fragility data, click over to the ""HAZUS Facility Model Data"" " & _
        "spreadsheet, or select the option from the Options Menu. From here you can edit the values we have input for specific facility models " & _
        "or create your own facility by adding your own row to the bottom of the sheet." & vbNewLine & vbNewLine & _
        "New Facility Types can be created from the Options menu."
        
    FacSheetForm.AdvUser.Caption = "Access General User Worksheet"
     
    If FacSheetForm.Visible = False Then
        FacSheetForm.Show
    End If
     

End Sub

Private Sub FacGenInfo()

    FacSheetForm.SheetInfoName.Caption = "The Facility Spreadsheet"
    FacSheetForm.DialogueBox.Text = "This spreadsheet can be used to to convert information about your facilities into " & _
        "XML format, which is readable by the ShakeCast application." & _
        vbNewLine & vbNewLine & _
        "This spreadsheet was made to be completed from left to right. Please hit tab or enter in order to submit your " & _
        "information to the spreadsheet. Some of the fields to " & _
        "the right will automatically fill as you move along. You may feel free to change " & _
        "any values we've supplied for you. " & vbNewLine & vbNewLine & _
        "If you are uncertain of the information you should be providing for a field, " & _
        "click on the ""More Info"" button at the top of that column. These hold information " & _
        "about our expectations for your input for each field." & vbNewLine & vbNewLine & _
        "When deleting a row, it is best to: select a couple cells in that row, highlight the delete drop-down menu, and select ""table rows"". If this isn't possible, you can highlight the entire row and hit the delete key." & vbNewLine & vbNewLine & _
        "When you are finished updating your facility information, hit Options button and select ""Export XML"" and hit Go. " & _
        "You will be prompted to select a save location and name for the file. The default save " & _
        "location is the folder this workbook is currently running in! " & _
        "This file can then easily be uploaded to ShakeCast by dragging and dropping it into " & _
        "the upload page." & vbNewLine & vbNewLine & _
        "It is also possible to export all facility, group, and user information in a single XML file, by clicking " & _
        "the ""Export Master XML"" button." & vbNewLine & vbNewLine & _
        "If you wish to define your own facility type, facility attributes, define facility components, or save your own fragility information, you can do so from the advanced user spreadsheet."
        
    FacSheetForm.AdvUser.Caption = "Access Advanced User Worksheet"

     
    If FacSheetForm.Visible = False Then
        FacSheetForm.Show
    End If

End Sub
