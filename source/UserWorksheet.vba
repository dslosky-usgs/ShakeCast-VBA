Attribute VB_Name = "UserWorksheet"
'' CheckUsers
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


Sub CheckUsers(ByVal checkRow As Integer)

' set our working spreadsheet to the User XML sheet
Set mySheet = Worksheets("User XML")

' keep the worksheet looking at the same cell while code executes
mySheet.ScrollArea = ActiveCell.Address

' make variable to track the active cell when we started this subroutine
Dim startActive As Range
Set startActive = ActiveCell

Dim lastRow As Integer
lastRow = mySheet.Cells(Rows.count, "A").End(xlUp).row ' where we stop!


' change color of all blank cells to white. The ones in group rows will have their color changed
' later
' Dim sheetRange As Range
' Set sheetRange = mySheet.Range("A:K")

' sheetRange.SpecialCells(xlCellTypeBlanks).Interior.ColorIndex = 2


' make the title banner stay the same color even though we are whiting out blank cells
Dim titleRange As Range
Set titleRange = mySheet.Range("A1:K2")

With titleRange.Interior
    .Color = RGB(196, 215, 155)
End With


    ' skip the headers. Since the active row goes first, if the uprow is a header, the active row
    ' will still get evaluated
If checkRow < 4 Then GoTo TheEnd

    ' for each row, check for user input

    If WorksheetFunction.CountBlank(mySheet.Range("A" & checkRow, "F" & checkRow)) = 6 Then
    

    ' The user input for the row is empty, so lets just clear the row
        mySheet.Range("A" & checkRow, "K" & checkRow).Clear
        mySheet.Range("A" & checkRow, "K" & checkRow).Locked = False
        
        GoTo TheEnd
    
    ElseIf WorksheetFunction.CountBlank(mySheet.Range("A" & checkRow, "E" & checkRow)) > 0 Then
    
        ' if row is incomplete, change color to bad
        ChangeColors "Bad", mySheet.Range("A" & checkRow, "J" & checkRow), "User"
        
        mySheet.Range("K" & checkRow).value = "Bad"
    
    Else
    
        ' check the special case of the ADMIN user
        If mySheet.Range("B" & checkRow) = "ADMIN" And IsEmpty(mySheet.Range("F" & checkRow)) Then
        
            ChangeColors "Bad", mySheet.Range("A" & checkRow, "J" & checkRow), "User"
            
            mySheet.Range("K" & checkRow).value = "Bad"
            
            GoTo TheEnd
        End If
    
        ' if row IS complete, change color to good!
        ChangeColors "Good", mySheet.Range("A" & checkRow, "J" & checkRow), "User"

        mySheet.Range("K" & checkRow).value = "Good"
End If

TheEnd:
' if ANY user input, supply autodata and group drop down menu
FillUserInfo checkRow




End Sub


Sub FillUserInfo(ByVal checkRow As Integer)

Set mySheet = Worksheets("User XML")
Set groupSheet = Worksheets("Notification XML")

' PAGER
If IsEmpty(mySheet.Range("H" & checkRow)) Or mySheet.Range("H" & checkRow).value = "example@example.com" Then
    mySheet.Range("H" & checkRow).value = mySheet.Range("E" & checkRow).value
End If

' EMAIL_HTML
If IsEmpty(mySheet.Range("I" & checkRow)) Or mySheet.Range("I" & checkRow).value = "example@example.com" Then
    mySheet.Range("I" & checkRow).value = mySheet.Range("E" & checkRow).value
End If

'EMAIL_ TEXT
If IsEmpty(mySheet.Range("J" & checkRow)) Or mySheet.Range("J" & checkRow).value = "example@example.com" Then
    mySheet.Range("J" & checkRow).value = mySheet.Range("E" & checkRow).value
End If


' create User Type and Group drop down menus
Set UserType = mySheet.Range("B" & checkRow)
Set groupName = mySheet.Range("G" & checkRow)

' create a string array with the user types
Dim UserTypes(0 To 1) As String
UserTypes(0) = "USER"
UserTypes(1) = "ADMIN"

' create a string array of all the group names. This means going into the group spreadsheet and reading
' all the names of the groups without counting repeated names
Dim GroupNames() As String
Set allGroupCells = groupSheet.Range("A:A")

Dim oldGroup As String
Dim curGroup As String
Dim GroupCount As Integer
Dim blankCount As Integer

GroupCount = 0
blankCount = 0
For Each groupCell In allGroupCells

    If groupCell.row < 4 Then GoTo NextGroupCell
    
    If IsEmpty(groupCell) Then
        blankCount = blankCount + 1
        GoTo NextGroupCell
    End If
    
    curGroup = groupCell.value
    
    If groupCell.row = 4 Then
    
        If groupSheet.Range("N" & groupCell.row).value = "Good" Then
            ReDim Preserve GroupNames(0 To 0)
            GroupNames(0) = curGroup
        End If
        
    ElseIf curGroup <> oldGroup Then
        
        If groupSheet.Range("N" & groupCell.row).value = "Good" Then
            GroupCount = GroupCount + 1
            ReDim Preserve GroupNames(0 To GroupCount)
        
            GroupNames(GroupCount) = curGroup
        
        End If
        
        ' reset blankCount to zero, since it is really trying to count the blank rows at the END of the
        ' spreadsheet
        blankCount = 0
        
    End If
        
    oldGroup = curGroup
NextGroupCell:
If blankCount > 10 Then GoTo QuitGroupLoop
Next groupCell
QuitGroupLoop:

' create user/admin drop down
'With UserType.Validation
'    .Delete
'    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
'    Operator:=xlBetween, Formula1:=Join(UserTypes, ",")
'    .IgnoreBlank = True
'    .InCellDropdown = True
'    .InputTitle = "User Type"
'    .ErrorTitle = ""
'    .InputMessage = "Please select a user type type from the drop-down list"
'    .ErrorMessage = ""
'    .ShowInput = True
'    .ShowError = True
'End With

Dim newGroupStr As String
newGroupStr = ""

For Each group In Split(groupName.value, ":")
    If InArray(GroupNames, group) Then
        If newGroupStr = "" Then
            newGroupStr = group
        Else
            newGroupStr = newGroupStr & ":" & group
        End If
    End If
Next group

If newGroupStr <> "" Then
    groupName.value = newGroupStr
Else
    groupName.value = Empty
End If

End Sub



'' UserXML
'' Daniel Slosky
'' Last Updated: 2/24/2016
'' Creates an XML document that holds all the information the user input into the User spreadsheet.
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
Sub UserXML(master As String, _
                    Optional ByVal docCount As Integer = 0, _
                    Optional ByVal overFlowCount As Integer = 0, _
                    Optional ByVal docMax As Integer = 15000, _
                    Optional ByVal docStr As String = "")


Dim docArr() As String

Dim getOS As String
getOS = Application.OperatingSystem

If master = "Master" Then
    docArr = Split(docStr, ",")
    GoTo MasterSkip1
End If


If master = "Master" Then GoTo MasterSkip1

Application.EnableEvents = False

'On Error GoTo XMLFinish
'On Error Resume Next

Close #2

'Dim refreshInfo() As Variant
'refreshInfo = Application.Run("RefreshFormulas") ' enter formulas into any fields that should hold formulas, but are empty

Set mySheet = Worksheets("User XML")

' change color of all blank cells to white. The ones in user rows will have their color changed
' later
'Dim sheetRange As Range
'Set sheetRange = mySheet.Range("A:K")
'sheetRange.SpecialCells(xlCellTypeBlanks).Interior.ColorIndex = 2


' We now get XML info from the worksheet UserXMLexport, and we have to move all the info over there
Dim xmlInfo() As Variant
xmlInfo = Application.Run("UserXMLTable")       ' This function populates the XML table in the HIDDEN
                                                 ' spreadsheet UserXMLexport



'                          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                 OPEN XML FILE FOR WRITING, AND DETERMINE THE START AND END CELLS TO BE EXAMINED

' open file location
Dim dir As String
dir = Application.ActiveWorkbook.Path

Dim docNum As Double

docNum = xmlInfo(0) / docMax

' we don't want the number of entries in the document to be exactly docMax, because then wierd stuff will happen
If WorksheetFunction.Ceiling(docNum, 1) = docNum Then
    docMax = docMax - 1
    docNum = infoAcc / docMax
End If

If docNum < 1 Then
    docNum = 1
    docStr = "UserXML.xml"
Else
    docNum = Application.WorksheetFunction.Ceiling(docNum, 1)
    docStr = "UserXML1.xml"
    For i = 2 To docNum
        docStr = docStr & "," & "UserXML" & i & ".xml"
    Next i
End If


ExportXML.FileDest.Text = dir
ExportXML.FileName = docStr
ExportXML.Show



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

Set XMLSheet = Worksheets("UserXMLexport")

Dim startCol As Integer                      ' Figure out which rows and columns we want to sort through
Dim startRow As Integer
Dim endCol As Integer
Dim endRow As Integer


startCol = 1                                 ' The first column that holds XML info
startRow = 4                                 ' The First Row with XML info


' The last column we will look at
' EndCol = XMLSheet.Cells(3, Columns.count).End(xlToLeft).column
endCol = 10


' This will pull cells with formulas and no values, but we can filter those out!
endRow = XMLSheet.Cells(Rows.count, "A").End(xlUp).row + 1




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''' BEGIN XML '''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim printStr As String
printStr = ""

If master = "Master" Then GoTo MasterSkip2
printStr = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbNewLine
MasterSkip2:

printStr = printStr & "<UserTable>"
           
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

Dim NextField As Range        ' Keeping track of the next field and subfield, we can determine when to close tags
Dim NextSubf As Range

oldField = ""
oldSubf = ""

' set up progress bar
Dim progressCount As Long
Dim progressWhen As Long
Dim pcntDone As Double

progressCount = 0
progressWhen = endRow * 0.01
pcntDone = 0
ProgressForm.ProcessName.Caption = "Exporting User XML"


' This loop will run through all the cells in the XML region of the document.
Dim row As Integer
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
    'Set TestCell = XMLSheet.Cells(row, 1) ' This is the first XML cell, which should be filled for any valid entry
    
        
        
        ' Since there IS data for this row, we start a new entry in the XML page
        printStr = vbTab & "<UserRow>" & vbNewLine
        
        ' Now we can run through all the columns in the row. Row 32 is the first one to hold XML data
        For col = 1 To endCol
        
            ' Assign the field, subfield, and header for the current cell.
            Set Field = XMLSheet.Cells(1, col)
            Set Subf = XMLSheet.Cells(2, col)
            Set Head = XMLSheet.Cells(3, col)
            Set CurCell = XMLSheet.Cells(row, col)
            
            
'                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                REMOVE NUMBERS FROM FIELD AND SUBFIELD NAMES AND HEADERS

            ' Some of the column heads have numbers in them because their names are duplicates. This gets rid of
            ' those numbers
            Dim FieldValue As String
            Dim SubfValue As String
            Dim HeadValue As String
            
            FieldValue = Field.value
            SubfValue = Subf.value
            HeadValue = Head.value
            CellValue = CurCell.value
            
            ' This will remove all numbers from the headers... but why am I doing it this way??
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
             
             Dim cellArray As Variant
             Dim CurCellValue As String
             Dim cellValCount As Integer
             
             CurCellValue = ""
             cellValCount = 0
             If InStr(CellValue, ";") Then
                cellArray = Split(CellValue, ";")
                
                For Each celval In cellArray
                
                    If cellValCount < 1 Then
                        CurCellValue = CurCellValue & celval
                        
                        cellValCount = cellValCount + 1
                        
                    Else
                        CurCellValue = CurCellValue & vbNewLine & celval
                        
                        cellValCount = cellValCount + 1
                        
                    End If
                
                Next celval
                
            Else
            
                CurCellValue = CurCell.value
                
            End If
'                      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            
            If Field.value = Empty Then            ' There can be no subfield without a field, this piece of data
                indent = 2                         ' only has its UserRow as a parent. We open no additional
                                                   ' XML tages for this data
                
                
                ' Save line into a string which can later be printed to XML
                printStr = printStr & vbTab & vbTab & "<" & HeadValue & _
                ">" & CurCellValue & "</" & Head & ">"
        
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
                    
                End If
                
                indent = indent + 1
                
                
                printStr = printStr & vbTab & vbTab & vbTab & _
                            "<" & HeadValue & ">" & CurCellValue & "</" & HeadValue & ">"
                
                            
            Else                                   ' The data is inside both a field and subfield. We will have to
                indent = 2                         ' check if either of them is open already
                
                ' check if the field has already been opened
                If Field.value <> oldField Then
                    printStr = vbTab & vbTab & "<" & FieldValue & ">" & vbNewLine
                End If
                
                indent = indent + 1
                
                ' check if the subfield has already been opened
                If Subf.value <> oldSubf Then
                    printStr = printStr & vbTab & vbTab & vbTab & "<" & SubfValue & ">" & vbNewLine
                End If
                
                indent = indent + 1
                
                ' print the info for the data
                printStr = printStr & vbTab & vbTab & vbTab & vbTab & _
                        "<" & HeadValue & ">" & CurCellValue & "</" & HeadValue & ">"
                
            End If
            
            
            ' check if we should close the subfield
            Set NextSubf = XMLSheet.Cells(2, col + 1)
            If Subf.value <> NextSubf.value And Subf.value <> Empty Then
            
                printStr = printStr & vbNewLine & vbTab & vbTab & vbTab & "</" & SubfValue & ">"
                
            End If
            
            ' check if we should close the field
            Set NextField = XMLSheet.Cells(1, col + 1)
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
           
        

        Next col


' When finished going through the columns, we close the User row
printStr = printStr & vbTab & "</UserRow>"


Print #2, printStr


overFlowCount = overFlowCount + 1
' check if a new XML document should be started
If overFlowCount > docMax - 1 Then

    Print #2, "</UserTable>"
    
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
                    "<UserTable>"
           
        Print #2, printStr
        
    ElseIf row + 1 < endRow Then
    
        printStr = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbNewLine & _
                    "<UserTable>"
                    
        Print #2, printStr
        
    Else: GoTo AfterXML
    
    End If
    
    
End If








' The software skips here if the first value in the XML table is blank
ContinueLoop:

' reset old field and old subfield
oldField = ""
oldSubf = ""

' iterate to the next row
progressCount = progressCount + 1
row = row + 1
Loop

ProgressForm.ProgressLabel.Width = ProgressForm.MaxWidth.Caption
ProgressForm.ProgressFrame.Caption = "100" & "%"

' Close the root XML tag
Print #2, "</UserTable>"


AfterXML:

If master = "Master" Then GoTo MasterSkip3
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''' Make Error Statement ''''''''''''''''''''''''''''''''''''''''
Dim errStr As String
errStr = MakeErrStr(xmlInfo, "User")

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
    "Users Accepted: " & xmlInfo(0) & _
    vbNewLine & _
    "Users Declined: " & xmlInfo(1) & _
    vbNewLine & vbNewLine & _
    "--------------------------------------------------------------------------" & vbNewLine & _
    "--------------------------------------------------------------------------"
      
        
DiaStr = DiaStr & _
    vbNewLine & vbNewLine & _
    "Errors: " & xmlInfo(1)

If errStr <> "" Then

    DiaStr = DiaStr & _
        vbNewLine & vbNewLine & _
        "Any Users that you attempted to include in your XML document that were rejected are highlighted " & _
        "in blue." & _
        vbNewLine & vbNewLine & _
        "The following cells contain invalid entries and are stopping some facilities from being included in the XML document: " & _
        vbNewLine & vbNewLine & _
        errStr

Else
    DiaStr = DiaStr & vbNewLine & vbNewLine & _
        "All User information has been converted to XML."

End If

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


Application.EnableEvents = True
MasterSkip3:

End Sub



'' UserXMLTable
'' Daniel Slosky
'' Last Updated: 2/12/2015
''
'' This software works in conjunction with FacilityXML to generate an XML document that holds
'' facility data for ShakeCast
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Function UserXMLTable()

'-------------------------------- Create Variables to Track Errors ---------------------------------

Dim xmlNums(0 To 1) As Variant ' this will be the return variable with two values,
                               ' the number of accepted and declined facilities
xmlNums(0) = 0
xmlNums(1) = 0

Dim ErrAdd() As Variant          ' ErrAdd is an array that holds the addresses of the cells with
                                       ' errors

'---------------------------------------------------------------------------------------------------




' create some handles to access the worksheets
Set mySheet = Worksheets("User XML")
Set XMLSheet = Worksheets("UserXMLexport")

' Start with a fresh XML table
Dim lastRow As Integer
Dim lastCol As Integer
lastRow = mySheet.Cells(Rows.count, "F").End(xlUp).row


LastXML = XMLSheet.Cells(Rows.count, "A").End(xlUp).row

If LastXML < 4 Then
    LastXML = 4
End If

XMLSheet.Range("A4:" & "K" & LastXML).Clear


' These values are used to determine starting and ending position for the loops used to read the
' facility worksheet

Dim startRow As Integer
Dim startCol As Integer
Dim endRow As Integer
Dim endCol As Integer

startRow = 4
startCol = 1
endRow = mySheet.Cells(Rows.count, "F").End(xlUp).row
endCol = 10

' Create variable to monitor how many rows of the XML table are filled. The integer keeps track of
' the NEXT ROW TO BE FILLED
Dim XMLcol As Integer
Dim XMLrow As Integer
XMLcol = 1
XMLrow = 4

' set up progress bar
Dim progressCount As Long
Dim progressWhen As Long
Dim pcntDone As Double

progressCount = 0
progressWhen = endRow * 0.01
pcntDone = 0
ProgressForm.ProcessName.Caption = "Making User XML Table"

Dim RowRange1 As String
Dim RowRange2 As String

Dim UserRow As Integer
UserRow = startRow
Do While UserRow < endRow + 1

    If progressCount > progressWhen Then
        
        pcntDone = ((UserRow - 4) / (endRow - 3))
        
        ProgressForm.ProgressLabel.Width = pcntDone * ProgressForm.MaxWidth.Caption
        ProgressForm.ProgressFrame.Caption = Round(pcntDone * 100, 0) & "%"
        
        DoEvents
        progressCount = 0
        
    End If

    ' gets rid of empty rows before we call them errors
    If IsEmpty(mySheet.Range("K" & UserRow)) Then GoTo ContinueLoop
    
    If mySheet.Range("K" & UserRow).value = "Bad" Then
        
        CheckUsers UserRow
        
        xmlNums(1) = xmlNums(1) + 1
        
        ReDim Preserve ErrAdd(0 To (xmlNums(1) - 1))
        ErrAdd(xmlNums(1) - 1) = UserRow
        
        
        GoTo ContinueLoop


    End If
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' If this code is executing, the row is FULL '
    ''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ' Count the number of facilities that will make it all the way to the XML
    xmlNums(0) = xmlNums(0) + 1
    
    XMLcol = 1
    For UserCol = startCol To endCol
        
            ' This is the cell that we are about to write to in the XML table
            Set XMLcell = XMLSheet.Cells(XMLrow, XMLcol)
            Set UserCell = mySheet.Cells(UserRow, UserCol)
        
            XMLcell.value = UserCell.value
            
            XMLcol = XMLcol + 1
            
'        End If
        
    Next UserCol
    
    XMLrow = XMLrow + 1
    
ContinueLoop:
progressCount = progressCount + 1
UserRow = UserRow + 1
    
Loop

ProgressForm.ProgressLabel.Width = ProgressForm.MaxWidth.Caption
ProgressForm.ProgressFrame.Caption = "100" & "%"

Dim exportInfo(0 To 3) As Variant

exportInfo(0) = xmlNums(0)
exportInfo(1) = xmlNums(1)
exportInfo(2) = ErrAdd

UserXMLTable = exportInfo
End Function


Sub UserXMLButton()
    UserXML "NotMaster"
End Sub

Sub UpdateGroupsButton()

Application.EnableEvents = False

' create some handles to access the worksheets
Set mySheet = Worksheets("User XML")

' Start with a fresh XML table
Dim lastRow As Integer
Dim lastCol As Integer
lastRow = mySheet.Cells(Rows.count, "F").End(xlUp).row


' These values are used to determine starting and ending position for the loops used to read the
' facility worksheet

Dim startRow As Integer
Dim startCol As Integer
Dim endRow As Integer
Dim endCol As Integer

startRow = 4
startCol = 1
endRow = mySheet.Cells(Rows.count, "F").End(xlUp).row
endCol = 10

' Create variable to monitor how many rows of the XML table are filled. The integer keeps track of
' the NEXT ROW TO BE FILLED
Dim XMLcol As Integer
Dim XMLrow As Integer
XMLcol = 1
XMLrow = 4

Dim RowRange1 As String
Dim RowRange2 As String

Dim progressCount As Long
Dim progressWhen As Long
Dim pcntDone As Double

progressCount = 0
progressWhen = endRow * 0.01
prcntDone = 0
ProgressForm.ProcessName.Caption = "Updating Worksheet"
    

Dim UserRow As Integer
UserRow = startRow
Do While UserRow < endRow + 1

    If progressCount > progressWhen Then
        
        pcntDone = ((UserRow - 4) / (endRow - 3))
        
        ProgressForm.ProgressLabel.Width = pcntDone * ProgressForm.MaxWidth.Caption
        ProgressForm.ProgressFrame.Caption = Round(pcntDone * 100, 0) & "%"
        
        DoEvents
        progressCount = 0
        
    End If

    CheckUsers UserRow
    
ContinueLoop:
progressCount = progressCount + 1
UserRow = UserRow + 1
    
Loop

ProgressForm.ProgressLabel.Width = ProgressForm.MaxWidth.Caption
ProgressForm.ProgressFrame.Caption = "100" & "%"

Unload ProgressForm

End Sub


Private Sub UserSheetInfo()
FacSheetForm.SheetInfoName.Caption = "The User Spreadsheet"
FacSheetForm.DialogueBox.Text = "This spreadsheet can be used to create user accounts for ShakeCast and link them to notification groups " & _
        "as defined in the previous spreadsheet. By filling this spreadsheet out from right to left, we will be " & _
        "able to check your inputs as you go. After you enter information into a cell, hit the tab or enter key " & _
        "to submit the information. If you click elsewhere on the spreadsheet while Excel is still in data entry " & _
        "mode, it is possible that we will not be able to validate your information." & vbNewLine & vbNewLine & _
        "You will only be able to select notification groups the those you defined in the notification spreadsheet, " & _
        "so it is important to fill that one out first. If a notification group is selected for a user then later " & _
        "deleted, the group name will be removed from that user's information before an XML document is exported. " & _
        "The expired notification group will also be removed from user rows in you hit the ""Update Spreadsheet"" button. " & _
        "Unless you plan to set up your own email server, just leave the delivery email addresses as ""shakecast@usgs.gov"". " & _
        "The ShakeCast system is preconfigured to send emails from this address." & vbNewLine & vbNewLine & _
        "If you have not filled out all the required fields for a specific user, their user row will be turned blue. " '& _

FacSheetForm.AdvUser.Visible = False
FacSheetForm.Show

End Sub

Private Sub makeGroupChecklist()
Set groupSheet = Worksheets("Notification XML")

' get group list
' create a string array of all the group names. This means going into the group spreadsheet and reading
' all the names of the groups without counting repeated names
Dim GroupNames() As String
Set allGroupCells = groupSheet.Range("A:A")

Dim oldGroup As String
Dim curGroup As String
Dim GroupCount As Integer
Dim blankCount As Integer

GroupCount = 0
blankCount = 0
For Each groupCell In allGroupCells

    If groupCell.row < 4 Then GoTo NextGroupCell
    
    If IsEmpty(groupCell) Then
        blankCount = blankCount + 1
        GoTo NextGroupCell
    End If
    
    curGroup = groupCell.value
    
    If groupCell.row = 4 Then
    
        If groupSheet.Range("N" & groupCell.row).value = "Good" Then
            ReDim Preserve GroupNames(0 To 0)
            GroupNames(0) = curGroup
        End If
        
    ElseIf curGroup <> oldGroup Then
        
        If groupSheet.Range("N" & groupCell.row).value = "Good" Then
            GroupCount = GroupCount + 1
            ReDim Preserve GroupNames(0 To GroupCount)
        
            GroupNames(GroupCount) = curGroup
        
        End If
        
        ' reset blankCount to zero, since it is really trying to count the blank rows at the END of the
        ' spreadsheet
        blankCount = 0
        
    End If
        
    oldGroup = curGroup
NextGroupCell:
If blankCount > 10 Then GoTo QuitGroupLoop
Next groupCell
QuitGroupLoop:

' get the groups from the current cell
Dim selGroups() As String
If InStr(ActiveCell.value, ":") Then
    selGroups = Split(ActiveCell.value, ":")
Else
    ReDim selGroups(0 To 0) As String
    selGroups(0) = ActiveCell.value
    
End If

' turn list into check boxes
Dim curColumn   As Long
Dim lastRow     As Long
Dim i           As Long
Dim a           As Integer
Dim chkBox      As MSForms.CheckBox

a = 0
On Error GoTo TheEnd
For i = 0 To UBound(GroupNames)
    
    
    If GroupNames(i) = "" Then
        a = a + 1
        GoTo NextOne
    End If

    Set chkBox = GroupCheckBox.Controls.Add("Forms.CheckBox.1", "CheckBox_" & i)
    chkBox.Caption = GroupNames(i)
    chkBox.Left = 5
    chkBox.Top = 5 + (i * 20 - a * 20)
    chkBox.Font.Size = 12
    chkBox.AutoSize = True
    
    ' select the right checkboxes
    If InArray(selGroups, GroupNames(i)) Then
        chkBox.value = True
    End If
    
NextOne:
Next i

' Keep the number of check boxes we've created to reference later in a hidden label
GroupCheckBox.GroupCount.Caption = i

TheEnd:
If Err Then
MsgBox "You have to define a notification group before using this column!"
End If

End Sub

Private Sub userUnlock()

    Set mySheet = Worksheets("User XML")

    ' figure out the used range of the workbook
    Dim startRow As Integer
    Dim endRow As Integer
    Dim startCol As String
    Dim endCol As String
    
    startRow = 4
    startCol = "A"
    endCol = "L"
    
    endRow = mySheet.Cells(Rows.count, "A").End(xlUp).row
    
    mySheet.Range(startCol & startRow, endCol & endRow).Locked = False
    

End Sub
