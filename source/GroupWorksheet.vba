Attribute VB_Name = "GroupWorksheet"
Sub CheckGroups()

Set mySheet = Worksheets("Notification XML")

mySheet.Unprotect

' keep the worksheet looking at the same cell while code executes
mySheet.ScrollArea = ActiveCell.Address

Dim lastRow As Integer

If mySheet.Cells(Rows.count, "N").End(xlUp).row > mySheet.Cells(Rows.count, "A").End(xlUp).row Then
    lastRow = mySheet.Cells(Rows.count, "N").End(xlUp).row ' where we stop!
Else
    lastRow = mySheet.Cells(Rows.count, "A").End(xlUp).row
End If

' make variable to track the active cell when we started this subroutine
Dim startActive As Range

Set startActive = ActiveCell

' make an array that holds all the handles to a group and an int to help vary the length
' of the array
Dim group() As Variant
Dim num As Integer

num = 0

' make array to hold empty cells and an int to vary the length of the array
Dim emptyCells As Variant
Dim emptyNum As Integer

emptyNum = 0

Dim cellRange As Range

' make some variables to keep track of the group name, and past group name
Dim curGroup As String
Dim pastGroup As String
Dim groupColor As String
Dim LetStr As String        ' Keeps track of the columns which we want to mirror: B C G J K L M
pastGroup = ""
groupColor = "Yellow"
LetStr = "BCGJKLM"


' determine if we should refresh the whole table or just the group
Set colA = mySheet.Range("A:A")
Dim endAddress As String


' white out all the empty cells
Dim sheetRange As Range
Set sheetRange = mySheet.Range("A:M")

sheetRange.SpecialCells(xlCellTypeBlanks).Interior.ColorIndex = 2


TheLoop:

' first we change back the color in the header
' make the title banner stay the same color even though we are whiting out blank cells
Dim titleRange As Range
Set titleRange = mySheet.Range("A1:P2")

If mySheet.Range("A2") = "General User" Then
    With titleRange.Interior
        .Color = RGB(196, 215, 155)
    End With
Else
    With titleRange.Interior
        .Color = RGB(192, 80, 77)
    End With
End If

' run through all the group names
For Each cell In colA.Cells
    
    curGroup = cell.Value               ' get the group name for the current row

    
    
    If cell.row > 4 Then
        testEmptyName pastGroup, (cell.row - 1)
    End If
    
    If curGroup <> pastGroup And cell.row >= 4 Then

    
        
    
        ' now we have an entire group in the variant so we can start messing with it
    
        ' check if the group is more than one row
        
        ' if not, change it's color and validate that the nececessary fields are filled
        If num = 1 Then
        
            ' select the cell range for the single row
            Set cellRange = Range("A" & (cell.row - 1), "P" & (cell.row - 1))
            
            ' validate that the single row group is valid
            
            If IsEmpty(mySheet.Range("A" & cell.row - 1)) Or _
                    IsEmpty(mySheet.Range("B" & cell.row - 1)) Or _
                    IsEmpty(mySheet.Range("C" & cell.row - 1)) Or _
                    IsEmpty(mySheet.Range("D" & cell.row - 1)) Or _
                    IsEmpty(mySheet.Range("G" & cell.row - 1)) Or _
                    IsEmpty(mySheet.Range("H" & cell.row - 1)) Then
                
                
                If WorksheetFunction.CountBlank(Range("A" & rowNum, "D" & rowNum)) > 3 Then
                
                    mySheet.Range("A" & (cell.row - 1), "P" & (cell.row - 1)).Clear
                    mySheet.Range("A" & (cell.row - 1), "P" & (cell.row - 1)).Locked = False
                    
                    If groupColor = "Yellow" Then
                        groupColor = "Red"
                    Else
                        groupColor = "Yellow"
                    End If
                    
                    GoTo NextGroup
                    
                End If
                
                FillGroup (cell.row - 1)
                
'                On Error GoTo MakeDropDownsS
'                If mySheet.Range("B" & cell.row - 1).Validation.Formula1 = "" Or _
'                         mySheet.Range("D" & cell.row - 1).Validation.Formula1 = "" Or _
'                         mySheet.Range("E" & cell.row - 1).Validation.Formula1 = "" Or _
'                         mySheet.Range("G" & cell.row - 1).Validation.Formula1 = "" Or _
'                         mySheet.Range("H" & cell.row - 1).Validation.Formula1 = "" Then
'
'MakeDropDownsS:
                    GroupDropDowns (cell.row - 1)
                        
                        
                    Err.Clear
'                End If
    
                If Not IsEmpty(mySheet.Range("A" & cell.row - 1)) And _
                        Not IsEmpty(mySheet.Range("B" & cell.row - 1)) And _
                        Not IsEmpty(mySheet.Range("C" & cell.row - 1)) And _
                        Not IsEmpty(mySheet.Range("D" & cell.row - 1)) And _
                        Not IsEmpty(mySheet.Range("G" & cell.row - 1)) And _
                        Not IsEmpty(mySheet.Range("H" & cell.row - 1)) Then GoTo GoodRow
                
                
                ChangeColors "Bad", Range("A" & (cell.row - 1), "M" & (cell.row - 1)), "Group"
                mySheet.Range("N" & (cell.row - 1)).Value = "Bad"
                mySheet.Range("P" & (cell.row - 1)).Value = "Blue"
                
                greyCells (cell.row - 1), (cell.row - 1)
                
'                If groupColor = "Yellow" Then
'                    groupColor = "Red"
'                Else
'                    groupColor = "Yellow"
'                End If

                GoTo NextGroup
                
            End If
GoodRow:
            
            ' change the group color to yellow
            If groupColor = "Yellow" Then
                With cellRange.Interior
'                    .Pattern = xlSolid
'                    .PatternColorIndex = xlAutomatic
'                    .Color = 6750207
'                    .TintAndShade = 0
'                    .PatternTintAndShade = 0
                    .ColorIndex = 36
                    
                End With
                
                mySheet.Range("P" & (cell.row - 1)).Value = "Yellow"
            ' change the group color to green
            Else
            
                With cellRange.Interior
'                    .Pattern = xlSolid
'                    .PatternColorIndex = xlAutomatic
'                    .Color = 5296274
'                    .TintAndShade = 0
'                    .PatternTintAndShade = 0
                    .Color = RGB(230, 166, 121)
                    
                End With
                
                
                mySheet.Range("P" & (cell.row - 1)).Value = "Red"
            End If

                mySheet.Range("N" & (cell.row - 1)).Value = "Good"
                

                greyCells (cell.row - 1), (cell.row - 1)
                FillGroup (cell.row - 1)
                GroupDropDowns (cell.row - 1)
            
    
        
        Else
        
        
            ' This is the cell range for the whole group... we might use this later
            'Set cellRange = Range("A" & (cell.row - num), "M" & (cell.row - 1))
            
            For rowNum = cell.row - num To cell.row - 1

                Set cellRange = Range("A" & rowNum, "M" & rowNum)
                
                If ActiveCell.column = 4 Or _
                    ActiveCell.column = 5 Then
                    
                    FillGroup (rowNum)
                    
                End If
                
                If rowNum = cell.row - num And _
                        Application.WorksheetFunction.CountBlank(mySheet.Range("A" & rowNum, "H" & rowNum)) > 1 Then
                    
                    FillGroup (rowNum)
                    GroupDropDowns (rowNum)
          
                ElseIf Application.WorksheetFunction.CountBlank(mySheet.Range("A" & rowNum, "H" & rowNum)) > 2 Then
               
                    FillGroup (rowNum)
                    GroupDropDowns (rowNum)

                End If
               
                ' we want to send the first row straight to ChangeColors if it
                ' is missing any of these things. Any following row needs less...
                If rowNum = cell.row - num And _
                    (IsEmpty(mySheet.Range("A" & rowNum)) Or _
                    IsEmpty(mySheet.Range("B" & rowNum)) Or _
                    IsEmpty(mySheet.Range("C" & rowNum)) Or _
                    IsEmpty(mySheet.Range("D" & rowNum)) Or _
                    IsEmpty(mySheet.Range("M" & rowNum))) Then
                
                    ChangeColors "Bad", Range("A" & rowNum, "M" & rowNum), "Group"
                    
                    mySheet.Range("N" & rowNum).Value = "Bad"
                    
                    greyCells rowNum, (cell.row - num)
                    
                    GoTo NextGroupItem
                    
                End If
                
                ' check the following rows for only the info that won't be assigned
                ' from the first row "BCGJKLM". If the first row is being evaluated
                ' it will pass all of these checks!
                If IsEmpty(mySheet.Range("A" & rowNum)) Or _
                    IsEmpty(mySheet.Range("D" & rowNum)) Or _
                    (IsEmpty(mySheet.Range("G" & rowNum)) And _
                    IsEmpty(mySheet.Range("D" & rowNum))) Or _
                    IsEmpty(mySheet.Range("M" & rowNum)) Or _
                    IsEmpty(mySheet.Range("A" & cell.row - num)) Or _
                    IsEmpty(mySheet.Range("B" & cell.row - num)) Or _
                    IsEmpty(mySheet.Range("C" & cell.row - num)) Or _
                    IsEmpty(mySheet.Range("D" & cell.row - num)) Or _
                    (IsEmpty(mySheet.Range("G" & cell.row - num)) And _
                    IsEmpty(mySheet.Range("D" & cell.row - num))) Or _
                    IsEmpty(mySheet.Range("H" & cell.row - num)) Then
                
                    ChangeColors "Bad", Range("A" & rowNum, "M" & rowNum), "Group"
                
                    mySheet.Range("N" & rowNum).Value = "Bad"
                    greyCells rowNum, cell.row - num
                    
                    GoTo NextGroupItem
                    
                ElseIf groupColor = "Yellow" Then
                    With cellRange.Interior
                         .ColorIndex = 36
                    End With
                
                    mySheet.Range("P" & rowNum).Value = "Yellow"
                Else
                    With cellRange.Interior
                         .Color = RGB(230, 166, 121)
                    End With
                    
                    mySheet.Range("P" & rowNum).Value = "Red"
            
                End If
                
                
                mySheet.Range("N" & rowNum).Value = "Good"
                
                ' mirror cells and black out unavailable parameters
                greyCells rowNum, cell.row - num
                
                        
NextGroupItem:
            Next rowNum

        
        
        End If
        
        
        ' switch colors for the next group
        If groupColor = "Yellow" Then
            groupColor = "Red"
        Else
            groupColor = "Yellow"
        End If
        
NextGroup:
        
        num = 0
        ReDim group(0 To num) As Variant

    End If
    
    
    If cell.row > 3 Then
        ReDim Preserve group(0 To num) As Variant     ' change the length of the array to hold another value
        group(num) = cell                   ' add the current cell to the
    
        num = num + 1
        
        pastGroup = curGroup
        
        
        ' test if we should keep the hidden autofilled columns or not
        For Each cellTest In Range("A" & cell.row & ":D" & cell.row)
        
            ' if any of the non-autofilled columns are filled, we skip clearing the hidden rows
            If Not IsEmpty(cellTest) Then
                GoTo TestDone
            End If
            
        Next cellTest
        mySheet.Range("A" & cell.row, "P" & cell.row).Clear
        mySheet.Range("A" & cell.row, "P" & cell.row).Locked = False
        
TestDone:
            
        
    End If
    
    
    If cell.row > lastRow + 1 Then GoTo endLoop
    
Next cell

endLoop:

' bring the view back to the current cell
On Error Resume Next
startActive.Activate
Err.Clear
'Application.Run "protectWorkbook"

End Sub


Sub greyCells(ByVal rowNum As Integer, _
                    startRow As Integer)

' use the groups worksheet
Set mySheet = Worksheets("Notification XML")

' string with the letters of the columns we want to
Dim LetStr As String
LetStr = "BCHIJL"

If rowNum > startRow Then

    For letNum = 1 To Len(LetStr)
    
        ' "C" is where the monitoring region lives, and we only want this defined once!
        If Mid(LetStr, letNum, 1) <> "C" Then
            mySheet.Range(Mid(LetStr, letNum, 1) & rowNum).Value = _
                mySheet.Range(Mid(LetStr, letNum, 1) & startRow).Value
        Else
        
            ' makes all group items following the first
            mySheet.Range(Mid(LetStr, letNum, 1) & rowNum) = Empty
        End If
                        
                        
                        
        With mySheet.Range(Mid(LetStr, letNum, 1) & rowNum).Interior
             .ColorIndex = 16
        End With
                        
                        
    Next letNum
    
    
End If



If mySheet.Range("D" & rowNum).Value = "NEW_EVENT" Then
                
    With mySheet.Range("E" & rowNum).Interior
        .ColorIndex = 16
    End With
            
ElseIf mySheet.Range("D" & rowNum).Value = "DAMAGE" Then
    
    With mySheet.Range("F" & rowNum).Interior
        .ColorIndex = 16
    End With
                
End If

mySheet.ScrollArea = ""

End Sub

Sub FillGroup(rowNum As Integer)

Set mySheet = Worksheets("Notification XML")

If WorksheetFunction.CountBlank(Range("A" & rowNum, "D" & rowNum)) < 4 Then

' autofil the minimum magnitude and inspection priority if appropriate
If mySheet.Range("D" & rowNum).Value = "NEW_EVENT" Then
    
    mySheet.Range("E" & rowNum) = Empty
                
    If IsEmpty(mySheet.Range("F" & rowNum).Value) Then
        mySheet.Range("F" & rowNum).Value = 3
    End If
            
ElseIf mySheet.Range("D" & rowNum).Value = "DAMAGE" Then
    
    mySheet.Range("F" & rowNum) = Empty
    
            
    If IsEmpty(mySheet.Range("E" & rowNum).Value) Then
        mySheet.Range("E" & rowNum).Value = "GREEN"
    End If
                
End If

' Event_type, Notification format, Product Type, METRIC, Aggregate Flag, Aggregate Group
If IsEmpty(mySheet.Range("G" & (rowNum))) Or _
        IsError(mySheet.Range("G" & (rowNum))) And Not _
        IsEmpty(mySheet.Range("A" & (rowNum))) Then
    mySheet.Range("G" & (rowNum)).Value = "ACTUAL"
ElseIf IsEmpty(mySheet.Range("A" & (rowNum))) Then
    mySheet.Range("G" & (rowNum)) = Empty
End If

If IsEmpty(mySheet.Range("H" & (rowNum))) Or _
        IsError(mySheet.Range("H" & (rowNum))) And Not _
        IsEmpty(mySheet.Range("A" & (rowNum))) Then
    mySheet.Range("H" & (rowNum)).Value = "Rich Content"
ElseIf IsEmpty(mySheet.Range("A" & (rowNum))) Then
    mySheet.Range("H" & (rowNum)) = Empty
End If

If IsEmpty(mySheet.Range("J" & (rowNum))) Or _
        IsError(mySheet.Range("J" & (rowNum))) Then
    mySheet.Range("J" & (rowNum)).Value = Empty
ElseIf IsEmpty(mySheet.Range("A" & (rowNum))) Then
    mySheet.Range("J" & (rowNum)) = Empty
End If
            
' the default for metric is PGA
'If IsEmpty(mySheet.Range("K" & (rowNum))) Or _
'            IsError(mySheet.Range("K" & (rowNum))) Then
'    mySheet.Range("K" & (rowNum)).Value = "PGA"
'End If

If mySheet.Range("D" & (rowNum)).Value = "DAMAGE" And IsEmpty(mySheet.Range("K" & (rowNum))) Then
    mySheet.Range("K" & (rowNum)).Value = Empty
ElseIf IsEmpty(mySheet.Range("D" & (rowNum))) Then
    mySheet.Range("K" & (rowNum)) = Empty
End If

If mySheet.Range("D" & (rowNum)).Value = "NEW_EVENT" And IsEmpty(mySheet.Range("K" & (rowNum))) Then
    mySheet.Range("K" & (rowNum)).Value = Empty
ElseIf IsEmpty(mySheet.Range("D" & (rowNum))) Then
    mySheet.Range("K" & (rowNum)) = Empty
End If
            
If IsEmpty(mySheet.Range("L" & (rowNum))) Or _
        IsError(mySheet.Range("L" & (rowNum))) And Not _
        IsEmpty(mySheet.Range("A" & (rowNum))) Then
    mySheet.Range("L" & (rowNum)).Value = 1
ElseIf IsEmpty(mySheet.Range("A" & (rowNum))) Then
    mySheet.Range("L" & (rowNum)) = Empty
End If
            
If IsEmpty(mySheet.Range("M" & (rowNum))) Or _
        IsError(mySheet.Range("M" & (rowNum))) And Not _
        IsEmpty(mySheet.Range("A" & (rowNum))) Then
    mySheet.Range("M" & (rowNum)).Value = "Default"
ElseIf IsEmpty(mySheet.Range("A" & (rowNum))) Then
    mySheet.Range("M" & (rowNum)) = Empty
End If

End If

End Sub

Sub GroupDropDowns(rowNum As Integer)

Set mySheet = Worksheets("Notification XML")

' add drop down menus to Facility Type, Notification Type, Inspection Priority, Event Type, and Notification Format
Set FacType = mySheet.Range("B" & (rowNum))
Set NotType = mySheet.Range("D" & (rowNum))
Set InsPrio = mySheet.Range("E" & (rowNum))
Set EvType = mySheet.Range("G" & (rowNum))
Set NotForm = mySheet.Range("H" & (rowNum))

' get all the facility types in a range element
Set LookUpSheet = Worksheets("ShakeCast Ref Lookup Values")

Dim lastFac As Integer
lastFac = LookUpSheet.Cells(Rows.count, "C").End(xlUp).row
Set FacTypeCells = LookUpSheet.Range("C1:C" & lastFac)

' create an array to hold the items for the Notification Type
Dim NotTypes(0 To 1) As String
NotTypes(0) = "NEW_EVENT"
NotTypes(1) = "DAMAGE"

' create an array for inspection priority
Dim InsPrios(0 To 3) As String
InsPrios(0) = "GREEN"
InsPrios(1) = "YELLOW"
InsPrios(2) = "ORANGE"
InsPrios(3) = "RED"

' create an array for event type
Dim EvTypes() As String
If mySheet.Range("A2").Value = "Advanced User" Then
    ReDim EvTypes(0 To 3)
    EvTypes(0) = "ACTUAL"
    EvTypes(1) = "SCENARIO"
    EvTypes(2) = "HEARTBEAT"
    EvTypes(3) = "ALL"
    
Else
    ReDim EvTypes(0 To 2)
    EvTypes(0) = "ACTUAL"
    EvTypes(1) = "SCENARIO"
    EvTypes(2) = "HEARTBEAT"
End If


' create an array for notification format
Dim NotForms(0 To 2) As String
NotForms(0) = "Rich Content"
NotForms(1) = "Plain Text"
NotForms(2) = "PAGER"


' now that we have the arrays defined, we can create the drop down menus
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

With NotType.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
    Operator:=xlBetween, Formula1:=Join(NotTypes, ",")
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = "Notification Type"
    .ErrorTitle = ""
    .InputMessage = "Please select a notification type from the drop-down list"
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True
End With

With InsPrio.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
    Operator:=xlBetween, Formula1:=Join(InsPrios, ",")
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = "Inspection Priority"
    .ErrorTitle = ""
    .InputMessage = "Please select an inspection priority from the drop-down list"
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True
End With

With EvType.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
    Operator:=xlBetween, Formula1:=Join(EvTypes, ",")
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = "Event Type"
    .ErrorTitle = ""
    .InputMessage = "Please select an event type from the drop-down list"
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True
End With

With NotForm.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
    Operator:=xlBetween, Formula1:=Join(NotForms, ",")
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = "Notification Format"
    .ErrorTitle = ""
    .InputMessage = "Please select a notification format from the drop-down list"
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True
End With
 


End Sub


Sub testEmptyName(curGroup As String, _
                    rowNum As Integer)

Set mySheet = Worksheets("Notification XML")


' check if the group name is empty, if so we have a bad group
If curGroup = "" Then
            
    ' test if we should keep the hidden autofilled columns or not
    For Each cellTest In mySheet.Range("B" & rowNum & ":G" & rowNum)
        
        ' if any of the non-autofilled columns are filled, we skip clearing the hidden rows
        If Not IsEmpty(cellTest) Then
            ChangeColors "Bad", mySheet.Range("A" & rowNum, "M" & rowNum), "Group"
            mySheet.Range("P" & rowNum).Value = "Blue"
            GoTo TestDone
        End If
            
    Next cellTest
            
    mySheet.Range("A" & rowNum, "P" & rowNum).Clear
    mySheet.Range("A" & rowNum, "P" & rowNum).Locked = False
End If
            
TestDone:


End Sub


'' GroupXML
'' Daniel Slosky
'' Last Updated: 2/24/2016
'' Creates an XML document that holds all the information the user input into the Group spreadsheet.
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
Sub GroupXML(master As String, _
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

'On Error GoTo XMLFinish
'On Error Resume Next


If master = "Master" Then GoTo MasterSkip1

Application.EnableEvents = False

Close #2

'Dim refreshInfo() As Variant
'refreshInfo = Application.Run("RefreshFormulas") ' enter formulas into any fields that should hold formulas, but are empty

' We now get XML info from the worksheet GroupXMLexport, and we have to move all the info over there

Dim xmlInfo() As Variant
xmlInfo = Application.Run("GroupXMLTable")       ' This function populates the XML table in the HIDDEN
                                                 ' spreadsheet GroupXMLexport



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
    docStr = "GroupXML.xml"
Else
    docNum = Application.WorksheetFunction.Ceiling(docNum, 1)
    docStr = "GroupXML1.xml"
    For i = 2 To docNum
        docStr = docStr & "," & "GroupXML" & i & ".xml"
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

Set XMLSheet = Worksheets("GroupXMLexport")

Dim startCol As Integer                      ' Figure out which rows and columns we want to sort through
Dim startRow As Integer
Dim endCol As Integer
Dim endRow As Integer


startCol = 1                                 ' The first column that holds XML info
startRow = 4                                 ' The First Row with XML info


' The last column we will look at
endCol = XMLSheet.Cells(3, Columns.count).End(xlToLeft).column

' This will pull cells with formulas and no values, but we can filter those out!
endRow = XMLSheet.Cells(Rows.count, "B").End(xlUp).row + 1




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''' BEGIN XML '''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim printStr As String
printStr = ""

If master = "Master" Then GoTo MasterSkip2
printStr = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbNewLine
MasterSkip2:

printStr = printStr & "<GroupTable>"
     
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
ProgressForm.ProcessName.Caption = "Exporting Group XML"


' This loop will run through all the cells in the XML region of the document.
Dim row As Integer
Dim indent As Integer
row = 4                   ' The first row that holds data
indent = 0                ' Keeps track of how many parents each tag has
Do While row < endRow


    'Set TestCell = XMLSheet.Cells(row, 1) ' This is the first XML cell, which should be filled for any valid entry
    If progressCount > progressWhen Then
        
        pcntDone = ((row - 4) / (endRow - 3))
        
        ProgressForm.ProgressLabel.Width = pcntDone * ProgressForm.MaxWidth.Caption
        ProgressForm.ProgressFrame.Caption = Round(pcntDone * 100, 0) & "%"
        
        DoEvents
        progressCount = 0
        
    End If
        
        
        ' Since there IS data for this row, we start a new entry in the XML page
        printStr = vbTab & "<GroupRow>" & vbNewLine
        
        ' Now we can run through all the columns in the row. Row 32 is the first one to hold XML data
        For col = 1 To endCol
        
            ' Assign the field, subfield, and header for the current cell.
            Set Field = XMLSheet.Cells(1, col)
            Set Subf = XMLSheet.Cells(2, col)
            Set Head = XMLSheet.Cells(3, col)
            Set CurCell = XMLSheet.Cells(row, col)
            
            ' Check to see if the cell is empty, and if we don't want the empty tags, we skip this column
            ' If IsEmpty(CurCell) And Head.Value = "POLY" Then GoTo NextColumn
            If IsEmpty(CurCell) Then GoTo NextColumn
            
            
'                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                REMOVE NUMBERS FROM FIELD AND SUBFIELD NAMES AND HEADERS

            ' Some of the column heads have numbers in them because their names are duplicates. This gets rid of
            ' those numbers
            Dim FieldValue As String
            Dim SubfValue As String
            Dim HeadValue As String
            
            FieldValue = Field.Value
            SubfValue = Subf.Value
            HeadValue = Head.Value
            CellValue = CurCell.Value
            
            ' This will remove all numbers from the headers... but why am I doing it this way??
            For i = 0 To 9

                    FieldValue = Replace(FieldValue, i, "")
                
                    SubfValue = Replace(SubfValue, i, "")
                    
                    HeadValue = Replace(HeadValue, i, "")
                
             Next i
             
            ' replace XML and HTML reserved characters
             CellValue = Replace(CellValue, "&", "&amp;")
             CellValue = Replace(CellValue, "<", "&lt;")
             CellValue = Replace(CellValue, ">", "&gt;")
             CellValue = Replace(CellValue, "'", "&apos;")
             
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
            
                CurCellValue = CellValue
                
            End If
'                      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            
            If Field.Value = Empty Then            ' There can be no subfield without a field, this piece of data
                indent = 2                         ' only has its GroupRow as a parent. We open no additional
                                                   ' XML tages for this data
                
                
                ' Save line into a string which can later be printed to XML
                printStr = printStr & vbTab & vbTab & "<" & HeadValue & _
                ">" & CurCellValue & "</" & Head & ">"
        
            ElseIf Field.Value <> Empty And Subf.Value = Empty Then
                indent = 2
                
                                                   ' In this case, the data resides inside of a field, but no sub-
                                                   ' field. This means we need to open the field tag if it hasn't
                                                   ' already been opened
                
                ' Check if the field has already been opened
                If Field.Value <> oldField Then
                    
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
                If Field.Value <> oldField Then
                    printStr = vbTab & vbTab & "<" & FieldValue & ">" & vbNewLine
                End If
                
                indent = indent + 1
                
                ' check if the subfield has already been opened
                If Subf.Value <> oldSubf Then
                    printStr = printStr & vbTab & vbTab & vbTab & "<" & SubfValue & ">" & vbNewLine
                End If
                
                indent = indent + 1
                
                ' print the info for the data
                printStr = printStr & vbTab & vbTab & vbTab & vbTab & _
                        "<" & HeadValue & ">" & CurCellValue & "</" & HeadValue & ">"
                
            End If
            
            
            ' check if we should close the subfield
            Set NextSubf = XMLSheet.Cells(2, col + 1)
            If Subf.Value <> NextSubf.Value And Subf.Value <> Empty Then
            
                printStr = printStr & vbNewLine & vbTab & vbTab & vbTab & "</" & SubfValue & ">"
                
            End If
            
            ' check if we should close the field
            Set NextField = XMLSheet.Cells(1, col + 1)
            If Field.Value <> NextField.Value And Field.Value <> Empty Then

                printStr = printStr & vbNewLine & vbTab & vbTab & "</" & FieldValue & ">"
                
            End If
            
            
            
            ' Print the string that we built!
            Print #2, printStr
            
            ' Sets the Field and subfield we just used as the old fields
            oldField = Field.Value
            oldSubf = Subf.Value
            
            ' Reset the print string for the next piece of data
            printStr = ""
           
        
NextColumn:
        Next col


' When finished going through the columns, we close the Group row
printStr = printStr & vbTab & "</GroupRow>"


Print #2, printStr

overFlowCount = overFlowCount + 1
' check if a new XML document should be started
If overFlowCount > docMax - 1 Then

    Print #2, "</GroupTable>"
    
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
                    "<GroupTable>"
           
        Print #2, printStr
        
    ElseIf row + 1 < endRow Then
    
        printStr = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbNewLine & _
                    "<GroupTable>"
                    
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
Print #2, "</GroupTable>"


AfterXML:

If master = "Master" Then GoTo MasterSkip3
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''' Make Error Statement ''''''''''''''''''''''''''''''''''''''''
Dim errStr As String
errStr = MakeErrStr(xmlInfo, "Group")

Dim DiaStr As String

DiaStr = "Your XML worksheet has been created. It's file location is: " & _
    vbNewLine & vbNewLine & XMLPath & _
    vbNewLine & vbNewLine & _
    "Groups Accepted: " & xmlInfo(0) & _
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
        "Any Groups that you attempted to include in your XML document that were rejected are highlighted " & _
        "in blue." & _
        vbNewLine & vbNewLine & _
        "The following cells contain invalid entries and are stopping some facilities from being included in the XML document: " & _
        vbNewLine & vbNewLine & _
        errStr

Else
    DiaStr = DiaStr & vbNewLine & vbNewLine & _
        "All Group information has been converted to XML."

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

Application.EnableEvents = True
MasterSkip3:

End Sub


'' GroupXMLTable
'' Daniel Slosky
'' Last Updated: 2/24/2015
''
'' This program will take information from the Notification XML sheet and move it to the
'' GroupXMLexport table
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Function GroupXMLTable()

'-------------------------------- Create Variables to Track Errors ---------------------------------

Dim xmlNums(0 To 1) As Variant ' this will be the return variable with two values,
                               ' the number of accepted and declined facilities
xmlNums(0) = 0
xmlNums(1) = 0

Dim ErrAdd() As Variant                ' ErrAdd is an array that holds the addresses of the cells with
                                       ' errors

'---------------------------------------------------------------------------------------------------




' create some handles to access the worksheets
Set notSheet = Worksheets("Notification XML")
Set XMLSheet = Worksheets("GroupXMLexport")

' Start with a fresh XML table
Dim lastRow As Integer
Dim LastCol As Integer
lastRow = notSheet.Cells(Rows.count, "A").End(xlUp).row


LastXML = XMLSheet.Cells(Rows.count, "A").End(xlUp).row

If LastXML < 4 Then
    LastXML = 4
End If

XMLSheet.Range("A4:" & "P" & LastXML).Clear

' These values are used to determine starting and ending position for the loops used to read the
' facility worksheet

Dim startRow As Integer
Dim startCol As Integer
Dim endRow As Integer
Dim endCol As Integer

startRow = 4
startCol = 1
endRow = notSheet.Cells(Rows.count, "A").End(xlUp).row
endCol = 13

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
ProgressForm.ProcessName.Caption = "Making Group XML Table"

Dim RowRange1 As String
Dim RowRange2 As String

Dim NotRow As Integer
NotRow = startRow
Do While NotRow < endRow + 1

    If progressCount > progressWhen Then
        
        pcntDone = ((NotRow - 4) / (endRow - 3))
        
        ProgressForm.ProgressLabel.Width = pcntDone * ProgressForm.MaxWidth.Caption
        ProgressForm.ProgressFrame.Caption = Round(pcntDone * 100, 0) & "%"
        
        DoEvents
        progressCount = 0
        
    End If

    RowRange1 = "A" & NotRow
    RowRange2 = "M" & NotRow
    
  
        'Set NotCell = NotSheet.Cells(NotRow, NotCol)

        If notSheet.Range("N" & NotRow).Value = "Bad" And Not _
                IsEmpty(notSheet.Range("N" & NotRow)) Then
        
            ' Count the number of facilities that will be rejected
            xmlNums(1) = xmlNums(1) + 1
                
                
            ' store bad group row numbers to generate errors later
            ReDim Preserve ErrAdd(0 To xmlNums(1) - 1) As Variant
            ErrAdd(xmlNums(1) - 1) = NotRow
            
            GoTo ContinueLoop
                
        ElseIf IsEmpty(notSheet.Range("N" & NotRow)) Then GoTo ContinueLoop
                
        End If
            

            
            



    ''''''''''''''''''''''''''''''''''''''''''''''
    ' If this code is executing, the row is GOOD '
    ''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ' Count the number of facilities that will make it all the way to the XML
    xmlNums(0) = xmlNums(0) + 1
    
    XMLcol = 1
    For NotCol = startCol To endCol

        ' This is the cell that we are about to write to in the XML table
        Set XMLcell = XMLSheet.Cells(XMLrow, XMLcol)
        Set NotCell = notSheet.Cells(NotRow, NotCol)
        
        Dim NotCellValue As String
        
        If NotCell.Value = "Rich Content" Then
        
            NotCellValue = "EMAIL_HTML"
            
        ElseIf NotCell.Value = "Plain Text" Then
            
            NotCellValue = "EMAIL_TEXT"
        
        Else
        
            NotCellValue = NotCell.Value
        End If
        
        
        
        XMLcell.Value = NotCellValue
            
        XMLcol = XMLcol + 1
        
    Next NotCol
    
    XMLrow = XMLrow + 1
    
ContinueLoop:
progressCount = progressCount + 1
NotRow = NotRow + 1
    
Loop

ProgressForm.ProgressLabel.Width = ProgressForm.MaxWidth.Caption
ProgressForm.ProgressFrame.Caption = "100" & "%"

Dim exportInfo(0 To 3) As Variant

exportInfo(0) = xmlNums(0)
exportInfo(1) = xmlNums(1)
exportInfo(2) = ErrAdd

GroupXMLTable = exportInfo
End Function



Sub GroupXMLButton()

    GroupXML "NotMaster"

End Sub

Private Sub groupUnlock()

    Set mySheet = Worksheets("Notification XML")

    ' figure out the used range of the workbook
    Dim startRow As Integer
    Dim endRow As Integer
    Dim startCol As String
    Dim endCol As String
    
    startRow = 4
    startCol = "A"
    endCol = "Q"
    
    endRow = mySheet.Cells(Rows.count, "A").End(xlUp).row
    
    mySheet.Range(startCol & startRow, endCol & endRow).Locked = False
    

End Sub

Sub GroupSheetInfo()

Set mySheet = Worksheets("Notification XML")

If mySheet.Range("A2") = "General User" Then
    GroupGenInfo
Else
    GroupAdvInfo
End If

FacSheetForm.Show

End Sub

Private Sub makeFacTypeChecklist()

AttCheckBox.Show

End Sub


Sub GroupGenInfo()

    FacSheetForm.AdvUser.Caption = "Access Advanced User Worksheet"
    FacSheetForm.SheetInfoName.Caption = "The Notification Spreadsheet"
    FacSheetForm.DialogueBox.Text = "This spreadsheet can be used to define when specific users will recieve earthquake notifications. " & _
        "Each group defined below will recieve earthquake notificaitons only when specific parameters are met. " & _
        "In order to add a user to a group, you must add the group name to the ""Notification Group"" section " & _
        "in the ""User XML"" spreadsheet." & vbNewLine & vbNewLine & _
        "A group will only recieve a notification when the parameters defined in this spreadsheet are met. " & _
        "One group can recieve notifications in multiple situations by creating multiple group rows with the " & _
        "same group name. For instance: " & vbNewLine & vbNewLine & _
        "CAL_BRIDGES" & vbTab & "BRIDGE" & vbTab & "..." & vbTab & "NEW_EVENT" & vbTab & "..." & vbNewLine & _
        "CAL_BRIDGES" & vbTab & "BRIDGE" & vbTab & "..." & vbTab & "DAMAGE" & vbTab & "..." & vbNewLine & vbNewLine & _
        "defines a single group (CAL_BRIDGES) that will be notified when a new event occurs or when a specified damage " & _
        "level occurs. If a group has multiple defining rows, they should all be input one under the other" & _
        ", rather than scattered around this document. The top row for any group will define aspects of that " & _
        "group that should be the same for all the group rows (i.e. Facility Type, Monitoring Region). You will find that you are unable to enter information into grey cells. This is to protect you from entering invalid information. " & _
        "By collecting your group rows together and inputting information from left to right, we will be able to " & _
        "check the validity of the information you're entering and ensure you aren't wasting time filling out " & _
        "non-required fields." & vbNewLine & vbNewLine & _
        "When you are finished filling out your notification information, you can export an XML document containing " & _
        "all the information by clicking ""Export XML"" at the top of the page. This file can then be dragged " & _
        "and dropped into ShakeCast. Your notification group information should be uploaded to ShakeCast after your " & _
        "facility information, but before your user information."




'If Not FacSheetForm.Visible Then
'    FacSheetForm.Show
'End If

End Sub

Sub GroupAdvInfo()

    FacSheetForm.AdvUser.Caption = "Access General User Worksheet"
    FacSheetForm.SheetInfoName.Caption = "The Advanced Notification Spreadsheet"
    FacSheetForm.DialogueBox.Text = "This spreadsheet allows you to enter more information about a notification group. You can now " & _
        "define the number of messages the ShakeCast report will be divided into (aggregate) and which of those " & _
        "will be grouped together (aggregate group). " & _
        "You also have the ability to select your notification template, the product type that will be delivered, " & _
        "and the shaking metric associated with each group."

'If Not FacSheetForm.Visible Then
'    FacSheetForm.Show
'End If

End Sub


