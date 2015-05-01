Attribute VB_Name = "AllSheets"
'' masterXMLexport
'' Daniel Slosky
'' Last Updated: 3/4/2014
''
'' Exports all of the facility, group, and user information as one file
''
''


Sub masterXMLexport()

Application.EnableEvents = False
Application.ScreenUpdating = False
' Close our write output, just in case it was left open
Close #2

' remember where we started
Set startActiveSheet = ActiveSheet
Set startActiveCell = ActiveCell

Dim facXMLInfo() As Variant
Dim groupXMLInfo() As Variant
Dim userXMLInfo() As Variant

' Make the XML tables for each piece of the spreadsheet
facXMLInfo = Application.Run("FacXMLTable")
groupXMLInfo = Application.Run("GroupXMLTable")
userXMLInfo = Application.Run("UserXMLTable")

' Calculate total accepted and total declided facilities
Dim infoAcc As Long
Dim infoDec As Integer

infoAcc = facXMLInfo(0) + groupXMLInfo(0) + userXMLInfo(0)

'                          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                 OPEN XML FILE FOR WRITING, AND DETERMINE THE START AND END CELLS TO BE EXAMINED

' open file location
Dim dir As String
dir = Application.ActiveWorkbook.Path
Dim docMax As Integer
Dim docNum As Double

docMax = 15000
docNum = infoAcc / docMax

Do While WorksheetFunction.Ceiling(docNum, 1) = docNum
    docMax = docMax - 1
    docNum = infoAcc / docMax
Loop

If docNum < 1 Then
    docNum = 1
    docStr = "MasterXML.xml"
Else
    docNum = Application.WorksheetFunction.Ceiling(docNum, 1)
    docStr = "MasterXML1.xml"
    For i = 2 To docNum
        docStr = docStr & "," & "MasterXML" & i & ".xml"
    Next i
End If



Dim getOS As String
getOS = Application.OperatingSystem

ExportXML.FileDest.Text = dir
ExportXML.FileName = docStr
ExportXML.Show


Dim docArr() As String
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


' Begin the document
printStr = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbNewLine & _
            "<Inventory>"

Print #2, printStr

'                          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim testOverFlow As Long

FacilityXML "Master", docCount, overFlowCount, docMax, docStr

' determine document and overflow count after Facility XML runs
testOverFlow = facXMLInfo(0)
Do While testOverFlow > docMax
   testOverFlow = testOverFlow - docMax
   docCount = docCount + 1
Loop

overFlowCount = testOverFlow

GroupXML "Master", docCount, overFlowCount, docMax, docStr

' determine document and overflow count after Group XML runs
testOverFlow = groupXMLInfo(0) + overFlowCount
Do While testOverFlow > docMax
   testOverFlow = testOverFlow - docMax
   docCount = docCount + 1
Loop

overFlowCount = testOverFlow

UserXML "Master", docCount, overFlowCount, docMax, docStr


printStr = "</Inventory>"

On Error Resume Next
Print #2, printStr

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''' Make Error Statement ''''''''''''''''''''''''''''''''''''''''
Dim errStr As String
Dim facErrStr As String
Dim groupErrStr As String
Dim userErrStr As String

facErrStr = MakeErrStr(facXMLInfo, "Facility")
groupErrStr = MakeErrStr(groupXMLInfo, "Group")
userErrStr = MakeErrStr(userXMLInfo, "User")

errStr = "Facility Spreadsheet: " & vbNewLine & vbNewLine & _
            facErrStr & vbNewLine & vbNewLine & _
            "Group Spreadsheet: " & vbNewLine & vbNewLine & _
            groupErrStr & vbNewLine & vbNewLine & _
            "User Spreadsheet: " & vbNewLine & vbNewLine & _
            userErrStr & vbNewLine & vbNewLine


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
    "Facilities Accepted: " & facXMLInfo(0) & vbNewLine & _
    "Groups Accepted: " & groupXMLInfo(0) & vbNewLine & _
    "Users Accepted: " & userXMLInfo(0) & vbNewLine & _
    vbNewLine & vbNewLine & _
    "Facilities Declined: " & facXMLInfo(1) & vbNewLine & _
    "Groups Declined: " & groupXMLInfo(1) & vbNewLine & _
    "Users Declined: " & userXMLInfo(1) & _
    vbNewLine & vbNewLine & _
    "--------------------------------------------------------------------------" & _
    vbNewLine & vbNewLine
      
        
DiaStr = DiaStr & _
    "--------------------------------------------------------------------------" & _
    vbNewLine & vbNewLine & _
    "Errors: " & (facXMLInfo(1) + groupXMLInfo(1) + userXMLInfo(1))

If errStr <> "" Then

    DiaStr = DiaStr & _
        vbNewLine & vbNewLine & _
        "Any facilities, groups, or users that you attempted to include in your XML document that were rejected are " & _
        "highlighted in blue." & _
        vbNewLine & vbNewLine & _
        "The following cells contain invalid entries and are stopping some entries from being included in the XML document: " & _
        vbNewLine & vbNewLine & _
        errStr

Else
    DiaStr = DiaStr & vbNewLine & vbNewLine & _
        "All facility information has been converted to XML."

End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''' End Error Statement ''''''''''''''''''''''''''''''''''''''''

startActiveSheet.Activate
ActiveCell.Activate

Unload ProgressForm

DialogueForm.DialogueBox.Text = DiaStr
DialogueForm.Show



XMLFinish:
Close #2

startActiveSheet.Activate

Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub



'' ChangeColors
'' Daniel Slosky
'' Last Updated: 2/11/2015
''
'' This software takes an input string that is either "Good" or "Bad". This describes how the color of the cells in the current row will be changed. If the row
'' is "Good", its color scheme are set to the normal spreadsheet colords. If not, they are set to a color that stands out. This will clearly indicate which rows
'' the user filled out completely and correctly
''
'' It also has an input RowCells, which is the range that needs to have its color changed
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Sub ChangeColors(GoodBad As String, _
                    RowCells As Range, _
                    whichSheet As String)
'
' ChangeColors Macro
'

'

Dim mySheet As Object

If whichSheet = "Facility" Then

    Set mySheet = Worksheets("Facility XML")
    Worksheets("Facility XML").Activate

ElseIf whichSheet = "Group" Then

    Set mySheet = Worksheets("Notification XML")
    Worksheets("Notification XML").Activate
    
ElseIf whichSheet = "User" Then

    Set mySheet = Worksheets("User XML")
    Worksheets("User XML").Activate
    
End If


' If the row is no good for our XML, we change it to a tourquiosey color to make it stand out
If GoodBad = "Bad" Then

    With RowCells.Interior
        .Color = RGB(146, 205, 220)
    End With
    
ElseIf GoodBad = "Advanced" Then

    RowCells.Select
    With Selection.Interior
        .Color = RGB(192, 80, 77)
    End With
    
ElseIf GoodBad = "Good" And whichSheet = "Facility" Then


' If the row is good for the XML, we have to break it up into sections to change the color of each column
' so that they match the rest of the color scheme

    Dim rowNum As Long
    rowNum = RowCells.row
    
    
    Dim ColorCells As Range
    
    Set ColorCells = mySheet.Range("A" & rowNum, "H" & rowNum)
    With ColorCells.Interior
         .Color = RGB(218, 238, 243)
    End With
            
    Set ColorCells = mySheet.Range("I" & rowNum, "M" & rowNum)
    With ColorCells.Interior
        .Color = RGB(235, 241, 222)
    End With
        
    Set ColorCells = mySheet.Range("N" & rowNum, "O" & rowNum)
    With ColorCells.Interior
        .Color = RGB(191, 191, 191)
    End With
    
    Set ColorCells = mySheet.Range("P" & rowNum, "R" & rowNum)
    With ColorCells.Interior
        .Color = RGB(0, 176, 80)
    End With
    
    Set ColorCells = mySheet.Range("S" & rowNum, "U" & rowNum)
    With ColorCells.Interior
        .Color = RGB(255, 255, 0)
    End With
    
    Set ColorCells = mySheet.Range("V" & rowNum, "X" & rowNum)
    With ColorCells.Interior
        .Color = RGB(255, 192, 0)
    End With
    
    Set ColorCells = mySheet.Range("Y" & rowNum, "AA" & rowNum)
    With ColorCells.Interior
        .Color = RGB(255, 0, 0)
    End With
    
    Set ColorCells = mySheet.Range("AB" & rowNum, "AD" & rowNum)
    With ColorCells.Interior
        .Color = RGB(242, 242, 242)
    End With
    
    Set ColorCells = mySheet.Range("AE" & rowNum, "AE" & rowNum)
    With ColorCells.Interior
        .Color = RGB(221, 217, 196)
    End With
    
        
ElseIf GoodBad = "Good" And whichSheet = "User" Then

    If RowCells.row Mod 2 = 0 Then
        With RowCells.Interior
            .ColorIndex = 36
        End With
        
    Else
        With RowCells.Interior
        .Color = RGB(230, 166, 121)
        End With
    End If
End If

End Sub


'' ExportDialogue
'' Daniel Slosky
'' Last Updated: 2/19/2015
''
'' This little guy will take all of the text in the dialogue spreadsheet and enter it into a
'' text document
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExportDialogue()

Close #3

On Error GoTo DiaFinish

' Open the file name we just created for writing
Dim Diapath As String
Diapath = DiaExportForm.TextBox1.Text


Open Diapath For Output As #3


'Set DiaSheet = Worksheets("Dialogue")       ' Open up the Dialogue worksheet
'Set DialogueBox = DiaSheet.Label21          ' Grab the label object that we use to display info

' Create a string that will be used to export the caption from the label
Dim exportStr As String
exportStr = DialogueForm.DialogueBox.Text

' Print to the file we created for this dialogue
Print #3, exportStr


' Create a pop-up to let the user know what happened
MsgBox "This Dialogue has been saved as: " & _
    vbNewLine & vbNewLine & _
    Diapath



DiaFinish:
If Err Then
    MsgBox "This dialogue could not be saved" & _
    vbNewLine & Err.Description
Else
    DiaExportForm.Hide
End If




' Close the file
Close #3

End Sub


Sub genericInfo()

MsgBox "Information on how to fill out this column: "

End Sub

'' MakeErrStr
'' Daniel Slosky
'' Last Updated: 2/20/2015
''
''
''
''
Function MakeErrStr(ErrAdd As Variant, _
                        whichSheet)

Dim rangeStart As String
Dim rangeEnd As String
Dim required As String

If whichSheet = "Facility" Then
    Set mySheet = Worksheets("Facility XML")
    rangeStart = "A"
    rangeEnd = "AD"
    
    required = "A,B,D,E,F,G,H,I,J,K,L,N,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AD"

ElseIf whichSheet = "Group" Then
    Set mySheet = Worksheets("Notification XML")
    rangeStart = "A"
    rangeEnd = "H"
    
ElseIf whichSheet = "User" Then
    Set mySheet = Worksheets("User XML")
    rangeStart = "A"
    rangeEnd = "K"
    
End If

Dim errStr As String
Dim tabSpace As String
Dim rowCheck As Integer


errStr = ""
tabSpace = "      "
rowCheck = 0


If ErrAdd(1) > 0 Then

    For Each row In ErrAdd(2)

        Set RowRange = mySheet.Range(rangeStart & row, rangeEnd & row)
    
        rowCheck = 0
    
        If row > 9 And row < 100 Then
            tabSpace = "    "
                
        ElseIf row > 99 And row < 1000 Then
            tabSpace = "   "
        
        ElseIf row > 999 And row < 10000 Then
            tabSpace = "  "
        
        ElseIf row > 9999 And row < 100000 Then
    
            tabSpace = " "
    
        End If
    
        For Each cell In RowRange.Cells
    
            If (IsEmpty(cell) Or IsError(cell)) And whichSheet = "Facility" Then
        
                If rowCheck < 1 Then
                    errStr = errStr & "Row " & row & ": " & tabSpace & cell.Address(False, False)
                Else
                    errStr = errStr & " :: " & cell.Address(False, False)
            
                End If
            
                rowCheck = rowCheck + 1
            
            ElseIf (IsEmpty(cell) Or IsError(cell)) And whichSheet = "Group" Then
            
                ' If the missing cell is one that should be covered by a different group row
                ' we want to skip it!
                If ((InStr(cell.Address, "B") Or InStr(cell.Address, "C") Or _
                        InStr(cell.Address, "H")) And _
                        mySheet.Range("A" & cell.row).Value = mySheet.Range("A" & cell.row - 1).Value) Or _
                        (InStr(cell.Address, "E") And mySheet.Range("D" & cell.row).Value = "NEW_EVENT") Or _
                        (InStr(cell.Address, "F") And mySheet.Range("D" & cell.row).Value = "DAMAGE") Or _
                        InStr(cell.Address, "I") Then GoTo NextCell
                        
            
                If rowCheck < 1 Then
                    errStr = errStr & "Row " & row & ": " & tabSpace & cell.Address(False, False)
                Else
                    errStr = errStr & " :: " & cell.Address(False, False)
            
                End If
            
                rowCheck = rowCheck + 1
                
            ElseIf (IsEmpty(cell) Or IsError(cell)) And whichSheet = "User" Then
            
                If (InStr(cell.Address, "F") And mySheet.Range("B" & cell.row).Value = "USER") Or _
                        InStr(cell.Address, "J") Then GoTo NextCell
            
                If rowCheck < 1 Then
                    errStr = errStr & "Row " & row & ": " & tabSpace & cell.Address(False, False)
                Else
                    errStr = errStr & " :: " & cell.Address(False, False)
            
                End If
            
                rowCheck = rowCheck + 1
            
            End If
    
    
            If rowCheck = 10 Or rowCheck = 20 Then
        
                If row < 10 Then
                    errStr = errStr & vbNewLine & "                      "
                ElseIf row > 9 And row < 100 Then
                    errStr = errStr & vbNewLine & "                        "
                ElseIf row > 99 And row < 1000 Then
                    errStr = errStr & vbNewLine & "                           "
                ElseIf row > 999 And row < 10000 Then
                    errStr = errStr & vbNewLine & "                              "
                ElseIf row > 9999 And row < 100000 Then
                    errStr = errStr & vbNewLine & "                                 "
                End If
            End If
    
NextCell:
        Next cell

        errStr = errStr & vbNewLine & vbNewLine
    
    Next row
    
End If

If errStr <> "" Then
    MakeErrStr = errStr
Else
    MakeErrStr = ""
End If
End Function


Sub FacilityXMLButton()

    FacilityXML "NotMaster"

End Sub


Function InArray(ByVal arr As Variant, ByVal stringToBeFound As String) As Boolean

    
    Dim arrStr As String
    arrStr = Join(arr, ",")
    
    On Error GoTo TheEnd
        
    If InStr(arrStr, "," & stringToBeFound & ",") Then
        InArray = True
    ElseIf arr(UBound(arr)) = stringToBeFound Then
        InArray = True
    ElseIf arr(0) = stringToBeFound Then
        InArray = True
    Else
        InArray = False
    End If
    
TheEnd:
If Err Then
    InArray = False
End If
End Function

Private Sub checkXMLchars(ByVal target As Range)

     For i = 1 To Len(target.Value)
        If InStr("&", Mid(target.Value, i, 1)) Then
            MsgBox "WARNING: We have detected an & in this cell. Unless you are using the ampersand as an escape character " & _
                "this text will not be presented correctly in the ShakeCast application."
                
            
            Exit Sub
        End If
     Next i
End Sub

Private Sub protectWorkbook()

    For Each Sheet In Application.ThisWorkbook.Sheets

        Sheet.Protect AllowFormattingCells:=True, AllowDeletingRows:=True, AllowInsertingRows:=True, UserInterfaceOnly:=True

    Next Sheet

    Application.EnableEvents = True

End Sub

Private Sub clearSheet()

Dim startRow As Integer
Dim endRow As Long

startRow = 5
endRow = 5
' A AF N K are the letters where each row is defined and evaluated for all worksheets
For Each letter In Split("A,AF,N,K", ",")
    
    If ActiveSheet.Cells(Rows.count, letter).End(xlUp).row > endRow Then
        endRow = ActiveSheet.Cells(Rows.count, letter).End(xlUp).row + 1
    End If
    
Next letter

On Error Resume Next

ActiveSheet.Rows("4:4").EntireRow.Clear
ActiveSheet.Rows("4:4").EntireRow.Locked = False
    
ActiveSheet.Rows(startRow & ":" & endRow).EntireRow.Delete
    
    
If ActiveSheet.Name = "Facility XML" Then
    CheckFacilities Worksheets("Facility XML").Range("A4")
ElseIf ActiveSheet.Name = "User XML" Then
    CheckUsers 4
End If

ActiveSheet.Range("A4").Activate

If Err Then
MsgBox "We couldn't find any rows!"
End If

End Sub

Private Sub progressBar()

    Dim process As String
    process = Worksheets("ShakeCast Ref Lookup Values").Range(Q2).Value

    If process = "FacilityXML" Then
        Unload OptionsForm
        
        ProgressForm.Show
        
    End If


End Sub
