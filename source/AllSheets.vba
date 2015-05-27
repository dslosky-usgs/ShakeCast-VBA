Attribute VB_Name = "AllSheets"
Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (dir(FileToTest) <> "")
End Function
Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then 'See above
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If
End Sub

Function removeWhite(ByVal inputStr As String)

    ' check that all whitespace is gone
    While Left(inputStr, 1) = " " Or Left(inputStr, 1) = vbTab Or Left(inputStr, 1) = vbNewLine
        inputStr = Mid(inputStr, 2, Len(inputStr) - 1)
    Wend
    
    While Right(inputStr, 1) = " " Or Right(inputStr, 1) = vbTab Or Right(inputStr, 1) = vbNewLine
        inputStr = Mid(inputStr, 1, Len(inputStr) - 1)
    Wend
    
    removeWhite = inputStr

End Function

Function rangeToArray(ByVal rangeAddress As String, _
                        ByVal sheetName As String)
                        
    Set mySheet = Worksheets(sheetName)
    
    Set myRange = mySheet.Range(rangeAddress)
    
    Dim finalArray() As String
    Dim colCount As Integer
    
    colCount = 0
    
    If myRange.Rows.count = 1 Then
    
        ReDim finalArray(0 To myRange.Columns.count - 1)
        
        For Each cell In myRange.Cells
        
            If Not IsError(cell) Then
                finalArray(colCount) = cell.value
            End If
            
            colCount = colCount + 1
        Next cell
                            
    Else
    End If
    
    rangeToArray = finalArray
End Function

' Allows us to use a file picker to get files on Mac and Windows
Function openFilePicker(Optional sPath As String) As String
Dim sFile As String
Dim sMacScript As String

    If InStr(Application.OperatingSystem, "Mac") Then
        If sPath = vbNullString Then
            sPath = "the path to documents folder"
        Else
            sPath = " alias """ & sPath & """"
        End If
        sMacScript = "set sFile to (choose file of type ({" & _
            """public.comma-separated-values-text"", ""public.item"",""public.text"", ""public.csv"", ""public.config""," & _
            """org.openxmlformats.spreadsheetml.sheet.macroenabled""}) with prompt " & _
            """Select a file to import"" default location " & sPath & ") as string" _
            & vbLf & _
            "return sFile"
         'Debug.Print sMacScript
         On Error Resume Next
         sFile = MacScript(sMacScript)

    Else
    
        'windows
'        sFile = Application.GetOpenFilename("CSV files,*.csv,Excel 2007 files,*.xlsx", 1, _
'            "Select file to import from", "&Import", False)

        With Application.FileDialog(msoFileDialogFilePicker)
            .Show
            If .SelectedItems.count = 0 Then
                MsgBox "Cancel Selected"
                Exit Function
            End If
            'fStr is the file path and name of the file you selected.
            sFile = .SelectedItems(1)
        End With


    End If

    openFilePicker = sFile
    
    End Function

Sub loadCSV()
    Dim fStr As String

    fStr = openFilePicker

'    With Application.FileDialog(msoFileDialogFilePicker)
'        .Show
'        If .SelectedItems.count = 0 Then
'            MsgBox "Cancel Selected"
'            Exit Sub
'        End If
'        'fStr is the file path and name of the file you selected.
'        fStr = .SelectedItems(1)
'    End With

    With ThisWorkbook.Sheets("Data").QueryTables.Add(Connection:= _
    "TEXT;" & fStr, Destination:=Range("$A$1"))
        .Name = "CAPTURE"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
'        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
'        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
'        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
'        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False

    End With
End Sub


Private Sub importCSV()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    'On Error GoTo ExitHandler
    
    ' clear data worksheet
    Set dataSheet = Worksheets("Data")
    dataSheet.Activate
    dataSheet.Unprotect
    
    clearSheet "Data"
    
    loadCSV
    
    Set dataSheet = Worksheets("Data")
    'Open csvFile For Input As #1
    
    Dim count As Integer            ' Count how many lines we're importing
    Dim lastRow As Integer         ' last line of data in the dataSheet
    Dim lastCol As Integer          ' last column of the imported data
    Dim startRow As Integer       ' where we start importing data
    Dim headCount As Integer    ' Counts the number of headers
    Dim sheetName As String     ' used to make sure we pipe information the right way
    Dim rowNum As Integer        ' keeps track of where we are in the worksheet
    Dim fields As String              ' this is where we will store the fields for each worksheet
    Dim lineArr() As String          ' used to read the csv one line at a time
    Dim replace As String           ' used to replace the comma delimeter
    Dim newStr As String            ' we use this to build a new csvLine without comma delimeters
    Dim sheetHead() As String    ' header for the worksheet
    Dim CSVHead() As String      ' header for the csv
    Dim LatLon As Double           ' store latitude and longitude before sticking them into the workbook
    
    
    ' get info out of a general info line and into workbook
    Dim lineCount As Integer
    Dim metricHead() As String
    Dim attArr() As String
    Dim savedAtts() As String
    Dim attStr As String
    Dim extraHeads As Integer
    

    
    lastRow = dataSheet.Cells(Rows.count, "A").End(xlUp).row
    lastCol = dataSheet.Cells(1, Columns.count).End(xlToLeft).Column
    startRow = 1
    
    ' set up progress bar
    Dim progressCount As Long
    Dim progressWhen As Long
    Dim pcntDone As Double
    
    progressCount = 0
    progressWhen = lastRow * 0.01
    pcntDone = 0
    ProgressForm.ProcessName.Caption = "Importing CSV"
    
    count = 0
    extraHeads = 0
    
    For csvCount = startRow To lastRow
    

        If progressCount > progressWhen Then
            
            pcntDone = ((csvCount) / (lastRow))
            
            ProgressForm.ProgressLabel.Width = pcntDone * ProgressForm.MaxWidth.Caption
            ProgressForm.ProgressFrame.Caption = Round(pcntDone * 100, 0) & "%"
            
            DoEvents
            progressCount = 0
            
        End If

                
        lastCol = dataSheet.Cells(csvCount, Columns.count).End(xlToLeft).Column
        lineArr = rangeToArray(dataSheet.Cells(csvCount, 1).Address & ":" & dataSheet.Cells(csvCount, lastCol).Address, "Data")
        
            '''''''''''''''' Create column association with the first row '''''''''''''


        If count = 0 Then
NewHead:
            headCount = 0
        
            ' save the head csv values to look at later
            CSVHead = lineArr
            
            ' determine which worksheet we're dealing with
            If InArray(lineArr, "EXTERNAL_FACILITY_ID") Then
                sheetName = "Facility XML"
                fields = "EXTERNAL_FACILITY_ID,FACILITY_TYPE,space,COMPONENT_CLASS,COMPONENT,FACILITY_NAME,DESCRIPTION,SHORT_NAME,GEOM_TYPE,LAT,LON,GEOM:DESCRIPTION,HAZUS,space,space,space,space,space,space,space,space,space,space,space"
                
            ElseIf InArray(lineArr, "POLY") Then
                sheetName = "Notification XML"
                fields = "GROUP_NAME,FACILITY_TYPE,POLY,NOTIFICATION_TYPE,DAMAGE_LEVEL,LIMIT_VALUE,EVENT_TYPE,DELIVERY_METHOD,MESSAGE_FORMAT,PRODUCT_TYPE,METRIC,AGGREGATE,AGGREGATE_GROUP"
            
                MsgBox "Try using a configuration file to import group notification information! If you are not trying to import groups, check to make sure your CSV headers are correct."
                GoTo ExitHandler
            ElseIf InArray(lineArr, "USERNAME") Then
                sheetName = "User XML"
                fields = "USERNAME,USER_TYPE,PASSWORD,FULL_NAME,EMAIL_ADDRESS,PHONE_NUMBER,GROUP,DELIVERY:EMAIL_HTML,DELIVERY:EMAIL_TEXT,DELIVERY:EMAIL_PAGER"
                
            Else
                MsgBox "ERROR: MISSING HEADERS"
                GoTo ExitHandler
            End If
            
            ' set up info for worksheet
            Set mySheet = Worksheets(sheetName)
            
            Dim sheetHeads() As String
            sheetHeads = Split(fields, ",")
            
            ' create an array to store the column numbers for each
            Dim colNums() As Variant
            
            ' get the start row for our input
            Dim lastInputRow As Integer
            lastInputRow = mySheet.Cells(Rows.count, "A").End(xlUp).row
            
            If IsEmpty(mySheet.Range("A" & lastInputRow)) Then
                lastInputRow = lastInputRow - 1
            End If
            
            ' make an array of the column numbers associated with the fields in the csv
            headCount = 0
            For Each Head In CSVHead
            
                ReDim Preserve colNums(0 To headCount)
                colNums(headCount) = ArrayIndex(sheetHeads, Head)
                
                headCount = headCount + 1
            Next Head
            
            GoTo NextLine
        End If
        
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' update the row
        rowNum = lastInputRow + count - (extraHeads * 2)
        
        ' check to see if we're dealing with a new header
        If InArray(lineArr, "EXTERNAL_FACILITY_ID") Then
            extraHeads = extraHeads + 1
            GoTo NewHead
        ElseIf InArray(lineArr, "POLY") Then
            extraHeads = extraHeads + 1
            GoTo NewHead
        ElseIf InArray(lineArr, "USERNAME") Then
            extraHeads = extraHeads + 1
            GoTo NewHead
        End If
        

        lineCount = 0
        attStr = ""
        
        For Each arrVal In lineArr
            If colNums(lineCount) <> -99 Then
                mySheet.Cells(rowNum, colNums(lineCount) + 1).value = arrVal
            
            ' get info from metric columns
            ElseIf sheetName = "Facility XML" Then
            
                If InStr(CSVHead(lineCount), "METRIC:") And UBound(Split(CSVHead(lineCount), ":")) + 1 = 3 Then
                
                    mySheet.Cells(rowNum, 13).value = "USER_DEFINED"
                
                    metricHead = Split(CSVHead(lineCount), ":")
                    
                    If metricHead(2) = "GREEN" Then
                    
                        mySheet.Cells(rowNum, 15).value = metricHead(1)
                        mySheet.Cells(rowNum, 16).value = arrVal
                        mySheet.Cells(rowNum, 17).value = 0.64
                        
                    ElseIf metricHead(2) = "YELLOW" Then
                    
                        mySheet.Cells(rowNum, 18).value = metricHead(1)
                        mySheet.Cells(rowNum, 19).value = arrVal
                        mySheet.Cells(rowNum, 20).value = 0.64
                    
                    ElseIf metricHead(2) = "ORANGE" Then
                    
                        mySheet.Cells(rowNum, 21).value = metricHead(1)
                        mySheet.Cells(rowNum, 22).value = arrVal
                        mySheet.Cells(rowNum, 23).value = 0.64
                    
                    ElseIf metricHead(2) = "RED" Then
                    
                        mySheet.Cells(rowNum, 24).value = metricHead(1)
                        mySheet.Cells(rowNum, 25).value = arrVal
                        mySheet.Cells(rowNum, 26).value = 0.64
                    
                    ElseIf metricHead(2) = "GREY" Then
                    
                        mySheet.Cells(rowNum, 27).value = metricHead(1)
                        mySheet.Cells(rowNum, 28).value = arrVal
                        mySheet.Cells(rowNum, 29).value = 0.64
                    
                    End If
                    
                ' get facility attribute information
                ElseIf InStr(CSVHead(lineCount), "ATTR:") And arrVal <> "" Then
                
                    attArr = Split(CSVHead(lineCount), ":")
                
                    ' stick attributes in facility XML cell
                    attStr = mySheet.Cells(rowNum, 30).value
                    If attStr = "" Then
                        mySheet.Cells(rowNum, 30).value = attArr(1) & ":" & arrVal
                    Else
                        mySheet.Cells(rowNum, 30).value = attStr & "%" & attArr(1) & ":" & arrVal
                    End If
                
                    ' add attribute to the attribute list
                    
                    Set ShakeSheet = Worksheets("ShakeCast Ref Lookup Values")
                    
                    If Not IsEmpty(ShakeSheet.Range("P2")) Then
                        savedAtts = Split(ShakeSheet.Range("P2").value, "%")
                        
                        If Not InArray(savedAtts, attArr(1)) Then
                            ShakeSheet.Range("P2").value = ShakeSheet.Range("P2").value & "%" & attArr(1)
                        End If
                    Else
                        ShakeSheet.Range("P2").value = attArr(1)
                    End If
                    
                ' Check for max and min latitude and longitudes
                ElseIf InStr(CSVHead(lineCount), "LAT") And arrVal <> "" Then
                    
                    LatLon = arrVal
                    
                    If InStr(CSVHead(lineCount), "MAX") Then

                        If IsEmpty(mySheet.Range("J" & rowNum)) Then
                            mySheet.Range("J" & rowNum).value = LatLon
                        ElseIf LatLon <> mySheet.Range("J" & rowNum).value Then
                            mySheet.Range("J" & rowNum).value = LatLon & ";" & mySheet.Range("J" & rowNum).value
                        End If

                    ElseIf InStr(CSVHead(lineCount), "MIN") Then

                        If IsEmpty(mySheet.Range("J" & rowNum)) Then
                            mySheet.Range("J" & rowNum).value = LatLon
                        ElseIf LatLon <> mySheet.Range("J" & rowNum).value Then
                            mySheet.Range("J" & rowNum).value = mySheet.Range("J" & rowNum).value & ";" & LatLon
                        End If


                    End If

                ElseIf InStr(CSVHead(lineCount), "LON") And arrVal <> "" Then
                    LatLon = arrVal
                    
                    If InStr(CSVHead(lineCount), "MAX") Then

                        If IsEmpty(mySheet.Range("K" & rowNum)) Then
                            mySheet.Range("K" & rowNum).value = LatLon
                        ElseIf LatLon <> mySheet.Range("K" & rowNum).value Then
                            mySheet.Range("K" & rowNum).value = LatLon & ";" & mySheet.Range("K" & rowNum).value
                        End If

                    ElseIf InStr(CSVHead(lineCount), "MIN") Then

                        If IsEmpty(mySheet.Range("K" & rowNum)) Then
                            mySheet.Range("K" & rowNum).value = LatLon
                        ElseIf LatLon <> mySheet.Range("K" & rowNum).value Then
                            mySheet.Range("K" & rowNum).value = mySheet.Range("K" & rowNum).value & ";" & LatLon
                        End If

                    End If
                End If
                
            ElseIf sheetName = "User XML" Then
            
                If InStr(CSVHead(lineCount), "GROUP:") And arrVal <> "" Then
                    mySheet.Cells(rowNum, 7).value = arrVal
                End If
            
            End If
            
            lineCount = lineCount + 1
            
        Next arrVal
        
        
        
NextLine:
        progressCount = progressCount + 1
        count = count + 1
    'Loop
    Next csvCount
        'End If
    
    ' update the worksheet
    If sheetName = "Facility XML" Then
        Worksheets("ShakeCast Ref Lookup Values").Range("Q2").value = "FacUpdate"
        Unload ProgressForm
        ProgressForm.Show
    ElseIf sheetName = "User XML" Then
        Worksheets("ShakeCast Ref Lookup Values").Range("Q2").value = "UserUpdate"
        Unload ProgressForm
        ProgressForm.Show
    End If
    
ExitHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Close #1
End Sub

Sub importConf()

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    On Error GoTo ExitHandler
    
    Close #1

    Dim fileStr As String
    fileStr = openFilePicker
    
    If fileStr = "" Then
        MsgBox "No file chosen!"
        GoTo ExitHandler
    End If
    
    Open fileStr For Input As #1
    
    ' keep track of where we are in the notification worksheet
    Dim startRow As Integer
    Dim rowNum As Integer
    
    Set mySheet = Worksheets("Notification XML")
    startRow = mySheet.Cells(Rows.count, "A").End(xlUp).row
    
    ' create some variables we'll need
    Dim sameLine As Boolean     ' determine if the current line is the same as the last line
    Dim tagName As String         ' the kind of data we're looking at
    Dim val As String
    Dim confSplit() As String
    Dim groupName As String              ' the name of the group we're looking at
    Dim notiCount As Integer      ' the number of notification rows in the group
    
    Dim fields As String
    
    fields = "GROUP_NAME,FACILITY_TYPE,POLY,NOTIFICATION_TYPE,DAMAGE_LEVEL,LIMIT_VALUE,EVENT_TYPE,DELIVERY_METHOD,MESSAGE_FORMAT,PRODUCT_TYPE,METRIC,AGGREGATE,AGGREGATE_GROUP"
    
    Dim sheetHeads() As String
    sheetHeads = Split(fields, ",")
    
    If IsEmpty(mySheet.Range("A" & startRow)) Then
        rowNum = startRow
    Else
        rowNum = startRow + 1
    End If
        
    
    While Not EOF(1)

        Line Input #1, confLine
        
        confLine = Trim(confLine)
        confLine = removeWhite(confLine)
        
        If Left(confLine, 1) = "#" Or confLine = "" Then GoTo NextInfo
        
        ' get rid of comments if there are any on the same line
        If InStr(confLine, "#") Then
            confLine = Split(confLine, "#")(0)
            confLine = removeWhite(confLine)
        End If
        
        ' if this is the same line: don't look for anything, just put the value check the existing tag
        If Not sameLine Then
            ' check for notification, or new group
            If InStr(confLine, "<") <> 0 And InStr(confLine, ">") <> 0 And InStr(confLine, "NOTIFICATION") = 0 And InStr(confLine, "Notification") = 0 And InStr(confLine, "notification") = 0 Then
                ' new group name
                groupName = Mid(confLine, 2, Len(confLine) - 2)
                
                ' rowNum = rowNum + 1
                
            ElseIf InStr(confLine, "<") <> 0 And InStr(confLine, ">") <> 0 And InStr(confLine, "/NOTIFICATION") = 0 Then
                ' new row
                'rowNum = rowNum + 1
                
                ' new notification group row
                mySheet.Range("A" & rowNum).value = groupName
                
            ElseIf InStr(confLine, "<") <> 0 And InStr(confLine, ">") <> 0 And InStr(confLine, "/NOTIFICATION") <> 0 Then
            
                ' close notification group
                 rowNum = rowNum + 1
                
            Else
            
                ' get header
                If InStr(confLine, vbTab) Then
                    confSplit = Split(confLine, vbTab)
                ElseIf InStr(confLine, " ") Then
                    confSplit = Split(confLine, " ")
                End If
            
                tagName = confSplit(0)
                
                ' create one value from the rest of the array in case there are other spaces or tabs
                confSplit(0) = ""
                
                val = Join(confSplit, "")
                
                tagName = removeWhite(tagName)
                val = removeWhite(val)
                
                If Right(val, 1) = "\" Then
                    val = Mid(val, 1, Len(val) - 1)
                    
                    val = removeWhite(val)
                    
                    sameLine = True
                End If
                
                If InArray(sheetHeads, tagName) Then
                
                    If val = "EMAIL_HTML" Then
                        val = "Rich Content"
                    ElseIf val = "EMAIL_TEXT" Then
                        val = "Plain Text"
                    End If
                    
                    mySheet.Cells(rowNum, ArrayIndex(sheetHeads, tagName) + 1).value = val
                End If
                
        ' if this is NOT the same line as the last line. Meaning the last line did not end with \
            ' find new group row: check if < & > are in line without "NOTIFICATION"
        
            ' find notification info: if notification = true (Means that we've seen <NOTIFICATION> and not </NOTIFICATION>)
        
            ' get other info if not notification: if notification = false
            
            End If
        Else
            ' this is the same line!! But the next one might not be!
            sameLine = False
            
            val = confLine
            val = removeWhite(val)
                
            If Right(val, 1) = "\" Then
                val = Mid(val, 1, Len(val) - 1)
                
                val = removeWhite(val)
                
                sameLine = True
            End If
            
            If InArray(sheetHeads, tagName) Then
                mySheet.Cells(rowNum, ArrayIndex(sheetHeads, tagName) + 1).value = mySheet.Cells(rowNum, ArrayIndex(sheetHeads, tagName) + 1).value & ";" & val
            End If
            
        End If
        
NextInfo:
        ' check if the next line will be the same as this one
        
    Wend
    
ExitHandler:
    Close #1
    
    Application.Run "CheckGroups"
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    

End Sub

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
            
    Set ColorCells = mySheet.Range("I" & rowNum, "L" & rowNum)
    With ColorCells.Interior
        .Color = RGB(235, 241, 222)
    End With
        
    Set ColorCells = mySheet.Range("M" & rowNum, "N" & rowNum)
    With ColorCells.Interior
        .Color = RGB(191, 191, 191)
    End With
    
    Set ColorCells = mySheet.Range("O" & rowNum, "Q" & rowNum)
    With ColorCells.Interior
        .Color = RGB(0, 176, 80)
    End With
    
    Set ColorCells = mySheet.Range("R" & rowNum, "T" & rowNum)
    With ColorCells.Interior
        .Color = RGB(255, 255, 0)
    End With
    
    Set ColorCells = mySheet.Range("U" & rowNum, "W" & rowNum)
    With ColorCells.Interior
        .Color = RGB(255, 192, 0)
    End With
    
    Set ColorCells = mySheet.Range("X" & rowNum, "Z" & rowNum)
    With ColorCells.Interior
        .Color = RGB(255, 0, 0)
    End With
    
    Set ColorCells = mySheet.Range("AA" & rowNum, "AC" & rowNum)
    With ColorCells.Interior
        .Color = RGB(242, 242, 242)
    End With
    
    Set ColorCells = mySheet.Range("AD" & rowNum, "AD" & rowNum)
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
                        mySheet.Range("A" & cell.row).value = mySheet.Range("A" & cell.row - 1).value) Or _
                        (InStr(cell.Address, "E") And mySheet.Range("D" & cell.row).value = "NEW_EVENT") Or _
                        (InStr(cell.Address, "F") And mySheet.Range("D" & cell.row).value = "DAMAGE") Or _
                        InStr(cell.Address, "I") Then GoTo NextCell
                        
            
                If rowCheck < 1 Then
                    errStr = errStr & "Row " & row & ": " & tabSpace & cell.Address(False, False)
                Else
                    errStr = errStr & " :: " & cell.Address(False, False)
            
                End If
            
                rowCheck = rowCheck + 1
                
            ElseIf (IsEmpty(cell) Or IsError(cell)) And whichSheet = "User" Then
            
                If (InStr(cell.Address, "F") And mySheet.Range("B" & cell.row).value = "USER") Or _
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

Function ArrayIndex(ByVal searchArray As Variant, ByVal value As String, Optional multi As Boolean)
    
    ' the function returns false if the index was not found
    ArrayIndex = -99
    
    ' test for the start position of the array
    Dim testIndex As Integer
    Dim testStr As String
    Dim count As Integer
    
    testIndex = 0
    
    On Error Resume Next
    testStr = searchArray(testIndex)
    
    If Err Then
        count = 1                  ' if there's an error, we know the array starts at 1, not 0
    Else
        count = 0
    End If
    
    Dim indexArray() As Variant    ' contains the index for each instance of the search value
    Dim indexCount As Integer      ' counts the number of instances
    
    indexCount = 0
    
    For Each arrVal In searchArray
        
        If arrVal = value Then
            If multi = True Then
            
                ReDim Preserve indexArray(0 To indexCount)
                
                indexArray(indexCount) = count
                indexCount = indexCount + 1
                
                ArrayIndex = indexArray
            Else
                ArrayIndex = count
                Exit Function
            End If
        End If
        
        count = count + 1
    Next arrVal
End Function

Private Sub checkXMLchars(ByVal target As Range)

     For i = 1 To Len(target.value)
        If InStr("&", Mid(target.value, i, 1)) Then
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

End Sub

Private Sub clearSheet(Optional sheetName As String)

Dim startRow As Integer
Dim endRow As Long

If sheetName = "Data" Then
    startRow = 1
Else
    startRow = 5
End If

endRow = 5
' A AF N K are the letters where each row is defined and evaluated for all worksheets
For Each letter In Split("A,AE,N,K", ",")
    
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

If Err And sheetName <> "Data" Then
MsgBox "We couldn't find any rows!"
End If

End Sub

Private Sub progressBar()

    Dim process As String
    process = Worksheets("ShakeCast Ref Lookup Values").Range(Q2).value

    If process = "FacilityXML" Then
        Unload OptionsForm
        
        ProgressForm.Show
        
    End If


End Sub


