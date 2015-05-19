Attribute VB_Name = "Tutorial"

Sub tutWindow()

For Each Control In TutForm.SecFrame.Controls
    If TypeOf Control Is MSForms.Label Then
        Control.BackColor = RGB(255, 255, 255)
        End If
Next Control

Dim tutNum As Integer
Dim tutDec As Integer

tutNum = TutForm.SecNum.Caption
tutDec = TutForm.SecDec.Caption

If tutNum = "0" And tutDec = "0" Then

    TutForm.DialogueBox = "Welcome to the ShakeCast Workbook tutorial!" & vbNewLine & vbNewLine & _
                    "This tutorial will walk you through how to fill out all the spreadsheets in this workbook, " & _
                    "beginning with the facility sheet. You can use the buttons at the bottom of this form to " & _
                    "navigate your way through the tutorial. You can also click on a specific section to the right, " & _
                    "and the tutorial will jump to that section."
                    
                    
ElseIf tutNum = "1" And tutDec = "0" Then

    TutForm.SecFrame.Label_0.BackColor = RGB(218, 238, 243)

    
ElseIf tutNum = "1" And tutDec = "1" Then
    
    TutForm.SecFrame.Label_1.BackColor = RGB(218, 238, 243)
    
ElseIf tutNum = "1" And tutDec = "2" Then
    
    TutForm.SecFrame.Label_2.BackColor = RGB(218, 238, 243)
    
    
ElseIf tutNum = "1" And tutDec = "3" Then
    
    TutForm.SecFrame.Label_3.BackColor = RGB(218, 238, 243)
    
    
ElseIf tutNum = "2" And tutDec = "0" Then
    
    TutForm.SecFrame.Label_4.BackColor = RGB(218, 238, 243)
    
    TutForm.DialogueBox.Text = "Welcome to the Notification Worksheet! Notification groups are used to determine who gets notifications, when. " & _
            "Each row in this spreadsheet represents a situation in which users will get a notification. For the purpose of this tutorial, lets create " & _
            "a notification group that will recieve alerts when there is an earthquake in California or when there is any ground shaking at the Golden Gate Bridge, which we just defined in the last part of the tutorial!" & vbNewLine & vbNewLine & _
            "We will need to make some more space to do this tutorial, so when you hit continue we will move the information you've input in rows 4-13 to another worksheet."
          
    
    
ElseIf tutNum = "2" And tutDec = "1" Then
    
    TutForm.SecFrame.Label_5.BackColor = RGB(218, 238, 243)
    
ElseIf tutNum = "2" And tutDec = "2" Then
    
    TutForm.SecFrame.Label_6.BackColor = RGB(218, 238, 243)
    
    
ElseIf tutNum = "3" And tutDec = "0" Then
    
    TutForm.SecFrame.Label_7.BackColor = RGB(218, 238, 243)
    
    TutForm.DialogueBox.Text = "Welcome to the User Worksheet! You can use this worksheet to define ShakeCast users and link them to the notification groups " & _
            "defined in the Notification Worksheet. We'll go through two quick examples for this sheet; we'll create a general user first, then an admin user. In order to do this, we will need to make room in this worksheet again! " & vbNewLine & vbNewLine & _
            "Hit Continue to start!"
    

ElseIf tutNum = "3" And tutDec = "1" Then
    
    TutForm.SecFrame.Label_8.BackColor = RGB(218, 238, 243)
    
    
ElseIf tutNum = "3" And tutDec = "2" Then
    
    TutForm.SecFrame.Label_9.BackColor = RGB(218, 238, 243)
    
Else
    TutForm.DialogueBox = "Finished!!"

End If


End Sub

Sub BuildTut()

Dim Secs As String       ' Secs will hold the names of all the sections with a comma delimeter
Dim SecsArr() As String

Secs = "Facility Worksheet,BRIDGE,Polyline,Polygon,Notification Worksheet,CAL_BRIDGES,Scenario Group,User Worksheet,USER,ADMIN"
SecsArr = Split(Secs, ",")

Set myFrame = TutForm.SecFrame

For i = 0 To UBound(SecsArr)

    Set lab = TutForm.SecFrame.Controls("Label_" & i)
    lab.Caption = SecsArr(i)
    lab.Top = 5 + (i * 25)
    lab.Font.Size = 12
    lab.Height = 23
    lab.Width = 160
    
    If SecsArr(i) <> "Facility Worksheet" And _
         SecsArr(i) <> "Notification Worksheet" And _
         SecsArr(i) <> "User Worksheet" Then
        
        lab.Left = 15
    End If

Next i

If ActiveSheet.Name = "Welcome" Then
    TutForm.SecNum.Caption = "0"
    TutForm.SecDec.Caption = "0"
    TutForm.InfoClick.Caption = "0"
End If

tutWindow

TutForm.Show

End Sub

Sub tutCont()

Dim tutNum As Integer
Dim tutDec As Integer
Dim tutInfo As Integer

tutNum = TutForm.SecNum.Caption
tutDec = TutForm.SecDec.Caption
tutInfo = TutForm.InfoClick.Caption

Dim tutRow As Integer
Dim copyRow As Integer


Set copySheet = Worksheets("ShakeCast Ref Lookup Values")

If tutNum = "1" Then
    Set mySheet = Worksheets("Facility XML")
    tutRow = 4
ElseIf tutNum = "2" Then
    Set mySheet = Worksheets("Notification XML")
    tutRow = 4
ElseIf tutNum = "3" Then
    Set mySheet = Worksheets("User XML")
    tutRow = 4
End If

Application.ScreenUpdating = False

If tutNum = "0" And tutDec = "0" Then

    Worksheets("Welcome").Activate

    TutForm.DialogueBox = "Welcome to the ShakeCast Workbook tutorial!" & vbNewLine & vbNewLine & _
                    "This tutorial will walk you through how to fill out all the spreadsheets in this workbook, " & _
                    "beginning with the facility sheet. You can use the buttons at the bottom of this form to " & _
                    "navigate your way through the tutorial. You can also click on a specific section to the right, " & _
                    "and the tutorial will jump to that section."
                    

ElseIf tutNum = "1" And tutDec = "0" Then

    TutForm.DialogueBox = "Facility Spreadsheet" & vbNewLine & vbNewLine & _
        "When you hit continue, you will see any data you've entered in row 4 dissapear. Don't Panic! All of your " & _
        "data has been recorded and will be replaced when you finish (or quit) this tutorial." & vbNewLine & vbNewLine & _
        "Hit continue to learn how to use the Facility Spreadsheet!"
        
ElseIf tutNum = "1" And tutDec = "1" Then
    
    If tutInfo = 0 Then
    
        If copySheet.Range("A99").value = "no" Then
            copyRow = 100
            
            copySheet.Range("A" & copyRow & ":" & "AF" & copyRow).value = _
                mySheet.Range("A" & tutRow & ":" & "AF" & tutRow).value
        
            mySheet.Range("A" & tutRow & ":" & "AF" & tutRow).Clear
            mySheet.Range("A" & tutRow & ":" & "AF" & tutRow).Locked = False
            
            Application.Run "CheckFacilities", mySheet.Range("A" & tutRow)
            
            copySheet.Range("A99").value = "yes"
        End If
        
            mySheet.Range("A" & tutRow & ":" & "AF" & tutRow).Clear
            mySheet.Range("A" & tutRow & ":" & "AF" & tutRow).Locked = False
            
            mySheet.Range("A4").Activate
        
        TutForm.DialogueBox.Text = "Okay, We've copied your information, and cleared some space to do a demo." & vbNewLine & vbNewLine & _
            "Let's create a new facility that we want to represent the Golden Gate Bridge in our ShakeCast system." & vbNewLine & vbNewLine & _
            "Hit Continue!"
        
    ElseIf tutInfo = 1 Then

        mySheet.Range("A" & tutRow, "AF" & tutRow).Clear
        mySheet.Range("A" & tutRow, "AF" & tutRow).Locked = False
        
        mySheet.Range("A" & tutRow).Select
        
        TutForm.DialogueBox.Text = "Facility ID Info:" & _
            vbNewLine & vbNewLine & _
            "This field is required." & _
            vbNewLine & vbNewLine & _
            "The facility ID can be any combination of numbers and letters, but should not have any spaces. It's used to keep track of your facilities in our database in conjunction " & _
            "with the facility type." & vbNewLine & vbNewLine & _
            "A facility ID can be the same for multiple facilities of different types (like a bridge and a dam), " & _
            "but should not be the same for two facilities of the same type. For example, two bridges should not have " & _
            "the facility ID ""112""." & _
            vbNewLine & vbNewLine & _
            "It turns out that advanced users can actually break this rule. If you wish to define multiple fragilities to a single " & _
            "facility, this can be done by entering multiple facility rows with the same Facility ID and Facility Type, but with different " & _
            "Components. Component names can be changed from the advanced user spreadsheet" & vbNewLine & vbNewLine & _
            "This field is to be filled in by the user. This field is also mandatory for all facilites. If it's not completed, " & _
            "this facility will not be uploaded to the ShakeCast system." & vbNewLine & vbNewLine & _
            "Let's just enter ""1"" for this facility, since it's our first one. (You don't have to do this, during the tutorial, we will fill out the spreadsheet for you)" & vbNewLine & vbNewLine & _
            "Hit Continue!"
            
            
        
    ElseIf tutInfo = 2 Then
    

    
        mySheet.Range("A" & tutRow).value = 1
        
        
        TutForm.DialogueBox.Text = "Check that out! We entered a Facility ID, and the whole facility row changed colors!" & vbNewLine & vbNewLine & _
            "When any row in this whole workbook turns that color of blue, we are trying to tell you that this row will not be uploaded " & _
            "to the ShakeCast system. There is enough information in this row to tell us that you've manually entered some data, but " & _
            "we need more information before this row can be uploaded to ShakeCast!" & vbNewLine & vbNewLine & _
            "So let's continue to enter information until the row changes colors again." & vbNewLine & vbNewLine & _
            "Hit Continue!"
            
    ElseIf tutInfo = 3 Then
        
        ' already present values
        mySheet.Range("A" & tutRow).value = 1
        
        mySheet.Range("B" & tutRow, "AF" & tutRow).Clear
        mySheet.Range("B" & tutRow, "AF" & tutRow).Locked = False
        
        mySheet.Range("B" & tutRow).Select

    
        TutForm.DialogueBox.Text = "Facility Type Info:" & _
            vbNewLine & vbNewLine & _
            "This is another required field." & vbNewLine & vbNewLine & _
            "The facility type must be selected from the drop down menu. If you don't see a drop down menu in the cell, make sure " & _
            "that you've already completed the Facility ID section" & vbNewLine & vbNewLine & _
            "You can click the arrow by the side of the cell to show to drop down menu. For the Golden Gate Bridge, we will choose ""BRIDGE""."
            
    ElseIf tutInfo = 4 Then
    
        ' already present values
        mySheet.Range("A" & tutRow).value = 1
    
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        
        TutForm.DialogueBox.Text = "Column ""C"" automatically populated when we completed the Facility Type section. There are actually " & _
            "a few other fields that have populated automatically as well, but you can only see those from the advanced user spreadsheet!" & vbNewLine & vbNewLine & _
            "You should never have to type any information in column ""C"", so let's move along to column ""F"". What happened to columns " & _
            """D"" and ""E""? They should only be edited by advanced users!" & vbNewLine & vbNewLine & _
            "Hit Continue!"
            
    ElseIf tutInfo = 5 Then
    
        ' already present values
        mySheet.Range("A" & tutRow).value = 1
        mySheet.Range("B" & tutRow).value = "BRIDGE"
    
        mySheet.Range("F" & tutRow, "AF" & tutRow).Clear
        mySheet.Range("F" & tutRow, "AF" & tutRow).Locked = False
    
        mySheet.Range("F" & tutRow).Select
        
        TutForm.DialogueBox.Text = "Facility Full Name Info:" & vbNewLine & vbNewLine & _
            "This field is required. " & _
            "Here you can enter a name for this facility that will appear in the ShakeCast application. For this example, we'll type in ""Golden Gate Bridge""."
        
    ElseIf tutInfo = 6 Then
    
        mySheet.Range("G" & tutRow, "AF" & tutRow).Clear
        mySheet.Range("G" & tutRow, "AF" & tutRow).Locked = False
        
        ' already present values
        mySheet.Range("A" & tutRow).value = 1
        mySheet.Range("B" & tutRow).value = "BRIDGE"
    
        mySheet.Range("F" & tutRow).value = "Golden Gate Bridge"
        
        mySheet.Range("G" & tutRow, "H" & tutRow).Select
        
        TutForm.DialogueBox.Text = "The next two columns are optional. A Facility Description can be shown in the ShakeCast application, but " & _
            "is not required to send earthquake notifications. A Facility Short Name is displayed when ShakeCast can't display all the characters " & _
            "in your Facility Full Name." & vbNewLine & vbNewLine & _
            "We will just enter a short description of the Golden Gate Bridge and nickname it ""GGB""."
        
    
    ElseIf tutInfo = 7 Then
    
        mySheet.Range("K" & tutRow, "AD" & tutRow).Clear
        mySheet.Range("K" & tutRow, "AD" & tutRow).Locked = False
        
        ' already present values
        mySheet.Range("A" & tutRow).value = 1
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("F" & tutRow).value = "Golden Gate Bridge"
    
        mySheet.Range("G" & tutRow).value = "a bridge in San Francisco"
        mySheet.Range("H" & tutRow).value = "GGB"
        
        mySheet.Range("K" & tutRow, "L" & tutRow).Select
        
        TutForm.DialogueBox.Text = "We are now entering the Map and Display section of the Facility Worksheet! Here we describe where the " & _
            "facility is, and what it looks like. For most facilities, you will only need to enter a single latitude, longitude point." & vbNewLine & vbNewLine & _
            "ShakeCast treats a point location as a facility with an area of one square mile. If your facility will not fit inside a square mile " & _
            "we recommend you use multiple points to describe your facility." & vbNewLine & vbNewLine & _
            "For the case of the Golden Gate Bridge, let's use google maps to drop a pin right in the center of the bridge. We can then use the point " & _
            "google gives us to fill in this information."
    ElseIf tutInfo = 8 Then
    
        mySheet.Range("M" & tutRow, "AF" & tutRow).Clear
        mySheet.Range("M" & tutRow, "AF" & tutRow).Locked = False
        
        ' already present values
        mySheet.Range("A" & tutRow).value = 1
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("F" & tutRow).value = "Golden Gate Bridge"
        mySheet.Range("G" & tutRow).value = "a bridge in San Francisco"
        mySheet.Range("H" & tutRow).value = "GGB"
    
        mySheet.Range("K" & tutRow).value = "37.819929"
        mySheet.Range("L" & tutRow).value = "-122.478255"
        
        mySheet.Range("M" & tutRow).Select
        
        TutForm.DialogueBox.Text = "Nice! We now have a point location for our facility. Did you notice color of the row changed back to normal? " & _
            "This means that we've input enough information for our data to be recognized as a facility by the ShakeCast application! We could stop here, " & _
            "but it's a good idea to enter fragility information for your facilities as well. Although it isn't mandatory, we'll go over the HTML Snippet section now. "
            
            
    ElseIf tutInfo = 9 Then
    
        ' already present values
        mySheet.Range("A" & tutRow).value = 1
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("F" & tutRow).value = "Golden Gate Bridge"
        mySheet.Range("G" & tutRow).value = "a bridge in San Francisco"
        mySheet.Range("H" & tutRow).value = "GGB"
        mySheet.Range("K" & tutRow).value = "37.819929"
        mySheet.Range("L" & tutRow).value = "-122.478255"
    
        TutForm.DialogueBox.Text = "The HTML Snippet gives ShakeCast some information to display on the map. This HTML information can also be included in the " & _
            "facility specific custom shaking report. For this case, I will just put in a simple title."

    ElseIf tutInfo = 10 Then
        
        mySheet.Range("N" & tutRow, "AF" & tutRow).Clear
        mySheet.Range("N" & tutRow, "AF" & tutRow).Locked = False
        
        ' already present values
        mySheet.Range("A" & tutRow).value = 1
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("F" & tutRow).value = "Golden Gate Bridge"
        mySheet.Range("G" & tutRow).value = "a bridge in San Francisco"
        mySheet.Range("H" & tutRow).value = "GGB"
        mySheet.Range("K" & tutRow).value = "37.819929"
        mySheet.Range("L" & tutRow).value = "-122.478255"
        
        mySheet.Range("M" & tutRow).value = "<h1>The Golden Gate Bridge</h1>"
    
        mySheet.Range("N" & tutRow).Select
        
        TutForm.DialogueBox.Text = "Some of you might notice that this entry contains some HTML and XML reserved characters. This is okay! We will convert all " & _
                "of these characters for you when you export your XML document. The only character we will not replace for you is the ampersand (&). " & _
                "We don't convert this character, just in case you are using it as an escape character. If you wish to include an ampersand in your facility name or any other field, type ""&amp"" (without the quotes) instead of ""&"". " & _
                "Now onto the facility fragility!" & vbNewLine & vbNewLine & _
                "We allow you to pick from many HAZUS model building types " & _
                "or to define your own fragilities. If you wish to define your own fragility information, you can do so from the " & _
                "options menu in the advanced user worksheet." & vbNewLine & vbNewLine & _
                "You can select a fragility model name from the drop down menu. If you've defined your own fragility model, it will appear here as well! " & _
                "For this case, since we don't really know the fragility of the Golden Gate Bridge (or at least I don't off hand), we will choose a generic " & _
                "fragility. It's our recommendation to do this for all facilities whose fragilities are currently unknown."
    
    ElseIf tutInfo = 11 Then
    
        mySheet.Range("O" & tutRow, "AF" & tutRow).Clear
        mySheet.Range("O" & tutRow, "AF" & tutRow).Locked = False
        
        ' already present values
        mySheet.Range("A" & tutRow).value = 1
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("F" & tutRow).value = "Golden Gate Bridge"
        mySheet.Range("G" & tutRow).value = "a bridge in San Francisco"
        mySheet.Range("H" & tutRow).value = "GGB"
        mySheet.Range("K" & tutRow).value = "37.819929"
        mySheet.Range("L" & tutRow).value = "-122.478255"
        mySheet.Range("M" & tutRow).value = "<h1>The Golden Gate Bridge</h1>"
    
        mySheet.Range("N" & tutRow).value = "GENERIC"
        
        TutForm.DialogueBox.Text = "As you can see, the fragility information automatically populates as soon as the model name is chosen. That's it! We've " & _
            "completed the entire facility row! Now I'd like to go back and talk about a couple different kinds of geometry definitions you can use for your " & _
            "facilities." & vbNewLine & vbNewLine & _
            "Hit Continue to learn how to define different geometry types for your facility, or click on ""Notification Spreadsheet"" to skip ahead " & _
            "and learn how to define notification groups."
            
    End If
    
ElseIf tutNum = "1" And tutDec = "2" Then
    
    If tutInfo = 0 Then
    
        mySheet.Range("K" & tutRow, "L" & tutRow).Select
        
        ' already present values
        mySheet.Range("A" & tutRow).value = 1
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("F" & tutRow).value = "Golden Gate Bridge"
        mySheet.Range("G" & tutRow).value = "a bridge in San Francisco"
        mySheet.Range("H" & tutRow).value = "GGB"
        mySheet.Range("K" & tutRow).value = "37.819929"
        mySheet.Range("L" & tutRow).value = "-122.478255"
        mySheet.Range("M" & tutRow).value = "<h1>The Golden Gate Bridge</h1>"
        mySheet.Range("N" & tutRow).value = "GENERIC"
    
        TutForm.DialogueBox.Text = "What would it look like if we wanted to use multiple points " & _
            "to describe this facility? As it turns out, the Golden Gate Bridge is actually around 9000 feet long, so it makes sense to use multiple " & _
            "points to describe this facility since it will not fit in a square mile box. This bridge is long, but skinny. By using one point on each end of the bridge, we can monitor those points as well as the space in between them." & vbNewLine & vbNewLine & _
            "The North end has the coordinates: " & vbNewLine & _
            "Latitude: 37.825777" & vbNewLine & _
            "Longitude: -122.479199" & vbNewLine & vbNewLine & _
            "The South end has the coordinates:" & vbNewLine & _
            "Latitude: 37.810284" & vbNewLine & _
            "Longitude: -122.477568" & vbNewLine & vbNewLine & _
            "We want to concatinate all of the latitude coordinates and all of the longitude coordinates so that they can be processed the right way. " & _
            "We use a semi-colon to seperate the individual coordinates. So in this case our entry into the latitude column will look like: " & vbNewLine & vbNewLine & _
            "37.825777;37.810284" & vbNewLine & vbNewLine & _
            "and our entry into the longitude column will look like: " & vbNewLine & vbNewLine & _
            "-122.479199;-122.477568" & vbNewLine & vbNewLine & "Hit Continue!"
    ElseIf tutInfo = 1 Then
    
        ' already present values
        mySheet.Range("A" & tutRow).value = 1
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("F" & tutRow).value = "Golden Gate Bridge"
        mySheet.Range("G" & tutRow).value = "a bridge in San Francisco"
        mySheet.Range("H" & tutRow).value = "GGB"
        mySheet.Range("M" & tutRow).value = "<h1>The Golden Gate Bridge</h1>"
        mySheet.Range("N" & tutRow).value = "GENERIC"
        
        mySheet.Range("K" & tutRow).value = "37.825777;37.810284"
        mySheet.Range("L" & tutRow).value = "-122.479199;-122.477568"
        
        TutForm.DialogueBox.Text = "Nice! Now the Golden Gate Bridge has what we call a ""Polyline"" geometry, which we define as more than one " & _
            "point that is not an enclosed shape. If we were to add another point and then the first point again, we would have a ""Polygon"" geometry." & vbNewLine & vbNewLine & _
            "Hit Continue to see a Polygon example and enter some fragility information!"

    End If
    
ElseIf tutNum = "1" And tutDec = "3" Then
    
    If tutInfo = 0 Then
    
        ' already present values
        mySheet.Range("A" & tutRow).value = 1
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("F" & tutRow).value = "Golden Gate Bridge"
        mySheet.Range("G" & tutRow).value = "a bridge in San Francisco"
        mySheet.Range("H" & tutRow).value = "GGB"
        mySheet.Range("M" & tutRow).value = "<h1>The Golden Gate Bridge</h1>"
        mySheet.Range("N" & tutRow).value = "GENERIC"
        mySheet.Range("K" & tutRow).value = "37.825777;37.810284"
        mySheet.Range("L" & tutRow).value = "-122.479199;-122.477568"
    
        TutForm.DialogueBox.Text = "Although the Golden Gate Bridge is best described as a Polyline, I'd like to do an example of the " & _
            "Polygon geometry." & vbNewLine & vbNewLine & _
            "The best way to do this is to make a shape that outlines the facility you are trying to describe. For this case, we will make a box that encloses " & _
            "the bridge. We can do this by selecting points at all four corners of the bridge. We will concatinate the points just like we did for the " & _
            "polyline, but in this case we will repeat the first point at the end. " & vbNewLine & vbNewLine & _
            "I used google maps to drop pins at all four corners of the bridge. It gave the following latitude/longitude points:" & vbNewLine & vbNewLine & _
            "Points: (37.832316, -122.480710), (37.832229, -122.480855), (37.807294, -122.475324), (37.807044, -122.475662)" & vbNewLine & vbNewLine & _
            "Now we can string all the latitude and longitude values together using a semi-colon: " & vbNewLine & vbNewLine & _
            "Latitude: 37.832316;37.832229;37.807294;37.807044;37.832316" & vbNewLine & _
            "Longitude: -122.480710;-122.480855;-122.475324;-122.475662;-122.480710" & vbNewLine & vbNewLine & _
            "Wait a second! We are only defining four points, but it looks like there are five numbers in both the latitude and longitude strings! " & _
            "Remember, in order to enclose the polygon, we have to redefine the first point after the last new point." & vbNewLine & vbNewLine & _
            "Hit Continue to submit these values!"
            
    ElseIf tutInfo = 1 Then
    
        ' already present values
        mySheet.Range("A" & tutRow).value = 1
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("F" & tutRow).value = "Golden Gate Bridge"
        mySheet.Range("G" & tutRow).value = "a bridge in San Francisco"
        mySheet.Range("H" & tutRow).value = "GGB"
        mySheet.Range("M" & tutRow).value = "<h1>The Golden Gate Bridge</h1>"
        mySheet.Range("N" & tutRow).value = "GENERIC"
    
        mySheet.Range("K" & tutRow).value = "37.832316;37.832229;37.807294;37.807044;37.832316"
        mySheet.Range("L" & tutRow).value = "-122.480710;-122.480855;-122.475324;-122.475662;-122.480710"
        
    ElseIf tutInfo = 2 Then
        TutForm.DialogueBox.Text = "That's it for the facility worksheet! If you've questions while you are filling out the worksheet " & _
            "click on the ""More Info"" buttons or refer back to this tutorial. " & _
            "Hit Continue when you are ready to start learning about the Notification Worksheet!"
    End If
    
ElseIf tutNum = "2" And tutDec = "0" Then
    
    mySheet.Range("A4").Activate
    
    If tutInfo = 1 Then
        TutForm.DialogueBox.Text = "Welcome to the Notification Worksheet! Notification groups are used to determine who gets notifications, when. " & _
            "Each row in this spreadsheet represents a situation in which users will get a notification. For the purpose of this tutorial, lets create " & _
            "a notification group that will recieve alerts when there is an earthquake in California or when there is any ground shaking at the Golden Gate Bridge, which we just defined in the last part of the tutorial!" & vbNewLine & vbNewLine & _
            "We will need to make some more space to do this tutorial, so when you hit continue we will move the information you've input in rows 4-8 to another worksheet."
            
    ElseIf tutInfo = 2 Then
        
    ElseIf tutInfo = 3 Then
    
    End If
    
ElseIf tutNum = "2" And tutDec = "1" Then
    
    If tutInfo = 0 Then
        If copySheet.Range("A199").value = "no" Then
            copyRow = 200

            copySheet.Range("A" & copyRow & ":" & "Q" & copyRow + 9).value = _
                mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).value

            mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Clear
            mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Locked = False


            copySheet.Range("A199").value = "yes"
        End If
        
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Clear
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Locked = False
        
        TutForm.DialogueBox.Text = "Alright, we've cleared some space for us to do our tutorial. Remember your data has been saved in another " & _
            "location, and when you quit the tutorial it will be moved back! " & vbNewLine & vbNewLine & _
            "Let's create a new group called CAL_BRIDGES. This group will receive alerts any time there's an earthquake in California or " & _
            "damage to any of the bridges we are uploading to shakecast." & vbNewLine & vbNewLine & _
            "Hit Continue to start building this Notification Group!"
            
    ElseIf tutInfo = 1 Then
        
        mySheet.Range("B" & tutRow & ":" & "Q" & tutRow + 9).Clear
        
        mySheet.Range("A" & tutRow + 1 & ":" & "Q" & tutRow + 9).Clear
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Locked = False
        
        mySheet.Range("A" & tutRow).value = "CAL_BRIDGES"
    
        mySheet.Range("B" & tutRow).Select
        
        TutForm.DialogueBox.Text = "We've started off by naming the notification group. Notice that some other information has automatically populated! The next piece of info we need to enter is the " & _
            "type of facility we would like to monitor. From the advanced user worksheet, you can add multiple facility types by secting " & _
            "the Facility Type column in a group row. In the General User worksheet, you can select the facility type from a drop down menu. " & vbNewLine & vbNewLine & _
            "For this example, we select the BRIDGE facility type."
        
    ElseIf tutInfo = 2 Then
        
        mySheet.Range("B" & tutRow & ":" & "Q" & tutRow + 9).Clear
        
        mySheet.Range("A" & tutRow + 1 & ":" & "Q" & tutRow + 9).Clear
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Locked = False
        
        Application.EnableEvents = False
        ' already present values
        mySheet.Range("A" & tutRow).value = "CAL_BRIDGES"
        Application.EnableEvents = True
        
        mySheet.Range("B" & tutRow).value = "BRIDGE"
    
        mySheet.Range("C" & tutRow).Select
        
        TutForm.DialogueBox.Text = "Done! Alright, now it's time to define the monitoring region. You will only receive notifications " & _
            "when there is shaking inside of this defined region. For this case, we will make an outline around California." & vbNewLine & vbNewLine & _
            "You can define a monitoring region with latitude and longitude points. It's best to seperate the latitude/longitude numbers that " & _
            "define a point by a space, and to seperate points by a semi-colon. This looks like: " & vbNewLine & vbNewLine & _
            "lat1 lon1;lat2 lon2;lat3 lon3;lat1 lon1" & vbNewLine & vbNewLine & _
            "Notice that the last point is the same as the first point. This is how the monitoring region is ""closed""." & vbNewLine & vbNewLine & _
            "The actual region that we're defining looks like:" & vbNewLine & vbNewLine & _
            "43 -126;39 -126;34 -123;31 -118;31 -113;36 -113;39 -118;43 -118;43 -126"
    
    
    ElseIf tutInfo = 3 Then
        
        mySheet.Range("C" & tutRow & ":" & "Q" & tutRow + 9).Clear
        
        mySheet.Range("A" & tutRow + 1 & ":" & "Q" & tutRow + 9).Clear
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Locked = False
        
        Application.EnableEvents = False
        ' already present values
        mySheet.Range("A" & tutRow).value = "CAL_BRIDGES"
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        Application.EnableEvents = True
        
        mySheet.Range("C" & tutRow).value = "43 -126;39 -126;34 -123;31 -118;31 -113;36 -113;39 -118;43 -118;43 -126"
    
        mySheet.Range("D" & tutRow).Select
        
        TutForm.DialogueBox.Text = "The next field is the notification type. This is where you define whether a new event or possible damage to a facility triggers an alert message to be sent. " & _
            "It's possible for a single group to receive notifications when new earthquakes occur as well as when there is possible damage to your facilities. " & _
            "In order for a group to recieve notifications for both of these situations, we must make a notification row for each. " & vbNewLine & vbNewLine & _
            "We'll set up this group to get notifications for new earthquakes as well as all damage probabilities. This first group row will define how notifications are sent for new earthquakes. We'll select NEW_EVENT from the drop down menu in this column."
    
    ElseIf tutInfo = 4 Then
        
        mySheet.Range("D" & tutRow & ":" & "Q" & tutRow + 9).Clear
        
        mySheet.Range("A" & tutRow + 1 & ":" & "Q" & tutRow + 9).Clear
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Locked = False
        
        Application.EnableEvents = False
        ' already present values
        mySheet.Range("A" & tutRow).value = "CAL_BRIDGES"
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("C" & tutRow).value = "43 -126;39 -126;34 -123;31 -118;31 -113;36 -113;39 -118;43 -118;43 -126"
        Application.EnableEvents = True
        
        mySheet.Range("D" & tutRow).value = "NEW_EVENT"
    
        mySheet.Range("E" & tutRow).Select
        
        TutForm.DialogueBox.Text = "Check that out! The inspection priority column turned grey when we selected NEW_EVENT. In this worksheet, " & _
            "you never need to enter any input into grey cells. Any information that could be entered into grey cells is unecessary and will be removed." & vbNewLine & vbNewLine & _
            "We'll talk about Inspection Priority in the next row, but for now let's move onto Minimum Magnitude, as that is the next relevent cell for this notification row." & vbNewLine & vbNewLine & _
            "The Minimum Magnitude defines the lowest magnitude earthquake that will trigger this notification. The default is 3; this is the lowest magnitude earthquake that can send you notifications for. " & _
            "For our case, let's just leave it at magnitude 3!"
    
    
    ElseIf tutInfo = 5 Then
        
        mySheet.Range("F" & tutRow & ":" & "Q" & tutRow + 9).Clear
        
        mySheet.Range("A" & tutRow + 1 & ":" & "Q" & tutRow + 9).Clear
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Locked = False
        
        Application.EnableEvents = False
        ' already present values
        mySheet.Range("A" & tutRow).value = "CAL_BRIDGES"
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("C" & tutRow).value = "43 -126;39 -126;34 -123;31 -118;31 -113;36 -113;39 -118;43 -118;43 -126"
        mySheet.Range("D" & tutRow).value = "NEW_EVENT"
        Application.EnableEvents = True
        
        mySheet.Range("F" & tutRow).value = "3"
    
        mySheet.Range("G" & tutRow).Select
    
        TutForm.DialogueBox.Text = "The next section has been filled out since the begining! The Event Type describes what kind of event " & _
            "will trigger an alert message. The default, ""ACTUAL"", will only send alerts when the actual earthquake hits. If you select ""SCENARIO"", this group will " & _
            "receive notifications only when you run scenarios. The ""HEARTBEAT"" option will cause a group to receive messages each day when the system does a check to make sure eveything is working!" & vbNewLine & vbNewLine & _
            "Since we only want this group to recieve notifications when an actual earthquake occurs, we will choose ""ACTUAL""."
            
    ElseIf tutInfo = 6 Then
    
'        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Clear
'        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Locked = False
    
        mySheet.Range("H" & tutRow).Select
        
'        Application.EnableEvents = False
'        ' already present values
'        mySheet.Range("A" & tutRow).Value = "CAL_BRIDGES"
'        mySheet.Range("B" & tutRow).Value = "BRIDGE"
'        mySheet.Range("C" & tutRow).Value = "43 -126;39 -126;34 -123;31 -118;31 -113;36 -113;39 -118;43 -118;43 -126"
'        mySheet.Range("D" & tutRow).Value = "NEW_EVENT"
'        mySheet.Range("F" & tutRow).Value = "3"
'        Application.EnableEvents = True
        

        
        TutForm.DialogueBox.Text = "This last column describes how the email you receive will look. All of these messages " & _
            "are sent by email. Rich Content messages use HTML formatting to make PDFs. Plain Text might be a better " & _
            "option if you are only looking for a barebones message. The PAGER option will send you a PAGER content message. For this example, we will leave Rich Content as our selection."
            
    ElseIf tutInfo = 7 Then
    
'        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Clear
'        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Locked = False
'
'        Application.EnableEvents = False
'        ' already present values
'        mySheet.Range("A" & tutRow).Value = "CAL_BRIDGES"
'        mySheet.Range("B" & tutRow).Value = "BRIDGE"
'        mySheet.Range("C" & tutRow).Value = "43 -126;39 -126;34 -123;31 -118;31 -113;36 -113;39 -118;43 -118;43 -126"
'        mySheet.Range("D" & tutRow).Value = "NEW_EVENT"
'        mySheet.Range("F" & tutRow).Value = "3"
'        Application.EnableEvents = True
        
        TutForm.DialogueBox.Text = "That's it, we've completed our first notification row! The notification group CAL_BRIDGES will now recieve " & _
            "notifications any time there is a new earthquake in the monitoring region. We also want to receive damage probabilities for our BRIDGE facilities, so we will " & _
            "create some more notification rows for that purpose."
            
            
    ElseIf tutInfo = 8 Then
    
        mySheet.Range("A" & tutRow + 1 & ":" & "Q" & tutRow + 9).Clear
        
        mySheet.Range("A" & tutRow + 2 & ":" & "Q" & tutRow + 9).Clear
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Locked = False

        Application.EnableEvents = False
        ' already present values
        mySheet.Range("A" & tutRow).value = "CAL_BRIDGES"
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("C" & tutRow).value = "43 -126;39 -126;34 -123;31 -118;31 -113;36 -113;39 -118;43 -118;43 -126"
        mySheet.Range("D" & tutRow).value = "NEW_EVENT"
        mySheet.Range("F" & tutRow).value = "3"
        mySheet.Range("G" & tutRow).value = "ACTUAL"
        mySheet.Range("H" & tutRow).value = "Rich Content"
        Application.EnableEvents = True
    
        mySheet.Range("A" & tutRow + 1).Select
    
        TutForm.DialogueBox.Text = "In order to add additional notifications to the group CAL_BRIDGES, we must add additional notification rows with the same group name. " & _
            "So, lets enter CAL_BRIDGES on line 5."
            
    ElseIf tutInfo = 9 Then
        
        mySheet.Range("A" & tutRow + 1 & ":" & "Q" & tutRow + 9).Clear
        
        mySheet.Range("A" & tutRow + 2 & ":" & "Q" & tutRow + 9).Clear
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Locked = False
    
        Application.EnableEvents = False
        ' already present values
        mySheet.Range("A" & tutRow).value = "CAL_BRIDGES"
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("C" & tutRow).value = "43 -126;39 -126;34 -123;31 -118;31 -113;36 -113;39 -118;43 -118;43 -126"
        mySheet.Range("D" & tutRow).value = "NEW_EVENT"
        mySheet.Range("F" & tutRow).value = "3"
        mySheet.Range("G" & tutRow).value = "ACTUAL"
        mySheet.Range("H" & tutRow).value = "Rich Content"
        Application.EnableEvents = True
    
        mySheet.Range("A" & tutRow + 1).value = "CAL_BRIDGES"
        
        TutForm.DialogueBox.Text = "Look at that! We've created a second notification row for this group and already there are more grey cells than we saw " & _
            "in the first row. This is because there is some information that we only want defined one time. For instance, we only want the monitoring region " & _
            "to be defined one time in order to minimize potential for user error. Although we won't let you define multiple facility types in the general user worksheet, " & _
            "this can be done from the advanced user worksheet." & vbNewLine & vbNewLine & _
            "The next column we have to fill out is the Notification Type. We already get notifications when new events occur, so let's make this row a DAMAGE notification. " & _
            "It's important to note that you will not be notified of damage to your facilities, but of POSSIBLE damage to your facilities."
        
    ElseIf tutInfo = 10 Then

        mySheet.Range("D" & tutRow + 1 & ":" & "Q" & tutRow + 9).Clear
        
        mySheet.Range("A" & tutRow + 2 & ":" & "Q" & tutRow + 9).Clear
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Locked = False
        
        Application.EnableEvents = False
        ' already present values
        mySheet.Range("A" & tutRow).value = "CAL_BRIDGES"
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("C" & tutRow).value = "43 -126;39 -126;34 -123;31 -118;31 -113;36 -113;39 -118;43 -118;43 -126"
        mySheet.Range("D" & tutRow).value = "NEW_EVENT"
        mySheet.Range("F" & tutRow).value = "3"
        mySheet.Range("G" & tutRow).value = "ACTUAL"
        mySheet.Range("H" & tutRow).value = "Rich Content"
        mySheet.Range("A" & tutRow + 1).value = "CAL_BRIDGES"
        Application.EnableEvents = True
    
        mySheet.Range("D" & tutRow + 1).value = "DAMAGE"
        mySheet.Range("D" & tutRow + 1).Select
        
        TutForm.DialogueBox.Text = "Notice that the next column otimatically populated ""Green""! These colors represent the possible severity of the damage to your facility. " & _
            "The order of severity is GREEN->YELLOW->ORANGE->RED. You will get a GREEN alert when there is possible minimal damage to one of your facilities, and a RED alert when there " & _
            "is possible severe damage to one of your facilities. YELLOW and ORANGE are somewhere in the middle. We actually want to recieve notifications for each of these alert levels, " & _
            "so lets make a notification row for each level. "
            
    ElseIf tutInfo = 11 Then
    
        mySheet.Range("E" & tutRow + 1 & ":" & "Q" & tutRow + 9).Clear
        
        mySheet.Range("A" & tutRow + 2 & ":" & "Q" & tutRow + 9).Clear
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Locked = False
    
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        
        mySheet.Range("A" & tutRow + 2).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 3).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 4).value = "CAL_BRIDGES"
    
        mySheet.Range("D" & tutRow + 2).value = "DAMAGE"
        mySheet.Range("D" & tutRow + 3).value = "DAMAGE"
        mySheet.Range("D" & tutRow + 4).value = "DAMAGE"
        
        mySheet.Range("E" & tutRow + 2).value = "YELLOW"
        mySheet.Range("E" & tutRow + 3).value = "ORANGE"
        
        
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        
        mySheet.Range("E" & tutRow + 4).value = "RED"
        mySheet.Range("E" & tutRow + 4).Select
        
        Application.EnableEvents = False
        ' already present values
        mySheet.Range("A" & tutRow).value = "CAL_BRIDGES"
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("C" & tutRow).value = "43 -126;39 -126;34 -123;31 -118;31 -113;36 -113;39 -118;43 -118;43 -126"
        mySheet.Range("D" & tutRow).value = "NEW_EVENT"
        mySheet.Range("F" & tutRow).value = "3"
        mySheet.Range("G" & tutRow).value = "ACTUAL"
        mySheet.Range("H" & tutRow).value = "Rich Content"
        mySheet.Range("A" & tutRow + 1).value = "CAL_BRIDGES"
        mySheet.Range("D" & tutRow + 1).value = "DAMAGE"
        Application.EnableEvents = True
    
        
        
        TutForm.DialogueBox.Text = "Now we have a notification row for each inspection priority. This means that the notification group will receieve an alert any time " & _
            "there is perspective damage to one of the facilities associated with the group. All I mean by ""associated with the group"" is that their facility type matches that of the " & _
            "notification group, and their defined location falls within the group's monitoring region." & vbNewLine & vbNewLine & _
            "That's it! We've defined our notification group. Notice now what happens when one of the notification rows is left incomplete."
            
    ElseIf tutInfo = 12 Then
    
        mySheet.Range("A" & tutRow + 5 & ":" & "Q" & tutRow + 9).Clear
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Locked = False
        
    
        Application.EnableEvents = False
        ' already present values
        mySheet.Range("A" & tutRow).value = "CAL_BRIDGES"
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("C" & tutRow).value = "43 -126;39 -126;34 -123;31 -118;31 -113;36 -113;39 -118;43 -118;43 -126"
        mySheet.Range("D" & tutRow).value = "NEW_EVENT"
        mySheet.Range("F" & tutRow).value = "3"
        mySheet.Range("G" & tutRow).value = "ACTUAL"
        mySheet.Range("H" & tutRow).value = "Rich Content"
        mySheet.Range("A" & tutRow + 1).value = "CAL_BRIDGES"
        mySheet.Range("D" & tutRow + 1).value = "DAMAGE"
        mySheet.Range("A" & tutRow + 2).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 3).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 4).value = "CAL_BRIDGES"
    
        mySheet.Range("D" & tutRow + 2).value = "DAMAGE"
        mySheet.Range("D" & tutRow + 3).value = "DAMAGE"
        mySheet.Range("D" & tutRow + 4).value = "DAMAGE"
        
        mySheet.Range("E" & tutRow + 2).value = "YELLOW"
        mySheet.Range("E" & tutRow + 3).value = "ORANGE"
        mySheet.Range("E" & tutRow + 4).value = "RED"
        Application.EnableEvents = True
    
        mySheet.Range("D" & tutRow + 2).value = Empty
        mySheet.Range("D" & tutRow + 2).Select
        
        TutForm.DialogueBox.Text = "You can see that only the single notification row is invalidated. All the rest of the notification specifications will still be " & _
            "uploaded to the ShakeCast application. Hit continue to see what happens when the first row is invalid!"
            
    ElseIf tutInfo = 13 Then
        
        mySheet.Range("A" & tutRow + 5 & ":" & "Q" & tutRow + 9).Clear
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Locked = False
        
        Application.EnableEvents = False
        ' already present values
        mySheet.Range("A" & tutRow).value = "CAL_BRIDGES"
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("C" & tutRow).value = "43 -126;39 -126;34 -123;31 -118;31 -113;36 -113;39 -118;43 -118;43 -126"
        mySheet.Range("D" & tutRow).value = "NEW_EVENT"
        mySheet.Range("F" & tutRow).value = "3"
        mySheet.Range("G" & tutRow).value = "ACTUAL"
        mySheet.Range("H" & tutRow).value = "Rich Content"
        mySheet.Range("A" & tutRow + 1).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 2).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 3).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 4).value = "CAL_BRIDGES"
        
        mySheet.Range("D" & tutRow + 2).value = "DAMAGE"
        mySheet.Range("D" & tutRow + 3).value = "DAMAGE"
        mySheet.Range("D" & tutRow + 4).value = "DAMAGE"
        
        mySheet.Range("E" & tutRow + 2).value = "YELLOW"
        mySheet.Range("E" & tutRow + 3).value = "ORANGE"
        mySheet.Range("E" & tutRow + 4).value = "RED"
        Application.EnableEvents = True
        
        mySheet.Range("D" & tutRow + 2).value = "DAMAGE"
        mySheet.Range("D" & tutRow).value = Empty
        mySheet.Range("D" & tutRow).Select
        
        TutForm.DialogueBox.Text = "Now you can see that the entire notification group is invalid! This happens because the first row in a notification group defines " & _
            "multiple pieces of information for the entire group. For instance, if all the valid rows were to be uploaded to ShakeCast still, they would be lacking a Monitoring Reigon! "
            
    End If
    
ElseIf tutNum = "2" And tutDec = "2" Then
           
    If tutInfo = 0 Then
    
        mySheet.Range("A" & tutRow + 5 & ":" & "Q" & tutRow + 9).Clear
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Locked = False
    
        Application.EnableEvents = False
        ' already present values
        mySheet.Range("A" & tutRow).value = "CAL_BRIDGES"
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("C" & tutRow).value = "43 -126;39 -126;34 -123;31 -118;31 -113;36 -113;39 -118;43 -118;43 -126"
        mySheet.Range("F" & tutRow).value = "3"
        mySheet.Range("G" & tutRow).value = "ACTUAL"
        mySheet.Range("H" & tutRow).value = "Rich Content"
        mySheet.Range("A" & tutRow + 1).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 2).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 3).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 4).value = "CAL_BRIDGES"
        
        mySheet.Range("D" & tutRow + 2).value = "DAMAGE"
        mySheet.Range("D" & tutRow + 3).value = "DAMAGE"
        mySheet.Range("D" & tutRow + 4).value = "DAMAGE"
        
        mySheet.Range("E" & tutRow + 2).value = "YELLOW"
        mySheet.Range("E" & tutRow + 3).value = "ORANGE"
        mySheet.Range("E" & tutRow + 4).value = "RED"
        
        mySheet.Range("D" & tutRow + 1).value = "DAMAGE"
        Application.EnableEvents = True
    
        mySheet.Range("D" & tutRow).Activate
        mySheet.Range("D" & tutRow).value = "NEW_EVENT"
    
        TutForm.DialogueBox.Text = "Some of you may find it useful to run scenarios. If you are planning to run scenarios, it's a good idea to make a notification group that exclusively gets " & _
            "scenario alerts. This way, you can be sure that none of your emergency responders are going to get a scenario alert and think that it's the real deal. Let's create a second group, " & _
            "identical to the first, that only receives notifictions when a scenario is run. So first let's copy the group we just created, select row 9, column A and hit paste. It's a good idea to " & _
            "copy rows in this workbook by clicking on the row number on the left-hand side of the screen."
            
    ElseIf tutInfo = 1 Then
    
        
        mySheet.Range("A" & tutRow + 5 & ":" & "Q" & tutRow + 9).Clear
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Locked = False
    
        Application.EnableEvents = False
        ' already present values
        mySheet.Range("A" & tutRow).value = "CAL_BRIDGES"
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("C" & tutRow).value = "43 -126;39 -126;34 -123;31 -118;31 -113;36 -113;39 -118;43 -118;43 -126"
        mySheet.Range("D" & tutRow).value = "NEW_EVENT"
        mySheet.Range("F" & tutRow).value = "3"
        mySheet.Range("G" & tutRow).value = "ACTUAL"
        mySheet.Range("H" & tutRow).value = "Rich Content"
        mySheet.Range("A" & tutRow + 1).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 2).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 3).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 4).value = "CAL_BRIDGES"
        
        mySheet.Range("D" & tutRow + 2).value = "DAMAGE"
        mySheet.Range("D" & tutRow + 3).value = "DAMAGE"
        mySheet.Range("D" & tutRow + 4).value = "DAMAGE"
        
        mySheet.Range("E" & tutRow + 2).value = "YELLOW"
        mySheet.Range("E" & tutRow + 3).value = "ORANGE"
        mySheet.Range("E" & tutRow + 4).value = "RED"
        
        mySheet.Range("D" & tutRow + 1).value = "DAMAGE"
        Application.EnableEvents = True
    
        mySheet.Range("A" & tutRow + 5 & ":" & "Q" & tutRow + 9).value = mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 4).value
        
        TutForm.DialogueBox.Text = "Now we should change the name of the group which we want to get scenario notifications to CAL_BRIDGES_SCENARIO. This way, it becomes obvious which users will get " & _
            "notifications when a scenario is run."
        
    ElseIf tutInfo = 2 Then
    
    
        Application.EnableEvents = False
        ' already present values
        mySheet.Range("A" & tutRow).value = "CAL_BRIDGES"
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("C" & tutRow).value = "43 -126;39 -126;34 -123;31 -118;31 -113;36 -113;39 -118;43 -118;43 -126"
        mySheet.Range("D" & tutRow).value = "NEW_EVENT"
        mySheet.Range("F" & tutRow).value = "3"
        mySheet.Range("G" & tutRow).value = "ACTUAL"
        mySheet.Range("H" & tutRow).value = "Rich Content"
        mySheet.Range("A" & tutRow + 1).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 2).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 3).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 4).value = "CAL_BRIDGES"
        
        mySheet.Range("D" & tutRow + 2).value = "DAMAGE"
        mySheet.Range("D" & tutRow + 3).value = "DAMAGE"
        mySheet.Range("D" & tutRow + 4).value = "DAMAGE"
        
        mySheet.Range("E" & tutRow + 2).value = "YELLOW"
        mySheet.Range("E" & tutRow + 3).value = "ORANGE"
        mySheet.Range("E" & tutRow + 4).value = "RED"
        
        mySheet.Range("D" & tutRow + 1).value = "DAMAGE"
        mySheet.Range("A" & tutRow + 5 & ":" & "Q" & tutRow + 9).value = mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 4).value
        Application.EnableEvents = True

        
        mySheet.Range("C" & tutRow + 5).value = Empty
        mySheet.Range("A" & tutRow + 5).value = "CAL_BRIDGES_SCENARIO"
        mySheet.Range("A" & tutRow + 6).value = "CAL_BRIDGES_SCENARIO"
        mySheet.Range("A" & tutRow + 7).value = "CAL_BRIDGES_SCENARIO"
        mySheet.Range("A" & tutRow + 8).value = "CAL_BRIDGES_SCENARIO"
        mySheet.Range("A" & tutRow + 9).value = "CAL_BRIDGES_SCENARIO"
        
        
        mySheet.Range("A" & tutRow + 5 & ":" & "A" & tutRow + 9).Activate
        
        TutForm.DialogueBox.Text = "Notice now how the color of the second group has changed. Since the group name has been changed, these notification rows are no longer associated with the monitoring region defined in row 4. Lets copy and paste that monitoring region down to row 9. "
    
    ElseIf tutInfo = 3 Then
    
        Application.EnableEvents = False
        ' already present values
        mySheet.Range("A" & tutRow).value = "CAL_BRIDGES"
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("C" & tutRow).value = "43 -126;39 -126;34 -123;31 -118;31 -113;36 -113;39 -118;43 -118;43 -126"
        mySheet.Range("D" & tutRow).value = "NEW_EVENT"
        mySheet.Range("F" & tutRow).value = "3"
        mySheet.Range("G" & tutRow).value = "ACTUAL"
        mySheet.Range("H" & tutRow).value = "Rich Content"
        mySheet.Range("A" & tutRow + 1).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 2).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 3).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 4).value = "CAL_BRIDGES"
        
        mySheet.Range("D" & tutRow + 2).value = "DAMAGE"
        mySheet.Range("D" & tutRow + 3).value = "DAMAGE"
        mySheet.Range("D" & tutRow + 4).value = "DAMAGE"
        
        mySheet.Range("E" & tutRow + 2).value = "YELLOW"
        mySheet.Range("E" & tutRow + 3).value = "ORANGE"
        mySheet.Range("E" & tutRow + 4).value = "RED"
        
        mySheet.Range("D" & tutRow + 1).value = "DAMAGE"
        mySheet.Range("A" & tutRow + 5 & ":" & "Q" & tutRow + 9).value = mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 4).value
        mySheet.Range("A" & tutRow + 5).value = "CAL_BRIDGES_SCENARIO"
        mySheet.Range("A" & tutRow + 6).value = "CAL_BRIDGES_SCENARIO"
        mySheet.Range("A" & tutRow + 7).value = "CAL_BRIDGES_SCENARIO"
        mySheet.Range("A" & tutRow + 8).value = "CAL_BRIDGES_SCENARIO"
        mySheet.Range("A" & tutRow + 9).value = "CAL_BRIDGES_SCENARIO"
        Application.EnableEvents = True
    
        mySheet.Range("C" & tutRow + 5).value = mySheet.Range("C" & tutRow).value
    
        mySheet.Range("C" & tutRow + 5).Select
    
        TutForm.DialogueBox.Text = "CAL_BRIDGES_SCENARIO should now be yellow. The groups defined in this worksheet will oscillate colors to help you see which notification rows are associated with each group." & vbNewLine & vbNewLine & _
            "Right now, we have two groups with the exact same notification parameters. In order to make it so that the SCENARIO group only gets notifications when scenarios are run, we " & _
            "change the event type to ""SCENARIO"" for all the notification rows associated with that group."
    
    ElseIf tutInfo = 4 Then
    
        Application.EnableEvents = False
        ' already present values
        mySheet.Range("A" & tutRow).value = "CAL_BRIDGES"
        mySheet.Range("B" & tutRow).value = "BRIDGE"
        mySheet.Range("C" & tutRow).value = "43 -126;39 -126;34 -123;31 -118;31 -113;36 -113;39 -118;43 -118;43 -126"
        mySheet.Range("D" & tutRow).value = "NEW_EVENT"
        mySheet.Range("F" & tutRow).value = "3"
        mySheet.Range("G" & tutRow).value = "ACTUAL"
        mySheet.Range("H" & tutRow).value = "Rich Content"
        mySheet.Range("A" & tutRow + 1).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 2).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 3).value = "CAL_BRIDGES"
        mySheet.Range("A" & tutRow + 4).value = "CAL_BRIDGES"
        
        mySheet.Range("D" & tutRow + 2).value = "DAMAGE"
        mySheet.Range("D" & tutRow + 3).value = "DAMAGE"
        mySheet.Range("D" & tutRow + 4).value = "DAMAGE"
        
        mySheet.Range("E" & tutRow + 2).value = "YELLOW"
        mySheet.Range("E" & tutRow + 3).value = "ORANGE"
        mySheet.Range("E" & tutRow + 4).value = "RED"
        
        mySheet.Range("D" & tutRow + 1).value = "DAMAGE"
        mySheet.Range("A" & tutRow + 5 & ":" & "Q" & tutRow + 9).value = mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 4).value
        mySheet.Range("A" & tutRow + 5).value = "CAL_BRIDGES_SCENARIO"
        mySheet.Range("A" & tutRow + 6).value = "CAL_BRIDGES_SCENARIO"
        mySheet.Range("A" & tutRow + 7).value = "CAL_BRIDGES_SCENARIO"
        mySheet.Range("A" & tutRow + 8).value = "CAL_BRIDGES_SCENARIO"
        mySheet.Range("A" & tutRow + 9).value = "CAL_BRIDGES_SCENARIO"
        mySheet.Range("C" & tutRow + 5).value = mySheet.Range("C" & tutRow).value
        Application.EnableEvents = True
    
        mySheet.Range("G" & tutRow + 5).value = "SCENARIO"
        mySheet.Range("G" & tutRow + 6).value = "SCENARIO"
        mySheet.Range("G" & tutRow + 7).value = "SCENARIO"
        mySheet.Range("G" & tutRow + 8).value = "SCENARIO"
        mySheet.Range("G" & tutRow + 9).value = "SCENARIO"
        
        mySheet.Range("G" & tutRow + 5 & ":" & "G" & tutRow + 9).Select
        
        TutForm.DialogueBox.Text = "Now we have two notification groups that are identical in every way, except one group recieves alerts when actual eathquakes occur and one group recieves " & _
            "alerts when scenarios are run. These groups are also explicitly named so you can be sure that only specific users receive scenario alerts." & vbNewLine & vbNewLine & _
            "That's all for the Notification Worksheet, now let's move on to the User Worksheet!"
    
    End If
    
    

ElseIf tutNum = "3" And tutDec = "0" Then
    
    mySheet.Range("A4").Activate
    If tutInfo = 1 Then
        TutForm.DialogueBox.Text = "Welcome to the User Worksheet! You can use this worksheet to define ShakeCast users and link them to the notification groups " & _
            "defined in the Notification Worksheet. We'll go through two quick examples for this sheet; we'll create a general user first, then an admin user. In order to do this, we will need to make room in this worksheet again! " & vbNewLine & vbNewLine & _
            "Hit Continue to start!"
            
    End If

ElseIf tutNum = "3" And tutDec = "1" Then
    
    If tutInfo = 0 Then

        If copySheet.Range("A299").value = "no" Then
            copyRow = 300

            copySheet.Range("A" & copyRow & ":" & "Q" & copyRow + 1).value = _
                mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 1).value

            mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 1).Clear
            mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 1).Locked = False


            copySheet.Range("A299").value = "yes"
            
        Else
            mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 1).Clear
            mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 1).Locked = False
        End If
        
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 1).Locked = False
        
        mySheet.Range("A" & tutRow).Activate
        
        Application.GoTo mySheet.Range("A1"), True
        
        
        TutForm.DialogueBox.Text = "Alright, we've cleared some space so let's define a general user. The first field we will fill out is the username. A username can consist of numbers and letters " & _
            "and underscores, but should not have any spaces. For this example, we'll use ""SCuser""."
    ElseIf tutInfo = 1 Then
    
        mySheet.Range("B" & tutRow & ":" & "J" & tutRow).Clear
        mySheet.Range("A" & tutRow + 1 & ":" & "J" & tutRow + 1).Clear
        
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 1).Locked = False
        
        mySheet.Range("A" & tutRow).value = "SCuser"
        
        mySheet.Range("B" & tutRow).Activate
    
        TutForm.DialogueBox.Text = "Now we can select the user type USER from the drop down menu. A general user can get email alerts and log onto the web interface for ShakeCast. An ADMIN has the power to " & _
            "add other users, notification groups, and facilities to the system. They can also upload and trigger scenarios."
    ElseIf tutInfo = 2 Then
        
        mySheet.Range("C" & tutRow & ":" & "J" & tutRow).Clear
        mySheet.Range("A" & tutRow + 1 & ":" & "J" & tutRow + 1).Clear
        
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 1).Locked = False
        
        Application.EnableEvents = False
        ' already defined values
        mySheet.Range("A" & tutRow).value = "SCuser"
        Application.EnableEvents = True
        
        mySheet.Range("B" & tutRow).value = "USER"
        
        mySheet.Range("C" & tutRow & ":" & "D" & tutRow).Select
        
        TutForm.DialogueBox.Text = "Now, I'm just going to add a password for the account and the name of the account holder."
        
    ElseIf tutInfo = 3 Then
    
        mySheet.Range("E" & tutRow & ":" & "J" & tutRow).Clear
        mySheet.Range("A" & tutRow + 1 & ":" & "J" & tutRow + 1).Clear
        
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 1).Locked = False
        
        Application.EnableEvents = False
        ' already defined values
        mySheet.Range("A" & tutRow).value = "SCuser"
        mySheet.Range("B" & tutRow).value = "USER"
        Application.EnableEvents = True
        
        mySheet.Range("C" & tutRow).value = "pass"
        mySheet.Range("D" & tutRow).value = "John Smith"
        
        mySheet.Range("E" & tutRow).Activate
        
        TutForm.DialogueBox.Text = "Next I will enter an email address for this account. This email address is the one that the ShakeCast team will use to " & _
            "contact this user if necessary. You will enter an email address later in the form for alerts to be sent to. I'll go ahead and enter an email " & _
            "address and skip the Phone Number column because it's not required for regular users."
            
    ElseIf tutInfo = 4 Then
    
        mySheet.Range("F" & tutRow & ":" & "J" & tutRow).Clear
        mySheet.Range("A" & tutRow + 1 & ":" & "J" & tutRow + 1).Clear
        
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 1).Locked = False
        
        Application.EnableEvents = False
        ' already defined values
        mySheet.Range("A" & tutRow).value = "SCuser"
        mySheet.Range("B" & tutRow).value = "USER"
        mySheet.Range("C" & tutRow).value = "pass"
        mySheet.Range("D" & tutRow).value = "John Smith"
        Application.EnableEvents = True
        
        mySheet.Range("E" & tutRow).value = "jsmith@gmail.com"
        
        
        TutForm.DialogueBox.Text = "Another user form will pop up when you hit continue. This happens any time you select a cell in the Notification Group Column. " & _
            "This form allows you to select which notification groups you would like a specific user to be associated with. If you just went through the Notification " & _
            "Worksheet portion of this tutorial, you will be able to select from at least the CAL_BRIDGES and CAL_BRIDGES_SCENARIO group. You can click the check " & _
            "boxes by the name of the group and hit ""Okay"". I actually want you to try out this process when you hit continue!"
            
    ElseIf tutInfo = 5 Then
    
        mySheet.Range("G" & tutRow & ":" & "J" & tutRow).Clear
        mySheet.Range("A" & tutRow + 1 & ":" & "J" & tutRow + 1).Clear
        
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 1).Locked = False
        
        Application.EnableEvents = False
        ' already defined values
        mySheet.Range("A" & tutRow).value = "SCuser"
        mySheet.Range("B" & tutRow).value = "USER"
        mySheet.Range("C" & tutRow).value = "pass"
        mySheet.Range("D" & tutRow).value = "John Smith"
        mySheet.Range("E" & tutRow).value = "jsmith@gmail.com"
        Application.EnableEvents = True
        
        TutForm.DialogueBox.Text = "Now you can select which groups you would like this user to be a part of! When you are finished, go ahead and hit continue."
        
        mySheet.Range("G" & tutRow).Activate
        

    ElseIf tutInfo = 6 Then
        mySheet.Range("H" & tutRow & ":" & "J" & tutRow).Clear
        mySheet.Range("A" & tutRow + 1 & ":" & "J" & tutRow + 1).Clear
        
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 1).Locked = False
        
        Application.GoTo mySheet.Range("G1"), True
        
        mySheet.Range("H" & tutRow & ":" & "J" & tutRow).Select
        
        Application.EnableEvents = False
        ' already defined values
        mySheet.Range("A" & tutRow).value = "SCuser"
        mySheet.Range("B" & tutRow).value = "USER"
        mySheet.Range("C" & tutRow).value = "pass"
        mySheet.Range("D" & tutRow).value = "John Smith"
        mySheet.Range("E" & tutRow).value = "jsmith@gmail.com"
        Application.EnableEvents = True
        
        TutForm.DialogueBox.Text = "Nice! Now we have to enter more email addresses for this user. We have already defined the one that the ShakeCast team will use to " & _
            "make contact if necessary, but we still have to define where ShakeCast products should be delivered. For this case, we will send all three types of content to the " & _
            "same email address as the one for ShakeCast contact."
            
    ElseIf tutInfo = 7 Then
        mySheet.Range("A" & tutRow + 1 & ":" & "J" & tutRow + 1).Clear
        
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 1).Locked = False
        
        mySheet.Range("H" & tutRow).Select
        
        Application.EnableEvents = False
        ' already defined values
        mySheet.Range("A" & tutRow).value = "SCuser"
        mySheet.Range("B" & tutRow).value = "USER"
        mySheet.Range("C" & tutRow).value = "pass"
        mySheet.Range("D" & tutRow).value = "John Smith"
        mySheet.Range("E" & tutRow).value = "jsmith@gmail.com"
        Application.EnableEvents = True
        
        mySheet.Range("H" & tutRow).value = "jsmith@gmail.com"
        mySheet.Range("I" & tutRow).value = "jsmith@gmail.com"
        mySheet.Range("J" & tutRow).value = "jsmith@gmail.com"
        
        
        
        TutForm.DialogueBox.Text = "That's all the information we need to define a general user! Now let's give John admin priledges!"
    
    End If
    
ElseIf tutNum = "3" And tutDec = "2" Then
    
    If tutInfo = 0 Then
    
        Application.EnableEvents = False
        ' already defined values
        mySheet.Range("A" & tutRow).value = "SCuser"
        mySheet.Range("B" & tutRow).value = "USER"
        mySheet.Range("C" & tutRow).value = "pass"
        mySheet.Range("D" & tutRow).value = "John Smith"
        mySheet.Range("E" & tutRow).value = "jsmith@gmail.com"
        mySheet.Range("H" & tutRow).value = "jsmith@gmail.com"
        mySheet.Range("I" & tutRow).value = "jsmith@gmail.com"
        mySheet.Range("J" & tutRow).value = "jsmith@gmail.com"
        Application.EnableEvents = True
        
        Application.GoTo mySheet.Range("A1"), True
        mySheet.Range("B" & tutRow).Select
        TutForm.DialogueBox.Text = "The first thing we have to do is change his User Type to ADMIN."
    ElseIf tutInfo = 1 Then
    
        Application.EnableEvents = False
        ' already defined values
        mySheet.Range("A" & tutRow).value = "SCuser"
        mySheet.Range("C" & tutRow).value = "pass"
        mySheet.Range("D" & tutRow).value = "John Smith"
        mySheet.Range("E" & tutRow).value = "jsmith@gmail.com"
        mySheet.Range("H" & tutRow).value = "jsmith@gmail.com"
        mySheet.Range("I" & tutRow).value = "jsmith@gmail.com"
        mySheet.Range("J" & tutRow).value = "jsmith@gmail.com"
        Application.EnableEvents = True
        
        mySheet.Range("B" & tutRow).value = "ADMIN"
    
        TutForm.DialogueBox.Text = "As you can see, the user row is no longer complete! We have to enter a phone number in order " & _
            "to upload an ADMIN user to ShakeCast. We actually have a user form that pops up when you click on the Phone Number column." & _
            "This allows us to make sure that you've input a valid phone number."
        
        
    ElseIf tutInfo = 2 Then
    
        Application.EnableEvents = False
        ' already defined values
        mySheet.Range("A" & tutRow).value = "SCuser"
        mySheet.Range("C" & tutRow).value = "pass"
        mySheet.Range("D" & tutRow).value = "John Smith"
        mySheet.Range("E" & tutRow).value = "jsmith@gmail.com"
        mySheet.Range("H" & tutRow).value = "jsmith@gmail.com"
        mySheet.Range("I" & tutRow).value = "jsmith@gmail.com"
        mySheet.Range("J" & tutRow).value = "jsmith@gmail.com"
        mySheet.Range("B" & tutRow).value = "ADMIN"
        Application.EnableEvents = True


        TutForm.DialogueBox.Text = "Go ahead an enter a phone number into the form to check out how it works! When you're finished, hit continue."
        mySheet.Range("F" & tutRow).Select
        
        
    End If
    
Else
    TutForm.DialogueBox = "Congratulations, you've finished the ShakeCast Workbook tutorial! Now it's time to " & _
        "head back to the Facility XML worksheet and start inputting your information. " & vbNewLine & vbNewLine & _
        "When you are finished inputting your information, click on the options menu and export an XML document. This " & _
        "document can be uploaded to the ShakeCast application. You can also try exporting a master XML document. This " & _
        "will export all of the information from the entire workbook into a single XML file."

End If

Application.ScreenUpdating = True

End Sub
Sub tutSec()

Dim tutNum As Integer
Dim tutDec As Integer

tutNum = TutForm.SecNum.Caption
tutDec = TutForm.SecDec.Caption
tutInfo = TutForm.InfoClick.Caption

Dim tutRow As Integer
Dim copyRow As Integer
copyRow = 100

If tutNum < 1 Then
    tutNum = 1
    tutDec = 0
    tutInfo = 0
    
    Set mySheet = Worksheets("Facility XML")
    mySheet.Activate
    
ElseIf tutNum = 1 And tutDec < 3 Then
    tutDec = tutDec + 1
    tutInfo = 0
    

ElseIf tutNum = 1 Then
    tutNum = tutNum + 1
    tutDec = 0
    tutInfo = 0
    
    Set mySheet = Worksheets("Notification XML")
    mySheet.Activate
        
    
ElseIf tutNum = 2 And tutDec < 2 Then
    tutDec = tutDec + 1
    tutInfo = 0
    

    
ElseIf tutNum = 2 Then
    tutNum = 3
    tutDec = 0
    tutInfo = 0
    
    Set mySheet = Worksheets("User XML")
    mySheet.Activate
        
ElseIf tutNum = 3 And tutDec < 2 Then
    tutDec = tutDec + 1
ElseIf tutNum = 3 Then
    tutNum = 4
    tutDec = 0
    tutInfo = 0
End If

TutForm.SecNum.Caption = tutNum
TutForm.SecDec.Caption = tutDec
TutForm.InfoClick.Caption = tutInfo

tutCont

End Sub

Private Sub copyFirstRows()

Set copySheet = Worksheets("ShakeCast Ref Lookup Values")
Dim copyRow As Integer
Dim tutRow As Integer

If copySheet.Range("A99").value = "no" Then

    Set mySheet = Worksheets("Facility XML")

    copyRow = 100
    tutRow = 4
    
    copySheet.Range("A" & copyRow & ":" & "AF" & copyRow).value = _
        mySheet.Range("A" & tutRow & ":" & "AF" & tutRow).value

    mySheet.Range("A" & tutRow & ":" & "AF" & tutRow).Clear
    mySheet.Range("A" & tutRow & ":" & "AF" & tutRow).Locked = False
    
    Application.Run "CheckFacilities", mySheet.Range("A" & tutRow)
    
    copySheet.Range("A99").value = "yes"
    
End If

If copySheet.Range("A199").value = "no" Then

    Set mySheet = Worksheets("Notification XML")

    copyRow = 200
    tutRow = 4

    copySheet.Range("A" & copyRow & ":" & "Q" & copyRow + 9).value = _
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).value

    mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Clear
    mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 9).Locked = False


    copySheet.Range("A199").value = "yes"
    
End If

If copySheet.Range("A299").value = "no" Then

    Set mySheet = Worksheets("User XML")

    copyRow = 300
    tutRow = 4
    copySheet.Range("A" & copyRow & ":" & "Q" & copyRow + 1).value = _
        mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 1).value

    mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 1).Clear
    mySheet.Range("A" & tutRow & ":" & "Q" & tutRow + 1).Locked = False


    copySheet.Range("A299").value = "yes"
    
End If
End Sub
