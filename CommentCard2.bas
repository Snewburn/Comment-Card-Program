'**************************************************************************************************************************
'   Programmer: Spencer Newburn
'   File: CommentCardV2.bas
'   Date: 12/07/2015
'   Programming Language: JustBASIC
'   Description: This Programm runs in the lobby of a Museum and collects customer satisfaction information.
'                   The information is stored in two RANDOM files and is accessed by the sister program: CommentPortal.bas
'
'**************************************************************************************************************************


nomainwin

'These are the global Variables that the User will affect on screen 2 and the default settings of
' thee of them.  These contain the answers to the questions of the survey.
global  Service$, Cleanliness$, Recommend$, FavoritePart$, AddSomething$, WhereFrom$, HowdHear$ 

Service$ = "Not Answered"
Cleanliness$ = "Not Answered"
Recommend$ = "Not Answered"

' This start screen stays up perpetually and contains the logo and welcome message
call StartScreen

wait

'This is the first screen that welcomes the user and invites comment
sub StartScreen

    WindowWidth = 800                           'This block contains the parameters for the #start window
    WindowHeight = 380
    UpperLeftX = (DisplayWidth/2) - 400
    UpperLeftY = (DisplayHeight/2) - 190
    BackgroundColor$ = "black"
    ForegroundColor$ = "white"
    loadbmp "Barn", "image001.bmp"
    graphicbox #start.graph, -2, -4, 360, 360
    statictext #start.statictext3, "Please Take A Minute To Tell         Us What You Think...", 430, 195, 300, 60
    statictext #start.statictext4, "   Thank you for visiting Old Town!", 403, 20, 400, 140
    bmpbutton #start.button5, "start.bmp", SurveyScreen, UL, 525, 275

    open "" for window_nf as #start

    print #start.graph, "fill white; flush"        'This block draws the image and set the typface for the #start window
    print #start.graph, "drawbmp Barn"
    print #start.graph, "flush"
    print #start, "font candara_bold_italic 0 30"
    print #start.statictext4, "!font candara_bold_italic 0 60"

    print #start, "trapclose TRPCLSstart"            'trapclose to sub CloseStart just below
    wait

end sub
sub TRPCLSstart handle$      'trapclose sub for #start
    unloadbmp "Barn"
    close #start
    end
end sub


' this is the main screen that collects the User input for five questions
sub SurveyScreen handle$
    unloadbmp "Barn"
    close #start

    WindowWidth = 864
    WindowHeight = 660
    UpperLeftX = (DisplayWidth/2) - 432
    UpperLeftY = (DisplayHeight/2) - 330
    BackgroundColor$ = "black"
    ForegroundColor$ = "white"
    TextboxColor$ = "black"
    loadbmp "Banner", "OTCB.bmp"
   radiobutton #main.radiobutton7, "", Q1Excellent, nil, 60, 291, 16, 20    'Question 1 RadioButtons
   radiobutton #main.radiobutton8, "", Q1Good, nil, 200, 291, 16, 20
   radiobutton #main.radiobutton9, "", Q1Fair, nil, 300, 291, 16, 20
   statictext #main.RB7statictext, "Excellent", 76, 290, 70, 20             'Question 1 Radiobutton labels
   statictext #main.RB8statictext, "Good", 216, 290, 60, 20
   statictext #main.RB9statictext, "Fair", 316, 290, 60, 20
   radiobutton #main.radiobutton3, "", Q2Excellent, nil, 60, 366, 16, 20    'Question 2 Radionbuttons
   radiobutton #main.radiobutton4, "", Q2Good, nil, 200, 366, 16, 20
   radiobutton #main.radiobutton5, "", Q2Fair, nil, 300, 366, 16, 20
   statictext #main.RB3statictext, "Excellent", 76, 365, 70, 20             'Question 3 Radiobutton labels
   statictext #main.RB4statictext, "Good", 216, 365, 60, 20
   statictext #main.RB5statictext, "Fair", 316, 365, 60, 20
   radiobutton #main.radiobutton15, "", Q3YES, nil, 118, 575, 16, 20        'Qestion 3 radiobuttons
   radiobutton #main.radiobutton18, "", Q3NO, nil, 214, 575, 16, 20
   statictext #main.RB15statictext, "YES", 135, 574, 60, 20                 'Question 3 radiobutton labels
   statictext #main.RB18statictext, "NO", 230, 574, 60, 20
   statictext #main.statictext11, "What was your favorite part of the museum?", 20, 415, 390, 30   'Questions for textbox response
   statictext #main.statictext17, "Is there anything you would like to            see added to the museum?", 478, 260, 304, 60
   statictext #main.statictext19, "Tell Us...", 490, 400, 205, 30

    textbox #main.textboxFP, 21, 470, 384, 40   'Textboxes for textbox questions
    statictext #main.TypeResponse1, "Type response here:", 21, 450, 350, 20
    textbox #main.textboxADD, 438, 341, 384, 40
    statictext #main.TypeResponse2, "Type response here:", 438, 321, 350, 20
    textbox #main.textboxWF, 490, 450, 330, 30
    statictext #main.TypeResponse3, "Where you're from:", 490, 430, 350, 20
    textbox #main.textboxHH, 490, 510, 330, 30
    statictext #main.TypeResponse4, "How you heard about us:", 490, 490, 350, 20

    graphicbox #main.graph, -1, 0, 860, 251     'Graphicbox for the banner picture "Colorado's backyard"

    bmpbutton #main.finished, "done.bmp", SurveyComplete, UL, 570, 560   'This survey complete button triggers the storeage of the information

    groupbox #main.groupbox2, "", 14, 260, 395, 67  'These are the groupboxes that contain the radiobuttons  on the three radiobutton questions
    groupbox #main.groupbox6, "", 14, 335, 395, 67
    groupbox #main.groupbox14, "", 14, 545, 395, 67

    statictext #main.linecover1, "", 14, 320, 395, 10   'these are the groupbox line covers...
    statictext #main.linecover2, "", 11, 260, 40, 155   'JustBasic gives no way to rid oneself of the box lines, so these statictext controls
    statictext #main.linecover3, "", 370, 260, 50, 155  'cover those lines up.
    statictext #main.linecover4, "", 14, 400, 415, 10
    statictext #main.linecover5, "", 14, 545, 20, 70
    statictext #main.linecover6, "", 14, 610, 395, 10
    statictext #main.linecover7, "", 385, 545, 30, 70
    statictext #main.Gbox2blinder, "" , 14, 260, 395, 5 'Question 1,2,3 groupbox blinders
    statictext #main.Gbox6blinder, "", 14, 335, 395, 5
    statictext #main.Gbox14blinder, "", 14, 545, 395, 5

    statictext #main.Gbox2question, " Rate the quality of our customer service ", 30, 260, 360, 30  'Question 1,2,3 Questions text
    statictext #main.Gbox6question, " How was the cleanliness of the museum?", 30, 335, 360, 30
    statictext #main.Gbox14question, " Would you recommend Old Town to a friend?", 20, 545, 400, 30

    open "" for window_nf as #main
    print #main, "trapclose TRPCLSmain"

    print #main.graph, "fill black; flush"
    print #main.graph, "drawbmp Banner"
    print #main.graph, "flush"

    print #main, "font candara_bold_italic 0 22"

    print #main.Gbox2question, "!font candara_bold_italic 0 26"
    print #main.Gbox6question, "!font candara_bold_italic 0 26"
    print #main.Gbox14question, "!font candara_bold_italic 0 26"

    print #main.statictext11, "!font candara_bold_italic 0 26"
    print #main.statictext17, "!font candara_bold_italic 0 26"
    print #main.statictext19, "!font candara_bold_italic 0 26"

    print #main.textboxFP, "!font candara_italic 0 30"
    print #main.textboxWF, "!font candara_italic 0 20"
    print #main.textboxADD, "!font candara_italic 0 30"
    print #main.textboxHH, "!font candara_italic 0 20"
    print #main.TypeResponse1, "!font Courier_New 0 16"
    print #main.TypeResponse2, "!font Courier_New 0 16"
    print #main.TypeResponse3, "!font Courier_New 0 20"
    print #main.TypeResponse4, "!font Courier_New 0 20"

    timer 900000, TimedOutMain    'This will cylce the program back to the start screen after 15 minutes activity or none. Leaving
                                    'the respondant only 15 minutes to complete the 6 questions. Designed to get the program off
                                    ' Main screen when it is left there.

end sub
sub TRPCLSmain handle$
    unloadbmp "Banner"
    close #main
    end
end sub


sub SurveyComplete handle$

    print #main.textboxFP, "!contents? FavoritePart$"
    if (FavoritePart$ = "") then FavoritePart$ = "Not Answered"
    print #main.textboxADD, "!contents? AddSomething$"
    if (AddSomething$ = "") then AddSomething$ = "Not Answered"
    print #main.textboxWF, "!contents? WhereFrom$"
    if (WhereFrom$ = "") then WhereFrom$ = "Not Answered"
    print #main.textboxHH, "!contents? HowdHear$"
    if (HowdHear$ = "") then HowdHear$ = "Not Answered"

    if (FavoritePart$ = "Not Answered") and (AddSomething$ = "Not Answered") and (WhereFrom$ = "Not Answered")_             'This if statement keeps blank cards from
     and (Service$ = "Not Answered") and (Cleanliness$ = "Not Answered") and (Recommend$ = "Not Answered")_
     and (HowdHear$ = "Not Answered") then call FinalScreen : exit sub  'being saved in stats or records.

    call UpdateRecords  'in order for the files to be synchronized, both of these must be executed
    call UpdateStats    'in the same run of the program and the records.dat and the stats.dat files must
                        'be created at the same time.
    call FinalScreen

end sub

sub FinalScreen     'This is the final screen...program pauses on this screen for 5 seconds and this cycles back to screen 1

    unloadbmp "Banner"
    close #main

    WindowWidth = 455
    WindowHeight = 477
    UpperLeftX = (DisplayWidth/2) - 227
    UpperLeftY = (DisplayHeight/2) - 238

    loadbmp "Barn2", "picture1.bmp"

    graphicbox #final.graph, -1, -1, 450, 450

    open "" for window_nf as #final

    print #final.graph, "fill black; flush"
    print #final.graph, "drawbmp Barn2"
    print #final.graph, "flush"
    print #final, "font ms_sans_serif 0 16"

    print #final, "trapclose TRPCLSfinal"

    timer 5000, ResetToStart

end sub
sub TRPCLSfinal handle$
    unloadbmp "Barn2"
    timer 0
    close #final
    end
end sub

sub TimedOutMain
    unloadbmp, "Banner"
    close #main
    timer 0
    Service$ = "Not Answered"
    Cleanliness$ = "Not Answered"
    Recommend$ = "Not Answered"
    call StartScreen
end sub

sub ResetToStart
    timer 0
    unloadbmp "Barn2"
    close #final

    Service$ = "Not Answered"
    Cleanliness$ = "Not Answered"
    Recommend$ = "Not Answered"

    call StartScreen
end sub


' this is the logic for the Customer Servive question
sub Q1Excellent handle$
    Service$ = "Excellent"
end sub
sub Q1Good handle$
    Service$ = "Good"
end sub
sub Q1Fair handle$
    Service$ = "Fair"
end sub

' This is the logic for the Cleanliness question
sub Q2Excellent handle$
    Cleanliness$ = "Excellent"
end sub
sub Q2Good handle$
    Cleanliness$ = "Good"
end sub
sub Q2Fair handle$
    Cleanliness$ = "Fair"
end sub
'this is the logic for the recommend question
sub Q3YES handle$
    Recommend$ = "Yes"
end sub
sub Q3NO handle$
    Recommend$ = "No"
end sub



'*********************************************************************************
'*************************FUNCTIONS **********************************************
'*********************************************************************************


'FUNCTION DESCRIPTION: Returns the number of records in a RANDOM (.dat) file.
dim FieldLength(30), RecordField$(100)
function OpenRandomFile(fileName$, fieldLengths$)
    while word$(fieldLengths$, fieldCount+1) <> ""
        fieldCount = fieldCount+1
    wend
    redim RecordField$(fieldCount)
    redim FieldLength(fieldCount+1)
    FieldLength(0) = fieldCount
    for i = 1 to fieldCount
        FieldLength(i) = val(word$(fieldLengths$, i))
        size = size+FieldLength(i)
    next i
    FieldLength(fieldCount+1) = size
    open fileName$ for binary as #randomarray
    OpenRandomFile = lof(#randomarray)/size
    close #randomarray
end function


'FUNCTION DESCRIPTION: Accepts the path and name of a file and returns non-zero if the file exists.
dim info$(10,10)
function fileExists(path$, filename$)
  'dimension the array info$( at the beginning of your program
  files path$, filename$, info$()
  fileExists = val(info$(0, 0))  'non zero is true
end function

'FUNCTION DESCRIPTION: This sub maintains the Stats.dat file for the use of the pie chart creation.
sub UpdateStats
    if (fileExists(DefaultDir$, "Stats.dat")=0) then    'This block goes into effect on the first run or the first run aftter
        open "Stats.dat" for random as #Stats LEN = 80  'Data Nuke from CommentPortal. It creates a blank record 1 in the stats.dat
            field #Stats,_                              'file so that the get #Stats statement later on works.
            10 as Q1EStat, 10 as Q1GStat, 10 as Q1FStat,_
            10 as Q2EStat, 10 as Q2GStat, 10 as Q2FStat,_
            10 as Q3YStat, 10 as Q3NStat
        Q1EStat = 0
        Q1GStat = 0
        Q1FStat = 0
        Q2EStat = 0
        Q2GStat = 0
        Q2FStat = 0
        Q3YStat = 0
        Q3NStat = 0
        put #Stats, 1
        close #Stats
    end if

    open "Stats.dat" for random as #Stats LEN = 80
        field #Stats,_
            10 as Q1EStat, 10 as Q1GStat, 10 as Q1FStat,_
            10 as Q2EStat, 10 as Q2GStat, 10 as Q2FStat,_
            10 as Q3YStat, 10 as Q3NStat
    get #Stats, 1
        if (Service$ = "Excellent") then Q1EStat = Q1EStat + 1
        if (Service$ = "Good") then Q1GStat = Q1GStat + 1
        if (Service$ = "Fair") then Q1FStat = Q1FStat + 1
        if (Cleanliness$ = "Excellent") then Q2EStat = Q2EStat + 1
        if (Cleanliness$ = "Good") then Q2GStat = Q2GStat + 1
        if (Cleanliness$ = "Fair") then Q2FStat = Q2FStat + 1
        if (Recommend$ = "Yes") then Q3YStat = Q3YStat + 1
        if (Recommend$ = "No") then Q3NStat = Q3NStat + 1
    put #Stats, 1
    close #Stats
end sub

'FUNCTION DESCRIPTION: This sub maintains the Records.dat file
sub UpdateRecords
    if fileExists(DefaultDir$, "Records.dat") then      'This if-else logic sets the counter i to the proper index
        i = OpenRandomFile("Records.dat", "10 100 1000 1000 1000 1000 12 12 12") + 1
    else
        i = 1
    end if
    'The following is the code for saveing the entry in a .dat Random Access File
    open "Records.dat" for random as #Records LEN = 4146
        field #Records,_
            10 as index, 100 as date$, 1000 as favPart$, 1000 as addSome$,_
            1000 as whFrom$, 1000 as howHear$, 12 as serve$, 12 as clean$, 12 as recom$
        index = i
        date$ = date$()
        favPart$ = FavoritePart$
        addSome$ = AddSomething$
        whFrom$ = WhereFrom$
        howHear$ = HowdHear$
        serve$ = Service$
        clean$ = Cleanliness$
        recom$ = Recommend$
        put #Records, i
    close #Records
end sub
