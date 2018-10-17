'Programmer: Spencer Newburn
'File: CommentPortal1.bas
'Date: 12/32/2015
'Description:   This program interacts with the two files (Stats.dat and Records.dat) created by the CommentPortal2 program.
'               The three main functions are to create a master list of the data, to create a report, and to delete the data.

'Next:
'       Adjust the position of elements on #ComMan.   Viewers should be the approx length that is right for the report; move arrows down to report location.

nomainwin

dim WindowCheck(6)          'This array stores indicators that the four windows are open (1)#Master, (2)#ComMan, (3)#report, (4)#report2, (5)#CommentSpread, (6)#report3
dim checkvalue(4)           'This array stores whether one of the three Include buttons has been pressed.
dim FavoriteComments$(100)  'These arrays stores the comments that are included by the User for the report. They redim when #ComMan is opened.
dim AddComments$(100)
dim WhereFrom$(100)
dim HowHear$(100)

call MainScreen             'Brings up the main start screen for the application and holds for User input.
wait

'*************************************************MAIN SCREEN*********************************************************************************
'This is the main start screen for the application. Contains the three options.
'DATA NUKE (Delete records.dat and stat.dat), MASTER LIST (Look at all entries in text form), COMMENT MANAGER (create form).

sub MainScreen
    WindowWidth = 1000                                                              'This block contains the parameters for the #main window.
    WindowHeight = 500
    UpperLeftX = (DisplayWidth/2)-(WindowWidth/2)
    UpperLeftY = (DisplayHeight/2)-(WindowHeight/2)
    BackgroundColor$ = "Black"
    ForegroundColor$ = "White"
    bmpbutton #main.bmpbutton1,"Bomb.bmp", CommentNuke, UL, 100, 131
    bmpbutton #main.bmpbutton3, "MagGlass.bmp", MasterList, UL, 435, 131
    bmpbutton #main.bmpbutton2, "Checklist.bmp", CommentManager, UL, 712, 131
    bmpbutton #main.bmpbutton8, "Bars.bmp", BarGraph, UL, 945, 425
    statictext #main.statictext4, "Old Town Comment Portal", 150.5, 30, 699, 75
    statictext #main.statictext5, "NUKE      DATA", 130, 315, 250, 140
    statictext #main.statictext6, "MASTER      LIST", 450, 325, 137, 70
    statictext #main.statictext7, "Comment  Manager", 722, 320, 140, 100

    open "" for window_nf as #main                                                  'Open command for the main window. Handle #main
    print #main, "trapclose TRPCLSmain"

    print #main.statictext4, "!font rio_oro 0 75"                                   'This block sets the typeface for the main screen's three
    print #main.statictext5, "!font The_End_Font 0 50"                              'button titles and banner.
    print #main.statictext6, "!font DJB_Get_Digital 0 40"
    print #main.statictext7, "!font Candara_bold 0 40"

end sub
sub TRPCLSmain handle$                                                              'This catches the User click on the close button and calls
    call CloseAll                                                                   'the closing subroutine. One of these exists right after
end sub                                                                             'the subroutine for each opening window and redirects to the
                                                                                    'actual closing routine located after the Comment Nuke sub
'***************************************************************************************************************************************************

'******************************************************MASTER LIST**********************************************************************************
sub MasterList handle$ 
    if (fileExists(DefaultDir$, "Records.dat")=0) then call NoData : exit sub       'Check for file existence before opening.
    if (WindowCheck(1)=1) then call WindowOpen : exit sub                           'if window is open, notice and exit sub.

    x = OpenRandomFile("Records.dat", "10 100 1000 1000 1000 1000 12 12 12")             'Assign counter variables 'x' (number of records)
    i = 1                                                                           'and i (index)
        open "Records.dat" for random as #Records LEN = 4146                        'The Records.dat is opened in order to print its contents into
            field #Records,_                                                        '#master. The information is looped through the number of records 'x'
                10 as index, 100 as date$, 1000 as favPart$, 1000 as addSome$,_
                1000 as whFrom$, 1000 as howHear$, 12 as serve$, 12 as clean$, 12 as recom$
            open "" for text as #master                                             'Open #master window
            print #master, "!trapclose TRPCLSmaster"
            call MasterIsOpen
                while (i<=x)
                    get #Records, i
                    print #master, index
                    print #master, date$
                    print #master, "'Favorite part?'                    "; favPart$ 
                    print #master, "'What would you add to Museum?'     "; addSome$ 
                    print #master, "'Where you're from:'                "; whFrom$
                    print #master, "'How did you hear about us?'        "; howHear$
                    print #master, "'Service?'                          "; serve$
                    print #master, "'Cleanliness?'                      "; clean$
                    print #master, "'Would you Recommend?'              "; recom$
                    print #master, ""
                    i=i+1
                wend
        close #Records
end sub
sub TRPCLSmaster handle$
    call CloseMaster
end sub
'***********************************************************************************************************************************************************************

'***********************************************COMMENT MANAGER*********************************************************************************************************
sub CommentManager handle$
    if (WindowCheck(2)=1) then call WindowOpen : exit sub                           'If the window is already open, displaay notice, and prevent opening.
    call ComManIsOpen                                                               'Set the windowcheck array slot as 1 for open.

    redim FavoriteComments$(100)                                                    'reset the arrays that are being used to select the comments for the report
    redim AddComments$(100)
    redim WhereFrom$(100)
    redim HowHear$(100)
    redim checkvalue(4)

    WindowWidth = 1220                                                              'This block contains the control parameters for the #ComMan window.
    WindowHeight = 720
    UpperLeftX = (DisplayWidth/2)-(WindowWidth/2)
    UpperLeftY = ((DisplayHeight/2)-(WindowHeight/2))/4
    BackgroundColor$ = "Black"
    ForegroundColor$ = "White"
    TextboxColor$ = "Black"
    TexteditorColor$ = "Black"
    loadbmp "Include", "Include.bmp"
    loadbmp "Check", "check.bmp"
    loadbmp "Blank", "blank.bmp"
    texteditor #ComMan.textedit1, 30, 61, 670, 140    'Displays what was your favorite part of the musem? response
    texteditor #ComMan.textedit2, 30, 276, 670, 140    'Displays is there anything you would like to see added? reponse
    texteditor #ComMan.textedit3, 197.5, 480, 335, 60     'Displays Where are you from? reponse
    texteditor #ComMan.textedit4, 197.5, 585, 335, 60     'Displays howd you hear about us? response
    bmpbutton #ComMan.bmpbutton5, "arrowright.bmp", ArrowFWD, UL, 958, 506
    bmpbutton #ComMan.bmpbutton9, "arrowleft.bmp", ArrowBCK, UL, 918, 506
    bmpbutton #ComMan.bmpbutton22, "Report.bmp", CreateReport, UL, 1065, 540
    bmpbutton #ComMan.bmpbutton19, "blank.bmp", IncludeFAV, UL, 710, 110
    bmpbutton #ComMan.bmpbutton20, "blank.bmp", IncludeADD, UL, 710, 321
    bmpbutton #ComMan.bmpbutton21, "blank.bmp", IncludeWHFROM, UL, 710, 500
    bmpbutton #ComMan.bmpbutton30, "blank.bmp", IncludeHOWHEAR, UL, 710, 595
    statictext #ComMan.statictext11, "What was your favorite part of the museum?", 30, 20, 500, 30
    statictext #ComMan.statictext12, "Is there anything you would like to see", 30, 215, 450, 25
    statictext #ComMan.statictext27, "added to the museum?", 30, 240, 450, 25
    statictext #ComMan.statictext13, "Where are you from?", 197.5, 445, 300, 30
    statictext #ComMan.statictext29, "How did you hear about us?", 197.5, 555, 300, 30
    statictext #ComMan.statictext26, "Comment      Manager", 920, 20, 400, 130
    statictext #ComMan.statictext23, "NAVIGATE     through the comments with the red and green arrows", 920, 130, 180, 120
    statictext #ComMan.statictext24, "ADD          comments to the reports with the 'Include' buttons", 920, 270, 180, 100
    statictext #ComMan.statictext25, "WHEN FINISHED click Create Report to see report and print", 920, 390, 180, 100
    textbox #ComMan.textbox15, 920.5, 542, 65, 60
    textbox #ComMan.textbox28, 903, 610, 100, 30

    open "Comment Manager" for window as #ComMan                                    'Open Comment Manager window.
    print #ComMan, "trapclose TRPCLScomman"

    print #ComMan, "font Candara_bold"                                              'Set the typefaces for the window.
    print #ComMan.statictext23, "!font Kronika 0 25"
    print #ComMan.statictext24, "!font Kronika 0 25"
    print #ComMan.statictext25, "!font Kronika 0 25"
    print #ComMan.statictext26, "!font rio_oro 0 50"
    print #ComMan.statictext11, "!font kronika 0 30"
    print #ComMan.statictext12, "!font kronika 0 30"
    print #ComMan.statictext13, "!font kronika 0 30"
    print #ComMan.statictext27, "!font kronika 0 30"
    print #ComMan.statictext29, "!font kronika 0 30"
    print #ComMan.textbox15, "!font 0 50"

    print #ComMan.bmpbutton19, "Disable"                                            'The three Include buttons are disabled and blank for the start of the window
    print #ComMan.bmpbutton20, "Disable"                                            'they come alive when you arrow over to the comments.
    print #ComMan.bmpbutton21, "Disable"

    print #ComMan.textbox15, 0                                                      'Comment counter display is set to 0
end sub
sub TRPCLScomman handle$                                                            'Catch the close button and call the actual close routine for this window.
    call CloseComMan
end sub
'***********************************************************************************************************************************************************************

'******************************************************ARROW FORWARD ***************************************************************************************************
sub ArrowFWD handle$                                                                'This subroutine brings up the next answers for each of the
    #ComMan.textbox15, "!contents? i"                                               'Three fields in the comment manager.
    x = OpenRandomFile("Records.dat", "10 100 1000 1000 1000 1000 12 12 12")             'The current card and the total number of cards is pulled

    if (i>x) then i = x - 1                                                         'these two statements handle User input into the index textbox that is outside the
    if (i<0) then notice "You cannot enter a negative number" : exit sub            'range of the records.dat file. Either higher than number of records or lower than 0

    i = i+1
    if (i = x+1) then notice "That is the END of the file" : exit sub               'If the button has been pressed and you are on in the last card already, notice is

    if (i>0) then                                                                   'given.
        print #ComMan.bmpbutton19, "Enable"                                         'Once the User arrows right above 0, the Include buttons come alive.
        print #ComMan.bmpbutton20, "Enable"
        print #ComMan.bmpbutton21, "Enable"
        print #ComMan.bmpbutton30, "Enable"
    end if

    if FavoriteComments$(i) = "" then                                               'These if else statements set the button and the checkvalue() values
        print #ComMan.bmpbutton19, "bitmap Include"                                 'depending on if the comment has already been included.
        checkvalue(1) = 0                                                           'if that slot in the comment array is empty, the include button shows "Include"
    else                                                                            'and sets that buttons checkvalue slot (example: FavoriteComments$() is checkvalue(1))
        print #ComMan.bmpbutton19, "bitmap Check"                                   'to the appropriate 1 or 2.
        checkvalue(1) = 1
    end if

    if AddComments$(i) = "" then                                                    'The checkvalue() slots are used by the Include button routines to determine what
        print #ComMan.bmpbutton20, "bitmap Include"                                 'action is taken when the button is pressed.
        checkvalue(2) = 0
    else
        print #ComMan.bmpbutton20, "bitmap Check"
        checkvalue(2) = 1
    end if

    if WhereFrom$(i) = "" then
        print #ComMan.bmpbutton21, "bitmap Include"
        checkvalue(3) = 0
    else
        print #ComMan.bmpbutton21, "bitmap Check"
        checkvalue(3) = 1
    end if

    if HowHear$(i) = "" then
        print #ComMan.bmpbutton30, "bitmap Include"
        checkvalue(4) = 0
    else
        print #ComMan.bmpbutton30, "bitmap Check"
        checkvalue(4) = 1
    end if

    print #ComMan.textedit1, "!cls"                                                 'The text editor controls are cleared in preparation for the printing of the next
    print #ComMan.textedit2, "!cls"                                                 'card.
    print #ComMan.textedit3, "!cls"
    print #ComMan.textedit4, "!cls"

    open "Records.dat" for random as #Records LEN = 4146                            'The file records.dat is opened and the next card's information is printed according
            field #Records,_                                                        'to the value previously given to 'i'.
                10 as index, 100 as date$, 1000 as favPart$, 1000 as addSome$,_
                1000 as whFrom$, 1000 as howHear$, 12 as serve$, 12 as clean$, 12 as recom$
            get #Records, i
            print #ComMan.textedit1, favPart$
            print #ComMan.textedit2, addSome$
            print #ComMan.textedit3, whFrom$
            print #ComMan.textedit4, howHear$
            print #ComMan.textbox28, date$
            print #ComMan.textbox15, i
    close #Records
end sub
'***********************************************************************************************************************************************************************

'******************************************************ARROW BACK*******************************************************************************************************
sub ArrowBCK handle$                                                                'This sub displays the next lowest card in the order.
    #ComMan.textbox15, "!contents? i"                                               'The current card's index number is pulled and assigned to 'i'.
    x = OpenRandomFile("Records.dat", "10 100 1000 1000 1000 1000 12 12 12")

    if (i = 0) then exit sub                                                        'If the User is already on '0' the sub is exited preventing furthur action.

    i = i - 1                                                                       ''i' is decreased by 1.
    if (i<0) then notice "Cannot go below 0" : exit sub                             'Prevents User from entering a negative number
    if (i>x) then i = x

    if (i=0) then
        print #ComMan.bmpbutton19, "bitmap Blank"                                   'In the event that the new 'i' is 0, the buttons are deactivated.
        print #ComMan.bmpbutton20, "bitmap Blank"
        print #ComMan.bmpbutton21, "bitmap Blank"
        print #ComMan.bmpbutton30, "bitmap Blank"
        print #ComMan.bmpbutton19, "Disable"
        print #ComMan.bmpbutton20, "Disable"
        print #ComMan.bmpbutton21, "Disable"
        print #ComMan.bmpbutton30, "Disable"
    else                                                                            'If the new 'i' is greater than zero, the four comment arrays arrays are checked
        if FavoriteComments$(i) = "" then                                           'for content.
            print #ComMan.bmpbutton19, "bitmap Include"
            checkvalue(1) = 0
        else
            print #ComMan.bmpbutton19, "bitmap Check"                               'if the xxxxComents$() has something in it (The User has put it there), then the
            checkvalue(1) = 1                                                       'check mark image is displayed and the checkvalue() is set to 1.  This is the 'else'
        end if

        if AddComments$(i) = "" then
            print #ComMan.bmpbutton20, "bitmap Include"
            checkvalue(2) = 0
        else                                                                        'The checkvalue() slots are used by the inlude button subs to determine what action
            print #ComMan.bmpbutton20, "bitmap Check"                               'include button takes.
            checkvalue(2) = 1
        end if

        if WhereFrom$(i) = "" then
            print #ComMan.bmpbutton21, "bitmap Include"
            checkvalue(3) = 0
        else
            print #ComMan.bmpbutton21, "bitmap Check"
            checkvalue(3) = 1
        end if

        if HowHear$(i) = "" then
            print #ComMan.bmpbutton30, "bitmap Include"
            checkvalue(4) = 0
        else
            print #ComMan.bmpbutton30, "bitmap Check"
            checkvalue(4) = 1
        end if
    end if
                                                                                    'The three texteditor controls are cleared in preparation for the printing of the
    print #ComMan.textedit1, "!cls"                                                 'new cards information
    print #ComMan.textedit2, "!cls"
    print #ComMan.textedit3, "!cls"
    print #ComMan.textedit4, "!cls"

    if (i = 0) then print #ComMan.textbox15, 0 : print #ComMan.textbox28, "" : exit sub 'if the new 'i' value is 0, 0 in printed to index textbox, the date textbox is
                                                                                        'cleared, and the sub is exited.
    open "Records.dat" for random as #Records LEN = 4146                                'The records.dat file is opened that the information for card 'i' is printed to
            field #Records,_                                                            'texteditor and texbox control.
                10 as index, 100 as date$, 1000 as favPart$, 1000 as addSome$,_
                1000 as whFrom$, 1000 as howHear$, 12 as serve$, 12 as clean$, 12 as recom$
            get #Records, i
            print #ComMan.textedit1, favPart$
            print #ComMan.textedit2, addSome$
            print #ComMan.textedit3, whFrom$
            print #ComMan.textedit4, howHear$
            print #ComMan.textbox28, date$
            print #ComMan.textbox15, i
    close #Records                                                                      'The records.dat file is closed.
end sub
'***********************************************************************************************************************************************************************

'******************************************************INCLUDE BUTTONS**************************************************************************************************
sub IncludeFAV handle$                                  'These subroutines capture the contents of the texteditor contols in #ComMan
    if (checkvalue(1) = 0) then
        checkvalue(1) = 1
        print #ComMan.bmpbutton19, "bitmap Check"
        print #ComMan.textedit1, "!lines x"
        i=1
        while (i<=x)                                    'This while loope in each routine here contructs the entry based on what is in the
            print #ComMan.textedit1, "!line "; i;" q$"  'Text Editors.
            if q$ = "" then exit while
            sentence$ = sentence$ + "\" + q$
            i=i+1
        wend
        print #ComMan.textbox15, "!contents? x"         'The comment is placed is the array spot matching its index number.
        FavoriteComments$(x) = sentence$
        exit sub
    end if
    if (checkvalue(1) = 1) then                         'This is a click to take away the comment from the report.
        checkvalue(1) = 0
        print #ComMan.bmpbutton19, "bitmap Include"
        print #ComMan.textbox15, "!contents? x"
        FavoriteComments$(x) = ""
    end if
end sub

sub IncludeADD handle$
    if (checkvalue(2) = 0) then
        checkvalue(2) = 1
        print #ComMan.bmpbutton20, "bitmap Check"
        print #ComMan.textedit2, "!lines x"
        i=1
        while (i<=x)
            print #ComMan.textedit2, "!line "; i;" q$"
            if q$ = "" then exit while
            sentence$ = sentence$ + "\" + q$
            i=i+1
        wend
        print #ComMan.textbox15, "!contents? x"
        AddComments$(x) = sentence$
        exit sub
    end if
    if (checkvalue(2) = 1) then
        checkvalue(2) = 0
        print #ComMan.bmpbutton20, "bitmap Include"
        print #ComMan.textbox15, "!contents? x"
        AddComments$(x) = ""
    end if
end sub

sub IncludeWHFROM handle$
    if (checkvalue(3) = 0) then
        checkvalue(3) = 1
        print #ComMan.bmpbutton21, "bitmap Check"
        print #ComMan.textedit3, "!lines x"
        i=1
        while (i<=x)
            print #ComMan.textedit3, "!line "; i;" q$"
            if q$ = "" then exit while
            sentence$ = sentence$ + "\" + q$
            i=i+1
        wend
        print #ComMan.textbox15, "!contents? x"
        WhereFrom$(x) = sentence$
        exit sub
    end if
    if (checkvalue(3) = 1) then
        checkvalue(3) = 0
        print #ComMan.bmpbutton21, "bitmap Include"
        print #ComMan.textbox15, "!contents? x"
        WhereFrom$(x) = ""
    end if
end sub

sub IncludeHOWHEAR handle$
    if (checkvalue(4) = 0) then
        checkvalue(4) = 1
        print #ComMan.bmpbutton30, "bitmap Check"
        print #ComMan.textedit4, "!lines x"
        i=1
        while (i<=x)
            print #ComMan.textedit4, "!line "; i;" q$"
            if q$ = "" then exit while
            sentence$ = sentence$ + "\" + q$
            i=i+1
        wend
        print #ComMan.textbox15, "!contents? x"
        HowHear$(x) = sentence$
        exit sub
    end if
    if (checkvalue(4) = 1) then
        checkvalue(4) = 0
        print #ComMan.bmpbutton30, "bitmap Include"
        print #ComMan.textbox15, "!contents? x"
        HowHear$(x) = ""
    end if
end sub
'************************************************************************************************************************************
'******************************************************REPORT************************************************************************
sub CreateReport handle$

    if (fileExists(DefaultDir$, "Records.dat")=0) or (fileExists(DefaultDir$, "Stats.dat")=0) then call NoData : exit sub
    if (WindowCheck(3)=1) or (WindowCheck(4)=1) or (WindowCheck(6)=1) then call WindowOpen : exit sub

    x = OpenRandomFile("Records.dat", "10 100 1000 1000 1000 1000 12 12 12")             'The records.dat file is opened to glean date
    open "Records.dat" for random as #Records LEN = 4146                            'for the first and last entries.
            field #Records,_
                10 as index, 100 as date$, 1000 as favPart$, 1000 as addSome$,_
                1000 as whFrom$, 1000 as howHear$, 12 as serve$, 12 as clean$, 12 as recom$
            get #Records, 1
                startDate$ = date$
            get #Records, x
                endDate$ = date$
    close #Records
    '----------------------------------------------------------------------DRAW STATS REPORT AS #report-------------------------------------------------
    call ReportIsOpen
    WindowWidth = 900                                                               'These are the window specs for #report
    WindowHeight = 700
    UpperLeftX = (DisplayWidth/2)-(WindowWidth/2)
    UpperLeftY = ((DisplayHeight/2)-(WindowHeight/2))/4
    menu #report, "&File", "&Create PDF", CreatePDF1

    open "Stats Report" for graphics as #report
    print #report, "trapclose TRPCLSreport"
                                                               'set the window open indicator to open
    #report "down"                                                              'the stats report is drawn
    #report "size 3"
    #report "backcolor black"
    #report "place 2 2; boxfilled 798 110"

    #report "place 180 60; down"
    #report "color white"
    #report "font rio_oro 0 55"
    #report "\Old Town Stats Report"

    #report "place 30 93"
    #report "font candara_bold 0 25"
    #report "\"; startDate$
    #report "place 170 93"
    #report "\"; endDate$ 
    #report "place 135 96"
    #report "font candara_italic 0 30"
    #report "\to"

    #report "place 400 96"
    #report "font candara_italic 0 30"
    #report "\Total number of comment cards:"
    #report "place 750 93; down"
    #report "font candara_bold 0 25"
    print #report, "\"; x
    #report "color black"

    #report "font Arial_bold 0 17"
    #report "size 1"
    #report "place 75 200"
    call drawCircle "Service"               'Custom pie charts funtion is used.
    #report "place 365 200"
    call drawCircle "Cleanliness"
    #report "place 650 200"
    call drawCircle "Recommend"

    #report "place 35 300"
    #report "backcolor white"                                              'these are the three radio button questions and the Where you from/howd you hear questions.
    print #report, "\Rate the quality of our\    customer service"
    #report "place 315 300"
    print #report, "\How was the cleanliness\      of the museum?"
    #report "place 595 300"
    print #report, "\Would you recommend\ Old Town to a friend?"
    #report "place 100 365"
    #report "font candara_italic 0 25"
    print #report, "\Where are you from?"
    #report "place 480 365"
    #report "font candara_italic 0 25"
    print #report, "\How did you hear about us?"

    #report "place 40 400"                  'This loop prints the User-selected responses stored in WhereFrom$()
    #report "font Arial_bold 0 17"

    FOR i=1 TO 100    '100 is the LEN of of the array                              'This FOR loop is copied from a justbasic tutorial by Welopaz on JB website
        FOR j=1 TO 99 '99 is one less than the LEN of array                              'the role is sorting the Array alphabetically
            IF WhereFrom$(j)>WhereFrom$(j+1) THEN 'Statement evaluates if the current slot is greater than the the next
                temp$=WhereFrom$(j)
                WhereFrom$(j)=WhereFrom$(j+1)
                WhereFrom$(j+1)=temp$
            END IF
        NEXT j
    NEXT i

    i=1                                 'print the WhereFrom$() to the report.
    while (i<=100)
        print #report, WhereFrom$(i)
        i = i +1
    wend

    #report "place 440 400"
    FOR i=1 TO 100                                  'This FOR loop is copied from a justbasic tutorial by Welopaz on JB website
        FOR j=1 TO 99                               'the role is sorting the array alphabetically
            IF HowHear$(j)>HowHear$(j+1) THEN
                temp$=HowHear$(j)
                HowHear$(j)=HowHear$(j+1)
                HowHear$(j+1)=temp$
            END IF
        NEXT j
    NEXT i

    i=1                                 'print the HowHear$() to the report.
    while (i<=100)
        print #report, HowHear$(i)
        i = i +1
    wend

    #report "size 2"
    #report "line 40 340 750 340"   'Horizontal line below the pie charts
    #report "line 400 360 400 1000" 'Vertical line down the middle
    #report "size 3"
    #report "line 2 110 2 1033"     'verticle line on left side
    #report "line 2 1033 797 1033"  'horizontal line at bottom
    #report "line 797 110 797 1033" 'verticle line on right side


    #report "flush"

    print #report, "vertscrollbar on 0 400"
    print #report, "horizscrollbar off"
    '-----------------------------------------------------------------END OF STATS REPORT----------------------------------------
    '--------------------------------START OF FAVORITES REPORT-------------------------------------------------------------------

    call Report2IsOpen
    WindowWidth = 900                                                               'These are the window specs for #report
    WindowHeight = 700
    UpperLeftX = (DisplayWidth/2)-(WindowWidth/2)
    UpperLeftY = (((DisplayHeight/2)-(WindowHeight/2))/4)+35
    menu #report2, "&File", "&Create PDF", CreatePDF2

    open "Favorite Part Report" for graphics as #report2
    print #report2, "trapclose TRPCLSreport2"
    #report2 "down"                                                              'the stats report is drawn
    #report2 "size 3"
    #report2 "backcolor black"
    #report2 "place 2 2; boxfilled 798 125"

    #report2 "place 300 40"         'Placing and draing of "Old Town Report"
    #report2 "color white"          'and Report question
    #report2 "font rio_oro 0 35"
    #report2 "\Old Town Report"
    #report2 "place 100 70"
    #report2 "\What was your favorite part of the museum?"

    #report2 "place 30 108"             'Draw the range of dates the report covers
    #report2 "font candara_bold 0 25"
    #report2 "\"; startDate$
    #report2 "place 170 108"
    #report2 "\"; endDate$ 
    #report2 "place 135 111"
    #report2 "font candara_italic 0 30"
    #report2 "\to"

    #report2 "place 400 111"            'Draw the total cards sentence.
    #report2 "font candara_italic 0 30"
    #report2 "\Total number of comment cards:"
    #report2 "place 750 108; down"
    #report2 "font candara_bold 0 25"
    print #report2, "\"; x
    #report2 "color black"

    #report2 "place 50 175"         'Draw the comments that were selected by the User
    #report2 "backcolor white"
    #report2 "font candara_bold 0 20"
    i=1
    while (i<=x)
        if FavoriteComments$(i) = "" then
            i = i+1
        else
            print #report2, FavoriteComments$(i)
            #report2 "up; north; turn 180; go 15; down"
            i = i+1
        end if
    wend

    #report2 "size 3"
    #report2 "line 2 110 2 1033"     'verticle line on left side
    #report2 "line 2 1033 797 1033"  'horizontal line at bottom
    #report2 "line 797 110 797 1033" 'verticle line on right side

    #report2 "flush"
    print #report2, "vertscrollbar on 0 400"
    print #report2, "horizscrollbar off"

    '-----------------------------------------------------------END FAVORITES REPORT-----------------------------------------
    '------------------------------START THINGS TO ADD REPORT----------------------------------------------------------------

    call Report3IsOpen
    WindowWidth = 900                                                               'These are the window specs for #report
    WindowHeight = 700
    UpperLeftX = (DisplayWidth/2)-(WindowWidth/2)
    UpperLeftY = (((DisplayHeight/2)-(WindowHeight/2))/4)+80
    menu #report3, "&File", "&Create PDF", CreatePDF3

    open "Add Anything Report" for graphics as #report3
    print #report3, "trapclose TRPCLSreport3"
    #report3 "down"                                                              'the stats report is drawn
    #report3 "size 3"
    #report3 "backcolor black"
    #report3 "place 2 2; boxfilled 798 125"

    #report3 "place 300 40"         'Placing and draing of "Old Town Report"
    #report3 "color white"          'and Report question
    #report3 "font rio_oro 0 33"
    #report3 "\Old Town Report"
    #report3 "place 30 70"
    #report3 "\Is there anything you would like to see added to the museum?"

    #report3 "place 30 108"             'Draw the range of dates the report covers
    #report3 "font candara_bold 0 25"
    #report3 "\"; startDate$
    #report3 "place 170 108"
    #report3 "\"; endDate$ 
    #report3 "place 135 111"
    #report3 "font candara_italic 0 30"
    #report3 "\to"

    #report3 "place 400 111"            'Draw the total cards sentence.
    #report3 "font candara_italic 0 30"
    #report3 "\Total number of comment cards:"
    #report3 "place 750 108; down"
    #report3 "font candara_bold 0 25"
    print #report3, "\"; x
    #report3 "color black"

    #report3 "place 50 175"         'Draw the comments that were selected by the User
    #report3 "backcolor white"
    #report3 "font candara_bold 0 20"
    i=1
    while (i<=x)
        if AddComments$(i) = "" then
            i = i+1
        else
            print #report3, AddComments$(i)
            #report3, "up; north; turn 180; go 10; down"
            i = i+1
        end if
    wend

    #report3 "size 3"
    #report3 "line 2 110 2 1033"     'verticle line on left side
    #report3 "line 2 1033 797 1033"  'horizontal line at bottom
    #report3 "line 797 110 797 1033" 'verticle line on right side

    #report3 "flush"
    print #report3, "vertscrollbar on 0 400"
    print #report3, "horizscrollbar off"

end sub

sub TRPCLSreport handle$
    call CloseReport
end sub
sub TRPCLSreport2 handle$
    call CloseReport2
end sub
sub TRPCLSreport3 handle$
    call CloseReport3
end sub

sub CreatePDF1
    #report "print svga"
end sub
sub CreatePDF2
    #report2 "print svga"
end sub
sub CreatePDF3
    #report3 "print svga"
end sub


'************************************************************************************************************************************

'*****************************************************COMMENT NUKE******************************************************************
sub CommentNuke handle$
    if (fileExists(DefaultDir$, "Records.dat")=0) and (fileExists(DefaultDir$, "Stats.dat")=0)then call NoData : exit sub

    playwave "CarHorn.wav", async
    confirm "WARNING - Erasing data will permanenty delete the responses collected by the program. Are you sure you want to delete all data?"; answer$
        if (answer$ = "no") then exit sub

    if (fileExists(DefaultDir$, "Records.dat")) then kill "Records.dat"
    if (fileExists(DefaultDir$, "Stats.dat")) then kill "Stats.dat"

    playwave "explosion.wav", async
    notice "Data has been erased"
end sub
'************************************************************************************************************************************

'******************************************************BAR GRAPH*********************************************************************
sub BarGraph handle$
    if (fileExists(DefaultDir$, "Records.dat")=0) then call NoData : exit sub
    if (WindowCheck(5)=1) then call WindowOpen : exit sub

    x = OpenRandomFile("Records.dat", "10 100 1000 1000 1000 1000 12 12 12")
    z=1

    open "Records.dat" for random as #Records LEN = 4146
        field #Records,_
                10 as index, 100 as date$, 1000 as favPart$, 1000 as addSome$,_
                1000 as whFrom$, 1000 as howHear$, 12 as serve$, 12 as clean$, 12 as recom$
        while (z<=x)
            get #Records, z
                if (left$(date$, 3) = "Jan") then aaa = aaa+1
                if (left$(date$, 3) = "Feb") then bbb = bbb+1
                if (left$(date$, 3) = "Mar") then ccc = ccc+1
                if (left$(date$, 3) = "Apr") then ddd = ddd+1
                if (left$(date$, 3) = "May") then eee = eee+1
                if (left$(date$, 3) = "Jun") then fff = fff+1
                if (left$(date$, 3) = "Jul") then ggg = ggg+1
                if (left$(date$, 3) = "Aug") then hhh = hhh+1
                if (left$(date$, 3) = "Sep") then iii = iii+1
                if (left$(date$, 3) = "Oct") then jjj = jjj+1
                if (left$(date$, 3) = "Nov") then kkk = kkk+1
                if (left$(date$, 3) = "Dec") then lll = lll+1
            z = z+1
        wend
    close #Records

    aa = int((aaa/x)*100)
    bb = int((bbb/x)*100)
    cc = int((ccc/x)*100)
    dd = int((ddd/x)*100)
    ee = int((eee/x)*100)
    ff = int((fff/x)*100)
    gg = int((ggg/x)*100)
    hh = int((hhh/x)*100)
    ii = int((iii/x)*100)
    jj = int((jjj/x)*100)
    kk = int((kkk/x)*100)
    ll = int((lll/x)*100)

    jan = 350 - (2.8*aa)                        ' 'a-l' is the input for the height of the bars.
    feb = 350 - (2.8*bb)
    mar = 350 - (2.8*cc)
    apr = 350 - (2.8*dd)
    may = 350 - (2.8*ee)
    jun = 350 - (2.8*ff)
    jul = 350 - (2.8*gg)
    aug = 350 - (2.8*hh)
    sep = 350 - (2.8*ii)
    oct = 350 - (2.8*jj)
    nov = 350 - (2.8*kk)
    dec = 350 - (2.8*ll)

    WindowWidth = 558
    WindowHeight = 460
    open "Comment Spread" for graphics_nsb_nf as #CommentSpread
    print #CommentSpread, "trapclose TRPCLScommentspread"
    call BarGraphIsOpen


    #CommentSpread, "down; color white"                         'Draw the background using a white lined black box.
    #CommentSpread, "backcolor black; size 20"
    #CommentSpread, "boxfilled 550 430"

    #CommentSpread, "place 80 80"
    #CommentSpread, "\";

    #CommentSpread, "place 50 50; font rio_oro 0 37"
    #CommentSpread, "\Ratio of comment cards per month"
    #CommentSpread, "\Total: "; x
    #CommentSpread, "font andale 0 17"

    #CommentSpread, "color white; size 3; line 30 70 30 350; line 30 350 517 350"   'Draw the "x/y" axis.

    #CommentSpread, "place 40 370; down"        'placement of the month abrivation and the % of cards from that month
    #CommentSpread, "\JAN"
    #CommentSpread, "place 40 390;"
    #CommentSpread, "\"; aa; "%"
    #CommentSpread, "place 40 410"
    #CommentSpread, "\("; aaa; ")"

    #CommentSpread, "place 80 370"
    #CommentSpread, "\FEB"
    #CommentSpread, "place 80 390;"
    #CommentSpread, "\"; bb; "%"
    #CommentSpread, "place 80 410"
    #CommentSpread, "\("; bbb; ")"

    #CommentSpread, "place 120 370"
    #CommentSpread, "\MAR"
    #CommentSpread, "place 120 390;"
    #CommentSpread, "\"; cc; "%"
    #CommentSpread, "place 120 410"
    #CommentSpread, "\("; ccc; ")"

    #CommentSpread, "place 160 370"
    #CommentSpread, "\APR"
    #CommentSpread, "place 160 390;"
    #CommentSpread, "\"; dd; "%"
    #CommentSpread, "place 160 410"
    #CommentSpread, "\("; ddd; ")"

    #CommentSpread, "place 200 370"
    #CommentSpread, "\MAY"
    #CommentSpread, "place 200 390;"
    #CommentSpread, "\"; ee; "%"
    #CommentSpread, "place 200 410"
    #CommentSpread, "\("; eee; ")"

    #CommentSpread, "place 240 370"
    #CommentSpread, "\JUN"
    #CommentSpread, "place 240 390;"
    #CommentSpread, "\"; ff; "%"
    #CommentSpread, "place 240 410"
    #CommentSpread, "\("; fff; ")"

    #CommentSpread, "place 280 370"
    #CommentSpread, "\JUL"
    #CommentSpread, "place 280 390;"
    #CommentSpread, "\"; gg; "%"
    #CommentSpread, "place 280 410"
    #CommentSpread, "\("; ggg; ")"

    #CommentSpread, "place 320 370"
    #CommentSpread, "\AUG"
    #CommentSpread, "place 320 390;"
    #CommentSpread, "\"; hh; "%"
    #CommentSpread, "place 320 410"
    #CommentSpread, "\("; hhh; ")"

    #CommentSpread, "place 360 370"
    #CommentSpread, "\SEP"
    #CommentSpread, "place 360 390;"
    #CommentSpread, "\"; ii; "%"
    #CommentSpread, "place 360 410"
    #CommentSpread, "\("; iii; ")"

    #CommentSpread, "place 400 370"
    #CommentSpread, "\OCT"
    #CommentSpread, "place 400 390;"
    #CommentSpread, "\"; jj; "%"
    #CommentSpread, "place 400 410"
    #CommentSpread, "\("; jjj; ")"

    #CommentSpread, "place 440 370"
    #CommentSpread, "\NOV"
    #CommentSpread, "place 440 390;"
    #CommentSpread, "\"; kk; "%"
    #CommentSpread, "place 440 410"
    #CommentSpread, "\("; kkk; ")"

    #CommentSpread, "place 480 370"
    #CommentSpread, "\DEC"
    #CommentSpread, "place 480 390;"
    #CommentSpread, "\"; ll; "%"
    #CommentSpread, "place 480 410"
    #CommentSpread, "\("; lll; ")"

    #CommentSpread, "backcolor white"           'Set the fill color for the bars
    #CommentSpread, "place 45 "; jan              'Set height for January
    #CommentSpread, "boxfilled 60 350"
    #CommentSpread, "place 85 "; feb              'Set height for February
    #CommentSpread, "boxfilled 100 350"
    #CommentSpread, "place 125 "; mar             'Set height for March
    #CommentSpread, "boxfilled 140 350"
    #CommentSpread, "place 165 "; apr             'Set height for April
    #CommentSpread, "boxfilled 180 350"
    #CommentSpread, "place 205 "; may             'Set height for May
    #CommentSpread, "boxfilled 220 350"
    #CommentSpread, "place 245 "; jun             'Set height for June
    #CommentSpread, "boxfilled 260 350"
    #CommentSpread, "place 285 "; jul             'Set height for July
    #CommentSpread, "boxfilled 300 350"
    #CommentSpread, "place 325 "; aug             'Set height for August
    #CommentSpread, "boxfilled 340 350"
    #CommentSpread, "place 365 "; sep             'Set height for September
    #CommentSpread, "boxfilled 380 350"
    #CommentSpread, "place 405 "; oct             'Set height for October
    #CommentSpread, "boxfilled 420 350"
    #CommentSpread, "place 445 "; nov             'Set height for November
    #CommentSpread, "boxfilled 460 350"
    #CommentSpread, "place 485 "; dec             'Set height for December
    #CommentSpread, "boxfilled 500 350"
    #CommentSpread, "flush"
wait
end sub
sub TRPCLScommentspread handle$
    call CloseBarGraph
end sub

'*****************************************************************************************************************************************************************
'These are the closing routines for all of the windows. They are triggered by
'TRPCLSxxx subs, (one for each window respectively.)
sub CloseAll
    close #main
    if (WindowCheck(1) = 1) then call CloseMaster
    if (WindowCheck(2) = 1) then call CloseComMan
    if (WindowCheck(3) = 1) then call CloseReport
    if (WindowCheck(4) = 1) then call CloseReport2
    if (WindowCheck(5) = 1) then call CloseBarGraph
    if (WindowCheck(6) = 1) then call CloseReport3
    end
end sub

sub CloseMaster
    call MasterIsClosed
    close #master
end sub

sub CloseComMan
    call ComManIsClosed
    unloadbmp "Include"
    unloadbmp "Check"
    unloadbmp "Blank"
    close #ComMan
    if (WindowCheck(3)=1) then call CloseReport
end sub

sub CloseReport
    call ReportIsClosed
    close #report
end sub

sub CloseReport2
    call Report2IsClosed
    close #report2
end sub

sub CloseReport3
    call Report3IsClosed
    close #report3
end sub

sub CloseBarGraph
    call BarGraphIsClosed
    close#CommentSpread
end sub

'These subs set the open window indicator for #Master (array[1]), #ComMan (array[2]), #report (array[3]), and #report2 (array[4]
sub MasterIsOpen
    WindowCheck(1) = 1
end sub
sub MasterIsClosed
    WindowCheck(1) = 0
end sub

sub ComManIsOpen
    WindowCheck(2) = 1
end sub
sub ComManIsClosed
    WindowCheck(2) = 0
end sub

sub ReportIsOpen
    WindowCheck(3) = 1
end sub
sub ReportIsClosed
    WindowCheck(3) = 0
end sub

sub Report2IsOpen
    WindowCheck(4) = 1
end sub
sub Report2IsClosed
    WindowCheck(4) = 0
end sub

sub Report3IsOpen
    WindowCheck(6) = 1
end sub
sub Report3IsClosed
    WindowCheck(6) = 0
end sub

sub BarGraphIsOpen
    WindowCheck(5) = 1
end sub
sub BarGraphIsClosed
    WindowCheck(5) = 0
end sub

'******************************************************************************************************************************************************************
'****************************FUNCTIONS*****************************************************************************************************************************
'******************************************************************************************************************************************************************

'FUNCTION DESCRIPTION: Accepts the name of  RECORD type file and its LEM and returns the number of entries
dim FieldLength(30), RecordField$(100)'This array is associated with OpenRandomFile()
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

' FUNCTION DESCRIPTION: (Framed as a subroutine) sends notification when there is no records.dat file
sub NoData
    notice "There is no data to work with"
end sub

'FUNCTION DESCRIPTION: Framed as subroutine) sends notification when a window is alredy open
sub WindowOpen
    notice "That window is already open"
end sub

'FUNCTION DESCRIPTION: processes the stats.dat file and draws one of three pie charts based on string passed
'                       ("Service", "Cleanliness", "Recommend")
sub drawCircle PieType$
open "Stats.dat" for random as #Stats LEN = 80                                          'The stats.dat file is opened and the string passed as "PieType$ is used to
    field #Stats,_                                                                      'determin which data is assigned to the variables x, y, and z.
            10 as Q1EStat, 10 as Q1GStat, 10 as Q1FStat,_
            10 as Q2EStat, 10 as Q2GStat, 10 as Q2FStat,_
            10 as Q3YStat, 10 as Q3NStat
        get #Stats, 1
    if (PieType$ = "Service") then      'Assign x,y,z based on what type of pie
        x = Q1EStat
        y = Q1GStat
        z = Q1FStat
    end if
    if (PieType$ = "Cleanliness") then
        x = Q2EStat
        y = Q2GStat
        z = Q2FStat
    end if
    if (PieType$ = "Recommend") then
        x = Q3YStat
        y = 0
        z = Q3NStat
    end if
close #Stats                                                                                            'Stats.dat file is closed
    NumRecords = x+y+z                                      'NumRecords is the total value of the 360 degree circle that is draw.
    if NumRecords = 0 then exit sub             'if there is no records of that kind the sub is exited meaning no circle is drawn.
    h = int((x/NumRecords)* 100)                'This sets the % that each wedge is prior to setting the actual
    i = int((y/NumRecords)* 100)                'degree-per-point value.
    j = int((z/NumRecords)* 100)
    if (h+i+j = 99) then h = h+1
    EntryDegreeValue = 360 / NumRecords
    x = EntryDegreeValue * x
    y = EntryDegreeValue * y
    z = EntryDegreeValue * z
        if (x > 0) then         'draw the wedges for the stat
            #report "backcolor blue"
            #report "piefilled 120 120 200 "; x   'This is the line above but with variables.
        end if
        if (y>0) then
            #report "backcolor yellow"                    'create the next wedge.
            #report "pieFilled 120 120 "; 200+x; " "; y
        end if
        if (z>0) then
            #report "backcolor red"
            #report "pieFilled 120 120 "; 200+x+y; " "; z
        end if
        #report "north; up; go 60; turn 90; go 80; down;" 'Draw the circles for the key
        #report "backcolor blue"
        #report "circlefilled 10"
        #report "up; turn 90; go 60; down;"
        #report "backcolor yellow;"
        if (PieType$ = "Recommend") then #report "backcolor red"
            #report "circlefilled 10"
        if (PieType$ = "Service") or (PieType$ = "Cleanliness") then
            #report "up; go 60; down;"
            #report "backcolor red;"
            #report "circlefilled 10"
        end if
            #report "backcolor white"      'Write the Labels and the xx%
        if (PieType$ = "Service") or (PieType$ = "Cleanliness") then
            #report "up; turn -90; go 15; turn -90; go 125; down;" 'facing north   'Sets the labels and %
            #report "\Excellent"                               'for Service and Cleanliness
            #report "\"; h; " %"
            #report "up; turn 180; go 30; down;"
            #report "\Good"
            #report "\"; i;" %"
            #report "up; go 30; down;"
            #report "\Fair"
            #report "\"; j; " %"
        end if
        if (PieType$ = "Recommend") then    'This logic set the labels for the Recommend yes/no stat.
            #report "up; turn -90; go 15; turn -90; go 65;"
            #report "\Yes"
            #report "\";h;" %"
            #report "up; turn 180; go 30; down;"
            #report "\No"
            #report "\";j;" %"
        end if
end sub




