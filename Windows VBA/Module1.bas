Attribute VB_Name = "Module1"
Sub goToHowToUse()
    ' Checks if macros are enabled
    ActivePresentation.Slides(1).Shapes("MacroDisabledText").Visible = False
    SlideShowWindows(1).View.GotoSlide 13
End Sub

Sub goToSetUp()
    ' Checks if macros are enabled
    ActivePresentation.Slides(1).Shapes("MacroDisabledText").Visible = False
    SlideShowWindows(1).View.GotoSlide 7
End Sub

Sub goToPuzzleBoard()
    ' Checks if macros are enabled
    ActivePresentation.Slides(1).Shapes("MacroDisabledText").Visible = False
    SlideShowWindows(1).View.GotoSlide 2
End Sub

Sub BGChange()
    Dim i As Integer
    Dim themeNumber
    Set themeNumber = ActivePresentation.Slides(2).Shapes("ThemeNumber").TextFrame.TextRange
    Dim puzzleBoardGradientBottom As Long
    Dim puzzleBoardGradientMiddle As Long
    Dim puzzleBoardGradientTop As Long
    Dim setFloorEdges As Long
    Dim setFloorMiddle As Long
    Dim setFloorLine As Long
    Dim wheelGradientMiddle As Long
    Dim wheelGradientTop As Long
    Dim categoryColor As Long
    Dim letterSelectorColor As Long
    If themeNumber.Text = "1" Then
        puzzleBoardGradientBottom = RGB(2, 127, 190)
        puzzleBoardGradientMiddle = RGB(94, 189, 208)
        puzzleBoardGradientTop = RGB(2, 127, 190)
        setFloorEdges = RGB(0, 51, 0)
        setFloorMiddle = RGB(0, 153, 0)
        setFloorLine = RGB(38, 100, 38)
        wheelGradientMiddle = RGB(2, 149, 156)
        wheelGradientTop = RGB(2, 89, 132)
        categoryColor = RGB(27, 91, 33)
        letterSelectorColor = RGB(185, 205, 229)
    ElseIf themeNumber.Text = "2" Then
        puzzleBoardGradientBottom = RGB(233, 115, 160)
        puzzleBoardGradientMiddle = RGB(250, 194, 210)
        puzzleBoardGradientTop = RGB(233, 115, 160)
        setFloorEdges = RGB(207, 73, 143)
        setFloorMiddle = RGB(229, 131, 194)
        setFloorLine = RGB(197, 103, 139)
        wheelGradientMiddle = RGB(214, 128, 161)
        wheelGradientTop = RGB(179, 51, 106)
        categoryColor = RGB(207, 61, 144)
        letterSelectorColor = RGB(248, 196, 223)
    ElseIf themeNumber.Text = "3" Then
        puzzleBoardGradientBottom = RGB(0, 0, 0)
        puzzleBoardGradientMiddle = RGB(0, 0, 0)
        puzzleBoardGradientTop = RGB(0, 0, 0)
        setFloorEdges = RGB(38, 38, 38)
        setFloorMiddle = RGB(87, 68, 35)
        setFloorLine = RGB(38, 38, 38)
        wheelGradientMiddle = RGB(0, 0, 0)
        wheelGradientTop = RGB(0, 0, 0)
        categoryColor = RGB(38, 38, 38)
        letterSelectorColor = RGB(127, 127, 127)
    Else:
        setFloorEdges = RGB(41, 38, 35)
        setFloorMiddle = RGB(125, 73, 126)
        setFloorLine = RGB(16, 37, 63)
        categoryColor = RGB(23, 55, 94)
        letterSelectorColor = RGB(79, 129, 189)
    End If
    With ActivePresentation.Slides(2)
        With .Shapes("BackDrop")
            If themeNumber.Text = "4" Then
                .Fill.Transparency = 1
            Else:
                .Fill.Transparency = 0
                .Fill.GradientStops.Insert puzzleBoardGradientTop, 0
                .Fill.GradientStops.Insert puzzleBoardGradientMiddle, 0.5
                .Fill.GradientStops.Insert puzzleBoardGradientBottom, 1
                .Fill.GradientStops.Delete (1)
                .Fill.GradientStops.Delete (1)
                .Fill.GradientStops.Delete (1)
            End If
        End With
        With .Shapes("SetFloor")
            .Fill.GradientStops.Insert setFloorEdges, 0
            .Fill.GradientStops.Insert setFloorMiddle, 0.5
            .Fill.GradientStops.Insert setFloorEdges, 1
            .Fill.GradientStops.Delete (1)
            .Fill.GradientStops.Delete (1)
            .Fill.GradientStops.Delete (1)
            .Line.ForeColor.RGB = setFloorLine
        End With
        With .Shapes("CategoryBox")
            .Fill.GradientStops.Insert categoryColor, 0
            .Fill.GradientStops.Insert categoryColor, 0.15
            .Fill.GradientStops.Insert categoryColor, 0.85
            .Fill.GradientStops.Insert categoryColor, 1
            .Fill.GradientStops.Delete (1)
            .Fill.GradientStops.Delete (1)
            .Fill.GradientStops.Delete (1)
            .Fill.GradientStops.Delete (1)
            .Fill.GradientStops(1).Transparency = 1
            .Fill.GradientStops(4).Transparency = 1
        End With
        With .Shapes("LetterSelectionOverlay")
            .Fill.ForeColor.RGB = letterSelectorColor
        End With
    End With
    For i = 3 To 6
        With ActivePresentation.Slides(i)
            With .Shapes("BackDrop")
                If themeNumber.Text = "4" Then
                    .Fill.Transparency = 1
                Else:
                    .Fill.Transparency = 0
                    .Fill.GradientStops.Insert wheelGradientTop, 0
                    .Fill.GradientStops.Insert wheelGradientMiddle, 0.5
                    .Fill.GradientStops.Insert puzzleBoardGradientTop, 1
                    .Fill.GradientStops.Delete (1)
                    .Fill.GradientStops.Delete (1)
                    .Fill.GradientStops.Delete (1)
                End If
            End With
        End With
    Next i
    If CInt(themeNumber.Text) < 4 Then
        themeNumber.Text = CStr(CInt(themeNumber.Text) + 1)
    Else:
        themeNumber.Text = "1"
    End If
End Sub

Sub ClearBoardButton()
    Dim i As Integer
    For i = 1 To 52
        ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).TextFrame.TextRange.Text = ""
        ActivePresentation.Slides(2).Shapes("PuzzleCache" & i).TextFrame.TextRange.Text = ""
        ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).Fill.ForeColor.RGB = RGB(24, 154, 80)
    Next i
    For j = 1 To 26
        ActivePresentation.Slides(2).Shapes("Letter" & j).Visible = False
        bringLetterBack (j)
    Next j
    ActivePresentation.Slides(2).Shapes("CategoryBox").TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(2).Shapes("LeftTab").Fill.ForeColor.RGB = RGB(41, 183, 233)
    ActivePresentation.Slides(2).Shapes("LeftTab").TextFrame.TextRange.Text = "Load Puzzle"
    ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = False
    resetBonusRound
End Sub

Sub PlayerRoundDollarSign(oSh As Shape)
    Dim i As Integer
    Dim j As Boolean
    For i = 1 To 3
        If ActivePresentation.Slides(2).Shapes("Player" & i & "RoundDollarSign Compatibility Layer").Name = oSh.Name Then
            j = True
            Exit For
        End If
    Next i
    If j = True Then
        Set roundDollarAmount = ActivePresentation.Slides(2).Shapes("Player" & i & "RoundDollarAmount").OLEFormat.Object
        Set playerName = ActivePresentation.Slides(2).Shapes("Player" & i & "Name").OLEFormat.Object

        If IsNumeric(roundDollarAmount) = False Then
            MsgBox ("Please remove non-numeric characters from " + playerName.Value + "'s score before using this action.")
        Else:
            addMoney = InputBox("Add money to " + playerName.Value + "'s score." & vbNewLine & vbNewLine & "Enter the number of letters, multiplied by the wheel value.", "Quick Add", ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text)
            If addMoney = vbNullString Then
                Exit Sub
            ElseIf IsNumeric(addMoney) = False Then
                Dim k As Boolean
                For c = 1 To Len(addMoney)
                    charInQ = Mid(addMoney, c, 1)
                
                    If StrComp(charInQ, "+") = 0 Or StrComp(charInQ, "-") = 0 Or StrComp(charInQ, "*") = 0 Or StrComp(charInQ, "x") = 0 Or StrComp(charInQ, "/") = 0 Then
                        theOperator = charInQ
                        splitMoney = Split(addMoney, theOperator)
                            If UBound(splitMoney) > 1 Or InStr(splitMoney(1), "+") > 0 Or InStr(splitMoney(1), "-") > 0 Or InStr(splitMoney(1), "*") > 0 Or InStr(splitMoney(1), "x") > 0 Or InStr(splitMoney(1), "/") > 0 Then
                                MsgBox ("You can only use one operator. If you're adding multiple numbers, try multiplying instead.")
                                Exit Sub
                            End If
                        k = True
                        Exit For
                    End If
                Next c
                If k = True Then
                    If IsNumeric(splitMoney(0)) = False Or IsNumeric(splitMoney(1)) = False Then
                        MsgBox ("There are non-numeric values in your expression. Please try again.")
                        Exit Sub
                    End If
                    If StrComp(theOperator, "+") = 0 Then
                        resultMoney = (splitMoney(0) + 0) + (splitMoney(1) + 0)
                    ElseIf StrComp(theOperator, "-") = 0 Then
                        resultMoney = (splitMoney(0) + 0) - (splitMoney(1) + 0)
                    ElseIf StrComp(theOperator, "*") = 0 Or StrComp(theOperator, "x") = 0 Then
                        resultMoney = (splitMoney(0) + 0) * (splitMoney(1) + 0)
                    Else:
                        If splitMoney(1) = 0 Then
                            MsgBox ("Whoa there, don't destroy the universe!")
                            Exit Sub
                        End If
                        resultMoney = (splitMoney(0) + 0) / (splitMoney(1) + 0)
                    End If
                    roundDollarAmount.Value = (roundDollarAmount + 0) + (resultMoney + 0)
                Else:
                MsgBox ("You can only enter numbers or arithmetic expressions here.")
                End If
        Else:
            roundDollarAmount.Value = (roundDollarAmount + 0) + (addMoney + 0)
        End If
    End If
End If
End Sub

Sub PlayerBuyaVowel(oSh As Shape)
    Dim i As Integer
    Dim j As Boolean
    Dim VOWELCOST As Integer
    VOWELCOST = CInt(ActivePresentation.Slides(9).Shapes("VowelPrice").TextFrame.TextRange.Text)
    For i = 1 To 3
        If ActivePresentation.Slides(2).Shapes("Player" & i & "BuyVowelButton").Name = oSh.Name Then
            j = True
            Exit For
        End If
    Next i
    If j = True Then
        Set roundDollarAmount = ActivePresentation.Slides(2).Shapes("Player" & i & "RoundDollarAmount").OLEFormat.Object
        Set playerName = ActivePresentation.Slides(2).Shapes("Player" & i & "Name").OLEFormat.Object
        If IsNumeric(roundDollarAmount) = False Then
            MsgBox ("Please remove non-numeric characters from " + playerName.Value + "'s score before buying a vowel.")
        ElseIf roundDollarAmount.Value < VOWELCOST Then
            MsgBox (playerName.Value + " cannot buy a vowel. Vowels cost $" + CStr(VOWELCOST) + ".")
        ElseIf roundDollarAmount.Value >= VOWELCOST Then
            roundDollarAmount.Value = (roundDollarAmount - VOWELCOST)
        End If
    End If
End Sub

Sub PlayerTransferTotals(oSh As Shape)
    Dim i As Integer
    Dim j As Boolean
    Dim HOUSEMINIMUM As Integer
    HOUSEMINIMUM = CInt(ActivePresentation.Slides(9).Shapes("HouseMinimum").TextFrame.TextRange.Text)
    For i = 1 To 3
        If ActivePresentation.Slides(2).Shapes("Player" & i & "TransferTotalsButton").Name = oSh.Name Then
            j = True
            Exit For
        End If
    Next i
    If j = True Then
        Set roundDollarAmount = ActivePresentation.Slides(2).Shapes("Player" & i & "RoundDollarAmount").OLEFormat.Object
        Set TotalsDollarAmount = ActivePresentation.Slides(2).Shapes("Player" & i & "TotalsDollarAmount").OLEFormat.Object
        If IsNumeric(TotalsDollarAmount) = True And IsNumeric(roundDollarAmount) = True Then
            If roundDollarAmount.Value < HOUSEMINIMUM Then
                shouldIHouse = MsgBox("The house minimum of $" + CStr(HOUSEMINIMUM) + " will be transferred.", vbOKCancel)
                If shouldIHouse = vbOK Then
                    TotalsDollarAmount.Value = (TotalsDollarAmount + 0) + HOUSEMINIMUM
                    wipeRoundScores
                Else:
                    Exit Sub
                End If
            Else:
                TotalsDollarAmount.Value = (TotalsDollarAmount + 0) + (roundDollarAmount + 0)
                wipeRoundScores
            End If
        End If
    End If
End Sub

Sub PlayerXButton(oSh As Shape)
    Dim i As Integer
    Dim j As Boolean
    For i = 1 To 3
        If ActivePresentation.Slides(2).Shapes("Player" & i & "XButton Compatibility Layer").Name = oSh.Name Then
            j = True
            Exit For
        End If
    Next i
    If j = True Then
        Set roundDollarAmount = ActivePresentation.Slides(2).Shapes("Player" & i & "RoundDollarAmount").OLEFormat.Object
        roundDollarAmount.Value = 0
    End If
End Sub

Sub DetermineMystery()
    ActivePresentation.Slides(4).Shapes("MysteryIndicator").TextFrame.TextRange.Text = ""
    wait (0.1)
    Randomize
    randomNumber = Int(2 * Rnd) + 1
    If randomNumber = 2 Then
        ActivePresentation.Slides(4).Shapes("MysteryIndicator").TextFrame.TextRange.Text = "$10,000"
    Else:
        ActivePresentation.Slides(4).Shapes("MysteryIndicator").TextFrame.TextRange.Text = "BANKRUPT"
    End If
    If ActivePresentation.Slides(4).Shapes("TheWheel").Rotation >= 112.5 And ActivePresentation.Slides(4).Shapes("TheWheel").Rotation < 127.5 Then
        ActivePresentation.Slides(4).Shapes("TheWheel").GroupItems(5).Fill.Transparency = 1
        ActivePresentation.Slides(4).Shapes("TheWheel2").GroupItems(5).Fill.Transparency = 1
    ElseIf ActivePresentation.Slides(4).Shapes("TheWheel").Rotation >= 292.5 And ActivePresentation.Slides(4).Shapes("TheWheel").Rotation < 307.5 Then
        ActivePresentation.Slides(4).Shapes("TheWheel").GroupItems(6).Fill.Transparency = 1
        ActivePresentation.Slides(4).Shapes("TheWheel2").GroupItems(5).Fill.Transparency = 1
    End If
End Sub

Sub ClearMysteryIndicator()
    ActivePresentation.Slides(4).Shapes("MysteryIndicator").TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(4).Shapes("TheWheel").GroupItems(5).Fill.Transparency = 0
    ActivePresentation.Slides(4).Shapes("TheWheel2").GroupItems(5).Fill.Transparency = 0
    ActivePresentation.Slides(4).Shapes("TheWheel").GroupItems(6).Fill.Transparency = 0
    ActivePresentation.Slides(4).Shapes("TheWheel2").GroupItems(6).Fill.Transparency = 0
End Sub

Sub TileChanger(i As Integer)
    Set oSh = ActivePresentation.Slides(8).Shapes("BoardTile" & CStr(i)).OLEFormat.Object
    If oSh.Value = "" Or oSh.Value = " " Then
        oSh.Value = ""
        oSh.BackColor = &H509A18
    Else:
        oSh.Value = UCase(oSh.Value)
        oSh.BackColor = &HFFFFFF
    End If
End Sub

Sub ErasePuzzleRow(oSh As Shape)
    If oSh.Name = "Eraser1" Then
        minim = 1
        maxim = 12
    ElseIf oSh.Name = "Eraser2" Then
        minim = 13
        maxim = 26
    ElseIf oSh.Name = "Eraser3" Then
        minim = 27
        maxim = 40
    ElseIf oSh.Name = "Eraser4" Then
        minim = 41
        maxim = 52
    End If
    For i = minim To maxim
        SlideShowWindows(1).View.Slide.Shapes("BoardTile" & CStr(i)).OLEFormat.Object.Value = ""
        SlideShowWindows(1).View.Slide.Shapes("BoardTile" & CStr(i)).OLEFormat.Object.BackColor = &H509A18
    Next i
End Sub

Sub EraseEntirePuzzle()
    shouldIEraseAll = MsgBox("Are you sure you want to delete the entire puzzle?", vbYesNo + vbDefaultButton2)
    If shouldIEraseAll = vbYes Then
    For i = 1 To 52
        SlideShowWindows(1).View.Slide.Shapes("BoardTile" & CStr(i)).OLEFormat.Object.Value = ""
        SlideShowWindows(1).View.Slide.Shapes("BoardTile" & CStr(i)).OLEFormat.Object.BackColor = &H509A18
    Next i
    SlideShowWindows(1).View.Slide.Shapes("CategoryBox").OLEFormat.Object.Value = ""
    Else
        Exit Sub
    End If
End Sub

Sub puzzleSetupFromOtherSlide(oSh As Shape)
    On Error GoTo errHandler
    Dim i As Integer
    Dim j As Boolean
    For i = 1 To 12
        If ActivePresentation.Slides(8).Shapes("LinkTo" & i).Name = oSh.Name Then
            j = True
            Exit For
        End If
    Next i
    If j = True Then
        savePuzzle
        placePuzzleToSetUp (i)
        shadeOccupiedPuzzles
        highlightCurrentPuzzle (i)
        SlideShowWindows(1).View.GotoSlide 8
    End If
    Exit Sub
errHandler:
    MsgBox "Cannot edit puzzles because ActiveX components are disabled." & vbNewLine & _
    "If you use Windows, check if ActiveX is enabled in Trust Center settings." & vbNewLine & _
    "If you use macOS, download the Mac version of Wheel of Fortune for PowerPoint.", 0, "Set Up Puzzles Error"
End Sub

Sub puzzleSetup(oSh As Shape)
    Dim i As Integer
    Dim j As Boolean
    For i = 1 To 12
        If ActivePresentation.Slides(8).Shapes("LinkTo" & i).Name = oSh.Name Then
            j = True
            Exit For
        End If
    Next i
    If j = True Then
        savePuzzle
        placePuzzleToSetUp (i)
        shadeOccupiedPuzzles
        highlightCurrentPuzzle (i)
    End If
End Sub

Sub EditVowelPrice(oClickedShape As Shape)
    On Error GoTo errHandler
    Dim oSh As Shape
    Dim sText As String
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    sText = InputBox("Edit the vowel price. The default price is $250.", "Edit Vowel Price", CInt(oSh.TextFrame.TextRange.Text))
    While IsNumeric(sText) = False And sText <> ""
    sText = InputBox("You can only enter numbers here. Try again:", "Edit Vowel Price", sText)
    Wend
    If sText = "" Then
        Exit Sub
    Else:
    newText = CInt(sText)
    oSh.TextFrame.TextRange.Text = "$" & newText
    End If
    Exit Sub
errHandler:
    MsgBox ("The vowel price cannot exceed $32767.")
End Sub

Sub EditHouseMinimum(oClickedShape As Shape)
    On Error GoTo errHandler
    Dim oSh As Shape
    Dim sText As String
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    sText = InputBox("Edit the house minimum. The default minimum is $1000.", "Edit House Minimum", CInt(oSh.TextFrame.TextRange.Text))
    While IsNumeric(sText) = False And sText <> ""
    sText = InputBox("You can only enter numbers here. Try again:", "Edit House Minimum", sText)
    Wend
    If sText = "" Then
        Exit Sub
    Else:
    newText = CInt(sText)
    oSh.TextFrame.TextRange.Text = "$" & newText
    End If
    Exit Sub
errHandler:
    MsgBox ("The house minimum cannot exceed $32767.")
End Sub

Sub confirmDeleteAllPuzzles()
    deletionConfirm = MsgBox("Are you ABSOLUTELY sure you want to delete all puzzles?", vbYesNo + vbDefaultButton2)
    If deletionConfirm = vbYes Then
        deleteAllPuzzles
    Else
        Exit Sub
    End If
End Sub

Private Sub highlightCurrentPuzzle(i)
    With ActivePresentation.Slides(8).Shapes("LinkTo" & CInt(i))
        .ZOrder msoBringToFront
        .Fill.ForeColor.RGB = RGB(250, 192, 144)
        .Line.ForeColor.RGB = RGB(228, 108, 10)
    End With
End Sub

Private Sub savePuzzle()
    Dim thereWasAPuzzle As Boolean
    thereWasAPuzzle = False
    For i = 1 To 12
        If ActivePresentation.Slides(8).Shapes("LinkTo" & CStr(i)).Fill.ForeColor.RGB = RGB(250, 192, 144) Then
            thereWasAPuzzle = True
            Exit For
        End If
    Next i
    If thereWasAPuzzle = True Then
        For j = 1 To 52
            ActivePresentation.Slides(10).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(j)).TextFrame.TextRange.Text = ActivePresentation.Slides(8).Shapes("BoardTile" & CStr(j)).OLEFormat.Object.Value
            ActivePresentation.Slides(10).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(j)).Fill.ForeColor.RGB = ActivePresentation.Slides(8).Shapes("BoardTile" & CStr(j)).OLEFormat.Object.BackColor
        Next j
        ActivePresentation.Slides(10).Shapes("PuzzleCategory" & CStr(i)).TextFrame.TextRange.Text = ActivePresentation.Slides(8).Shapes("CategoryBox").OLEFormat.Object.Value
    End If
End Sub

Private Sub placePuzzleToSetUp(i)
    For n = 1 To 52
        ActivePresentation.Slides(8).Shapes("BoardTile" & CStr(n)).OLEFormat.Object.Value = ActivePresentation.Slides(10).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(n)).TextFrame.TextRange.Text
    Next n
    ActivePresentation.Slides(8).Shapes("CategoryBox").OLEFormat.Object.Value = ActivePresentation.Slides(10).Shapes("PuzzleCategory" & CStr(i)).TextFrame.TextRange.Text
End Sub

Private Sub deleteAllPuzzles()
    For i = 1 To 12
        For j = 1 To 52
            ActivePresentation.Slides(10).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(j)).TextFrame.TextRange.Text = ""
            ActivePresentation.Slides(10).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(j)).Fill.ForeColor.RGB = RGB(24, 154, 80)
        Next j
        ActivePresentation.Slides(10).Shapes("PuzzleCategory" & CStr(i)).TextFrame.TextRange.Text = ""
    Next i
    For k = 1 To 52
        ActivePresentation.Slides(8).Shapes("BoardTile" & CStr(k)).OLEFormat.Object.Value = ""
    Next k
    ActivePresentation.Slides(8).Shapes("CategoryBox").OLEFormat.Object.Value = ""
    shadeOccupiedPuzzles
End Sub

Private Sub shadeOccupiedPuzzles()
    For p = 1 To 12
        blankPuzzle = True
        For q = 1 To 52
            If ActivePresentation.Slides(10).Shapes("PuzzleSolution" & CStr(p) & "-" & CStr(q)).TextFrame.TextRange.Text <> "" Then
                blankPuzzle = False
                Exit For
            End If
        Next q
        If ActivePresentation.Slides(10).Shapes("PuzzleCategory" & CStr(p)).TextFrame.TextRange.Text <> "" Then
            blankPuzzle = False
        End If
        For r = 7 To 9
           If blankPuzzle = False Then
               With ActivePresentation.Slides(r).Shapes("LinkTo" & CStr(p))
                   .Fill.ForeColor.RGB = RGB(146, 224, 204)
                   .Line.ForeColor.RGB = RGB(55, 96, 146)
               End With
           Else:
               With ActivePresentation.Slides(r).Shapes("LinkTo" & CStr(p))
                   .Fill.ForeColor.RGB = RGB(149, 179, 215)
                   .Line.ForeColor.RGB = RGB(55, 96, 146)
               End With
           End If
        Next r
    Next p
End Sub

Sub OnSlideShowTerminate(oWn As SlideShowWindow)
    savePuzzle
    resetBonusRound
    toggleBonusRound (False)
    For j = 3 To 6
        If ActivePresentation.Slides(j).Shapes("BackOval").Visible = False Then
            With ActivePresentation.Slides(j)
                .Shapes("BackOval").Visible = True
                .Shapes("BackText").Visible = True
            End With
        End If
    Next j
    ActivePresentation.Slides(1).Shapes("MacroDisabledText").Visible = True
    cancelSwap
End Sub

Sub OnSlideShowPageChange(ByVal SSW As SlideShowWindow)
        If SlideShowWindows(1).View.LastSlideViewed.SlideIndex = 8 Then
            savePuzzle
            shadeOccupiedPuzzles
            If Application.Version = 12# Then
                If SSW.View.CurrentShowPosition = 9 Then
                    ppt2007RefreshFix
                End If
            End If
        ElseIf SlideShowWindows(1).View.LastSlideViewed.SlideIndex = 2 Then
            ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = False
        Else:
            Exit Sub
        End If
End Sub

' PowerPoint 2007 has a screen refresh bug, so we need this code to work around it: http://www.pptalchemy.co.uk/PowerPoint_Screen%20_Refresh.html
Private Sub ppt2007RefreshFix()
    Dim osld As Slide
    Set osld = ActivePresentation.SlideShowWindow.View.Slide
    With osld.Shapes.AddTextbox(msoTextOrientationHorizontal, 1, 1, 1, 1)
        .Delete
    End With
End Sub

Private Function puzzleExists(i) As Boolean
    Dim puzzleBoolean As Boolean
    Dim m As Integer
    puzzleBoolean = False
    For m = 1 To 52
        If ActivePresentation.Slides(10).Shapes("PuzzleSolution" & i & "-" & m).Fill.ForeColor.RGB = RGB(255, 255, 255) Then
            puzzleBoolean = True
            Exit For
        End If
    Next m
    If ActivePresentation.Slides(10).Shapes("PuzzleCategory" & i).TextFrame.TextRange.Text <> "" Then
        puzzleBoolean = True
    End If
    puzzleExists = puzzleBoolean
End Function

Sub LoadPuzzleOrSolve()
    If ActivePresentation.Slides(2).Shapes("LeftTab").TextFrame.TextRange.Text = "Load Puzzle" Then
        Dim noPuzzlesExist As Boolean
        noPuzzlesExist = True
        For j = 1 To 12
            If puzzleExists(j) = True Then
                noPuzzlesExist = False
                Exit For
            End If
        Next j
        If noPuzzlesExist = True Then
            MsgBox ("No puzzles were found. Create puzzles using Set Up Puzzles on the top right of this slide.")
            Exit Sub
        End If
        numberToLoad = InputBox("Enter the puzzle number to load:", "Load Puzzle")
        While IsNumeric(numberToLoad) = False Or numberToLoad < 1 Or numberToLoad > 12
            If numberToLoad = "" Then
                Exit Sub
            Else:
                numberToLoad = InputBox("Valid puzzle numbers range from 1 to 12. Try again:", "Load Puzzle", numberToLoad)
            End If
        Wend
            loadPuzzle (CInt(numberToLoad))
    Else:
        Dim alreadySolved As Boolean
        alreadySolved = True
        For i = 1 To 52
            If ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).TextFrame.TextRange.Text <> ActivePresentation.Slides(2).Shapes("PuzzleCache" & i).TextFrame.TextRange.Text Then
                alreadySolved = False
                Exit For
            End If
        Next i
        If alreadySolved = True Or ActivePresentation.Slides(9).Shapes("ConfirmSolve").TextFrame.TextRange.Text = "off" Then
            solvePuzzle
            Exit Sub
        Else:
        solveConfirm = MsgBox("Are you sure you want to reveal the puzzle?", vbYesNo + vbDefaultButton1)
            If solveConfirm = vbYes Then
                solvePuzzle
            Else
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub loadPuzzle(i)
    If puzzleExists(i) = False Then
        MsgBox ("No puzzle found for number " & i & ".")
        Exit Sub
    End If
    ClearBoardButton
    For j = 1 To 52
        ActivePresentation.Slides(2).Shapes("PuzzleBoard" & j).Fill.ForeColor.RGB = ActivePresentation.Slides(10).Shapes("PuzzleSolution" & i & "-" & j).Fill.ForeColor.RGB
        ActivePresentation.Slides(2).Shapes("PuzzleCache" & j).TextFrame.TextRange.Text = ActivePresentation.Slides(10).Shapes("PuzzleSolution" & i & "-" & j).TextFrame.TextRange.Text
        If isLetter(ActivePresentation.Slides(10).Shapes("PuzzleSolution" & i & "-" & j).TextFrame.TextRange.Text) = False Then
            ActivePresentation.Slides(2).Shapes("PuzzleBoard" & j).TextFrame.TextRange.Text = ActivePresentation.Slides(10).Shapes("PuzzleSolution" & i & "-" & j).TextFrame.TextRange.Text
        End If
    Next j
    ActivePresentation.Slides(2).Shapes("CategoryBox").TextFrame.TextRange.Text = ActivePresentation.Slides(10).Shapes("PuzzleCategory" & i).TextFrame.TextRange.Text
    For k = 1 To 26
        ActivePresentation.Slides(2).Shapes("Letter" & k).Visible = True
    Next k
    ActivePresentation.Slides(2).Shapes("LeftTab").Fill.ForeColor.RGB = RGB(198, 159, 48)
    ActivePresentation.Slides(2).Shapes("LeftTab").TextFrame.TextRange.Text = "Solve"
    ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = ""
End Sub

Private Sub solvePuzzle()
    For i = 1 To 52
        ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).TextFrame.TextRange.Text = ActivePresentation.Slides(2).Shapes("PuzzleCache" & i).TextFrame.TextRange.Text
        If ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).Fill.ForeColor.RGB = RGB(0, 0, 255) Then
            ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).Fill.ForeColor.RGB = RGB(255, 255, 255)
        End If
    Next i
    For j = 1 To 26
        ActivePresentation.Slides(2).Shapes("Letter" & j).Visible = False
    Next j
    ActivePresentation.Slides(2).Shapes("LeftTab").Fill.ForeColor.RGB = RGB(41, 183, 233)
    ActivePresentation.Slides(2).Shapes("LeftTab").TextFrame.TextRange.Text = "Load Puzzle"
    ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = False
End Sub

Private Function isLetter(strValue As String) As Boolean
    For i = 1 To Len(strValue)
        If Asc(Mid(strValue, 1, 1)) < 65 Or Asc(Mid(strValue, 1, 1)) > 90 Then
            isLetter = False
        Else:
            isLetter = True
        End If
    Next i
End Function

Sub guessLetter(oSh As Shape)
    Dim i As Integer
    Dim j As Boolean
    For i = 1 To 26
        If ActivePresentation.Slides(2).Shapes("Letter" & i).Name = oSh.Name Then
            j = True
            Exit For
        End If
    Next i
    If j = True Then
        If ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text <> "" Then
            Dim theLetter As String
            Dim letterCount As Integer
            letterCount = 0
            theLetter = ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text
            If ActivePresentation.Slides(2).Shapes("BonusOutline").Line.Transparency = 0 Then
                If Len(ActivePresentation.Slides(2).Shapes("BonusLetters").TextFrame.TextRange.Text) >= 5 Then
                    MsgBox ("The contestant can only choose four letters (or five if he or she has a wild card). Use the spiral arrow button to remove letters if necessary.")
                    Exit Sub
                Else:
                    ActivePresentation.Slides(2).Shapes("BonusLetters").TextFrame.TextRange.Text = ActivePresentation.Slides(2).Shapes("BonusLetters").TextFrame.TextRange.Text + theLetter
                    ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = ""
                    ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text = ""
                    ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = False
                    Exit Sub
                End If
            End If
            For k = 1 To 52
                If ActivePresentation.Slides(2).Shapes("PuzzleCache" & k).TextFrame.TextRange.Text = theLetter Then
                    If ActivePresentation.Slides(9).Shapes("BlueTiles").TextFrame.TextRange.Text = "off" Then
                        ActivePresentation.Slides(2).Shapes("PuzzleBoard" & k).TextFrame.TextRange.Text = theLetter
                    Else:
                        ActivePresentation.Slides(2).Shapes("PuzzleBoard" & k).Fill.ForeColor.RGB = RGB(0, 0, 255)
                    End If
                    letterCount = letterCount + 1
                End If
            Next k
            If letterCount = 0 Then
                ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = True
                ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = ""
                ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text = ""
                Exit Sub
            End If
            ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text = ""
            ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = letterCount & "*"
            ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = False
        End If
    End If
End Sub

Sub revealLetter(oSh As Shape)
    If ActivePresentation.Slides(2).Shapes("LeftTab").TextFrame.TextRange.Text = "Load Puzzle" Then
        LoadPuzzleOrSolve
    Else:
        Dim i As Integer
        Dim j As Boolean
        For i = 1 To 52
            If ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).Name = oSh.Name Then
                j = True
                Exit For
            End If
        Next i
        If j = True Then
            If ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).Fill.ForeColor.RGB = RGB(0, 0, 255) Then
                ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).TextFrame.TextRange.Text = ActivePresentation.Slides(2).Shapes("PuzzleCache" & i).TextFrame.TextRange.Text
                ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).Fill.ForeColor.RGB = RGB(255, 255, 255)
            End If
        End If
    End If
End Sub

Private Sub bringLetterBack(i)
    ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text = Chr(i + 64)
End Sub

Private Function wait(PauseTime)
     Start = Timer
     Do While Timer < Start + PauseTime
         DoEvents
     Loop
End Function

Sub wipeOnClose()
    deleteAllPuzzles
    ClearBoardButton
    ClearMysteryIndicator
    wipeRoundScores
    For i = 1 To 3
        ActivePresentation.Slides(2).Shapes("Player" & i & "Name").OLEFormat.Object.Value = "Player " & i
        ActivePresentation.Slides(2).Shapes("Player" & i & "TotalsDollarAmount").OLEFormat.Object.Value = "0"
    Next i
    If ActivePresentation.Slides(3).Shapes("TheWheel").GroupItems(4).Fill.Transparency = 1 Then
        toggleWildCard
    End If
    For j = 3 To 6
        ActivePresentation.Slides(j).Shapes("TheWheel").Rotation = 0
    Next j
    ActivePresentation.SlideShowWindow.View.Exit
End Sub

Private Sub wipeRoundScores()
    Dim i As Integer
    For i = 1 To 3
        ActivePresentation.Slides(2).Shapes("Player" & i & "RoundDollarAmount").OLEFormat.Object.Value = "0"
    Next i
End Sub

Sub toggleWildCard()
    Dim transparentLevel As Integer
    Dim wildString As String
    If ActivePresentation.Slides(3).Shapes("TheWheel").GroupItems(4).Fill.Transparency = 0 Then
        transparentLevel = 1
        wildString = "Show"
    Else:
        transparentLevel = 0
        wildString = "Remove"
    End If
    For i = 3 To 5
        ActivePresentation.Slides(i).Shapes("TheWheel").GroupItems(4).Fill.Transparency = transparentLevel
        ActivePresentation.Slides(i).Shapes("TheWheel2").GroupItems(4).Fill.Transparency = transparentLevel
        ActivePresentation.Slides(i).Shapes("ToggleWildCard").TextFrame.TextRange.Text = wildString + " Wild Card"
    Next i
End Sub

Sub toggleOnOff(oClickedShape As Shape)
  Dim oSh As Shape
    Dim sText As String
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    If oSh.TextFrame.TextRange.Text = "on" Then
        oSh.TextFrame.TextRange.Text = "off"
        oSh.Fill.ForeColor.RGB = RGB(217, 150, 148)
    Else:
        oSh.TextFrame.TextRange.Text = "on"
        oSh.Fill.ForeColor.RGB = RGB(195, 214, 155)
    End If
End Sub

Sub shiftRight(oClickedShape As Shape)
  Dim oSh As Shape
    Dim sText As String
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    Dim minim As Integer
    Dim maxim As Integer
    If oSh.Name = "ShiftRight1" Then
        minim = 1
        maxim = 12
    ElseIf oSh.Name = "ShiftRight2" Then
        minim = 13
        maxim = 26
    ElseIf oSh.Name = "ShiftRight3" Then
        minim = 27
        maxim = 40
    ElseIf oSh.Name = "ShiftRight4" Then
        minim = 41
        maxim = 52
    End If
    If ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(maxim)).OLEFormat.Object.Value = "" Then
        For i = maxim - 1 To minim Step -1
            ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(i + 1)).OLEFormat.Object.Value = ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(i)).OLEFormat.Object.Value
        Next i
        ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(i + 1)).OLEFormat.Object.Value = ""
    End If
End Sub

Sub shiftLeft(oClickedShape As Shape)
  Dim oSh As Shape
    Dim sText As String
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    Dim minim As Integer
    Dim maxim As Integer
    If oSh.Name = "ShiftLeft1" Then
        minim = 1
        maxim = 12
    ElseIf oSh.Name = "ShiftLeft2" Then
        minim = 13
        maxim = 26
    ElseIf oSh.Name = "ShiftLeft3" Then
        minim = 27
        maxim = 40
    ElseIf oSh.Name = "ShiftLeft4" Then
        minim = 41
        maxim = 52
    End If
    If ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(minim)).OLEFormat.Object.Value = "" Then
        For i = minim + 1 To maxim
            ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(i - 1)).OLEFormat.Object.Value = ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(i)).OLEFormat.Object.Value
        Next i
        ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(i - 1)).OLEFormat.Object.Value = ""
    End If
End Sub

Sub shiftUp()
    Dim i As Integer, j As Integer
    Dim blockerTiles
    blockerTiles = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 26)
    For i = LBound(blockerTiles) To UBound(blockerTiles)
        If ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(blockerTiles(i))).OLEFormat.Object.Value <> "" Then
            Exit Sub
        End If
    Next i
    For j = 14 To 25
        ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(j - 13)).OLEFormat.Object.Value = _
        ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(j)).OLEFormat.Object.Value
    Next j
    For j = 27 To 40
        ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(j - 14)).OLEFormat.Object.Value = _
        ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(j)).OLEFormat.Object.Value
    Next j
    ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(27)).OLEFormat.Object.Value = ""
    ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(40)).OLEFormat.Object.Value = ""
    For j = 41 To 52
        ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(j - 13)).OLEFormat.Object.Value = _
        ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(j)).OLEFormat.Object.Value
        ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(j)).OLEFormat.Object.Value = ""
    Next j
End Sub

Sub shiftDown()
    Dim i As Integer, j As Integer
    Dim blockerTiles
    blockerTiles = Array(27, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52)
    For i = LBound(blockerTiles) To UBound(blockerTiles)
        If ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(blockerTiles(i))).OLEFormat.Object.Value <> "" Then
            Exit Sub
        End If
    Next i
    For j = 28 To 39
        ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(j + 13)).OLEFormat.Object.Value = _
        ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(j)).OLEFormat.Object.Value
    Next j
    For j = 13 To 26
        ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(j + 14)).OLEFormat.Object.Value = _
        ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(j)).OLEFormat.Object.Value
    Next j
    ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(13)).OLEFormat.Object.Value = ""
    ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(26)).OLEFormat.Object.Value = ""
    For j = 1 To 12
        ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(j + 13)).OLEFormat.Object.Value = _
        ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(j)).OLEFormat.Object.Value
        ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(j)).OLEFormat.Object.Value = ""
    Next j
End Sub

Sub RSTLNE()
    If ActivePresentation.Slides(2).Shapes("Letter1").Visible = False Then
        MsgBox ("Please load a new puzzle before starting the bonus round.")
    ElseIf ActivePresentation.Slides(2).Shapes("RSTLNEOutline").Line.Transparency = 0 And ActivePresentation.Slides(2).Shapes("Letter1").Visible = True Then
        guessLetterViaFunction (18)
        guessLetterViaFunction (19)
        guessLetterViaFunction (20)
        guessLetterViaFunction (12)
        guessLetterViaFunction (14)
        guessLetterViaFunction (5)
        ActivePresentation.Slides(2).Shapes("BonusBox").Fill.ForeColor.RGB = RGB(225, 129, 75)
        ActivePresentation.Slides(2).Shapes("RSTLNEBox").Fill.ForeColor.RGB = RGB(166, 166, 166)
        ActivePresentation.Slides(2).Shapes("RSTLNEOutline").Line.Transparency = 1
        ActivePresentation.Slides(2).Shapes("BonusOutline").Line.Transparency = 0
    End If
End Sub

Private Sub guessLetterViaFunction(i)
Dim theLetter As String
theLetter = ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text
If theLetter = "" Then
    Exit Sub
End If
    For k = 1 To 52
        If ActivePresentation.Slides(2).Shapes("PuzzleCache" & k).TextFrame.TextRange.Text = theLetter Then
            If ActivePresentation.Slides(9).Shapes("BlueTiles").TextFrame.TextRange.Text = "off" Then
                ActivePresentation.Slides(2).Shapes("PuzzleBoard" & k).TextFrame.TextRange.Text = theLetter
            Else:
                ActivePresentation.Slides(2).Shapes("PuzzleBoard" & k).Fill.ForeColor.RGB = RGB(0, 0, 255)
            End If
        End If
    Next k
    ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = False
End Sub

Private Sub resetBonusRound()
    If ActivePresentation.Slides(2).Shapes("BonusOutline").Line.Transparency = 0 Then
        While Len(ActivePresentation.Slides(2).Shapes("BonusLetters").TextFrame.TextRange.Text) > 0
            removeBonusLetter
        Wend
    End If
    ActivePresentation.Slides(2).Shapes("BonusBox").Fill.ForeColor.RGB = RGB(166, 166, 166)
    ActivePresentation.Slides(2).Shapes("RSTLNEBox").Fill.ForeColor.RGB = RGB(225, 129, 75)
    ActivePresentation.Slides(2).Shapes("RSTLNEOutline").Line.Transparency = 0
    ActivePresentation.Slides(2).Shapes("BonusOutline").Line.Transparency = 1
    ActivePresentation.Slides(2).Shapes("BonusLetters").TextFrame.TextRange.Text = ""
End Sub

Sub removeBonusLetter()
    If ActivePresentation.Slides(2).Shapes("BonusLetters").TextFrame.TextRange.Text <> "" Then
        letterToReturn = Right(ActivePresentation.Slides(2).Shapes("BonusLetters").TextFrame.TextRange.Text, 1)
        ActivePresentation.Slides(2).Shapes("BonusLetters").TextFrame.TextRange.Text = Left(ActivePresentation.Slides(2).Shapes("BonusLetters").TextFrame.TextRange.Text, Len(ActivePresentation.Slides(2).Shapes("BonusLetters").TextFrame.TextRange.Text) - 1)
        For i = 1 To Len(letterToReturn)
            letterNumber = Asc(Mid(letterToReturn, 1, 1)) - 64
        Next i
        ActivePresentation.Slides(2).Shapes("Letter" & letterNumber).TextFrame.TextRange.Text = letterToReturn
    End If
End Sub

Sub guessBonusLetters()
    If ActivePresentation.Slides(2).Shapes("BonusOutline").Line.Transparency = 0 And ActivePresentation.Slides(2).Shapes("Letter1").Visible = True Then
        If ActivePresentation.Slides(2).Shapes("BonusLetters").TextFrame.TextRange.Text = "" Then
            MsgBox ("No letters found. Use the letter selector to input the letters the contestant chooses for the bonus round.")
            Exit Sub
        End If
        Dim letterExist As Boolean
        letterExist = False
        For i = 1 To Len(ActivePresentation.Slides(2).Shapes("BonusLetters").TextFrame.TextRange.Text)
            letterToGuess = Mid(ActivePresentation.Slides(2).Shapes("BonusLetters").TextFrame.TextRange.Text, i, 1)
            For k = 1 To 52
                If ActivePresentation.Slides(2).Shapes("PuzzleCache" & k).TextFrame.TextRange.Text = letterToGuess Then
                    If ActivePresentation.Slides(9).Shapes("BlueTiles").TextFrame.TextRange.Text = "off" Then
                        ActivePresentation.Slides(2).Shapes("PuzzleBoard" & k).TextFrame.TextRange.Text = letterToGuess
                    Else:
                        ActivePresentation.Slides(2).Shapes("PuzzleBoard" & k).Fill.ForeColor.RGB = RGB(0, 0, 255)
                    End If
                    letterExist = True
                End If
            Next k
        Next i
        If letterExist = False Then
            ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = True
        Else:
            ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = False
        End If
        ActivePresentation.Slides(2).Shapes("BonusBox").Fill.ForeColor.RGB = RGB(166, 166, 166)
        ActivePresentation.Slides(2).Shapes("BonusOutline").Line.Transparency = 1
    End If
End Sub

Sub toggleRound(oClickedShape As Shape)
  Dim oSh As Shape
    Dim sText As String
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    If oSh.Name = "Fourth Round" Then
        toggleBonusRound (True)
    Else:
        toggleBonusRound (False)
    End If
End Sub

Private Sub toggleBonusRound(i)
    ActivePresentation.Slides(2).Shapes("BonusTimerCircle").Visible = i
    ActivePresentation.Slides(2).Shapes("BonusTimerOverlay").Visible = i
    ActivePresentation.Slides(2).Shapes("BonusBox").Visible = i
    ActivePresentation.Slides(2).Shapes("RSTLNEBox").Visible = i
    ActivePresentation.Slides(2).Shapes("BonusOutline").Visible = i
    ActivePresentation.Slides(2).Shapes("RSTLNEOutline").Visible = i
    ActivePresentation.Slides(2).Shapes("RSTLNE").Visible = i
    ActivePresentation.Slides(2).Shapes("HelpBonus").Visible = i
    ActivePresentation.Slides(2).Shapes("ResetBonus").Visible = i
    ActivePresentation.Slides(2).Shapes("BonusLetters").Visible = i
    ActivePresentation.Slides(2).Shapes("Bonus" & 10).Visible = i
    ActivePresentation.Slides(2).Shapes("BonusBox").TextFrame.TextRange.Text = ""
End Sub

Sub bonusRoundBlock()
    MsgBox ("Please reset the bonus round timer before switching rounds.")
End Sub

Sub randomSpin()
    If ActivePresentation.SlideShowWindow.View.Slide.Shapes("BackOval").Visible = True Then
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("BackOval").Visible = False
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("BackText").Visible = False
        Exit Sub
    Else:
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("BackOval").Visible = True
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("BackText").Visible = True
        Dim rand As Integer
        Randomize
        rand = Int((3599 + 1) * Rnd)
        Dim realRand As Double
        realRand = rand / 10
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("TheWheel").Rotation = realRand
    End If
End Sub

Sub puzzleSwapChoose(oSh As Shape)
    Dim i As Integer
    Dim j As Boolean
    For i = 1 To 12
        If ActivePresentation.Slides(10).Shapes("Swap" & i).Name = oSh.Name Then
            j = True
            Exit For
        End If
    Next i
    If j = True Then
        If ActivePresentation.Slides(10).Shapes("Swap" & i).Fill.ForeColor.RGB = RGB(255, 255, 0) Then
            cancelSwap
        ElseIf ActivePresentation.Slides(10).Shapes("Swap" & i).Rotation = 180 Then
            puzzleSwap (i)
        Else:
            ActivePresentation.Slides(10).Shapes("Swap" & i).Fill.ForeColor.RGB = RGB(255, 255, 0)
            ActivePresentation.Slides(10).Shapes("SwapAlert").TextFrame.TextRange.Text = "                   Swap puzzle " & i & " with…"
            ActivePresentation.Slides(10).Shapes("SwapAlert").Visible = True
            ActivePresentation.Slides(10).Shapes("SwapCancel").Visible = True
            ActivePresentation.Slides(10).Shapes("GoBack").Visible = False
            For k = 1 To 12
                If k <> i Then
                    ActivePresentation.Slides(10).Shapes("Swap" & k).Rotation = 180
                End If
            Next k
            For l = 1 To 3
                ActivePresentation.Slides(10).Shapes("Blocker" & l).Visible = True
            Next l
        End If
    End If
End Sub

Sub cancelSwap()
    For k = 1 To 12
        ActivePresentation.Slides(10).Shapes("Swap" & k).Rotation = 0
        ActivePresentation.Slides(10).Shapes("Swap" & k).Fill.ForeColor.RGB = RGB(113, 168, 213)
    Next k
    For l = 1 To 3
        ActivePresentation.Slides(10).Shapes("Blocker" & l).Visible = False
    Next l
    ActivePresentation.Slides(10).Shapes("GoBack").Visible = True
    ActivePresentation.Slides(10).Shapes("SwapAlert").Visible = False
    ActivePresentation.Slides(10).Shapes("SwapCancel").Visible = False
    ActivePresentation.Slides(10).Shapes("SwapAlert").TextFrame.TextRange.Text = ""
End Sub

Private Sub puzzleSwap(i)
    For k = 1 To 12
        If ActivePresentation.Slides(10).Shapes("Swap" & k).Rotation = 0 Then
            Exit For
        End If
    Next k
    For l = 1 To 52
        ActivePresentation.Slides(10).Shapes("PuzzleSolutionSwap-" & CStr(l)).TextFrame.TextRange.Text = ActivePresentation.Slides(10).Shapes("PuzzleSolution" & CStr(k) & "-" & CStr(l)).TextFrame.TextRange.Text
        ActivePresentation.Slides(10).Shapes("PuzzleSolutionSwap-" & CStr(l)).Fill.ForeColor.RGB = ActivePresentation.Slides(10).Shapes("PuzzleSolution" & CStr(k) & "-" & CStr(l)).Fill.ForeColor.RGB
    Next l
    ActivePresentation.Slides(10).Shapes("PuzzleCategorySwap").TextFrame.TextRange.Text = ActivePresentation.Slides(10).Shapes("PuzzleCategory" & CStr(k)).TextFrame.TextRange.Text
    For m = 1 To 52
        ActivePresentation.Slides(10).Shapes("PuzzleSolution" & CStr(k) & "-" & CStr(m)).TextFrame.TextRange.Text = ActivePresentation.Slides(10).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(m)).TextFrame.TextRange.Text
        ActivePresentation.Slides(10).Shapes("PuzzleSolution" & CStr(k) & "-" & CStr(m)).Fill.ForeColor.RGB = ActivePresentation.Slides(10).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(m)).Fill.ForeColor.RGB
    Next m
    ActivePresentation.Slides(10).Shapes("PuzzleCategory" & CStr(k)).TextFrame.TextRange.Text = ActivePresentation.Slides(10).Shapes("PuzzleCategory" & CStr(i)).TextFrame.TextRange.Text
    For n = 1 To 52
        ActivePresentation.Slides(10).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(n)).TextFrame.TextRange.Text = ActivePresentation.Slides(10).Shapes("PuzzleSolutionSwap-" & CStr(n)).TextFrame.TextRange.Text
        ActivePresentation.Slides(10).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(n)).Fill.ForeColor.RGB = ActivePresentation.Slides(10).Shapes("PuzzleSolutionSwap-" & CStr(n)).Fill.ForeColor.RGB
    Next n
    ActivePresentation.Slides(10).Shapes("PuzzleCategory" & CStr(i)).TextFrame.TextRange.Text = ActivePresentation.Slides(10).Shapes("PuzzleCategorySwap").TextFrame.TextRange.Text
    For o = 1 To 52
        ActivePresentation.Slides(10).Shapes("PuzzleSolutionSwap-" & CStr(o)).TextFrame.TextRange.Text = ""
        ActivePresentation.Slides(10).Shapes("PuzzleSolutionSwap-" & CStr(o)).Fill.ForeColor.RGB = RGB(24, 154, 80)
    Next o
    ActivePresentation.Slides(10).Shapes("PuzzleCategorySwap").TextFrame.TextRange.Text = ""
    cancelSwap
    shadeOccupiedPuzzles
End Sub
