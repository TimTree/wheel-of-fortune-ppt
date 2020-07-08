Attribute VB_Name = "Module1"
Option Explicit

Sub goToHowToUse()
    ' Allows slide to advance if macros are enabled
    ActivePresentation.Slides(1).Shapes("MacroDisabledText").Visible = False
    savePuzzle
    shadeOccupiedPuzzles
    SlideShowWindows(1).View.GotoSlide 17 + ActivePresentation.SectionProperties.SlidesCount(4)
End Sub

Sub goToSetUp()
    ' Checks if macros are enabled
    ActivePresentation.Slides(1).Shapes("MacroDisabledText").Visible = False
    savePuzzle
    shadeOccupiedPuzzles
    SlideShowWindows(1).View.GotoSlide 7
End Sub

Sub goToPuzzleBoard()
    ' Checks if macros are enabled
    ActivePresentation.Slides(1).Shapes("MacroDisabledText").Visible = False
    savePuzzle
    shadeOccupiedPuzzles
    SlideShowWindows(1).View.GotoSlide 2
End Sub

Sub BGChange()
    Dim i As Integer
    Dim themeNumber
    Set themeNumber = ActivePresentation.Slides(9).Shapes("Backdrop").TextFrame.TextRange
    Dim puzzleBoardGradientBottom As Long, puzzleBoardGradientMiddle As Long, puzzleBoardGradientTop As Long
    Dim setFloorEdges As Long, setFloorMiddle As Long, setFloorLine As Long, wheelGradientMiddle As Long
    Dim wheelGradientTop As Long, categoryColor As Long, letterSelectorColor As Long
    If themeNumber.Text = "studio" Then
        puzzleBoardGradientBottom = RGB(2, 127, 190)
        puzzleBoardGradientMiddle = RGB(94, 189, 208)
        puzzleBoardGradientTop = RGB(2, 127, 190)
        setFloorEdges = RGB(0, 51, 0)
        setFloorMiddle = RGB(0, 153, 0)
        setFloorLine = RGB(38, 100, 38)
        wheelGradientMiddle = RGB(20, 121, 152)
        wheelGradientTop = RGB(3, 96, 143)
        categoryColor = RGB(27, 91, 33)
        letterSelectorColor = RGB(185, 205, 229)
        themeNumber.Text = "stadium"
    ElseIf themeNumber.Text = "stadium" Then
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
        themeNumber.Text = "valentine's"
    ElseIf themeNumber.Text = "valentine's" Then
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
        themeNumber.Text = "blackout"
    Else:
        setFloorEdges = RGB(41, 38, 35)
        setFloorMiddle = RGB(125, 73, 126)
        setFloorLine = RGB(16, 37, 63)
        categoryColor = RGB(23, 55, 94)
        letterSelectorColor = RGB(79, 129, 189)
        themeNumber.Text = "studio"
    End If
    With ActivePresentation.Slides(2)
        With .Shapes("BackDrop")
            If themeNumber.Text = "studio" Then
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
        With .Shapes("ValuePanel")
            .Fill.ForeColor.RGB = letterSelectorColor
        End With
    End With
    For i = 3 To 6
        With ActivePresentation.Slides(i)
            With .Shapes("BackDrop")
                If themeNumber.Text = "studio" Then
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
End Sub

Sub ClearBoardButton()
    Dim i As Integer, j As Integer
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
    ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = False
    ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text = "WHEEL OF FORTUNE"
    resetBonusRound
    disableFinalSpin
End Sub

Sub editPlayerName(oClickedShape As Shape)
    Dim oSh As Shape, sText As String, i As Integer, j As Boolean
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    For i = 1 To 4
        If ActivePresentation.Slides(2).Shapes("Player" & i & "Name").Name = oSh.Name Then
            j = True
            Exit For
        End If
    Next i
    If j = True Then
        sText = InputBox("Edit " & ActivePresentation.Slides(2).Shapes("Player" & i & "Name").TextFrame.TextRange.Text & "'s name:", "Edit Player Name", oSh.TextFrame.TextRange.Text)
        If sText = "" Then
        Else:
        oSh.TextFrame.TextRange.Text = sText
        End If
    End If
End Sub

Sub PlayerRoundDollarSign(oClickedShape As Shape)
    Dim oSh As Shape, sText As String, i As Integer, j As Boolean
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    For i = 1 To 4
        If ActivePresentation.Slides(2).Shapes("Player" & i & "RoundDollarSign").Name = oSh.Name Then
            j = True
            Exit For
        End If
    Next i
    If j = True Then
        sText = InputBox("Manually edit " & ActivePresentation.Slides(2).Shapes("Player" & i & "Name").TextFrame.TextRange.Text & "'s round score:", "Manually Edit Round Score", ActivePresentation.Slides(2).Shapes("Player" & i & "RoundScore").TextFrame.TextRange.Text)
        While IsNumeric(sText) = False And sText <> ""
            sText = InputBox("You can only enter numbers here. Try again:", "Manually Edit Round Score", sText)
        Wend
        If sText = "" Then
        Else:
            ActivePresentation.Slides(2).Shapes("Player" & i & "RoundScore").TextFrame.TextRange.Text = CLng(sText)
        End If
    End If
End Sub

Sub PlayerTotalsDollarSign(oClickedShape As Shape)
    Dim oSh As Shape, sText As String, i As Integer, j As Boolean
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    For i = 1 To 4
        If ActivePresentation.Slides(2).Shapes("Player" & i & "TotalsDollarSign").Name = oSh.Name Then
            j = True
            Exit For
        End If
    Next i
    If j = True Then
        sText = InputBox("Manually edit " & ActivePresentation.Slides(2).Shapes("Player" & i & "Name").TextFrame.TextRange.Text & "'s totals score:", "Manually Edit Totals Score", ActivePresentation.Slides(2).Shapes("Player" & i & "TotalsScore").TextFrame.TextRange.Text)
        While IsNumeric(sText) = False And sText <> ""
            sText = InputBox("You can only enter numbers here. Try again:", "Manually Edit Totals Score", sText)
        Wend
        If sText = "" Then
        Else:
            ActivePresentation.Slides(2).Shapes("Player" & i & "TotalsScore").TextFrame.TextRange.Text = CLng(sText)
        End If
    End If
End Sub

Sub PlayerAddFromValuePanel(oClickedShape As Shape)
    Dim oSh As Shape, sText As String, i As Integer, j As Boolean
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    For i = 1 To 4
        If ActivePresentation.Slides(2).Shapes("Player" & i & "RoundScoreCompatibility").Name = oSh.Name Then
            j = True
            Exit For
        End If
    Next i
    If j = True Then
        If ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text = "WHEEL OF FORTUNE" Then
            MsgBox "During a game, click here to add the amount shown on the value panel" & vbNewLine & _
            "(currently reads WHEEL OF FORTUNE) to " & ActivePresentation.Slides(2).Shapes("Player" & i & "Name").TextFrame.TextRange.Text & "'s round score." _
            , 0, "Add to " & ActivePresentation.Slides(2).Shapes("Player" & i & "Name").TextFrame.TextRange.Text & "'s Round Score"
            Exit Sub
        Else:
            If ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = "" Then
                MsgBox "There's nothing to add because the value panel is empty." & vbNewLine & _
                "Spin the wheel or manually set the spun wheel value on the Value Panel first." _
                , 0, "Add to " & ActivePresentation.Slides(2).Shapes("Player" & i & "Name").TextFrame.TextRange.Text & "'s Round Score"
            Else:
                Dim effectiveWheelValue As Long
                If ActivePresentation.Slides(2).Shapes("FinalSpinButton").Fill.ForeColor.RGB = RGB(225, 129, 75) Then
                    effectiveWheelValue = CLng(ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text) + 1000
                Else:
                    effectiveWheelValue = CLng(ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text)
                End If
                If ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = "" Then
                    ActivePresentation.Slides(2).Shapes("Player" & i & "RoundScore").TextFrame.TextRange.Text = _
                    CLng(ActivePresentation.Slides(2).Shapes("Player" & i & "RoundScore").TextFrame.TextRange.Text) + _
                    effectiveWheelValue
                Else:
                    If ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = "10000" Then
                        ActivePresentation.Slides(2).Shapes("Player" & i & "RoundScore").TextFrame.TextRange.Text = _
                        CLng(ActivePresentation.Slides(2).Shapes("Player" & i & "RoundScore").TextFrame.TextRange.Text) + _
                        effectiveWheelValue
                    Else:
                        ActivePresentation.Slides(2).Shapes("Player" & i & "RoundScore").TextFrame.TextRange.Text = _
                        CLng(ActivePresentation.Slides(2).Shapes("Player" & i & "RoundScore").TextFrame.TextRange.Text) + _
                        CLng(ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text) * _
                        effectiveWheelValue
                    End If
                End If
            ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = ""
            setValuePanelDisplay
            End If
        End If
    End If
End Sub

Sub PlayerBuyaVowel(oSh As Shape)
    Dim i As Integer, j As Boolean, RoundDollarAmount, playerName, VOWELCOST As Long
    VOWELCOST = CLng(ActivePresentation.Slides(9).Shapes("VowelPrice").TextFrame.TextRange.Text)
    For i = 1 To 4
        If ActivePresentation.Slides(2).Shapes("Player" & i & "BuyVowelButton").Name = oSh.Name Then
            j = True
            Exit For
        End If
    Next i
    If j = True Then
        Set RoundDollarAmount = ActivePresentation.Slides(2).Shapes("Player" & i & "RoundScore")
        Set playerName = ActivePresentation.Slides(2).Shapes("Player" & i & "Name")
        If IsNumeric(RoundDollarAmount.TextFrame.TextRange.Text) = False Then
            MsgBox "Please remove non-numeric characters from " + playerName.TextFrame.TextRange.Text + "'s score before buying a vowel." _
            , 0, playerName.TextFrame.TextRange.Text + " Buy Vowel"
        ElseIf CLng(RoundDollarAmount.TextFrame.TextRange.Text) < VOWELCOST Then
            MsgBox playerName.TextFrame.TextRange.Text + " cannot buy a vowel. Vowels cost $" + CStr(VOWELCOST) + "." _
            , 0, playerName.TextFrame.TextRange.Text + " Buy Vowel"
        ElseIf CLng(RoundDollarAmount.TextFrame.TextRange.Text) >= VOWELCOST Then
            RoundDollarAmount.TextFrame.TextRange.Text = (CLng(RoundDollarAmount.TextFrame.TextRange.Text) - VOWELCOST)
        End If
    End If
End Sub

Sub PlayerTransferTotals(oSh As Shape)
    Dim i As Integer, j As Boolean, RoundDollarAmount, TotalsDollarAmount, HOUSEMINIMUM As Long, shouldIHouse
    HOUSEMINIMUM = CLng(ActivePresentation.Slides(9).Shapes("HouseMinimum").TextFrame.TextRange.Text)
    For i = 1 To 4
        If ActivePresentation.Slides(2).Shapes("Player" & i & "TransferTotalsButton").Name = oSh.Name Then
            j = True
            Exit For
        End If
    Next i
    If j = True Then
        Set RoundDollarAmount = ActivePresentation.Slides(2).Shapes("Player" & i & "RoundScore")
        Set TotalsDollarAmount = ActivePresentation.Slides(2).Shapes("Player" & i & "TotalsScore")
        If IsNumeric(TotalsDollarAmount.TextFrame.TextRange.Text) = True And IsNumeric(RoundDollarAmount.TextFrame.TextRange.Text) = True Then
            If CLng(RoundDollarAmount.TextFrame.TextRange.Text) < HOUSEMINIMUM Then
                shouldIHouse = MsgBox("The house minimum of $" + CStr(HOUSEMINIMUM) + " will be transferred.", vbOKCancel, _
                "Confirm House Minimum Transfer")
                If shouldIHouse = vbOK Then
                    TotalsDollarAmount.TextFrame.TextRange.Text = CLng(TotalsDollarAmount.TextFrame.TextRange.Text) + HOUSEMINIMUM
                    wipeRoundScores
                Else:
                    Exit Sub
                End If
            Else:
                TotalsDollarAmount.TextFrame.TextRange.Text = CLng(TotalsDollarAmount.TextFrame.TextRange.Text) + CLng(RoundDollarAmount.TextFrame.TextRange.Text)
                wipeRoundScores
            End If
        End If
    End If
End Sub

Sub PlayerXButton(oSh As Shape)
    Dim i As Integer, j As Boolean, RoundDollarAmount
    For i = 1 To 4
        If ActivePresentation.Slides(2).Shapes("Player" & i & "XButtonCompatibility").Name = oSh.Name Then
            j = True
            Exit For
        End If
    Next i
    If j = True Then
        Set RoundDollarAmount = ActivePresentation.Slides(2).Shapes("Player" & i & "RoundScore")
        RoundDollarAmount.TextFrame.TextRange.Text = 0
    End If
End Sub

Sub TogglePlayerItem(oSh As Shape)
    If oSh.Fill.Transparency = 1 Then
        oSh.Fill.Transparency = 0
    Else:
        oSh.Fill.Transparency = 1
    End If
End Sub

Sub DetermineMystery()
    Dim randomNumber As Integer
    ActivePresentation.Slides(4).Shapes("MysteryIndicator").TextFrame.TextRange.Text = ""
    wait (0.1)
    Randomize
    randomNumber = Int(2 * Rnd) + 1
    If randomNumber = 2 Then
        ActivePresentation.Slides(4).Shapes("MysteryIndicator").TextFrame.TextRange.Text = "$10,000"
        ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = "10000"
        ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = ""
        setValuePanelDisplay
    Else:
        ActivePresentation.Slides(4).Shapes("MysteryIndicator").TextFrame.TextRange.Text = "Bankrupt"
        ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = ""
        ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = ""
        setValuePanelDisplay
    End If
    If ActivePresentation.Slides(4).Shapes("WheelValue").TextFrame.TextRange.Text = "Mystery 1" Then
        ActivePresentation.Slides(4).Shapes("TheWheel").GroupItems("MysteryWedge1").Fill.Transparency = 1
    ElseIf ActivePresentation.Slides(4).Shapes("WheelValue").TextFrame.TextRange.Text = "Mystery 2" Then
        ActivePresentation.Slides(4).Shapes("TheWheel").GroupItems("MysteryWedge2").Fill.Transparency = 1
    End If
End Sub

Sub ClearMysteryIndicator()
    ActivePresentation.Slides(4).Shapes("MysteryIndicator").TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(4).Shapes("TheWheel").GroupItems("MysteryWedge1").Fill.Transparency = 0
    ActivePresentation.Slides(4).Shapes("TheWheel").GroupItems("MysteryWedge2").Fill.Transparency = 0
End Sub

Sub ErasePuzzleRow(oSh As Shape)
    Dim i As Integer, minim As Integer, maxim As Integer
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
        SlideShowWindows(1).View.Slide.Shapes("SetUpPuzzle" & CStr(i)).TextFrame.TextRange.Text = ""
        SlideShowWindows(1).View.Slide.Shapes("SetUpPuzzle" & CStr(i)).Fill.ForeColor.RGB = RGB(24, 154, 80)
    Next i
End Sub

Sub EraseEntirePuzzle()
    Dim i As Integer, shouldIEraseAll
    shouldIEraseAll = MsgBox("Are you sure you want to delete the entire puzzle?", vbYesNo + vbDefaultButton2, "Confirm Puzzle Delete")
    If shouldIEraseAll = vbYes Then
    For i = 1 To 52
        SlideShowWindows(1).View.Slide.Shapes("SetUpPuzzle" & CStr(i)).TextFrame.TextRange.Text = ""
        SlideShowWindows(1).View.Slide.Shapes("SetUpPuzzle" & CStr(i)).Fill.ForeColor.RGB = RGB(24, 154, 80)
    Next i
        SlideShowWindows(1).View.Slide.Shapes("SetUpPuzzleCategory").TextFrame.TextRange.Text = ""
    Else
        Exit Sub
    End If
End Sub

Sub puzzleSetupFromOtherSlide(oClickedShape As Shape)
    On Error GoTo errHandler
    Dim oSh As Shape
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    savePuzzle
    placePuzzleToSetUp (CInt(oSh.TextFrame.TextRange.Text))
    shadeOccupiedPuzzles
    highlightCurrentPuzzle (CInt(oSh.TextFrame.TextRange.Text))
    SlideShowWindows(1).View.GotoSlide 8
    Exit Sub
errHandler:
    MsgBox "Cannot edit puzzles because ActiveX components are disabled." & vbNewLine & _
    "If you use Windows, check if ActiveX is enabled in Trust Center settings." & vbNewLine & _
    "If you use macOS, download the Mac version of Wheel of Fortune for PowerPoint.", 0, "Set Up Puzzles Error"
End Sub

Sub puzzleSetup(oClickedShape As Shape)
    Dim oSh As Shape
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    savePuzzle
    placePuzzleToSetUp (CInt(oSh.TextFrame.TextRange.Text))
    shadeOccupiedPuzzles
    highlightCurrentPuzzle (CInt(oSh.TextFrame.TextRange.Text))
End Sub

Sub puzzleSetupJump(num As Integer)
    placePuzzleToSetUp (num)
    shadeOccupiedPuzzles
    highlightCurrentPuzzle (num)
    SlideShowWindows(1).View.GotoSlide 8
End Sub

Sub PuzzleSetupFromAllPuzzles(oClickedShape As Shape)
    Dim oSh As Shape
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    exactPuzzleRow (Int((CInt(oSh.TextFrame.TextRange.Text) - 1) / 12))
    puzzleSetupJump (CInt(oSh.TextFrame.TextRange.Text))
End Sub

Sub EditVowelPrice(oClickedShape As Shape)
    On Error GoTo errHandler
    Dim oSh As Shape, sText As String, newText As Long
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    sText = InputBox("Edit the vowel price. The default price is $250.", "Edit Vowel Price", CLng(oSh.TextFrame.TextRange.Text))
    While IsNumeric(sText) = False And sText <> ""
        sText = InputBox("You can only enter numbers here. Try again:", "Edit Vowel Price", sText)
    Wend
    If sText = "" Then
        Exit Sub
    Else:
        newText = CLng(sText)
        oSh.TextFrame.TextRange.Text = "$" & newText
    End If
    Exit Sub
errHandler:
    MsgBox "The vowel price cannot exceed $2147483647.", 0, "Edit Vowel Price Error"
End Sub

Sub EditHouseMinimum(oClickedShape As Shape)
    On Error GoTo errHandler
    Dim oSh As Shape, sText As String, newText As Long
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    sText = InputBox("Edit the house minimum. The default minimum is $1000.", "Edit House Minimum", CLng(oSh.TextFrame.TextRange.Text))
    While IsNumeric(sText) = False And sText <> ""
    sText = InputBox("You can only enter numbers here. Try again:", "Edit House Minimum", sText)
    Wend
    If sText = "" Then
        Exit Sub
    Else:
    newText = CLng(sText)
    oSh.TextFrame.TextRange.Text = "$" & newText
    End If
    Exit Sub
errHandler:
    MsgBox "The house minimum cannot exceed $2147483647.", 0, "Edit House Minimum Error"
End Sub

Sub confirmDeleteAllPuzzles()
    Dim deletionConfirm
    deletionConfirm = MsgBox("Are you ABSOLUTELY sure you want to delete all puzzles?", vbYesNo + vbDefaultButton2, "Confirm Delete All Puzzles")
    If deletionConfirm = vbYes Then
        deleteAllPuzzles
        MsgBox "Successfully deleted all puzzles.", 0, "Confirm Delete All Puzzles"
    Else
        Exit Sub
    End If
End Sub

Private Sub highlightCurrentPuzzle(i As Integer)
    With ActivePresentation.Slides(8).Shapes("LinkTo" & i)
        .ZOrder msoBringToFront
        .Fill.ForeColor.RGB = RGB(250, 192, 144)
        .Line.ForeColor.RGB = RGB(228, 108, 10)
    End With
End Sub

Private Sub savePuzzle()
    Dim thereWasAPuzzle As Boolean, PuzzleIndex As Integer, i As Integer, j As Integer
    thereWasAPuzzle = False
    PuzzleIndex = CInt(ActivePresentation.Slides(7).Shapes("CurrentPuzzleRowIndex").TextFrame.TextRange.Text)
    For i = 1 To 12
        If ActivePresentation.Slides(8).Shapes("LinkTo" & CStr(i + (12 * PuzzleIndex))).Fill.ForeColor.RGB = RGB(250, 192, 144) Then
            thereWasAPuzzle = True
            Exit For
        End If
    Next i
    If thereWasAPuzzle = True Then
        For j = 1 To 52
            ActivePresentation.Slides(11 + PuzzleIndex).Shapes("PuzzleSolution" & CStr(i + (12 * PuzzleIndex)) & "-" & CStr(j)).TextFrame.TextRange.Text = ActivePresentation.Slides(8).Shapes("SetUpPuzzle" & CStr(j)).TextFrame.TextRange.Text
            ActivePresentation.Slides(11 + PuzzleIndex).Shapes("PuzzleSolution" & CStr(i + (12 * PuzzleIndex)) & "-" & CStr(j)).Fill.ForeColor.RGB = ActivePresentation.Slides(8).Shapes("SetUpPuzzle" & CStr(j)).Fill.ForeColor.RGB
        Next j
        ActivePresentation.Slides(11 + PuzzleIndex).Shapes("PuzzleCategory" & CStr(i + (12 * PuzzleIndex))).TextFrame.TextRange.Text = ActivePresentation.Slides(8).Shapes("SetUpPuzzleCategory").TextFrame.TextRange.Text
    End If
End Sub

Private Sub placePuzzleToSetUp(i As Integer)
    Dim PuzzleIndex As Integer, n As Integer
    PuzzleIndex = CInt(ActivePresentation.Slides(7).Shapes("CurrentPuzzleRowIndex").TextFrame.TextRange.Text)
    For n = 1 To 52
        ActivePresentation.Slides(8).Shapes("SetUpPuzzle" & CStr(n)).TextFrame.TextRange.Text = ActivePresentation.Slides(11 + PuzzleIndex).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(n)).TextFrame.TextRange.Text
        ActivePresentation.Slides(8).Shapes("SetUpPuzzle" & CStr(n)).Fill.ForeColor.RGB = ActivePresentation.Slides(11 + PuzzleIndex).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(n)).Fill.ForeColor.RGB
    Next n
    ActivePresentation.Slides(8).Shapes("SetUpPuzzleCategory").TextFrame.TextRange.Text = ActivePresentation.Slides(11 + PuzzleIndex).Shapes("PuzzleCategory" & CStr(i)).TextFrame.TextRange.Text
End Sub

Private Sub deleteAllPuzzles()
    Dim s As Integer, i As Integer, j As Integer, k As Integer
    s = 11 + ActivePresentation.SectionProperties.SlidesCount(4) - 1
    Do Until s = 11
        ActivePresentation.Slides(s).Delete
        s = 11 + ActivePresentation.SectionProperties.SlidesCount(4) - 1
    Loop
    ActivePresentation.Slides(11).Shapes("NextAllPuzzles").Visible = msoFalse
    For i = 1 To 12
        For j = 1 To 52
            ActivePresentation.Slides(11).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(j)).TextFrame.TextRange.Text = ""
            ActivePresentation.Slides(11).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(j)).Fill.ForeColor.RGB = RGB(24, 154, 80)
        Next j
        ActivePresentation.Slides(11).Shapes("PuzzleCategory" & CStr(i)).TextFrame.TextRange.Text = ""
    Next i
    For k = 1 To 52
        ActivePresentation.Slides(8).Shapes("SetupPuzzle" & CStr(k)).TextFrame.TextRange.Text = ""
        ActivePresentation.Slides(8).Shapes("SetUpPuzzle" & CStr(k)).Fill.ForeColor.RGB = RGB(24, 154, 80)
    Next k
    ActivePresentation.Slides(8).Shapes("SetUpPuzzleCategory").TextFrame.TextRange.Text = ""
    exactPuzzleRow (0)
    shadeOccupiedPuzzles
End Sub

Private Sub shadeOccupiedPuzzles()
    Dim PuzzleIndex As Integer, p As Integer, q As Integer, r As Integer, blankPuzzle As Boolean
    PuzzleIndex = CInt(ActivePresentation.Slides(7).Shapes("CurrentPuzzleRowIndex").TextFrame.TextRange.Text)
    For p = 1 To 12
        blankPuzzle = True
        For q = 1 To 52
            If ActivePresentation.Slides(11 + PuzzleIndex).Shapes("PuzzleSolution" & CStr(p + (12 * PuzzleIndex)) & "-" & CStr(q)).TextFrame.TextRange.Text <> "" Then
                blankPuzzle = False
                Exit For
            End If
        Next q
        If ActivePresentation.Slides(11 + PuzzleIndex).Shapes("PuzzleCategory" & CStr(p + (12 * PuzzleIndex))).TextFrame.TextRange.Text <> "" Then
            blankPuzzle = False
        End If
        For r = 7 To 9
           If blankPuzzle = False Then
               With ActivePresentation.Slides(r).Shapes("LinkTo" & CStr(p + (12 * PuzzleIndex)))
                   .Fill.ForeColor.RGB = RGB(146, 224, 204)
                   .Line.ForeColor.RGB = RGB(55, 96, 146)
               End With
           Else:
               With ActivePresentation.Slides(r).Shapes("LinkTo" & CStr(p + (12 * PuzzleIndex)))
                   .Fill.ForeColor.RGB = RGB(149, 179, 215)
                   .Line.ForeColor.RGB = RGB(55, 96, 146)
               End With
           End If
        Next r
    Next p
End Sub

Private Sub removeWheelAnimations()
    Dim j As Integer, x As Integer
    For j = 3 To 6
        For x = ActivePresentation.Slides(j).TimeLine.MainSequence.Count To 1 Step -1
            ActivePresentation.Slides(j).TimeLine.MainSequence.Item(x).Delete
        Next x
    Next j
End Sub

Sub OnSlideShowTerminate(oWn As SlideShowWindow)
    removeWheelAnimations
    savePuzzle
    shadeOccupiedPuzzles
    resetBonusRound
    toggleBonusRound (False)
    toggleFourthRound (False)
    ActivePresentation.Slides(1).Shapes("MacroDisabledText").Visible = True
End Sub

Sub goToHowToUseFromSetUpPuzzles()
    savePuzzle
    shadeOccupiedPuzzles
    SlideShowWindows(1).View.GotoSlide 17 + ActivePresentation.SectionProperties.SlidesCount(4)
End Sub

Sub goToPuzzleBoardFromSetUpPuzzles()
    savePuzzle
    shadeOccupiedPuzzles
    SlideShowWindows(1).View.GotoSlide 2
End Sub

Sub goToSettingsFromSetUpPuzzles()
    savePuzzle
    shadeOccupiedPuzzles
    SlideShowWindows(1).View.GotoSlide 9
End Sub

Private Function puzzleExists(i As Integer) As Boolean
    Dim PuzzleNumberIndex As Integer, puzzleBoolean As Boolean, m As Integer
    PuzzleNumberIndex = Int((i - 1) / 12)
    puzzleBoolean = False
    If PuzzleNumberIndex + 1 > ActivePresentation.SectionProperties.SlidesCount(4) Then
        puzzleExists = puzzleBoolean
        Exit Function
    End If
    For m = 1 To 52
        If ActivePresentation.Slides(11 + PuzzleNumberIndex).Shapes("PuzzleSolution" & i & "-" & m).Fill.ForeColor.RGB = RGB(255, 255, 255) Then
            puzzleBoolean = True
            Exit For
        End If
    Next m
    If ActivePresentation.Slides(11 + PuzzleNumberIndex).Shapes("PuzzleCategory" & i).TextFrame.TextRange.Text <> "" Then
        puzzleBoolean = True
    End If
    puzzleExists = puzzleBoolean
End Function

Sub LoadPuzzleOrSolve()
    If ActivePresentation.Slides(2).Shapes("LeftTab").TextFrame.TextRange.Text = "Load Puzzle" Then
        Dim noPuzzlesExist As Boolean, numAllPuzzlesSlides As Integer, numberToLoad, j As Integer, m As Integer
        noPuzzlesExist = True
        numAllPuzzlesSlides = ActivePresentation.SectionProperties.SlidesCount(4)
        For m = 0 To numAllPuzzlesSlides - 1
            For j = (1 + (12 * m)) To (12 + (12 * m))
                If puzzleExists(j) = True Then
                    noPuzzlesExist = False
                    Exit For
                End If
            Next j
        Next m
        If noPuzzlesExist = True Then
            MsgBox "No puzzles were found. Create puzzles using Set Up Puzzles on the top right of this slide.", 0, "Load Puzzle"
            Exit Sub
        End If
        numberToLoad = InputBox("Enter the puzzle number to load:", "Load Puzzle")
        While IsNumeric(numberToLoad) = False:
            If numberToLoad = "" Then
                Exit Sub
            Else:
                numberToLoad = InputBox("Please enter a number:", "Load Puzzle", numberToLoad)
            End If
        Wend
            loadPuzzle (CInt(numberToLoad))
    Else:
        Dim alreadySolved As Boolean, solveConfirm, i As Integer
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
        solveConfirm = MsgBox("Are you sure you want to reveal the puzzle?", vbYesNo + vbDefaultButton1, "Confirm Puzzle Solve")
            If solveConfirm = vbYes Then
                solvePuzzle
            Else
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub loadPuzzle(i As Integer)
    On Error GoTo errHandler
    If puzzleExists(i) = False Then
        MsgBox "No puzzle found for number " & i & ".", 0, "Load Puzzle Error"
        Exit Sub
    End If
    ClearBoardButton
    Dim PuzzleNumberIndex As Integer, j As Integer, k As Integer
    PuzzleNumberIndex = Int((i - 1) / 12)
    For j = 1 To 52
        ActivePresentation.Slides(2).Shapes("PuzzleBoard" & j).Fill.ForeColor.RGB = ActivePresentation.Slides(11 + PuzzleNumberIndex).Shapes("PuzzleSolution" & i & "-" & j).Fill.ForeColor.RGB
        ActivePresentation.Slides(2).Shapes("PuzzleCache" & j).TextFrame.TextRange.Text = ActivePresentation.Slides(11 + PuzzleNumberIndex).Shapes("PuzzleSolution" & i & "-" & j).TextFrame.TextRange.Text
        If isLetter(ActivePresentation.Slides(11 + PuzzleNumberIndex).Shapes("PuzzleSolution" & i & "-" & j).TextFrame.TextRange.Text) = False Then
            ActivePresentation.Slides(2).Shapes("PuzzleBoard" & j).TextFrame.TextRange.Text = ActivePresentation.Slides(11 + PuzzleNumberIndex).Shapes("PuzzleSolution" & i & "-" & j).TextFrame.TextRange.Text
        End If
    Next j
    ActivePresentation.Slides(2).Shapes("CategoryBox").TextFrame.TextRange.Text = ActivePresentation.Slides(11 + PuzzleNumberIndex).Shapes("PuzzleCategory" & i).TextFrame.TextRange.Text
    For k = 1 To 26
        ActivePresentation.Slides(2).Shapes("Letter" & k).Visible = True
    Next k
    ActivePresentation.Slides(2).Shapes("LeftTab").Fill.ForeColor.RGB = RGB(198, 159, 48)
    ActivePresentation.Slides(2).Shapes("LeftTab").TextFrame.TextRange.Text = "Solve"
    ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(10).Shapes("LoadPuzzleChime").ActionSettings(ppMouseClick).SoundEffect.Play
    Exit Sub
errHandler:
    Exit Sub
End Sub

Private Sub solvePuzzle()
    On Error GoTo errHandler
    Dim i As Integer, j As Integer
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
    ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text = "WHEEL OF FORTUNE"
    ActivePresentation.Slides(10).Shapes("SolvePuzzleChime").ActionSettings(ppMouseClick).SoundEffect.Play
    Exit Sub
errHandler:
    Exit Sub
End Sub

Private Function isLetter(strValue As String) As Boolean
    Dim i As Integer
    For i = 1 To Len(strValue)
        If Asc(Mid(strValue, 1, 1)) < 65 Or Asc(Mid(strValue, 1, 1)) > 90 Then
            isLetter = False
        Else:
            isLetter = True
        End If
    Next i
End Function

Private Function isVowel(strValue As String) As Boolean
    Dim i As Integer
    For i = 1 To Len(strValue)
        If Asc(Mid(strValue, 1, 1)) = 65 Or Asc(Mid(strValue, 1, 1)) = 69 _
        Or Asc(Mid(strValue, 1, 1)) = 73 Or Asc(Mid(strValue, 1, 1)) = 79 _
        Or Asc(Mid(strValue, 1, 1)) = 85 Then
            isVowel = True
        Else:
            isVowel = False
        End If
    Next i
End Function

Sub guessLetter(oSh As Shape)
    On Error GoTo errHandler
    Dim i As Integer, j As Boolean, k As Integer
    For i = 1 To 26
        If ActivePresentation.Slides(2).Shapes("Letter" & i).Name = oSh.Name Then
            j = True
            Exit For
        End If
    Next i
    If j = True Then
        If ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text <> "" Then
            Dim theLetter As String, letterCount As Integer
            letterCount = 0
            theLetter = ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text
            If ActivePresentation.Slides(2).Shapes("BonusOutline").Line.Transparency = 0 Then
                If Len(ActivePresentation.Slides(2).Shapes("BonusLetters").TextFrame.TextRange.Text) >= 5 Then
                    MsgBox "The contestant can only choose four letters (or five if he or she has a wild card). Use the spiral arrow button to remove letters if necessary.", _
                    0, "Add Bonus Round Letter Error"
                    Exit Sub
                Else:
                    ActivePresentation.Slides(2).Shapes("BonusLetters").TextFrame.TextRange.Text = ActivePresentation.Slides(2).Shapes("BonusLetters").TextFrame.TextRange.Text + theLetter
                    ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = ""
                    ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = ""
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
                If ActivePresentation.Slides(2).Shapes("FinalSpinButton").Line.Transparency = 0 Then
                Else:
                    ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = ""
                End If
                ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = True
                ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = ""
                ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text = ""
                setValuePanelDisplay
                ActivePresentation.Slides(10).Shapes("GuessLetterWrong").ActionSettings(ppMouseClick).SoundEffect.Play
                Exit Sub
            End If
            If isVowel(theLetter) Then
                ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = ""
            Else:
                ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = letterCount
            End If
            ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text = ""
            ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = False
            setValuePanelDisplay
            ActivePresentation.Slides(10).Shapes("GuessLetterCorrect").ActionSettings(ppMouseClick).SoundEffect.Play
            Exit Sub
        End If
    End If
errHandler:
    Exit Sub
End Sub

Sub revealLetter(oSh As Shape)
    If ActivePresentation.Slides(2).Shapes("LeftTab").TextFrame.TextRange.Text = "Load Puzzle" Then
        LoadPuzzleOrSolve
    Else:
        Dim i As Integer, j As Boolean
        For i = 1 To 52
            If ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).Name = oSh.Name Then
                j = True
                Exit For
            End If
        Next i
        If j = True Then
            If ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).Fill.ForeColor.RGB <> RGB(24, 154, 80) Then
                ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).TextFrame.TextRange.Text = ActivePresentation.Slides(2).Shapes("PuzzleCache" & i).TextFrame.TextRange.Text
                ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).Fill.ForeColor.RGB = RGB(255, 255, 255)
            End If
        End If
    End If
End Sub

Private Sub bringLetterBack(i As Integer)
    ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text = Chr(i + 64)
End Sub

Sub EditSetUpPuzzle(oClickedShape As Shape)
    Dim oSh As Shape
    Dim sText As String
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    sText = InputBox("Type the letter for this puzzle board tile:", "Set Up Puzzle", oSh.TextFrame.TextRange.Text)
    While Len(sText) > 1
    sText = InputBox("Only one letter per tile. Try again:", "Set Up Puzzle", sText)
    Wend
    If Len(sText) = 1 And Not sText = " " Then
    oSh.TextFrame.TextRange.Text = UCase(sText)
    oSh.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Else:
    oSh.TextFrame.TextRange.Text = ""
    oSh.Fill.ForeColor.RGB = RGB(24, 154, 80)
    End If
End Sub

Sub EditSetUpCategory(oClickedShape As Shape)
    Dim oSh As Shape
    Dim sText As String
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    sText = InputBox("Type the category:", "Set Up Category", oSh.TextFrame.TextRange.Text)
    oSh.TextFrame.TextRange.Text = UCase(sText)
End Sub

Private Function wait(PauseTime As Double)
    Dim start
    start = Timer
    Do While Timer < start + PauseTime
        DoEvents
    Loop
End Function

Sub wipeOnClose()
    deleteAllPuzzles
    ClearBoardButton
    ClearMysteryIndicator
    wipeRoundScores
    Dim i As Integer, j As Integer
    For i = 1 To 4
        ActivePresentation.Slides(2).Shapes("Player" & i & "Name").TextFrame.TextRange.Text = "Player " & i
        ActivePresentation.Slides(2).Shapes("Player" & i & "TotalsScore").TextFrame.TextRange.Text = "0"
        If ActivePresentation.Slides(2).Shapes("Player" & i & "WildCard").Fill.Transparency = 0 Then
            ActivePresentation.Slides(2).Shapes("Player" & i & "WildCard").Fill.Transparency = 1
        End If
        If ActivePresentation.Slides(2).Shapes("Player" & i & "GiftTag").Fill.Transparency = 0 Then
            ActivePresentation.Slides(2).Shapes("Player" & i & "GiftTag").Fill.Transparency = 1
        End If
    Next i
    If ActivePresentation.Slides(3).Shapes("TheWheel").GroupItems("WildCard").Fill.Transparency = 1 Then
        toggleWildCard
    End If
    If ActivePresentation.Slides(3).Shapes("TheWheel").GroupItems("GiftTag").Fill.Transparency = 0 Then
        toggleGiftTag
    End If
    If ActivePresentation.Slides(2).Shapes("NumPlayers").TextFrame.TextRange.Text = "4 Players" Then
        TogglePlayers
    End If
    For j = 3 To 6
        ActivePresentation.Slides(j).Shapes("WheelValue").TextFrame.TextRange.Text = ""
    Next j
    ActivePresentation.SlideShowWindow.View.Exit
End Sub

Private Sub wipeRoundScores()
    Dim i As Integer
    For i = 1 To 4
        ActivePresentation.Slides(2).Shapes("Player" & i & "RoundScore").TextFrame.TextRange.Text = "0"
    Next i
End Sub

Sub toggleWildCard()
    Dim transparentLevel As Integer, wildString As String, i As Integer
    If ActivePresentation.Slides(3).Shapes("TheWheel").GroupItems("WildCard").Fill.Transparency = 0 Then
        transparentLevel = 1
        wildString = "Add"
    Else:
        transparentLevel = 0
        wildString = "Remove"
    End If
    For i = 3 To 5
        ActivePresentation.Slides(i).Shapes("TheWheel").GroupItems("WildCard").Fill.Transparency = transparentLevel
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

Sub toggle10000Wedge()
    Dim i As Integer, j As Integer
    If ActivePresentation.Slides(9).Shapes("10000Wedge").TextFrame.TextRange.Text = "on" Then
        ActivePresentation.Slides(9).Shapes("10000Wedge").TextFrame.TextRange.Text = "off"
        ActivePresentation.Slides(9).Shapes("10000Wedge").Fill.ForeColor.RGB = RGB(217, 150, 148)
        For i = 3 To 5
            ActivePresentation.Slides(i).Shapes("TheWheel").GroupItems("10000Wedge").Fill.Transparency = 1
        Next i
    Else:
        ActivePresentation.Slides(9).Shapes("10000Wedge").TextFrame.TextRange.Text = "on"
        ActivePresentation.Slides(9).Shapes("10000Wedge").Fill.ForeColor.RGB = RGB(195, 214, 155)
        For j = 3 To 5
            ActivePresentation.Slides(j).Shapes("TheWheel").GroupItems("10000Wedge").Fill.Transparency = 0
        Next j
    End If
End Sub

Sub toggleGiftTag()
    Dim transparentLevel As Integer, giftString As String, i As Integer
    If ActivePresentation.Slides(3).Shapes("TheWheel").GroupItems("GiftTag").Fill.Transparency = 0 Then
        transparentLevel = 1
        giftString = "Add"
    Else:
        transparentLevel = 0
        giftString = "Remove"
    End If
    For i = 3 To 5
        ActivePresentation.Slides(i).Shapes("TheWheel").GroupItems("GiftTag").Fill.Transparency = transparentLevel
        ActivePresentation.Slides(i).Shapes("ToggleGiftTag").TextFrame.TextRange.Text = giftString + " Gift Tag"
    Next i
End Sub

Sub toggleWheelValues()
    Dim i As Integer, j As Integer, k As Integer, m As Integer
    If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
        ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "classic ($300 min)"
        For i = 3 To 6
            For j = 1 To 7
                ActivePresentation.Slides(i).Shapes("TheWheel").GroupItems("ClassicWedge" & CStr(j)).Fill.Transparency = 0
            Next j
        Next i
    Else:
        ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)"
        For k = 3 To 6
            For m = 1 To 7
                ActivePresentation.Slides(k).Shapes("TheWheel").GroupItems("ClassicWedge" & CStr(m)).Fill.Transparency = 1
            Next m
        Next k
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
    Dim minim As Integer, maxim As Integer, i As Integer
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
    If ActivePresentation.Slides(8).Shapes("SetUpPuzzle" + CStr(maxim)).TextFrame.TextRange.Text = "" Then
        For i = maxim - 1 To minim Step -1
            ActivePresentation.Slides(8).Shapes("SetUpPuzzle" + CStr(i + 1)).TextFrame.TextRange.Text = ActivePresentation.Slides(8).Shapes("SetUpPuzzle" + CStr(i)).TextFrame.TextRange.Text
            ActivePresentation.Slides(8).Shapes("SetUpPuzzle" + CStr(i + 1)).Fill.ForeColor.RGB = ActivePresentation.Slides(8).Shapes("SetUpPuzzle" + CStr(i)).Fill.ForeColor.RGB
        Next i
        ActivePresentation.Slides(8).Shapes("SetUpPuzzle" + CStr(i + 1)).TextFrame.TextRange.Text = ""
        ActivePresentation.Slides(8).Shapes("SetUpPuzzle" + CStr(i + 1)).Fill.ForeColor.RGB = RGB(24, 154, 80)
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
    Dim minim As Integer, maxim As Integer, i As Integer
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
    If ActivePresentation.Slides(8).Shapes("SetUpPuzzle" + CStr(minim)).TextFrame.TextRange.Text = "" Then
        For i = minim + 1 To maxim
            ActivePresentation.Slides(8).Shapes("SetUpPuzzle" + CStr(i - 1)).TextFrame.TextRange.Text = ActivePresentation.Slides(8).Shapes("SetUpPuzzle" + CStr(i)).TextFrame.TextRange.Text
            ActivePresentation.Slides(8).Shapes("SetUpPuzzle" + CStr(i - 1)).Fill.ForeColor.RGB = ActivePresentation.Slides(8).Shapes("SetUpPuzzle" + CStr(i)).Fill.ForeColor.RGB
        Next i
        ActivePresentation.Slides(8).Shapes("SetUpPuzzle" + CStr(i - 1)).TextFrame.TextRange.Text = ""
        ActivePresentation.Slides(8).Shapes("SetUpPuzzle" + CStr(i - 1)).Fill.ForeColor.RGB = RGB(24, 154, 80)
    End If
End Sub

Sub RSTLNE()
    If ActivePresentation.Slides(2).Shapes("Letter1").Visible = False Then
        MsgBox "Please load a new puzzle before starting the bonus round.", 0, "Start Bonus Round Error"
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

Private Sub guessLetterViaFunction(i As Integer)
    Dim theLetter As String, k As Integer
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
    Dim letterToReturn As String, i As Integer, letterNumber As Integer
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
    On Error GoTo errHandler
    Dim letterExist As Boolean, letterToGuess As String, i As Integer, k As Integer
    If ActivePresentation.Slides(2).Shapes("BonusOutline").Line.Transparency = 0 And ActivePresentation.Slides(2).Shapes("Letter1").Visible = True Then
        If ActivePresentation.Slides(2).Shapes("BonusLetters").TextFrame.TextRange.Text = "" Then
            MsgBox "No letters found. Use the letter selector to input the letters the contestant chooses for the bonus round.", 0, "Guess Bonus Round Letters Error"
            Exit Sub
        End If
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
        ActivePresentation.Slides(2).Shapes("BonusBox").Fill.ForeColor.RGB = RGB(166, 166, 166)
        ActivePresentation.Slides(2).Shapes("BonusOutline").Line.Transparency = 1
        If letterExist = False Then
            ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = True
            Exit Sub
        Else:
            ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = False
            ActivePresentation.Slides(10).Shapes("GuessLetterCorrect").ActionSettings(ppMouseClick).SoundEffect.Play
            Exit Sub
        End If
    End If
errHandler:
    Exit Sub
End Sub

Sub toggleRound(oClickedShape As Shape)
  Dim oSh As Shape
    Dim sText As String
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    If oSh.Name = "Express Round" Then
        toggleFourthRound (True)
    ElseIf oSh.Name = "Fourth Round" Then
        toggleFourthRound (False)
        toggleBonusRound (True)
    Else:
        toggleBonusRound (False)
    End If
End Sub

Private Sub toggleFourthRound(i As Boolean)
    ActivePresentation.Slides(2).Shapes("FinalSpinButton").Visible = i
    ActivePresentation.Slides(2).Shapes("FinalSpinDescription").Visible = i
    ActivePresentation.Slides(2).Shapes("FinalSpinOverlay").Visible = i
    If i = False Then
        disableFinalSpin
    End If
End Sub

Sub toggleFinalSpin()
    If ActivePresentation.Slides(2).Shapes("FinalSpinButton").Visible = True Then
        If ActivePresentation.Slides(2).Shapes("FinalSpinButton").Line.Transparency = 1 Then
            ActivePresentation.Slides(2).Shapes("FinalSpinButton").Fill.ForeColor.RGB = RGB(225, 129, 75)
            ActivePresentation.Slides(2).Shapes("FinalSpinButton").Fill.Transparency = 0
            ActivePresentation.Slides(2).Shapes("FinalSpinButton").Line.Transparency = 0
            If ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text <> "WHEEL OF FORTUNE" Then
                setValuePanelDisplay
            End If
        Else:
            disableFinalSpin
        End If
    End If
End Sub

Private Sub disableFinalSpin()
    ActivePresentation.Slides(2).Shapes("FinalSpinButton").Fill.ForeColor.RGB = RGB(166, 166, 166)
    ActivePresentation.Slides(2).Shapes("FinalSpinButton").Fill.Transparency = 0.5
    ActivePresentation.Slides(2).Shapes("FinalSpinButton").Line.Transparency = 1
    If ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text <> "WHEEL OF FORTUNE" Then
        setValuePanelDisplay
    End If
End Sub

Private Sub toggleBonusRound(i As Boolean)
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
    On Error GoTo errHandler
    Dim x As Integer, rand As Integer, realRand As Double, effNew As Effect, effNew2 As Effect, effNew3 As Effect, effNew4 As Effect
    ' Remove wheel animations
    For x = ActivePresentation.SlideShowWindow.View.Slide.TimeLine.MainSequence.Count To 1 Step -1
        ActivePresentation.SlideShowWindow.View.Slide.TimeLine.MainSequence.Item(x).Delete
    Next x
    ' Clear letter counter
    ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = False
    Randomize
    rand = Int((3599 + 1) * Rnd)
    realRand = 1800 + rand / 10
    Set effNew = ActivePresentation.SlideShowWindow.View.Slide.TimeLine.MainSequence _
        .AddEffect(Shape:=ActivePresentation.SlideShowWindow.View.Slide.Shapes("TheWheel"), effectId:=msoAnimEffectSpin, trigger:=msoAnimTriggerWithPrevious)
    With effNew
        .Timing.Duration = 3.6
        .Timing.SmoothEnd = msoTrue
        .Timing.Decelerate = 1
        .EffectParameters.Amount = realRand
    End With
    Set effNew2 = ActivePresentation.SlideShowWindow.View.Slide.TimeLine.MainSequence _
        .AddEffect(Shape:=ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue"), effectId:=msoAnimEffectAppear, trigger:=msoAnimTriggerWithPrevious)
    With effNew2
        .Timing.TriggerDelayTime = 3.6
    End With
    Set effNew3 = ActivePresentation.SlideShowWindow.View.Slide.TimeLine.MainSequence _
        .AddEffect(Shape:=ActivePresentation.SlideShowWindow.View.Slide.Shapes("BackText"), effectId:=msoAnimEffectAppear, trigger:=msoAnimTriggerWithPrevious)
    With effNew3
        .Timing.TriggerDelayTime = 3.6
    End With
    Set effNew4 = ActivePresentation.SlideShowWindow.View.Slide.TimeLine.MainSequence _
        .AddEffect(Shape:=ActivePresentation.SlideShowWindow.View.Slide.Shapes("BackOval"), effectId:=msoAnimEffectAppear, trigger:=msoAnimTriggerWithPrevious)
    With effNew4
        .Timing.TriggerDelayTime = 3.6
    End With
    Call Module2.youLandedOn(rand, ActivePresentation.SlideShowWindow.View.Slide.SlideNumber)
    If IsNumeric(ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text) Then
        ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = CLng(ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text)
    ElseIf ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Mystery 1" _
    Or ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Mystery 2" _
    Or ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Express" Then
        ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = 1000
    ElseIf ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Free Play" Then
        ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = 500
    ElseIf ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Bankrupt" _
    Or ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Lose a Turn" Then
        ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = ""
    End If
    setValuePanelDisplay
    ActivePresentation.Slides(10).Shapes("SpinWheel").ActionSettings(ppMouseClick).SoundEffect.Play
    Exit Sub
errHandler:
    Exit Sub
End Sub

Sub manuallySetValuePanel()
    Dim sText As String
    If ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text = "WHEEL OF FORTUNE" Then
        MsgBox "Load a puzzle first before manually setting the spun wheel value.", 0, "Manually Set Value Panel"
        Exit Sub
    End If
    sText = InputBox("Manually set the spun wheel value:", "Manually Set Value Panel", ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text)
    While IsNumeric(sText) = False And sText <> ""
        sText = InputBox("You can only enter numbers here. Try again:", "Manually Set Value Panel", sText)
    Wend
    If sText = "" Then
    Else:
        If CLng(sText) > 10000 Or CLng(sText) < 1 Then
            MsgBox "Wheel values must range from 1 to 10000.", 0, "Manually Set Value Panel"
            Exit Sub
        Else:
            ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = CLng(sText)
            ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = False
            setValuePanelDisplay
        End If
    End If
End Sub

Private Sub setValuePanelDisplay()
    Dim spunWheelValue, letterCounter, effectiveWheelValue As Long
    Set spunWheelValue = ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange
    Set letterCounter = ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange
    If spunWheelValue.Text <> "" Then
        If ActivePresentation.Slides(2).Shapes("FinalSpinButton").Line.Transparency = 1 Then
            effectiveWheelValue = CLng(spunWheelValue.Text)
        Else:
            effectiveWheelValue = CLng(spunWheelValue.Text) + 1000
        End If
        ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text = "$" & CStr(effectiveWheelValue)
    Else:
        ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text = ""
        Exit Sub
    End If
    If letterCounter.Text <> "" And IsNumeric(letterCounter.Text) Then
        If spunWheelValue.Text = "10000" Then
            ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text = ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text _
        & " (no multiplier)"
        Else:
            ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text = ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text _
        & " * " & letterCounter.Text & " = $" & CLng(effectiveWheelValue) * CLng(letterCounter.Text)
        End If
    End If
End Sub

Sub puzzleSwapChoose(oClickedShape As Shape)
    Dim oSh As Shape
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    Dim numberToSwap
    numberToSwap = InputBox("Swap this puzzle with puzzle number:", "Puzzle Swap")
    While IsNumeric(numberToSwap) = False:
        If numberToSwap = "" Then
            Exit Sub
        Else:
            numberToSwap = InputBox("Please enter a number:", "Puzzle Swap", numberToSwap)
        End If
    Wend
    puzzleSwap CInt(oSh.TextFrame.TextRange.Text), CInt(numberToSwap)
End Sub

Private Sub puzzleSwap(i As Integer, j As Integer)
    Dim iPuzzleNumberIndex As Integer, jPuzzleNumberIndex As Integer, l As Integer, m As Integer, n As Integer, o As Integer
    iPuzzleNumberIndex = Int((i - 1) / 12)
    jPuzzleNumberIndex = Int((j - 1) / 12)
    If jPuzzleNumberIndex + 1 <= ActivePresentation.SectionProperties.SlidesCount(4) Then
        ' Move initial puzzle to swap to cache
        For l = 1 To 52
            ActivePresentation.Slides(9).Shapes("PuzzleSolutionSwap-" & CStr(l)).TextFrame.TextRange.Text = ActivePresentation.Slides(11 + iPuzzleNumberIndex).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(l)).TextFrame.TextRange.Text
            ActivePresentation.Slides(9).Shapes("PuzzleSolutionSwap-" & CStr(l)).Fill.ForeColor.RGB = ActivePresentation.Slides(11 + iPuzzleNumberIndex).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(l)).Fill.ForeColor.RGB
        Next l
        ActivePresentation.Slides(9).Shapes("PuzzleCategorySwap").TextFrame.TextRange.Text = ActivePresentation.Slides(11 + iPuzzleNumberIndex).Shapes("PuzzleCategory" & CStr(i)).TextFrame.TextRange.Text
        ' Overwrite contents of initial puzzle number with the puzzle number you want to swap with
        For m = 1 To 52
            ActivePresentation.Slides(11 + iPuzzleNumberIndex).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(m)).TextFrame.TextRange.Text = ActivePresentation.Slides(11 + jPuzzleNumberIndex).Shapes("PuzzleSolution" & CStr(j) & "-" & CStr(m)).TextFrame.TextRange.Text
            ActivePresentation.Slides(11 + iPuzzleNumberIndex).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(m)).Fill.ForeColor.RGB = ActivePresentation.Slides(11 + jPuzzleNumberIndex).Shapes("PuzzleSolution" & CStr(j) & "-" & CStr(m)).Fill.ForeColor.RGB
        Next m
        ActivePresentation.Slides(11 + iPuzzleNumberIndex).Shapes("PuzzleCategory" & CStr(i)).TextFrame.TextRange.Text = ActivePresentation.Slides(11 + jPuzzleNumberIndex).Shapes("PuzzleCategory" & CStr(j)).TextFrame.TextRange.Text
        ' Overwrite contents of the puzzle number you want to swap with with what's in the cache
        For n = 1 To 52
            ActivePresentation.Slides(11 + jPuzzleNumberIndex).Shapes("PuzzleSolution" & CStr(j) & "-" & CStr(n)).TextFrame.TextRange.Text = ActivePresentation.Slides(9).Shapes("PuzzleSolutionSwap-" & CStr(n)).TextFrame.TextRange.Text
            ActivePresentation.Slides(11 + jPuzzleNumberIndex).Shapes("PuzzleSolution" & CStr(j) & "-" & CStr(n)).Fill.ForeColor.RGB = ActivePresentation.Slides(9).Shapes("PuzzleSolutionSwap-" & CStr(n)).Fill.ForeColor.RGB
        Next n
        ActivePresentation.Slides(11 + jPuzzleNumberIndex).Shapes("PuzzleCategory" & CStr(j)).TextFrame.TextRange.Text = ActivePresentation.Slides(9).Shapes("PuzzleCategorySwap").TextFrame.TextRange.Text
        ' Clear the cache
        For o = 1 To 52
            ActivePresentation.Slides(9).Shapes("PuzzleSolutionSwap-" & CStr(o)).TextFrame.TextRange.Text = ""
            ActivePresentation.Slides(9).Shapes("PuzzleSolutionSwap-" & CStr(o)).Fill.ForeColor.RGB = RGB(24, 154, 80)
        Next o
        ActivePresentation.Slides(9).Shapes("PuzzleCategory" & CStr(j)).TextFrame.TextRange.Text = ""
    Else:
        MsgBox "The puzzle number you want to swap with does not exist. Generate more puzzle numbers with the right arrow next to the puzzle numbers in Set Up Puzzles.", 0, "Puzzle Swap"
    End If
End Sub

Private Sub addPuzzleRow(num As Integer)
    Dim k As Integer, l As Integer
    ActivePresentation.Slides(11).Duplicate
    For k = 1 To 12
        ActivePresentation.Slides(12).Shapes("LinkTo" & k).TextFrame.TextRange.Text = (k + (12 * num))
        ActivePresentation.Slides(12).Shapes("LinkTo" & k).Name = "LinkTo" & (k + (12 * num))
        ActivePresentation.Slides(12).Shapes("Swap" & k).TextFrame.TextRange.Text = (k + (12 * num))
        ActivePresentation.Slides(12).Shapes("Swap" & k).Name = "Swap" & (k + (12 * num))
        ActivePresentation.Slides(12).Shapes("PuzzleCategory" & k).TextFrame.TextRange.Text = ""
        ActivePresentation.Slides(12).Shapes("PuzzleCategory" & k).Name = "PuzzleCategory" & (k + (12 * num))
        ActivePresentation.Slides(12).Shapes("PrevAllPuzzles").Visible = msoTrue
        ActivePresentation.Slides(12).Shapes("NextAllPuzzles").Visible = msoFalse
        For l = 1 To 52
            ActivePresentation.Slides(12).Shapes("PuzzleSolution" & CStr(k) & "-" & CStr(l)).TextFrame.TextRange.Text = ""
            ActivePresentation.Slides(12).Shapes("PuzzleSolution" & CStr(k) & "-" & CStr(l)).Fill.ForeColor.RGB = RGB(24, 154, 80)
            ActivePresentation.Slides(12).Shapes("PuzzleSolution" & CStr(k) & "-" & CStr(l)).Name = "PuzzleSolution" & CStr(k + (12 * num)) & "-" & CStr(l)
        Next l
    Next k
    ActivePresentation.Slides(12).MoveTo toPos:=11 + ActivePresentation.SectionProperties.SlidesCount(4) - 1
    ActivePresentation.Slides(11 + ActivePresentation.SectionProperties.SlidesCount(4) - 2).Shapes("NextAllPuzzles").Visible = msoTrue
End Sub

Sub nextPuzzleRow()
    Dim r As Integer, p As Integer, RowIndex As Integer
    savePuzzle
    RowIndex = CInt(ActivePresentation.Slides(7).Shapes("CurrentPuzzleRowIndex").TextFrame.TextRange.Text)
    For r = 7 To 9
        For p = 1 To 12
            With ActivePresentation.Slides(r).Shapes("LinkTo" & CStr(p + (12 * RowIndex)))
                .Name = "LinkTo" & CStr(p + (12 * (RowIndex + 1)))
                .TextFrame.TextRange.Text = CStr(p + (12 * (RowIndex + 1)))
            End With
        Next p
        ActivePresentation.Slides(r).Shapes("PrevPuzzleRow").Visible = msoTrue
    Next r
    ActivePresentation.Slides(7).Shapes("CurrentPuzzleRowIndex").TextFrame.TextRange.Text = RowIndex + 1
    If ActivePresentation.SectionProperties.SlidesCount(4) <= RowIndex + 1 Then
        addPuzzleRow (RowIndex + 1)
    End If
    puzzleSetupJump (1 + (12 * (RowIndex + 1)))
End Sub

Sub prevPuzzleRow()
    Dim r As Integer, p As Integer, RowIndex As Integer
    savePuzzle
    RowIndex = CInt(ActivePresentation.Slides(7).Shapes("CurrentPuzzleRowIndex").TextFrame.TextRange.Text)
    For r = 7 To 9
        For p = 1 To 12
            With ActivePresentation.Slides(r).Shapes("LinkTo" & CStr(p + (12 * RowIndex)))
                .Name = "LinkTo" & CStr(p + (12 * (RowIndex - 1)))
                .TextFrame.TextRange.Text = CStr(p + (12 * (RowIndex - 1)))
            End With
        Next p
        If RowIndex - 1 = 0 Then
            ActivePresentation.Slides(r).Shapes("PrevPuzzleRow").Visible = msoFalse
        End If
    Next r
    ActivePresentation.Slides(7).Shapes("CurrentPuzzleRowIndex").TextFrame.TextRange.Text = RowIndex - 1
    puzzleSetupJump (1 + (12 * (RowIndex - 1)))
    
End Sub

Private Sub exactPuzzleRow(num As Integer)
    Dim RowIndex As Integer, r As Integer, p As Integer
    RowIndex = CInt(ActivePresentation.Slides(7).Shapes("CurrentPuzzleRowIndex").TextFrame.TextRange.Text)
    For r = 7 To 9
        For p = 1 To 12
            With ActivePresentation.Slides(r).Shapes("LinkTo" & CStr(p + (12 * RowIndex)))
                .Name = "LinkTo" & CStr(p + (12 * num))
                .TextFrame.TextRange.Text = CStr(p + (12 * num))
            End With
        Next p
        If num = 0 Then
            ActivePresentation.Slides(r).Shapes("PrevPuzzleRow").Visible = msoFalse
        Else:
            ActivePresentation.Slides(r).Shapes("PrevPuzzleRow").Visible = msoTrue
        End If
    Next r
    ActivePresentation.Slides(7).Shapes("CurrentPuzzleRowIndex").TextFrame.TextRange.Text = num
End Sub

Sub TogglePlayers()
    Dim i As Integer, j As Integer
    If ActivePresentation.Slides(2).Shapes("NumPlayers").TextFrame.TextRange.Text = "3 Players" Then
        For i = 1 To 4:
            With ActivePresentation.Slides(2)
                .Shapes("Player" & i & "BuyVowelButton").Left = 85.61977 + 157.4233 * (i - 1)
                .Shapes("Player" & i & "TransferTotalsButton").Left = 85.65504 + 157.4233 * (i - 1)
                .Shapes("Player" & i & "RoundDollarSign").Left = 103.5772 + 157.4233 * (i - 1)
                .Shapes("Player" & i & "TotalsDollarSign").Left = 103.5772 + 157.4233 * (i - 1)
                .Shapes("Player" & i & "RoundScore").Left = 121.5772 + 157.4233 * (i - 1)
                .Shapes("Player" & i & "RoundScoreCompatibility").Left = 121.5772 + 157.4233 * (i - 1)
                .Shapes("Player" & i & "TotalsScore").Left = 117.9772 + 157.4233 * (i - 1)
                .Shapes("Player" & i & "Name").Left = 92.63976 + 157.4233 * (i - 1)
                .Shapes("Player" & i & "XButton").Left = 197.8109 + 157.4233 * (i - 1)
                .Shapes("Player" & i & "XButtonCompatibility").Left = 197.8109 + 157.4233 * (i - 1)
                .Shapes("Player" & i & "WildCard").Left = 220.1326 + 157.4233 * (i - 1)
                .Shapes("Player" & i & "GiftTag").Left = 217.3 + 157.4233 * (i - 1)
            End With
        Next i
        With ActivePresentation.Slides(2)
            .Shapes("RoundTotals").Left = 15.44449
            .Shapes("NumPlayers").Left = 19.75614
            .Shapes("AddRemovePlayers").Left = 66.01803
            .Shapes("NumPlayers").TextFrame.TextRange.Text = "4 Players"
            .Shapes("AddRemovePlayers").TextFrame.TextRange.Text = "-"
        End With
    Else:
        For j = 1 To 3:
        With ActivePresentation.Slides(2)
            .Shapes("Player" & j & "BuyVowelButton").Left = 118.98 + 185.5516 * (j - 1)
            .Shapes("Player" & j & "TransferTotalsButton").Left = 119.0153 + 185.5516 * (j - 1)
            .Shapes("Player" & j & "RoundDollarSign").Left = 136.9374 + 185.5516 * (j - 1)
            .Shapes("Player" & j & "TotalsDollarSign").Left = 136.9374 + 185.5516 * (j - 1)
            .Shapes("Player" & j & "RoundScore").Left = 154.9374 + 185.5516 * (j - 1)
            .Shapes("Player" & j & "RoundScoreCompatibility").Left = 154.9374 + 185.5516 * (j - 1)
            .Shapes("Player" & j & "TotalsScore").Left = 151.3376 + 185.5516 * (j - 1)
            .Shapes("Player" & j & "Name").Left = 126 + 185.5516 * (j - 1)
            .Shapes("Player" & j & "XButton").Left = 231.1711 + 185.5516 * (j - 1)
            .Shapes("Player" & j & "XButtonCompatibility").Left = 231.1711 + 185.5516 * (j - 1)
            .Shapes("Player" & j & "WildCard").Left = 253.4928 + 185.5516 * (j - 1)
            .Shapes("Player" & j & "GiftTag").Left = 250.6602 + 185.5516 * (j - 1)
        End With
        Next j
        With ActivePresentation.Slides(2)
            .Shapes("Player" & 4 & "BuyVowelButton").Left = 118.98 + 185.5516 * (4.4 - 1)
            .Shapes("Player" & 4 & "TransferTotalsButton").Left = 119.0153 + 185.5516 * (4.4 - 1)
            .Shapes("Player" & 4 & "RoundDollarSign").Left = 136.9374 + 185.5516 * (4.4 - 1)
            .Shapes("Player" & 4 & "TotalsDollarSign").Left = 136.9374 + 185.5516 * (4.4 - 1)
            .Shapes("Player" & 4 & "RoundScore").Left = 154.9374 + 185.5516 * (4.4 - 1)
            .Shapes("Player" & 4 & "RoundScoreCompatibility").Left = 154.9374 + 185.5516 * (4.4 - 1)
            .Shapes("Player" & 4 & "TotalsScore").Left = 151.3376 + 185.5516 * (4.4 - 1)
            .Shapes("Player" & 4 & "Name").Left = 126 + 185.5516 * (4.4 - 1)
            .Shapes("Player" & 4 & "XButton").Left = 231.1711 + 185.5516 * (4.4 - 1)
            .Shapes("Player" & 4 & "XButtonCompatibility").Left = 231.1711 + 185.5516 * (4.4 - 1)
            .Shapes("Player" & 4 & "WildCard").Left = 253.4928 + 185.5516 * (4.4 - 1)
            .Shapes("Player" & 4 & "GiftTag").Left = 250.6602 + 185.5516 * (4.4 - 1)
            .Shapes("RoundTotals").Left = 30.63039
            .Shapes("NumPlayers").Left = 34.94205
            .Shapes("AddRemovePlayers").Left = 81.20393
            .Shapes("NumPlayers").TextFrame.TextRange.Text = "3 Players"
            .Shapes("AddRemovePlayers").TextFrame.TextRange.Text = "+"
        End With
    End If
End Sub
