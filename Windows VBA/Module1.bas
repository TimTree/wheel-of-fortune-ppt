Attribute VB_Name = "Module1"
Option Explicit

Sub goToHowToUse()
    ' Allows slide to advance if macros are enabled
    ActivePresentation.Slides(1).Shapes("MacroDisabledText").Visible = False
    ' Allows slide to advance if not PowerPoint 2007
    If Val(Application.Version) <= 12 Then
        MsgBox "PowerPoint 2007 is no longer supported. Upgrade to 2010 or newer, or download an earlier version of Wheel of Fortune for PowerPoint.", 0, "PowerPoint 2007 Not Supported"
        Exit Sub
    End If
    savePuzzleAndShadeOccupiedPuzzles
    ActivePresentation.SlideShowSettings.Run.View.PointerType = ppSlideShowPointerArrow
    SlideShowWindows(1).View.GotoSlide 18 + ActivePresentation.SectionProperties.SlidesCount(4)
End Sub

Sub goToSetUp()
    ' Checks if macros are enabled
    ActivePresentation.Slides(1).Shapes("MacroDisabledText").Visible = False
    ' Allows slide to advance if not PowerPoint 2007
    If Val(Application.Version) <= 12 Then
        MsgBox "PowerPoint 2007 is no longer supported. Upgrade to 2010 or newer, or download an earlier version of Wheel of Fortune for PowerPoint.", 0, "PowerPoint 2007 Not Supported"
        Exit Sub
    End If
    savePuzzleAndShadeOccupiedPuzzles
    ActivePresentation.SlideShowSettings.Run.View.PointerType = ppSlideShowPointerArrow
    SlideShowWindows(1).View.GotoSlide 7
End Sub

Sub goToPuzzleBoard()
    ' Checks if macros are enabled
    ActivePresentation.Slides(1).Shapes("MacroDisabledText").Visible = False
    ' Allows slide to advance if not PowerPoint 2007
    If Val(Application.Version) <= 12 Then
        MsgBox "PowerPoint 2007 is no longer supported. Upgrade to 2010 or newer, or download an earlier version of Wheel of Fortune for PowerPoint.", 0, "PowerPoint 2007 Not Supported"
        Exit Sub
    End If
    savePuzzleAndShadeOccupiedPuzzles
    ActivePresentation.SlideShowSettings.Run.View.PointerType = ppSlideShowPointerArrow
    SlideShowWindows(1).View.GotoSlide 2
End Sub

Sub BGChange()
    Dim i As Integer
    Dim themeNumber
    Set themeNumber = ActivePresentation.Slides(9).Shapes("Backdrop").TextFrame.TextRange
    Dim puzzleBoardGradientBottom As Long, puzzleBoardGradientMiddle As Long, puzzleBoardGradientTop As Long
    Dim setFloorEdges As Long, setFloorMiddle As Long, setFloorLine As Long, wheelGradientMiddle As Long
    Dim wheelGradientTop As Long, wheelGradientBottom As Long, helpColor As Long, categoryColor As Long, letterSelectorColor As Long
    If themeNumber.Text = "studio" Then
        puzzleBoardGradientBottom = RGB(2, 127, 190)
        puzzleBoardGradientMiddle = RGB(94, 189, 208)
        puzzleBoardGradientTop = RGB(2, 127, 190)
        setFloorEdges = RGB(0, 51, 0)
        setFloorMiddle = RGB(0, 153, 0)
        setFloorLine = RGB(38, 100, 38)
        wheelGradientBottom = RGB(44, 87, 17)
        wheelGradientMiddle = RGB(34, 138, 46)
        wheelGradientTop = RGB(44, 87, 17)
        categoryColor = RGB(27, 91, 33)
        letterSelectorColor = RGB(139, 193, 229)
        helpColor = RGB(44, 64, 58)
        themeNumber.Text = "stadium"
    ElseIf themeNumber.Text = "stadium" Then
        puzzleBoardGradientBottom = RGB(233, 91, 81)
        puzzleBoardGradientMiddle = RGB(246, 161, 192)
        puzzleBoardGradientTop = RGB(233, 91, 91)
        setFloorEdges = RGB(197, 75, 87)
        setFloorMiddle = RGB(212, 130, 153)
        setFloorLine = RGB(197, 103, 139)
        wheelGradientBottom = RGB(210, 116, 125)
        wheelGradientMiddle = RGB(212, 98, 98)
        wheelGradientTop = RGB(210, 116, 125)
        categoryColor = RGB(181, 69, 69)
        letterSelectorColor = RGB(248, 196, 223)
        helpColor = RGB(119, 69, 69)
        themeNumber.Text = "valentine's"
    ElseIf themeNumber.Text = "valentine's" Then
        puzzleBoardGradientBottom = RGB(120, 114, 200)
        puzzleBoardGradientMiddle = RGB(60, 67, 212)
        puzzleBoardGradientTop = RGB(41, 14, 158)
        setFloorEdges = RGB(163, 163, 163)
        setFloorMiddle = RGB(207, 207, 207)
        setFloorLine = RGB(29, 36, 141)
        wheelGradientBottom = RGB(146, 148, 180)
        wheelGradientMiddle = RGB(169, 169, 191)
        wheelGradientTop = RGB(146, 148, 180)
        categoryColor = RGB(118, 127, 146)
        letterSelectorColor = RGB(107, 107, 169)
        helpColor = RGB(81, 77, 133)
        themeNumber.Text = "winter"
    ElseIf themeNumber.Text = "winter" Then
        puzzleBoardGradientBottom = RGB(0, 0, 0)
        puzzleBoardGradientMiddle = RGB(0, 0, 0)
        puzzleBoardGradientTop = RGB(0, 0, 0)
        setFloorEdges = RGB(38, 38, 38)
        setFloorMiddle = RGB(87, 68, 35)
        setFloorLine = RGB(38, 38, 38)
        wheelGradientBottom = RGB(0, 0, 0)
        wheelGradientMiddle = RGB(0, 0, 0)
        wheelGradientTop = RGB(0, 0, 0)
        categoryColor = RGB(38, 38, 38)
        letterSelectorColor = RGB(127, 127, 127)
        helpColor = RGB(179, 162, 199)
        themeNumber.Text = "blackout"
    Else:
        setFloorEdges = RGB(41, 38, 35)
        setFloorMiddle = RGB(125, 73, 126)
        setFloorLine = RGB(16, 37, 63)
        categoryColor = RGB(23, 55, 94)
        letterSelectorColor = RGB(79, 129, 189)
        helpColor = RGB(179, 162, 199)
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
            .Shapes("Help").Fill.ForeColor.RGB = helpColor
            With .Shapes("BackDrop")
                If themeNumber.Text = "studio" Then
                    .Fill.Transparency = 1
                Else:
                    .Fill.Transparency = 0
                    .Fill.GradientStops.Insert wheelGradientTop, 0
                    .Fill.GradientStops.Insert wheelGradientMiddle, 0.5
                    .Fill.GradientStops.Insert wheelGradientBottom, 1
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
    For j = 1 To 27
        ActivePresentation.Slides(2).Shapes("Letter" & j).Visible = False
        bringLetterBack (j)
    Next j
    ActivePresentation.Slides(2).Shapes("TossUpBanner").Visible = False
    ActivePresentation.Slides(2).Shapes("CategoryBox").TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(2).Shapes("LeftTab").Fill.ForeColor.RGB = RGB(41, 183, 233)
    ActivePresentation.Slides(2).Shapes("LeftTab").TextFrame.TextRange.Text = "Load Puzzle"
    ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = False
    ActivePresentation.Slides(2).Shapes("NoMoreVowels").Visible = False
    ActivePresentation.Slides(2).Shapes("NoMoreConsonants").Visible = False
    ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text = ActivePresentation.Slides(9).Shapes("GameName").TextFrame.TextRange.Text
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
        Do While IsNumeric(sText) = False And sText <> ""
            sText = InputBox("You can only enter numbers here. Try again:", "Manually Edit Round Score", sText)
        Loop
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
        If ActivePresentation.Slides(2).Shapes("Player" & i & "TotalsDollarSign").Name = oSh.Name Or _
        ActivePresentation.Slides(2).Shapes("Player" & i & "TotalsScore").Name = oSh.Name Then
            j = True
            Exit For
        End If
    Next i
    If j = True Then
        sText = InputBox("Manually edit " & ActivePresentation.Slides(2).Shapes("Player" & i & "Name").TextFrame.TextRange.Text & "'s totals score:", "Manually Edit Totals Score", ActivePresentation.Slides(2).Shapes("Player" & i & "TotalsScore").TextFrame.TextRange.Text)
        Do While IsNumeric(sText) = False And sText <> ""
            sText = InputBox("You can only enter numbers here. Try again:", "Manually Edit Totals Score", sText)
        Loop
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
        If ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text = ActivePresentation.Slides(9).Shapes("GameName").TextFrame.TextRange.Text Then
            MsgBox "During a game, click here to add the amount shown on the value panel" & vbNewLine & _
            "(currently reads " & ActivePresentation.Slides(9).Shapes("GameName").TextFrame.TextRange.Text & ") to " & ActivePresentation.Slides(2).Shapes("Player" & i & "Name").TextFrame.TextRange.Text & "'s round score." _
            , 0, "Add to " & ActivePresentation.Slides(2).Shapes("Player" & i & "Name").TextFrame.TextRange.Text & "'s Round Score"
            Exit Sub
        Else:
            If ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = "" Then
                MsgBox "There's nothing to add because the value panel is empty." & vbNewLine & _
                "Spin the wheel or manually set the spun wheel value on the Value Panel first." _
                , 0, "Add to " & ActivePresentation.Slides(2).Shapes("Player" & i & "Name").TextFrame.TextRange.Text & "'s Round Score"
                Exit Sub
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
                        If ActivePresentation.Slides(10).Shapes("WheelItems").TextFrame.TextRange.Text = "once" And _
                        ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text = "** $10000 **" Then
                            If ActivePresentation.Slides(3).Shapes("TheWheel").GroupItems("10000Wedge").Fill.Transparency = 0 Then
                                toggle10000Wedge
                            End If
                        End If
                    Else:
                        ActivePresentation.Slides(2).Shapes("Player" & i & "RoundScore").TextFrame.TextRange.Text = _
                        CLng(ActivePresentation.Slides(2).Shapes("Player" & i & "RoundScore").TextFrame.TextRange.Text) + _
                        CLng(ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text) * _
                        effectiveWheelValue
                    End If
                End If
            If ActivePresentation.Slides(2).Shapes("TossUpBanner").Visible = True Then
                solvePuzzle
                Exit Sub
            End If
            ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = ""
            setValuePanelDisplay
            End If
        End If
    End If
End Sub

Sub PlayerBuyaVowel(oSh As Shape)
    Dim i As Integer, j As Boolean, RoundDollarAmount, playerName, VOWELCOST As Long
    VOWELCOST = CLng(Replace(ActivePresentation.Slides(9).Shapes("VowelPrice").TextFrame.TextRange.Text, "$", ""))
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
    HOUSEMINIMUM = CLng(Replace(ActivePresentation.Slides(9).Shapes("HouseMinimum").TextFrame.TextRange.Text, "$", ""))
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
        If ActivePresentation.Slides(10).Shapes("WheelItems").TextFrame.TextRange.Text = "once" Then
            If oSh.Name Like "*WildCard" Then
                If ActivePresentation.Slides(3).Shapes("TheWheel").GroupItems("WildCard").Fill.Transparency = 0 Then
                    toggleWildCard
                End If
            ElseIf oSh.Name Like "*GiftTag" Then
                If ActivePresentation.Slides(3).Shapes("TheWheel").GroupItems("GiftTag").Fill.Transparency = 0 Then
                    toggleGiftTag
                End If
            End If
        End If
    Else:
        oSh.Fill.Transparency = 1
    End If
End Sub

Sub DetermineMystery()
    Dim randomNumber As Integer, noMysteryWedgeWarning
    If ActivePresentation.Slides(4).Shapes("WheelValue").TextFrame.TextRange.Text = "Mystery 1" Then
        ActivePresentation.Slides(4).Shapes("TheWheel").GroupItems("MysteryWedge1").Fill.Transparency = 1
    ElseIf ActivePresentation.Slides(4).Shapes("WheelValue").TextFrame.TextRange.Text = "Mystery 2" Then
        ActivePresentation.Slides(4).Shapes("TheWheel").GroupItems("MysteryWedge2").Fill.Transparency = 1
    Else:
        noMysteryWedgeWarning = MsgBox("The wheel is not on a Mystery wedge. This feature is intended to 'flip' the Mystery wedge landed on." & vbNewLine & vbNewLine & _
        "Do you still want to use this feature for another purpose?", vbYesNo + vbDefaultButton2, "Flip Mystery Wedge Warning")
        If noMysteryWedgeWarning = vbNo Then
            Exit Sub
        End If
    End If
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
End Sub

Sub ClearMysteryIndicator()
    ActivePresentation.Slides(4).Shapes("MysteryIndicator").TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(4).Shapes("TheWheel").GroupItems("MysteryWedge1").Fill.Transparency = 0
    ActivePresentation.Slides(4).Shapes("TheWheel").GroupItems("MysteryWedge2").Fill.Transparency = 0
End Sub

Sub TileChanger(i As Integer)
    Dim oSh
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
        SlideShowWindows(1).View.Slide.Shapes("BoardTile" & CStr(i)).OLEFormat.Object.Value = ""
    Next i
End Sub

Sub EraseEntirePuzzle()
    Dim i As Integer, shouldIEraseAll
    shouldIEraseAll = MsgBox("Are you sure you want to delete the entire puzzle?", vbYesNo + vbDefaultButton2, "Confirm Puzzle Delete")
    If shouldIEraseAll = vbYes Then
        For i = 1 To 52
            SlideShowWindows(1).View.Slide.Shapes("BoardTile" & CStr(i)).OLEFormat.Object.Value = ""
        Next i
        SlideShowWindows(1).View.Slide.Shapes("CategoryBox").OLEFormat.Object.Value = ""
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
    placePuzzleToSetUp (CInt(oSh.TextFrame.TextRange.Text))
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
    savePuzzleAndShadeOccupiedPuzzles
    placePuzzleToSetUp (CInt(oSh.TextFrame.TextRange.Text))
    highlightCurrentPuzzle (CInt(oSh.TextFrame.TextRange.Text))
End Sub

Sub puzzleSetupJump(num As Integer)
    placePuzzleToSetUp (num)
    shadeOccupiedPuzzlesFull
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
    sText = InputBox("Edit the vowel price. The default price is $250.", "Edit Vowel Price", CLng(Replace(oSh.TextFrame.TextRange.Text, "$", "")))
    Do While IsNumeric(sText) = False And sText <> ""
        sText = InputBox("You can only enter numbers here. Try again:", "Edit Vowel Price", sText)
    Loop
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
    sText = InputBox("Edit the house minimum. The default minimum is $1000.", "Edit House Minimum", CLng(Replace(oSh.TextFrame.TextRange.Text, "$", "")))
    Do While IsNumeric(sText) = False And sText <> ""
        sText = InputBox("You can only enter numbers here. Try again:", "Edit House Minimum", sText)
    Loop
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
            ActivePresentation.Slides(12 + PuzzleIndex).Shapes("PuzzleSolution" & CStr(i + (12 * PuzzleIndex)) & "-" & CStr(j)).TextFrame.TextRange.Text = ActivePresentation.Slides(8).Shapes("BoardTile" & CStr(j)).OLEFormat.Object.Value
            ActivePresentation.Slides(12 + PuzzleIndex).Shapes("PuzzleSolution" & CStr(i + (12 * PuzzleIndex)) & "-" & CStr(j)).Fill.ForeColor.RGB = ActivePresentation.Slides(8).Shapes("BoardTile" & CStr(j)).OLEFormat.Object.BackColor
        Next j
        ActivePresentation.Slides(12 + PuzzleIndex).Shapes("PuzzleCategory" & CStr(i + (12 * PuzzleIndex))).TextFrame.TextRange.Text = ActivePresentation.Slides(8).Shapes("CategoryBox").OLEFormat.Object.Value
    End If
End Sub

Private Sub placePuzzleToSetUp(i As Integer)
    Dim PuzzleIndex As Integer, n As Integer
    PuzzleIndex = CInt(ActivePresentation.Slides(7).Shapes("CurrentPuzzleRowIndex").TextFrame.TextRange.Text)
    For n = 1 To 52
        ActivePresentation.Slides(8).Shapes("BoardTile" & CStr(n)).OLEFormat.Object.Value = ActivePresentation.Slides(12 + PuzzleIndex).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(n)).TextFrame.TextRange.Text
    Next n
    ActivePresentation.Slides(8).Shapes("CategoryBox").OLEFormat.Object.Value = ActivePresentation.Slides(12 + PuzzleIndex).Shapes("PuzzleCategory" & CStr(i)).TextFrame.TextRange.Text
End Sub

Private Sub deleteAllPuzzles()
    Dim s As Integer, i As Integer, j As Integer, k As Integer
    s = 12 + ActivePresentation.SectionProperties.SlidesCount(4) - 1
    Do Until s = 12
        ActivePresentation.Slides(s).Delete
        s = 12 + ActivePresentation.SectionProperties.SlidesCount(4) - 1
    Loop
    ActivePresentation.Slides(12).Shapes("NextAllPuzzles").Visible = msoFalse
    For i = 1 To 12
        For j = 1 To 52
            ActivePresentation.Slides(12).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(j)).TextFrame.TextRange.Text = ""
            ActivePresentation.Slides(12).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(j)).Fill.ForeColor.RGB = RGB(24, 154, 80)
        Next j
        ActivePresentation.Slides(12).Shapes("PuzzleCategory" & CStr(i)).TextFrame.TextRange.Text = ""
    Next i
    For k = 1 To 52
        ActivePresentation.Slides(8).Shapes("BoardTile" & CStr(k)).OLEFormat.Object.Value = ""
    Next k
    ActivePresentation.Slides(8).Shapes("CategoryBox").OLEFormat.Object.Value = ""
    exactPuzzleRow (0)
    shadeOccupiedPuzzlesFull
End Sub

Private Sub savePuzzleAndShadeOccupiedPuzzles()
    Dim PuzzleIndex As Integer, r As Integer, blankPuzzle As Boolean, i As Integer, j As Integer, thereWasAPuzzle As Boolean
    PuzzleIndex = CInt(ActivePresentation.Slides(7).Shapes("CurrentPuzzleRowIndex").TextFrame.TextRange.Text)
    thereWasAPuzzle = False
    blankPuzzle = True
    For i = 1 To 12
        If ActivePresentation.Slides(8).Shapes("LinkTo" & CStr(i + (12 * PuzzleIndex))).Fill.ForeColor.RGB = RGB(250, 192, 144) Then
            thereWasAPuzzle = True
            Exit For
        End If
    Next i
    If thereWasAPuzzle = True Then
        For j = 1 To 52
            ActivePresentation.Slides(12 + PuzzleIndex).Shapes("PuzzleSolution" & CStr(i + (12 * PuzzleIndex)) & "-" & CStr(j)).TextFrame.TextRange.Text = ActivePresentation.Slides(8).Shapes("BoardTile" & CStr(j)).OLEFormat.Object.Value
            ActivePresentation.Slides(12 + PuzzleIndex).Shapes("PuzzleSolution" & CStr(i + (12 * PuzzleIndex)) & "-" & CStr(j)).Fill.ForeColor.RGB = ActivePresentation.Slides(8).Shapes("BoardTile" & CStr(j)).OLEFormat.Object.BackColor
            If ActivePresentation.Slides(12 + PuzzleIndex).Shapes("PuzzleSolution" & CStr(i + (12 * PuzzleIndex)) & "-" & CStr(j)).TextFrame.TextRange.Text <> "" Then
                blankPuzzle = False
            End If
        Next j
        ActivePresentation.Slides(12 + PuzzleIndex).Shapes("PuzzleCategory" & CStr(i + (12 * PuzzleIndex))).TextFrame.TextRange.Text = ActivePresentation.Slides(8).Shapes("CategoryBox").OLEFormat.Object.Value
        If ActivePresentation.Slides(12 + PuzzleIndex).Shapes("PuzzleCategory" & CStr(i + (12 * PuzzleIndex))).TextFrame.TextRange.Text <> "" Then
            blankPuzzle = False
        End If
        For r = 7 To 9
           If blankPuzzle = False Then
               With ActivePresentation.Slides(r).Shapes("LinkTo" & CStr(i + (12 * PuzzleIndex)))
                   .Fill.ForeColor.RGB = RGB(146, 224, 204)
                   .Line.ForeColor.RGB = RGB(55, 96, 146)
               End With
           Else:
               With ActivePresentation.Slides(r).Shapes("LinkTo" & CStr(i + (12 * PuzzleIndex)))
                   .Fill.ForeColor.RGB = RGB(149, 179, 215)
                   .Line.ForeColor.RGB = RGB(55, 96, 146)
               End With
           End If
        Next r
    End If
End Sub

Private Sub shadeOccupiedPuzzlesFull()
    Dim PuzzleIndex As Integer, p As Integer, q As Integer, r As Integer, blankPuzzle As Boolean
    PuzzleIndex = CInt(ActivePresentation.Slides(7).Shapes("CurrentPuzzleRowIndex").TextFrame.TextRange.Text)
    For p = 1 To 12
        blankPuzzle = True
        For q = 1 To 52
            If ActivePresentation.Slides(12 + PuzzleIndex).Shapes("PuzzleSolution" & CStr(p + (12 * PuzzleIndex)) & "-" & CStr(q)).TextFrame.TextRange.Text <> "" Then
                blankPuzzle = False
                Exit For
            End If
        Next q
        If ActivePresentation.Slides(12 + PuzzleIndex).Shapes("PuzzleCategory" & CStr(p + (12 * PuzzleIndex))).TextFrame.TextRange.Text <> "" Then
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
    savePuzzleAndShadeOccupiedPuzzles
    resetBonusRound
    toggleBonusRound (False)
    toggleFourthRound (False)
    ActivePresentation.Slides(1).Shapes("MacroDisabledText").Visible = True
End Sub

Sub goToHowToUseFromSetUpPuzzles()
    savePuzzleAndShadeOccupiedPuzzles
    SlideShowWindows(1).View.GotoSlide 18 + ActivePresentation.SectionProperties.SlidesCount(4)
End Sub

Sub goToPuzzleBoardFromSetUpPuzzles()
    savePuzzleAndShadeOccupiedPuzzles
    SlideShowWindows(1).View.GotoSlide 2
End Sub

Sub goToSettingsFromSetUpPuzzles()
    savePuzzleAndShadeOccupiedPuzzles
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
        If ActivePresentation.Slides(12 + PuzzleNumberIndex).Shapes("PuzzleSolution" & i & "-" & m).Fill.ForeColor.RGB = RGB(255, 255, 255) Then
            puzzleBoolean = True
            Exit For
        End If
    Next m
    If ActivePresentation.Slides(12 + PuzzleNumberIndex).Shapes("PuzzleCategory" & i).TextFrame.TextRange.Text <> "" Then
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
        numberToLoad = InputBox("Enter the puzzle number to load (1, 2 etc)." & vbNewLine & vbNewLine & _
        "Append T to the number to load the puzzle as a Toss-Up (1T, 2T, etc).", "Load Puzzle", ActivePresentation.Slides(2).Shapes("NextPuzzleToLoad").TextFrame.TextRange.Text)
        Do While IsNumeric(Replace(UCase(numberToLoad), "T", "")) = False:
            If numberToLoad = "" Then
                Exit Sub
            Else:
                numberToLoad = InputBox("Please enter a number, or a number with a T:", "Load Puzzle", numberToLoad)
            End If
        Loop
            If InStr(UCase(numberToLoad), "T") > 0 Then
                Call loadPuzzle(CInt(Replace(UCase(numberToLoad), "T", "")), True)
            Else:
                Call loadPuzzle(CInt(numberToLoad), False)
            End If
    Else:
        If ActivePresentation.Slides(2).Shapes("TossUpBanner").Visible = True Then
            MsgBox "To reveal the Toss-Up, add the Toss-Up value to a player's round score." & vbNewLine & vbNewLine & _
            "To end the Toss-Up without scoring, use Clear Board on the top right of this slide.", 0, "Solve Toss-Up Puzzle Note"
        Else:
            solvePuzzle
            Exit Sub
        End If
    End If
End Sub

Private Sub loadPuzzle(i As Integer, isTossUp As Boolean)
    On Error GoTo errHandler
    Dim SpanishNError As Boolean, sText As String
    SpanishNError = False
    If puzzleExists(i) = False Then
        MsgBox "No puzzle found for number " & i & ".", 0, "Load Puzzle Error"
        Exit Sub
    End If
    If isTossUp = True Then
        sText = InputBox("How much is the Toss-Up worth in dollars?", "Set Toss-Up Value", "1000")
        Do Until sText = ""
            If IsNumeric(sText) = False Then
                GoTo notValidTossUpValue
            ElseIf CLng(sText) > 10000 Or CLng(sText) < 1 Then
                GoTo notValidTossUpValue
            Else:
                Exit Do
            End If
notValidTossUpValue:
            sText = InputBox("The Toss-Up value must be a number between 1 and 10000.", "Set Toss-Up Value", sText)
        Loop
        If sText = "" Then
            Exit Sub
        End If
    End If
    ClearBoardButton
    Dim PuzzleNumberIndex As Integer, j As Integer, k As Integer
    PuzzleNumberIndex = Int((i - 1) / 12)
    For j = 1 To 52
        ActivePresentation.Slides(2).Shapes("PuzzleBoard" & j).Fill.ForeColor.RGB = ActivePresentation.Slides(12 + PuzzleNumberIndex).Shapes("PuzzleSolution" & i & "-" & j).Fill.ForeColor.RGB
        If ActivePresentation.Slides(12 + PuzzleNumberIndex).Shapes("PuzzleSolution" & i & "-" & j).TextFrame.TextRange.Text <> "" Then
            ActivePresentation.Slides(2).Shapes("PuzzleCache" & j).TextFrame.TextRange.Text = ActivePresentation.Slides(12 + PuzzleNumberIndex).Shapes("PuzzleSolution" & i & "-" & j).TextFrame.TextRange.Text
            If isLetter(ActivePresentation.Slides(12 + PuzzleNumberIndex).Shapes("PuzzleSolution" & i & "-" & j).TextFrame.TextRange.Text) = False Then
                ActivePresentation.Slides(2).Shapes("PuzzleBoard" & j).TextFrame.TextRange.Text = ActivePresentation.Slides(12 + PuzzleNumberIndex).Shapes("PuzzleSolution" & i & "-" & j).TextFrame.TextRange.Text
            End If
            If ActivePresentation.Slides(12 + PuzzleNumberIndex).Shapes("PuzzleSolution" & i & "-" & j).TextFrame.TextRange.Text = ActivePresentation.Slides(9).Shapes("SpanN").TextFrame.TextRange.Text And _
            ActivePresentation.Slides(9).Shapes("SpanishN").TextFrame.TextRange.Text = "off" Then
               SpanishNError = True
               toggleSpanishN
            End If
        End If
    Next j
    ActivePresentation.Slides(2).Shapes("CategoryBox").TextFrame.TextRange.Text = ActivePresentation.Slides(12 + PuzzleNumberIndex).Shapes("PuzzleCategory" & i).TextFrame.TextRange.Text
    If isTossUp = False Then
        For k = 1 To 27
            If k < 27 Or (k = 27 And ActivePresentation.Slides(9).Shapes("SpanishN").TextFrame.TextRange.Text = "on") Then
                ActivePresentation.Slides(2).Shapes("Letter" & k).Visible = True
            End If
        Next k
        ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text = ""
        ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = ""
    Else:
        ActivePresentation.Slides(2).Shapes("TossUpBanner").Visible = True
        ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = CLng(sText)
        setValuePanelDisplay
    End If
    ActivePresentation.Slides(2).Shapes("LeftTab").Fill.ForeColor.RGB = RGB(198, 159, 48)
    ActivePresentation.Slides(2).Shapes("LeftTab").TextFrame.TextRange.Text = "Solve"
    ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(2).Shapes("NextPuzzleToLoad").TextFrame.TextRange.Text = i + 1
    If SpanishNError Then
        MsgBox "This puzzle uses the letter " & ActivePresentation.Slides(9).Shapes("SpanN").TextFrame.TextRange.Text & ", but the Spanish " _
        & ActivePresentation.Slides(9).Shapes("SpanN").TextFrame.TextRange.Text & " setting was disabled. This setting has automatically been enabled." & vbNewLine & vbNewLine & _
        "You can re-disable the Spanish " & ActivePresentation.Slides(9).Shapes("SpanN").TextFrame.TextRange.Text & " in the Settings slide.", 0, "Spanish " & ActivePresentation.Slides(9).Shapes("SpanN").TextFrame.TextRange.Text & " Note"
    End If
    ActivePresentation.Slides(11).Shapes("LoadPuzzleChime").ActionSettings(ppMouseClick).SoundEffect.Play
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
    For j = 1 To 27
        ActivePresentation.Slides(2).Shapes("Letter" & j).Visible = False
    Next j
        ActivePresentation.Slides(2).Shapes("TossUpBanner").Visible = False
    ActivePresentation.Slides(2).Shapes("LeftTab").Fill.ForeColor.RGB = RGB(41, 183, 233)
    ActivePresentation.Slides(2).Shapes("LeftTab").TextFrame.TextRange.Text = "Load Puzzle"
    ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = False
    ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text = ActivePresentation.Slides(9).Shapes("GameName").TextFrame.TextRange.Text
    ActivePresentation.Slides(2).Shapes("NoMoreVowels").Visible = False
    ActivePresentation.Slides(2).Shapes("NoMoreConsonants").Visible = False
    ActivePresentation.Slides(11).Shapes("SolvePuzzleChime").ActionSettings(ppMouseClick).SoundEffect.Play
    Exit Sub
errHandler:
    Exit Sub
End Sub

Private Function isLetter(strValue As String) As Boolean
    Dim i As Integer
    Dim extendedChars As String
    extendedChars = ActivePresentation.Slides(9).Shapes("ExtendedChars").TextFrame.TextRange.Text
    For i = 1 To Len(strValue)
        If (Asc(Mid(strValue, 1, 1)) < 65 Or Asc(Mid(strValue, 1, 1)) > 90) And InStr(extendedChars, Mid(strValue, 1, 1)) = 0 Then
            isLetter = False
        Else:
            isLetter = True
        End If
    Next i
End Function

Private Function isVowel(strValue As String) As Boolean
    Dim vowelChars As String
    vowelChars = ActivePresentation.Slides(9).Shapes("VowelChars").TextFrame.TextRange.Text
    Dim i As Integer
    For i = 1 To Len(strValue)
        If InStr(vowelChars, Mid(strValue, 1, 1)) > 0 Then
            isVowel = True
        Else:
            isVowel = False
        End If
    Next i
End Function

Private Function lettersMatch(letter1 As String, letterSelectorLetter As String) As Boolean
    Dim extendedChars As String
    Select Case letterSelectorLetter
        Case "A"
            extendedChars = ActivePresentation.Slides(9).Shapes("AChars").TextFrame.TextRange.Text
        Case "C"
            extendedChars = ActivePresentation.Slides(9).Shapes("CChars").TextFrame.TextRange.Text
        Case "E"
            extendedChars = ActivePresentation.Slides(9).Shapes("EChars").TextFrame.TextRange.Text
        Case "I"
            extendedChars = ActivePresentation.Slides(9).Shapes("IChars").TextFrame.TextRange.Text
        Case "O"
            extendedChars = ActivePresentation.Slides(9).Shapes("OChars").TextFrame.TextRange.Text
        Case "U"
            extendedChars = ActivePresentation.Slides(9).Shapes("UChars").TextFrame.TextRange.Text
        Case Else
            extendedChars = letterSelectorLetter
    End Select
    If InStr(extendedChars, letter1) > 0 Then
        lettersMatch = True
    Else:
        lettersMatch = False
    End If
End Function

Private Function isInArray(theString As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = theString Then
            isInArray = True
            Exit Function
        End If
    Next i
    isInArray = False
End Function

Sub guessLetter(oSh As Shape)
    On Error GoTo errHandler
    Dim i As Integer, j As Boolean, k As Integer
    For i = 1 To 27
        If ActivePresentation.Slides(2).Shapes("Letter" & i).Name = oSh.Name Then
            j = True
            Exit For
        End If
    Next i
    If j = True Then
        If ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text <> "" Then
            Dim theLetter As String, letterCount As Integer, vowelsRemaining As Boolean, consonantsRemaining As Boolean
            letterCount = 0
            vowelsRemaining = False
            consonantsRemaining = False
            theLetter = ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text
            If ActivePresentation.Slides(2).Shapes("BonusOutline").Line.Transparency = 0 Then
                If Len(ActivePresentation.Slides(2).Shapes("BonusLetters").TextFrame.TextRange.Text) >= 5 Then
                    MsgBox "The contestant can only choose four letters (or five if he or she has a Wild Card). Use the spiral arrow button to remove letters if necessary.", _
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
                If ActivePresentation.Slides(2).Shapes("PuzzleBoard" & k).Fill.ForeColor.RGB = RGB(255, 255, 255) And ActivePresentation.Slides(2).Shapes("PuzzleBoard" & k).TextFrame.TextRange.Text = "" Then
                    If lettersMatch(ActivePresentation.Slides(2).Shapes("PuzzleCache" & k).TextFrame.TextRange.Text, theLetter) Then
                        If ActivePresentation.Slides(9).Shapes("BlueTiles").TextFrame.TextRange.Text = "off" Then
                            ActivePresentation.Slides(2).Shapes("PuzzleBoard" & k).TextFrame.TextRange.Text = ActivePresentation.Slides(2).Shapes("PuzzleCache" & k).TextFrame.TextRange.Text
                        Else:
                            ActivePresentation.Slides(2).Shapes("PuzzleBoard" & k).Fill.ForeColor.RGB = RGB(0, 0, 255)
                        End If
                        letterCount = letterCount + 1
                    ElseIf isVowel(ActivePresentation.Slides(2).Shapes("PuzzleCache" & k).TextFrame.TextRange.Text) Then
                        vowelsRemaining = True
                    Else:
                        consonantsRemaining = True
                    End If
                End If
            Next k
            If letterCount = 0 Then
                If ActivePresentation.Slides(2).Shapes("FinalSpinButton").Line.Transparency <> 0 Then
                    ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = ""
                End If
                ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = True
                ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = ""
                ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text = ""
                setValuePanelDisplay
                If ActivePresentation.Slides(2).Shapes("FinalSpinButton").Line.Transparency <> 0 Then
                    ActivePresentation.Slides(11).Shapes("GuessLetterWrong").ActionSettings(ppMouseClick).SoundEffect.Play
                End If
                Exit Sub
            End If
            If isVowel(theLetter) Then
                ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = ""
                If vowelsRemaining = False And consonantsRemaining = True And _
                ActivePresentation.Slides(9).Shapes("NoMoreVowels").TextFrame.TextRange.Text = "on" Then
                    ActivePresentation.Slides(2).Shapes("NoMoreVowels").Visible = True
                End If
            Else:
                ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = letterCount
                If vowelsRemaining = True And consonantsRemaining = False And _
                ActivePresentation.Slides(9).Shapes("NoMoreVowels").TextFrame.TextRange.Text = "on" Then
                    ActivePresentation.Slides(2).Shapes("NoMoreConsonants").Visible = True
                End If
            End If
            ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text = ""
            ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = False
            setValuePanelDisplay
            If ActivePresentation.Slides(2).Shapes("FinalSpinButton").Line.Transparency <> 0 Then
                ActivePresentation.Slides(11).Shapes("GuessLetterCorrect").ActionSettings(ppMouseClick).SoundEffect.Play
            End If
            Exit Sub
        End If
    End If
errHandler:
    Exit Sub
End Sub

Sub revealLetter(oSh As Shape)
On Error GoTo errHandler
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
            If ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).Fill.ForeColor.RGB = RGB(0, 0, 255) Then
                ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).TextFrame.TextRange.Text = ActivePresentation.Slides(2).Shapes("PuzzleCache" & i).TextFrame.TextRange.Text
                ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).Fill.ForeColor.RGB = RGB(255, 255, 255)
                Exit Sub
            ElseIf ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).Fill.ForeColor.RGB <> RGB(24, 154, 80) Then
                ' If a letter was already selected in letter selector, do nothing
                Dim k As Integer
                For k = 1 To 27:
                    If ActivePresentation.Slides(2).Shapes("Letter" & k).TextFrame.TextRange.Text = "" Then
                        Exit Sub
                    End If
                Next k
                ' If no letters were selected, instantiate a toss-up.
                ' Prompt for toss-up value at stake if there's no value currently in Value Panel
                If ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = "" Then
                    Dim sText As String
                    sText = InputBox("How much is the Toss-Up worth in dollars?", "Set Toss-Up Value", "1000")
                    Do Until sText = ""
                        If IsNumeric(sText) = False Then
                            GoTo notValidTossUpValue
                        ElseIf CLng(sText) > 10000 Or CLng(sText) < 1 Then
                            GoTo notValidTossUpValue
                        Else:
                            Exit Do
                        End If
notValidTossUpValue:
                        sText = InputBox("The toss-up value must be a number between 1 and 10000.", "Set Toss-Up Value", sText)
                    Loop
                    If sText = "" Then
                        Exit Sub
                    Else:
                        ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = CLng(sText)
                        ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = False
                        setValuePanelDisplay
                    End If
                End If
                If ActivePresentation.Slides(2).Shapes("TossUpBanner").Visible = False Then
                    ' Hide the letter selector during a toss-up
                    Dim m As Integer
                    For m = 1 To 27:
                        ActivePresentation.Slides(2).Shapes("Letter" & m).Visible = False
                        ActivePresentation.Slides(2).Shapes("TossUpBanner").Visible = True
                    Next m
                End If
                Dim n As Integer, isFirstReveal As Boolean
                isFirstReveal = True
                For n = 1 To 52
                    If ActivePresentation.Slides(2).Shapes("PuzzleBoard" & n).Fill.ForeColor.RGB = RGB(255, 255, 255) Then
                        If isLetter(ActivePresentation.Slides(2).Shapes("PuzzleBoard" & n).TextFrame.TextRange.Text) Then
                            isFirstReveal = False
                        End If
                    End If
                Next n
                ' Reveal letter
                ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).TextFrame.TextRange.Text = ActivePresentation.Slides(2).Shapes("PuzzleCache" & i).TextFrame.TextRange.Text
                If isFirstReveal Then
                    ActivePresentation.Slides(11).Shapes("TossUpMusic").ActionSettings(ppMouseClick).SoundEffect.Play
                    Exit Sub
                End If
            End If
        End If
    End If
errHandler:
    Exit Sub
End Sub

Private Sub bringLetterBack(i As Integer)
    If i = 27 Then
        ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text = ActivePresentation.Slides(9).Shapes("SpanN").TextFrame.TextRange.Text
    Else:
        ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text = Chr(i + 64)
    End If
End Sub

Private Function wait(PauseTime As Double)
    Dim start
    start = Timer
    Do While Timer < start + PauseTime
        DoEvents
    Loop
End Function

Sub wipeScores()
    wipeOnClose (False)
End Sub

Sub wipeAllWarning()
    Dim wipeAllConfirm
    wipeAllConfirm = MsgBox("This will reset the template to a clean slate, preserving only your settings. Are you ABSOLUTELY sure you want this?", vbYesNo + vbDefaultButton2, "Confirm Wipe All Puzzles/Scores")
    If wipeAllConfirm = vbYes Then
        wipeOnClose (True)
    Else:
        Exit Sub
    End If
End Sub

Private Sub wipeOnClose(wipeAll As Boolean)
    If wipeAll = True Then
        deleteAllPuzzles
    End If
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
    restoreWheelItems
    If ActivePresentation.Slides(2).Shapes("NumPlayers").TextFrame.TextRange.Text = "4 Players" Or _
    ActivePresentation.Slides(2).Shapes("NumPlayers").TextFrame.TextRange.Text = "2 Players" Then
        TogglePlayers (3)
    End If
    For j = 3 To 6
        ActivePresentation.Slides(j).Shapes("WheelValue").TextFrame.TextRange.Text = ""
    Next j
    ActivePresentation.Slides(2).Shapes("NextPuzzleToLoad").TextFrame.TextRange.Text = "1"
    ActivePresentation.SlideShowWindow.View.Exit
End Sub

Private Sub wipeRoundScores()
    Dim i As Integer
    For i = 1 To 4
        ActivePresentation.Slides(2).Shapes("Player" & i & "RoundScore").TextFrame.TextRange.Text = "0"
    Next i
End Sub

Sub toggleWildCard()
    Dim i As Integer
    For i = 3 To 5
        ActivePresentation.Slides(i).Shapes("TheWheel").GroupItems("WildCard").Fill.Transparency = 1
        ActivePresentation.Slides(i).Shapes("RestoreWheelItems").Visible = True
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

Sub toggleSpanishN()
    If ActivePresentation.Slides(9).Shapes("SpanishN").TextFrame.TextRange.Text = "on" Then
        ActivePresentation.Slides(9).Shapes("SpanishN").TextFrame.TextRange.Text = "off"
        ActivePresentation.Slides(9).Shapes("SpanishN").Fill.ForeColor.RGB = RGB(217, 150, 148)
        ActivePresentation.Slides(2).Shapes("Letter27").Visible = False
        ActivePresentation.Slides(2).Shapes("LetterSecondRowGroup").Left = 30.18748
    Else:
        ActivePresentation.Slides(9).Shapes("SpanishN").TextFrame.TextRange.Text = "on"
        ActivePresentation.Slides(9).Shapes("SpanishN").Fill.ForeColor.RGB = RGB(195, 214, 155)
        If ActivePresentation.Slides(2).Shapes("Letter1").Visible = True Then
            ActivePresentation.Slides(2).Shapes("Letter27").Visible = True
        End If
        ActivePresentation.Slides(2).Shapes("LetterSecondRowGroup").Left = 39.18748
    End If
End Sub

Sub toggle10000Wedge()
    Dim i As Integer
    For i = 3 To 5
        ActivePresentation.Slides(i).Shapes("TheWheel").GroupItems("10000Wedge").Fill.Transparency = 1
        ActivePresentation.Slides(i).Shapes("RestoreWheelItems").Visible = True
    Next i
End Sub

Sub toggleGiftTag()
    Dim i As Integer
    For i = 3 To 5
        ActivePresentation.Slides(i).Shapes("TheWheel").GroupItems("GiftTag").Fill.Transparency = 1
        ActivePresentation.Slides(i).Shapes("RestoreWheelItems").Visible = True
    Next i
End Sub

Sub restoreWheelItems()
    Dim i As Integer
    For i = 3 To 5
        If ActivePresentation.Slides(10).Shapes("WildCard").TextFrame.TextRange.Text = "on" Then
            ActivePresentation.Slides(i).Shapes("TheWheel").GroupItems("WildCard").Fill.Transparency = 0
        End If
        If ActivePresentation.Slides(10).Shapes("10000Wedge").TextFrame.TextRange.Text = "on" Then
            ActivePresentation.Slides(i).Shapes("TheWheel").GroupItems("10000Wedge").Fill.Transparency = 0
        End If
        If ActivePresentation.Slides(10).Shapes("GiftTag").TextFrame.TextRange.Text = "on" Then
            ActivePresentation.Slides(i).Shapes("TheWheel").GroupItems("GiftTag").Fill.Transparency = 0
        End If
        ActivePresentation.Slides(i).Shapes("RestoreWheelItems").Visible = False
    Next i
End Sub

Sub toggleWheelValues()
    Dim i As Integer, j As Integer, k As Integer, m As Integer
    If ActivePresentation.Slides(10).Shapes("WheelValues").TextFrame.TextRange.Text = "$500" Then
        ActivePresentation.Slides(10).Shapes("WheelValues").TextFrame.TextRange.Text = "$300"
        For i = 3 To 6
            For j = 1 To 7
                ActivePresentation.Slides(i).Shapes("TheWheel").GroupItems("ClassicWedge" & CStr(j)).Fill.Transparency = 0
            Next j
        Next i
    Else:
        ActivePresentation.Slides(10).Shapes("WheelValues").TextFrame.TextRange.Text = "$500"
        For k = 3 To 6
            For m = 1 To 7
                ActivePresentation.Slides(k).Shapes("TheWheel").GroupItems("ClassicWedge" & CStr(m)).Fill.Transparency = 1
            Next m
        Next k
    End If
End Sub

Sub toggleFreePlay(oClickedShape As Shape)
  Dim oSh As Shape, i As Integer, j As Integer
    Dim sText As String
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    Call toggleOnOff(oSh)
    If ActivePresentation.Slides(10).Shapes("FreePlayWedge").TextFrame.TextRange.Text = "off" Then
        For i = 3 To 6
            ActivePresentation.Slides(i).Shapes("TheWheel").GroupItems("Orange950").Fill.Transparency = 0
        Next i
    Else:
        For j = 3 To 6
            ActivePresentation.Slides(j).Shapes("TheWheel").GroupItems("Orange950").Fill.Transparency = 1
        Next j
    End If
End Sub

Sub toggleBankrupts(oClickedShape As Shape)
    Dim oSh As Shape, sText As String, i As Integer, j As Boolean
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    For i = 3 To 6
        If ActivePresentation.Slides(10).Shapes("Slide" & i & "Bankrupts").Name = oSh.Name Then
            j = True
            Exit For
        End If
    Next i
    If j = True Then
        If oSh.TextFrame.TextRange.Text = "2" Then
            oSh.TextFrame.TextRange.Text = "1"
            ActivePresentation.Slides(i).Shapes("TheWheel").GroupItems("Purple600").Fill.Transparency = 0
        Else:
            oSh.TextFrame.TextRange.Text = "2"
            ActivePresentation.Slides(i).Shapes("TheWheel").GroupItems("Purple600").Fill.Transparency = 1
        End If
    End If
End Sub

Sub toggle5Wedge(oClickedShape As Shape)
  Dim oSh As Shape, i As Integer, j As Integer
    Dim sText As String
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    Call toggleOnOff(oSh)
    If ActivePresentation.Slides(10).Shapes("5Wedge").TextFrame.TextRange.Text = "off" Then
        For i = 3 To 6
            ActivePresentation.Slides(i).Shapes("TheWheel").GroupItems("Purple5").Fill.Transparency = 1
        Next i
    Else:
        For j = 3 To 6
            ActivePresentation.Slides(j).Shapes("TheWheel").GroupItems("Purple5").Fill.Transparency = 0
        Next j
    End If
End Sub

Sub toggleWheelItems(oClickedShape As Shape)
  Dim oSh As Shape, i As Integer, j As Integer
    Dim sText As String
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    Call toggleOnOff(oSh)
    If oSh.TextFrame.TextRange.Text = "off" Then
        For i = 3 To 5
            ActivePresentation.Slides(i).Shapes("TheWheel").GroupItems(oSh.Name).Fill.Transparency = 1
        Next i
    Else:
        For j = 3 To 5
            ActivePresentation.Slides(j).Shapes("TheWheel").GroupItems(oSh.Name).Fill.Transparency = 0
        Next j
    End If
End Sub

Sub toggleClaimable()
    Dim i As Integer, j As Integer, k As Integer, m As Integer
    If ActivePresentation.Slides(10).Shapes("WheelItems").TextFrame.TextRange.Text = "once" Then
        ActivePresentation.Slides(10).Shapes("WheelItems").TextFrame.TextRange.Text = "multiple"
    Else:
        ActivePresentation.Slides(10).Shapes("WheelItems").TextFrame.TextRange.Text = "once"
    End If
End Sub

Sub toggleShotClock(oClickedShape As Shape)
    On Error GoTo errHandler
    Dim oSh As Shape, sText As String, currentShotClockTime As Integer, newShotClockTime As Integer
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    If oSh.TextFrame.TextRange.Text = "none" Then
        currentShotClockTime = 0
    Else:
        currentShotClockTime = CInt(Replace(oSh.TextFrame.TextRange.Text, " seconds", ""))
    End If
    sText = InputBox("The shot clock helps you enforce time limits for player decisions." & vbNewLine & vbNewLine & _
    "Enter a number from 1 to 30 to enable and set the shot clock's time limit in seconds, or 0 to disable the shot clock.", "Configure Shot Clock", CStr(currentShotClockTime))
    Do While IsNumeric(sText) = False And sText <> ""
        sText = InputBox("You can only enter numbers here. Try again:", "Configure Shot Clock", sText)
    Loop
    If sText = "" Then
        Exit Sub
    Else:
        newShotClockTime = CInt(sText)
        If newShotClockTime = 0 Then
            Dim i As Integer
            oSh.TextFrame.TextRange.Text = "none"
            ActivePresentation.Slides(2).Shapes("ShotClockBaseNumber").TextFrame.TextRange.Text = ""
            ActivePresentation.Slides(2).Shapes("ShotClockBaseNumber").Visible = False
            ActivePresentation.Slides(2).Shapes("ShotClockBase").Visible = False
            ActivePresentation.Slides(2).Shapes("ShotClockOverlay").Visible = False
            For i = 0 To 29
                ActivePresentation.Slides(2).Shapes("ShotClock" & i).Visible = False
            Next i
        ElseIf newShotClockTime > 30 Or newShotClockTime < 0 Then
            MsgBox "The shot clock time limit cannot exceed 30 seconds.", 0, "Configure Shot Clock Error"
            Exit Sub
        Else:
            Dim j As Integer, k As Integer
            oSh.TextFrame.TextRange.Text = CStr(newShotClockTime) & " seconds"
            ActivePresentation.Slides(2).Shapes("ShotClockBaseNumber").TextFrame.TextRange.Text = CStr(newShotClockTime)
            ActivePresentation.Slides(2).Shapes("ShotClockBaseNumber").Visible = True
            ActivePresentation.Slides(2).Shapes("ShotClockBase").Visible = True
            ActivePresentation.Slides(2).Shapes("ShotClockOverlay").Visible = True
            For k = 0 To 29
                ActivePresentation.Slides(2).Shapes("ShotClock" & k).Visible = False
            Next k
            For j = 0 To newShotClockTime - 1
                ActivePresentation.Slides(2).Shapes("ShotClock" & j).Visible = True
            Next j
        End If
        ' PowerPoint 2010 has a bug that breaks triggers when modifying its shapes' visibilities. The workaround for this version is to restart the slide show.
        If Val(Application.Version) <= 14 Then
            Dim m As Integer, n As Integer
            ActivePresentation.SlideShowWindow.View.Exit
            For m = 1 To 8
               ActivePresentation.Slides(m).SlideShowTransition.Hidden = msoTrue
            Next m
            ActivePresentation.SlideShowSettings.Run
            For n = 1 To 8
               ActivePresentation.Slides(n).SlideShowTransition.Hidden = msoFalse
            Next n
        End If
    End If
    Exit Sub
errHandler:
    MsgBox "The shot clock time limit cannot exceed 30 seconds.", 0, "Configure Shot Clock Error"
End Sub

Sub editGameName(oClickedShape As Shape)
    Dim oSh As Shape, sText As String
    For Each oSh In SlideShowWindows(1).View.Slide.Shapes
        If oSh.Name = oClickedShape.Name Then
            Exit For
        End If
    Next
    sText = InputBox("Edit the game name. The default name is WHEEL OF FORTUNE.", "Edit Game Name", oSh.TextFrame.TextRange.Text)
    Do While InStr(sText, "$") > 0 And sText <> ""
        sText = InputBox("The game name cannot contain the $ sign to prevent confusion with wheel values. Try again:", "Edit Game Name", sText)
    Loop
    If Trim(sText) = "" Then
        Exit Sub
    Else:
        If ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text = oSh.TextFrame.TextRange.Text Then
            ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text = Trim(UCase(sText))
        End If
        oSh.TextFrame.TextRange.Text = Trim(UCase(sText))
        If UCase(sText) <> "WHEEL OF FORTUNE" Then
            ActivePresentation.Slides(1).Shapes("WheelofFortuneLogo").Visible = False
        Else:
            ActivePresentation.Slides(1).Shapes("WheelofFortuneLogo").Visible = True
        End If
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

Sub puzzleScribe()
    Dim i As Integer, scribeWarning, scribeWarningConfirm
    For i = 1 To 52
        If SlideShowWindows(1).View.Slide.Shapes("BoardTile" & CStr(i)).OLEFormat.Object.Value <> "" Then
            scribeWarning = True
            Exit For
        End If
    Next i
    If scribeWarning = True Then
        scribeWarningConfirm = MsgBox("Warning: Puzzle Scribe will overwrite the existing puzzle. Do you want to continue?", vbYesNo + vbDefaultButton1, "Puzzle Scribe Overwrite Warning")
        If scribeWarningConfirm = vbYes Then
            puzzleScribe2 (True)
            Exit Sub
        Else:
            Exit Sub
        End If
    End If
    puzzleScribe2 (False)
End Sub

Private Sub puzzleScribe2(clearBoard As Boolean)
    Dim sText As String, isValidPuzzle As Boolean, puzzleSplitted As Variant
    Dim i As Integer, iPointer As Integer
    Dim puzzleRowLengths(3) As Variant
    Dim erroredRow As Integer
    isValidPuzzle = False
    sText = InputBox("Type your puzzle here, and it'll automatically write onto the tiles. Separate rows with |." & vbNewLine & vbNewLine & _
    "Example" & vbNewLine & "puzzle scribe | saves me time", "Puzzle Scribe")
    Do Until sText = ""
        puzzleSplitted = Split(sText, "|")
        If UBound(puzzleSplitted) + 1 <= 2 Then
            puzzleRowLengths(0) = -1
            For i = LBound(puzzleSplitted) To UBound(puzzleSplitted)
                If Len(Trim(puzzleSplitted(i))) > 14 Then
                    erroredRow = i + 1
                    GoTo notValidPuzzle
                Else:
                    puzzleRowLengths(i + 1) = Len(Trim(puzzleSplitted(i)))
                End If
            Next i
            puzzleRowLengths(3) = -1
            Exit Do
        ElseIf UBound(puzzleSplitted) + 1 = 3 Then
            iPointer = 0
            For i = LBound(puzzleSplitted) To UBound(puzzleSplitted)
                If Len(Trim(puzzleSplitted(i))) > 14 Then
                    erroredRow = i + 1
                    GoTo notValidPuzzle
                ElseIf Len(Trim(puzzleSplitted(i))) > 12 Then
                    If iPointer = 0 Then
                        puzzleRowLengths(0) = -1
                        iPointer = 1
                    ElseIf iPointer = 3 Then
                        erroredRow = i + 1
                        GoTo notValidPuzzle
                    End If
                        puzzleRowLengths(iPointer) = Len(Trim(puzzleSplitted(i)))
                Else:
                    If iPointer = 0 Then
                        puzzleRowLengths(3) = -1
                    End If
                    puzzleRowLengths(iPointer) = Len(Trim(puzzleSplitted(i)))
                End If
                iPointer = iPointer + 1
            Next i
            Exit Do
        ElseIf UBound(puzzleSplitted) + 1 = 4 Then
            For i = LBound(puzzleSplitted) To UBound(puzzleSplitted)
                If Len(Trim(puzzleSplitted(i))) > 14 Then
                    erroredRow = i + 1
                    GoTo notValidPuzzle
                ElseIf Len(Trim(puzzleSplitted(i))) > 12 Then
                    If i = 0 Or i = 3 Then
                        erroredRow = i + 1
                        GoTo notValidPuzzle
                    Else:
                        puzzleRowLengths(i) = Len(Trim(puzzleSplitted(i)))
                    End If
                Else:
                    puzzleRowLengths(i) = Len(Trim(puzzleSplitted(i)))
                End If
            Next i
            Exit Do
        Else:
            erroredRow = 5
            GoTo notValidPuzzle
        End If
notValidPuzzle:
        If erroredRow < 5 Then
            sText = InputBox("Row " & erroredRow & " is too long for the puzzle board. Try again." & vbNewLine & vbNewLine & _
            "Type your puzzle here, and it'll automatically write onto the tiles. Separate rows with |." & vbNewLine & vbNewLine & _
            "Example" & vbNewLine & "puzzle scribe | saves me time", "Puzzle Scribe", sText)
        Else:
            sText = InputBox("This puzzle has more than the allowed four rows. Try again." & vbNewLine & vbNewLine & _
            "Type your puzzle here, and it'll automatically write onto the tiles. Separate rows with |." & vbNewLine & vbNewLine & _
            "Example" & vbNewLine & "puzzle scribe | saves me time", "Puzzle Scribe", sText)
        End If
    Loop
    If sText = "" Then
        Exit Sub
    Else:
        Dim j As Integer, k As Integer, n As Integer, p As Integer, maxRowLength As Integer
        maxRowLength = 0
        ' Clear puzzle board if necessary
        If clearBoard = True Then
            For j = 1 To 52
                SlideShowWindows(1).View.Slide.Shapes("BoardTile" & CStr(j)).OLEFormat.Object.Value = ""
            Next j
        End If
        ' Determine maximum puzzle row length for letter placement
        For k = LBound(puzzleRowLengths) To UBound(puzzleRowLengths)
            If puzzleRowLengths(k) > maxRowLength Then
                maxRowLength = puzzleRowLengths(k)
            End If
        Next k
        ' Inscribe puzzle onto tiles
        Dim rowsProcessed As Integer, startingTile As Integer
        rowsProcessed = 0
        For n = 0 To 3
            If puzzleRowLengths(n) > -1 Then
                If n = 0 Then
                    startingTile = 7 - CInt(maxRowLength / 2 + 0.0001)
                    If startingTile < 1 Then
                        startingTile = 1
                    End If
                ElseIf n = 1 Then
                    startingTile = 20 - CInt(maxRowLength / 2 + 0.0001)
                ElseIf n = 2 Then
                    startingTile = 34 - CInt(maxRowLength / 2 + 0.0001)
                ElseIf n = 3 Then
                    startingTile = 47 - CInt(maxRowLength / 2 + 0.0001)
                    If startingTile < 41 Then
                        startingTile = 41
                    End If
                End If
                For p = 1 To Len(Trim(puzzleSplitted(rowsProcessed)))
                    SlideShowWindows(1).View.Slide.Shapes("BoardTile" & CStr(startingTile + p - 1)).OLEFormat.Object.Value = Mid(Trim(puzzleSplitted(rowsProcessed)), p, 1)
                Next p
                rowsProcessed = rowsProcessed + 1
            End If
        Next n
    End If
    Exit Sub
End Sub

Sub puzzleProperties()
    Dim i As Integer, currentLetter As String, convertedLetter As String, RSTLNELetters As Integer, uniqueConsonants As Variant, uniqueVowels As Variant
    Dim AChars As String, CChars As String, EChars As String, IChars As String, OChars As String, UChars As String
    AChars = ActivePresentation.Slides(9).Shapes("AChars").TextFrame.TextRange.Text
    CChars = ActivePresentation.Slides(9).Shapes("CChars").TextFrame.TextRange.Text
    EChars = ActivePresentation.Slides(9).Shapes("EChars").TextFrame.TextRange.Text
    IChars = ActivePresentation.Slides(9).Shapes("IChars").TextFrame.TextRange.Text
    OChars = ActivePresentation.Slides(9).Shapes("OChars").TextFrame.TextRange.Text
    UChars = ActivePresentation.Slides(9).Shapes("UChars").TextFrame.TextRange.Text
    Dim totalLetters As Integer, totalVowels As Integer
    totalLetters = 0
    totalVowels = 0
    RSTLNELetters = 0
    For i = 1 To 52
        currentLetter = ActivePresentation.Slides(8).Shapes("BoardTile" + CStr(i)).OLEFormat.Object.Value
        ' Check if current tile has a valid letter
        If isLetter(currentLetter) = True Then
            ' Convert letter to non-accented form
            If InStr(AChars, currentLetter) > 0 Then
                convertedLetter = "A"
            ElseIf InStr(CChars, currentLetter) > 0 Then
                convertedLetter = "C"
            ElseIf InStr(EChars, currentLetter) > 0 Then
                convertedLetter = "E"
            ElseIf InStr(IChars, currentLetter) > 0 Then
                convertedLetter = "I"
            ElseIf InStr(OChars, currentLetter) > 0 Then
                convertedLetter = "O"
            ElseIf InStr(UChars, currentLetter) > 0 Then
                convertedLetter = "U"
            Else:
                convertedLetter = currentLetter
            End If
            ' Check if letter is a vowel
            If isVowel(convertedLetter) = True Then
                ' Check uniqueness of vowel
                If IsEmpty(uniqueVowels) Then
                    ReDim uniqueVowels(0)
                    uniqueVowels(0) = convertedLetter
                ElseIf isInArray(convertedLetter, uniqueVowels) = False Then
                    ReDim Preserve uniqueVowels(UBound(uniqueVowels) + 1)
                    uniqueVowels(UBound(uniqueVowels)) = convertedLetter
                End If
                ' Add vowel to counter
                totalVowels = totalVowels + 1
            Else:
                ' Check uniqueness of consonant
                If IsEmpty(uniqueConsonants) Then
                    ReDim uniqueConsonants(0)
                    uniqueConsonants(0) = convertedLetter
                ElseIf isInArray(convertedLetter, uniqueConsonants) = False Then
                    ReDim Preserve uniqueConsonants(UBound(uniqueConsonants) + 1)
                    uniqueConsonants(UBound(uniqueConsonants)) = convertedLetter
                End If
            End If
            ' Check if letter is RSTLNE
            If convertedLetter = "R" Or convertedLetter = "S" Or convertedLetter = "T" _
            Or convertedLetter = "L" Or convertedLetter = "N" Or convertedLetter = "E" Then
                RSTLNELetters = RSTLNELetters + 1
            End If
            ' Add letter to total letters count
            totalLetters = totalLetters + 1
        End If
    Next i
    ' Calculate unique letters
    Dim uniqueVowelsCount As Integer, uniqueConsonantsCount As Integer
    If IsEmpty(uniqueVowels) Then
        uniqueVowelsCount = 0
    Else:
        uniqueVowelsCount = UBound(uniqueVowels) + 1
    End If
    If IsEmpty(uniqueConsonants) Then
        uniqueConsonantsCount = 0
    Else:
        uniqueConsonantsCount = UBound(uniqueConsonants) + 1
    End If
    ' Calculate RSTLNE Ratio
    Dim RSTLNERatio As Integer
    If RSTLNELetters = 0 And totalLetters = 0 Then
        RSTLNERatio = 0
    Else:
        RSTLNERatio = CInt((RSTLNELetters / totalLetters) * 100 + 0.00000001)
    End If
    ' Output Puzzle Properties
    MsgBox "TOTAL LETTERS: " & totalLetters & vbNewLine & _
    "Consonants: " & totalLetters - totalVowels & vbNewLine & _
    "Vowels: " & totalVowels & vbNewLine & vbNewLine & _
    "UNIQUE LETTERS: " & (uniqueConsonantsCount + uniqueVowelsCount) & vbNewLine & _
    "Consonants: " & uniqueConsonantsCount & vbNewLine & _
    "Vowels: " & uniqueVowelsCount & vbNewLine & vbNewLine & _
    "RSTLNE RATIO: " & RSTLNERatio & "%", 0, "Puzzle Properties"
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
        If ActivePresentation.Slides(2).Shapes("PuzzleBoard" & k).Fill.ForeColor.RGB = RGB(255, 255, 255) And ActivePresentation.Slides(2).Shapes("PuzzleBoard" & k).TextFrame.TextRange.Text = "" Then
            If lettersMatch(ActivePresentation.Slides(2).Shapes("PuzzleCache" & k).TextFrame.TextRange.Text, theLetter) Then
                If ActivePresentation.Slides(9).Shapes("BlueTiles").TextFrame.TextRange.Text = "off" Then
                    ActivePresentation.Slides(2).Shapes("PuzzleBoard" & k).TextFrame.TextRange.Text = ActivePresentation.Slides(2).Shapes("PuzzleCache" & k).TextFrame.TextRange.Text
                Else:
                    ActivePresentation.Slides(2).Shapes("PuzzleBoard" & k).Fill.ForeColor.RGB = RGB(0, 0, 255)
                End If
            End If
        End If
    Next k
    ActivePresentation.Slides(2).Shapes("Letter" & i).TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = False
End Sub

Private Sub resetBonusRound()
    If ActivePresentation.Slides(2).Shapes("BonusOutline").Line.Transparency = 0 Then
        Do While Len(ActivePresentation.Slides(2).Shapes("BonusLetters").TextFrame.TextRange.Text) > 0
            removeBonusLetter
        Loop
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
                If ActivePresentation.Slides(2).Shapes("PuzzleBoard" & k).Fill.ForeColor.RGB = RGB(255, 255, 255) And ActivePresentation.Slides(2).Shapes("PuzzleBoard" & k).TextFrame.TextRange.Text = "" Then
                    If lettersMatch(ActivePresentation.Slides(2).Shapes("PuzzleCache" & k).TextFrame.TextRange.Text, letterToGuess) Then
                        If ActivePresentation.Slides(9).Shapes("BlueTiles").TextFrame.TextRange.Text = "off" Then
                            ActivePresentation.Slides(2).Shapes("PuzzleBoard" & k).TextFrame.TextRange.Text = ActivePresentation.Slides(2).Shapes("PuzzleCache" & k).TextFrame.TextRange.Text
                        Else:
                            ActivePresentation.Slides(2).Shapes("PuzzleBoard" & k).Fill.ForeColor.RGB = RGB(0, 0, 255)
                        End If
                        letterExist = True
                    End If
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
            ActivePresentation.Slides(11).Shapes("GuessLetterCorrect").ActionSettings(ppMouseClick).SoundEffect.Play
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
    If i = False Then
        disableFinalSpin
    End If
End Sub

Sub toggleFinalSpin()
    On Error GoTo errHandler
    If ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text = ActivePresentation.Slides(9).Shapes("GameName").TextFrame.TextRange.Text Then
        MsgBox "Load a puzzle first before starting the Final Spin.", 0, "Final Spin Error"
        Exit Sub
    End If
    If ActivePresentation.Slides(2).Shapes("FinalSpinButton").Visible = True Then
        If ActivePresentation.Slides(2).Shapes("FinalSpinButton").Line.Transparency = 1 Then
            ActivePresentation.Slides(2).Shapes("FinalSpinButton").Fill.ForeColor.RGB = RGB(225, 129, 75)
            ActivePresentation.Slides(2).Shapes("FinalSpinButton").Fill.Transparency = 0
            ActivePresentation.Slides(2).Shapes("FinalSpinButton").Line.Transparency = 0
            ActivePresentation.Slides(2).Shapes("FInalSpinBanner").Visible = True
            ActivePresentation.Slides(2).Shapes("BlackCover").Visible = True
            ActivePresentation.Slides(2).Shapes("ManualFinalSpin").Visible = True
            ActivePresentation.Slides(11).Shapes("FinalSpinAlert").ActionSettings(ppMouseClick).SoundEffect.Play
            Exit Sub
        Else:
            disableFinalSpin
        End If
    End If
errHandler:
    Exit Sub
End Sub

Private Sub disableFinalSpin()
    ActivePresentation.Slides(2).Shapes("FinalSpinButton").Fill.ForeColor.RGB = RGB(166, 166, 166)
    ActivePresentation.Slides(2).Shapes("FinalSpinButton").Fill.Transparency = 0.5
    ActivePresentation.Slides(2).Shapes("FinalSpinButton").Line.Transparency = 1
    ActivePresentation.Slides(2).Shapes("BlackCover").Visible = False
    ActivePresentation.Slides(2).Shapes("FinalSpinBanner").Visible = False
    ActivePresentation.Slides(2).Shapes("ManualFinalSpin").Visible = False
    If ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text <> ActivePresentation.Slides(9).Shapes("GameName").TextFrame.TextRange.Text Then
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
    MsgBox "Please reset the bonus round timer before switching rounds.", 0, "Exit Bonus Round Error"
End Sub

Sub randomSpin()
    On Error GoTo errHandler
    Dim x As Integer, rand As Integer, realRand As Double, effNew As Effect, effNew2 As Effect, effNew3 As Effect, effNew4 As Effect
    ' Clear letter counter
    ActivePresentation.Slides(2).Shapes("LetterCounter").TextFrame.TextRange.Text = ""
    ActivePresentation.Slides(2).Shapes("LetterSelectionOverlay2").Visible = False
Spin:
    ' Remove wheel animations
    For x = ActivePresentation.SlideShowWindow.View.Slide.TimeLine.MainSequence.Count To 1 Step -1
        ActivePresentation.SlideShowWindow.View.Slide.TimeLine.MainSequence.Item(x).Delete
    Next x
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
    If IsNumeric(Replace(ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text, "$", "")) Then
        ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text = CLng(Replace(ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text, "$", ""))
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
    ' If Final Spin, respin if spun value is non-numeric
    If ActivePresentation.Slides(2).Shapes("FinalSpinButton").Visible = True Then
        If ActivePresentation.Slides(2).Shapes("FinalSpinButton").Line.Transparency = 0 And _
        Not IsNumeric(Replace(ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text, "$", "")) Then
            GoTo Spin
        End If
    End If
    setValuePanelDisplay
    ActivePresentation.Slides(11).Shapes("SpinWheel").ActionSettings(ppMouseClick).SoundEffect.Play
    Exit Sub
errHandler:
    Exit Sub
End Sub

Sub manuallySetValuePanel()
    Dim sText As String
    If ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text = ActivePresentation.Slides(9).Shapes("GameName").TextFrame.TextRange.Text Then
        MsgBox "Load a puzzle first before manually setting the spun wheel value.", 0, "Manually Set Value Panel"
        Exit Sub
    End If
    sText = InputBox("Manually set the spun wheel value:", "Manually Set Value Panel", ActivePresentation.Slides(2).Shapes("SpunWheelValue").TextFrame.TextRange.Text)
    Do While IsNumeric(sText) = False And sText <> ""
        sText = InputBox("You can only enter numbers here. Try again:", "Manually Set Value Panel", sText)
    Loop
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
            ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text = "** " & ActivePresentation.Slides(2).Shapes("ValuePanel").TextFrame.TextRange.Text _
        & " **"
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
    Do While IsNumeric(numberToSwap) = False:
        If numberToSwap = "" Then
            Exit Sub
        Else:
            numberToSwap = InputBox("Please enter a number:", "Puzzle Swap", numberToSwap)
        End If
    Loop
    puzzleSwap CInt(oSh.TextFrame.TextRange.Text), CInt(numberToSwap)
End Sub

Private Sub puzzleSwap(i As Integer, j As Integer)
    Dim iPuzzleNumberIndex As Integer, jPuzzleNumberIndex As Integer, l As Integer, m As Integer, n As Integer, o As Integer
    iPuzzleNumberIndex = Int((i - 1) / 12)
    jPuzzleNumberIndex = Int((j - 1) / 12)
    If jPuzzleNumberIndex + 1 <= ActivePresentation.SectionProperties.SlidesCount(4) Then
        ' Move initial puzzle to swap to cache
        For l = 1 To 52
            ActivePresentation.Slides(9).Shapes("PuzzleSolutionSwap-" & CStr(l)).TextFrame.TextRange.Text = ActivePresentation.Slides(12 + iPuzzleNumberIndex).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(l)).TextFrame.TextRange.Text
            ActivePresentation.Slides(9).Shapes("PuzzleSolutionSwap-" & CStr(l)).Fill.ForeColor.RGB = ActivePresentation.Slides(12 + iPuzzleNumberIndex).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(l)).Fill.ForeColor.RGB
        Next l
        ActivePresentation.Slides(9).Shapes("PuzzleCategorySwap").TextFrame.TextRange.Text = ActivePresentation.Slides(12 + iPuzzleNumberIndex).Shapes("PuzzleCategory" & CStr(i)).TextFrame.TextRange.Text
        ' Overwrite contents of initial puzzle number with the puzzle number you want to swap with
        For m = 1 To 52
            ActivePresentation.Slides(12 + iPuzzleNumberIndex).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(m)).TextFrame.TextRange.Text = ActivePresentation.Slides(12 + jPuzzleNumberIndex).Shapes("PuzzleSolution" & CStr(j) & "-" & CStr(m)).TextFrame.TextRange.Text
            ActivePresentation.Slides(12 + iPuzzleNumberIndex).Shapes("PuzzleSolution" & CStr(i) & "-" & CStr(m)).Fill.ForeColor.RGB = ActivePresentation.Slides(12 + jPuzzleNumberIndex).Shapes("PuzzleSolution" & CStr(j) & "-" & CStr(m)).Fill.ForeColor.RGB
        Next m
        ActivePresentation.Slides(12 + iPuzzleNumberIndex).Shapes("PuzzleCategory" & CStr(i)).TextFrame.TextRange.Text = ActivePresentation.Slides(12 + jPuzzleNumberIndex).Shapes("PuzzleCategory" & CStr(j)).TextFrame.TextRange.Text
        ' Overwrite contents of the puzzle number you want to swap with with what's in the cache
        For n = 1 To 52
            ActivePresentation.Slides(12 + jPuzzleNumberIndex).Shapes("PuzzleSolution" & CStr(j) & "-" & CStr(n)).TextFrame.TextRange.Text = ActivePresentation.Slides(9).Shapes("PuzzleSolutionSwap-" & CStr(n)).TextFrame.TextRange.Text
            ActivePresentation.Slides(12 + jPuzzleNumberIndex).Shapes("PuzzleSolution" & CStr(j) & "-" & CStr(n)).Fill.ForeColor.RGB = ActivePresentation.Slides(9).Shapes("PuzzleSolutionSwap-" & CStr(n)).Fill.ForeColor.RGB
        Next n
        ActivePresentation.Slides(12 + jPuzzleNumberIndex).Shapes("PuzzleCategory" & CStr(j)).TextFrame.TextRange.Text = ActivePresentation.Slides(9).Shapes("PuzzleCategorySwap").TextFrame.TextRange.Text
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
    ActivePresentation.Slides(12).Duplicate
    For k = 1 To 12
        ActivePresentation.Slides(13).Shapes("LinkTo" & k).TextFrame.TextRange.Text = (k + (12 * num))
        ActivePresentation.Slides(13).Shapes("LinkTo" & k).Name = "LinkTo" & (k + (12 * num))
        ActivePresentation.Slides(13).Shapes("Swap" & k).TextFrame.TextRange.Text = (k + (12 * num))
        ActivePresentation.Slides(13).Shapes("Swap" & k).Name = "Swap" & (k + (12 * num))
        ActivePresentation.Slides(13).Shapes("PuzzleCategory" & k).TextFrame.TextRange.Text = ""
        ActivePresentation.Slides(13).Shapes("PuzzleCategory" & k).Name = "PuzzleCategory" & (k + (12 * num))
        For l = 1 To 52
            If ActivePresentation.Slides(13).Shapes("PuzzleSolution" & CStr(k) & "-" & CStr(l)).TextFrame.TextRange.Text <> "" Then
                ActivePresentation.Slides(13).Shapes("PuzzleSolution" & CStr(k) & "-" & CStr(l)).TextFrame.TextRange.Text = ""
                ActivePresentation.Slides(13).Shapes("PuzzleSolution" & CStr(k) & "-" & CStr(l)).Fill.ForeColor.RGB = RGB(24, 154, 80)
            End If
            ActivePresentation.Slides(13).Shapes("PuzzleSolution" & CStr(k) & "-" & CStr(l)).Name = "PuzzleSolution" & CStr(k + (12 * num)) & "-" & CStr(l)
        Next l
    Next k
    ActivePresentation.Slides(13).Shapes("PrevAllPuzzles").Visible = msoTrue
    ActivePresentation.Slides(13).Shapes("NextAllPuzzles").Visible = msoFalse
    ActivePresentation.Slides(13).MoveTo toPos:=12 + ActivePresentation.SectionProperties.SlidesCount(4) - 1
    ActivePresentation.Slides(12 + ActivePresentation.SectionProperties.SlidesCount(4) - 2).Shapes("NextAllPuzzles").Visible = msoTrue
End Sub

Sub nextPuzzleRow()
    Dim r As Integer, p As Integer, RowIndex As Integer
    savePuzzle
    RowIndex = CInt(ActivePresentation.Slides(7).Shapes("CurrentPuzzleRowIndex").TextFrame.TextRange.Text)
    ActivePresentation.Slides(7).Shapes("CurrentPuzzleRowIndex").TextFrame.TextRange.Text = RowIndex + 1
    If ActivePresentation.SectionProperties.SlidesCount(4) <= RowIndex + 1 Then
        addPuzzleRow (RowIndex + 1)
    End If
    For r = 7 To 9
        For p = 1 To 12
            With ActivePresentation.Slides(r).Shapes("LinkTo" & CStr(p + (12 * RowIndex)))
                .Name = "LinkTo" & CStr(p + (12 * (RowIndex + 1)))
                .TextFrame.TextRange.Text = CStr(p + (12 * (RowIndex + 1)))
            End With
        Next p
        ActivePresentation.Slides(r).Shapes("PrevPuzzleRow").Visible = msoTrue
    Next r
    puzzleSetupJump (1 + (12 * (RowIndex + 1)))
End Sub

Sub prevPuzzleRow()
    Dim r As Integer, p As Integer, RowIndex As Integer
    savePuzzle
    RowIndex = CInt(ActivePresentation.Slides(7).Shapes("CurrentPuzzleRowIndex").TextFrame.TextRange.Text)
    ActivePresentation.Slides(7).Shapes("CurrentPuzzleRowIndex").TextFrame.TextRange.Text = RowIndex - 1
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

Sub TogglePlayersPlus()
    If ActivePresentation.Slides(2).Shapes("NumPlayers").TextFrame.TextRange.Text = "2 Players" Then
        TogglePlayers (3)
    ElseIf ActivePresentation.Slides(2).Shapes("NumPlayers").TextFrame.TextRange.Text = "3 Players" Then
        TogglePlayers (4)
    End If
End Sub

Sub TogglePlayersMinus()
    If ActivePresentation.Slides(2).Shapes("NumPlayers").TextFrame.TextRange.Text = "4 Players" Then
        TogglePlayers (3)
    ElseIf ActivePresentation.Slides(2).Shapes("NumPlayers").TextFrame.TextRange.Text = "3 Players" Then
        TogglePlayers (2)
    End If
End Sub

Private Sub TogglePlayers(numPlayers As Integer)
    Dim i As Integer, j As Integer, k As Integer, m As Integer
    If numPlayers = 4 Then
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
            .Shapes("NumPlayers").Left = 34.35677
            .Shapes("RemovePlayers").Left = 19.42866
            .Shapes("NumPlayers").TextFrame.TextRange.Text = "4 Players"
            .Shapes("AddPlayers").Visible = False
            .Shapes("RemovePlayers").Visible = True
        End With
    ElseIf numPlayers = 3 Then
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
            .Shapes("NumPlayers").Left = 42.07858
            .Shapes("AddPlayers").Left = 88.9752
            .Shapes("RemovePlayers").Left = 27.15047
            .Shapes("NumPlayers").TextFrame.TextRange.Text = "3 Players"
            .Shapes("AddPlayers").Visible = True
            .Shapes("RemovePlayers").Visible = True
        End With
    Else:
        For k = 1 To 2:
        With ActivePresentation.Slides(2)
            .Shapes("Player" & k & "BuyVowelButton").Left = 193.0415 + 185.5516 * (k - 1)
            .Shapes("Player" & k & "TransferTotalsButton").Left = 193.0768 + 185.5516 * (k - 1)
            .Shapes("Player" & k & "RoundDollarSign").Left = 210.9989 + 185.5516 * (k - 1)
            .Shapes("Player" & k & "TotalsDollarSign").Left = 210.9989 + 185.5516 * (k - 1)
            .Shapes("Player" & k & "RoundScore").Left = 228.9989 + 185.5516 * (k - 1)
            .Shapes("Player" & k & "RoundScoreCompatibility").Left = 228.9989 + 185.5516 * (k - 1)
            .Shapes("Player" & k & "TotalsScore").Left = 225.3991 + 185.5516 * (k - 1)
            .Shapes("Player" & k & "Name").Left = 200.0615 + 185.5516 * (k - 1)
            .Shapes("Player" & k & "XButton").Left = 305.2326 + 185.5516 * (k - 1)
            .Shapes("Player" & k & "XButtonCompatibility").Left = 305.2326 + 185.5516 * (k - 1)
            .Shapes("Player" & k & "WildCard").Left = 327.5543 + 185.5516 * (k - 1)
            .Shapes("Player" & k & "GiftTag").Left = 324.7217 + 185.5516 * (k - 1)
        End With
        Next k
        For m = 3 To 4:
        With ActivePresentation.Slides(2)
            .Shapes("Player" & m & "BuyVowelButton").Left = 193.0415 + 185.5516 * (1.4 + m - 1)
            .Shapes("Player" & m & "TransferTotalsButton").Left = 193.0768 + 185.5516 * (1.4 + m - 1)
            .Shapes("Player" & m & "RoundDollarSign").Left = 210.9989 + 185.5516 * (1.4 + m - 1)
            .Shapes("Player" & m & "TotalsDollarSign").Left = 210.9989 + 185.5516 * (1.4 + m - 1)
            .Shapes("Player" & m & "RoundScore").Left = 228.9989 + 185.5516 * (1.4 + m - 1)
            .Shapes("Player" & m & "RoundScoreCompatibility").Left = 228.9989 + 185.5516 * (1.4 + m - 1)
            .Shapes("Player" & m & "TotalsScore").Left = 225.3991 + 185.5516 * (1.4 + m - 1)
            .Shapes("Player" & m & "Name").Left = 200.0615 + 185.5516 * (1.4 + m - 1)
            .Shapes("Player" & m & "XButton").Left = 305.2326 + 185.5516 * (1.4 + m - 1)
            .Shapes("Player" & m & "XButtonCompatibility").Left = 305.2326 + 185.5516 * (1.4 + m - 1)
            .Shapes("Player" & m & "WildCard").Left = 327.5543 + 185.5516 * (1.4 + m - 1)
            .Shapes("Player" & m & "GiftTag").Left = 324.7217 + 185.5516 * (1.4 + m - 1)
        End With
        Next m
        With ActivePresentation.Slides(2)
            .Shapes("RoundTotals").Left = 30.63039
            .Shapes("NumPlayers").Left = 34.57827
            .Shapes("AddPlayers").Left = 81.47488
            .Shapes("NumPlayers").TextFrame.TextRange.Text = "2 Players"
            .Shapes("AddPlayers").Visible = True
            .Shapes("RemovePlayers").Visible = False
        End With
    End If
End Sub

Sub ExplainVowelPrice()
    MsgBox "Choose how much it costs to buy a vowel.", 0, "Vowel Price Setting"
End Sub

Sub ExplainHouseMinimum()
    MsgBox "Choose how much the player wins if their round total is below the given amount.", 0, "House Minimum Setting"
End Sub

Sub ExplainShotClock()
    MsgBox "Add an optional in-game timer to the puzzle board with a given amount of seconds. Use it to enforce time limits on player decisions.", 0, "Shot Clock Setting"
End Sub

Sub ExplainBlueTiles()
    MsgBox "When a correct letter is called in the puzzle, choose whether the puzzle board tiles light up blue (requiring a click to reveal the letter) or if the letters show up instantly. The default is on.", 0, "Blue Tiles Setting"
End Sub

Sub ExplainNoMoreVowels()
    MsgBox "Choose whether to inform players when there are no more vowels or no more consonants in the puzzle. The default is on.", 0, "No More Vowels Setting"
End Sub

Sub ExplainClaimable()
    MsgBox "When a player collects a wheel item, choose whether the game removes the item from the wheel (claimable once) or leaves it for others to earn (claimable multiple). The default is once.", 0, "Claimable Setting"
End Sub

Sub ExplainBaseValue()
    MsgBox "Choose whether the base monetary value on the wheel is $300 or $500. The default is $300." & vbNewLine & vbNewLine & _
    "The $300 base offers a wider range of wheel values. The $500 base aligns with the actual show's wheel values (as of this writing).", 0, "Base Value Setting"
End Sub

Sub ExplainBackdrop()
    MsgBox "Choose the background scenery of the puzzle board and wheel.", 0, "Backdrop Setting"
End Sub

Sub ExplainSpanishN()
    MsgBox "If your puzzles are in Spanish, enable this setting to make " & ActivePresentation.Slides(9).Shapes("SpanN").TextFrame.TextRange.Text & " a selectable letter.", 0, "Spanish " & ActivePresentation.Slides(9).Shapes("SpanN").TextFrame.TextRange.Text & " Setting"
End Sub

Sub ExplainGameName()
    MsgBox "Edit the name of the game that displays on the puzzle board." & vbNewLine & vbNewLine & _
    "Game names other than WHEEL OF FORTUNE will also hide the Wheel of Fortune logo from the title slide.", 0, "Game Name Setting"
End Sub

Sub ExplainFreePlay()
    MsgBox "Choose whether to use the Free Play wedge on the wheel. The default is off." & vbNewLine & vbNewLine & _
    "Landing on Free Play lets the player call a consonant at $500 or a vowel for free. The player then gets another turn regardless if the letter they called is in the puzzle.", 0, "Free Play Wedge Setting"
End Sub

Sub ExplainNumberofBankrupts()
    MsgBox "Choose how many Bankrupt wedges are applied to each round. The default is 2 for all rounds.", 0, "Number of Bankrupts Setting"
End Sub

Sub ExplainWildCard()
    MsgBox "Choose whether to use the Wild Card item on the wheel. The default is off." & vbNewLine & vbNewLine & _
    "The Wild Card, when used by the player, lets them call another consonant worth the value of their prior spin. The player can also use it during the Bonus Round to call an extra consonant.", 0, "Wild Card Setting"
End Sub

Sub Explain10000Wedge()
    MsgBox "Choose whether to use the $10000 Wedge item on the wheel. The default is off." & vbNewLine & vbNewLine & _
    "The $10000 in the wedge is a thin sliver flanked by two thin Bankrupts. If a player lands on $10000 and calls a letter in the puzzle, they earn exactly $10000, regardless of the amount of the letter in the puzzle.", 0, "$10000 Wedge Setting"
End Sub

Sub ExplainGiftTag()
    MsgBox "Choose whether to use the Gift Tag item on the wheel. The default is off." & vbNewLine & vbNewLine & _
    "The Gift Tag is an auxiliary item that gives the player a prize of your choice.", 0, "Gift Tag Setting"
End Sub

Sub Explain5Wedge()
    MsgBox "Optionally place a $5 wedge on the wheel. A joke that's unfortunate and fortunate for the player (at least it's not a Bankrupt)." & vbNewLine & vbNewLine & _
    "The default is off.", 0, "$5 Wedge Setting"
End Sub

Sub doFinalSpin()
    ActivePresentation.Slides(2).Shapes("BlackCover").Visible = False
    ActivePresentation.Slides(2).Shapes("FinalSpinBanner").Visible = False
    ActivePresentation.Slides(2).Shapes("ManualFinalSpin").Visible = False
    SlideShowWindows(1).View.GotoSlide 6
    randomSpin
End Sub

Sub manualFinalSpin()
    manuallySetValuePanel
    setValuePanelDisplay
    ActivePresentation.Slides(2).Shapes("BlackCover").Visible = False
    ActivePresentation.Slides(2).Shapes("FinalSpinBanner").Visible = False
    ActivePresentation.Slides(2).Shapes("ManualFinalSpin").Visible = False
End Sub

Sub revealTossUpLetter()
    On Error GoTo errHandler
    Dim blankTiles As Variant, i As Integer, rand As Integer, isFirstReveal As Boolean
    isFirstReveal = True
    For i = 1 To 52
        If ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).Fill.ForeColor.RGB = RGB(255, 255, 255) Then
            If isLetter(ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).TextFrame.TextRange.Text) Then
                isFirstReveal = False
            End If
            If ActivePresentation.Slides(2).Shapes("PuzzleBoard" & i).TextFrame.TextRange.Text = "" Then
                If IsEmpty(blankTiles) Then
                    ReDim blankTiles(0)
                    blankTiles(0) = i
                Else:
                    ReDim Preserve blankTiles(UBound(blankTiles) + 1)
                    blankTiles(UBound(blankTiles)) = i
                End If
            End If
        End If
    Next i
    If IsEmpty(blankTiles) Then
        MsgBox "There are no more letters to reveal in this puzzle.", 0, "Reveal a Letter Error"
        Exit Sub
    Else:
        Randomize
        rand = Int((UBound(blankTiles) + 1) * Rnd)
        ActivePresentation.Slides(2).Shapes("PuzzleBoard" & blankTiles(rand)).TextFrame.TextRange.Text = ActivePresentation.Slides(2).Shapes("PuzzleCache" & blankTiles(rand)).TextFrame.TextRange.Text
        If isFirstReveal Then
            ActivePresentation.Slides(11).Shapes("TossUpMusic").ActionSettings(ppMouseClick).SoundEffect.Play
            Exit Sub
        End If
    End If
errHandler:
    Exit Sub
End Sub



