Attribute VB_Name = "Module2"
Option Explicit

Sub youLandedOn(degrees As Integer, wheelType As Integer)
    ' Determines the wheel value you've landed on based on degrees spun and the type of wheel used.
    Select Case degrees
    Case 0 To 149
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "Bankrupt"
    Case 150 To 299
        If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$900"
        Else:
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$300"
        End If
    Case 300 To 449
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$500"
    Case 450 To 599
        If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$650"
        Else:
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$450"
        End If
    Case 600 To 749
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$500"
    Case 750 To 899
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$800"
    Case 900 To 1049
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "Lose a Turn"
    Case 1050 To 1199
        If wheelType = 4 Then ' Mystery Round
            If ActivePresentation.SlideShowWindow.View.Slide.Shapes("TheWheel").GroupItems("MysteryWedge1").Fill.Transparency = 1 Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$700"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "Mystery 1"
            End If
        Else:
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$700"
        End If
    Case 1200 To 1349
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "Free Play"
    Case 1350 To 1499
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$650"
    Case 1500 To 1649
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "Bankrupt"
    Case 1650 To 1799
        If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$600"
        Else:
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$900"
        End If
    Case 1800 To 1949
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$500"
    Case 1950 To 2099
        If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$550"
        Else:
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$350"
        End If
    Case 2100 To 2249
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$600"
    ' Begin $10,000 wedge
    Case 2250 To 2299
        If wheelType = 6 Then ' Fourth Round
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$500"
        ElseIf ActivePresentation.SlideShowWindow.View.Slide.Shapes("TheWheel").GroupItems("10000Wedge").Fill.Transparency = 1 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$500"
        Else:
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "Bankrupt"
        End If
    Case 2300 To 2349
        If wheelType = 6 Then ' Fourth Round
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$500"
        ElseIf ActivePresentation.SlideShowWindow.View.Slide.Shapes("TheWheel").GroupItems("10000Wedge").Fill.Transparency = 1 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$500"
        Else:
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$10,000"
        End If
    Case 2350 To 2399
        If wheelType = 6 Then ' Fourth Round
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$500"
        ElseIf ActivePresentation.SlideShowWindow.View.Slide.Shapes("TheWheel").GroupItems("10000Wedge").Fill.Transparency = 1 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$500"
        Else:
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "Bankrupt"
        End If
    ' End $10,000 wedge
    Case 2400 To 2549
        If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$700"
        Else:
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$400"
        End If
    Case 2550 To 2699
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$500"
    Case 2700 To 2849
        If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$650"
        Else:
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$800"
        End If
    Case 2850 To 2999
        If wheelType = 4 Then ' Mystery Round
            If ActivePresentation.SlideShowWindow.View.Slide.Shapes("TheWheel").GroupItems("MysteryWedge2").Fill.Transparency = 1 Then
                If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                    ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$600"
                Else:
                    ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$300"
                End If
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "Mystery 2"
            End If
        Else:
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$600"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$300"
            End If
        End If
    Case 3000 To 3149
        If wheelType = 5 Then ' Express Round
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "Express"
        Else:
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$700"
        End If
    Case 3150 To 3299
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$900"
    Case 3300 To 3449
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$500"
    Case Else
        If wheelType = 3 Then ' First Round
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$2500"
        ElseIf wheelType = 4 Or wheelType = 5 Then ' Mystery/Express Round
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$3500"
        Else: ' Fourth Round
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange = "$5000"
        End If
    End Select
End Sub

