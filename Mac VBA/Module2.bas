Attribute VB_Name = "Module2"
Option Explicit

Sub youLandedOn(degrees As Double, wheelType As Integer)
    ' Determines the wheel value you've landed on based on degrees spun and the type of wheel used.
    ' Yes, there's a bunch of if statements which isn't very good performance wise (O(n)), but I don't
    ' know how to implement a hash table in PowerPoint VBA for O(1) performance.
    ' If you know how to improve the performance of this function, send a pull request:
    ' https://github.com/TimTree/wheel-of-fortune-ppt/pulls
    degrees = degrees - 1800
    If wheelType = 3 Then
        If degrees < 15 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Bankrupt"
        ElseIf degrees < 30 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$900"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$300"
            End If
        ElseIf degrees < 45 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        ElseIf degrees < 60 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$650"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$450"
            End If
        ElseIf degrees < 75 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        ElseIf degrees < 90 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$800"
        ElseIf degrees < 105 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Lose a Turn"
        ElseIf degrees < 120 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$700"
        ElseIf degrees < 135 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Free Play"
        ElseIf degrees < 150 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$650"
        ElseIf degrees < 165 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Bankrupt"
        ElseIf degrees < 180 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$600"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$900"
            End If
        ElseIf degrees < 195 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        ElseIf degrees < 210 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$550"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$350"
            End If
        ElseIf degrees < 225 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$600"
        ElseIf degrees < 230 Then
            If ActivePresentation.SlideShowWindow.View.Slide.Shapes("TheWheel").GroupItems("10000Wedge").Fill.Transparency = 1 Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Bankrupt"
            End If
        ElseIf degrees < 235 Then
            If ActivePresentation.SlideShowWindow.View.Slide.Shapes("TheWheel").GroupItems("10000Wedge").Fill.Transparency = 1 Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$10,000"
            End If
        ElseIf degrees < 240 Then
            If ActivePresentation.SlideShowWindow.View.Slide.Shapes("TheWheel").GroupItems("10000Wedge").Fill.Transparency = 1 Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Bankrupt"
            End If
        ElseIf degrees < 255 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$700"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$400"
            End If
        ElseIf degrees < 270 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        ElseIf degrees < 285 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$650"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$800"
            End If
        ElseIf degrees < 300 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$600"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$300"
            End If
        ElseIf degrees < 315 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$700"
        ElseIf degrees < 330 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$900"
        ElseIf degrees < 345 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        Else:
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$2500"
        End If
    ' Mystery Round
    ElseIf wheelType = 4 Then
        If degrees < 15 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Bankrupt"
        ElseIf degrees < 30 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$900"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$300"
            End If
        ElseIf degrees < 45 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        ElseIf degrees < 60 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$650"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$450"
            End If
        ElseIf degrees < 75 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        ElseIf degrees < 90 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$800"
        ElseIf degrees < 105 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Lose a Turn"
        ElseIf degrees < 120 Then
            If ActivePresentation.SlideShowWindow.View.Slide.Shapes("TheWheel").GroupItems("MysteryWedge1").Fill.Transparency = 1 Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$700"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Mystery 1"
            End If
        ElseIf degrees < 135 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Free Play"
        ElseIf degrees < 150 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$650"
        ElseIf degrees < 165 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Bankrupt"
        ElseIf degrees < 180 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$600"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$900"
            End If
        ElseIf degrees < 195 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        ElseIf degrees < 210 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$550"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$350"
            End If
        ElseIf degrees < 225 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$600"
        ElseIf degrees < 230 Then
            If ActivePresentation.SlideShowWindow.View.Slide.Shapes("TheWheel").GroupItems("10000Wedge").Fill.Transparency = 1 Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Bankrupt"
            End If
        ElseIf degrees < 235 Then
            If ActivePresentation.SlideShowWindow.View.Slide.Shapes("TheWheel").GroupItems("10000Wedge").Fill.Transparency = 1 Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$10,000"
            End If
        ElseIf degrees < 240 Then
            If ActivePresentation.SlideShowWindow.View.Slide.Shapes("TheWheel").GroupItems("10000Wedge").Fill.Transparency = 1 Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Bankrupt"
            End If
        ElseIf degrees < 255 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$700"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$400"
            End If
        ElseIf degrees < 270 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        ElseIf degrees < 285 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$650"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$800"
            End If
        ElseIf degrees < 300 Then
            If ActivePresentation.SlideShowWindow.View.Slide.Shapes("TheWheel").GroupItems("MysteryWedge2").Fill.Transparency = 1 Then
                If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                    ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$600"
                Else:
                    ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$300"
                End If
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Mystery 2"
            End If
        ElseIf degrees < 315 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$700"
        ElseIf degrees < 330 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$900"
        ElseIf degrees < 345 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        Else:
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$3500"
        End If
    ' Express Round
    ElseIf wheelType = 5 Then
        If degrees < 15 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Bankrupt"
        ElseIf degrees < 30 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$900"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$300"
            End If
        ElseIf degrees < 45 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        ElseIf degrees < 60 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$650"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$450"
            End If
        ElseIf degrees < 75 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        ElseIf degrees < 90 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$800"
        ElseIf degrees < 105 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Lose a Turn"
        ElseIf degrees < 120 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$700"
        ElseIf degrees < 135 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Free Play"
        ElseIf degrees < 150 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$650"
        ElseIf degrees < 165 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Bankrupt"
        ElseIf degrees < 180 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$600"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$900"
            End If
        ElseIf degrees < 195 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        ElseIf degrees < 210 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$550"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$350"
            End If
        ElseIf degrees < 225 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$600"
        ElseIf degrees < 230 Then
            If ActivePresentation.SlideShowWindow.View.Slide.Shapes("TheWheel").GroupItems("10000Wedge").Fill.Transparency = 1 Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Bankrupt"
            End If
        ElseIf degrees < 235 Then
            If ActivePresentation.SlideShowWindow.View.Slide.Shapes("TheWheel").GroupItems("10000Wedge").Fill.Transparency = 1 Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$10,000"
            End If
        ElseIf degrees < 240 Then
            If ActivePresentation.SlideShowWindow.View.Slide.Shapes("TheWheel").GroupItems("10000Wedge").Fill.Transparency = 1 Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Bankrupt"
            End If
        ElseIf degrees < 255 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$700"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$400"
            End If
        ElseIf degrees < 270 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        ElseIf degrees < 285 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$650"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$800"
            End If
        ElseIf degrees < 300 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$600"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$300"
            End If
        ElseIf degrees < 315 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Express"
        ElseIf degrees < 330 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$900"
        ElseIf degrees < 345 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        Else:
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$3500"
        End If
    ' Fourth Round
    ElseIf wheelType = 6 Then
        If degrees < 15 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Bankrupt"
        ElseIf degrees < 30 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$900"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$300"
            End If
        ElseIf degrees < 45 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        ElseIf degrees < 60 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$650"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$450"
            End If
        ElseIf degrees < 75 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        ElseIf degrees < 90 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$800"
        ElseIf degrees < 105 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Lose a Turn"
        ElseIf degrees < 120 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$700"
        ElseIf degrees < 135 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Free Play"
        ElseIf degrees < 150 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$650"
        ElseIf degrees < 165 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "Bankrupt"
        ElseIf degrees < 180 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$600"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$900"
            End If
        ElseIf degrees < 195 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        ElseIf degrees < 210 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$550"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$350"
            End If
        ElseIf degrees < 225 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$600"
        ElseIf degrees < 240 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        ElseIf degrees < 255 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$700"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$400"
            End If
        ElseIf degrees < 270 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        ElseIf degrees < 285 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$650"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$800"
            End If
        ElseIf degrees < 300 Then
            If ActivePresentation.Slides(9).Shapes("WheelValues").TextFrame.TextRange.Text = "current ($500 min)" Then
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$600"
            Else:
                ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$300"
            End If
        ElseIf degrees < 315 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$700"
        ElseIf degrees < 330 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$900"
        ElseIf degrees < 345 Then
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$500"
        Else:
            ActivePresentation.SlideShowWindow.View.Slide.Shapes("WheelValue").TextFrame.TextRange.Text = "$5000"
        End If
    End If
End Sub
