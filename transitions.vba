Sub SquareScaleTransition()
    ' needs better name
    ' assumes a slide width and height of 960 * 540. 
    ' (16:9 aspect ratio, units are arbitrary and probably pixels)
    ' creates 16 * 9 squares, then creates effect that scales them, staggered diagonally
    
    Dim slide As slide
    Dim x, y As Integer ' loop variables
    Dim shape As shape
    Dim enter As Boolean ' enter effect / exit effect
    
    Set slide = ActiveWindow.View.slide
    
    result = MsgBox("Do you want to use an enter effect? Yes = Enter effect, No = Exit effect", vbYesNo + vbQuestion, "Square Scale Transition")
    If result = vbYes Then
        enter = True
    ElseIf result = vbNo Then
        enter = False
    End If

    For x = 0 To (960 / 60) - 1 ' the 0 and the -1 is for alignment
        For y = 0 To (540 / 60) - 1 ' ditto
            Set shape = slide.Shapes.AddShape(msoShapeRectangle, x * 60, y * 60, 60, 60) ' make new shape
            With slide.TimeLine.MainSequence.AddEffect(shape, msoAnimEffectZoom, , msoAnimTriggerWithPrevious) ' add zoom effect
                .Timing.duration = 0.5 ' found that 0.5 seconds works, change if you want
                
                ' staggered effects:
                ' add the x and y coordinates and you get the delay time
                ' then transform it to get a shorter delay time
                ' example:
                '   0 1 2 3 4 5 6 7
                '  |_______________
                '0 |0 1 2 3 4 5 6 7
                '1 |1 2 3 4 5 6 7 8
                '2 |2 3 4 5 6 7 8 9
                '3 |3 4 5 6 7 8 9 10
                .Timing.TriggerDelayTime = (x + y) / 100
                
                
                ' True means it transtitions into the slide
                ' False means it transitions out of the slide
                If enter Then
                    .Exit = msoTrue
                    .EffectParameters.Direction = msoAnimDirectionOut
                Else:
                    .Exit = msoFalse
                    .EffectParameters.Direction = msoAnimDirectionIn
                End If
            End With
                
        Next
    Next
End Sub
Sub SquareScaleTransition_Image()
    ' needs better name
    ' assumes a slide width and height of 960 * 540. 
    ' (16:9 aspect ratio, units are arbitrary and probably pixels)
    ' with the user's selection image, creates 16 * 9 copies of the image
    ' then creates effect that scales them, staggered diagonally

    ' to use: select your target image and run the macro
    Dim slide As slide
    Dim x, y As Integer
    Dim shp As shape
    Dim enter As Boolean
    
    Dim baseShape As shape ' user selection
    
    Set slide = ActiveWindow.View.slide
    
    With ActiveWindow.Selection
        If .Type = ppSelectionShapes Then
            If .ShapeRange.count > 0 Then
                Set baseShape = .ShapeRange(1)
            End If
        End If
    End With
        
    If baseShape Is Nothing Then
        MsgBox "Please select an image first!", vbExclamation
        Exit Sub
    End If

    result = MsgBox("Do you want to use an enter effect? Yes = Enter effect, No = Exit effect", vbYesNo + vbQuestion, "Square Scale Transition")
    If result = vbYes Then
        enter = True
    ElseIf result = vbNo Then
        enter = False
    End If

    ' generate squares
    For x = 0 To (960 / 60) - 1 ' the 0 and the -1 is for alignment
        For y = 0 To (540 / 60) - 1 ' same here
            Set shp = baseShape.Duplicate(1)
            shp.Width = 60
            shp.Height = 60
            shp.Left = x * 60
            shp.Top = y * 60
            With slide.TimeLine.MainSequence.AddEffect(shp, msoAnimEffectZoom, , msoAnimTriggerWithPrevious) ' add zoom effect
                .Timing.duration = 0.5

                ' stagger animation
                .Timing.TriggerDelayTime = (x + y) / 50 ' this is very cool
                
                ' enter / exit
                If enter Then
                    .Exit = msoTrue
                    .EffectParameters.Direction = msoAnimDirectionOut
                Else:
                    .Exit = msoFalse
                    .EffectParameters.Direction = msoAnimDirectionIn
                End If
            End With
                
        Next
    Next
End Sub

Sub SquareScaleTransition_CopyAnim()
    ' needs better name
    ' assumes a slide width and height of 960 * 540. 
    ' (16:9 aspect ratio, units are arbitrary and probably pixels)
    ' with the user's selection image, creates 16 * 9 copies of the image
    ' then creates effect that scales them, staggered diagonally

    ' to use, select your target shape, with your desired animations
    ' (pulse/color change/whip/whatever)

    Dim sld As slide
    Dim x, y As Integer
    Dim shp As shape
    
    Dim baseShape As shape
    Dim newAnim As effect
    
    Set sld = ActiveWindow.View.slide
    
    With ActiveWindow.Selection
        If .Type = ppSelectionShapes Then
            If .ShapeRange.count > 0 Then
                Set baseShape = .ShapeRange(1)
            End If
        End If
    End With
        
    If baseShape Is Nothing Then
        MsgBox "Please select an image first!", vbExclamation
        Exit Sub
    End If
    
    Dim baseAnims As New Collection
    
    For Each possibleAnim In sld.TimeLine.MainSequence
        If possibleAnim.shape.Id = baseShape.Id Then
            ' got it
            baseAnims.Add possibleAnim
        End If
    Next
    If baseAnims.count = 0 Then
        MsgBox "Shape does not have any animations - cancelling..."
    End If
    
    
    ' generate squares
    For x = 0 To (960 / 60) - 1 ' the 0 and the -1 is for alignment
        For y = 0 To (540 / 60) - 1 ' ditto
            Set shp = baseShape.Duplicate(1)
            shp.Width = 60
            shp.Height = 60
            shp.Left = x * 60
            shp.Top = y * 60
            
            ' shape already has anims
            ' find existing animations and then adjust the delay
                For Each anim In sld.TimeLine.MainSequence
                    If anim.shape.Id = shp.Id Then
                        ' gottem
                        anim.Timing.TriggerDelayTime = (x + y) / 50 + anim.Timing.TriggerDelayTime
                    End If
                Next
        Next
    Next
End Sub

Sub ApplyPositionBasedAnimation()
    ' done by big gpt
    ' with all selected shapes, adds staggered animations diagonally
    ' customize animation yourself if yw
    Dim sld As slide
    Dim shp As shape
    Dim eff As effect

    
    Dim minX As Single, maxX As Single
    Dim minY As Single, maxY As Single
    
    Dim normX As Single, normY As Single
    Dim delay As Single
    
    Dim duration As Single
    duration = 0.5
    
    Set sld = ActiveWindow.View.slide
    
    ' Validate selection
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Select some shapes first!", vbExclamation
        Exit Sub
    End If
    
    If ActiveWindow.Selection.ShapeRange.count = 0 Then Exit Sub
    
    ' --- Find bounds (for normalization) ---
    minX = 999999: maxX = -999999
    minY = 999999: maxY = -999999
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If shp.Left < minX Then minX = shp.Left
        If shp.Left > maxX Then maxX = shp.Left
        If shp.Top < minY Then minY = shp.Top
        If shp.Top > maxY Then maxY = shp.Top
    Next
    
    ' Avoid division by zero
    If maxX = minX Then maxX = minX + 1
    If maxY = minY Then maxY = minY + 1
    
    ' --- Apply animation ---
    For Each shp In ActiveWindow.Selection.ShapeRange
        
        ' Normalize position (0 ? 1)
        normX = (shp.Left - minX) / (maxX - minX)
        normY = (shp.Top - minY) / (maxY - minY)
        
        ' Combine X + Y influence (diagonal wave)
        delay = (normX + normY) / 2 * duration
        
        ' —— ADD ANIMATION ——
        ' replace effectId with your desired effect Id
        ' TODO: improve selection
        Set eff = sld.TimeLine.MainSequence.AddEffect( _
            shape:=shp, _
            effectId:=msoAnimEffectWhip, _
            Trigger:=msoAnimTriggerWithPrevious)
        
        ' Apply timing
        With eff.Timing
            .duration = duration
            .TriggerDelayTime = delay
        End With
        
    Next
    
End Sub

Sub RectangleTransition()
    ' self explanatory, creates rectangles that fly out
    
    Dim slide As slide
    Dim slideWidth, slideHeight As Integer
    Dim i As Integer
    Dim count As Integer
    Dim along As Integer
    Dim shape As shape
    Dim duration, delay As Single
    
    Set slide = ActiveWindow.View.slide
    slideWidth = 960
    slideHeight = 540

    ' input
    count = InputBox("Enter the number of rectangles to create.", "Rectangle Creator")
    along = MsgBox("Select Yes to create WIDE RECTANGLES" & vbCrLf & "Select No to create TALL RECTANGLES", vbYesNoCancel, "Rectangle Creator")
    duration = InputBox("Enter the duration for each animation.", "Rectangle Creator")
    delay = InputBox("Enter the delay for each animation.", "Rectangle Creator")
    
    If along = vbCancel Then
        Exit Sub
    End If
    
    If along = vbYes Then
        ' WIDE RECTANGLES
        For i = 0 To count - 1
            Set shape = slide.Shapes.AddShape(msoShapeRectangle, 0, i * slideHeight / count, slideWidth, slideHeight / count) ' make new shape
            With slide.TimeLine.MainSequence.AddEffect(shape, msoAnimEffectFly, , msoAnimTriggerWithPrevious)
                .Timing.duration = duration
                .Timing.TriggerDelayTime = delay * i
            End With
        Next
    ElseIf along = vbNo Then
        ' TALL RECTANGLES
        For i = 0 To count - 1
            Set shape = slide.Shapes.AddShape(msoShapeRectangle, i * slideWidth / count, 0, slideWidth / count, slideHeight) ' make new shape
            With slide.TimeLine.MainSequence.AddEffect(shape, msoAnimEffectFly, , msoAnimTriggerWithPrevious)
                .Timing.duration = duration
                .Timing.TriggerDelayTime = delay * i
        End With
        Next
    End If
End Sub


Sub CreateRectangles()
    'creates rectangles, can create along length or height

    Dim slide As slide
    Dim slideWidth, slideHeight As Integer
    Dim i As Integer
    Dim count As Integer
    Dim along As Integer
    Dim shape As shape
    
    Set slide = ActiveWindow.View.slide
    slideWidth = 960
    slideHeight = 540

    ' input
    count = InputBox("Enter the number of rectangles to create.", "Rectangle Creator")
    along = MsgBox("Do you want wide rectangles?" & vbCrLf & "Select Yes to create WIDE RECTANGLES" & vbCrLf & "Select No to create TALL RECTANGLES", vbYesNoCancel, "Selection")
    If along = vbCancel Then
        Exit Sub
    End If
    
    If along = vbYes Then
        ' WIDE RECTANGLES
        For i = 0 To count - 1
            Set shape = slide.Shapes.AddShape(msoShapeRectangle, 0, i * slideHeight / count, slideWidth, slideHeight / count) ' make new shape
        Next
    ElseIf along = vbNo Then
        ' TALL RECTANGLES
        For i = 0 To count - 1
            Set shape = slide.Shapes.AddShape(msoShapeRectangle, i * slideWidth / count, 0, slideWidth / count, slideHeight) ' make new shape
        Next
    End If
End Sub

Sub GetWidthHeight()
    ' mainly for debugging purposes, use if your slide size is not 960 * 540 (16:9)
    With ActiveWindow.Selection
        MsgBox ("Width: " & .ShapeRange(1).Width & " Height: " & .ShapeRange(1).Height)
    End With
End Sub

Sub ConcentricRectangles()
    Dim slide As slide
    Dim slideWidth, slideHeight As Integer
    Dim i As Integer
    Dim count As Integer
    Dim shape As shape
    
    Set slide = ActiveWindow.View.slide
    slideWidth = 960
    slideHeight = 540

    ' input
    count = InputBox("Enter the number of concentric rectangles to create.", "Rectangle Creator")
    For i = 0 To count - 1
        Set shape = slide.Shapes.AddShape(msoShapeRectangle, slideWidth / 2, slideHeight / 2, (count - i) * 100, (count - i) * 100) ' make new shape
        shape.Left = slideWidth / 2 - shape.Width / 2
        shape.Top = slideHeight / 2 - shape.Height / 2
    Next
End Sub

' this is what i used to create that portal effect from geometry dash btw
' if you want that effect:
' change text color to none/transparent
' change shape to a circle (elipse) with gradient that goes from 0 to transparent
' (ensure that gradient goes to transparent before 100%, e.g. 0%=green, 67%=transparent, 100% = transparent)'
' stretch the shapes so they long, move them up if needed
' in this order: select shapes, group, duplicate, align with original group, move below slide, flip the group, ungroup, flip all shapes
' then move the new shapes to a new slide and add morph
Sub CreateSquareParticles()
    ' generates squares of varying sizes above the slide

    Dim sld As slide
    Dim shp As shape
    Dim i As Integer
    
    Dim slideWidth As Single
    Dim slideHeight As Single
    
    Set sld = ActivePresentation.Slides(ActiveWindow.View.slide.SlideIndex)
    
    slideWidth = ActivePresentation.PageSetup.slideWidth
    slideHeight = ActivePresentation.PageSetup.slideHeight
    
    Randomize
    
    For i = 1 To 80
        
        Dim size As Single
        size = 5 + Rnd * 10
        
        Dim startX As Single
        Dim startY As Single
        
        ' more horizontal spread control
        startX = Rnd * slideWidth
        
        ' vertical spread ABOVE slide (not all at same line)
        startY = -(Rnd * 400 + 10)
        
        Set shp = sld.Shapes.AddShape( _
            msoShapeRectangle, _
            startX, _
            startY, _
            size, _
            size)
        
        shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
        shp.Line.Visible = msoFalse
        
        shp.TextFrame.TextRange.text = "P" & i
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
        shp.TextFrame.VerticalAnchor = msoAnchorMiddle
        
    Next i

End Sub

