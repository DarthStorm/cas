' macros for text splitting, it is reccomended to put the next four macros together in a module

Sub SplitTextToLetters()
    ' big gpt does stuff
    ' usage: select text box and run macro to split text into individual characters
    ' destructive - deletes the original text box

    Dim sld As slide
    Set sld = ActiveWindow.View.slide
    
    Dim txtShape As shape
    
    ' FIXED SELECTION
    With ActiveWindow.Selection
        If .Type <> ppSelectionShapes Or .ShapeRange.count <> 1 Then
            MsgBox "Select a single text box."
            Exit Sub
        End If
        
        If Not .ShapeRange(1).HasTextFrame Then
            MsgBox "Selected shape must be a text box."
            Exit Sub
        End If
        
        Set txtShape = .ShapeRange(1)
    End With
    
    Dim tr As TextRange
    Set tr = txtShape.TextFrame.TextRange
    
    ' calibration
    Dim testChar As TextRange
    Set testChar = tr.Characters(1, 1)
    
    Dim testShape As shape
    Set testShape = sld.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        testChar.BoundLeft, _
        testChar.BoundTop, _
        testChar.BoundWidth, _
        testChar.BoundHeight)
    
    With testShape.TextFrame
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
    End With
    
    testShape.TextFrame.TextRange.text = testChar.text
    
    Dim offsetX As Double
    Dim offsetY As Double
    
    offsetX = testShape.Left - testChar.BoundLeft
    offsetY = testShape.Top - testChar.BoundTop
    
    testShape.Delete
    
    ' main loop
    Dim i As Long
    Dim charShape As shape
    
    For i = 1 To tr.length
        
        Dim ch As String
        ch = tr.Characters(i, 1).text
        
        If ch <> vbCr And ch <> vbLf Then
            
            Dim cLeft As Double
            Dim cTop As Double
            
            cLeft = tr.Characters(i, 1).BoundLeft
            cTop = tr.Characters(i, 1).BoundTop
            
            Set charShape = sld.Shapes.AddTextbox( _
                msoTextOrientationHorizontal, _
                cLeft, _
                cTop, _
                tr.Characters(i, 1).BoundWidth, _
                tr.Characters(i, 1).BoundHeight)
            
            With charShape.TextFrame
                .MarginLeft = 0
                .MarginRight = 0
                .MarginTop = 0
                .MarginBottom = 0
            End With
            
            charShape.TextFrame.TextRange.text = ch
            
            With charShape.TextFrame.TextRange.Font
                .Name = tr.Characters(i, 1).Font.Name
                .size = tr.Characters(i, 1).Font.size
            End With
            
            charShape.Left = cLeft - offsetX
            charShape.Top = cTop - offsetY
            
        End If
        
    Next i
    
    txtShape.Delete

End Sub

Sub SplitTextToWords()
    ' big gpt made ts
    ' usage: select text box and run macro to split text into individual words (seperated by spaces)
    ' destructive - deletes the original text box

    Dim sld As slide
    Set sld = ActiveWindow.View.slide
    
    Dim txtShape As shape
    
    ' ? FIXED SELECTION
    With ActiveWindow.Selection
        If .Type <> ppSelectionShapes Or .ShapeRange.count <> 1 Then
            MsgBox "Select a single text box."
            Exit Sub
        End If
        
        If Not .ShapeRange(1).HasTextFrame Then
            MsgBox "Selected shape must be a text box."
            Exit Sub
        End If
        
        Set txtShape = .ShapeRange(1)
    End With
    
    Dim tr As TextRange
    Set tr = txtShape.TextFrame.TextRange
    
    Dim i As Long
    Dim wShape As shape
    
    For i = 1 To tr.Words.count
        
        Dim rawText As String
        rawText = tr.Words(i).text
        
        If Trim(rawText) <> "" Then
            
            Dim wLeft As Double
            Dim wTop As Double
            
            wLeft = tr.Words(i).BoundLeft
            wTop = tr.Words(i).BoundTop
            
            Set wShape = sld.Shapes.AddTextbox( _
                msoTextOrientationHorizontal, _
                wLeft, _
                wTop, _
                tr.Words(i).BoundWidth + 2, _
                tr.Words(i).BoundHeight)
            
            With wShape.TextFrame
                .MarginLeft = 0
                .MarginRight = 0
                .MarginTop = 0
                .MarginBottom = 0
                .WordWrap = msoFalse
                .AutoSize = ppAutoSizeShapeToFitText
            End With
            
            wShape.TextFrame.TextRange.text = Trim(rawText)
            
            With wShape.TextFrame.TextRange.Font
                .Name = tr.Words(i).Characters(1).Font.Name
                .size = tr.Words(i).Characters(1).Font.size
            End With
            
        End If
        
    Next i
    
    txtShape.Delete

End Sub

Sub TextBomb()    
    ' i hope you notice a trend
    ' usage: select text box and another shape as origin
    ' run macro to split text into individual letters, then launch them into the stratosphere
    ' (out of the slide)
    ' deletes the original text box, but not the origin shape so you can use later

    Dim sld As slide
    Set sld = ActiveWindow.View.slide
    
    Dim slideW As Double: slideW = sld.Master.Width
    Dim slideH As Double: slideH = sld.Master.Height
    
    Dim txtShape As shape
    Dim originShape As shape
    
    ' SELECT 2 SHAPES
    With ActiveWindow.Selection
        If .Type <> ppSelectionShapes Or .ShapeRange.count <> 2 Then
            MsgBox "Select exactly 2 shapes: a text box and an origin shape."
            Exit Sub
        End If
        
        If .ShapeRange(1).HasTextFrame Then
            Set txtShape = .ShapeRange(1)
            Set originShape = .ShapeRange(2)
        ElseIf .ShapeRange(2).HasTextFrame Then
            Set txtShape = .ShapeRange(2)
            Set originShape = .ShapeRange(1)
        Else
            MsgBox "One shape must be a text box."
            Exit Sub
        End If
    End With
    
    ' ORIGIN = CENTER OF SHAPE
    Dim originX As Double
    Dim originY As Double
    
    originX = originShape.Left + originShape.Width / 2
    originY = originShape.Top + originShape.Height / 2
    
    Dim tr As TextRange
    Set tr = txtShape.TextFrame.TextRange
    
    ' CALIBRATION
    Dim testChar As TextRange
    Set testChar = tr.Characters(1, 1)
    
    Dim testShape As shape
    Set testShape = sld.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        testChar.BoundLeft, _
        testChar.BoundTop, _
        testChar.BoundWidth, _
        testChar.BoundHeight)
    
    With testShape.TextFrame
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
    End With
    
    testShape.TextFrame.TextRange.text = testChar.text
    
    Dim offsetX As Double
    Dim offsetY As Double
    
    offsetX = testShape.Left - testChar.BoundLeft
    offsetY = testShape.Top - testChar.BoundTop
    
    testShape.Delete
    
    ' MAIN LOOP
    Dim maxDist As Double
    maxDist = Sqr(slideW ^ 2 + slideH ^ 2)
    
    Dim i As Long
    Dim charShape As shape
    
    For i = 1 To tr.length
        
        Dim ch As String
        ch = tr.Characters(i, 1).text
        
        If ch <> vbCr And ch <> vbLf Then
            
            Dim cLeft As Double
            Dim cTop As Double
            
            cLeft = tr.Characters(i, 1).BoundLeft
            cTop = tr.Characters(i, 1).BoundTop
            
            Dim cWidth As Double
            Dim cHeight As Double
            
            cWidth = tr.Characters(i, 1).BoundWidth
            cHeight = tr.Characters(i, 1).BoundHeight
            
            Set charShape = sld.Shapes.AddTextbox( _
                msoTextOrientationHorizontal, _
                cLeft, _
                cTop, _
                cWidth, _
                cHeight)
            
            With charShape.TextFrame
                .MarginLeft = 0
                .MarginRight = 0
                .MarginTop = 0
                .MarginBottom = 0
            End With
            
            charShape.TextFrame.TextRange.text = ch
            
            With charShape.TextFrame.TextRange.Font
                .Name = tr.Characters(i, 1).Font.Name
                .size = tr.Characters(i, 1).Font.size
            End With
            
            ' Alignment fix
            charShape.Left = cLeft - offsetX
            charShape.Top = cTop - offsetY
            
            ' EXPLOSION
            
            Dim dx As Double
            Dim dy As Double
            
            dx = charShape.Left - originX
            dy = charShape.Top - originY
            
            Dim length As Double
            length = Sqr(dx * dx + dy * dy)
            
            If length <> 0 Then
                Dim ux As Double
                Dim uy As Double
                
                ux = dx / length
                uy = dy / length
                
                Dim extra As Double
                extra = charShape.TextFrame.TextRange.Font.size * 2
                
                charShape.Left = charShape.Left + ux * (maxDist + extra)
                charShape.Top = charShape.Top + uy * (maxDist + extra)
            End If
            
        End If
        
    Next i
    
    txtShape.Delete

End Sub


Sub WordBomb()    
    ' at this point ill probably stop crediting chatgpt
    ' usage: select text box and another shape as origin
    ' run macro to split text into individual words (seperated by spaces), then launch them into outer space
    ' (out of the slide)
    ' deletes the original text box, but not the origin shape so you can use later


    Dim sld As slide
    Set sld = ActiveWindow.View.slide
    
    Dim slideW As Double: slideW = sld.Master.Width
    Dim slideH As Double: slideH = sld.Master.Height
    
    Dim txtShape As shape
    Dim originShape As shape
    
    With ActiveWindow.Selection
        If .Type <> ppSelectionShapes Or .ShapeRange.count <> 2 Then
            MsgBox "Select exactly 2 shapes: a text box and an origin shape."
            Exit Sub
        End If
        
        If .ShapeRange(1).HasTextFrame Then
            Set txtShape = .ShapeRange(1)
            Set originShape = .ShapeRange(2)
        ElseIf .ShapeRange(2).HasTextFrame Then
            Set txtShape = .ShapeRange(2)
            Set originShape = .ShapeRange(1)
        Else
            MsgBox "One shape must be a text box."
            Exit Sub
        End If
    End With
    
    Dim originX As Double
    Dim originY As Double
    
    originX = originShape.Left + originShape.Width / 2
    originY = originShape.Top + originShape.Height / 2
    
    Dim tr As TextRange
    Set tr = txtShape.TextFrame.TextRange
    
    Dim maxDist As Double
    maxDist = Sqr(slideW ^ 2 + slideH ^ 2)
    
    Dim i As Long
    Dim wShape As shape
    
    For i = 1 To tr.Words.count
        
        Dim wordText As String
        wordText = Trim(tr.Words(i).text)
        
        If wordText <> "" Then
            
            Dim wLeft As Double
            Dim wTop As Double
            
            wLeft = tr.Words(i).BoundLeft
            wTop = tr.Words(i).BoundTop
            
            Dim wWidth As Double
            Dim wHeight As Double
            
            wWidth = tr.Words(i).BoundWidth + 2
            wHeight = tr.Words(i).BoundHeight
            
            Set wShape = sld.Shapes.AddTextbox( _
                msoTextOrientationHorizontal, _
                wLeft, _
                wTop, _
                wWidth, _
                wHeight)
            
            With wShape.TextFrame
                .MarginLeft = 0
                .MarginRight = 0
                .MarginTop = 0
                .MarginBottom = 0
                .WordWrap = msoFalse
                .AutoSize = ppAutoSizeShapeToFitText
            End With
            
            wShape.TextFrame.TextRange.text = wordText
            
            With wShape.TextFrame.TextRange.Font
                .Name = tr.Words(i).Characters(1).Font.Name
                .size = tr.Words(i).Characters(1).Font.size
            End With
            
            ' EXPLOSION
            
            Dim dx As Double
            Dim dy As Double
            
            dx = wLeft - originX
            dy = wTop - originY
            
            Dim length As Double
            length = Sqr(dx * dx + dy * dy)
            
            If length <> 0 Then
                
                Dim ux As Double
                Dim uy As Double
                
                ux = dx / length
                uy = dy / length
                
                Dim extra As Double
                extra = wShape.TextFrame.TextRange.Font.size * 3
                
                wShape.Left = wLeft + ux * (maxDist + extra)
                wShape.Top = wTop + uy * (maxDist + extra)
                
            End If
            
        End If
        
    Next i
    
    txtShape.Delete
End Sub
