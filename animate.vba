' easing stuff, it is reccomended to put this in its own module
Sub Animate()
    ' gotta give it to big gpt, i made one at first then i got it to rewrite and make better
    ' usage:
    ' plan out your animation by selecting one base shape
    ' and making keyframes by duplicating the shape and moving/rotating/scaling it
    ' keyframes are not deleted so you can use them for other things
    ' then select your base shape, and while holding shift select your keyframes in order
    ' (if working in selection pane hold ctrl instead)
    ' with shapes selected, run the macro
    ' select easing when it asks you
    Dim slid As Slide
    Dim base As Shape, target As Shape
    Dim i As Integer, j As Integer
    Dim t As Double, easeVal As Double
    
    ' --- CONFIGURATION ---
    Dim duration As Single: duration = 1
    Dim stepCount As Integer: stepCount = 60
    Dim stepLength As Single: stepLength = duration / stepCount
    Dim slideWidth As Integer: slideWidth = 960
    Dim slideHeight As Integer: slideHeight = 540

    ' --- USER SELECTION ---
    ' 0=None, 1=Linear, others = custom easing
    Dim moveType As Integer: moveType = GetEasingChoice("Move")
    Dim scaleType As Integer: scaleType = GetEasingChoice("Scale")
    Dim rotType As Integer: rotType = GetEasingChoice("Rotation")
    
    If moveType = 0 And scaleType = 0 And rotType = 0 Then Exit Sub

    ' --- TRACKING VARIABLES ---
    Dim currL As Single: currL = 0
    Dim currT As Single: currT = 0
    Dim currR As Single: currR = 0
    Dim currSX As Single: currSX = 100
    Dim currSY As Single: currSY = 100
    
    Set slid = ActiveWindow.View.Slide
    
    With ActiveWindow.Selection
        If .Type <> ppSelectionShapes Or .ShapeRange.Count < 2 Then
            MsgBox "Select Base, then Target(s)."
            Exit Sub
        End If
        
        Set base = .ShapeRange(1)
        
        For i = 2 To .ShapeRange.Count
            Set target = .ShapeRange(i)
            
            ' --- TARGET VALUES ---
            Dim tL As Single: tL = target.Left - base.Left - (base.Width / 2) + (target.Width / 2)
            Dim tT As Single: tT = target.Top - base.Top - (base.Height / 2) + (target.Height / 2)
            Dim tr As Single: tr = target.Rotation - base.Rotation
            Dim tSX As Single: tSX = (target.Width / base.Width) * 100
            Dim tSY As Single: tSY = (target.Height / base.Height) * 100
            
            Dim sL As Single: sL = currL
            Dim sT As Single: sT = currT
            Dim sR As Single: sR = currR
            Dim sSX As Single: sSX = currSX
            Dim sSY As Single: sSY = currSY
            
            ' =========================
            ' MOVE
            ' =========================
            If moveType > 0 Then
                If moveType = 1 Then
                    ' Linear (single animation)
                    With slid.TimeLine.MainSequence.AddEffect(base, msoAnimEffectPathDown, , msoAnimTriggerWithPrevious)
                        .Behaviors(1).MotionEffect.Path = "M " & currL / slideWidth & " " & currT / slideHeight & _
                                                          " L " & tL / slideWidth & " " & tT / slideHeight
                        .Timing.Duration = duration
                        .Timing.TriggerDelayTime = (i - 2) * duration
                    End With
                    currL = tL: currT = tT
                Else
                    ' Stepped easing
                    For j = 1 To stepCount
                        t = j / stepCount
                        easeVal = ApplyEasing(moveType, t)
                        
                        Dim nL As Single: nL = sL + (easeVal * (tL - sL))
                        Dim nT As Single: nT = sT + (easeVal * (tT - sT))
                        
                        With slid.TimeLine.MainSequence.AddEffect(base, msoAnimEffectPathDown, , msoAnimTriggerWithPrevious)
                            .Behaviors(1).MotionEffect.Path = "M " & currL / slideWidth & " " & currT / slideHeight & _
                                                              " L " & nL / slideWidth & " " & nT / slideHeight
                            .Timing.Duration = stepLength
                            .Timing.TriggerDelayTime = ((i - 2) * duration) + ((j - 1) * stepLength)
                        End With
                        
                        currL = nL: currT = nT
                    Next j
                End If
            End If
            
            ' =========================
            ' ROTATION
            ' =========================
            If rotType > 0 Then
                If rotType = 1 Then
                    With slid.TimeLine.MainSequence.AddEffect(base, msoAnimEffectSpin, , msoAnimTriggerWithPrevious)
                        .Behaviors(1).RotationEffect.From = currR
                        .Behaviors(1).RotationEffect.To = tr
                        .Timing.Duration = duration
                        .Timing.TriggerDelayTime = (i - 2) * duration
                    End With
                    currR = tr
                Else
                    For j = 1 To stepCount
                        t = j / stepCount
                        easeVal = ApplyEasing(rotType, t)
                        
                        Dim nR As Single: nR = sR + (easeVal * (tr - sR))
                        
                        With slid.TimeLine.MainSequence.AddEffect(base, msoAnimEffectSpin, , msoAnimTriggerWithPrevious)
                            .Behaviors(1).RotationEffect.From = currR
                            .Behaviors(1).RotationEffect.To = nR
                            .Timing.Duration = stepLength
                            .Timing.TriggerDelayTime = ((i - 2) * duration) + ((j - 1) * stepLength)
                        End With
                        
                        currR = nR
                    Next j
                End If
            End If
            
            ' =========================
            ' SCALE
            ' =========================
            If scaleType > 0 Then
                If scaleType = 1 Then
                    With slid.TimeLine.MainSequence.AddEffect(base, msoAnimEffectGrowShrink, , msoAnimTriggerWithPrevious)
                        .Behaviors(1).ScaleEffect.FromX = currSX
                        .Behaviors(1).ScaleEffect.ToX = tSX
                        .Behaviors(1).ScaleEffect.FromY = currSY
                        .Behaviors(1).ScaleEffect.ToY = tSY
                        .Timing.Duration = duration
                        .Timing.TriggerDelayTime = (i - 2) * duration
                    End With
                    currSX = tSX: currSY = tSY
                Else
                    For j = 1 To stepCount
                        t = j / stepCount
                        easeVal = ApplyEasing(scaleType, t)
                        
                        Dim nSX As Single: nSX = sSX + (easeVal * (tSX - sSX))
                        Dim nSY As Single: nSY = sSY + (easeVal * (tSY - sSY))
                        
                        With slid.TimeLine.MainSequence.AddEffect(base, msoAnimEffectGrowShrink, , msoAnimTriggerWithPrevious)
                            .Behaviors(1).ScaleEffect.FromX = currSX
                            .Behaviors(1).ScaleEffect.ToX = nSX
                            .Behaviors(1).ScaleEffect.FromY = currSY
                            .Behaviors(1).ScaleEffect.ToY = nSY
                            .Timing.Duration = stepLength
                            .Timing.TriggerDelayTime = ((i - 2) * duration) + ((j - 1) * stepLength)
                        End With
                        
                        currSX = nSX: currSY = nSY
                    Next j
                End If
            End If
            
        Next i
    End With
End Sub

' helper functions for Animate()
Function ApplyEasing(choice As Integer, t As Double) As Double
    Select Case choice
        Case 1: ApplyEasing = t ' Linear
        Case 2: ApplyEasing = EaseBackOut(t)
        Case 3: ApplyEasing = EaseBackIn(t)
        Case 4: ApplyEasing = JumpOut(t)
        Case Else: ApplyEasing = t
    End Select
End Function

Function GetEasingChoice(propName As String) As Integer
    Dim m As String
    m = "Choose Easing for " & propName & ":" & vbCrLf & _
        "0: None (Skip)" & vbCrLf & "1: Linear" & vbCrLf & _
        "2: EaseBackOut" & vbCrLf & "3: EaseBackIn" & vbCrLf & _
        "4: JumpOut"
    Dim ans As String: ans = InputBox(m, propName, "2")
    If ans = "" Then GetEasingChoice = 0 Else GetEasingChoice = Val(ans)
End Function

' easing functions, add your own easing functions for Animate() here
Function EaseBackOut(t As Double) As Double
    Dim c1 As Double: c1 = 1.70158: Dim c3 As Double: c3 = c1 + 1
    EaseBackOut = 1 + c3 * (t - 1) ^ 3 + c1 * (t - 1) ^ 2
End Function

Function EaseBackIn(t As Double) As Double
    Dim c1 As Double: c1 = 1.70158: Dim c3 As Double: c3 = c1 + 1
    EaseBackIn = c3 * t ^ 3 - c1 * t ^ 2
End Function

Function JumpOut(t As Double) As Double
    JumpOut = 2.92 * (t ^ 2) - (1.92 * t)
End Function