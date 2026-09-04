Sub Pop3CountriesPerTick()

    ' Ovals only: appear + grow & shrink — staggered left → right.

    Dim sld As Slide
    Dim shp As Shape
    Dim shapes() As Shape
    Dim i As Long, j As Long
    Dim temp As Shape
    Dim n As Long
    Dim seq As Sequence
    Dim effAppear As Effect
    Dim effPulse As Effect
    Dim waveDelay As Double
    Dim appearDuration As Double
    Dim pulseDuration As Double
    Dim scalePct As Single
    Dim delay As Double

    Set sld = ActiveWindow.View.Slide
    Set seq = sld.TimeLine.MainSequence

    Do While seq.Count > 0
        seq(1).Delete
    Loop

    ' collect ovals only
    n = 0
    For Each shp In sld.Shapes
        If shp.Type = msoAutoShape Then
            If shp.AutoShapeType = msoShapeOval Then
                n = n + 1
            End If
        End If
    Next shp

    If n = 0 Then Exit Sub

    ReDim shapes(1 To n)
    n = 0
    For Each shp In sld.Shapes
        If shp.Type = msoAutoShape Then
            If shp.AutoShapeType = msoShapeOval Then
                n = n + 1
                Set shapes(n) = shp
            End If
        End If
    Next shp

    ' sort left → right
    For i = 1 To n - 1
        For j = i + 1 To n
            If shapes(i).Left > shapes(j).Left Then
                Set temp = shapes(i)
                Set shapes(i) = shapes(j)
                Set shapes(j) = temp
            End If
        Next j
    Next i

    waveDelay = 0.12         ' gap between each oval appearing (seconds)
    appearDuration = 0.01    ' appear (near-instant)
    pulseDuration = 0.6      ' grow half; AutoReverse adds shrink
    scalePct = 125           ' peak size (% of original)

    For i = 1 To n

        ' Absolute delays from slide start — avoids chaining to prior ovals
        delay = (i - 1) * waveDelay

        ' Appear (hidden in slideshow until this runs)
        Set effAppear = seq.AddEffect(Shape:=shapes(i), _
                                      effectId:=msoAnimEffectAppear, _
                                      trigger:=msoAnimTriggerWithPrevious)

        With effAppear.Timing
            .Duration = appearDuration
            .TriggerDelayTime = delay
        End With

        ' Grow then shrink — starts right after this oval's appear
        Set effPulse = seq.AddEffect(Shape:=shapes(i), _
                                     effectId:=msoAnimEffectGrowShrink, _
                                     trigger:=msoAnimTriggerWithPrevious)

        With effPulse.Behaviors(1).ScaleEffect
            .ByX = scalePct
            .ByY = scalePct
        End With

        With effPulse.Timing
            .Duration = pulseDuration
            .TriggerDelayTime = delay + appearDuration
            .AutoReverse = msoTrue
            .Accelerate = 0.25
            .Decelerate = 0.25
        End With

    Next i

End Sub
