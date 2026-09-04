Sub AddOvalColorFade()

    ' Appends Fill Color animation to each oval (keeps appear/pulse animations).
    ' Run after Pop3CountriesPerTick. Colors run on next click, left → right.

    Dim sld As Slide
    Dim shp As Shape
    Dim shapes() As Shape
    Dim effColors() As Effect
    Dim colorAssign() As Long
    Dim i As Long, j As Long, k As Long
    Dim temp As Shape
    Dim n As Long
    Dim seq As Sequence
    Dim eff As Effect
    Dim waveDelay As Double
    Dim fadeDuration As Double
    Dim delay As Double
    Dim colors(1 To 3) As Long
    Dim fromColor As Long

    Set sld = ActiveWindow.View.Slide
    Set seq = sld.TimeLine.MainSequence

    ' remove old fill-color / custom color effects on ovals (safe to re-run)
    For k = seq.Count To 1 Step -1
        Set eff = seq(k)
        If eff.Shape.Type = msoAutoShape Then
            If eff.Shape.AutoShapeType = msoShapeOval Then
                If eff.EffectType = msoAnimEffectChangeFillColor _
                    Or eff.EffectType = msoAnimEffectCustom Then
                    eff.Delete
                End If
            End If
        End If
    Next k

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
    ReDim effColors(1 To n)
    ReDim colorAssign(1 To n)

    n = 0
    For Each shp In sld.Shapes
        If shp.Type = msoAutoShape Then
            If shp.AutoShapeType = msoShapeOval Then
                n = n + 1
                Set shapes(n) = shp
            End If
        End If
    Next shp

    For i = 1 To n - 1
        For j = i + 1 To n
            If shapes(i).Left > shapes(j).Left Then
                Set temp = shapes(i)
                Set shapes(i) = shapes(j)
                Set shapes(j) = temp
            End If
        Next j
    Next i

    colors(1) = RGB(235, 60, 150)    ' #EB3C96
    colors(2) = RGB(255, 200, 50)    ' #FFC832
    colors(3) = RGB(45, 190, 205)    ' #2DBECD

    waveDelay = 0.12
    fadeDuration = 0.6

    Randomize Timer
    For i = 1 To n
        colorAssign(i) = colors(Int(Rnd * 3) + 1)
    Next i

    ' Pass 1: add Fill Color effect + timing per oval
    For i = 1 To n

        delay = (i - 1) * waveDelay

        With shapes(i).Fill
            .Visible = msoTrue
            .Solid
            If .ForeColor.Type = msoColorTypeRGB Then
                fromColor = .ForeColor.RGB
            Else
                fromColor = RGB(255, 255, 255)
            End If
            .ForeColor.RGB = fromColor
        End With

        If i = 1 Then
            Set effColors(i) = seq.AddEffect(Shape:=shapes(i), _
                                             effectId:=msoAnimEffectChangeFillColor, _
                                             trigger:=msoAnimTriggerOnPageClick)
        Else
            Set effColors(i) = seq.AddEffect(Shape:=shapes(i), _
                                             effectId:=msoAnimEffectChangeFillColor, _
                                             trigger:=msoAnimTriggerWithPrevious)
        End If

        With effColors(i).Timing
            .Duration = fadeDuration
            .TriggerDelayTime = delay
        End With

    Next i

    ' Pass 2: set target color per effect (avoids PowerPoint reusing Color2)
    For i = 1 To n
        effColors(i).EffectParameters.Color2.RGB = colorAssign(i)
    Next i

End Sub
