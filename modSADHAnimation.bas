
Attribute VB_Name = "modSADHAnimation"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public animEnabled As Boolean
Private nextTick As Date

' ======= Public entry points =======
Public Sub SetupSADH()
    ' Creates progress bars and a Play/Pause button on DASHBOARD
    EnsureProgressBars
    AddPlayPauseButton
    MsgBox "SAḌH animation controls added on DASHBOARD." & vbCrLf & _
           "Use ToggleAnimation to Play/Pause.", vbInformation, "SAḌH Animation"
End Sub

Public Sub SplashShow()
    ' Minimal animated splash screen (3s) on DASHBOARD
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("DASHBOARD")
    Dim shpBG As Shape, shpTitle As Shape, shpBar As Shape, shpFill As Shape
    
    DeleteShapeIfExists ws, "SADH_SplashBG"
    DeleteShapeIfExists ws, "SADH_SplashTitle"
    DeleteShapeIfExists ws, "SADH_SplashBar"
    DeleteShapeIfExists ws, "SADH_SplashFill"
    
    Set shpBG = ws.Shapes.AddShape(msoShapeRoundedRectangle, 120, 80, 520, 260)
    shpBG.Name = "SADH_SplashBG"
    With shpBG
        .Fill.ForeColor.RGB = RGB(16, 24, 40)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = ""
    End With
    
    Set shpTitle = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 140, 110, 480, 80)
    shpTitle.Name = "SADH_SplashTitle"
    With shpTitle.TextFrame2
        .TextRange.Text = "SAḌH — FTE Billing Dashboard" & vbCrLf & "Loading..."
        .TextRange.Characters(1, 26).Font.Size = 24
        .TextRange.Characters(1, 26).Font.Bold = msoTrue
        .TextRange.Characters(1, 26).Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.Characters(29).Font.Size = 12
        .TextRange.Characters(29).Font.Fill.ForeColor.RGB = RGB(200, 210, 230)
    End With
    shpTitle.Line.Visible = msoFalse
    
    Set shpBar = ws.Shapes.AddShape(msoShapeRectangle, 160, 220, 440, 14)
    shpBar.Name = "SADH_SplashBar"
    With shpBar
        .Fill.ForeColor.RGB = RGB(50, 70, 110)
        .Line.Visible = msoFalse
    End With
    Set shpFill = ws.Shapes.AddShape(msoShapeRectangle, 160, 220, 1, 14)
    shpFill.Name = "SADH_SplashFill"
    With shpFill
        .Fill.ForeColor.RGB = RGB(99, 142, 198)
        .Line.Visible = msoFalse
    End With
    
    Dim steps As Long: steps = 60 ' ~2s
    Dim i As Long
    For i = 1 To steps
        shpFill.Width = 440 * (i / steps)
        DoEvents
        Sleep 30
    Next i
    
    ' brief glow pulse
    Dim p As Long
    For p = 1 To 12
        shpFill.Glow.Radius = 4 + ((p Mod 6) - 3) * 0.5
        DoEvents
        Sleep 30
    Next p
    
    DeleteShapeIfExists ws, "SADH_SplashBG"
    DeleteShapeIfExists ws, "SADH_SplashTitle"
    DeleteShapeIfExists ws, "SADH_SplashBar"
    DeleteShapeIfExists ws, "SADH_SplashFill"
End Sub

Public Sub ToggleAnimation()
    If animEnabled Then
        StopAnimation
    Else
        StartAnimation
    End If
End Sub

Public Sub StartAnimation()
    animEnabled = True
    EnsureProgressBars
    ScheduleNextTick 0.15
    UpdatePlayPauseButton
End Sub

Public Sub StopAnimation()
    animEnabled = False
    On Error Resume Next
    Application.OnTime EarliestTime:=nextTick, Procedure:="modSADHAnimation.AnimateTick", _
        Schedule:=False
    On Error GoTo 0
    UpdatePlayPauseButton
End Sub

' ======= Animation loop =======
Public Sub AnimateTick()
    On Error GoTo SafeExit
    If Not animEnabled Then Exit Sub
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("DASHBOARD")
    
    AnimateBar ws, "Service Level", "pbSL", 0, 1
    AnimateBar ws, "AHT (sec)", "pbAHT", 0, 1, True ' invert: smaller is better
    AnimateBar ws, "Occupancy", "pbOCC", 0, 1
    AnimateBar ws, "Conformance", "pbCONF", 0, 1
    AnimateBar ws, "Utilization", "pbUTIL", 0, 1
    AnimateBar ws, "FTE Billed (Avg/day)", "pbFTE", 0, 2 ' 2 FTE/day rough scale
    
SafeExit:
    If animEnabled Then ScheduleNextTick 0.15
End Sub

Private Sub ScheduleNextTick(ByVal seconds As Double)
    nextTick = Now + seconds / 86400#
    Application.OnTime EarliestTime:=nextTick, Procedure:="modSADHAnimation.AnimateTick", _
        Schedule:=True
End Sub

' ======= Drawing helpers =======
Private Sub EnsureProgressBars()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("DASHBOARD")
    MakeBar ws, "Service Level", "pbSL"
    MakeBar ws, "AHT (sec)", "pbAHT"
    MakeBar ws, "Occupancy", "pbOCC"
    MakeBar ws, "Conformance", "pbCONF"
    MakeBar ws, "Utilization", "pbUTIL"
    MakeBar ws, "FTE Billed (Avg/day)", "pbFTE"
End Sub

Private Sub MakeBar(ws As Worksheet, ByVal labelText As String, ByVal nameFill As String)
    Dim r As Long: r = FindLabelRow(ws, labelText)
    If r = 0 Then Exit Sub
    
    ' Place bars one row below the KPI value (which is B{r+1})
    Dim topY As Double: topY = ws.Rows(r + 1).Top + 6
    Dim leftX As Double: leftX = ws.Columns("E").Left
    Dim widthW As Double: widthW = ws.Columns("H").Left - leftX - 10
    Dim heightH As Double: heightH = 10
    
    ' Background
    Dim nameBG As String: nameBG = nameFill & "_bg"
    DeleteShapeIfExists ws, nameBG
    With ws.Shapes.AddShape(msoShapeRectangle, leftX, topY, widthW, heightH)
        .Name = nameBG
        .Fill.ForeColor.RGB = RGB(235, 240, 250)
        .Line.Visible = msoFalse
    End With
    
    ' Fill
    DeleteShapeIfExists ws, nameFill
    With ws.Shapes.AddShape(msoShapeRectangle, leftX, topY, 1, heightH)
        .Name = nameFill
        .Fill.ForeColor.RGB = RGB(99, 142, 198)
        .Line.Visible = msoFalse
    End With
End Sub

Private Sub AnimateBar(ws As Worksheet, ByVal labelText As String, ByVal nameFill As String, _
                       ByVal minVal As Double, ByVal maxVal As Double, Optional invert As Boolean = False)
    Dim r As Long: r = FindLabelRow(ws, labelText)
    If r = 0 Then Exit Sub
    
    Dim valCell As Range: Set valCell = ws.Range("B" & (r + 1))
    Dim targetCell As Range: Set targetCell = ws.Range("D" & (r + 2))
    
    Dim v As Double: v = NzD(valCell.Value2)
    Dim t As Double: t = NzD(targetCell.Value2)
    If t = 0 Then t = maxVal
    
    Dim pct As Double
    If invert Then
        If v <= 0 Then
            pct = 1
        Else
            pct = Application.Min(1#, t / v)
        End If
    Else
        pct = Application.Max(0#, Application.Min(1#, v / IIf(t = 0, maxVal, t)))
    End If
    
    Dim nameBG As String: nameBG = nameFill & "_bg"
    Dim w As Double: w = 0
    If ShapeExists(ws, nameBG) Then w = ws.Shapes(nameBG).Width
    
    If ShapeExists(ws, nameFill) Then
        Dim cur As Double: cur = ws.Shapes(nameFill).Width
        Dim targetW As Double: targetW = w * pct
        Dim stepW As Double: stepW = (targetW - cur) * 0.25 ' easing
        ws.Shapes(nameFill).Width = cur + stepW
        ws.Shapes(nameFill).Fill.ForeColor.RGB = BlendRGB(RGB(248, 105, 107), RGB(99, 190, 123), pct)
        ws.Shapes(nameFill).Glow.Radius = 2 + 4 * pct
    End If
End Sub

Private Function FindLabelRow(ws As Worksheet, ByVal labelText As String) As Long
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    Dim r As Long
    For r = 1 To lastRow
        If Nz(ws.Cells(r, "B").Value2) = labelText Then
            FindLabelRow = r
            Exit Function
        End If
    Next r
End Function

Private Function ShapeExists(ws As Worksheet, ByVal nm As String) As Boolean
    On Error Resume Next
    ShapeExists = Not ws.Shapes(nm) Is Nothing
    On Error GoTo 0
End Function

Private Sub DeleteShapeIfExists(ws As Worksheet, ByVal nm As String)
    On Error Resume Next
    ws.Shapes(nm).Delete
    On Error GoTo 0
End Sub

Private Function Nz(ByVal v) As String
    If IsError(v) Then
        Nz = ""
    ElseIf IsEmpty(v) Then
        Nz = ""
    Else
        Nz = CStr(v)
    End If
End Function

Private Function NzD(ByVal v) As Double
    On Error Resume Next
    If IsError(v) Or IsEmpty(v) Or v = "" Then
        NzD = 0#
    Else
        NzD = CDbl(v)
    End If
    On Error GoTo 0
End Function

Private Function BlendRGB(ByVal c1 As Long, ByVal c2 As Long, ByVal t As Double) As Long
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    r1 = (c1 And &HFF): g1 = (c1 \ &H100 And &HFF): b1 = (c1 \ &H10000 And &HFF)
    r2 = (c2 And &HFF): g2 = (c2 \ &H100 And &HFF): b2 = (c2 \ &H10000 And &HFF)
    
    Dim r As Long, g As Long, b As Long
    r = r1 + (r2 - r1) * t
    g = g1 + (g2 - g1) * t
    b = b1 + (b2 - b1) * t
    
    BlendRGB = RGB(r, g, b)
End Function

Private Sub AddPlayPauseButton()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("DASHBOARD")
    DeleteShapeIfExists ws, "btnPlayPause"
    Dim topY As Double: topY = ws.Range("B2").Top
    Dim leftX As Double: leftX = ws.Range("H4").Left
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftX, topY, 110, 28)
    btn.Name = "btnPlayPause"
    btn.OnAction = "modSADHAnimation.ToggleAnimation"
    UpdatePlayPauseButton
End Sub

Private Sub UpdatePlayPauseButton()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("DASHBOARD")
    If Not ShapeExists(ws, "btnPlayPause") Then Exit Sub
    With ws.Shapes("btnPlayPause")
        .TextFrame2.TextRange.Text = IIf(animEnabled, "⏸ Pause", "▶ Play")
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .Fill.ForeColor.RGB = IIf(animEnabled, RGB(230, 240, 250), RGB(99, 142, 198))
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(20, 30, 40)
        .Line.Visible = msoFalse
    End With
End Sub
