Attribute VB_Name = "Game"
Public Sub MakePaddles()
Dim Dif As Single, t As Single
Dim MusFade As Single

MusFade = MusicVol
MusFade = MusFade / 255

For t = 255 To 0 Step -1
SetPlayBackVolume t * MusFade
Next t

LoadMusic App.Path + "\music\game.mid"
PlayMusic
SetPlayBackVolume (MusicVol)
DSPosition.CurrentPosition = CMPos

PHSize = PSize / 2
OHSize = OSize / 2
GameOn = True
PXPos = 0
OXPos = 0

'make user paddle first

With GetRect
    .Top = 0
    .bottom = 128
    .Left = 0
    .Right = 16
End With
PutPoint.X = 0: PutPoint.Y = 0
D3DDevice.CopyRects PMakeSurf, GetRect, 1, UPaddleSurf, PutPoint

Dif = PSize - 32

With GetRect
    .Top = 0
    .bottom = 128
    .Left = 24
    .Right = 48
End With
PutPoint.X = 16 + Dif: PutPoint.Y = 0
D3DDevice.CopyRects PMakeSurf, GetRect, 1, UPaddleSurf, PutPoint

If Dif > 0 Then
    With GetRect
        .Top = 0
        .bottom = 128
        .Left = 16
        .Right = 24
    End With
    For t = 0 To Dif - 2
    PutPoint.X = 16 + t: PutPoint.Y = 0
    D3DDevice.CopyRects PMakeSurf, GetRect, 1, UPaddleSurf, PutPoint
    Next t
End If

'make Opponent paddle next

With GetRect
    .Top = 0
    .bottom = 128
    .Left = 48
    .Right = 64
End With
PutPoint.X = 0: PutPoint.Y = 0
D3DDevice.CopyRects PMakeSurf, GetRect, 1, OPaddleSurf, PutPoint

Dif = OSize - 32

With GetRect
    .Top = 0
    .bottom = 128
    .Left = 72
    .Right = 96
End With
PutPoint.X = 16 + Dif: PutPoint.Y = 0
D3DDevice.CopyRects PMakeSurf, GetRect, 1, OPaddleSurf, PutPoint

If Dif > 0 Then
    With GetRect
        .Top = 0
        .bottom = 128
        .Left = 64
        .Right = 72
    End With
    For t = 0 To Dif - 2
    PutPoint.X = 16 + t: PutPoint.Y = 0
    D3DDevice.CopyRects PMakeSurf, GetRect, 1, OPaddleSurf, PutPoint
    Next t
End If

BringPaddles

End Sub

Public Sub BringPaddles()
Dim t As Single

PlayDSound "over", 256, 256

For t = 256 To 256 - 32 Step -1
PYPos = t
OYPos = -t
Render1
Next t

PlayDSound "click1", 256, 256

DoReady

End Sub

Public Sub DoReady()
Dim t As Single
PScore = 0
OScore = 0


LoadMap "ready"
PlayDSound "over", 256, 256

For t = 0 To 256 Step MenuSpeed
DoMenu t, 360 - t * Ang
Render1
DoEvents
Next t
PlayDSound "click1", 256, 256

Dim a As Single, Dif As Single
Dif = 0.5: a = 0
Clicked = False
Do While Clicked = False
DoMenu 256, 360 - a * Ang
a = a + Dif
If a > 10 Then Dif = -0.5
If a < -10 Then Dif = 0.5
DoEvents
Render1
Loop
PlayDSound "click2", 256, 256

For t = 256 To 0 Step -MenuSpeed
DoMenu t, 360 + t * Ang
Render1
DoEvents
Next t

DoGo
End Sub
Public Sub DoGo()
Dim t As Single
LoadMap "go"
PlayDSound "over", 256, 256

For t = 0 To 256 Step MenuSpeed
DoMenu t, 360 - t * Ang
Render1
DoEvents
Next t
PlayDSound "click1", 256, 256

Dim a As Single, Dif As Single
Dif = 0.5: a = 0
Clicked = False
Do While Clicked = False
DoMenu 256, 360 - a * Ang
a = a + Dif
If a > 10 Then Dif = -0.5
If a < -10 Then Dif = 0.5
DoEvents
Render1
Loop
PlayDSound "click2", 256, 256

For t = 256 To 0 Step -MenuSpeed
DoMenu t, 360 + t * Ang
Render1
DoEvents
Next t

BringBall
End Sub
Public Sub MainGameLoop()

Do
DoEvents
Render1
BallUpdate
Loop

End Sub

Public Sub BringBall()
Dim t As Single
BXpos = 0
BYpos = 0
PlayDSound "ballin", 256, 256

Randomize Timer
BAng = Int(Rnd * 2) + 1
If BAng = 1 Then BAng = 180
If BAng = 2 Then BAng = 0

For t = 0 To BSize Step 0.25
BGSize = t
Render1
DoEvents
Next t

MainGameLoop
End Sub
Public Sub BallUpdate()
Dim rate As Single, aSpeed As Single

'Move
Rotate BAng, 0, BInc
BXpos = BXpos + RotX
BYpos = BYpos + RotY

rate = Int(Rnd * ORate)
rate = rate * 0.01
aSpeed = OSpeed * rate
aSpeed = OSpeed - aSpeed

If BXpos < OXPos Then OXPos = OXPos - aSpeed
If BXpos > OXPos Then OXPos = OXPos + aSpeed
If OXPos < -256 + OHSize Then OXPos = -256 + OHSize
If OXPos > 256 - OHSize Then OXPos = 256 - OHSize

CheckUPaddle
CheckOPaddle
CheckWalls
CheckDown
Checkup
End Sub

Public Sub CheckUPaddle()
Dim cx1 As Single, cx2 As Single, cy1 As Single, cy2 As Single
cy1 = 224 - BSize: cy2 = 232 - BSize
cx1 = PXPos - PHSize - BSize
cx2 = PXPos + PHSize + BSize

If BYpos >= cy1 And BYpos <= cy2 And BXpos >= cx1 And BXpos <= cx2 Then
PlayDSound "bounce", BXpos + 256, 256
If BAng > 0 And BAng < 90 Then
BAng = 180 - BAng
BAng = BAng + PChange
End If
If BAng < 0 And BAng > -90 Then
BAng = BAng * -1
BAng = -180 + BAng
BAng = BAng + PChange
End If
If BAng = 0 Then BAng = 180: BAng = BAng + PChange

If BAng > 180 Then
BAng = BAng - 180
BAng = -180 + BAng
End If
If BAng < -179 Then
BAng = BAng * -1
BAng = BAng - 180
BAng = -180 + BAng
End If
If BAng < 130 And BAng > 0 Then BAng = 130
If BAng > -130 And BAng < 0 Then BAng = -130

End If

End Sub

Public Sub CheckOPaddle()
Dim cx1 As Single, cx2 As Single, cy1 As Single, cy2 As Single
cy1 = -224 + BSize: cy2 = -232 + BSize
cx1 = OXPos - OHSize - BSize
cx2 = OXPos + OHSize + BSize

If BYpos <= cy1 And BYpos >= cy2 And BXpos >= cx1 And BXpos <= cx2 Then
PlayDSound "bounce", BXpos + 256, 256
If BAng = 180 Then BAng = 0

If BAng > 0 Then
BAng = 180 - BAng
End If

If BAng < 0 Then
BAng = BAng * -1
BAng = -180 + BAng
End If

End If

End Sub

Public Sub CheckWalls()
Dim cx As Single

cx = BXpos - BSize
If cx < -256 Then
PlayDSound "bounce", BXpos + 256, 256
BAng = BAng * -1
End If

cx = BXpos + BSize
If cx > 256 Then
PlayDSound "bounce", BXpos + 256, 256
BAng = BAng * -1
End If

End Sub
Public Sub CheckDown()
If BYpos >= 256 + BGSize Then
OScore = OScore + 1
TakeOutPaddles
End If

End Sub
Public Sub Checkup()
If BYpos <= -256 - BGSize Then
PScore = PScore + 1
TakeOutPaddles
End If

End Sub

Public Sub TakeOutPaddles()
Dim t As Single

PlayDSound "over", 256, 256

For t = 256 - 32 To 256 Step 1
PYPos = t
OYPos = -t
Render1
Next t

PlayDSound "click1", 256, 256
DoBallOut

End Sub

Public Sub BringPaddlesTwo()
Dim t As Single
Dim m As Long

For m = MusicVol To 0 Step -1
SetPlayBackVolume (m)
Next m

LoadMusic App.Path + "\music\game.mid"
PlayMusic
DSPosition.CurrentPosition = CMPos

For m = 0 To MusicVol
SetPlayBackVolume (m)
Next m

PlayDSound "over", 256, 256
For t = 256 To 256 - 32 Step -1
PYPos = t
OYPos = -t
Render1
Next t

PlayDSound "click1", 256, 256

BringBall

End Sub

Public Sub DoBallOut()
Dim t As Single, ir As Single, ig As Single, ib As Single
Dim m As Long
LoadMap "ballout"
PlayDSound "over", 256, 256

For m = MusicVol To 0 Step -1
SetPlayBackVolume (m)
Next m
CMPos = DSPosition.CurrentPosition
LoadMusic App.Path + "\music\final.mid"
PlayMusic
SetPlayBackVolume (MusicVol)

ir = (250 - BackRed) / 256
ig = (BackGreen / 256)
ib = (BackBlue / 256)
For t = 0 To 256 Step MenuSpeed
BackCol = CLong(BackRed + ir * t, BackGreen - ig * t, BackBlue - ib * t)
DoMenu t, 360 - t * Ang
Render1
DoEvents
Next t
PlayDSound "click1", 256, 256

Dim a As Single, Dif As Single
Dif = 0.5: a = 0
Clicked = False
Do While Clicked = False
DoMenu 256, 360 - a * Ang
a = a + Dif
If a > 10 Then Dif = -0.5
If a < -10 Then Dif = 0.5
DoEvents
Render1
Loop
PlayDSound "click2", 256, 256

For t = 256 To 0 Step -MenuSpeed
BackCol = CLong(BackRed + ir * t, BackGreen - ig * t, BackBlue - ib * t)
DoMenu t, 360 + t * Ang
Render1
DoEvents
Next t

DoScoreMenu
End Sub

Public Sub DoScoreMenu()
Dim m As Long
PXPos = 0
OXPos = 0
Dim t As Single
LoadMap "score"
PlayDSound "over", 256, 256

'DoValueBar 1, 6, 9, PScore
'DoValueBar 1, 9, 9, OScore

With GetRect
    .Top = 112
    .bottom = .Top + 16
    .Left = 96
    .Right = .Left + 16
End With

'do player score
If PScore <> 0 Then
For t = 1 To PScore
PutPoint.X = t * 16 + OffX: PutPoint.Y = 6 * 16 + OffY
D3DDevice.CopyRects DrawSurf(0), GetRect, 1, MenuSurf, PutPoint
Next t
End If

'do opponent score
If OScore <> 0 Then
For t = 1 To OScore
PutPoint.X = t * 16 + OffX: PutPoint.Y = 9 * 16 + OffY
D3DDevice.CopyRects DrawSurf(0), GetRect, 1, MenuSurf, PutPoint
Next t
End If

For t = 0 To 256 Step MenuSpeed
DoMenu t, 360 - t * Ang
Render1
DoEvents
Next t
PlayDSound "click1", 256, 256
PointerOn = True
'set up cursors
With CurRects(0)
    .Top = OffY + 13 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(1)
    .Top = OffY + 14 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With

CurrentCur = -1
Cursors = 1
 CurType(0, 1) = 0: CurType(0, 2) = 1
 CurType(1, 1) = 0: CurType(1, 2) = 2

MLoop:
Clicked = False
Do While Clicked = False
DoEvents
Render1
Loop
If CurrentCur = -1 Then GoTo MLoop

Cursors = -1
PointerOn = False
Highlight = False
PlayDSound "click2", 256, 256

For t = 256 To 0 Step -MenuSpeed
DoMenu t, 360 + t * Ang
Render1
DoEvents
Next t

If OScore = 10 Then DoLost
If PScore = 10 Then DoWin

If CurrentCur = 0 Then BringPaddlesTwo
If CurrentCur = 1 Then
For m = MusicVol To 0 Step -1
SetPlayBackVolume (m)
Next m
LoadMusic App.Path + "\music\intro.mid"
PlayMusic
DoMainMenu
End If


End Sub

Public Sub DoLost()
Dim t As Single, ir As Single, ig As Single, ib As Single
Dim m As Long
LoadMap "lost"
PlayDSound "over", 256, 256

ir = (250 - BackRed) / 256
ig = (BackGreen / 256)
ib = (BackBlue / 256)
For t = 0 To 256 Step MenuSpeed
BackCol = CLong(BackRed + ir * t, BackGreen - ig * t, BackBlue - ib * t)
DoMenu t, 360 - t * Ang
Render1
DoEvents
Next t
PlayDSound "click1", 256, 256

Dim a As Single, Dif As Single
Dif = 0.5: a = 0
Clicked = False
Do While Clicked = False
DoMenu 256, 360 - a * Ang
a = a + Dif
If a > 10 Then Dif = -0.5
If a < -10 Then Dif = 0.5
DoEvents
Render1
Loop
PlayDSound "click2", 256, 256

For t = 256 To 0 Step -MenuSpeed
BackCol = CLong(BackRed + ir * t, BackGreen - ig * t, BackBlue - ib * t)
DoMenu t, 360 + t * Ang
Render1
DoEvents
Next t

For m = MusicVol To 0 Step -1
SetPlayBackVolume (m)
Next m
LoadMusic App.Path + "\music\intro.mid"
PlayMusic
SetPlayBackVolume (MusicVol)
DoMainMenu

End Sub

Public Sub DoWin()
Dim t As Single, ir As Single, ig As Single, ib As Single
Dim m As Long
LoadMap "won"
PlayDSound "over", 256, 256

ir = (BackRed / 256)
ig = (BackGreen / 256)
ib = (250 - BackBlue) / 256
For t = 0 To 256 Step MenuSpeed
BackCol = CLong(BackRed - ir * t, BackGreen - ig * t, BackBlue + ib * t)
DoMenu t, 360 - t * Ang
Render1
DoEvents
Next t
PlayDSound "click1", 256, 256

Dim a As Single, Dif As Single
Dif = 0.5: a = 0
Clicked = False
Do While Clicked = False
DoMenu 256, 360 - a * Ang
a = a + Dif
If a > 10 Then Dif = -0.5
If a < -10 Then Dif = 0.5
DoEvents
Render1
Loop
PlayDSound "click2", 256, 256

For t = 256 To 0 Step -MenuSpeed
BackCol = CLong(BackRed - ir * t, BackGreen - ig * t, BackBlue + ib * t)
DoMenu t, 360 + t * Ang
Render1
DoEvents
Next t

For m = MusicVol To 0 Step -1
SetPlayBackVolume (m)
Next m
LoadMusic App.Path + "\music\intro.mid"
PlayMusic
SetPlayBackVolume (MusicVol)
DoMainMenu

End Sub

