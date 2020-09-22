Attribute VB_Name = "Menu"
Public Sub Intro1()
Dim t As Single

For t = 0 To 255
BackCol = CInc(t)

Render1
DoEvents
Next t

LoadMap "logoyr"
MenuOn = True
'
DoLogo1

End Sub

Public Sub DoLogo1()
Dim t As Single

PlayDSound "over", 256, 256

Ang = 360 / 256
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

DoLogo2

End Sub
Public Sub DoLogo2()
Dim t As Single
LoadMap "logo"
PlayDSound "over", 256, 256

For t = 0 To 256 Step MenuSpeed
DoMenu t, 360 - t * Ang
Render1
DoEvents
Next t
PlayDSound "buda", 256, 256

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

DoLogo3
End Sub
Public Sub DoLogo3()
Dim t As Single
LoadMap "presents"
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

DoLogo4
End Sub
Public Sub DoLogo4()
Dim t As Single
LoadMap "title"
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

DoMainMenu
End Sub
Public Sub DoMainMenu()
Dim t As Single
LoadMap "mainmenu"
PlayDSound "over", 256, 256

For t = 0 To 256 Step MenuSpeed
DoMenu t, 360 - t * Ang
Render1
DoEvents
Next t
PlayDSound "click1", 256, 256
PointerOn = True
'set up cursors
With CurRects(0)
    .Top = OffY + 5 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(1)
    .Top = OffY + 6 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(2)
    .Top = OffY + 7 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(3)
    .Top = OffY + 8 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With

CurrentCur = -1
Cursors = 3
 CurType(0, 1) = 0: CurType(0, 2) = 1
 CurType(1, 1) = 0: CurType(1, 2) = 1
 CurType(2, 1) = 0: CurType(2, 2) = 1
 CurType(3, 1) = 0: CurType(3, 2) = 2

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

If CurrentCur = 0 Then MakePaddles
If CurrentCur = 1 Then DoOptMenu
If CurrentCur = 2 Then DoAboutMenu
If CurrentCur = 3 Then DoQuitMenu

End Sub
Public Sub DoAboutMenu()
Dim t As Single
LoadMap "aboutmenu"
PlayDSound "over", 256, 256

For t = 0 To 256 Step MenuSpeed
DoMenu t, 360 - t * Ang
Render1
DoEvents
Next t
PlayDSound "buda", 256, 256
PointerOn = True
'set up cursors
With CurRects(0)
    .Top = OffY + 17 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With

CurrentCur = -1
Cursors = 0
 CurType(0, 1) = 0: CurType(0, 2) = 1

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

If CurrentCur = 0 Then DoMainMenu

End Sub
Public Sub DoQuitMenu()
Dim t As Single
LoadMap "quitmenu"
PlayDSound "over", 256, 256

For t = 0 To 256 Step MenuSpeed
DoMenu t, 360 - t * Ang
Render1
DoEvents
Next t
PlayDSound "click1", 256, 256
PointerOn = True
'set up cursors
With CurRects(0)
    .Top = OffY + 6 * 16
    .bottom = .Top + 16
    .Left = OffX + 8 * 16
    .Right = .Left + 16
End With
With CurRects(1)
    .Top = OffY + 7 * 16
    .bottom = .Top + 16
    .Left = OffX + 8 * 16
    .Right = .Left + 16
End With

CurrentCur = -1
Cursors = 1
 CurType(0, 1) = 0: CurType(0, 2) = 2
 CurType(1, 1) = 0: CurType(1, 2) = 1

MLoop:
Clicked = False
Do While Clicked = False
DoEvents
Render1
Loop
If CurrentCur = -1 Then GoTo MLoop
Cursors = -1

PlayDSound "click2", 256, 256
If CurrentCur = 0 Then DoExit
PointerOn = False
Highlight = False

For t = 256 To 0 Step -MenuSpeed
DoMenu t, 360 + t * Ang
Render1
DoEvents
Next t

If CurrentCur = 1 Then DoMainMenu



End Sub
Public Sub DoExit()
Dim t As Single
Dim MusFade As Single

MusFade = MusicVol
MusFade = MusFade / 255

PointerOn = False
Highlight = False

For t = 255 To 0 Step -1
SetPlayBackVolume t * MusFade
MenuCol = RGB(t, t, t)
BackCol = CInc(t)
DoBack
DoMenu 256, 0
Render1
DoEvents
Next t
Main.Terminate

End Sub

Public Sub DoOptMenu()
Dim t As Single
LoadMap "Optmenu"
PlayDSound "over", 256, 256

For t = 0 To 256 Step MenuSpeed
DoMenu t, 360 - t * Ang
Render1
DoEvents
Next t
PlayDSound "click1", 256, 256
PointerOn = True
'set up cursors
With CurRects(0)
    .Top = OffY + 5 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(1)
    .Top = OffY + 6 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(2)
    .Top = OffY + 7 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(3)
    .Top = OffY + 8 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With

With CurRects(4)
    .Top = OffY + 10 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With

CurrentCur = -1
Cursors = 4
 CurType(0, 1) = 0: CurType(0, 2) = 1
 CurType(1, 1) = 0: CurType(1, 2) = 1
 CurType(2, 1) = 0: CurType(2, 2) = 1
 CurType(3, 1) = 0: CurType(3, 2) = 1
 CurType(4, 1) = 0: CurType(4, 2) = 1

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

If CurrentCur = 0 Then DoSpeedMenu
If CurrentCur = 1 Then DoSizeMenu
If CurrentCur = 2 Then DoSndMenu
If CurrentCur = 3 Then DoConFigMenu
If CurrentCur = 4 Then DoMainMenu


End Sub

Public Sub DoSndMenu()
Dim t As Single
LoadMap "sndmenu"
PlayDSound "over", 256, 256

DoValueBar 2, 6, 10, 10 - Int(Abs(SoundVol / 500))
DoValueBar 2, 9, 10, Int((MusicVol) / 10)

For t = 0 To 256 Step MenuSpeed
DoMenu t, 360 - t * Ang
Render1
DoEvents
Next t
PlayDSound "click1", 256, 256
PointerOn = True
'set up cursors
With CurRects(0)
    .Top = OffY + 6 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(1)
    .Top = OffY + 6 * 16
    .bottom = .Top + 16
    .Left = OffX + 13 * 16
    .Right = .Left + 16
End With
With CurRects(2)
    .Top = OffY + 9 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(3)
    .Top = OffY + 9 * 16
    .bottom = .Top + 16
    .Left = OffX + 13 * 16
    .Right = .Left + 16
End With
With CurRects(4)
    .Top = OffY + 13 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(5)
    .Top = OffY + 14 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With

CurrentCur = -1
Cursors = 5
 CurType(0, 1) = 3: CurType(0, 2) = 4
 CurType(1, 1) = 5: CurType(1, 2) = 6
 CurType(2, 1) = 3: CurType(2, 2) = 4
 CurType(3, 1) = 5: CurType(3, 2) = 6
 CurType(4, 1) = 0: CurType(4, 2) = 1
 CurType(5, 1) = 0: CurType(5, 2) = 2

MLoop:
Clicked = False
Do While Clicked = False
DoEvents
Render1
Loop
If CurrentCur = -1 Then GoTo MLoop
PlayDSound "click2", 256, 256

If CurrentCur = 0 And SoundVol <> -5000 Then
SoundVol = SoundVol - 500
MaxVolume = SoundVol
DoValueBar 2, 6, 10, 10 - Int(Abs(SoundVol / 500))
End If
If CurrentCur = 1 And SoundVol <> 0 Then
SoundVol = SoundVol + 500
MaxVolume = SoundVol
DoValueBar 2, 6, 10, 10 - Int(Abs(SoundVol / 500))
End If

If CurrentCur = 2 And MusicVol <> 0 Then
MusicVol = MusicVol - 10
SetPlayBackVolume (MusicVol)
DoValueBar 2, 9, 10, Int((MusicVol) / 10)
End If
If CurrentCur = 3 And MusicVol <> 100 Then
MusicVol = MusicVol + 10
SetPlayBackVolume (MusicVol)
DoValueBar 2, 9, 10, Int((MusicVol) / 10)
End If

If CurrentCur < 4 Then GoTo MLoop

Cursors = -1
PointerOn = False
Highlight = False

For t = 256 To 0 Step -MenuSpeed
DoMenu t, 360 + t * Ang
Render1
DoEvents
Next t
If CurrentCur = 4 Then SaveUser: DoOptMenu
If CurrentCur = 5 Then LoadUser: DoOptMenu


End Sub

Public Sub DoSizeMenu()
Dim t As Single
LoadMap "sizemenu"
PlayDSound "over", 256, 256

DoValueBar 2, 6, 10, (BSize - 8) / 4
DoValueBar 2, 9, 10, (PSize - 32) / 8
DoValueBar 2, 12, 10, (OSize - 32) / 8

For t = 0 To 256 Step MenuSpeed
DoMenu t, 360 - t * Ang
Render1
DoEvents
Next t
PlayDSound "click1", 256, 256
PointerOn = True
'set up cursors
With CurRects(0)
    .Top = OffY + 6 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(1)
    .Top = OffY + 6 * 16
    .bottom = .Top + 16
    .Left = OffX + 13 * 16
    .Right = .Left + 16
End With
With CurRects(2)
    .Top = OffY + 9 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(3)
    .Top = OffY + 9 * 16
    .bottom = .Top + 16
    .Left = OffX + 13 * 16
    .Right = .Left + 16
End With
With CurRects(4)
    .Top = OffY + 12 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(5)
    .Top = OffY + 12 * 16
    .bottom = .Top + 16
    .Left = OffX + 13 * 16
    .Right = .Left + 16
End With
With CurRects(6)
    .Top = OffY + 16 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(7)
    .Top = OffY + 17 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With

CurrentCur = -1
Cursors = 7
 CurType(0, 1) = 3: CurType(0, 2) = 4
 CurType(1, 1) = 5: CurType(1, 2) = 6
 CurType(2, 1) = 3: CurType(2, 2) = 4
 CurType(3, 1) = 5: CurType(3, 2) = 6
 CurType(4, 1) = 3: CurType(4, 2) = 4
 CurType(5, 1) = 5: CurType(5, 2) = 6
 CurType(6, 1) = 0: CurType(6, 2) = 1
 CurType(7, 1) = 0: CurType(7, 2) = 2

MLoop:
Clicked = False
Do While Clicked = False
DoEvents
Render1
Loop
If CurrentCur = -1 Then GoTo MLoop
PlayDSound "click2", 256, 256

If CurrentCur = 0 And BSize <> 8 Then
BSize = BSize - 4
DoValueBar 2, 6, 10, (BSize - 8) / 4
End If
If CurrentCur = 1 And BSize <> 48 Then
BSize = BSize + 4
DoValueBar 2, 6, 10, (BSize - 8) / 4
End If


If CurrentCur = 2 And PSize <> 32 Then
PSize = PSize - 8
DoValueBar 2, 9, 10, (PSize - 32) / 8
End If
If CurrentCur = 3 And PSize <> 112 Then
PSize = PSize + 8
DoValueBar 2, 9, 10, (PSize - 32) / 8
End If

If CurrentCur = 4 And OSize <> 32 Then
OSize = OSize - 8
DoValueBar 2, 12, 10, (OSize - 32) / 8
End If
If CurrentCur = 5 And OSize <> 112 Then
OSize = OSize + 8
DoValueBar 2, 12, 10, (OSize - 32) / 8
End If


If CurrentCur < 6 Then GoTo MLoop

Cursors = -1
PointerOn = False
Highlight = False

For t = 256 To 0 Step -MenuSpeed
DoMenu t, 360 + t * Ang
Render1
DoEvents
Next t
If CurrentCur = 6 Then SaveUser: DoOptMenu
If CurrentCur = 7 Then LoadUser: DoOptMenu


End Sub



Public Sub DoConFigMenu()
Dim t As Single, MS As Single
LoadMap "configmenu"
PlayDSound "over", 256, 256

If MenuSpeed = 1 Then MS = 0
If MenuSpeed = 2 Then MS = 1
If MenuSpeed = 4 Then MS = 2
If MenuSpeed = 8 Then MS = 3
If MenuSpeed = 16 Then MS = 4

DoValueBar 2, 6, 10, BackTile - 1
DoValueBar 2, 11, 10, BackRed / 25
DoValueBar 2, 13, 10, BackGreen / 25
DoValueBar 2, 15, 10, BackBlue / 25
DoValueBar 2, 21, 4, MS

For t = 0 To 256 Step MenuSpeed
DoMenu t, 360 - t * Ang
Render1
DoEvents
Next t
PlayDSound "click1", 256, 256
PointerOn = True
'set up cursors
With CurRects(0)
    .Top = OffY + 6 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(1)
    .Top = OffY + 6 * 16
    .bottom = .Top + 16
    .Left = OffX + 13 * 16
    .Right = .Left + 16
End With
With CurRects(2)
    .Top = OffY + 11 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(3)
    .Top = OffY + 11 * 16
    .bottom = .Top + 16
    .Left = OffX + 13 * 16
    .Right = .Left + 16
End With
With CurRects(4)
    .Top = OffY + 13 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(5)
    .Top = OffY + 13 * 16
    .bottom = .Top + 16
    .Left = OffX + 13 * 16
    .Right = .Left + 16
End With
With CurRects(6)
    .Top = OffY + 15 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(7)
    .Top = OffY + 15 * 16
    .bottom = .Top + 16
    .Left = OffX + 13 * 16
    .Right = .Left + 16
End With
With CurRects(8)
    .Top = OffY + 21 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(9)
    .Top = OffY + 21 * 16
    .bottom = .Top + 16
    .Left = OffX + 7 * 16
    .Right = .Left + 16
End With
With CurRects(10)
    .Top = OffY + 25 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(11)
    .Top = OffY + 26 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With

CurrentCur = -1
Cursors = 11
 CurType(0, 1) = 3: CurType(0, 2) = 4
 CurType(1, 1) = 5: CurType(1, 2) = 6
 CurType(2, 1) = 3: CurType(2, 2) = 4
 CurType(3, 1) = 5: CurType(3, 2) = 6
 CurType(4, 1) = 3: CurType(4, 2) = 4
 CurType(5, 1) = 5: CurType(5, 2) = 6
 CurType(6, 1) = 3: CurType(6, 2) = 4
 CurType(7, 1) = 5: CurType(7, 2) = 6
 CurType(8, 1) = 3: CurType(8, 2) = 4
 CurType(9, 1) = 5: CurType(9, 2) = 6
 CurType(10, 1) = 0: CurType(10, 2) = 1
 CurType(11, 1) = 0: CurType(11, 2) = 2

MLoop:
Clicked = False
Do While Clicked = False
DoEvents
BackCol = CLong(BackRed, BackGreen, BackBlue)
Render1
Loop
If CurrentCur = -1 Then GoTo MLoop
PlayDSound "click2", 256, 256

If CurrentCur = 0 And BackTile <> 1 Then
BackTile = BackTile - 1
SmokeTile = BackTile / 2
DoValueBar 2, 6, 10, BackTile - 1
End If
If CurrentCur = 1 And Depth <> 11 Then
BackTile = BackTile + 1
SmokeTile = BackTile / 2
DoValueBar 2, 6, 10, BackTile - 1
End If

If CurrentCur = 2 And BackRed <> 0 Then
BackRed = BackRed - 25
DoValueBar 2, 11, 10, BackRed / 25
End If
If CurrentCur = 3 And BackRed <> 250 Then
BackRed = BackRed + 25
DoValueBar 2, 11, 10, BackRed / 25
End If

If CurrentCur = 4 And BackGreen <> 0 Then
BackGreen = BackGreen - 25
DoValueBar 2, 13, 10, BackGreen / 25
End If
If CurrentCur = 5 And BackGreen <> 250 Then
BackGreen = BackGreen + 25
DoValueBar 2, 13, 10, BackGreen / 25
End If

If CurrentCur = 6 And BackBlue <> 0 Then
BackBlue = BackBlue - 25
DoValueBar 2, 15, 10, BackBlue / 25
End If
If CurrentCur = 7 And BackBlue <> 250 Then
BackBlue = BackBlue + 25
DoValueBar 2, 15, 10, BackBlue / 25
End If

If CurrentCur = 8 And MS <> 0 Then
MS = MS - 1
DoValueBar 2, 21, 4, MS
End If
If CurrentCur = 9 And MS <> 4 Then
MS = MS + 1
DoValueBar 2, 21, 4, MS
End If

If CurrentCur < 10 Then GoTo MLoop

Cursors = -1
PointerOn = False
Highlight = False

If MS = 0 Then MenuSpeed = 1
If MS = 1 Then MenuSpeed = 2
If MS = 2 Then MenuSpeed = 4
If MS = 3 Then MenuSpeed = 8
If MS = 4 Then MenuSpeed = 16

For t = 256 To 0 Step -MenuSpeed
DoMenu t, 360 + t * Ang
Render1
DoEvents
Next t


If CurrentCur = 10 Then SaveUser: DoOptMenu
If CurrentCur = 11 Then LoadUser: DoOptMenu


End Sub


Public Sub DoSpeedMenu()
Dim t As Single
LoadMap "speedmenu"
PlayDSound "over", 256, 256

DoValueBar 2, 6, 10, (BInc - 1) / 0.5
DoValueBar 2, 9, 10, (OSpeed - 1) / 0.5
DoValueBar 2, 12, 10, ORate / 10

For t = 0 To 256 Step MenuSpeed
DoMenu t, 360 - t * Ang
Render1
DoEvents
Next t
PlayDSound "click1", 256, 256
PointerOn = True
'set up cursors
With CurRects(0)
    .Top = OffY + 6 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(1)
    .Top = OffY + 6 * 16
    .bottom = .Top + 16
    .Left = OffX + 13 * 16
    .Right = .Left + 16
End With
With CurRects(2)
    .Top = OffY + 9 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(3)
    .Top = OffY + 9 * 16
    .bottom = .Top + 16
    .Left = OffX + 13 * 16
    .Right = .Left + 16
End With
With CurRects(4)
    .Top = OffY + 12 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(5)
    .Top = OffY + 12 * 16
    .bottom = .Top + 16
    .Left = OffX + 13 * 16
    .Right = .Left + 16
End With
With CurRects(6)
    .Top = OffY + 16 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With
With CurRects(7)
    .Top = OffY + 17 * 16
    .bottom = .Top + 16
    .Left = OffX + 16
    .Right = .Left + 16
End With

CurrentCur = -1
Cursors = 7
 CurType(0, 1) = 3: CurType(0, 2) = 4
 CurType(1, 1) = 5: CurType(1, 2) = 6
 CurType(2, 1) = 3: CurType(2, 2) = 4
 CurType(3, 1) = 5: CurType(3, 2) = 6
 CurType(4, 1) = 3: CurType(4, 2) = 4
 CurType(5, 1) = 5: CurType(5, 2) = 6
 CurType(6, 1) = 0: CurType(6, 2) = 1
 CurType(7, 1) = 0: CurType(7, 2) = 2

MLoop:
Clicked = False
Do While Clicked = False
DoEvents
Render1
Loop
If CurrentCur = -1 Then GoTo MLoop
PlayDSound "click2", 256, 256

If CurrentCur = 0 And BInc <> 1 Then
BInc = BInc - 0.5
DoValueBar 2, 6, 10, (BInc - 1) / 0.5
End If
If CurrentCur = 1 And BInc <> 6 Then
BInc = BInc + 0.5
DoValueBar 2, 6, 10, (BInc - 1) / 0.5
End If

If CurrentCur = 2 And OSpeed <> 1 Then
OSpeed = OSpeed - 0.5
DoValueBar 2, 9, 10, (OSpeed - 1) / 0.5
End If
If CurrentCur = 3 And OSpeed <> 6 Then
OSpeed = OSpeed + 0.5
DoValueBar 2, 9, 10, (OSpeed - 1) / 0.5
End If

If CurrentCur = 4 And ORate <> 0 Then
ORate = ORate - 10
DoValueBar 2, 12, 10, ORate / 10
End If
If CurrentCur = 5 And ORate <> 100 Then
ORate = ORate + 10
DoValueBar 2, 12, 10, ORate / 10
End If


If CurrentCur < 6 Then GoTo MLoop

Cursors = -1
PointerOn = False
Highlight = False

For t = 256 To 0 Step -MenuSpeed
DoMenu t, 360 + t * Ang
Render1
DoEvents
Next t
If CurrentCur = 6 Then SaveUser: DoOptMenu
If CurrentCur = 7 Then LoadUser: DoOptMenu


End Sub

