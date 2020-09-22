VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Pong 3D - Booda 2002"
   ClientHeight    =   9330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Main.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   622
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   819
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox DPic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   9000
      Left            =   120
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   120
      Width           =   12000
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CImg_Click(Index As Integer)

End Sub

Private Sub DPic_Click()
Clicked = True
End Sub

Private Sub DPic_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Terminate

End Sub

Private Sub DPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim t As Integer, tx As Single, ty As Single
Dim gx As Single, gy As Single
Dim tt As Single
PointerX = X - 6
PointerY = Y
CurX = PointerX - 56
CurY = PointerY + 16

If Cursors <> -1 Then
LastCur = CurrentCur
CurrentCur = -1
For t = 0 To Cursors
If CheckPoint(CurRects(t), CurX, CurY) = True Then CurrentCur = t
Next t
End If

If CurrentCur = -1 Then Highlight = False

If LastCur <> CurrentCur And CurrentCur <> -1 Then
tt = CurType(CurrentCur, 2)
If tt = 0 Then gx = 48: gy = 112
If tt = 1 Then gx = 64: gy = 112
If tt = 2 Then gx = 80: gy = 112
If tt = 3 Then gx = 48: gy = 48
If tt = 4 Then gx = 32: gy = 48
If tt = 5 Then gx = 96: gy = 48
If tt = 6 Then gx = 112: gy = 48

Dim Px As Single, Py As Single
Px = CurRects(CurrentCur).Left + 64
Py = CurRects(CurrentCur).Top - 16

HighStrip(0) = CreateTLV(Px, Py, 0, 1, &HFFFFFF, 0, CTTC(128, gx), CTTC(128, gy))
HighStrip(1) = CreateTLV(Px + 16, Py, 0, 1, &HFFFFFF, 0, CTTC(128, gx + 16), CTTC(128, gy))
HighStrip(2) = CreateTLV(Px, Py + 16, 0, 1, &HFFFFFF, 0, CTTC(128, gx), CTTC(128, gy + 16))
HighStrip(3) = CreateTLV(Px + 16, Py + 16, 0, 1, &HFFFFFF, 0, CTTC(128, gx + 16), CTTC(128, gy + 16))

Highlight = True

End If


If GameOn = True Then
PXPos = X - 320

PChange = X - OldPX

If PXPos < -256 + PHSize Then PXPos = -256 + PHSize
If PXPos > 256 - PHSize Then PXPos = 256 - PHSize

OldPX = X
End If


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If Key = 27 Then Terminate

End Sub

Private Sub Form_Load()
Main.Width = Screen.Width
Main.Height = Screen.Height
DPic.Top = 0
DPic.Left = 0
DPic.Width = Main.ScaleWidth
DPic.Height = Main.ScaleHeight

XOrigin = 320
YOrigin = 240

LoadUser

Cursors = -1
InitGraphics
InitGeometry
InitSound 50, SoundVol, 512, 200

SetWaveDirectory App.Path + "\wavs\"
LoadMusic App.Path + "\music\intro.mid"
PlayMusic
SetPlayBackVolume (MusicVol)
CMPos = 7
Running = True
GameOn = False

MenuOn = False
PointerOn = False
Highlight = False
MenuCol = &HFFFFFF

Me.Show
Main.DPic.SetFocus
DoEvents
Intro1

End Sub

Public Sub Terminate()
Set D3DDevice = Nothing
Set D3D = Nothing
Set DX = Nothing

TerminateSound
TerminateMusic


Unload Me
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
Terminate
End Sub


