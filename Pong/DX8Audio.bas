Attribute VB_Name = "DX8Audio"
'volume is:
    'max = -10,000
    'min = -5,000
'panning is:
    'left = -10,000
    'middle = 0
    'right = 10,000
'Freqs is:
    'min = 100
    'max = 100,000

Public DS As DirectSound8
Public DSBuffer() As DirectSoundSecondaryBuffer8
Public DSEnum As DirectSoundEnum8

Public CurrentBuffer As Long
Public MaxBuffers As Long
Public MaxVolume As Long
Public FieldSize As Long
Public Decay As Long
Public WorkBuffer As Long
Public WaveDirectory As String


Public Sub InitSound(MaxBuff As Long, MaxVol As Long, Field As Long, Dec As Long)
On Error GoTo BailOut:
Set DSEnum = DX.GetDSEnum
Set DS = DX.DirectSoundCreate(DSEnum.GetGuid(1))
DS.SetCooperativeLevel Main.DPic.hWnd, DSSCL_NORMAL
MaxBuffers = MaxBuff
MaxVolume = MaxVol
FieldSize = Field
Decay = Dec

ReDim DSBuffer(MaxBuffers)
Dim DSBDesc As DSBUFFERDESC
Dim t As Long
For t = 0 To MaxBuffers
    DSBDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    Set DSBuffer(t) = DS.CreateSoundBufferFromFile(App.Path + "\default.wav", DSBDesc)
Next t
Exit Sub
BailOut:
End
End Sub

Public Sub LoadSound(Pos As Long, File As String)
Dim DSBDesc As DSBUFFERDESC
    DSBDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    Set DSBuffer(Pos) = DS.CreateSoundBufferFromFile(WaveDirectory + File + ".wav", DSBDesc)
End Sub

Public Sub PlayIt(Pos As Long, Looping As Boolean)
DSBuffer(Pos).SetCurrentPosition 0
If Looping = True Then
    DSBuffer(Pos).Play DSBPLAY_LOOPING
Else
    DSBuffer(Pos).Play DSBPLAY_DEFAULT
End If
End Sub

Public Sub PauseIt(Pos As Long)
DSBuffer(Pos).Stop
End Sub

Public Sub StopIt(Pos As Long)
DSBuffer(Pos).Stop
DSBuffer(Pos).SetCurrentPosition 0
End Sub

Public Sub TerminateSound()
Dim t As Long
For t = 0 To MaxBuffers
Set DSBuffer(t) = Nothing
Next t
Set DSEnum = Nothing
Set DS = Nothing
Set DX = Nothing
End Sub

Public Function GetBuffer() As Long
    Do While DSBuffer(CurrentBuffer).GetStatus = DSBSTATUS_PLAYING
    CurrentBuffer = CurrentBuffer + 1
    If CurrentBuffer > MaxBuffers Then CurrentBuffer = 0
    Loop
GetBuffer = CurrentBuffer
End Function

Public Sub SetFrequency(Pos As Long, Freqs As Long)
DSBuffer(Pos).SetFrequency Freqs
End Sub

Public Sub SetPan(Pos As Long, Pan As Long)
DSBuffer(Pos).SetPan Pan
End Sub

Public Sub SetVolume(Pos As Long, Volume As Long)
DSBuffer(Pos).SetVolume Volume
End Sub

Public Sub SetWaveDirectory(File As String)
WaveDirectory = File
End Sub

Public Sub PlayB(File As String, Optional Volume As Long, Optional Pan As Long)
WorkBuffer = GetBuffer()

LoadSound WorkBuffer, File

SetVolume WorkBuffer, Volume

SetPan WorkBuffer, Pan

PlayIt WorkBuffer, False

End Sub
'The rest of this code is for Booda's 2D directional sound

Public Sub PlayDSound(Name As String, Xpos As Single, Ypos As Single)
WorkBuffer = GetBuffer
LoadSound WorkBuffer, Name
SetPan WorkBuffer, GetDirectionalPan(Xpos)
SetVolume WorkBuffer, GetVolumeDecay(Xpos, Ypos)
PlayIt WorkBuffer, False
End Sub
Public Function GetDirectionalPan(Xpos As Single) As Long
Dim HalfPoint As Long
Dim IDP As Long
Dim XDef As Long
Dim FinalPan As Long
If Xpos > -1 And Xpos < FieldSize + 1 Then
HalfPoint = FieldSize / 2
IDP = 1000 / HalfPoint
XDef = HalfPoint - Xpos
FinalPan = XDef * IDP
FinalPan = FinalPan * -1
End If
If Xpos < 0 Then FinalPan = -1000
If Xpos > FieldSize Then FinalPan = 1000
GetDirectionalPan = FinalPan
End Function

Public Function GetVolumeDecay(Xpos As Single, Ypos As Single) As Long
Dim Xres As Long, Yres As Long
Dim DecRes As Long
Dim IDV As Long
Dim FinalRes As Long
If Xpos > -1 And Xpos < FieldSize + 1 And Ypos > -1 And Ypos < FieldSize + 1 Then GetVolumeDecay = MaxVolume
'Check X
If Xpos < 0 Then Xres = Xpos * -1
If Xpos > FieldSize Then Xres = Xpos - FieldSize
If Xpos > -1 And Xpos < FieldSize + 1 Then Xres = 0
'checky
If Ypos < 0 Then Yres = Ypos * -1
If Ypos > FieldSize Then Yres = Ypos - FieldSize
If Ypos > -1 And Ypos < FieldSize + 1 Then Yres = 0
'check to see wich is larger
If Xres > Yres Then
DecRes = Xres
Else
DecRes = Yres
End If
If DecRes > Decay Then DecRes = Decay
'Find IncrementalDecayVolume
IDV = 5000 / Decay
DecRes = DecRes * IDV
FinalRes = MaxVolume - DecRes
GetVolumeDecay = FinalRes
End Function


