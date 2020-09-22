Attribute VB_Name = "DirectX8Graphics"
Public DX As DirectX8 'Main directx object
Public D3D As Direct3D8 'This is the main 3d object
Public D3DDevice As Direct3DDevice8 'The 3d hardware
'This represents the vertex format we will use
Public Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

'This is your vertex structure,identical to directx7
Public Type TLVertex
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    color As Long
    specular As Long
    tu As Single
    tv As Single
End Type
Public Type GScreen
    X As Single
    Y As Single
End Type

'These polygons are going to be similar to surfaces
Public WorkStrip(0 To 3) As TLVertex 'These are temp polygons
Public BorderL(0 To 3) As TLVertex  'Borderpolygons
Public BorderR(0 To 3) As TLVertex  'Borderpolygons

Public BackStrip(0 To 3) As TLVertex
Public BackStrip1(0 To 3) As TLVertex
Public BackStrip2(0 To 3) As TLVertex
Public BackStrip3(0 To 3) As TLVertex
Public SmokeStrip(0 To 3) As TLVertex
Public SmokeStrip1(0 To 3) As TLVertex
Public SmokeStrip2(0 To 3) As TLVertex
Public SmokeStrip3(0 To 3) As TLVertex
Public MenuStrip(0 To 3) As TLVertex
Public PointerStrip(0 To 3) As TLVertex
Public UPaddleStrip(0 To 3) As TLVertex
Public OPaddleStrip(0 To 3) As TLVertex
Public HighStrip(0 To 3) As TLVertex
Public OrbStrip(0 To 3) As TLVertex
Public OrbSStrip(0 To 3) As TLVertex

'these are the textures for the polys
Public D3DX As D3DX8 'Helper library for textures
Public BackTex As Direct3DTexture8
Public SmokeTex As Direct3DTexture8
Public DrawTex(0 To 2) As Direct3DTexture8
Public MenuTex As Direct3DTexture8
Public BorderTex As Direct3DTexture8
Public OrbTex As Direct3DTexture8
Public OrbSTex As Direct3DTexture8
Public UPaddleTex As Direct3DTexture8
Public OPaddleTex As Direct3DTexture8
Public PMakeTex As Direct3DTexture8

'These are the surfaces
'Public LSurf(0 To 2) As Direct3DSurface8
Public DrawSurf(0 To 2) As Direct3DSurface8
Public MenuSurf As Direct3DSurface8
Public UPaddleSurf As Direct3DSurface8
Public OPaddleSurf As Direct3DSurface8
Public PMakeSurf As Direct3DSurface8


'These are the rects and points for surface manipulation
Public GetRect As RECT
Public PutPoint As Point

'#########
'## FONTS ##
'#########
Public MainFont As D3DXFont
Public MainFontDesc As IFont
Public TextRect As RECT
Public fnt As New StdFont
'########################
'## FRAME RATE CALCULATIONS ##
'#######################
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public FPS_LastCheck As Long
Public FPS_Count As Long
Public FPS_Current As Integer

Public Sub InitGraphics()
'On Error GoTo bailout:
Dim DispMode As D3DDISPLAYMODE
Dim D3DWindow As D3DPRESENT_PARAMETERS
Dim ColorKeyVal As Long

Set DX = New DirectX8
Set D3D = DX.Direct3DCreate()
Set D3DX = New D3DX8

DispMode.Format = D3DFMT_R5G6B5
DispMode.Width = 640
DispMode.Height = 480

D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
D3DWindow.BackBufferCount = 1
D3DWindow.BackBufferFormat = DispMode.Format
D3DWindow.BackBufferWidth = 640
D3DWindow.BackBufferHeight = 480
D3DWindow.hDeviceWindow = Main.hWnd
'D3DWindow.Windowed = True

Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Main.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DWindow)

D3DDevice.SetVertexShader FVF

D3DDevice.SetRenderState D3DRS_LIGHTING, False


'ColorKeyValue
ColorKeyVal = &HFF0000FF         'Blue, the only color I use, hence blue screen

'Load textures
Set BackTex = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\pix\back.dds", 128, 128, 0, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, ColorKeyVal, ByVal 0, ByVal 0)
Set SmokeTex = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\pix\smoke.dds", 256, 256, 0, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, ColorKeyVal, ByVal 0, ByVal 0)
Set DrawTex(0) = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\pix\font.dds", 128, 128, 0, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, ColorKeyVal, ByVal 0, ByVal 0)
Set DrawTex(1) = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\pix\title.dds", 128, 128, 0, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, ColorKeyVal, ByVal 0, ByVal 0)
Set MenuTex = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\pix\title.dds", 512, 512, 0, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, ColorKeyVal, ByVal 0, ByVal 0)
Set BorderTex = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\pix\border.dds", 16, 16, 0, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, ColorKeyVal, ByVal 0, ByVal 0)
Set OrbTex = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\pix\orb.dds", 128, 128, 0, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, ColorKeyVal, ByVal 0, ByVal 0)
Set OrbSTex = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\pix\orbs.dds", 128, 128, 0, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, ColorKeyVal, ByVal 0, ByVal 0)
Set UPaddleTex = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\pix\pmake.dds", 128, 128, 0, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, ColorKeyVal, ByVal 0, ByVal 0)
Set OPaddleTex = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\pix\pmake.dds", 128, 128, 0, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, ColorKeyVal, ByVal 0, ByVal 0)
Set PMakeTex = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\pix\pmake.dds", 128, 128, 0, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, ColorKeyVal, ByVal 0, ByVal 0)


'set pointers to surfaces
Set DrawSurf(0) = DrawTex(0).GetSurfaceLevel(0)
Set DrawSurf(1) = DrawTex(1).GetSurfaceLevel(0)
Set MenuSurf = MenuTex.GetSurfaceLevel(0)
Set UPaddleSurf = UPaddleTex.GetSurfaceLevel(0)
Set OPaddleSurf = OPaddleTex.GetSurfaceLevel(0)
Set PMakeSurf = PMakeTex.GetSurfaceLevel(0)


'#########
'## FONTS ##
'#########
fnt.Name = "Verdana"
fnt.Size = 12
Set MainFontDesc = fnt
Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)


Exit Sub
BailOut:
End
End Sub

Public Function CreateTLV(X As Single, Y As Single, Z As Single, rhw As Single, color As Long, specular As Long, tu As Single, tv As Single) As TLVertex
CreateTLV.X = X
CreateTLV.Y = Y
CreateTLV.Z = Z
CreateTLV.rhw = rhw
CreateTLV.color = color
CreateTLV.specular = specular
CreateTLV.tu = tu
CreateTLV.tv = tv

End Function

Public Sub InitGeometry()


'make top border
'Borders(0, 0) = CreateTLV(0, 0, 0, 1, &HFFFFFFFF, 0, 0, CTTC(32, 24))
'Borders(0, 1) = CreateTLV(320, 0, 0, 1, &HFFFFFFFF, 0, 10, CTTC(32, 24))
'Borders(0, 2) = CreateTLV(0, 8, 0, 1, &HFFFFFFFF, 0, 0, 1)
'Borders(0, 3) = CreateTLV(320, 8, 0, 1, &HFFFFFFFF, 0, 10, 1)


'make left border
BorderL(0) = CreateTLV(0, 0, 0, 1, &H0, 0, 0, 0)
BorderL(1) = CreateTLV(64, 0, 0, 1, &H0, 0, 1, 0)
BorderL(2) = CreateTLV(0, 480, 0, 1, &H0, 0, 0, 1)
BorderL(3) = CreateTLV(64, 480, 0, 1, &H0, 0, 1, 1)

'make right border
BorderR(0) = CreateTLV(576, 0, 0, 1, &H0, 0, 0, 0)
BorderR(1) = CreateTLV(640, 0, 0, 1, &H0, 0, 1, 0)
BorderR(2) = CreateTLV(576, 480, 0, 1, &H0, 0, 0, 1)
BorderR(3) = CreateTLV(640, 480, 0, 1, &H0, 0, 1, 1)

End Sub

Public Function CLong(Red As Single, Green As Single, Blue As Single) As Long
Dim r As Long, g As Long, b As Long
r = Red: g = Green: b = Blue
g = g * 256
r = r * 256 * 256
CLong = r + g + b
End Function
Public Function CInc(Val As Single) As Long
Dim RTemp As Single, GTemp As Single, BTemp As Single

RTemp = BackRed / 250
GTemp = BackGreen / 250
BTemp = BackBlue / 250

CInc = CLong(RTemp * Val, GTemp * Val, BTemp * Val)

End Function
Public Function CTTC(Total As Single, Current As Single) As Single
Dim IncTex As Single
IncTex = 1 / Total
CTTC = IncTex * Current
End Function

Public Sub TerminateGraphics()
Set D3DDevice = Nothing
Set D3D = Nothing
Set DX = Nothing

End Sub
Public Function Collide(rect1 As RECT, rect2 As RECT) As Boolean
'always put smaller rect first
Dim x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer, cx As Integer, cy As Integer
x1 = rect2.Left
x2 = rect2.Right
y1 = rect2.Top
y2 = rect2.bottom

'do topleftpoint
cx = rect1.Left
cy = rect1.Top
If cx > x1 And cx < x2 And cy > y1 And cy < y2 Then Collide = True: Exit Function

'do bottomleftpoint
cx = rect1.Left
cy = rect1.bottom
If cx > x1 And cx < x2 And cy > y1 And cy < y2 Then Collide = True: Exit Function

'do bottomrightpoint
cx = rect1.Right
cy = rect1.bottom
If cx > x1 And cx < x2 And cy > y1 And cy < y2 Then Collide = True: Exit Function

End Function

Public Function CheckPoint(TRect As RECT, Px As Single, Py As Single) As Boolean
If Px > TRect.Left And Px < TRect.Right And Py > TRect.Top And Py < TRect.bottom Then
CheckPoint = True
Else: CheckPoint = False
End If

End Function
Public Sub LoadMap(FileName As String)
FileName = App.Path + "\maps\" + FileName + ".bdm"

Dim mX As Single, mY As Single
Dim ts As Single, RX As Single, RY As Single, ct As Integer
Dim GridX As Single, GridY As Single, Dum As Single

'clear screen first
For mX = 0 To 31
For mY = 0 To 31
With GetRect
    .Top = 0
    .bottom = .Top + 16
    .Left = 0
    .Right = .Left + 16
End With
PutPoint.X = mX * 16: PutPoint.Y = mY * 16
D3DDevice.CopyRects DrawSurf(0), GetRect, 1, MenuSurf, PutPoint
Next mY, mX

Open FileName For Input As #1
Input #1, GridX, GridY
OffX = (512 - (GridX * 16)) / 2 - 8
OffY = (512 - (GridY * 16)) / 2 - 8
For mX = 0 To GridX
For mY = 0 To GridY
Input #1, ts, Dum, Dum
ct = Int(ts / 64)
ts = ts - (ct * 64)
RY = Int(ts / 8)
RX = ts - (RY * 8)
RY = RY * 16: RX = RX * 16
With GetRect
    .Top = RY
    .bottom = RY + 16
    .Left = RX
    .Right = RX + 16
End With
PutPoint.X = mX * 16 + OffX: PutPoint.Y = mY * 16 + OffY
D3DDevice.CopyRects DrawSurf(ct), GetRect, 1, MenuSurf, PutPoint
Next mY, mX
Close #1
End Sub

Public Sub DoBack()

BackStrip(0) = CreateTLV(64, -16, 0, 1, BackCol, 0, 0, 0)
BackStrip(1) = CreateTLV(576, -16, 0, 1, BackCol, 0, BackTile, 0)
BackStrip(2) = CreateTLV(64, 496, 0, 1, BackCol, 0, 0, BackTile)
BackStrip(3) = CreateTLV(576, 496, 0, 1, BackCol, 0, BackTile, BackTile)
End Sub

Public Sub DoSmoke()

SmokeOff = SmokeOff + 0.0005
If SmokeOff >= 1 Then SmokeOff = 0

SmokeStrip(0) = CreateTLV(64, -16, 0, 1, SmokeCol, 0, SmokeOff, SmokeOff)
SmokeStrip(1) = CreateTLV(576, -16, 0, 1, SmokeCol, 0, SmokeOff + SmokeTile, SmokeOff)
SmokeStrip(2) = CreateTLV(64, 496, 0, 1, SmokeCol, 0, SmokeOff, SmokeOff + SmokeTile)
SmokeStrip(3) = CreateTLV(576, 496, 0, 1, SmokeCol, 0, SmokeOff + SmokeTile, SmokeOff + SmokeTile)

End Sub
Public Sub DoPaddles()
Dim PG As Single
PGOff = PGOff + 0.1
If PGOff >= 8 Then PGOff = 0

PG = Int(PGOff) * 16
'user paddles
UPaddleStrip(0) = CreateTLV(XOrigin + PXPos - PHSize, YOrigin + PYPos, 0, 1, &HFFFFFF, 0, 0, CTTC(128, PG))
UPaddleStrip(1) = CreateTLV(XOrigin + PXPos + PHSize + 8, YOrigin + PYPos, 0, 1, &HFFFFFF, 0, CTTC(128, PSize + 8), CTTC(128, PG))
UPaddleStrip(2) = CreateTLV(XOrigin + PXPos - PHSize, YOrigin + PYPos + 16, 0, 1, &HFFFFFF, 0, 0, CTTC(128, PG + 16))
UPaddleStrip(3) = CreateTLV(XOrigin + PXPos + PHSize + 8, YOrigin + PYPos + 16, 0, 1, &HFFFFFF, 0, CTTC(128, PSize + 8), CTTC(128, PG + 16))

PG = 112 - Int(PGOff) * 16
'opponent paddles
OPaddleStrip(0) = CreateTLV(XOrigin + OXPos - OHSize, YOrigin + OYPos - 8, 0, 1, &HFFFFFF, 0, 0, CTTC(128, PG))
OPaddleStrip(1) = CreateTLV(XOrigin + OXPos + OHSize + 8, YOrigin + OYPos - 8, 0, 1, &HFFFFFF, 0, CTTC(128, OSize + 8), CTTC(128, PG))
OPaddleStrip(2) = CreateTLV(XOrigin + OXPos - OHSize, YOrigin + OYPos + 8, 0, 1, &HFFFFFF, 0, 0, CTTC(128, PG + 16))
OPaddleStrip(3) = CreateTLV(XOrigin + OXPos + OHSize + 8, YOrigin + OYPos + 8, 0, 1, &HFFFFFF, 0, CTTC(128, OSize + 8), CTTC(128, PG + 16))

End Sub
Public Sub DoMenu(OffSet As Single, ZAng As Single)

Rotate ZAng, -OffSet, -OffSet
MenuStrip(0) = CreateTLV(XOrigin + RotX, YOrigin + RotY, 0, 1, MenuCol, 0, 0, 0)

Rotate ZAng, OffSet, -OffSet
MenuStrip(1) = CreateTLV(XOrigin + RotX, YOrigin + RotY, 0, 1, MenuCol, 0, 1, 0)

Rotate ZAng, -OffSet, OffSet
MenuStrip(2) = CreateTLV(XOrigin + RotX, YOrigin + RotY, 0, 1, MenuCol, 0, 0, 1)

Rotate ZAng, OffSet, OffSet
MenuStrip(3) = CreateTLV(XOrigin + RotX, YOrigin + RotY, 0, 1, MenuCol, 0, 1, 1)

End Sub
Public Sub DoPointer()
PointerStrip(0) = CreateTLV(PointerX, PointerY, 0, 1, &HFFFFFF, 0, 0, CTTC(128, 48))
PointerStrip(1) = CreateTLV(PointerX + 32, PointerY, 0, 1, &HFFFFFF, 0, CTTC(128, 32), CTTC(128, 48))
PointerStrip(2) = CreateTLV(PointerX, PointerY + 48, 0, 1, &HFFFFFF, 0, 0, CTTC(128, 96))
PointerStrip(3) = CreateTLV(PointerX + 32, PointerY + 48, 0, 1, &HFFFFFF, 0, CTTC(128, 32), CTTC(128, 96))
End Sub
Public Sub DoBall()
Dim gx As Single, gy As Single
Dim bx As Single, by As Single

'do graphic position first
BGIncX = BGIncX + 0.25
If BGIncX >= 4 Then BGIncX = 0: BGIncY = BGIncY + 1
If BGIncY >= 4 Then BGIncY = 0
gx = Int(BGIncX) * 32
gy = Int(BGIncY) * 32

'Do Shadow First
bx = XOrigin + BXpos + 4
by = YOrigin + BYpos + 4

OrbSStrip(0) = CreateTLV(bx - BGSize, by - BGSize, 0, 1, &HFFFFFF, 0, CTTC(128, gx), CTTC(128, gy))
OrbSStrip(1) = CreateTLV(bx + BGSize, by - BGSize, 0, 1, &HFFFFFF, 0, CTTC(128, gx + 32), CTTC(128, gy))
OrbSStrip(2) = CreateTLV(bx - BGSize, by + BGSize, 0, 1, &HFFFFFF, 0, CTTC(128, gx), CTTC(128, gy + 32))
OrbSStrip(3) = CreateTLV(bx + BGSize, by + BGSize, 0, 1, &HFFFFFF, 0, CTTC(128, gx + 32), CTTC(128, gy + 32))

'Do ball last
bx = XOrigin + BXpos
by = YOrigin + BYpos

OrbStrip(0) = CreateTLV(bx - BGSize, by - BGSize, 0, 1, &HFFFFFF, 0, CTTC(128, gx), CTTC(128, gy))
OrbStrip(1) = CreateTLV(bx + BGSize, by - BGSize, 0, 1, &HFFFFFF, 0, CTTC(128, gx + 32), CTTC(128, gy))
OrbStrip(2) = CreateTLV(bx - BGSize, by + BGSize, 0, 1, &HFFFFFF, 0, CTTC(128, gx), CTTC(128, gy + 32))
OrbStrip(3) = CreateTLV(bx + BGSize, by + BGSize, 0, 1, &HFFFFFF, 0, CTTC(128, gx + 32), CTTC(128, gy + 32))


End Sub
Public Sub Render1()
If Running = True Then

If DSPosition.CurrentPosition >= DSPosition.Duration Then
DSPosition.CurrentPosition = 0
PlayMusic
End If

DoBack
DoSmoke

D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, &H0, 1#, 0

'Begin scene
D3DDevice.BeginScene

    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
   D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
   D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
'set up and draw layers if enabled

'MakeScreenStrip
D3DDevice.SetTexture 0, BackTex
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, BackStrip(0), Len(BackStrip(0))

D3DDevice.SetTexture 0, SmokeTex
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, SmokeStrip(0), Len(SmokeStrip(0))

If GameOn = True Then
'user paddle
DoBall
D3DDevice.SetTexture 0, OrbSTex
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, OrbSStrip(0), Len(OrbSStrip(0))
D3DDevice.SetTexture 0, OrbTex
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, OrbStrip(0), Len(OrbStrip(0))

DoPaddles
D3DDevice.SetTexture 0, UPaddleTex
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, UPaddleStrip(0), Len(UPaddleStrip(0))

D3DDevice.SetTexture 0, OPaddleTex
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, OPaddleStrip(0), Len(OPaddleStrip(0))

End If

If MenuOn = True Then
D3DDevice.SetTexture 0, MenuTex
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, MenuStrip(0), Len(MenuStrip(0))
End If

If Highlight = True Then
D3DDevice.SetTexture 0, DrawTex(0)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, HighStrip(0), Len(HighStrip(0))
End If

If PointerOn = True Then
DoPointer
D3DDevice.SetTexture 0, DrawTex(1)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, PointerStrip(0), Len(PointerStrip(0))
End If

'DrawBorders
D3DDevice.SetTexture 0, BorderTex
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, BorderL(0), Len(BorderL(0))
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, BorderR(0), Len(BorderR(0))

'//Draw the frame rate
    TextRect.Top = 0
    TextRect.bottom = 20
    TextRect.Right = 75
'    D3DX.DrawText MainFont, &HFFFFFFFF, PScore, TextRect, DT_TOP Or DT_LEFT
'    TextRect.Top = 20: TextRect.bottom = 40
'    D3DX.DrawText MainFont, &HFFFFFFFF, OScore, TextRect, DT_TOP Or DT_LEFT
    

D3DDevice.EndScene

D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

End If

End Sub

Public Sub DoValueBar(Xpos As Single, Ypos As Single, Max As Single, Val As Single)
Dim t As Single
Dim gx As Single, gy As Single

For t = 0 To Max

If t <= Val Then
    gx = 64: gy = 48
Else: gx = 80: gy = 48
End If

With GetRect
    .Top = gy
    .bottom = gy + 16
    .Left = gx
    .Right = gx + 16
End With

PutPoint.X = OffX + Xpos * 16 + t * 16
PutPoint.Y = OffY + Ypos * 16

D3DDevice.CopyRects DrawSurf(0), GetRect, 1, MenuSurf, PutPoint

Next t

End Sub

Public Function TransX(XVal As Single, ZVal As Single) As Single
Dim Z As Single
If ZVal >= 0 Then Z = -ZVal
If ZVal <= 0 Then Z = ZVal
Z = Z + 250
Z = 500 - Z

If ZVal >= 0 Then TransX = 320 + (XVal / Z) * 250
If ZVal <= 0 Then TransX = 320 + (XVal / Z) * 250

End Function

Public Function TransY(YVal As Single, ZVal As Single) As Single
Dim Z As Single
If ZVal >= 0 Then Z = -ZVal
If ZVal <= 0 Then Z = ZVal
Z = Z + 250
Z = 500 - Z

If ZVal >= 0 Then TransY = 240 + (YVal / Z) * 250
If ZVal <= 0 Then TransY = 240 + (YVal / Z) * 250
End Function

Public Function FogColor(ZVal As Single) As Long
Dim IC As Single
If Z >= 0 Then
IC = 255 / 1664
IC = Int(ZVal * IC)
If IC > 255 Then IC = 255
FogColor = RGB(255 - IC, 255 - IC, 255 - IC)
Else: FogColor = RGB(255, 255, 255)
End If

End Function

Public Function Rotate(ZAng As Single, XVal As Single, YVal As Single)
Dim Zn As Single, xx As Single, yy As Single

Zn = ZAng * (3.141592654 / 180)

xx = XVal: yy = YVal
RotX = xx * Cos(Zn) - yy * Sin(Zn)
RotY = yy * Cos(Zn) + xx * Sin(Zn)

End Function
