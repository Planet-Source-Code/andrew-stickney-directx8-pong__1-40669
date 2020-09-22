Attribute VB_Name = "Globals"
Public Running As Boolean
Public GameOn As Boolean

Public MenuOn As Boolean
Public PointerOn As Boolean
Public Highlight As Boolean
Public Clicked As Boolean
Public OffX As Single, OffY As Single

Public PointerX As Single, PointerY As Single

Public SmokeCol As Long
Public SmokeOff As Single
Public SmokeTile As Single
Public SmokeZ As Single

Public BackCol As Long
Public BackOff As Single
Public BackTile As Single
Public BackZ As Single
Public BackRed As Single, BackGreen As Single, BackBlue As Single

Public MenuCol As Long

Public MusicVol As Long
Public SoundVol As Long
Public MenuSpeed As Long

Public Cursors As Single
Public CurType(15, 2) As Single
Public CurrentCur As Single
Public LastCur As Single
Public CurRects(15) As RECT
Public CurX As Single, CurY As Single

Public BSize As Single
Public BAng As Single
Public BXpos As Single
Public BYpos As Single
Public BInc As Single
Public BGIncX As Single
Public BGIncY As Single
Public BGSize As Single

Public PSize As Single
Public OSize As Single
Public PHSize As Single
Public OHSize As Single
Public PXPos As Single
Public PYPos As Single
Public OXPos As Single
Public OYPos As Single
Public PGOff As Single
Public OSpeed As Single
Public ORate As Single
Public OldPX As Single
Public PChange As Single


Public RotX As Single, RotY As Single
Public XOrigin As Single, YOrigin As Single
Public Ang As Single
Public PScore As Single
Public OScore As Single
Public CMPos As Single



Public Sub SaveUser()
Open App.Path + "\User.bdat" For Output As #1
Print #1, MusicVol, SoundVol, MenuSpeed
Print #1, BSize, BInc, PSize, OSize, OSpeed, ORate
Print #1, BackTile, SmokeTile, BackRed, BackGreen, BackBlue
Close #1

End Sub

Public Sub LoadUser()
Open App.Path + "\User.bdat" For Input As #1
Input #1, MusicVol, SoundVol, MenuSpeed
Input #1, BSize, BInc, PSize, OSize, OSpeed, ORate
Input #1, BackTile, SmokeTile, BackRed, BackGreen, BackBlue
Close #1

SetPlayBackVolume (MusicVol)
MaxVolume = SoundVol
BackCol = CLong(BackRed, BackGreen, BackBlue)

End Sub
