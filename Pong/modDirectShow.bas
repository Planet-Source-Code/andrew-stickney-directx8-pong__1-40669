Attribute VB_Name = "DX8Music"
'==========================================
'
'       NAME: modDirectShow.bas
'       DESC: MP3 playing module for background music.
'
'       AUTHOR: Jack Hoxley
'       DATE: 29th July 2001
'
'       REQUIRES: Microsoft DirectX 8.0 Runtime libraries
'
'==========================================
'Updated to play Midi files by Drew"Da Booda"Stickney

'Volume is 100 max to 0 min
'balance is -5000 left to 5000 right
'speed is 50 slow to 220 fast

Option Explicit

'//DirectShow Objects
    Public DSAudio  As IBasicAudio         'Basic Audio Objectt
    Public DSEvent As IMediaEvent        'MediaEvent Object
    Public DSControl As IMediaControl    'MediaControl Object
    Public DSPosition As IMediaPosition 'MediaPosition Object



Public Function TerminateMusic() As Boolean
On Error GoTo BailOut:

    If ObjPtr(DSControl) > 0 Then
        DSControl.Stop
    End If
                
    If ObjPtr(DSAudio) Then Set DSAudio = Nothing
    If ObjPtr(DSEvent) Then Set DSEvent = Nothing
    If ObjPtr(DSControl) Then Set DSControl = Nothing
    If ObjPtr(DSPosition) Then Set DSPosition = Nothing
                
    TerminateMusic = True
    Exit Function
BailOut:
    TerminateMusic = False
    Debug.Print "ERROR: modDirectShow.TerminateEngine()"
    Debug.Print "     ", Err.Number, Err.Description
End Function


Public Function LoadMusic(FileName As String) As Boolean
On Error GoTo BailOut:

    '//0. Any variables
    
    '//1. Destroy existing instances
        If Not (TerminateMusic() = True) Then GoTo BailOut:
        
    '//2. Setup a filter graph for the file
        Set DSControl = New FilgraphManager
        Call DSControl.RenderFile(FileName)
    
    '//3. Setup the basic audio object
        Set DSAudio = DSControl
        DSAudio.Volume = 0
        DSAudio.Balance = 0
    
    '//4. Setup the media event and position objects
        Set DSEvent = DSControl
        Set DSPosition = DSControl
        If ObjPtr(DSPosition) Then DSPosition.Rate = 1#
        DSPosition.CurrentPosition = 0
           
    '//5. Done!

    LoadMusic = True
    Exit Function
BailOut:
    LoadMusic = False
    Debug.Print "ERROR: modDirectShow.Loadmusic()"
    Debug.Print "     ", Err.Number, Err.Description
End Function


Public Function SetPlayBackSpeed(Speed As Single) As Boolean
On Error GoTo BailOut:

    If ObjPtr(DSPosition) > 0 Then
        DSPosition.Rate = Speed
    End If

    SetPlayBackSpeed = True
    Exit Function
BailOut:
    SetPlayBackSpeed = False
    Debug.Print "ERROR: modDirectShow.SetPlayBackSpeed()"
    Debug.Print "     ", Err.Number, Err.Description
End Function


Public Function SetPlayBackVolume(Volume As Long) As Boolean
On Error GoTo BailOut:
    'invert volume
    Volume = 100 - Volume
    '//Set the new volume
    If ObjPtr(DSControl) > 0 Then
        DSAudio.Volume = Volume * (-Volume * 0.25) ' makes it in the 0 to -4000 range (-4000 is almost silent)
    End If

    SetPlayBackVolume = True
    Exit Function
BailOut:
    SetPlayBackVolume = False
    Debug.Print "ERROR: modDirectShow.SetPlayBackVolume()"
    Debug.Print "     ", Err.Number, Err.Description
End Function

Public Function SetPlayBackBalance(Balance As Long) As Boolean
On Error GoTo BailOut:

    If ObjPtr(DSControl) > 0 Then
        DSAudio.Balance = Balance
    End If

    SetPlayBackBalance = True
    Exit Function
BailOut:
    SetPlayBackBalance = False
    Debug.Print "ERROR: modDirectShow.SetPlayBackBalance()"
    Debug.Print "     ", Err.Number, Err.Description
End Function


Public Function PlayMusic() As Boolean
On Error GoTo BailOut:

    DSControl.Run

    PlayMusic = True
    Exit Function
BailOut:
    PlayMusic = False
    Debug.Print "ERROR: modDirectShow.PlayMP3()"
    Debug.Print "     ", Err.Number, Err.Description
End Function

Public Function StopMusic() As Boolean
On Error GoTo BailOut:
    
    DSControl.Stop
    DSPosition.CurrentPosition = 0 'set it back to the beginning
    StopMusic = True
    Exit Function
BailOut:
    StopMusic = False
    Debug.Print "ERROR: modDirectShow.StopMP3()"
    Debug.Print "     ", Err.Number, Err.Description
End Function

Public Function PauseMusic() As Boolean
On Error GoTo BailOut:
    
    DSControl.Stop

    PauseMusic = True
    Exit Function
BailOut:
    PauseMusic = False
    Debug.Print "ERROR: modDirectShow.PauseMP3()"
    Debug.Print "     ", Err.Number, Err.Description
End Function
