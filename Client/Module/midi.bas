Attribute VB_Name = "midi"

Option Explicit

Public Kind
Public lNote As Long
Public rc As Long
Public hMidi As Long
Public Channel As Long
Public numDevices As Long
Public curDevice As Long
Public KeyMap(255) As Long
Public Pitch As Long
Public Velocity As Long
Public noteLong As Long
Public Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Public Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long

Public Sub PlayNote(ByVal Note As Long)
On Error Resume Next
Dim midimsg As Long

If Note = 88 Then
    Sustain True
Else
    midimsg = &H90 + Channel + ((Pitch + Note) * &H100) + (Velocity * &H10000)
    midiOutShortMsg hMidi, midimsg
    frmPiano.pKey(Note - 1).BackColor = vbRed
End If

    lNote = Note
    
End Sub

Public Sub StopNote(ByVal Note As Long)

On Error Resume Next

Dim midimsg As Long

If Note = 88 Then

    Sustain False
    
Else

    midimsg = &H80 + ((Pitch + Note) * &H100) + Channel
    midiOutShortMsg hMidi, midimsg

    If frmPiano.pKey(Note - 1).Tag = "1" Then
        frmPiano.pKey(Note - 1).BackColor = vbWhite
    Else
        frmPiano.pKey(Note - 1).BackColor = vbBlack
    End If
    
End If

    If Note = lNote Then lNote = 0
    
End Sub

Public Sub Sustain(Active As Boolean) 'Piano Pedal

On Error Resume Next

If Active Then
    midiOutShortMsg hMidi, (&HB0 + Channel + &H4000 + &H7F0000)
Else
    midiOutShortMsg hMidi, (&HB0 + Channel + &H4000)
End If

End Sub

Public Sub InitializInstrument(Value) 'Initializ Instrument

On Error Resume Next

Dim midimsg As Long
Kind = Value

If Value = 128 Then 'Percussion
    Channel = 9
Else
    Channel = 0
    midimsg = (Value * &H100) + &HC0 + Channel
    midiOutShortMsg hMidi, midimsg
End If
    
End Sub

Public Sub KeyMapping(Value)

On Error Resume Next

Dim temp() As String
Dim x As Long

    For x = 300 To 347
        temp = Split(LoadResString(x), ",")
        KeyMap(CLng(temp(0))) = CLng(temp(Value))
    Next x
    
KeyMap(16) = 88

End Sub

Public Sub InitializeMidi() 'Initialize Midi

On Error Resume Next

Dim x As Long

midiOutClose hMidi
rc = midiOutOpen(hMidi, curDevice, 0, 0, 0)

If rc = 4 Then
    MsgBox "미디오류"
End If
    
End Sub
