Attribute VB_Name = "MIDI"
'MIDI演奏API
Declare Sub mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long)
Public Sub MIDIClose(Optional Als As String = "midifile")
    mciSendString "close """ & Als, """", 0, 0
End Sub
Public Sub MIDILoop(Optional Als As String = "midifile")
    If MIDIStopped(Als) Then mciSendString "play " & Als & " from 0", "", 0, 0
End Sub
'ファイルを開いてから再生
Public Sub MIDIOpenAndPlay(mciFile As String, Optional Als As String = "midifile")
    mciSendString "stop """ & Als & """", "", 0, 0
    mciSendString "close """ & Als & """", "", 0, 0
    mciSendString "open """ & mciFile & """ type sequencer alias """ & Als & """", "", 0, 0
    mciSendString "play """ & Als & """ from 0", "", 0, 0
End Sub
'ファイルを開く
Public Sub MIDIOpen(MIDIFile As String, Optional Als As String = "midifile")
    mciSendString "open """ & MIDIFile & """ type sequencer alias """ & Als & """", "", 0, 0
End Sub
Public Sub MIDIPlay(Optional Als As String = "midifile")
    mciSendString "play """ & Als & """ from 0", "", 0, 0
End Sub
Public Function MIDIStatus(Optional Als As String = "midifile") As String
    Dim Buf As String * 256
    mciSendString "status """ & Als & """ mode", Buf, 256, 0
    MIDIStatus = Buf
End Function
Public Function MIDIStopped(Optional Als As String = "midifile") As Boolean
    MIDIStopped = (Left$(LCase$(MIDIStatus(Als)), 7) = "stopped")
End Function
Public Sub MIDIStop(Optional Als As String = "midifile")
    mciSendString "stop """ & Als & """", "", 0, 0
End Sub
