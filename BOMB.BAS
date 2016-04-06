Type Bakudan
    X           As Integer
    Y           As Integer
End Type

Type PlayerData
    X           As Integer
    Y           As Integer
    Stock       As Integer
    Blink       As Integer
End Type

Type SettingData
    Stock       As Integer
    Jikan       As Integer
    Timing      As Integer
    WWin        As Integer
    BWin        As Integer
    Key(7)      As Integer
End Type

Global Setting  As SettingData

Sub DataClear ()
    Setting.Stock = 5
    Setting.Jikan = 600
    Setting.Timing = 5
    Setting.WWin = 0
    Setting.BWin = 0
    Setting.Key(0) = &H26
    Setting.Key(1) = &H25
    Setting.Key(2) = &H27
    Setting.Key(3) = &H28
    Setting.Key(4) = 104
    Setting.Key(5) = 100
    Setting.Key(6) = 102
    Setting.Key(7) = 98
End Sub

Sub DataLoad ()
    On Error Resume Next
    If Dir(MyPath & "Bomb.ini") = "" Then
        DataClear
        Exit Sub
    End If
    Open MyPath & "Bomb.INI" For Binary As #1
        Get #1, , Setting
    Close
End Sub

Sub DataSave ()
    Open MyPath & "Bomb.INI" For Binary As #1
        Put #1, , Setting
    Close
    End
End Sub

