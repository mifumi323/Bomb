Attribute VB_Name = "Misc"
Option Explicit

Dim C As Integer
Public MyPath As String
Public OK As Boolean

Public x As Integer, y As Integer
Public Function ComOpen(ByVal FileNumber As Integer) As String
    Do
        If EOF(FileNumber) = True Then Exit Function
        Input #FileNumber, ComOpen
        If ComOpen <> "" And Left(ComOpen, 1) <> ";" Then Exit Do
    Loop
End Function
Public Function GetPathFromFileName(ByVal FileName As String) As String
    GetPathFromFileName = Left$(FileName, Len(FileName) - InStr(1, Reverse(FileName), "\", vbTextCompare))
End Function
Public Function Reverse(ByVal Str As String) As String
'    Dim Temp As String
'    For x = 0 To Len(Str) - 1
'        Temp = Temp & Mid$(Str, Len(Str) - x, 1)
'    Next
    Reverse = StrReverse(Str)
End Function
Public Sub SetMax(ByVal A As Integer, ByVal B As Integer)
    If A > B Then B = A
End Sub
Public Function TimeFormat(ByVal Second As Integer) As String
    Dim Hours As Integer, Minutes As Integer, Seconds As Integer
    Hours = Second \ 3600
    Minutes = (Second - Hours * 3600) \ 60
    Seconds = Second - Hours * 3600 - Minutes * 60
    TimeFormat = Format$(Hours, "00") & ":" & Format(Minutes, "00") & ":" & Format(Seconds, "00")
End Function
Public Sub WaitOK()
    OK = False
    Do Until OK
        DoEvents
    Loop
End Sub
Public Sub DblCheck()
    If App.PrevInstance = True Then End
End Sub
Public Function Directory(ByVal F As String) As String
    If F = "" Or Dir(F) = "" Then
        Directory = ""
    Else
        Directory = F
    End If
End Function
Public Function LastList(ByVal L As Object) As Boolean
    If L.ListIndex = L.ListCount - 1 Then
        LastList = True
    Else
        LastList = 0
    End If
End Function
Public Function MakePath(ByVal Path As String) As String
    MakePath = Path
    If Right$(MakePath, 1) <> "\" Then MakePath = MakePath & "\"
End Function
Public Sub MoveFormOwner(F As Form, OF As Form)
    x = OF.Left + (OF.Width - F.Width) / 2
    y = OF.Top + (OF.Height - F.Height) / 2
    F.Move x, y
End Sub
Public Sub MoveFormOwnerTo(F As Form, OF As Form, ByVal W As Integer, ByVal H As Integer)
    FormSize F, W, H
    MoveFormOwner F, OF
End Sub
Public Sub MsgError(ByVal Message As String)
    MsgBox Message, vbExclamation, "ƒGƒ‰["
End Sub
Public Sub Swap(A, B)
    Dim C
    C = A
    A = B
    B = C
End Sub
Public Sub InputBoxS(ByVal Msg As String, Words As String)
    Do
        Words = InputBox(Msg)
    Loop Until Words
End Sub
Public Sub InputBoxI(ByVal Msg As String, Number As Integer)
    Do
        Number = Val(InputBox(Msg))
    Loop Until Number
End Sub
Public Sub InputBoxL(ByVal Msg As String, Number As Long)
    Do
        Number = Val(InputBox(Msg))
    Loop Until Number
End Sub
Public Sub FormSize(F As Form, ByVal W As Integer, ByVal H As Integer)
    x = F.Width - F.ScaleWidth * Screen.TwipsPerPixelX
    y = F.Height - F.ScaleHeight * Screen.TwipsPerPixelY
    F.Width = W * Screen.TwipsPerPixelX + x
    F.Height = H * Screen.TwipsPerPixelY + y
End Sub
Public Sub MoveForm(F As Form)
    x = (Screen.Width - F.Width) / 2
    y = (Screen.Height - F.Height) / 2
    F.Move x, y
End Sub
Public Sub MoveFormTo(F As Form, ByVal W As Integer, ByVal H As Integer)
    FormSize F, W, H
    MoveForm F
End Sub
Public Sub SetPath()
    MyPath = MakePath(App.Path)
End Sub
Public Sub Inc(Variable As Long, Optional ByVal Value As Long = 1, Optional ByVal Max)
    Variable = Variable + Value
    If Not IsMissing(Max) Then
        If Variable > Max Then Variable = Max
    End If
End Sub
Public Sub Dec(Variable As Long, Optional ByVal Value As Long = 1, Optional ByVal Min)
    Variable = Variable - Value
    If Not IsMissing(Min) Then
        If Variable < Min Then Variable = Min
    End If
End Sub
