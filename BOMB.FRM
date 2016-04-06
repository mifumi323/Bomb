VERSION 5.00
Begin VB.Form Bomb 
   Appearance      =   0  'ﾌﾗｯﾄ
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "Bomber丸"
   ClientHeight    =   4515
   ClientLeft      =   1080
   ClientTop       =   1755
   ClientWidth     =   6270
   ForeColor       =   &H80000008&
   Icon            =   "BOMB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   4515
   ScaleWidth      =   6270
   Begin VB.Timer TimerBlink 
      Interval        =   200
      Left            =   3360
      Top             =   0
   End
   Begin VB.Timer TimerBomb 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   2880
      Top             =   0
   End
   Begin VB.Timer TimerMsg 
      Left            =   2400
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1920
      Top             =   0
   End
   Begin VB.Shape BombH 
      BackColor       =   &H000080FF&
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   255
      Left            =   3240
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Shape BombV 
      BackColor       =   &H000080FF&
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   975
      Left            =   2880
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape BombHB 
      BackColor       =   &H000000FF&
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   255
      Left            =   3240
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Shape BombVB 
      BackColor       =   &H000000FF&
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   975
      Left            =   2520
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   36
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   720
      Left            =   1920
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image Bomb 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   480
      Left            =   1920
      Picture         =   "BOMB.frx":030A
      Top             =   480
      Width           =   480
   End
   Begin VB.Image PicBlack 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   480
      Index           =   3
      Left            =   1440
      Picture         =   "BOMB.frx":0614
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image PicBlack 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   480
      Index           =   2
      Left            =   1440
      Picture         =   "BOMB.frx":091E
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image PicBlack 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   480
      Index           =   1
      Left            =   960
      Picture         =   "BOMB.frx":0C28
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image PicBlack 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   480
      Index           =   0
      Left            =   1440
      Picture         =   "BOMB.frx":0F32
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image PicWhite 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   480
      Index           =   3
      Left            =   480
      Picture         =   "BOMB.frx":123C
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image PicWhite 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   480
      Index           =   2
      Left            =   480
      Picture         =   "BOMB.frx":1546
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image PicWhite 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   480
      Index           =   1
      Left            =   0
      Picture         =   "BOMB.frx":1850
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image PicWhite 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   480
      Index           =   0
      Left            =   480
      Picture         =   "BOMB.frx":1B5A
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Black 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   480
      Left            =   960
      Picture         =   "BOMB.frx":1E64
      Top             =   0
      Width           =   480
   End
   Begin VB.Image White 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   480
      Left            =   0
      Picture         =   "BOMB.frx":216E
      Top             =   0
      Width           =   480
   End
   Begin VB.Menu M_Game 
      Caption         =   "ゲーム(&G&ｹ)"
      Begin VB.Menu M_Start 
         Caption         =   "スタート(&S&ｽ)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu M_Set 
         Caption         =   "設定(&E&ｾ)"
      End
      Begin VB.Menu S 
         Caption         =   "-"
      End
      Begin VB.Menu M_Exit 
         Caption         =   "終了(&X&ｼ)"
      End
   End
End
Attribute VB_Name = "Bomb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub PlaySound Lib "winmm.dll" (ByVal FileName As String, ByVal hMod As Long, ByVal Flag As Long)

Dim P1          As PlayerData
Dim P2          As PlayerData
Dim Bom         As Bakudan

Dim Started     As Integer
Dim Nokori      As Integer

Private Sub Bakuhatsu()
    PlaySound MyPath & "Bomb.WAV", 0, 1
    BombVB.Move 480 * Bom.x + 30, 0, 420, ScaleHeight
    BombV.Move 480 * Bom.x + 60, 0, 360, ScaleHeight
    BombHB.Move 0, 480 * Bom.y + 30, ScaleWidth, 420
    BombH.Move 0, 480 * Bom.y + 60, ScaleWidth, 360
    BombVB.Visible = True
    BombV.Visible = True
    BombHB.Visible = True
    BombH.Visible = True
    TimerBomb.Enabled = True
    If Bom.x = P1.x Or Bom.y = P1.y And P1.Blink = 0 Then WLose
    If Bom.x = P2.x Or Bom.y = P2.y And P2.Blink = 0 Then BLose
    Hantei
End Sub

'黒を元の位置に戻します。
Private Sub BlackPlace()
    Black.Picture = PicBlack(3).Picture
    Black.Move 5760, 3840
    P2.x = 12: P2.y = 8
End Sub

Private Sub BLose()
    P2.Stock = P2.Stock - 1
    P2.Blink = 10
    BlackPlace
End Sub

Private Sub BombMove(x, y)
    Bom.x = x: Bom.y = y
    If Bom.x < 0 Then Bom.x = 1
    If Bom.x > 12 Then Bom.x = 11
    If Bom.y < 0 Then Bom.y = 1
    If Bom.y > 8 Then Bom.y = 7
    Bomb.Move 480 * Bom.x, 480 * Bom.y
    If Bom.x = P1.x And Bom.y = P1.y And P1.Blink = 0 Then WLose
    If Bom.x = P2.x And Bom.y = P2.y And P2.Blink = 0 Then BLose
End Sub

'爆弾を元の位置に戻します。
Private Sub BombPlace()
    BombMove 6, 4
End Sub

Private Sub BWalk(KeyCode)
    Select Case KeyCode
    Case 104                        '上
        p = 0
        If P2.x = P1.x And P2.y - 1 = P1.y Then GoTo A
        If P2.y > 0 Then P2.y = P2.y - 1
        If Bom.y = P2.y And Bom.x = P2.x Then
            BombMove Bom.x, Bom.y - 1
        End If
    Case 100                        '左
        p = 1
        If P2.x - 1 = P1.x And P2.y = P1.y Then GoTo A
        If P2.x > 0 Then P2.x = P2.x - 1
        If Bom.y = P2.y And Bom.x = P2.x Then
            BombMove Bom.x - 1, Bom.y
        End If
    Case 102                        '右
        p = 2
        If P2.x + 1 = P1.x And P2.y = P1.y Then GoTo A
        If P2.x < 12 Then P2.x = P2.x + 1
        If Bom.y = P2.y And Bom.x = P2.x Then
            BombMove Bom.x + 1, Bom.y
        End If
    Case 98                        '下
        p = 3
        If P2.x = P1.x And P2.y + 1 = P1.y Then GoTo A
        If P2.y < 8 Then P2.y = P2.y + 1
        If Bom.y = P2.y And Bom.x = P2.x Then
            BombMove Bom.x, Bom.y + 1
        End If
    Case Else: Exit Sub
    End Select
    Black.Move 480 * P2.x, 480 * P2.y
A:
    Black.Picture = PicBlack(p).Picture
End Sub

'ゲーム状況を初期化します。
Private Sub First()
    WhitePlace
    BlackPlace
    Bom.x = 6: Bom.y = 4
    Bomb.Move 2880, 1920
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim B As Integer, W As Integer
    If Started Then
        WWalk KeyCode
        BWalk KeyCode
        Hantei
    End If
End Sub

Private Sub Form_Load()
    SetPath
    DataLoad
    dx = Width - ScaleWidth
    dy = Height - ScaleHeight
    Move 0, 0, 6240 + dx, 4320 + dy
    BackColor = &HC000&
    Label1.Move ScaleWidth / 2 - Label1.Width / 2, ScaleHeight / 2 - Label1.Height / 2
    MIDIOpen MyPath & "bomb.mid"
    First
    MoveForm Me
    Show
    Do
        If Left(M_Start.Caption, 1) = "ポ" Then
            If MIDIStopped() Then MIDIPlay
        End If
        DoEvents
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MIDIStop
    MIDIClose
    DataSave
    End
End Sub

Private Sub Hantei()
    If P1.Stock = 0 Then A = 1
    If P2.Stock = 0 Then A = A + 2
    Select Case A
    Case 0
        Exit Sub
    Case 1
        Msg = "黒の勝ちです！"
        Setting.BWin = Setting.BWin + 1
    Case 2
        Msg = "白の勝ちです！"
        Setting.WWin = Setting.WWin + 1
    Case 3
        Msg = "引き分けです！"
        Setting.BWin = Setting.BWin + 1
        Setting.WWin = Setting.WWin + 1
    End Select
    MsgBox Msg, 48, "ゲームセット"
    Timer1.Enabled = False
    Started = False
    M_Start.Caption = "スタート(&S&ｽ)"
    MIDIStop
    M_Set.Enabled = True
End Sub

Private Sub M_Exit_Click()
    Unload Me
End Sub

Private Sub M_Set_Click()
    BombSet.Show 1
End Sub

Private Sub M_Start_Click()
    Select Case Left(M_Start.Caption, 1)
    Case "ス"
        Message "スタート", 3
        Started = True
        Timer1.Enabled = True
        Nokori = Setting.Jikan
        Caption = "Bomber丸　残り：" & Nokori & "秒"
        P1.Stock = Setting.Stock
        P2.Stock = Setting.Stock
        P1.Blink = 0
        P2.Blink = 0
        First
        MIDIPlay
        M_Start.Caption = "ポーズ(&P&ﾎ)"
        M_Set.Enabled = False
    Case "ポ"
        Started = False
        Timer1.Enabled = False
        M_Start.Caption = "再開(&S&ｽ)"
        M_Set.Enabled = True
        Message "ポーズ", 0
        MIDIStop
    Case "再"
        Started = True
        Timer1.Enabled = True
        M_Start.Caption = "ポーズ(&P&ﾎ)"
        M_Set.Enabled = False
        Message "再開", 3
        MIDIPlay
    End Select
End Sub

Private Sub Message(Msg$, Timing%)
    Label1.Caption = Msg$
    Label1.Visible = True
    TimerMsg.Interval = 1000 * Timing%
End Sub

Private Sub Timer1_Timer()
    Nokori = Nokori - 1
    If Nokori = 0 Then
        Caption = "Bomber丸　終わり"
        Timer1.Enabled = False
        Started = False
        M_Start.Caption = "スタート(&S&ｽ)"
        M_Set.Enabled = True
        Message "タイムアップ", 3
        MsgBox "時間切れ！", 48, "引き分け"
    Else
        Caption = "Bomber丸　残り：" & Nokori & "秒"
        If Nokori <= 10 Then Message Str(Nokori), 1
        If Nokori Mod Setting.Timing = 0 Then Bakuhatsu
    End If
End Sub

Private Sub TimerBlink_Timer()
    If P1.Blink Then
        P1.Blink = P1.Blink - 1
        White.Visible = Not White.Visible
    End If
    If P2.Blink Then
        P2.Blink = P2.Blink - 1
        Black.Visible = Not Black.Visible
    End If
End Sub

Private Sub TimerBomb_Timer()
    BombVB.Visible = False
    BombV.Visible = False
    BombHB.Visible = False
    BombH.Visible = False
    TimerBomb.Enabled = False
    If Started Then BombPlace
End Sub

Private Sub TimerMsg_Timer()
    Label1.Visible = False
    TimerMsg.Interval = 0
End Sub

'白を元の位置に戻します。
Private Sub WhitePlace()
    White.Picture = PicWhite(3).Picture
    White.Move 0, 0
    P1.x = 0: P1.y = 0
End Sub

Private Sub WLose()
    P1.Stock = P1.Stock - 1
    P1.Blink = 10
    WhitePlace
End Sub

Private Sub WWalk(KeyCode%)
    Select Case KeyCode%
    Case &H26                       '上
        p = 0
        If P2.x = P1.x And P2.y + 1 = P1.y Then GoTo B
        If P1.y > 0 Then P1.y = P1.y - 1
        If Bom.y = P1.y And Bom.x = P1.x Then
            BombMove Bom.x, Bom.y - 1
        End If
    Case &H25                       '左
        p = 1
        If P2.x + 1 = P1.x And P2.y = P1.y Then GoTo B
        If P1.x > 0 Then P1.x = P1.x - 1
        If Bom.y = P1.y And Bom.x = P1.x Then
            BombMove Bom.x - 1, Bom.y
        End If
    Case &H27                       '右
        p = 2
        If P2.x - 1 = P1.x And P2.y = P1.y Then GoTo B
        If P1.x < 12 Then P1.x = P1.x + 1
        If Bom.y = P1.y And Bom.x = P1.x Then
            BombMove Bom.x + 1, Bom.y
        End If
    Case &H28                       '下
        p = 3
        If P2.x = P1.x And P2.y - 1 = P1.y Then GoTo B
        If P1.y < 8 Then P1.y = P1.y + 1
        If Bom.y = P1.y And Bom.x = P1.x Then
            BombMove Bom.x, Bom.y + 1
        End If
    Case Else: Exit Sub
    End Select
    White.Move 480 * P1.x, 480 * P1.y
B:
    White.Picture = PicWhite(p).Picture
End Sub

