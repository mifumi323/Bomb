VERSION 5.00
Begin VB.Form BombSet 
   Appearance      =   0  'ﾌﾗｯﾄ
   BackColor       =   &H80000005&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "設定"
   ClientHeight    =   3105
   ClientLeft      =   2295
   ClientTop       =   2970
   ClientWidth     =   4005
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "BOMBSET.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   3105
   ScaleWidth      =   4005
   Begin VB.CommandButton Customize 
      Appearance      =   0  'ﾌﾗｯﾄ
      Caption         =   "ｷｰ&ｶｽﾀﾏｲｽﾞ"
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Clear_ 
      Appearance      =   0  'ﾌﾗｯﾄ
      Caption         =   "&ｸﾘｱ"
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Cancel 
      Appearance      =   0  'ﾌﾗｯﾄ
      Cancel          =   -1  'True
      Caption         =   "&ｷｬﾝｾﾙ"
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton OK 
      Appearance      =   0  'ﾌﾗｯﾄ
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3000
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin VB.HScrollBar STime 
      Height          =   270
      LargeChange     =   100
      Left            =   120
      Max             =   1000
      TabIndex        =   5
      Top             =   1920
      Value           =   600
      Width           =   2775
   End
   Begin VB.HScrollBar SStock 
      Height          =   270
      LargeChange     =   5
      Left            =   120
      Max             =   30
      Min             =   1
      TabIndex        =   3
      Top             =   1200
      Value           =   5
      Width           =   2775
   End
   Begin VB.HScrollBar STiming 
      Height          =   270
      Left            =   120
      Max             =   10
      Min             =   5
      TabIndex        =   1
      Top             =   480
      Value           =   5
      Width           =   2775
   End
   Begin VB.Label BWin 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label3 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "："
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label WWin 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   2760
      Width           =   495
   End
   Begin VB.Image B 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   480
      Left            =   1800
      Picture         =   "BOMBSET.frx":030A
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image W 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   480
      Left            =   1080
      Picture         =   "BOMBSET.frx":0614
      Top             =   2280
      Width           =   480
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "対戦成績"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label LTime 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      Caption         =   "制限時間(&I&ｾ)：600秒"
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label LStock 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      Caption         =   "ストック数(&S&ｽ)：5"
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label LTiming 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      Caption         =   "爆発までの時間(&T&ﾀ)：5秒"
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "BombSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub Clear__Click()
    Msg = "データと設定が最初の状態に戻ります。" & Chr(10) & "初期化してもよろしいですか？"
    If MsgBox(Msg, 36, "データクリア") = 6 Then
        DataClear
        STiming.Value = Setting.Timing
        SStock.Value = Setting.Stock
        STime.Value = Setting.Jikan
        WWin.Caption = Setting.WWin
        BWin.Caption = Setting.BWin
    End If
End Sub

Private Sub Customize_Click()
    BombKey.Show 1
End Sub

Private Sub Form_Load()
    STime.Min = STiming.Value * SStock * 2
    STiming.Value = Setting.Timing
    SStock.Value = Setting.Stock
    STime.Value = Setting.Jikan
    WWin.Caption = Setting.WWin
    BWin.Caption = Setting.BWin
    MoveForm Me
End Sub

Private Sub OK_Click()
    Setting.Timing = STiming.Value
    Setting.Stock = SStock.Value
    Setting.Jikan = STime.Value
    Unload Me
End Sub

Private Sub SStock_Change()
    LStock.Caption = "ストック数(&S&ｽ)：" & SStock.Value
    STime.Min = STiming.Value * SStock * 2
End Sub

Private Sub STime_Change()
    LTime.Caption = "制限時間(&I&ｾ)：" & STime.Value & "秒"
End Sub

Private Sub STiming_Change()
    LTiming.Caption = "爆発までの時間(&T&ﾀ)：" & STiming.Value & "秒"
    STime.Min = STiming.Value * SStock * 2
End Sub

