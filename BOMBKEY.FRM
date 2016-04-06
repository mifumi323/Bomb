VERSION 5.00
Begin VB.Form BombKey 
   Appearance      =   0  'ﾌﾗｯﾄ
   BackColor       =   &H80000005&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "ｷｰｶｽﾀﾏｲｽﾞ"
   ClientHeight    =   1680
   ClientLeft      =   3015
   ClientTop       =   1860
   ClientWidth     =   2910
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   1680
   ScaleWidth      =   2910
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   2655
      TabIndex        =   2
      Top             =   480
      Width           =   2655
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Caption         =   "閉じる(&C)"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   300
      Left            =   120
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "BombKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
    Label1.Caption = "KeyCode:" & Setting.Key(Combo1.ListIndex)
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    MoveForm Me
    Combo1.AddItem "白　上"
    Combo1.AddItem "白　左"
    Combo1.AddItem "白　右"
    Combo1.AddItem "白　下"
    Combo1.AddItem "黒　上"
    Combo1.AddItem "黒　左"
    Combo1.AddItem "黒　右"
    Combo1.AddItem "黒　下"
    Combo1.ListIndex = 0
End Sub

Private Sub Label1_Click()
    Pic1.SetFocus
End Sub

Private Sub Pic1_GotFocus()
    Label1.Caption = "KeyCode:" & Setting.Key(Combo1.ListIndex) & Chr(10) & Chr(13) & "キーを押してください"
End Sub

Private Sub Pic1_KeyDown(KeyCode As Integer, Shift As Integer)
    Setting.Key(Combo1.ListIndex) = KeyCode
    Label1.Caption = "KeyCode:" & Setting.Key(Combo1.ListIndex)
    Combo1.SetFocus
End Sub

Private Sub Pic1_LostFocus()
    Label1.Caption = "KeyCode:" & Setting.Key(Combo1.ListIndex)
End Sub

