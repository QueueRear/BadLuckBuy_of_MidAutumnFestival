VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Ê¥µ®µ¹Ã¹¹í"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   9315
   ScaleWidth      =   11985
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   10920
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   10920
      Top             =   960
   End
   Begin VB.Image Image3 
      Height          =   1560
      Left            =   6120
      Picture         =   "Form1.frx":3C853
      Top             =   7680
      Width           =   1560
   End
   Begin VB.Image Image2 
      Height          =   1560
      Left            =   4080
      Picture         =   "Form1.frx":3D40B
      Top             =   7440
      Width           =   1560
   End
   Begin VB.Image Image1 
      Height          =   5955
      Left            =   2280
      Top             =   1080
      Visible         =   0   'False
      Width           =   6180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ÇëÇ×ÊÖ»½³öÕâ´ÎµÄµ¹Ã¹¹í°É£¡"
      BeginProperty Font 
         Name            =   "ºÚÌå"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00157AB0&
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Private Sub Image2_Click()
 Timer1.Enabled = True
 Image1.Visible = True
 Label1.Caption = "ÇëÇ×ÊÖ»½³öÕâ´ÎµÄµ¹Ã¹¹í°É£¡"
End Sub

Private Sub Image3_Click()
 Timer1.Enabled = False
 Image1.Picture = LoadPicture("D:\×ÀÃæ\ÖÐÇïµ¹Ã¹¹í\¹·ÑÛ.jpg")
 Timer2.Enabled = True
End Sub

Private Sub Image3_DblClick()
 Image1.Visible = False
 Label1.Caption = "ÇëÇ×ÊÖ»½³öÕâ´ÎµÄµ¹Ã¹¹í°É£¡"
 Timer2.Enabled = False
End Sub

Private Sub Timer1_Timer()
 Randomize
 a = LTrim(Str(Int(Rnd() * 101 + 1)))
 Image1.Picture = LoadPicture("D:\×ÀÃæ\ÖÐÇïµ¹Ã¹¹í\" + a + ".jpg")
End Sub

Private Sub Timer2_Timer()
 Image1.Picture = LoadPicture("D:\×ÀÃæ\ÖÐÇïµ¹Ã¹¹í\" + a + ".jpg")
 Label1.Caption = "µ¹Ã¹¹íÔÚ´Ë£¡"
 Timer2.Enabled = False
End Sub
