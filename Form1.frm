VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4275
   ClientLeft      =   4110
   ClientTop       =   465
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   4
      Left            =   405
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   270
      ScaleWidth      =   7020
      TabIndex        =   6
      Top             =   3030
      Width           =   7020
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   345
         TabIndex        =   10
         Top             =   30
         Width           =   585
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   4
         Left            =   45
         Picture         =   "Form1.frx":62FA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   3
      Left            =   405
      Picture         =   "Form1.frx":BF24
      ScaleHeight     =   195
      ScaleWidth      =   4095
      TabIndex        =   5
      Top             =   2115
      Width           =   4095
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Index           =   1
         Left            =   345
         TabIndex        =   8
         Top             =   0
         Width           =   585
      End
      Begin VB.Image Image1 
         Height          =   195
         Index           =   3
         Left            =   15
         Picture         =   "Form1.frx":E90A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Index           =   2
      Left            =   405
      Picture         =   "Form1.frx":112BC
      ScaleHeight     =   345
      ScaleWidth      =   8310
      TabIndex        =   4
      Top             =   3465
      Width           =   8310
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Index           =   4
         Left            =   330
         TabIndex        =   11
         Top             =   90
         Width           =   585
      End
      Begin VB.Image Image1 
         Height          =   345
         Index           =   2
         Left            =   30
         Picture         =   "Form1.frx":1A87E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   1
      Left            =   405
      Picture         =   "Form1.frx":23DE4
      ScaleHeight     =   270
      ScaleWidth      =   5430
      TabIndex        =   3
      Top             =   2550
      Width           =   5430
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Index           =   2
         Left            =   315
         TabIndex        =   9
         Top             =   45
         Width           =   585
      End
      Begin VB.Image Image1 
         Height          =   270
         Index           =   1
         Left            =   30
         Picture         =   "Form1.frx":28AA6
         Stretch         =   -1  'True
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Timer"
      Height          =   270
      Left            =   630
      TabIndex        =   2
      Top             =   750
      Value           =   1  'Checked
      Width           =   750
   End
   Begin VB.HScrollBar HScroll1 
      Enabled         =   0   'False
      Height          =   270
      Left            =   645
      SmallChange     =   5
      TabIndex        =   1
      Top             =   1185
      Width           =   4065
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   600
      Top             =   225
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   0
      Left            =   405
      Picture         =   "Form1.frx":2D6D8
      ScaleHeight     =   195
      ScaleWidth      =   4110
      TabIndex        =   0
      Top             =   1725
      Width           =   4110
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Index           =   0
         Left            =   375
         TabIndex        =   7
         Top             =   0
         Width           =   585
      End
      Begin VB.Image Image1 
         Height          =   195
         Index           =   0
         Left            =   30
         Picture         =   "Form1.frx":300F2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   120
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim nbcontrol As Integer
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Timer1.Enabled = True
        HScroll1.Enabled = False
    Else
        Timer1.Enabled = False
        HScroll1.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    i = 0
    nbcontrol = 4 '---> nbre of scrollbar (0 to 4)
    Me.BackColor = &H8000000F
    HScroll1.Max = 100
End Sub

Private Sub HScroll1_Change()
    HScroll1_Scroll
End Sub


Private Sub HScroll1_Scroll()

    For j = 0 To nbcontrol
        Image1(j).Width = HScroll1.Value * (Picture1(j).Width - (Image1(j).Left * 2)) / 100
        Label1(j).Caption = HScroll1.Value & " %"
    Next j
    
    Me.Caption = HScroll1.Value & " %"

End Sub
Private Sub Timer1_Timer()
    i = i + 1
    If i > 100 Then
        i = 0
    End If
    
    For j = 0 To nbcontrol
        Image1(j).Width = i * (Picture1(j).Width - (Image1(j).Left * 2)) / 100 'Image1.Width + 10
        Label1(j).Caption = i & " %"
    Next j
    
    Me.Caption = i & " %"
End Sub


