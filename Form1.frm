VERSION 5.00
Begin VB.Form frmPump 
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   12540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":0000
   MousePointer    =   99  'Benutzerdefiniert
   Picture         =   "Form1.frx":030A
   ScaleHeight     =   836
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   604
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox ckPump 
      BackColor       =   &H00008000&
      Height          =   270
      Left            =   7950
      Style           =   1  'Grafisch
      TabIndex        =   13
      ToolTipText     =   "Pump"
      Top             =   2715
      Width           =   405
   End
   Begin VB.Timer tmrTick 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3825
      Top             =   6795
   End
   Begin VB.PictureBox picDP 
      Appearance      =   0  '2D
      BackColor       =   &H00000080&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   60
      Index           =   2
      Left            =   4320
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   7
      Top             =   4395
      Width           =   45
   End
   Begin VB.PictureBox picDP 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   75
      Index           =   1
      Left            =   4305
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   6
      Top             =   3420
      Width           =   45
   End
   Begin VB.PictureBox picDP 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   75
      Index           =   0
      Left            =   4620
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   5
      Top             =   2565
      Width           =   45
   End
   Begin VB.VScrollBar scrPrice 
      Height          =   360
      Left            =   3345
      Max             =   0
      Min             =   32767
      TabIndex        =   4
      Top             =   4185
      Value           =   1239
      Width           =   240
   End
   Begin Projekt1.Counter cntPrice 
      Height          =   300
      Left            =   3705
      TabIndex        =   3
      Top             =   4215
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      ForeColor       =   128
      Value           =   1239
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Digits          =   4
   End
   Begin Projekt1.Counter cntLiters 
      Height          =   525
      Left            =   3375
      TabIndex        =   1
      Top             =   2205
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   926
      BackColor       =   8421504
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Digits          =   5
   End
   Begin VB.CommandButton btExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   3330
      TabIndex        =   0
      Top             =   10905
      Width           =   1485
   End
   Begin Projekt1.Counter cntEuro 
      Height          =   525
      Left            =   3375
      TabIndex        =   2
      Top             =   3075
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   926
      BackColor       =   8421504
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Digits          =   5
   End
   Begin VB.Label lbl 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Super Plus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   405
      Index           =   5
      Left            =   3240
      TabIndex        =   14
      Top             =   1605
      Width           =   1800
   End
   Begin VB.Label lbl 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cent"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   3
      Left            =   4650
      TabIndex        =   12
      Top             =   4245
      Width           =   330
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "UMGPET"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   795
      Index           =   4
      Left            =   2805
      TabIndex        =   11
      Top             =   345
      Width           =   2625
   End
   Begin VB.Label lbl 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Preis per Liter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   2
      Left            =   3645
      TabIndex        =   10
      Top             =   3975
      Width           =   1005
   End
   Begin VB.Label lbl 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Euro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Index           =   1
      Left            =   5100
      TabIndex        =   9
      Top             =   3135
      Width           =   765
   End
   Begin VB.Label lbl 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Liter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Index           =   0
      Left            =   5115
      TabIndex        =   8
      Top             =   2265
      Width           =   690
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "UMGPET"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   795
      Index           =   6
      Left            =   2775
      TabIndex        =   15
      Top             =   315
      Width           =   2625
   End
End
Attribute VB_Name = "frmPump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Trimmer     As cTrimmer
Private LastState   As Boolean

Private Sub btExit_Click()

    Unload Me

End Sub

Private Sub ckPump_Click()

    If LastState Then
        scrPrice_Change
    End If
    LastState = (ckPump = vbUnchecked)
    tmrTick.Enabled = Not LastState
    scrPrice.Enabled = LastState

End Sub

Private Sub Form_Load()

    Set Trimmer = New cTrimmer
    Trimmer.TrimForm Me

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        Trimmer.GrabForm Me
    End If

End Sub

Private Sub lbl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub scrPrice_Change()

    cntPrice = scrPrice
    cntLiters = 0
    cntEuro = 0

End Sub

Private Sub scrPrice_Scroll()

    scrPrice_Change

End Sub

Private Sub tmrTick_Timer()

    cntEuro = cntEuro + 0.1
    cntLiters = cntEuro / scrPrice * 100

End Sub

':) Ulli's VB Code Formatter V2.23.17 (2008-Jan-22 22:40)  Decl: 4  Code: 56  Total: 60 Lines
':) CommentOnly: 2 (3,3%)  Commented: 0 (0%)  Empty: 23 (38,3%)  Max Logic Depth: 2
