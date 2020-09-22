VERSION 5.00
Object = "{D35AEA85-7919-4C6A-ADA9-3539868BBB6E}#1.0#0"; "ExtendedTimer.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extended Timer"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   3810
   StartUpPosition =   3  'Windows Default
   Begin ExtendedTimer.ExtTimer ExtTimer1 
      Left            =   3060
      Top             =   240
      _ExtentX        =   688
      _ExtentY        =   688
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start Timer"
      Height          =   375
      Left            =   2220
      TabIndex        =   7
      Top             =   1140
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Timer Test"
      Height          =   1455
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   3675
      Begin VB.TextBox txtSeconds 
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Text            =   "1"
         Top             =   1080
         Width           =   915
      End
      Begin VB.TextBox txtMinutes 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Text            =   "0"
         Top             =   720
         Width           =   915
      End
      Begin VB.TextBox txtHour 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Text            =   "0"
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Seconds:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1140
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Minutes:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hours:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   465
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const ES_NUMBER = &H2000&


Private Sub Command1_Click()
If Command1.Caption = "&Start Timer" Then Command1.Caption = "&Stop Timer" Else Command1.Caption = "&Start Timer"
 ExtTimer1.Hours = txtHour
  ExtTimer1.Minutes = txtMinutes
   ExtTimer1.Seconds = txtSeconds
    ExtTimer1.Enabled = Not (ExtTimer1.Enabled)
End Sub

Private Sub ExtTimer1_Timer()
 Debug.Print "Timer1", Now
End Sub

Private Sub Form_Load()
Dim esStyle&
 esStyle = GetWindowLong(txtHour.hwnd, GWL_STYLE)
  esStyle = esStyle Or ES_NUMBER
   SetWindowLong txtHour.hwnd, GWL_STYLE, esStyle
    SetWindowLong txtMinutes.hwnd, GWL_STYLE, esStyle
     SetWindowLong txtSeconds.hwnd, GWL_STYLE, esStyle
End Sub


