VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4365
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5115
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   90
      Top             =   3870
   End
   Begin VB.CommandButton butOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1830
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000A&
      BorderWidth     =   2
      X1              =   1140
      X2              =   4860
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   1170
      Picture         =   "frmAbout.frx":0442
      Top             =   150
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ovib@osclabs.ro"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2850
      TabIndex        =   5
      Top             =   3180
      Width           =   2025
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can contact me at:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2010
      TabIndex        =   4
      Top             =   2940
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   210
      X2              =   4860
      Y1              =   810
      Y2              =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ToDo Reminder Add-In"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   240
      TabIndex        =   3
      Top             =   180
      Width           =   4665
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version: 1.0.15"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2850
      TabIndex        =   2
      Top             =   900
      Width           =   1830
   End
   Begin VB.Image Image1 
      Height          =   945
      Left            =   300
      Picture         =   "frmAbout.frx":0515
      Top             =   1050
      Width           =   1440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":149A
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   2070
      TabIndex        =   1
      Top             =   1410
      Width           =   3015
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   900
      Picture         =   "frmAbout.frx":1527
      Top             =   3000
      Width           =   705
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'displays the dialog and returns TRUE if OK and False if cancel was pressed or window was closed by other means (alt+f4 etc)
Public Function dlgShow(Optional frmParent As Form = Nothing) As Boolean

Me.Show vbModal, frmParent

End Function
Private Sub butOK_Click()

Unload Me

End Sub

Private Sub Form_Load()

Me.Caption = "About " & App.Title
lblVersion.Caption = "Version:" & Str$(App.Major) & "." & LTrim$(Str$(App.Minor)) & "." & LTrim$(Str$(App.Revision))

'center dialog
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2

'set on top if main window on top (to open in front of it)
'can't set it always on top because it will raise an error in me.show vbmodal
'OnTop Me, CBool(GetSetting(sAppName$, sRegSection$, "AlwaysOnTop", -1))

End Sub


Private Sub Timer1_Timer()

OnTop Me, CBool(GetSetting(sAppName$, sRegSection$, "AlwaysOnTop", -1))

Timer1.Enabled = False

End Sub


