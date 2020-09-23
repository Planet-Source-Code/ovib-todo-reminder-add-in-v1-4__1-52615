VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   6450
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8850
   ClipControls    =   0   'False
   Icon            =   "frmHelp.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   30
      Top             =   30
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Why use ToDo Reminder Add-In"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   60
      TabIndex        =   8
      Top             =   870
      Width           =   4335
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp.frx":014A
         Height          =   2535
         Left            =   180
         TabIndex        =   9
         Top             =   270
         Width           =   4020
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Define your custom tags"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2505
      Left            =   60
      TabIndex        =   5
      Top             =   3870
      Width           =   8685
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp.frx":03FC
         Height          =   1245
         Left            =   4530
         TabIndex        =   7
         Top             =   270
         Width           =   4020
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp.frx":0532
         Height          =   1245
         Left            =   150
         TabIndex        =   6
         Top             =   270
         Width           =   4020
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   885
         Left            =   150
         Picture         =   "frmHelp.frx":067A
         Top             =   1560
         Width           =   6915
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Valid tagged lines"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   4410
      TabIndex        =   3
      Top             =   870
      Width           =   4335
      Begin VB.Image imgExample 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1755
         Left            =   150
         Picture         =   "frmHelp.frx":0C54
         Top             =   1080
         Width           =   4020
      End
      Begin VB.Label lblExplanation 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp.frx":15F2
         Height          =   825
         Left            =   180
         TabIndex        =   4
         Top             =   270
         Width           =   4020
      End
   End
   Begin VB.CommandButton butAbout 
      Height          =   465
      Left            =   6480
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmHelp.frx":16C4
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "About..."
      Top             =   180
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton butOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   7530
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   1530
      Picture         =   "frmHelp.frx":17C6
      Top             =   150
      Width           =   240
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   210
      X2              =   6960
      Y1              =   750
      Y2              =   750
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
      Left            =   600
      TabIndex        =   1
      Top             =   180
      Width           =   4665
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'displays the dialog and returns TRUE if OK and False if cancel was pressed or window was closed by other means (alt+f4 etc)
Public Function dlgShow(Optional frmParent As Form = Nothing) As Boolean

Me.Show vbModal, frmParent

End Function

Private Sub butAbout_Click()

frmAbout.dlgShow

End Sub

Private Sub butOK_Click()

Unload Me

End Sub

Private Sub Form_Load()

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



