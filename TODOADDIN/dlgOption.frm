VERSION 5.00
Begin VB.Form dlgOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ToDo Add-In options"
   ClientHeight    =   3180
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4140
   ClipControls    =   0   'False
   Icon            =   "dlgOption.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkToolbarAlign 
      Caption         =   "Toolbar aligned on the left"
      Height          =   225
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   2235
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3510
      Top             =   1170
   End
   Begin VB.CommandButton butRefresh 
      Caption         =   "&Refersh list"
      Height          =   345
      Left            =   180
      TabIndex        =   3
      Top             =   2730
      Width           =   2175
   End
   Begin VB.ListBox lstTags 
      Height          =   1620
      Left            =   180
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton butAbout 
      Height          =   465
      Left            =   3420
      MaskColor       =   &H00FF00FF&
      Picture         =   "dlgOption.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "About..."
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CheckBox chkAlwaysOnTop 
      Caption         =   "Always on &top"
      Height          =   225
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   1845
   End
   Begin VB.CommandButton butCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2820
      TabIndex        =   6
      Top             =   2700
      Width           =   1215
   End
   Begin VB.CommandButton butOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2820
      TabIndex        =   5
      Top             =   2250
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tags took into account now:"
      Height          =   195
      Left            =   210
      TabIndex        =   7
      Top             =   870
      Width           =   2040
   End
End
Attribute VB_Name = "dlgOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dlgCancel As Boolean
Private dlgVBInstance As vbide.VBE
Private Sub butAbout_Click()

frmAbout.dlgShow

End Sub

Private Sub butCancel_Click()

Unload Me

End Sub


Private Sub butOK_Click()

dlgCancel = False

'save setting
SaveSetting sAppName$, sRegSection$, "AlwaysOnTop", Str$(chkAlwaysOnTop.Value = vbChecked)
SaveSetting sAppName$, sRegSection$, "ToolbarLeft", Str$(chkToolbarAlign.Value = vbChecked)

Unload Me

End Sub


Private Sub butRefresh_Click()

'fill the collection with tags
Set colTagList = GetTagsCol(dlgVBInstance, sTagListSignature$)
'fill the list with tags
FillListWithTags lstTags, colTagList

End Sub

Private Sub Form_Load()

'center dialog
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2

dlgCancel = True

If CBool(GetSetting(sAppName$, sRegSection$, "AlwaysOnTop", -1)) Then
    chkAlwaysOnTop.Value = vbChecked
    'OnTop Me, True 'can't set it always on top because it will raise an error in me.show vbmodal
End If

If CBool(GetSetting(sAppName$, sRegSection$, "ToolbarLeft", -1)) Then
    chkToolbarAlign.Value = vbChecked
End If

FillListWithTags lstTags, colTagList

End Sub



Private Sub Timer1_Timer()

OnTop Me, CBool(GetSetting(sAppName$, sRegSection$, "AlwaysOnTop", -1))

Timer1.Enabled = False

End Sub

'displays the dialog and returns TRUE if OK and False if cancel was pressed or window was closed by other means (alt+f4 etc)
Public Function dlgShow(objVBInstance As vbide.VBE, Optional frmParent As Form = Nothing) As Boolean

Set dlgVBInstance = objVBInstance

Me.Show vbModal, frmParent

'return the result (true if ok, false if cancel was pressed or window was closed by other means)
dlgShow = Not dlgCancel

Set dlgVBInstance = Nothing

End Function

Private Sub FillListWithTags(lstTags As ListBox, colTags As Collection)

With lstTags
    .Clear
    Dim lTag As Long
    For lTag = 1 To colTags.Count
        .AddItem colTags(lTag)
    Next lTag
End With 'lsttags

End Sub

