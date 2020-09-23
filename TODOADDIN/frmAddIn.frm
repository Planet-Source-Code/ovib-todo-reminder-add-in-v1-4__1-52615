VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddIn 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "TODO list"
   ClientHeight    =   1995
   ClientLeft      =   4740
   ClientTop       =   4920
   ClientWidth     =   7995
   Icon            =   "frmAddIn.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   6870
      Top             =   1230
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":000C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":011E
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":0230
            Key             =   "SortA"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":0342
            Key             =   "SortD"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":0454
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":0566
            Key             =   "Option"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbToolbar1 
      Align           =   3  'Align Left
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageKey        =   "Exit"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh (CTRL+R)"
            ImageKey        =   "Refresh"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SortA"
            Object.ToolTipText     =   "Sort Ascending"
            ImageKey        =   "SortA"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SortD"
            Object.ToolTipText     =   "Sort descending"
            ImageKey        =   "SortD"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Option"
            Object.ToolTipText     =   "Options"
            ImageKey        =   "Option"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageKey        =   "Help"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid flxTODO 
      Height          =   1695
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   2990
      _Version        =   393216
      FixedCols       =   0
      MergeCells      =   3
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Image imgMenuBar 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   1260
      Picture         =   "frmAddIn.frx":0678
      Top             =   1740
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Menu mnuGridPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuCurrentTag 
         Caption         =   "Switch tag"
         Begin VB.Menu mnuTag 
            Caption         =   "Select a tag"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuGoTo 
         Caption         =   "Go to line in code pane (2x click)"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuAlwaysOnTop 
         Caption         =   "Always on top"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public VBInstance As vbide.VBE
Public Connect As Connect


Private Sub flxTODO_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    'rightclik on the grid, check if the clicked line contains a todo item
    If flxTODO.MouseRow > 1 Then
        'clicked line contains a todo item
            
        'select current line
        With flxTODO
            .Row = .MouseRow
            .Col = 0
            .ColSel = .Cols - 1
            .RowSel = .Row
        End With 'flxtodo
        
        mnuGoTo.Enabled = True
    Else
        mnuGoTo.Enabled = False
    End If
    
    PopupMenu mnuGridPopUp
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If (Shift = vbCtrlMask And KeyCode = Asc("R")) Or KeyCode = vbKeyF5 Then
    mnuRefresh_Click
ElseIf KeyCode = vbKeyF1 Then
    mnuHelp_Click
End If

End Sub


Private Sub mnuAbout_Click()

frmAbout.dlgShow

End Sub


Private Sub mnuAlwaysOnTop_Click()

'toggle checked property
mnuAlwaysOnTop.Checked = Not mnuAlwaysOnTop.Checked
'set window onTop state
OnTop Me, mnuAlwaysOnTop.Checked
'save setting
SaveSetting sAppName$, sRegSection$, "AlwaysOnTop", Str$(mnuAlwaysOnTop.Checked)

End Sub


Private Sub mnuGoTo_Click()

'check if in the line with the number specified in flxTODO.RowData from module with name in the first column of the flexgrid
'there still is a todo tag (it was there when the grid was filled but may eventually be gone).
'if there still is the todo tag in that line, will go to that line (if in edit mode)
'if there isn't a todo tag anymore, ask for refreshing the grid (the grid content being outdated)

Const sMsgAskForGridRefresh$ = "Do you want to refresh the TODO list ?"

If IDEMode(VBInstance) = vbext_vm_Design Then
    If flxTODO.Row > 1 Or (flxTODO.Row = 1 And flxTODO.RowData(1) <> ROWDATAEMPTYLINE) Then
        'get the project name, the module name and the line number for the todo tag containd in the current flexgrid line
        'also get the todo string so i can check if the line code still contains the same todo tag (or the grid is oudated)
        
        Dim sModuleName As String
        Dim sToDoText As String
        Dim sProjectName As String
        Dim lCodeLine As Long
        
        With flxTODO
            lCodeLine = .RowData(.Row)
            sProjectName$ = .TextMatrix(.Row, 0)
            sModuleName$ = .TextMatrix(.Row, 1)
            sToDoText$ = .TextMatrix(.Row, 3)
        End With 'flxtodo
        
        Dim sMsgErr As String 'store the string that will be displayed in the msgbox if the grid is outdated
        Dim bDesync As Boolean
        
        Dim objToDo As cToDo
        Set objToDo = ToDoGet(VBInstance, sProjectName$, sModuleName$, lCodeLine, sMsgErr$, bDesync, sToDoTag$)
        
        If objToDo Is Nothing Then
            'tag not found in the specified line (module was changed)
            If bDesync Then
                If vbYes = MsgBox(sMsgErr$ & vbCrLf$ & sMsgAskForGridRefresh$, vbYesNo + vbInformation) Then
                    'refresh the information displayed in the grid
                    RefreshGrid
                End If
            Else
                'we are not in desing mode
                MsgBox sMsgErr$, vbInformation
            End If
        Else
            'we've found a TODO tag in the specified line, check if it is the same with the one in that line when the grid was filled
            'in linia specificata am gasit un TODO, verific daca este vechiul todo
            If sToDoText$ = objToDo.GetTODOText(sToDoTag$) Then
                'the line is the same, I can jump to it in editor
                If Not JumpToLine(VBInstance, sProjectName$, sModuleName$, lCodeLine) Then
                    'there is another todo in the specified line
                     If vbYes = MsgBox("The jump to the specified line failed." & vbCrLf$ & sMsgAskForGridRefresh$, vbYesNo + vbInformation) Then
                        'refresh the information displayed in the grid
                        RefreshGrid
                    End If
                End If
            Else
                'there is another tag in the specified line (other than the one which existed when the grid was filled)
                 If vbYes = MsgBox(sMsgErr$ & vbCrLf$ & sMsgAskForGridRefresh$, vbYesNo + vbInformation) Then
                    'refresh the information displayed in the grid
                    RefreshGrid
                End If
            End If
        End If
        
    Else
        Beep
    End If
Else
    MsgBox "You can jump to the specified line only in design mode.", vbInformation
End If

End Sub


Private Sub mnuHelp_Click()

Help

End Sub

Private Sub mnuOptions_Click()

StartOptionsDialog

End Sub

Private Sub mnuRefresh_Click()

LoadTagsMenu colTagList

FillGridWithTODO flxTODO, VBInstance, sToDoTag$

End Sub


Private Sub mnuTag_Click(Index As Integer)

UncheckAllTagMenuItems

mnuTag(Index).Checked = True

'set current tag from mnu caption
sToDoTag$ = UCase$(mnuTag(Index).Caption)
 
'refresh the grid based on new current tag
RefreshGrid

End Sub

Private Sub tlbToolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
Select Case Button.Key
    
    Case "Exit"
        Connect.Hide
    
    Case "Refresh"
        mnuRefresh_Click
        
    Case "SortA"
        'sorting (ascending)
        With flxTODO
            If .Rows > 2 Then
                'if there are 2 or less lines in the grid, there is no reason for sorting
                .Row = 2
                .RowSel = .Rows - 1
                .Sort = flexSortStringNoCaseAscending
            End If
        End With 'flxtodo
        
    Case "SortD"
        'sorting (descending)
        With flxTODO
            If .Rows > 2 Then
                'if there are 2 or less lines in the grid, there is no reason for sorting
                .Row = 2
                .RowSel = .Rows - 1
                .Sort = flexSortStringNoCaseDescending
            End If
        End With 'flxtodo
        
    Case "Option"
        mnuOptions_Click
    
    Case "Help"
        Call Help
        

End Select

End Sub


Private Sub flxTODO_DblClick()

mnuGoTo_Click

End Sub


Private Sub Form_Load()

'initialize to current tag
sToDoTag$ = sTODOFixedSignature$

'pun the form on the screen in previous position
FormSizeRestore Me, sAppName$, sRegSection$
FormPositionRestore Me, sAppName$, sRegSection$

SetToolbarPosition

'set window on top
SetOnTopState

RefreshWindowState

End Sub


Private Sub SetGridFormat(flxGrid As MSFlexGrid, sCurrentTag As String)

With flxGrid
    .FormatString = "<Project|<Module |<Procedure |<ToDo"
    .ColWidth(0) = 1100 'project
    .ColWidth(1) = 1100 'module
    .ColWidth(2) = 1900 'procedure
    '.ColWidth(3) = 3160 'todo text
    .TextMatrix(0, 3) = sCurrentTag$
End With 'flxgrid

'fill the entire space of the grid with last col
AdjustColWidthToFit flxTODO, flxTODO.Cols - 1

End Sub


'read all TODO tags from all open priojects and put them in the grid
Private Sub FillGridWithTODO(flxGrid As MSFlexGrid, VBinst As vbide.VBE, sToDoTag As String)

Screen.MousePointer = vbHourglass

'add colon to tag string
Dim sToDoLocalTag As String
sToDoLocalTag$ = Trim$(UCase$(sToDoTag$)) & ":" 'a tag must be followed by ":" to be valid

If IDEMode(VBinst) = vbext_vm_Design Then
    'we are in design mode so we can fill the grid
    
    With flxGrid
        .Redraw = False
        .Rows = 2
        .Clear
        .Refresh
    End With 'flxgrid
    
    'set the columns / heading
    SetGridFormat flxGrid, sToDoTag$
    
    Dim lNoOfItems As Long
    lNoOfItems = 0
    
    'iterate through all open projects
    Dim objProject As vbide.VBProject
    For Each objProject In VBinst.VBProjects
    
        'iterate through each code module from the current project
        Dim objModule As vbide.VBComponent
        For Each objModule In objProject.VBComponents
        
            If objModule.Type = vbext_ct_VBForm Or _
               objModule.Type = vbext_ct_ClassModule Or _
               objModule.Type = vbext_ct_StdModule Or _
               objModule.Type = vbext_ct_VBMDIForm Or _
               objModule.Type = vbext_ct_UserControl Or _
               objModule.Type = vbext_ct_ActiveXDesigner Then
            
                'store the total number of lines in the current module and the number of lines in the declaration part of the module
                'so I can fill "DECLARATIONS" as procedure name for lines that belong to the declarations section
                
                Dim lNoDeclLines As Long
                Dim lNoLinesOfCode As Long
                lNoLinesOfCode = objModule.CodeModule.CountOfLines
                lNoDeclLines = objModule.CodeModule.CountOfDeclarationLines
                
                Dim lStartLine As Long
                Dim lEndLine As Long
                Dim lStartCol As Long
                Dim lEndCol As Long
                lStartLine = 1
                lEndLine = -1
                lStartCol = 1
                lEndCol = -1
                Do Until Not objModule.CodeModule.Find("'*" & sToDoLocalTag$, lStartLine, lStartCol, lEndLine, lEndCol, False, False, True)
                                
                    Dim sCodeLine As String
                    Dim sCodeLineUpperCase As String
        
                    sCodeLine$ = Trim$(objModule.CodeModule.Lines(lStartLine, 1))
                    sCodeLineUpperCase$ = UCase$(sCodeLine$)
                    
                    Dim lPosToDo As Long
                    lPosToDo = InStr(1, sCodeLineUpperCase$, sToDoLocalTag$, vbTextCompare)
        
                    If lPosToDo <> 0 Then
                        'a tag was found in the current line
        
                        Dim sProcName As String
                        If lStartLine <= lNoDeclLines Then
                            sProcName$ = "Declarations"
                        Else
        
                            Dim lProcType As Long
                            lProcType = vbext_pk_Proc
        
                            sProcName$ = objModule.CodeModule.ProcOfLine(lStartLine, lProcType)
        
                            'ProcOfLine returns the name and the type of the procedure
                            'based upon this type i can diferentiate between procedure let/set/get (this is not possible only by name)
                            If lProcType <> vbext_pk_Proc Then
                                'it is a property let/get/set
                                Select Case lProcType
                                    Case vbext_pk_Get
                                        sProcName$ = sProcName$ & " (Prop.Get)"
                                    Case vbext_pk_Set
                                        sProcName$ = sProcName$ & " (Prop.Set)"
                                    Case vbext_pk_Let
                                        sProcName$ = sProcName$ & " (Prop.Let)"
                                End Select
                            End If
                        End If
        
                        flxGrid.AddItem objProject.Name & vbTab$ & _
                                        objModule.Name & vbTab$ & _
                                        sProcName$ & vbTab$ & _
                                        Trim$(Mid$(sCodeLine$, lPosToDo + Len(sToDoLocalTag$)))
        
                        'count the no.of items
                        lNoOfItems = lNoOfItems + 1
                        
                        'store the number of the line (in module)
                        flxGrid.RowData(flxGrid.Rows - 1) = lStartLine
                    
                    End If
                    
                    'move to the next line and continue the search until the eof
                    lStartLine = lStartLine + 1
                    lStartCol = 1
                    lEndLine = -1
                    lEndCol = -1
                Loop
                
            End If 'module can contain code ?
        Next objModule
    Next objProject
    
    With flxGrid
        'remove first nonempty line
        If lNoOfItems <> 0 Then
            .RemoveItem 1
            .MergeCells = flexMergeRestrictColumns
            .MergeCol(0) = True
            .MergeCol(1) = True
        Else
            Const MSGNOITEMS$ = " - there are no items to be displayed -"
            .TextMatrix(1, 0) = MSGNOITEMS$
            .TextMatrix(1, 1) = MSGNOITEMS$
            .TextMatrix(1, 2) = MSGNOITEMS$
            .TextMatrix(1, 3) = MSGNOITEMS$
            .MergeCells = flexMergeRestrictRows
            .MergeRow(1) = True
            .RowData(1) = ROWDATAEMPTYLINE
        End If
                
        .Redraw = True
    End With 'flxgrid
Else
    MsgBox "You can refresh the list only in design mode.", vbInformation
End If

Me.Caption = "TODO Add-in (current tag: " & sToDoTag$ & "," & Str$(lNoOfItems) & " items)"

Screen.MousePointer = vbDefault

End Sub


Private Sub Form_Resize()

On Error Resume Next
flxTODO.Move flxTODO.Left, flxTODO.Top, Me.ScaleWidth - flxTODO.Left, Me.ScaleHeight - flxTODO.Top
'fill the entire space of the grid with last col
AdjustColWidthToFit flxTODO, flxTODO.Cols - 1

End Sub


Private Sub RefreshGrid()

FillGridWithTODO flxTODO, VBInstance, sToDoTag$

End Sub


Private Sub Form_Unload(Cancel As Integer)

FormSizeSave Me, sAppName$, sRegSection$
FormPositionSave Me, sAppName$, sRegSection$

End Sub



Private Sub StartOptionsDialog()

Dim bDlgOk As Boolean
bDlgOk = dlgOption.dlgShow(VBInstance, Me)
If bDlgOk Then
    'set the updated alwaysontop state
    SetOnTopState
    SetToolbarPosition
    RefreshWindowState
End If
    
End Sub

Private Sub SetOnTopState()

Dim bOnTopState As Boolean
bOnTopState = CBool(GetSetting(sAppName$, sRegSection$, "AlwaysOnTop", -1))
mnuAlwaysOnTop.Checked = bOnTopState
OnTop Me, bOnTopState

End Sub

Private Sub Help()

frmHelp.dlgShow Me

End Sub

Private Sub LoadTagsMenu(colTags As Collection)

'we must have at least one visible menu item to be able to unload all the rest
mnuTag(0).Visible = True

Dim lItem As Long
On Error Resume Next
For lItem = 1 To MAXTAGSNO
    Unload mnuTag(lItem)
Next lItem
On Error GoTo 0

If Not colTags Is Nothing Then
    For lItem = 1 To colTags.Count
        Load mnuTag(lItem)
        mnuTag(lItem).Caption = UCase$(Trim$(colTags.Item(lItem)))
        mnuTag(lItem).Enabled = True
        mnuTag(lItem).Visible = True
        mnuTag(lItem).Checked = False
    Next lItem
End If

'we have at least one other item so we can hide this one (the collection is always non-empty)
mnuTag(0).Visible = False

mnuTag(1).Checked = True 'set default value
sToDoTag$ = UCase$(Trim$(mnuTag(1).Caption))

End Sub

Private Sub UncheckAllTagMenuItems()

Dim lItem As Long
On Error Resume Next
For lItem = 1 To MAXTAGSNO
    mnuTag(lItem).Checked = False
Next lItem
On Error GoTo 0

End Sub

Private Sub RefreshWindowState()

flxTODO.Visible = False
Me.Refresh

'fill the collection with tags
Set colTagList = GetTagsCol(VBInstance, sTagListSignature$)

'populate popup menu with tags
LoadTagsMenu colTagList

'populate the grid
RefreshGrid

flxTODO.Move flxTODO.Left, flxTODO.Top, Me.ScaleWidth - flxTODO.Left, Me.ScaleHeight - flxTODO.Top
'fill the entire space of the grid with last col
AdjustColWidthToFit flxTODO, flxTODO.Cols - 1

flxTODO.Visible = True

End Sub

Private Sub SetToolbarPosition()

Dim ToolBarWidth As Long

If CBool(GetSetting(sAppName$, sRegSection$, "ToolbarLeft", -1)) Then
    'toolbad aligned to the left
    tlbToolbar1.Align = vbAlignLeft
    ToolBarWidth = tlbToolbar1.ButtonWidth
    If tlbToolbar1.BorderStyle = ccFixedSingle Then
        ToolBarWidth = ToolBarWidth + 2 * Screen.TwipsPerPixelX
    End If
    flxTODO.Move ToolBarWidth, 0
Else
    'toolbar aligned to the top
    tlbToolbar1.Align = vbAlignTop
    ToolBarWidth = tlbToolbar1.ButtonHeight
    If tlbToolbar1.BorderStyle = ccFixedSingle Then
        ToolBarWidth = ToolBarWidth + 2 * Screen.TwipsPerPixelX
    End If
    flxTODO.Move 0, ToolBarWidth
End If

End Sub
