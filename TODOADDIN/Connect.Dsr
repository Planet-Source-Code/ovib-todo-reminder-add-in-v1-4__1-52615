VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9945
   ClientLeft      =   1215
   ClientTop       =   3765
   ClientWidth     =   6585
   _ExtentX        =   11615
   _ExtentY        =   17542
   _Version        =   393216
   Description     =   "TODO manager Add-In"
   DisplayName     =   "TODO Add-In"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed            As Boolean
Dim mcbMenuCommandBar           As Office.CommandBarControl
Dim mfrmAddIn                   As frmAddIn
Public VBInstance               As vbide.VBE
Attribute VBInstance.VB_VarHelpID = -1
Public WithEvents MenuHandler   As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
Public WithEvents myFileControlEvent As FileControlEvents
Attribute myFileControlEvent.VB_VarHelpID = -1
Sub Hide()
    
On Error Resume Next

FormDisplayed = False
Unload mfrmAddIn
   
End Sub

Sub Show()
  
On Error Resume Next

If mfrmAddIn Is Nothing Then
    Set mfrmAddIn = New frmAddIn
End If

Set mfrmAddIn.VBInstance = VBInstance
Set mfrmAddIn.Connect = Me
FormDisplayed = True

mfrmAddIn.Show

End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = Application
    
    'this will raise events for refreshing the data
    'todo: uncomment next line to intercept file adding / removing from current project
    'Set myFileControlEvent = VBInstance.Events.FileControlEvents(VBInstance.ActiveVBProject)
    
    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods
    'Debug.Print VBInstance.FullName

    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar("TODO Add-In", "TODO Add-In")
        'sink the event
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
  
    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'set this to display the form on connect
            Me.Show
        End If
    End If
  
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload mfrmAddIn
    Set mfrmAddIn = Nothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    
If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
    'set this to display the form on connect
    Me.Show
End If

End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    
    Me.Show

End Sub

Private Function AddToAddInCommandBar(sCaption As String, sDescription As String) As Office.CommandBarControl

'Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
Dim cbMenuCommandBar As Office.CommandBarButton  'command bar object
Dim cbMenu As Object

On Error GoTo AddToAddInCommandBarErr

'see if we can find the Add-Ins menu
Set cbMenu = VBInstance.CommandBars("Add-Ins")
If Not cbMenu Is Nothing Then
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    
    'set the caption
    cbMenuCommandBar.Caption = sCaption$
    
    'todo: uncomment next line when there we have the keyboard shortcut working
    'cbMenuCommandBar.ShortcutText = "Ctrl+Shift+T"
    'todo: add a shortcut (not just the text)
    
    'add the icon for this program to the command bar
    'Clipboard.SetData mfrmAddIn.imgMenuBar.Picture
    'cbMenuCommandBar.PasteFace
    cbMenuCommandBar.DescriptionText = sDescription$
    
    Set AddToAddInCommandBar = cbMenuCommandBar

End If

AddToAddInCommandBarErr:
On Error GoTo 0

End Function

Private Sub myFileControlEvent_AfterAddFile(ByVal VBProject As vbide.VBProject, ByVal FileType As vbide.vbext_FileType, ByVal FileName As String)

FilesInProjectChanged

End Sub


Private Sub myFileControlEvent_AfterCloseFile(ByVal VBProject As vbide.VBProject, ByVal FileType As vbide.vbext_FileType, ByVal FileName As String, ByVal WasDirty As Boolean)

FilesInProjectChanged

End Sub


Private Sub myFileControlEvent_AfterRemoveFile(ByVal VBProject As vbide.VBProject, ByVal FileType As vbide.vbext_FileType, ByVal FileName As String)

FilesInProjectChanged

End Sub


Private Sub FilesInProjectChanged()

If FormDisplayed Then
    Me.Hide
End If

End Sub
