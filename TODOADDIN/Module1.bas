Attribute VB_Name = "Module1"
Option Explicit

'todoaddintaglist: bug,todo_v2,todo_qa

Public Const sTODOAddInTagListTag$ = "TODOAddInTagList:"
Public Const sAppName$ = "TODO VB6-AddIn"

Public Const sTagListSignature$ = "TODOADDINTAGLIST" 'tag for the line containing custom tags list
Public Const sTODOFixedSignature$ = "TODO" 'fixed, uneditable tag, MUST BE IN UPPERCASE here
Public Const MAXTAGSNO As Long = 41

'collection of strings containing all the tags defined in the current project + the fixed one - TODO
'this collection is filled with SET coltaglist=GetTagsCol()
Public colTagList As Collection

Public sToDoTag As String 'current tag

Public Const sRegSection$ = "Mainform"
Public Const ROWDATAEMPTYLINE = -1

'Always on top
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

'The GetSystemMetrics function retrieves various system metrics (widths and heights of display elements) and system configuration settings
'All dimensions retrieved by GetSystemMetrics are in pixels.
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXHSCROLL = 21 'Width, in pixels, of the arrow bitmap on a horizontal scroll bar
Public Const SM_CYHSCROLL = 3  'height, in pixels, of a horizontal scroll bar.
Public Const SM_CXVSCROLL = 2 'Width, in pixels, of a vertical scroll bar
Public Const SM_CYVSCROLL = 20 'height, in pixels, of the arrow bitmap on a vertical scroll bar.


'adjunst the width of the specified flexgrid column so that the entire width of the grid will be filled
'and a space will remain to the right for vertical scrollbar
Public Sub AdjustColWidthToFit(flxGrid As MSFlexGrid, ByVal lCol As Long)

Const ColMinWidth = 100 'minimum column width

'just in case, check if the column is within range
If lCol >= 0 And lCol < flxGrid.Cols Then
    
    'compute the total width of the coluns
    Dim TotalColWidth As Long
    TotalColWidth = 0
    
    Dim lColLocal As Long
    For lColLocal = 0 To flxGrid.Cols - 1
        TotalColWidth = TotalColWidth + flxGrid.ColWidth(lColLocal)
    Next lColLocal

    'contains the total width that can be used for columns (exlcuding the width of the vertical scrollbar)
    Dim UsableGridWidth As Single
    UsableGridWidth = FlexGridScaleWidth(flxGrid)

    Dim NewColWidth As Single
    NewColWidth = UsableGridWidth - (TotalColWidth - flxGrid.ColWidth(lCol))
    
    If UsableGridWidth < TotalColWidth Then
        'make the lcol narrower
        If NewColWidth < ColMinWidth Then
            NewColWidth = ColMinWidth
        End If
    End If
    
    flxGrid.ColWidth(lCol) = NewColWidth
End If

End Sub


'returns the width of the usable interior area of the flexgrid without the space that can be taken by vertical scrollbar
Public Function FlexGridScaleWidth(flxGrid As MSFlexGrid) As Long

Dim UsableGridWidth As Long
UsableGridWidth = flxGrid.Width - Screen.TwipsPerPixelX * GetSystemMetrics(SM_CXVSCROLL) - Screen.TwipsPerPixelX
If flxGrid.BorderStyle = flexBorderSingle Then
    UsableGridWidth = UsableGridWidth - 2 * Screen.TwipsPerPixelX
    If flxGrid.Appearance = flex3D Then
        'daca e 3d, inca 2 pixeli in fiecare parte sint ocupati de border
        UsableGridWidth = UsableGridWidth - 4 * Screen.TwipsPerPixelX
    End If
End If

FlexGridScaleWidth = UsableGridWidth

End Function


'restore the position of a window that was previously saved with SaveFormPosition
Public Sub FormPositionRestore(frmX As Form, AppName As String, Section As String)

Const Margin = 500

With frmX
    Dim topINI As Single
    Dim leftINI As Single
    leftINI = GetSetting(AppName$, Section$, "Left", (Screen.Width - .Width) \ 2)
    topINI = GetSetting(AppName$, Section$, "Top", (Screen.Height - .Height) \ 2)
    
    'keep the window on the screen if the resolution was lowered
    If leftINI > Screen.Width - Margin Then
        leftINI = Screen.Width - Margin 'Margine = a small portion so that I can grab the window with the mouse
    Else
        If leftINI < 0 Then
            leftINI = 0
        End If
    End If
    If topINI > Screen.Height - Margin Then
        topINI = Screen.Height - Margin 'Margine =  a small portion so that I can grab the window with the mouse
    Else
        If topINI < 0 Then
            topINI = 0
        End If
    End If
    If .WindowState = vbNormal Then
        .Move leftINI, topINI
    End If
End With

End Sub


'save the position of the frmX in the INI file
Public Sub FormPositionSave(frmX As Form, AppName As String, Section As String)

With frmX
    SaveSetting AppName$, Section$, "Left", .Left
    SaveSetting AppName$, Section$, "Top", .Top
End With 'frmX

End Sub


Public Sub FormSizeRestore(frmX As Form, AppName$, Section$)

With frmX
    Dim WidthINI As Single
    Dim HeightINI As Single
    WidthINI = GetSetting(AppName$, Section$, "Width", -1)
    HeightINI = GetSetting(AppName$, Section$, "Height", -1)
    
    If WidthINI <> -1 And HeightINI <> -1 Then
        On Error Resume Next
        .Width = WidthINI
        .Height = HeightINI
        On Error GoTo 0
    End If
End With

End Sub


'makes parameter form always on top
Public Sub OnTop(frmFormX As Form, Optional bOnTop As Boolean = True)

Dim wx As Integer
Dim wy As Integer
Dim wcx As Integer
Dim wcy As Integer

Dim tpRatioX As Single
Dim tpRatioY As Single
tpRatioX = Screen.TwipsPerPixelX
tpRatioY = Screen.TwipsPerPixelY

With frmFormX
    wx = .Left \ tpRatioX
    wy = .Top \ tpRatioY
    wcx = .Width \ tpRatioX
    wcy = .Height \ tpRatioY
    
    Dim lOnTopValue As Long
    If bOnTop Then
        lOnTopValue = HWND_TOPMOST
    Else
        lOnTopValue = HWND_NOTOPMOST
    End If
    SetWindowPos .hwnd, lOnTopValue, wx, wy, wcx, wcy, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End With 'frmFormX

End Sub

Public Sub FormSizeSave(frmX As Form, AppName$, Section$)

With frmX
    SaveSetting AppName$, Section$, "Width", .Width
    SaveSetting AppName$, Section$, "Height", .Height
End With

End Sub


'==================================
'Returns the mode the IDE is in:
' vbext_vm_Run = "Run mode"
' vbext_vm_Break = "Break Mode"
' vbext_vm_Design = "Design Mode"
'==================================
Public Function IDEMode(VBinst As vbide.VBE) As Long

Dim lMode As Long
lMode = vbext_vm_Design

'vbinst is nothing when called from a modal window
If VBinst.CommandBars("Run").Controls("End").Enabled Then
    ' The IDE is at least in run mode
    lMode = vbext_vm_Run
    If VBinst.CommandBars("Run").Controls("Break").Enabled = False Then
        ' The IDE is in Break mode
        lMode = vbext_vm_Break
    End If
End If

IDEMode = lMode

End Function


Public Sub NotReady()

MsgBox "The selected option is not implemented (yet)"

End Sub


'returns the TODO tag from the specified line (first if there are more than one in that line)
'renturns nothing if there is no TODO tag in the specified line
'returns an error message when there is no valid todo to return
'returns bDesyncronize = true if we have no todo in the specified line although we are in design mode
Public Function ToDoGet(VBinst As vbide.VBE, sProject As String, sModule As String, lLineNo As Long, sMsgErr As String, bDesyncronize As Boolean, sTag As String) As cToDo

Set ToDoGet = Nothing

If IDEMode(VBinst) = vbext_vm_Design Then
    'we are in design mode, we can continue
    
    Dim objProject As vbide.VBProject
    Set objProject = VBinst.VBProjects(sProject)
    If objProject Is Nothing Then
        'the project from which the specified line was part is no longer loaded
        sMsgErr$ = "Project " & sProject$ & " from which the specified line took part isn't loaded in IDE anymore."
        bDesyncronize = True
    Else
        'the specified project still exists, find the specified module
        Dim objModule As vbide.VBComponent
        Set objModule = objProject.VBComponents(sModule$)
        If objModule Is Nothing Then
            'the module from which the specified line was part is no longer loaded
            sMsgErr$ = "Module " & sModule$ & " doesn't exist anymore in project " & sProject$ & "."
            bDesyncronize = True
        Else
            'the module still exists, read the line with the given number and check if there is the same todo as it was when the grid was loaded
            Dim sFullLine As String
            Dim sLineOfCodeUpperCase As String

            sFullLine$ = objModule.CodeModule.Lines(lLineNo, 1)
            sLineOfCodeUpperCase$ = UCase$(Trim$(sFullLine$))
            
            'search where in the line is the todo tag and get the text after the todo tag to compare it with the one in the grid
            Dim lPosToDo As Long
            lPosToDo = InStr(1, sLineOfCodeUpperCase$, sTag$, vbTextCompare)

            If lPosToDo = 0 Then
                sMsgErr$ = "The specified line doesn't contain a valid TODO tag."
                bDesyncronize = True
            Else
                'we found a todo tag in the current line, return the associated data in an object (cToDo)
                Set ToDoGet = New cToDo
                With ToDoGet
                    .FullTODOLine = sFullLine$
                    .LineNo = lLineNo
                    .ModuleName = sModule$
                    .ProjectName = sProject
                End With 'todoget
                bDesyncronize = False
            End If
        End If 'specified module still in IDE ?
        
        Set objModule = Nothing
        
    End If 'specified module still in IDE ?
    
    Set objProject = Nothing
    
Else
    'we are not in design mode, we cannot get enough data
    sMsgErr$ = "You can get a TODO line only in design mode."
    bDesyncronize = False
End If

End Function


'returns true if the jump to the specified line succeeded
Public Function JumpToLine(VBinst As vbide.VBE, sProject As String, sModule As String, lLineNo As Long) As Boolean

JumpToLine = False

If IDEMode(VBinst) = vbext_vm_Design Then
    'we are in design mode so we can continue
    Dim objProject As vbide.VBProject
    Set objProject = VBinst.VBProjects(sProject)
    If Not objProject Is Nothing Then
        'the specified project still exists, find the module from which the selected line is part
        Dim objModule As vbide.VBComponent
        Set objModule = objProject.VBComponents(sModule$)
        If Not objModule Is Nothing Then
            'we are in the specified project / module, jump to the line with the number lLineNo
            objModule.Activate
            
            With objModule.CodeModule.CodePane
                .Show
                .TopLine = lLineNo
                Call .SetSelection(lLineNo, 1, lLineNo, 255)
            End With
            
            Set objModule = Nothing
            JumpToLine = True
        End If
        Set objModule = Nothing
    End If
    Set objProject = Nothing
End If

End Function

Public Function GetTagsCol(VBinst As vbide.VBE, sToDoTagListTag As String) As Collection

Screen.MousePointer = vbHourglass

If IDEMode(VBinst) <> vbext_vm_Design Then
    Set GetTagsCol = Nothing
    MsgBox "You can refresh the tag list only in design mode.", vbInformation
Else
    'we are in design mode so we can fill the grid
        
    'add colon to tag string
    Dim sLocalTag As String
    sLocalTag$ = UCase$(Trim$(sToDoTagListTag)) & ":" 'a tag must be followed by ":" to be valid
        
    Dim colTagDefLines As Collection 'will store all the lines containing tag definitions list
    Set colTagDefLines = New Collection
        
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
            
                Dim lStartLine As Long
                Dim lEndLine As Long
                Dim lStartCol As Long
                Dim lEndCol As Long
                lStartLine = 1
                lEndLine = -1
                lStartCol = 1
                lEndCol = -1
                Do Until Not objModule.CodeModule.Find("'*" & sLocalTag$, lStartLine, lStartCol, lEndLine, lEndCol, False, False, True)
                                
                    Dim sCodeLine As String
                    Dim sCodeLineUpperCase As String
        
                    sCodeLine$ = Trim$(objModule.CodeModule.Lines(lStartLine, 1))
                    sCodeLineUpperCase$ = UCase$(sCodeLine$)
                    
                    Dim lPosToDo As Long
                    lPosToDo = InStr(1, sCodeLineUpperCase$, sLocalTag$, vbTextCompare)
        
                    If lPosToDo <> 0 Then
                        'a tag was found in the current line
                        colTagDefLines.Add UCase$(Trim$(Mid$(sCodeLine$, lPosToDo + Len(sLocalTag$))))
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
    
    '***************************************************************
    'fill te collection with tags from the collection with tag lines
    '***************************************************************
    Dim colGetTagsCol As Collection
    Set colGetTagsCol = New Collection
    
    'add fixed item
    colGetTagsCol.Add sTODOFixedSignature$
    
    Dim sTempLine As String
    Dim sTagArray() As String
    'we'll add max. MAXTAGSNO tags
    Dim lTagLine As Long
    For lTagLine = 1 To colTagDefLines.Count
        'replace ; with , (";" will be an accepted separator) (other tolerated separator can be added here)
        sTempLine$ = Replace$(colTagDefLines(lTagLine), ";", ",")
        'get rid of all the spaces in the line
        sTempLine$ = Replace$(sTempLine$, " ", vbNullString$)
        'split line
        sTagArray() = Split(sTempLine$, ",", MAXTAGSNO)
        If Not IsEmpty(sTagArray()) Then
            Dim lTagInLine As Long
            For lTagInLine = 0 To UBound(sTagArray())
                'add each tag to collection -tags already in uppercase and no spaces
                
                'the only invalid tag is sToDoTagListTag$ that is reserved for taglists
                If sTagArray$(lTagInLine) <> sToDoTagListTag$ Then
                    AddTagToCollection colGetTagsCol, sTagArray$(lTagInLine)
                End If
                If colGetTagsCol.Count = MAXTAGSNO Then
                    Exit For
                End If
            Next lTagInLine
        End If
        If colGetTagsCol.Count = MAXTAGSNO Then
            Exit For
        End If
    Next lTagLine
    
    'return the result
    Set GetTagsCol = colGetTagsCol
    
    Set colGetTagsCol = Nothing
    Set colTagDefLines = Nothing
End If

Screen.MousePointer = vbDefault

End Function


'add sTag to the cColTags collection
'avoid duplicates (stag is aleady in uppercase and there are no spaces to be trimmed)
Private Sub AddTagToCollection(colTags As Collection, sTag As String)

If Len(sTag$) <> 0 Then
    'try to add only non-empty strings
    Dim bFound As Boolean
    bFound = False
    
    Dim lColIndex As Long
    For lColIndex = 1 To colTags.Count
        If sTag$ = colTags(lColIndex) Then
            bFound = True
            Exit For
        End If
    Next lColIndex
    
    If Not bFound Then
        colTags.Add sTag$
    End If
End If

End Sub
