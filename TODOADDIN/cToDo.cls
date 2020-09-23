VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToDo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarProjectName As String 'local copy
Private mvarModuleName As String 'local copy
Private mvarLineNo As Long 'local copy
Private mvarFullTODOLine As String 'local copy

Public Property Let FullTODOLine(ByVal vData As String)
Attribute FullTODOLine.VB_Description = "linia TODO completa (fara parse, inclusiv cod daca e vorba de comentariu adaugat dupa linia de cod)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FullTODOLine = 5
    mvarFullTODOLine = vData
End Property


Public Property Get FullTODOLine() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FullTODOLine
    FullTODOLine = mvarFullTODOLine
End Property



Public Property Let LineNo(ByVal vData As Long)
Attribute LineNo.VB_Description = "numarul liniei din cadrul modulului in care se afla TODO curent"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LineNo = 5
    mvarLineNo = vData
End Property


Public Property Get LineNo() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LineNo
    LineNo = mvarLineNo
End Property



Public Property Let ModuleName(ByVal vData As String)
Attribute ModuleName.VB_Description = "Numele modulului in care se afla TODO curent"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ModuleName = 5
    mvarModuleName = vData
End Property


Public Property Get ModuleName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ModuleName
    ModuleName = mvarModuleName
End Property


Public Property Let ProjectName(ByVal vData As String)
Attribute ProjectName.VB_Description = "Numele proiectului in care se afla modulul care contine linia TODO curenta"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProjectName = 5
    mvarProjectName = vData
End Property


Public Property Get ProjectName() As String

ProjectName = mvarProjectName

End Property


Public Function GetTODOText(ByVal sTag As String) As String

Dim sTagLocal As String
sTagLocal$ = UCase$(sTag$ & ":")

If Len(Me.FullTODOLine) = 0 Then
    'we have not enough data filled in the class members
    GetTODOText$ = vbNullString$
Else

    Dim sLineOfCodeUpperCase As String
    sLineOfCodeUpperCase$ = UCase$(Me.FullTODOLine)
    
    'search where in the line is the TODO tag and get the text from that point (not parsed - including the tag)
    Dim lPozToDo As Long
    lPozToDo = InStr(1, sLineOfCodeUpperCase$, sTagLocal$, vbTextCompare)
    
    If lPozToDo <> 0 Then
        'we found a todo tag in the current line, check if the text after the tag is the same as the text in the grid
        GetTODOText = Trim$(Mid$(Me.FullTODOLine, lPozToDo + Len(sTagLocal$)))
    Else
        GetTODOText = vbNullString$
    End If
End If

End Function

