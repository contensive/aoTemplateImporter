VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JoomlaClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'
'
Friend Sub ImportJoomlaTemplate(Main As Object, TemplateID As Long, WorkingPathPage As String)
'Friend Sub ImportJoomlaTemplate(Main As MainClass, TemplateID As Long, WorkingPathPage As String)
    '
    Dim CS As Long
    Dim Doc As String
    Dim BasePath As String
    Dim DocBody As String
    Dim DocHead As String
    Dim SourceHost As String
    Dim Pos As Long
    '
    Doc = Main.ReadVirtualFile(WorkingPathPage)
    If Doc = "" Then
        Main.AddUserError ("The Joomla package is empty")
    Else
        'Call ImportJoomlaTemplate_Convert(Doc, BasePath, SourceHost, TemplateID, ManageStyles, BuildVersion)
        Call Main.ClearPageTemplateCache
        Main.AddUserError ("Now Import it")
    End If
    '
End Sub

