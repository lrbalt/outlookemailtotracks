VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TodoForm 
   Caption         =   "Tracks Todo"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   OleObjectBlob   =   "TodoForm10.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TodoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' greg jarman   20080821   v1.0     initial version
'
'

' the URL of the tracks installation, required.
Const sURL = "http://your.tracks.host/tracks/"
 
' set username and password here, required.
Const sUsername = "userid"
Const sPassword = "password"

' proxy server address and port number, in the form proxy.server.com:1234. Set to "" for none
Const sProxy = ""

' set this to true if you want to refresh the projects and contexts each time a new todo is created.
' otherwise the data is gathered the first time a todo is created after outlook is opened.
Const Update_Projects_And_Contexts_Each_Time = True

' internal variables
Private Type Context
  Name As String
  id As Integer
  Position As Integer
End Type

Private Type Project
  Name As String
  id As Integer
  Position As Integer
  State As String
End Type

Dim Projects() As Project
Dim ActiveProjects() As Project
Dim Projects_Length As Integer
Public Projects_Loaded As Boolean

Dim Contexts() As Context
Dim Contexts_Length As Integer
Public Contexts_Loaded As Boolean

Const HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = 0


Private Function CreateWinHttpRequest() As Object
  Dim WinHttpRequest As Object
  
  Set WinHttpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
  
  Set CreateWinHttpRequest = WinHttpRequest
End Function

Private Sub ConfigureWinHttpRequest(WinHttpRequest As Variant)
  WinHttpRequest.SetCredentials sUsername, sPassword, HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
  WinHttpRequest.SetRequestHeader "Content-type", "text/xml"
  If Len(sProxy) > 0 Then
    WinHttpRequest.SetProxy 2, sProxy, ""
  End If
  'WinHttpRequest.SetClientCertificate ("")
End Sub

' Download the contexts from the Tracks server and add them to the Projects array
Private Sub DownloadContexts()
  Dim WinHttpRequest As Object
  Dim XMLDoc As Object
  Dim oRoot As Object
  Dim oContext As Object
  Dim oChild As Object
  
  Set WinHttpRequest = CreateWinHttpRequest()
  WinHttpRequest.Open "GET", sURL & "contexts.xml", False
  ConfigureWinHttpRequest WinHttpRequest
  WinHttpRequest.Send
    
  If Not WinHttpRequestSucceeded(WinHttpRequest) Then
    MsgBox "DownloadContexts Failed: " & WinHttpRequest.StatusText
    Exit Sub
  End If
    
    Set XMLDoc = CreateObject("MSXML2.DOMDocument")
    XMLDoc.validateOnParse = False
    XMLDoc.loadXML (WinHttpRequest.ResponseText)
    Set oRoot = XMLDoc.documentElement
    Contexts_Length = oRoot.childNodes.Length
    ReDim Contexts(0 To Contexts_Length - 1) As Context
    
    Dim sName
    Dim iPosition
    Dim iId
    For Each oContext In oRoot.childNodes
      For Each oChild In oContext.childNodes
        If oChild.nodeName = "name" Then
          sName = oChild.nodeTypedValue
        ElseIf oChild.nodeName = "position" Then
          iPosition = Val(oChild.nodeTypedValue)
        ElseIf oChild.nodeName = "id" Then
          iId = Val(oChild.nodeTypedValue)
        End If
      Next oChild
      Contexts(iPosition - 1).Name = sName
      Contexts(iPosition - 1).Position = iPosition
      Contexts(iPosition - 1).id = iId
    Next oContext
    
    Contexts_Loaded = True
    
    Set WinHttpRequest = Nothing
    Set oRoot = Nothing
    Set XMLDoc = Nothing
End Sub

' Download the projects from the Tracks server and add them to the Projects array
Private Sub DownloadProjects()
  Dim WinHttpRequest As Object
  Dim XMLDoc As Object
  Dim oRoot As Object
  Dim oProject As Object
  Dim oChild As Object
  
    Set WinHttpRequest = CreateWinHttpRequest()
    WinHttpRequest.Open "GET", sURL & "projects.xml", False
    ConfigureWinHttpRequest WinHttpRequest
    WinHttpRequest.Send
    
    If Not WinHttpRequestSucceeded(WinHttpRequest) Then
      MsgBox "DownloadProjects Failed: " & WinHttpRequest.StatusText
      Exit Sub
    End If
  
    Set XMLDoc = CreateObject("MSXML2.DOMDocument")
    XMLDoc.validateOnParse = False
    XMLDoc.loadXML (WinHttpRequest.ResponseText)
    Set oRoot = XMLDoc.documentElement
    Projects_Length = oRoot.childNodes.Length
    ReDim Projects(0 To Projects_Length - 1) As Project
    
    Dim sName
    Dim iPosition
    Dim iId
    Dim sState
    For Each oProject In oRoot.childNodes
      For Each oChild In oProject.childNodes
        If oChild.nodeName = "name" Then
          sName = oChild.nodeTypedValue
        ElseIf oChild.nodeName = "position" Then
          iPosition = Val(oChild.nodeTypedValue)
        ElseIf oChild.nodeName = "id" Then
          iId = Val(oChild.nodeTypedValue)
        ElseIf oChild.nodeName = "state" Then
          sState = oChild.nodeTypedValue
        End If
      Next oChild
      Projects(iPosition - 1).Name = sName
      Projects(iPosition - 1).Position = iPosition
      Projects(iPosition - 1).id = iId
      Projects(iPosition - 1).State = sState
    Next oProject
    
    Projects_Loaded = True
    
    Set WinHttpRequest = Nothing
    Set oRoot = Nothing
    Set XMLDoc = Nothing
End Sub



Private Sub PopulateProjectListBox()
  Dim old_index As Integer
  
  old_index = TodoForm.ProjectListBox.ListIndex
  TodoForm.ProjectListBox.Clear
  TodoForm.ProjectListBox.Style = fmStyleDropDownList
    
  Dim i
  Dim count As Integer
  ReDim ActiveProjects(0 To Projects_Length - 1)
  count = 0
  
  For i = 0 To Projects_Length - 1
    If Projects(i).State = "active" Then
      TodoForm.ProjectListBox.AddItem Projects(i).Name
      ActiveProjects(count) = Projects(i)
      count = count + 1
    End If
  Next i
  
  TodoForm.ProjectListBox.ListIndex = old_index
End Sub

Private Sub PopulateContextListBox()
  Dim old_index As Integer
  old_index = TodoForm.ContextListBox.ListIndex
    
  TodoForm.ContextListBox.Clear
  TodoForm.ContextListBox.Style = fmStyleDropDownList
    
  Dim i
  For i = 0 To Contexts_Length - 1
    TodoForm.ContextListBox.AddItem Contexts(i).Name
  Next i
  
  TodoForm.ContextListBox.ListIndex = old_index
End Sub

Private Function WinHttpRequestSucceeded(WinHttpRequest As Variant)
  WinHttpRequestSucceeded = (WinHttpRequest.Status >= 200) And (WinHttpRequest.Status <= 299)
End Function

Private Function HTMLEncode(ByVal Text As String, Optional HardSpaces As Boolean = False) As String
  Dim i As Integer
  Dim ch As String
  Dim NewString As String


  For i = 1 To Len(Text)
    ch = Mid$(Text, i, 1)
    Select Case ch
          Case " "
              If HardSpaces Then ch = "&nbsp;"
          Case """"
              ch = "&quot;"
          Case "&"
                ch = "&amp;"
          Case "<"
              ch = "&lt;"
          Case ">"
              ch = "&gt;"
          Case " " To "~"
              ' Not one we already processed but
              ' but in the normal display range
          Case Else
              ch = "&#" & Asc(ch) & ";"
      End Select
      NewString = NewString & ch
  Next
  HTMLEncode = NewString
End Function

Private Sub CreateTodo(Description As String, Notes As String, ContextId As Integer, ProjectId As Integer)
  Dim WinHttpRequest As Object
  Dim sData As String
  Dim aBody() As Byte
    
  Set WinHttpRequest = CreateWinHttpRequest()
  WinHttpRequest.Open "POST", sURL & "todos.xml", False
  ConfigureWinHttpRequest WinHttpRequest
    
  sData = "<todo><description>" + Description + "</description>"
  sData = sData & "<notes>" & Notes & "</notes>"
  sData = sData & "<context_id>" & Str(ContextId) & "</context_id>"
    
  If ProjectId > 0 Then
    sData = sData & "<project_id>" & Str(ProjectId) & "</project_id>"
  End If

  sData = sData & "</todo>" & vbCrLf
    
  aBody = StrConv(sData, vbFromUnicode)
  WinHttpRequest.Send CByte(aBody)
    
  If Not WinHttpRequestSucceeded(WinHttpRequest) Then
    MsgBox "CreateTodo Failed: " & WinHttpRequest.StatusText
  End If
    
  Set WinHttpRequest = Nothing
End Sub

Private Sub AddActionButton_Click()
  If Len(DescriptionTextBox.Text) = 0 Then
    MsgBox "Must have a description!", vbExclamation
  ElseIf ContextListBox.ListIndex = -1 Then
    MsgBox "Must have a context!", vbExclamation
  Else
    Dim ContextId As Integer
    Dim ProjectId As Integer
    ContextId = Contexts(ContextListBox.ListIndex).id
    ProjectId = 0
    If ProjectListBox.ListIndex > 0 Then
      ProjectId = ActiveProjects(ProjectListBox.ListIndex).id
    End If
    
    CreateTodo DescriptionTextBox.Text, NotesTextBox.Text, ContextId, ProjectId
    FormReset
    TodoForm.Hide
  End If
End Sub

Private Sub FormReset()
  DescriptionTextBox.Text = ""
  NotesTextBox.Text = ""
End Sub

Private Sub CancelButton_Click()
  FormReset
  TodoForm.Hide
End Sub


Private Sub ClearProject_Click()
  ProjectListBox.ListIndex = -1
End Sub

Private Sub UserForm_Initialize()
  Contexts_Loaded = False
  Projects_Loaded = False
End Sub

Private Sub UserForm_Activate()
  If (Projects_Loaded = False) Or Update_Projects_And_Contexts_Each_Time Then
    DownloadProjects
    PopulateProjectListBox
    Projects_Loaded = True
  End If
    
  If (Contexts_Loaded = False) Or Update_Projects_And_Contexts_Each_Time Then
    DownloadContexts
    PopulateContextListBox
    Contexts_Loaded = True
  End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  ' Prevents use of the Close button, so we don't have to download the Projects and Contexts again
  TodoForm.Hide
  FormReset
  Cancel = True
End Sub
