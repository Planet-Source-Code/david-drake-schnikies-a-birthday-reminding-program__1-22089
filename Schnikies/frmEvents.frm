VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AgentCtl.dll"
Begin VB.Form frmEvents 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Schnikies Event Editor"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3975
   ClipControls    =   0   'False
   Icon            =   "frmEvents.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSelectEvent 
      Caption         =   "Select Event:"
      Height          =   675
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.ComboBox cboSelectEvent 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   180
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   315
      Left            =   2640
      TabIndex        =   13
      Top             =   3120
      Width           =   1275
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   315
      Left            =   60
      TabIndex        =   11
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   315
      Left            =   1320
      TabIndex        =   12
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Frame fraEvent 
      Caption         =   "Event:"
      Height          =   1815
      Left            =   60
      TabIndex        =   2
      Top             =   900
      Width           =   3850
      Begin VB.TextBox txtWarning 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1140
         TabIndex        =   7
         Top             =   1380
         Width           =   555
      End
      Begin VB.TextBox txtDay 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3240
         TabIndex        =   6
         Top             =   1020
         Width           =   435
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         ItemData        =   "frmEvents.frx":0442
         Left            =   1140
         List            =   "frmEvents.frx":046A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1020
         Width           =   2055
      End
      Begin VB.ComboBox cboEvent 
         Height          =   315
         ItemData        =   "frmEvents.frx":04D0
         Left            =   1140
         List            =   "frmEvents.frx":04DD
         TabIndex        =   4
         Top             =   660
         Width           =   2535
      End
      Begin VB.TextBox txtPerson 
         Height          =   315
         Left            =   1140
         TabIndex        =   3
         Top             =   300
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "(days)"
         Height          =   255
         Left            =   1740
         TabIndex        =   18
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Warning:"
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label lblMonth 
         Caption         =   "Date:"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Event:"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   720
         Width           =   915
      End
      Begin VB.Label lblPerson 
         Caption         =   "Person:"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.TextBox txtRecord 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   750
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Record 1/1"
      Top             =   2760
      Width           =   2465
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      Height          =   315
      Left            =   3240
      TabIndex        =   10
      Top             =   2760
      Width           =   675
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<<"
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   2760
      Width           =   675
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3960
      Y1              =   0
      Y2              =   0
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   0
      Top             =   120
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuSettingsCheckNow 
         Caption         =   "&Check Now"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettingsStartup 
         Caption         =   "&Check On Startup"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettingsExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'For Setup Instructions, see Readme.txt

' Constants required by SHGetSpecialFolderLocation API call:
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Type Events
    Person As String
    EventDate As String
    EventType As String
    WarningDays As Integer
End Type

Private Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

Private mobjFS As FileSystemObject
Private mobjEvents() As Events
Private WithEvents mobjXMLParser As clsXMLFileStrParser
Attribute mobjXMLParser.VB_VarHelpID = -1
Public mobjCloseMe As Object
Public mobjAgent As IAgentCtlCharacterEx

Private mlngEventCount As Long
Private mlngEventPointer As Long
Private mstrStartupPath As String

Private Sub Agent1_RequestComplete(ByVal Request As Object)
    Select Case Request
        Case mobjCloseMe
            Agent1.Characters.Unload (AGENT)
            If mbStartUp Then Unload Me
    End Select
End Sub

Private Function ValidateCurrent(DisplayMessage As Boolean) As Boolean
    txtPerson.Text = Trim$(txtPerson.Text)
    cboEvent.Text = Trim$(cboEvent.Text)
    
    If Len(txtPerson.Text) = 0 Then
        If DisplayMessage Then
            Beep
            txtPerson.SetFocus
            MsgBox "Please enter a Person for this event!", vbInformation + vbOKOnly, "Schnikies!"
        End If
        Exit Function
    End If
        
    If Len(cboEvent.Text) = 0 Then
        If DisplayMessage Then
            Beep
            cboEvent.SetFocus
            MsgBox "Please enter the Event Type!", vbInformation + vbOKOnly, "Schnikies!"
        End If
        Exit Function
    End If
    
    If cboMonth.ListIndex = -1 Then
        If DisplayMessage Then
            Beep
            cboMonth.SetFocus
            MsgBox "Please enter the Month for this Event!", vbInformation + vbOKOnly, "Schnikies!"
        End If
        Exit Function
    End If
    
    If Val(txtDay.Text) <= 0 Then
        If DisplayMessage Then
            Beep
            txtDay.SetFocus
            MsgBox "Please enter a valid Day for this Event!", vbInformation + vbOKOnly, "Schnikies!"
        End If
        Exit Function
    End If
        
    Select Case cboMonth.ListIndex
        Case 0, 2, 4, 6, 7, 9, 11
            If Val(txtDay.Text) > 31 Then
                If DisplayMessage Then
                    Beep
                    txtDay.SetFocus
                    MsgBox "Please enter a valid Day for this Event!", vbInformation + vbOKOnly, "Schnikies!"
                End If
                Exit Function
            End If
            
        Case 3, 5, 8, 10
            If Val(txtDay.Text) > 30 Then
                If DisplayMessage Then
                    Beep
                    txtDay.SetFocus
                    MsgBox "Please enter a valid Day for this Event!", vbInformation + vbOKOnly, "Schnikies!"
                End If
                Exit Function
            End If
        
        Case 1
            If Val(txtDay.Text) > 29 Then
                If DisplayMessage Then
                    Beep
                    txtDay.SetFocus
                    MsgBox "Please enter a valid Day for this Event!", vbInformation + vbOKOnly, "Schnikies!"
                End If
                Exit Function
            End If
    End Select
    
    If Val(txtWarning.Text) < 0 Or Val(txtWarning.Text) > 365 Or Not IsNumeric(txtWarning.Text) Then
        If DisplayMessage Then
            Beep
            txtWarning.SetFocus
            MsgBox "Please enter a valid number of Warning days for this Event!", vbInformation + vbOKOnly, "Schnikies!"
        End If
        Exit Function
    End If
    
    ValidateCurrent = True
End Function

Private Sub cboSelectEvent_Click()
    If cboSelectEvent.ListIndex > -1 And cboSelectEvent.ListIndex <> mlngEventPointer Then
        Call SaveCurrent
        mlngEventPointer = cboSelectEvent.ListIndex
        Call DisplayEvent
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim i As Long
    
    If mlngEventCount > 0 Then
        If MsgBox("Are you sure that you want to delete the current record?", vbInformation + vbOKCancel, "Schnikies!") = vbCancel Then Exit Sub
    
        mlngEventCount = mlngEventCount - 1
        
        For i = mlngEventPointer To mlngEventCount - 1
            With mobjEvents(i)
                .Person = mobjEvents(i + 1).Person
                .EventDate = mobjEvents(i + 1).EventDate
                .EventType = mobjEvents(i + 1).EventType
                .WarningDays = mobjEvents(i + 1).WarningDays
            End With
        Next i
        
        If mlngEventPointer > mlngEventCount - 1 Then mlngEventPointer = mlngEventCount - 1
        ReDim Preserve mobjEvents(0 To mlngEventCount - 1)
        
        Call DisplayEvent
        cboSelectEvent.ListIndex = -1
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    If Not ValidateCurrent(True) Then Exit Sub

    Call SaveCurrent
    ReDim Preserve mobjEvents(0 To mlngEventCount)
    mlngEventPointer = mlngEventCount
    mlngEventCount = mlngEventCount + 1
    'Default Values
    With mobjEvents(mlngEventPointer)
        .EventType = "Birthday"
        .WarningDays = 14
    End With
    
    Call DisplayEvent
    cboSelectEvent.ListIndex = -1
    txtPerson.SetFocus
End Sub

Private Sub cmdNext_Click()
    If mlngEventPointer < mlngEventCount - 1 Then
        If Not ValidateCurrent(True) Then Exit Sub
            
        Call SaveCurrent
        mlngEventPointer = mlngEventPointer + 1
        Call DisplayEvent
        cboSelectEvent.ListIndex = -1
    End If
End Sub

Private Sub cmdPrev_Click()
    If mlngEventPointer > 0 Then
        If Not ValidateCurrent(True) Then Exit Sub
        
        Call SaveCurrent
        mlngEventPointer = mlngEventPointer - 1
        Call DisplayEvent
        cboSelectEvent.ListIndex = -1
    End If
End Sub

Private Sub SaveCurrent()
    With mobjEvents(mlngEventPointer)
        .Person = txtPerson.Text
        .EventDate = CStr(cboMonth.ListIndex + 1) & "/" & txtDay.Text & "/00"
        .EventType = cboEvent.Text
        .WarningDays = Val(txtWarning.Text)
    End With
End Sub

Private Sub Form_Load()
    Dim sXMLPaths(0 To 0) As String
    
    If mbStartUp Then Exit Sub
    
    Set mobjXMLParser = New clsXMLFileStrParser
    Set mobjFS = New FileSystemObject
        
    If mobjFS.FileExists(App.Path & "\Events.xml") Then
        sXMLPaths(0) = "XML\EVENT"
        mobjXMLParser.ParseXMLFile App.Path & "\Events.xml", sXMLPaths
    Else
        mlngEventCount = 1
        ReDim mobjEvents(0 To mlngEventCount - 1)
        'Default Values
        With mobjEvents(mlngEventPointer)
            .Person = "Enter Name Here"
            .EventType = "Birthday"
            .EventDate = #1/1/2000#
            .WarningDays = 14
        End With
    End If
    
    Set mobjXMLParser = Nothing
    
    'Fill Select Combo
    For mlngEventPointer = 0 To mlngEventCount - 1
        With mobjEvents(mlngEventPointer)
            cboSelectEvent.AddItem .Person & "'s " & .EventType
        End With
    Next mlngEventPointer
    
    mlngEventPointer = 0
    Call DisplayEvent
    
    'Set Startup Path
    mstrStartupPath = GetStartDir() 'Get Start Menu Path
    If Len(mstrStartupPath) > 0 Then mstrStartupPath = mstrStartupPath & "\StartUp\Schnikies.LNK"
    mnuSettingsStartup.Checked = mobjFS.FileExists(mstrStartupPath)
End Sub

Private Sub DisplayEvent()
    With mobjEvents(mlngEventPointer)
        txtPerson.Text = .Person
        cboEvent.Text = .EventType
        
        If IsDate(.EventDate) Then
            cboMonth.ListIndex = Month(.EventDate) - 1
            txtDay.Text = Day(.EventDate)
        Else
            cboMonth.ListIndex = -1
            txtDay.Text = ""
        End If
        
        txtWarning.Text = CStr(.WarningDays)
    End With
    
    txtRecord.Text = "Record " & CStr(mlngEventPointer + 1) & "/" & CStr(mlngEventCount)
End Sub

Private Sub Form_Terminate()
    Set mobjFS = Nothing
    Erase mobjEvents
    Set mobjXMLParser = Nothing
    Set mobjCloseMe = Nothing
    Set mobjAgent = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim objTS As TextStream
    Dim i As Long
    
    On Error Resume Next
    
    If Not mbStartUp Then
        If Not ValidateCurrent(True) Then
            Cancel = 1
            Exit Sub
        End If
    End If
    
    Call SaveCurrent
    
    With mobjFS
        If .FileExists(App.Path & "\Events.xml") Then
            .CopyFile App.Path & "\Events.xml", App.Path & "\Events.bak", True
            .DeleteFile App.Path & "\Events.xml"
        End If
        
        Set objTS = .OpenTextFile(App.Path & "\Events.xml", ForAppending, True)
    End With
    
    With objTS
        .WriteLine "<XML>"
        
        For i = 0 To mlngEventCount - 1
            .WriteLine "<EVENT><PERSON>" & mobjEvents(i).Person & "</PERSON>" & "<EVENTDATE>" & mobjEvents(i).EventDate & "</EVENTDATE>" & "<EVENTTYPE>" & mobjEvents(i).EventType & "</EVENTTYPE>" & "<WARNING>" & CStr(mobjEvents(i).WarningDays) & "</WARNING></EVENT>"
        Next i
        
        .WriteLine "</XML>"
    End With
    Set objTS = Nothing
End Sub

Private Sub mnuSettingsStartup_Click()
    If mnuSettingsStartup.Checked Then
        'Delete StartUp Icon if Exists
        If mobjFS.FileExists(mstrStartupPath) Then mobjFS.DeleteFile mstrStartupPath, True
        
        'Un-Check Menu Option
        mnuSettingsStartup.Checked = False
    Else
        'Create Startup Icon
        fCreateShellLink "Startup", "Schnikies", App.Path & "\Schnikies.exe", "/S"
        
        'Check Menu Option
        mnuSettingsStartup.Checked = True
    End If
End Sub


Private Sub mnuSettingsCheckNow_Click()
    On Error GoTo MainErr
    Dim objXMLParser As clsSchnikiesParser

    If ValidateCurrent(True) = False Then Exit Sub
    
    'Save Settings
    Call Form_Unload(0)
    
    Set objXMLParser = New clsSchnikiesParser
    Call objXMLParser.ParseXML(App.Path & "\Events.xml")
    Set objXMLParser = Nothing
    Exit Sub

MainErr:
    Set objXMLParser = Nothing
    MsgBox Err.Description, vbCritical + vbOKOnly, App.Title
End Sub

Private Sub mnuSettingsExit_Click()
    Call cmdExit_Click
End Sub

Private Sub mobjXMLParser_XMLNode(XMLPath As String, XMLContent As String)
    Dim i As Long
    Dim sPerson As String
    
    If StrComp(XMLPath, "XML\EVENT", vbTextCompare) = 0 Then
        ReDim Preserve mobjEvents(0 To mlngEventCount)
        
        mlngEventCount = mlngEventCount + 1
        sPerson = parseXMLVal(XMLContent, "Person")
        
        'Sort Alphabetically
        If mlngEventCount > 1 Then
            Call BinarySearch(sPerson)
        Else
            mlngEventPointer = 0
        End If
        
        If mlngEventPointer < mlngEventCount - 1 Then Call ShiftRight
        
        With mobjEvents(mlngEventPointer)
            .Person = parseXMLVal(XMLContent, "Person")
            .EventDate = parseXMLVal(XMLContent, "EventDate")
            .EventType = parseXMLVal(XMLContent, "EventType")
            .WarningDays = Val(parseXMLVal(XMLContent, "Warning"))
        End With
        
    End If
End Sub

Private Sub ShiftRight()
    Dim i As Long
    
    For i = mlngEventCount - 1 To mlngEventPointer Step -1
        If i = 0 Then Exit For
        With mobjEvents(i)
            .Person = mobjEvents(i - 1).Person
            .EventDate = mobjEvents(i - 1).EventDate
            .EventType = mobjEvents(i - 1).EventType
            .WarningDays = mobjEvents(i - 1).WarningDays
        End With
    Next i
End Sub

Private Sub BinarySearch(Person As String)
    Dim lHigh As Long
    Dim lMid As Long
    Dim lLow As Long
    
    lHigh = mlngEventCount - 2
    lLow = 0
    
    Do
        lMid = lLow + ((lHigh - lLow) \ 2)
        Select Case StrComp(Person, mobjEvents(lMid).Person, vbTextCompare)
            Case -1
                lHigh = lMid
                If lLow = lHigh Then
                    mlngEventPointer = lHigh
                    GoTo Limit
                End If
            Case 1
                lLow = lMid + 1
                If lLow > lHigh Then
                    mlngEventPointer = lLow
                    GoTo Limit
                End If
            Case 0
                mlngEventPointer = lMid
                GoTo Limit
        End Select
    Loop
Limit:

End Sub

Private Sub txtDay_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
        Case Is < 48
            KeyAscii = 0
        Case Is > 57
            KeyAscii = 0
        Case Else
            If Len(txtDay.Text) = 2 And txtDay.SelLength = 0 Then KeyAscii = 0
    End Select
End Sub

Private Sub txtWarning_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
        Case Is < 48
            KeyAscii = 0
        Case Is > 57
            KeyAscii = 0
        Case Else
            If Len(txtWarning.Text) = 3 And txtWarning.SelLength = 0 Then KeyAscii = 0
    End Select
End Sub

Private Function GetStartDir() As String
    Dim sPath As String
    Dim IDL As ITEMIDLIST
    Const NOERROR = 0
    Const MAX_LENGTH = 260
    Const USER_START = 2

    If SHGetSpecialFolderLocation(Me.hWnd, USER_START, IDL) = NOERROR Then
        sPath = Space$(MAX_LENGTH)
        
        If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath) Then GetStartDir = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
    End If
End Function
