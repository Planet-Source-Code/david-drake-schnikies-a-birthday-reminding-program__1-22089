Attribute VB_Name = "modMain"
Option Explicit

Public Const AGENT = "Peedy"
Public Const AGENTFILE = "Peedy.ACS"

Public mbStartUp As Boolean

Public Sub Main()
    On Error GoTo MainErr
    Dim objXMLParser As clsSchnikiesParser
    
    'First check to see if we are in Startup Mode or Editing Mode
    mbStartUp = Len(ParseCommandLine(Command$, "/")) > 0
    
    If mbStartUp Then
        Set objXMLParser = New clsSchnikiesParser
        Call objXMLParser.ParseXML(App.Path & "\Events.xml")
        Set objXMLParser = Nothing
    Else
        frmEvents.Show
    End If
    Exit Sub

MainErr:
    Set objXMLParser = Nothing
    MsgBox Err.Description, vbCritical + vbOKOnly, App.Title
End Sub

Private Function ParseCommandLine(sCmdLine As String, sOpt1 As String) As String
   'Description:  Takes the given commandline, and option and returns the parameter
   'Parameters:   sCmdLine - String that contains the option and parameter combined
   '              sOption  - Three leter option string (/m:,/p:, etc)
   'Returns:      String with all characters between the last of the option and the next space
   '              sCmdLine with the option and parameter removed.
   Dim sOpt As String
   Dim sTemp As String
   Dim nStartPos As Integer
   Dim bAlreadyTried As Boolean
   
   bAlreadyTried = False
   sOpt = UCase$(sOpt1)

TryAgain:
   If InStr(sCmdLine, sOpt) Then
      ParseCommandLine = ""
      nStartPos = InStr(sCmdLine, sOpt) + Len(sOpt1)
      
      sTemp = Mid$(sCmdLine, nStartPos, 1)
      Do While sTemp <> " " And nStartPos <= Len(sCmdLine)
         ParseCommandLine = ParseCommandLine & sTemp
         nStartPos = nStartPos + 1
         sTemp = Mid$(sCmdLine, nStartPos, 1)
      Loop
   Else
      If Not bAlreadyTried Then
         sOpt = LCase$(sOpt1)
         bAlreadyTried = True
         GoTo TryAgain
      End If
      ParseCommandLine = ""
   End If
End Function

