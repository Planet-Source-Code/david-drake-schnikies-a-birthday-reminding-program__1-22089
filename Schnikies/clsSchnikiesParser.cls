VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSchnikiesParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjXMLFileParser As clsXMLFileStrParser
Attribute mobjXMLFileParser.VB_VarHelpID = -1
Private mdEventDate As Date
Private msEventPerson As String
Private msEventDesc As String
Private mbNotify As Boolean
Private mlLocation As Long
Private mlEventCountDown As Long

Private Sub Class_Initialize()
    Dim x As IAgentCtlRequest
    Set mobjXMLFileParser = New clsXMLFileStrParser
    frmEvents.Agent1.Characters.Load AGENT, AGENTFILE
    Set frmEvents.mobjAgent = frmEvents.Agent1.Characters(AGENT)
    frmEvents.mobjAgent.LanguageID = &H409
End Sub

Private Sub Class_Terminate()
    Set mobjXMLFileParser = Nothing
    If Not frmEvents.mobjAgent Is Nothing Then
        Set frmEvents.mobjCloseMe = frmEvents.mobjAgent.Hide
        Set frmEvents.mobjAgent = Nothing
    End If
End Sub

Public Sub ParseXML(File As String)
    On Error GoTo ParseErr
    Dim sXMLNodes(1 To 1) As String

    'Seed Randomizer
    Rnd (-1)
    Randomize (CLng(Format(Now, "snndd")))
    
    If Not Exists(File) Then
        MsgBox "Unable to find Events.xml in the following location: " & Chr$(13) & File, vbCritical + vbOKOnly, "Schnikies Error!!"
        End
    End If
    
    sXMLNodes(1) = "XML\EVENT"
    
    Call mobjXMLFileParser.ParseXMLFile(File, sXMLNodes)
    Exit Sub

ParseErr:
    MsgBox Err.Description, vbCritical, App.Title
End Sub

Private Sub mobjXMLFileParser_XMLNode(XMLPath As String, XMLContent As String)
    Dim sTemp As String
    
    Select Case XMLPath
        Case "XML\EVENT"
            sTemp = parseXMLVal(XMLContent, "EVENTDATE")
            
            If IsDate(sTemp) Then
                mdEventDate = CDate(sTemp)
                mdEventDate = CDate(CStr(Month(mdEventDate)) & "/" & CStr(Day(mdEventDate)) & "/" & CStr(Year(Date)))
                If DateDiff("d", Date, mdEventDate) < 0 Then mdEventDate = CDate(CStr(Month(mdEventDate)) & "/" & CStr(Day(mdEventDate)) & "/" & CStr(Year(Date) + 1))

                msEventPerson = parseXMLVal(XMLContent, "PERSON")
                msEventDesc = parseXMLVal(XMLContent, "EVENTTYPE")
                mlEventCountDown = Val(parseXMLVal(XMLContent, "WARNING"))
                
                Call NotifyEvent
            End If
    End Select
End Sub

Private Sub NotifyEvent()
    Dim msg As String
    Dim Temp1 As String
    Dim Temp2 As String
    Dim CountDownDate As Date
    Dim Count As Long
    Dim CountDown As Long
        
    If Not IsDate(mdEventDate) Or mdEventDate = #12:00:00 AM# Then GoTo ExitNotify
    
    CountDownDate = DateAdd("d", mlEventCountDown, Date)
    
    Count = DateDiff("d", CountDownDate, mdEventDate)
    CountDown = -1 * mlEventCountDown
    
    If Count <= 0 And Count >= CountDown Then
        With frmEvents.mobjAgent
        .MoveTo Int(CStr((Rnd * ((Screen.Width / Screen.TwipsPerPixelX) - 200)))), Int(CStr((Rnd * ((Screen.Height / Screen.TwipsPerPixelY) - 200))))
        
        Count = mlEventCountDown + Count
        If Count = 0 Then
            msg = "Today is " & msEventPerson & "'s " & msEventDesc & "!" '& Chr$(13) & Chr$(13)
        ElseIf Count > 1 Then
            msg = "There " & msg & "are " & Count & " Days " & "until " & msEventPerson & "'s " & msEventDesc & "."
        Else
            msg = "Tomorrow is " & msEventPerson & "'s " & msEventDesc & "!"
        End If
        
        If Not .Visible Then
            .Show
            .Play "Explain"
            .Speak "Schnikies!"
            .Play "restpose"
        End If
        
        Select Case Int(CStr(Rnd * 9)) + 1
            Case 1
                .Play "Process"
            Case 2
                .Play "Think"
            Case 3
                .Play "Read"
                .Speak msg
                .Play "ReadReturn"
                GoTo AfterSpeek
            Case 4
                .Play "suggest"
            Case 5
                .Play "announce"
            Case 6
                .Play "surprised"
            Case 7
                .Play "search"
            Case 8
                .Play "getattention"
            Case 9
                .Play "Write"
                .Speak msg
                .Play "WriteReturn"
                GoTo AfterSpeek
        End Select
        
        .Speak msg
AfterSpeek:
        
        If Count = 0 Then
            .Play "restpose"
            .Play "Explain"
            Select Case UCase$(msEventDesc)
                Case "BIRTHDAY"
                    Temp1 = "birth-"
                    Temp2 = "day"
                    
                Case "NAME DAY", "NAMEDAY"
                    Temp1 = "name-"
                    Temp2 = "day"
                    
                Case Else
                    Temp1 = ""
                    Temp2 = ""
            End Select
            
            If Len(Temp1) > 0 Then
                msg = "\chr=" + Chr$(34) + "monotone" + Chr$(34)
                .Speak msg + "\\spd=130\\pit=65\Hap-\pit=65\p \pit=73\" & Temp1 & "\pit=65\" & Temp2 & " \spd=55\\pit=87\to \pit=82\" & msEventPerson & "."
                .Speak msg + "\\spd=130\\pit=65\Hap-\pit=65\p \pit=73\" & Temp1 & "\pit=65\" & Temp2 & " \spd=55\\pit=98\to \pit=87\" & msEventPerson & "."
                .Speak msg + "\\spd=130\\pit=65\Hap-\pit=65\p \pit=131\" & Temp1 & "\pit=110\" & Temp2 & ", \spd=50\\pit=87\from \pit=82\me-\pit=73\ee."
                .Play "restpose"
                .Play "blink"
                .Play "blink"
                Select Case Int(CStr(Rnd * 2)) + 1
                    Case 1
                        .Play "pleased"
                        .Play "Explain"
                    Case 2
                        .Play "stoplistening"
                End Select
                
                .Speak msg + "\\spd=130\\pit=117\Hap-\pit=117\p \pit=110\" & Temp1 & "\pit=87\" & Temp2 & " \spd=55\\pit=98\to \pit=87\" & msEventPerson & "."
                Select Case Int(CStr(Rnd * 3)) + 1
                    Case 1
                        .Play "greet"
                    Case 2
                        .Play "Uncertain"
                    Case 3
                        .Play "acknowledge"
                End Select
            End If
        End If
        .Play "Blink"
        .Play "restpose"
        End With
    End If
    
ExitNotify:
    mdEventDate = #12:00:00 AM#
    msEventDesc = ""
    msEventPerson = ""
    mlEventCountDown = 0
End Sub

Private Function Exists(FileName As String) As Boolean
On Error GoTo Done
    If Len(Dir(FileName)) > 0 Then Exists = True
Done:
End Function

