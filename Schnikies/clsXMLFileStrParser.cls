VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXMLFileStrParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event XMLNode(XMLPath As String, XMLContent As String)

Public Function ParseXMLFile(FilePath As String, XMLPaths() As String) As Boolean
    On Error GoTo ParseXMLFile_Err
    Dim lMaxPaths As Long
    Dim lMinPaths As Long
    Dim lXMLPathStart As Long
    Dim lXMLPathIndex As Long
    Dim i As Long
    Dim j As Long
    Dim XMLPath As String
    Dim XMLRoot As String
    Dim XMLContent As String
    Dim XMLString As String
    Dim XMLSubString As String
    Dim XMLPreviousNode As String
    Dim XMLCurrentNode As String
    Dim sTempNode As String
    Dim objFS As FileSystemObject
    Dim objTS As TextStream
    
    lMinPaths = LBound(XMLPaths)
    lMaxPaths = UBound(XMLPaths)
    If lMinPaths < 0 Or lMaxPaths < 0 Then Err.Raise 1, "[clsXMLFileStrParser][ParseXMLFile]", "Invalid XML Path Array Encountered."

    Set objFS = New FileSystemObject
    
    'Validate Parameters
    If Not objFS.FileExists(FilePath) Then Err.Raise 1, "[clsXMLFileStrParser][ParseXMLFile]", "File '" & FilePath & "' does not exists."
    If Not IsArray(XMLPaths) Then Err.Raise 1, "[clsXMLFileStrParser][ParseXMLFile]", "Invalid XML Path Array Encountered."
    
    Set objTS = objFS.OpenTextFile(FilePath, ForReading)
    XMLContent = objTS.ReadAll
    XMLContent = Replace$(XMLContent, vbCrLf, "")
    
    For j = lMinPaths To lMaxPaths
        XMLPath = XMLPaths(j)
        
        'Validate path
        If Len(Trim$(XMLPath)) = 0 Then Err.Raise 1, "[clsXMLFileStrParser][ParseXMLFile]", "Invalid XML Path."
        If Left$(XMLPath, 1) = "\" Then Err.Raise 1, "[clsXMLFileStrParser][ParseXMLFile]", "Invalid XML Path Encountered."
        If Right$(XMLPath, 1) = "\" Then Err.Raise 1, "[clsXMLFileStrParser][ParseXMLFile]", "Invalid XML Path Encountered."
        If InStr(1, XMLPath, "\\") > 0 Then Err.Raise 1, "[clsXMLFileStrParser][ParseXMLFile]", "Invalid XML Path Encountered."
    
        lXMLPathStart = 1
            
        'Get Root Node
        lXMLPathIndex = InStr(lXMLPathStart, XMLPath, "\", vbTextCompare)
        If lXMLPathIndex = 0 Then
            XMLRoot = XMLPath
            RaiseEvent XMLNode(XMLPath, parseXMLVal(XMLContent, XMLPath))
            GoTo NextPath
        End If
        
        sTempNode = Mid$(XMLPath, lXMLPathStart, lXMLPathIndex - lXMLPathStart)
        If Len(Trim$(sTempNode)) = 0 Then Err.Raise 1, "[clsXMLFileStrParser][ParseXMLFile]", "Invalid XML Path Encountered."
    
        'On first Path, handle Root node specially
        If j = lMinPaths Then
            XMLRoot = sTempNode
            
        'Validate that all paths have the same root
        ElseIf StrComp(XMLRoot, sTempNode, vbTextCompare) <> 0 Then
            Err.Raise 1, "[clsXMLFileStrParser][ParseXMLFile]", "XML Paths Contain different Roots."
        End If
        
        'Parse Child Nodes
        lXMLPathStart = lXMLPathIndex + 1
        lXMLPathIndex = InStr(lXMLPathStart, XMLPath, "\", vbTextCompare)
        If lXMLPathIndex = 0 Then
            If lXMLPathStart < Len(XMLPath) Then
                lXMLPathIndex = Len(XMLPath) + 1
                sTempNode = Mid$(XMLPath, lXMLPathStart, lXMLPathIndex - lXMLPathStart)
            Else
                sTempNode = ""
            End If
        Else
            sTempNode = Mid$(XMLPath, lXMLPathStart, lXMLPathIndex - lXMLPathStart)
        End If
        
        If Len(sTempNode) > 0 Then
            i = 0
            
            Do
                If Len(Trim$(sTempNode)) = 0 Then Err.Raise 1, "[clsXMLFileStrParser][ParseXMLFile]", "Invalid XML Path Encountered."
                i = i + 1
                
                Select Case i
                    Case 1
                        XMLPreviousNode = XMLRoot
                        XMLCurrentNode = sTempNode
                        XMLString = parseXMLVal(XMLContent, XMLRoot)
                        
                    Case 2
                        XMLString = parseXMLVal(XMLContent, XMLCurrentNode)
                        XMLPreviousNode = XMLCurrentNode
                        XMLCurrentNode = sTempNode
                    
                    Case Is > 2
                        XMLString = parseXMLVal(XMLString, XMLCurrentNode)
                        XMLPreviousNode = XMLCurrentNode
                        XMLCurrentNode = sTempNode
                End Select
                
                
                'Get Next Node
                lXMLPathStart = lXMLPathIndex + 1
                lXMLPathIndex = InStr(lXMLPathStart, XMLPath, "\", vbTextCompare)
                If lXMLPathIndex = 0 Then
                    If lXMLPathStart < Len(XMLPath) Then
                        lXMLPathIndex = Len(XMLPath) + 1
                        sTempNode = Mid$(XMLPath, lXMLPathStart, lXMLPathIndex - lXMLPathStart)
                    Else
                        sTempNode = ""
                    End If
                Else
                    sTempNode = Mid$(XMLPath, lXMLPathStart, lXMLPathIndex - lXMLPathStart)
                End If
                
            Loop While Len(sTempNode) > 0
            
            
            While Len(XMLString) > 0
                If GetSubString(XMLString, XMLCurrentNode, sTempNode) Then _
                    RaiseEvent XMLNode(XMLPath, sTempNode)
            Wend
        End If

NextPath:
    Next j
        
ParseXMLFile_Exit:
    Set objTS = Nothing
    Set objFS = Nothing
    Exit Function
    
ParseXMLFile_Err:
    Set objTS = Nothing
    Set objFS = Nothing
    Err.Raise Err.Number, Err.Source, "[clsXMLFileStrParser][ParseXMLFile]" & Err.Description
End Function

Private Function GetSubString(ByRef XML As String, Node As String, ByRef Contents As String) As Boolean
    On Error GoTo GetSubString_Err
    Dim lStartIndex As Long
    Dim lStopIndex As Long
    Dim sJob As String
    
    lStartIndex = InStr(1, XML, "<" & Node & ">", vbTextCompare)
    
Retry:
    If lStartIndex <= 0 Then
        Contents = ""
        XML = ""
        Exit Function
    End If
    
    lStopIndex = InStr(lStartIndex + Len(Node) + 2, XML, "</" & Node & ">", vbTextCompare)
    If lStopIndex <= lStartIndex Then
        Contents = ""
        XML = ""
        Exit Function
    End If
    
    Contents = Mid$(XML, lStartIndex + Len(Node) + 2, lStopIndex - lStartIndex - Len(Node) - 2)
    XML = Trim$(Mid$(XML, lStopIndex + Len(Node) + 3, Len(XML) - lStopIndex - Len(Node) - 2))
    GetSubString = True
    
GetSubString_Exit:
    Exit Function
    
GetSubString_Err:
    Err.Raise Err.Number, "[GetSubString]", "[GetSubString]" & Err.Description
    Contents = ""
    Resume GetSubString_Exit
End Function

