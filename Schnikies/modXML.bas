Attribute VB_Name = "modXML"
Option Explicit

Public Function parseXMLVal(ByVal strXML As String, ByVal strNode As String) As String
    On Error GoTo ErrHandler
    
    Dim lngStartPos As Long
    Dim lngEndPos As Long
    
    lngStartPos = InStr(1, strXML, "<" & strNode & ">", vbTextCompare)
    lngStartPos = lngStartPos + Len(strNode) + 2
    lngEndPos = InStr(1, strXML, "</" & strNode & ">", vbTextCompare)
    If lngStartPos > 0 And lngStartPos < lngEndPos Then
        parseXMLVal = Mid$(strXML, lngStartPos, lngEndPos - lngStartPos)
    Else
        parseXMLVal = vbNullString
    End If
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[modXML.parseXMLVal]" & Err.Description
End Function

