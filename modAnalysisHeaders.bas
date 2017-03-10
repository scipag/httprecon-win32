Attribute VB_Name = "modAnalysisHeaders"
Option Explicit

Public Function GetHeaderValue(ByVal sInput As String, ByVal sHeader As String, Optional ByRef bFuzzySearch = True) As String
    Dim sHeaderValue As String
    Dim iHeaderTitle As Integer
    Dim iHeaderStart As Integer
    Dim iHeaderEnd As Integer
    Dim iDelimiterPosition As Integer
    
    iDelimiterPosition = InStr(1, sHeader, ":", vbBinaryCompare)
    
    If (iDelimiterPosition = 0) Then
        sHeader = sHeader & ":"
    End If
    
    If (bFuzzySearch = True) Then
        iHeaderTitle = InStr(1, LCase$(sInput), LCase$(sHeader), vbBinaryCompare)
    Else
        iHeaderTitle = InStr(1, sInput, sHeader, vbBinaryCompare)
    End If
    
    If (iHeaderTitle <> 0) Then
        iHeaderStart = iHeaderTitle + Len(sHeader)
        iHeaderEnd = InStr(iHeaderStart, sInput, vbCrLf, vbBinaryCompare)
        
        If (iHeaderEnd > 0) Then
            iHeaderEnd = iHeaderEnd - iHeaderStart
        End If
        sHeaderValue = Mid$(sInput, iHeaderStart, iHeaderEnd)
    End If
    
    GetHeaderValue = Trim(sHeaderValue)
End Function

Public Function GetHeaderSpace(ByRef sInput As String) As Integer
    Dim iDelimiterPosition As Integer

    iDelimiterPosition = InStr(1, sInput, ":", vbBinaryCompare)

    If (Mid$(sInput, iDelimiterPosition + 1, 1) = " ") Then
        GetHeaderSpace = 1
    Else
        GetHeaderSpace = 0
    End If
End Function

Public Function GetHeaderOrder(ByRef sInput As String, ByRef sIgnore As String) As String
    Dim sIgnoreArray() As String
    Dim iIgnoreCount As Integer
    Dim bFiltered As Boolean
    Dim sHeaderLines() As String
    Dim iHeaderCount As Integer
    Dim i As Integer
    Dim j As Integer
    Dim sHeaderName As String
    Dim iHeaderNameEnd As Integer
    Dim cHeaderNames As Concat
    
    Set cHeaderNames = New Concat
    
    sIgnoreArray = Split(sIgnore, "|", , vbBinaryCompare)
    iIgnoreCount = UBound(sIgnoreArray)
    
    sHeaderLines = Split(sInput, vbCrLf, , vbBinaryCompare)
    iHeaderCount = UBound(sHeaderLines)

    For i = 0 To iHeaderCount
        iHeaderNameEnd = InStr(1, sHeaderLines(i), ":", vbBinaryCompare)
        
        If (iHeaderNameEnd <> 0) Then
            sHeaderName = Mid(sHeaderLines(i), 1, iHeaderNameEnd - 1)
            
            bFiltered = False
            For j = 0 To iIgnoreCount
                If (InStrB(1, LCase$(sHeaderName), LCase$(sIgnoreArray(j)), vbBinaryCompare)) Then
                    bFiltered = True
                    Exit For
                End If
            Next j
        
            If (bFiltered = False) Then
                cHeaderNames.Concat sHeaderName
                
                If (LenB(sHeaderLines(i + 1))) Then
                    cHeaderNames.Concat ","
                End If
            End If
        End If
    Next i
    
    GetHeaderOrder = cHeaderNames.Value
End Function

Public Function GetHeaderCapitalAfterDash(ByRef sInput As String) As Integer
    Dim sHeaders() As String
    Dim iHeaderCount As Integer
    Dim i As Integer
    Dim j As Integer
    Dim iDashPosition As Integer
    Dim iCounterDashes As Integer
    Dim iCounterCapital As Integer
    
    sHeaders = Split(GetHeaderOrder(sInput, vbNullString), ",", , vbBinaryCompare)
    iHeaderCount = UBound(sHeaders)
    
    For i = 0 To iHeaderCount
        iDashPosition = InStr(1, sHeaders(i), "-", vbBinaryCompare)
        
        If (iDashPosition <> 0) Then
            iCounterDashes = iCounterDashes + 1
            For j = 65 To 90
                If (ChrW$(j) = Mid(sHeaders(i), iDashPosition + 1, 1)) Then
                    iCounterCapital = iCounterCapital + 1
                    Exit For
                End If
            Next j
        End If
    Next i
    
    If ((iCounterDashes / 2) < iCounterCapital) Then
        GetHeaderCapitalAfterDash = 1
    Else
        GetHeaderCapitalAfterDash = 0
    End If
End Function
