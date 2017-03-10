Attribute VB_Name = "modReportingCsv"
Option Explicit

Public Function GenerateCsvReport(ByRef iHitlistSize As Integer) As String
    Call ChangeStatusBar("Generate CSV Report...")
    GenerateCsvReport = GenerateHitListCsv(frmMain.lsvResults, iHitlistSize, ";")
    Call ChangeStatusBarDone
End Function

Public Function GenerateHitListCsv(ByRef lSource As ListView, ByRef iCount As Integer, Optional ByRef sDelimiter As String = ";") As String
    Dim cResults As Concat
    Dim iListItemCount As Integer
    Dim i As Integer
    
    Set cResults = New Concat
    
    iListItemCount = lSource.ListItems.Count
    
    If (iListItemCount > iCount) Then
        iListItemCount = iCount
    End If
    
    With cResults
        .Concat "Position" & sDelimiter & "Name" & sDelimiter & "Hits" & sDelimiter & "Match" & vbCrLf
        For i = 1 To iListItemCount
             .Concat i & sDelimiter
             .Concat lSource.ListItems(i).ListSubItems(1).Text & sDelimiter
             .Concat lSource.ListItems(i).ListSubItems(2).Text & sDelimiter
             .Concat Round(lSource.ListItems(i).ListSubItems(3).Text, 2) & "%" & vbCrLf
        Next i
    
        GenerateHitListCsv = .Value
    End With
End Function

