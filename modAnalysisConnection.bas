Attribute VB_Name = "modAnalysisConnection"
Option Explicit

Public Function GetConnection(ByRef sInput As String) As String
    GetConnection = GetHeaderValue(sInput, "Connection")
End Function
