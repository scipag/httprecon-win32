Attribute VB_Name = "modListViewSort"
Option Explicit

Public Sub ListViewSort(ByRef lListView As ListView, ByRef ColumnHeader As MSComctlLib.ColumnHeader, ByRef iSortOrder As Integer)
    Dim l As Long
    Dim sFormat As String
    Dim sData() As String
    Dim lIndex As Long
    
    On Error Resume Next
    
    With lListView
        lIndex = ColumnHeader.Index - 1
        
        Select Case LCase$(ColumnHeader.Tag)
        Case "number"
            sFormat = String(30, "0") & "." & String(30, "0")
            
            With .ListItems
                If (lIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lIndex)
                            .Tag = .Text & ChrW$(0) & .Tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), _
                                        sFormat)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(.Text), _
                                        sFormat))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            .Tag = .Text & ChrW$(0) & .Tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), _
                                        sFormat)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(.Text), _
                                        sFormat))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                End If
            End With
            
            .SortOrder = iSortOrder
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            
            With .ListItems
                If (lIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lIndex)
                            sData = Split(.Tag, ChrW$(0))
                            .Text = sData(0)
                            .Tag = sData(1)
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            sData = Split(.Tag, ChrW$(0))
                            .Text = sData(0)
                            .Tag = sData(1)
                        End With
                    Next l
                End If
            End With
        
        Case Else
            .SortOrder = iSortOrder
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
        End Select
    End With
End Sub

Private Function InvNumber(ByRef sNumber As String) As String
    Dim i As Integer
    Dim iNumberLength As Integer
    
    iNumberLength = Len(sNumber)
    
    For i = 1 To iNumberLength
        Select Case Mid$(sNumber, i, 1)
        Case "-": Mid$(sNumber, i, 1) = " "
        Case "0": Mid$(sNumber, i, 1) = "9"
        Case "1": Mid$(sNumber, i, 1) = "8"
        Case "2": Mid$(sNumber, i, 1) = "7"
        Case "3": Mid$(sNumber, i, 1) = "6"
        Case "4": Mid$(sNumber, i, 1) = "5"
        Case "5": Mid$(sNumber, i, 1) = "4"
        Case "6": Mid$(sNumber, i, 1) = "3"
        Case "7": Mid$(sNumber, i, 1) = "2"
        Case "8": Mid$(sNumber, i, 1) = "1"
        Case "9": Mid$(sNumber, i, 1) = "0"
        End Select
    Next
    
    InvNumber = sNumber
End Function

