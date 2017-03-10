Attribute VB_Name = "modComboboxAutocomplete"
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const CB_FINDSTRING As Long = &H14C

Public Sub ComboAutoComplete(ByRef SourceCtl As VB.ComboBox, ByRef KeyAscii As Integer, ByRef LeftOffPos As Long)
    Dim iStart As Long
    Dim sSearchKey As String
    
    With SourceCtl
        Select Case ChrW$(KeyAscii)
          Case vbBack
          Case Else
            If ChrW$(KeyAscii) <> vbBack Then
              .SelText = ChrW$(KeyAscii)
              
              iStart = .SelStart
              
              If LeftOffPos <> 0 Then
                .SelStart = LeftOffPos
                iStart = LeftOffPos
              End If
              
              sSearchKey = CStr(Left$(.Text, iStart))
              .ListIndex = SendMessage(.hwnd, CB_FINDSTRING, -1, _
                  ByVal CStr(Left$(.Text, iStart)))
              
              If .ListIndex = -1 Then
                LeftOffPos = Len(sSearchKey)
              End If
              
              .SelStart = iStart
              .SelLength = Len(.Text)
              LeftOffPos = 0
              
              KeyAscii = 0
            End If
        End Select
    End With
End Sub

