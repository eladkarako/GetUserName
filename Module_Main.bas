Attribute VB_Name = "Module_Main"
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Function get_current_username() As String
  On Error Resume Next
  get_current_username = ""
  
  Dim sBuffer As String
  sBuffer = String$(255, 0)
  Call GetUserName(sBuffer, 255)
  
  get_current_username = Left$(sBuffer, InStr(vbNull, sBuffer, vbNullChar, vbBinaryCompare) - 1)
End Function

Public Sub Main()
    WriteStdOut get_current_username()
End Sub

