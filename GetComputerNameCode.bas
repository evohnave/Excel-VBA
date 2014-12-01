Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName _
    Lib "kernel32" Alias "GetComputerNameA" _
    ( _
      ByVal lpBuffer As String, _
      nSize As Long _
     ) _
     As Long
     
Public Function WhoAmI() As String
    
    Dim dwLen As Long
    Dim strString As String
    'Create a buffer
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    'Get the computer name
    GetComputerName strString, dwLen
    'get only the actual data
    WhoAmI = Left(strString, dwLen)

End Function
