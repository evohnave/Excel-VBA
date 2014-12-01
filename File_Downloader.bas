Option Explicit
 'Declarations
Private Declare Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
                                ByVal szURL As String, _
                                ByVal szFileName As String, _
                                ByVal dwReserved As Long, _
                                ByVal lpfnCB As Long) _
    As Long

Sub DownloadFile(strURL As String, strSaveFileName As String)
     
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                                                         '''
'''  Version 1.0                                                            '''
'''  Date: 20131019                                                         '''
'''                                                                         '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                                                         '''
'''  Subroutine Version                                                     '''
'''  Downloads strSaveFileName from strURL                                  '''
'''                                                                         '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
    Dim lngReturnValue As Long
    Dim lngCaller As Long
    Dim lngReserved As Long
    Dim lngCallBack As Long
    
    'pCaller aka lngCaller can be set to 0
    lngCallBack = 0
    
    'strURL is szURL
    
    'strSaveFileName is szFileName
    
    'dwReserved aka lngReserved must be 0
    lngReserved = 0
    
    'lpfnCB aka lngCallback isn't necessary as the API is asynchronous
    '  and will not return control to VBA until it is done
    lngCallBack = 0
    
    lngReturnValue = URLDownloadToFile(lngCaller, _
                                       strURL, _
                                       strSaveFileName, _
                                       lngReserved, _
                                       lngCallBack)
     
    If lngReturnValue <> 0 Then
        MsgBox Prompt:="File Not Found!", _
               Buttons:=vbOKOnly, _
               Title:="File Not Found!"
    End If

End Sub

Function FileDownloader(strURL As String, strSaveFileName As String) As Boolean
     
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                                                         '''
'''  Version 1.0                                                            '''
'''  Date: 20131019                                                         '''
'''                                                                         '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                                                         '''
'''  Function Version                                                       '''
'''  Downloads strSaveFileName from strURL                                  '''
'''  Returns TRUE for success, FALSE for failure                            '''
'''                                                                         '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim lngReturnValue As Long
    Dim lngCaller As Long
    Dim lngReserved As Long
    Dim lngCallBack As Long
    
    'Default - result was failure
    FileDownloader = False
    
    'pCaller aka lngCaller can be set to 0
    lngCallBack = 0
    
    'strURL is szURL
    
    'strSaveFileName is szFileName
    
    'dwReserved aka lngReserved must be 0
    lngReserved = 0
    
    'lpfnCB aka lngCallback isn't necessary as the API is asynchronous
    '  and will not return control to VBA until it is done
    lngCallBack = 0
    
    lngReturnValue = URLDownloadToFile(lngCaller, _
                                       strURL, _
                                       strSaveFileName, _
                                       lngReserved, _
                                       lngCallBack)
    
    If lngReturnValue = 0 Then FileDownloader = True
     
End Function
