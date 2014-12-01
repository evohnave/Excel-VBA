Function FileLister(strFldr As String, ByRef FileList() As Variant) As Variant

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''                                                                           ''
''  This requires a reference to                                             ''
''    Windows Script Host Object Model                                       ''
''    C:\Windows\SysWOW64\wshom.ocx                                          ''
''    Not necessarily in this folder everywhere                              ''
''                                                                           ''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''                                                                           ''
''  This function will pass back a nx1 array FileList() of filenames in the  ''
''    folder strFolder.  If there are no files in the folder it will return  ''
''    a value of FALSE.  Any errors also return FALSE.                       ''
''                                                                           ''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim fso As FileSystemObject
Dim fldr As Folder
Dim i As Long
Dim fl As File

'Default value for function is TRUE
FileLister = True

'Create object
Set fso = CreateObject("Scripting.FileSystemObject")

'Generic Error Handling
On Error GoTo Err_Handler

'Test that the target folder exists
If Not fso.FolderExists(strFldr) Then GoTo Fn_False

'OK, the folder exists
Set fldr = fso.GetFolder(strFldr)

'Does it have any files?
If fldr.Files.Count = 0 Then GoTo Fn_False

'OK, files exist.  Let's list 'em!
ReDim FileList(1 To fldr.Files.Count)

i = 0
For Each fl In fldr.Files
    i = i + 1
    FileList(i) = fl.Name
Next fl

GoTo Fn_Exit

Err_Handler:
'Uncomment the msgbox if you want to figure out a problem
'MsgBox Prompt:="Error  = " & Err.Number & ", " & Err.Description, _
'       Buttons:=vbOKOnly, _
'       Title:="Error Generated!"

'If you get here then there was a problem
Fn_False:
FileLister = False

Fn_Exit:
Set fl = Nothing
Set fldr = Nothing
Set fso = Nothing

End Function

Sub UseFileLister()

'Example showing how to use FileLister

Dim strFldr As String
Dim result As Variant
Dim FileList() As Variant

strFldr = InputBox(Prompt:="Please enter folder name", _
                   Title:="Give me a folder")

result = FileLister(strFldr, FileList)

For i = 1 To UBound(FileList)
Debug.Print i & "=" & FileList(i)
Next i

End Sub

Function ReferenceLister(ByRef RefList() As Variant) As Variant

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''                                                                           ''
''  This requires the VBE to be trusted                                      ''
''    This is done in the Trust Center                                       ''
''                                                                           ''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''                                                                           ''
''  This function will pass back a nx3 array RefList() of references in the  ''
''    VBE Project.  If there are no references it will return a value of     ''
''    FALSE.  Any errors also return FALSE.                                  ''
''                                                                           ''
''  RefList(,1) is .Name                                                     ''
''  RefList(,2) is .FullPath                                                 ''
''  RefList(,3) is .Description                                              ''
''                                                                           ''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim i As Long

On Error GoTo Err_Handler

'Default value is TRUE
ReferenceLister = True

With Application.VBE.ActiveVBProject

If Not (.References.Count > 0) Then GoTo Fn_False

'OK, at least 1 reference
ReDim RefList(1 To .References.Count, 1 To 3)

For i = 1 To .References.Count
    RefList(i, 1) = .References.Item(i).Name
    RefList(i, 2) = .References.Item(i).FullPath
    RefList(i, 3) = .References.Item(i).Description
Next i

End With

GoTo Fn_Exit

Err_Handler:
'Uncomment the msgbox if you want to figure out a problem
'MsgBox Prompt:="Error  = " & Err.Number & ", " & Err.Description, _
'       Buttons:=vbOKOnly, _
'       Title:="Error Generated!"

Fn_False:
ReferenceLister = False

Fn_Exit:

End Function

Sub UseRefLister()

Dim RefList() As Variant
Dim result As Boolean
Dim i As Long

result = ReferenceLister(RefList)

If result Then

For i = 1 To UBound(RefList, 1)
Debug.Print RefList(i, 1), RefList(i, 2), RefList(i, 3)
Next i

End If

End Sub

Function CheckRefDesc(strRefDesc As String) As Boolean

'strRef is the description of a reference
'  For example, Windows Script Host Object Model

Dim RefList() As Variant
Dim i As Long

'Default value is FALSE
CheckRefDesc = False

If Not ReferenceLister(RefList) Then Exit Function
'If you get here then there is at least one reference

For i = 1 To UBound(RefList, 1)
    If RefList(i, 3) = strRefDesc Then
        CheckRefDesc = True
        Exit Function
    End If
Next i

End Function
