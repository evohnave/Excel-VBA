Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Declare Function EmptyClipboard Lib "user32" () As Long
Declare Function CopyEnhMetaFileA Lib "gdi32" (ByVal hENHSrc As Long, ByVal lpszFile As String) As Long
Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hemf As Long) As Long
 
Const CF_ENHMETAFILE As Long = 14
Const cInitialFilename = "Picture1.emf"
Const cFileFilter = "Enhanced Windows Metafile (*.emf), *.emf"
 
Public Sub SaveAsEMF()
    
'------------------------------------------------------------------------------|
' Select something prior to running this macro, which will save the selection  |
' as an EMF File                                                               |
'------------------------------------------------------------------------------|
' Thanks to Rob van Gelder for this code, which he posted on Daily Dose of     |
' Excel on May 05, 2012 at:                                                    |
' http://www.dailydoseofexcel.com/archives/2012/05/05/copy-chart-as-a-picture/ |
'------------------------------------------------------------------------------|
    
    Dim var As Variant, lng As Long
 
    var = Application.GetSaveAsFilename(cInitialFilename, cFileFilter)
    
    If VarType(var) <> vbBoolean Then
        On Error Resume Next
        Selection.Copy
        OpenClipboard 0
        lng = GetClipboardData(CF_ENHMETAFILE)
        lng = CopyEnhMetaFileA(lng, var)
        'EmptyClipboard
        'CloseClipboard
        DeleteEnhMetaFile lng
        On Error GoTo 0
    End If

End Sub

Sub rSaveAsEMF(control As IRibbonControl)

    SaveAsEMF

End Sub
