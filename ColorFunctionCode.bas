Attribute VB_Name = "ColorFunctionCode"
Function ColorFunction(rColor As Range, rRange As Range, Optional SUM As Boolean)
'
' ColorFunction either counts the number of cells in a range that match a certain
'   color or sums the contents of cells in a range that match that color
'
' rColor - The cell with the color you're looking for
' rRange - The range you're looking at analyzing
' SUM    - TRUE  - sum the range contents that match the rColor
'        - FALSE - count the number that match rColor
'
Dim rCell As Range
Dim lCol As Long
Dim vResult
lCol = rColor.Interior.ColorIndex
If SUM = True Then
For Each rCell In rRange
If rCell.Interior.ColorIndex = lCol Then
vResult = WorksheetFunction.SUM(rCell, vResult)
End If
Next rCell
Else
For Each rCell In rRange
If rCell.Interior.ColorIndex = lCol Then
vResult = 1 + vResult
End If
Next rCell
End If
ColorFunction = vResult
End Function
