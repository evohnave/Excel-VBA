Sub Refresh_All_Data_Connections()

'-----------------------------------------------------------------------------'
'-----------------------------------------------------------------------------'
'-----                                                                   -----'
'----- Refreshes all data connections and provides notification on the   -----'
'-----   status bar while doing it.                                      -----'
'-----                                                                   -----'
'-----------------------------------------------------------------------------'
'-----------------------------------------------------------------------------'

Dim bBackground As Boolean
Dim iCount As Integer
Dim i As Integer

iCount = ThisWorkbook.Connections.Count

Application.StatusBar = "Refreshing All Connections"

i = 0

For Each objConnection In ThisWorkbook.Connections
        
    i = i + 1
    
    'Notification
    Application.StatusBar = "Refreshing '" & _
        objConnection.Name & "' (" & i & " of " & _
        iCount & ")"
    
    'Get current background-refresh value
    bBackground = objConnection.OLEDBConnection.BackgroundQuery

    'Temporarily disable background-refresh
    objConnection.OLEDBConnection.BackgroundQuery = False

    'Refresh this connection
    objConnection.Refresh

    'Set background-refresh value back to original value
    objConnection.OLEDBConnection.BackgroundQuery = bBackground

Next

Application.StatusBar = False

End Sub
