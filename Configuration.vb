'Initializing variables

'Declaring configuration variables
Public columnID As Integer
Public columnFolderPath As Integer
Public columnFileName As Integer
Public columnVideoLink As Integer
Public columnFirstTrigger As Integer
Public columnTriggerID As Integer
Public columnSeverity As Integer
Public columnModule As Integer
Public columnCause As Integer
Public columnViewerNotes As Integer
Public columnViewer As Integer
Public columnFlag As Integer
Public columnReviewed As Integer

Public tomraConnectInstallationNo As Integer

Function Configuration_Init()
    'Set column number here...
    columnID = 1
    columnFolderPath = 2
    columnFileName = 3
    columnVideoLink = 9
    columnFirstTrigger = 11
    columnTriggerID = 5
    columnSeverity = 6
    columnModule = 7
    columnCause = 8
    columnViewerNotes = 13
    columnViewer = 12
    columnFlag = 18
    columnReviewed = 19
    
    tomraConnectInstallationNo = 28210
    
End Function
