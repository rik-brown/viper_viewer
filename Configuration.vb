'Initializing variables

'Declaring configuration variables
Public columnItemNumber As Integer
Public columnFilePath As Integer
Public columnFileName As Integer
Public columnVideoLink As Integer
Public columnFirstTrigger As Integer
Public columnCause As Integer
Public columnSeverity As Integer
Public columnModule As Integer
Public columnCauseExplanation As Integer
Public columnViewed As Integer
Public columnViewerNotes As Integer
Public columnViewedBy As Integer
Public columnFlag As Integer
Public columnClosed As Integer

Public tomraConnectInstallationNo As Integer

Function Configuration_Init()
    'Set column number here...
    columnItemNumber = 2
    columnFilePath = 3
    columnFileName = 4
    columnVideoLink = 5
    columnFirstTrigger = 6
    columnCause = 7
    columnSeverity = 8
    columnModule = 9
    columnCauseExplanation = 10
    columnViewed = 11
    columnViewerNotes = 12
    columnViewedBy = 13
    columnFlag = 14
    columnClosed = 15
    
    tomraConnectInstallationNo = 28210
    
End Function
