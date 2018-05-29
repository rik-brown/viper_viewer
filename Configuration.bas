Attribute VB_Name = "Configuration"
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
    columnItemNumber = 1
    columnFilePath = 2
    columnFileName = 3
    columnVideoLink = 9
    columnFirstTrigger = 11
    columnCause = 5
    columnSeverity = 6
    columnModule = 7
    columnCauseExplanation = 8
    columnViewed = 17
    columnViewerNotes = 13
    columnViewedBy = 12
    columnFlag = 18
    columnClosed = 19
    
    tomraConnectInstallationNo = 28210
    
End Function
