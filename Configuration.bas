Attribute VB_Name = "Configuration"
'Initializing variables

'Declaring configuration variables
Public columnID As Integer
Public columnFolderPath As Integer
Public columnFileName As Integer

Public columnTriggerID As Integer
Public columnSeverity As Integer
Public columnModule As Integer
Public columnCauseExplanation As Integer

Public columnFirstTrigger As Integer
Public columnViewer As Integer
Public columnViewerNotes As Integer
Public columnIssueID As Integer

Public columnFlag As Integer
Public columnReviewed As Integer

Public tomraConnectInstallationNo As Integer

Function Configuration_Init()
    'Set column number here...
    columnID = 1
    columnFolderPath = 2
    columnFileName = 3
    
    columnTriggerID = 5
    columnSeverity = 6
    columnModule = 7
    columnCauseExplanation = 8
    
    columnFirstTrigger = 11
    columnViewer = 12
    columnViewerNotes = 13
    columnIssueID = 14
    
    columnFlag = 18
    columnReviewed = 19
    
    tomraConnectInstallationNo = 28210
    
End Function
