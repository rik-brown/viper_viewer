Attribute VB_Name = "Viper"
'API for VLC Plugin:
'https://wiki.videolan.org/Documentation:WebPlugin/

'Current Video Variables
Dim currentRow As Integer
Dim nextRow As Integer
Dim prevRow As Integer


Sub Viper_Refresh()

    'Fetching flag alternatives
    Dim flagList As Collection
    
    Set flagList = Helpers.CollectUniques(ActiveSheet.Range("N:N"))
    'Set flagList = Helpers.CollectUniques(ActiveSheet.ListObjects("Triggers").ListColumns(14).DataBodyRange.Select)
    
    ViperForm.issueSelecter.Clear
    For Each flagType In flagList
        ViperForm.issueSelecter.AddItem flagType
    Next
    
    'Inserting data into form
    ViperForm.UserName.Caption = Application.UserName
    ViperForm.IDLabel.Caption = ActiveSheet.Cells(currentRow, Configuration.columnID).Value
    ViperForm.CurrentItemLabel.Caption = ActiveSheet.Cells(currentRow, Configuration.columnID).Value
    ViperForm.SeverityLabel.Caption = "Severity:" & ActiveSheet.Cells(currentRow, Configuration.columnSeverity).Value
    ViperForm.ModuleLabel.Caption = "Module:" & ActiveSheet.Cells(currentRow, Configuration.columnModule).Value
    ViperForm.FirstTriggerLabel.Caption = ActiveSheet.Cells(currentRow, Configuration.columnFirstTrigger).Value
    ViperForm.UrlLabel.Caption = ActiveSheet.Cells(currentRow, Configuration.columnFolderPath).Value
    ViperForm.CauseLabel.Caption = ActiveSheet.Cells(currentRow, Configuration.columnTriggerID).Value + " - " + ActiveSheet.Cells(currentRow, Configuration.columnCauseExplanation).Value
    ViperForm.issueSelecter.Text = ActiveSheet.Cells(currentRow, Configuration.columnIssueID).Value
    
    If ActiveSheet.Cells(currentRow, Configuration.columnReviewed).Value <> "" Then
        ViperForm.ReviewedCheckBox.Value = True
        Else: ViperForm.ReviewedCheckBox.Value = False
    End If
    
    If ActiveSheet.Cells(currentRow, Configuration.columnFlag).Value <> "" Then
        ViperForm.FlagCheckBox.Value = True
        Else: ViperForm.FlagCheckBox.Value = False
    End If
    
    ViperForm.ViewerNotesTextbox.Text = ActiveSheet.Cells(currentRow, Configuration.columnViewerNotes).Value
       
    'Safeguards to avoid crash if values are not initialized...
    'If prevRow > 0 Then
    '    ViperForm.prevButton.Caption = "" & ActiveSheet.Cells(prevRow, Configuration.columnID).Value & " << Previous"
    'End If
    
    'If nextRow > 0 Then
    '    ViperForm.nextButton.Caption = "Next >> " & ActiveSheet.Cells(nextRow, Configuration.columnID).Value
    'End If
    
    'ViperForm.currentRowLabel = currentRow
    
    
    '-- Loading video --
    'If ViperForm.ViperPlayer.Playlist.isPlaying Then
    '    ViperForm.ViperPlayer.Playlist.stop
    'End If
    
    'Getting File Path
    Dim filePath As String
    filePath = "file:" & ActiveSheet.Cells(currentRow, Configuration.columnFolderPath).Value & "/video/" & ActiveSheet.Cells(currentRow, Configuration.columnFileName).Value & ".avi"
    filePath = Replace(filePath, "\", "/")
    
    'Starting Video
    ViperForm.ViperPlayer.Playlist.items.Clear
    ViperForm.ViperPlayer.Playlist.Add (filePath)
    ViperForm.ViperPlayer.Playlist.Play
    
    'ViperForm.isPlaying.Text = ViperForm.ViperPlayer.Playlist.isPlaying
    
    'If ViperForm.ViperPlayer.Playlist.isPlaying Then
    '    ViperForm.isPlaying.Text = filePath
    'End If
        
End Sub

Sub Viper_Save()

    Call Viper_Stop

    If IsNumeric(ViperForm.IDLabel.Caption) Then
        
        ActiveSheet.Cells(currentRow, Configuration.columnViewer).Value = ViperForm.UserName.Caption
        ActiveSheet.Cells(currentRow, Configuration.columnIssueID).Value = ViperForm.issueSelecter.Text
        
        If ViperForm.ReviewedCheckBox.Value = True Then
            ActiveSheet.Cells(currentRow, Configuration.columnReviewed).Value = "x"
        Else
            ActiveSheet.Cells(currentRow, Configuration.columnReviewed).Value = ""
        End If
        
        If ViperForm.FlagCheckBox.Value = True Then
            ActiveSheet.Cells(currentRow, Configuration.columnFlag).Value = "x"
        Else
            ActiveSheet.Cells(currentRow, Configuration.columnFlag).Value = ""
        End If
        
    ActiveSheet.Cells(currentRow, Configuration.columnViewerNotes).Value = ViperForm.ViewerNotesTextbox.Text
    
    End If

End Sub

Sub Viper_FindNext()

    Dim i As Integer
    
    For i = currentRow To 10000
        'Next row
        'i = i + 1
        If ActiveSheet.Rows(i + 1).EntireRow.Hidden = False Then
            nextRow = i + 1
            Call Viper_Refresh
            Exit For
        End If
    Next i
    
End Sub

Sub Viper_FindPrev()
    
    Dim i As Integer

    For i = currentRow - 1 To 1 Step -1
        'Next row
        If ActiveSheet.Rows(i).EntireRow.Hidden = False Then
            prevRow = i
            Call Viper_Refresh
            Exit For
        End If
    Next i
    
End Sub


Sub Viper_Next()    'Normal and hidden funcs can be combined with a showHidden input var.
    Call Viper_Save
    
    currentRow = nextRow
    Call Viper_FindNext
    Call Viper_FindPrev
    
    Call Viper_Refresh
End Sub

Sub Viper_Prev()
    Call Viper_Save
    
    currentRow = prevRow
    If currentRow < 1 Then currentRow = 1
    
    Call Viper_FindNext
    Call Viper_FindPrev
    
    Call Viper_Refresh
End Sub

Sub Viper_NextHidden()
    Call Viper_Save
    
    currentRow = currentRow + 1
    Call Viper_FindNext
    Call Viper_FindPrev
    
    Call Viper_Refresh
End Sub

Sub Viper_PrevHidden()
    Call Viper_Save
    
    currentRow = currentRow - 1
    If currentRow < 1 Then currentRow = 1
    
    Call Viper_FindNext
    Call Viper_FindPrev
    
    Call Viper_Refresh
End Sub

Sub Viper_SetSpeed(speed As Double)
    ViperForm.ViperPlayer.input.Rate = speed
End Sub

Sub Viper_Restart()
    'ViperForm.ViperPlayer.Position (0#)
    ViperForm.ViperPlayer.Playlist.stop
    ViperForm.ViperPlayer.Playlist.Play
End Sub

Sub Viper_TogglePause()
    ViperForm.ViperPlayer.Playlist.togglePause
End Sub

Sub Viper_Stop()
    ViperForm.ViperPlayer.Playlist.stop
End Sub

Sub Viper_ToggleSubtitles()
    If ViperForm.ViperPlayer.video.subtitle > 0 Then
        ViperForm.ViperPlayer.video.subtitle = 0
    Else
       ViperForm.ViperPlayer.video.subtitle = 1
    End If
End Sub

Sub Viper_Fullscreen()
    ViperForm.ViperPlayer.video.fullscreen = True
End Sub


Sub Viper_Open(row As Integer)
    
    'Initializing
    Call Configuration.Configuration_Init
    currentRow = row

    Call Viper_FindNext
    Call Viper_FindPrev
        
    Call Viper_Refresh
    
    ViperForm.Show

End Sub
