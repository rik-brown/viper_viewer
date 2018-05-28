
Private Sub CancelExitButton_Click()
    ViperForm.Hide
End Sub


Private Sub ConsumerSessonLink_Click()
    ActiveWorkbook.FollowHyperlink Address:=("https://www1.tomraconnect.com/page/installation/consumersession/list?installation=" & Configuration.tomraConnectInstallationNo), NewWindow:=True
    'ActiveWorkbook.FollowHyperlink Address:="https://www1.tomraconnect.com/page/installation/view?installation=28210", NewWindow:=True
End Sub

Private Sub FullscreenButton_Click()
    Call Viper.Viper_Fullscreen
End Sub

Private Sub InstallationDetailsLink_Click()
    ActiveWorkbook.FollowHyperlink Address:=("https://www1.tomraconnect.com/page/installation/details?installation=" & Configuration.tomraConnectInstallationNo), NewWindow:=True
    'ActiveWorkbook.FollowHyperlink Address:="https://www1.tomraconnect.com/page/installation/details?installation=28210", NewWindow:=True
End Sub

Private Sub nextButton_Click()
    Call Viper.Viper_Next
End Sub

Private Sub NextHiddenButton_Click()
    Call Viper.Viper_NextHidden
End Sub

Private Sub OkExitButton_Click()
    Call Viper.Viper_Save
    ViperForm.Hide
End Sub

Private Sub PauseButton_Click()
    Call Viper.Viper_TogglePause
End Sub

Private Sub RestartButton_Click()
    Call Viper.Viper_Restart
End Sub

Private Sub PlayPauseButton_Click()
    Call Viper.Viper_TogglePause
End Sub

Private Sub prevButton_Click()
    Call Viper.Viper_Prev
End Sub

Private Sub PrevHiddenButton_Click()
    Call Viper.Viper_PrevHidden
End Sub

Private Sub SpeedButton025x_Click()
    Call Viper.Viper_SetSpeed(0.25)
End Sub

Private Sub SpeedButton05x_Click()
    Call Viper.Viper_SetSpeed(0.5)
End Sub

Private Sub SpeedButton1x_Click()
    Call Viper.Viper_SetSpeed(1#)
End Sub

Private Sub SpeedButton2x_Click()
    Call Viper.Viper_SetSpeed(2#)
End Sub

Private Sub StatusHistoryLink_Click()
    ActiveWorkbook.FollowHyperlink Address:=("https://www1.tomraconnect.com/page/installation/status?installation=" & Configuration.tomraConnectInstallationNo), NewWindow:=True
    'ActiveWorkbook.FollowHyperlink Address:="https://www1.tomraconnect.com/page/installation/status?installation=28210", NewWindow:=True
End Sub

Private Sub SubtitlesButton_Click()
    Call Viper.Viper_ToggleSubtitles
End Sub

Private Sub UrlLabel_Click()
    ActiveWorkbook.FollowHyperlink Address:=ViperForm.UrlLabel.Caption, NewWindow:=True
End Sub