Class PageSetupSystem

#Region "语言"
    Private Sub SelectCurrentLanguage()
        For i As Integer = 0 To ComboBackgroundSuit.Items.Count - 1
            Dim item As MyComboBoxItem = CType(ComboBackgroundSuit.Items(i), MyComboBoxItem)
            If item.Tag.Equals(Lang) Then
                ComboBackgroundSuit.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub

    Private Sub RefreshLang() Handles ComboBackgroundSuit.SelectionChanged
        If Not IsLoaded Then Exit Sub
        If Not ComboBackgroundSuit.IsLoaded Then Exit Sub
        Dim TargetLang As String = CType(ComboBackgroundSuit.SelectedItem, MyComboBoxItem).Tag
        If TargetLang.Equals(Lang) Then Exit Sub
        If HasRunningMinecraft OrElse McLaunchLoader.State = LoadState.Loading Then
            Hint(GetLang("LangPageSetupSystemHintCloseGameBeforeChangeLanguage"))
            SelectCurrentLanguage()
            Exit Sub
        End If
        If HasDownloadingTask() Then
            Hint(GetLang("LangPageSetupSystemHintFinishDownloadTaskBeforeChangeLanguage"))
            SelectCurrentLanguage()
            Exit Sub
        End If
        Lang = TargetLang
        Application.Current.Resources.MergedDictionaries(1) = New ResourceDictionary With {.Source = New Uri("pack://application:,,,/Resources/Language/" & Lang & ".xaml", UriKind.RelativeOrAbsolute)}
        If Lang.Equals("zh-MEME") Then MyMsgBox($"此语言仅供娱乐，请勿当真{vbCr}此語言僅供娛樂，請勿當真{vbCr}This language is for entertainment only, please don't take it seriously", IsWarn:=True)
        WriteReg("Lang", Lang)
        MyMsgBox(GetLang("LangPageSetupSystemDialogContentLanguageRestart"), ForceWait:=True)
        Process.Start(New ProcessStartInfo(PathWithName))
        FormMain.EndProgramForce()
    End Sub

    Private Sub HelpTranslate(sender As Object, e As EventArgs) Handles BtnHelpTranslate.Click
        OpenWebsite("https://weblate.tangge233.cn/engage/PCL/")
    End Sub
#End Region

    Private Shadows IsLoaded As Boolean = False

    Private Sub PageSetupSystem_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        '重复加载部分
        PanBack.ScrollToHome()
        '语言
        SelectCurrentLanguage()

        '非重复加载部分
        If IsLoaded Then Exit Sub
        IsLoaded = True

        AniControlEnabled += 1
        Reload()
        SliderLoad()
        AniControlEnabled -= 1

    End Sub
    Public Sub Reload()

        '下载
        SliderDownloadThread.Value = Setup.Get("ToolDownloadThread")
        SliderDownloadSpeed.Value = Setup.Get("ToolDownloadSpeed")
        ComboDownloadVersion.SelectedIndex = Setup.Get("ToolDownloadVersion")

        'Mod 与整合包
        ComboDownloadTranslate.SelectedIndex = Setup.Get("ToolDownloadTranslate")
        'ComboDownloadMod.SelectedIndex = Setup.Get("ToolDownloadMod")
        ComboModLocalNameStyle.SelectedIndex = Setup.Get("ToolModLocalNameStyle")
        CheckDownloadIgnoreQuilt.Checked = Setup.Get("ToolDownloadIgnoreQuilt")
        CheckDownloadClipboard.Checked = Setup.Get("ToolDownloadClipboard")

        'Minecraft 更新提示
        CheckUpdateRelease.Checked = Setup.Get("ToolUpdateRelease")
        CheckUpdateSnapshot.Checked = Setup.Get("ToolUpdateSnapshot")

        '辅助设置
        CheckHelpChinese.Checked = Setup.Get("ToolHelpLanguage")

        '系统设置
        ComboSystemUpdate.SelectedIndex = Setup.Get("SystemSystemUpdate")
        If Val(Environment.OSVersion.Version.ToString().Split(".")(2)) >= 19042 Then
            ComboSystemUpdateBranch.SelectedIndex = Setup.Get("SystemSystemUpdateBranch")
        Else '不满足系统要求
            ComboSystemUpdateBranch.Items.Clear()
            ComboSystemUpdateBranch.Items.Add("Legacy")
            ComboSystemUpdateBranch.SelectedIndex = 0
            ComboSystemUpdateBranch.ToolTip = "由于你的 Windows 版本过低，不满足新版本要求，只能获取 Legacy 分支的更新。&#xa;升级到 Windows 10 20H2 或以上版本以获取最新更新。"
            ComboSystemUpdateBranch.IsEnabled = False
        End If
        ComboSystemActivity.SelectedIndex = Setup.Get("SystemSystemActivity")
        ComboSystemServer.SelectedIndex = Setup.Get("SystemSystemServer")
        TextSystemCache.Text = Setup.Get("SystemSystemCache")
        CheckSystemDisableHardwareAcceleration.Checked = Setup.Get("SystemDisableHardwareAcceleration")
        SliderAniFPS.Value = Setup.Get("UiAniFPS")

        '网络
        TextSystemHttpProxy.Text = Setup.Get("SystemHttpProxy")
        CheckDownloadCert.Checked = Setup.Get("ToolDownloadCert")

        '调试选项
        CheckDebugMode.Checked = Setup.Get("SystemDebugMode")
        SliderDebugAnim.Value = Setup.Get("SystemDebugAnim")
        CheckDebugDelay.Checked = Setup.Get("SystemDebugDelay")
        CheckDebugSkipCopy.Checked = Setup.Get("SystemDebugSkipCopy")

    End Sub

    '初始化
    Public Sub Reset()
        Try
            Setup.Reset("ToolDownloadThread")
            Setup.Reset("ToolDownloadSpeed")
            Setup.Reset("ToolDownloadVersion")
            Setup.Reset("ToolDownloadTranslate")
            Setup.Reset("ToolDownloadIgnoreQuilt")
            Setup.Reset("ToolDownloadClipboard")
            Setup.Reset("ToolDownloadMod")
            Setup.Reset("ToolModLocalNameStyle")
            Setup.Reset("ToolUpdateRelease")
            Setup.Reset("ToolUpdateSnapshot")
            Setup.Reset("ToolHelpLanguage")
            Setup.Reset("SystemDebugMode")
            Setup.Reset("SystemDebugAnim")
            Setup.Reset("SystemDebugDelay")
            Setup.Reset("SystemDebugSkipCopy")
            Setup.Reset("SystemSystemCache")
            Setup.Reset("SystemSystemUpdate")
            Setup.Reset("SystemSystemServer")
            Setup.Reset("SystemSystemActivity")
            Setup.Reset("SystemDisableHardwareAcceleration")
            Setup.Reset("SystemHttpProxy")
            Setup.Reset("ToolDownloadCert")
            Setup.Reset("UiAniFPS")

            Log("[Setup] 已初始化启动器页设置")
            Hint(GetLang("LangPageSetupSystemLaunchResetSuccess"), HintType.Finish, False)
        Catch ex As Exception
            Log(ex, GetLang("LangPageSetupSystemLaunchResetFail"), LogLevel.Msgbox)
        End Try

        Reload()
    End Sub

    '将控件改变路由到设置改变
    Private Shared Sub CheckBoxChange(sender As MyCheckBox, e As Object) Handles CheckDebugMode.Change, CheckDebugDelay.Change, CheckDebugSkipCopy.Change, CheckUpdateRelease.Change, CheckUpdateSnapshot.Change, CheckHelpChinese.Change, CheckDownloadIgnoreQuilt.Change, CheckDownloadCert.Change, CheckDownloadClipboard.Change, CheckSystemDisableHardwareAcceleration.Change
        If AniControlEnabled = 0 Then Setup.Set(sender.Tag, sender.Checked)
    End Sub
    Private Shared Sub SliderChange(sender As MySlider, e As Object) Handles SliderDebugAnim.Change, SliderDownloadThread.Change, SliderDownloadSpeed.Change, SliderAniFPS.Change
        If AniControlEnabled = 0 Then Setup.Set(sender.Tag, sender.Value)
    End Sub
    Private Shared Sub ComboChange(sender As MyComboBox, e As Object) Handles ComboDownloadVersion.SelectionChanged, ComboModLocalNameStyle.SelectionChanged, ComboDownloadTranslate.SelectionChanged, ComboSystemUpdate.SelectionChanged, ComboSystemUpdateBranch.SelectionChanged, ComboSystemActivity.SelectionChanged, ComboSystemServer.SelectionChanged
        If AniControlEnabled = 0 Then Setup.Set(sender.Tag, sender.SelectedIndex)
    End Sub
    Private Shared Sub TextBoxChange(sender As MyTextBox, e As Object) Handles TextSystemCache.ValidatedTextChanged, TextSystemHttpProxy.TextChanged
        If AniControlEnabled = 0 Then Setup.Set(sender.Tag, sender.Text)
    End Sub

    Private Sub StartClipboardListening() Handles CheckDownloadClipboard.Change
        If CheckDownloadClipboard.Checked Then
            RunInNewThread(Sub() CompClipboard.ClipboardListening())
        End If
    End Sub

    '滑动条
    Private Sub SliderLoad()
        SliderDownloadThread.GetHintText = Function(v) v + 1
        SliderDownloadSpeed.GetHintText =
        Function(v)
            Select Case v
                Case Is <= 14
                    Return (v + 1) * 0.1 & " M/s"
                Case Is <= 31
                    Return (v - 11) * 0.5 & " M/s"
                Case Is <= 41
                    Return (v - 21) & " M/s"
                Case Else
                    Return GetLang("LangPageSetupSystemDownloadSpeedUnlimited")
            End Select
        End Function
        SliderDebugAnim.GetHintText = Function(v) If(v > 29, GetLang("LangPageSetupSystemDebugAnimSpeedDisable"), (v / 10 + 0.1) & "x")
        SliderAniFPS.GetHintText =
            Function(v)
                Return $"{v + 1} FPS"
            End Function
    End Sub
    Private Sub SliderDownloadThread_PreviewChange(sender As Object, e As RouteEventArgs) Handles SliderDownloadThread.PreviewChange
        If SliderDownloadThread.Value < 100 Then Exit Sub
        If Not Setup.Get("HintDownloadThread") Then
            Setup.Set("HintDownloadThread", True)
            MyMsgBox(GetLang("LangPageSetupSystemDownloadSpeedDialogThreadTooMuchContent"), GetLang("LangDialogTitleWarning"), GetLang("LangDialogBtnIC"), IsWarn:=True)
        End If
    End Sub

    '硬件加速
    Private Sub Check_DisableHardwareAcceleration(sender As Object, user As Boolean) Handles CheckSystemDisableHardwareAcceleration.Change
        Hint("此项变更将在重启 PCL 后生效")
    End Sub

    '调试模式
    Private Sub CheckDebugMode_Change() Handles CheckDebugMode.Change
        If AniControlEnabled = 0 Then Hint(GetLang("LangPageSetupSystemDebugNeedRestart"),, False)
    End Sub

    '自动更新
    Private Sub ComboSystemActivity_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ComboSystemActivity.SelectionChanged
        If AniControlEnabled <> 0 Then Exit Sub
        If ComboSystemActivity.SelectedIndex <> 2 Then Exit Sub
        If MyMsgBox(GetLang("LangPageSetupSystemLaunchDialogAnnouncementSilentContent"), GetLang("LangDialogTitleWarning"), GetLang("LangPageSetupSystemLaunchDialogAnnouncementBtnConfirm"), GetLang("LangDialogBtnCancel"), IsWarn:=True) = 2 Then
            ComboSystemActivity.SelectedItem = e.RemovedItems(0)
        End If
    End Sub
    Private Sub ComboSystemUpdate_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ComboSystemUpdate.SelectionChanged
        If AniControlEnabled <> 0 Then Exit Sub
        If ComboSystemUpdate.SelectedIndex <> 3 Then Exit Sub
        If MyMsgBox(GetLang("LangPageSetupSystemLaunchDialogAnnouncementDisableContent"), GetLang("LangDialogTitleWarning"), GetLang("LangPageSetupSystemLaunchDialogAnnouncementBtnConfirm"), GetLang("LangDialogBtnCancel"), IsWarn:=True) = 2 Then
            ComboSystemUpdate.SelectedItem = e.RemovedItems(0)
        End If
    End Sub
    Private Sub ComboSystemUpdateBranch_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ComboSystemUpdateBranch.SelectionChanged
        If AniControlEnabled <> 0 Then Exit Sub
        If ComboSystemUpdateBranch.SelectedIndex <> 1 Then Exit Sub
        If MyMsgBox("你正在切换启动器更新通道到 Fast Ring。" & vbCrLf &
                    "Fast Ring 可以提供下个版本更新内容的预览，但可能会包含未经充分测试的功能，稳定性欠佳。" & vbCrLf & vbCrLf &
                    "在升级到 Fast Ring 版本后，如果你选择切换到 Slow Ring，需要等待下一个 Slow Ring 版本发布，在这期间不会提供更新。" & vbCrLf &
                    "该选项仅推荐具有一定基础知识和能力的用户选择。如果你正在制作整合包，请使用 Slow Ring！", "继续之前...", "我已知晓", "取消", IsWarn:=True) = 2 Then
            ComboSystemUpdateBranch.SelectedItem = e.RemovedItems(0)
        Else
            UpdateCheckByButton()
        End If
    End Sub
    Private Sub BtnSystemUpdate_Click(sender As Object, e As EventArgs) Handles BtnSystemUpdate.Click
        UpdateCheckByButton()
    End Sub
    ''' <summary>
    ''' 启动器是否已经是最新版？
    ''' 若返回 Nothing，则代表无更新缓存文件或出错。
    ''' </summary>
    Public Shared Function IsLauncherNewest() As Boolean?
        Try
            '确认服务器公告是否正常
            Dim ServerContent As String = ReadFile(PathTemp & "Cache\Notice.cfg")
            If ServerContent.Split("|").Count < 3 Then Return Nothing
            '确认是否为最新
#If RELEASE Then
            Dim NewVersionCode As Integer = ServerContent.Split("|")(2)
#Else
            Dim NewVersionCode As Integer = ServerContent.Split("|")(1)
#End If
            Return NewVersionCode <= VersionCode
        Catch ex As Exception
            Log(ex, GetLang("LangPageSetupSystemSystemLaunchUpdateFail"), LogLevel.Feedback)
            Return Nothing
        End Try
    End Function

#Region "导出 / 导入设置"

    Private Sub BtnSystemSettingExp_Click(sender As Object, e As MouseButtonEventArgs) Handles BtnSystemSettingExp.Click
        Hint(GetLang("LangPageSetupSystemInDev"))
    End Sub
    Private Sub BtnSystemSettingImp_Click(sender As Object, e As MouseButtonEventArgs) Handles BtnSystemSettingImp.Click
        Hint(GetLang("LangPageSetupSystemInDev"))
    End Sub

#End Region

End Class
