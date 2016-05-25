Imports PSE
Imports CLRUSB

Namespace Config

    Enum SelectedDevice As Integer
        None = 0
        Keyboard = 1
    End Enum

    Class CLR_Config
        Private SettingsFile As FreezeDataHelper = Nothing
        Public IniFolderPath As String = "inis"
        'Params
        Public EnableLog As Boolean = False
        Public Port1SelectedDevice As SelectedDevice = SelectedDevice.Keyboard
        Public Port1Options As IConfigData = Nothing

        Public Port2SelectedDevice As SelectedDevice = SelectedDevice.None
        Public Port2Options As IConfigData = Nothing

        Public Sub Configure()
            LoadConfig()

            Dim LogToggel As New SingleToggle("Enable Logging (Developer use only)")
            LogToggel.ValueEnabled = EnableLog

            Dim DeviceOptionsString As String() = {"No Device", "Keyboard"}
            Dim Port1Deleg(DeviceOptionsString.Length - 1) As [Delegate]
            Port1Deleg(1) = Sub() ConfigKeyboard(1)
            Dim Port1Select As New ComboConfig("Device in Port 1", DeviceOptionsString, Port1Deleg)
            Port1Select.ValueSelected = Port1SelectedDevice

            Dim EmptyDeviceOptionsString As String() = {"No Device"}
            Dim Port2Deleg(EmptyDeviceOptionsString.Length - 1) As [Delegate]
            Dim Port2Select As New ComboConfig("Device in Port 2", EmptyDeviceOptionsString, Port2Deleg)
            Port2Select.ValueSelected = Port2SelectedDevice

            Dim ConForm As New DynamicConfigForm()
            ConForm.AddConfigControl(LogToggel)
            ConForm.AddConfigControl(Port1Select)
            ConForm.AddConfigControl(Port2Select)
            ConForm.ShowDialog()

            If ConForm.Accepted Then
                EnableLog = LogToggel.ValueEnabled
                Port1SelectedDevice = CType(Port1Select.ValueSelected, SelectedDevice)
                Port2SelectedDevice = CType(Port2Select.ValueSelected, SelectedDevice)
            End If

            ReloadDevConfig(1, Port1SelectedDevice)
            ReloadDevConfig(2, Port2SelectedDevice)

            SetLoggingState()
            CLR_USB.CreateDevices()

            SaveConfig()
        End Sub

        Protected Sub ConfigKeyboard(Port As Integer)
            'Make Sure Configured Device settings are loaded
            ReloadDevConfig(Port, SelectedDevice.Keyboard)

            Dim Data As ConfigDataKeyboard = Nothing
            Select Case Port
                Case 1
                    Data = DirectCast(Port1Options, ConfigDataKeyboard)
                Case 2
                    Data = DirectCast(Port2Options, ConfigDataKeyboard)
            End Select

            Dim TranslateOptions As String() = {"No", "To JPN layout"}
            Dim TranslateSelect As New ComboConfig("Translate Input?", TranslateOptions, Nothing)
            TranslateSelect.ValueSelected = Data.ToRegion
            Dim RawAPI As New USB.Keyboard.RAW_Keyboard()

            Const RawAPISelectionStart As Integer = 1

            Dim APIOptions(RawAPI.ListOfDevices.Count) As String
            APIOptions(0) = "Windows Messaging"
            Dim DeviceNames As List(Of String) = RawAPI.ListOfDevices.Keys.ToList()
            Dim index As Integer = RawAPISelectionStart
            Dim SelectedIndex As Integer = Data.UseAPI
            For Each Name As String In DeviceNames
                APIOptions(index) = "Raw API : " + RawAPI.ListOfDevices(Name).DeviceDesc

                If SelectedIndex = -1 Then
                    If Name = Data.RawAPIKeyboard Then
                        SelectedIndex = index
                    End If
                End If

                index += 1
            Next

            RawAPI.Dispose()

            If SelectedIndex = -1 Then
                SelectedIndex = RawAPISelectionStart
            End If

            Dim APISelect As New ComboConfig("Use API", APIOptions, Nothing)
            APISelect.ValueSelected = SelectedIndex

            Dim ConForm As New DynamicConfigForm()
            ConForm.Text = "Configure Port " & Port & " Keyboard"
            ConForm.AddConfigControl(TranslateSelect)
            ConForm.AddConfigControl(APISelect)
            ConForm.ShowDialog()

            If ConForm.Accepted Then
                Data.ToRegion = CType(TranslateSelect.ValueSelected, USB.Keyboard.EmuMode)

                If APISelect.ValueSelected >= RawAPISelectionStart Then
                    Data.UseAPI = USB.Keyboard.EnumAPI.RAW
                    Data.RawAPIKeyboard = DeviceNames(APISelect.ValueSelected - RawAPISelectionStart)
                Else
                    Data.UseAPI = CType(APISelect.ValueSelected, USB.Keyboard.EnumAPI)
                End If

            End If

            Data.SaveConfig(Port, SettingsFile)
        End Sub

        Protected Sub ReloadDevConfig(Port As Integer, DeviceType As SelectedDevice)
            Dim NewData As IConfigData = Nothing
            Select Case DeviceType
                Case SelectedDevice.None

                Case SelectedDevice.Keyboard
                    NewData = New ConfigDataKeyboard()
                    NewData.LoadConfig(Port, SettingsFile)
            End Select
            Select Case Port
                Case 1
                    Port1Options = NewData
                Case 2
                    Port2Options = NewData
            End Select
        End Sub

        Public Sub LoadConfig()
            Dim IniPath As String = IniFolderPath
            If IniFolderPath.EndsWith(IO.Path.DirectorySeparatorChar) Then
                IniPath += "USB_CLR.ini"
            Else
                IniPath += IO.Path.DirectorySeparatorChar + "USB_CLR.ini"
            End If

            SettingsFile = New FreezeDataHelper

            Try
                Dim Data As Byte() = IO.File.ReadAllBytes(IniPath)
                SettingsFile.FromBytes(Data, True)
            Catch ex As Exception
                SettingsFile = New FreezeDataHelper
                Log_Info("Failed to open " & IniPath)
                SaveConfig()
            End Try

            EnableLog = ReadBoolErrorSafe("Logging", EnableLog)
            Port1SelectedDevice = CType(ReadInt32ErrorSafe("Port1.SelectedDev", Port1SelectedDevice), SelectedDevice)
            Port2SelectedDevice = CType(ReadInt32ErrorSafe("Port2.SelectedDev", Port2SelectedDevice), SelectedDevice)

            ReloadDevConfig(1, Port1SelectedDevice)
            ReloadDevConfig(2, Port2SelectedDevice)

            SetLoggingState()
            CLR_USB.CreateDevices()
        End Sub
        Protected Sub SaveConfig()
            Dim IniPath As String = IniFolderPath
            If IniFolderPath.EndsWith(IO.Path.DirectorySeparatorChar) Then
                IniPath += "USB_CLR.ini"
            Else
                IniPath += IO.Path.DirectorySeparatorChar + "USB_CLR.ini"
            End If

            SettingsFile.SetBoolValue("Logging", EnableLog, True)
            Dim int As Integer = Port1SelectedDevice
            SettingsFile.SetInt32Value("Port1.SelectedDev", int, True)
            int = Port2SelectedDevice
            SettingsFile.SetInt32Value("Port2.SelectedDev", int, True)

            If Not IsNothing(Port1Options) Then
                Port1Options.SaveConfig(1, SettingsFile)
            End If
            If Not IsNothing(Port2Options) Then
                Port2Options.SaveConfig(2, SettingsFile)
            End If

            Dim Data As Byte() = SettingsFile.ToBytes(True)
            Try
                IO.File.WriteAllBytes(IniPath, Data)
            Catch ex As Exception
                Log_Error("Failed to open " & IniPath)
            End Try
        End Sub

        Private Sub SetLoggingState()
            If EnableLog Then
                CLR_PSE_PluginLog.SetFileLevel(SourceLevels.All)
            Else
                CLR_PSE_PluginLog.SetFileLevel(SourceLevels.Critical)
            End If
        End Sub

        Private Shared Sub Log_Error(str As String)
            CLR_PSE_PluginLog.WriteLine(TraceEventType.[Error], CInt(USBLogSources.PluginInterface), str)
        End Sub
        Private Shared Sub Log_Info(str As String)
            CLR_PSE_PluginLog.WriteLine(TraceEventType.Information, CInt(USBLogSources.PluginInterface), str)
        End Sub
        Private Shared Sub Log_Verb(str As String)
            CLR_PSE_PluginLog.WriteLine(TraceEventType.Verbose, CInt(USBLogSources.PluginInterface), str)
        End Sub

#Region "SetWrappers"
        Private Function ReadBoolErrorSafe(key As String, ByVal bool As Boolean) As Boolean
            Try
                Dim ret As Boolean = bool
                SettingsFile.SetBoolValue(key, ret, False)
                Return ret
            Catch ex As Exception
                Return bool
            End Try
        End Function
        Private Function ReadInt32ErrorSafe(key As String, ByVal si As Integer) As Integer
            Try
                Dim ret As Integer = si
                SettingsFile.SetInt32Value(key, ret, False)
                Return ret
            Catch ex As Exception
                Return si
            End Try
        End Function
#End Region

    End Class
End Namespace
