Imports CLRUSB.USB.Keyboard
Namespace Config
    Class ConfigDataKeyboard
        Implements IConfigData

        Public ToRegion As EmuMode = EmuMode.PassThough
        Public UseAPI As EnumAPI = EnumAPI.WM
        Public RawAPIKeyboard As String = ""

        Public Sub SaveConfig(Port As Integer, Data As FreezeDataHelper) Implements IConfigData.SaveConfig
            SaveLoad(Port, Data, True)
        End Sub

        Public Sub LoadConfig(Port As Integer, Data As FreezeDataHelper) Implements IConfigData.LoadConfig
            SaveLoad(Port, Data, False)
        End Sub

        Private Sub SaveLoad(Port As Integer, Data As FreezeDataHelper, Saving As Boolean)
            Try
                Dim refToR As Integer = ToRegion
                Data.SetInt32Value("Port" & Port & ".Keyboard.ToRegion", refToR, Saving)
                If Not Saving Then
                    ToRegion = CType(refToR, EmuMode)
                End If
            Catch
            End Try
            Try
                Dim refUAPI As Integer = UseAPI
                Data.SetInt32Value("Port" & Port & ".Keyboard.UseAPI", refUAPI, Saving)
                If Not Saving Then
                    UseAPI = CType(refUAPI, EnumAPI)
                End If
            Catch
            End Try

            Try
                Data.SetStringValue("Port" & Port & ".Keyboard.RawKeyboard", RawAPIKeyboard, Saving)
            Catch
            End Try

        End Sub
    End Class
End Namespace
