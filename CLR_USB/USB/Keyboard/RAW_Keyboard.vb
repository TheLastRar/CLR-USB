Imports System.Runtime.InteropServices
Imports System.Windows.Forms

''Credits for this section goes to http://www.codeproject.com/Articles/17123/Using-Raw-Input-from-C-to-handle-multiple-keyboard
''And LilyPad for how to actully get the Input Events

Namespace USB.Keyboard
    Class RAW_Keyboard
        Inherits API_Keyboard

        <DllImport("user32.dll", SetLastError:=True)> _
        Private Shared Function RegisterRawInputDevices(pRawInputDevice As RawInputDevice(), numberDevices As UInteger, size As UInteger) As Boolean
        End Function
        <DllImport("user32.dll", SetLastError:=True)> _
        Private Shared Function GetRawInputDeviceList(pRawInputDeviceList As IntPtr, ByRef numberDevices As UInteger, size As UInteger) As UInteger
        End Function
        <DllImport("user32.dll", SetLastError:=True)> _
        Private Shared Function GetRawInputDeviceInfo(hDevice As IntPtr, command As RawInputDeviceInfo, pData As IntPtr, ByRef size As UInteger) As UInteger
        End Function
        <DllImport("user32.dll", SetLastError:=True)> _
        Private Shared Function RegisterDeviceNotification(hRecipient As IntPtr, notificationFilter As IntPtr, flags As DeviceNotification) As IntPtr
        End Function
        <DllImport("user32.dll", SetLastError:=True)> _
        Private Shared Function UnregisterDeviceNotification(handle As IntPtr) As IntPtr
        End Function
        <DllImport("user32.dll", SetLastError:=True)> _
        Private Shared Function GetRawInputData(hRawInput As IntPtr, command As DataCommand, <Out> pData As IntPtr, <[In], Out> ByRef size As Integer, sizeHeader As Integer) As IntPtr
        End Function
        <DllImport("User32.dll", SetLastError:=True)> _
        Private Shared Function GetRawInputData(hRawInput As IntPtr, command As DataCommand, <Out> ByRef buffer As InputData, <[In], Out> ByRef size As Integer, cbSizeHeader As Integer) As Integer
        End Function
        ''
        '<DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall, SetLastError:=True)> _
        'Friend Shared Function SetProp(hWdn As IntPtr, lpstring As String, hData As IntPtr) As Boolean
        'End Function
        '<DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)> _
        'Friend Shared Function GetProp(hWdn As IntPtr, lpstring As String) As IntPtr
        'End Function
        <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall, SetLastError:=True)> _
        Private Shared Function SetWindowLong(hWdn As IntPtr, nIndex As Integer, dwNewLong As IntPtr) As IntPtr
        End Function '32bit
        <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall, SetLastError:=True)> _
        Private Shared Function SetWindowLongPtr(hWdn As IntPtr, nIndex As Integer, dwNewLong As IntPtr) As IntPtr
        End Function '64bit
        <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)> _
        Private Shared Function CallWindowProc(lpPrevWndFunc As IntPtr, hWdn As IntPtr, msg As UInteger, wParam As UIntPtr, lParam As IntPtr) As IntPtr
        End Function
        <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)> _
        Private Shared Function DefWindowProc(hWdn As IntPtr, msg As UInteger, wParam As UIntPtr, lParam As IntPtr) As IntPtr
        End Function
        '<DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)> _
        'Private Shared Function RegisterWindowMessage(lpProcName As String) As UInteger
        'End Function
        <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall, SetLastError:=True)> _
        Private Shared Function SendMessage(hWdn As IntPtr, msg As UInteger, wParam As UIntPtr, lParam As IntPtr) As IntPtr
        End Function

#Region "BiggerThanIExpectedListOfEnums/Structures,OhAndAClassForGoodMesure"

        Public Structure RawInputDeviceDictEntry
            Public DeviceDesc As String
            Public DeviceHandle As IntPtr
        End Structure

        Private Structure BroadcastDeviceInterface
            Public dbcc_size As Int32
            Public BroadcastDeviceType As BroadcastDeviceType
            Private dbcc_reserved As Int32
            Public dbcc_classguid As Guid
            Public dbcc_name As Char
        End Structure

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure Rawinputdevicelist
            Public hDevice As IntPtr
            Public dwType As UInteger
        End Structure

        <StructLayout(LayoutKind.Explicit)> _
        Private Structure RawData
            <FieldOffset(0)> _
            Friend mouse As Rawmouse
            <FieldOffset(0)> _
            Friend keyboard As Rawkeyboard
            <FieldOffset(0)> _
            Friend hid As Rawhid
        End Structure

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure InputData
            Public header As Rawinputheader ' 64 bit header size is 24  32 bit the header size is 16
            Public data As RawData ' Creating the rest in a struct allows the header size to align correctly for 32 or 64 bit
        End Structure

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure Rawinputheader
            Public dwType As UInteger ' Type of raw input (RIM_TYPEHID 2, RIM_TYPEKEYBOARD 1, RIM_TYPEMOUSE 0)
            Public dwSize As UInteger ' Size in bytes of the entire input packet of data. This includes RAWINPUT plus possible extra input reports in the RAWHID variable length array. 
            Public hDevice As IntPtr ' A handle to the device generating the raw input data. 
            Public wParam As IntPtr ' RIM_INPUT 0 if input occurred while application was in the foreground else RIM_INPUTSINK 1 if it was not.
        End Structure

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure Rawhid
            Public dwSizHid As UInteger
            Public dwCount As UInteger
            Public bRawData As Byte
        End Structure

        <StructLayout(LayoutKind.Explicit)> _
        Private Structure Rawmouse
            <FieldOffset(0)> _
            Public usFlags As UShort
            <FieldOffset(4)> _
            Public ulButtons As UInteger
            <FieldOffset(4)> _
            Public usButtonFlags As UShort
            <FieldOffset(6)> _
            Public usButtonData As UShort
            <FieldOffset(8)> _
            Public ulRawButtons As UInteger
            <FieldOffset(12)> _
            Public lLastX As Integer
            <FieldOffset(16)> _
            Public lLastY As Integer
            <FieldOffset(20)> _
            Public ulExtraInformation As UInteger
        End Structure

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure Rawkeyboard
            Public Makecode As UShort ' Scan code from the key depression
            Public Flags As UShort ' One or more of RI_KEY_MAKE, RI_KEY_BREAK, RI_KEY_E0, RI_KEY_E1
            Public Reserved As UShort ' Always 0    
            Public VKey As UShort ' Virtual Key Code
            Public Message As UInteger ' Corresponding Windows message for exmaple (WM_KEYDOWN, WM_SYASKEYDOWN etc)
            Public ExtraInformation As UInteger ' The device-specific addition information for the event (seems to always be zero for keyboards)
        End Structure

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure RawInputDevice
            Public UsagePage As HidUsagePage
            Public Usage As HidUsage
            Public Flags As RawInputDeviceFlags
            Public Target As IntPtr
        End Structure

        Private Enum BroadcastDeviceType
            DBT_DEVTYP_OEM = 0
            DBT_DEVTYP_DEVNODE = 1
            DBT_DEVTYP_VOLUME = 2
            DBT_DEVTYP_PORT = 3
            DBT_DEVTYP_NET = 4
            DBT_DEVTYP_DEVICEINTERFACE = 5
            DBT_DEVTYP_HANDLE = 6
        End Enum

        Private Enum DeviceNotification
            ''' <summary>The hRecipient parameter is a window handle.</summary>
            DEVICE_NOTIFY_WINDOW_HANDLE = &H0
            ''' <summary>The hRecipient parameter is a service status handle.</summary>
            DEVICE_NOTIFY_SERVICE_HANDLE = &H1
            ''' <summary>
            ''' Notifies the recipient of device interface events for all device interface classes. (The dbcc_classguid member is ignored.)
            ''' This value can be used only if the dbch_devicetype member is DBT_DEVTYP_DEVICEINTERFACE.
            '''</summary>
            DEVICE_NOTIFY_ALL_INTERFACE_CLASSES = &H4
        End Enum

        Private Enum DataCommand As UInteger
            RID_HEADER = &H10000005 ' Get the header information from the RAWINPUT structure.
            RID_INPUT = &H10000003 ' Get the raw data from the RAWINPUT structure.
        End Enum

        Private Enum HidUsagePage As UShort
            ''' <summary>Unknown usage page.</summary>
            UNDEFINED = &H0
            ''' <summary>Generic desktop controls.</summary>
            GENERIC = &H1
            ''' <summary>Simulation controls.</summary>
            SIMULATION = &H2
            ''' <summary>Virtual reality controls.</summary>
            VR = &H3
            ''' <summary>Sports controls.</summary>
            SPORT = &H4
            ''' <summary>Games controls.</summary>
            GAME = &H5
            ''' <summary>Keyboard controls.</summary>
            KEYBOARD = &H7
        End Enum

        Private Enum HidUsage As UShort
            ''' <summary>Unknown usage.</summary>
            Undefined = &H0
            ''' <summary>Pointer</summary>
            Pointer = &H1
            ''' <summary>Mouse</summary>
            Mouse = &H2
            ''' <summary>Joystick</summary>
            Joystick = &H4
            ''' <summary>Game Pad</summary>
            Gamepad = &H5
            ''' <summary>Keyboard</summary>
            Keyboard = &H6
            ''' <summary>Keypad</summary>
            Keypad = &H7
            ''' <summary>Muilt-axis Controller</summary>
            SystemControl = &H80
            ''' <summary>Tablet PC controls</summary>
            Tablet = &H80
            ''' <summary>Consumer</summary>
            Consumer = &HC
        End Enum

        <Flags()> _
        Private Enum RawInputDeviceFlags
            ''' <summary>No flags.</summary>
            NONE = 0
            ''' <summary>If set, this removes the top level collection from the inclusion list. This tells the operating system to stop reading from a device which matches the top level collection.</summary>
            REMOVE = &H1
            ''' <summary>If set, this specifies the top level collections to exclude when reading a complete usage page. This flag only affects a TLC whose usage page is already specified with PageOnly.</summary>
            EXCLUDE = &H10
            ''' <summary>If set, this specifies all devices whose top level collection is from the specified UsagePage. Note that Usage must be zero. To exclude a particular top level collection, use Exclude.</summary>
            PAGEONLY = &H20
            ''' <summary>If set, this prevents any devices specified by UsagePage or Usage from generating legacy messages. This is only for the mouse and keyboard.</summary>
            NOLEGACY = &H30
            ''' <summary>If set, this enables the caller to receive the input even when the caller is not in the foreground. Note that WindowHandle must be specified.</summary>
            INPUTSINK = &H100
            ''' <summary>If set, the mouse button click does not activate the other window.</summary>
            CAPTUREMOUSE = &H200
            ''' <summary>If set, the application-defined keyboard device hotkeys are not handled. However, the system hotkeys; for example, ALT+TAB and CTRL+ALT+DEL, are still handled. By default, all keyboard hotkeys are handled. NoHotKeys can be specified even if NoLegacy is not specified and WindowHandle is NULL.</summary>
            NOHOTKEYS = &H200
            ''' <summary>If set, application keys are handled.  NoLegacy must be specified.  Keyboard only.</summary>
            APPKEYS = &H400
            ''' <summary> If set, this enables the caller to receive input in the background only if the foreground application
            ''' does not process it. In other words, if the foreground application is not registered for raw input,
            ''' then the background application that is registered will receive the input.
            ''' </summary>
            EXINPUTSINK = &H1000
            DEVNOTIFY = &H2000
        End Enum

        Private Enum RawInputDeviceInfo As UInteger
            RIDI_DEVICENAME = &H20000007
            RIDI_DEVICEINFO = &H2000000B
            PREPARSEDDATA = &H20000005
        End Enum

        Private Class DeviceType
            Public Const RimTypemouse As Integer = 0
            Public Const RimTypekeyboard As Integer = 1
            Public Const RimTypeHid As Integer = 2
        End Class

        Private Class RegistryAccess
            Public Shared Function GetDeviceKey(device As String) As Microsoft.Win32.RegistryKey
                Dim split As String() = device.Substring(4).Split("#"c)

                Dim ClassCode As String = split(0) '// ACPI (Class code)
                Dim subClassCode As String = split(1) '// PNP0303 (SubClass code)
                Dim protocolCode As String = split(2) '// 3&13c0b0c5&0 (Protocol code)
                Return Microsoft.Win32.Registry.LocalMachine.OpenSubKey(String.Format("System\CurrentControlSet\Enum\{0}\{1}\{2}", ClassCode, subClassCode, protocolCode))
            End Function
        End Class
#End Region

        Shared ReadOnly DeviceInterfaceHid As New Guid("4D1E55B2-F16F-11CF-88CB-001111000030")
        Private Const KEYBOARD_OVERRUN_MAKE_CODE As Integer = &HFF

        Private Const RI_KEY_BREAK As Integer = &H1 '// Key Up (Key Down = 0)
        Private Const RI_KEY_E0 As Integer = &H2 '// Left version of the key
        Private Const SC_SHIFT_R As Integer = &H36
        Private Const WM_ACTIVATE As UInteger = (&H6)
        Private Const WM_INPUT As UInteger = (&HFF)
        Private Const WM_USB_DEVICECHANGE As UInteger = (&H219)
        Private Const GWLP_WNDPROC As Integer = -4

        Private Const WM_USER_PING As UInteger = &H700
        Private PING_RET As IntPtr = New IntPtr(1234)

        Private DNotifyID As IntPtr = IntPtr.Zero

        'SubClass Stuff
        Private myGSWndProcHandle As GCHandle
        Private myGSWndProcPointer As IntPtr = IntPtr.Zero
        Private eatenGSWndProcPointer As IntPtr = IntPtr.Zero

        'Private myGuiWndProcHandle As GCHandle
        'Private myGuiWndProcPointer As IntPtr = IntPtr.Zero
        'Private eatenGuiWndProcPointer As IntPtr = IntPtr.Zero
        'end of subclass stuff
        Public ListOfDevices As New Dictionary(Of String, RawInputDeviceDictEntry)
        Private TargetDeviceString As String = ""
        Private TargetDeviceID As IntPtr = IntPtr.Zero

        Private rAPI_hWnd As IntPtr
        Private GUI_hWnd As IntPtr

        Private sentry As New Object

        Public Sub New()
            MyBase.New(IntPtr.Zero)
            EnumerateDevices()
        End Sub

        Public Sub New(_hWnd As IntPtr, Optional targetKeyboard As String = "")
            MyBase.New(_hWnd)
            TargetDeviceString = targetKeyboard

            GUI_hWnd = GetTopParent(PCSX2_hWnd)
            rAPI_hWnd = PCSX2_hWnd

            'GSWndProc
            GSSubClassHook()

            ''GUIWndProc
            'If rAPI_hWnd <> GUI_hWnd Then
            '    GuiSubClassHook()
            'End If
            ''done subclassing

            Dim rid(1 - 1) As RawInputDevice

            rid(0).UsagePage = HidUsagePage.GENERIC
            rid(0).Usage = HidUsage.Keyboard
            'rid(0).Flags = RawInputDeviceFlags.INPUTSINK
            rid(0).Flags = RawInputDeviceFlags.NONE
            rid(0).Target = rAPI_hWnd

            If Not (RegisterRawInputDevices(rid, CUInt(rid.Length), CUInt(Marshal.SizeOf(rid(0))))) Then
                Throw New Exception("Failed to register raw input device(s). Error: " + Marshal.GetLastWin32Error().ToString())
            End If

            SetTargetDevice()

            DNotifyID = RegisterForDeviceNotification(rAPI_hWnd)
            If DNotifyID = IntPtr.Zero Then
                Throw New Exception("Registration for device notifications Failed. Error: " + Marshal.GetLastWin32Error().ToString())
            End If

            'WM_USER_Beat = RegisterWindowMessage("CLRUSB_RawAPI_Beat")
            Dim beatThread As New Threading.Thread(AddressOf HeartBeat)
            beatThread.Start()
        End Sub

        Private Sub HeartBeat()
            Dim endBeat As Boolean = False
            'Console.Error.WriteLine("Starting Heart Beat")
            Dim HeartFails As Integer = 0
            Do
                Threading.Thread.Sleep(2000) 'every 2 seconds
                SyncLock sentry
                    If eatenGSWndProcPointer = IntPtr.Zero Then
                        'Console.Error.WriteLine("Ending Heart Beat")
                        endBeat = True
                    Else
                        If GetForegroundWindow() <> GUI_hWnd Then
                            Continue Do
                        End If
                        Dim ret As IntPtr = SendMessage(rAPI_hWnd, WM_USER_PING, UIntPtr.Zero, IntPtr.Zero)
                        If Not ret = PING_RET Then
                            Console.Error.WriteLine("WndProc Heart Beat Failed")
                            Console.Error.WriteLine("USB Plugin RawAPI capture has probably stopped")
                            ' Console.Error.WriteLine("Pause and Resume emulation to recapture")
                            Console.Error.WriteLine("Attempting to recapture")
                            AllSubClassUnHook()
                            GSSubClassHook()
                            'endBeat = True
                        End If
                    End If
                End SyncLock
            Loop Until endBeat
        End Sub

        Public Delegate Function SubClassProcDelegate(ByVal hwnd As IntPtr, ByVal msg As UInteger, ByVal wParam As UIntPtr, ByVal lParam As IntPtr) As IntPtr
        Protected Function GSWndProc(ByVal _hWnd As IntPtr, ByVal msg As UInteger, ByVal wParam As UIntPtr, ByVal lParam As IntPtr) As IntPtr
            If rAPI_hWnd <> _hWnd Then
                Console.Error.WriteLine("Mismatched window handles")
            End If
            Select Case msg
                Case WM_INPUT
                    Handle_WM_INPUT(lParam)
                Case WM_USB_DEVICECHANGE
                    SetTargetDevice()
                Case WM_USER_PING
                    'Console.Error.WriteLine("Beat")
                    Return PING_RET
            End Select

            Return CallWindowProc(eatenGSWndProcPointer, _hWnd, msg, wParam, lParam)
        End Function

        'Protected Function GuiWndProc(ByVal _hWnd As IntPtr, ByVal msg As UInteger, ByVal wParam As UIntPtr, ByVal lParam As IntPtr) As IntPtr
        '    If GUI_hWnd <> _hWnd Then
        '        Console.Error.WriteLine("Mismatched window handles")
        '    End If
        '    Console.Error.WriteLine("Got WM on GUI")
        '    Select Case msg
        '        Case WM_INPUT
        '            Console.Error.WriteLine("Got WM_INPUT on GUI")
        '        Case WM_ACTIVATE
        '            'Console.Error.WriteLine("Got Activate")
        '    End Select

        '    Return CallWindowProc(eatenGuiWndProcPointer, _hWnd, msg, wParam, lParam)
        'End Function

        Private Sub Handle_WM_INPUT(ByVal lParam As IntPtr)
            Dim _rawBuffer As InputData

            Dim dwSize As Integer = 0
            GetRawInputData(lParam, DataCommand.RID_INPUT, IntPtr.Zero, dwSize, Marshal.SizeOf(GetType(Rawinputheader)))

            If (dwSize <> GetRawInputData(lParam, DataCommand.RID_INPUT, _rawBuffer, dwSize, Marshal.SizeOf(GetType(Rawinputheader)))) Then
                'Console.Error.WriteLine("Error getting the rawinput buffer")
                Return
            End If

            Dim virtualKey As Integer = _rawBuffer.data.keyboard.VKey
            Dim makeCode As Integer = _rawBuffer.data.keyboard.Makecode
            Dim flags As Integer = _rawBuffer.data.keyboard.Flags
            Dim hDevice As IntPtr = _rawBuffer.header.hDevice
            If TargetDeviceID <> IntPtr.Zero And TargetDeviceID <> hDevice Then
                Return 'If Target Keyboard is missing, don't filter
            End If

            If (virtualKey = KEYBOARD_OVERRUN_MAKE_CODE) Then Return

            Dim isE0BitSet As Boolean = ((flags And RI_KEY_E0) <> 0)

            Dim isBreakBitSet As Boolean = ((flags And RI_KEY_BREAK) <> 0)

            If isBreakBitSet Then
                RaiseKeyUp(VirturalKeyCorrection(_rawBuffer, virtualKey, isE0BitSet, makeCode))
            Else
                RaiseKeyDown(VirturalKeyCorrection(_rawBuffer, virtualKey, isE0BitSet, makeCode))
            End If
        End Sub

        Private Function VirturalKeyCorrection(_rawbuffer As InputData, virturalKey As Integer, isE0BitSet As Boolean, makeCode As Integer) As Keys
            Dim correctedKey As Keys = CType(virturalKey, Keys)

            If _rawbuffer.header.hDevice = IntPtr.Zero Then
                '// When hDevice is 0 and the vkey is VK_CONTROL indicates the ZOOM key
                If (_rawbuffer.data.keyboard.VKey = Keys.ControlKey) Then
                    correctedKey = Keys.Zoom
                End If
            Else
                Select Case virturalKey
                    '// Right-hand CTRL and ALT have their e0 bit set 
                    Case Keys.ControlKey
                        correctedKey = If(isE0BitSet, Keys.RControlKey, Keys.LControlKey)
                    Case Keys.Menu
                        correctedKey = If(isE0BitSet, Keys.RMenu, Keys.LMenu)
                    Case Keys.ShiftKey
                        correctedKey = If(makeCode = SC_SHIFT_R, Keys.RShiftKey, Keys.LShiftKey)
                    Case Else
                        '
                End Select
            End If
            Return correctedKey
        End Function

        Private Function RegisterForDeviceNotification(parent As IntPtr) As IntPtr
            Dim usbNotifyHandle As IntPtr = IntPtr.Zero
            Dim bdi As New BroadcastDeviceInterface()
            bdi.dbcc_size = Marshal.SizeOf(bdi)
            bdi.BroadcastDeviceType = BroadcastDeviceType.DBT_DEVTYP_DEVICEINTERFACE
            bdi.dbcc_classguid = DeviceInterfaceHid

            Dim mem As IntPtr = IntPtr.Zero
            Try
                mem = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(BroadcastDeviceInterface)))
                Marshal.StructureToPtr(bdi, mem, False)
                usbNotifyHandle = RegisterDeviceNotification(parent, mem, DeviceNotification.DEVICE_NOTIFY_WINDOW_HANDLE)
            Catch ex As Exception
                Console.Error.WriteLine("Registration for device notifications Failed. Error: {0}", Marshal.GetLastWin32Error())
                Console.Error.WriteLine(ex.StackTrace)
            Finally
                Marshal.FreeHGlobal(mem)
            End Try

            If usbNotifyHandle = IntPtr.Zero Then
                Console.Error.WriteLine("Registration for device notifications Failed. Error: {0}", Marshal.GetLastWin32Error())
            End If

            Return usbNotifyHandle
        End Function

        Private Sub SetTargetDevice()
            EnumerateDevices()
            If TargetDeviceString = "" Then
                TargetDeviceID = IntPtr.Zero
                Return
            End If
            If ListOfDevices.ContainsKey(TargetDeviceString) Then
                TargetDeviceID = ListOfDevices(TargetDeviceString).DeviceHandle
            Else
                TargetDeviceID = IntPtr.Zero 'Target Keyboard missing
            End If
        End Sub

        Public Sub EnumerateDevices()
            ListOfDevices.Clear()

            'add Global/Fake Keyboard (?)

            Dim deviceCount As UInteger
            Dim dwSize As Integer = Marshal.SizeOf(GetType(Rawinputdevicelist))

            If (GetRawInputDeviceList(IntPtr.Zero, deviceCount, CUInt(dwSize)) = 0) Then
                Dim pRawInputDeviceList As IntPtr = Marshal.AllocHGlobal(CInt(dwSize * deviceCount))
                GetRawInputDeviceList(pRawInputDeviceList, deviceCount, CUInt(dwSize))

                For i As UInteger = 0 To deviceCount - 1UI
                    Dim pcbSize As UInteger = 0
                    '// On Window 8 64bit when compiling against .Net > 3.5 using .ToInt32 you will generate an arithmetic overflow. Leave as it is for 32bit/64bit applications
                    Dim rid As Rawinputdevicelist = CType(Marshal.PtrToStructure(New IntPtr((pRawInputDeviceList.ToInt64() + (dwSize * i))), GetType(Rawinputdevicelist)), Rawinputdevicelist)

                    GetRawInputDeviceInfo(rid.hDevice, RawInputDeviceInfo.RIDI_DEVICENAME, IntPtr.Zero, pcbSize)

                    'if (pcbSize <= 0) continue; when pcbSize was an integer
                    If (pcbSize = 0 Or pcbSize > Integer.MaxValue) Then Continue For

                    Dim pData As IntPtr = Marshal.AllocHGlobal(CInt(pcbSize))
                    GetRawInputDeviceInfo(rid.hDevice, RawInputDeviceInfo.RIDI_DEVICENAME, pData, pcbSize)
                    Dim deviceName As String = Marshal.PtrToStringAnsi(pData)

                    If (rid.dwType = DeviceType.RimTypekeyboard) Then ' OrElse rid.dwType = DeviceType.RimTypeHid) Then

                        Dim DeviceDesc As String = GetDeviceDescription(deviceName)
                        If Not (DeviceDesc = "Terminal Server Keyboard Driver") Then
                            Dim DevInfo As RawInputDeviceDictEntry
                            DevInfo.DeviceDesc = DeviceDesc
                            DevInfo.DeviceHandle = rid.hDevice
                            ListOfDevices.Add(deviceName, DevInfo)
                        End If

                    End If
                    Marshal.FreeHGlobal(pData)
                Next
                Marshal.FreeHGlobal(pRawInputDeviceList)
                Return
            End If
            Throw New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error)
        End Sub

        Private Shared Function GetDeviceDescription(device As String) As String
            Dim deviceKey As Microsoft.Win32.RegistryKey = RegistryAccess.GetDeviceKey(device)
            Dim deviceDesc As String = deviceKey.GetValue("DeviceDesc").ToString()
            deviceDesc = deviceDesc.Substring(deviceDesc.IndexOf(";"c) + 1)

            Return deviceDesc
        End Function

        Private Sub GSSubClassHook()
            Dim newGSWndProc As SubClassProcDelegate = New SubClassProcDelegate(AddressOf GSWndProc) 'SubClassProcDelegate = New SubClassProcDelegate(AddressOf GSWndProc)
            myGSWndProcHandle = GCHandle.Alloc(newGSWndProc) 'Prevent GC
            Dim GSfp As IntPtr = Marshal.GetFunctionPointerForDelegate(newGSWndProc)
            eatenGSWndProcPointer = SubClassHook(rAPI_hWnd, GSfp)
            myGSWndProcPointer = GSfp 'keep track for debug purposes
        End Sub
        'Private Sub GuiSubClassHook()
        '    'Console.Error.WriteLine("Adding my WndProc Subclass")
        '    Dim newGuiWndProc As SubClassProcDelegate = New SubClassProcDelegate(AddressOf GuiWndProc)
        '    myGuiWndProcHandle = GCHandle.Alloc(newGuiWndProc) 'Prevent GC
        '    Dim Guifp As IntPtr = Marshal.GetFunctionPointerForDelegate(newGuiWndProc)
        '    eatenGuiWndProcPointer = SubClassHook(GUI_hWnd, Guifp)
        '    myGuiWndProcPointer = Guifp 'keep track for debug purposes
        'End Sub
        Private Sub AllSubClassUnHook()
            SyncLock sentry
                'GUI
                'SubClassUnHook(GUI_hWnd, eatenGuiWndProcPointer, myGuiWndProcPointer)
                'If myGuiWndProcHandle.IsAllocated Then
                '    myGuiWndProcHandle.Free()
                'End If

                'GS
                SubClassUnHook(rAPI_hWnd, eatenGSWndProcPointer, myGSWndProcPointer)
                If myGSWndProcHandle.IsAllocated Then
                    myGSWndProcHandle.Free()
                    myGSWndProcHandle = Nothing
                End If
            End SyncLock
        End Sub
        Private Function SubClassHook(ByVal WindowHandle As IntPtr, ByVal FpToAdd As IntPtr) As IntPtr
            Console.Error.WriteLine("Adding my WndProc Subclass")
            Dim ret As IntPtr
            If (IntPtr.Size = 4) Then
                ret = SetWindowLong(WindowHandle, GWLP_WNDPROC, FpToAdd)
            Else
                ret = SetWindowLongPtr(WindowHandle, GWLP_WNDPROC, FpToAdd)
            End If

            If (ret = IntPtr.Zero) Then
                Throw New Exception("Failed to SetWindowLong(Ptr). Error: " + Marshal.GetLastWin32Error().ToString())
            End If
            Console.Error.WriteLine("Completed successfully")
            Return ret
        End Function
        Private Sub SubClassUnHook(ByVal WindowHandle As IntPtr, ByRef FpToReturn As IntPtr, ByVal ExpectedReturnFp As IntPtr)
            If Not FpToReturn = IntPtr.Zero Then
                Console.Error.WriteLine("Removing my WndProc Subclass")
                Dim ret As IntPtr
                If (IntPtr.Size = 4) Then
                    ret = SetWindowLong(WindowHandle, GWLP_WNDPROC, FpToReturn)
                Else
                    ret = SetWindowLongPtr(WindowHandle, GWLP_WNDPROC, FpToReturn)
                End If

                If (ExpectedReturnFp = IntPtr.Zero) Then
                    FpToReturn = IntPtr.Zero
                    Return
                End If

                If ret = ExpectedReturnFp Then
                    Console.Error.WriteLine("Returned WndProc is mine")
                Else
                    If ret = IntPtr.Zero Then
                        Console.Error.WriteLine("Error, Unexpected return value, Win32Error: " & Marshal.GetLastWin32Error)
                    Else
                        'Console.Error.WriteLine("Returned WndProc isn't mine, take it back!")
                        Console.Error.WriteLine("Error, Unexpected return value, undoing operation")
                        If (IntPtr.Size = 4) Then
                            ret = SetWindowLong(WindowHandle, GWLP_WNDPROC, ret)
                        Else
                            ret = SetWindowLongPtr(WindowHandle, GWLP_WNDPROC, ret)
                        End If
                        If ret = IntPtr.Zero Then
                            Console.Error.WriteLine("Error, Undo failed, Win32Error: " & Marshal.GetLastWin32Error)
                        End If
                    End If
                End If

                FpToReturn = IntPtr.Zero
            End If
        End Sub

        Public Overrides Sub Close()
            If Not DNotifyID = IntPtr.Zero Then
                UnregisterDeviceNotification(DNotifyID)
                DNotifyID = IntPtr.Zero

                'Leave Inputs running for lilypad (?)
                'Lilypad seems to recover just fine
                Dim rid(1 - 1) As RawInputDevice

                rid(0).UsagePage = HidUsagePage.GENERIC
                rid(0).Usage = HidUsage.Keyboard
                rid(0).Flags = RawInputDeviceFlags.REMOVE
                rid(0).Target = rAPI_hWnd

                If Not (RegisterRawInputDevices(rid, CUInt(rid.Length), CUInt(Marshal.SizeOf(rid(0))))) Then
                    'Don't Care
                    'Throw New Exception("Failed to register raw input device(s). Error: " + Marshal.GetLastWin32Error().ToString())
                End If
            End If

            'Return The WndProc
            'Both RawAPI mode and lilypad subclass the WndProc
            'If we are opened and closed in altetnate order (and we are) then all should be fine
            '
            'However, if lilypad is configured while emulation was running, only that plugin is reset (close+open) 
            'and we lose our place on the subclass chain
            'I have no idea what to do about that,
            'so I'm going to carry on blisfully.

            ''Gui
            'If Not eatenGSWndProcPointer = IntPtr.Zero Then
            '    If (IntPtr.Size = 4) Then
            '        SetWindowLong(GUI_hWnd, GWLP_WNDPROC, eatenGuiWndProcPointer)
            '    Else
            '        SetWindowLongPtr(GUI_hWnd, GWLP_WNDPROC, eatenGuiWndProcPointer)
            '    End If
            '    eatenGuiWndProcPointer = IntPtr.Zero
            'End If
            'If myGuiWndProcHandle.IsAllocated Then
            '    myGuiWndProcHandle.Free()
            'End If

            'GS
            AllSubClassUnHook()
        End Sub
    End Class

End Namespace
