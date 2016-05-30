Imports System.Windows.Forms
Imports PSE

Namespace USB.Keyboard
    Enum EmuMode As Integer
        PassThough = 0
        ChangeToJPN = 1
        ChangeToUS = 2
    End Enum

    Enum EnumAPI As Integer
        RAW = -1
        WM = 0
        AsyncKeys = 1
        DInput = 2 'Is a wrapper around Raw API, No plans to support it
    End Enum

    Class USB_Keyboard
        Inherits USB_Device
        '/* HID interface requests */
        Const GET_REPORT As Integer = ClassInterfaceRequest Or &H1
        Const GET_IDLE As Integer = ClassInterfaceRequest Or &H2
        Const GET_PROTOCOL As Integer = ClassInterfaceRequest Or &H3
        'Const GET_INTERFACE As Integer = InterfaceRequest Or &HA
        Const SET_REPORT As Integer = ClassInterfaceOutRequest Or &H9
        Const SET_IDLE As Integer = ClassInterfaceOutRequest Or &HA
        Const SET_PROTOCOL As Integer = ClassInterfaceOutRequest Or &HB
        Const SET_INTERFACE As Integer = InterfaceOutRequest Or &H11

        Const LED_NUMPAD As Byte = (1 << 0)
        Const LED_CAPS As Byte = (1 << 1)
        Const LED_SCROLL As Byte = (1 << 2)
        'Const LED_COMPOSE As Byte = (1 << 3)
        'Const LED_KANA As Byte = (1 << 4)

        Const PollingRateInFrames As Byte = &HA
        Const PollingRateInms As Integer = PollingRateInFrames * 1 'in ms for low/full Speed, 125us for high speed
        Const CountryCode As Byte = 0 'UK=32, US=33 Japan=15, 00 = unkown
        Const RegionByte0 As Byte = &H9 'for string descriptors
        Const RegionByte1 As Byte = &H8

        Dim KeyToUSB(256 - 1) As Byte
        Dim KeysPressed As New List(Of Keys)
        Dim PastKeysPressed As New List(Of Keys)
        'Dim KeyQueue As New List(Of Keys)

        Dim Idle_MaxDuration As Integer = 500
        Dim Idle_timer As Integer = 0

        'LEDs
        Dim LastLEDState As Byte = &HFF
        Dim NumLockOn As Boolean = False 'Some games (JakX) start with NumLock off
        Dim CapsLockOn As Boolean = False
        Dim ScrollLockOn As Boolean = False

        Private ToRegion As EmuMode = EmuMode.PassThough
        Private UseAPI As EnumAPI = EnumAPI.WM
        Private RawAPIKeyboard As String = ""

#Region "qemu_keyboard_dev_descriptor"
        Dim qemu_keyboard_dev_descriptor As Byte() = {
    &H12,
    &H1,
    &H10, &H0,
    &H0,
    &H0,
    &H0,
    &H8,
    &H27, &H6,
    &H1, &H0,
    &H0, &H0,
    &H3,
    &H2,
    &H1,
    &H1}

        '/* u8 bLength; */
        '/* u8 bDescriptorType; Device
        '/* u16 bcdUSB; v1.0 */ 'nuvee has 0x01 0x01
        '/* u8 bDeviceClass; */
        '/* u8 bDeviceSubClass; */
        '/* u8 bDeviceProtocol; [ low/full speeds only ] */
        '/* u8 bMaxPacketSize0; 8 Byte
        '/* u16 idVendor; */ 'nuvee has diffrent values
        '/* u16 idProduct; */ 'nuvee has diffrent values
        '/* u16 bcdDevice */ 'nuvee has diffrent values
        '/* u8 iManufacturer; */ 'nuvee has diffrent values
        '/* u8 iProduct; */ 'nuvee has diffrent values
        '/* u8 iSerialNumber; */ 'nuvee has diffrent values
        '/* u8 bNumConfigurations; */

#End Region


#Region "qemu_keyboard_config_descriptor"
        Dim qemu_keyboard_config_descriptor As Byte() = {
    &H9,
    &H2,
    &H22, &H0,
    &H1,
    &H1,
    &H4,
    &HA0, _
 _
 _
 _
    50, _
 _
 _
 _
    &H9,
    &H4,
    &H0,
    &H0,
    &H1,
    &H3,
    &H1,
    &H1,
    &H5, _
 _
    &H9,
    &H21,
    &H1, &H0,
    CountryCode,
    &H1,
    &H22,
    50, 0, _
 _
    &H7,
    &H5,
    &H81,
    &H3,
    &H8, &H0,
    PollingRateInFrames}
        '/* one configuration */
        '/* u8 bLength; */
        '/* u8 bDescriptorType; Configuration */
        '/* u16 wTotalLength; */
        '/* u8 bNumInterfaces; (1) */
        '/* u8 bConfigurationValue; */
        '/* u8 iConfiguration; */
        '/* u8 bmAttributes;                                'nuvee has this set to 0x80
        '    Bit 7: must be set,
        '    6: Self-powered,
        '    5: Remote wakeup,
        '    4..0: resvd */
        '/* u8 MaxPower; */

        '    /* USB 1.1:
        '    * USB 2.0, single TT organization (mandatory):
        '    * one interface, protocol 0
        '    *
        '    * USB 2.0, multiple TT organization (optional):
        '    * two interfaces, protocols 1 (like single TT)
        '    * and 2 (multiple TT mode) ... config is
        '    * sometimes settable
        '    * NOT IMPLEMENTED
        '    */
        '    /* one interface */
        '/* u8 if_bLength; */
        '/* u8 if_bDescriptorType; Interface */
        '/* u8 if_bInterfaceNumber; */
        '/* u8 if_bAlternateSetting; */
        '/* u8 if_bNumEndpoints; */
        '/* u8 if_bInterfaceClass; */
        '/* u8 if_bInterfaceSubClass; */
        '/* u8 if_bInterfaceProtocol; [usb1.1 or single tt] */
        '/* u8 if_iInterface; */
        '    /* HID descriptor */
        '/* u8 bLength; */
        '/* u8 bDescriptorType; */
        '/* u16 HID_class */
        '/* u8 country_code */ 'Look into this
        '/* u8 num_descriptors */
        '/* u8 type; Report */
        '/* u16 len */
        '    /* one endpoint (status change endpoint) */
        '/* u8 ep_bLength; */
        '/* u8 ep_bDescriptorType; Endpoint */
        '/* u8 ep_bEndpointAddress; IN Endpoint 1 */
        '/* u8 ep_bmAttributes; Interrupt */
        '/* u16 ep_wMaxPacketSize; */ 'nuvee has 0x08 0x00 (and now we do)
        '/* u8 ep_bInterval; (255ms -- usb 2.0 spec) */ 'nuvee has 0x08
#End Region

#Region "qemu_keyboard_hid_report_descriptor"
        Dim qemu_keyboard_hid_report_descriptor As Byte() = {
    &H5, &H1,
    &H9, &H6,
    &HA1, &H1,
        &H5, &H7,
        &H19, &HE0,
        &H29, &HE7,
        &H15, &H0,
        &H25, &H1,
        &H75, &H1,
        &H95, &H8,
        &H81, &H2,
        &H95, &H1,
        &H75, &H8,
        &H81, &H3,
        &H95, &H5,
        &H75, &H1,
        &H5, &H8,
        &H19, &H1,
        &H29, &H5,
        &H91, &H2,
        &H95, &H1,
        &H75, &H3,
        &H91, &H3,
        &H95, &H6,
        &H75, &H8,
        &H15, &H0,
        &H25, &H65,
        &H5, &H7,
        &H19, &H0,
        &H29, &H65,
        &H81, &H0,
    &HC0} '// END_COLLECTION

        '    '// USAGE_PAGE (Generic Desktop)
        '    '// USAGE (Keyboard)
        '    '// COLLECTION (Application)
        '    '// USAGE_PAGE (Keyboard)
        '    '// USAGE_MINIMUM (Keyboard Left Control)
        '    '// USAGE_MAXIMUM (Keyboard Right GUI)
        '    '// LOGICAL_MINIMUM (0)
        '    '// LOGICAL_MAXIMUM (1)
        '    '// REPORT_SIZE (1)
        '    '// REPORT_COUNT (8)
        '    '// INPUT (Data,Var,Abs)
        '    '// REPORT_COUNT (1)
        '    '// REPORT_SIZE (8)
        '    '// INPUT (Cnst,Var,Abs)
        '    '// REPORT_COUNT (5)
        '    '// REPORT_SIZE (1)
        '    '// USAGE_PAGE (LEDs)
        '    '// USAGE_MINIMUM (Num Lock)
        '    '// USAGE_MAXIMUM (Kana)
        '    '// OUTPUT (Data,Var,Abs)
        '    '// REPORT_COUNT (1)
        '    '// REPORT_SIZE (3)
        '    '// OUTPUT (Cnst,Var,Abs)
        '    '// REPORT_COUNT (6)
        '    '// REPORT_SIZE (8)
        '    '// LOGICAL_MINIMUM (0)
        '    '// LOGICAL_MAXIMUM (101)
        '    '// USAGE_PAGE (Keyboard)
        '    '// USAGE_MINIMUM (Reserved (no event indicated))
        '    '// USAGE_MAXIMUM (Keyboard Application)
        '    '// INPUT (Data,Ary,Abs)
#End Region

        Private Sub InitUKKey()
            KeyToUSB(Keys.Oem8) = &H35 'the `¬¦ key, acts as the key that toggels jpn mode
            KeyToUSB(Keys.D1) = &H1E 'ALL SHIFTED NUMBER KEYS DIFFER
            KeyToUSB(Keys.D2) = &H1F
            KeyToUSB(Keys.D3) = &H20
            KeyToUSB(Keys.D4) = &H21
            KeyToUSB(Keys.D5) = &H22
            KeyToUSB(Keys.D6) = &H23
            KeyToUSB(Keys.D7) = &H24
            KeyToUSB(Keys.D8) = &H25
            KeyToUSB(Keys.D9) = &H26
            KeyToUSB(Keys.D0) = &H27

            KeyToUSB(Keys.OemMinus) = &H2D '-_
            KeyToUSB(Keys.Oemplus) = &H2E '=+

            KeyToUSB(Keys.Back) = &H2A 'backspace key
            KeyToUSB(Keys.Tab) = &H2B
            KeyToUSB(Keys.Q) = &H14
            KeyToUSB(Keys.W) = &H1A
            KeyToUSB(Keys.E) = &H8
            KeyToUSB(Keys.R) = &H15
            KeyToUSB(Keys.T) = &H17
            KeyToUSB(Keys.Y) = &H1C
            KeyToUSB(Keys.U) = &H18
            KeyToUSB(Keys.I) = &HC
            KeyToUSB(Keys.O) = &H12
            KeyToUSB(Keys.P) = &H13

            KeyToUSB(Keys.OemOpenBrackets) = &H2F '[{
            KeyToUSB(Keys.OemCloseBrackets) = &H30 '}}

            KeyToUSB(Keys.CapsLock) = &H39

            KeyToUSB(Keys.A) = &H4
            KeyToUSB(Keys.S) = &H16
            KeyToUSB(Keys.D) = &H7
            KeyToUSB(Keys.F) = &H9
            KeyToUSB(Keys.G) = &HA
            KeyToUSB(Keys.H) = &HB
            KeyToUSB(Keys.J) = &HD
            KeyToUSB(Keys.K) = &HE
            KeyToUSB(Keys.L) = &HF
            KeyToUSB(Keys.OemSemicolon) = &H33
            KeyToUSB(Keys.Oemtilde) = &H34 'is the '@ key
            KeyToUSB(Keys.Oem7) = &H32 'is the #~ key
            KeyToUSB(Keys.Enter) = &H28

            'KeyToUSB(Keys.LShiftKey) = &HE1
            KeyToUSB(Keys.Oem5) = &H64 'the \| button (?), Check this button
            KeyToUSB(Keys.Z) = &H1D
            KeyToUSB(Keys.X) = &H1B
            KeyToUSB(Keys.C) = &H6
            KeyToUSB(Keys.V) = &H19
            KeyToUSB(Keys.B) = &H5
            KeyToUSB(Keys.N) = &H11
            KeyToUSB(Keys.M) = &H10
            KeyToUSB(Keys.Oemcomma) = &H36 ',<
            KeyToUSB(Keys.OemPeriod) = &H37 '.>
            KeyToUSB(Keys.OemQuestion) = &H38 '/? (?)
            'KeyToUSB(Keys.RShiftKey) = &HE5

            'don't send system keys as both a system key and as part of the first byte
            'KeyToUSB(Keys.LControlKey) = &HE0
            'KeyToUSB(Keys.LMenu) = &HE2
            KeyToUSB(Keys.Space) = &H2C
            'KeyToUSB(Keys.RMenu) = &HE6 'right alt key (alt gr?)
            'KeyToUSB(Keys.RControlKey) = &HE4

            KeyToUSB(Keys.Insert) = &H49
            KeyToUSB(Keys.Delete) = &H4C
            KeyToUSB(Keys.Left) = &H50
            KeyToUSB(Keys.Home) = &H4A
            KeyToUSB(Keys.End) = &H4D
            KeyToUSB(Keys.Up) = &H52
            KeyToUSB(Keys.Down) = &H51
            KeyToUSB(Keys.PageUp) = &H4B
            KeyToUSB(Keys.PageDown) = &H4E
            KeyToUSB(Keys.Right) = &H4F

            'numpad
            KeyToUSB(Keys.NumLock) = &H53
            KeyToUSB(Keys.NumPad7) = &H5F
            KeyToUSB(Keys.NumPad4) = &H5C
            KeyToUSB(Keys.NumPad1) = &H59
            KeyToUSB(Keys.Divide) = &H54
            KeyToUSB(Keys.NumPad8) = &H60
            KeyToUSB(Keys.NumPad5) = &H5D
            KeyToUSB(Keys.NumPad2) = &H5A
            KeyToUSB(Keys.NumPad0) = &H62
            KeyToUSB(Keys.Multiply) = &H55
            KeyToUSB(Keys.NumPad9) = &H61
            KeyToUSB(Keys.NumPad6) = &H5E
            KeyToUSB(Keys.NumPad3) = &H5B
            KeyToUSB(Keys.Decimal) = &H5C
            KeyToUSB(Keys.Subtract) = &H56
            KeyToUSB(Keys.Add) = &H57
            'NumPadEnter (is the extended variant of the regular enter)

            KeyToUSB(Keys.Escape) = &H29
            'function keys

            'printscreen
            KeyToUSB(Keys.Scroll) = &H47
            KeyToUSB(Keys.Pause) = &H48

            'KeyToUSB(Keys.LWin) = &HE3
            'KeyToUSB(Keys.RWin) = &HE7
            KeyToUSB(Keys.Apps) = &H65

            'unkown positions


            'Extra Keys Needed for Translation to JPN
            'These keys are Undefined acording to the VK Spec
            'so I am using it as a constant mapping to emulated keys
            'JPN
            KeyToUSB(&H3A) = &H34 'JPN :*
            KeyToUSB(&H3B) = &H32 'JPN ]}
            KeyToUSB(&H3C) = &H87 'JPN \_ AKA (Yen _)
        End Sub

        Private Sub InitUSKey()
            InitUKKey()
            'apply changes

            KeyToUSB(Keys.Oemtilde) = &H35 'is the `~ key on US keyboards 'nuuvee puts a diffrent key here(?))

            KeyToUSB(Keys.Separator) = &H31 'is the \| key on US keyboards (?)

            KeyToUSB(Keys.OemQuotes) = &H34 'is the single/double qoute key (also OEM7)

            KeyToUSB(Keys.OemBackslash) = &H32 'Either the angle bracket or the backslash key on the RT 102-key keyboard

        End Sub

        Dim hWndTop As IntPtr
        Dim keyboard_grabbed As Boolean = False
        Dim CaptureKeyboard As Boolean = True
        Dim WithEvents hostKeyboard As API_Keyboard

        Sub New(ByVal Config As Config.ConfigDataKeyboard)
            MyBase.New()
            ToRegion = Config.ToRegion
            UseAPI = Config.UseAPI
            RawAPIKeyboard = Config.RawAPIKeyboard

            Speed = USB_SPEED_FULL
            devname = "Generic USB Keyboard"
            PastKeysPressed.Add(Keys.None) 'dummy key to reset key state
        End Sub

        Protected Function keyboard_poll(ByRef buf As Byte(), len As Integer, isInterrupt As Boolean) As Integer
            Dim l As Integer
            Dim KeysPressedConverted As List(Of Integer) = Nothing
            If Not (keyboard_grabbed) Then
                'hookinto keyboard
                'InitUKKey()
                InitUSKey()
                keyboard_grabbed = True
            End If

            SyncLock KeysPressed
                'PollAPI()
                'check if forground
                If hWndTop <> Utils.GetForegroundWindow() Then
                    KeysPressed.Clear()
                End If

                If KeysPressed.Contains(Keys.RControlKey) Then
                    CaptureKeyboard = Not CaptureKeyboard
                    If CaptureKeyboard Then
                        Log_Info("Keyboard Capture Enabled")
                    Else
                        Log_Info("Keyboard Capture Disabled")
                    End If
                    KeysPressed.Remove(Keys.RControlKey)
                End If

                If CaptureKeyboard = False Then
                    KeysPressed.Clear()
                End If

                'Handle idle state
                If isInterrupt Then
                    If Idle_MaxDuration <> 0 Then Idle_timer += PollingRateInms

                    If Idle_MaxDuration = 0 OrElse Idle_timer < Idle_MaxDuration Then
                        'only update on changed state
                        If PollIsUpdate() = False Then Return USB_RET_NAK
                    End If
                    Idle_timer = 0
                    'continue as normal
                End If
                PastKeysPressed.Clear()
                If KeysPressed.Contains(Keys.None) Then KeysPressed.Remove(Keys.None) 'Remove dummy key
                PastKeysPressed.AddRange(KeysPressed)

                'Normal Reporting
                KeysPressedConverted = TranslateKeys()
            End SyncLock
            l = 0

            buf(l) = 0
            l = 1
            'system keys
            If KeysPressedConverted.Contains(Keys.LControlKey) Then buf(0) = CByte(buf(0) Or (1 << 0))
            If KeysPressedConverted.Contains(Keys.LShiftKey) Then buf(0) = CByte(buf(0) Or (1 << 1))
            If KeysPressedConverted.Contains(Keys.LMenu) Then buf(0) = CByte(buf(0) Or (1 << 2)) 'L alt
            If KeysPressedConverted.Contains(Keys.LWin) Then buf(0) = CByte(buf(0) Or (1 << 3))

            If KeysPressedConverted.Contains(Keys.RControlKey) Then buf(0) = CByte(buf(0) Or (1 << 4))
            If KeysPressedConverted.Contains(Keys.RShiftKey) Then buf(0) = CByte(buf(0) Or (1 << 5))
            If KeysPressedConverted.Contains(Keys.RMenu) Then buf(0) = CByte(buf(0) Or (1 << 6)) 'R alt
            If KeysPressedConverted.Contains(Keys.RWin) Then buf(0) = CByte(buf(0) Or (1 << 7))
            buf(l) = 0
            l = 2

            For Each key As Keys In KeysPressedConverted
                'If key = Keys.None Then
                '    Continue For
                'End If
                'key down check
                If (l = 8) Then
                    buf(1) = 1 'TODO: implement RollOver report (byte 2-8 set to 1)
                    Exit For
                End If
                Dim uc As Byte = KeyToUSB(key)
                If ((uc > 0)) Then
                    buf(l) = uc
                    l += 1
                End If
            Next

            While (l < 8)
                buf(l) = 0
                l += 1
            End While
            Return l
        End Function

        Protected Overrides Function handle_control(ByVal request As Integer, ByVal value As Integer, ByVal index As Integer, ByVal length As Integer, ByRef data() As Byte) As Integer
            Dim ret As Integer = 0
            'Dim buf(8 - 1) As Byte
            Select Case (request)
                'Standard Device Requests
                Case DeviceRequest Or USB_REQ_GET_STATUS
                    data(0) = CByte((1 << USB_DEVICE_SELF_POWERED) Or (remote_wakeup << USB_DEVICE_REMOTE_WAKEUP))
                    data(1) = &H0
                    ret = 2
                Case DeviceOutRequest Or USB_REQ_CLEAR_FEATURE
                    If (value = USB_DEVICE_REMOTE_WAKEUP) Then
                        remote_wakeup = 0
                    Else
                        GoTo fail
                    End If
                    ret = 0
                Case DeviceOutRequest Or USB_REQ_SET_FEATURE
                    If (value = USB_DEVICE_REMOTE_WAKEUP) Then
                        remote_wakeup = 1
                    Else
                        GoTo fail
                    End If
                    ret = 0
                Case DeviceOutRequest Or USB_REQ_SET_ADDRESS
                    addr = CByte(value)
                    ret = 0
                Case DeviceRequest Or USB_REQ_GET_DESCRIPTOR
                    'value is the descriptor type/id
                    Select Case (value >> 8)
                        Case USB_DT_DEVICE
                            Utils.memcpy(data, 0, qemu_keyboard_dev_descriptor, 0,
                                   qemu_keyboard_dev_descriptor.Count)
                            ret = qemu_keyboard_dev_descriptor.Count
                        Case USB_DT_CONFIG
                            Utils.memcpy(data, 0, qemu_keyboard_config_descriptor, 0,
                                   qemu_keyboard_config_descriptor.Count)
                            ret = qemu_keyboard_config_descriptor.Count
                        Case USB_DT_STRING
                            'index is the language ID
                            Select Case (value And &HFF)
                                Case 0
                                    '/* language ids */
                                    data(0) = 4 'blength in bytes
                                    data(1) = 3 'bDescriptorType
                                    data(2) = RegionByte0 'wLANGID[0].2
                                    data(3) = RegionByte1 'wLANGID[0].1
                                    ret = 4
                                Case 1 'see qemu_keyboard_dev_descriptor
                                    '/* serial number */
                                    ret = set_usb_string(data, "1")
                                Case 2 'see qemu_keyboard_dev_descriptor
                                    '/* product description */
                                    ret = set_usb_string(data, devname)
                                Case 3 'see qemu_keyboard_dev_descriptor
                                    '/* vendor description */
                                    ret = set_usb_string(data, "CLR_USB/PCSX2/QEMU")
                                Case 4 'see qemu_keyboard_config_descriptor
                                    ret = set_usb_string(data, "HID Keyboard")
                                Case 5 'see qemu_keyboard_config_descriptor /* u8 if_iInterface; */
                                    ret = set_usb_string(data, "Endpoint1 Interrupt Pipe")
                                Case Else
                                    GoTo fail
                            End Select
                        Case Else
                            GoTo fail
                    End Select
                Case DeviceRequest Or USB_REQ_GET_CONFIGURATION
                    data(0) = 1
                    ret = 1
                Case DeviceOutRequest Or USB_REQ_SET_CONFIGURATION
                    'Setting the device config
                    'Zero means unconfigured
                    '1 is the 1st (and only) config
                    If (value <> 1) Then
                        GoTo fail
                    End If
                    ret = 0
                Case DeviceRequest Or USB_REQ_GET_INTERFACE
                    data(0) = 0
                    ret = 1
                Case DeviceOutRequest Or USB_REQ_SET_INTERFACE
                    ret = 0
                    '/* hid specific requests */
                    Select Case (value >> 8)
                        Case &H22
                            Utils.memcpy(data, 0, qemu_keyboard_hid_report_descriptor, 0,
                                   qemu_keyboard_hid_report_descriptor.Count)
                            ret = qemu_keyboard_hid_report_descriptor.Count
                        Case Else
                            GoTo fail
                    End Select
                    'Interface requests
                Case InterfaceOutRequest Or USB_REQ_SET_INTERFACE
                    'Setting the interface altanative setting
                    'We only have 1 (index 0)
                    If (value <> 0) Then
                        GoTo fail
                    End If
                    ret = 0
                    'class interface requests
                Case SET_PROTOCOL 'Boot vs report protocol
                    ret = 0
                Case GET_REPORT
                    ret = keyboard_poll(data, length, False)
                    'case Get_IDLE,PROTOCOL
                Case SET_REPORT 'LEDs
                    'investigate report IDs/type (maybe)
                    NumLockOn = (data(0) And LED_NUMPAD) <> 0
                    CapsLockOn = (data(0) And LED_CAPS) <> 0
                    ScrollLockOn = (data(0) And LED_SCROLL) <> 0
                    If (data(0) <> LastLEDState) Then
                        LastLEDState = data(0)
                        If NumLockOn Then
                            Log_Info("NumPad ON")
                        Else
                            Log_Info("NumPad OFF")
                        End If
                        If CapsLockOn Then
                            Log_Info("CapsLock ON")
                        Else
                            Log_Info("CapsLock OFF")
                        End If
                        If ScrollLockOn Then
                            Log_Info("ScrollLock ON")
                        Else
                            Log_Info("ScrollLock OFF")
                        End If

                        If (data(0) And (LED_NUMPAD Or LED_CAPS Or LED_SCROLL)) = 0 AndAlso data(0) <> 0 Then
                            Log_Error("Unkown LED ON")
                            'GoTo fail
                        End If
                    End If
                    ret = 0
                Case SET_IDLE
                    Idle_MaxDuration = (value >> 8) * 4
                    ret = 0
                Case Else
fail:
                    Log_Error("STALL with request=0x" & request.ToString("X") & ", value=" & value & ", index=" & index & ", length = " & length)
                    ret = USB_RET_STALL
            End Select
            Return ret
        End Function

        Protected Overrides Function handle_data(ByVal pid As Integer, ByVal devep As Byte, ByRef data() As Byte, ByVal len As Integer) As Integer
            Dim ret As Integer = 0
            Select Case (pid)
                Case USB_TOKEN_IN
                    If (devep = 1) Then
                        ret = keyboard_poll(data, len, True)
                    Else
                        GoTo fail
                    End If
                    'case USB_TOKEN_OUT
                Case Else
fail:
                    ret = USB_RET_STALL
                    Log_Error("STALL on data")
            End Select
            Return ret
        End Function

        'Private Sub SyncLockKeys()
        '    'Sync Caps/Numlock/scroll lock
        '    If Control.IsKeyLocked(Keys.CapsLock) <> CapsOn Then
        '        KeyQueue.Add(Keys.CapsLock)
        '    End If
        '    If Control.IsKeyLocked(Keys.NumLock) <> NumLockOn Then
        '        KeyQueue.Add(Keys.NumLock)
        '    End If
        '    If Control.IsKeyLocked(Keys.Scroll) <> ScrollLck Then
        '        KeyQueue.Add(Keys.Scroll)
        '    End If
        'End Sub

        'Private Sub PollAPI()
        '    If UseEvents Then
        '        'do nothing
        '    Else
        '        'fill KeysPressed with GetAsyncKeys Data
        '        'or other polled api
        '    End If
        'End Sub
        Private Function PollIsUpdate() As Boolean
            If KeysPressed.Count = 0 AndAlso PastKeysPressed.Count = 0 Then Return False

            If KeysPressed.Count = PastKeysPressed.Count Then
                Dim Differ As Boolean = False
                For ki As Integer = 0 To KeysPressed.Count - 1
                    If Not (KeysPressed(ki) = PastKeysPressed(ki)) Then
                        Differ = True
                        Exit For
                    End If
                Next
                Return Differ
            End If
            Return True
        End Function
        'TODO: sync num/caps lock state bettween host-PS2
        Private Function TranslateKeys() As List(Of Integer)
            Dim KeysConverted As New List(Of Integer)
            Dim Shift As Boolean = KeysPressed.Contains(Keys.RShiftKey) Or KeysPressed.Contains(Keys.LShiftKey)
            Dim AltGr As Boolean = KeysPressed.Contains(Keys.LControlKey) And KeysPressed.Contains(Keys.RMenu)

            Dim NeedShift As Integer = 0 '1 = down, 0 = passthough, -1 = Up

            For index As Integer = 0 To KeysPressed.Count - 1
                Dim Key As Integer = KeysPressed(index)

                Select Case ToRegion
                    Case EmuMode.PassThough
                        KeysConverted.Add(Key)
                    Case EmuMode.ChangeToJPN
                        NeedShift = TranslateToJPN(KeysConverted, Key, Shift, AltGr)
                    Case EmuMode.ChangeToUS
                        'Not Implemented
                        KeysConverted.Add(Key)
                End Select
            Next

            'Adjust Shift
            If NeedShift = 1 And Not Shift Then
                KeysConverted.Add(Keys.LShiftKey)
            End If
            If NeedShift = -1 And Shift Then
                KeysConverted.Remove(Keys.LShiftKey)
                KeysConverted.Remove(Keys.RShiftKey)
            End If

            Return KeysConverted
        End Function

        Private Function TranslateToJPN(ByRef KeysConverted As List(Of Integer), ByVal Key As Integer, Shift As Boolean, AltGr As Boolean) As Integer
            'returns need shift as an int
            Dim NeedShift As Integer = 0
            If KeyToUSB(Key) = &H35 Then 'keep change mode unchanged
                KeysConverted.Add(Key)
                Return 0
            End If

            Select Case API_Keyboard.GetCharsFromKeys(CType(Key, Keys), Shift, AltGr).ToLower
                'Case "!"
                '    KeysConverted.Add(Key)
                Case """"
                    NeedShift = 1
                    If Not (KeysConverted.Contains(Keys.D2)) Then
                        KeysConverted.Add(Keys.D2)
                    End If
                Case "#"
                    NeedShift = 1
                    If Not (KeysConverted.Contains(Keys.D3)) Then
                        KeysConverted.Add(Keys.D3)
                    End If
                    'Case "$"
                    '    KeysConverted.Add(Key)
                Case "%"
                    NeedShift = 1
                    If Not (KeysConverted.Contains(Keys.D5)) Then
                        KeysConverted.Add(Keys.D5)
                    End If
                Case "&"
                    NeedShift = 1
                    If Not (KeysConverted.Contains(Keys.D6)) Then
                        KeysConverted.Add(Keys.D6)
                    End If
                Case "'"
                    NeedShift = 1
                    If Not (KeysConverted.Contains(Keys.D7)) Then
                        KeysConverted.Add(Keys.D7)
                    End If
                Case "("
                    NeedShift = 1
                    If Not (KeysConverted.Contains(Keys.D8)) Then
                        KeysConverted.Add(Keys.D8)
                    End If
                Case ")"
                    NeedShift = 1
                    If Not (KeysConverted.Contains(Keys.D9)) Then
                        KeysConverted.Add(Keys.D9)
                    End If
                    'Case "-"
                    '    KeysConverted.Add(Key)
                Case "="
                    NeedShift = 1
                    If Not (KeysConverted.Contains(Keys.OemMinus)) Then
                        KeysConverted.Add(Keys.OemMinus)
                    End If
                Case "^"
                    NeedShift = -1
                    If Not (KeysConverted.Contains(Keys.Oemplus)) Then
                        KeysConverted.Add(Keys.Oemplus)
                    End If
                Case "~"
                    NeedShift = 1
                    If Not (KeysConverted.Contains(Keys.Oemplus)) Then
                        KeysConverted.Add(Keys.Oemplus)
                    End If
                Case "¥" 'Why did i even add this?
                    NeedShift = -1
                    If Not (KeysConverted.Contains(Keys.Oem5)) Then
                        KeysConverted.Add(Keys.Oem5)
                    End If
                Case "|"
                    NeedShift = 1
                    If Not (KeysConverted.Contains(Keys.Oem5)) Then
                        KeysConverted.Add(Keys.Oem5)
                    End If
                Case "@"
                    NeedShift = -1
                    If Not (KeysConverted.Contains(Keys.OemOpenBrackets)) Then
                        KeysConverted.Add(Keys.OemOpenBrackets)
                    End If
                Case "`"
                    NeedShift = 1
                    If Not (KeysConverted.Contains(Keys.OemOpenBrackets)) Then
                        KeysConverted.Add(Keys.OemOpenBrackets)
                    End If
                Case "["
                    NeedShift = -1
                    If Not (KeysConverted.Contains(Keys.OemCloseBrackets)) Then
                        KeysConverted.Add(Keys.OemCloseBrackets)
                    End If
                Case "{"
                    NeedShift = 1
                    If Not (KeysConverted.Contains(Keys.OemCloseBrackets)) Then
                        KeysConverted.Add(Keys.OemCloseBrackets)
                    End If
                Case ";"
                    NeedShift = -1
                    If Not (KeysConverted.Contains(Keys.OemSemicolon)) Then
                        KeysConverted.Add(Keys.OemSemicolon)
                    End If
                Case "+"
                    NeedShift = 1
                    If Not (KeysConverted.Contains(Keys.OemSemicolon)) Then
                        KeysConverted.Add(Keys.OemSemicolon)
                    End If
                Case ":"
                    NeedShift = -1
                    If Not (KeysConverted.Contains(&H3A)) Then
                        KeysConverted.Add(&H3A)
                    End If
                Case "*"
                    NeedShift = 1
                    If Not (KeysConverted.Contains(&H3A)) Then
                        KeysConverted.Add(&H3A)
                    End If
                Case "]"
                    NeedShift = -1
                    If Not (KeysConverted.Contains(&H3B)) Then
                        KeysConverted.Add(&H3B)
                    End If
                Case "}"
                    NeedShift = 1
                    If Not (KeysConverted.Contains(&H3B)) Then
                        KeysConverted.Add(&H3B)
                    End If
                    'Case ","
                    '    KeysConverted.Add(Key)
                    'Case "<"
                    '    KeysConverted.Add(Key)
                    'Case "."
                    '    KeysConverted.Add(Key)
                    'Case ">"
                    '    KeysConverted.Add(Key)
                    'Case "/"
                    '    KeysConverted.Add(Key)
                    'Case "?"
                    '    KeysConverted.Add(Key)
                Case "\" 'AKA another Yen symbol
                    NeedShift = -1
                    If Not (KeysConverted.Contains(&H3C)) Then
                        KeysConverted.Add(&H3C)
                    End If
                Case "_"
                    NeedShift = 1
                    If Not (KeysConverted.Contains(&H3C)) Then
                        KeysConverted.Add(&H3C)
                    End If
                Case Else
                    If Not KeysConverted.Contains(Key) Then
                        KeysConverted.Add(Key)
                    End If
            End Select
            Return NeedShift
        End Function

        Public Overrides Sub handle_reset()
            'release of any handles held by this instance
            MyBase.handle_reset()
            KeysPressed.Clear()
            PastKeysPressed.Clear()

            Idle_MaxDuration = 500
            Idle_timer = 0

            'LEDs
            NumLockOn = False
            CapsLockOn = False
            ScrollLockOn = False
            CaptureKeyboard = True
            PastKeysPressed.Add(Keys.None) 'dummy key to fire first input
        End Sub

        Public Overrides Function open(_hWnd As IntPtr) As Integer
            hWndTop = Utils.GetTopParent(_hWnd)

            'hookinto APIs here (if needed)
            Select Case UseAPI
                Case EnumAPI.WM
                    hostKeyboard = New WM_Keyboard(_hWnd)
                Case EnumAPI.RAW
                    hostKeyboard = New RAW_Keyboard(_hWnd, RawAPIKeyboard)
                Case Else
                    Throw New NotImplementedException("Unknow API ID")
            End Select
            '
            Log_Verb("Open Keyboard")
            CaptureKeyboard = True
            Return 0
        End Function
        Public Overrides Sub close()
            'unhook APIs here (if needed)
            hostKeyboard.Dispose()
            hostKeyboard = Nothing
            Log_Verb("Close Keyboard")
        End Sub
        Private Sub KbHook_KeyDown(ByVal Key As Keys) Handles hostKeyboard.KeyDown
            SyncLock KeysPressed
                If Not KeysPressed.Contains(Key) Then
                    KeysPressed.Add(Key)
                End If
            End SyncLock
        End Sub
        Private Sub KbHook_KeyUp(ByVal Key As Keys) Handles hostKeyboard.KeyUp
            SyncLock KeysPressed
                KeysPressed.Remove(Key)
            End SyncLock
        End Sub

        Public Overrides Sub Freeze(ByRef freezedata As FreezeDataHelper, index As Integer, save As Boolean)
            freezedata.SetInt32Value("OHCI.P" & index & ".KB.Idle_MaxDuration", Idle_MaxDuration, save)
            freezedata.SetBoolValue("OHCI.P" & index & ".KB.NumLockOn", NumLockOn, save)
            freezedata.SetBoolValue("OHCI.P" & index & ".KB.CapsLockOn", CapsLockOn, save)
            freezedata.SetBoolValue("OHCI.P" & index & ".KB.ScrollLockOn", ScrollLockOn, save)

            MyBase.Freeze(freezedata, index, save)
            CaptureKeyboard = True
            PastKeysPressed.Add(Keys.None) 'dummy key to reset key state
        End Sub

        Private Shared Sub Log_Error(str As String)
            CLR_PSE_PluginLog.WriteLine(TraceEventType.[Error], (USBLogSources.USBKeyboard), str)
        End Sub
        Private Shared Sub Log_Info(str As String)
            CLR_PSE_PluginLog.WriteLine(TraceEventType.Information, USBLogSources.USBKeyboard, str)
        End Sub
        Private Shared Sub Log_Verb(str As String)
            CLR_PSE_PluginLog.WriteLine(TraceEventType.Verbose, (USBLogSources.USBKeyboard), str)
        End Sub
    End Class
End Namespace
