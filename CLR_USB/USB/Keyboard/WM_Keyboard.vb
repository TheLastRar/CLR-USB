Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports PSE

Namespace USB.Keyboard
    Class WM_Keyboard
        Inherits API_Keyboard

        Private Class NativeMethods
            <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall, SetLastError:=True)>
            Public Shared Function SetWindowsHookEx(ByVal idHook As Integer, ByVal HookProc As KBDLLHookProc, ByVal hMod As IntPtr, ByVal dwThreadId As Integer) As IntPtr
            End Function
            <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
            Public Shared Function CallNextHookEx(ByVal idHook As IntPtr, ByVal nCode As Integer, ByVal wParam As UIntPtr, ByVal lParam As IntPtr) As IntPtr
            End Function
            <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
            Public Shared Function UnhookWindowsHookEx(ByVal idHook As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
            End Function
        End Class

        <StructLayout(LayoutKind.Sequential)>
        Private Structure KBDLLHOOKSTRUCT
            Public vkCode As UInt32
            Public scanCode As UInt32
            Public flags As KBDLLHOOKSTRUCTFlags
            Public time As UInt32
            Public dwExtraInfo As UIntPtr
        End Structure

        <Flags()>
        Private Enum KBDLLHOOKSTRUCTFlags As UInt32
            LLKHF_EXTENDED = &H1
            LLKHF_INJECTED = &H10
            LLKHF_ALTDOWN = &H20
            LLKHF_UP = &H80
        End Enum

        Protected Const WH_KEYBOARD_LL As Integer = 13
        Protected Const HC_ACTION As Integer = 0

        Private WM_KEYDOWN As ULong = (&H100)
        Private WM_KEYUP As ULong = (&H101)
        Private WM_SYSKEYDOWN As ULong = (&H104)
        Private WM_SYSKEYUP As ULong = (&H105)

        Private Delegate Function KBDLLHookProc(ByVal nCode As Integer, ByVal wParam As UIntPtr, ByVal lParam As IntPtr) As IntPtr

        Private KBDLLHookProcDelegate As KBDLLHookProc = New KBDLLHookProc(AddressOf KeyboardHookProc)
        Private HHookID As IntPtr = IntPtr.Zero

        Protected Function KeyboardHookProc(ByVal nCode As Integer, ByVal wParam As UIntPtr, ByVal lParam As IntPtr) As IntPtr
            Try
                ''Profile Code
                Dim startUtcNow As Date = Date.UtcNow
                ''
                If (nCode = HC_ACTION) Then
                    Dim struct As KBDLLHOOKSTRUCT
                    Select Case wParam.ToUInt64
                        Case WM_KEYDOWN, WM_SYSKEYDOWN
                            struct = CType(Marshal.PtrToStructure(lParam, struct.GetType()), KBDLLHOOKSTRUCT)
                            RaiseKeyDown(CType(struct.vkCode, Keys))
                        Case WM_KEYUP, WM_SYSKEYUP
                            struct = CType(Marshal.PtrToStructure(lParam, struct.GetType()), KBDLLHOOKSTRUCT)
                            RaiseKeyUp(CType(struct.vkCode, Keys))
                    End Select
                End If
                ''Profile Code
                Dim endUtcNow As Date = Date.UtcNow
                If ((endUtcNow - startUtcNow).TotalSeconds > 1.0) Then
                    Log_Error("KeyboardProc Took Too Long")
                End If
                ''
            Catch e As Exception
                CLR_PSE_PluginLog.MsgBoxError(e)
                Throw
            End Try
            Return NativeMethods.CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam)
        End Function

        Public Sub New(_hWnd As IntPtr)
            MyBase.New(_hWnd)
            'http://stackoverflow.com/a/17898148
            ''SetWindowsHookEx() is a bit awkward for the low-level hooks. It requires a valid module handle, and
            ''checks it, but doesn't actually use it. This got fixed in Windows, somewhere around Win7 SP1.
            ''Also works with nullptr on latest WinVista

            Dim hMod As IntPtr = IntPtr.Zero
            HHookID = NativeMethods.SetWindowsHookEx(WH_KEYBOARD_LL, KBDLLHookProcDelegate, hMod, 0)

            If HHookID = IntPtr.Zero Then
                Throw New Exception("Could not set keyboard hook. Error: " + Marshal.GetLastWin32Error.ToString())
            End If
        End Sub

        Public Overrides Sub Close()
            If Not HHookID = IntPtr.Zero Then
                NativeMethods.UnhookWindowsHookEx(HHookID)
                HHookID = IntPtr.Zero
            End If
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
