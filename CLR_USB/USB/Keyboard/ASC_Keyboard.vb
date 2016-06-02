Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Namespace USB.Keyboard
    Class GetAsyncKey_Keyboard
        Inherits API_Keyboard

        Private Class NativeMethods
            <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
            Public Shared Function GetAsyncKeyState(ByVal vkey As Integer) As UInt16
            End Function
        End Class

        Dim oldState(255 - 1) As Byte
        Private GUI_hWnd As IntPtr

        Public Overrides Sub Poll()
            ''Sends a KeyDown that is not Key.None
            ''But is not mapped to a key
            ''Try removing vkeycodes like Controlm shift, menu

            ''Start from 8 to ignore
            ''Mouse buttons and VK_CANCEL

            Dim newState(255 - 1) As Byte

            If GUI_hWnd <> Utils.GetForegroundWindow() Then
                oldState = newState
                Return
            End If

            For key As Integer = 8 To 255 - 1
                newState(key) = CByte((NativeMethods.GetAsyncKeyState(key) >> 8) >> 7)
            Next

            For key As Integer = 8 To 255 - 1
                If (newState(key) <> oldState(key)) Then
                    If newState(key) = 1 Then
                        RaiseKeyDown(CType(key, Keys))
                    Else
                        RaiseKeyUp(CType(key, Keys))
                    End If
                End If
            Next

            oldState = newState
        End Sub

        Public Sub New(_hWnd As IntPtr)
            MyBase.New(_hWnd)
            GUI_hWnd = Utils.GetTopParent(PCSX2_hWnd)
        End Sub

        Public Overrides Sub Close()

        End Sub

    End Class
End Namespace
