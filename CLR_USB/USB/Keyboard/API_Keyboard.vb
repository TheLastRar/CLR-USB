Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Namespace USB.Keyboard
    MustInherit Class API_Keyboard
        Implements IDisposable
        'LPARAM is a typedef for LONG_PTR which is a long (signed 32-bit) on win32 and __int64 (signed 64-bit) on x86_64.
        'WPARAM is a typedef for UINT_PTR which is an unsigned int (unsigned 32-bit) on win32 and unsigned __int64 (unsigned 64-bit) on x86_64.

        Private Class NativeMethods
            <DllImport("user32.dll")>
            Public Overloads Shared Function ToUnicode(virtualKeyCode As UInteger, scanCode As UInteger, keyboardState As Byte(),
            <Out, MarshalAs(UnmanagedType.LPWStr, SizeConst:=64)> receivingBuffer As Text.StringBuilder,
            bufferSize As Integer, flags As UInteger) As Integer
            End Function
        End Class

        '<DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)> _
        'Public Overloads Shared Function GetAsyncKeyState(ByVal vkey As Integer) As Short
        'End Function

        Public Shared Event KeyDown(ByVal Key As Keys)
        Public Shared Event KeyUp(ByVal Key As Keys)

        Protected Sub RaiseKeyDown(ByVal Key As Keys)
            RaiseEvent KeyDown(Key)
        End Sub
        Protected Sub RaiseKeyUp(ByVal Key As Keys)
            RaiseEvent KeyUp(Key)
        End Sub

        Protected PCSX2_hWnd As IntPtr

        Public Sub New(_hWnd As IntPtr)
            PCSX2_hWnd = _hWnd
        End Sub

        Public Shared Function GetCharsFromKeys(ByVal keys As Keys, ByVal shift As Boolean, ByVal altGr As Boolean) As String
            Dim buf As New Text.StringBuilder(256)
            Dim keyboardState(256 - 1) As Byte
            If shift Then
                keyboardState(CType(Keys.ShiftKey, Integer)) = 255
            End If
            If altGr Then
                keyboardState(CType(Keys.ControlKey, Integer)) = 255
                keyboardState(CType(Keys.Menu, Integer)) = 255
            End If
            NativeMethods.ToUnicode(CType(keys, UInteger), 0, keyboardState, buf, 256, 0)
            Return buf.ToString
        End Function

        Public MustOverride Sub Close()

#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If
                'TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                'TODO: set large fields to null.
                Close()
            End If
            disposedValue = True
        End Sub

        Protected Overrides Sub Finalize()
            Dispose(False)
            MyBase.Finalize()
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region
    End Class
End Namespace
