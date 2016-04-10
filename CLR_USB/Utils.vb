Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Module Utils
    Public Sub memcpy(ByRef target As Byte(), ByVal targetstartindex As Integer, ByRef source As Byte(), ByVal sourcestartindex As Integer, ByVal num As Integer)
        For x As Integer = 0 To num - 1
            target(targetstartindex + x) = source(sourcestartindex + x)
        Next
    End Sub

    <DllImport("user32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetForegroundWindow() As IntPtr
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)> _
    Private Function GetWindowLong(hWnd As IntPtr, nIndex As Integer) As IntPtr
    End Function '32bit
    <DllImport("user32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)> _
    Private Function GetWindowLongPtr(hWnd As IntPtr, nIndex As Integer) As IntPtr
    End Function '64bit
    <DllImport("user32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Private Function GetParent(hWnd As IntPtr) As IntPtr
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Private Function IsWindow(hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function
    '<DllImport("Kernel32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    'Private Function IsBadReadPtr(lp As IntPtr, ucb As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
    'End Function
    Private Const GWL_STYLE As Integer = -16
    Private Const WS_CHILD As Long = &H40000000L
    Public Function GetTopParent(hWnd As IntPtr) As IntPtr
        Dim hWndTop As IntPtr = hWnd
        While (GetWindowLongAuto(hWndTop, GWL_STYLE).ToInt64 And WS_CHILD) <> 0
            hWndTop = GetParent(hWndTop)
        End While
        Return hWndTop
    End Function
    Private Function GetWindowLongAuto(hWnd As IntPtr, nIndex As Integer) As IntPtr
        If (IntPtr.Size = 4) Then
            Return GetWindowLong(hWnd, nIndex)
        Else
            Return GetWindowLongPtr(hWnd, nIndex)
        End If
    End Function
    Public Function IsAWindow(hWnd As IntPtr) As Boolean
        Return IsWindow(hWnd)
    End Function
    ''Not safe
    'Public Function IsBadPtrPtr(lp As IntPtr) As Boolean
    '    Return IsBadReadPtr(lp, IntPtr.Size)
    'End Function
End Module
