Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "powershell -command ""$hwnd = (Get-Process msedge).MainWindowHandle; Add-Type -Name Win32 -Namespace Native -MemberDefinition '[DllImport(""user32.dll"")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);'; [Native.Win32]::SetWindowPos($hwnd, -1, 0, 0, 0, 0, 0x0003)"""
