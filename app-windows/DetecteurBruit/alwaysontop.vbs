Set shell = CreateObject("WScript.Shell")

cmd = "powershell -NoProfile -WindowStyle Hidden -Command " & Chr(34) & _
      "$sig='[DllImport(""user32.dll"")]public static extern bool SetWindowPos(" & _
      "IntPtr hWnd,IntPtr hWndInsertAfter,int X,int Y,int cx,int cy,uint uFlags);';" & _
      "Add-Type -Name Win32 -Namespace Native -MemberDefinition $sig;" & _
      "$proc = Get-Process msedge | Sort-Object StartTime -Descending | Select-Object -First 1;" & _
      "$hwnd = $proc.MainWindowHandle;" & _
      "[Native.Win32]::SetWindowPos($hwnd,[IntPtr](-1),0,0,0,0,0x0003)" & Chr(34)

shell.Run cmd, 0, False
