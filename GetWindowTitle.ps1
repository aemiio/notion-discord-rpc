Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class Win32 {
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int count);
    }
"@

function Get-WindowTitle {
    $hWnd = [Win32]::GetForegroundWindow()
    $title = New-Object -TypeName System.Text.StringBuilder -ArgumentList 256
    [Win32]::GetWindowText($hWnd, $title, $title.Capacity) | Out-Null
    $title.ToString()
}

Get-WindowTitle
