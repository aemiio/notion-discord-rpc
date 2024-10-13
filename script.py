from subprocess import Popen, PIPE
from pypresence import Presence
import time

# Use a list for the Popen command arguments
powershell_command = [
    'powershell.exe', '-Command',
    'Add-Type @"\n    using System;\n    using System.Runtime.InteropServices;\n    public class '
    'Win32 {\n        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]\n    '
    '    public static extern IntPtr GetForegroundWindow();\n        [DllImport("user32.dll", '
    'CharSet = CharSet.Auto, SetLastError = true)]\n        public static extern int '
    'GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int count);\n    '
    '}\n"@\n\nfunction Get-WindowTitle {\n    $hWnd = [Win32]::GetForegroundWindow()\n    $title '
    '= New-Object -TypeName System.Text.StringBuilder -ArgumentList 256\n    ['
    'Win32]::GetWindowText($hWnd, $title, $title.Capacity) | Out-Null\n    $title.ToString('
    ')\n}\n\nGet-WindowTitle'
]

# Run PowerShell script to get the active window title
#(title, tError)
(tError) = Popen(powershell_command, stdout=PIPE, universal_newlines=True).communicate()


client_id = "861651444234846258"  # application ID
RPC = Presence(client_id=client_id)
RPC.connect()

# use the same name that you used when uploading the image
RPC.update(large_image="notion_1024", large_text="Notion",
           small_image="keeb_512", small_text="⋆𐙚₊˚⊹♡!", start=int(time.time()),
           details=("Editing: " + "(っ- ‸ - ς)")) #("Editing: " + title)

while 1:
    time.sleep(15)  # Can only update presence every 15 seconds
