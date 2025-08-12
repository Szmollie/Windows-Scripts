
# Load the MoveWindow function from user32.dll
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class User32 {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
}
"@

# List of VNC files to open
$vncFiles = @(
    "ASB_I4_2534N.vnc",
    "ASB_I5_2535N.vnc",
    "ASB_II5_2511N.vnc",
    "ASB_II6_2510N.vnc",
    "ER_III5_2505N.vnc",
    "ER_III6_2507N.vnc",
    "ER_IV_5_2537N.vnc",
    "FL_I1_2506N.vnc",
    "FL_I3_2504N.vnc",
    "FL_V1_2536N.vnc",
    "RT_I4_2540N.vnc",
    "RT_II2_2528N.vnc",
    "RT_II4_2529N.vnc",
    "RT_I2_2501N.vnc"
    # Add the rest of your 14 files here
)

# Define window positions (adjust as needed for your screen layout)
$positions = @(
    @{X=0; Y=0},
    @{X=350; Y=0},
    @{X=700; Y=0},
    @{X=0; Y=300},
    @{X=350; Y=300},
    @{X=700; Y=300},
    @{X=0; Y=600},
    @{X=350; Y=600},
    @{X=700; Y=600},
    @{X=0; Y=900},
    @{X=350; Y=900},
    @{X=700; Y=900},
    @{X=0; Y=1200},
    @{X=350; Y=1200},
    @{X=700; Y=1200}
    

    # Add more positions for additional windows
)

# Launch and arrange each VNC session
for ($i = 0; $i -lt $vncFiles.Count; $i++) {
    $filePath = "T:\misc\VNC_Login_List\$($vncFiles[$i])"
    $proc = Start-Process $filePath -PassThru
    Start-Sleep -Seconds 6

    # Refresh the process to ensure MainWindowHandle is available
    $proc.Refresh()

    if ($proc.MainWindowHandle -ne 0) {
        $pos = $positions[$i]
        [User32]::MoveWindow($proc.MainWindowHandle, $pos.X, $pos.Y, 370, 320, $true)
    } else {
        Write-Host "Window handle not found for $($vncFiles[$i])"
    }
}
