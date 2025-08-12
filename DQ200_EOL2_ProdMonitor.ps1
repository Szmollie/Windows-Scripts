####################################################################################################
# Open DAD2506N
Start-Process "cmd.exe" -ArgumentList '/c "C:\Program Files\SolarWinds\Dameware Mini Remote Control x64\DWRCC.exe" -c: -h: -m:DAD2506N.ot1.vitesco.com -u:svo14419 -p:DebrecenTesteng2025+ -d:OT1 -a:2 -prxa:10.116.119.98 -prxp:6127' -WindowStyle Hidden

# Wait for Excel to open
Start-Sleep -Seconds 10

# Get the window ID of Excel
$window1 = Get-Process | Where-Object { $_.MainWindowTitle -like "*Dameware*" }

# Move the window to a specific position (e.g., x=100, y=100)
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class User32 {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
}
"@
# ($window.MainWindowHandle, Xpos_Screen[map:(left)-5<y<2000(right)],Ypos_Screen[map:(top)600<y<2000(bottom)],x,WinHeight[map:(short)500<y<1200(tall)]
[User32]::MoveWindow($window1[0].MainWindowHandle, 0, 0, 600, 500, $true)


####################################################################################################
# Open DAD2536N
Start-Process "cmd.exe" -ArgumentList '/c "C:\Program Files\SolarWinds\Dameware Mini Remote Control x64\DWRCC.exe" -c: -h: -m:DAD2536N.ot1.vitesco.com -u:svo14419 -p:DebrecenTesteng2025+ -d:OT1 -a:2 -prxa:10.116.119.98 -prxp:6127' -WindowStyle Hidden

# Wait for Excel to open
Start-Sleep -Seconds 10

# Get the window ID of Excel
$window2 = Get-Process | Where-Object { $_.MainWindowTitle -like "*Dameware*" }

# Move the window to a specific position (e.g., x=100, y=100)

####################################################################################################
# Open DAD2501N
Start-Process "cmd.exe" -ArgumentList '/c "C:\Program Files\SolarWinds\Dameware Mini Remote Control x64\DWRCC.exe" -c: -h: -m:DAD2501N.ot1.vitesco.com -u:svo14419 -p:DebrecenTesteng2025+ -d:OT1 -a:2 -prxa:10.116.119.98 -prxp:6127' -WindowStyle Hidden

# Wait for Excel to open
Start-Sleep -Seconds 10

# Get the window ID of Excel
$window3 = Get-Process | Where-Object { $_.MainWindowTitle -like "*DameWare*" }

# Move the window to a specific position (e.g., x=100, y=100)
# ($window.MainWindowHandle, Xpos_Screen[map:(left)-5<y<2000(right)],Ypos_Screen[map:(top)600<y<2000(bottom)],x,WinHeight[map:(short)500<y<1200(tall)]
[User32]::MoveWindow($window3[2].MainWindowHandle, 1200, 0, 600, 500, $true)


####################################################################################################
# Open DAD2507N
Start-Process "cmd.exe" -ArgumentList '/c "C:\Program Files\SolarWinds\Dameware Mini Remote Control x64\DWRCC.exe" -c: -h: -m:DAD2507N.ot1.vitesco.com -u:svo14419 -p:DebrecenTesteng2025+ -d:OT1 -a:2 -prxa:10.116.119.98 -prxp:6127' -WindowStyle Hidden

# Wait for Excel to open
Start-Sleep -Seconds 10

# Get the window ID of Excel
$window4 = Get-Process | Where-Object { $_.MainWindowTitle -like "*DameWare*" }

# Move the window to a specific position (e.g., x=100, y=100)
# ($window.MainWindowHandle, Xpos_Screen[map:(left)-5<y<2000(right)],Ypos_Screen[map:(top)600<y<2000(bottom)],x,WinHeight[map:(short)500<y<1200(tall)]
[User32]::MoveWindow($window4[3].MainWindowHandle, 1200, 520, 600, 500, $true)








<#
# Define MoveWindow once
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class User32 {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
}
"@

function Start-DamewareSession {
    param (
        [string]$machine,
        [int]$x,
        [int]$y
    )

    $args = "/c `"C:\Program Files\SolarWinds\Dameware Mini Remote Control x64\DWRCC.exe`" -c: -h: -m:$machine -u:svo14419 -p:DebrecenTesteng2025+ -d:OT1 -a:2 -prxa:10.116.119.98 -prxp:6127"
    $proc = Start-Process "cmd.exe" -ArgumentList $args -WindowStyle Hidden -PassThru
    Start-Sleep -Seconds 10
    $proc.Refresh()
    if ($proc.MainWindowHandle -ne 0) {
        [User32]::MoveWindow($proc.MainWindowHandle, $x, $y, 600, 500, $true)
    } else {
        Write-Host "Could not find window for $machine"
    }
}

# Launch sessions
Start-DamewareSession -machine "DAD2506N.ot1.vitesco.com" -x 0 -y 0
Start-DamewareSession -machine "DAD2536N.ot1.vitesco.com" -x 600 -y 0
Start-DamewareSession -machine "DAD2501N.ot1.vitesco.com" -x 1200 -y 0
Start-DamewareSession -machine "DAD2507N.ot1.vitesco.com" -x 1200 -y 520

#>
