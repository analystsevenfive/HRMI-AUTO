<#
วนเปิดหน้า Time Attendance > บันทึกข้อมูลเวลา > ประมวลผลเวลาการทำงาน
ทีละองค์กรตามรหัสในรายการ HRMI และหยุดก่อนสั่งคำนวณจริง

ตั้ง environment variables user และ password ก่อนรัน -AutoLogin
#>
[CmdletBinding()]
param(
    [string]$ExePath = 'C:\Program Files (x86)\Prosoft\Prosoft HRMI\Bin\HRMI.exe',
    [int]$WaitSeconds = 30,
    [string[]]$Organizations = @('[BB]', '[BK]', '[NTS]', '[PK]', '[SM]'),
    [string]$LogPath = '',
    [switch]$CalibrateTimeAttendance,
    [switch]$AutoLogin
)

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($LogPath)) { $LogPath = Join-Path $PSScriptRoot 'logs' }
$script:LogFile = $null
$script:CurrentOrganization = ''
$script:CurrentStep = 'เริ่มต้น'

function Write-HrmiLog {
    param(
        [Parameter(Mandatory)][string]$Step,
        [Parameter(Mandatory)][string]$Status,
        [string]$Message = '',
        [string]$Organization = $script:CurrentOrganization,
        [string]$Details = ''
    )
    $entry = [PSCustomObject]@{
        Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
        Step = $Step
        Status = $Status
        Organization = $Organization
        Message = $Message
        Details = $Details
    }
    if ($script:LogFile) {
        $entry | Export-Csv -LiteralPath $script:LogFile -NoTypeInformation -Append -Encoding UTF8
    }
    $display = "[$($entry.Timestamp)] [$Status] [$Step]"
    if ($Organization) { $display += " [$Organization]" }
    if ($Message) { $display += " $Message" }
    if ($Status -eq 'ERROR') { Write-Warning $display } else { Write-Host $display }
}

function Start-HrmiStep {
    param([Parameter(Mandatory)][string]$Name, [string]$Message = '')
    $script:CurrentStep = $Name
    Write-HrmiLog -Step $Name -Status 'START' -Message $Message
}

function Complete-HrmiStep {
    param([Parameter(Mandatory)][string]$Name, [string]$Message = '')
    Write-HrmiLog -Step $Name -Status 'OK' -Message $Message
}

try {
    $logDirectory = [IO.Path]::GetFullPath($LogPath)
    if (-not (Test-Path -LiteralPath $logDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
    }
    $script:LogFile = Join-Path $logDirectory ("HRMI-{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
    Write-HrmiLog -Step 'เริ่มต้น' -Status 'INFO' -Message "เริ่มสคริปต์; Log: $script:LogFile"
} catch {
    throw "สร้างไฟล์ log ไม่สำเร็จ: $($_.Exception.Message)"
}

try {
if (-not (Test-Path -LiteralPath $ExePath -PathType Leaf)) { throw "ไม่พบ HRMI.exe: $ExePath" }

# โหลดค่า .env ข้างสคริปต์โดยไม่พิมพ์ค่าลงหน้าจอหรือ log
$envFile = Join-Path $PSScriptRoot '.env'
if (Test-Path -LiteralPath $envFile -PathType Leaf) {
    foreach ($line in Get-Content -LiteralPath $envFile -Encoding UTF8) {
        if ($line -match '^\s*(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$') {
            $key = $matches[1]
            $value = $matches[2]
            if ($value.Length -ge 2 -and (($value[0] -eq '"' -and $value[$value.Length - 1] -eq '"') -or ($value[0] -eq "'" -and $value[$value.Length - 1] -eq "'"))) {
                $value = $value.Substring(1, $value.Length - 2)
            }
            [Environment]::SetEnvironmentVariable($key, $value, 'Process')
        }
    }
}

$hrmiUser = if (-not [string]::IsNullOrWhiteSpace($env:HRMI_USER)) { $env:HRMI_USER } else { $env:user }
$hrmiPassword = if (-not [string]::IsNullOrWhiteSpace($env:HRMI_PASSWORD)) { $env:HRMI_PASSWORD } else { $env:password }
if ($AutoLogin -and ([string]::IsNullOrWhiteSpace($hrmiUser) -or [string]::IsNullOrWhiteSpace($hrmiPassword))) {
    throw 'กำหนด HRMI_USER และ HRMI_PASSWORD ใน environment ก่อนใช้ -AutoLogin (รองรับ user/password ด้วย)'
}
if ($AutoLogin -and $hrmiUser.Trim() -in @('ชื่อผู้ใช้', 'username', 'user name', 'ใส่ชื่อผู้ใช้')) {
    throw 'ค่า HRMI_USER ยังเป็นข้อความตัวอย่าง ไม่ใช่ชื่อเข้า HRMI จริง จึงหยุดก่อนกรอกข้อมูล'
}
if ($AutoLogin -and $hrmiPassword.Trim() -in @('รหัสผ่าน', 'password', 'ใส่รหัสผ่าน')) {
    throw 'ค่า HRMI_PASSWORD ยังเป็นข้อความตัวอย่าง ไม่ใช่รหัสผ่านจริง จึงหยุดก่อนกรอกข้อมูล'
}

$process = Get-Process -Name HRMI -ErrorAction SilentlyContinue |
    Where-Object { $_.Path -eq $ExePath } | Select-Object -First 1
if (-not $process) {
    Start-HrmiStep -Name 'เปิดโปรแกรม HRMI' -Message "Path: $ExePath"
    $process = Start-Process -FilePath $ExePath -WorkingDirectory (Split-Path -Parent $ExePath) -PassThru
    Write-Host 'กำลังเปิด Prosoft HRMI...'
    Complete-HrmiStep -Name 'เปิดโปรแกรม HRMI' -Message "ProcessId: $($process.Id)"
} else {
    Start-HrmiStep -Name 'เชื่อมต่อโปรแกรม HRMI ที่เปิดอยู่'
    Write-Host 'ใช้หน้าต่าง Prosoft HRMI ที่เปิดอยู่แล้ว'
    Complete-HrmiStep -Name 'เชื่อมต่อโปรแกรม HRMI ที่เปิดอยู่' -Message "ProcessId: $($process.Id)"
}

Add-Type @'
using System;
using System.Runtime.InteropServices;
public class HRMIStartupWindow {
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
    [DllImport("user32.dll", SetLastError=true)] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    public static IntPtr FindLargest(uint expectedPid) {
        IntPtr found = IntPtr.Zero; long largest = 0;
        EnumWindows((h, l) => { uint pid; GetWindowThreadProcessId(h, out pid); if (pid != expectedPid || !IsWindowVisible(h)) return true;
            RECT r; if (!GetWindowRect(h, out r)) return true; long area = (long)(r.Right-r.Left)*(r.Bottom-r.Top);
            if (area > largest) { largest=area; found=h; } return true; }, IntPtr.Zero);
        return found;
    }
}
'@

Start-HrmiStep -Name 'รอหน้าต่าง HRMI' -Message "Timeout: $WaitSeconds วินาที"
$windowDeadline = (Get-Date).AddSeconds($WaitSeconds)
do {
    try { $process.Refresh() } catch { }
    if ($process.HasExited) { throw 'โปรเซส HRMI ปิดตัวก่อนหน้าต่าง Login พร้อมใช้งาน' }
    $startupWindowHandle = [HRMIStartupWindow]::FindLargest([uint32]$process.Id)
    if ($startupWindowHandle -ne [IntPtr]::Zero) { break }
    Start-Sleep -Milliseconds 400
} while ((Get-Date) -lt $windowDeadline)
if ($startupWindowHandle -eq [IntPtr]::Zero) { throw 'รอหน้าต่าง HRMI ไม่สำเร็จ' }
$startupWindowInfo = Get-Process -Id $process.Id
Complete-HrmiStep -Name 'รอหน้าต่าง HRMI' -Message "พบ visible window handle: $startupWindowHandle; ProcessTitle: $($startupWindowInfo.MainWindowTitle)"

Add-Type -Namespace HRMI -Name Window -MemberDefinition @'
[System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
public struct POINT { public int X; public int Y; }
public struct MSLLHOOKSTRUCT { public POINT pt; public uint mouseData; public uint flags; public uint time; public System.IntPtr dwExtraInfo; }
public delegate bool EnumChildProc(System.IntPtr hWnd, System.IntPtr lParam);
public delegate bool EnumWindowsProc(System.IntPtr hWnd, System.IntPtr lParam);
public delegate System.IntPtr LowLevelMouseProc(int nCode, System.IntPtr wParam, System.IntPtr lParam);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool SetForegroundWindow(System.IntPtr hWnd);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern System.IntPtr GetForegroundWindow();
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool BringWindowToTop(System.IntPtr hWnd);
[System.Runtime.InteropServices.DllImport("user32.dll", SetLastError=true)]
public static extern bool SetWindowPos(System.IntPtr hWnd, System.IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint flags);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool ShowWindow(System.IntPtr hWnd, int command);
[System.Runtime.InteropServices.DllImport("user32.dll", SetLastError=true)]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);
[System.Runtime.InteropServices.DllImport("kernel32.dll")]
public static extern uint GetCurrentThreadId();
[System.Runtime.InteropServices.DllImport("user32.dll", SetLastError=true)]
public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool attach);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern void keybd_event(byte virtualKey, byte scanCode, uint flags, System.UIntPtr extraInfo);
public static bool ActivateWindow(System.IntPtr hWnd) {
    if (hWnd == System.IntPtr.Zero) return false;
    uint pid; uint targetThread = GetWindowThreadProcessId(hWnd, out pid);
    uint currentThread = GetCurrentThreadId();
    bool attached = currentThread != targetThread && AttachThreadInput(currentThread, targetThread, true);
    try {
        ShowWindow(hWnd, 9);
        BringWindowToTop(hWnd); SetForegroundWindow(hWnd);
        System.Threading.Thread.Sleep(200);
        if (GetForegroundWindow() != hWnd) {
            keybd_event(0x12, 0, 0, System.UIntPtr.Zero);
            try { BringWindowToTop(hWnd); SetForegroundWindow(hWnd); }
            finally { keybd_event(0x12, 0, 0x0002, System.UIntPtr.Zero); }
            System.Threading.Thread.Sleep(200);
        }
        return GetForegroundWindow() == hWnd;
    } finally {
        if (attached) AttachThreadInput(currentThread, targetThread, false);
    }
}
public static void RestoreWindowZOrder(System.IntPtr hWnd) {
    SetWindowPos(hWnd, new System.IntPtr(-2), 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0040);
}
public static string ClickWindowClient(System.IntPtr hWnd, int x, int y) {
    POINT p = new POINT(); p.X=x; p.Y=y;
    System.IntPtr target = hWnd;
    for (int i=0; i<16; i++) {
        System.IntPtr child = ChildWindowFromPointEx(target, p, 0x0001 | 0x0002 | 0x0004);
        if (child == System.IntPtr.Zero || child == target) break;
        if (!ClientToScreen(target, ref p) || !ScreenToClient(child, ref p)) break;
        target = child;
    }
    System.Text.StringBuilder cls = new System.Text.StringBuilder(256);
    GetClassName(target, cls, cls.Capacity);
    System.IntPtr lp = (System.IntPtr)((p.Y << 16) | (p.X & 0xffff));
    bool moved = PostMessage(target, 0x0200, System.IntPtr.Zero, lp);
    bool down = PostMessage(target, 0x0201, new System.IntPtr(1), lp);
    bool up = PostMessage(target, 0x0202, System.IntPtr.Zero, lp);
    return "targetHwnd=" + target + "; class=" + cls + "; client=" + p.X + "," + p.Y + "; posted=" + moved + "/" + down + "/" + up;
}
[System.Runtime.InteropServices.DllImport("user32.dll", CharSet=System.Runtime.InteropServices.CharSet.Unicode)]
public static extern int GetWindowText(System.IntPtr hWnd, System.Text.StringBuilder text, int count);
[System.Runtime.InteropServices.DllImport("user32.dll", CharSet=System.Runtime.InteropServices.CharSet.Unicode)]
public static extern System.IntPtr FindWindow(string className, string windowName);
[System.Runtime.InteropServices.DllImport("user32.dll", SetLastError=true)]
public static extern bool GetWindowRect(System.IntPtr hWnd, out RECT rect);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool SetCursorPos(int x, int y);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool GetCursorPos(out POINT point);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern short GetAsyncKeyState(int virtualKey);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern System.IntPtr WindowFromPoint(POINT point);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool ScreenToClient(System.IntPtr hWnd, ref POINT point);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern System.IntPtr ChildWindowFromPointEx(System.IntPtr hWnd, POINT point, uint flags);
[System.Runtime.InteropServices.DllImport("user32.dll", CharSet=System.Runtime.InteropServices.CharSet.Unicode)]
public static extern int GetClassName(System.IntPtr hWnd, System.Text.StringBuilder className, int maxCount);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, System.UIntPtr extraInfo);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool GetClientRect(System.IntPtr hWnd, out RECT rect);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool ClientToScreen(System.IntPtr hWnd, ref POINT point);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool SetProcessDPIAware();
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern System.IntPtr SetThreadDpiAwarenessContext(System.IntPtr dpiContext);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool IsWindowVisible(System.IntPtr hWnd);
 [System.Runtime.InteropServices.DllImport("user32.dll", CharSet=System.Runtime.InteropServices.CharSet.Auto, SetLastError=true)]
public static extern System.IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc callback, System.IntPtr module, uint threadId);
[System.Runtime.InteropServices.DllImport("user32.dll", CharSet=System.Runtime.InteropServices.CharSet.Auto, SetLastError=true)]
public static extern bool UnhookWindowsHookEx(System.IntPtr hook);
[System.Runtime.InteropServices.DllImport("user32.dll", CharSet=System.Runtime.InteropServices.CharSet.Auto)]
public static extern System.IntPtr CallNextHookEx(System.IntPtr hook, int code, System.IntPtr wParam, System.IntPtr lParam);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool EnumChildWindows(System.IntPtr parent, EnumChildProc callback, System.IntPtr lParam);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool EnumWindows(EnumWindowsProc callback, System.IntPtr lParam);
[System.Runtime.InteropServices.DllImport("user32.dll", CharSet=System.Runtime.InteropServices.CharSet.Unicode)]
public static extern int GetWindowTextLength(System.IntPtr hWnd);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern System.IntPtr SendMessage(System.IntPtr hWnd, uint msg, System.IntPtr wParam, System.IntPtr lParam);
[System.Runtime.InteropServices.DllImport("user32.dll", SetLastError=true)]
public static extern bool PostMessage(System.IntPtr hWnd, uint msg, System.IntPtr wParam, System.IntPtr lParam);
public static string ClickAtScreen(int x, int y, uint expectedPid) {
    POINT p = new POINT(); p.X = x; p.Y = y;
    System.IntPtr target = WindowFromPoint(p);
    uint targetPid;
    if (target == System.IntPtr.Zero || GetWindowThreadProcessId(target, out targetPid) == 0 || targetPid != expectedPid)
        throw new InvalidOperationException("ตำแหน่งคลิกไม่ได้อยู่บนหน้าต่าง HRMI");
    if (!SetCursorPos(x, y)) throw new InvalidOperationException("ย้าย cursor ไปยังตำแหน่งคลิกไม่ได้");
    System.Text.StringBuilder cls = new System.Text.StringBuilder(256);
    GetClassName(target, cls, cls.Capacity);
    if (!ScreenToClient(target, ref p)) return "hwnd=" + target + "; class=" + cls + "; screen-to-client-failed";
    System.IntPtr lp = (System.IntPtr)((p.Y << 16) | (p.X & 0xffff));
    mouse_event(0x0002, 0, 0, 0, System.UIntPtr.Zero);
    mouse_event(0x0004, 0, 0, 0, System.UIntPtr.Zero);
    return "hwnd=" + target + "; class=" + cls + "; client=" + p.X + "," + p.Y;
}
public static System.IntPtr FindProcessWindowContaining(uint expectedPid, string wanted) {
    System.IntPtr found = System.IntPtr.Zero;
    EnumWindows((hWnd, lParam) => {
        uint pid; GetWindowThreadProcessId(hWnd, out pid);
        if (pid != expectedPid || !IsWindowVisible(hWnd)) return true;
        int length = GetWindowTextLength(hWnd);
        if (length <= 0) return true;
        var text = new System.Text.StringBuilder(length + 1);
        GetWindowText(hWnd, text, text.Capacity);
        if (text.ToString().IndexOf(wanted, System.StringComparison.OrdinalIgnoreCase) >= 0) { found = hWnd; return false; }
        return true;
    }, System.IntPtr.Zero);
    return found;
}
public static string GetChildWindowText(System.IntPtr parent) {
    var values = new System.Collections.Generic.List<string>();
    EnumChildWindows(parent, (hWnd, lParam) => {
        int length = GetWindowTextLength(hWnd);
        if (length <= 0) return true;
        var text = new System.Text.StringBuilder(length + 1);
        GetWindowText(hWnd, text, text.Capacity);
        if (!System.String.IsNullOrWhiteSpace(text.ToString())) values.Add(text.ToString().Trim());
        return true;
    }, System.IntPtr.Zero);
    return System.String.Join(" | ", values.ToArray());
}
public static System.IntPtr FindLargestProcessWindow(uint expectedPid) {
    System.IntPtr found = System.IntPtr.Zero;
    long largestArea = 0;
    EnumWindows((hWnd, lParam) => {
        uint pid; GetWindowThreadProcessId(hWnd, out pid);
        if (pid != expectedPid || !IsWindowVisible(hWnd)) return true;
        RECT rect;
        if (!GetWindowRect(hWnd, out rect)) return true;
        long width = rect.Right - rect.Left, height = rect.Bottom - rect.Top;
        long area = width * height;
        if (area > largestArea) { largestArea = area; found = hWnd; }
        return true;
    }, System.IntPtr.Zero);
    return found;
}
public static System.IntPtr FindChildByText(System.IntPtr parent, string wanted) {
    System.IntPtr found = System.IntPtr.Zero;
    EnumChildWindows(parent, (hWnd, lParam) => {
        int length = GetWindowTextLength(hWnd);
        if (length <= 0) return true;
        var text = new System.Text.StringBuilder(length + 1);
        GetWindowText(hWnd, text, text.Capacity);
        if (System.String.Equals(text.ToString().Trim(), wanted, System.StringComparison.OrdinalIgnoreCase)) { found = hWnd; return false; }
        return true;
    }, System.IntPtr.Zero);
    return found;
}
public static System.IntPtr GetTopLevelWindow(System.IntPtr hWnd) {
    System.IntPtr current = hWnd;
    while (current != System.IntPtr.Zero) {
        System.IntPtr parent = GetParent(current);
        if (parent == System.IntPtr.Zero) return current;
        current = parent;
    }
    return hWnd;
}
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern System.IntPtr GetParent(System.IntPtr hWnd);
public static int SelectOrgAndConfirm(int rowIndex, uint expectedPid) {
    System.IntPtr dialog = FindProcessWindowContaining(expectedPid, "Organization");
    if (dialog == System.IntPtr.Zero) return 1;
    uint pid; GetWindowThreadProcessId(dialog, out pid);
    if (pid != expectedPid) return 2;
    RECT r; if (!GetWindowRect(dialog, out r)) return 3;
    int width = r.Right - r.Left, height = r.Bottom - r.Top;
    if (width < 250 || height < 180 || rowIndex < 0 || rowIndex > 4) return 4;
    if (!ActivateWindow(dialog)) return 6;
    int rowX = r.Left + (int)(width * 0.32);
    int rowY = r.Top + (int)(height * (0.245 + rowIndex * 0.072));
    ClickAtScreen(rowX, rowY, expectedPid);
    System.Threading.Thread.Sleep(200);
    System.IntPtr ok = FindChildByText(dialog, "OK");
    if (ok != System.IntPtr.Zero) SendMessage(ok, 0x00F5, System.IntPtr.Zero, System.IntPtr.Zero);
    System.Threading.Thread.Sleep(350);
    if (IsWindowVisible(dialog)) {
        // HRMI ใช้ปุ่ม WinForms แบบ custom; ถ้า BM_CLICK ไม่ปิด ให้คลิกตำแหน่งปุ่ม OK
        int okX = r.Left + (int)(width * 0.72), okY = r.Top + (int)(height * 0.91);
        ClickAtScreen(okX, okY, expectedPid);
    }
    for (int i=0; i<30 && IsWindowVisible(dialog); i++) System.Threading.Thread.Sleep(200);
    return IsWindowVisible(dialog) ? 5 : 0;
}
'@
[void][HRMI.Window]::SetProcessDPIAware()
[void][HRMI.Window]::SetThreadDpiAwarenessContext([IntPtr](-4))
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
$foregroundTitleBuffer = New-Object System.Text.StringBuilder 256
[void][HRMI.Window]::GetWindowText([HRMI.Window]::GetForegroundWindow(), $foregroundTitleBuffer, $foregroundTitleBuffer.Capacity)
$foregroundTitle = $foregroundTitleBuffer.ToString()
$organizationWindowAtStart = [HRMI.Window]::FindProcessWindowContaining([uint32]$process.Id, 'Organization')
$loginWindowAtStart = [HRMI.Window]::FindProcessWindowContaining([uint32]$process.Id, 'Login')
$atOrganizationPicker = ($organizationWindowAtStart -ne [IntPtr]::Zero)
$atLoginWindow = ($loginWindowAtStart -ne [IntPtr]::Zero)
$resumeFromAuthenticatedMain = (-not $atOrganizationPicker -and -not $atLoginWindow -and
    [HRMI.Window]::FindLargestProcessWindow([uint32]$process.Id) -ne [IntPtr]::Zero)

function Get-HrmiRoot {
    param([int]$TimeoutSeconds = 30)
    $until = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        try {
            $process.Refresh()
            $hwnd = [HRMI.Window]::FindLargestProcessWindow([uint32]$process.Id)
            if ($hwnd -ne [IntPtr]::Zero) { return [System.Windows.Automation.AutomationElement]::RootElement }
        } catch { }
        Start-Sleep -Milliseconds 400
    } while ((Get-Date) -lt $until)
    return $null
}

function Find-HrmiText {
    param($Root, [string]$Text)
    if (-not $Root) { return @() }
    try {
        $condition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ProcessIdProperty, [int]$process.Id)
        $items = $Root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)
        $matches = @()
        foreach ($item in $items) {
            try {
                $name = $item.Current.Name -replace '\s+', ''
                $wanted = $Text -replace '\s+', ''
                if ($name -and $name.Contains($wanted)) { $matches += $item }
            } catch { }
        }
        return $matches
    } catch { return @() }
}

function Invoke-HrmiControl {
    param($Element)
    foreach ($pattern in @([System.Windows.Automation.SelectionItemPattern]::Pattern,
                           [System.Windows.Automation.InvokePattern]::Pattern)) {
        try {
            $p = $Element.GetCurrentPattern($pattern)
            if ($p -is [System.Windows.Automation.SelectionItemPattern]) { $p.Select() } else { $p.Invoke() }
            return $true
        } catch { }
    }
    return $false
}

function Invoke-HrmiNavigation {
    param([string]$Text)
    $alert = [HRMI.Window]::FindProcessWindowContaining([uint32]$process.Id, 'Alert')
    if ($alert -ne [IntPtr]::Zero) {
        $alertText = [HRMI.Window]::GetChildWindowText($alert)
        Write-HrmiLog -Step 'ปิดหน้าต่าง Alert' -Status 'START' -Message 'ปิดหน้าต่างแจ้งเตือนที่บังแผงเมนู' -Details $alertText
        [void][HRMI.Window]::SendMessage($alert, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)
        Start-Sleep -Milliseconds 400
        if ([HRMI.Window]::IsWindowVisible($alert)) { throw "ปิดหน้าต่าง Alert ไม่สำเร็จ: $alertText" }
        Write-HrmiLog -Step 'ปิดหน้าต่าง Alert' -Status 'OK' -Message 'ปิดหน้าต่างแจ้งเตือนแล้ว' -Details $alertText
    }
    $root = Get-HrmiRoot
    $controls = @(Find-HrmiText -Root $root -Text $Text)
    foreach ($candidate in $controls) {
        try {
            $b = $candidate.Current.BoundingRectangle
            Write-HrmiLog -Step 'ตรวจตำแหน่งเมนู' -Status 'INFO' -Message $Text -Details ("name={0}; rect=({1},{2},{3},{4}); type={5}" -f $candidate.Current.Name,$b.Left,$b.Top,$b.Width,$b.Height,$candidate.Current.ControlType.ProgrammaticName)
        } catch { }
    }
    $control = $controls | Where-Object {
        try { [string]::Equals($_.Current.Name.Trim(), $Text, [System.StringComparison]::OrdinalIgnoreCase) } catch { $false }
    } | Select-Object -First 1
    if ($control -and (Invoke-HrmiControl -Element $control)) { return 'UIAutomation' }

    if ($control) {
        Write-HrmiLog -Step 'ตรวจพบชื่อเมนู' -Status 'INFO' -Message $Text -Details 'พบ label ใน UIA แต่ไม่มี InvokePattern; จะไม่คลิกจาก bounding rectangle ที่ยังไม่ยืนยัน'
    }
    if (-not $CalibrateTimeAttendance) {
        throw 'เมนู Time Attendance เป็น custom control ที่ UI Automation คลิกไม่ได้; ใช้ -CalibrateTimeAttendance เพื่อบันทึกการคลิกจริงหนึ่งครั้ง'
    }

    $hwnd = [HRMI.Window]::FindLargestProcessWindow([uint32]$process.Id)
    if ($hwnd -eq [IntPtr]::Zero) { throw 'ไม่พบหน้าต่าง HRMI สำหรับ calibration' }
    $configPath = Join-Path $PSScriptRoot 'hrmi-navigation.json'
    $existingConfig = $null
    if (Test-Path -LiteralPath $configPath -PathType Leaf) {
        try { $existingConfig = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch { }
    }
    $navigator = @(Find-HrmiText -Root (Get-HrmiRoot) -Text 'Navigator Explorer Bar' | Where-Object {
        try { -not $_.Current.IsOffscreen } catch { $false }
    }) | Select-Object -First 1
    $navBoundsBeforeClick = if ($navigator) { $navigator.Current.BoundingRectangle } else { $null }
    if ($existingConfig -and $existingConfig.Version -eq 3 -and $existingConfig.NavigatorRelative) {
        try {
            if (-not $navigator) { throw 'ไม่พบแถบ Navigator' }
            $navBounds = $navigator.Current.BoundingRectangle
            $relativeX = [double]$existingConfig.NavigatorRelative.X
            $relativeY = [double]$existingConfig.NavigatorRelative.Y
            if ($relativeX -lt 0 -or $relativeX -gt 1 -or $relativeY -lt 0 -or $relativeY -gt 1) { throw 'ค่าพิกัดสัมพัทธ์ไม่ถูกต้อง' }
            $point = New-Object HRMI.Window+POINT
            $point.X = [int]($navBounds.Left + $navBounds.Width * $relativeX)
            $point.Y = [int]($navBounds.Top + $navBounds.Height * $relativeY)
            if (-not [HRMI.Window]::ActivateWindow($hwnd)) { throw 'นำ HRMI มาไว้ด้านหน้าไม่ได้' }
            $hit = [HRMI.Window]::ClickAtScreen($point.X, $point.Y, [uint32]$process.Id)
            Start-Sleep -Milliseconds 650
            if (-not (Wait-HrmiVisibleText -Text 'Time Attendance Flow' -TimeoutSeconds 3)) { throw "ไม่พบหน้า Time Attendance Flow หลังคลิก; $hit" }
            return "พิกัดสัมพัทธ์ Navigator=($relativeX,$relativeY); $hit"
        } catch {
            Write-HrmiLog -Step 'ตรวจพิกัด Time Attendance' -Status 'INFO' -Message "พิกัดเดิมใช้ไม่ได้; ปรับเทียบใหม่: $($_.Exception.Message)"
        }
    } elseif ($existingConfig) {
        Write-HrmiLog -Step 'ตรวจพิกัด Time Attendance' -Status 'INFO' -Message 'พิกัดรูปแบบเก่า; ปรับเทียบใหม่'
    }

    $bounds = New-Object HRMI.Window+RECT
    if (-not [HRMI.Window]::GetWindowRect($hwnd, [ref]$bounds)) { throw 'อ่านตำแหน่งหน้าต่าง HRMI ไม่ได้' }
    Write-Host ''
    Write-Host 'Calibration: คลิกเมนู Time Attendance ในหน้าต่าง HRMI ด้วยตนเองหนึ่งครั้ง' -ForegroundColor Yellow
    Write-Host 'ตัวจับจะเก็บพิกัดเฉพาะเมื่อคลิกลงบนหน้าต่างหลัก HRMI; จากนั้นรอเมนูย่อยปรากฏ'
    $script:CalibrationClick = $null
    $deadline = (Get-Date).AddSeconds(120)
    $wasDown = ([HRMI.Window]::GetAsyncKeyState(1) -band 0x8000) -ne 0
    while (-not $script:CalibrationClick -and (Get-Date) -lt $deadline) {
        $isDown = ([HRMI.Window]::GetAsyncKeyState(1) -band 0x8000) -ne 0
        if ($isDown -and -not $wasDown) {
            $point = New-Object HRMI.Window+POINT
            if ([HRMI.Window]::GetCursorPos([ref]$point)) {
                $hit = [HRMI.Window]::WindowFromPoint($point)
                $hitPid = [uint32]0
                if ($hit -ne [IntPtr]::Zero -and
                    [HRMI.Window]::GetWindowThreadProcessId($hit, [ref]$hitPid) -ne 0 -and
                    $hitPid -eq [uint32]$process.Id) {
                    $script:CalibrationClick = $point
                }
            }
        }
        $wasDown = $isDown
        Start-Sleep -Milliseconds 20
    }
    if (-not $script:CalibrationClick) { throw 'หมดเวลา 120 วินาที: ยังไม่ได้รับคลิกบนหน้าต่างหลัก HRMI' }
    Start-Sleep -Milliseconds 650
    $checkRoot = Get-HrmiRoot -TimeoutSeconds 5
    $checkMenu = @(Find-HrmiText -Root $checkRoot -Text 'Time Attendance Flow' | Where-Object {
        try { $_.Current.IsOffscreen -eq $false } catch { $false }
    }) | Select-Object -First 1
    if (-not $checkMenu) { throw 'Calibration ไม่ผ่าน: หลังผู้ใช้คลิกแล้ว UIA ยังไม่พบหน้า Time Attendance Flow; ยังไม่บันทึกพิกัด' }
    $screenPoint = $script:CalibrationClick
    if (-not $navigator) { throw 'ไม่พบแถบ Navigator สำหรับบันทึกพิกัดสัมพัทธ์' }
    $navBounds = $navBoundsBeforeClick
    if ($navBounds.Width -le 0 -or $navBounds.Height -le 0) { throw 'ขนาดแถบ Navigator ไม่ถูกต้อง' }
    $relativeX = ($screenPoint.X - $navBounds.Left) / $navBounds.Width
    $relativeY = ($screenPoint.Y - $navBounds.Top) / $navBounds.Height
    if ($relativeX -lt 0 -or $relativeX -gt 1 -or $relativeY -lt 0 -or $relativeY -gt 1) {
        throw 'จุดที่คลิกไม่อยู่ในแถบ Navigator; ยังไม่บันทึกพิกัด'
    }
    $config = [ordered]@{
        Version = 3
        NavigatorRelative = [ordered]@{ X = $relativeX; Y = $relativeY }
        VerifiedSubmenu = 'Time Attendance Flow'
        CalibratedAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss.fffK')
    }
    $config | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $configPath -Encoding UTF8
    Write-HrmiLog -Step 'ปรับเทียบเมนู Time Attendance' -Status 'OK' -Message 'บันทึกตำแหน่งสัมพัทธ์หลังยืนยันหน้า Flow ปรากฏ' -Details ("navigator=({0},{1}); config={2}" -f $relativeX,$relativeY,$configPath)
    return "ผู้ใช้คลิกและยืนยันหน้า Flow แล้ว; บันทึกตำแหน่งสัมพัทธ์ Navigator=($relativeX,$relativeY)"
}

function Invoke-HrmiNamedControl {
    param($Root, [string]$Text)
    $control = @(Find-HrmiText -Root $Root -Text $Text) | Select-Object -First 1
    if (-not $control) {
        Write-HrmiLog -Step 'ค้นหา control' -Status 'ERROR' -Message "ไม่พบ $Text ใน UI Automation"
        return $false
    }
    if (Invoke-HrmiControl -Element $control) { return $true }
    try {
        $bounds = $control.Current.BoundingRectangle
        if ($bounds.IsEmpty -or $bounds.Width -le 4 -or $bounds.Height -le 4) { return $false }
        $x = [int]($bounds.Left + ($bounds.Width / 2))
        $y = [int]($bounds.Top + ($bounds.Height / 2))
        $hwnd = [HRMI.Window]::FindProcessWindowContaining([uint32]$process.Id, 'Time Attendance Flow')
        if ($hwnd -eq [IntPtr]::Zero) { $hwnd = [HRMI.Window]::FindLargestProcessWindow([uint32]$process.Id) }
        [void][HRMI.Window]::ActivateWindow($hwnd)
        Start-Sleep -Milliseconds 200
        [void][HRMI.Window]::ClickAtScreen($x, $y, [uint32]$process.Id)
        Start-Sleep -Milliseconds 300
        return $true
    } catch {
        Write-HrmiLog -Step 'คลิก control' -Status 'ERROR' -Message "$Text : $($_.Exception.Message)"
        return $false
    }
}

function Wait-HrmiVisibleText {
    param([string]$Text, [int]$TimeoutSeconds = 8)
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $root = Get-HrmiRoot -TimeoutSeconds 1
        $visible = @(Find-HrmiText -Root $root -Text $Text | Where-Object {
            try { -not $_.Current.IsOffscreen } catch { $false }
        })
        if ($visible.Count -gt 0) { return $true }
        Start-Sleep -Milliseconds 300
    } while ((Get-Date) -lt $deadline)
    return $false
}

function Wait-HrmiProcessingWindow {
    param([int]$TimeoutSeconds = 8)
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $root = Get-HrmiRoot -TimeoutSeconds 1
        $windows = @(Find-HrmiText -Root $root -Text 'ประมวลผลเวลาการทำงาน' | Where-Object {
            try {
                $_.Current.ControlType -eq [System.Windows.Automation.ControlType]::Window -and
                -not $_.Current.IsOffscreen
            } catch { $false }
        })
        if ($windows.Count -gt 0) { return $true }
        Start-Sleep -Milliseconds 300
    } while ((Get-Date) -lt $deadline)
    return $false
}

function Send-HrmiLogin {
    Start-HrmiStep -Name 'เข้าสู่ระบบ HRMI'
    $process.Refresh()
    $loginHwnd = [HRMI.Window]::FindProcessWindowContaining([uint32]$process.Id, 'Login')
    if ($loginHwnd -eq [IntPtr]::Zero) { $loginHwnd = $process.MainWindowHandle }
    if ($loginHwnd -eq 0) { throw 'ไม่พบหน้าต่าง HRMI สำหรับ Login' }
    $activated = $false
    for ($attempt = 1; $attempt -le 5 -and -not $activated; $attempt++) {
        $activated = [HRMI.Window]::ActivateWindow($loginHwnd)
        if (-not $activated) { Start-Sleep -Milliseconds 300 }
    }
    if (-not $activated) { throw 'ไม่สามารถนำหน้าต่าง Login ของ HRMI มาไว้ด้านหน้า หลังลอง 5 ครั้ง จึงไม่ส่งข้อมูลรับรอง' }
    Start-Sleep -Milliseconds 400
    # ในฟอร์ม HRMI จุดเริ่มโฟกัสคือปุ่ม Login; ย้อนสองช่องไป Password,
    # กรอกแล้วถอยไป User จากนั้น Tab ไปปุ่ม Login เพื่อส่งฟอร์ม
    [System.Windows.Forms.SendKeys]::SendWait('+{TAB}+{TAB}')
    [System.Windows.Forms.SendKeys]::SendWait('^a')
    [System.Windows.Forms.SendKeys]::SendWait((ConvertTo-SendKeysLiteral $hrmiPassword))
    [System.Windows.Forms.SendKeys]::SendWait('+{TAB}')
    [System.Windows.Forms.SendKeys]::SendWait('^a')
    [System.Windows.Forms.SendKeys]::SendWait((ConvertTo-SendKeysLiteral $hrmiUser))
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}{TAB}{ENTER}')
    $deadline = (Get-Date).AddSeconds([Math]::Max(10, $WaitSeconds))
    $state = 'รอผล Login'
    do {
        Start-Sleep -Milliseconds 500
        $process.Refresh()
        $titleBuffer = New-Object System.Text.StringBuilder 256
        [void][HRMI.Window]::GetWindowText([HRMI.Window]::GetForegroundWindow(), $titleBuffer, $titleBuffer.Capacity)
        $title = $titleBuffer.ToString()
        $organizationHwnd = [HRMI.Window]::FindProcessWindowContaining([uint32]$process.Id, 'Organization')
        if ($organizationHwnd -ne [IntPtr]::Zero) { $state = 'Organization'; break }
        $alertHwnd = [HRMI.Window]::FindProcessWindowContaining([uint32]$process.Id, 'Alert')
        if ($alertHwnd -eq [IntPtr]::Zero) { $alertHwnd = [HRMI.Window]::FindProcessWindowContaining([uint32]$process.Id, 'Information') }
        if ($alertHwnd -ne [IntPtr]::Zero) { $state = 'Alert'; break }
        $loginWindow = [HRMI.Window]::FindProcessWindowContaining([uint32]$process.Id, 'Login')
        if ($loginWindow -ne [IntPtr]::Zero) { $state = 'Login'; continue }
        if ($title -match 'Organization') { $state = 'Organization'; break }
        if ($title -match 'Alert') { $state = 'Alert'; break }
        if ($title -match 'Login') { $state = 'Login'; continue }
    } while ((Get-Date) -lt $deadline)
    if ($state -eq 'Alert') {
        $alertHwnd = [HRMI.Window]::FindProcessWindowContaining([uint32]$process.Id, 'Alert')
        if ($alertHwnd -eq [IntPtr]::Zero) { $alertHwnd = [HRMI.Window]::FindProcessWindowContaining([uint32]$process.Id, 'Information') }
        $alertText = ''
        if ($alertHwnd -ne [IntPtr]::Zero) {
            $alertText = [HRMI.Window]::GetChildWindowText($alertHwnd)
            try {
                $alertRoot = [System.Windows.Automation.AutomationElement]::FromHandle($alertHwnd)
                $alertItems = @($alertRoot.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition))
                $uiaText = @($alertItems | ForEach-Object { try { $_.Current.Name } catch { '' } } | Where-Object { $_ }) -join ' | '
                if ($uiaText) { $alertText = (@($alertText, $uiaText | Where-Object { $_ }) -join ' | ') }
            } catch { }
        }
        Write-HrmiLog -Step 'เข้าสู่ระบบ HRMI' -Status 'ERROR' -Message 'HRMI แสดง Alert หลังส่ง Login' -Details $alertText
        throw "Login ถูกขัดจังหวะด้วย Alert: $alertText"
    }
    if ($state -ne 'Organization') {
        Write-HrmiLog -Step 'เข้าสู่ระบบ HRMI' -Status 'ERROR' -Message "ไม่ได้ไปถึงหน้าเลือกองค์กร; state=$state; mainTitle=$($process.MainWindowTitle)"
        throw "ตรวจไม่พบหน้าเลือกองค์กรหลัง Login (state=$state; title=$($process.MainWindowTitle))"
    }
    Complete-HrmiStep -Name 'เข้าสู่ระบบ HRMI' -Message 'ตรวจพบหน้า Organization แล้ว (ไม่ได้บันทึกข้อมูลรับรอง)'
}

function ConvertTo-SendKeysLiteral {
    param([string]$Value)
    $builder = New-Object System.Text.StringBuilder
    foreach ($character in $Value.ToCharArray()) {
        $key = switch ([string]$character) {
            '+'  { '{+}'; break }
            '^'  { '{^}'; break }
            '%'  { '{%}'; break }
            '~'  { '{~}'; break }
            '('  { '{(}'; break }
            ')'  { '{)}'; break }
            '{'  { '{{}'; break }
            '}'  { '{}}'; break }
            default { [string]$character }
        }
        [void]$builder.Append($key)
    }
    return $builder.ToString()
}

function Select-HrmiOrganization {
    param([string]$Code)
    $script:CurrentOrganization = $Code
    Start-HrmiStep -Name 'เลือกองค์กร' -Message "แถวรายการ: $Code"
    $rowIndex = switch ($Code) {
        '[BB]' { 0 }
        '[BK]' { 1 }
        '[NTS]' { 2 }
        '[PK]' { 3 }
        '[SM]' { 4 }
        default { throw "ไม่รู้จักรหัสองค์กร $Code" }
    }
    $result = [HRMI.Window]::SelectOrgAndConfirm($rowIndex, [uint32]$process.Id)
    if ($result -ne 0) {
        $reasons = @('สำเร็จ', 'ไม่พบหน้าต่าง Organization', 'หน้าต่างไม่ใช่ของ HRMI', 'อ่านขนาดหน้าต่างไม่ได้', 'ขนาดหรือแถวไม่ถูกต้อง', 'ปุ่ม OK ไม่ปิดหน้าต่าง', 'นำหน้าต่าง Organization มาไว้ด้านหน้าไม่ได้')
        throw "เลือกองค์กร $Code ไม่สำเร็จ: $($reasons[$result]) (code $result)"
    }
    Write-Host "เลือกและยืนยันองค์กร $Code แล้ว"
    Start-Sleep -Seconds 2
    Complete-HrmiStep -Name 'เลือกองค์กร' -Message "ยืนยันแถวสำหรับ $Code แล้ว"
}

function Open-HrmiProcessingPage {
    Start-HrmiStep -Name 'เปิดเมนู Time Attendance'
    $navigationMethod = Invoke-HrmiNavigation -Text 'Time Attendance'
    if (-not (Wait-HrmiVisibleText -Text 'Time Attendance Flow')) { throw 'เปิด Time Attendance แล้ว แต่ไม่พบหน้า Time Attendance Flow' }
    Complete-HrmiStep -Name 'เปิดเมนู Time Attendance' -Message "วิธีคลิก: $navigationMethod"

    Start-HrmiStep -Name 'เปิดหน้าประมวลผลเวลาการทำงาน'
    $root = Get-HrmiRoot
    if (-not (Invoke-HrmiNamedControl -Root $root -Text 'ประมวลผลเวลาการทำงาน')) { throw 'เปิดหน้าประมวลผลเวลาการทำงานไม่สำเร็จ: หา control ไม่พบหรือไม่มีตำแหน่งคลิก' }
    if (-not (Wait-HrmiProcessingWindow)) { throw 'ไม่พบหน้าต่างประมวลผลเวลาการทำงานหลังส่งคำสั่งเปิด' }
    Complete-HrmiStep -Name 'เปิดหน้าประมวลผลเวลาการทำงาน' -Message 'พบข้อความหน้าประมวลผล; ยังไม่ได้กดเริ่มคำนวณ'
}

if ($AutoLogin -and $atLoginWindow) {
    $script:CurrentStep = 'เข้าสู่ระบบ HRMI'
    Write-Host 'กำลังเข้าสู่ระบบ HRMI ด้วยข้อมูลจาก environment variables...'
    Send-HrmiLogin
} elseif ($AutoLogin -and $atOrganizationPicker) {
    $script:CurrentStep = 'ตรวจสอบหน้าเลือกองค์กร'
    Write-HrmiLog -Step $script:CurrentStep -Status 'OK' -Message 'พบหน้าเลือกองค์กรที่เปิดค้างอยู่; ข้ามการ Login'
    Write-Host 'พบหน้าเลือกองค์กรที่เปิดค้างอยู่ จะทำต่อจากหน้านี้โดยไม่กรอก Login ซ้ำ'
} elseif ($resumeFromAuthenticatedMain) {
    throw 'HRMI เปิดอยู่ที่หน้าหลัก แต่ตรวจยืนยันองค์กรปัจจุบันไม่ได้; กลับไปหน้าเลือกองค์กรก่อนรันเพื่อป้องกันการระบุองค์กรผิด'
} elseif (-not $atOrganizationPicker -and ($foregroundTitle -match 'Login' -or $process.MainWindowTitle -match 'Login')) {
    $script:CurrentStep = 'รอ Login ด้วยตนเอง'
    Write-HrmiLog -Step $script:CurrentStep -Status 'ERROR' -Message 'HRMI อยู่หน้า Login; รันใหม่ด้วย -AutoLogin หลังตั้ง HRMI_USER และ HRMI_PASSWORD หรือ Login ด้วยตนเองก่อนรัน'
    throw 'HRMI ยังอยู่หน้า Login จึงไม่สามารถเลือกองค์กรได้'
}

$results = @()
for ($i = 0; $i -lt $Organizations.Count; $i++) {
    $org = $Organizations[$i]
    $script:CurrentOrganization = $org
    Write-HrmiLog -Step 'เริ่มองค์กร' -Status 'START' -Message "องค์กร $($i + 1) จาก $($Organizations.Count)"
    Write-Host "[$($i + 1)/$($Organizations.Count)] เลือกองค์กร $org"
    try {
        Select-HrmiOrganization -Code $org
        Open-HrmiProcessingPage
        Write-Host "เปิดหน้าประมวลผลของ $org แล้ว (ยังไม่ได้กดเริ่มคำนวณ)"
        Write-HrmiLog -Step 'ประมวลผลองค์กร' -Status 'OK' -Message 'เปิดหน้าประมวลผลแล้ว; ยังไม่ได้เริ่มคำนวณ'
        $results += [PSCustomObject]@{ Organization = $org; Status = 'เปิดหน้าประมวลผลแล้ว' }
    } catch {
        Write-HrmiLog -Step $script:CurrentStep -Status 'ERROR' -Message $_.Exception.Message -Details $_.ScriptStackTrace
        Write-Warning "หยุดที่องค์กร $org : $($_.Exception.Message)"
        $results += [PSCustomObject]@{ Organization = $org; Status = "ผิดพลาด: $($_.Exception.Message)" }
        throw
    }

    if ($i -lt ($Organizations.Count - 1)) {
        # เมนู File > Login ใน HRMI ใช้กลับไปหน้าเลือกองค์กร
        Start-HrmiStep -Name 'กลับไปหน้าเลือกองค์กร' -Message 'ส่ง File > Login'
        [void][HRMI.Window]::SetForegroundWindow($process.MainWindowHandle)
        Start-Sleep -Milliseconds 300
        [System.Windows.Forms.SendKeys]::SendWait('%f')
        Start-Sleep -Milliseconds 300
        [System.Windows.Forms.SendKeys]::SendWait('l')
        Start-Sleep -Seconds 2
        $returnRoot = Get-HrmiRoot -TimeoutSeconds $WaitSeconds
        $returnTitle = ''
        try { $returnTitle = $returnRoot.Current.Name } catch { }
        if ($returnTitle -match 'Login') {
            Write-Host 'HRMI ขอ Login อีกครั้งหลัง File > Login'
            Send-HrmiLogin
        }
        Complete-HrmiStep -Name 'กลับไปหน้าเลือกองค์กร' -Message "หน้าต่าง: $returnTitle"
    }
}

Write-Host ''
$results | Format-Table -AutoSize
Write-Host 'สคริปต์ไม่กดปุ่มเริ่มประมวลผล: ตรวจสอบบริษัทและงวดบนหน้าจอก่อนคำนวณ'
Write-HrmiLog -Step 'จบสคริปต์' -Status 'INFO' -Message "เสร็จ; องค์กรที่สรุป: $($results.Count)"
$script:CurrentStep = 'จบสคริปต์'
} catch {
    Write-HrmiLog -Step $script:CurrentStep -Status 'FATAL' -Message $_.Exception.Message -Details $_.ScriptStackTrace
    throw
}
