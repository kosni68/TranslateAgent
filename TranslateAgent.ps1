param(
    [string]$HostIp = "127.0.0.1",
    [int]$Port = 1234,
    [string]$Lang = "de"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Level,
        [Parameter(Mandatory=$true)][string]$Message,
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )
    $ts = Get-Date -Format "HH:mm:ss"
    Write-Host "[$ts][$Level] $Message" -ForegroundColor $Color
}

function New-Stopwatch {
    return [System.Diagnostics.Stopwatch]::StartNew()
}

function New-TimingSession {
    return @{
        Total = [System.Diagnostics.Stopwatch]::StartNew()
        Last  = [System.Diagnostics.Stopwatch]::StartNew()
    }
}

function Log-Time {
    param(
        [Parameter(Mandatory=$true)]$Timing,
        [Parameter(Mandatory=$true)][string]$Message
    )
    $delta = $Timing.Last.ElapsedMilliseconds
    $total = $Timing.Total.ElapsedMilliseconds
    $Timing.Last.Restart()
    Write-Host ("[+{0}ms | Total: {1}ms] {2}" -f $delta, $total, $Message) -ForegroundColor DarkCyan
}

function New-Stopwatch {
    return [System.Diagnostics.Stopwatch]::StartNew()
}

function Stop-Log {
    param(
        [Parameter(Mandatory=$true)][System.Diagnostics.Stopwatch]$Stopwatch,
        [Parameter(Mandatory=$true)][string]$Message,
        [ConsoleColor]$Color = [ConsoleColor]::DarkCyan
    )
    $Stopwatch.Stop()
    Write-Log -Level "TIME" -Message ("{0} ({1} ms)" -f $Message, $Stopwatch.ElapsedMilliseconds) -Color $Color
}

function Ensure-STA {
    $state = [System.Threading.Thread]::CurrentThread.ApartmentState
    if ($state -ne "STA") {
        Write-Log -Level "WARN" -Message "Current thread apartment state is $state. Relaunching in STA for clipboard/SendKeys." -Color Yellow
        $args = @()
        if ($PSCommandPath) {
            $args += "-File"
            $args += "`"$PSCommandPath`""
        }
        if ($args.Count -gt 0) {
            if ($PSVersionTable.PSEdition -eq "Core") {
                Start-Process -FilePath "pwsh" -ArgumentList @("-Sta") + $args
            } else {
                Start-Process -FilePath "powershell.exe" -ArgumentList @("-STA") + $args
            }
            exit 0
        }
    }
}

Ensure-STA

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class Win32 {
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    public const uint KEYEVENTF_KEYUP = 0x0002;
    public const byte VK_CONTROL = 0x11;
    public const byte VK_C = 0x43;
    public const byte VK_V = 0x56;
}
"@

function Prompt-StartupConfig {
    param(
        [string]$CurrentHostIp,
        [int]$CurrentPort,
        [string]$CurrentLang
    )
    Write-Host ""
    Write-Host "=== Startup Configuration ===" -ForegroundColor Green
    Write-Host "Press Enter to keep the current value." -ForegroundColor DarkGray

    $hostInput = Read-Host "Host IP [$CurrentHostIp]"
    if (-not [string]::IsNullOrWhiteSpace($hostInput)) {
        $CurrentHostIp = $hostInput
    }

    $portInput = Read-Host "Port [$CurrentPort]"
    if (-not [string]::IsNullOrWhiteSpace($portInput)) {
        $parsed = 0
        if ([int]::TryParse($portInput, [ref]$parsed)) {
            $CurrentPort = $parsed
        } else {
            Write-Log -Level "WARN" -Message "Invalid port, keeping $CurrentPort" -Color Yellow
        }
    }

    $langInput = Read-Host "Target language code [$CurrentLang]"
    if (-not [string]::IsNullOrWhiteSpace($langInput)) {
        $CurrentLang = $langInput.Trim()
    }

    return @{
        HostIp = $CurrentHostIp
        Port = $CurrentPort
        Lang = $CurrentLang
    }
}

# C# code for global hotkey listener
$csharp = @"
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class HotkeyForm : Form
{
    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    
    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    
    private const int WM_HOTKEY = 0x0312;
    private const uint MOD_CONTROL = 0x0002;
    private const uint MOD_ALT = 0x0001;
    private const uint VK_T = 0x54;
    private const int HOTKEY_ID = 1;
    
    public event EventHandler HotkeyTriggered;
    
    public HotkeyForm()
    {
        this.Text = "TranslateAgent";
        this.Width = 1;
        this.Height = 1;
        this.Opacity = 0;
        this.ShowInTaskbar = false;
        this.FormBorderStyle = FormBorderStyle.None;
        this.StartPosition = FormStartPosition.Manual;
        this.Location = new System.Drawing.Point(-2000, -2000);
    }
    
    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        bool result = RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL | MOD_ALT, VK_T);
        if (!result)
        {
            throw new Exception("Failed to register hotkey Ctrl+Alt+T");
        }
        System.Console.WriteLine("[DEBUG] Hotkey registered successfully");
    }
    
    protected override void OnHandleDestroyed(EventArgs e)
    {
        UnregisterHotKey(this.Handle, HOTKEY_ID);
        base.OnHandleDestroyed(e);
    }
    
    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_HOTKEY && (int)m.WParam == HOTKEY_ID)
        {
            System.Console.WriteLine("[DEBUG] Hotkey triggered!");
            if (HotkeyTriggered != null)
            {
                HotkeyTriggered.Invoke(this, EventArgs.Empty);
            }
        }
        base.WndProc(ref m);
    }
}
"@

Add-Type -TypeDefinition $csharp -Language CSharp -ReferencedAssemblies @("System.Windows.Forms", "System.Drawing")

function Resolve-LanguageName([string]$code) {
    switch ($code.ToLower()) {
        "de" { "German" }
        "fr" { "French" }
        "en" { "English" }
        "es" { "Spanish" }
        "it" { "Italian" }
        "pt" { "Portuguese" }
        "nl" { "Dutch" }
        "pl" { "Polish" }
        "sv" { "Swedish" }
        "da" { "Danish" }
        "no" { "Norwegian" }
        "fi" { "Finnish" }
        "cs" { "Czech" }
        default { $code }
    }
}

function Get-LmStudioModel([string]$baseUrl) {
    try {
        Write-Log -Level "INFO" -Message "Fetching LM Studio models from $baseUrl" -Color DarkCyan
        $resp = Invoke-RestMethod -Method GET -Uri "$baseUrl/v1/models" -TimeoutSec 5
        if ($resp.data -and $resp.data.Count -gt 0) {
            return [string]$resp.data[0].id
        }
    } catch {
        Write-Log -Level "WARN" -Message "Could not fetch models from LM Studio: $($_.Exception.Message)" -Color Yellow
    }
    return ""
}

function Invoke-LmStudioTranslateAndCorrect {
    param(
        [Parameter(Mandatory=$true)][string]$baseUrl,
        [Parameter(Mandatory=$true)][string]$model,
        [Parameter(Mandatory=$true)][string]$targetLangCode,
        [Parameter(Mandatory=$true)][string]$inputText
    )

    $targetName = Resolve-LanguageName $targetLangCode

    $system = "You are a professional proofreader and translator. Follow instructions precisely."

    $user = "Task: 1) Correct the input text in its original language (spelling, grammar, punctuation). Do NOT change meaning. 2) Translate the corrected text into $targetName ($targetLangCode). Output ONLY the final translated text. Input: $inputText"

    $body = @{
        model = $model
        messages = @(
            @{ role="system"; content=$system },
            @{ role="user"; content=$user }
        )
        temperature = 0.2
        top_p = 0.9
        stream = $false
    } | ConvertTo-Json -Depth 6

    Write-Log -Level "INFO" -Message "Sending request to LM Studio ($targetLangCode)" -Color DarkCyan
    $resp = Invoke-RestMethod -Method POST -Uri "$baseUrl/v1/chat/completions" -ContentType "application/json" -Body $body -TimeoutSec 30
    $out = $resp.choices[0].message.content
    return ($out -as [string]).Trim()
}

function Get-ClipboardText {
    try {
        if ([System.Windows.Forms.Clipboard]::ContainsText()) {
            return [System.Windows.Forms.Clipboard]::GetText()
        }
    } catch { }
    return ""
}

function Set-ClipboardText([string]$text) {
    try {
        [System.Windows.Forms.Clipboard]::SetText($text)
        return $true
    } catch {
        return $false
    }
}

function Set-ClipboardTextRobust([string]$text) {
    $attempts = 0
    do {
        try {
            [System.Windows.Forms.Clipboard]::SetText($text)
            return $true
        } catch {
            try {
                Set-Clipboard -Value $text -ErrorAction Stop
                return $true
            } catch {
                $attempts++
                Start-Sleep -Milliseconds 80
            }
        }
    } while ($attempts -lt 6)
    return $false
}

function Send-CtrlKey([string]$key) {
    [System.Windows.Forms.SendKeys]::SendWait("^{$key}")
}

function Send-CtrlC {
    [Win32]::keybd_event([Win32]::VK_CONTROL, 0, 0, [UIntPtr]::Zero)
    [Win32]::keybd_event([Win32]::VK_C, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 20
    [Win32]::keybd_event([Win32]::VK_C, 0, [Win32]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
    [Win32]::keybd_event([Win32]::VK_CONTROL, 0, [Win32]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
}

function Send-CtrlV {
    [Win32]::keybd_event([Win32]::VK_CONTROL, 0, 0, [UIntPtr]::Zero)
    [Win32]::keybd_event([Win32]::VK_V, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 20
    [Win32]::keybd_event([Win32]::VK_V, 0, [Win32]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
    [Win32]::keybd_event([Win32]::VK_CONTROL, 0, [Win32]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
}

function Get-SelectedText([string]$savedClip) {
    $cleared = $false
    if (Set-ClipboardTextRobust "") {
        $cleared = $true
        Write-Log -Level "DEBUG" -Message "Clipboard cleared before capture" -Color DarkGray
    } else {
        Write-Log -Level "WARN" -Message "Failed to clear clipboard before capture" -Color Yellow
    }
    Send-CtrlC
    Start-Sleep -Milliseconds 80
    $attempts = 0
    do {
        $current = Get-ClipboardText
        if (-not [string]::IsNullOrEmpty($current)) {
            if ($cleared -or $current -ne $savedClip) {
            Write-Log -Level "DEBUG" -Message ("Clipboard updated, length: {0}" -f $current.Length) -Color DarkGray
            return $current
        }
        }
        $attempts++
        Start-Sleep -Milliseconds 60
    } while ($attempts -lt 12)
    Write-Log -Level "WARN" -Message "Clipboard did not change after Ctrl+C. Selection may be missing or clipboard blocked." -Color Yellow
    return ""
}

# Create invisible form for hotkey registration
$config = Prompt-StartupConfig -CurrentHostIp $HostIp -CurrentPort $Port -CurrentLang $Lang
$HostIp = $config.HostIp
$Port = $config.Port
$Lang = $config.Lang

$form = New-Object HotkeyForm

# Register hotkey event
$form.add_HotkeyTriggered({
    $timing = New-TimingSession
    # Save current clipboard
    $savedClip = Get-ClipboardText
    
    Write-Log -Level "INFO" -Message "Hotkey triggered" -Color Cyan
    $targetHwnd = [Win32]::GetForegroundWindow()
    if ($targetHwnd -ne [IntPtr]::Zero) {
        [Win32]::SetForegroundWindow($targetHwnd) | Out-Null
        Start-Sleep -Milliseconds 50
        Write-Log -Level "DEBUG" -Message ("Restored foreground window: 0x{0:X}" -f $targetHwnd.ToInt64()) -Color DarkGray
    } else {
        Write-Log -Level "WARN" -Message "Could not get foreground window handle" -Color Yellow
    }
    
    # Copy selection
    Write-Log -Level "DEBUG" -Message "Sending Ctrl+C to capture selection" -Color DarkGray
    $selected = Get-SelectedText -savedClip $savedClip
    Log-Time -Timing $timing -Message "Capture selection"
    
    if (-not [string]::IsNullOrWhiteSpace($selected)) {
        try {
            Write-Log -Level "INFO" -Message ("Processing text ({0} chars)..." -f $selected.Length) -Color Yellow
            
            $translated = Invoke-LmStudioTranslateAndCorrect -baseUrl $BaseUrl -model $Model -targetLangCode $Lang -inputText $selected
            Log-Time -Timing $timing -Message "LM Studio request"
            
            if (-not [string]::IsNullOrWhiteSpace($translated)) {
                Set-ClipboardTextRobust $translated | Out-Null
                Start-Sleep -Milliseconds 50
                Send-CtrlV
                Start-Sleep -Milliseconds 100
                Log-Time -Timing $timing -Message "Paste result"
                Write-Log -Level "INFO" -Message "Done" -Color Green
            } else {
                Set-ClipboardTextRobust $savedClip | Out-Null
                [console]::beep(400, 150)
                Write-Log -Level "WARN" -Message "Empty result from model" -Color Yellow
            }
        }
        catch {
            Set-ClipboardTextRobust $savedClip | Out-Null
            [console]::beep(300, 100)
            Write-Log -Level "ERROR" -Message $($_.Exception.Message) -Color Red
        }
    } else {
        [console]::beep(800, 100)
        Write-Log -Level "WARN" -Message "No text selected" -Color Yellow
    }
    
    # Restore clipboard
    if (-not [string]::IsNullOrWhiteSpace($savedClip)) {
        Start-Sleep -Milliseconds 50
        Set-ClipboardTextRobust $savedClip | Out-Null
        Log-Time -Timing $timing -Message "Restore clipboard"
    }
    Log-Time -Timing $timing -Message "Total hotkey workflow"
})

$BaseUrl = "http://$HostIp`:$Port"

# Auto-detect model
$Model = Get-LmStudioModel -baseUrl $BaseUrl
if ([string]::IsNullOrWhiteSpace($Model)) {
    Write-Log -Level "ERROR" -Message "Could not detect model from LM Studio at $BaseUrl" -Color Red
    Write-Log -Level "WARN" -Message "Please ensure LM Studio is running with a model loaded." -Color Yellow
    exit 1
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "TranslateAgent - Global Hotkey Active" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Hotkey:     Ctrl + Alt + T" -ForegroundColor Cyan
Write-Host "Language:   $Lang ($(Resolve-LanguageName $Lang))" -ForegroundColor Cyan
Write-Host "Model:      $Model" -ForegroundColor Cyan
Write-Host "LM Studio:  $BaseUrl" -ForegroundColor Cyan
Write-Host ""
Write-Host "Usage:" -ForegroundColor Yellow
Write-Host "  1. Select text in any application"
Write-Host "  2. Press Ctrl+Alt+T"
Write-Host "  3. The text will be corrected and translated"
Write-Host "  4. The selection will be replaced with the result"
Write-Host ""
Write-Host "Press Ctrl+C to stop the agent" -ForegroundColor Yellow
Write-Host ""

try {
    [System.Windows.Forms.Application]::Run($form)
}
finally {
    $form.Dispose()
}

