param(
    [string]$HostIp = "127.0.0.1",
    [int]$Port = 1234,
    [string]$Lang = "de"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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
        $resp = Invoke-RestMethod -Method GET -Uri "$baseUrl/v1/models" -TimeoutSec 5
        if ($resp.data -and $resp.data.Count -gt 0) {
            return [string]$resp.data[0].id
        }
    } catch {
        Write-Host "Warning: Could not fetch models from LM Studio" -ForegroundColor Yellow
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

function Send-CtrlKey([string]$key) {
    [System.Windows.Forms.SendKeys]::SendWait("^{$key}")
}

# Create invisible form for hotkey registration
$form = New-Object HotkeyForm

# Register hotkey event
$form.add_HotkeyTriggered({
    # Save current clipboard
    $savedClip = Get-ClipboardText
    
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Hotkey triggered!" -ForegroundColor Cyan
    
    # Copy selection
    Send-CtrlKey "c"
    Start-Sleep -Milliseconds 150
    $selected = Get-ClipboardText
    
    if (-not [string]::IsNullOrWhiteSpace($selected)) {
        try {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Processing text..." -ForegroundColor Yellow
            
            $translated = Invoke-LmStudioTranslateAndCorrect -baseUrl $BaseUrl -model $Model -targetLangCode $Lang -inputText $selected
            
            if (-not [string]::IsNullOrWhiteSpace($translated)) {
                Set-ClipboardText $translated | Out-Null
                Start-Sleep -Milliseconds 50
                Send-CtrlKey "v"
                Start-Sleep -Milliseconds 100
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Done!" -ForegroundColor Green
            } else {
                Set-ClipboardText $savedClip | Out-Null
                [console]::beep(400, 150)
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Empty result" -ForegroundColor Yellow
            }
        }
        catch {
            Set-ClipboardText $savedClip | Out-Null
            [console]::beep(300, 100)
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] ERROR: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        [console]::beep(800, 100)
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] No text selected" -ForegroundColor Yellow
    }
    
    # Restore clipboard
    if (-not [string]::IsNullOrWhiteSpace($savedClip)) {
        Start-Sleep -Milliseconds 50
        Set-ClipboardText $savedClip | Out-Null
    }
})

$BaseUrl = "http://$HostIp`:$Port"

# Auto-detect model
$Model = Get-LmStudioModel -baseUrl $BaseUrl
if ([string]::IsNullOrWhiteSpace($Model)) {
    Write-Host "ERROR: Could not detect model from LM Studio at $BaseUrl" -ForegroundColor Red
    Write-Host "Please ensure LM Studio is running with a model loaded." -ForegroundColor Yellow
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

