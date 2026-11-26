# ============================================================
#  DQA.ps1  -  Device Quality Assurance (Automated GUI)
#  DSU IT Support 
#  Author: Sandeep Pokharel
# ============================================================

# --- DATA STORAGE BEGINS HERE ---
$Global:SavedEmails = @()
$Global:DatabaseCSV = @'
Index,TechnicianEmail,SerialNumber,DurationMins,Charging,Screen,Touchscreen,NetworkAdapters,Keyboard,MouseTrackpad,VideoPorts,AudioOutput,Microphone,Camera,USBPorts,PalmRest,Backplate,BaseAndVents,Hinge,Notes,FinalStatus
'@
$Global:AppSettingsJson = '{}'
# --- DATA STORAGE ENDS HERE ---

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Shared keyboard focus cue for the app's WPF windows and dashboard menu.
$script:DqaFocusVisual = [Windows.Markup.XamlReader]::Parse(@'
<Style xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" TargetType="Control">
    <Setter Property="Template">
        <Setter.Value>
            <ControlTemplate TargetType="Control">
                <Border BorderBrush="#00779E" BorderThickness="2" CornerRadius="5" Margin="-1"
                        IsHitTestVisible="False" SnapsToDevicePixels="True"/>
            </ControlTemplate>
        </Setter.Value>
    </Setter>
</Style>
'@)

function Set-DqaFocusVisual([System.Windows.FrameworkElement]$Element) {
    $Element.Resources[[System.Windows.SystemParameters]::FocusVisualStyleKey] = $script:DqaFocusVisual
    # Some Windows themes set a focus style directly on controls such as buttons.
    $Element.Add_PreviewGotKeyboardFocus({
            param($sender, $e)
            $target = $e.NewFocus -as [System.Windows.FrameworkElement]
            if ($null -ne $target -and $null -ne $target.FocusVisualStyle) {
                $target.SetCurrentValue([System.Windows.FrameworkElement]::FocusVisualStyleProperty, $script:DqaFocusVisual)
            }
        })
}

# ------------------------------------------------------------
# DSU DIALOGS (replace the small stock Windows message boxes)
# Large, touch-friendly buttons in DSU colors; always on top.
# Buttons: array of @{ Text; Result; Style; IsDefault; IsCancel }
#   Style = Primary | Secondary | Danger | Success
# Returns the clicked button's Result, or $DefaultResult if closed with Esc.
# ------------------------------------------------------------
function Show-DsuDialog {
    param(
        [string]$Title,
        [string]$Message,
        [string]$Kind = "Information",   # Information | Question | Warning | Error | Success
        [array]$Buttons,
        [string]$DefaultResult = "Cancel",
        [int]$Width = 420,
        [string]$Position = "Center"      # "Bottom" keeps the middle of the screen visible (camera preview)
    )

    switch ($Kind) {
        "Warning"  { $glyph = "&#xE7BA;"; $fg = "#8A5A00"; $bg = "#FFF4DB" }
        "Error"    { $glyph = "&#xE783;"; $fg = "#B3261E"; $bg = "#FBE9E8" }
        "Question" { $glyph = "&#xE9CE;"; $fg = "#00779E"; $bg = "#E6F6FC" }
        "Success"  { $glyph = "&#xE930;"; $fg = "#1E7B3A"; $bg = "#E6F4EA" }
        default    { $glyph = "&#xE946;"; $fg = "#00779E"; $bg = "#E6F6FC" }
    }

    $dlgXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="$Width" SizeToContent="Height" WindowStartupLocation="CenterScreen" Topmost="True" ResizeMode="NoResize"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent" ShowInTaskbar="False"
        FontFamily="Segoe UI" FontSize="13" SnapsToDevicePixels="True" UseLayoutRounding="True">
    <Window.Resources>
        <Style x:Key="DlgBtn" TargetType="Button">
            <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Padding" Value="14,0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Grid>
                            <Border x:Name="ring" CornerRadius="11" BorderThickness="2" BorderBrush="Transparent" Margin="-3"/>
                            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8" Padding="{TemplateBinding Padding}">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Opacity" Value="0.88"/></Trigger>
                            <Trigger Property="IsPressed" Value="True"><Setter TargetName="bd" Property="Opacity" Value="0.75"/></Trigger>
                            <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="ring" Property="BorderBrush" Value="#00A9E0"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="TitleCloseX" TargetType="Button">
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Width="30" Height="30" CornerRadius="15" Background="Transparent">
                            <TextBlock Text="&#xE711;" FontFamily="Segoe MDL2 Assets" FontSize="11" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#1D4E6C"/></Trigger>
                            <Trigger Property="IsPressed" Value="True"><Setter TargetName="bd" Property="Background" Value="#00A9E0"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid Margin="16">
        <Border CornerRadius="12" Background="White">
            <Border.Effect><DropShadowEffect BlurRadius="20" ShadowDepth="4" Direction="270" Opacity="0.30"/></Border.Effect>
        </Border>
        <Border CornerRadius="12" Background="White" BorderBrush="#D5D8DA" BorderThickness="1">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Border x:Name="DlgTitleBar" Grid.Row="0" Background="#003552" CornerRadius="11,11,0,0" Padding="16,6,8,6" Cursor="SizeAll">
                    <DockPanel>
                        <Button x:Name="DlgCloseX" DockPanel.Dock="Right" Style="{StaticResource TitleCloseX}" ToolTip="Close"/>
                        <TextBlock x:Name="DlgTitle" Foreground="White" FontWeight="SemiBold" FontSize="14" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
                    </DockPanel>
                </Border>
                <Border Grid.Row="1" Height="3" Background="#00A9E0"/>
                <Grid Grid.Row="2" Margin="18,15,18,14">
                    <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                    <Border Width="34" Height="34" CornerRadius="17" Background="$bg" VerticalAlignment="Top" Margin="0,0,13,0">
                        <TextBlock Text="$glyph" FontFamily="Segoe MDL2 Assets" FontSize="16" Foreground="$fg" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <StackPanel Grid.Column="1" VerticalAlignment="Center">
                        <TextBlock x:Name="DlgMessage" TextWrapping="Wrap" Foreground="#1F2326" FontSize="13.5" LineHeight="19"/>
                        <TextBlock x:Name="DlgDetail" TextWrapping="Wrap" Foreground="#5A5F66" FontSize="12.5" LineHeight="18" Margin="0,4,0,0" Visibility="Collapsed"/>
                    </StackPanel>
                </Grid>
                <Border Grid.Row="3" Background="#F6F7F7" CornerRadius="0,0,11,11" BorderBrush="#E6E8E9" BorderThickness="0,1,0,0" Padding="14,9">
                    <WrapPanel x:Name="DlgButtons" HorizontalAlignment="Right"/>
                </Border>
            </Grid>
        </Border>
    </Grid>
</Window>
"@
    $dlg = [Windows.Markup.XamlReader]::Parse($dlgXaml)
    Set-DqaFocusVisual $dlg
    $dlg.Title = $Title
    $dlg.Tag = $DefaultResult
    $dlg.FindName("DlgTitle").Text = $Title
    # "Question`n`nExplanation": the first paragraph is shown bold, the rest in softer grey underneath
    $msgBlock = $dlg.FindName("DlgMessage")
    $cut = $Message.IndexOf("`n`n")
    if ($cut -gt 0) {
        $msgBlock.Text = $Message.Substring(0, $cut)
        $msgBlock.FontWeight = "SemiBold"
        $detailBlock = $dlg.FindName("DlgDetail")
        $detailBlock.Text = $Message.Substring($cut + 2).Trim()
        $detailBlock.Visibility = "Visible"
    }
    else { $msgBlock.Text = $Message }
    $dlg.FindName("DlgCloseX").Add_Click({ param($s, $e) [System.Windows.Window]::GetWindow($s).Close() })
    $dlg.FindName("DlgTitleBar").Add_MouseLeftButtonDown({
            param($s, $e)
            try { [System.Windows.Window]::GetWindow($s).DragMove() } catch { }
        })

    # Bg / text / border per button style. Main actions are filled; alternatives are white with an outline.
    $palette = @{
        Primary       = @("#00A9E0", "White", "#00A9E0")
        Success       = @("#28A745", "White", "#28A745")
        Danger        = @("#DC3545", "White", "#DC3545")
        DangerOutline = @("White", "#C82333", "#E3A1A8")
        Secondary     = @("White", "#1F2326", "#C5CBD3")
    }
    $btnPanel = $dlg.FindName("DlgButtons")
    $defaultBtn = $null
    foreach ($b in $Buttons) {
        $btn = New-Object System.Windows.Controls.Button
        $btn.Content = $b.Text
        $btn.Style = $dlg.FindResource("DlgBtn")
        $styleName = if ($b.Style -and $palette.ContainsKey($b.Style)) { $b.Style } else { "Secondary" }
        $btn.Background = $palette[$styleName][0]
        $btn.Foreground = $palette[$styleName][1]
        $btn.BorderBrush = $palette[$styleName][2]
        $btn.BorderThickness = 1
        $btn.MinWidth = 84
        $btn.Height = 34
        $btn.Margin = "8,3,0,3"
        $btn.Tag = $b.Result
        if ($b.IsDefault) { $btn.IsDefault = $true; $defaultBtn = $btn }
        if ($b.IsCancel) { $btn.IsCancel = $true }
        $btn.Add_Click({
                param($s, $e)
                $w = [System.Windows.Window]::GetWindow($s)
                $w.Tag = $s.Tag
                $w.Close()
            })
        $btnPanel.Children.Add($btn) | Out-Null
    }

    if ($Position -eq "Bottom") { $dlg.WindowStartupLocation = "Manual" }
    $dlg.Add_Loaded({
            param($s, $e)
            if ($Position -eq "Bottom") {
                $area = [System.Windows.SystemParameters]::WorkArea
                $s.Left = $area.Left + ($area.Width - $s.ActualWidth) / 2
                $s.Top = $area.Bottom - $s.ActualHeight - 12
            }
            $s.Activate() | Out-Null
            if ($defaultBtn) { $defaultBtn.Focus() | Out-Null }
        }.GetNewClosure())

    $dlg.ShowDialog() | Out-Null
    return [string]$dlg.Tag
}

# Drop-in replacement for MessageBox::Show(text, caption, buttons, icon).
# Returns "OK" / "Yes" / "No" / "Cancel" as strings, so existing
# comparisons like ($res -eq "Yes") behave exactly as before.
function Show-DsuMessage {
    param([string]$Text, [string]$Caption = "", [string]$Buttons = "OK", [string]$Icon = "Information")
    $kind = switch -Regex ($Icon) {
        "Warning|Exclamation" { "Warning"; break }
        "Error|Hand|Stop"     { "Error"; break }
        "Question"            { "Question"; break }
        default               { "Information" }
    }
    switch ($Buttons) {
        "YesNo" {
            $yesStyle = if ($kind -eq "Warning") { "Danger" } else { "Primary" }
            $btns = @(
                @{ Text = "Yes"; Result = "Yes"; Style = $yesStyle; IsDefault = $true },
                @{ Text = "No"; Result = "No"; Style = "Secondary"; IsCancel = $true }
            )
            $default = "No"
        }
        "OKCancel" {
            $btns = @(
                @{ Text = "OK"; Result = "OK"; Style = "Primary"; IsDefault = $true },
                @{ Text = "Cancel"; Result = "Cancel"; Style = "Secondary"; IsCancel = $true }
            )
            $default = "Cancel"
        }
        default {
            $btns = @(@{ Text = "OK"; Result = "OK"; Style = "Primary"; IsDefault = $true; IsCancel = $true })
            $default = "OK"
        }
    }
    return Show-DsuDialog -Title $Caption -Message $Text -Kind $kind -Buttons $btns -DefaultResult $default
}

# Two-answer test question with custom button text, e.g. "Works" / "Broken".
# Returns "Yes" (first button), "No" (second button) or "Cancel" (Esc).
function Show-DsuChoice {
    param([string]$Text, [string]$Title, [string]$YesText = "Yes", [string]$NoText = "No", [string]$Position = "Center")
    $btns = @(
        @{ Text = $YesText; Result = "Yes"; Style = "Success"; IsDefault = $true },
        @{ Text = $NoText; Result = "No"; Style = "DangerOutline" }
    )
    return Show-DsuDialog -Title $Title -Message $Text -Kind "Question" -Buttons $btns -DefaultResult "Cancel" -Position $Position
}

# Static shim so every former MessageBox call keeps its call shape:
#   [DsuMessageBox]::Show(text, caption, buttons, icon)
Add-Type -TypeDefinition @"
using System;
public static class DsuMessageBox {
    public static Func<string, string, string, string, string> Handler;
    public static string Show(string text) { return Handler(text, "", "OK", "Information"); }
    public static string Show(string text, string caption) { return Handler(text, caption, "OK", "Information"); }
    public static string Show(string text, string caption, string buttons) { return Handler(text, caption, buttons, "Information"); }
    public static string Show(string text, string caption, string buttons, string icon) { return Handler(text, caption, buttons, icon); }
}
"@ -ErrorAction SilentlyContinue
[DsuMessageBox]::Handler = [Func[string, string, string, string, string]] {
    param($t, $c, $b, $i)
    [string](Show-DsuMessage -Text $t -Caption $c -Buttons $b -Icon $i)
}

# ------------------------------------------------------------
# EXE vs PS1 DETECTION & JSON DATA LOADING
# ------------------------------------------------------------
$Global:IsRunningAsExe = $false
$Global:JsonDataPath = $null
$script:DataPath = $PSCommandPath
$script:PersistedDataText = $MyInvocation.MyCommand.ScriptBlock.Ast.Extent.Text

$exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
if ($exePath -and $exePath -match '\.exe$' -and $exePath -notmatch 'powershell|pwsh') {
    $Global:IsRunningAsExe = $true
    $exeDir = [System.IO.Path]::GetDirectoryName($exePath)
    $Global:JsonDataPath = [System.IO.Path]::Combine($exeDir, 'DQA_NC_Data.json')

    $script:DataPath = $Global:JsonDataPath
    $script:PersistedDataText = $null

    # Load saved data from JSON; if that file is damaged, fall back to the backup copy
    foreach ($candidate in @($Global:JsonDataPath, "$($Global:JsonDataPath).bak")) {
        if (-not (Test-Path -LiteralPath $candidate)) { continue }
        try {
            $text = [IO.File]::ReadAllText($candidate)
            if ($candidate -eq $Global:JsonDataPath) { $script:PersistedDataText = $text }
            $savedJson = $text | ConvertFrom-Json
            if ($null -ne $savedJson.SavedEmails) {
                $Global:SavedEmails = @($savedJson.SavedEmails)
            }
            if ($savedJson.DatabaseCSV) {
                $Global:DatabaseCSV = $savedJson.DatabaseCSV
            }
            if ($savedJson.AppSettingsJson) {
                $Global:AppSettingsJson = [string]$savedJson.AppSettingsJson
            }
            break
        } catch {
            # This file is corrupt: try the next one (backup), otherwise start fresh
        }
    }
}

# ------------------------------------------------------------
# FORCE DPI AWARENESS (Fixes blurry graphics on laptops like Lenovo)
# ------------------------------------------------------------
try {
    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public static class DPIHelper {
        [DllImport("user32.dll")]
        public static extern bool SetProcessDPIAware();
    }
"@ -ErrorAction SilentlyContinue
    [DPIHelper]::SetProcessDPIAware() | Out-Null
}
catch { }

# ------------------------------------------------------------
# GLOBAL CONFIGURATION
# ------------------------------------------------------------
$Global:AdminPinHash = "c1f330d0aff31c1c87403f1e4347bcc21aff7c179908723535f2b31723702525"

# ------------------------------------------------------------
# APP SETTINGS (gear button in the header)
# Saved as one JSON line inside the data block (PS1) or in
# DQA_NC_Data.json (EXE). Files saved before Settings existed
# have no AppSettingsJson line and simply get these defaults.
# ------------------------------------------------------------
$Global:AppSettingsDefaults = [ordered]@{
    AudioVolume      = 25        # % system volume during the audio test
    AudioSound       = "Classic chime"
    TouchTargets     = 5         # circles in the touchscreen test
    TouchSeconds     = 5         # time limit for the touchscreen test
    WifiAutoConnect  = $true     # connect to Guest Wi-Fi at startup when offline
    SaveSound        = $true     # chime when an inspection is saved
    RememberLastTech = $true     # pre-fill the last technician email at startup
    LastTechnician   = ""
    ExportMode       = "Auto"    # Auto (USB, else Desktop) | Desktop | Folder
    ExportFolder     = ""
    SettingsSavedAt  = ""        # last successful Settings write (UTC)
}

function Initialize-AppSettings {
    $Global:AppSettings = [ordered]@{}
    foreach ($k in $Global:AppSettingsDefaults.Keys) { $Global:AppSettings[$k] = $Global:AppSettingsDefaults[$k] }
    if (-not [string]::IsNullOrWhiteSpace($Global:AppSettingsJson)) {
        try {
            $saved = $Global:AppSettingsJson | ConvertFrom-Json
            foreach ($prop in $saved.PSObject.Properties) {
                if ($Global:AppSettings.Contains($prop.Name)) { $Global:AppSettings[$prop.Name] = $prop.Value }
            }
        }
        catch { }   # corrupt settings line: keep the defaults
    }
}

function Get-Setting([string]$Name) { return $Global:AppSettings[$Name] }


Initialize-AppSettings

# ------------------------------------------------------------
# AUTO-CONNECT GUEST WIFI (IF NO INTERNET)
# ------------------------------------------------------------
try {
    $isConnected = $false
    # Quick check: If there are any active network profiles, we are connected to Ethernet or some Wi-Fi
    $netProfiles = Get-NetConnectionProfile -ErrorAction SilentlyContinue
    if ($netProfiles) { 
        $isConnected = $true 
    }
    
    # Settings > "Connect to Guest Wi-Fi at startup"
    if (-not $isConnected -and (Get-Setting "WifiAutoConnect")) {
        $wlanProfiles = netsh wlan show profiles
        if (-not ($wlanProfiles -match "\bGuest\b")) {
            $xmlProfile = @"
<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
    <name>Guest</name>
    <SSIDConfig>
        <SSID>
            <name>Guest</name>
        </SSID>
    </SSIDConfig>
    <connectionType>ESS</connectionType>
    <connectionMode>auto</connectionMode>
    <MSM>
        <security>
            <authEncryption>
                <authentication>open</authentication>
                <encryption>none</encryption>
                <useOneX>false</useOneX>
            </authEncryption>
        </security>
    </MSM>
</WLANProfile>
"@
            $xmlPath = "$env:TEMP\Guest_WiFi_Profile.xml"
            $xmlProfile | Out-File -FilePath $xmlPath -Encoding ascii
            netsh wlan add profile filename="$xmlPath" | Out-Null
            Start-Sleep -Milliseconds 500
            Remove-Item -Path $xmlPath -ErrorAction SilentlyContinue
        }
        
        netsh wlan connect name="Guest" ssid="Guest" 2>&1 | Out-Null
        
        # Suppressing the Wi-Fi error as requested. The script will simply fail silently if not connected.
        
        # We removed the 4-second sleep here! The Wi-Fi will connect in the background 
        # while the GUI continues to load instantly.
    }
}
catch {
    [DsuMessageBox]::Show("Auto-connect Wi-Fi failed: $($_.Exception.Message)", "Wi-Fi Error", "OK", "Warning") | Out-Null
}

# ------------------------------------------------------------
# DURATION TRACKER & ASYNC AUDIO SETUP (WITH CORE AUDIO API)
# ------------------------------------------------------------
$Global:sessionStopwatch = New-Object System.Diagnostics.Stopwatch
$Global:isTimerPaused = $true

# Adjusted C# code to be fully compatible with PowerShell 5.1's compiler
$AudioCode = @'
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

public class AudioHelper {
    // Audio Recording Helper
    [DllImport("winmm.dll", EntryPoint="mciSendStringA", CharSet=CharSet.Ansi)]
    public static extern int mciSendString(string cmd, IntPtr ret, int len, IntPtr hwnd);
    
    // Query version of mciSendString (returns the reply text, e.g. the live input level)
    [DllImport("winmm.dll", EntryPoint="mciSendStringA", CharSet=CharSet.Ansi)]
    private static extern int mciSendStringQuery(string cmd, StringBuilder ret, int len, IntPtr hwnd);
    public static string MciQuery(string cmd) {
        StringBuilder sb = new StringBuilder(128);
        int r = mciSendStringQuery(cmd, sb, sb.Capacity, IntPtr.Zero);
        return r == 0 ? sb.ToString() : "";
    }

    // Each test sound is a stereo WAV with a left-only cue first, then a right-only cue.
    public const double ChimeSeconds = 4.0;
    public const double LeftUntil = 1.40;
    public static byte[] BuildChimeWav(string sound) {
        const int rate = 44100;
        int n = (int)(rate * ChimeSeconds);
        double[] left = new double[n];
        double[] right = new double[n];
        // Each note contains its start time, frequency, decay time, and relative gain.
        double[,] notes = {
            { 0.00, 523.25, 0.45, 1.00 }, { 0.30, 659.25, 0.45, 1.00 },
            { 0.60, 783.99, 0.45, 1.00 }, { 0.90, 1046.50, 0.35, 1.00 },
            { LeftUntil, 783.99, 0.40, 0.80 }, { 1.70, 659.25, 0.40, 0.80 },
            { 2.15, 523.25, 0.85, 0.55 }, { 2.15, 659.25, 0.85, 0.50 },
            { 2.15, 783.99, 0.85, 0.45 }, { 2.15, 1046.50, 0.85, 0.35 }
        };
        double attackSeconds = 0.008;
        double holdSeconds = 0;
        double secondPartialRatio = 2.0;
        double thirdPartialRatio = 3.0;
        double secondPartialGain = 0.25;
        double thirdPartialGain = 0.08;

        switch (sound) {
            case "Bright rise":
                notes = new double[,] {
                    { 0.00, 587.33, 0.16, 0.92 }, { 0.18, 739.99, 0.17, 0.82 },
                    { 0.37, 880.00, 0.21, 0.88 }, { 0.68, 987.77, 0.47, 1.00 },
                    { LeftUntil, 880.00, 0.17, 0.93 }, { 1.58, 739.99, 0.17, 0.86 },
                    { 1.77, 659.25, 0.20, 0.89 }, { 2.08, 587.33, 0.55, 1.00 }
                };
                holdSeconds = 0.035;
                secondPartialGain = 0.33;
                thirdPartialGain = 0.10;
                break;
            case "High cascade":
                notes = new double[,] {
                    { 0.00, 1046.50, 0.18, 0.96 }, { 0.23, 880.00, 0.15, 0.85 },
                    { 0.39, 783.99, 0.17, 0.84 }, { 0.58, 659.25, 0.49, 1.00 },
                    { LeftUntil, 880.00, 0.18, 0.96 }, { 1.63, 783.99, 0.15, 0.85 },
                    { 1.79, 659.25, 0.17, 0.84 }, { 1.98, 523.25, 0.56, 1.00 }
                };
                attackSeconds = 0.012;
                secondPartialRatio = 2.20;
                thirdPartialRatio = 3.26;
                secondPartialGain = 0.28;
                thirdPartialGain = 0.06;
                break;
            case "Major cadence":
                notes = new double[,] {
                    { 0.00, 587.33, 0.20, 0.94 }, { 0.24, 659.25, 0.22, 0.88 },
                    { 0.48, 783.99, 0.26, 0.94 }, { 0.76, 880.00, 0.36, 1.00 },
                    { LeftUntil, 783.99, 0.22, 0.98 }, { 1.66, 698.46, 0.20, 0.91 }, { 1.92, 493.88, 0.21, 0.86 },
                    { 2.18, 523.25, 0.64, 0.80 }, { 2.18, 659.25, 0.58, 0.39 }, { 2.18, 783.99, 0.54, 0.28 }
                };
                attackSeconds = 0.014;
                secondPartialGain = 0.29;
                thirdPartialGain = 0.08;
                break;
            case "Echo melody":
                notes = new double[,] {
                    { 0.00, 523.25, 0.19, 1.00 }, { 0.37, 783.99, 0.17, 0.90 },
                    { 0.53, 659.25, 0.16, 0.89 }, { 0.69, 783.99, 0.42, 0.95 },
                    { LeftUntil, 587.33, 0.19, 1.00 }, { 1.77, 880.00, 0.17, 0.90 },
                    { 1.93, 783.99, 0.16, 0.89 }, { 2.09, 698.46, 0.48, 0.95 }
                };
                attackSeconds = 0.016;
                holdSeconds = 0.060;
                secondPartialGain = 0.31;
                thirdPartialGain = 0.07;
                break;
            case "Signal fanfare":
                notes = new double[,] {
                    { 0.00, 659.25, 0.11, 0.98 }, { 0.17, 659.25, 0.11, 0.92 },
                    { 0.34, 880.00, 0.25, 0.96 }, { 0.70, 1108.73, 0.32, 1.00 },
                    { LeftUntil, 739.99, 0.11, 0.98 }, { 1.57, 739.99, 0.11, 0.92 },
                    { 1.74, 987.77, 0.25, 0.96 }, { 2.10, 880.00, 0.49, 1.00 }
                };
                attackSeconds = 0.010;
                holdSeconds = 0.045;
                secondPartialGain = 0.38;
                thirdPartialGain = 0.12;
                break;
        }

        for (int k = 0; k < notes.GetLength(0); k++) {
            double start = notes[k, 0];
            double frequency = notes[k, 1];
            double decay = notes[k, 2];
            double gain = notes[k, 3];
            bool onLeft = start < LeftUntil;
            int s0 = (int)(start * rate);
            double[] buf = onLeft ? left : right;
            // Fade out the left channel at the switch so the cues never overlap.
            int stop = onLeft ? Math.Min(n, (int)(LeftUntil * rate)) : n;
            int fadeLen = (int)(0.06 * rate);
            for (int i = s0; i < stop; i++) {
                double t = (i - s0) / (double)rate;
                double attack = Math.Min(1.0, t / attackSeconds);
                double release = Math.Max(0, t - holdSeconds);
                double env = gain * attack * Math.Exp(-release / decay);
                if (env < 0.0005 && t > 0.05) { break; }
                if (onLeft && i > stop - fadeLen) { env *= (stop - i) / (double)fadeLen; }
                double w = 2.0 * Math.PI * frequency * t;
                double v = Math.Sin(w)
                    + secondPartialGain * Math.Sin(secondPartialRatio * w) * Math.Exp(-t / 0.15)
                    + thirdPartialGain * Math.Sin(thirdPartialRatio * w) * Math.Exp(-t / 0.08);
                buf[i] += env * v;
            }
        }
        double peak = 0;
        for (int i = 0; i < n; i++) {
            if (Math.Abs(left[i]) > peak) { peak = Math.Abs(left[i]); }
            if (Math.Abs(right[i]) > peak) { peak = Math.Abs(right[i]); }
        }
        int fadeOut = (int)(0.4 * rate);
        short[] pcm = new short[n * 2];   // interleaved L, R
        for (int i = 0; i < n; i++) {
            double fade = (i > n - fadeOut) ? (n - i) / (double)fadeOut : 1.0;
            double l = (peak > 0 ? left[i] / peak : 0) * 0.8 * fade;
            double r = (peak > 0 ? right[i] / peak : 0) * 0.8 * fade;
            pcm[i * 2] = (short)Math.Round(l * 32767.0);
            pcm[i * 2 + 1] = (short)Math.Round(r * 32767.0);
        }
        return BuildWav(pcm, rate, 2);
    }

    private static byte[] BuildWav(short[] pcm, int rate, int channels) {
        int dataLen = pcm.Length * 2;
        int blockAlign = channels * 2;
        byte[] b = new byte[44 + dataLen];
        Encoding.ASCII.GetBytes("RIFF").CopyTo(b, 0);
        BitConverter.GetBytes(36 + dataLen).CopyTo(b, 4);
        Encoding.ASCII.GetBytes("WAVEfmt ").CopyTo(b, 8);
        BitConverter.GetBytes(16).CopyTo(b, 16);                  // fmt chunk size
        BitConverter.GetBytes((short)1).CopyTo(b, 20);            // PCM
        BitConverter.GetBytes((short)channels).CopyTo(b, 22);
        BitConverter.GetBytes(rate).CopyTo(b, 24);
        BitConverter.GetBytes(rate * blockAlign).CopyTo(b, 28);   // byte rate
        BitConverter.GetBytes((short)blockAlign).CopyTo(b, 32);
        BitConverter.GetBytes((short)16).CopyTo(b, 34);           // bits per sample
        Encoding.ASCII.GetBytes("data").CopyTo(b, 36);
        BitConverter.GetBytes(dataLen).CopyTo(b, 40);
        Buffer.BlockCopy(pcm, 0, b, 44, dataLen);
        return b;
    }

    // Cleans up the 16-bit PCM mic recording so the playback sounds smooth and clear:
    //  - removes DC offset (the low hum/"thump" some laptop mics add)
    //  - boosts a quiet recording (up to maxGain) based on the loud part of the speech,
    //    so a single clap spike does not stop the voice from being boosted
    //  - soft-limits peaks instead of hard clipping (no crackle)
    //  - fades in the first 20 ms (no click at the start)
    // Returns the ORIGINAL peak level (0.0 - 1.0), or -1 if the file could not be read.
    public static double NormalizeWav(string path, double maxGain) {
        try {
            byte[] b = System.IO.File.ReadAllBytes(path);
            int pos = 12, bits = 0, rate = 44100, dataPos = -1, dataLen = 0;
            while (pos + 8 <= b.Length) {
                string id = Encoding.ASCII.GetString(b, pos, 4);
                int len = BitConverter.ToInt32(b, pos + 4);
                if (id == "fmt " && pos + 8 + 16 <= b.Length) {
                    rate = BitConverter.ToInt32(b, pos + 8 + 4);
                    bits = BitConverter.ToInt16(b, pos + 8 + 14);
                }
                if (id == "data") { dataPos = pos + 8; dataLen = Math.Min(len, b.Length - dataPos); break; }
                if (len < 0) { break; }
                pos += 8 + len + (len % 2);
            }
            if (dataPos < 0 || bits != 16) { return -1; }
            int count = dataLen / 2;
            if (count < 2) { return -1; }
            double[] x = new double[count];
            double mean = 0;
            for (int i = 0; i < count; i++) { x[i] = BitConverter.ToInt16(b, dataPos + i * 2) / 32768.0; mean += x[i]; }
            mean /= count;

            // Peak and 99.5th-percentile level (histogram of |sample| after DC removal)
            int[] hist = new int[1024];
            double peak = 0;
            for (int i = 0; i < count; i++) {
                x[i] -= mean;
                double a = Math.Abs(x[i]);
                if (a > peak) { peak = a; }
                hist[Math.Min(1023, (int)(a * 1024))]++;
            }
            int target = (int)(count * 0.995), acc = 0, bin = 0;
            for (bin = 0; bin < 1024; bin++) { acc += hist[bin]; if (acc >= target) { break; } }
            double loud = Math.Max((bin + 1) / 1024.0, 1e-4);

            double gain = Math.Max(1.0, Math.Min(maxGain, 0.6 / loud));
            int fadeIn = Math.Min(count, rate / 50);
            for (int i = 0; i < count; i++) {
                double v = x[i] * gain;
                double av = Math.Abs(v);
                if (av > 0.8) { v = Math.Sign(v) * (0.8 + 0.19 * Math.Tanh((av - 0.8) / 0.19)); }
                if (i < fadeIn) { v *= i / (double)fadeIn; }
                short s = (short)Math.Round(Math.Max(-1.0, Math.Min(1.0, v)) * 32767.0);
                b[dataPos + i * 2] = (byte)(s & 0xFF);
                b[dataPos + i * 2 + 1] = (byte)((s >> 8) & 0xFF);
            }
            System.IO.File.WriteAllBytes(path, b);
            return Math.Min(1.0, peak);
        } catch { return -1; }
    }

    // Core Audio API for Direct Volume Control
    [ComImport, Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAudioEndpointVolume {
        int RegisterControlChangeNotify(IntPtr pNotify);
        int UnregisterControlChangeNotify(IntPtr pNotify);
        int GetChannelCount(out int pnChannelCount);
        int SetMasterVolumeLevel(float fLevelDB, IntPtr pguidEventContext);
        int SetMasterVolumeLevelScalar(float fLevel, IntPtr pguidEventContext);
        int GetMasterVolumeLevel(out float pfLevelDB);
        int GetMasterVolumeLevelScalar(out float pfLevel);
        int SetChannelVolumeLevel(uint nChannel, float fLevelDB, IntPtr pguidEventContext);
        int SetChannelVolumeLevelScalar(uint nChannel, float fLevel, IntPtr pguidEventContext);
        int GetChannelVolumeLevel(uint nChannel, out float pfLevelDB);
        int GetChannelVolumeLevelScalar(uint nChannel, out float pfLevel);
        int SetMute(bool bMute, IntPtr pguidEventContext);
        int GetMute(out bool pbMute);
    }

    [ComImport, Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDevice {
        int Activate(ref Guid id, int clsCtx, IntPtr activationParams, out IAudioEndpointVolume ppInterface);
    }

    [ComImport, Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDeviceEnumerator {
        int EnumAudioEndpoints(int dataFlow, int stateMask, out IntPtr ppDevices);
        int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice ppEndpoint);
    }

    [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    public class MMDeviceEnumeratorComObject { }

    public static void SetVolume(int targetLevel) {
        try {
            IMMDeviceEnumerator enumerator = (IMMDeviceEnumerator)new MMDeviceEnumeratorComObject();
            IMMDevice device = null;
            enumerator.GetDefaultAudioEndpoint(0, 1, out device);
            Guid epvid = typeof(IAudioEndpointVolume).GUID;
            IAudioEndpointVolume volume = null;
            device.Activate(ref epvid, 1, IntPtr.Zero, out volume);
            
            // Instantly un-mutes if currently muted
            volume.SetMute(false, IntPtr.Zero);
            
            // Instantly sets volume to exact percentage
            float levelScalar = targetLevel / 100f;
            volume.SetMasterVolumeLevelScalar(levelScalar, IntPtr.Zero);
        } catch { }
    }
}
'@
Add-Type -TypeDefinition $AudioCode -ErrorAction SilentlyContinue

# ------------------------------------------------------------
# KEYBOARD TESTER HOOK
# While the keyboard tester is the window in front, every key press is captured
# here and NOT passed on to Windows: the Windows key no longer opens Start, the
# Copilot key no longer opens Copilot, Print Screen no longer opens the Snipping
# Tool, Alt+Tab / Alt+F4 cannot leave or close the tester, and volume keys do not
# change the volume. The tester reads the captured keys and lights them up.
# The hook is removed when the tester closes (and Windows removes it if the app
# exits). Ctrl+Alt+Del and Win+L can never be blocked - Windows protects them.
# ------------------------------------------------------------
if (-not ('KbdHook' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

public static class KbdHook {
    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct KBDLLHOOKSTRUCT { public uint vkCode; public uint scanCode; public uint flags; public uint time; public IntPtr dwExtraInfo; }

    private struct Entry { public int Vk; public int Time; }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);
    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);
    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();
    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr GetModuleHandle(string lpModuleName);

    private const int WH_KEYBOARD_LL = 13;
    private const int VK_LWIN = 0x5B, VK_LSHIFT = 0xA0;
    public const int ComboFlag = 0x10000;   // set on a key that arrived as part of a laptop action-key burst

    private static IntPtr hook = IntPtr.Zero;
    private static IntPtr target = IntPtr.Zero;
    private static readonly LowLevelKeyboardProc proc = HookCallback;   // kept alive for the hook's lifetime
    private static readonly List<Entry> pending = new List<Entry>();
    private static readonly object gate = new object();

    public static bool IsActive { get { return hook != IntPtr.Zero; } }

    // Starts capturing; keys are only swallowed while 'window' is the foreground window
    public static bool Start(IntPtr window) {
        target = window;
        if (hook != IntPtr.Zero) { return true; }
        IntPtr mod = GetModuleHandle(Process.GetCurrentProcess().MainModule.ModuleName);
        hook = SetWindowsHookEx(WH_KEYBOARD_LL, proc, mod, 0);
        return hook != IntPtr.Zero;
    }

    public static void Stop() {
        if (hook != IntPtr.Zero) { UnhookWindowsHookEx(hook); hook = IntPtr.Zero; }
        target = IntPtr.Zero;
        lock (gate) { pending.Clear(); }
    }

    private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam) {
        if (nCode >= 0 && target != IntPtr.Zero && GetForegroundWindow() == target) {
            KBDLLHOOKSTRUCT k = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));
            int msg = wParam.ToInt32();
            if (msg == 0x100 || msg == 0x104) {   // WM_KEYDOWN / WM_SYSKEYDOWN
                lock (gate) {
                    int vk = (int)k.vkCode;
                    int t = (int)k.time;
                    bool isModifier = (vk >= 0xA0 && vk <= 0xA5) || vk == VK_LWIN || vk == 0x5C;
                    if (!isModifier) {
                        // A laptop action key (Copilot = Win+Shift+F23, display = Win+P, ...) sends Win (+Shift) and a key
                        // within a few milliseconds - far faster than a person can. Treat that burst as ONE key press:
                        // drop the Win/Shift it sent and flag the key with ComboFlag.
                        bool winBurst = pending.Exists(delegate (Entry e) { return e.Vk == VK_LWIN && unchecked(t - e.Time) < 60; });
                        if (winBurst) {
                            pending.RemoveAll(delegate (Entry e) { return (e.Vk == VK_LWIN || e.Vk == VK_LSHIFT) && unchecked(t - e.Time) < 200; });
                            vk |= ComboFlag;
                        }
                    }
                    Entry n; n.Vk = vk; n.Time = t;
                    pending.Add(n);
                }
            }
            return (IntPtr)1;   // swallow the key (down AND up) so Windows never acts on it
        }
        return CallNextHookEx(hook, nCode, wParam, lParam);
    }

    // Returns the captured keys. Win and Shift are held back ~150 ms in case they are part of an action-key burst.
    public static int[] Drain() {
        lock (gate) {
            int now = Environment.TickCount;
            List<int> ready = new List<int>();
            List<Entry> keep = new List<Entry>();
            foreach (Entry e in pending) {
                bool hold = (e.Vk == VK_LWIN || e.Vk == VK_LSHIFT) && unchecked(now - e.Time) < 150;
                if (hold) { keep.Add(e); } else { ready.Add(e.Vk); }
            }
            pending.Clear();
            pending.AddRange(keep);
            return ready.ToArray();
        }
    }
}
'@
}

# ------------------------------------------------------------
# SELF-MODIFYING SAVE FUNCTION (Saves data inside this script)
# ------------------------------------------------------------
# The lock covers the whole read/check/write operation.
$script:DqaCsvColumns = @(
    "Index", "TechnicianEmail", "SerialNumber", "DurationMins", "Charging", "Screen",
    "Touchscreen", "NetworkAdapters", "Keyboard", "MouseTrackpad", "VideoPorts", "AudioOutput",
    "Microphone", "Camera", "USBPorts", "PalmRest", "Backplate", "BaseAndVents", "Hinge", "Notes", "FinalStatus"
)

function ConvertTo-DqaDatabaseCsv($Records) {
    $index = 0
    $lines = $Records | ForEach-Object {
        $row = $_ | Select-Object -Property $script:DqaCsvColumns
        $row.Index = ++$index
        $row
    } | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1
    $header = $script:DqaCsvColumns -join ','
    if ($index -eq 0) { return $header }
    return $header + [Environment]::NewLine + ($lines -join [Environment]::NewLine)
}

function Get-DqaRecordKey($Record) {
    return "$($Record.Index)|$($Record.SerialNumber)"
}

function Test-DqaInspectionFailed($Record) {
    foreach ($column in $script:DqaCsvColumns[4..18]) {
        if ($Record.$column -in "Defective", "Needs Repair") { return $true }
    }
    return $Record.Notes -match '\bCHECK:\s*(?:Drive failing \(SMART\)|Storage drive unhealthy\b|(?:SSD|HDD|SCM|drive) \d+ GB[^:|\r\n]*:\s*Unhealthy - replace drive)'
}

function Write-FileSafely([string]$Path, [string]$Text, [System.Text.Encoding]$Encoding) {
    $tmp = "$Path.$([guid]::NewGuid().ToString('N')).saving"
    $bak = "$Path.bak"
    try {
        [IO.File]::WriteAllText($tmp, $Text, $Encoding)
        if ([IO.File]::ReadAllText($tmp, $Encoding) -cne $Text) {
            throw "The temporary save file could not be verified."
        }
        if (-not [IO.File]::Exists($Path)) {
            [IO.File]::Move($tmp, $Path)
            return
        }

        try { [IO.File]::Replace($tmp, $Path, $bak) }
        catch {
            # Some removable drives cannot use File.Replace. Keep a backup and publish by renaming.
            $errorCode = $_.Exception.GetBaseException().HResult -band 0xFFFF
            if ($errorCode -notin 1, 50, 120) { throw }
            [IO.File]::Copy($Path, $bak, $true)
            $previous = "$tmp.previous"
            [IO.File]::Move($Path, $previous)
            try { [IO.File]::Move($tmp, $Path) }
            catch {
                [IO.File]::Move($previous, $Path)
                throw
            }
            Remove-Item -LiteralPath $previous -Force -ErrorAction SilentlyContinue
        }
    }
    finally {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }
}

function Update-ScriptData {
    param(
        [string]$DatabaseCSV = $Global:DatabaseCSV,
        [System.Collections.IDictionary]$Settings = $Global:AppSettings,
        [string[]]$SavedEmails = $Global:SavedEmails,
        [scriptblock]$BeforeWrite
    )

    $dataPath = $script:DataPath
    if ([string]::IsNullOrWhiteSpace($dataPath)) { throw "Could not locate the DQA data file." }
    $lock = $null
    try {
        try {
            $lock = [IO.FileStream]::new("$dataPath.lock", [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite,
                [IO.FileShare]::None, 1, [IO.FileOptions]::DeleteOnClose)
        }
        catch {
            $errorCode = $_.Exception.GetBaseException().HResult -band 0xFFFF
            if ($errorCode -in 32, 33) { throw "Another DQA window is saving. Try again in a moment." }
            throw
        }

        $rawText = if ([IO.File]::Exists($dataPath)) { [IO.File]::ReadAllText($dataPath) } else { $null }
        if (($null -eq $rawText) -ne ($null -eq $script:PersistedDataText) -or
            ($null -ne $rawText -and -not [string]::Equals($rawText, $script:PersistedDataText, [StringComparison]::Ordinal))) {
            throw "The DQA file changed since this window opened. Keep any unsaved inspection details and reopen the latest file before trying again."
        }

        $cleanEmails = @($SavedEmails | Select-Object -Unique | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $settingsJson = $Settings | ConvertTo-Json -Compress
        if ($Global:IsRunningAsExe) {
            $dataObj = @{
                SavedEmails     = $cleanEmails
                DatabaseCSV     = $DatabaseCSV
                AppSettingsJson = $settingsJson
                LastSaved       = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            }
            $newText = $dataObj | ConvertTo-Json -Depth 5
        }
        else {
            if ($null -eq $rawText) { throw "Could not locate the script file to save data." }
            $emailItems = ($cleanEmails | ForEach-Object { "'" + ([string]$_).Replace("'", "''") + "'" }) -join ", "
            $settingsLiteral = $settingsJson.Replace("'", "''")
            $newDataBlock = "`$Global:SavedEmails = @($emailItems)`r`n`$Global:DatabaseCSV = @'`r`n$($DatabaseCSV.Trim())`r`n'@`r`n`$Global:AppSettingsJson = '$settingsLiteral'"
            $regex = '(?ms)^(# -{3} DATA STORAGE BEGINS HERE -{3})\r?$.*?^(# -{3} DATA STORAGE ENDS HERE -{3})(?=\r?$)'
            if ([regex]::Matches($rawText, $regex).Count -ne 1) { throw "Could not identify one embedded data block." }
            $replacement = "`$1`r`n" + $newDataBlock.Replace('$', '$$') + "`r`n`$2"
            $newText = $rawText -replace $regex, $replacement

            $parseErrors = $null
            [void][Management.Automation.Language.Parser]::ParseInput($newText, [ref]$null, [ref]$parseErrors)
            if ($parseErrors.Count -gt 0) { throw "The updated script did not pass its syntax check." }
        }

        if ($BeforeWrite) { & $BeforeWrite }
        Write-FileSafely -Path $dataPath -Text $newText -Encoding ([Text.UTF8Encoding]::new($true))
        $script:PersistedDataText = $newText
        $Global:DatabaseCSV = $DatabaseCSV
        $Global:AppSettings = $Settings
        $Global:SavedEmails = $cleanEmails
        $Global:AppSettingsJson = $settingsJson
    }
    finally {
        if ($lock) { $lock.Dispose() }
    }
}

function Write-DqaExport($Records, [string]$Path) {
    $tmp = "$Path.$([guid]::NewGuid().ToString('N')).saving"
    try {
        $Records | Export-Csv -LiteralPath $tmp -NoTypeInformation -Encoding UTF8 -NoClobber -ErrorAction Stop
        $expected = ($Records | ConvertTo-Csv -NoTypeInformation) -join [Environment]::NewLine
        if ([IO.File]::ReadAllText($tmp) -cne ($expected + [Environment]::NewLine)) {
            throw "The exported CSV could not be verified."
        }
        [IO.File]::Move($tmp, $Path)
    }
    finally {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }
}

# ------------------------------------------------------------
# PIN INPUT PROMPT
# ------------------------------------------------------------
function Get-PinInput {
    # DSU-styled PIN prompt. WPF sizes itself to its content, so nothing is cut off at
    # high display scaling (the old fixed-pixel WinForms box clipped its buttons).
    # Returns the entered PIN, or $null when cancelled.
    $pinXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Admin PIN" Width="350" SizeToContent="Height" WindowStartupLocation="CenterScreen" Topmost="True" ResizeMode="NoResize"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent" ShowInTaskbar="False"
        FontFamily="Segoe UI" FontSize="14" SnapsToDevicePixels="True" UseLayoutRounding="True">
    <Window.Resources>
        <Style x:Key="PinBtn" TargetType="Button">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="MinWidth" Value="84"/>
            <Setter Property="Height" Value="34"/>
            <Setter Property="Padding" Value="14,0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Opacity" Value="0.88"/></Trigger>
                            <Trigger Property="IsPressed" Value="True"><Setter TargetName="bd" Property="Opacity" Value="0.75"/></Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter TargetName="bd" Property="Opacity" Value="0.45"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="PinBoxStyle" TargetType="PasswordBox">
            <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="PasswordBox">
                        <Border x:Name="Bd" CornerRadius="10" Background="White" BorderBrush="#C5CBD3" BorderThickness="1.5">
                            <ScrollViewer x:Name="PART_ContentHost" Focusable="False" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#00A9E0"/></Trigger>
                            <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#00A9E0"/><Setter TargetName="Bd" Property="BorderThickness" Value="2"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="TitleCloseX" TargetType="Button">
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Width="30" Height="30" CornerRadius="15" Background="Transparent">
                            <TextBlock Text="&#xE711;" FontFamily="Segoe MDL2 Assets" FontSize="11" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#1D4E6C"/></Trigger>
                            <Trigger Property="IsPressed" Value="True"><Setter TargetName="bd" Property="Background" Value="#00A9E0"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid Margin="16">
        <Border CornerRadius="12" Background="White">
            <Border.Effect><DropShadowEffect BlurRadius="20" ShadowDepth="4" Direction="270" Opacity="0.30"/></Border.Effect>
        </Border>
        <Border CornerRadius="12" Background="White" BorderBrush="#D5D8DA" BorderThickness="1">
            <Grid>
                <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                <Border x:Name="PinTitleBar" Grid.Row="0" Background="#003552" CornerRadius="11,11,0,0" Padding="16,6,8,6" Cursor="SizeAll">
                    <DockPanel>
                    <Button x:Name="PinCloseX" DockPanel.Dock="Right" Style="{StaticResource TitleCloseX}" ToolTip="Cancel"/>
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <TextBlock Text="&#xE72E;" FontFamily="Segoe MDL2 Assets" FontSize="16" Foreground="#00A9E0" VerticalAlignment="Center" Margin="0,0,10,0"/>
                        <TextBlock Text="Security Verification" Foreground="White" FontWeight="SemiBold" FontSize="14" VerticalAlignment="Center"/>
                    </StackPanel>
                    </DockPanel>
                </Border>
                <Border Grid.Row="1" Height="3" Background="#00A9E0"/>
                <StackPanel Grid.Row="2" Margin="24,20,24,18">
                    <TextBlock Text="Enter the 4-digit Admin PIN to continue." Foreground="#1F2326" Margin="0,0,0,12"/>
                    <PasswordBox x:Name="PinBox" Style="{StaticResource PinBoxStyle}" MaxLength="4" PasswordChar="&#x25CF;" FontSize="22" Height="44"/>
                    <TextBlock x:Name="PinHint" Text="Digits only" FontSize="12" Foreground="#6C757D" Margin="0,6,0,0" HorizontalAlignment="Center"/>
                </StackPanel>
                <Border Grid.Row="3" Background="#F6F7F7" CornerRadius="0,0,11,11" BorderBrush="#E6E8E9" BorderThickness="0,1,0,0" Padding="14,9">
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="PinCancel" Content="Cancel" Style="{StaticResource PinBtn}" Background="White" Foreground="#1F2326" BorderBrush="#C5CBD3" BorderThickness="1" IsCancel="True"/>
                        <Button x:Name="PinVerify" Content="Verify" Style="{StaticResource PinBtn}" Background="#00A9E0" IsDefault="True" IsEnabled="False" Margin="10,0,0,0"/>
                    </StackPanel>
                </Border>
            </Grid>
        </Border>
    </Grid>
</Window>
'@
    $pinWin = [Windows.Markup.XamlReader]::Parse($pinXaml)
    Set-DqaFocusVisual $pinWin
    $pinBox = $pinWin.FindName("PinBox")
    $verifyBtn = $pinWin.FindName("PinVerify")
    $pinWin.Tag = $null

    # Digits only (typing and pasting); Verify is enabled once 4 digits are entered
    $pinBox.Add_PreviewTextInput({
            param($s, $e)
            if ($e.Text -notmatch '^[0-9]+$') { $e.Handled = $true }
        })
    [System.Windows.DataObject]::AddPastingHandler($pinBox, [System.Windows.DataObjectPastingEventHandler] {
            param($s, $e)
            $txt = [string]$e.DataObject.GetData([System.Windows.DataFormats]::Text)
            if ($txt -notmatch '^[0-9]+$') { $e.CancelCommand() }
        })
    $pinBox.Add_PasswordChanged({ $verifyBtn.IsEnabled = ($pinBox.Password.Length -eq 4) })

    $verifyBtn.Add_Click({
            $pinWin.Tag = $pinBox.Password
            $pinWin.Close()
        })
    $pinWin.FindName("PinCancel").Add_Click({ $pinWin.Close() })
    $pinWin.FindName("PinCloseX").Add_Click({ $pinWin.Close() })
    $pinWin.FindName("PinTitleBar").Add_MouseLeftButtonDown({ try { $pinWin.DragMove() } catch { } })
    $pinWin.Add_ContentRendered({ $pinWin.Activate() | Out-Null; $pinBox.Focus() | Out-Null })

    $pinWin.ShowDialog() | Out-Null
    if ($null -eq $pinWin.Tag) { return $null }
    return [string]$pinWin.Tag
}

# ------------------------------------------------------------
# PIN HASH VERIFICATION
# ------------------------------------------------------------
function Test-PinValid {
    param([string]$EnteredPin)
    if ([string]::IsNullOrWhiteSpace($EnteredPin)) { return $false }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $hashBytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($EnteredPin))
    $hashString = [BitConverter]::ToString($hashBytes).Replace("-", "").ToLower()
    return $hashString -eq $Global:AdminPinHash
}

# ------------------------------------------------------------
# TITLE-CASE HELPER
# ------------------------------------------------------------
function Format-TitleCaseEmail {
    param([string]$email)
    if (-not $email -or -not $email.Contains("@")) { return $email }
    $parts = $email -split "@", 2
    $localPart = $parts[0]
    $domain = $parts[1]
    $words = $localPart -split "\."
    $titled = $words | ForEach-Object {
        if ($_.Length -gt 0) { $_.Substring(0, 1).ToUpper() + $_.Substring(1).ToLower() }
        else { $_ }
    }
    return ($titled -join ".") + "@" + $domain
}

# ------------------------------------------------------------
# NATIVE KEYBOARD TESTER
# ------------------------------------------------------------
function Show-KeyboardTester {
    $kbdXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Native Keyboard Tester" Height="650" Width="1100" WindowStartupLocation="CenterScreen" Background="#EAEAED" Topmost="True" Focusable="True" SnapsToDevicePixels="True" UseLayoutRounding="True">
    <Window.Resources>
        <!-- Standard Key Style -->
        <Style TargetType="Border">
            <Setter Property="Background" Value="White"/>
            <Setter Property="CornerRadius" Value="3"/>
            <Setter Property="Margin" Value="3"/>
            <Setter Property="BorderBrush" Value="#D0D0D0"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        
        <!-- Standard Text Style -->
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="#A0A0A0"/>
            <Setter Property="HorizontalAlignment" Value="Center"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="TextAlignment" Value="Center"/>
        </Style>

        <!-- Modifier Text Style -->
        <Style x:Key="ModText" TargetType="TextBlock" BasedOn="{StaticResource {x:Type TextBlock}}">
            <Setter Property="FontSize" Value="12"/>
        </Style>

        <Style x:Key="DarkKey" TargetType="Border" BasedOn="{StaticResource {x:Type Border}}">
            <Setter Property="Background" Value="#4d4f53"/>
            <Setter Property="BorderBrush" Value="#003552"/>
        </Style>
        <Style x:Key="DarkText" TargetType="TextBlock" BasedOn="{StaticResource {x:Type TextBlock}}">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="14"/>
        </Style>
        
        <!-- No Border for gaps -->
        <Style x:Key="Gap" TargetType="Border">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Effect" Value="{x:Null}"/>
        </Style>
        
        <Style TargetType="Button">
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="10" Padding="18,0">
                            <Border.Effect><DropShadowEffect BlurRadius="10" ShadowDepth="2" Direction="270" Opacity="0.22"/></Border.Effect>
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.9"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.75"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <!-- Keyboard layout -->
            <RowDefinition Height="*"/>    <!-- Spacer -->
            <RowDefinition Height="Auto"/> <!-- Buttons -->
        </Grid.RowDefinitions>

        <Grid Grid.Row="0" HorizontalAlignment="Center" VerticalAlignment="Top">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/> <!-- Main Block -->
                <ColumnDefinition Width="30"/>   <!-- Gap -->
                <ColumnDefinition Width="Auto"/> <!-- Nav Block -->
            </Grid.ColumnDefinitions>
            
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/> <!-- Row 0 -->
                <RowDefinition Height="Auto"/> <!-- Row 1 -->
                <RowDefinition Height="Auto"/> <!-- Row 2 -->
                <RowDefinition Height="Auto"/> <!-- Row 3 -->
                <RowDefinition Height="Auto"/> <!-- Row 4 -->
                <RowDefinition Height="Auto"/> <!-- Row 5 -->
            </Grid.RowDefinitions>

            <!-- ===================== MAIN BLOCK ===================== -->
            
            <!-- ROW 0: esc, F-keys, eject -->
            <StackPanel Grid.Row="0" Grid.Column="0" Orientation="Horizontal" Margin="0,0,0,10">
                <Border x:Name="Key_Escape" Width="50" Height="35"><TextBlock Text="esc" Style="{StaticResource ModText}"/></Border>
                <Border Width="10" Style="{StaticResource Gap}"/>
                <Border x:Name="Key_F1" Width="50" Height="35"><TextBlock Text="F1" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_F2" Width="50" Height="35"><TextBlock Text="F2" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_F3" Width="50" Height="35"><TextBlock Text="F3" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_F4" Width="50" Height="35"><TextBlock Text="F4" Style="{StaticResource ModText}"/></Border>
                <Border Width="10" Style="{StaticResource Gap}"/>
                <Border x:Name="Key_F5" Width="50" Height="35"><TextBlock Text="F5" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_F6" Width="50" Height="35"><TextBlock Text="F6" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_F7" Width="50" Height="35"><TextBlock Text="F7" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_F8" Width="50" Height="35"><TextBlock Text="F8" Style="{StaticResource ModText}"/></Border>
                <Border Width="10" Style="{StaticResource Gap}"/>
                <Border x:Name="Key_F9" Width="50" Height="35"><TextBlock Text="F9" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_F10" Width="50" Height="35"><TextBlock Text="F10" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_F11" Width="50" Height="35"><TextBlock Text="F11" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_F12" Width="50" Height="35"><TextBlock Text="F12" Style="{StaticResource ModText}"/></Border>
                <Border Width="10" Style="{StaticResource Gap}"/>
                <Border x:Name="Key_Sleep" Width="50" Height="35" Style="{StaticResource DarkKey}"><TextBlock Text="&#x23FE;" Style="{StaticResource DarkText}"/></Border>
            </StackPanel>

            <!-- ROW 1: Numbers -->
            <StackPanel Grid.Row="1" Grid.Column="0" Orientation="Horizontal">
                <Border x:Name="Key_Oem3" Width="50" Height="50"><TextBlock Text="~&#x0a;`" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_D1" Width="50" Height="50"><TextBlock Text="!&#x0a;1" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_D2" Width="50" Height="50"><TextBlock Text="@&#x0a;2" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_D3" Width="50" Height="50"><TextBlock Text="#&#x0a;3" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_D4" Width="50" Height="50"><TextBlock Text="$&#x0a;4" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_D5" Width="50" Height="50"><TextBlock Text="%&#x0a;5" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_D6" Width="50" Height="50"><TextBlock Text="^&#x0a;6" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_D7" Width="50" Height="50"><TextBlock Text="&amp;&#x0a;7" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_D8" Width="50" Height="50"><TextBlock Text="*&#x0a;8" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_D9" Width="50" Height="50"><TextBlock Text="(&#x0a;9" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_D0" Width="50" Height="50"><TextBlock Text=")&#x0a;0" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_OemMinus" Width="50" Height="50"><TextBlock Text="_&#x0a;-" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_OemPlus" Width="50" Height="50"><TextBlock Text="+&#x0a;=" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_Back" Width="78" Height="50"><TextBlock Text="backspace" Style="{StaticResource ModText}"/></Border>
            </StackPanel>

            <!-- ROW 2: QWERTY -->
            <StackPanel Grid.Row="2" Grid.Column="0" Orientation="Horizontal">
                <Border x:Name="Key_Tab" Width="78" Height="50"><TextBlock Text="tab" Style="{StaticResource ModText}" HorizontalAlignment="Left" Margin="10,0,0,0"/></Border>
                <Border x:Name="Key_Q" Width="50" Height="50"><TextBlock Text="Q"/></Border>
                <Border x:Name="Key_W" Width="50" Height="50"><TextBlock Text="W"/></Border>
                <Border x:Name="Key_E" Width="50" Height="50"><TextBlock Text="E"/></Border>
                <Border x:Name="Key_R" Width="50" Height="50"><TextBlock Text="R"/></Border>
                <Border x:Name="Key_T" Width="50" Height="50"><TextBlock Text="T"/></Border>
                <Border x:Name="Key_Y" Width="50" Height="50"><TextBlock Text="Y"/></Border>
                <Border x:Name="Key_U" Width="50" Height="50"><TextBlock Text="U"/></Border>
                <Border x:Name="Key_I" Width="50" Height="50"><TextBlock Text="I"/></Border>
                <Border x:Name="Key_O" Width="50" Height="50"><TextBlock Text="O"/></Border>
                <Border x:Name="Key_P" Width="50" Height="50"><TextBlock Text="P"/></Border>
                <Border x:Name="Key_OemOpenBrackets" Width="50" Height="50"><TextBlock Text="{}{&#x0a;[" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_OemCloseBrackets" Width="50" Height="50"><TextBlock Text="}&#x0a;]" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_Oem5" Width="50" Height="50"><TextBlock Text="|&#x0a;\" Style="{StaticResource ModText}"/></Border>
            </StackPanel>

            <!-- ROW 3: ASDF -->
            <StackPanel Grid.Row="3" Grid.Column="0" Orientation="Horizontal">
                <Border x:Name="Key_Capital" Width="92" Height="50"><TextBlock Text="caps lock" Style="{StaticResource ModText}" HorizontalAlignment="Left" Margin="10,0,0,0"/></Border>
                <Border x:Name="Key_A" Width="50" Height="50"><TextBlock Text="A"/></Border>
                <Border x:Name="Key_S" Width="50" Height="50"><TextBlock Text="S"/></Border>
                <Border x:Name="Key_D" Width="50" Height="50"><TextBlock Text="D"/></Border>
                <Border x:Name="Key_F" Width="50" Height="50"><TextBlock Text="F"/></Border>
                <Border x:Name="Key_G" Width="50" Height="50"><TextBlock Text="G"/></Border>
                <Border x:Name="Key_H" Width="50" Height="50"><TextBlock Text="H"/></Border>
                <Border x:Name="Key_J" Width="50" Height="50"><TextBlock Text="J"/></Border>
                <Border x:Name="Key_K" Width="50" Height="50"><TextBlock Text="K"/></Border>
                <Border x:Name="Key_L" Width="50" Height="50"><TextBlock Text="L"/></Border>
                <Border x:Name="Key_Oem1" Width="50" Height="50"><TextBlock Text=":&#x0a;;" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_OemQuotes" Width="50" Height="50"><TextBlock Text="&quot;&#x0a;'" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_Return" Width="86" Height="50"><TextBlock Text="enter" Style="{StaticResource ModText}" HorizontalAlignment="Right" Margin="0,0,10,0"/></Border>
            </StackPanel>

            <!-- ROW 4: ZXCV -->
            <StackPanel Grid.Row="4" Grid.Column="0" Orientation="Horizontal">
                <Border x:Name="Key_LeftShift" Width="120" Height="50"><TextBlock Text="shift" Style="{StaticResource ModText}" HorizontalAlignment="Left" Margin="10,0,0,0"/></Border>
                <Border x:Name="Key_Z" Width="50" Height="50"><TextBlock Text="Z"/></Border>
                <Border x:Name="Key_X" Width="50" Height="50"><TextBlock Text="X"/></Border>
                <Border x:Name="Key_C" Width="50" Height="50"><TextBlock Text="C"/></Border>
                <Border x:Name="Key_V" Width="50" Height="50"><TextBlock Text="V"/></Border>
                <Border x:Name="Key_B" Width="50" Height="50"><TextBlock Text="B"/></Border>
                <Border x:Name="Key_N" Width="50" Height="50"><TextBlock Text="N"/></Border>
                <Border x:Name="Key_M" Width="50" Height="50"><TextBlock Text="M"/></Border>
                <Border x:Name="Key_OemComma" Width="50" Height="50"><TextBlock Text="&lt;&#x0a;," Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_OemPeriod" Width="50" Height="50"><TextBlock Text="&gt;&#x0a;." Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_OemQuestion" Width="50" Height="50"><TextBlock Text="?&#x0a;/" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_RightShift" Width="114" Height="50"><TextBlock Text="shift" Style="{StaticResource ModText}" HorizontalAlignment="Right" Margin="0,0,10,0"/></Border>
            </StackPanel>

            <!-- ROW 5: laptop bottom row (fn, ctrl, windows, alt, space, alt, copilot, ctrl) & arrows -->
            <StackPanel Grid.Row="5" Grid.Column="0" Orientation="Horizontal">
                <Border x:Name="Key_Fn" Width="50" Height="50" Background="#F4F4F4" BorderBrush="#C8C8C8" Opacity="0.75" ToolTip="Windows cannot detect the fn key itself"><TextBlock Text="fn" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_LeftCtrl" Width="55" Height="50"><TextBlock Text="ctrl" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_LWin" Width="55" Height="50"><TextBlock Text="windows" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_LeftAlt" Width="50" Height="50"><TextBlock Text="alt" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_Space" Width="237" Height="50"><TextBlock Text=""/></Border>
                <Border x:Name="Key_RightAlt" Width="50" Height="50"><TextBlock Text="alt" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_Copilot" Width="50" Height="50"><TextBlock Text="copilot" Style="{StaticResource ModText}" FontSize="11"/></Border>
                <Border x:Name="Key_RightCtrl" Width="55" Height="50"><TextBlock Text="ctrl" Style="{StaticResource ModText}"/></Border>
                <Border Width="11" Style="{StaticResource Gap}"/>
                <!-- Arrows Block -->
                <Border x:Name="Key_Left" Width="45" Height="25" VerticalAlignment="Bottom" Margin="3,3,1,3"><TextBlock Text="&#x25C0;" Style="{StaticResource ModText}"/></Border>
                <StackPanel VerticalAlignment="Bottom" Margin="1,3,1,3">
                    <Border x:Name="Key_Up" Width="45" Height="23" Margin="0,0,0,2"><TextBlock Text="&#x25B2;" Style="{StaticResource ModText}"/></Border>
                    <Border x:Name="Key_Down" Width="45" Height="23" Margin="0,2,0,0"><TextBlock Text="&#x25BC;" Style="{StaticResource ModText}"/></Border>
                </StackPanel>
                <Border x:Name="Key_Right" Width="45" Height="25" VerticalAlignment="Bottom" Margin="1,3,3,3"><TextBlock Text="&#x25B6;" Style="{StaticResource ModText}"/></Border>
            </StackPanel>

            <!-- ===================== NAV BLOCK ===================== -->
            
            <StackPanel Grid.Row="0" Grid.Column="2" Orientation="Horizontal" Margin="0,0,0,10">
                <Border x:Name="Key_Snapshot" Width="50" Height="35"><TextBlock Text="print&#x0a;screen" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_Scroll" Width="50" Height="35"><TextBlock Text="scroll&#x0a;lock" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_Pause" Width="50" Height="35"><TextBlock Text="pause" Style="{StaticResource ModText}"/></Border>
            </StackPanel>

            <StackPanel Grid.Row="1" Grid.Column="2" Orientation="Horizontal">
                <Border x:Name="Key_Insert" Width="50" Height="50"><TextBlock Text="insert" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_Home" Width="50" Height="50"><TextBlock Text="home" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_PageUp" Width="50" Height="50"><TextBlock Text="page&#x0a;up" Style="{StaticResource ModText}"/></Border>
            </StackPanel>

            <StackPanel Grid.Row="2" Grid.Column="2" Orientation="Horizontal">
                <Border x:Name="Key_Delete" Width="50" Height="50"><TextBlock Text="delete" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_End" Width="50" Height="50"><TextBlock Text="end" Style="{StaticResource ModText}"/></Border>
                <Border x:Name="Key_PageDown" Width="50" Height="50"><TextBlock Text="page&#x0a;down" Style="{StaticResource ModText}"/></Border>
            </StackPanel>

            <StackPanel Grid.Row="3" Grid.Column="2" Orientation="Horizontal">
                <Border Width="50" Style="{StaticResource Gap}" Margin="3"/>
                <Border x:Name="Key_NumLock" Width="50" Height="50"><TextBlock Text="num&#x0a;lock" Style="{StaticResource ModText}"/></Border>
                <Border Width="50" Style="{StaticResource Gap}" Margin="3"/>
            </StackPanel>
        </Grid>

        <!-- Fn Lock / F-row mode (judged from the last F-row key pressed) -->
        <Border x:Name="FnModeBox" Grid.Row="1" HorizontalAlignment="Center" VerticalAlignment="Bottom" Background="#F0F1F3" BorderThickness="0" CornerRadius="16" Padding="18,7" Margin="0,16,0,0">
            <TextBlock x:Name="FnModeText" Text="Fn Lock: press any F-key to check" Foreground="#4D4F53" FontSize="14" FontWeight="SemiBold"/>
        </Border>

        <!-- Instructions & Close Buttons -->
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,30,0,10">
            <Button x:Name="BtnFail" Width="200" Height="48" Background="#DC3545" Foreground="White" Margin="0,0,16,0"><StackPanel Orientation="Horizontal"><TextBlock Text="&#xE711;" FontFamily="Segoe MDL2 Assets" FontSize="15" FontWeight="Normal" Foreground="White" VerticalAlignment="Center" Margin="0,0,10,0"/><TextBlock Text="Fail (Broken Keys)" FontSize="15" FontWeight="SemiBold" Foreground="White" VerticalAlignment="Center"/></StackPanel></Button>
            <Button x:Name="BtnPass" Width="200" Height="48" Background="#28A745" Foreground="White" Margin="0,0,16,0"><StackPanel Orientation="Horizontal"><TextBlock Text="&#xE73E;" FontFamily="Segoe MDL2 Assets" FontSize="15" FontWeight="Normal" Foreground="White" VerticalAlignment="Center" Margin="0,0,10,0"/><TextBlock Text="Pass (All Work)" FontSize="15" FontWeight="SemiBold" Foreground="White" VerticalAlignment="Center"/></StackPanel></Button>
            <Button x:Name="BtnClose" Width="200" Height="48" Background="#4D4F53" Foreground="White"><StackPanel Orientation="Horizontal"><TextBlock Text="&#xE72B;" FontFamily="Segoe MDL2 Assets" FontSize="15" FontWeight="Normal" Foreground="White" VerticalAlignment="Center" Margin="0,0,10,0"/><TextBlock Text="Close (Cancel)" FontSize="15" FontWeight="SemiBold" Foreground="White" VerticalAlignment="Center"/></StackPanel></Button>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = [System.Xml.XmlNodeReader]::new([xml]$kbdXaml)
    $kbdWindow = [Windows.Markup.XamlReader]::Load($reader)
    Set-DqaFocusVisual $kbdWindow

    $global:kbdResult = $null

    $kbdWindow.FindName("BtnPass").Add_Click({
            $global:kbdResult = "Operational"
            $kbdWindow.Close()
        })

    $kbdWindow.FindName("BtnFail").Add_Click({
            $global:kbdResult = "Defective"
            $kbdWindow.Close()
        })
    
    $kbdWindow.FindName("BtnClose").Add_Click({
            $global:kbdResult = $null
            $kbdWindow.Close()
        })

    $bgBrush = (New-Object System.Windows.Media.BrushConverter).ConvertFrom("#39FF14")
    $fnMode = $kbdWindow.FindName("FnModeText")
    $fnModeBox = $kbdWindow.FindName("FnModeBox")
    # What to press to switch the F-row to F1-F12, for this laptop's brand
    $maker = ""
    try { $maker = [string](Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).Manufacturer } catch { }
    $fnSwitchTip = switch -Regex ($maker) {
        "Dell|Lenovo" { "press Fn+Esc"; break }
        "Microsoft" { "press the Fn key"; break }
        default { "hold Fn or turn on Fn Lock" }
    }

    # Bottom line: which mode the F-row is in (judged from the last F-row key that arrived) and what to do
    $setFnMode = {
        param([bool]$FunctionKeys)
        if ($FunctionKeys) {
            $fnMode.Text = "Fn Lock: ON  -  F1-F12 mode"
            $fnModeBox.Background = "#E6F4EA"; $fnMode.Foreground = "#1E7B3A"
        }
        else {
            $fnMode.Text = "Fn Lock: OFF  -  Media mode  -  $fnSwitchTip to test F1-F12"
            $fnModeBox.Background = "#FFF4DB"; $fnMode.Foreground = "#8A5A00"
        }
    }

    # Media / action keys the F-row sends when Fn Lock is off
    $mediaKeys = @("VolumeMute", "VolumeDown", "VolumeUp", "MediaPlayPause", "MediaNextTrack", "MediaPreviousTrack", "MediaStop",
        "BrowserBack", "BrowserForward", "BrowserRefresh", "BrowserStop", "BrowserSearch", "BrowserFavorites", "BrowserHome",
        "LaunchMail", "SelectMedia", "LaunchApplication1", "LaunchApplication2", "Sleep")

    # An F-row key in media mode (Mute, Volume, a display-switch shortcut, ...): it lights nothing,
    # the bottom line just switches to Media mode so the technician knows to switch to F1-F12.
    $registerAction = {
        param([string]$ActionId)
        & $setFnMode $false
    }

    # Lights up one key on the drawing
    $markKey = {
        param([string]$keyStr)
        if ($keyStr -match '^F([1-9]|1[0-2])$') { & $setFnMode $true }
        if ($mediaKeys -contains $keyStr) { & $registerAction $keyStr; return }
        $border = $kbdWindow.FindName("Key_$keyStr")
        if (-not $border) {
            # Duplicate WPF Key enum names (like Oem6 vs OemCloseBrackets) and special keys
            $alias = switch ($keyStr) {
                "Oem6" { "OemCloseBrackets" } "Oem4" { "OemOpenBrackets" } "OemPipe" { "Oem5" }
                "OemSemicolon" { "Oem1" } "Oem7" { "OemQuotes" } "Oem2" { "OemQuestion" } "OemTilde" { "Oem3" }
                "Prior" { "PageUp" } "Next" { "PageDown" } "PrintScreen" { "Snapshot" } "Enter" { "Return" }
                "CapsLock" { "Capital" } "F23" { "Copilot" } "RWin" { "LWin" }
                default { $null }
            }
            if ($alias) { $border = $kbdWindow.FindName("Key_$alias") }
        }
        if ($border) { $border.Background = $bgBrush }
    }

    # Fallback path: keys that reach the window normally (only when the hook is not active)
    $kbdWindow.Add_PreviewKeyDown({
            param($src, $e)
            $keyStr = $e.Key.ToString()
            if ($keyStr -eq "System") { $keyStr = $e.SystemKey.ToString() }
            & $markKey $keyStr
            # Prevent Enter, Space, Escape and System keys from doing their default behavior (like closing the window or clicking buttons)
            $e.Handled = $true
        })

    $kbdWindow.Add_PreviewKeyUp({
            param($src, $e)
            $e.Handled = $true
        })

    # Main path: the hook captures (and blocks) keys while the tester is in front; read them every 30 ms
    $hookTimer = New-Object System.Windows.Threading.DispatcherTimer
    $hookTimer.Interval = [TimeSpan]::FromMilliseconds(30)
    $hookTimer.Add_Tick({
            foreach ($code in [KbdHook]::Drain()) {
                $isCombo = ($code -band [KbdHook]::ComboFlag) -ne 0
                $vk = $code -band 0xFFFF
                $keyStr = [System.Windows.Input.KeyInterop]::KeyFromVirtualKey($vk).ToString()
                if ($keyStr -eq "F23") { & $markKey "F23" }                   # Copilot key
                elseif ($isCombo) { & $registerAction "combo-$vk" }           # action key such as display switch (Win+P)
                else { & $markKey $keyStr }
            }
        })

    $kbdWindow.Add_ContentRendered({
            $kbdWindow.Activate() | Out-Null
            $kbdWindow.Focus() | Out-Null
            try {
                $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper($kbdWindow)).Handle
                if ([KbdHook]::Start($hwnd)) { $hookTimer.Start() }
            }
            catch { }
        })

    $kbdWindow.Add_Closed({
            $hookTimer.Stop()
            try { [KbdHook]::Stop() } catch { }
        })
    $kbdWindow.ShowDialog() | Out-Null
    return $global:kbdResult
}

# ------------------------------------------------------------
# CUSTOM TOUCHSCREEN PASS DIALOG
# ------------------------------------------------------------
function Show-TouchscreenPassDialog {
    param([string]$Message)
    # Large DSU-styled buttons so it is easy to tap with a finger. Returns "OK" (Pass) or "Retry" (Redo).
    $btns = @(
        @{ Text = "Pass"; Result = "OK"; Style = "Success"; IsDefault = $true },
        @{ Text = "Redo"; Result = "Retry"; Style = "Secondary" }
    )
    return Show-DsuDialog -Title "Touchscreen Verified" -Message $Message -Kind "Success" -Buttons $btns -DefaultResult "Retry"
}

# ------------------------------------------------------------
# NATIVE TOUCHSCREEN TESTER (tap every circle before time runs out)
# Number of circles and time limit come from Settings (default 5 / 5 s).
# Pass = ALL circles tapped. Result dialogs are shown only after the
# tester window has closed: a dialog opened inside a touch event can get
# stuck behind the topmost full-screen window and ignore finger taps.
# ------------------------------------------------------------
function Show-TouchscreenTester {
    $target = [int](Get-Setting "TouchTargets"); if ($target -lt 1) { $target = 5 }
    $seconds = [double](Get-Setting "TouchSeconds"); if ($seconds -lt 1) { $seconds = 5 }

    $touchXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Touchscreen Tester" Height="600" Width="900" WindowStartupLocation="CenterScreen" Background="#1a1a2e" Topmost="True" WindowState="Maximized" SnapsToDevicePixels="True" UseLayoutRounding="True">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="10" Padding="18,0">
                            <Border.Effect><DropShadowEffect BlurRadius="10" ShadowDepth="2" Direction="270" Opacity="0.22"/></Border.Effect>
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.9"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.75"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" HorizontalAlignment="Center" Margin="0,25,0,0">
            <TextBlock x:Name="TouchInstruction" Text="&#x1F446; TAP ALL THE CIRCLES!" FontSize="36" FontWeight="Bold" Foreground="White" HorizontalAlignment="Center"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,10,0,0">
                <TextBlock x:Name="TouchTimer" Text="5.0s" FontSize="28" FontWeight="Bold" Foreground="#FFC107" Margin="0,0,30,0"/>
                <TextBlock x:Name="TouchCounter" Text="Tapped: 0 / 5" FontSize="22" Foreground="White"/>
            </StackPanel>
        </StackPanel>
        <Canvas x:Name="TouchCanvas" Grid.Row="1" Background="Transparent" Margin="20"/>
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,0,0,25">
            <Button x:Name="BtnTouchFail" Width="220" Height="48" Background="#DC3545" Foreground="White" Margin="0,0,16,0"><StackPanel Orientation="Horizontal"><TextBlock Text="&#xE711;" FontFamily="Segoe MDL2 Assets" FontSize="15" FontWeight="Normal" Foreground="White" VerticalAlignment="Center" Margin="0,0,10,0"/><TextBlock Text="Fail (Touch Broken)" FontSize="15" FontWeight="SemiBold" Foreground="White" VerticalAlignment="Center"/></StackPanel></Button>
            <Button x:Name="BtnTouchClose" Width="220" Height="48" Background="#4D4F53" Foreground="White"><StackPanel Orientation="Horizontal"><TextBlock Text="&#xE72B;" FontFamily="Segoe MDL2 Assets" FontSize="15" FontWeight="Normal" Foreground="White" VerticalAlignment="Center" Margin="0,0,10,0"/><TextBlock Text="Close (Cancel)" FontSize="15" FontWeight="SemiBold" Foreground="White" VerticalAlignment="Center"/></StackPanel></Button>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = [System.Xml.XmlNodeReader]::new([xml]$touchXaml)
    $touchWindow = [Windows.Markup.XamlReader]::Load($reader)
    Set-DqaFocusVisual $touchWindow

    $global:touchCount = 0
    $global:touchTarget = $target
    $global:touchSeconds = $seconds
    $global:touchDone = $false
    $global:touchOutcome = $null     # AllTapped | Timeout | Defective | $null (cancelled)
    $global:touchStopwatch = $null

    $touchCanvas = $touchWindow.FindName("TouchCanvas")
    $touchCounterText = $touchWindow.FindName("TouchCounter")
    $touchTimerText = $touchWindow.FindName("TouchTimer")
    $touchInstructionText = $touchWindow.FindName("TouchInstruction")
    $touchTimerText.Text = "{0:F1}s" -f $seconds
    $touchCounterText.Text = "Tapped: 0 / $target"

    $circleColors = @("#FF6B6B", "#4ECDC4", "#45B7D1", "#96CEB4", "#FFEAA7", "#DDA0DD", "#F7DC6F", "#74B9FF")
    $brushConverter = New-Object System.Windows.Media.BrushConverter

    # Closes the tester shortly after it finishes, so the final message is visible for a moment
    $closeTimer = New-Object System.Windows.Threading.DispatcherTimer
    $closeTimer.Interval = [TimeSpan]::FromMilliseconds(700)
    $closeTimer.Add_Tick({ $closeTimer.Stop(); $touchWindow.Close() })

    # Countdown
    $touchTickTimer = New-Object System.Windows.Threading.DispatcherTimer
    $touchTickTimer.Interval = [TimeSpan]::FromMilliseconds(100)
    $touchTickTimer.Add_Tick({
            if ($global:touchDone) { $touchTickTimer.Stop(); return }
            $remaining = $global:touchSeconds - $global:touchStopwatch.Elapsed.TotalSeconds
            if ($remaining -le 0) {
                $global:touchDone = $true
                $global:touchOutcome = "Timeout"
                $touchTickTimer.Stop()
                $touchTimerText.Text = "0.0s"
                $touchInstructionText.Text = "TIME'S UP!"
                $touchInstructionText.Foreground = $brushConverter.ConvertFrom("#E74C3C")
                $touchCanvas.Children.Clear()
                $closeTimer.Start()
            }
            else {
                $touchTimerText.Text = "{0:F1}s" -f $remaining
                if ($remaining -le 2) { $touchTimerText.Foreground = $brushConverter.ConvertFrom("#E74C3C") }
            }
        })

    # Only physical touch input can register a target.
    $registerHit = {
        param($circle)
        if ($global:touchDone) { return }
        if ($circle.Visibility -ne "Visible") { return }
        $circle.Visibility = "Collapsed"
        foreach ($child in $touchCanvas.Children) {
            if ($child -is [System.Windows.Controls.TextBlock] -and $child.Tag -eq ($circle.Tag + "_lbl")) {
                $child.Visibility = "Collapsed"
            }
        }
        $global:touchCount++
        $touchCounterText.Text = "Tapped: $($global:touchCount) / $($global:touchTarget)"
        if ($global:touchCount -ge $global:touchTarget) {
            $global:touchDone = $true
            $global:touchOutcome = "AllTapped"
            $touchTickTimer.Stop()
            $touchInstructionText.Text = "$([char]0x2714) ALL CIRCLES TAPPED!"
            $touchInstructionText.Foreground = $brushConverter.ConvertFrom("#2ECC71")
            $closeTimer.Start()
        }
    }

    $touchWindow.FindName("BtnTouchFail").Add_Click({
            $global:touchDone = $true
            $global:touchOutcome = "Defective"
            $touchWindow.Close()
        })

    $touchWindow.FindName("BtnTouchClose").Add_Click({
            $global:touchDone = $true
            $global:touchOutcome = $null
            $touchWindow.Close()
        })

    $touchWindow.Add_Closed({
            $touchTickTimer.Stop()
            $closeTimer.Stop()
        })

    $touchWindow.Add_ContentRendered({
            $global:touchStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $global:touchCount = 0
            $touchTickTimer.Start()

            $circleSize = 80
            $rnd = New-Object System.Random
            $canvasW = [Math]::Max(400, $touchCanvas.ActualWidth)
            $canvasH = [Math]::Max(300, $touchCanvas.ActualHeight)
            $maxX = [int]($canvasW - $circleSize - 10)
            $maxY = [int]($canvasH - $circleSize - 10)
            $placedPositions = @()

            for ($i = 0; $i -lt $global:touchTarget; $i++) {
                # Find non-overlapping position
                $attempts = 0
                do {
                    $x = $rnd.Next(10, $maxX)
                    $y = $rnd.Next(10, $maxY)
                    $tooClose = $false
                    foreach ($pos in $placedPositions) {
                        if ([Math]::Abs($x - $pos[0]) -lt ($circleSize + 20) -and [Math]::Abs($y - $pos[1]) -lt ($circleSize + 20)) {
                            $tooClose = $true; break
                        }
                    }
                    $attempts++
                } while ($tooClose -and $attempts -lt 50)
                $placedPositions += , @($x, $y)

                $color = $circleColors[$i % $circleColors.Count]
                $ellipse = New-Object System.Windows.Shapes.Ellipse
                $ellipse.Width = $circleSize
                $ellipse.Height = $circleSize
                $ellipse.Fill = $brushConverter.ConvertFrom($color)
                $ellipse.Cursor = [System.Windows.Input.Cursors]::Hand
                $ellipse.Tag = "circle_$i"

                $shadow = New-Object System.Windows.Media.Effects.DropShadowEffect
                $shadow.Color = [System.Windows.Media.Colors]::White
                $shadow.BlurRadius = 20
                $shadow.ShadowDepth = 0
                $shadow.Opacity = 0.6
                $ellipse.Effect = $shadow

                # Number label inside circle
                $numLabel = New-Object System.Windows.Controls.TextBlock
                $numLabel.Text = ($i + 1).ToString()
                $numLabel.FontSize = 28
                $numLabel.FontWeight = "Bold"
                $numLabel.Foreground = $brushConverter.ConvertFrom("#FFFFFF")
                $numLabel.HorizontalAlignment = "Center"
                $numLabel.VerticalAlignment = "Center"
                $numLabel.IsHitTestVisible = $false

                $ellipse.Add_TouchDown({
                        param($evtSender, $e)
                        & $registerHit $evtSender
                        $e.Handled = $true
                    })

                [System.Windows.Controls.Canvas]::SetLeft($ellipse, $x)
                [System.Windows.Controls.Canvas]::SetTop($ellipse, $y)
                $touchCanvas.Children.Add($ellipse) | Out-Null

                $numLabel.Tag = "circle_${i}_lbl"
                [System.Windows.Controls.Canvas]::SetLeft($numLabel, $x + ($circleSize / 2) - 10)
                [System.Windows.Controls.Canvas]::SetTop($numLabel, $y + ($circleSize / 2) - 18)
                $touchCanvas.Children.Add($numLabel) | Out-Null
            }
        })

    $touchWindow.ShowDialog() | Out-Null

    # Results are asked AFTER the tester window has closed (see note above)
    if ($global:touchOutcome -eq "AllTapped") {
        $res = Show-TouchscreenPassDialog -Message "All $target circles tapped.`nThe touchscreen is working."
        if ($res -eq "OK") { return "Operational" }
        return "Retry"
    }
    if ($global:touchOutcome -eq "Timeout") {
        $res = [DsuMessageBox]::Show(
            "You tapped $($global:touchCount) of $target circles in $seconds seconds.`n`nYes = try again`nNo = mark the touchscreen Defective",
            "Touchscreen Test - Not Completed", "YesNo", "Warning")
        if ($res -eq "Yes") { return "Retry" }
        return "Defective"
    }
    if ($global:touchOutcome -eq "Defective") { return "Defective" }
    return $null
}

# ------------------------------------------------------------
# XAML INTERFACE
# ------------------------------------------------------------
$Global:LogoBase64 = "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
"

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="DSU IT Support - Automated DQA System" Height="780" Width="1100" WindowStartupLocation="CenterScreen" Background="#F0F3F6" FontFamily="Din Medium, Roboto, Helvetica, Arial, sans-serif">
    <Window.Resources>
        <DataTemplate x:Key="ColoredComboBoxTemplate">
            <Border x:Name="bd" Padding="5,0" CornerRadius="4" HorizontalAlignment="Stretch" MinHeight="22">
                <TextBlock x:Name="txt" Text="{Binding}" VerticalAlignment="Center" Foreground="Black" FontSize="13"/>
            </Border>
            <DataTemplate.Triggers>
                <DataTrigger Binding="{Binding}" Value="Defective">
                    <Setter TargetName="bd" Property="Background" Value="#F8D7DA"/>
                    <Setter TargetName="txt" Property="FontWeight" Value="Bold"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding}" Value="Needs Repair">
                    <Setter TargetName="bd" Property="Background" Value="#F8D7DA"/>
                    <Setter TargetName="txt" Property="FontWeight" Value="Bold"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding}" Value="Operational">
                    <Setter TargetName="bd" Property="Background" Value="#D4EDDA"/>
                    <Setter TargetName="txt" Property="FontWeight" Value="SemiBold"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding}" Value="Acceptable">
                    <Setter TargetName="bd" Property="Background" Value="#D4EDDA"/>
                    <Setter TargetName="txt" Property="FontWeight" Value="SemiBold"/>
                </DataTrigger>
            </DataTemplate.Triggers>
        </DataTemplate>
        <!-- Rounded dropdowns (works for the editable Technician Email box too) -->
        <Style TargetType="ComboBox">
            <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
            <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
            <Setter Property="ItemTemplate" Value="{StaticResource ColoredComboBoxTemplate}"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Padding" Value="4,3,28,3"/>
            <Setter Property="VerticalContentAlignment" Value="Stretch"/>
            <Setter Property="Background" Value="#F8FAFC"/>
            <Setter Property="BorderBrush" Value="#C5CBD3"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton x:Name="ToggleButton" Focusable="False" ClickMode="Press"
                                          IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}">
                                <ToggleButton.Template>
                                    <ControlTemplate TargetType="ToggleButton">
                                        <Border x:Name="Bd" CornerRadius="6"
                                                Background="{Binding Background, RelativeSource={RelativeSource AncestorType=ComboBox}}"
                                                BorderBrush="{Binding BorderBrush, RelativeSource={RelativeSource AncestorType=ComboBox}}"
                                                BorderThickness="{Binding BorderThickness, RelativeSource={RelativeSource AncestorType=ComboBox}}">
                                            <Path HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,11,0" Data="M 0 0 L 4.5 4.5 L 9 0" Stroke="#4D4F53" StrokeThickness="1.6"/>
                                        </Border>
                                        <ControlTemplate.Triggers>
                                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#00A9E0"/></Trigger>
                                            <Trigger Property="IsChecked" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#00A9E0"/></Trigger>
                                        </ControlTemplate.Triggers>
                                    </ControlTemplate>
                                </ToggleButton.Template>
                            </ToggleButton>
                            <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False" Margin="{TemplateBinding Padding}"
                                              Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                                              HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="{TemplateBinding VerticalContentAlignment}"/>
                            <TextBox x:Name="PART_EditableTextBox" Style="{x:Null}" FocusVisualStyle="{x:Null}" Visibility="Hidden" Margin="6,0,28,0" VerticalAlignment="Center"
                                     Background="Transparent" BorderThickness="0" Focusable="True" FontSize="{TemplateBinding FontSize}"
                                     IsReadOnly="{Binding IsReadOnly, RelativeSource={RelativeSource TemplatedParent}}"/>
                            <Popup x:Name="PART_Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                <Border Background="White" BorderBrush="#C5CBD3" BorderThickness="1" CornerRadius="6" Margin="0,3,0,0" Padding="3"
                                        MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}"
                                        MaxHeight="{TemplateBinding MaxDropDownHeight}" SnapsToDevicePixels="True">
                                    <ScrollViewer SnapsToDevicePixels="True">
                                        <ItemsPresenter KeyboardNavigation.DirectionalNavigation="Contained"/>
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsEditable" Value="True">
                                <Setter TargetName="PART_EditableTextBox" Property="Visibility" Value="Visible"/>
                                <Setter TargetName="ContentSite" Property="Visibility" Value="Hidden"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.55"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <!-- The dropdown the technician is on gets a blue border and a light blue tint -->
                <Trigger Property="IsKeyboardFocusWithin" Value="True">
                    <Setter Property="BorderBrush" Value="#00A9E0"/>
                    <Setter Property="Background" Value="#EAF7FD"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <!-- Rounded rows inside the open dropdown list -->
        <Style TargetType="ComboBoxItem">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border x:Name="Bd" CornerRadius="4" Padding="2,1" Margin="0,1" Background="Transparent">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="True"><Setter TargetName="Bd" Property="Background" Value="#E6F6FC"/></Trigger>
                            <Trigger Property="IsSelected" Value="True"><Setter TargetName="Bd" Property="Background" Value="#D2EEF9"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Rounded text boxes -->
        <Style TargetType="TextBox">
            <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
            <Setter Property="Background" Value="#F8FAFC"/>
            <Setter Property="BorderBrush" Value="#C5CBD3"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="4,2"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="Bd" CornerRadius="6" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" SnapsToDevicePixels="True">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" Focusable="False" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#00A9E0"/></Trigger>
                            <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#00A9E0"/></Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter TargetName="Bd" Property="Opacity" Value="0.55"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Rounded tabs -->
        <!-- Tabs: both the same off-white as the page; the selected one has a DSU-blue top line and navy text -->
        <Style TargetType="TabItem">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Grid Margin="0,0,4,0">
                            <Border x:Name="Bd" Background="White" BorderBrush="#C5CBD3" BorderThickness="1,1,1,0" CornerRadius="8,8,0,0" Padding="14,7">
                                <ContentPresenter x:Name="Hdr" ContentSource="Header" HorizontalAlignment="Center" VerticalAlignment="Center" TextElement.Foreground="#6C757D"/>
                            </Border>
                            <Border x:Name="Accent" Height="3" VerticalAlignment="Top" CornerRadius="8,8,0,0" Background="Transparent"/>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Accent" Property="Background" Value="#00A9E0"/>
                                <Setter TargetName="Bd" Property="Margin" Value="0,0,0,-1"/>
                                <Setter TargetName="Hdr" Property="TextElement.Foreground" Value="#003552"/>
                                <Setter Property="Panel.ZIndex" Value="1"/>
                            </Trigger>
                            <MultiTrigger>
                                <MultiTrigger.Conditions>
                                    <Condition Property="IsMouseOver" Value="True"/>
                                    <Condition Property="IsSelected" Value="False"/>
                                </MultiTrigger.Conditions>
                                <Setter TargetName="Hdr" Property="TextElement.Foreground" Value="#003552"/>
                                <Setter TargetName="Bd" Property="BorderBrush" Value="#00A9E0"/>
                            </MultiTrigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#00a9e0"/>
            <Setter Property="Foreground" Value="#ffffff"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Padding" Value="5"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="6" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.85"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.7"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="StatCardStyle" TargetType="Border">
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#E3E6EA"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="10"/>
            <Setter Property="Margin" Value="6"/>
            <Setter Property="Padding" Value="16,14"/>
            <Setter Property="Effect">
                <Setter.Value>
                    <DropShadowEffect Color="Black" Direction="270" ShadowDepth="2" BlurRadius="12" Opacity="0.08"/>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Rounded page frame under the tabs -->
        <Style TargetType="TabControl">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabControl">
                        <Grid KeyboardNavigation.TabNavigation="Local">
                            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                            <TabPanel x:Name="HeaderPanel" Grid.Row="0" Panel.ZIndex="1" Margin="0" IsItemsHost="True" KeyboardNavigation.TabIndex="1" Background="Transparent"/>
                            <Border Grid.Row="1" Background="{TemplateBinding Background}" BorderBrush="#C5CBD3" BorderThickness="1" CornerRadius="0,10,10,10" Padding="4,6,4,6" SnapsToDevicePixels="True">
                                <ContentPresenter x:Name="PART_SelectedContentHost" ContentSource="SelectedContent"/>
                            </Border>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Slim rounded scrollbars (no arrow buttons; drag, click the track or use the mouse wheel) -->
        <Style x:Key="ScrollPageButton" TargetType="RepeatButton">
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="IsTabStop" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="RepeatButton"><Border Background="Transparent"/></ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="ScrollThumb" TargetType="Thumb">
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="IsTabStop" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Thumb">
                        <Border x:Name="tb" CornerRadius="4" Background="#B8BEC4"/>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="tb" Property="Background" Value="#8E959C"/></Trigger>
                            <Trigger Property="IsDragging" Value="True"><Setter TargetName="tb" Property="Background" Value="#00A9E0"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Width" Value="12"/>
            <Setter Property="MinWidth" Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollBar">
                        <Border CornerRadius="6" Background="#EEF0F2" Margin="2,2,2,2">
                            <Track x:Name="PART_Track" IsDirectionReversed="True">
                                <Track.DecreaseRepeatButton><RepeatButton Style="{StaticResource ScrollPageButton}" Command="ScrollBar.PageUpCommand"/></Track.DecreaseRepeatButton>
                                <Track.Thumb><Thumb Style="{StaticResource ScrollThumb}" Margin="1"/></Track.Thumb>
                                <Track.IncreaseRepeatButton><RepeatButton Style="{StaticResource ScrollPageButton}" Command="ScrollBar.PageDownCommand"/></Track.IncreaseRepeatButton>
                            </Track>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="Orientation" Value="Horizontal">
                    <Setter Property="Width" Value="Auto"/>
                    <Setter Property="MinWidth" Value="0"/>
                    <Setter Property="Height" Value="12"/>
                    <Setter Property="MinHeight" Value="12"/>
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="ScrollBar">
                                <Border CornerRadius="6" Background="#EEF0F2" Margin="2,2,2,2">
                                    <Track x:Name="PART_Track" IsDirectionReversed="False">
                                        <Track.DecreaseRepeatButton><RepeatButton Style="{StaticResource ScrollPageButton}" Command="ScrollBar.PageLeftCommand"/></Track.DecreaseRepeatButton>
                                        <Track.Thumb><Thumb Style="{StaticResource ScrollThumb}" Margin="1"/></Track.Thumb>
                                        <Track.IncreaseRepeatButton><RepeatButton Style="{StaticResource ScrollPageButton}" Command="ScrollBar.PageRightCommand"/></Track.IncreaseRepeatButton>
                                    </Track>
                                </Border>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                </Trigger>
            </Style.Triggers>
        </Style>
        <!-- Dashboard table: navy header row with sort arrow; column resize grips kept -->
        <Style x:Key="HeaderGripper" TargetType="Thumb">
            <Setter Property="Width" Value="6"/>
            <Setter Property="Cursor" Value="SizeWE"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Thumb"><Border Background="Transparent"/></ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#003552"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Padding" Value="10,9,10,9"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="DataGridColumnHeader">
                        <Grid>
                            <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="#1D4E6C" BorderThickness="0,0,1,0" Padding="{TemplateBinding Padding}">
                                <DockPanel>
                                    <Path x:Name="SortArrow" DockPanel.Dock="Right" Visibility="Collapsed" Width="8" Height="5" Margin="8,0,0,0"
                                          Fill="#00A9E0" Stretch="Fill" Data="M 0 0 L 1 0 L 0.5 1 Z" VerticalAlignment="Center" RenderTransformOrigin="0.5,0.5"/>
                                    <ContentPresenter VerticalAlignment="Center" HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" RecognizesAccessKey="False"/>
                                </DockPanel>
                            </Border>
                            <Thumb x:Name="PART_LeftHeaderGripper" HorizontalAlignment="Left" Style="{StaticResource HeaderGripper}"/>
                            <Thumb x:Name="PART_RightHeaderGripper" HorizontalAlignment="Right" Style="{StaticResource HeaderGripper}"/>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#0B4A6E"/></Trigger>
                            <Trigger Property="SortDirection" Value="Ascending">
                                <Setter TargetName="SortArrow" Property="Visibility" Value="Visible"/>
                                <Setter TargetName="SortArrow" Property="RenderTransform"><Setter.Value><RotateTransform Angle="180"/></Setter.Value></Setter>
                            </Trigger>
                            <Trigger Property="SortDirection" Value="Descending">
                                <Setter TargetName="SortArrow" Property="Visibility" Value="Visible"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Dashboard cells: padded, vertically centred; the row colour (PASSED/FAILED) shows through -->
        <Style TargetType="DataGridCell">
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="DataGridCell">
                        <Border x:Name="Bd" Background="Transparent" Padding="10,0">
                            <ContentPresenter VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True"><Setter Property="Foreground" Value="Black"/></Trigger>
            </Style.Triggers>
        </Style>
        <!-- "I have wiped down the computer: YES" check box: rounded box, green fill + white check mark when ticked -->
        <Style x:Key="WipeCheck" TargetType="CheckBox">
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CheckBox">
                        <StackPanel Orientation="Horizontal" Background="Transparent">
                            <Grid Width="24" Height="24" VerticalAlignment="Center">
                                <Border x:Name="box" CornerRadius="6" Background="White" BorderBrush="#8E959C" BorderThickness="2"/>
                                <Path x:Name="tick" Data="M 6 12.5 L 10.5 17 L 18.5 7.5" Stroke="White" StrokeThickness="2.6"
                                      StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round" Visibility="Collapsed"/>
                            </Grid>
                            <ContentPresenter Margin="9,0,0,0" VerticalAlignment="Center" RecognizesAccessKey="True"/>
                        </StackPanel>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="box" Property="BorderBrush" Value="#28A745"/></Trigger>
                            <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="box" Property="BorderBrush" Value="#00A9E0"/></Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="box" Property="Background" Value="#28A745"/>
                                <Setter TargetName="box" Property="BorderBrush" Value="#28A745"/>
                                <Setter TargetName="tick" Property="Visibility" Value="Visible"/>
                            </Trigger>
                            <MultiTrigger>
                                <MultiTrigger.Conditions>
                                    <Condition Property="IsChecked" Value="True"/>
                                    <Condition Property="IsKeyboardFocused" Value="True"/>
                                </MultiTrigger.Conditions>
                                <Setter TargetName="box" Property="BorderBrush" Value="#00A9E0"/>
                            </MultiTrigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.5"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Round header buttons (settings gear, timer play/pause): soft white fill on the navy header, DSU blue on hover -->
        <Style x:Key="HeaderRoundBtn" TargetType="Button">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontFamily" Value="Segoe MDL2 Assets"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bg" Background="#26FFFFFF" BorderBrush="#40FFFFFF" BorderThickness="1" CornerRadius="999">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bg" Property="Background" Value="#00A9E0"/><Setter TargetName="bg" Property="BorderBrush" Value="#00A9E0"/></Trigger>
                            <Trigger Property="IsPressed" Value="True"><Setter TargetName="bg" Property="Background" Value="#0087B3"/><Setter TargetName="bg" Property="BorderBrush" Value="#0087B3"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
        
        <Border Grid.Row="0" Background="#003552" Padding="20,15">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Vertical" VerticalAlignment="Center">
                    <TextBlock Text="Device Quality Assurance System" FontSize="26" FontWeight="Bold" Foreground="#00a9e0"/>
                    <TextBlock Text="Dakota State University - IT Support" FontSize="14" Foreground="White" Margin="0,2,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center" HorizontalAlignment="Right">
                    <Button x:Name="SettingsBtn" Style="{StaticResource HeaderRoundBtn}" Content="&#xE713;" FontSize="17" Width="40" Height="40" Margin="0,0,16,0" ToolTip="Settings" VerticalAlignment="Center"/>
                    <Border x:Name="InspectionTimerPanel" CornerRadius="24" Background="#1AFFFFFF" BorderBrush="#33FFFFFF" BorderThickness="1" Padding="16,5,6,5" Margin="0,0,16,0" VerticalAlignment="Center">
                        <StackPanel Orientation="Horizontal">
                            <StackPanel VerticalAlignment="Center" Margin="0,0,12,0">
                                <TextBlock Text="ELAPSED TIME" FontSize="10" FontWeight="SemiBold" Foreground="#9CC3D8" HorizontalAlignment="Left"/>
                                <TextBlock x:Name="LiveTimerLabel" Text="00:00" FontWeight="Bold" FontSize="22" Foreground="#DC3545" HorizontalAlignment="Left" Margin="0,-2,0,0"/>
                            </StackPanel>
                            <Button x:Name="PauseTimerBtn" Style="{StaticResource HeaderRoundBtn}" Content="&#xE768;" FontSize="13" Width="36" Height="36" ToolTip="Pause/Resume Timer" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Border>
                    <Image x:Name="LogoImage" Height="50" Margin="0,0,10,0" RenderOptions.BitmapScalingMode="HighQuality"/>
                </StackPanel>
            </Grid>
        </Border>

        <TabControl x:Name="MainTabControl" Grid.Row="1" Margin="10" Background="White">
            
            <TabItem Header=" New Inspection (Automated) " FontSize="14" FontWeight="SemiBold">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel>
                        <Border Background="#FFF3CD" BorderBrush="#FFEEBA" BorderThickness="1" CornerRadius="5" Margin="20,15,20,0" Padding="12">
                            <TextBlock Text="PRE-CHECK: Make sure to connect to Guest Wi-Fi, plug in the charger, and see if the microphone is ON or sound level is muted." 
                                       Foreground="#856404" FontWeight="Bold" FontSize="14" TextWrapping="Wrap"/>
                        </Border>

                        <Grid Margin="20,15,20,20">
                            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                            
                            <StackPanel Grid.Column="0" Margin="0,0,15,0">
                                <TextBlock Text="System Automation" FontSize="18" FontWeight="Bold" Foreground="#003552" Margin="0,0,0,10"/>
                                
                                <TextBlock Text="Technician Email (Triggers Auto-Detect):" FontWeight="SemiBold"/>
                                <ComboBox x:Name="TechEmailInput" IsEditable="True" Height="28" Margin="0,0,0,4" TabIndex="1"/>
                                <TextBlock x:Name="DetectStatus" Text="Detecting hardware..." FontStyle="Italic" FontSize="12" Foreground="#00779E" Margin="0,0,0,4" Visibility="Collapsed"/>
                                <Border Height="6"/>

                                <StackPanel Orientation="Horizontal" Margin="0,0,0,0">
                                    <TextBlock Text="Serial Number (Auto):" FontWeight="SemiBold"/>
                                </StackPanel>
                                <TextBox x:Name="SerialInput" Height="28" Margin="0,0,0,10" TabIndex="2"/>

                                <TextBlock Text="Charging Status (Auto):" FontWeight="SemiBold"/><ComboBox x:Name="cbCharging" Height="28" Margin="0,0,0,10" TabIndex="3"/>
                                <TextBlock Text="Network Adapters (Auto):" FontWeight="SemiBold"/><ComboBox x:Name="cbNetwork" Height="28" Margin="0,0,0,15" TabIndex="4"/>

                                <TextBlock Text="Interactive Tests" FontSize="18" FontWeight="Bold" Foreground="#003552" Margin="0,10,0,5"/>
                                <Button x:Name="BtnAutoTest" Content="Auto-Test Hardware" Height="35" Width="200" HorizontalAlignment="Left" Margin="0,0,0,2" Background="#00A9E0" Foreground="White" FontWeight="Bold" TabIndex="5"/>
                                <TextBlock Text="Click the button to automatically execute sequentially interactive tests." Foreground="Gray" FontSize="11" Margin="0,0,0,15" TextWrapping="Wrap"/>
                                
                                <StackPanel Margin="0,0,0,15">
                                    <Button x:Name="BtnTestCamera" Content="1. Launch Camera" Margin="0,0,0,5" Height="30" Width="200" HorizontalAlignment="Left" Background="#00A9E0" Foreground="White" FontWeight="Bold" TabIndex="7"/>
                                    <TextBlock Text="Camera Status:" FontWeight="SemiBold"/>
                                    <ComboBox x:Name="cbCamera" Height="28" Margin="0,0,0,15" TabIndex="6"/>
                                    
                                    <Button x:Name="BtnTestAudio" Content="2. Test Audio/Mic" Margin="0,0,0,5" Height="30" Width="200" HorizontalAlignment="Left" Background="#00A9E0" Foreground="White" FontWeight="Bold" TabIndex="10"/>
                                    <TextBlock Text="Audio Status:" FontWeight="SemiBold"/>
                                    <ComboBox x:Name="cbAudio" Height="28" Margin="0,0,0,5" TabIndex="8"/>
                                    <TextBlock Text="Mic Status:" FontWeight="SemiBold"/>
                                    <ComboBox x:Name="cbMic" Height="28" Margin="0,0,0,15" TabIndex="9"/>
                                    <TextBlock x:Name="AudioStatusLabel" Visibility="Collapsed" FontWeight="Bold" Margin="0,0,0,4" Height="22" Padding="5,0,5,0" VerticalAlignment="Center"/>
                                    <!-- Live mic level bar (visible only while recording) -->
                                    <StackPanel x:Name="MicMeter" Orientation="Horizontal" Margin="0,0,0,11" Visibility="Collapsed">
                                        <TextBlock Text="Mic level" FontSize="11" Foreground="#4D4F53" VerticalAlignment="Center" Margin="0,0,8,0"/>
                                        <Border x:Name="MicMeterTrack" Width="220" Height="8" CornerRadius="4" Background="#E3E6EA" VerticalAlignment="Center">
                                            <Border x:Name="MicMeterFill" Width="0" HorizontalAlignment="Left" CornerRadius="4" Background="#28A745"/>
                                        </Border>
                                    </StackPanel>
                                    
                                    <Button x:Name="BtnTestKeys" Content="3. Keyboard Web Test" Margin="0,0,0,5" Height="30" Width="200" HorizontalAlignment="Left" Background="#00A9E0" Foreground="White" FontWeight="Bold" TabIndex="12"/>
                                    <TextBlock Text="Keyboard Status:" FontWeight="SemiBold"/>
                                    <ComboBox x:Name="cbKeyboard" Height="28" Margin="0,0,0,15" TabIndex="11"/>
                                    
                                    <Button x:Name="BtnTestTouch" Content="4. Test Touchscreen" Margin="0,0,0,5" Height="30" Width="200" HorizontalAlignment="Left" Background="#00A9E0" Foreground="White" FontWeight="Bold" TabIndex="12"/>
                                    <TextBlock Text="Touchscreen Status:" FontWeight="SemiBold"/>
                                    <ComboBox x:Name="cbTouch" Height="28" Margin="0,0,0,0" TabIndex="13"/>
                                </StackPanel>

                            </StackPanel>

                            <StackPanel Grid.Column="1" Margin="15,0,0,0">
                                <Grid Margin="0,0,0,10">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Grid.Column="0" Text="External Ports (Auto-Detected)" FontSize="18" FontWeight="Bold" Foreground="#003552" VerticalAlignment="Center" Margin="0"/>
                                    <Button x:Name="BtnAutoFill" Grid.Column="1" Content="Auto-Fill Remaining to Operational" Height="30" HorizontalAlignment="Right" Background="#00A9E0" Foreground="White" Margin="10,0,0,0" Padding="10,0,10,0" TabIndex="13"/>
                                </Grid>
                                <StackPanel Margin="0,0,0,15">
                                    <TextBlock Text="USB Ports:"/><ComboBox x:Name="cbUSB" Height="26" Margin="0,0,0,5" TabIndex="14"/>
                                    <TextBlock Text="Video Ports (Ext. Monitor):"/><ComboBox x:Name="cbVideo" Height="26" TabIndex="15"/>
                                </StackPanel>

                                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                                    <TextBlock Text="Other Hardware" FontSize="18" FontWeight="Bold" Foreground="#003552" VerticalAlignment="Center"/>
                                </StackPanel>
                                <StackPanel Margin="0,0,0,15">
                                    <TextBlock Text="Mouse/Trackpad:"/><ComboBox x:Name="cbMouse" Height="26" Margin="0,0,0,5" TabIndex="16"/>
                                    <TextBlock Text="Screen (Visual):"/><ComboBox x:Name="cbScreen" Height="26" TabIndex="17"/>
                                </StackPanel>

                                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                                    <TextBlock Text="Cosmetic Inspection" FontSize="18" FontWeight="Bold" Foreground="#003552" VerticalAlignment="Center"/>
                                </StackPanel>
                                
                                <TextBlock Text="Palm Rest:"/><ComboBox x:Name="cbPalm" Height="26" Margin="0,0,0,5" TabIndex="19"/>
                                <TextBlock Text="Backplate:"/><ComboBox x:Name="cbBackplate" Height="26" Margin="0,0,0,5" TabIndex="20"/>
                                <TextBlock Text="Base and Vents:"/><ComboBox x:Name="cbBase" Height="26" Margin="0,0,0,5" TabIndex="21"/>
                                <TextBlock Text="Hinge:"/><ComboBox x:Name="cbHinge" Height="26" Margin="0,0,0,15" TabIndex="22"/>

                                <StackPanel Orientation="Horizontal" Margin="0,0,0,2">
                                    <TextBlock Text="Additional Notes:" FontSize="16" FontWeight="Bold" Foreground="#003552"/>
                                    <TextBlock x:Name="NotesCounter" Text="(0 chars)" Foreground="#747676" Margin="10,2,0,0" VerticalAlignment="Bottom"/>
                                </StackPanel>
                                <TextBox x:Name="NotesInput" Height="60" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" Margin="0,0,0,15" TabIndex="23"/>
                                
                                <StackPanel x:Name="AdditionalTimePanel" Orientation="Horizontal" Margin="0,0,0,10">
                                    <TextBlock Text="Additional Pre-DQA Time (mins):" VerticalAlignment="Center" FontWeight="SemiBold" Margin="0,0,10,0"/>
                                    <TextBox x:Name="AdditionalTimeInput" Width="60" Height="26" VerticalContentAlignment="Center" HorizontalContentAlignment="Center" TabIndex="24"/>
                                </StackPanel>
                                <StackPanel x:Name="DurationCorrectionPanel" Visibility="Collapsed" Margin="0,0,0,10">
                                    <StackPanel Orientation="Horizontal">
                                        <TextBlock Text="Correct total duration (min)" VerticalAlignment="Center" FontWeight="SemiBold" Margin="0,0,10,0"/>
                                        <TextBox x:Name="DurationCorrectionInput" Width="80" Height="26" VerticalContentAlignment="Center" HorizontalContentAlignment="Center" AutomationProperties.Name="Correct total duration in minutes"/>
                                    </StackPanel>
                                    <TextBlock x:Name="SavedDurationHint" Foreground="#5A5F66" TextWrapping="Wrap" Margin="0,4,0,0"/>
                                </StackPanel>
                                
                                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,0,0,10">
                                    <TextBlock Text="I have wiped down the computer :" FontSize="16" Foreground="#003552" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                    <CheckBox x:Name="rbWipeYes" Style="{StaticResource WipeCheck}" Content="YES" Foreground="#28A745" FontWeight="Bold" FontSize="16" Margin="0,0,10,0" VerticalAlignment="Center" TabIndex="25"/>
                                </StackPanel>

                                <Button x:Name="SaveInspectionBtn" Content="SAVE INSPECTION" Height="45" FontSize="16" Margin="0,0,0,8" IsEnabled="False" Opacity="0.5" TabIndex="26"/>
                                <Button x:Name="ResetFormBtn" Content="Reset / Clear Form" Height="32" FontSize="13" Background="#4d4f53" Foreground="White" FontWeight="Bold" BorderThickness="0" TabIndex="27"/>
                                <TextBlock x:Name="LastSavedLabel" Text=" " Foreground="#28A745" FontWeight="Bold" Margin="0,10,0,0" HorizontalAlignment="Center"/>
                            </StackPanel>
                        </Grid>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <TabItem Header=" Dashboard &amp; History " FontSize="14" FontWeight="SemiBold">
                <Grid Margin="10">
                    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                    <Grid Grid.Row="0" Margin="0,0,0,15">
                        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                        <Border Grid.Column="0" Style="{StaticResource StatCardStyle}"><DockPanel><Border DockPanel.Dock="Left" Width="46" Height="46" CornerRadius="23" Background="#E6EEF3" Margin="0,0,14,0" VerticalAlignment="Center"><TextBlock Text="&#xE8FD;" FontFamily="Segoe MDL2 Assets" FontSize="20" Foreground="#003552" HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><StackPanel VerticalAlignment="Center"><TextBlock Text="TOTAL RUNS" FontSize="11" FontWeight="SemiBold" Foreground="#6C757D"/><TextBlock x:Name="StatTotal" FontSize="28" FontWeight="Bold" Foreground="#003552"/></StackPanel></DockPanel></Border>
                        <Border Grid.Column="1" Style="{StaticResource StatCardStyle}"><DockPanel><Border DockPanel.Dock="Left" Width="46" Height="46" CornerRadius="23" Background="#E6F4EA" Margin="0,0,14,0" VerticalAlignment="Center"><TextBlock Text="&#xE73E;" FontFamily="Segoe MDL2 Assets" FontSize="20" Foreground="#1E7B3A" HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><StackPanel VerticalAlignment="Center"><TextBlock Text="PASSED" FontSize="11" FontWeight="SemiBold" Foreground="#6C757D"/><TextBlock x:Name="StatPassed" FontSize="28" FontWeight="Bold" Foreground="#1E7B3A"/></StackPanel></DockPanel></Border>
                        <Border Grid.Column="2" Style="{StaticResource StatCardStyle}"><DockPanel><Border DockPanel.Dock="Left" Width="46" Height="46" CornerRadius="23" Background="#FBE9EA" Margin="0,0,14,0" VerticalAlignment="Center"><TextBlock Text="&#xE711;" FontFamily="Segoe MDL2 Assets" FontSize="20" Foreground="#C82333" HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><StackPanel VerticalAlignment="Center"><TextBlock Text="FAILED" FontSize="11" FontWeight="SemiBold" Foreground="#6C757D"/><TextBlock x:Name="StatFailed" FontSize="28" FontWeight="Bold" Foreground="#C82333"/></StackPanel></DockPanel></Border>
                        <Border Grid.Column="3" Style="{StaticResource StatCardStyle}"><DockPanel><Border DockPanel.Dock="Left" Width="46" Height="46" CornerRadius="23" Background="#E6F6FC" Margin="0,0,14,0" VerticalAlignment="Center"><TextBlock Text="&#xE916;" FontFamily="Segoe MDL2 Assets" FontSize="20" Foreground="#00779E" HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><StackPanel VerticalAlignment="Center"><TextBlock Text="AVG DURATION" FontSize="11" FontWeight="SemiBold" Foreground="#6C757D"/><StackPanel Orientation="Horizontal"><TextBlock x:Name="StatAvg" FontSize="28" FontWeight="Bold" Foreground="#00779E"/><TextBlock Text="min" FontSize="14" FontWeight="SemiBold" Foreground="#6C757D" VerticalAlignment="Bottom" Margin="5,0,0,6"/></StackPanel></StackPanel></DockPanel></Border>
                    </Grid>
                    <!-- Toolbar: search + filter on the left; Edit/Delete appear only when a row is selected; Export on the right.
                         Row shortcuts: double-click or Enter = Edit, right-click = Edit/Delete menu, Delete key = Delete. Clear All lives in Settings. -->
                    <DockPanel Grid.Row="1" Margin="5,0,5,10" LastChildFill="False">
                        <Button x:Name="ExportBtn" DockPanel.Dock="Right" Content="Export to CSV" Width="130" Height="32" Background="#003552" Foreground="White" FontWeight="Bold" BorderThickness="0" Cursor="Hand" ToolTip="Export the records shown below (Ctrl+E)"/>
                        <StackPanel DockPanel.Dock="Left" Orientation="Horizontal">
                            <!-- Search: rounded box with icon, placeholder and a clear (X) button -->
                            <Border x:Name="SearchFrame" Width="230" Height="32" Margin="0,0,10,0" CornerRadius="8" Background="White" BorderThickness="1">
                                <Border.Style>
                                    <Style TargetType="Border">
                                        <Setter Property="BorderBrush" Value="#C5CBD3"/>
                                        <Style.Triggers>
                                            <Trigger Property="IsMouseOver" Value="True"><Setter Property="BorderBrush" Value="#00A9E0"/></Trigger>
                                            <Trigger Property="IsKeyboardFocusWithin" Value="True"><Setter Property="BorderBrush" Value="#00A9E0"/></Trigger>
                                        </Style.Triggers>
                                    </Style>
                                </Border.Style>
                                <Grid>
                                    <TextBlock Text="&#xE721;" FontFamily="Segoe MDL2 Assets" FontSize="13" Foreground="#6C757D" Margin="11,0,0,0" HorizontalAlignment="Left" VerticalAlignment="Center" IsHitTestVisible="False"/>
                                    <TextBox x:Name="SearchBox" Background="Transparent" BorderThickness="0" Padding="0" Margin="32,0,30,0" VerticalContentAlignment="Center" FontSize="13"/>
                                    <TextBlock x:Name="SearchPlaceholder" Text="Search serial number..." Foreground="#8A8F95" FontSize="13" Margin="34,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False"/>
                                    <Button x:Name="SearchClearBtn" Content="&#xE711;" FontFamily="Segoe MDL2 Assets" FontSize="10" FontWeight="Normal" Width="22" Height="22" Padding="0" Margin="0,0,5,0"
                                            HorizontalAlignment="Right" VerticalAlignment="Center" Background="#EEF0F2" Foreground="#4D4F53" Visibility="Collapsed" ToolTip="Clear search"/>
                                </Grid>
                            </Border>
                            <!-- Filter: rounded dropdown with a funnel icon -->
                            <Grid Width="150" Height="32" Margin="0,0,16,0">
                                <ComboBox x:Name="FilterBox" Height="32" Padding="32,0,28,0" VerticalContentAlignment="Center" ToolTip="Show all, only PASSED or only FAILED">
                                    <ComboBoxItem Content="All Results" IsSelected="True"/><ComboBoxItem Content="PASSED only"/><ComboBoxItem Content="FAILED only"/>
                                </ComboBox>
                                <TextBlock Text="&#xE71C;" FontFamily="Segoe MDL2 Assets" FontSize="13" Foreground="#6C757D" Margin="12,0,0,0" HorizontalAlignment="Left" VerticalAlignment="Center" IsHitTestVisible="False"/>
                            </Grid>
                            <StackPanel x:Name="RowActions" Orientation="Horizontal" Visibility="Collapsed">
                                <Button x:Name="EditBtn" Content="Edit" Width="80" Height="32" Background="#17A2B8" Foreground="White" FontWeight="SemiBold" BorderThickness="0" Margin="0,0,8,0" Cursor="Hand" ToolTip="Edit the selected record (or double-click the row)"/>
                                <Button x:Name="DeleteRowBtn" Content="Delete" Width="80" Height="32" Background="#FD7E14" Foreground="White" FontWeight="SemiBold" BorderThickness="0" Cursor="Hand" ToolTip="Delete the selected record (PIN required)"/>
                            </StackPanel>
                        </StackPanel>
                    </DockPanel>
                    <!-- Rounded table: frame + a rounded mask that clips the grid (headers and row colours) to the corners -->
                    <Border Grid.Row="2" Margin="5" CornerRadius="10" BorderBrush="#C5CBD3" BorderThickness="1" Background="White" SnapsToDevicePixels="True">
                        <Grid>
                            <Border x:Name="ResultsGridMask" CornerRadius="9" Background="White"/>
                            <DataGrid x:Name="ResultsGrid" AutoGenerateColumns="True" IsReadOnly="True" Background="White" BorderThickness="0" HeadersVisibility="Column" GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="#E3E6EA" RowHeight="32" SelectionMode="Single"
                                      ScrollViewer.HorizontalScrollBarVisibility="Auto" ScrollViewer.PanningMode="Both">
                                <DataGrid.OpacityMask>
                                    <VisualBrush Visual="{Binding ElementName=ResultsGridMask}"/>
                                </DataGrid.OpacityMask>
                                <DataGrid.RowStyle>
                                    <Style TargetType="DataGridRow">
                                        <Style.Triggers>
                                            <DataTrigger Binding="{Binding FinalStatus}" Value="PASSED">
                                                <Setter Property="Background" Value="#D4EDDA"/>
                                            </DataTrigger>
                                            <DataTrigger Binding="{Binding FinalStatus}" Value="FAILED">
                                                <Setter Property="Background" Value="#F8D7DA"/>
                                            </DataTrigger>
                                            <!-- Selected row: light DSU blue (listed last so it wins over the PASSED/FAILED colours) -->
                                            <Trigger Property="IsSelected" Value="True">
                                                <Setter Property="Background" Value="#BFE6F6"/>
                                            </Trigger>
                                        </Style.Triggers>
                                    </Style>
                                </DataGrid.RowStyle>
                            </DataGrid>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>
            
        </TabControl>
    </Grid>
</Window>
"@

$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
Set-DqaFocusVisual $window
$window.WindowState = 'Maximized'

try {
    $logoBytes = [Convert]::FromBase64String($Global:LogoBase64)
    $ms = New-Object System.IO.MemoryStream($logoBytes, 0, $logoBytes.Length)
    $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
    $bitmap.BeginInit()
    $bitmap.StreamSource = $ms
    $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $bitmap.EndInit()
    $bitmap.Freeze()
    $ms.Close()
    $window.FindName("LogoImage").Source = $bitmap
}
catch { Write-Host "Failed to load base64 logo" }

# ------------------------------------------------------------
# MAP UI ELEMENTS
# ------------------------------------------------------------
$techEmailInput = $window.FindName("TechEmailInput")
$serialInput = $window.FindName("SerialInput")
$notesInput = $window.FindName("NotesInput")
$notesCounter = $window.FindName("NotesCounter")
$audioStatusLabel = $window.FindName("AudioStatusLabel")
$lastSavedLabel = $window.FindName("LastSavedLabel")
$resultsGrid = $window.FindName("ResultsGrid")
$searchBox = $window.FindName("SearchBox")
$searchPlaceholder = $window.FindName("SearchPlaceholder")
$saveInspectionBtn = $window.FindName("SaveInspectionBtn")
$additionalTimeInput = $window.FindName("AdditionalTimeInput")
$durationCorrectionInput = $window.FindName("DurationCorrectionInput")
$liveTimerLabel = $window.FindName("LiveTimerLabel")
$pauseTimerBtn = $window.FindName("PauseTimerBtn")

$pauseTimerBtn.Add_Click({
        if ($Global:isTimerPaused) {
            $Global:sessionStopwatch.Start()
            $Global:isTimerPaused = $false
            $pauseTimerBtn.Content = "$([char]0xE769)"
            $liveTimerLabel.Foreground = "#FFC107"
        
            # Immediate UI update when resuming
            $elapsed = $Global:sessionStopwatch.Elapsed
            $mins = [int][math]::Floor($elapsed.TotalMinutes)
            $secs = $elapsed.Seconds
            $liveTimerLabel.Text = "{0:00}:{1:00}" -f $mins, $secs
        
            if ($null -ne $Global:uiTimer) { $Global:uiTimer.Start() }
        }
        else {
            $Global:sessionStopwatch.Stop()
            $Global:isTimerPaused = $true
            $pauseTimerBtn.Content = "$([char]0xE768)"
            $liveTimerLabel.Foreground = "#DC3545"
            if ($null -ne $Global:uiTimer) { $Global:uiTimer.Stop() }
        }
    })

$autoResumeAction = {
    param($src, $e)
    if ($null -eq $Global:EditingRecord -and $Global:isTimerPaused) {
        $Global:sessionStopwatch.Start()
        $Global:isTimerPaused = $false
        $window.FindName("PauseTimerBtn").Content = "$([char]0xE769)"
        $window.FindName("LiveTimerLabel").Foreground = "#FFC107"
        if ($null -ne $Global:uiTimer) { $Global:uiTimer.Start() }
    }
}
$window.FindName("MainTabControl").Add_PreviewMouseDown($autoResumeAction)
$window.FindName("MainTabControl").Add_PreviewKeyDown($autoResumeAction)

if ($null -ne $Global:uiTimer) {
    $Global:uiTimer.Stop()
}
$Global:uiTimer = New-Object System.Windows.Threading.DispatcherTimer
$Global:uiTimer.Interval = [TimeSpan]::FromSeconds(1)
$Global:uiTimer.Add_Tick({
        if (-not $Global:isTimerPaused) {
            $elapsed = $Global:sessionStopwatch.Elapsed
            $mins = [int][math]::Floor($elapsed.TotalMinutes)
            $secs = $elapsed.Seconds
            $liveTimerLabel.Text = "{0:00}:{1:00}" -f $mins, $secs
        }
    })
$Global:uiTimer.Start()

$techEmailInput.ItemsSource = $Global:SavedEmails

$hwOptions = @("", "Operational", "Defective", "Not Applicable")
$cosmeticOptions = @("", "Acceptable", "Needs Repair", "Not Applicable")

# Set Tab Order for Smooth Keyboard Navigation
$tabOrder = @(
    "TechEmailInput", "SerialInput", "cbCharging", "cbNetwork", 
    "BtnAutoTest", "BtnTestCamera", "cbCamera", "BtnTestAudio", "cbAudio", "cbMic", 
    "BtnTestKeys", "cbKeyboard", "BtnTestTouch", "cbTouch",
    "BtnAutoFill", "cbUSB", "cbVideo", "cbMouse", "cbScreen", 
    "cbPalm", "cbBackplate", "cbBase", "cbHinge", 
    "NotesInput", "AdditionalTimeInput", "DurationCorrectionInput", "rbWipeYes",
    "SaveInspectionBtn", "ResetFormBtn"
)
for ($i = 0; $i -lt $tabOrder.Count; $i++) {
    $ctrl = $window.FindName($tabOrder[$i])
    if ($ctrl) { $ctrl.TabIndex = ($i + 1) }
}

$bc = New-Object System.Windows.Media.BrushConverter
$Global:redBgBrush = $bc.ConvertFrom("#F8D7DA")
$Global:greenBgBrush = $bc.ConvertFrom("#D4EDDA")
$Global:whiteBgBrush = $bc.ConvertFrom("#F8FAFC")
$Global:blackTextBrush = $bc.ConvertFrom("#000000")

$hwBoxes = @("cbCharging", "cbScreen", "cbTouch", "cbNetwork", "cbKeyboard", "cbMouse", "cbVideo", "cbAudio", "cbMic", "cbCamera", "cbUSB")
foreach ($box in $hwBoxes) { $window.FindName($box).ItemsSource = $hwOptions; $window.FindName($box).SelectedIndex = 0 }

$cosmeticBoxes = @("cbPalm", "cbBackplate", "cbBase", "cbHinge")
foreach ($box in $cosmeticBoxes) { $window.FindName($box).ItemsSource = $cosmeticOptions; $window.FindName($box).SelectedIndex = 0 }


# ------------------------------------------------------------
# FORM COMPLETION CHECKER (Disables Save until complete)
# ------------------------------------------------------------
$script:InspectionGeneration = 0
$script:AutoDetectGeneration = $null
$script:AutoDetectPending = 0
$script:InspectionStorageFailure = ""

function Test-FormCompletion {
    $allFilled = $script:AutoDetectPending -eq 0 -and $Global:audioState -le 0
    
    if ([string]::IsNullOrWhiteSpace($techEmailInput.Text) -or [string]::IsNullOrWhiteSpace($serialInput.Text)) { 
        $allFilled = $false 
    }
    
    $reqBoxes = @("cbCharging", "cbScreen", "cbTouch", "cbNetwork", "cbKeyboard", "cbMouse", "cbVideo", "cbAudio", "cbMic", "cbCamera", "cbUSB", "cbPalm", "cbBackplate", "cbBase", "cbHinge")
    foreach ($box in $reqBoxes) {
        $selectedVal = $window.FindName($box).SelectedItem
        if ([string]::IsNullOrWhiteSpace($selectedVal) -or $selectedVal -eq "") { 
            $allFilled = $false 
        }
    }
    
    if ($window.FindName("rbWipeYes").IsChecked -ne $true) {
        $allFilled = $false
    }
    
    if ($allFilled) {
        $saveInspectionBtn.IsEnabled = $true
        $saveInspectionBtn.Opacity = 1.0
    }
    else {
        $saveInspectionBtn.IsEnabled = $false
        $saveInspectionBtn.Opacity = 0.5
    }
}

# ------------------------------------------------------------
# QOL: AUTO-FOCUS NEXT EMPTY COMBOBOX AFTER SELECTION
# ------------------------------------------------------------
$Global:ComboBoxOrder = @(
    "cbCharging", "cbNetwork", "cbCamera", "cbAudio", "cbMic", 
    "cbKeyboard", "cbTouch", "cbUSB", "cbVideo", "cbMouse", 
    "cbScreen", "cbPalm", "cbBackplate", "cbBase", "cbHinge"
)

function Move-ToNextEmptyComboBox {
    $focused = [System.Windows.Input.Keyboard]::FocusedElement
    if ($focused -isnot [System.Windows.Controls.ComboBox]) { return }
    $currentName = $focused.Name
    
    # Only auto-advance if a real value was selected (not blank)
    $selectedVal = $focused.SelectedItem
    if ([string]::IsNullOrWhiteSpace($selectedVal) -or $selectedVal -eq "") { return }
    
    $currentIdx = [array]::IndexOf($Global:ComboBoxOrder, $currentName)
    if ($currentIdx -lt 0) { return }
    
    for ($i = $currentIdx + 1; $i -lt $Global:ComboBoxOrder.Count; $i++) {
        $nextBox = $window.FindName($Global:ComboBoxOrder[$i])
        if ($nextBox) {
            $nextVal = $nextBox.SelectedItem
            if ([string]::IsNullOrWhiteSpace($nextVal) -or $nextVal -eq "") {
                $nextBox.Focus() | Out-Null
                return
            }
        }
    }
    # If all ComboBoxes are filled, focus Notes
    $window.FindName("NotesInput").Focus() | Out-Null
}

$window.FindName("rbWipeYes").Add_Click({ 
        Test-FormCompletion 
    })
$emailTextProperty = [ComponentModel.DependencyPropertyDescriptor]::FromProperty(
    [System.Windows.Controls.ComboBox]::TextProperty, [System.Windows.Controls.ComboBox])
$emailTextHandler = [EventHandler]{ Test-FormCompletion }
$emailTextProperty.AddValueChanged($techEmailInput, $emailTextHandler)
$window.Add_Closed({ $emailTextProperty.RemoveValueChanged($techEmailInput, $emailTextHandler) })
foreach ($box in $hwBoxes) { 
    $window.FindName($box).Add_SelectionChanged({ 
            Test-FormCompletion
            if (-not $Global:SkipAutoJump) { Move-ToNextEmptyComboBox }
        }) 
}
foreach ($box in $cosmeticBoxes) { 
    $window.FindName($box).Add_SelectionChanged({ 
            Test-FormCompletion
            if (-not $Global:SkipAutoJump) { Move-ToNextEmptyComboBox }
        }) 
}


# ------------------------------------------------------------
# DASHBOARD CLEANUP
# ------------------------------------------------------------
$resultsGrid.Add_AutoGeneratingColumn({
        param($evtSender, $e)
        if ($e.PropertyName -in @("Id", "RunDate", "start date", "Status")) {
            $e.Cancel = $true
        }
    })

# ------------------------------------------------------------
# DASHBOARD HORIZONTAL SCROLLING
# ------------------------------------------------------------
# WPF does not route native horizontal wheel messages as MouseWheel events.
$dashboardScrollCode = @'
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Threading;

public sealed class DashboardHorizontalScroll {
    private const int MouseWheel = 0x020A;
    private const int MouseHorizontalWheel = 0x020E;
    private const int ShiftModifier = 0x0004;

    private readonly DataGrid grid;
    // Keep the delegate alive with the window for the lifetime of the hook.
    private readonly HwndSourceHook wheelHook;
    private double? pendingOffset;

    private DashboardHorizontalScroll(DataGrid grid) {
        this.grid = grid;
        this.wheelHook = OnWindowMessage;
    }

    public static void Attach(Window window, DataGrid grid) {
        var scrolling = new DashboardHorizontalScroll(grid);
        window.SourceInitialized += scrolling.AttachHook;
    }

    private void AttachHook(object sender, EventArgs args) {
        HwndSource source = PresentationSource.FromVisual((Window)sender) as HwndSource;
        if (source != null) { source.AddHook(wheelHook); }
    }

    private IntPtr OnWindowMessage(IntPtr hwnd, int message, IntPtr wParam, IntPtr lParam, ref bool handled) {
        bool horizontal = message == MouseHorizontalWheel;
        bool shiftWheel = message == MouseWheel && (wParam.ToInt64() & ShiftModifier) != 0;
        if ((!horizontal && !shiftWheel) || !grid.IsVisible || !grid.IsEnabled) { return IntPtr.Zero; }

        // The message coordinates are signed screen pixels, including on secondary monitors.
        long coordinates = lParam.ToInt64();
        Point screenPoint = new Point(unchecked((short)coordinates), unchecked((short)(coordinates >> 16)));
        Point gridPoint = grid.PointFromScreen(screenPoint);
        if (!new Rect(grid.RenderSize).Contains(gridPoint)) { return IntPtr.Zero; }

        ScrollViewer viewer = grid.Template == null ? null : grid.Template.FindName("DG_ScrollViewer", grid) as ScrollViewer;
        int delta = unchecked((short)(wParam.ToInt64() >> 16));
        if (viewer == null || viewer.ScrollableWidth <= 0 || delta == 0) { return IntPtr.Zero; }

        // Combine high-resolution wheel messages that arrive before WPF lays out the grid.
        bool enqueue = !pendingOffset.HasValue;
        double offset = pendingOffset ?? viewer.HorizontalOffset;
        pendingOffset = Math.Max(0, Math.Min(viewer.ScrollableWidth, offset + (horizontal ? delta : -delta)));
        if (enqueue) {
            viewer.Dispatcher.BeginInvoke(DispatcherPriority.Input, new Action(delegate {
                double target = pendingOffset.GetValueOrDefault();
                pendingOffset = null;
                if (viewer.IsVisible) { viewer.ScrollToHorizontalOffset(target); }
            }));
        }
        handled = true;
        return IntPtr.Zero;
    }
}
'@
Add-Type -TypeDefinition $dashboardScrollCode -ReferencedAssemblies @(
    [System.Windows.Window].Assembly.Location
    [System.Windows.Media.Visual].Assembly.Location
    [System.Windows.Point].Assembly.Location
    [System.Xaml.XamlReader].Assembly.Location
)
[DashboardHorizontalScroll]::Attach($window, $resultsGrid)

# ------------------------------------------------------------
# DASHBOARD ROW ACTIONS
#   Select a row  -> Edit / Delete buttons appear in the toolbar
#   Double-click or Enter -> Edit    Right-click -> Edit / Delete menu    Delete key -> Delete
# ------------------------------------------------------------
function Invoke-DashButton([string]$Name) {
    $window.FindName($Name).RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
}

$resultsGrid.Add_SelectionChanged({
        $window.FindName("RowActions").Visibility = if ($null -ne $resultsGrid.SelectedItem) { "Visible" } else { "Collapsed" }
    })

# True when the mouse event happened on a data row (not the header or empty space)
function Test-OnDataRow($OriginalSource) {
    $node = $OriginalSource
    while ($null -ne $node -and -not ($node -is [System.Windows.Controls.DataGridRow])) {
        if ($node -is [System.Windows.Media.Visual] -or $node -is [System.Windows.Media.Media3D.Visual3D]) { $node = [System.Windows.Media.VisualTreeHelper]::GetParent($node) }
        else { $node = $node.Parent }
    }
    return ($null -ne $node)
}

$resultsGrid.Add_MouseDoubleClick({
        param($s, $e)
        if ($null -ne $resultsGrid.SelectedItem -and (Test-OnDataRow $e.OriginalSource)) { Invoke-DashButton "EditBtn" }
    })

$resultsGrid.Add_PreviewKeyDown({
        param($s, $e)
        if ($null -eq $resultsGrid.SelectedItem) { return }
        if ($e.Key -eq "Delete") { Invoke-DashButton "DeleteRowBtn"; $e.Handled = $true }
        elseif ($e.Key -eq "Return") { Invoke-DashButton "EditBtn"; $e.Handled = $true }
    })

$gridMenu = New-Object System.Windows.Controls.ContextMenu
Set-DqaFocusVisual $gridMenu
$menuEdit = New-Object System.Windows.Controls.MenuItem
$menuEdit.Header = "Edit record"
$menuEdit.InputGestureText = "Double-click"
$menuEdit.Add_Click({ Invoke-DashButton "EditBtn" })
$menuDelete = New-Object System.Windows.Controls.MenuItem
$menuDelete.Header = "Delete record..."
$menuDelete.InputGestureText = "Del"
$menuDelete.Add_Click({ Invoke-DashButton "DeleteRowBtn" })
$gridMenu.Items.Add($menuEdit) | Out-Null
$gridMenu.Items.Add($menuDelete) | Out-Null
$gridMenu.Add_Opened({
        $hasRow = ($null -ne $resultsGrid.SelectedItem)
        $menuEdit.IsEnabled = $hasRow
        $menuDelete.IsEnabled = $hasRow
    })
$resultsGrid.ContextMenu = $gridMenu

# ------------------------------------------------------------
# QOL: AUTO-UPPERCASE SERIAL NUMBER
# ------------------------------------------------------------
$serialInput.Add_TextChanged({
        $upperText = $serialInput.Text.ToUpper()
        if ($serialInput.Text -cne $upperText) {
            $currStart = $serialInput.SelectionStart
            $serialInput.Text = $upperText
            $serialInput.SelectionStart = $currStart
        }
        
        $val = $serialInput.Text.Trim()
        
        $isDup = $false
        if (-not [string]::IsNullOrWhiteSpace($val) -and -not [string]::IsNullOrWhiteSpace($Global:DatabaseCSV)) {
            $existing = @($Global:DatabaseCSV | ConvertFrom-Csv | ForEach-Object { ([string]$_.SerialNumber).Trim() })
            if ($val -in $existing) {
                if ($null -eq $Global:EditingRecord -or ([string]$Global:EditingRecord.SerialNumber).Trim() -ne $val) {
                    $isDup = $true
                }
            }
        }
        
        if ($isDup) {
            $serialInput.Background = $Global:redBgBrush
        }
        else {
            $serialInput.Background = $Global:whiteBgBrush
        }
        
        Test-FormCompletion
    })



# ------------------------------------------------------------
# QOL: NOTES CHARACTER COUNTER
# ------------------------------------------------------------
$notesInput.Add_TextChanged({
        $count = $notesInput.Text.Length
        # Hard limit at 500 characters
        if ($count -gt 500) {
            $caretPos = [Math]::Min($notesInput.SelectionStart, 500)
            $notesInput.Text = $notesInput.Text.Substring(0, 500)
            $notesInput.SelectionStart = $caretPos
            $count = 500
        }
        $notesCounter.Text = "($count / 500 chars)"
        if ($count -gt 400) {
            $notesCounter.Foreground = "#DC3545" # Red - approaching limit
        }
        elseif ($count -gt 250) {
            $notesCounter.Foreground = "#FD7E14" # Orange - warning
        }
        else {
            $notesCounter.Foreground = "Gray"
        }
    })

# ------------------------------------------------------------
# QOL: SEARCH PLACEHOLDER TEXT VISIBILITY & AUTO-FILTER
# ------------------------------------------------------------
$searchBox.Add_TextChanged({
        if ($searchBox.Text.Length -gt 0) {
            $searchPlaceholder.Visibility = "Hidden"
            $window.FindName("SearchClearBtn").Visibility = "Visible"
            $window.FindName("FilterBox").SelectedIndex = 0
        }
        else {
            $searchPlaceholder.Visibility = "Visible"
            $window.FindName("SearchClearBtn").Visibility = "Collapsed"
        }
        Update-Dashboard
    })
$window.FindName("SearchClearBtn").Add_Click({
        $searchBox.Text = ""
        $searchBox.Focus() | Out-Null
    })

# ------------------------------------------------------------
# BACKGROUND WORK (keeps the window responsive)
# Hardware queries (CIM/WMI) can take several seconds. They used to run on
# the UI thread and froze the window; they now run in a small background
# runspace pool and the results are applied on the UI thread.
# ------------------------------------------------------------
$Global:BgPool = [runspacefactory]::CreateRunspacePool(1, 4)
$Global:BgPool.Open()
$Global:BgJobs = New-Object System.Collections.ArrayList
$Global:BgJobTimer = New-Object System.Windows.Threading.DispatcherTimer
$Global:BgJobTimer.Interval = [TimeSpan]::FromMilliseconds(150)
$Global:BgJobTimer.Add_Tick({
        foreach ($job in @($Global:BgJobs)) {
            if (-not $job.Handle.IsCompleted) { continue }
            $Global:BgJobs.Remove($job) | Out-Null
            $result = $null
            try { $result = @($job.PS.EndInvoke($job.Handle))[0] } catch { }
            $job.PS.Dispose()
            try { & $job.OnDone $result $job.Context } catch { }
        }
        if ($Global:BgJobs.Count -eq 0) { $Global:BgJobTimer.Stop() }
    })

function Start-BackgroundWork([scriptblock]$Work, [scriptblock]$OnDone, $Context) {
    $ps = [powershell]::Create()
    try {
        $ps.RunspacePool = $Global:BgPool
        $ps.AddScript($Work.ToString()) | Out-Null
        $job = @{ PS = $ps; Handle = $ps.BeginInvoke(); OnDone = $OnDone; Context = $Context }
    }
    catch {
        $ps.Dispose()
        & $OnDone $null $Context
        return
    }
    $Global:BgJobs.Add($job) | Out-Null
    $Global:BgJobTimer.Start()
}

# Adds one short bullet line to the Notes box (never the same line twice).
# Callers write "- text"; the dash is shown as a bullet point (U+2022, built at runtime so
# this file stays plain ASCII). Lines that need attention read "CHECK: ...".
$Global:NoteBullet = [string][char]0x2022
function Add-AutoNote([string]$Line, [switch]$First) {
    if ($Line.StartsWith("- ")) { $Line = $Global:NoteBullet + " " + $Line.Substring(2) }
    $current = $notesInput.Text.TrimEnd()
    if ($current.Contains($Line)) { return }
    $notesInput.Text = if ($current.Length -eq 0) { $Line }
        elseif ($First) { $Line + "`r`n" + $current }
        else { $current + "`r`n" + $Line }
}

# Quick part: serial and charging (usually well under a second)
$Global:AutoDetectQuickWork = {
    $r = @{ Serial = "UNKNOWN"; Charging = $null }
    try { $r.Serial = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber.Trim() } catch { $r.Serial = "UNKNOWN" }
    try {
        $batteries = @(Get-CimInstance -ClassName Win32_Battery -ErrorAction Stop)
        if ($batteries.Count -eq 0) { $r.Charging = "Not Applicable" }
        else {
            $unconfirmed = @($batteries | Where-Object { $_.BatteryStatus -notin 3, 6, 7, 8, 9 })
            # A full battery can report status 2 when the charge controller is idle on AC.
            if ($unconfirmed.Count -eq 0 -or
                (@($unconfirmed | Where-Object { $_.BatteryStatus -ne 2 -or $_.EstimatedChargeRemaining -ne 100 }).Count -eq 0 -and
                    [System.Windows.Forms.SystemInformation]::PowerStatus.PowerLineStatus -eq "Online")) {
                $r.Charging = "Operational"
            }
        }
    }
    catch { }
    return $r
}

# Network: an adapter being "up" is not enough - check the laptop can really reach the internet.
# Uses the same test page Windows itself uses (Microsoft's connectivity check); nothing is sent.
# Tries 3 times over ~10 s, because the Guest Wi-Fi may still be connecting right after startup.
$Global:NetCheckWork = {
    $r = @{ AdapterUp = $null; Internet = $false; Problem = "" }
    try {
        # Same data as Get-NetAdapter (Status Up, hidden adapters excluded) without loading the slow NetAdapter module
        $adapters = Get-CimInstance -Namespace root/StandardCimv2 -ClassName MSFT_NetAdapter -Filter "InterfaceOperationalStatus=1 AND Hidden=FALSE" -ErrorAction Stop
        $r.AdapterUp = [bool]$adapters
    }
    catch { $r.AdapterUp = $null }
    if (-not $r.AdapterUp) { return $r }

    for ($i = 0; $i -lt 3 -and -not $r.Internet; $i++) {
        if ($i -gt 0) { Start-Sleep -Seconds 4 }
        try {
            $req = [System.Net.HttpWebRequest]::Create("http://www.msftconnecttest.com/connecttest.txt")
            $req.Timeout = 4000
            $req.ReadWriteTimeout = 4000
            $req.AllowAutoRedirect = $false   # a Wi-Fi sign-in page redirects instead of answering
            $resp = $req.GetResponse()
            $reader = New-Object System.IO.StreamReader($resp.GetResponseStream())
            $body = $reader.ReadToEnd()
            $resp.Close()
            if ($body.Trim() -eq "Microsoft Connect Test") { $r.Internet = $true; $r.Problem = "" }
            else { $r.Problem = "signin" }
        }
        catch { $r.Problem = "unreachable" }
    }
    return $r
}

# Deep part: monitors, USB, camera/audio presence, battery health, device errors, drive health
$Global:AutoDetectWork = {
    $r = @{ VideoOK = $false; UsbOK = $false; Notes = @(); StorageFailure = ""; BatteryHealth = $null; FullMWh = 0; DesignMWh = 0; CycleCount = 0 }

    # External monitor (PNPClass 'Monitor' is the most reliable for HDMI detection)
    try {
        $pnpMonitors = Get-CimInstance Win32_PnPEntity -Filter "PNPClass='Monitor' AND ConfigManagerErrorCode=0" -ErrorAction SilentlyContinue
        $r.VideoOK = (@($pnpMonitors).Count -gt 1)
    }
    catch { }

    $pnpDevices = @(Get-CimInstance Win32_PnPEntity -ErrorAction SilentlyContinue)

    # USB flash drive (USBSTOR) or USB mouse/keyboard (mouhid/kbdhid) with a VID/PID, or a removable disk
    try {
        $externalUsbDevices = $pnpDevices | Where-Object { ($_.Service -match 'mouhid|kbdhid|USBSTOR') -and ($_.DeviceID -match 'VID_[0-9A-F]+&PID_') }
        $usbLogical = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=2" -ErrorAction SilentlyContinue
        $r.UsbOK = [bool]($externalUsbDevices -or $usbLogical)
    }
    catch { }

    # Camera & audio hardware presence
    try {
        $cam = $pnpDevices | Where-Object { $_.PNPClass -match "Camera|Image" -and $_.ConfigManagerErrorCode -eq 0 }
        if (-not $cam) { $r.Notes += "- CHECK: Camera not detected" }
        $audio = $pnpDevices | Where-Object { $_.PNPClass -match "Media" -and $_.ConfigManagerErrorCode -eq 0 }
        if (-not $audio) { $r.Notes += "- CHECK: Sound / mic device not detected" }
    }
    catch { }

    # Battery health: full-charge capacity vs. the capacity it had when new (mWh).
    # powercfg's battery report works without admin rights; the WMI design-capacity
    # class often fails ("Generic failure") on standard accounts, so it is only a fallback.
    try {
        $reportPath = Join-Path $env:TEMP "DQA_BatteryReport.xml"
        & powercfg.exe /batteryreport /xml /output "$reportPath" 2>&1 | Out-Null
        if (Test-Path -LiteralPath $reportPath) {
            [xml]$report = Get-Content -LiteralPath $reportPath -Raw
            Remove-Item -LiteralPath $reportPath -Force -ErrorAction SilentlyContinue
            $bat = @($report.BatteryReport.Batteries.Battery)[0]
            if ($bat -and [double]$bat.DesignCapacity -gt 0 -and [double]$bat.FullChargeCapacity -gt 0) {
                $r.DesignMWh = [double]$bat.DesignCapacity
                $r.FullMWh = [double]$bat.FullChargeCapacity
                $r.CycleCount = [int]$bat.CycleCount
            }
        }
    }
    catch { }
    if ($r.DesignMWh -le 0) {
        try {
            $full = Get-CimInstance -Namespace root\wmi -ClassName BatteryFullChargedCapacity -ErrorAction SilentlyContinue | Select-Object -First 1
            $design = Get-CimInstance -Namespace root\wmi -ClassName BatteryStaticData -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($full -and $design -and $design.DesignedCapacity -gt 0 -and $full.FullChargedCapacity -gt 0) {
                $r.FullMWh = [double]$full.FullChargedCapacity
                $r.DesignMWh = [double]$design.DesignedCapacity
            }
        }
        catch { }
    }
    if ($r.DesignMWh -gt 0 -and $r.FullMWh -gt 0) { $r.BatteryHealth = ($r.FullMWh / $r.DesignMWh) * 100 }

    # Device Manager errors (USB & general)
    try {
        $errorDevices = $pnpDevices | Where-Object { $_.ConfigManagerErrorCode -gt 0 }
        if ($errorDevices | Where-Object { $_.PNPClass -eq "USB" }) {
            $r.Notes += "- CHECK: USB port error in Device Manager"
        }
        $errNames = (($errorDevices | Where-Object { $_.PNPClass -ne "USB" }) | Select-Object -ExpandProperty Name) -join ", "
        if ($errNames) { if ($errNames.Length -gt 90) { $errNames = $errNames.Substring(0, 87) + "..." }; $r.Notes += "- CHECK: Device errors: $errNames" }
    }
    catch { }

    # Drive health (SMART)
    try {
        $failingDrives = Get-CimInstance Win32_DiskDrive -ErrorAction SilentlyContinue | Where-Object { $_.Status -match "Fail|Error" -or $_.Status -eq "Pred Fail" }
        if ($failingDrives) { $r.StorageFailure = "- CHECK: Drive failing (SMART) - replace drive" }
    }
    catch { }

    # Storage health: what each internal drive reports about itself (USB drives are skipped)
    try {
        $healthNames = @{ 0 = "Healthy"; 1 = "Warning"; 2 = "Unhealthy"; 5 = "Unknown" }
        $busNames = @{ 17 = "NVMe"; 11 = "SATA"; 8 = "RAID"; 10 = "SAS"; 12 = "SD card"; 13 = "MMC" }
        $mediaNames = @{ 3 = "HDD"; 4 = "SSD"; 5 = "SCM" }
        $disks = @(Get-CimInstance -Namespace root\Microsoft\Windows\Storage -ClassName MSFT_PhysicalDisk -ErrorAction Stop | Where-Object { $_.BusType -ne 7 })
        $sysVol = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$($env:SystemDrive)'" -ErrorAction SilentlyContinue
        $freeText = ""
        if ($sysVol -and $sysVol.Size -gt 0) {
            $freeGB = [math]::Round($sysVol.FreeSpace / 1GB)
            $freeText = ", $freeGB GB free"
            if (($sysVol.FreeSpace / $sysVol.Size) -lt 0.10) { $r.Notes += "- CHECK: Low disk space ($freeGB GB free)" }
        }
        foreach ($d in $disks) {
            $health = if ($healthNames.ContainsKey([int]$d.HealthStatus)) { $healthNames[[int]$d.HealthStatus] } else { "Unknown" }
            $bus = if ($busNames.ContainsKey([int]$d.BusType)) { " " + $busNames[[int]$d.BusType] } else { "" }
            $media = if ($mediaNames.ContainsKey([int]$d.MediaType)) { $mediaNames[[int]$d.MediaType] } else { "drive" }
            $line = "- $media $([math]::Round($d.Size / 1GB)) GB${bus}: $health$freeText"
            if ($health -eq "Warning" -or $health -eq "Unhealthy") { $line = "- CHECK: $media $([math]::Round($d.Size / 1GB)) GB${bus}: $health - replace drive" }
            if ($health -eq "Unhealthy" -and -not $r.StorageFailure) {
                $r.StorageFailure = "- CHECK: Storage drive unhealthy - replace drive"
            }
            $r.Notes += $line
            $freeText = ""   # only mention free space once
        }
    }
    catch { }

    # Windows version and activation (filtered by the Windows application ID: ~1 s instead of ~35 s)
    try {
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $displayVersion = ""
        try { $displayVersion = " " + (Get-ItemPropertyValue "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name DisplayVersion -ErrorAction Stop) } catch { }
        $winText = ($os.Caption -replace '^Microsoft ', '' -replace '^Windows ', 'Win ') + $displayVersion
        $licNames = @{ 0 = "NOT activated"; 1 = "Activated"; 2 = "NOT activated"; 3 = "NOT activated"; 4 = "NOT activated (non-genuine)"; 5 = "NOT activated"; 6 = "NOT activated" }
        $actText = "activation unknown"
        try {
            $lic = Get-CimInstance SoftwareLicensingProduct -Filter "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL" -ErrorAction Stop | Select-Object -First 1
            if ($lic -and $licNames.ContainsKey([int]$lic.LicenseStatus)) { $actText = $licNames[[int]$lic.LicenseStatus] }
        }
        catch { }
        $line = "- ${winText}: $actText"
        if ($actText -like "NOT activated*") { $line = "- CHECK: ${winText}: $actText" }
        $r.Notes += $line
        if ($os.Caption -match 'Windows 10') { $r.Notes += "- CHECK: Windows 10 is unsupported - upgrade to Win 11" }
    }
    catch { }

    # Security: TPM chip, Secure Boot, BitLocker on the system drive
    try {
        $tpm = Get-CimInstance Win32_PnPEntity -Filter "PNPClass='SecurityDevices'" -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'Trusted Platform' } | Select-Object -First 1
        if ($tpm) {
            $tpmVer = if ($tpm.Name -match '(\d\.\d)') { $Matches[1] } else { "" }
            $tpmText = "TPM $tpmVer " + $(if ($tpm.ConfigManagerErrorCode -eq 0) { "OK" } else { "error" })
        }
        else { $tpmText = "No TPM" }

        $sb = $null
        try { $sb = Get-ItemPropertyValue "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State" -Name UEFISecureBootEnabled -ErrorAction Stop } catch { }
        $sbText = if ($sb -eq 1) { "Secure Boot ON" } elseif ($sb -eq 0) { "Secure Boot OFF" } else { "Legacy BIOS (no Secure Boot)" }

        $blText = "BitLocker unknown"
        try {
            $blCode = (New-Object -ComObject Shell.Application).NameSpace("$($env:SystemDrive)\").Self.ExtendedProperty('System.Volume.BitLockerProtection')
            $blText = switch ([int]$blCode) { 1 { "BitLocker ON" } 6 { "BitLocker ON" } 2 { "BitLocker OFF" } 3 { "BitLocker encrypting" } 4 { "BitLocker decrypting" } 5 { "BitLocker suspended" } default { "BitLocker unknown" } }
        }
        catch { }

        $line = "- $sbText | $blText"
        if ($sb -ne 1) { $line = "- CHECK: $sbText | $blText" }
        $r.Notes += $line
        # TPM is only mentioned when there is a problem (Windows 11 needs TPM 2.0)
        if ($tpmText.Trim() -ne "TPM 2.0 OK") { $r.Notes += "- CHECK: $($tpmText.Trim()) - Win 11 needs TPM 2.0" }
    }
    catch { }

    # Wireless & sensors: Wi-Fi signal (antenna check), Bluetooth radio, fingerprint / face sensor
    try {
        $parts = @()
        $problem = $false

        $wlan = @(netsh wlan show interfaces 2>$null)
        if (($wlan -join " ") -match 'There is no wireless interface|Wireless AutoConfig Service .* is not running') {
            $parts += "No Wi-Fi adapter"; $problem = $true
        }
        else {
            $sigLine = $wlan | Where-Object { $_ -match '^\s*Signal\s*:' } | Select-Object -First 1
            if ($sigLine -and ($sigLine -split ':', 2)[1].Trim() -match '(\d+)') {
                $sig = [int]$Matches[1]
                if ($sig -lt 30) { $parts += "Wi-Fi weak $sig% (check antenna)"; $problem = $true } else { $parts += "Wi-Fi $sig%" }
            }
            else { $parts += "Wi-Fi not connected" }
        }

        # The Bluetooth radio itself is a USB or PCI device in the Bluetooth class
        $bt = @(Get-CimInstance Win32_PnPEntity -Filter "PNPClass='Bluetooth'" -ErrorAction Stop | Where-Object { $_.DeviceID -match '^(USB|PCI)\\' })
        if ($bt.Count -eq 0) { $parts += "No Bluetooth"; $problem = $true }
        elseif (@($bt | Where-Object { $_.ConfigManagerErrorCode -ne 0 }).Count -gt 0) { $parts += "Bluetooth error"; $problem = $true }
        else { $parts += "Bluetooth OK" }

        $bio = @(Get-CimInstance Win32_PnPEntity -Filter "PNPClass='Biometric'" -ErrorAction Stop)
        if ($bio.Count -eq 0) { $parts += "No fingerprint/face sensor" }
        elseif (@($bio | Where-Object { $_.ConfigManagerErrorCode -ne 0 }).Count -gt 0) { $parts += "Fingerprint/face error"; $problem = $true }
        else { $parts += "Fingerprint/face OK" }

        $r.Notes += $(if ($problem) { "- CHECK: " } else { "- " }) + ($parts -join " | ")
    }
    catch { }

    return $r
}

# Writes the battery health into the Additional Notes (with a warning when it is below the Settings limit)
function Show-BatteryHealth($r) {
    if ($null -eq $r -or $null -eq $r.BatteryHealth) { return }   # no battery data (e.g. desktop): add nothing
    $pct = [int][math]::Round($r.BatteryHealth)
    $warn = 40   # fixed limit: below 40% of original capacity gets a "- CHECK:" line (not a Settings option)
    $cycles = if ([int]$r.CycleCount -gt 0) { ", $($r.CycleCount) cycles" } else { "" }
    # Only noted when the battery is below the limit; a healthy battery adds nothing
    if ($pct -lt $warn) { Add-AutoNote "- CHECK: Battery health $pct%$cycles - replace battery" }
}

# ------------------------------------------------------------
# 1. AUTO-DETECT (Triggers only after Email is provided)
# ------------------------------------------------------------
function Reset-InspectionDetection {
    if ($Global:audioState -gt 0) {
        $Global:AutoTestInProgress = $false
        Stop-AudioTest
    }
    $script:InspectionGeneration++
    $script:AutoDetectGeneration = $null
    $script:AutoDetectPending = 0
    $script:InspectionStorageFailure = ""
    $window.FindName("DetectStatus").Visibility = "Collapsed"
}

function Complete-AutoDetectPart {
    $script:AutoDetectPending--
    $detectStatus = $window.FindName("DetectStatus")
    if ($script:AutoDetectPending -eq 0) { $detectStatus.Visibility = "Collapsed" }
    else { $detectStatus.Text = "Still checking internet, ports, battery and device health..." }
    Test-FormCompletion
}

function Start-AutoDetect {
    if ([string]::IsNullOrWhiteSpace($techEmailInput.Text) -or $null -ne $Global:EditingRecord) { return }
    if ($script:AutoDetectGeneration -eq $script:InspectionGeneration) { return }

    $generation = $script:InspectionGeneration
    $script:AutoDetectGeneration = $generation
    $script:AutoDetectPending = 3
    $detectStatus = $window.FindName("DetectStatus")
    $detectStatus.Text = "Detecting hardware... (serial, charging, internet)"
    $detectStatus.Visibility = "Visible"
    Test-FormCompletion

    Start-BackgroundWork -Work $Global:AutoDetectQuickWork -Context $generation -OnDone {
        param($result, $generation)
        if ($generation -ne $script:InspectionGeneration) { return }
        try {
            if ($null -eq $result) { return }
            # Keep a serial or charging result the technician has already entered.
            if ([string]::IsNullOrWhiteSpace($serialInput.Text) -or $serialInput.Text -eq "UNKNOWN") {
                $serialInput.Text = $result.Serial
            }
            $chargingBox = $window.FindName("cbCharging")
            if ([string]::IsNullOrWhiteSpace($chargingBox.SelectedItem)) {
                if ($result.Charging) { $chargingBox.SelectedItem = $result.Charging }
                else { Add-AutoNote "- CHECK: Confirm charging with the AC adapter and select a result" }
            }
        }
        finally { Complete-AutoDetectPart }
    }

    Start-BackgroundWork -Work $Global:NetCheckWork -Context $generation -OnDone {
        param($result, $generation)
        if ($generation -ne $script:InspectionGeneration) { return }
        try {
            if ($null -eq $result) { return }
            $networkBox = $window.FindName("cbNetwork")
            if (-not [string]::IsNullOrWhiteSpace($networkBox.SelectedItem)) { return }

            if ($null -eq $result.AdapterUp) { Add-AutoNote "- CHECK: Network adapter status unavailable; check manually" }
            elseif (-not $result.AdapterUp) {
                $networkBox.SelectedItem = "Defective"
                Add-AutoNote "- CHECK: No network adapter connected"
            }
            elseif ($result.Internet) { $networkBox.SelectedItem = "Operational" }
            else {
                $why = if ($result.Problem -eq "signin") { "network wants a sign-in page" } else { "no internet" }
                Add-AutoNote "- CHECK: Network connected but $why - check by hand"
            }
        }
        finally { Complete-AutoDetectPart }
    }

    Start-BackgroundWork -Work $Global:AutoDetectWork -Context $generation -OnDone {
        param($result, $generation)
        if ($generation -ne $script:InspectionGeneration) { return }
        try {
            if ($null -eq $result) { return }
            Set-DetectedPortResults $result
            $script:InspectionStorageFailure = [string]$result.StorageFailure
            if ($script:InspectionStorageFailure) { Add-AutoNote $script:InspectionStorageFailure -First }
            foreach ($line in @($result.Notes)) { Add-AutoNote $line }
            Show-BatteryHealth $result
        }
        finally { Complete-AutoDetectPart }
    }
}

function Set-DetectedPortResults($Result) {
    $videoBox = $window.FindName("cbVideo")
    $usbBox = $window.FindName("cbUSB")
    if ($Result.VideoOK -and [string]::IsNullOrWhiteSpace($videoBox.SelectedItem)) {
        $videoBox.SelectedItem = "Operational"
    }
    if ($Result.UsbOK -and [string]::IsNullOrWhiteSpace($usbBox.SelectedItem)) {
        $usbBox.SelectedItem = "Operational"
    }
}

$techEmailInput.Add_LostFocus({ Start-AutoDetect })
$techEmailInput.Add_SelectionChanged({ Start-AutoDetect })
$acceptTechnicianEmail = {
    $serialInput.Focus() | Out-Null
    Start-AutoDetect
}
# WPF can handle Enter itself when accepting a saved email. Advance in that case too.
$techEmailInput.AddHandler(
    [System.Windows.Input.Keyboard]::PreviewKeyDownEvent,
    [System.Windows.Input.KeyEventHandler] {
        param($sender, $e)
        if ($e.Key -ne [System.Windows.Input.Key]::Return) { return }

        if ($sender.IsDropDownOpen) {
            # Let the combo box accept the highlighted email before moving focus.
            $sender.Dispatcher.BeginInvoke(
                [Action]$acceptTechnicianEmail, [System.Windows.Threading.DispatcherPriority]::Input
            ) | Out-Null
        }
        else {
            $e.Handled = $true
            & $acceptTechnicianEmail
        }
    },
    $true
)

# ------------------------------------------------------------
# 2. CAMERA & KEYBOARD TESTS (Interactive Prompts)
# ------------------------------------------------------------
$Global:AutoTestInProgress = $false

# Camera test: opens the Windows Camera app, asks at the bottom of the screen
# (so the live preview stays visible), then closes the Camera app again.
function Invoke-CameraTest {
    $cameraWasOpen = [bool](Get-Process -Name "WindowsCamera" -ErrorAction SilentlyContinue)
    try { Start-Process "microsoft.windows.camera:" } catch { }
    Start-Sleep -Seconds 1
    $res = Show-DsuChoice -Text "Is the camera preview showing a clear, live picture?" -Title "Camera Test" -YesText "Camera works" -NoText "Camera broken" -Position "Bottom"
    if ($res -eq "Yes") { $window.FindName("cbCamera").SelectedItem = "Operational" }
    elseif ($res -eq "No") { $window.FindName("cbCamera").SelectedItem = "Defective" }
    # Close the Camera app only if this test opened it
    if (-not $cameraWasOpen) { Get-Process -Name "WindowsCamera" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue }
    $window.Activate() | Out-Null
}

$window.FindName("BtnAutoTest").Add_Click({
        if ($Global:audioState -ne 0) { return }   # an audio test is already running
        # 1. Camera test
        Invoke-CameraTest

        # 2. Start Audio Test & flag for sequential keyboard -> touchscreen tests
        $Global:AutoTestInProgress = $true
        Start-AudioTest
    })
$window.FindName("BtnTestCamera").Add_Click({ Invoke-CameraTest })

$window.FindName("BtnTestKeys").Add_Click({ 
        $res = Show-KeyboardTester
        if ($res -eq "Operational") { $window.FindName("cbKeyboard").SelectedItem = "Operational" }
        elseif ($res -eq "Defective") { $window.FindName("cbKeyboard").SelectedItem = "Defective" }
    })

# Touchscreen test: hardware check first, then the tap test (repeats on Redo / Retry)
function Invoke-TouchTest {
    try {
        $touchDevice = Get-CimInstance -ClassName Win32_PnPEntity -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "touch screen|touch digitizer" }
        if (-not $touchDevice) {
            [DsuMessageBox]::Show("No touchscreen hardware detected in Device Manager.", "Hardware Not Found", "OK", "Warning") | Out-Null
            $window.FindName("cbTouch").SelectedItem = "Defective"
            return
        }
    }
    catch { }

    do {
        $res = Show-TouchscreenTester
        if ($res -eq "Operational") { $window.FindName("cbTouch").SelectedItem = "Operational" }
        elseif ($res -eq "Defective") { $window.FindName("cbTouch").SelectedItem = "Defective" }
    } while ($res -eq "Retry")
}

$window.FindName("BtnTestTouch").Add_Click({ Invoke-TouchTest })

# ------------------------------------------------------------
# 3. ASYNC AUDIO & MIC TEST (soft chime + voice recording)
#    The mic records for 6 s. A soft chime plays from the speakers right
#    after recording starts (speaker test); then the technician speaks or
#    claps. The recording is cleaned up (hum removed, quiet audio boosted,
#    peaks soft-limited) and played back: you hear the chime + your voice.
#    States: 0 idle | 1 recording | 2 playback
# ------------------------------------------------------------
$Global:audioTimer = New-Object System.Windows.Threading.DispatcherTimer
$Global:audioState = 0
$Global:tempAudioPath = $null
$Global:soundPlayer = New-Object System.Media.SoundPlayer
# 6 s total: the 4 s chime + about 2 s for the technician's voice or clap
$Global:RecordSeconds = 6
$Global:recordClock = $null

# Generate each sound on first use and keep it in memory for previews and tests.
$script:TestSoundWavs = @{}

function Start-Chime([string]$Sound = (Get-Setting "AudioSound")) {
    try {
        if (-not $script:TestSoundWavs.ContainsKey($Sound)) {
            $script:TestSoundWavs[$Sound] = [AudioHelper]::BuildChimeWav($Sound)
        }
        $chimeStream = [System.IO.MemoryStream]::new($script:TestSoundWavs[$Sound], $false)
        $Global:chimePlayer = New-Object System.Media.SoundPlayer -ArgumentList $chimeStream
        $Global:chimePlayer.Load()
        $Global:chimePlayer.Play()
        return $true
    }
    catch { return $false }
}

# Live mic level bar (polls the recorder while it records)
$Global:micMeterTimer = New-Object System.Windows.Threading.DispatcherTimer
$Global:micMeterTimer.Interval = [TimeSpan]::FromMilliseconds(80)
$Global:micMeterTimer.Add_Tick({
        # Status cue: listen while the chime plays, then a "speak now" countdown
        if ($null -ne $Global:recordClock) {
            $elapsed = $Global:recordClock.Elapsed.TotalSeconds
            $chimeEnd = [AudioHelper]::ChimeSeconds + 0.15
            if (-not $Global:SpeakerCueStarted -and $elapsed -lt $chimeEnd) {
                $statusText = " Speaker test sound unavailable. "
                $statusColor = "#856404"
            }
            elseif ($elapsed -lt ([AudioHelper]::LeftUntil + 0.15)) {
                $statusText = " Sound on the LEFT speaker... "
                $statusColor = "#003552"
            }
            elseif ($elapsed -lt $chimeEnd) {
                $statusText = " Sound on the RIGHT speaker... "
                $statusColor = "#00779E"
            }
            else {
                $left = [Math]::Max(0, [Math]::Ceiling($Global:RecordSeconds - $elapsed))
                $statusText = if ($Global:MicCaptureError) { " Microphone recording unavailable. " }
                    else { " Speak or clap NOW - say ""Mic test 1, 2, 3""  ($($left)s) " }
                $statusColor = "#DC3545"
            }
            if ($audioStatusLabel.Text -ne $statusText) {
                $audioStatusLabel.Text = $statusText
                $audioStatusLabel.Background = $statusColor
            }
        }

        if ($Global:MicCaptureError) { return }
        $raw = [AudioHelper]::MciQuery("status recsound level")
        $value = 0
        if (-not [int]::TryParse($raw, [ref]$value)) {
            # This audio driver doesn't report a live level: hide the bar rather than show a dead meter
            $Global:micMeterMisses++
            if ($Global:micMeterMisses -ge 3) { $window.FindName("MicMeter").Visibility = "Collapsed" }
            return
        }
        $Global:micMeterMisses = 0
        # 16-bit level 0..32767; scaled up so normal speech fills most of the bar
        $fraction = [Math]::Min(1.0, [Math]::Sqrt([Math]::Abs($value) / 32768.0) * 1.6)
        $track = $window.FindName("MicMeterTrack")
        $window.FindName("MicMeterFill").Width = [Math]::Max(0, $track.ActualWidth * $fraction)
    })

# Plain-language mic level from the recording's peak (0.0 - 1.0)
function Get-MicLevelText([double]$Level) {
    if ($Level -lt 0) { return "unknown" }
    if ($Level -lt 0.02) { return "almost nothing recorded" }
    if ($Level -lt 0.10) { return "low (boosted for playback)" }
    return "good"
}

function Set-TestButtonsEnabled([bool]$Enabled) {
    $window.FindName("BtnTestAudio").IsEnabled = $Enabled
    $window.FindName("BtnAutoTest").IsEnabled = $Enabled
}

function Set-AudioStatus([string]$Text, [string]$Color) {
    $audioStatusLabel.Text = " $Text "
    $audioStatusLabel.Background = $Color
    $audioStatusLabel.Foreground = "White"
    $audioStatusLabel.Visibility = "Visible"
}

function Start-AudioTest {
    if ($Global:audioState -ne 0) { return }
    $Global:audioState = 1
    $Global:tempAudioPath = Join-Path ([IO.Path]::GetTempPath()) "DQA_MicTest_$([guid]::NewGuid().ToString('N')).wav"
    $Global:MicCaptureError = ""
    $Global:MicLevelText = "unknown"

    Set-TestButtonsEnabled $false
    $window.FindName("cbMic").SelectedIndex = 0
    $window.FindName("cbAudio").SelectedIndex = 0
    Test-FormCompletion

    [AudioHelper]::SetVolume([int](Get-Setting "AudioVolume"))
    Set-AudioStatus "Sound on the LEFT speaker..." "#003552"
    $window.FindName("MicMeterFill").Width = 0
    $Global:micMeterMisses = 0

    [AudioHelper]::mciSendString("close recsound", [IntPtr]::Zero, 0, [IntPtr]::Zero) | Out-Null
    $opened = [AudioHelper]::mciSendString("open new type waveaudio alias recsound", [IntPtr]::Zero, 0, [IntPtr]::Zero)
    if ($opened -eq 0) {
        $fmt = [AudioHelper]::mciSendString("set recsound time format ms bitspersample 16 channels 1 samplespersec 44100 bytespersec 88200 alignment 2", [IntPtr]::Zero, 0, [IntPtr]::Zero)
        if ($fmt -ne 0) {
            foreach ($cmd in @("set recsound bitspersample 16", "set recsound samplespersec 44100", "set recsound channels 1", "set recsound bytespersec 88200", "set recsound alignment 2")) {
                [AudioHelper]::mciSendString($cmd, [IntPtr]::Zero, 0, [IntPtr]::Zero) | Out-Null
            }
        }
        $recorded = [AudioHelper]::mciSendString("record recsound", [IntPtr]::Zero, 0, [IntPtr]::Zero)
        if ($recorded -ne 0) {
            $Global:MicCaptureError = "Windows could not start recording from the selected input."
            [AudioHelper]::mciSendString("close recsound", [IntPtr]::Zero, 0, [IntPtr]::Zero) | Out-Null
        }
    }
    else { $Global:MicCaptureError = "Windows could not open the selected microphone." }
    $window.FindName("MicMeter").Visibility = if ($Global:MicCaptureError) { "Collapsed" } else { "Visible" }

    # Give the recorder a moment to start so the chime is captured too.
    Start-Sleep -Milliseconds 150
    $Global:SpeakerCueStarted = Start-Chime
    if (-not $Global:SpeakerCueStarted) { Set-AudioStatus "Speaker test sound unavailable." "#856404" }
    $Global:recordClock = [Diagnostics.Stopwatch]::StartNew()
    $Global:audioTimer.Interval = [TimeSpan]::FromSeconds($Global:RecordSeconds)
    $Global:audioTimer.Start()
    $Global:micMeterTimer.Start()
}

# Re-enables the buttons and continues Auto-Test (Keyboard -> Touchscreen)
function Complete-AudioTest {
    Stop-AudioTest
    if ($Global:AutoTestInProgress) {
        $Global:AutoTestInProgress = $false
        $resKey = Show-KeyboardTester
        if ($resKey -eq "Operational") { $window.FindName("cbKeyboard").SelectedItem = "Operational" }
        elseif ($resKey -eq "Defective") { $window.FindName("cbKeyboard").SelectedItem = "Defective" }
        Invoke-TouchTest
    }
}

# Also used when the inspection changes or the window closes mid-test.
function Stop-AudioTest {
    $Global:audioTimer.Stop()
    $Global:micMeterTimer.Stop()
    $Global:recordClock = $null
    try { $Global:soundPlayer.Stop() } catch { }
    try { if ($Global:chimePlayer) { $Global:chimePlayer.Stop() } } catch { }
    [AudioHelper]::mciSendString("close recsound", [IntPtr]::Zero, 0, [IntPtr]::Zero) | Out-Null
    if ($Global:tempAudioPath) {
        Remove-Item -LiteralPath $Global:tempAudioPath -Force -ErrorAction SilentlyContinue
        $Global:tempAudioPath = $null
    }
    $Global:audioState = 0
    $audioStatusLabel.Visibility = "Collapsed"
    $window.FindName("MicMeter").Visibility = "Collapsed"
    Set-TestButtonsEnabled $true
    Test-FormCompletion
}
$window.Add_Closing({ Stop-AudioTest })

$Global:audioTimer.Add_Tick({
        $Global:audioTimer.Stop()

        if ($Global:audioState -eq 1) {
            $Global:micMeterTimer.Stop()
            $Global:recordClock = $null
            $window.FindName("MicMeter").Visibility = "Collapsed"

            $playMs = 1
            if (-not $Global:MicCaptureError) {
                $saveCmd = 'save recsound "' + $Global:tempAudioPath + '"'
                $stopped = [AudioHelper]::mciSendString("stop recsound", [IntPtr]::Zero, 0, [IntPtr]::Zero)
                $saved = [AudioHelper]::mciSendString($saveCmd, [IntPtr]::Zero, 0, [IntPtr]::Zero)
                [AudioHelper]::mciSendString("close recsound", [IntPtr]::Zero, 0, [IntPtr]::Zero) | Out-Null
                if ($stopped -ne 0 -or $saved -ne 0 -or -not [IO.File]::Exists($Global:tempAudioPath)) {
                    $Global:MicCaptureError = "Windows could not save a new microphone recording."
                }
                else {
                    try {
                        $level = [AudioHelper]::NormalizeWav($Global:tempAudioPath, 8.0)
                        $Global:soundPlayer = [System.Media.SoundPlayer]::new($Global:tempAudioPath)
                        $Global:soundPlayer.Load()
                        $Global:soundPlayer.Play()
                        $Global:MicLevelText = Get-MicLevelText $level
                        Set-AudioStatus "Playing back the recording...  Mic level: $($Global:MicLevelText)" "#003552"
                        $playMs = ($Global:RecordSeconds * 1000) + 400
                    }
                    catch { $Global:MicCaptureError = "The microphone recording could not be played back." }
                }
            }

            $Global:audioState = 2
            $Global:audioTimer.Interval = [TimeSpan]::FromMilliseconds($playMs)
            $Global:audioTimer.Start()
        }
        elseif ($Global:audioState -eq 2) {
            $audioStatusLabel.Visibility = "Collapsed"
            $window.Activate() | Out-Null

            if ($Global:SpeakerCueStarted) {
                # The test sound plays on the LEFT speaker first, then on the RIGHT.
                $speakerBtns = @(
                    @{ Text = "Both sides"; Result = "Both"; Style = "Success"; IsDefault = $true },
                    @{ Text = "Left only"; Result = "Left"; Style = "DangerOutline" },
                    @{ Text = "Right only"; Result = "Right"; Style = "DangerOutline" },
                    @{ Text = "Neither"; Result = "None"; Style = "DangerOutline" }
                )
                $resAudio = Show-DsuDialog -Title "Speaker Test" -Kind "Question" -Buttons $speakerBtns -DefaultResult "Cancel" -Width 480 `
                    -Message "Which speakers played the test sound?`n`nThe sound plays on the LEFT first, then on the RIGHT."
                switch ($resAudio) {
                    "Both" { $window.FindName("cbAudio").SelectedItem = "Operational" }
                    "Left" { $window.FindName("cbAudio").SelectedItem = "Defective"; Add-AutoNote "- CHECK: Right speaker silent" }
                    "Right" { $window.FindName("cbAudio").SelectedItem = "Defective"; Add-AutoNote "- CHECK: Left speaker silent" }
                    "None" { $window.FindName("cbAudio").SelectedItem = "Defective"; Add-AutoNote "- CHECK: No sound from either speaker" }
                }
            }
            else {
                [DsuMessageBox]::Show("The stereo test sound could not play. Retry or check both speakers manually.", "Speaker Test", "OK", "Warning") | Out-Null
            }

            if ($Global:MicCaptureError) {
                [DsuMessageBox]::Show("$($Global:MicCaptureError) Retry or check the microphone manually.", "Microphone Test", "OK", "Warning") | Out-Null
            }
            else {
                $resMic = Show-DsuChoice -Text "Did you hear your voice or clap in the playback?`n`nMic level: $($Global:MicLevelText)" -Title "Microphone Test" -YesText "Yes, mic works" -NoText "No"
                if ($resMic -eq "Yes") { $window.FindName("cbMic").SelectedItem = "Operational" }
                elseif ($resMic -eq "No") { $window.FindName("cbMic").SelectedItem = "Defective" }
            }

            Complete-AudioTest
        }
    })

$window.FindName("BtnTestAudio").Add_Click({ Start-AudioTest })

# ------------------------------------------------------------
# AUTO-FILL REMAINING TO OPERATIONAL (Excludes Auto-Detected & Interactive Tests)
# ------------------------------------------------------------
$window.FindName("BtnAutoFill").Add_Click({
        # Excludes cbCharging and cbNetwork (which auto-detect) 
        # Excludes Cam, Audio, Mic, Keyboard, Touchscreen (which have interactive prompts)
        $hwBoxesToFill = @("cbScreen", "cbMouse", "cbVideo", "cbUSB")
        foreach ($box in $hwBoxesToFill) {
            if ([string]::IsNullOrWhiteSpace($window.FindName($box).Text)) { 
                $window.FindName($box).SelectedItem = "Operational" 
            }
        }
    
        $cosmeticBoxesToFill = @("cbPalm", "cbBackplate", "cbBase", "cbHinge")
        foreach ($box in $cosmeticBoxesToFill) {
            if ([string]::IsNullOrWhiteSpace($window.FindName($box).Text)) { 
                $window.FindName($box).SelectedItem = "Acceptable" 
            }
        }
        Test-FormCompletion
    })

# ------------------------------------------------------------
# 4. SAVE & DASHBOARD LOGIC (Modifies Script Content)
# ------------------------------------------------------------
function Set-InspectionDurationMode {
    $editing = $null -ne $Global:EditingRecord
    $additionalTimeInput.Clear()
    $durationCorrectionInput.Clear()
    $window.FindName("AdditionalTimePanel").Visibility = if ($editing) { "Collapsed" } else { "Visible" }
    $window.FindName("InspectionTimerPanel").Visibility = if ($editing) { "Collapsed" } else { "Visible" }
    $window.FindName("DurationCorrectionPanel").Visibility = if ($editing) { "Visible" } else { "Collapsed" }

    if ($editing) {
        $Global:sessionStopwatch.Stop()
        $Global:isTimerPaused = $true
        $Global:uiTimer.Stop()
        $savedDuration = $Global:EditingRecord.DurationMins
        $window.FindName("SavedDurationHint").Text = if ([string]::IsNullOrWhiteSpace($savedDuration)) {
            "No duration was recorded. Leave blank to keep it empty."
        }
        else {
            "Saved duration is {0} min. Leave blank to keep it." -f $savedDuration
        }
    }
}

$window.FindName("SaveInspectionBtn").Add_Click({
        if ($script:AutoDetectPending -gt 0 -or $Global:audioState -gt 0) { return }
        $recordToReplace = $Global:EditingRecord
        if ($null -ne $recordToReplace -and $script:EditingDatabaseCsv -cne $Global:DatabaseCSV) {
            [DsuMessageBox]::Show("Inspection history changed while this record was open. Reopen it from the dashboard to edit it.", "Inspection Changed", "OK", "Warning") | Out-Null
            return
        }
        $keepSavedDuration = $false
        if ($null -ne $recordToReplace) {
            $keepSavedDuration = [string]::IsNullOrWhiteSpace($durationCorrectionInput.Text)
            $finalDuration = $recordToReplace.DurationMins
            if (-not $keepSavedDuration) {
                $correctedDuration = 0.0
                $validDuration = [double]::TryParse($durationCorrectionInput.Text, [System.Globalization.NumberStyles]::Float,
                    [System.Globalization.CultureInfo]::CurrentCulture, [ref]$correctedDuration)
                if (-not $validDuration -or [double]::IsNaN($correctedDuration) -or
                    [double]::IsInfinity($correctedDuration) -or $correctedDuration -lt 0.1) {
                    $message = "Enter at least {0} minutes, or clear the correction field to keep the saved duration." -f 0.1
                    [DsuMessageBox]::Show($message, "Invalid duration", "OK", "Warning") | Out-Null
                    $durationCorrectionInput.Focus() | Out-Null
                    return
                }
                $finalDuration = [math]::Round($correctedDuration, 1).ToString([System.Globalization.CultureInfo]::InvariantCulture)
            }
        }
        else {
            $calcDurationMins = $Global:sessionStopwatch.Elapsed.TotalMinutes
            $addTime = 0.0
            if (-not [string]::IsNullOrWhiteSpace($additionalTimeInput.Text) -and [double]::TryParse($additionalTimeInput.Text, [ref]$addTime)) {
                $calcDurationMins += $addTime
            }

            $finalDuration = [math]::Max(0.1, [math]::Round($calcDurationMins, 1))
        }

        $rawEmail = $techEmailInput.Text.Trim()
        $rawEmail = Format-TitleCaseEmail $rawEmail

        if ([string]::IsNullOrWhiteSpace($rawEmail) -or [string]::IsNullOrWhiteSpace($serialInput.Text)) {
            [DsuMessageBox]::Show("Technician Email and Serial Number are required.", "Error", "OK", "Warning")
            return
        }

        $savedSerial = $serialInput.Text.Trim()
        $allData = [System.Collections.ArrayList]@($Global:DatabaseCSV | ConvertFrom-Csv)
        $existingRecord = $allData | Where-Object {
            ([string]$_.SerialNumber).Trim() -eq $savedSerial
        } | Select-Object -First 1
        $editingSameSerial = $null -ne $recordToReplace -and ([string]$recordToReplace.SerialNumber).Trim() -eq $savedSerial

        if ($existingRecord -and -not $editingSameSerial) {
            if ($null -ne $recordToReplace) {
                [DsuMessageBox]::Show("Another inspection already uses this serial number. Enter a different serial or edit that record from the dashboard.", "Serial Already Saved", "OK", "Warning") | Out-Null
                return
            }
            $dupBtns = @(
                @{ Text = "Replace old record"; Result = "Yes"; Style = "Danger" },
                @{ Text = "Cancel"; Result = "No"; Style = "Secondary"; IsDefault = $true; IsCancel = $true }
            )
            $dupConfirm = Show-DsuDialog -Title "Serial Already Saved" -Kind "Warning" -Buttons $dupBtns -DefaultResult "No" -Width 460 `
                -Message "Serial number '$savedSerial' already has a saved inspection.`n`nResult: $($existingRecord.FinalStatus)`nTechnician: $($existingRecord.TechnicianEmail)`nDuration: $($existingRecord.DurationMins) min`n`nReplace it with this new inspection?"
            if ($dupConfirm -ne "Yes") { return }
            $recordToReplace = $existingRecord
            $keepSavedDuration = $false
        }

        if ($script:InspectionStorageFailure) { Add-AutoNote $script:InspectionStorageFailure -First }
        $R = [ordered]@{
            Index = 0; TechnicianEmail = $rawEmail; SerialNumber = $savedSerial; DurationMins = $finalDuration
            Charging = $window.FindName("cbCharging").Text; Screen = $window.FindName("cbScreen").Text
            Touchscreen = $window.FindName("cbTouch").Text; NetworkAdapters = $window.FindName("cbNetwork").Text
            Keyboard = $window.FindName("cbKeyboard").Text; MouseTrackpad = $window.FindName("cbMouse").Text
            VideoPorts = $window.FindName("cbVideo").Text; AudioOutput = $window.FindName("cbAudio").Text
            Microphone = $window.FindName("cbMic").Text; Camera = $window.FindName("cbCamera").Text
            USBPorts = $window.FindName("cbUSB").Text
            PalmRest = $window.FindName("cbPalm").Text; Backplate = $window.FindName("cbBackplate").Text
            BaseAndVents = $window.FindName("cbBase").Text; Hinge = $window.FindName("cbHinge").Text
            Notes = ($notesInput.Text.Trim() -replace "`r`n", " | " -replace "`n", " | " -replace "`r", " | ")
        }

        $FinalStatusDisplay = if (Test-DqaInspectionFailed ([pscustomobject]$R)) { "FAILED" } else { "PASSED" }
        $R["FinalStatus"] = $FinalStatusDisplay
        $row = [pscustomobject]$R

        if ($null -ne $recordToReplace) {
            $recordKey = Get-DqaRecordKey $recordToReplace
            $matchIndex = -1
            for ($i = 0; $i -lt $allData.Count; $i++) {
                if ((Get-DqaRecordKey $allData[$i]) -eq $recordKey) { $matchIndex = $i; break }
            }
            if ($matchIndex -lt 0) {
                [DsuMessageBox]::Show("This inspection changed while it was open. Reopen it from the dashboard to edit it.", "Inspection Changed", "OK", "Warning") | Out-Null
                return
            }
            if ($keepSavedDuration) { $row.DurationMins = $allData[$matchIndex].DurationMins }
            $allData[$matchIndex] = $row
        }
        else {
            [void]$allData.Add($row)
        }

        $addEmail = $rawEmail -notin $Global:SavedEmails
        $nextEmails = @($Global:SavedEmails)
        if ($addEmail) { $nextEmails += $rawEmail }
        $nextSettings = [ordered]@{}
        foreach ($key in $Global:AppSettings.Keys) { $nextSettings[$key] = $Global:AppSettings[$key] }
        if (Get-Setting "RememberLastTech") { $nextSettings["LastTechnician"] = $rawEmail }

        try {
            Update-ScriptData -DatabaseCSV (ConvertTo-DqaDatabaseCsv $allData) -SavedEmails $nextEmails -Settings $nextSettings
        }
        catch {
            [DsuMessageBox]::Show("The inspection could not be saved. Your entries are still on this form.`n`n$($_.Exception.Message)", "Save Error", "OK", "Error") | Out-Null
            return
        }
        if ($addEmail) { $techEmailInput.ItemsSource = $Global:SavedEmails }
        $techEmailInput.Text = $rawEmail
        $Global:EditingRecord = $null
        Reset-InspectionDetection
        Set-InspectionDurationMode
        $saveInspectionBtn.Content = "SAVE INSPECTION"

        # Settings > "Play a sound when an inspection is saved"
        if (Get-Setting "SaveSound") {
            if (Test-Path "C:\Windows\Media\chimes.wav") { (New-Object System.Media.SoundPlayer "C:\Windows\Media\chimes.wav").Play() } else { [System.Media.SystemSounds]::Beep.Play() }
        }
    
        $lastSavedLabel.Text = "SUCCESS: $savedSerial saved as $FinalStatusDisplay at $((Get-Date).ToString('hh:mm tt'))!"
        $lastSavedLabel.Foreground = "#39FF14" # Bright Green
        
        # Flash the save button background slightly
        $window.FindName("SaveInspectionBtn").Background = "#28A745"
        [System.Threading.Tasks.Task]::Delay(1000).ContinueWith([System.Action[System.Threading.Tasks.Task]] {
                try {
                    [System.Windows.Application]::Current.Dispatcher.Invoke({
                            $window.FindName("SaveInspectionBtn").ClearValue([System.Windows.Controls.Button]::BackgroundProperty)
                            $lastSavedLabel.Foreground = "#28A745"
                        })
                } catch {}
            })

        $serialInput.Text = ""; $notesInput.Text = ""; $audioStatusLabel.Text = " "; $audioStatusLabel.Background = "Transparent"
        $window.FindName("rbWipeYes").IsChecked = $false
        $Global:sessionStopwatch = New-Object System.Diagnostics.Stopwatch
        $Global:isTimerPaused = $true
        $window.FindName("LiveTimerLabel").Foreground = "#DC3545"
        $window.FindName("LiveTimerLabel").Text = "00:00"
        $window.FindName("PauseTimerBtn").Content = "$([char]0xE768)"
        if ($null -ne $Global:uiTimer) { $Global:uiTimer.Start() }
    
        foreach ($box in $hwBoxes) { $window.FindName($box).SelectedIndex = 0 }
        foreach ($box in $cosmeticBoxes) { $window.FindName($box).SelectedIndex = 0 }
    
        Test-FormCompletion
        Update-Dashboard
    })

function Update-Dashboard {
    if (-not [string]::IsNullOrWhiteSpace($Global:DatabaseCSV)) {
        $data = @($Global:DatabaseCSV | ConvertFrom-Csv)
    }
    else { $data = @() }

    foreach ($row in $data) {
        $statusText = if ($row.FinalStatus -eq "FAILED" -or (Test-DqaInspectionFailed $row)) { "FAILED" } else { "PASSED" }
        $row | Add-Member -MemberType NoteProperty -Name "FinalStatus" -Value $statusText -Force
        if ($row.DurationMins) { $row.DurationMins = [double]$row.DurationMins }
        if ($row.Index) { $row.Index = [int]$row.Index }
    }
    
    $data = $data | Sort-Object -Property Index

    $window.FindName("StatTotal").Text = $data.Count
    $window.FindName("StatPassed").Text = @($data | Where-Object { $_.FinalStatus -eq "PASSED" }).Count
    $window.FindName("StatFailed").Text = @($data | Where-Object { $_.FinalStatus -eq "FAILED" }).Count
    if ($data.Count -gt 0) { $window.FindName("StatAvg").Text = [math]::Round((($data | Measure-Object -Property DurationMins -Average).Average), 1) } 
    else { $window.FindName("StatAvg").Text = "0" }

    $search = $searchBox.Text.Trim().ToLower()
    $filter = ($window.FindName("FilterBox").SelectedItem).Content
    $filtered = @($data)

    if ($search) { $filtered = @($filtered | Where-Object { $_.SerialNumber -like "*$search*" }) }
    if ($filter -eq "PASSED only") { $filtered = @($filtered | Where-Object { $_.FinalStatus -eq "PASSED" }) }
    if ($filter -eq "FAILED only") { $filtered = @($filtered | Where-Object { $_.FinalStatus -eq "FAILED" }) }

    $resultsGrid.ItemsSource = $filtered
}

# Refresh the dashboard every time its tab is opened (replaces the Refresh button).
# SelectionChanged also bubbles up from the ComboBoxes and the grid, so only react to the tab control itself.
$window.FindName("MainTabControl").Add_SelectionChanged({
        param($s, $e)
        if ($e.OriginalSource -eq $s -and $s.SelectedIndex -eq 1) { Update-Dashboard }
    })

# ------------------------------------------------------------
# RESET / CLEAR FORM BUTTON
# ------------------------------------------------------------
$window.FindName("ResetFormBtn").Add_Click({
        if ($serialInput.Text -ne "" -or $notesInput.Text -ne "") {
            $res = [DsuMessageBox]::Show("Are you sure you want to clear all entered data?", "Confirm Reset", "YesNo", "Warning")
            if ($res -ne "Yes") { return }
        }
        Reset-InspectionDetection
        $serialInput.Text = ""
        $notesInput.Text = ""
        $Global:EditingRecord = $null
        Set-InspectionDurationMode
        $saveInspectionBtn.Content = "SAVE INSPECTION"
        $audioStatusLabel.Text = " "
        $audioStatusLabel.Background = "Transparent"
        # Settings > "Remember the last technician": keep the email, otherwise clear it
        $keptEmail = ""
        if (Get-Setting "RememberLastTech") {
            $keptEmail = $techEmailInput.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($keptEmail)) { $keptEmail = [string](Get-Setting "LastTechnician") }
        }
        $techEmailInput.Text = $keptEmail
        $lastSavedLabel.Text = " "
        $Global:sessionStopwatch = New-Object System.Diagnostics.Stopwatch
        $Global:isTimerPaused = $true
        $window.FindName("LiveTimerLabel").Foreground = "#DC3545"
        $window.FindName("LiveTimerLabel").Text = "00:00"
        $window.FindName("PauseTimerBtn").Content = "$([char]0xE768)"
        if ($null -ne $Global:uiTimer) { $Global:uiTimer.Start() }
        foreach ($box in $hwBoxes) { $window.FindName($box).SelectedIndex = 0 }
        foreach ($box in $cosmeticBoxes) { $window.FindName($box).SelectedIndex = 0 }
        $window.FindName("rbWipeYes").IsChecked = $false
        Test-FormCompletion
        # Email kept: detect the hardware again right away (runs in the background)
        if (-not [string]::IsNullOrWhiteSpace($keptEmail)) { Start-AutoDetect }
    })

$window.FindName("FilterBox").Add_SelectionChanged({ Update-Dashboard })

# ------------------------------------------------------------
# EXPORT ONLY (Creates the single output CSV on the USB)
# ------------------------------------------------------------
$window.FindName("ExportBtn").Add_Click({
        $exportData = @($resultsGrid.ItemsSource | Select-Object -Property $script:DqaCsvColumns)
        if ($exportData.Count -eq 0) {
            [DsuMessageBox]::Show("No data to export.", "Export", "OK", "Information") | Out-Null
            return
        }

        $exportMode = [string](Get-Setting "ExportMode")
        $customFolder = [string](Get-Setting "ExportFolder")
        $exportDir = [Environment]::GetFolderPath("Desktop")
        if ($exportMode -eq "Folder" -and $customFolder -and (Test-Path -LiteralPath $customFolder -PathType Container -ErrorAction SilentlyContinue)) {
            $exportDir = $customFolder
        }
        elseif ($exportMode -ne "Desktop") {
            $usbCheck = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=2" -ErrorAction SilentlyContinue |
                Where-Object { $_.Size -gt 0 } | Select-Object -First 1).DeviceID
            if ($usbCheck) { $exportDir = "$usbCheck\" }
        }
        if ([string]::IsNullOrWhiteSpace($exportDir)) {
            [DsuMessageBox]::Show("No export folder is available. Choose a folder in Settings.", "Export Error", "OK", "Error") | Out-Null
            return
        }

        $allData = @($Global:DatabaseCSV | ConvertFrom-Csv)
        $exportKeys = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($record in $exportData) { [void]$exportKeys.Add((Get-DqaRecordKey $record)) }
        $remainingData = @($allData | Where-Object { -not $exportKeys.Contains((Get-DqaRecordKey $_)) })
        if ($allData.Count - $remainingData.Count -ne $exportData.Count -or $exportKeys.Count -ne $exportData.Count) {
            [DsuMessageBox]::Show("The displayed rows no longer match inspection history. Refresh the dashboard and try again.", "Export Error", "OK", "Error") | Out-Null
            Update-Dashboard
            return
        }

        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
        $suffix = [guid]::NewGuid().ToString('N').Substring(0, 8)
        $savePath = [IO.Path]::Combine($exportDir, "DQA_Output_$($stamp)_$suffix.csv")
        $exportState = @{ Published = $false }
        try {
            Update-ScriptData -DatabaseCSV (ConvertTo-DqaDatabaseCsv $remainingData) -BeforeWrite {
                Write-DqaExport -Records $exportData -Path $savePath
                $exportState.Published = $true
            }
        }
        catch {
            if ($exportState.Published) {
                $message = "The CSV was saved to:`n$savePath`n`nInspection history could not be updated. The rows are still in DQA.`n`n$($_.Exception.Message)"
            }
            else {
                $message = "The CSV could not be exported. The rows are still in DQA.`n`n$($_.Exception.Message)"
            }
            [DsuMessageBox]::Show($message, "Export Error", "OK", "Error") | Out-Null
            return
        }
        Update-Dashboard
        [DsuMessageBox]::Show("Output successfully exported to:`n$savePath`n`nExported records have been removed from the database.", "Export Complete", "OK", "Information") | Out-Null
        Invoke-Item -LiteralPath $exportDir -ErrorAction SilentlyContinue
    })

# ------------------------------------------------------------
# EDIT SELECTED ROW
# ------------------------------------------------------------
$window.FindName("EditBtn").Add_Click({
        $selectedItem = $resultsGrid.SelectedItem
        if ($null -eq $selectedItem) {
            [DsuMessageBox]::Show("Please select a laptop record from the dashboard to edit.", "No Selection", "OK", "Warning")
            return
        }

        Reset-InspectionDetection
        $Global:EditingRecord = $selectedItem
        $script:EditingDatabaseCsv = $Global:DatabaseCSV
    
        $techEmailInput.Text = $selectedItem.TechnicianEmail
        $serialInput.Text = $selectedItem.SerialNumber
    
        $hwBoxesToMap = @{
            "cbCharging"="Charging"; "cbScreen"="Screen"; "cbTouch"="Touchscreen"; "cbNetwork"="NetworkAdapters";
            "cbKeyboard"="Keyboard"; "cbMouse"="MouseTrackpad"; "cbVideo"="VideoPorts"; "cbAudio"="AudioOutput";
            "cbMic"="Microphone"; "cbCamera"="Camera"; "cbUSB"="USBPorts"; "cbPalm"="PalmRest";
            "cbBackplate"="Backplate"; "cbBase"="BaseAndVents"; "cbHinge"="Hinge"
        }
        foreach ($key in $hwBoxesToMap.Keys) { $window.FindName($key).SelectedItem = $selectedItem.($hwBoxesToMap[$key]) }
    
        $notesInput.Text = $selectedItem.Notes
        Set-InspectionDurationMode
    
        $saveInspectionBtn.Content = "UPDATE INSPECTION"
        $window.FindName("MainTabControl").SelectedIndex = 0
        Test-FormCompletion
    })

# ------------------------------------------------------------
# DELETE SELECTED ROW (Requires PIN)
# ------------------------------------------------------------
$window.FindName("DeleteRowBtn").Add_Click({
        $selectedItem = $resultsGrid.SelectedItem
    
        if ($null -eq $selectedItem) {
            [DsuMessageBox]::Show("Please select a laptop record from the dashboard to delete.", "No Selection", "OK", "Warning")
            return
        }

        $enteredPin = Get-PinInput
    
        if (Test-PinValid -EnteredPin $enteredPin) {
            $allData = @($Global:DatabaseCSV | ConvertFrom-Csv)
            $selectedKey = Get-DqaRecordKey $selectedItem
            $remainingData = @($allData | Where-Object { (Get-DqaRecordKey $_) -ne $selectedKey })
            if ($allData.Count - $remainingData.Count -ne 1) {
                [DsuMessageBox]::Show("Could not locate the exact record in the database.", "Error", "OK", "Error") | Out-Null
                return
            }
            try {
                Update-ScriptData -DatabaseCSV (ConvertTo-DqaDatabaseCsv $remainingData)
            }
            catch {
                [DsuMessageBox]::Show("The record could not be deleted.`n`n$($_.Exception.Message)", "Delete Error", "OK", "Error") | Out-Null
                return
            }
            Update-Dashboard
            [DsuMessageBox]::Show("Laptop record '$($selectedItem.SerialNumber)' has been removed from inspection history.", "Success", "OK", "Information") | Out-Null
        }
        elseif ($null -ne $enteredPin) {
            [DsuMessageBox]::Show("Incorrect PIN entered. Deletion canceled.", "Security Alert", "OK", "Error")
        }
    })

# ------------------------------------------------------------
# CLEAR ALL HISTORY (Settings > Data; confirmation + Admin PIN)
# ------------------------------------------------------------
function Invoke-ClearAllHistory {
    $count = 0
    if (-not [string]::IsNullOrWhiteSpace($Global:DatabaseCSV)) { $count = @($Global:DatabaseCSV | ConvertFrom-Csv).Count }
    if ($count -eq 0) {
        [DsuMessageBox]::Show("There are no saved records to clear.", "Clear All History", "OK", "Information") | Out-Null
        return $false
    }
    $btns = @(
        @{ Text = "Delete all $count records"; Result = "Yes"; Style = "Danger" },
        @{ Text = "Cancel"; Result = "No"; Style = "Secondary"; IsDefault = $true; IsCancel = $true }
    )
    $sure = Show-DsuDialog -Title "Clear All History" -Kind "Warning" -Buttons $btns -DefaultResult "No" -Width 460 `
        -Message "This permanently deletes all $count saved inspection records from this tool.`n`nExport them first if you still need them. The Admin PIN is required next."
    if ($sure -ne "Yes") { return $false }

    $enteredPin = Get-PinInput
    if (Test-PinValid -EnteredPin $enteredPin) {
        try { Update-ScriptData -DatabaseCSV ($script:DqaCsvColumns -join ',') }
        catch {
            [DsuMessageBox]::Show("Inspection history could not be cleared.`n`n$($_.Exception.Message)", "Clear History Error", "OK", "Error") | Out-Null
            return $false
        }
        Update-Dashboard
        [DsuMessageBox]::Show("All inspection history has been cleared.", "Success", "OK", "Information") | Out-Null
        return $true
    }
    elseif ($null -ne $enteredPin) {
        [DsuMessageBox]::Show("Incorrect PIN entered. Nothing was deleted.", "Security Alert", "OK", "Error") | Out-Null
    }
    return $false
}
# ------------------------------------------------------------
# KEYBOARD SHORTCUTS
# ------------------------------------------------------------

$window.Add_KeyDown({
        param($src, $e)
        
        # --- Ctrl+Key Shortcuts ---
        if ([System.Windows.Input.Keyboard]::Modifiers -eq [System.Windows.Input.ModifierKeys]::Control) {
            switch ($e.Key) {
                'S' {
                    # Ctrl+S: Save/Update inspection (if enabled)
                    if ($saveInspectionBtn.IsEnabled) {
                        $window.FindName("SaveInspectionBtn").RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
                    }
                    $e.Handled = $true
                    return
                }
                'R' {
                    # Ctrl+R: Reset/Clear form
                    $window.FindName("ResetFormBtn").RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
                    $e.Handled = $true
                    return
                }
                'E' {
                    # Ctrl+E: Export to CSV
                    $window.FindName("ExportBtn").RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
                    $e.Handled = $true
                    return
                }
                'D1' {
                    # Ctrl+1: Switch to Inspection tab
                    $window.FindName("MainTabControl").SelectedIndex = 0
                    $e.Handled = $true
                    return
                }
                'D2' {
                    # Ctrl+2: Switch to Dashboard tab
                    $window.FindName("MainTabControl").SelectedIndex = 1
                    $e.Handled = $true
                    return
                }
                'F' {
                    # Ctrl+F: Focus search box (switch to Dashboard if needed)
                    $window.FindName("MainTabControl").SelectedIndex = 1
                    $searchBox.Focus() | Out-Null
                    $e.Handled = $true
                    return
                }
                'A' {
                    # Ctrl+A: Auto-fill remaining (only on Inspection tab)
                    if ($window.FindName("MainTabControl").SelectedIndex -eq 0) {
                        $window.FindName("BtnAutoFill").RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
                        $e.Handled = $true
                        return
                    }
                }
            }
        }

        # --- Single Key Shortcuts ---
        if ([System.Windows.Input.Keyboard]::Modifiers -eq [System.Windows.Input.ModifierKeys]::None) {
            if ($e.Key -eq 'F') {
                $focus = [System.Windows.Input.Keyboard]::FocusedElement
                if ($focus -isnot [System.Windows.Controls.TextBox] -and $focus -isnot [System.Windows.Controls.PasswordBox]) {
                    if ($window.WindowState -eq 'Maximized') {
                        $window.WindowState = 'Normal'
                    } else {
                        $window.WindowState = 'Maximized'
                    }
                    $e.Handled = $true
                    return
                }
            }
        }

        # Handle ComboBox cycling with Enter
        if ($e.Key -eq 'Enter') {
            $focus = [System.Windows.Input.Keyboard]::FocusedElement
            if ($focus -is [System.Windows.Controls.ComboBox]) {
                $cb = $focus
                # Exclude tech email and filter dropdowns
                if ($cb.Name -eq "TechEmailInput" -or $cb.Name -eq "FilterBox") { return }
            
                $items = $cb.ItemsSource
                if ($items -and $items.Count -gt 0) {
                    # Find current index and increment, looping back to 0
                    $idx = [array]::IndexOf($items, $cb.SelectedItem)
                    $nextIdx = $idx + 1
                    if ($nextIdx -ge $items.Count) { $nextIdx = 0 }
                    $Global:SkipAutoJump = $true
                    $cb.SelectedItem = $items[$nextIdx]
                    $Global:SkipAutoJump = $false
                    $e.Handled = $true
                }
            }
        }
    })

# ------------------------------------------------------------
# SETTINGS WINDOW (gear button in the header)
# Changes apply as they are saved to this script or its EXE data file.
# ------------------------------------------------------------
function Show-SettingsWindow {
    [xml]$setXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Settings" Width="1112" SizeToContent="Height" ResizeMode="NoResize" WindowStyle="None"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False" Background="Transparent" AllowsTransparency="True"
        FontFamily="Segoe UI" FontSize="13" SnapsToDevicePixels="True" UseLayoutRounding="True">
    <Window.Resources>
        <!-- Card: white box; its colored header band is the section's color code -->
        <Style x:Key="Card" TargetType="Border">
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#E3E6EA"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="Margin" Value="0"/>
        </Style>
        <Style x:Key="CardHead" TargetType="Border">
            <Setter Property="CornerRadius" Value="7,7,0,0"/>
            <Setter Property="BorderThickness" Value="0,0,0,3"/>
            <Setter Property="Padding" Value="14,9"/>
        </Style>
        <Style x:Key="CardIcon" TargetType="TextBlock">
            <Setter Property="FontFamily" Value="Segoe MDL2 Assets"/>
            <Setter Property="FontSize" Value="16"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="Margin" Value="0,0,10,0"/>
        </Style>
        <Style x:Key="CardTitle" TargetType="TextBlock">
            <Setter Property="FontSize" Value="15"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>
        <Style x:Key="CardBody" TargetType="StackPanel">
            <Setter Property="Margin" Value="16,12,16,14"/>
        </Style>
        <Style x:Key="Label" TargetType="TextBlock">
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="#1F2326"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>
        <Style x:Key="Hint" TargetType="TextBlock">
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Foreground" Value="#6C757D"/>
            <Setter Property="TextWrapping" Value="Wrap"/>
        </Style>
        <Style x:Key="Btn" TargetType="Button">
            <Setter Property="Background" Value="#00A9E0"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="14,0"/>
            <Setter Property="Height" Value="32"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Opacity" Value="0.88"/></Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter TargetName="bd" Property="Opacity" Value="0.45"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Close (X) button in the navy header -->
        <Style x:Key="CloseX" TargetType="Button">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="Transparent" CornerRadius="6">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#00A9E0"/></Trigger>
                            <Trigger Property="IsPressed" Value="True"><Setter TargetName="bd" Property="Background" Value="#0087B3"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- On/off switch: the knob slides and the track fades from grey to DSU blue -->
        <Style x:Key="Switch" TargetType="CheckBox">
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CheckBox">
                        <Grid x:Name="root" Width="44" Height="24" Background="Transparent">
                            <Border x:Name="trackOff" CornerRadius="12" Background="#C9CED4"/>
                            <Border x:Name="trackOn" CornerRadius="12" Background="#00A9E0" Opacity="0"/>
                            <Ellipse x:Name="knob" Width="18" Height="18" Fill="White" HorizontalAlignment="Left" Margin="3,0,0,0">
                                <Ellipse.Effect><DropShadowEffect BlurRadius="4" ShadowDepth="1" Direction="270" Opacity="0.30"/></Ellipse.Effect>
                                <Ellipse.RenderTransform><TranslateTransform X="0"/></Ellipse.RenderTransform>
                            </Ellipse>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Trigger.EnterActions>
                                    <BeginStoryboard>
                                        <Storyboard>
                                            <DoubleAnimation Storyboard.TargetName="knob" Storyboard.TargetProperty="(UIElement.RenderTransform).(TranslateTransform.X)" To="20" Duration="0:0:0.15"/>
                                            <DoubleAnimation Storyboard.TargetName="trackOn" Storyboard.TargetProperty="Opacity" To="1" Duration="0:0:0.15"/>
                                        </Storyboard>
                                    </BeginStoryboard>
                                </Trigger.EnterActions>
                                <Trigger.ExitActions>
                                    <BeginStoryboard>
                                        <Storyboard>
                                            <DoubleAnimation Storyboard.TargetName="knob" Storyboard.TargetProperty="(UIElement.RenderTransform).(TranslateTransform.X)" To="0" Duration="0:0:0.15"/>
                                            <DoubleAnimation Storyboard.TargetName="trackOn" Storyboard.TargetProperty="Opacity" To="0" Duration="0:0:0.15"/>
                                        </Storyboard>
                                    </BeginStoryboard>
                                </Trigger.ExitActions>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="root" Property="Opacity" Value="0.88"/></Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter TargetName="root" Property="Opacity" Value="0.5"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Radio buttons: round ring, DSU-blue dot when selected -->
        <Style TargetType="RadioButton">
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Foreground" Value="#1F2326"/>
            <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="RadioButton">
                        <Grid Background="Transparent">
                            <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                            <Grid Width="18" Height="18" VerticalAlignment="Top" Margin="0,0,0,0">
                                <Ellipse x:Name="ring" Stroke="#8E959C" StrokeThickness="1.5" Fill="White"/>
                                <Ellipse x:Name="dot" Width="8" Height="8" Fill="#00A9E0" Opacity="0"/>
                            </Grid>
                            <ContentPresenter Grid.Column="1" Margin="9,0,0,0" VerticalAlignment="Top" RecognizesAccessKey="True"/>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="ring" Property="Stroke" Value="#00A9E0"/></Trigger>
                            <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="ring" Property="Stroke" Value="#00A9E0"/></Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="ring" Property="Stroke" Value="#00A9E0"/>
                                <Setter TargetName="ring" Property="StrokeThickness" Value="2"/>
                                <Setter TargetName="dot" Property="Opacity" Value="1"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.5"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Rounded text boxes (email input, custom folder) -->
        <Style TargetType="TextBox">
            <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#C5CBD3"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="Bd" CornerRadius="6" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" SnapsToDevicePixels="True">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" Focusable="False" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden" VerticalAlignment="{TemplateBinding VerticalContentAlignment}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#00A9E0"/></Trigger>
                            <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#00A9E0"/></Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter TargetName="Bd" Property="Opacity" Value="0.55"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Rounded email list with rounded row highlights -->
        <Style TargetType="ListBox">
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#C5CBD3"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBox">
                        <Border CornerRadius="6" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Padding="3" SnapsToDevicePixels="True">
                            <ScrollViewer Focusable="False" HorizontalScrollBarVisibility="{TemplateBinding ScrollViewer.HorizontalScrollBarVisibility}" VerticalScrollBarVisibility="Auto">
                                <ItemsPresenter/>
                            </ScrollViewer>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ListBoxItem">
            <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                        <Border x:Name="Bd" CornerRadius="4" Padding="8,4" Margin="0,1" Background="Transparent">
                            <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#E6F6FC"/></Trigger>
                            <Trigger Property="IsSelected" Value="True"><Setter TargetName="Bd" Property="Background" Value="#D2EEF9"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Slim rounded scrollbars (same as the main screen) -->
        <Style x:Key="SetScrollPageButton" TargetType="RepeatButton">
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="IsTabStop" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="RepeatButton"><Border Background="Transparent"/></ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="SetScrollThumb" TargetType="Thumb">
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="IsTabStop" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Thumb">
                        <Border x:Name="tb" CornerRadius="4" Background="#B8BEC4"/>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="tb" Property="Background" Value="#8E959C"/></Trigger>
                            <Trigger Property="IsDragging" Value="True"><Setter TargetName="tb" Property="Background" Value="#00A9E0"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Width" Value="12"/>
            <Setter Property="MinWidth" Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollBar">
                        <Border CornerRadius="6" Background="#EEF0F2" Margin="2">
                            <Track x:Name="PART_Track" IsDirectionReversed="True">
                                <Track.DecreaseRepeatButton><RepeatButton Style="{StaticResource SetScrollPageButton}" Command="ScrollBar.PageUpCommand"/></Track.DecreaseRepeatButton>
                                <Track.Thumb><Thumb Style="{StaticResource SetScrollThumb}" Margin="1"/></Track.Thumb>
                                <Track.IncreaseRepeatButton><RepeatButton Style="{StaticResource SetScrollPageButton}" Command="ScrollBar.PageDownCommand"/></Track.IncreaseRepeatButton>
                            </Track>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="Orientation" Value="Horizontal">
                    <Setter Property="Width" Value="Auto"/>
                    <Setter Property="MinWidth" Value="0"/>
                    <Setter Property="Height" Value="12"/>
                    <Setter Property="MinHeight" Value="12"/>
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="ScrollBar">
                                <Border CornerRadius="6" Background="#EEF0F2" Margin="2">
                                    <Track x:Name="PART_Track" IsDirectionReversed="False">
                                        <Track.DecreaseRepeatButton><RepeatButton Style="{StaticResource SetScrollPageButton}" Command="ScrollBar.PageLeftCommand"/></Track.DecreaseRepeatButton>
                                        <Track.Thumb><Thumb Style="{StaticResource SetScrollThumb}" Margin="1"/></Track.Thumb>
                                        <Track.IncreaseRepeatButton><RepeatButton Style="{StaticResource SetScrollPageButton}" Command="ScrollBar.PageRightCommand"/></Track.IncreaseRepeatButton>
                                    </Track>
                                </Border>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="KeyCap" TargetType="Border">
            <Setter Property="Background" Value="#EEF3F7"/>
            <Setter Property="BorderBrush" Value="#C5CBD3"/>
            <Setter Property="BorderThickness" Value="1,1,1,2"/>
            <Setter Property="CornerRadius" Value="5"/>
            <Setter Property="Padding" Value="7,2"/>
            <Setter Property="Margin" Value="0,0,0,7"/>
            <Setter Property="MinWidth" Value="62"/>
            <Setter Property="HorizontalAlignment" Value="Left"/>
        </Style>
        <Style x:Key="KeyCapText" TargetType="TextBlock">
            <Setter Property="FontSize" Value="11.5"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="#003552"/>
            <Setter Property="HorizontalAlignment" Value="Center"/>
        </Style>
        <Style x:Key="KeyDesc" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#1F2326"/>
            <Setter Property="VerticalAlignment" Value="Top"/>
            <Setter Property="Margin" Value="10,2,0,7"/>
        </Style>
    </Window.Resources>

    <!-- Rounded window: shadow layer + rounded panel (the window itself is transparent) -->
    <Grid Margin="16">
        <Border CornerRadius="12" Background="#F8FAFC">
            <Border.Effect><DropShadowEffect BlurRadius="22" ShadowDepth="4" Direction="270" Opacity="0.35"/></Border.Effect>
        </Border>
        <Border CornerRadius="12" Background="#F8FAFC" BorderBrush="#003552" BorderThickness="1">
    <DockPanel>
        <!-- Header: drag to move; X closes Settings -->
        <Border x:Name="SetHeader" DockPanel.Dock="Top" Background="#003552" CornerRadius="11,11,0,0" Padding="20,10,10,10">
            <DockPanel>
                <Button x:Name="BtnCloseX" DockPanel.Dock="Right" Style="{StaticResource CloseX}" Width="38" Height="38" ToolTip="Close settings (Esc)">
                    <TextBlock Text="&#xE711;" FontFamily="Segoe MDL2 Assets" FontSize="14" Foreground="White"/>
                </Button>
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="&#xE713;" FontFamily="Segoe MDL2 Assets" FontSize="20" Foreground="#00A9E0" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <TextBlock Text="Settings" FontSize="20" FontWeight="Bold" Foreground="White" VerticalAlignment="Center"/>
                </StackPanel>
            </DockPanel>
        </Border>
        <Border DockPanel.Dock="Top" Height="3" Background="#00A9E0"/>

        <Border DockPanel.Dock="Bottom" Background="White" BorderBrush="#E3E6EA" BorderThickness="0,1,0,0" CornerRadius="0,0,11,11" Padding="20,12">
            <TextBlock x:Name="SettingsSaveStatus" TextWrapping="Wrap" FontWeight="SemiBold"/>
        </Border>

        <ScrollViewer VerticalScrollBarVisibility="Auto">
            <!-- Even 3 x 2 grid: every card in a row has the same height, all gaps are 16 px -->
            <Grid Margin="20,18,20,18">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/><ColumnDefinition Width="16"/><ColumnDefinition Width="*"/><ColumnDefinition Width="16"/><ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/><RowDefinition Height="16"/><RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Row 1 -->
                <Border Style="{StaticResource Card}" Grid.Row="0" Grid.Column="0">
                    <DockPanel>
                        <Border DockPanel.Dock="Top" Style="{StaticResource CardHead}" Background="#EEF3F7" BorderBrush="#00A9E0">
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="&#xE9D9;" Style="{StaticResource CardIcon}" Foreground="#00A9E0"/>
                                <TextBlock Text="Tests" Style="{StaticResource CardTitle}" Foreground="#003552"/>
                            </StackPanel>
                        </Border>
                        <StackPanel Style="{StaticResource CardBody}">
                            <TextBlock Text="Test volume" Style="{StaticResource Label}" Margin="0,0,0,4"/>
                            <DockPanel Margin="0,0,0,14">
                                <Button x:Name="BtnPreviewChime" DockPanel.Dock="Right" Content="Preview" Style="{StaticResource Btn}" Height="28" Margin="10,0,0,0"/>
                                <TextBlock x:Name="SetVolumeText" DockPanel.Dock="Right" Text="25%" FontWeight="Bold" Width="40" TextAlignment="Right" VerticalAlignment="Center"/>
                                <Slider x:Name="SetVolume" Minimum="5" Maximum="100" TickFrequency="5" IsSnapToTickEnabled="True" VerticalAlignment="Center"/>
                            </DockPanel>

                            <DockPanel Margin="0,0,0,8">
                                <ComboBox x:Name="SetAudioSound" DockPanel.Dock="Right" Width="142" Height="28"
                                          AutomationProperties.Name="Test sound"/>
                                <TextBlock Text="Test sound" Style="{StaticResource Label}"/>
                            </DockPanel>
                            <DockPanel Margin="0,0,0,8">
                                <ComboBox x:Name="SetTouchTargets" DockPanel.Dock="Right" Width="90" Height="28"/>
                                <TextBlock Text="Touch circles" Style="{StaticResource Label}"/>
                            </DockPanel>
                            <DockPanel Margin="0,0,0,4">
                                <ComboBox x:Name="SetTouchSeconds" DockPanel.Dock="Right" Width="90" Height="28"/>
                                <TextBlock Text="Touch time (s)" Style="{StaticResource Label}"/>
                            </DockPanel>
                        </StackPanel>
                    </DockPanel>
                </Border>

                <Border Style="{StaticResource Card}" Grid.Row="0" Grid.Column="2">
                    <DockPanel>
                        <Border DockPanel.Dock="Top" Style="{StaticResource CardHead}" Background="#EEF3F7" BorderBrush="#00A9E0">
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="&#xE715;" Style="{StaticResource CardIcon}" Foreground="#00A9E0"/>
                                <TextBlock Text="Technician Emails" Style="{StaticResource CardTitle}" Foreground="#003552"/>
                            </StackPanel>
                        </Border>
                        <StackPanel Style="{StaticResource CardBody}">
                            <DockPanel Margin="0,0,0,8">
                                <Button x:Name="BtnRemoveEmail" DockPanel.Dock="Right" Content="Remove" Style="{StaticResource Btn}" Background="#4D4F53" Width="80" VerticalAlignment="Top" Margin="8,0,0,0"/>
                                <ListBox x:Name="EmailList" Height="96" ScrollViewer.HorizontalScrollBarVisibility="Disabled">
                                    <ListBox.ItemTemplate>
                                        <DataTemplate>
                                            <TextBlock Text="{Binding}" TextTrimming="CharacterEllipsis" ToolTip="{Binding}"/>
                                        </DataTemplate>
                                    </ListBox.ItemTemplate>
                                </ListBox>
                            </DockPanel>
                            <DockPanel Margin="0,0,0,12">
                                <Button x:Name="BtnAddEmail" DockPanel.Dock="Right" Content="Add" Style="{StaticResource Btn}" Background="#00A9E0" Width="80" Margin="8,0,0,0"/>
                                <Grid>
                                    <TextBox x:Name="NewEmailBox" Height="32" VerticalContentAlignment="Center" Padding="8,0"/>
                                    <TextBlock x:Name="NewEmailHint" Text="first.last@trojans.dsu.edu" Foreground="#9A9C9E" Margin="11,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False"/>
                                </Grid>
                            </DockPanel>
                            <DockPanel>
                                <CheckBox x:Name="SetRememberTech" DockPanel.Dock="Right" Style="{StaticResource Switch}" Margin="8,0,0,0"/>
                                <StackPanel>
                                    <TextBlock Text="Remember last technician" Style="{StaticResource Label}"/>
                                </StackPanel>
                            </DockPanel>
                        </StackPanel>
                    </DockPanel>
                </Border>

                <Border Style="{StaticResource Card}" Grid.Row="0" Grid.Column="4">
                    <DockPanel>
                        <Border DockPanel.Dock="Top" Style="{StaticResource CardHead}" Background="#EEF3F7" BorderBrush="#00A9E0">
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="&#xE701;" Style="{StaticResource CardIcon}" Foreground="#00A9E0"/>
                                <TextBlock Text="General" Style="{StaticResource CardTitle}" Foreground="#003552"/>
                            </StackPanel>
                        </Border>
                        <StackPanel Style="{StaticResource CardBody}">
                            <DockPanel Margin="0,0,0,12">
                                <CheckBox x:Name="SetWifi" DockPanel.Dock="Right" Style="{StaticResource Switch}" Margin="8,0,0,0"/>
                                <StackPanel>
                                    <TextBlock Text="Auto-connect Guest Wi-Fi" Style="{StaticResource Label}" TextWrapping="Wrap"/>
                                </StackPanel>
                            </DockPanel>
                            <DockPanel>
                                <CheckBox x:Name="SetSaveSound" DockPanel.Dock="Right" Style="{StaticResource Switch}" Margin="8,0,0,0"/>
                                <StackPanel>
                                    <TextBlock Text="Sound on save" Style="{StaticResource Label}" TextWrapping="Wrap"/>
                                </StackPanel>
                            </DockPanel>
                        </StackPanel>
                    </DockPanel>
                </Border>

                <!-- Row 2 -->
                <Border Style="{StaticResource Card}" Grid.Row="2" Grid.Column="0">
                    <DockPanel>
                        <Border DockPanel.Dock="Top" Style="{StaticResource CardHead}" Background="#EEF3F7" BorderBrush="#00A9E0">
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="&#xE898;" Style="{StaticResource CardIcon}" Foreground="#00A9E0"/>
                                <TextBlock Text="Export" Style="{StaticResource CardTitle}" Foreground="#003552"/>
                            </StackPanel>
                        </Border>
                        <StackPanel Style="{StaticResource CardBody}">
                            <RadioButton x:Name="ExportAuto" GroupName="ExportDest" Margin="0,0,0,6">
                                <TextBlock Text="USB drive, else Desktop" TextWrapping="Wrap"/>
                            </RadioButton>
                            <RadioButton x:Name="ExportDesktop" GroupName="ExportDest" Content="Desktop" Margin="0,0,0,6"/>
                            <RadioButton x:Name="ExportFolder" GroupName="ExportDest" Content="Custom folder" Margin="0,0,0,6"/>
                            <DockPanel>
                                <Button x:Name="BtnBrowse" DockPanel.Dock="Right" Content="Browse..." Style="{StaticResource Btn}" Background="#00A9E0" Height="30" Margin="8,0,0,0"/>
                                <TextBox x:Name="ExportFolderText" IsReadOnly="True" Height="30" VerticalContentAlignment="Center" FontSize="12" Background="#F8F9FA"/>
                            </DockPanel>
                        </StackPanel>
                    </DockPanel>
                </Border>

                <Border Style="{StaticResource Card}" Grid.Row="2" Grid.Column="2">
                    <DockPanel>
                        <Border DockPanel.Dock="Top" Style="{StaticResource CardHead}" Background="#EEF3F7" BorderBrush="#00A9E0">
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="&#xE74D;" Style="{StaticResource CardIcon}" Foreground="#00A9E0"/>
                                <TextBlock Text="Data" Style="{StaticResource CardTitle}" Foreground="#003552"/>
                            </StackPanel>
                        </Border>
                        <StackPanel Style="{StaticResource CardBody}">
                            <TextBlock x:Name="RecordCountText" FontWeight="SemiBold" Foreground="#003552" Margin="0,0,0,4"/>
                            <TextBlock Text="Permanent. Admin PIN required." Style="{StaticResource Hint}" Margin="0,0,0,10"/>
                            <Button x:Name="BtnClearAll" Content="Clear All History" Style="{StaticResource Btn}" Background="#DC3545" HorizontalAlignment="Left"/>
                        </StackPanel>
                    </DockPanel>
                </Border>

                <Border Style="{StaticResource Card}" Grid.Row="2" Grid.Column="4">
                    <DockPanel>
                        <Border DockPanel.Dock="Top" Style="{StaticResource CardHead}" Background="#EEF3F7" BorderBrush="#00A9E0">
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="&#xE765;" Style="{StaticResource CardIcon}" Foreground="#00A9E0"/>
                                <TextBlock Text="Keyboard Shortcuts" Style="{StaticResource CardTitle}" Foreground="#003552"/>
                            </StackPanel>
                        </Border>
                        <Grid Margin="16,12,16,7">
                            <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                            <Border Grid.Row="0" Style="{StaticResource KeyCap}"><TextBlock Text="Ctrl+S" Style="{StaticResource KeyCapText}"/></Border>
                            <TextBlock Grid.Row="0" Grid.Column="1" Text="Save" Style="{StaticResource KeyDesc}"/>
                            <Border Grid.Row="1" Style="{StaticResource KeyCap}"><TextBlock Text="Ctrl+R" Style="{StaticResource KeyCapText}"/></Border>
                            <TextBlock Grid.Row="1" Grid.Column="1" Text="Reset form" Style="{StaticResource KeyDesc}"/>
                            <Border Grid.Row="2" Style="{StaticResource KeyCap}"><TextBlock Text="Ctrl+A" Style="{StaticResource KeyCapText}"/></Border>
                            <TextBlock Grid.Row="2" Grid.Column="1" Text="Auto-fill" Style="{StaticResource KeyDesc}"/>
                            <Border Grid.Row="3" Style="{StaticResource KeyCap}"><TextBlock Text="Ctrl+E" Style="{StaticResource KeyCapText}"/></Border>
                            <TextBlock Grid.Row="3" Grid.Column="1" Text="Export CSV" Style="{StaticResource KeyDesc}"/>
                            <Border Grid.Row="4" Style="{StaticResource KeyCap}"><TextBlock Text="Ctrl+F" Style="{StaticResource KeyCapText}"/></Border>
                            <TextBlock Grid.Row="4" Grid.Column="1" Text="Search" Style="{StaticResource KeyDesc}"/>
                            <Border Grid.Row="5" Style="{StaticResource KeyCap}"><TextBlock Text="Ctrl+1 / 2" Style="{StaticResource KeyCapText}"/></Border>
                            <TextBlock Grid.Row="5" Grid.Column="1" Text="Switch tabs" Style="{StaticResource KeyDesc}"/>
                            <Border Grid.Row="6" Style="{StaticResource KeyCap}"><TextBlock Text="Enter" Style="{StaticResource KeyCapText}"/></Border>
                            <TextBlock Grid.Row="6" Grid.Column="1" Text="Next value" Style="{StaticResource KeyDesc}"/>
                        </Grid>
                    </DockPanel>
                </Border>
            </Grid>
        </ScrollViewer>
    </DockPanel>
        </Border>
    </Grid>
</Window>
'@
    $setWin = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $setXaml))
    Set-DqaFocusVisual $setWin
    # Share the main window's dropdown styles. Settings dimensions come from its XAML.
    $setWin.Resources[[System.Windows.Controls.ComboBox]] = $window.Resources[[System.Windows.Controls.ComboBox]]
    $setWin.Resources[[System.Windows.Controls.ComboBoxItem]] = $window.Resources[[System.Windows.Controls.ComboBoxItem]]
    $setWin.Owner = $window
    # Never larger than the screen (the cards scroll on small screens)
    $workArea = [System.Windows.SystemParameters]::WorkArea
    $setWin.Width = [Math]::Min(1112, $workArea.Width)
    $setWin.MaxHeight = $workArea.Height - 20

    # Default position: centred across the main window, just below its navy header (the header and timer stay visible)
    $setWin.WindowStartupLocation = "Manual"
    $setWin.Add_Loaded({
            try {
                $source = [System.Windows.PresentationSource]::FromVisual($window)
                $toDip = $source.CompositionTarget.TransformFromDevice
                $root = $window.Content
                $rootTopLeft = $toDip.Transform($root.PointToScreen((New-Object System.Windows.Point(0, 0))))
                $tabsTop = $toDip.Transform($window.FindName("MainTabControl").PointToScreen((New-Object System.Windows.Point(0, 0)))).Y
                $left = $rootTopLeft.X + ($root.ActualWidth - $setWin.ActualWidth) / 2
                $top = $tabsTop + 4      # the visible panel starts ~20 px below the header (16 px is the shadow margin)
                $wa = [System.Windows.SystemParameters]::WorkArea
                if ($top + $setWin.ActualHeight -gt $wa.Bottom) { $top = [Math]::Max($wa.Top, $wa.Bottom - $setWin.ActualHeight) }
                $left = [Math]::Min([Math]::Max($wa.Left, $left), [Math]::Max($wa.Left, $wa.Right - $setWin.ActualWidth))
                $setWin.Left = $left
                $setWin.Top = $top
            }
            catch {
                # Fallback: centre on screen
                $wa = [System.Windows.SystemParameters]::WorkArea
                $setWin.Left = $wa.Left + ($wa.Width - $setWin.ActualWidth) / 2
                $setWin.Top = $wa.Top + ($wa.Height - $setWin.ActualHeight) / 2
            }
        })

    $c = @{}
    foreach ($n in @("SetVolume", "SetVolumeText", "SetAudioSound", "BtnPreviewChime", "SetTouchTargets", "SetTouchSeconds",
            "EmailList", "BtnRemoveEmail", "NewEmailBox", "NewEmailHint", "BtnAddEmail", "SetRememberTech",
            "ExportAuto", "ExportDesktop", "ExportFolder", "ExportFolderText", "BtnBrowse",
            "SetWifi", "SetSaveSound", "SettingsSaveStatus", "BtnClearAll", "RecordCountText", "BtnCloseX", "SetHeader")) {
        $c[$n] = $setWin.FindName($n)
    }

    # ---- Load the current values ----
    $c.SetVolume.Value = [int](Get-Setting "AudioVolume")
    $c.SetVolumeText.Text = "$([int]$c.SetVolume.Value)%"
    $c.SetAudioSound.ItemsSource = @(
        $Global:AppSettingsDefaults.AudioSound, "Bright rise", "High cascade", "Major cadence", "Echo melody", "Signal fanfare"
    )
    $c.SetAudioSound.SelectedItem = [string](Get-Setting "AudioSound")
    if ($null -eq $c.SetAudioSound.SelectedItem) { $c.SetAudioSound.SelectedIndex = 0 }


    $c.SetTouchTargets.ItemsSource = @(3, 4, 5, 6, 7, 8, 9, 10)
    $c.SetTouchTargets.SelectedItem = [int](Get-Setting "TouchTargets")
    if ($null -eq $c.SetTouchTargets.SelectedItem) { $c.SetTouchTargets.SelectedItem = 5 }
    $c.SetTouchSeconds.ItemsSource = @(3, 4, 5, 6, 7, 8, 10, 15, 20)
    $c.SetTouchSeconds.SelectedItem = [int](Get-Setting "TouchSeconds")
    if ($null -eq $c.SetTouchSeconds.SelectedItem) { $c.SetTouchSeconds.SelectedItem = 5 }

    # The dialog edits its own list, then publishes it after a successful write.
    $workEmails = New-Object System.Collections.ObjectModel.ObservableCollection[string]
    foreach ($em in $Global:SavedEmails) { if (-not [string]::IsNullOrWhiteSpace($em)) { $workEmails.Add([string]$em) } }
    $c.EmailList.ItemsSource = $workEmails
    $c.SetRememberTech.IsChecked = [bool](Get-Setting "RememberLastTech")

    switch ([string](Get-Setting "ExportMode")) {
        "Desktop" { $c.ExportDesktop.IsChecked = $true }
        "Folder" { $c.ExportFolder.IsChecked = $true }
        default { $c.ExportAuto.IsChecked = $true }
    }
    $c.ExportFolderText.Text = [string](Get-Setting "ExportFolder")

    $c.SetWifi.IsChecked = [bool](Get-Setting "WifiAutoConnect")
    $c.SetSaveSound.IsChecked = [bool](Get-Setting "SaveSound")

    $readDraft = {
        [ordered]@{
            Settings = [ordered]@{
                AudioVolume      = [int]$c.SetVolume.Value
                AudioSound       = [string]$c.SetAudioSound.SelectedItem
                TouchTargets     = [int]$c.SetTouchTargets.SelectedItem
                TouchSeconds     = [int]$c.SetTouchSeconds.SelectedItem
                WifiAutoConnect  = [bool]$c.SetWifi.IsChecked
                SaveSound        = [bool]$c.SetSaveSound.IsChecked
                RememberLastTech = [bool]$c.SetRememberTech.IsChecked
                LastTechnician   = if ($c.SetRememberTech.IsChecked) { Get-Setting "LastTechnician" } else { "" }
                ExportMode       = if ($c.ExportDesktop.IsChecked) { "Desktop" } elseif ($c.ExportFolder.IsChecked) { "Folder" } else { "Auto" }
                ExportFolder     = $c.ExportFolderText.Text.Trim()
            }
            Emails = [string[]]$workEmails
        }
    }

    $saveState = @{
        Snapshot = (& $readDraft | ConvertTo-Json -Compress -Depth 4)
        Failed = $false
    }
    $volumeSaveTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $volumeSaveTimer.Interval = [TimeSpan]::FromMilliseconds(250)

    $showSaveStatus = {
        param([bool]$FolderRequired)
        $c.SettingsSaveStatus.ToolTip = $null
        $savedAt = [DateTimeOffset]::MinValue
        if ($FolderRequired) {
            $c.SettingsSaveStatus.Text = "Choose an existing folder for custom exports."
            $c.SettingsSaveStatus.Foreground = "#805500"
        }
        elseif ([DateTimeOffset]::TryParse([string](Get-Setting "SettingsSavedAt"), [ref]$savedAt)) {
            $c.SettingsSaveStatus.Text = "Saved for next launch ($($savedAt.ToLocalTime().ToString('g')))"
            $c.SettingsSaveStatus.Foreground = "#17643A"
        }
        else {
            $c.SettingsSaveStatus.Text = "Settings save automatically"
            $c.SettingsSaveStatus.Foreground = "#4D4F53"
        }
    }

    $saveSettings = {
        $volumeSaveTimer.Stop()
        $draft = & $readDraft
        $folder = $draft.Settings["ExportFolder"]
        $folderRequired = $draft.Settings["ExportMode"] -eq "Folder" -and
            ([string]::IsNullOrWhiteSpace($folder) -or -not (Test-Path -LiteralPath $folder -PathType Container -ErrorAction SilentlyContinue))
        if ($folderRequired) {
            # Keep the saved destination while the technician chooses a valid folder.
            $draft.Settings["ExportMode"] = [string](Get-Setting "ExportMode")
            $draft.Settings["ExportFolder"] = [string](Get-Setting "ExportFolder")
        }

        $snapshot = $draft | ConvertTo-Json -Compress -Depth 4
        if ($snapshot -cne $saveState.Snapshot) {
            $draft.Settings["SettingsSavedAt"] = [DateTimeOffset]::UtcNow.ToString("o")
            $emailsChanged = ($Global:SavedEmails -join [char]10) -cne ($draft.Emails -join [char]10)
            try {
                Update-ScriptData -Settings $draft.Settings -SavedEmails $draft.Emails
            }
            catch {
                $saveState.Failed = $true
                $c.SettingsSaveStatus.Text = "Couldn't save settings. $($_.Exception.Message)"
                $c.SettingsSaveStatus.Foreground = "#B42318"
                $c.SettingsSaveStatus.ToolTip = $_.Exception.Message
                return
            }

            $saveState.Snapshot = $snapshot
            # Refresh the email dropdown only when the list changes.
            if ($emailsChanged) {
                $typed = $techEmailInput.Text
                $techEmailInput.ItemsSource = $Global:SavedEmails
                $techEmailInput.Text = $typed
            }
        }

        $saveState.Failed = $false
        & $showSaveStatus $folderRequired
    }
    & $showSaveStatus

    # Batch slider movement into one write. Other controls save on change.
    $volumeSaveTimer.Add_Tick($saveSettings)
    $c.SetVolume.Add_ValueChanged({
            $c.SetVolumeText.Text = "$([int]$c.SetVolume.Value)%"
            $volumeSaveTimer.Stop()
            $volumeSaveTimer.Start()
        })
    $c.SetAudioSound.Add_SelectionChanged($saveSettings)
    $c.SetTouchTargets.Add_SelectionChanged($saveSettings)
    $c.SetTouchSeconds.Add_SelectionChanged($saveSettings)
    foreach ($name in @("SetWifi", "SetSaveSound", "SetRememberTech")) {
        $c[$name].Add_Checked($saveSettings)
        $c[$name].Add_Unchecked($saveSettings)
    }
    foreach ($name in @("ExportAuto", "ExportDesktop", "ExportFolder")) {
        $c[$name].Add_Checked($saveSettings)
    }
    $workEmails.Add_CollectionChanged($saveSettings)

    $c.BtnPreviewChime.Add_Click({
            if ($Global:audioState -gt 0) {
                [DsuMessageBox]::Show("Wait for the current audio test to finish before previewing a sound.", "Test Sound", "OK", "Information") | Out-Null
                return
            }
            [AudioHelper]::SetVolume([int]$c.SetVolume.Value)
            if (-not (Start-Chime -Sound $c.SetAudioSound.SelectedItem)) {
                [DsuMessageBox]::Show("The selected sound could not play. Check the Windows audio output.", "Test Sound", "OK", "Warning") | Out-Null
            }
        })

    $c.NewEmailBox.Add_TextChanged({
            $c.NewEmailHint.Visibility = if ([string]::IsNullOrEmpty($c.NewEmailBox.Text)) { "Visible" } else { "Collapsed" }
        })

    $addEmail = {
        $new = $c.NewEmailBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($new)) { return }
        if ($new -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
            [DsuMessageBox]::Show("Please enter a valid email address.", "Technician Emails", "OK", "Warning") | Out-Null
            return
        }
        $new = Format-TitleCaseEmail $new
        if (@($workEmails | Where-Object { $_ -ieq $new }).Count -gt 0) {
            [DsuMessageBox]::Show("That email is already in the list.", "Technician Emails", "OK", "Information") | Out-Null
            return
        }
        $workEmails.Add($new)
        $c.NewEmailBox.Text = ""
    }
    $c.BtnAddEmail.Add_Click({ & $addEmail })
    $c.NewEmailBox.Add_KeyDown({
            param($s, $e)
            if ($e.Key -eq "Return") { & $addEmail; $e.Handled = $true }
        })

    $c.BtnRemoveEmail.Add_Click({
            $sel = $c.EmailList.SelectedItem
            if ($sel) { $workEmails.Remove([string]$sel) | Out-Null }
        })

    $c.BtnBrowse.Add_Click({
            $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
            $dlg.Description = "Choose where exported CSV files are saved"
            if (-not [string]::IsNullOrWhiteSpace($c.ExportFolderText.Text) -and
                (Test-Path -LiteralPath $c.ExportFolderText.Text -PathType Container)) {
                $dlg.SelectedPath = $c.ExportFolderText.Text
            }
            if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $c.ExportFolderText.Text = $dlg.SelectedPath
                if ($c.ExportFolder.IsChecked) { & $saveSettings }
                else { $c.ExportFolder.IsChecked = $true }
            }
        })

    # Data: record count + Clear All History (acts immediately, with its own confirmation and PIN)
    $updateCount = {
        $n = 0
        if (-not [string]::IsNullOrWhiteSpace($Global:DatabaseCSV)) { $n = @($Global:DatabaseCSV | ConvertFrom-Csv).Count }
        $c.RecordCountText.Text = "$n saved record$(if ($n -ne 1) { 's' })"
        $c.BtnClearAll.IsEnabled = ($n -gt 0)
    }
    & $updateCount
    $c.BtnClearAll.Add_Click({
            Invoke-ClearAllHistory | Out-Null
            & $updateCount
        })

    $c.BtnCloseX.Add_Click({ $setWin.Close() })
    $setWin.Add_KeyDown({
            param($sender, $e)
            if ($e.Key -eq "Escape") {
                $e.Handled = $true
                $setWin.Close()
            }
        })
    $setWin.Add_Closing({
            if ($volumeSaveTimer.IsEnabled -or $saveState.Failed) { & $saveSettings }
        })
    # The window has no Windows title bar, so drag it by the navy header.
    $c.SetHeader.Add_MouseLeftButtonDown({ try { $setWin.DragMove() } catch { } })

    $setWin.ShowDialog() | Out-Null
}

$window.FindName("SettingsBtn").Add_Click({ Show-SettingsWindow })

# ------------------------------------------------------------
# DYNAMIC HARDWARE AUTO-DETECTION POLLER (USB & VIDEO)
# ------------------------------------------------------------
# Runs in the background (it used to freeze the window for a moment every 2 seconds)
$Global:HwPollWork = {
    $r = @{ VideoOK = $false; UsbOK = $false }
    try {
        $pnpMonitors = Get-CimInstance Win32_PnPEntity -Filter "PNPClass='Monitor' AND ConfigManagerErrorCode=0" -ErrorAction SilentlyContinue
        $r.VideoOK = (@($pnpMonitors).Count -gt 1)
    }
    catch { }
    try {
        # USB flash drive (USBSTOR) or USB mouse/keyboard (mouhid/kbdhid) with a VID/PID, or a removable disk
        $externalUsbDevices = Get-CimInstance Win32_PnPEntity -Filter "Service='mouhid' OR Service='kbdhid' OR Service='USBSTOR'" -ErrorAction SilentlyContinue |
        Where-Object { $_.DeviceID -match 'VID_[0-9A-F]+&PID_' }
        $usbLogical = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=2" -ErrorAction SilentlyContinue
        $r.UsbOK = [bool]($externalUsbDevices -or $usbLogical)
    }
    catch { }
    return $r
}
$Global:HwPollRunning = $false
$hwPollTimer = New-Object System.Windows.Threading.DispatcherTimer
$hwPollTimer.Interval = [TimeSpan]::FromSeconds(2)
$hwPollTimer.Add_Tick({
        if ($null -ne $Global:EditingRecord -or [string]::IsNullOrWhiteSpace($techEmailInput.Text)) { return }
        # Only fill unanswered fields, and leave explicitly selected results alone.
        if ($window.FindName("cbVideo").SelectedItem -and $window.FindName("cbUSB").SelectedItem) { return }
        if ($Global:HwPollRunning) { return }

        $Global:HwPollRunning = $true
        Start-BackgroundWork -Work $Global:HwPollWork -Context $script:InspectionGeneration -OnDone {
            param($result, $generation)
            $Global:HwPollRunning = $false
            if ($generation -ne $script:InspectionGeneration -or $null -ne $Global:EditingRecord -or $null -eq $result) { return }
            Set-DetectedPortResults $result
        }
    })
$hwPollTimer.Start()
$window.Add_Closed({
        try { $hwPollTimer.Stop(); $Global:BgJobTimer.Stop(); $Global:BgPool.Close() } catch { }
    })


# Settings > "Remember the last technician": pre-fill the email at startup
if ((Get-Setting "RememberLastTech") -and -not [string]::IsNullOrWhiteSpace([string](Get-Setting "LastTechnician"))) {
    $techEmailInput.Text = [string](Get-Setting "LastTechnician")
}

Update-Dashboard
[System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke(
    [Action] { $techEmailInput.Focus() | Out-Null },
    [System.Windows.Threading.DispatcherPriority]::Background
) | Out-Null
$window.ShowDialog() | Out-Null