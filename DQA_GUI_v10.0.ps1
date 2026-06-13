# ============================================================
#  DQA_GUI.ps1  -  Device Quality Assurance (Automated GUI)
#  DSU IT Support 
#  Author: Sandeep Pokharel
# ============================================================

# --- DATA STORAGE BEGINS HERE ---
$Global:SavedEmails = @("Sandeep.Pokharel@trojans.dsu.edu", "Sandesh.Dhakal@trojans.dsu.edu", "Tyler.Steele@trojans.dsu.edu", "Riley.Wermers@trojans.dsu.edu")
$Global:DatabaseCSV = @"
TechnicianEmail,SerialNumber,DurationHours,Charging,Screen,Touchscreen,NetworkAdapters,Keyboard,MouseTrackpad,VideoPorts,AudioOutput,Microphone,Camera,USBPorts,PalmRest,Backplate,BaseAndVents,Hinge,Notes,FinalStatus
"@
# --- DATA STORAGE ENDS HERE ---

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ------------------------------------------------------------
# DURATION TRACKER & ASYNC AUDIO SETUP (WITH CORE AUDIO API)
# ------------------------------------------------------------
$Global:sessionStartTime = Get-Date

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
    
    public static void PlayScale() {
        int[] notes = { 523, 659, 784, 1047, 784, 659, 523 };
        for(int i = 0; i < 5; i++) {
            foreach (int n in notes) { Console.Beep(n, 250); }
            System.Threading.Thread.Sleep(150);
        }
        Console.Beep(1000, 1000);
    }
    public static void PlayScaleAsync() { Task.Run(() => PlayScale()); }

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
# SELF-MODIFYING SAVE FUNCTION (Saves data inside this script)
# ------------------------------------------------------------
function Update-ScriptData {
    $scriptPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Definition }
    if (-not $scriptPath -or -not (Test-Path $scriptPath)) { 
        [System.Windows.MessageBox]::Show("Could not locate script path to save data.", "Error", "OK", "Error")
        return 
    }

    $lines = Get-Content $scriptPath
    $newLines = @()
    $inDataBlock = $false

    foreach ($line in $lines) {
        if ($line -match "^# --- DATA STORAGE BEGINS HERE ---") {
            $newLines += $line
            $cleanEmails = @($Global:SavedEmails | Select-Object -Unique | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            $emailItems  = ($cleanEmails | ForEach-Object { '"' + $_ + '"' }) -join ", "
            $newLines   += ('$Global:SavedEmails = @(' + $emailItems + ')')
            
            $newLines += "`$Global:DatabaseCSV = @`""
            $newLines += $Global:DatabaseCSV.Trim()
            $newLines += "`"@"
            $inDataBlock = $true
        }
        elseif ($line -match "^# --- DATA STORAGE ENDS HERE ---") {
            $newLines += $line
            $inDataBlock = $false
        }
        elseif (-not $inDataBlock) {
            $newLines += $line
        }
    }
    Set-Content -Path $scriptPath -Value ($newLines -join "`r`n") -Force
}

# ------------------------------------------------------------
# PIN INPUT PROMPT
# ------------------------------------------------------------
function Get-PinInput {
    $pinForm = New-Object System.Windows.Forms.Form
    $pinForm.Text = "Security Verification"
    $pinForm.Size = New-Object System.Drawing.Size(280,140)
    $pinForm.StartPosition = "CenterScreen"
    $pinForm.FormBorderStyle = "FixedDialog"
    $pinForm.MaximizeBox = $false
    $pinForm.MinimizeBox = $false
    $pinForm.TopMost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,15)
    $label.Size = New-Object System.Drawing.Size(240,20)
    $label.Text = "Enter 4-digit PIN to authorize:"

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,40)
    $textBox.Size = New-Object System.Drawing.Size(240,20)
    $textBox.PasswordChar = '*'
    $textBox.MaxLength = 4

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(90,70)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "Verify"
    $okButton.DialogResult = "OK"

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(175,70)
    $cancelButton.Size = New-Object System.Drawing.Size(75,25)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = "Cancel"

    $pinForm.Controls.Add($label)
    $pinForm.Controls.Add($textBox)
    $pinForm.Controls.Add($okButton)
    $pinForm.Controls.Add($cancelButton)
    $pinForm.AcceptButton = $okButton
    $pinForm.CancelButton = $cancelButton

    $result = $pinForm.ShowDialog()
    if ($result -eq "OK") { return $textBox.Text } else { return $null }
}

# ------------------------------------------------------------
# TITLE-CASE HELPER
# ------------------------------------------------------------
function Format-TitleCaseEmail {
    param([string]$email)
    if (-not $email -or -not $email.Contains("@")) { return $email }
    $parts     = $email -split "@", 2
    $localPart = $parts[0]
    $domain    = $parts[1]
    $words     = $localPart -split "\."
    $titled    = $words | ForEach-Object {
        if ($_.Length -gt 0) { $_.Substring(0,1).ToUpper() + $_.Substring(1).ToLower() }
        else { $_ }
    }
    return ($titled -join ".") + "@" + $domain
}

# ------------------------------------------------------------
# DELETE TECHNICIAN EMAIL DIALOG
# ------------------------------------------------------------
function Show-DeleteEmailDialog {
    $delForm = New-Object System.Windows.Forms.Form
    $delForm.Text = "Delete Technician Email"
    $delForm.Size = New-Object System.Drawing.Size(380, 200)
    $delForm.StartPosition = "CenterScreen"
    $delForm.FormBorderStyle = "FixedDialog"
    $delForm.MaximizeBox = $false
    $delForm.MinimizeBox = $false
    $delForm.TopMost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 15)
    $label.Size = New-Object System.Drawing.Size(340, 20)
    $label.Text = "Select email to permanently remove:"
    $delForm.Controls.Add($label)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10, 40)
    $listBox.Size = New-Object System.Drawing.Size(340, 80)
    foreach ($e in $Global:SavedEmails) { $listBox.Items.Add($e) | Out-Null }
    $delForm.Controls.Add($listBox)

    $deleteBtn = New-Object System.Windows.Forms.Button
    $deleteBtn.Location = New-Object System.Drawing.Point(100, 130)
    $deleteBtn.Size = New-Object System.Drawing.Size(80, 28)
    $deleteBtn.Text = "Delete"
    $deleteBtn.DialogResult = "OK"
    $delForm.Controls.Add($deleteBtn)

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Location = New-Object System.Drawing.Point(195, 130)
    $cancelBtn.Size = New-Object System.Drawing.Size(80, 28)
    $cancelBtn.Text = "Cancel"
    $cancelBtn.DialogResult = "Cancel"
    $delForm.Controls.Add($cancelBtn)

    $delForm.AcceptButton = $deleteBtn
    $delForm.CancelButton = $cancelBtn

    $result = $delForm.ShowDialog()
    if ($result -eq "OK" -and $listBox.SelectedItem) {
        return $listBox.SelectedItem
    }
    return $null
}

# ------------------------------------------------------------
# XAML INTERFACE
# ------------------------------------------------------------
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="DSU IT Support - Automated DQA System" Height="780" Width="1100" WindowStartupLocation="CenterScreen" Background="#F4F6F9" FontFamily="Segoe UI">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#FFC72C"/>
            <Setter Property="Foreground" Value="#002D62"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
        
        <Border Grid.Row="0" Background="#002D62" Padding="20,15">
            <StackPanel Orientation="Vertical" VerticalAlignment="Center">
                <TextBlock Text="Device Quality Assurance System" FontSize="26" FontWeight="Bold" Foreground="#FFC72C"/>
                <TextBlock Text="Dakota State University - IT Support" FontSize="14" Foreground="White" Margin="0,2,0,0"/>
            </StackPanel>
        </Border>

        <TabControl Grid.Row="1" Margin="10" Background="White">
            
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
                                <TextBlock Text="System Automation" FontSize="18" FontWeight="Bold" Foreground="#002D62" Margin="0,0,0,10"/>
                                
                                <TextBlock Text="Technician Email (Triggers Auto-Detect):" FontWeight="SemiBold"/>
                                <ComboBox x:Name="TechEmailInput" IsEditable="True" Height="28" Margin="0,0,0,10"/>
                                
                                <StackPanel Orientation="Horizontal" Margin="0,0,0,0">
                                    <TextBlock Text="Serial Number (Auto):" FontWeight="SemiBold"/>
                                    <TextBlock Text="(Press Enter to check duplicate)" Foreground="Gray" FontSize="11" Margin="5,2,0,0"/>
                                </StackPanel>
                                <TextBox x:Name="SerialInput" Height="28" Margin="0,0,0,10" Background="#E9ECEF"/>
                                
                                <TextBlock Text="Charging Status (Auto):" FontWeight="SemiBold"/><ComboBox x:Name="cbCharging" Height="28" Margin="0,0,0,10"/>
                                <TextBlock Text="Network Adapters (Auto):" FontWeight="SemiBold"/><ComboBox x:Name="cbNetwork" Height="28" Margin="0,0,0,15"/>

                                <TextBlock Text="Interactive Tests" FontSize="18" FontWeight="Bold" Foreground="#002D62" Margin="0,10,0,10"/>
                                
                                <Grid Margin="0,0,0,15">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    
                                    <Button x:Name="BtnTestCamera" Grid.Column="0" Grid.Row="0" Content="1. Launch Camera" Margin="0,0,5,5" Height="30"/>
                                    <Button x:Name="BtnTestAudio" Grid.Column="1" Grid.Row="0" Content="2. Test Audio/Mic" Margin="5,0,5,5" Height="30"/>
                                    <Button x:Name="BtnTestKeys" Grid.Column="2" Grid.Row="0" Content="3. Keyboard Web Test" Margin="5,0,0,5" Height="30"/>
                                    
                                    <TextBlock Text="Camera Status:" Grid.Column="0" Grid.Row="1" FontSize="12" Foreground="#666" Margin="0,0,5,2"/>
                                    <TextBlock Text="Audio/Mic Status:" Grid.Column="1" Grid.Row="1" FontSize="12" Foreground="#666" Margin="5,0,5,2"/>
                                    <TextBlock Text="Keyboard Status:" Grid.Column="2" Grid.Row="1" FontSize="12" Foreground="#666" Margin="5,0,0,2"/>

                                    <ComboBox x:Name="cbCamera" Grid.Column="0" Grid.Row="2" Height="26" Margin="0,0,5,0"/>
                                    <StackPanel Grid.Column="1" Grid.Row="2" Margin="5,0,5,0">
                                        <ComboBox x:Name="cbAudio" Height="26" Margin="0,0,0,4"/>
                                        <ComboBox x:Name="cbMic" Height="26"/>
                                    </StackPanel>
                                    <ComboBox x:Name="cbKeyboard" Grid.Column="2" Grid.Row="2" Height="26" Margin="5,0,0,0" VerticalAlignment="Top"/>
                                </Grid>
                                
                                <TextBlock x:Name="AudioStatusLabel" Text=" " FontWeight="Bold" Margin="0,0,0,10" Height="20"/>

                                <TextBlock Text="Other Hardware:" FontWeight="SemiBold" Foreground="#888" Margin="0,5,0,5"/>
                                <Grid>
                                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                    <StackPanel Grid.Column="0" Margin="0,0,5,0">
                                        <TextBlock Text="Mouse/Trackpad:"/><ComboBox x:Name="cbMouse" Height="26" Margin="0,0,0,5"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="1" Margin="5,0,0,0">
                                        <TextBlock Text="USB Ports:"/><ComboBox x:Name="cbUSB" Height="26"/>
                                    </StackPanel>
                                </Grid>
                            </StackPanel>

                            <StackPanel Grid.Column="1" Margin="15,0,0,0">
                                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                                    <TextBlock Text="Manual &amp; Cosmetic Inspection" FontSize="18" FontWeight="Bold" Foreground="#002D62" VerticalAlignment="Center"/>
                                    <Button x:Name="BtnAutoFill" Content="Auto-Fill Remaining to Operational" Height="25" Width="220" Margin="15,0,0,0" Background="#28A745" Foreground="White" FontSize="12"/>
                                </StackPanel>
                                
                                <TextBlock Text="Screen (Visual):"/><ComboBox x:Name="cbScreen" Height="26" Margin="0,0,0,5"/>
                                <TextBlock Text="Touchscreen:"/><ComboBox x:Name="cbTouch" Height="26" Margin="0,0,0,5"/>
                                <TextBlock Text="Video Ports (Ext. Monitor):"/><ComboBox x:Name="cbVideo" Height="26" Margin="0,0,0,15"/>

                                <TextBlock Text="Palm Rest:"/><ComboBox x:Name="cbPalm" Height="26" Margin="0,0,0,5"/>
                                <TextBlock Text="Backplate:"/><ComboBox x:Name="cbBackplate" Height="26" Margin="0,0,0,5"/>
                                <TextBlock Text="Base and Vents:"/><ComboBox x:Name="cbBase" Height="26" Margin="0,0,0,5"/>
                                <TextBlock Text="Hinge:"/><ComboBox x:Name="cbHinge" Height="26" Margin="0,0,0,15"/>

                                <StackPanel Orientation="Horizontal" Margin="0,0,0,2">
                                    <TextBlock Text="Additional Notes:" FontSize="16" FontWeight="Bold" Foreground="#002D62"/>
                                    <TextBlock x:Name="NotesCounter" Text="(0 chars)" Foreground="Gray" Margin="10,2,0,0" VerticalAlignment="Bottom"/>
                                </StackPanel>
                                <TextBox x:Name="NotesInput" Height="60" TextWrapping="Wrap" AcceptsReturn="True" Margin="0,0,0,15"/>

                                <Button x:Name="SaveInspectionBtn" Content="SAVE INSPECTION" Height="45" FontSize="16" Margin="0,0,0,8" IsEnabled="False" Opacity="0.5"/>
                                <Button x:Name="ResetFormBtn" Content="Reset / Clear Form" Height="32" FontSize="13" Background="#6C757D" Foreground="White" FontWeight="Bold" BorderThickness="0"/>
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
                        <Border Grid.Column="0" Background="#002D62" CornerRadius="5" Margin="5" Padding="15"><StackPanel><TextBlock x:Name="StatTotal" FontSize="28" FontWeight="Bold" Foreground="White"/><TextBlock Text="Total Runs" Foreground="#FFC72C"/></StackPanel></Border>
                        <Border Grid.Column="1" Background="#002D62" CornerRadius="5" Margin="5" Padding="15"><StackPanel><TextBlock x:Name="StatPassed" FontSize="28" FontWeight="Bold" Foreground="#28A745"/><TextBlock Text="Passed" Foreground="#FFC72C"/></StackPanel></Border>
                        <Border Grid.Column="2" Background="#002D62" CornerRadius="5" Margin="5" Padding="15"><StackPanel><TextBlock x:Name="StatFailed" FontSize="28" FontWeight="Bold" Foreground="#DC3545"/><TextBlock Text="Failed" Foreground="#FFC72C"/></StackPanel></Border>
                        <Border Grid.Column="3" Background="#002D62" CornerRadius="5" Margin="5" Padding="15"><StackPanel><TextBlock x:Name="StatAvg" FontSize="28" FontWeight="Bold" Foreground="White"/><TextBlock Text="Avg Duration (min)" Foreground="#FFC72C"/></StackPanel></Border>
                    </Grid>
                    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="5,0,0,10">
                        <Grid Width="130" Height="30" Margin="0,0,10,0">
                            <TextBox x:Name="SearchBox" Background="Transparent" VerticalContentAlignment="Center"/>
                            <TextBlock x:Name="SearchPlaceholder" Text="Search serial..." Foreground="#888" Margin="6,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False"/>
                        </Grid>
                        <ComboBox x:Name="FilterBox" Width="100" Height="30" Margin="0,0,10,0">
                            <ComboBoxItem Content="All Results" IsSelected="True"/><ComboBoxItem Content="PASSED only"/><ComboBoxItem Content="FAILED only"/>
                        </ComboBox>
                        <Button x:Name="RefreshBtn" Content="Refresh Page" Width="100" Height="30" Margin="0,0,10,0"/>
                        <Button x:Name="ExportBtn" Content="Export to CSV" Width="110" Height="30" Margin="0,0,10,0"/>
                        <Button x:Name="DeleteRowBtn" Content="Delete Selected" Width="110" Height="30" Background="#FD7E14" Foreground="White" Margin="0,0,10,0"/>
                        <Button x:Name="DeleteDbBtn" Content="Clear All" Width="90" Height="30" Background="#DC3545" Foreground="White" Margin="0,0,10,0"/>
                        <Button x:Name="DeleteEmailBtn" Content="Delete Email" Width="100" Height="30" Background="#6C757D" Foreground="White" FontWeight="Bold" BorderThickness="0"/>
                    </StackPanel>
                    <DataGrid x:Name="ResultsGrid" Grid.Row="2" AutoGenerateColumns="True" IsReadOnly="True" Margin="5" Background="White">
                        <DataGrid.RowStyle>
                            <Style TargetType="DataGridRow">
                                <Style.Triggers>
                                    <DataTrigger Binding="{Binding FinalStatus}" Value="PASSED">
                                        <Setter Property="Background" Value="#D4EDDA"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding FinalStatus}" Value="FAILED">
                                        <Setter Property="Background" Value="#F8D7DA"/>
                                    </DataTrigger>
                                </Style.Triggers>
                            </Style>
                        </DataGrid.RowStyle>
                    </DataGrid>
                </Grid>
            </TabItem>
            
        </TabControl>
    </Grid>
</Window>
"@

$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

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

$techEmailInput.ItemsSource = $Global:SavedEmails

$hwOptions = @("", "Operational", "Defective", "Not Applicable")
$cosmeticOptions = @("", "Acceptable", "Needs Repair", "Not Applicable")

$hwBoxes = @("cbCharging", "cbScreen", "cbTouch", "cbNetwork", "cbKeyboard", "cbMouse", "cbVideo", "cbAudio", "cbMic", "cbCamera", "cbUSB")
foreach ($box in $hwBoxes) { $window.FindName($box).ItemsSource = $hwOptions; $window.FindName($box).SelectedIndex = 0 }

$cosmeticBoxes = @("cbPalm", "cbBackplate", "cbBase", "cbHinge")
foreach ($box in $cosmeticBoxes) { $window.FindName($box).ItemsSource = $cosmeticOptions; $window.FindName($box).SelectedIndex = 0 }


# ------------------------------------------------------------
# FORM COMPLETION CHECKER (Disables Save until complete)
# ------------------------------------------------------------
function Check-FormCompletion {
    $allFilled = $true
    
    if ([string]::IsNullOrWhiteSpace($techEmailInput.Text) -or [string]::IsNullOrWhiteSpace($serialInput.Text)) { 
        $allFilled = $false 
    }
    
    $reqBoxes = @("cbCharging","cbScreen","cbTouch","cbNetwork","cbKeyboard","cbMouse","cbVideo","cbAudio","cbMic","cbCamera","cbUSB","cbPalm","cbBackplate","cbBase","cbHinge")
    foreach ($box in $reqBoxes) {
        $selectedVal = $window.FindName($box).SelectedItem
        if ([string]::IsNullOrWhiteSpace($selectedVal) -or $selectedVal -eq "") { 
            $allFilled = $false 
        }
    }
    
    if ($allFilled) {
        $saveInspectionBtn.IsEnabled = $true
        $saveInspectionBtn.Opacity = 1.0
    } else {
        $saveInspectionBtn.IsEnabled = $false
        $saveInspectionBtn.Opacity = 0.5
    }
}

$techEmailInput.Add_KeyUp({ Check-FormCompletion })
$serialInput.Add_TextChanged({ Check-FormCompletion })
foreach ($box in $hwBoxes) { $window.FindName($box).Add_SelectionChanged({ Check-FormCompletion }) }
foreach ($box in $cosmeticBoxes) { $window.FindName($box).Add_SelectionChanged({ Check-FormCompletion }) }


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
# QOL: AUTO-UPPERCASE SERIAL NUMBER
# ------------------------------------------------------------
$serialInput.Add_TextChanged({
    $upperText = $serialInput.Text.ToUpper()
    if ($serialInput.Text -cne $upperText) {
        $currStart = $serialInput.SelectionStart
        $serialInput.Text = $upperText
        $serialInput.SelectionStart = $currStart
    }
})

# ------------------------------------------------------------
# QOL: DUPLICATE CHECK ON "ENTER" KEY
# ------------------------------------------------------------
$serialInput.Add_KeyDown({
    param($evtSender, $e)
    if ($e.Key -eq 'Return' -or $e.Key -eq 'Enter') {
        $val = $serialInput.Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($val)) {
            $existing = @()
            if (-not [string]::IsNullOrWhiteSpace($Global:DatabaseCSV)) {
                $existing = @($Global:DatabaseCSV | ConvertFrom-Csv | Select-Object -ExpandProperty SerialNumber -ErrorAction SilentlyContinue)
            }
            if ($val -in $existing) {
                [System.Windows.MessageBox]::Show("Serial number '$val' already exists in the database.", "Duplicate Found", "OK", "Warning")
            } else {
                [System.Windows.MessageBox]::Show("Serial number '$val' is new. No duplicates found.", "Check Passed", "OK", "Information")
            }
        }
    }
})

# ------------------------------------------------------------
# QOL: NOTES CHARACTER COUNTER
# ------------------------------------------------------------
$notesInput.Add_TextChanged({
    $count = $notesInput.Text.Length
    $notesCounter.Text = "($count chars)"
    if ($count -gt 250) { 
        $notesCounter.Foreground = "#DC3545" # Red 
    } else { 
        $notesCounter.Foreground = "Gray" 
    }
})

# ------------------------------------------------------------
# QOL: SEARCH PLACEHOLDER TEXT VISIBILITY
# ------------------------------------------------------------
$searchBox.Add_TextChanged({
    if ($searchBox.Text.Length -gt 0) {
        $searchPlaceholder.Visibility = "Hidden"
    } else {
        $searchPlaceholder.Visibility = "Visible"
    }
    Update-Dashboard
})

# ------------------------------------------------------------
# 1. AUTO-DETECT (Triggers only after Email is provided)
# ------------------------------------------------------------
function Start-AutoDetect {
    if ([string]::IsNullOrWhiteSpace($techEmailInput.Text)) { return }
    if ($serialInput.Text -ne "" -and $serialInput.Text -ne "UNKNOWN") { return }

    try { $serialInput.Text = (Get-WmiObject Win32_BIOS).SerialNumber.Trim() } catch { $serialInput.Text = "UNKNOWN" }
    
    try {
        $battery = Get-WmiObject -Class Win32_Battery -ErrorAction Stop
        if ($battery -and $battery.BatteryStatus -in 2,3) { $window.FindName("cbCharging").SelectedItem = "Operational" }
        else { $window.FindName("cbCharging").SelectedItem = "Defective" }
    } catch { $window.FindName("cbCharging").SelectedItem = "Not Applicable" }

    try {
        $adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
        if ($adapters) { $window.FindName("cbNetwork").SelectedItem = "Operational" }
        else { $window.FindName("cbNetwork").SelectedItem = "Defective" }
    } catch { $window.FindName("cbNetwork").SelectedItem = "Not Applicable" }
    
    Check-FormCompletion
}

$techEmailInput.Add_LostFocus({ Start-AutoDetect })
$techEmailInput.Add_SelectionChanged({ Start-AutoDetect })

# ------------------------------------------------------------
# 2. CAMERA & KEYBOARD TESTS (Interactive Prompts)
# ------------------------------------------------------------
$window.FindName("BtnTestCamera").Add_Click({ 
    Start-Process "microsoft.windows.camera:" 
    $res = [System.Windows.MessageBox]::Show("Did the camera work properly?", "Camera Test", "YesNo", "Question")
    if ($res -eq "Yes") { $window.FindName("cbCamera").SelectedItem = "Operational" }
    elseif ($res -eq "No") { $window.FindName("cbCamera").SelectedItem = "Defective" }
})

$window.FindName("BtnTestKeys").Add_Click({ 
    Start-Process "https://keyboardchecker.com/" 
    $res = [System.Windows.MessageBox]::Show("Did all the keys on the keyboard work properly?", "Keyboard Test", "YesNo", "Question")
    if ($res -eq "Yes") { $window.FindName("cbKeyboard").SelectedItem = "Operational" }
    elseif ($res -eq "No") { $window.FindName("cbKeyboard").SelectedItem = "Defective" }
})

# ------------------------------------------------------------
# 3. ASYNC AUDIO & MIC TEST (With Direct Volume API & Prompts)
# ------------------------------------------------------------
$Global:audioTimer = New-Object System.Windows.Threading.DispatcherTimer
$Global:audioTimer.Interval = [TimeSpan]::FromSeconds(5)
$Global:audioState = 0
$Global:tempAudioPath = "$env:TEMP\DQA_MicTest.wav"
$Global:soundPlayer = New-Object System.Media.SoundPlayer

$Global:audioTimer.Add_Tick({
    $Global:audioTimer.Stop()
    
    if ($Global:audioState -eq 1) {
        $saveCmd = 'save recsound "' + $Global:tempAudioPath + '"'
        [AudioHelper]::mciSendString("stop recsound",  [IntPtr]::Zero, 0, [IntPtr]::Zero) | Out-Null
        [AudioHelper]::mciSendString($saveCmd,         [IntPtr]::Zero, 0, [IntPtr]::Zero) | Out-Null
        [AudioHelper]::mciSendString("close recsound", [IntPtr]::Zero, 0, [IntPtr]::Zero) | Out-Null

        $audioStatusLabel.Text = "[PLAYING] Audio playing back... Listen to your speakers."
        $audioStatusLabel.Foreground = "#002D62"

        if (Test-Path $Global:tempAudioPath) {
            $Global:soundPlayer.SoundLocation = $Global:tempAudioPath
            $Global:soundPlayer.Play() 
        }
        
        $Global:audioState = 2
        $Global:audioTimer.Start() 
        
    } elseif ($Global:audioState -eq 2) {
        $audioStatusLabel.Text = "[DONE] Audio test complete."
        $audioStatusLabel.Foreground = "#28A745"
        Remove-Item $Global:tempAudioPath -Force -ErrorAction SilentlyContinue
        
        $resAudio = [System.Windows.MessageBox]::Show("Did you hear the audio playing from speakers or not?", "Speaker Test", "YesNo", "Question")
        if ($resAudio -eq "Yes") { 
            $window.FindName("cbAudio").SelectedItem = "Operational"
        } elseif ($resAudio -eq "No") { 
            $window.FindName("cbAudio").SelectedItem = "Defective"
        }

        $resMic = [System.Windows.MessageBox]::Show("Did you hear audio playback sound and your voice or clap?", "Microphone Test", "YesNo", "Question")
        if ($resMic -eq "Yes") { 
            $window.FindName("cbMic").SelectedItem = "Operational"
        } elseif ($resMic -eq "No") { 
            $window.FindName("cbMic").SelectedItem = "Defective"
        }
    }
})

$window.FindName("BtnTestAudio").Add_Click({
    
    # Automatically force unmute and set system volume to exactly 50% using Core Audio API
    [AudioHelper]::SetVolume(50)

    $audioStatusLabel.Text = "[RECORDING] 5s... SPEAK or CLAP now!"
    $audioStatusLabel.Foreground = "#DC3545"
    
    [AudioHelper]::mciSendString("open new type waveaudio alias recsound", [IntPtr]::Zero, 0, [IntPtr]::Zero) | Out-Null
    [AudioHelper]::mciSendString("record recsound", [IntPtr]::Zero, 0, [IntPtr]::Zero) | Out-Null
    [AudioHelper]::PlayScaleAsync()
    
    $Global:audioState = 1
    $Global:audioTimer.Start()
})

# ------------------------------------------------------------
# AUTO-FILL REMAINING TO OPERATIONAL (Excludes Auto-Detected & Interactive Tests)
# ------------------------------------------------------------
$window.FindName("BtnAutoFill").Add_Click({
    # Excludes cbCharging and cbNetwork (which auto-detect) 
    # Excludes Cam, Audio, Mic, Keyboard (which have interactive prompts)
    $hwBoxesToFill = @("cbScreen","cbTouch","cbMouse","cbVideo","cbUSB")
    foreach ($box in $hwBoxesToFill) {
        if ([string]::IsNullOrWhiteSpace($window.FindName($box).Text)) { 
            $window.FindName($box).SelectedItem = "Operational" 
        }
    }
    
    $cosmeticBoxesToFill = @("cbPalm","cbBackplate","cbBase","cbHinge")
    foreach ($box in $cosmeticBoxesToFill) {
        if ([string]::IsNullOrWhiteSpace($window.FindName($box).Text)) { 
            $window.FindName($box).SelectedItem = "Acceptable" 
        }
    }
    Check-FormCompletion
})

# ------------------------------------------------------------
# 4. SAVE & DASHBOARD LOGIC (Modifies Script Content)
# ------------------------------------------------------------
$window.FindName("SaveInspectionBtn").Add_Click({
    
    $durationSpan = (Get-Date) - $Global:sessionStartTime
    $calcDuration = [math]::Round($durationSpan.TotalHours, 4)
    if ($calcDuration -le 0) { $calcDuration = 0.01 }

    $rawEmail = $techEmailInput.Text.Trim()
    if ($rawEmail -match "(?i)^sandeep") { $rawEmail = "Sandeep.Pokharel@trojans.dsu.edu" }
    elseif ($rawEmail -match "(?i)^sandesh") { $rawEmail = "Sandesh.Dhakal@trojans.dsu.edu" }
    $rawEmail = Format-TitleCaseEmail $rawEmail
    $techEmailInput.Text = $rawEmail

    if ([string]::IsNullOrWhiteSpace($rawEmail) -or [string]::IsNullOrWhiteSpace($serialInput.Text)) {
        [System.Windows.MessageBox]::Show("Technician Email and Serial Number are required.", "Error", "OK", "Warning")
        return
    }

    $savedSerial = $serialInput.Text

    $existingSerials = @()
    if (-not [string]::IsNullOrWhiteSpace($Global:DatabaseCSV)) {
        $existingSerials = @($Global:DatabaseCSV | ConvertFrom-Csv | Select-Object -ExpandProperty SerialNumber -ErrorAction SilentlyContinue)
    }
    if ($savedSerial -in $existingSerials) {
        $dupConfirm = [System.Windows.MessageBox]::Show(
            "Serial number '$savedSerial' already exists in the database.`n`nDo you still want to save this inspection?",
            "Duplicate Serial Number", "YesNo", "Warning"
        )
        if ($dupConfirm -ne "Yes") { return }
    }

    $R = [ordered]@{
        TechnicianEmail = $rawEmail; SerialNumber = $savedSerial; Duration = $calcDuration
        Charging = $window.FindName("cbCharging").Text; Screen = $window.FindName("cbScreen").Text
        Touchscreen = $window.FindName("cbTouch").Text; NetworkAdapters = $window.FindName("cbNetwork").Text
        Keyboard = $window.FindName("cbKeyboard").Text; MouseTrackpad = $window.FindName("cbMouse").Text
        VideoPorts = $window.FindName("cbVideo").Text; AudioOutput = $window.FindName("cbAudio").Text
        Microphone = $window.FindName("cbMic").Text; Camera = $window.FindName("cbCamera").Text
        USBPorts = $window.FindName("cbUSB").Text
        PalmRest = $window.FindName("cbPalm").Text; Backplate = $window.FindName("cbBackplate").Text
        BaseAndVents = $window.FindName("cbBase").Text; Hinge = $window.FindName("cbHinge").Text
        Notes = $notesInput.Text
    }

    $failFound = $false
    foreach ($val in $R.Values) { if ($val -eq "Defective" -or $val -eq "Needs Repair") { $failFound = $true } }
    $FinalStatusDisplay = if ($failFound) { "FAILED" } else { "PASSED" }

    $R.Add("FinalStatus", $FinalStatusDisplay)

    if ($rawEmail -notin $Global:SavedEmails) {
        $Global:SavedEmails += $rawEmail
        $techEmailInput.ItemsSource = $Global:SavedEmails
    }

    $csvLine = ConvertTo-Csv -InputObject ([PSCustomObject]$R) -NoTypeInformation | Select-Object -Last 1
    $Global:DatabaseCSV += "`r`n" + $csvLine
    
    Update-ScriptData

    [System.Windows.MessageBox]::Show("Inspection saved inside the script! Status: $FinalStatusDisplay", "Success", "OK", "Information")
    
    $lastSavedLabel.Text = "Last saved: $savedSerial at $((Get-Date).ToString('hh:mm tt'))"

    $serialInput.Text = ""; $notesInput.Text = ""; $audioStatusLabel.Text = " "
    $Global:sessionStartTime = Get-Date 
    
    foreach ($box in $hwBoxes) { $window.FindName($box).SelectedIndex = 0 }
    foreach ($box in $cosmeticBoxes) { $window.FindName($box).SelectedIndex = 0 }
    
    Check-FormCompletion
    Update-Dashboard
})

function Update-Dashboard {
    if (-not [string]::IsNullOrWhiteSpace($Global:DatabaseCSV)) {
        $data = @($Global:DatabaseCSV | ConvertFrom-Csv)
    } else { $data = @() }

    foreach ($row in $data) {
        $isFail = $false
        foreach ($prop in $row.psobject.properties) {
            if ($prop.Value -in @("Defective", "Needs Repair")) { $isFail = $true; break }
        }
        $statusText = if ($isFail) { "FAILED" } else { "PASSED" }
        $row | Add-Member -MemberType NoteProperty -Name "FinalStatus" -Value $statusText -Force
        if ($row.DurationHours) { $row.DurationHours = [double]$row.DurationHours }
    }
    [array]::Reverse($data)

    $window.FindName("StatTotal").Text = $data.Count
    $window.FindName("StatPassed").Text = @($data | Where-Object { $_.FinalStatus -eq "PASSED" }).Count
    $window.FindName("StatFailed").Text = @($data | Where-Object { $_.FinalStatus -eq "FAILED" }).Count
    if ($data.Count -gt 0) { $window.FindName("StatAvg").Text = [math]::Round((($data | Measure-Object -Property DurationHours -Average).Average * 60), 1) } 
    else { $window.FindName("StatAvg").Text = "0" }

    $search = $searchBox.Text.Trim().ToLower()
    $filter = ($window.FindName("FilterBox").SelectedItem).Content
    $filtered = @($data)

    if ($search) { $filtered = @($filtered | Where-Object { $_.SerialNumber -like "*$search*" }) }
    if ($filter -eq "PASSED only") { $filtered = @($filtered | Where-Object { $_.FinalStatus -eq "PASSED" }) }
    if ($filter -eq "FAILED only") { $filtered = @($filtered | Where-Object { $_.FinalStatus -eq "FAILED" }) }

    $resultsGrid.ItemsSource = $filtered
}

$window.FindName("RefreshBtn").Add_Click({ Update-Dashboard })

# ------------------------------------------------------------
# RESET / CLEAR FORM BUTTON
# ------------------------------------------------------------
$window.FindName("ResetFormBtn").Add_Click({
    $serialInput.Text        = ""
    $notesInput.Text         = ""
    $audioStatusLabel.Text   = " "
    $techEmailInput.Text     = ""
    $lastSavedLabel.Text     = " "
    $Global:sessionStartTime = Get-Date
    foreach ($box in $hwBoxes) { $window.FindName($box).SelectedIndex = 0 }
    foreach ($box in $cosmeticBoxes) { $window.FindName($box).SelectedIndex = 0 }
    Check-FormCompletion
})

# ------------------------------------------------------------
# DELETE TECHNICIAN EMAIL BUTTON
# ------------------------------------------------------------
$window.FindName("DeleteEmailBtn").Add_Click({
    $emailToDelete = Show-DeleteEmailDialog
    if ($emailToDelete) {
        $confirm = [System.Windows.MessageBox]::Show(
            "Permanently remove '$emailToDelete' from the dropdown?",
            "Confirm Delete", "YesNo", "Warning"
        )
        if ($confirm -eq "Yes") {
            $Global:SavedEmails = @($Global:SavedEmails | Where-Object { $_ -ne $emailToDelete })
            $techEmailInput.ItemsSource = $Global:SavedEmails
            Update-ScriptData
            [System.Windows.MessageBox]::Show("Email removed and script updated.", "Done", "OK", "Information")
        }
    }
})
$window.FindName("FilterBox").Add_SelectionChanged({ Update-Dashboard })

# ------------------------------------------------------------
# EXPORT ONLY (Creates the single output CSV on the USB)
# ------------------------------------------------------------
$window.FindName("ExportBtn").Add_Click({
    $exportData = @($resultsGrid.ItemsSource)
    if ($exportData.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No data to export.", "Export", "OK", "Information")
        return
    }
    
    $usbCheck = (Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DriveType -eq 2 } | Select-Object -First 1).DeviceID
    $exportDir = if ($usbCheck) { $usbCheck } else { "$env:USERPROFILE\Desktop" }
    
    $savePath = "$exportDir\DQA_Output_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    
    $exportData | Select-Object * -ExcludeProperty Id, RunDate, "start date", Status | Export-Csv -Path $savePath -NoTypeInformation -Force
    [System.Windows.MessageBox]::Show("Output successfully exported to:`n$savePath", "Export Complete", "OK", "Information")
})

# ------------------------------------------------------------
# DELETE SELECTED ROW (Requires PIN: 5555)
# ------------------------------------------------------------
$window.FindName("DeleteRowBtn").Add_Click({
    $selectedItem = $resultsGrid.SelectedItem
    
    if ($null -eq $selectedItem) {
        [System.Windows.MessageBox]::Show("Please select a laptop record from the dashboard to delete.", "No Selection", "OK", "Warning")
        return
    }

    $enteredPin = Get-PinInput
    
    if ($enteredPin -eq "5555") {
        $allData = [System.Collections.ArrayList]@($Global:DatabaseCSV | ConvertFrom-Csv)
        $matchIndex = -1
        
        for ($i = 0; $i -lt $allData.Count; $i++) {
            if ($allData[$i].SerialNumber -eq $selectedItem.SerialNumber -and $allData[$i].DurationHours -eq $selectedItem.DurationHours) {
                $matchIndex = $i
                break
            }
        }
        
        if ($matchIndex -ge 0) {
            $allData.RemoveAt($matchIndex)
            
            $csvHeader = "TechnicianEmail,SerialNumber,DurationHours,Charging,Screen,Touchscreen,NetworkAdapters,Keyboard,MouseTrackpad,VideoPorts,AudioOutput,Microphone,Camera,USBPorts,PalmRest,Backplate,BaseAndVents,Hinge,Notes,FinalStatus"
            
            if ($allData.Count -gt 0) {
                $newCsvData = ($allData | Select-Object -Property TechnicianEmail,SerialNumber,DurationHours,Charging,Screen,Touchscreen,NetworkAdapters,Keyboard,MouseTrackpad,VideoPorts,AudioOutput,Microphone,Camera,USBPorts,PalmRest,Backplate,BaseAndVents,Hinge,Notes,FinalStatus | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1) -join "`r`n"
                $Global:DatabaseCSV = $csvHeader + "`r`n" + $newCsvData
            } else {
                $Global:DatabaseCSV = $csvHeader
            }

            Update-ScriptData
            Update-Dashboard
            [System.Windows.MessageBox]::Show("Laptop record '$($selectedItem.SerialNumber)' has been securely deleted.", "Success", "OK", "Information")
        } else {
            [System.Windows.MessageBox]::Show("Could not locate the exact record in the database.", "Error", "OK", "Error")
        }
    }
    elseif ($null -ne $enteredPin) {
        [System.Windows.MessageBox]::Show("Incorrect PIN entered. Deletion canceled.", "Security Alert", "OK", "Error")
    }
})

# ------------------------------------------------------------
# CLEAR DASHBOARD (Requires PIN: 5555)
# ------------------------------------------------------------
$window.FindName("DeleteDbBtn").Add_Click({
    $enteredPin = Get-PinInput
    
    if ($enteredPin -eq "5555") {
        $Global:DatabaseCSV = "TechnicianEmail,SerialNumber,DurationHours,Charging,Screen,Touchscreen,NetworkAdapters,Keyboard,MouseTrackpad,VideoPorts,AudioOutput,Microphone,Camera,USBPorts,PalmRest,Backplate,BaseAndVents,Hinge,Notes,FinalStatus"
        Update-ScriptData
        Update-Dashboard
        [System.Windows.MessageBox]::Show("Laptop history securely cleared from script memory.", "Success", "OK", "Information")
    }
    elseif ($null -ne $enteredPin) {
        [System.Windows.MessageBox]::Show("Incorrect PIN entered. Deletion canceled.", "Security Alert", "OK", "Error")
    }
})

Update-Dashboard
$window.ShowDialog() | Out-Null
