# .Net methods for hiding/showing the console in the background
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

function Show-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()

    # Hide = 0,
    # ShowNormal = 1,
    # ShowMinimized = 2,
    # ShowMaximized = 3,
    # Maximize = 3,
    # ShowNormalNoActivate = 4,
    # Show = 5,
    # Minimize = 6,
    # ShowMinNoActivate = 7,
    # ShowNoActivate = 8,
    # Restore = 9,
    # ShowDefault = 10,
    # ForceMinimized = 11

    [Console.Window]::ShowWindow($consolePtr, 4)
}

function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}

$height = 495
$width = 900
$consolewidth = 300

$player = New-Object System.Media.SoundPlayer "G:\activate-windows\lib\audio\dreamtown.wav"
$player.PlayLooping()

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Windows 10 Enterprise KeyGen'

$form.Size = New-Object System.Drawing.Size($width,$height)
$form.MaximumSize = New-Object System.Drawing.Size($width,$height)
$form.MinimumSize = New-Object System.Drawing.Size(($consolewidth+40),$height)

$form.StartPosition = 'CenterScreen'

$iconBase64      = 'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAIAAAD8GO2jAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjAuMTnU1rJkAAAB50lEQVRIS7WWzytEURTHZ2FhaWFhYWFhYWFhaWFh6c+wsGCapJBJU0hRSrOgLBVSmkQoSpqyUJISapIFJU0i1KQp1PG9826vO9+Z97Pr22dz3pxzv/PO/fUSkvxfOLYOx9bh2DocBzPZKlPtku2VuS7JtMlwIydUw7EnGO50Rd4ehPRdlru87GWUMZVU4LgOqLza0cP56PdHNga4NthgsUc+i3qIQOH9qDzAYLZTyiVd7Ah/8/lGztclNyLbY5JfksKxeuiKRvAzwOy93OsyR+9Pam6HajLHm5UfXhTQT34GqDH1eCGjTZxjkmqQmQ5+6Gdgth5NQLsoIRwca7AoTZ1k1UM0p7Y/QXCsof4sdOvRrRlgeZgyu49eY/0cTDO76ShzcJnTQ0O0NvA2XioWqjIrcKy5PdQ1kLl90CKsVC9F2GiYVVPuiQaDiRb1TzGWw9eHzoEiGGwO6hpHWFRe04vuu4pggCPIFPwowSWmAXa/qdKr5zaOaQBwyps6W+UEh/gGtasFlrjCKC2+ATia15WucHpf76vna/2ylVLntnmeRzbApsUQ4RXZwAFnAC7eMMLNSrWhDEC6RfUac1D3+sRew87HhVzvC4PjYLBecRxhDsByn/qEoYRqOLYOx9bh2DocWyaZ+APgBBKhVfsHwAAAAABJRU5ErkJggg=='
$iconBytes       = [Convert]::FromBase64String($iconBase64)
$stream          = [System.IO.MemoryStream]::new($iconBytes, 0, $iconBytes.Length)

$form.Icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))

$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(12,420)
$button.Size = New-Object System.Drawing.Size($consolewidth,23)
$button.ForeColor = "White"
$button.BackColor = "Black"
$button.FlatStyle = 'Flat'
$button.Text = 'Activate'

$form.Controls.Add($button)

$button.Add_Click(
    {
        # Check Windows Edition

        $sku = (Get-WmiObject Win32_OperatingSystem).OperatingSystemSKU

        if (($sku -ne 4) -and ($sku -ne 1)) {
            $textBox.AppendText("[ERROR]: This Edition of Windows does not support Hyper-V. Please use Pro or Enterprise editions of Windows.`r`n")
            return
        }

        # Check if Hyper-V is enabled and enable it if it is not
        if ((Get-WindowsOptionalFeature -FeatureName Microsoft-Hyper-V-All -Online).State -eq "Enabled") {
            $textBox.AppendText("[OK]: Hyper-V is enabled...`r`n")
        } else {
            $textBox.AppendText("[OK]: Hyper-V will be enabled...`r`n")
            Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V -All
            Restart-Computer
        }

        # Configure Virtual Switches
        $switch = (Get-VMSwitch -Name "Vlmcsd Switch" -ErrorAction SilentlyContinue -ErrorVariable SwitchError)

        if ($SwitchError) {
            $textBox.AppendText("[OK]: Creating virtual switch...`r`n")
            New-VMSwitch -SwitchType Internal -Name "Vlmcsd Switch" | Out-Null
            New-NetIPAddress -IPAddress 10.100.4.1 -PrefixLength 24 -InterfaceAlias “vEthernet (Vlmcsd Switch)” | Out-Null
            New-NetNat -Name "Vlmcsd Network" -InternalIPInterfaceAddressPrefix 10.100.4.0/24 -ErrorAction SilentlyContinue | Out-Null
        } else {
            $textBox.AppendText("[OK]: Virtual switch found...`r`n")
        }

        # Check if Virtual Machine was already imported
        Get-VM -Name "debian-vlmcsd" -ErrorAction SilentlyContinue -ErrorVariable VirtualMachineError | Out-Null

        if ($VirtualMachineError) {
            # Import Virtual Machine
            $textBox.AppendText("[OK]: Importing virtual machine...`r`n")
            Import-VM -Path "$PSScriptRoot\debian-vlmcsd\Virtual Machines\59089D84-7221-43B7-8611-C460BC0A690C.vmcx" -Copy -GenerateNewId -ErrorAction SilentlyContinue -ErrorVariable VirtualMachineImportError | Out-Null

            if ($VirtualMachineImportError) {
                $textBox.AppendText("[ERROR]: Error importing virtual machine...`r`n")
                $textBox.AppendText("[ERROR]: $VirtualMachineImportError`r`n")
                Return
            }

        } else {
            $textBox.AppendText("[OK]: Virtual machine found...`r`n")
        }

        # Start Virtual Machine
        if ((Get-VM -Name "debian-vlmcsd").state -ne "Running") {
            $textBox.AppendText("[OK]: Starting virtual machine...`r`n")
            Start-VM -Name "debian-vlmcsd"
            Start-Sleep -Seconds 15
        }


        # Set KMS IP address
        if (Test-Connection -Count 1 (Get-VMNetworkAdapter -VMName "debian-vlmcsd").IPAddresses[0] -Quiet) {
            $textBox.AppendText("[OK]: Setting KMS IP...`r`n")
            cscript //B "slmgr.vbs" /skms 10.100.4.2
        }

        # Set Generic Volume License Key
        $textBox.AppendText("[OK]: Setting Windows Enterprise Generic Volume License Key...`r`n")
        cscript //B "slmgr.vbs" /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43

        # Activate Windows
        $textBox.AppendText("[OK]: Activating Windows...`r`n")
        #cscript //B "slmgr.vbs" /ato
        $textBox.AppendText("[OK]: Windows has been activated!`r`n")

    }
)

$textBox = New-Object System.Windows.Forms.RichTextBox
$textBox.Location = New-Object System.Drawing.Point(12,12)
$textBox.Size = New-Object System.Drawing.Size($consolewidth,400)
$textBox.WordWrap = $true
$textBox.Multiline = $true
$textBox.ForeColor = "Green"
$textBox.BackColor = "Black"
#$textBox.BorderStyle = 'None'
$textBox.ReadOnly = $true

$form.Controls.Add($textBox)
$form.Topmost = $true
$form.Add_Shown({
    $textBox.Select()
    Hide-Console
})

$file = (Get-Item 'G:\activate-windows\lib\img\highwayman3.jpg')
$img = [System.Drawing.Image]::Fromfile($file);

[System.Windows.Forms.Application]::EnableVisualStyles();

$picture = new-object Windows.Forms.PictureBox
$picture.Location = New-Object System.Drawing.Size(0,1)
$picture.Size = New-Object System.Drawing.Size($img.Width,$img.Height)
$picture.Image = $img

$form.Controls.Add($picture)

$result = $form.ShowDialog()

$stream.Dispose()
$form.Dispose()
$player.Stop()