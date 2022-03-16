$player = New-Object System.Media.SoundPlayer "G:\activate-windows\lib\audio\dreamtown.wav"
$player.PlayLooping()

# Check Windows Edition
$sku = (Get-WmiObject Win32_OperatingSystem).OperatingSystemSKU

if (($sku -ne 4) -and ($sku -ne 1)) {
    Write-Host "[ERROR]: This Edition of Windows does not support Hyper-V. Please use Pro or Enterprise editions of Windows." -ForegroundColor Red
    return
}

# Check if Hyper-V is enabled and enable it if it is not
if ((Get-WindowsOptionalFeature -FeatureName Microsoft-Hyper-V-All -Online).State -eq "Enabled") {
    Write-Host "[OK]: Hyper-V is enabled..." -ForegroundColor Green
} else {
    Write-Host "[OK]: Hyper-V will be enabled..." -ForegroundColor Green
    Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V -All
    Restart-Computer
}

# Configure Virtual Switches
$switch = (Get-VMSwitch -Name "Vlmcsd Switch" -ErrorAction SilentlyContinue -ErrorVariable SwitchError)

if ($SwitchError) {
    Write-Host "[OK]: Creating virtual switch..." -ForegroundColor Green
    New-VMSwitch -SwitchType Internal -Name "Vlmcsd Switch" | Out-Null
    New-NetIPAddress -IPAddress 10.100.4.1 -PrefixLength 24 -InterfaceAlias “vEthernet (Vlmcsd Switch)” | Out-Null
    New-NetNat -Name "Vlmcsd Network" -InternalIPInterfaceAddressPrefix 10.100.4.0/24 -ErrorAction SilentlyContinue | Out-Null
} else {
    Write-Host "[OK]: Virtual switch found..." -ForegroundColor Green
}

# Check if Virtual Machine was already imported
Get-VM -Name "debian-vlmcsd" -ErrorAction SilentlyContinue -ErrorVariable VirtualMachineError | Out-Null

if ($VirtualMachineError) {
    # Import Virtual Machine
    Write-Host "[OK]: Importing virtual machine..." -ForegroundColor Green
    Import-VM -Path "$PSScriptRoot\debian-vlmcsd\Virtual Machines\59089D84-7221-43B7-8611-C460BC0A690C.vmcx" -Copy -GenerateNewId -ErrorAction SilentlyContinue -ErrorVariable VirtualMachineImportError | Out-Null

    if ($VirtualMachineImportError) {
        Write-Host "[ERROR]: Error importing virtual machine..." -ForegroundColor Red
        Write-Host "[ERROR]: $VirtualMachineImportError" -ForegroundColor Red
        Return
    }

} else {
    Write-Host "[OK]: Virtual machine found..." -ForegroundColor Green
}

# Start Virtual Machine
if ((Get-VM -Name "debian-vlmcsd").state -ne "Running") {
    Write-Host "[OK]: Starting virtual machine..." -ForegroundColor Green
    Start-VM -Name "debian-vlmcsd"
    Start-Sleep -Seconds 15
}


# Set KMS IP address
if (Test-Connection -Count 1 (Get-VMNetworkAdapter -VMName "debian-vlmcsd").IPAddresses[0] -Quiet) {
    Write-Host "[OK]: Setting KMS IP..." -ForegroundColor Green
    cscript //B "slmgr.vbs" /skms 10.100.4.2
}

# Set Generic Volume License Key
Write-Host "[OK]: Setting Windows Enterprise Generic Volume License Key..." -ForegroundColor Green
cscript //B "slmgr.vbs" /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43

# Activate Windows
Write-Host "[OK]: Activating Windows..." -ForegroundColor Green
cscript //B "slmgr.vbs" /ato


$player.Stop()

#Write-Host "[SUCCESS]: Press any key to close..." -ForegroundColor Green
#Read-Host