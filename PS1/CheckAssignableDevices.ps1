# These variables are device properties. For details, see pciprop.h in the Windows Driver Kit headers.
$devpkey_PciDevice_DeviceType                              = "{3AB22E31-8264-4b4e-9AF5-A8D2D8E33E62}  1"
$devpkey_PciDevice_BaseClass                               = "{3AB22E31-8264-4b4e-9AF5-A8D2D8E33E62}  3"
$devpkey_PciDevice_RequiresReservedMemoryRegion            = "{3AB22E31-8264-4b4e-9AF5-A8D2D8E33E62}  34"
$devpkey_PciDevice_AcsCompatibleUpHierarchy                = "{3AB22E31-8264-4b4e-9AF5-A8D2D8E33E62}  31"

# DeviceType constants
$devprop_PciDevice_DeviceType_PciConventional              = 0
$devprop_PciDevice_DeviceType_PciX                         = 1
$devprop_PciDevice_DeviceType_PciExpressEndpoint           = 2
$devprop_PciDevice_DeviceType_PciExpressLegacyEndpoint     = 3
$devprop_PciDevice_DeviceType_PciExpressRootComplexIntegratedEndpoint = 4
$devprop_PciDevice_DeviceType_PciExpressTreatedAsPci       = 5

# BridgeType constants
$devprop_PciDevice_BridgeType_PciConventional              = 6
$devprop_PciDevice_BridgeType_PciX                         = 7
$devprop_PciDevice_BridgeType_PciExpressRootPort           = 8
$devprop_PciDevice_BridgeType_PciExpressUpstreamSwitchPort = 9
$devprop_PciDevice_BridgeType_PciExpressDownstreamSwitchPort = 10
$devprop_PciDevice_BridgeType_PciExpressToPciXBridge       = 11
$devprop_PciDevice_BridgeType_PciXToExpressBridge          = 12
$devprop_PciDevice_BridgeType_PciExpressTreatedAsPci       = 13
$devprop_PciDevice_BridgeType_PciExpressEventCollector     = 14

# ACS hierarchy compatibility
$devprop_PciDevice_AcsCompatibleUpHierarchy_NotSupported    = 0
$devprop_PciDevice_AcsCompatibleUpHierarchy_SingleFunctionSupported = 1
$devprop_PciDevice_AcsCompatibleUpHierarchy_NoP2PSupported  = 2
$devprop_PciDevice_AcsCompatibleUpHierarchy_Supported       = 3

# BaseClass values from WDK wdm.h
$devprop_PciDevice_BaseClass_DisplayCtlr                   = 3

Write-Host "Executing SurveyDDA.ps1, revision 1"
Write-Host "Generating a list of PCI Express endpoint devices"

# Get only present PCI devices
$pnpdevs = Get-PnpDevice -PresentOnly
$pcidevs = $pnpdevs | Where-Object { $_.InstanceId -like "PCI*" }

foreach ($pcidev in $pcidevs) {

    # Check for reserved memory region requirement
    $rmrr = ($pcidev | Get-PnpDeviceProperty $devpkey_PciDevice_RequiresReservedMemoryRegion).Data
    if ($rmrr -ne 0) {
        Write-Host -ForegroundColor Red -BackgroundColor Black "BIOS requires that this device remain attached to BIOS-owned memory. Not assignable."
        continue
    }

    # Check ACS compatibility up hierarchy
    $acsUp = ($pcidev | Get-PnpDeviceProperty $devpkey_PciDevice_AcsCompatibleUpHierarchy).Data
    if ($acsUp -eq $devprop_PciDevice_AcsCompatibleUpHierarchy_NotSupported) {
        Write-Host -ForegroundColor Red -BackgroundColor Black "Traffic from this device may be redirected to other devices in the system. Not assignable."
        continue
    }

    # Determine device type
    $devtype = ($pcidev | Get-PnpDeviceProperty $devpkey_PciDevice_DeviceType).Data
    switch ($devtype) {
        $devprop_PciDevice_DeviceType_PciExpressEndpoint {
            Write-Host "Express Endpoint -- more secure."
        }
        $devprop_PciDevice_DeviceType_PciExpressRootComplexIntegratedEndpoint {
            Write-Host "Embedded Endpoint -- less secure."
        }
        $devprop_PciDevice_DeviceType_PciExpressLegacyEndpoint {
            $devBaseClass = ($pcidev | Get-PnpDeviceProperty $devpkey_PciDevice_BaseClass).Data
            if ($devBaseClass -eq $devprop_PciDevice_BaseClass_DisplayCtlr) {
                Write-Host "Legacy Express Endpoint -- graphics controller."
            }
            else {
                Write-Host -ForegroundColor Red -BackgroundColor Black "Legacy, non-VGA PCI device. Not assignable."
                continue
            }
        }
        $devprop_PciDevice_DeviceType_PciExpressTreatedAsPci {
            Write-Host -ForegroundColor Red -BackgroundColor Black "BIOS kept control of PCI Express for this device. Not assignable."
            continue
        }
        default {
            Write-Host -ForegroundColor Red -BackgroundColor Black "Old-style PCI device, switch port, etc. Not assignable."
            continue
        }
    }

    # Get the stable location path
    $locationpath = ($pcidev | Get-PnpDeviceProperty DEVPKEY_Device_LocationPaths).Data[0]

    # Warn if device is disabled
    if ($pcidev.ConfigManagerErrorCode -eq "CM_PROB_DISABLED") {
        Write-Host -ForegroundColor Yellow -BackgroundColor Black "Device is Disabled, unable to check resource requirements. It may be assignable."
        Write-Host -ForegroundColor Yellow -BackgroundColor Black "Enable the device and rerun this script to confirm."
        Write-Host $locationpath
        continue
    }

    # Check IRQ assignments
    $escapedId    = "*" + $pcidev.PNPDeviceID.Replace("\", "\\") + "*"
    $irqAssignments = Get-WmiObject -Query "SELECT * FROM Win32_PnPAllocatedResource" |
                      Where-Object { $_.__RELPATH -like "*Win32_IRQResource*" -and $_.Dependent -like $escapedId }

    if ($irqAssignments.Count -eq 0) {
        Write-Host -ForegroundColor Green -BackgroundColor Black "    And it has no interrupts at all -- assignment can work."
    }
    else {
        $msiAssignments = $irqAssignments | Where-Object { $_.Antecedent -like "*IRQNumber=42949*" }
        if ($msiAssignments.Count -eq 0) {
            Write-Host -ForegroundColor Red -BackgroundColor Black "All of the interrupts are line-based, no assignment can work."
            continue
        }
        else {
            Write-Host -ForegroundColor Green -BackgroundColor Black "    And its interrupts are message-based, assignment can work."
        }
    }

    # Calculate MMIO requirements
    $mmioAssignments = Get-WmiObject -Query "SELECT * FROM Win32_PnPAllocatedResource" |
                       Where-Object { $_.__RELPATH -like "*Win32_DeviceMemoryAddress*" -and $_.Dependent -like $escapedId }

    $mmioTotal = 0
    foreach ($mem in $mmioAssignments) {
        $baseAdd = ($mem.Antecedent -split '"')[1]
        $range   = Get-WmiObject -Query "SELECT * FROM Win32_DeviceMemoryAddress WHERE StartingAddress = $baseAdd"
        $mmioTotal += ($range.EndingAddress - $range.StartingAddress)
    }

    if ($mmioTotal -eq 0) {
        Write-Host -ForegroundColor Green -BackgroundColor Black "    And it has no MMIO space"
        $mmioMB = 0
    }
    else {
        $mmioMB = [math]::Ceiling($mmioTotal / 1MB)
        Write-Host -ForegroundColor Green -BackgroundColor Black "    And it requires at least: $mmioMB MB of MMIO gap space"
    }

    # Output device friendly name, location, and MMIO info
    "$($pcidev.FriendlyName)`nLocation Path: $locationpath`nMMIO Space: $mmioMB MiB`n"
}

# Check host SR-IOV (Discrete Device Assignment) support
if (-not (Get-VMHost).IovSupport) {
    Write-Host ""
    Write-Host "Unfortunately, this machine doesn't support using them in a VM."
    Write-Host ""
    (Get-VMHost).IovSupportReasons
}
