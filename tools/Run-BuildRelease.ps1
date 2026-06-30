# Automation for Build-Release.vbs (handles SetupWorkbook MsgBox).
$ErrorActionPreference = 'Stop'
$repo = 'C:\Users\ringo\Documents\GitHub\GIT_VERSION_SCU Label Printing'
$target = Join-Path $repo 'MedicationDispensing.xlsm'
$bootstrap = Join-Path $repo 'Broken_PrettyPrint_MedicationDispensing.xlsm'
$bas = Join-Path $repo 'MedParser.bas'

if (-not (Test-Path $target)) {
    if (-not (Test-Path $bootstrap)) {
        throw "Missing MedicationDispensing.xlsm and bootstrap workbook."
    }
    Copy-Item $bootstrap $target -Force
}

if (-not (Test-Path $bas)) { throw "Missing required file: $bas" }

$shell = New-Object -ComObject WScript.Shell
$dismiss = $true
$timer = [System.Threading.Timer]::new({
    if (-not $script:dismiss) { return }
    foreach ($title in @('Microsoft Excel', 'Build-Release', 'Saturday Clinic - Setup Complete')) {
        if ($shell.AppActivate($title)) {
            $shell.SendKeys('{ENTER}')
        }
    }
}, $null, 8000, 2000)

$xl = $null
$wb = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $true
    $xl.DisplayAlerts = $false
    $xl.AutomationSecurity = 1

    $wb = $xl.Workbooks.Open($target)

    try {
        $null = $wb.VBProject
    } catch {
        throw "Cannot access VBA project. Enable Trust access to the VBA project object model in Excel Trust Center."
    }

    try { [void]$wb.VBProject.VBComponents.Remove($wb.VBProject.VBComponents.Item('MedParser')) } catch {}
    [void]$wb.VBProject.VBComponents.Import($bas)

    try {
        $macro = "'$($wb.Name)'!SetupWorkbook"
        $wb.Activate() | Out-Null
        $xl.Run($macro)
    } catch {
        throw "SetupWorkbook failed: $($_.Exception.Message). Open VBA (Alt+F11) and use Debug > Compile VBAProject to see compile errors."
    }

    $wb.Save()
    $wb.Close($false)
    $xl.Quit()

    $info = Get-Item $target
    Write-Output "Build succeeded."
    Write-Output "Updated: $($info.FullName)"
    Write-Output "Size: $($info.Length) bytes"
    Write-Output "Saved: $($info.LastWriteTime)"
    exit 0
}
catch {
    Write-Output "Build failed: $($_.Exception.Message)"
    if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
    if ($xl -ne $null) { try { $xl.Quit() } catch {} }
    exit 1
}
finally {
    $script:dismiss = $false
    if ($timer -ne $null) { $timer.Dispose() }
    if ($wb -ne $null) { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) }
    if ($xl -ne $null) { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
