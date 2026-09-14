# Create a repository-independent MSI compound file for Windows SIP signing parity.
param(
    [Parameter(Mandatory)][string]$OutputMsi
)

$ErrorActionPreference = "Stop"
$parent = Split-Path -Parent $OutputMsi
if ($parent) {
    New-Item -ItemType Directory -Force -Path $parent | Out-Null
}
if (Test-Path -LiteralPath $OutputMsi) {
    Remove-Item -LiteralPath $OutputMsi -Force
}

$installer = New-Object -ComObject WindowsInstaller.Installer
try {
    # msiOpenDatabaseModeCreateDirect = 3. Commit writes a valid empty MSI database; the test only
    # exercises the MSI subject interface package and does not need to install a product.
    $database = $installer.OpenDatabase($OutputMsi, 3)
    $database.Commit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($database)

    if (-not (Test-Path -LiteralPath $OutputMsi)) {
        throw "Windows Installer did not create $OutputMsi"
    }

    # Reopening read-only proves the generated fixture is accepted by Windows Installer itself.
    $database = $installer.OpenDatabase($OutputMsi, 0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($database)
}
finally {
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($installer)
}

Write-Host "Created minimal MSI parity fixture: $OutputMsi"
