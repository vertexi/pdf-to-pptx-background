param(
    [string]$Python = "python"
)

$ErrorActionPreference = "Stop"
Push-Location $PSScriptRoot

try {
    $pythonPrefix = (& $Python -c "import sys; print(sys.prefix)").Trim()
    if ($LASTEXITCODE -ne 0) {
        throw "Unable to determine the Python installation directory."
    }

    $pyInstallerArgs = @(
        "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name", "PDF-to-PPTX-Background",
        "--collect-all", "pymupdf",
        "--collect-data", "pptx"
    )

    # Conda keeps these dependencies outside Python's DLLs directory.
    $condaBin = Join-Path $pythonPrefix "Library\bin"
    foreach ($dllName in @("tcl86t.dll", "tk86t.dll", "liblzma.dll", "libbz2.dll")) {
        $dllPath = Join-Path $condaBin $dllName
        if (Test-Path -LiteralPath $dllPath) {
            $pyInstallerArgs += @("--add-binary", "$dllPath;.")
        }
    }

    $pyInstallerArgs += ".\pdf_to_pptx_background_gui.py"
    & $Python @pyInstallerArgs

    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller exited with code $LASTEXITCODE."
    }

    Write-Host ""
    Write-Host "Created: $PSScriptRoot\dist\PDF-to-PPTX-Background.exe"
}
finally {
    Pop-Location
}
