param(
    [string]$Entry = "pdf_tool/app.py",
    [string]$Name = "PDFTool"
)

$ErrorActionPreference = "Stop"

# Ensure pyinstaller is available
if (-not (Get-Command pyinstaller -ErrorAction SilentlyContinue)) {
    Write-Host "Installing pyinstaller..."
    pip install pyinstaller | Out-Host
}

# Clean old builds
if (Test-Path build) { Remove-Item build -Recurse -Force }
if (Test-Path dist) { Remove-Item dist -Recurse -Force }
if (Test-Path "$Name.spec") { Remove-Item "$Name.spec" -Force }

# Build
pyinstaller --noconfirm --clean --name $Name --onefile --windowed `
    --hidden-import pdf2docx `
    --hidden-import pdf2docx.converter `
    --hidden-import pdf2docx.parser `
    --hidden-import fitz `
    --hidden-import PIL `
    --hidden-import PIL.Image `
    --hidden-import pypdf `
    --hidden-import pypdf.generic `
    --hidden-import pypdf.reader `
    --hidden-import pypdf.writer `
    $Entry | Out-Host

Write-Host "Build finished. See dist/$Name/ or dist/$Name.exe"



