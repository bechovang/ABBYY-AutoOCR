# ABBYY FineReader 16 Auto OCR Script
# Tu dong nhan dang PDF va xuat ra file TXT

#Requires -Version 5.1

# ===== CAU HINH =====
# Lay duong dan thu muc hien tai cua script
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

$Config = @{
    ABBYYPath = "C:\Program Files\ABBYY FineReader 16"
    ABBYYExe = "FineReader.exe"
    InputFolder = Join-Path $ScriptPath "OCR_Input"
    OutputFolder = Join-Path $ScriptPath "OCR_Output"
    Languages = @("Vietnamese", "English")
    OutputFormat = "TXT"
    Encoding = "UTF-8"
}

# ===== HAM HELPER =====
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )
    
    $colors = @{
        "Success" = "Green"
        "Error" = "Red"
        "Warning" = "Yellow"
        "Info" = "Cyan"
    }
    
    Write-Host $Message -ForegroundColor $colors[$Type]
}

function Initialize-Folders {
    param(
        [string]$InputPath,
        [string]$OutputPath
    )
    
    if (-not (Test-Path $InputPath)) {
        New-Item -ItemType Directory -Path $InputPath -Force | Out-Null
        Write-ColorOutput "Da tao thu muc input: $InputPath" "Success"
    }
    
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        Write-ColorOutput "Da tao thu muc output: $OutputPath" "Success"
    }
}

function Test-ABBYYInstalled {
    param([string]$Path)
    
    # Thu tim FineReader.exe hoac FineReaderEngine.exe
    $possibleExes = @("FineReader.exe", "FineReaderEngine.exe")
    
    foreach ($exe in $possibleExes) {
        $exePath = Join-Path $Path $exe
        if (Test-Path $exePath) {
            $Config.ABBYYExe = $exe
            return $true
        }
    }
    
    # Tim kiem thay the
    $possiblePaths = @(
        "C:\Program Files\ABBYY FineReader 16",
        "C:\Program Files (x86)\ABBYY FineReader 16",
        "C:\Program Files\ABBYY\FineReader 16",
        "C:\Program Files\ABBYY FineReader PDF 16",
        "C:\Program Files (x86)\ABBYY FineReader PDF 16"
    )
    
    foreach ($p in $possiblePaths) {
        foreach ($exe in $possibleExes) {
            $testPath = Join-Path $p $exe
            if (Test-Path $testPath) {
                $Config.ABBYYPath = $p
                $Config.ABBYYExe = $exe
                return $true
            }
        }
    }
    
    return $false
}

# ===== PHUONG PHAP 1: SU DUNG COM INTERFACE =====
function Start-OCRWithCOM {
    param(
        [string]$InputFolder,
        [string]$OutputFolder
    )
    
    Write-ColorOutput "`n===== BAT DAU OCR BANG COM INTERFACE =====" "Info"
    
    try {
        # Khoi tao ABBYY Engine
        Write-ColorOutput "Dang khoi tao ABBYY Engine..." "Info"
        $Engine = New-Object -ComObject "FineReader.Engine"
        
        if ($null -eq $Engine) {
            throw "Khong the khoi tao ABBYY Engine"
        }
        
        Write-ColorOutput "ABBYY Engine da san sang" "Success"
        
        # Thiet lap tham so xu ly
        $ProcessingParams = $Engine.CreateProcessingParams()
        $ProcessingParams.SetPredefinedTextDocumentProcessingParams()
        
        # Cau hinh ngon ngu
        foreach ($lang in $Config.Languages) {
            $ProcessingParams.Recognition.RecognitionParams.TextLanguage.AddLanguage($lang)
        }
        
        # Cau hinh output
        $ProcessingParams.OutputFormat.TextExportParams.Encoding = $Config.Encoding
        
        # Lay danh sach file PDF
        $pdfFiles = Get-ChildItem -Path $InputFolder -Filter "*.pdf"
        
        if ($pdfFiles.Count -eq 0) {
            Write-ColorOutput "Khong tim thay file PDF nao trong thu muc input" "Warning"
            return
        }
        
        Write-ColorOutput "`nTim thay $($pdfFiles.Count) file PDF" "Info"
        $processedCount = 0
        $errorCount = 0
        
        # Xu ly tung file
        foreach ($file in $pdfFiles) {
            $outputFile = Join-Path $OutputFolder "$($file.BaseName).txt"
            
            Write-ColorOutput "`n[$($processedCount + 1)/$($pdfFiles.Count)] Dang xu ly: $($file.Name)" "Info"
            
            try {
                # Tao FRDocument
                $Document = $Engine.CreateFRDocument()
                $Document.AddImageFile($file.FullName, $null)
                
                # Nhan dang
                Write-Host "  -> Dang nhan dang van ban..." -NoNewline
                $Document.Process($ProcessingParams)
                Write-Host " OK" -ForegroundColor Green
                
                # Export
                Write-Host "  -> Dang xuat file TXT..." -NoNewline
                $Document.Export($outputFile, "TextExport", $null)
                Write-Host " OK" -ForegroundColor Green
                
                $Document.Close()
                
                Write-ColorOutput "  Hoan thanh: $($file.BaseName).txt" "Success"
                $processedCount++
            }
            catch {
                Write-ColorOutput "  Loi: $_" "Error"
                $errorCount++
            }
        }
        
        # Don dep
        $Engine.DeinitializeEngine()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Engine) | Out-Null
        
        # Tong ket
        Write-ColorOutput "`n===== KET QUA =====" "Info"
        Write-ColorOutput "Thanh cong: $processedCount file" "Success"
        if ($errorCount -gt 0) {
            Write-ColorOutput "Loi: $errorCount file" "Error"
        }
    }
    catch {
        Write-ColorOutput "Loi COM: $_" "Error"
        Write-ColorOutput "Co the ABBYY chua duoc cai dat hoac license khong hop le" "Warning"
    }
}

# ===== PHUONG PHAP 2: SU DUNG COMMAND LINE =====
function Start-OCRWithCLI {
    param(
        [string]$InputFolder,
        [string]$OutputFolder,
        [string]$ABBYYPath
    )
    
    Write-ColorOutput "`n===== BAT DAU OCR BANG COMMAND LINE =====" "Info"
    
    $engineExe = Join-Path $ABBYYPath $Config.ABBYYExe
    
    if (-not (Test-Path $engineExe)) {
        Write-ColorOutput "Khong tim thay $($Config.ABBYYExe) tai: $engineExe" "Error"
        
        # Thu khoi dong ABBYY GUI de nguoi dung tu chay
        Write-ColorOutput "`nThu khoi dong ABBYY GUI..." "Warning"
        try {
            Start-Process $engineExe
            Write-ColorOutput "Da mo ABBYY FineReader. Vui long su dung Hot Folder hoac xu ly thu cong." "Info"
        } catch {
            Write-ColorOutput "Khong the khoi dong ABBYY: $_" "Error"
        }
        return
    }
    
    # Lay danh sach file PDF
    $pdfFiles = Get-ChildItem -Path $InputFolder -Filter "*.pdf"
    
    if ($pdfFiles.Count -eq 0) {
        Write-ColorOutput "Khong tim thay file PDF nao trong thu muc input" "Warning"
        return
    }
    
    Write-ColorOutput "Tim thay $($pdfFiles.Count) file PDF`n" "Info"
    $processedCount = 0
    
    foreach ($file in $pdfFiles) {
        $outputFile = Join-Path $OutputFolder "$($file.BaseName).txt"
        
        Write-ColorOutput "[$($processedCount + 1)/$($pdfFiles.Count)] Dang xu ly: $($file.Name)" "Info"
        
        $arguments = @(
            "/if `"$($file.FullName)`""
            "/of `"$outputFile`""
            "/tet UTF8"
            "/tel Vietnamese,English"
            "/quit"
        )
        
        try {
            $process = Start-Process -FilePath $engineExe -ArgumentList ($arguments -join " ") -Wait -PassThru -NoNewWindow
            
            if ($process.ExitCode -eq 0) {
                Write-ColorOutput "  Hoan thanh: $($file.BaseName).txt" "Success"
                $processedCount++
            } else {
                Write-ColorOutput "  Loi voi exit code: $($process.ExitCode)" "Error"
            }
        }
        catch {
            Write-ColorOutput "  Loi: $_" "Error"
        }
    }
    
    Write-ColorOutput "`n===== HOAN THANH: $processedCount/$($pdfFiles.Count) file =====" "Success"
}

# ===== MAIN SCRIPT =====
function Main {
    Clear-Host
    Write-ColorOutput "================================================" "Info"
    Write-ColorOutput "  ABBYY FINEREADER 16 AUTO OCR TOOL" "Info"
    Write-ColorOutput "  PowerShell Automation Script" "Info"
    Write-ColorOutput "================================================" "Info"
    
    # Kiem tra va tao thu muc
    Initialize-Folders -InputPath $Config.InputFolder -OutputPath $Config.OutputFolder
    
    Write-ColorOutput "`nCau hinh:" "Info"
    Write-Host "  Input:  $($Config.InputFolder)"
    Write-Host "  Output: $($Config.OutputFolder)"
    Write-Host "  Ngon ngu: $($Config.Languages -join ', ')"
    
    # Kiem tra ABBYY
    if (-not (Test-ABBYYInstalled -Path $Config.ABBYYPath)) {
        Write-ColorOutput "`nKhong tim thay ABBYY FineReader 16!" "Error"
        Write-ColorOutput "Vui long cai dat hoac kiem tra duong dan" "Warning"
        Read-Host "`nNhan Enter de thoat"
        return
    }
    
    Write-ColorOutput "`nABBYY FineReader 16 da duoc tim thay" "Success"
    Write-Host "  Duong dan: $($Config.ABBYYPath)"
    Write-Host "  File exe: $($Config.ABBYYExe)"
    
    # Chon phuong phap
    Write-ColorOutput "`nChon phuong phap OCR:" "Info"
    Write-Host "  1. COM Interface (Khuyen nghi - On dinh)"
    Write-Host "  2. Command Line (Nhanh hon)"
    Write-Host "  3. Thoat"
    
    $choice = Read-Host "`nNhap lua chon (1-3)"
    
    switch ($choice) {
        "1" {
            Start-OCRWithCOM -InputFolder $Config.InputFolder -OutputFolder $Config.OutputFolder
        }
        "2" {
            Start-OCRWithCLI -InputFolder $Config.InputFolder -OutputFolder $Config.OutputFolder -ABBYYPath $Config.ABBYYPath
        }
        "3" {
            Write-ColorOutput "Thoat chuong trinh" "Info"
            return
        }
        default {
            Write-ColorOutput "Lua chon khong hop le!" "Error"
        }
    }
    
    Write-ColorOutput "`nKet qua da duoc luu tai: $($Config.OutputFolder)" "Success"
    Read-Host "`nNhan Enter de thoat"
}

# Chay script
Main