# ABBYY FineReader 16 Auto OCR Script
# Tự động nhận dạng PDF và xuất ra file TXT

#Requires -Version 5.1

# ===== CẤU HÌNH =====
$Config = @{
    ABBYYPath = "C:\Program Files (x86)\ABBYY FineReader 16"
    InputFolder = "C:\OCR_Input"
    OutputFolder = "C:\OCR_Output"
    Languages = @("Vietnamese", "English")
    OutputFormat = "TXT"
    Encoding = "UTF-8"
}

# ===== HÀM HELPER =====
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
        Write-ColorOutput "✓ Đã tạo thư mục input: $InputPath" "Success"
    }
    
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        Write-ColorOutput "✓ Đã tạo thư mục output: $OutputPath" "Success"
    }
}

function Test-ABBYYInstalled {
    param([string]$Path)
    
    $enginePath = Join-Path $Path "FineReaderEngine.exe"
    
    if (Test-Path $enginePath) {
        return $true
    }
    
    # Tìm kiếm thay thế
    $possiblePaths = @(
        "C:\Program Files\ABBYY FineReader 16",
        "C:\Program Files (x86)\ABBYY FineReader 16",
        "C:\Program Files\ABBYY\FineReader 16"
    )
    
    foreach ($p in $possiblePaths) {
        $testPath = Join-Path $p "FineReaderEngine.exe"
        if (Test-Path $testPath) {
            $Config.ABBYYPath = $p
            return $true
        }
    }
    
    return $false
}

# ===== PHƯƠNG PHÁP 1: SỬ DỤNG COM INTERFACE =====
function Start-OCRWithCOM {
    param(
        [string]$InputFolder,
        [string]$OutputFolder
    )
    
    Write-ColorOutput "`n===== BẮT ĐẦU OCR BẰNG COM INTERFACE =====" "Info"
    
    try {
        # Khởi tạo ABBYY Engine
        Write-ColorOutput "Đang khởi tạo ABBYY Engine..." "Info"
        $Engine = New-Object -ComObject "FineReader.Engine"
        
        if ($null -eq $Engine) {
            throw "Không thể khởi tạo ABBYY Engine"
        }
        
        Write-ColorOutput "✓ ABBYY Engine đã sẵn sàng" "Success"
        
        # Thiết lập tham số xử lý
        $ProcessingParams = $Engine.CreateProcessingParams()
        $ProcessingParams.SetPredefinedTextDocumentProcessingParams()
        
        # Cấu hình ngôn ngữ
        foreach ($lang in $Config.Languages) {
            $ProcessingParams.Recognition.RecognitionParams.TextLanguage.AddLanguage($lang)
        }
        
        # Cấu hình output
        $ProcessingParams.OutputFormat.TextExportParams.Encoding = $Config.Encoding
        
        # Lấy danh sách file PDF
        $pdfFiles = Get-ChildItem -Path $InputFolder -Filter "*.pdf"
        
        if ($pdfFiles.Count -eq 0) {
            Write-ColorOutput "Không tìm thấy file PDF nào trong thư mục input" "Warning"
            return
        }
        
        Write-ColorOutput "`nTìm thấy $($pdfFiles.Count) file PDF" "Info"
        $processedCount = 0
        $errorCount = 0
        
        # Xử lý từng file
        foreach ($file in $pdfFiles) {
            $outputFile = Join-Path $OutputFolder "$($file.BaseName).txt"
            
            Write-ColorOutput "`n[$($processedCount + 1)/$($pdfFiles.Count)] Đang xử lý: $($file.Name)" "Info"
            
            try {
                # Tạo FRDocument
                $Document = $Engine.CreateFRDocument()
                $Document.AddImageFile($file.FullName, $null)
                
                # Nhận dạng
                Write-Host "  → Đang nhận dạng văn bản..." -NoNewline
                $Document.Process($ProcessingParams)
                Write-Host " ✓" -ForegroundColor Green
                
                # Export
                Write-Host "  → Đang xuất file TXT..." -NoNewline
                $Document.Export($outputFile, "TextExport", $null)
                Write-Host " ✓" -ForegroundColor Green
                
                $Document.Close()
                
                Write-ColorOutput "  ✓ Hoàn thành: $($file.BaseName).txt" "Success"
                $processedCount++
            }
            catch {
                Write-ColorOutput "  ✗ Lỗi: $_" "Error"
                $errorCount++
            }
        }
        
        # Dọn dẹp
        $Engine.DeinitializeEngine()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Engine) | Out-Null
        
        # Tổng kết
        Write-ColorOutput "`n===== KẾT QUẢ =====" "Info"
        Write-ColorOutput "Thành công: $processedCount file" "Success"
        if ($errorCount -gt 0) {
            Write-ColorOutput "Lỗi: $errorCount file" "Error"
        }
    }
    catch {
        Write-ColorOutput "Lỗi COM: $_" "Error"
        Write-ColorOutput "Có thể ABBYY chưa được cài đặt hoặc license không hợp lệ" "Warning"
    }
}

# ===== PHƯƠNG PHÁP 2: SỬ DỤNG COMMAND LINE =====
function Start-OCRWithCLI {
    param(
        [string]$InputFolder,
        [string]$OutputFolder,
        [string]$ABBYYPath
    )
    
    Write-ColorOutput "`n===== BẮT ĐẦU OCR BẰNG COMMAND LINE =====" "Info"
    
    $engineExe = Join-Path $ABBYYPath "FineReaderEngine.exe"
    
    if (-not (Test-Path $engineExe)) {
        Write-ColorOutput "Không tìm thấy FineReaderEngine.exe tại: $engineExe" "Error"
        return
    }
    
    # Lấy danh sách file PDF
    $pdfFiles = Get-ChildItem -Path $InputFolder -Filter "*.pdf"
    
    if ($pdfFiles.Count -eq 0) {
        Write-ColorOutput "Không tìm thấy file PDF nào trong thư mục input" "Warning"
        return
    }
    
    Write-ColorOutput "Tìm thấy $($pdfFiles.Count) file PDF`n" "Info"
    $processedCount = 0
    
    foreach ($file in $pdfFiles) {
        $outputFile = Join-Path $OutputFolder "$($file.BaseName).txt"
        
        Write-ColorOutput "[$($processedCount + 1)/$($pdfFiles.Count)] Đang xử lý: $($file.Name)" "Info"
        
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
                Write-ColorOutput "  ✓ Hoàn thành: $($file.BaseName).txt" "Success"
                $processedCount++
            } else {
                Write-ColorOutput "  ✗ Lỗi với exit code: $($process.ExitCode)" "Error"
            }
        }
        catch {
            Write-ColorOutput "  ✗ Lỗi: $_" "Error"
        }
    }
    
    Write-ColorOutput "`n===== HOÀN THÀNH: $processedCount/$($pdfFiles.Count) file =====" "Success"
}

# ===== MAIN SCRIPT =====
function Main {
    Clear-Host
    Write-ColorOutput @"
╔═══════════════════════════════════════════╗
║  ABBYY FINEREADER 16 AUTO OCR TOOL       ║
║  PowerShell Automation Script             ║
╚═══════════════════════════════════════════╝
"@ "Info"
    
    # Kiểm tra và tạo thư mục
    Initialize-Folders -InputPath $Config.InputFolder -OutputPath $Config.OutputFolder
    
    Write-ColorOutput "`nCấu hình:" "Info"
    Write-Host "  Input:  $($Config.InputFolder)"
    Write-Host "  Output: $($Config.OutputFolder)"
    Write-Host "  Ngôn ngữ: $($Config.Languages -join ', ')"
    
    # Kiểm tra ABBYY
    if (-not (Test-ABBYYInstalled -Path $Config.ABBYYPath)) {
        Write-ColorOutput "`nKhông tìm thấy ABBYY FineReader 16!" "Error"
        Write-ColorOutput "Vui lòng cài đặt hoặc kiểm tra đường dẫn" "Warning"
        Read-Host "`nNhấn Enter để thoát"
        return
    }
    
    Write-ColorOutput "`n✓ ABBYY FineReader 16 đã được tìm thấy" "Success"
    
    # Chọn phương pháp
    Write-ColorOutput "`nChọn phương pháp OCR:" "Info"
    Write-Host "  1. COM Interface (Khuyến nghị - Ổn định)"
    Write-Host "  2. Command Line (Nhanh hơn)"
    Write-Host "  3. Thoát"
    
    $choice = Read-Host "`nNhập lựa chọn (1-3)"
    
    switch ($choice) {
        "1" {
            Start-OCRWithCOM -InputFolder $Config.InputFolder -OutputFolder $Config.OutputFolder
        }
        "2" {
            Start-OCRWithCLI -InputFolder $Config.InputFolder -OutputFolder $Config.OutputFolder -ABBYYPath $Config.ABBYYPath
        }
        "3" {
            Write-ColorOutput "Thoát chương trình" "Info"
            return
        }
        default {
            Write-ColorOutput "Lựa chọn không hợp lệ!" "Error"
        }
    }
    
    Write-ColorOutput "`nKết quả đã được lưu tại: $($Config.OutputFolder)" "Success"
    Read-Host "`nNhấn Enter để thoát"
}

# Chạy script
Main