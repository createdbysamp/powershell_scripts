# List of critical DLLs for Postman
$dlls = @(
    "kernel32.dll","ntdll.dll","user32.dll","gdi32.dll","advapi32.dll","shell32.dll","ole32.dll",
    "ws2_32.dll","rpcrt4.dll","d3d11.dll","dxgi.dll","ucrtbase.dll","msvcp140.dll","vcruntime140.dll"
)

$system32 = "$env:SystemRoot\System32"
$results = @()

$results = foreach ($dll in $dlls) {
    $path = Join-Path $system32 $dll
    if (Test-Path $path) {
        $sig = Get-AuthenticodeSignature $path
        [PSCustomObject]@{
            DLL    = $dll
            Path   = $path
            Status = if ($sig.Status -eq "Valid") { "OK" } else { "Signature Issue" }
        }
    } else {
        [PSCustomObject]@{
            DLL    = $dll
            Path   = "Not Found"
            Status = "Missing"
        }
    }
}

$results | Format-Table -AutoSize


# # Define Postman installation path
# $postmanPath = "$env:LOCALAPPDATA\Postman"

# # List of Electron-related DLLs typically required
# $electronDlls = @(
#     "ffmpeg.dll",
#     "libEGL.dll",
#     "libGLESv2.dll",
#     "vk_swiftshader.dll",
#     "d3dcompiler_47.dll"
# )

# Write-Host "Checking Electron DLLs in $postmanPath`n"

# foreach ($dll in $electronDlls) {
#     $dllPath = Join-Path $postmanPath $dll
#     if (Test-Path $dllPath) {
#         Write-Host "$dll : Found" -ForegroundColor Green
#     } else {
#         Write-Host "$dll : Missing" -ForegroundColor Red
#     }
# }

# # Optional: Check if Postman executable exists
# $postmanExe = Join-Path $postmanPath "Postman.exe"
# if (Test-Path $postmanExe) {
#     Write-Host "`nPostman executable found at $postmanExe" -ForegroundColor Cyan
# } else {
#     Write-Host "`nPostman executable NOT found!" -ForegroundColor Yellow
# }


# Validate TEMP and TMP paths
$envVars = @("TEMP", "TMP")

foreach ($var in $envVars) {
    $path = [System.Environment]::GetEnvironmentVariable($var, "User")
    if (-not $path) {
        Write-Host "$var is NOT set for User" -ForegroundColor Red
    } elseif (Test-Path $path) {
        try {
            $testFile = Join-Path $path "perm_test.txt"
            New-Item -Path $testFile -ItemType File -Force | Out-Null
            Remove-Item $testFile -Force
            Write-Host "$var path ($path) is valid and writable" -ForegroundColor Green
        } catch {
            Write-Host "$var path ($path) exists but is NOT writable" -ForegroundColor Yellow
        }
    } else {
        Write-Host "$var path ($path) does NOT exist" -ForegroundColor Red
    }
}


# Validate write permissions on LocalAppData
$localAppData = $env:LOCALAPPDATA
Write-Host "Checking write permissions for $localAppData"

try {
    $testFile = Join-Path $localAppData "perm_test.txt"
    New-Item -Path $testFile -ItemType File -Force | Out-Null
    Remove-Item $testFile -Force
    Write-Host "Write permissions OK for $localAppData" -ForegroundColor Green
} catch {
    Write-Host "Cannot write to $localAppData. Check permissions or run installer as Admin." -ForegroundColor Red
}
