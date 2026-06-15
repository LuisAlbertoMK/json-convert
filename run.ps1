<#
.SYNOPSIS
  Analiza y limpia el proyecto json-convert con un solo comando.
.DESCRIPTION
  Fase 1 - ANALISIS: verifica Python, dependencias, tests, git status, tamano.
  Fase 2 - LIMPIEZA: elimina __pycache__, *.log, *.pyc, temporales.
  Fase 3 - REPORTE: resumen con scores y espacio liberado.
#>

param(
    [switch]$SkipTests,
    [switch]$SkipCleanup
)

# ── Config ──
$ROOT = Split-Path -Parent $MyInvocation.MyCommand.Path
$GREEN = "Green"; $RED = "Red"; $YELLOW = "Yellow"
$CYAN = "Cyan"; $MAGENTA = "Magenta"

# ── Helpers ──
function Write-Step($Title, $Color = $CYAN) {
    Write-Host "`n===========================================" -ForegroundColor $Color
    Write-Host "  $Title" -ForegroundColor $Color
    Write-Host "===========================================" -ForegroundColor $Color
}
function Write-Ok($Msg) { Write-Host "  [OK] $Msg" -ForegroundColor $GREEN }
function Write-Warn($Msg) { Write-Host "  [!] $Msg" -ForegroundColor $YELLOW }
function Write-Err($Msg) { Write-Host "  [FAIL] $Msg" -ForegroundColor $RED }
function Write-Info($Msg) { Write-Host "       $Msg" -ForegroundColor $MAGENTA }

$exitCode = 0
$errors = @(); $warnings = @()
$totalBefore = 0; $totalAfter = 0

# ═══════════════════════════════════════════
# FASE 1: ANALISIS
# ═══════════════════════════════════════════
Write-Step "FASE 1: ANALISIS DEL PROYECTO"

# 1.1 - Python
Write-Host "`n[1/5] Python" -ForegroundColor $YELLOW
try {
    $pyVer = python --version 2>&1
    if ($LASTEXITCODE -eq 0 -and $pyVer -match 'Python (\d+)\.(\d+)') {
        $major = [int]$Matches[1]; $minor = [int]$Matches[2]
        if ($major -ge 3 -and $minor -ge 9) {
            Write-Ok "Python $major.$minor (OK)"
        } else {
            Write-Warn "Python $major.$minor requiere 3.9+"
            $warnings += "Python version"
        }
    } else {
        Write-Err "No se pudo detectar version de Python"
        $errors += "Python no encontrado"
    }
} catch {
    Write-Err "Python no encontrado"
    $errors += "Python no encontrado"
}

# 1.2 - Dependencias
Write-Host "`n[2/5] Dependencias" -ForegroundColor $YELLOW
$depsOk = $true
try {
    $v = python -c "import openpyxl; print(openpyxl.__version__)" 2>$null
    if ($LASTEXITCODE -eq 0) { Write-Ok "openpyxl $v" } else { throw }
} catch {
    Write-Warn "openpyxl no instalado ejecuta install.bat primero"
    $depsOk = $false; $warnings += "openpyxl no instalado"
}
try {
    $v = python -c "import playwright; print(playwright.__version__)" 2>$null
    if ($LASTEXITCODE -eq 0) { Write-Ok "playwright $v" } else { throw }
} catch {
    Write-Warn "playwright no instalado ejecuta install.bat primero"
    $depsOk = $false; $warnings += "playwright no instalado"
}

# 1.3 - Tests
Write-Host "`n[3/5] Tests" -ForegroundColor $YELLOW
if (-not $SkipTests -and $depsOk) {
    $testSuites = @(
        @{File="test_parse.py"; Name="Unitarios (parse)"}
        @{File="test_gen_urls.py"; Name="Unitarios (gen_urls)"}
        @{File="test_extract_aa.py"; Name="Integracion (extract_aa)"}
        @{File="test_integration.py"; Name="Integracion (browser)"}
    )
    $allPassed = $true; $totalTests = 0; $passedTests = 0; $failedTests = 0

    foreach ($suite in $testSuites) {
        $path = Join-Path $ROOT $suite.File
        if (-not (Test-Path $path)) {
            Write-Warn "  No se encuentra $($suite.File) saltando"
            continue
        }
        Write-Host "  --> $($suite.Name): " -NoNewline
        $result = python -m pytest "$path" -q --tb=no 2>&1
        $ok = $LASTEXITCODE -eq 0

        $summary = "?"; $summaryFail = "0"
        if ($result -match '(\d+) passed') { $summary = $Matches[1] }
        if ($result -match '(\d+) failed') { $summaryFail = $Matches[1] }

        if ($ok) {
            Write-Host "OK ($summary passed)" -ForegroundColor $GREEN
            $passedTests += [int]$summary
        } else {
            Write-Host "FAIL ($summaryFail failed)" -ForegroundColor $RED
            $allPassed = $false; $failedTests += [int]$summaryFail
            $result | Select-String "FAILED" | ForEach-Object { Write-Info $_.Line.Trim() }
        }
        $totalTests += [int]$summary + [int]$summaryFail
    }

    if ($allPassed) { Write-Ok "Todos los tests pasaron ($totalTests tests)" }
    else { Write-Err "$failedTests tests fallaron de $totalTests"; $errors += "Tests fallidos: $failedTests" }
} elseif (-not $depsOk) { Write-Warn "Tests saltados por dependencias faltantes" }
else { Write-Info "Tests saltados (--SkipTests)" }

# 1.4 - Git status
Write-Host "`n[4/5] Git status" -ForegroundColor $YELLOW
$gitDir = Join-Path $ROOT ".git"
if (Test-Path $gitDir) {
    Push-Location $ROOT
    $status = git status --short 2>&1
    $branch = git branch --show-current 2>&1

    $ahead = git rev-list --count '@{u}'..HEAD 2>&1
    if ($LASTEXITCODE -ne 0) { $ahead = 0 }
    $behind = git rev-list --count HEAD..'@{u}' 2>&1
    if ($LASTEXITCODE -ne 0) { $behind = 0 }

    Write-Host "  Rama: " -NoNewline; Write-Host $branch -ForegroundColor $CYAN
    if ([string]::IsNullOrWhiteSpace($status)) {
        Write-Ok "Working tree limpio"
    } else {
        $stLines = ($status | Measure-Object -Line).Lines
        Write-Warn "$stLines archivos modificados/sin seguimiento"
        $status | ForEach-Object { Write-Info $_ }
    }
    if ($ahead -gt 0) { Write-Info "$ahead commits por pushear" }
    if ($behind -gt 0) { Write-Warn "$behind commits detras del remoto" }
    Pop-Location
} else {
    Write-Info "No es un repositorio git"
}

# 1.5 - Tamano del proyecto
Write-Host "`n[5/5] Tamano del proyecto" -ForegroundColor $YELLOW
function Get-DirSize($Path) {
    $sum = 0
    Get-ChildItem -LiteralPath $Path -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object { $sum += $_.Length }
    return $sum
}
$totalBefore = Get-DirSize $ROOT
Write-Host "  Peso total: " -NoNewline
Write-Host "$([math]::Round($totalBefore / 1MB, 2)) MB" -ForegroundColor $CYAN

# ═══════════════════════════════════════════
# FASE 2: LIMPIEZA
# ═══════════════════════════════════════════
if (-not $SkipCleanup) {
    Write-Step "FASE 2: LIMPIEZA"
    $cleanedBytes = 0; $cleanedFiles = 0

    # NOTA: usamos Where-Object con Extension para filtrar, NO -Include con -LiteralPath
    # (PS 5.1 ignora -Include cuando -LiteralPath no tiene wildcard y devuelve TODO)

    # 2.1 - __pycache__
    Write-Host "`n[1/4] __pycache__" -ForegroundColor $YELLOW
    $cacheDirs = Get-ChildItem -LiteralPath $ROOT -Directory -Filter "__pycache__" -Recurse -ErrorAction SilentlyContinue
    if ($cacheDirs.Count -gt 0) {
        foreach ($d in $cacheDirs) {
            $size = Get-DirSize $d.FullName
            Remove-Item -LiteralPath $d.FullName -Recurse -Force -ErrorAction SilentlyContinue
            if (-not (Test-Path $d.FullName)) {
                Write-Info "  Eliminado: $($d.FullName.Replace($ROOT,'')) ($([math]::Round($size/1KB,1)) KB)"
                $cleanedBytes += $size; $cleanedFiles++
            }
        }
    } else { Write-Ok "Nada que limpiar" }

    # 2.2 - *.pyc / *.pyo (usando Where-Object + Extension, seguro)
    Write-Host "`n[2/4] Archivos .pyc / .pyo" -ForegroundColor $YELLOW
    $compiled = Get-ChildItem -LiteralPath $ROOT -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Extension -match '^\.py[co]$' }
    $count = 0
    foreach ($f in $compiled) {
        $cleanedBytes += $f.Length
        Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue; $count++
    }
    if ($count -gt 0) { Write-Info "Eliminados $count archivos compilados"; $cleanedFiles += $count }
    else { Write-Ok "Nada que limpiar" }

    # 2.3 - *.log
    Write-Host "`n[3/4] Archivos .log" -ForegroundColor $YELLOW
    $logs = Get-ChildItem -LiteralPath $ROOT -Filter "*.log" -ErrorAction SilentlyContinue
    $count = 0
    foreach ($f in $logs) {
        $cleanedBytes += $f.Length
        Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue; $count++
    }
    if ($count -gt 0) { Write-Info "Eliminados $count archivos de log"; $cleanedFiles += $count }
    else { Write-Ok "Nada que limpiar" }

    # 2.4 - .history/
    Write-Host "`n[4/4] .history/" -ForegroundColor $YELLOW
    $histPath = Join-Path $ROOT ".history"
    if (Test-Path $histPath) {
        $size = Get-DirSize $histPath
        Remove-Item -LiteralPath $histPath -Recurse -Force -ErrorAction SilentlyContinue
        if (-not (Test-Path $histPath)) {
            Write-Info "Eliminado .history/ ($([math]::Round($size/1KB,1)) KB)"
            $cleanedBytes += $size; $cleanedFiles++
        }
    } else { Write-Ok "Nada que limpiar" }

    $totalAfter = $totalBefore - $cleanedBytes
    Write-Host "`n  Espacio recuperado: " -NoNewline
    if ($cleanedBytes -gt 1MB) { Write-Host "$([math]::Round($cleanedBytes/1MB,2)) MB" -ForegroundColor $GREEN }
    elseif ($cleanedBytes -gt 0) { Write-Host "$([math]::Round($cleanedBytes/1KB,1)) KB" -ForegroundColor $GREEN }
    else { Write-Host "0 bytes" -ForegroundColor $GREEN }
    Write-Host "  Archivos eliminados: $cleanedFiles" -ForegroundColor $GREEN
} else {
    $totalAfter = $totalBefore
}

# ═══════════════════════════════════════════
# FASE 3: REPORTE FINAL
# ═══════════════════════════════════════════
Write-Step "FASE 3: REPORTE FINAL" $GREEN

if ($errors.Count -eq 0 -and $warnings.Count -eq 0) {
    Write-Host "  Todo OK proyecto saludable" -ForegroundColor $GREEN
} else {
    if ($errors.Count -gt 0) {
        Write-Host "`n  ERRORES ($($errors.Count)):" -ForegroundColor $RED
        $errors | ForEach-Object { Write-Host "    * $_" -ForegroundColor $RED }
        $exitCode = 1
    }
    if ($warnings.Count -gt 0) {
        Write-Host "`n  ADVERTENCIAS ($($warnings.Count)):" -ForegroundColor $YELLOW
        $warnings | ForEach-Object { Write-Host "    * $_" -ForegroundColor $YELLOW }
    }
}

# Mostrar peso final solo si cambio
if ($totalAfter -ne $totalBefore) {
    Write-Host "`n  Peso final del proyecto: $([math]::Round($totalAfter / 1MB, 2)) MB" -ForegroundColor $CYAN
}
Write-Host "`n  Comandos utiles:" -ForegroundColor $MAGENTA
Write-Host "    .\run.ps1 -SkipTests       solo limpieza" -ForegroundColor $MAGENTA
Write-Host "    .\run.ps1 -SkipCleanup     solo analisis" -ForegroundColor $MAGENTA

Write-Host "`n===========================================" -ForegroundColor $GREEN
if ($exitCode -eq 0) { Write-Host "  FINALIZADO proyecto OK" -ForegroundColor $GREEN }
else { Write-Host "  FINALIZADO CON ERRORES revisa arriba" -ForegroundColor $RED }
Write-Host "===========================================" -ForegroundColor $GREEN

exit $exitCode
