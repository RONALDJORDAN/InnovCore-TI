# ==============================================================================
# INNOVCORE TI - AUTO-SCANNER E CAPTURA AUTOMATICA DE MAC WI-FI DE TABLETS (ADB)
# Desenvolvido por Jordan | InnovTech & InnovCore TI
# Plug & Play: Conectou o tablet via USB -> Detecta -> Verifica MAC -> Captura com Retries
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::InputEncoding = [System.Text.Encoding]::UTF8 } catch {}
try { $Host.UI.RawUI.WindowTitle = "InnovCore TI - Auto-Scanner de MAC Wi-Fi Android [ADB Plug & Play]" } catch {}

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$rootDir = if ($env:INNOVCORE_USB_DIR -and (Test-Path $env:INNOVCORE_USB_DIR)) { 
    $env:INNOVCORE_USB_DIR.TrimEnd('\') 
} else { 
    Split-Path $scriptDir -Parent 
}
$pastaDadosMobile = Join-Path $rootDir "dados\mobile"
if (-not (Test-Path $pastaDadosMobile)) {
    New-Item -ItemType Directory -Path $pastaDadosMobile -Force | Out-Null
}
$csvSaida = Join-Path $pastaDadosMobile "relatorio_macs_tablets.csv"

function Exibir-BannerAdb {
    Write-Host " +======================================================================+" -ForegroundColor Cyan
    Write-Host " |  ___ _   _ _   _  _____     ______ _____ ____  _____   _____ ___     |" -ForegroundColor Yellow
    Write-Host " | |_ _| \ | | \ | |/ _ \ \   / / ___/ _ \|  _ \| ____| |_   _|_ _|    |" -ForegroundColor Yellow
    Write-Host " |  | ||  \| |  \| | | | \ \ / / |  | | | | |_) |  _|     | |  | |     |" -ForegroundColor Yellow
    Write-Host " |  | || |\  | |\  | |_| |\ V /| |__| |_| |  _ <| |___    | |  | |     |" -ForegroundColor Yellow
    Write-Host " | |___|_| \_|_| \_|\___/  \_/  \____\___/|_| \_\_____|   |_| |___|    |" -ForegroundColor Yellow
    Write-Host " |                                                                      |" -ForegroundColor Cyan
    Write-Host " |   AUTO-SCANNER DE TABLETS ANDROID - CAPTURA DE MAC WI-FI EM 1 SEGUNDO|" -ForegroundColor Green
    Write-Host " |              Desenvolvido por Jordan | Edicao Corporativa             |" -ForegroundColor White
    Write-Host " +======================================================================+" -ForegroundColor Cyan
}

function Localizar-Adb {
    $possiveis = @(
        (Join-Path $scriptDir "platform-tools\adb.exe"),
        (Join-Path $scriptDir "..\platform-tools\adb.exe"),
        (Join-Path $scriptDir "..\..\Inno_RPA\tools\platform-tools\adb.exe"),
        (Join-Path $scriptDir "..\tools\platform-tools\adb.exe"),
        (Join-Path $scriptDir "tools\platform-tools\adb.exe"),
        "C:\Users\Administrator\Pictures\projetos\Projetos Innov\Inno_RPA\tools\platform-tools\adb.exe",
        "$env:USERPROFILE\AppData\Local\Android\Sdk\platform-tools\adb.exe",
        "C:\adb\adb.exe"
    )
    foreach ($p in $possiveis) {
        if (Test-Path $p) {
            return (Resolve-Path $p).Path
        }
    }
    $cmd = Get-Command adb.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

function Testar-MacValido {
    param([string]$m)
    if ([string]::IsNullOrWhiteSpace($m)) { return $false }
    $m = $m.Trim().ToUpper()
    if ($m -notmatch '^([0-9A-F]{2}[:-]){5}([0-9A-F]{2})$') { return $false }
    if ($m -eq "00:00:00:00:00:00" -or $m -eq "FF:FF:FF:FF:FF:FF" -or $m -eq "02:00:00:00:00:00") {
        return $false
    }
    return $true
}

function Normalizar-Mac {
    param([string]$m)
    if ([string]::IsNullOrWhiteSpace($m)) { return "N/A" }
    return ($m.Trim().ToUpper() -replace '-', ':')
}

function Carregar-CacheMacsExistentes {
    param([string[]]$arquivosCsv)
    $cache = @{}
    foreach ($csv in $arquivosCsv) {
        if (Test-Path $csv) {
            try {
                $linhas = Get-Content $csv -Encoding UTF8 -ErrorAction SilentlyContinue
                if (-not $linhas -or $linhas.Count -eq 0) { continue }
                
                $nomeBase = [System.IO.Path]::GetFileName($csv)
                $header = $linhas[0] -split ';'
                $idxSerial = -1
                $idxMac = -1
                $idxMarca = -1
                $idxModelo = -1
                $idxPat = -1
                
                for ($i = 0; $i -lt $header.Count; $i++) {
                    $h = $header[$i].Trim()
                    if ($h -match '^(?i)(?:numero_serie|num_serie|serial|serial_no|ns)$') { $idxSerial = $i }
                    elseif ($h -match '^(?i)(?:mac_rede|mac_wifi|mac|endereco_mac)$') { $idxMac = $i }
                    elseif ($h -match '^(?i)(?:marca|fabricante)$') { $idxMarca = $i }
                    elseif ($h -match '^(?i)modelo$') { $idxModelo = $i }
                    elseif ($h -match '^(?i)patrimonio$') { $idxPat = $i }
                }

                if ($idxSerial -ge 0 -and $idxMac -ge 0) {
                    for ($k = 1; $k -lt $linhas.Count; $k++) {
                        $linha = $linhas[$k]
                        if ([string]::IsNullOrWhiteSpace($linha)) { continue }
                        $cols = $linha -split ';'
                        if ($cols.Count -gt [Math]::Max($idxSerial, $idxMac)) {
                            $s = $cols[$idxSerial].Trim()
                            $m = $cols[$idxMac].Trim()
                            if ($s -and (Testar-MacValido $m)) {
                                $obj = [PSCustomObject]@{
                                    Serial     = $s
                                    MAC        = Normalizar-Mac $m
                                    Marca      = if ($idxMarca -ge 0 -and $cols.Count -gt $idxMarca) { $cols[$idxMarca].Trim() } else { "N/A" }
                                    Modelo     = if ($idxModelo -ge 0 -and $cols.Count -gt $idxModelo) { $cols[$idxModelo].Trim() } else { "N/A" }
                                    Patrimonio = if ($idxPat -ge 0 -and $cols.Count -gt $idxPat) { $cols[$idxPat].Trim() } else { "N/A" }
                                    Origem     = $nomeBase
                                }
                                $cache[$s.ToUpper()] = $obj
                                $cache[$s] = $obj
                            }
                        }
                    }
                } else {
                    foreach ($linha in $linhas) {
                        if ([string]::IsNullOrWhiteSpace($linha)) { continue }
                        $parts = $linha -split ';'
                        if ($parts.Count -ge 2) {
                            $serialCand = $null
                            $macCand = $null
                            foreach ($p in $parts) {
                                $val = $p.Trim()
                                if (Testar-MacValido $val) {
                                    $macCand = Normalizar-Mac $val
                                } elseif ($val -match '^[a-zA-Z0-9_-]{6,30}$' -and $val -notmatch '(?i)serial|data|patrimonio|innovtech') {
                                    if (-not $serialCand) { $serialCand = $val }
                                }
                            }
                            if ($serialCand -and $macCand) {
                                $obj = [PSCustomObject]@{
                                    Serial     = $serialCand
                                    MAC        = $macCand
                                    Marca      = "N/A"
                                    Modelo     = "N/A"
                                    Patrimonio = "N/A"
                                    Origem     = $nomeBase
                                }
                                $cache[$serialCand.ToUpper()] = $obj
                                $cache[$serialCand] = $obj
                            }
                        }
                    }
                }
            } catch {}
        }
    }
    return $cache
}

function Obter-MacComRetries {
    param(
        [string]$adbPath,
        [string]$serial,
        [int]$maxRetries = 6,
        [scriptblock]$onProgress = $null
    )

    # 0. Garante que o rádio Wi-Fi está ENERGIZADO
    # Quando o Wi-Fi está desligado, o Android desliga a energia do chip físico (power gating).
    # Isso faz a interface wlan0 sumir e o Linux responder apenas com interfaces virtuais (dummy0).
    # Ligando o rádio preventivamente via ADB, a antena é energizada e o MAC real de fábrica é exposto.
    try {
        & $adbPath -s $serial shell "cmd wifi set-wifi-enabled enabled 2>/dev/null || svc wifi enable 2>/dev/null" 2>$null
        Start-Sleep -Milliseconds 600
    } catch {}

    $mac = "N/A"
    for ($tentativa = 1; $tentativa -le $maxRetries; $tentativa++) {
        if ($onProgress) {
            & $onProgress "Tentativa $tentativa de $maxRetries..."
        }

        # 1. Estratégia A: dumpsys wifi (mFactoryMacAddress ou mWifiInfo MAC de hardware)
        try {
            $dump = & $adbPath -s $serial shell "dumpsys wifi" 2>$null
            if ($dump) {
                $mFac = [regex]::Match($dump, "(?i)mFactoryMacAddress\s*[=:]\s*([0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2})")
                if ($mFac.Success -and (Testar-MacValido $mFac.Groups[1].Value)) {
                    return (Normalizar-Mac $mFac.Groups[1].Value)
                }
                $mWifi = [regex]::Match($dump, "(?i)mWifiInfo.*?MAC:\s*([0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2})")
                if ($mWifi.Success -and (Testar-MacValido $mWifi.Groups[1].Value)) {
                    return (Normalizar-Mac $mWifi.Groups[1].Value)
                }
                $mMac = [regex]::Match($dump, "(?i)mMacAddress\s*[=:]\s*([0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2})")
                if ($mMac.Success -and (Testar-MacValido $mMac.Groups[1].Value)) {
                    $c = Normalizar-Mac $mMac.Groups[1].Value
                    if ($c -ne "02:00:00:00:00:00") { return $c }
                }
            }
        } catch {}

        # 2. Estratégia B: ip addr show wlan0 (EXCLUSIVAMENTE interface wlan0)
        try {
            $ipL = & $adbPath -s $serial shell "ip addr show wlan0 2>/dev/null || ip link show dev wlan0 2>/dev/null" 2>$null
            if ($ipL) {
                $wlanMatch = [regex]::Match($ipL, "(?i)link/ether\s+([0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2})")
                if ($wlanMatch.Success -and (Testar-MacValido $wlanMatch.Groups[1].Value)) {
                    return (Normalizar-Mac $wlanMatch.Groups[1].Value)
                }
            }
        } catch {}

        # 3. Estratégia C: sysfs do driver Wi-Fi (wlan0)
        $sysfsPaths = @(
            "/sys/class/net/wlan0/address",
            "/sys/devices/virtual/net/wlan0/address",
            "/sys/devices/platform/soc/*/wlan*/net/wlan0/address"
        )
        foreach ($p in $sysfsPaths) {
            try {
                $addr = (& $adbPath -s $serial shell "cat $p 2>/dev/null").Trim()
                if (Testar-MacValido $addr) { return (Normalizar-Mac $addr) }
            } catch {}
        }

        # 4. Estratégia D: getprop de hardware / bootloader
        $props = @("ro.boot.wifimac", "persist.sys.wifi.mac", "ro.ril.oem.wifimac", "ro.boot.mac")
        foreach ($pr in $props) {
            try {
                $pv = (& $adbPath -s $serial shell "getprop $pr 2>/dev/null").Trim()
                if (Testar-MacValido $pv) { return (Normalizar-Mac $pv) }
            } catch {}
        }

        # 5. Estratégia E: Re-ativação forçada de rádio e leitura wlan0
        try {
            & $adbPath -s $serial shell "input keyevent KEYCODE_WAKEUP" 2>$null
            & $adbPath -s $serial shell "cmd wifi set-wifi-enabled enabled 2>/dev/null || svc wifi enable 2>/dev/null" 2>$null
            Start-Sleep -Milliseconds 850
            $ipL2 = & $adbPath -s $serial shell "ip addr show wlan0 2>/dev/null" 2>$null
            if ($ipL2) {
                $m2 = [regex]::Match($ipL2, "(?i)link/ether\s+([0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2})")
                if ($m2.Success -and (Testar-MacValido $m2.Groups[1].Value)) {
                    return (Normalizar-Mac $m2.Groups[1].Value)
                }
            }
        } catch {}

        # 6. Estratégia F: Logcat recente do driver Wi-Fi
        try {
            $logs = & $adbPath -s $serial shell "logcat -d -t 1500" 2>$null
            if ($logs) {
                $matches = [regex]::Matches($logs, "(?i)(?:wlan|wifi|mac|efuse|nvram)[^a-f0-9\r\n]{1,30}([0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2})")
                foreach ($m in $matches) {
                    $cand = $m.Groups[1].Value
                    if (Testar-MacValido $cand) { return (Normalizar-Mac $cand) }
                }
            }
        } catch {}

        # 7. Estratégia G (Fallback avançado): Abre tela de Device Info
        if ($tentativa -ge 3) {
            try {
                & $adbPath -s $serial shell "am start -a android.settings.DEVICE_INFO_SETTINGS" 2>$null
                Start-Sleep -Milliseconds 1200
                $ipL3 = & $adbPath -s $serial shell "ip addr show wlan0 2>/dev/null" 2>$null
                if ($ipL3) {
                    $m3 = [regex]::Match($ipL3, "(?i)link/ether\s+([0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2})")
                    if ($m3.Success -and (Testar-MacValido $m3.Groups[1].Value)) {
                        $mac = Normalizar-Mac $m3.Groups[1].Value
                    }
                }
                & $adbPath -s $serial shell "input keyevent 3" 2>$null
                if (Testar-MacValido $mac) { return $mac }
            } catch {}
        }

        Start-Sleep -Milliseconds 800
    }

    return "N/A"
}

# Início da Execução
try { Clear-Host } catch {}
Exibir-BannerAdb
Write-Host ""

$adb = Localizar-Adb
if (-not $adb) {
    Write-Host " [!] Executavel ADB (adb.exe) nao encontrado no sistema." -ForegroundColor Red
    Write-Host "     Verifique se os platform-tools estao na pasta ou adicione o ADB ao PATH." -ForegroundColor Yellow
    Write-Host ""
    Write-Host " Pressione ENTER para sair..." -ForegroundColor Gray
    Read-Host | Out-Null
    exit 1
}

Write-Host " [*] Localizacao do ADB: " -NoNewline -ForegroundColor Gray
Write-Host "$adb" -ForegroundColor Green
Write-Host " [*] Carregando historico de MACs ja identificados anteriormente..." -ForegroundColor Gray

$cacheProcessados = Carregar-CacheMacsExistentes -arquivosCsv @(
    $csvSaida,
    (Join-Path $rootDir "dados\desktop\inventario_innovtech.csv"),
    (Join-Path $rootDir "dados\desktop\inventario_innovcore.csv"),
    (Join-Path $rootDir "inventario_innovtech.csv"),
    (Join-Path $rootDir "inventario_innovcore.csv")
)
$totalCadastrados = ($cacheProcessados.Keys | Where-Object { $cacheProcessados[$_].MAC } | Select-Object -Unique).Count
Write-Host " [+] Base com $totalCadastrados aparelho(s) previamente catalogado(s) com MAC." -ForegroundColor DarkGray
Write-Host ""
Write-Host " +----------------------------------------------------------------------------------------------+" -ForegroundColor Cyan
Write-Host " |  [PLUG & PLAY] MODO AUTO-SCAN ATIVO: Conecte qualquer tablet Android via cabo USB!           |" -ForegroundColor Green
Write-Host " |  1. Assim que conectar, identifica o tablet e verifica se ja possui endereco MAC registrado  |" -ForegroundColor White
Write-Host " |  2. Se ja tiver MAC cadastrado, avisa na hora com bip sonoro de confirmacao                  |" -ForegroundColor Yellow
Write-Host " |  3. Se nao tiver MAC, tenta a captura automatica com retries progressivos ate obter sucesso   |" -ForegroundColor Cyan
Write-Host " |  >> Pressione 'Q' ou 'ESC' a qualquer momento para pausar ou sair do monitor de bancada <<   |" -ForegroundColor DarkGray
Write-Host " +----------------------------------------------------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""

$dispositivosConectadosAnteriores = @()
$notificadosConectados = @{}
$totalCapturadosNestaSessao = 0

while ($true) {
    # Permite sair pressionando 'Q' ou 'ESC' sem travar quando executado interativamente
    $keyPressed = $null
    try {
        if ([Console]::KeyAvailable) {
            $keyPressed = [Console]::ReadKey($true)
        }
    } catch {}
    if ($keyPressed) {
        if ($keyPressed.KeyChar -eq 'q' -or $keyPressed.KeyChar -eq 'Q' -or $keyPressed.Key -eq [ConsoleKey]::Escape) {
            Write-Host "`n [!] Auto-scan finalizado pelo operador." -ForegroundColor Yellow
            break
        }
    }

    # Varre aparelhos conectados via ADB
    $rawDevices = & $adb devices -l
    $lines = $rawDevices -split "\r?\n"
    $atuais = @()

    foreach ($line in $lines) {
        if ($line -match "^([a-zA-Z0-9_-]+)\s+device\b") {
            $atuais += $Matches[1]
        } elseif ($line -match "^([a-zA-Z0-9_-]+)\s+unauthorized\b") {
            $devUnauth = $Matches[1]
            if (-not $notificadosConectados.ContainsKey("UNAUTH_$devUnauth")) {
                Write-Host " [!] Tablet [$devUnauth] detectado como NAO AUTORIZADO!" -ForegroundColor Red
                Write-Host "     Olhe a tela do tablet e marque 'Sempre permitir deste computador'." -ForegroundColor Yellow
                try { [Console]::Beep(800, 300) } catch {}
                $notificadosConectados["UNAUTH_$devUnauth"] = $true
            }
        }
    }

    # Detecta aparelhos que foram desconectados
    foreach ($ant in $dispositivosConectadosAnteriores) {
        if ($atuais -notcontains $ant) {
            Write-Host " [-] Tablet [$ant] desconectado da porta USB. Bancada pronta para o proximo tablet!" -ForegroundColor DarkYellow
            if ($notificadosConectados.ContainsKey($ant)) { $notificadosConectados.Remove($ant) }
            if ($notificadosConectados.ContainsKey("UNAUTH_$ant")) { $notificadosConectados.Remove("UNAUTH_$ant") }
        }
    }
    $dispositivosConectadosAnteriores = $atuais

    # Processa cada tablet conectado
    foreach ($dev in $atuais) {
        # Se já tratou este dispositivo enquanto permanece conectado, não repete
        if ($notificadosConectados.ContainsKey($dev)) {
            continue
        }

        # 1. Obtém o número de série de hardware
        $hwSerial = (& $adb -s $dev shell "getprop ro.serialno 2>/dev/null").Trim()
        if (-not $hwSerial) { $hwSerial = (& $adb -s $dev shell "getprop ro.boot.serialno 2>/dev/null").Trim() }
        if (-not $hwSerial) { $hwSerial = $dev }

        $marca = (& $adb -s $dev shell "getprop ro.product.manufacturer 2>/dev/null").Trim()
        if (-not $marca) { $marca = (& $adb -s $dev shell "getprop ro.product.brand 2>/dev/null").Trim() }
        if (-not $marca) { $marca = "Android" }

        $modelo = (& $adb -s $dev shell "getprop ro.product.model 2>/dev/null").Trim()
        $androidVer = (& $adb -s $dev shell "getprop ro.build.version.release 2>/dev/null").Trim()

        $bat = "N/A"
        try {
            $batRaw = & $adb -s $dev shell "dumpsys battery" 2>$null
            $mBat = [regex]::Match($batRaw, '(?i)level:\s*(\d+)')
            if ($mBat.Success) { $bat = "$($mBat.Groups[1].Value)%" }
        } catch {
            $bat = "N/A"
        }

        # 2. VERIFICA SE JÁ TEM O ENDEREÇO MAC DAQUELE DISPOSITIVO
        $cadastrado = $null
        if ($cacheProcessados.ContainsKey($hwSerial.ToUpper())) {
            $cadastrado = $cacheProcessados[$hwSerial.ToUpper()]
        } elseif ($cacheProcessados.ContainsKey($dev.ToUpper())) {
            $cadastrado = $cacheProcessados[$dev.ToUpper()]
        }

        if ($cadastrado -and (Testar-MacValido $cadastrado.MAC)) {
            Write-Host " ==============================================================================================" -ForegroundColor Cyan
            Write-Host " [DISPOSITIVO JA CADASTRADO] " -NoNewline -ForegroundColor Green
            Write-Host "N/S Hardware: [$hwSerial] | ID ADB: [$dev]" -ForegroundColor Yellow
            Write-Host "     Modelo:        " -NoNewline -ForegroundColor Gray
            Write-Host "$marca $modelo " -NoNewline -ForegroundColor White
            Write-Host "(Android $androidVer) | Bateria: " -NoNewline -ForegroundColor Gray
            Write-Host "$bat" -ForegroundColor Cyan
            Write-Host "     MAC Wi-Fi:     " -NoNewline -ForegroundColor White
            Write-Host "$($cadastrado.MAC)" -NoNewline -ForegroundColor Green
            $infoOrigem = if ($cadastrado.Patrimonio -and $cadastrado.Patrimonio -ne 'N/A') { "Patrimonio: $($cadastrado.Patrimonio) | Base: $($cadastrado.Origem)" } else { "Base: $($cadastrado.Origem)" }
            Write-Host " ($infoOrigem)" -ForegroundColor DarkGray
            Write-Host "     Status:        Este dispositivo JA possui endereco MAC registrado!" -ForegroundColor White
            Write-Host "     Acao:          Nenhuma acao necessaria. Pode desconectar e plugar o proximo tablet!" -ForegroundColor Green
            Write-Host " ==============================================================================================" -ForegroundColor Cyan
            Write-Host ""

            # Bip sonoro duplo indicando que já está cadastrado
            try { [Console]::Beep(1800, 140); Start-Sleep -Milliseconds 60; [Console]::Beep(2200, 180) } catch {}

            # Copia para Área de Transferência
            $clipTexto = "$hwSerial;$($cadastrado.MAC);$marca;$modelo"
            try { Set-Clipboard -Value $clipTexto } catch { try { $clipTexto | clip.exe } catch {} }

            $notificadosConectados[$dev] = $true
            continue
        }

        # 3. DISPOSITIVO SEM MAC CADASTRADO: RECOLHE AUTOMATICAMENTE COM RETRIES
        Write-Host " ==============================================================================================" -ForegroundColor Cyan
        Write-Host " [NOVO TABLET CONECTADO - SEM MAC CADASTRADO] " -ForegroundColor Green
        Write-Host "     N/S Hardware:  [$hwSerial] | ID ADB: [$dev]" -ForegroundColor Yellow
        Write-Host "     Modelo:        $marca $modelo (Android $androidVer) | Bateria: $bat" -ForegroundColor White
        Write-Host "     Status:        O endereco MAC Wi-Fi deste aparelho ainda NAO consta no cadastro!" -ForegroundColor Yellow
        Write-Host "     Acao:          Recolhendo MAC Wi-Fi automaticamente (tentando com retries de hardware)..." -ForegroundColor Cyan

        $macCapturado = "N/A"
        $ciclo = 1
        $maxCiclos = 8
        $tempoTotalMs = 0

        while (-not (Testar-MacValido $macCapturado) -and $ciclo -le $maxCiclos) {
            # Confere se o tablet ainda está fisicamente plugado
            $devsAtuais = & $adb devices
            $aindaPlugado = ($devsAtuais | Where-Object { $_ -match "^$([regex]::Escape($dev))\s+device\b" }).Count -gt 0
            if (-not $aindaPlugado) {
                Write-Host " [-] Dispositivo desconectado durante as tentativas de captura." -ForegroundColor DarkYellow
                break
            }

            if ($ciclo -gt 1) {
                Write-Host " [RETRY $ciclo/$maxCiclos] Acordando o tablet e tentando novamente..." -ForegroundColor Yellow
                & $adb -s $dev shell "input keyevent KEYCODE_WAKEUP" 2>$null
                Start-Sleep -Milliseconds 600
            }

            $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $macCapturado = Obter-MacComRetries -adbPath $adb -serial $dev -maxRetries 6 -onProgress {
                param($msg)
                Write-Host "       ... $msg" -ForegroundColor Gray
            }
            $stopwatch.Stop()
            $tempoTotalMs += [Math]::Round($stopwatch.Elapsed.TotalMilliseconds, 0)

            if (Testar-MacValido $macCapturado) {
                break
            }

            $ciclo++
            if ($ciclo -le $maxCiclos) {
                Write-Host " [!] MAC ainda nao liberado pelo driver de Wi-Fi do tablet." -ForegroundColor DarkYellow
                Write-Host "     Ligando a tela do tablet e tentando novamente em 2 segundos... (Tentativa $ciclo de $maxCiclos)" -ForegroundColor Cyan
                & $adb -s $dev shell "input keyevent KEYCODE_WAKEUP" 2>$null
                Start-Sleep -Seconds 2
            }
        }

        if (Testar-MacValido $macCapturado) {
            # Bip sonoro de confirmação
            try { [Console]::Beep(1400, 220) } catch {}

            Write-Host ""
            Write-Host " [SUCESSO ABSOLUTO] " -NoNewline -ForegroundColor Green
            Write-Host "MAC Wi-Fi Capturado: " -NoNewline -ForegroundColor White
            Write-Host "$macCapturado" -NoNewline -ForegroundColor Green
            Write-Host " (Tempo total: ${tempoTotalMs}ms)" -ForegroundColor DarkGray

            # Armazena no cache para não reprocessar repetidamente
            $novoObj = [PSCustomObject]@{
                Serial     = $hwSerial
                MAC        = $macCapturado
                Marca      = $marca
                Modelo     = $modelo
                Patrimonio = "N/A"
                Origem     = [System.IO.Path]::GetFileName($csvSaida)
            }
            $cacheProcessados[$hwSerial.ToUpper()] = $novoObj
            $cacheProcessados[$dev.ToUpper()] = $novoObj
            $cacheProcessados[$hwSerial] = $novoObj
            $cacheProcessados[$dev] = $novoObj
            $notificadosConectados[$dev] = $true
            $totalCapturadosNestaSessao++

            # Grava no relatório CSV
            $existeCsv = Test-Path $csvSaida
            $linhasCsv = @()
            if (-not $existeCsv) {
                $linhasCsv += "Data_Registro;Serial;MAC_WiFi;Fabricante;Modelo;Android;Bateria"
            }
            $linhasCsv += "$((Get-Date).ToString('dd/MM/yyyy HH:mm:ss'));$hwSerial;$macCapturado;$marca;$modelo;$androidVer;$bat"
            $linhasCsv | Out-File -FilePath $csvSaida -Append -Encoding UTF8

            # Copia para Área de Transferência
            $clipTexto = "$hwSerial;$macCapturado;$marca;$modelo"
            try { Set-Clipboard -Value $clipTexto } catch { try { $clipTexto | clip.exe } catch {} }

            Write-Host " [OK] Gravado no relatorio: $csvSaida" -ForegroundColor Green
            Write-Host " [OK] Dados copiados para a Area de Transferencia (Ctrl + V)!" -ForegroundColor Green
            Write-Host " [-> PROXIMO] Tablet finalizado com exito! Pode desconectar e plugar o proximo." -ForegroundColor Yellow
            Write-Host " ==============================================================================================" -ForegroundColor Cyan
            Write-Host ""
        } else {
            # Se não conseguiu capturar e ainda está plugado, não marca como notificado
            # Avisa o operador e tentará novamente no próximo ciclo
            Write-Host " [ATENCAO] Nao foi possivel obter o MAC deste tablet apos $maxCiclos ciclos de tentativas." -ForegroundColor Red
            Write-Host "     Verifique se o tablet esta desbloqueado e com a opcao 'Sempre permitir deste computador' autorizada." -ForegroundColor Yellow
            Write-Host "     O sistema tentara novamente de forma automatica assim que a tela for destravada..." -ForegroundColor Gray
            Write-Host " ==============================================================================================" -ForegroundColor Cyan
            Write-Host ""
        }
    }

    Start-Sleep -Milliseconds 1000
}

Write-Host ""
Write-Host " Sessao encerrada. Total de novos tablets capturados: $totalCapturadosNestaSessao" -ForegroundColor Green
Write-Host " Relatorio salvo em: $csvSaida" -ForegroundColor Yellow
Write-Host " Pressione ENTER para fechar a janela..." -ForegroundColor Gray
Read-Host | Out-Null
