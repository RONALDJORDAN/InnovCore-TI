# ==============================================================================
# INNOVCORE TI - SISTEMA DE GESTAO DE ATIVOS E INVENTARIO DE HARDWARE
# Desenvolvido por Jordan | Versao 2.0 (Edicao Corporativa Avancada)
# Compativel com Windows 10 e Windows 11 (32 e 64 bits) | PowerShell 5.1 & 7+
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::InputEncoding = [System.Text.Encoding]::UTF8 } catch {}
try { $Host.UI.RawUI.WindowTitle = "InnovCore TI - Gestao de Ativos & Inventario [v2.0]" } catch {}

# Identifica o diretorio raiz do Pen Drive / pacote
$scriptRootDir = if ($env:INNOVCORE_USB_DIR -and (Test-Path $env:INNOVCORE_USB_DIR)) { 
    $env:INNOVCORE_USB_DIR.TrimEnd('\') 
} elseif ($PSScriptRoot) { 
    if ($PSScriptRoot -match '(?i)[\\/]core$') {
        Split-Path $PSScriptRoot -Parent
    } else {
        $PSScriptRoot
    }
} elseif ($MyInvocation.MyCommand.Path) { 
    $parent = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path)
    if ($parent -match '(?i)[\\/]core$') {
        Split-Path $parent -Parent
    } else {
        $parent
    }
} else { 
    (Get-Location).Path 
}

# Localizacao inteligente da base de dados CSV de Desktops / Computadores
$pastaDadosDesktop   = Join-Path $scriptRootDir "dados\desktop"
$caminhoCsvDesktop   = Join-Path $pastaDadosDesktop "inventario_innovtech.csv"
$caminhoCsvPrincipal = Join-Path $pastaDadosDesktop "inventario_innovcore.csv"
$caminhoCsvLegacy    = Join-Path $scriptRootDir "inventario_innovtech.csv"

# Se ja existir a base antiga e nao a nova, migra suavemente sem perdas
$caminhoCsv = if (Test-Path $caminhoCsvDesktop) {
    $caminhoCsvDesktop
} elseif (Test-Path $caminhoCsvPrincipal) {
    $caminhoCsvPrincipal
} elseif (Test-Path $caminhoCsvLegacy) {
    $caminhoCsvLegacy
} else {
    $caminhoCsvDesktop
}

# Pastas de exportacoes e backups
$pastaExport = Join-Path $scriptRootDir "Exportacoes"
$pastaBackup = Join-Path $scriptRootDir "Backups"

# Definicao de Cabecalho Completo Oficial InnovCore TI (Compativel com Excel PT-BR)
$cabecalhoOficial = @(
    "Data_Registro",
    "Empresa",
    "Patrimonio",
    "Usuario",
    "Funcao",
    "Setor",
    "Tipo_Equipamento",
    "Marca",
    "Modelo",
    "Numero_Serie",
    "Ano_Fabricacao",
    "Nome_Computador",
    "Processador",
    "Memoria_RAM",
    "Armazenamento",
    "Placa_Video",
    "Sistema_Operacional",
    "IP_Rede",
    "MAC_Rede"
)
$cabecalhoCsvString = $cabecalhoOficial -join ';'

$utf8ComBOM = New-Object System.Text.UTF8Encoding($true)
$utf8SemBOM = New-Object System.Text.UTF8Encoding($false)

# ==============================================================================
# IDENTIDADE VISUAL E BANNER CORPORATIVO
# ==============================================================================
function Exibir-LogoPrincipal {
    Write-Host " +======================================================================+" -ForegroundColor Cyan
    Write-Host " |  ___ _   _ _   _  _____     ______ _____ ____  _____   _____ ___     |" -ForegroundColor Yellow
    Write-Host " | |_ _| \ | | \ | |/ _ \ \   / / ___/ _ \|  _ \| ____| |_   _|_ _|    |" -ForegroundColor Yellow
    Write-Host " |  | ||  \| |  \| | | | \ \ / / |  | | | | |_) |  _|     | |  | |     |" -ForegroundColor Yellow
    Write-Host " |  | || |\  | |\  | |_| |\ V /| |__| |_| |  _ <| |___    | |  | |     |" -ForegroundColor Yellow
    Write-Host " | |___|_| \_|_| \_|\___/  \_/  \____\___/|_| \_\_____|   |_| |___|    |" -ForegroundColor Yellow
    Write-Host " |                                                                      |" -ForegroundColor Cyan
    Write-Host " |             INNOVCORE TI - GESTAO DE ATIVOS E INVENTARIO             |" -ForegroundColor Green
    Write-Host " |                 Desenvolvido por Jordan | Versao 2.0                 |" -ForegroundColor White
    Write-Host " +======================================================================+" -ForegroundColor Cyan
}

function Pausar-Tela {
    param([string]$mensagem = " Pressione ENTER ou ESC para voltar ao menu...")
    Write-Host ""
    Write-Host $mensagem -ForegroundColor Gray
    try {
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } catch {
        Read-Host | Out-Null
    }
}

function Ler-OpcaoOuEsc {
    param(
        [string]$prompt = "",
        [string]$valorPadrao = "",
        [string]$corPrompt = "Cyan"
    )

    if ($prompt) {
        Write-Host $prompt -NoNewline -ForegroundColor $corPrompt
    }
    if ($valorPadrao) {
        Write-Host " [$valorPadrao] " -NoNewline -ForegroundColor Gray
    } else {
        Write-Host " " -NoNewline
    }

    try {
        if ([Console]::IsInputRedirected) {
            $val = Read-Host
            if ([string]::IsNullOrWhiteSpace($val)) { return $valorPadrao }
            return $val.Trim()
        }

        $buffer = ""
        while ($true) {
            $keyInfo = [Console]::ReadKey($true)

            # TECLA ESC -> Retorna '0' imediatamente para voltar ao menu
            if ($keyInfo.Key -eq [System.ConsoleKey]::Escape) {
                Write-Host " [ESC - Voltar]" -ForegroundColor Yellow
                return "0"
            }

            # TECLA ENTER
            if ($keyInfo.Key -eq [System.ConsoleKey]::Enter) {
                Write-Host ""
                if ([string]::IsNullOrWhiteSpace($buffer)) {
                    return $valorPadrao
                }
                return $buffer.Trim()
            }

            # TECLA BACKSPACE
            if ($keyInfo.Key -eq [System.ConsoleKey]::Backspace) {
                if ($buffer.Length -gt 0) {
                    $buffer = $buffer.Substring(0, $buffer.Length - 1)
                    [Console]::Write("`b `b")
                }
                continue
            }

            # CARACTERES DIGITADOS
            if ($keyInfo.KeyChar -and -not [char]::IsControl($keyInfo.KeyChar)) {
                $buffer += $keyInfo.KeyChar
                [Console]::Write($keyInfo.KeyChar)
            }
        }
    } catch {
        $val = Read-Host
        if ([string]::IsNullOrWhiteSpace($val)) { return $valorPadrao }
        return $val.Trim()
    }
}

# ==============================================================================
# MANIPULACAO ROBUSTA DE ARQUIVOS E BANCO CSV
# ==============================================================================
function Obter-RegistrosCsv {
    param([string]$caminho)
    $lista = @()
    if (Test-Path $caminho) {
        try {
            $raw = [System.IO.File]::ReadAllText($caminho, [System.Text.Encoding]::UTF8)
            $linhas = $raw -split "\r?\n" | Where-Object { ![string]::IsNullOrWhiteSpace($_) }
            if ($linhas.Count -gt 1) {
                $cabecalhoLinha = ($linhas[0] -replace "^\uFEFF", "")
                $cabecalhos = $cabecalhoLinha -split ';'
                for ($i = 1; $i -lt $linhas.Count; $i++) {
                    $valores = $linhas[$i] -split ';'
                    $obj = [PSCustomObject]@{}
                    for ($k = 0; $k -lt $cabecalhos.Count; $k++) {
                        $prop = $cabecalhos[$k].Trim()
                        $val = if ($k -lt $valores.Count) { $valores[$k].Trim() } else { "" }
                        $obj | Add-Member -MemberType NoteProperty -Name $prop -Value $val -Force
                    }
                    foreach ($propOficial in $script:cabecalhoOficial) {
                        if (-not ($obj.PSObject.Properties[$propOficial])) {
                            $obj | Add-Member -MemberType NoteProperty -Name $propOficial -Value "N/A" -Force
                        }
                    }
                    $lista += $obj
                }
            }
        } catch {
            Write-Host " [!] Aviso ao ler CSV: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    return $lista
}

function Fazer-BackupCsv {
    param([string]$caminho)
    if (Test-Path $caminho) {
        try {
            if (!(Test-Path $script:pastaBackup)) {
                New-Item -ItemType Directory -Path $script:pastaBackup -Force | Out-Null
            }
            $dataHora = Get-Date -Format "yyyyMMdd_HHmmss"
            $nomeBase = [System.IO.Path]::GetFileNameWithoutExtension($caminho)
            $destinoBak = Join-Path $script:pastaBackup "$($nomeBase)_backup_$($dataHora).csv"
            Copy-Item -Path $caminho -Destination $destinoBak -Force
        } catch {}
    }
}

function Salvar-RegistrosCsv {
    param([string]$caminho, [array]$registros)
    $limparCampo = { 
        param($c) 
        if ($null -ne $c) { 
            ($c.ToString() -replace ';', ',') -replace "[\r\n]+", " " 
        } else { 
            "" 
        } 
    }
    
    # 1. Garante que o diretorio de destino existe
    $pastaDestino = [System.IO.Path]::GetDirectoryName($caminho)
    if ($pastaDestino -and !(Test-Path $pastaDestino)) {
        try { New-Item -ItemType Directory -Path $pastaDestino -Force | Out-Null } catch {}
    }
    
    # 2. Faz backup preventivo se o arquivo ja existir
    Fazer-BackupCsv -caminho $caminho
    
    # 3. Monta o conteudo do CSV
    $linhasTexto = [System.Collections.Generic.List[string]]::new()
    $linhasTexto.Add($script:cabecalhoCsvString)
    foreach ($r in $registros) {
        $linha = @(
            (& $limparCampo $r.Data_Registro),
            (& $limparCampo $r.Empresa),
            (& $limparCampo $r.Patrimonio),
            (& $limparCampo $r.Usuario),
            (& $limparCampo $r.Funcao),
            (& $limparCampo $r.Setor),
            (& $limparCampo $r.Tipo_Equipamento),
            (& $limparCampo $r.Marca),
            (& $limparCampo $r.Modelo),
            (& $limparCampo $r.Numero_Serie),
            (& $limparCampo $r.Ano_Fabricacao),
            (& $limparCampo $r.Nome_Computador),
            (& $limparCampo $r.Processador),
            (& $limparCampo $r.Memoria_RAM),
            (& $limparCampo $r.Armazenamento),
            (& $limparCampo $r.Placa_Video),
            (& $limparCampo $r.Sistema_Operacional),
            (& $limparCampo $r.IP_Rede),
            (& $limparCampo $r.MAC_Rede)
        ) -join ';'
        $linhasTexto.Add($linha)
    }
    $conteudoFinal = ($linhasTexto -join "`r`n") + "`r`n"
    
    # 4. Loop de gravacao ultra-resiliente
    $salvo = $false
    $tentativa = 1
    $caminhoAlvo = $caminho
    
    while (-not $salvo) {
        if (Test-Path $caminhoAlvo) {
            try { (Get-Item $caminhoAlvo -Force).Attributes = 'Normal' } catch {}
        }
        
        try {
            [System.IO.File]::WriteAllText($caminhoAlvo, $conteudoFinal, $script:utf8ComBOM)
            $salvo = $true
            $script:caminhoCsv = $caminhoAlvo
        } catch {
            Write-Host ""
            Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Red
            Write-Host " | [!] AVISO: NAO FOI POSSIVEL GRAVAR NA PLANILHA CSV                    |" -ForegroundColor Yellow
            Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Red
            Write-Host "   Arquivo: $caminhoAlvo" -ForegroundColor Cyan
            Write-Host "   Motivo:  $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "   Possiveis causas:" -ForegroundColor White
            Write-Host "   1. O arquivo CSV esta ABERTO no Excel, VS Code ou outro editor." -ForegroundColor Gray
            Write-Host "   2. A unidade do Pen Drive esta temporariamente protegida." -ForegroundColor Gray
            Write-Host ""
            Write-Host "   [T] Tentar salvar novamente (Apos fechar o Excel/Editor)" -ForegroundColor Green
            Write-Host "   [N] Salvar em um NOVO arquivo no Pen Drive" -ForegroundColor Cyan
            Write-Host "   [E] Salvar copia de EMERGENCIA na Area de Trabalho" -ForegroundColor Yellow
            Write-Host ""
            Write-Host " Escolha uma opcao " -NoNewline -ForegroundColor Cyan
            $opcaoErro = Read-Host "[T]"
            if ([string]::IsNullOrWhiteSpace($opcaoErro)) { $opcaoErro = "T" }
            
            switch ($opcaoErro.Trim().ToUpper()) {
                "N" {
                    $timestampNovo = Get-Date -Format "yyyyMMdd_HHmmss"
                    $caminhoAlvo = Join-Path $script:internalDir "inventario_innovcore_$timestampNovo.csv"
                    Write-Host "   -> Tentando gravar em: $caminhoAlvo" -ForegroundColor Cyan
                }
                "E" {
                    $caminhoAlvo = Join-Path $env:USERPROFILE "Desktop\inventario_innovcore_emergencia.csv"
                    Write-Host "   -> Salvando na Area de Trabalho: $caminhoAlvo" -ForegroundColor Yellow
                }
                default {
                    Write-Host "   -> Tentando gravar novamente..." -ForegroundColor Gray
                }
            }
            $tentativa++
        }
    }
}

# ==============================================================================
# CALCULO INTELIGENTE DE PATRIMONIO SEQUENCIAL
# ==============================================================================
function Obter-ProximoPatrimonio {
    param([array]$registros)
    if (!$registros -or $registros.Count -eq 0) {
        return "PAT-0001"
    }
    
    $maiorNumero = 0
    foreach ($r in $registros) {
        if ($r.Patrimonio -match 'PAT-(\d+)') {
            $num = [int]$Matches[1]
            if ($num -gt $maiorNumero) {
                $maiorNumero = $num
            }
        }
    }
    
    if ($maiorNumero -eq 0) {
        $maiorNumero = $registros.Count
    }
    
    $proximo = $maiorNumero + 1
    return "PAT-" + $proximo.ToString("D4")
}

# ==============================================================================
# AUDITORIA E TESTE DE DUPLICIDADE EM TEMPO REAL
# ==============================================================================
function Testar-DuplicidadeEquipamento {
    param(
        [array]$registros,
        [string]$patrimonio,
        [string]$numeroSerie,
        [string]$macRede,
        [string]$nomeComputador
    )
    
    if (!$registros -or $registros.Count -eq 0) { return $null }
    
    foreach ($r in $registros) {
        $motivo = $null
        
        # 1. Verifica Patrimonio
        if (![string]::IsNullOrWhiteSpace($patrimonio) -and ![string]::IsNullOrWhiteSpace($r.Patrimonio)) {
            if ($r.Patrimonio.Trim().ToUpper() -eq $patrimonio.Trim().ToUpper()) {
                $motivo = "Patrimonio repetido ($patrimonio)"
            }
        }
        
        # 2. Verifica Serial Number (se nao for N/A ou generico)
        if (-not $motivo -and ![string]::IsNullOrWhiteSpace($numeroSerie) -and $numeroSerie.Trim() -notmatch '(?i)^n/a|none|default|to be filled|0000$') {
            if (![string]::IsNullOrWhiteSpace($r.Numero_Serie) -and $r.Numero_Serie.Trim().ToUpper() -eq $numeroSerie.Trim().ToUpper()) {
                $motivo = "Numero de Serie da BIOS identico ($numeroSerie)"
            }
        }
        
        # 3. Verifica Endereco MAC (se nao for N/A)
        if (-not $motivo -and ![string]::IsNullOrWhiteSpace($macRede) -and $macRede.Trim() -notmatch '(?i)^n/a$') {
            if (![string]::IsNullOrWhiteSpace($r.MAC_Rede) -and $r.MAC_Rede.Trim().ToUpper() -eq $macRede.Trim().ToUpper()) {
                $motivo = "Placa de Rede / MAC Address identico ($macRede)"
            }
        }
        
        if ($motivo) {
            return [PSCustomObject]@{
                RegistroExistente = $r
                Motivo            = $motivo
            }
        }
    }
    
    return $null
}

# ==============================================================================
# SCANNER COMPLETO DE HARDWARE E SISTEMA (MODO AUTOMATICO)
# ==============================================================================
function Coletar-HardwareCompleto {
    $dataColeta = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    $hostname = $env:COMPUTERNAME
    
    # BIOS & Serial
    $bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
    $serialBios = if ($bios.SerialNumber) { $bios.SerialNumber.Trim() } else { "" }
    
    $bb = Get-CimInstance Win32_BaseBoard -ErrorAction SilentlyContinue
    $serialPlaca = if ($bb.SerialNumber) { $bb.SerialNumber.Trim() } else { "" }
    
    $numeroSerie = if ($serialBios -and $serialBios -notmatch '(?i)Default|To be filled|None|000000|System Serial Number') { 
        $serialBios 
    } elseif ($serialPlaca -and $serialPlaca -notmatch '(?i)Default|To be filled|None|000000') { 
        $serialPlaca 
    } elseif ($serialBios) {
        $serialBios
    } else { 
        "N/A" 
    }
    
    # Ano de Fabricacao / Data BIOS
    $anoFabricacao = if ($bios.ReleaseDate) {
        if ($bios.ReleaseDate -is [datetime]) {
            $bios.ReleaseDate.Year.ToString()
        } else {
            $dataStr = $bios.ReleaseDate.ToString()
            if ($dataStr -match '^\d{4}') { $Matches[0] } else { $dataStr.Substring(0, [math]::Min(4, $dataStr.Length)) }
        }
    } else {
        "N/A"
    }
    
    # Computador / Fabricante / Modelo
    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
    $marca = if ($cs.Manufacturer) { $cs.Manufacturer.Trim() } else { "N/A" }
    $modelo = if ($cs.Model) { $cs.Model.Trim() } else { "N/A" }
    
    if ($marca -match '(?i)System manufacturer|To be filled|Default' -and $bb.Manufacturer) {
        $marca = $bb.Manufacturer.Trim()
    }
    if ($modelo -match '(?i)System Product Name|To be filled|Default' -and $bb.Product) {
        $modelo = $bb.Product.Trim()
    }
    
    # Tipo de Equipamento (Notebook, Desktop, etc)
    $chassis = Get-CimInstance Win32_SystemEnclosure -ErrorAction SilentlyContinue
    $tipoChassisNum = if ($chassis.ChassisTypes) { $chassis.ChassisTypes[0] } else { 0 }
    $tipoEquipamento = switch ($tipoChassisNum) {
        { $_ -in 8, 9, 10, 11, 12, 14, 18, 21, 31, 32 } { "Notebook / Laptop" }
        { $_ -in 3, 4, 5, 6, 7, 15, 16 } { "Desktop / PC" }
        { $_ -in 13 } { "All-in-One" }
        { $_ -in 23 } { "Servidor" }
        default { "Computador" }
    }
    
    # Processador (CPU)
    $cpu = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
    $cpuNome = if ($cpu.Name) { ($cpu.Name.Trim() -replace '\s+', ' ') } else { "N/A" }
    
    # Memoria RAM
    $ramModulos = Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue
    $ramTotalBytes = 0
    $ramDetalhes = @()
    if ($ramModulos) {
        foreach ($m in $ramModulos) {
            $ramTotalBytes += $m.Capacity
            $capGB = [math]::Round($m.Capacity / 1GB, 0)
            $speed = if ($m.Speed) { "$($m.Speed)MHz" } else { "" }
            $ramDetalhes += "$capGB GB $speed".Trim()
        }
    }
    $ramTotalGB = if ($ramTotalBytes -gt 0) { 
        [math]::Round($ramTotalBytes / 1GB, 0) 
    } elseif ($cs.TotalPhysicalMemory) { 
        [math]::Round($cs.TotalPhysicalMemory / 1GB, 0) 
    } else { 
        0 
    }
    $ramFormatada = if ($ramTotalGB -gt 0) {
        if ($ramDetalhes.Count -gt 0) { "$ramTotalGB GB RAM ($($ramDetalhes -join ' + '))" } else { "$ramTotalGB GB RAM" }
    } else { "N/A" }
    
    # Armazenamento (Discos)
    $discos = Get-CimInstance Win32_DiskDrive -ErrorAction SilentlyContinue
    $listaDiscos = @()
    if ($discos) {
        foreach ($d in $discos) {
            $tamGB = [math]::Round($d.Size / 1GB, 0)
            $tipoMedia = ""
            try {
                $pd = Get-PhysicalDisk | Where-Object { $_.DeviceId -eq $d.Index } -ErrorAction SilentlyContinue
                if ($pd.MediaType) { $tipoMedia = "[$($pd.MediaType)] " }
            } catch {}
            $listaDiscos += "$tipoMedia$($d.Model.Trim()) ($tamGB GB)"
        }
    }
    $armazenamento = if ($listaDiscos.Count -gt 0) { $listaDiscos -join " | " } else { "N/A" }
    
    # Placa de Video (GPU)
    $gpus = Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue
    $gpuNomes = ($gpus | Where-Object { $_.Name } | ForEach-Object { $_.Name.Trim() } | Select-Object -Unique)
    $placaVideo = if ($gpuNomes) { $gpuNomes -join " | " } else { "N/A" }
    
    # Sistema Operacional
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $osNome = if ($os.Caption) { $os.Caption.Trim() } else { "Windows" }
    $osArch = if ($os.OSArchitecture) { $os.OSArchitecture } else { "64-bit" }
    $osBuild = if ($os.BuildNumber) { $os.BuildNumber } else { "" }
    $sistemaOperacional = "$osNome ($osArch) - Build $osBuild"
    
    # Rede (IP e MAC)
    $rede = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" -ErrorAction SilentlyContinue | Select-Object -First 1
    $ipV4 = if ($rede.IPAddress) { ($rede.IPAddress | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' } | Select-Object -First 1) } else { "N/A" }
    $macAddr = if ($rede.MACAddress) { $rede.MACAddress } else { "N/A" }
    
    return [PSCustomObject]@{
        Data_Registro       = $dataColeta
        Nome_Computador     = $hostname
        Tipo_Equipamento    = $tipoEquipamento
        Marca               = $marca
        Modelo              = $modelo
        Numero_Serie        = $numeroSerie
        Ano_Fabricacao      = $anoFabricacao
        Processador         = $cpuNome
        Memoria_RAM         = $ramFormatada
        Armazenamento       = $armazenamento
        Placa_Video         = $placaVideo
        Sistema_Operacional = $sistemaOperacional
        IP_Rede             = $ipV4
        MAC_Rede            = $macAddr
    }
}

# ==============================================================================
# SUBSISTEMA DE CAPTURA E INVENTARIO DE TABLETS ANDROID (ADB)
# ==============================================================================
function Localizar-Adb {
    $baseDir = if ($env:INNOVCORE_USB_DIR -and (Test-Path $env:INNOVCORE_USB_DIR)) { 
        $env:INNOVCORE_USB_DIR.TrimEnd('\') 
    } else { 
        $scriptRootDir 
    }
    $possiveis = @(
        (Join-Path $baseDir "mobile\platform-tools\adb.exe"),
        (Join-Path $baseDir "adb\platform-tools\adb.exe"),
        (Join-Path $baseDir "platform-tools\adb.exe"),
        (Join-Path $baseDir "..\mobile\platform-tools\adb.exe"),
        (Join-Path $baseDir "..\adb\platform-tools\adb.exe"),
        (Join-Path $baseDir "..\..\Inno_RPA\tools\platform-tools\adb.exe"),
        (Join-Path $baseDir "..\Inno_RPA\tools\platform-tools\adb.exe"),
        (Join-Path $baseDir "tools\platform-tools\adb.exe"),
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

function Obter-MacWifiAndroid {
    param([string]$adbPath, [string]$serial)

    # 1. Estrategia A: dumpsys wifi (mFactoryMacAddress de hardware)
    try {
        $dumpWifi = & $adbPath -s $serial shell "dumpsys wifi" 2>$null
        if ($dumpWifi) {
            $mFac = [regex]::Match($dumpWifi, "(?i)mFactoryMacAddress\s*[=:]\s*([0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2})")
            if ($mFac.Success -and (Testar-MacValido $mFac.Groups[1].Value)) {
                return (Normalizar-Mac $mFac.Groups[1].Value)
            }
            $mMac = [regex]::Match($dumpWifi, "(?i)mMacAddress\s*[=:]\s*([0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2})")
            if ($mMac.Success -and (Testar-MacValido $mMac.Groups[1].Value)) {
                $cand = Normalizar-Mac $mMac.Groups[1].Value
                if ($cand -ne "02:00:00:00:00:00") { return $cand }
            }
        }
    } catch {}

    # 2. Estrategia B: ip link show wlan0
    try {
        $ipLink = & $adbPath -s $serial shell "ip link show wlan0" 2>$null
        if ($ipLink) {
            $mLink = [regex]::Match($ipLink, "(?i)link/ether\s+([0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2})")
            if ($mLink.Success -and (Testar-MacValido $mLink.Groups[1].Value)) {
                return (Normalizar-Mac $mLink.Groups[1].Value)
            }
        }
    } catch {}

    # 3. Estrategia C: sysfs do driver Wi-Fi
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

    # 4. Estrategia D: getprop de hardware / bootloader
    $props = @("ro.boot.wifimac", "persist.sys.wifi.mac", "ro.ril.oem.wifimac", "ro.boot.mac")
    foreach ($pr in $props) {
        try {
            $pv = (& $adbPath -s $serial shell "getprop $pr 2>/dev/null").Trim()
            if (Testar-MacValido $pv) { return (Normalizar-Mac $pv) }
        } catch {}
    }

    # 5. Estrategia E: Ativacao rapida temporaria do radio Wi-Fi
    try {
        & $adbPath -s $serial shell "svc wifi enable" 2>$null
        Start-Sleep -Milliseconds 850
        $ipLink2 = & $adbPath -s $serial shell "ip link show wlan0" 2>$null
        if ($ipLink2) {
            $m2 = [regex]::Match($ipLink2, "(?i)link/ether\s+([0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2})")
            if ($m2.Success -and (Testar-MacValido $m2.Groups[1].Value)) {
                $mac = Normalizar-Mac $m2.Groups[1].Value
            }
        }
        if (-not $mac) {
            $addr2 = (& $adbPath -s $serial shell "cat /sys/class/net/wlan0/address 2>/dev/null").Trim()
            if (Testar-MacValido $addr2) { $mac = Normalizar-Mac $addr2 }
        }
        & $adbPath -s $serial shell "svc wifi disable" 2>$null
        if ($mac) { return $mac }
    } catch {}

    # 6. Estrategia F: Logcat recente do driver
    try {
        $logs = & $adbPath -s $serial shell "logcat -d -t 1500" 2>$null
        if ($logs) {
            $matches = [regex]::Matches($logs, "(?i)(?:wlan|wifi|mac|efuse|nvram|address)[^a-f0-9\r\n]{1,30}([0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2})")
            foreach ($m in $matches) {
                $cand = $m.Groups[1].Value
                if (Testar-MacValido $cand) { return (Normalizar-Mac $cand) }
            }
        }
    } catch {}

    return "N/A"
}

function Obter-MacComRetries {
    param(
        [string]$adbPath,
        [string]$serial,
        [int]$maxRetries = 6,
        [scriptblock]$onProgress = $null
    )

    $mac = "N/A"
    for ($tentativa = 1; $tentativa -le $maxRetries; $tentativa++) {
        if ($onProgress) {
            & $onProgress "Tentativa $tentativa de $maxRetries..."
        }

        # 1. Estratégia A: dumpsys wifi (mFactoryMacAddress de hardware)
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
                $mL = [regex]::Match($ipL, "(?i)link/ether\s+([0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2})")
                if ($mL.Success -and (Testar-MacValido $mL.Groups[1].Value)) {
                    return (Normalizar-Mac $mL.Groups[1].Value)
                }
            }
        } catch {}

        # 3. Estratégia C: sysfs do driver Wi-Fi
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

        # 5. Estratégia E: Ativação rápida temporária do rádio Wi-Fi
        try {
            & $adbPath -s $serial shell "input keyevent KEYCODE_WAKEUP" 2>$null
            & $adbPath -s $serial shell "svc wifi enable" 2>$null
            Start-Sleep -Milliseconds 850
            $ipL2 = & $adbPath -s $serial shell "ip link show wlan0" 2>$null
            if ($ipL2) {
                $m2 = [regex]::Match($ipL2, "(?i)link/ether\s+([0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2})")
                if ($m2.Success -and (Testar-MacValido $m2.Groups[1].Value)) {
                    $mac = Normalizar-Mac $m2.Groups[1].Value
                }
            }
            if (-not (Testar-MacValido $mac)) {
                $addr2 = (& $adbPath -s $serial shell "cat /sys/class/net/wlan0/address 2>/dev/null").Trim()
                if (Testar-MacValido $addr2) { $mac = Normalizar-Mac $addr2 }
            }
            & $adbPath -s $serial shell "svc wifi disable" 2>$null
            if (Testar-MacValido $mac) { return $mac }
        } catch {}

        # 6. Estratégia F: Logcat recente do driver
        try {
            $logs = & $adbPath -s $serial shell "logcat -d -t 1500" 2>$null
            if ($logs) {
                $matches = [regex]::Matches($logs, "(?i)(?:wlan|wifi|mac|efuse|nvram|address)[^a-f0-9\r\n]{1,30}([0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2}[:-][0-9a-f]{2})")
                foreach ($m in $matches) {
                    $cand = $m.Groups[1].Value
                    if (Testar-MacValido $cand) { return (Normalizar-Mac $cand) }
                }
            }
        } catch {}

        # 7. Estratégia G: Tela de configurações
        if ($tentativa -ge 3) {
            try {
                & $adbPath -s $serial shell "am start -a android.settings.DEVICE_INFO_SETTINGS" 2>$null
                Start-Sleep -Milliseconds 1200
                $ipL3 = & $adbPath -s $serial shell "ip link show wlan0" 2>$null
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

function Coletar-HardwareAndroidTablet {
    param(
        [string]$adbPath,
        [string]$serial,
        [string]$macPreExistente = $null
    )

    $dataColeta = (Get-Date).ToString("dd/MM/yyyy HH:mm:ss")
    
    # Serial de hardware
    $hwSerial = (& $adbPath -s $serial shell "getprop ro.serialno 2>/dev/null").Trim()
    if (-not $hwSerial) { $hwSerial = (& $adbPath -s $serial shell "getprop ro.boot.serialno 2>/dev/null").Trim() }
    if (-not $hwSerial) { $hwSerial = $serial }

    # Marca / Fabricante
    $marca = (& $adbPath -s $serial shell "getprop ro.product.manufacturer 2>/dev/null").Trim()
    if (-not $marca) { $marca = (& $adbPath -s $serial shell "getprop ro.product.brand 2>/dev/null").Trim() }
    if (-not $marca) { $marca = "Android Tablet" }
    
    # Modelo
    $modelo = (& $adbPath -s $serial shell "getprop ro.product.model 2>/dev/null").Trim()
    if (-not $modelo) { $modelo = "Tablet" }

    # Ano de Fabricacao
    $ano = (Get-Date).Year.ToString()
    try {
        $buildDate = (& $adbPath -s $serial shell "getprop ro.build.date.utc 2>/dev/null").Trim()
        if ($buildDate -and $buildDate -match '^\d+$') {
            $epoch = [DateTimeOffset]::FromUnixTimeSeconds([int64]$buildDate).DateTime.Year
            if ($epoch -ge 2018 -and $epoch -le 2030) { $ano = $epoch.ToString() }
        }
    } catch {}

    # Hostname
    $hostname = (& $adbPath -s $serial shell "getprop net.hostname 2>/dev/null").Trim()
    if (-not $hostname) { $hostname = "TABLET-$hwSerial" }

    # Processador (SoC)
    $soc = (& $adbPath -s $serial shell "getprop ro.board.platform 2>/dev/null").Trim()
    if (-not $soc) { $soc = (& $adbPath -s $serial shell "getprop ro.hardware 2>/dev/null").Trim() }
    $processador = if ($soc) { "ARM SoC ($soc)" } else { "ARM Multi-Core Processor" }

    # Memoria RAM
    $ramStr = "N/A"
    try {
        $memRaw = & $adbPath -s $serial shell "cat /proc/meminfo" 2>$null
        if ($memRaw -match 'MemTotal:\s+(\d+)\s+kB') {
            $kb = [int64]$Matches[1]
            $gb = [Math]::Round($kb / 1024 / 1024, 0)
            if ($gb -lt 1) { $ramStr = "$([Math]::Round($kb / 1024, 0)) MB RAM" }
            else { $ramStr = "$gb GB RAM" }
        }
    } catch {}

    # Armazenamento Interno
    $armStr = "N/A"
    try {
        $dfData = & $adbPath -s $serial shell "df -h /data" 2>$null
        $lines = $dfData -split "\r?\n" | Where-Object { $_ -match '/data' }
        if ($lines.Count -gt 0) {
            $cols = $lines[0] -split '\s+'
            if ($cols.Count -ge 2) { $armStr = "Armazenamento Interno ($($cols[1]))" }
        }
    } catch {}

    # Placa de Video
    $gpu = if ($soc) { "GPU Integrada ($soc)" } else { "GPU Integrada" }

    # Sistema Operacional
    $androidVer = (& $adbPath -s $serial shell "getprop ro.build.version.release 2>/dev/null").Trim()
    $sdkVer = (& $adbPath -s $serial shell "getprop ro.build.version.sdk 2>/dev/null").Trim()
    $so = if ($androidVer) { "Android $androidVer (API $sdkVer)" } else { "Android OS" }

    # Bateria
    $batStr = ""
    try {
        $bat = & $adbPath -s $serial shell "dumpsys battery" 2>$null
        if ($bat -match 'level:\s*(\d+)') {
            $batStr = "$($Matches[1])%"
        }
    } catch {}

    # IP de Rede
    $ip = (& $adbPath -s $serial shell "getprop dhcp.wlan0.ipaddress 2>/dev/null").Trim()
    if (-not $ip -or $ip -notmatch '^\d+\.\d+\.\d+\.\d+$') {
        $ip = "Wi-Fi Desconectado / USB"
    }

    # MAC Wi-Fi de Hardware: verifica se ja possui ou faz varredura com retries
    $mac = "N/A"
    if (Testar-MacValido $macPreExistente) {
        $mac = Normalizar-Mac $macPreExistente
        Write-Host "      -> MAC Wi-Fi: $mac (recuperado do historico)" -ForegroundColor Green
    } else {
        Write-Host "      -> Capturando MAC Wi-Fi com retries automaticos..." -ForegroundColor Yellow
        $mac = Obter-MacComRetries -adbPath $adbPath -serial $serial -maxRetries 6 -onProgress {
            param($msg)
            Write-Host "         ... $msg" -ForegroundColor Gray
        }
        if (Testar-MacValido $mac) {
            try { [Console]::Beep(1400, 200) } catch {}
            Write-Host "      -> [SUCESSO] MAC Wi-Fi: $mac" -ForegroundColor Green
        } else {
            Write-Host "      -> [FALHA] Nao foi possivel obter o MAC." -ForegroundColor Red
        }
    }

    return [PSCustomObject]@{
        Data_Registro       = $dataColeta
        Nome_Computador     = $hostname
        Tipo_Equipamento    = "Tablet Android"
        Marca               = $marca
        Modelo              = $modelo
        Numero_Serie        = $hwSerial
        Ano_Fabricacao      = $ano
        Processador         = $processador
        Memoria_RAM         = $ramStr
        Armazenamento       = $armStr
        Placa_Video         = $gpu
        Sistema_Operacional = $so
        IP_Rede             = $ip
        MAC_Rede            = $mac
        Bateria             = $batStr
    }
}

function Iniciar-CadastroAndroidTablet {
    param([array]$registrosAtuais, [string]$caminhoCsv)

    $adb = Localizar-Adb
    if (-not $adb) {
        try { Clear-Host } catch {}
        Exibir-LogoPrincipal
        Write-Host " |       >>> INVENTARIO DE TABLETS ANDROID - ADB NAO LOCALIZADO <<<       |" -ForegroundColor Red
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host ""
        Write-Host " [!] Executavel ADB (adb.exe) nao foi localizado no sistema ou no Pen Drive." -ForegroundColor Red
        Write-Host "     Certifique-se de que a pasta 'platform-tools' esta presente ou adicione o ADB ao PATH." -ForegroundColor Yellow
        Write-Host ""
        Pausar-Tela
        return
    }

    $loopTablets = $true
    while ($loopTablets) {
        try { Clear-Host } catch {}
        Exibir-LogoPrincipal
        Write-Host " |       >>> NOVO CADASTRO - TABLETS ANDROID (USB / ADB) <<<            |" -ForegroundColor Magenta
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host ""
        Write-Host " [*] Localizacao do ADB: $adb" -ForegroundColor Gray
        Write-Host " [*] Verificando tablets conectados na porta USB..." -ForegroundColor Yellow
        Write-Host ""

        $rawDevices = & $adb devices -l
        $lines = $rawDevices -split "\r?\n"
        $dispositivos = @()

        foreach ($line in $lines) {
            if ($line -match "^([a-zA-Z0-9_-]+)\s+device\b") {
                $dispositivos += $Matches[1]
            } elseif ($line -match "^([a-zA-Z0-9_-]+)\s+unauthorized\b") {
                Write-Host " [!] Tablet [$($Matches[1])] detectado como NAO AUTORIZADO!" -ForegroundColor Red
                Write-Host "     Olhe a tela do tablet e marque 'Sempre permitir deste computador'." -ForegroundColor Yellow
            } elseif ($line -match "^([a-zA-Z0-9_-]+)\s+offline\b") {
                Write-Host " [!] Tablet [$($Matches[1])] detectado como OFFLINE! Reconecte o cabo USB." -ForegroundColor Red
            }
        }

        if ($dispositivos.Count -eq 0) {
            Write-Host " [?] Nenhum tablet Android com Depuracao USB ativa encontrado no momento." -ForegroundColor Yellow
            Write-Host ""
            Write-Host " INSTRUCOES PARA CONEXAO:" -ForegroundColor Cyan
            Write-Host "  1. Conecte o tablet ao computador usando um cabo USB confiavel." -ForegroundColor White
            Write-Host "  2. No tablet, acerte as 'Opcoes do desenvolvedor' e ative 'Depuracao USB'." -ForegroundColor White
            Write-Host "  3. Quando surgir a mensagem na tela do tablet, clique em 'Permitir'." -ForegroundColor White
            Write-Host ""
            Write-Host "  [R] Tentar Novamente / Escanear Novamente" -ForegroundColor Green
            Write-Host "  [0] Voltar ao Menu Principal" -ForegroundColor Gray
            Write-Host ""
            Write-Host "  Escolha " -NoNewline -ForegroundColor Cyan
            $r = Read-Host "[R]"
            if ($r.Trim() -eq "0") {
                $loopTablets = $false
                break
            }
            continue
        }

        Write-Host " [+] $($dispositivos.Count) tablet(s) conectado(s) via USB detectado(s)!`n" -ForegroundColor Green

        # Se houver mais de 1, permite escolher
        $alvoSerial = $dispositivos[0]
        if ($dispositivos.Count -gt 1) {
            Write-Host " Foram encontrados multiplos tablets conectados:" -ForegroundColor Yellow
            for ($k = 0; $k -lt $dispositivos.Count; $k++) {
                $devSerial = $dispositivos[$k]
                $devModel = (& $adb -s $devSerial shell "getprop ro.product.model 2>/dev/null").Trim()
                $devBrand = (& $adb -s $devSerial shell "getprop ro.product.manufacturer 2>/dev/null").Trim()
                Write-Host "   [$($k + 1)] $devSerial - $devBrand $devModel" -ForegroundColor Cyan
            }
            Write-Host "   [0] Cancelar e Voltar" -ForegroundColor Gray
            Write-Host ""
            Write-Host " Escolha o tablet para cadastrar " -NoNewline -ForegroundColor Cyan
            $escolhaDev = Read-Host "[1]"
            if ($escolhaDev.Trim() -eq "0") { break }
            $numDev = 1
            if ([int]::TryParse($escolhaDev.Trim(), [ref]$numDev) -and $numDev -ge 1 -and $numDev -le $dispositivos.Count) {
                $alvoSerial = $dispositivos[$numDev - 1]
            }
        }

        # 1. Verifica se já temos o número de série e MAC deste tablet no inventário ou histórico
        $hwSerialPre = (& $adb -s $alvoSerial shell "getprop ro.serialno 2>/dev/null").Trim()
        if (-not $hwSerialPre) { $hwSerialPre = (& $adb -s $alvoSerial shell "getprop ro.boot.serialno 2>/dev/null").Trim() }
        if (-not $hwSerialPre) { $hwSerialPre = $alvoSerial }

        $macExistente = $null
        $regExistente = $registrosAtuais | Where-Object { 
            ($_.Numero_Serie -and $_.Numero_Serie.Trim().ToUpper() -eq $hwSerialPre.Trim().ToUpper()) -or
            ($_.Numero_Serie -and $_.Numero_Serie.Trim().ToUpper() -eq $alvoSerial.Trim().ToUpper())
        } | Select-Object -First 1

        if ($regExistente -and (Testar-MacValido $regExistente.MAC_Rede)) {
            $macExistente = $regExistente.MAC_Rede
            Write-Host " [INFO] Endereco MAC Wi-Fi ja localizado no inventario para este tablet: $macExistente" -ForegroundColor Cyan
        }

        Write-Host " [*] Coletando informacoes de hardware e MAC Wi-Fi do tablet [$alvoSerial]..." -ForegroundColor Yellow
        $hw = Coletar-HardwareAndroidTablet -adbPath $adb -serial $alvoSerial -macPreExistente $macExistente

        Write-Host "`n  -> Fabricante / Modelo: $($hw.Marca) $($hw.Modelo)" -ForegroundColor White
        Write-Host "  -> Numero de Serie:    $($hw.Numero_Serie)" -ForegroundColor Yellow
        Write-Host "  -> Endereco MAC Wi-Fi: $($hw.MAC_Rede)" -ForegroundColor $(if ($hw.MAC_Rede -ne 'N/A') { 'Green' } else { 'Red' })
        Write-Host "  -> Sistema Operacional: $($hw.Sistema_Operacional)" -ForegroundColor White
        Write-Host "  -> Memoria / Armaz.:   $($hw.Memoria_RAM) | $($hw.Armazenamento)" -ForegroundColor Gray
        if ($hw.Bateria) {
            Write-Host "  -> Nivel de Bateria:   $($hw.Bateria)" -ForegroundColor Cyan
        }
        Write-Host ""

        if ($hw.MAC_Rede -eq 'N/A') {
            Write-Host " [!] O MAC Wi-Fi nao foi obtido automaticamente." -ForegroundColor Red
            Write-Host "     Deseja tentar novamente capturar o MAC agora? (S/N) " -NoNewline -ForegroundColor Yellow
            $respRetryMac = Read-Host "[S]"
            if ($respRetryMac.Trim() -match '^(?i)s|sim|y|yes$|^$') {
                Write-Host " [*] Re-tentando captura do MAC com ativacao forcada de radio..." -ForegroundColor Cyan
                $macNovo = Obter-MacComRetries -adbPath $adb -serial $alvoSerial -maxRetries 6 -onProgress {
                    param($msg)
                    Write-Host "         ... $msg" -ForegroundColor Gray
                }
                if (Testar-MacValido $macNovo) {
                    $hw.MAC_Rede = $macNovo
                    Write-Host " [OK - SUCESSO] MAC Wi-Fi capturado: $macNovo" -ForegroundColor Green
                    try { [Console]::Beep(1400, 200) } catch {}
                } else {
                    Write-Host "     Deseja digitar o MAC manualmente? (Deixe em branco para pular): " -NoNewline -ForegroundColor Gray
                    $macMan = Read-Host
                    if (Testar-MacValido $macMan) {
                        $hw.MAC_Rede = Normalizar-Mac $macMan
                        Write-Host " [OK] MAC Wi-Fi definido manualmente: $($hw.MAC_Rede)" -ForegroundColor Green
                    }
                }
            }
        }

        # Teste de duplicidade de equipamento
        $duplicado = Testar-DuplicidadeEquipamento -registros $registrosAtuais -serial $hw.Numero_Serie -mac $hw.MAC_Rede -patrimonio ""
        if ($duplicado) {
            Write-Host " [ATENCAO] Este tablet JA CONSTA no inventario!" -ForegroundColor Red
            Write-Host "   Patrimonio: $($duplicado.Patrimonio) | Colaborador: $($duplicado.Usuario) ($($duplicado.Setor))" -ForegroundColor Yellow
            Write-Host "   Deseja atualizar o registro existente com os novos dados de hardware? (S/N) " -NoNewline -ForegroundColor Cyan
            $respDup = Read-Host "[S]"
            if ($respDup.Trim() -notmatch '^(?i)s|sim|y|yes$') {
                Write-Host " Operacao cancelada pelo usuario." -ForegroundColor Yellow
                Pausar-Tela
                continue
            }
            $duplicado.Data_Registro       = $hw.Data_Registro
            $duplicado.Tipo_Equipamento    = $hw.Tipo_Equipamento
            $duplicado.Marca               = $hw.Marca
            $duplicado.Modelo              = $hw.Modelo
            $duplicado.Numero_Serie        = $hw.Numero_Serie
            $duplicado.Processador         = $hw.Processador
            $duplicado.Memoria_RAM         = $hw.Memoria_RAM
            $duplicado.Armazenamento       = $hw.Armazenamento
            $duplicado.Sistema_Operacional = $hw.Sistema_Operacional
            $duplicado.MAC_Rede            = $hw.MAC_Rede
            
            Salvar-RegistrosCsv -caminho $caminhoCsv -registros $registrosAtuais
            Write-Host "`n [OK] Registro existente de $($duplicado.Patrimonio) atualizado com sucesso!" -ForegroundColor Green
            Pausar-Tela
            continue
        }

        # Dados Administrativos do Cadastro
        $patrimonioSugerido = Obter-ProximoPatrimonio -registros $registrosAtuais

        $setoresPadrao = @(
            "Operacional", "Pedagogico", "Financeiro", "Marketing", "Comercial",
            "Administrativo", "Administrativo / Financeiro", "Logistica",
            "Motoristas", "Servicos Gerais / Limpeza", "TI", "RH"
        )
        $setoresHistorico = @()
        if ($registrosAtuais.Count -gt 0) {
            $setoresHistorico = @($registrosAtuais | ForEach-Object { $_.Setor.Trim() } | Where-Object { $_ -ne "" -and $_ -notmatch '^\d+$' } | Select-Object -Unique)
        }
        $listaSetores = @()
        foreach ($s in $setoresHistorico) { if ($listaSetores -notcontains $s) { $listaSetores += $s } }
        foreach ($s in $setoresPadrao) { if ($listaSetores -notcontains $s) { $listaSetores += $s } }

        # [1] EMPRESA
        Write-Host " [1/5] IDENTIFICACAO DA EMPRESA" -ForegroundColor Yellow
        Write-Host "   Pressione ENTER para manter " -NoNewline -ForegroundColor Gray
        Write-Host "'InnovTech'" -ForegroundColor Green
        $empresaInput = Read-Host "   Empresa"
        $empresa = if ([string]::IsNullOrWhiteSpace($empresaInput)) { "InnovTech" } else { $empresaInput.Trim() }

        # [2] USUARIO
        Write-Host "`n [2/5] DADOS DO USUARIO / RESPONSAVEL / ALUNO" -ForegroundColor Yellow
        $usuario = ""
        while ([string]::IsNullOrWhiteSpace($usuario)) {
            $usuario = Read-Host "   Nome do Responsavel ou Setor"
            if ([string]::IsNullOrWhiteSpace($usuario)) {
                Write-Host "   (!) Por favor, digite o nome do usuario ou responsavel." -ForegroundColor Red
            }
        }

        # [3] FUNCAO
        Write-Host "`n [3/5] FUNCAO / CARGO" -ForegroundColor Yellow
        Write-Host "   Pressione ENTER para manter " -NoNewline -ForegroundColor Gray
        Write-Host "'Operador de Tablet'" -ForegroundColor Green
        $funcaoInput = Read-Host "   Funcao / Cargo"
        $funcao = if ([string]::IsNullOrWhiteSpace($funcaoInput)) { "Operador de Tablet" } else { $funcaoInput.Trim() }

        # [4] SETOR
        Write-Host "`n [4/5] SETOR / SALA / DEPARTAMENTO" -ForegroundColor Yellow
        for ($i = 0; $i -lt $listaSetores.Count; $i++) {
            $numFormatado = ($i + 1).ToString().PadLeft(2, ' ')
            $ehHistorico = ($setoresHistorico -contains $listaSetores[$i])
            $corSetor = if ($ehHistorico) { "Green" } else { "White" }
            $qtdSetor = @($registrosAtuais | Where-Object { $_.Setor -eq $listaSetores[$i] }).Count
            $tagQtd = if ($qtdSetor -gt 0) { " ($qtdSetor Ativos)" } else { "" }
            Write-Host "     [$numFormatado] " -NoNewline -ForegroundColor Cyan
            Write-Host "$($listaSetores[$i])$tagQtd" -ForegroundColor $corSetor
        }
        Write-Host "     [ 0] Digitar um outro setor novo" -ForegroundColor Gray
        
        $setor = ""
        while ([string]::IsNullOrWhiteSpace($setor)) {
            $setorInput = Read-Host "`n   Escolha o numero ou digite o nome do setor"
            if ([string]::IsNullOrWhiteSpace($setorInput)) {
                Write-Host "   (!) O setor nao pode ficar vazio." -ForegroundColor Red
                continue
            }
            $numSetor = 0
            if ([int]::TryParse($setorInput.Trim(), [ref]$numSetor)) {
                if ($numSetor -ge 1 -and $numSetor -le $listaSetores.Count) {
                    $setor = $listaSetores[$numSetor - 1]
                } elseif ($numSetor -eq 0) {
                    while ([string]::IsNullOrWhiteSpace($setor)) {
                        $setor = Read-Host "   Digite o nome do novo setor"
                    }
                } else {
                    $setor = $setorInput.Trim()
                }
            } else {
                $setor = $setorInput.Trim()
            }
        }

        # [5] PATRIMONIO
        Write-Host "`n [5/5] NUMERO DE PATRIMONIO" -ForegroundColor Yellow
        Write-Host "   Pressione ENTER para aceitar o sugerido: " -NoNewline -ForegroundColor Gray
        Write-Host "[$patrimonioSugerido]" -ForegroundColor Green
        $patrimonioInput = Read-Host "   Patrimonio"
        $patrimonio = if ([string]::IsNullOrWhiteSpace($patrimonioInput)) { $patrimonioSugerido } else { $patrimonioInput.Trim() }

        # Criacao do Objeto Novo
        $novoItem = [PSCustomObject]@{
            Data_Registro       = $hw.Data_Registro
            Empresa             = $empresa
            Patrimonio          = $patrimonio
            Usuario             = $usuario
            Funcao              = $funcao
            Setor               = $setor
            Tipo_Equipamento    = $hw.Tipo_Equipamento
            Marca               = $hw.Marca
            Modelo              = $hw.Modelo
            Numero_Serie        = $hw.Numero_Serie
            Ano_Fabricacao      = $hw.Ano_Fabricacao
            Nome_Computador     = $hw.Nome_Computador
            Processador         = $hw.Processador
            Memoria_RAM         = $hw.Memoria_RAM
            Armazenamento       = $hw.Armazenamento
            Placa_Video         = $hw.Placa_Video
            Sistema_Operacional = $hw.Sistema_Operacional
            IP_Rede             = $hw.IP_Rede
            MAC_Rede            = $hw.MAC_Rede
        }

        $listaAtualizada = @($registrosAtuais) + $novoItem
        Salvar-RegistrosCsv -caminho $caminhoCsv -registros $listaAtualizada
        $registrosAtuais = $listaAtualizada

        # Copia resumo para area de transferencia
        $resumoTexto = "$($novoItem.Patrimonio);$($novoItem.Usuario);$($novoItem.Setor);$($novoItem.Marca);$($novoItem.Modelo);$($novoItem.Numero_Serie);$($novoItem.MAC_Rede)"
        try { Set-Clipboard -Value $resumoTexto } catch { try { $resumoTexto | clip.exe } catch {} }

        Write-Host "`n [OK] Tablet cadastrado com sucesso no inventario do Pen Drive!" -ForegroundColor Green
        Write-Host "      Patrimonio: $patrimonio | N/S: $($hw.Numero_Serie) | MAC: $($hw.MAC_Rede)" -ForegroundColor Yellow
        Write-Host " [OK] Informacoes copiadas para a Area de Transferencia!" -ForegroundColor Green
        Write-Host ""
        Write-Host " Deseja cadastrar outro tablet agora? (S/N) " -NoNewline -ForegroundColor Cyan
        $outro = Read-Host "[S]"
        if ($outro.Trim() -notmatch '^(?i)s|sim|y|yes$') {
            $loopTablets = $false
        }
    }
}

# ==============================================================================
# MOTOR DE ENTRADA MANUAL COMPLETA (NOVO RECURSO)
# ==============================================================================
function Iniciar-CadastroManual {
    param([array]$registrosAtuais, [string]$caminhoCsv)
    
    try { Clear-Host } catch {}
    Exibir-LogoPrincipal
    Write-Host " |             >>> NOVO CADASTRO - ENTRADA MANUAL <<<                   |" -ForegroundColor Yellow
    Write-Host " +======================================================================+" -ForegroundColor Cyan
    Write-Host "  (Preencha as informacoes de qualquer computador ou equipamento)`n" -ForegroundColor Gray
    
    # 1. Proximo numero de patrimonio sequencial inteligente
    $patrimonioSugerido = Obter-ProximoPatrimonio -registros $registrosAtuais
    
    # 2. Lista Completa de Setores
    $setoresPadrao = @(
        "Operacional", "Pedagogico", "Financeiro", "Marketing", "Comercial",
        "Administrativo", "Administrativo / Financeiro", "Logistica",
        "Motoristas", "Servicos Gerais / Limpeza", "TI", "RH"
    )
    $setoresHistorico = @()
    if ($registrosAtuais.Count -gt 0) {
        $setoresHistorico = @($registrosAtuais | ForEach-Object { $_.Setor.Trim() } | Where-Object { $_ -ne "" -and $_ -notmatch '^\d+$' } | Select-Object -Unique)
    }
    $listaSetores = @()
    foreach ($s in $setoresHistorico) { if ($listaSetores -notcontains $s) { $listaSetores += $s } }
    foreach ($s in $setoresPadrao) { if ($listaSetores -notcontains $s) { $listaSetores += $s } }
    
    # [1] EMPRESA
    Write-Host " [1/14] IDENTIFICACAO DA EMPRESA" -ForegroundColor Yellow
    Write-Host "   Pressione ENTER para manter " -NoNewline -ForegroundColor Gray
    Write-Host "'InnovTech'" -ForegroundColor Green
    $empresaInput = Read-Host "   Empresa"
    $empresa = if ([string]::IsNullOrWhiteSpace($empresaInput)) { "InnovTech" } else { $empresaInput.Trim() }
    
    # [2] COLABORADOR
    Write-Host "`n [2/14] DADOS DO USUARIO / RESPONSAVEL" -ForegroundColor Yellow
    $usuario = ""
    while ([string]::IsNullOrWhiteSpace($usuario)) {
        $usuario = Read-Host "   Nome Completo do Colaborador"
        if ([string]::IsNullOrWhiteSpace($usuario)) {
            Write-Host "   (!) Por favor, digite o nome do colaborador." -ForegroundColor Red
        }
    }
    
    # [3] FUNCAO / CARGO
    Write-Host "`n [3/14] FUNCAO / CARGO" -ForegroundColor Yellow
    $funcao = ""
    while ([string]::IsNullOrWhiteSpace($funcao)) {
        $funcao = Read-Host "   Cargo / Funcao (Ex: Professor, Operador, Motorista, Analista)"
        if ([string]::IsNullOrWhiteSpace($funcao)) {
            Write-Host "   (!) Por favor, digite a funcao." -ForegroundColor Red
        }
    }
    
    # [4] SETOR
    Write-Host "`n [4/14] SETOR / DEPARTAMENTO" -ForegroundColor Yellow
    Write-Host "   Escolha um setor da lista ou digite um novo:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $listaSetores.Count; $i++) {
        $numFormatado = ($i + 1).ToString().PadLeft(2, ' ')
        $ehHistorico = ($setoresHistorico -contains $listaSetores[$i])
        $corSetor = if ($ehHistorico) { "Green" } else { "White" }
        $qtdSetor = @($registrosAtuais | Where-Object { $_.Setor -eq $listaSetores[$i] }).Count
        $tagQtd = if ($qtdSetor -gt 0) { " ($qtdSetor PCs)" } else { "" }
        Write-Host "     [$numFormatado] " -NoNewline -ForegroundColor Cyan
        Write-Host "$($listaSetores[$i])$tagQtd" -ForegroundColor $corSetor
    }
    Write-Host "     [ 0] Digitar um outro setor novo" -ForegroundColor Gray
    
    $setor = ""
    while ([string]::IsNullOrWhiteSpace($setor)) {
        $setorInput = Read-Host "`n   Escolha o numero ou digite o nome do setor"
        if ([string]::IsNullOrWhiteSpace($setorInput)) {
            Write-Host "   (!) O setor nao pode ficar vazio." -ForegroundColor Red
            continue
        }
        $numSetor = 0
        if ([int]::TryParse($setorInput.Trim(), [ref]$numSetor)) {
            if ($numSetor -ge 1 -and $numSetor -le $listaSetores.Count) {
                $setor = $listaSetores[$numSetor - 1]
            } elseif ($numSetor -eq 0) {
                while ([string]::IsNullOrWhiteSpace($setor)) {
                    $setor = Read-Host "   Digite o nome do novo setor"
                }
            } else {
                $setor = $setorInput.Trim()
            }
        } else {
            $setor = $setorInput.Trim()
        }
    }
    Write-Host "   -> Setor selecionado: " -NoNewline -ForegroundColor Gray
    Write-Host "$setor" -ForegroundColor Green
    
    # [5] PATRIMONIO
    Write-Host "`n [5/14] NUMERO DE PATRIMONIO" -ForegroundColor Yellow
    Write-Host "   Pressione ENTER para aceitar o sugerido: " -NoNewline -ForegroundColor Gray
    Write-Host "[$patrimonioSugerido]" -ForegroundColor Green
    $patrimonioInput = Read-Host "   Patrimonio"
    $patrimonio = if ([string]::IsNullOrWhiteSpace($patrimonioInput)) { $patrimonioSugerido } else { $patrimonioInput.Trim() }
    
    # [6] TIPO DE EQUIPAMENTO
    Write-Host "`n [6/14] TIPO DE EQUIPAMENTO" -ForegroundColor Yellow
    Write-Host "   [1] Notebook / Laptop" -ForegroundColor White
    Write-Host "   [2] Desktop / PC" -ForegroundColor White
    Write-Host "   [3] All-in-One" -ForegroundColor White
    Write-Host "   [4] Servidor" -ForegroundColor White
    Write-Host "   [5] Outro (Digitar personalizado)" -ForegroundColor Gray
    $tipoOpcao = Read-Host "   Escolha o tipo [1]"
    $tipoEquipamento = switch ($tipoOpcao.Trim()) {
        "2" { "Desktop / PC" }
        "3" { "All-in-One" }
        "4" { "Servidor" }
        "5" { 
            $customTipo = Read-Host "   Digite o tipo de equipamento"
            if ([string]::IsNullOrWhiteSpace($customTipo)) { "Computador" } else { $customTipo.Trim() }
        }
        default { "Notebook / Laptop" }
    }
    
    # [7] MARCA / FABRICANTE
    Write-Host "`n [7/14] FABRICANTE / MARCA" -ForegroundColor Yellow
    $marcaInput = Read-Host "   Marca (Ex: Dell, Lenovo, HP, Samsung, Positivo, Acer, Apple)"
    $marca = if ([string]::IsNullOrWhiteSpace($marcaInput)) { "N/A" } else { $marcaInput.Trim() }
    
    # [8] MODELO
    Write-Host "`n [8/14] MODELO DO EQUIPAMENTO" -ForegroundColor Yellow
    $modeloInput = Read-Host "   Modelo (Ex: Latitude 3420, ThinkCentre M70, Inspiron 15)"
    $modelo = if ([string]::IsNullOrWhiteSpace($modeloInput)) { "N/A" } else { $modeloInput.Trim() }
    
    # [9] NUMERO DE SERIE
    Write-Host "`n [9/14] NUMERO DE SERIE / SERVICE TAG" -ForegroundColor Yellow
    $serialInput = Read-Host "   Numero de Serie / Service Tag"
    $numeroSerie = if ([string]::IsNullOrWhiteSpace($serialInput)) { "N/A" } else { $serialInput.Trim() }
    
    # [10] ANO DE FABRICACAO
    Write-Host "`n [10/14] ANO DE FABRICACAO" -ForegroundColor Yellow
    $anoPadrao = (Get-Date).Year.ToString()
    Write-Host "   Pressione ENTER para ano atual: " -NoNewline -ForegroundColor Gray
    Write-Host "[$anoPadrao]" -ForegroundColor Green
    $anoInput = Read-Host "   Ano"
    $anoFabricacao = if ([string]::IsNullOrWhiteSpace($anoInput)) { $anoPadrao } else { $anoInput.Trim() }
    
    # [11] HOSTNAME
    Write-Host "`n [11/14] NOME DO COMPUTADOR (HOSTNAME)" -ForegroundColor Yellow
    $hostInput = Read-Host "   Hostname (Ex: TI-NOTE-01, FIN-PC-02)"
    $nomeComputador = if ([string]::IsNullOrWhiteSpace($hostInput)) { "N/A" } else { $hostInput.Trim() }
    
    # [12] PROCESSADOR (CPU)
    Write-Host "`n [12/14] PROCESSADOR (CPU)" -ForegroundColor Yellow
    $cpuInput = Read-Host "   Processador (Ex: Intel Core i5-1135G7, AMD Ryzen 5 5500U)"
    $processador = if ([string]::IsNullOrWhiteSpace($cpuInput)) { "N/A" } else { $cpuInput.Trim() }
    
    # [13] MEMORIA RAM
    Write-Host "`n [13/14] MEMORIA RAM" -ForegroundColor Yellow
    $ramInput = Read-Host "   Memoria RAM (Ex: 8 GB RAM, 16 GB RAM DDR4)"
    $memoriaRam = if ([string]::IsNullOrWhiteSpace($ramInput)) { "N/A" } else { $ramInput.Trim() }
    
    # [14] ARMAZENAMENTO & SISTEMA
    Write-Host "`n [14/14] ARMAZENAMENTO & SISTEMA OPERACIONAL" -ForegroundColor Yellow
    $diskInput = Read-Host "   Armazenamento (Ex: [SSD] 256GB NVMe, [SSD] 512GB, 1TB HDD)"
    $armazenamento = if ([string]::IsNullOrWhiteSpace($diskInput)) { "N/A" } else { $diskInput.Trim() }
    
    $gpuInput = Read-Host "   Placa de Video (Opcional - ENTER para Integrada/N/A)"
    $placaVideo = if ([string]::IsNullOrWhiteSpace($gpuInput)) { "Integrada / N/A" } else { $gpuInput.Trim() }
    
    $osInput = Read-Host "   Sistema Operacional (Ex: Windows 11 Pro 64-bit, Windows 10)"
    $sistemaOperacional = if ([string]::IsNullOrWhiteSpace($osInput)) { "Windows 11 Pro (64-bit)" } else { $osInput.Trim() }
    
    $ipInput = Read-Host "   Endereco IP (Opcional - ENTER para N/A)"
    $ipRede = if ([string]::IsNullOrWhiteSpace($ipInput)) { "N/A" } else { $ipInput.Trim() }
    
    $macInput = Read-Host "   Endereco MAC (Opcional - ENTER para N/A)"
    $macRede = if ([string]::IsNullOrWhiteSpace($macInput)) { "N/A" } else { $macInput.Trim() }
    
    $dataRegistro = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    
    # RESUMO COMPLETO NA TELA
    $divisoria = "=" * 64
    $resumoTexto = @(
        $divisoria,
        "       INNOVCORE TI - RESUMO DO INVENTARIO MANUAL ($empresa)",
        $divisoria,
        "  - Empresa:            $empresa",
        "  - Patrimonio:         $patrimonio",
        "  - Usuario:            $usuario",
        "  - Funcao / Cargo:     $funcao",
        "  - Setor:              $setor",
        "",
        "  - Tipo Equipamento:   $tipoEquipamento",
        "  - Marca:              $marca",
        "  - Modelo:             $modelo",
        "  - Numero de Serie:    $numeroSerie",
        "  - Ano de Fabricacao:  $anoFabricacao",
        "  - Computador (Host):  $nomeComputador",
        "  - Processador (CPU):  $processador",
        "  - Memoria RAM:        $memoriaRam",
        "  - Armazenamento:      $armazenamento",
        "  - Placa de Video:     $placaVideo",
        "  - Sistema Operacional:$sistemaOperacional",
        "  - Endereco IP / MAC:  $ipRede | $macRede",
        "  - Data do Registro:   $dataRegistro",
        $divisoria
    ) -join "`r`n"
    
    try { Clear-Host } catch {}
    Exibir-LogoPrincipal
    Write-Host ""
    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
    Write-Host " |              RESUMO DAS ESPECIFICACOES (CADASTRO MANUAL)             |" -ForegroundColor Yellow
    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
    Write-Host "   EMPRESA:          " -NoNewline -ForegroundColor Gray; Write-Host "$empresa" -ForegroundColor White
    Write-Host "   PATRIMONIO:       " -NoNewline -ForegroundColor Gray; Write-Host "$patrimonio" -ForegroundColor Green
    Write-Host "   COLABORADOR:      " -NoNewline -ForegroundColor Gray; Write-Host "$usuario" -ForegroundColor White
    Write-Host "   CARGO / FUNCAO:   " -NoNewline -ForegroundColor Gray; Write-Host "$funcao" -ForegroundColor White
    Write-Host "   SETOR:            " -NoNewline -ForegroundColor Gray; Write-Host "$setor" -ForegroundColor Green
    Write-Host "  ----------------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "   TIPO EQUIPAMENTO: " -NoNewline -ForegroundColor Gray; Write-Host "$tipoEquipamento" -ForegroundColor Cyan
    Write-Host "   FABRICANTE/MARCA: " -NoNewline -ForegroundColor Gray; Write-Host "$marca" -ForegroundColor Cyan
    Write-Host "   MODELO:           " -NoNewline -ForegroundColor Gray; Write-Host "$modelo" -ForegroundColor Cyan
    Write-Host "   NUMERO DE SERIE:  " -NoNewline -ForegroundColor Gray; Write-Host "$numeroSerie" -ForegroundColor Yellow
    Write-Host "   ANO FABRICACAO:   " -NoNewline -ForegroundColor Gray; Write-Host "$anoFabricacao" -ForegroundColor Yellow
    Write-Host "   NOME DO PC (HOST):" -NoNewline -ForegroundColor Gray; Write-Host "$nomeComputador" -ForegroundColor White
    Write-Host "   PROCESSADOR:      " -NoNewline -ForegroundColor Gray; Write-Host "$processador" -ForegroundColor White
    Write-Host "   MEMORIA RAM:      " -NoNewline -ForegroundColor Gray; Write-Host "$memoriaRam" -ForegroundColor White
    Write-Host "   ARMAZENAMENTO:    " -NoNewline -ForegroundColor Gray; Write-Host "$armazenamento" -ForegroundColor White
    Write-Host "   DATA DO CADASTRO: " -NoNewline -ForegroundColor Gray; Write-Host "$dataRegistro" -ForegroundColor Gray
    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
    
    # Copiar para Clipboard
    try { Set-Clipboard -Value $resumoTexto } catch { try { $resumoTexto | clip.exe } catch {} }
    Write-Host ""
    Write-Host " [OK] Informacoes COPIADAS para a Area de Transferencia (Ctrl + V)!" -ForegroundColor Green
    
    # Salvar no CSV
    $novoItem = [PSCustomObject]@{
        Data_Registro       = $dataRegistro
        Empresa             = $empresa
        Patrimonio          = $patrimonio
        Usuario             = $usuario
        Funcao              = $funcao
        Setor               = $setor
        Tipo_Equipamento    = $tipoEquipamento
        Marca               = $marca
        Modelo              = $modelo
        Numero_Serie        = $numeroSerie
        Ano_Fabricacao      = $anoFabricacao
        Nome_Computador     = $nomeComputador
        Processador         = $processador
        Memoria_RAM         = $memoriaRam
        Armazenamento       = $armazenamento
        Placa_Video         = $placaVideo
        Sistema_Operacional = $sistemaOperacional
        IP_Rede             = $ipRede
        MAC_Rede            = $macRede
    }
    
    # Verificacao de Duplicidade Inteligente
    $dup = Testar-DuplicidadeEquipamento -registros $registrosAtuais -patrimonio $novoItem.Patrimonio -numeroSerie $novoItem.Numero_Serie -macRede $novoItem.MAC_Rede -nomeComputador $novoItem.Nome_Computador
    
    if ($dup) {
        $itemAntigo = $dup.RegistroExistente
        Write-Host ""
        Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Yellow
        Write-Host " |               [!] ALERTA: DUPLICIDADE DETECTADA                      |" -ForegroundColor Yellow
        Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Yellow
        Write-Host "   Motivo:             " -NoNewline -ForegroundColor Gray; Write-Host "$($dup.Motivo)" -ForegroundColor Red
        Write-Host "   Registro Anterior:  " -NoNewline -ForegroundColor Gray; Write-Host "$($itemAntigo.Patrimonio) - $($itemAntigo.Usuario) ($($itemAntigo.Setor))" -ForegroundColor Cyan
        Write-Host "   Data Anterior:      " -NoNewline -ForegroundColor Gray; Write-Host "$($itemAntigo.Data_Registro)" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "   Como deseja proceder?" -ForegroundColor Yellow
        Write-Host "     [1] ATUALIZAR cadastro existente (Sobrescrever dados com backup)" -ForegroundColor Green
        Write-Host "     [2] MANTER AMBOS e gerar novo numero de patrimonio" -ForegroundColor Cyan
        Write-Host "     [0] CANCELAR gravacao deste cadastro" -ForegroundColor Gray
        Write-Host ""
        $opcDup = Read-Host "   Escolha uma opcao [1]"
        if ([string]::IsNullOrWhiteSpace($opcDup)) { $opcDup = "1" }
        
        if ($opcDup.Trim() -eq "1") {
            $novoItem.Patrimonio = $itemAntigo.Patrimonio
            $listaAtualizada = @($registrosAtuais | Where-Object { $_.Patrimonio -ne $itemAntigo.Patrimonio }) + $novoItem
            Salvar-RegistrosCsv -caminho $caminhoCsv -registros $listaAtualizada
            Write-Host "`n [OK] Cadastro $($novoItem.Patrimonio) ATUALIZADO com sucesso no Pen Drive!" -ForegroundColor Green
            Pausar-Tela -mensagem " Pressione ENTER para voltar ao menu..."
            return
        } elseif ($opcDup.Trim() -eq "2") {
            $novoPat = Obter-ProximoPatrimonio -registros $registrosAtuais
            $novoItem.Patrimonio = $novoPat
            $listaAtualizada = @($registrosAtuais) + $novoItem
            Salvar-RegistrosCsv -caminho $caminhoCsv -registros $listaAtualizada
            Write-Host "`n [OK] Gravado como NOVO equipamento com patrimonio: $novoPat!" -ForegroundColor Green
            Pausar-Tela -mensagem " Pressione ENTER para voltar ao menu..."
            return
        } else {
            Write-Host "`n [!] Operacao cancelada pelo usuario. Nenhum dado alterado." -ForegroundColor Yellow
            Pausar-Tela -mensagem " Pressione ENTER para voltar ao menu..."
            return
        }
    }
    
    $listaAtualizada = @($registrosAtuais) + $novoItem
    Salvar-RegistrosCsv -caminho $caminhoCsv -registros $listaAtualizada
    
    Write-Host " [OK] Cadastro MANUAL gravado com sucesso na planilha do Pen Drive!" -ForegroundColor Green
    Write-Host "      Arquivo: " -NoNewline -ForegroundColor Gray
    Write-Host "$caminhoCsv" -ForegroundColor Yellow
    
    Pausar-Tela -mensagem " Cadastro manual concluido! Pressione ENTER para voltar ao menu..."
}

# ==============================================================================
# MOTOR DE EXPORTACAO EXCEL UTF-8 / HTML / RELATORIOS
# ==============================================================================
function Exportar-InventarioExcel {
    param(
        [array]$registros,
        [string]$tipo = "CSV_EXCEL",
        [string]$filtroSetor = "",
        [switch]$abrirAposExportar
    )
    
    if (!(Test-Path $script:pastaExport)) {
        New-Item -ItemType Directory -Path $script:pastaExport -Force | Out-Null
    }
    
    $registrosFiltrados = if ([string]::IsNullOrWhiteSpace($filtroSetor)) {
        $registros
    } else {
        @($registros | Where-Object { $_.Setor -eq $filtroSetor })
    }
    
    if ($registrosFiltrados.Count -eq 0) {
        Write-Host " [!] Nenhum registro encontrado para exportacao com os criterios definidos." -ForegroundColor Yellow
        return $null
    }
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $sufixoFiltro = if ($filtroSetor) { "_$($filtroSetor -replace '[^\w]', '_')" } else { "_Geral" }
    
    switch ($tipo) {
        "CSV_EXCEL" {
            # Planilha CSV UTF-8 com BOM e delimitador ';' (Abre direto no Excel PT-BR sem erros de codificacao)
            $nomeArquivo = "Inventario_InnovCore$($sufixoFiltro)_$($timestamp).csv"
            $caminhoArquivo = Join-Path $script:pastaExport $nomeArquivo
            
            $limparCampo = { param($c) if ($null -ne $c) { ($c.ToString() -replace ';', ',') -replace "[\r\n]+", " " } else { "" } }
            
            $linhasExport = [System.Collections.Generic.List[string]]::new()
            $linhasExport.Add($script:cabecalhoCsvString)
            foreach ($r in $registrosFiltrados) {
                $linha = @(
                    (& $limparCampo $r.Data_Registro),
                    (& $limparCampo $r.Empresa),
                    (& $limparCampo $r.Patrimonio),
                    (& $limparCampo $r.Usuario),
                    (& $limparCampo $r.Funcao),
                    (& $limparCampo $r.Setor),
                    (& $limparCampo $r.Tipo_Equipamento),
                    (& $limparCampo $r.Marca),
                    (& $limparCampo $r.Modelo),
                    (& $limparCampo $r.Numero_Serie),
                    (& $limparCampo $r.Ano_Fabricacao),
                    (& $limparCampo $r.Nome_Computador),
                    (& $limparCampo $r.Processador),
                    (& $limparCampo $r.Memoria_RAM),
                    (& $limparCampo $r.Armazenamento),
                    (& $limparCampo $r.Placa_Video),
                    (& $limparCampo $r.Sistema_Operacional),
                    (& $limparCampo $r.IP_Rede),
                    (& $limparCampo $r.MAC_Rede)
                ) -join ';'
                $linhasExport.Add($linha)
            }
            
            [System.IO.File]::WriteAllText($caminhoArquivo, (($linhasExport -join "`r`n") + "`r`n"), $script:utf8ComBOM)
            
            Write-Host "`n [OK] Planilha Excel CSV UTF-8 gerada com sucesso!" -ForegroundColor Green
            Write-Host "      Arquivo: " -NoNewline -ForegroundColor Gray
            Write-Host "$caminhoArquivo" -ForegroundColor Yellow
            Write-Host "      Total de maquinas exportadas: " -NoNewline -ForegroundColor Gray
            Write-Host "$($registrosFiltrados.Count)" -ForegroundColor Green
            
            if ($abrirAposExportar) {
                try { Invoke-Item $caminhoArquivo } catch {}
            }
            return $caminhoArquivo
        }
        
        "HTML_EXCEL" {
            # Planilha Formatada em HTML Corporativo InnovCore (Abre perfeitamente no Excel e em qualquer Navegador)
            $nomeArquivo = "Relatorio_InnovCore_Formatado$($sufixoFiltro)_$($timestamp).xls"
            $caminhoArquivo = Join-Path $script:pastaExport $nomeArquivo
            
            $dataGeracao = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
            $totalMaquinas = $registrosFiltrados.Count
            
            $linhasHtml = ""
            $num = 1
            foreach ($r in $registrosFiltrados) {
                $bgClass = if ($num % 2 -eq 0) { "even" } else { "odd" }
                $linhasHtml += "                <tr class=""$bgClass"">`r`n"
                $linhasHtml += "                    <td align=""center""><b>$num</b></td>`r`n"
                $linhasHtml += "                    <td class=""pat"">$($r.Patrimonio)</td>`r`n"
                $linhasHtml += "                    <td><b>$($r.Usuario)</b></td>`r`n"
                $linhasHtml += "                    <td>$($r.Setor)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Funcao)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Marca)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Modelo)</td>`r`n"
                $linhasHtml += "                    <td class=""serial"">$($r.Numero_Serie)</td>`r`n"
                $linhasHtml += "                    <td align=""center"">$($r.Ano_Fabricacao)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Nome_Computador)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Processador)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Memoria_RAM)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Armazenamento)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Placa_Video)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Sistema_Operacional)</td>`r`n"
                $linhasHtml += "                    <td>$($r.IP_Rede)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Data_Registro)</td>`r`n"
                $linhasHtml += "                </tr>`r`n"
                $num++
            }
            
            $escopoTxt = if ($filtroSetor) { "Setor $filtroSetor" } else { "Inventario Completo" }
            $conteudoHtml = @(
                '<!DOCTYPE html>',
                '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">',
                '<head>',
                '    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">',
                '    <style>',
                '        body { font-family: Arial, sans-serif; font-size: 11pt; color: #1e293b; background-color: #f8fafc; }',
                '        .header-box { background-color: #0f172a; color: #ffffff; padding: 18px 24px; border-radius: 8px; margin-bottom: 20px; }',
                '        .header-title { font-size: 20pt; font-weight: bold; margin: 0; color: #38bdf8; }',
                '        .header-subtitle { font-size: 11pt; color: #cbd5e1; margin-top: 5px; }',
                '        .stats-bar { background-color: #e2e8f0; padding: 10px 15px; border-radius: 6px; font-weight: bold; margin-bottom: 15px; }',
                '        table { border-collapse: collapse; width: 100%; border: 1px solid #cbd5e1; background-color: #ffffff; }',
                '        th { background-color: #0284c7; color: #ffffff; font-weight: bold; padding: 10px 8px; font-size: 10.5pt; text-align: left; border: 1px solid #0369a1; }',
                '        td { padding: 8px; border: 1px solid #e2e8f0; font-size: 9.5pt; }',
                '        tr.even { background-color: #f8fafc; }',
                '        tr.odd { background-color: #ffffff; }',
                '        .pat { color: #0284c7; font-weight: bold; }',
                '        .serial { color: #b45309; font-weight: 600; }',
                '        .footer { font-size: 9pt; color: #64748b; margin-top: 20px; text-align: right; }',
                '    </style>',
                '</head>',
                '<body>',
                '    <div class="header-box">',
                "        <h1 class=""header-title"">INNOVCORE TI - RELATORIO DE INVENTARIO E ATIVOS</h1>",
                "        <div class=""header-subtitle"">Empresa: InnovTech | Gerado em: $dataGeracao | Escopo: $escopoTxt</div>",
                '    </div>',
                '    <div class="stats-bar">',
                "        TOTAL DE EQUIPAMENTOS CATALOGADOS: $totalMaquinas | SISTEMA: INNOVCORE TI v2.0",
                '    </div>',
                '    <table>',
                '        <thead>',
                '            <tr>',
                '                <th>#</th>',
                '                <th>Patrimonio</th>',
                '                <th>Colaborador</th>',
                '                <th>Setor</th>',
                '                <th>Cargo / Funcao</th>',
                '                <th>Marca</th>',
                '                <th>Modelo</th>',
                '                <th>N Serie</th>',
                '                <th>Ano</th>',
                '                <th>Hostname</th>',
                '                <th>Processador</th>',
                '                <th>Memoria RAM</th>',
                '                <th>Armazenamento</th>',
                '                <th>Placa de Video</th>',
                '                <th>Sistema Operacional</th>',
                '                <th>IPv4</th>',
                '                <th>Data Coleta</th>',
                '            </tr>',
                '        </thead>',
                '        <tbody>',
                $linhasHtml,
                '        </tbody>',
                '    </table>',
                '    <div class="footer">',
                '        Relatorio gerado automaticamente por <b>InnovCore TI v2.0</b> - Sistema Corporativo de Gestao de Ativos.',
                '    </div>',
                '</body>',
                '</html>'
            ) -join "`r`n"
            [System.IO.File]::WriteAllText($caminhoArquivo, $conteudoHtml, $script:utf8ComBOM)
            
            Write-Host "`n [OK] Planilha Formatada para Excel (.xls/HTML) gerada com sucesso!" -ForegroundColor Green
            Write-Host "      Arquivo: " -NoNewline -ForegroundColor Gray
            Write-Host "$caminhoArquivo" -ForegroundColor Yellow
            Write-Host "      Visual corporativo de alta qualidade pronto para apresentacoes." -ForegroundColor Cyan
            
            if ($abrirAposExportar) {
                try { Invoke-Item $caminhoArquivo } catch {}
            }
            return $caminhoArquivo
        }
        
        "TXT_RELATORIO" {
            # Relatorio Executivo em Texto Formatado
            $nomeArquivo = "Relatorio_Executivo_InnovCore$($sufixoFiltro)_$($timestamp).txt"
            $caminhoArquivo = Join-Path $script:pastaExport $nomeArquivo
            
            $dataGeracao = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
            $divisoria = "=" * 78
            $subdivisoria = "-" * 78
            $escopoTxt = if ($filtroSetor) { "Setor $filtroSetor" } else { "Inventario Geral Completo" }
            $totalQtd = $registrosFiltrados.Count
            
            $txt = @(
                $divisoria,
                "           INNOVCORE TI - RELATORIO CONSOLIDADO DE GESTAO DE ATIVOS",
                $divisoria,
                " Data de Emissao: $dataGeracao",
                " Total de Maquinas: $totalQtd",
                " Escopo: $escopoTxt",
                $divisoria,
                ""
            ) -join "`r`n"
            
            $i = 1
            foreach ($r in $registrosFiltrados) {
                $txt += "[$i] PATRIMONIO: $($r.Patrimonio) | COLABORADOR: $($r.Usuario)`r`n"
                $txt += "$subdivisoria`r`n"
                $txt += "  - Empresa:            $($r.Empresa)`r`n"
                $txt += "  - Setor:              $($r.Setor)`r`n"
                $txt += "  - Cargo / Funcao:     $($r.Funcao)`r`n"
                $txt += "  - Tipo Equipamento:   $($r.Tipo_Equipamento)`r`n"
                $txt += "  - Fabricante / Marca: $($r.Marca)`r`n"
                $txt += "  - Modelo:             $($r.Modelo)`r`n"
                $txt += "  - Numero de Serie:    $($r.Numero_Serie)`r`n"
                $txt += "  - Ano de Fabricacao:  $($r.Ano_Fabricacao)`r`n"
                $txt += "  - Nome do Host:       $($r.Nome_Computador)`r`n"
                $txt += "  - Processador:        $($r.Processador)`r`n"
                $txt += "  - Memoria RAM:        $($r.Memoria_RAM)`r`n"
                $txt += "  - Armazenamento:      $($r.Armazenamento)`r`n"
                $txt += "  - Placa de Video:     $($r.Placa_Video)`r`n"
                $txt += "  - Sistema Operacional:$($r.Sistema_Operacional)`r`n"
                $txt += "  - Endereco IP / MAC:  $($r.IP_Rede) | $($r.MAC_Rede)`r`n"
                $txt += "  - Registrado em:      $($r.Data_Registro)`r`n`r`n"
                $i++
            }
            
            $txt += "$divisoria`r`n INNOVCORE TI - Desenvolvido por Jordan para InnovTech`r`n$divisoria`r`n"
            [System.IO.File]::WriteAllText($caminhoArquivo, $txt, $script:utf8ComBOM)
            
            Write-Host "`n [OK] Relatorio Executivo em Texto gerado com sucesso!" -ForegroundColor Green
            Write-Host "      Arquivo: " -NoNewline -ForegroundColor Gray
            Write-Host "$caminhoArquivo" -ForegroundColor Yellow
            
            if ($abrirAposExportar) {
                try { Invoke-Item $caminhoArquivo } catch {}
            }
            return $caminhoArquivo
        }
    }
}

# ==============================================================================
# MENU DE EXPORTACAO DEDICADO
# ==============================================================================
function Menu-Exportacao {
    param([array]$registros)
    
    if ($registros.Count -eq 0) {
        try { Clear-Host } catch {}
        Exibir-LogoPrincipal
        Write-Host " |                 >>> CENTRAL DE EXPORTACOES <<<                       |" -ForegroundColor Yellow
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  [!] Nao ha registros cadastrados para exportar no momento." -ForegroundColor Yellow
        Write-Host "      Realize o cadastro de equipamentos primeiro na opcao [1] ou [2]." -ForegroundColor Gray
        Pausar-Tela
        return
    }
    
    $emExportacao = $true
    while ($emExportacao) {
        try { Clear-Host } catch {}
        Exibir-LogoPrincipal
        Write-Host " |            >>> CENTRAL DE EXPORTACOES INNOVCORE TI <<<               |" -ForegroundColor Yellow
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host " |                                                                      |" -ForegroundColor Cyan
        Write-Host " |  " -NoNewline -ForegroundColor Cyan
        Write-Host "[1]" -NoNewline -ForegroundColor Green
        Write-Host " Exportar Planilha Excel UTF-8 (.CSV com BOM para Excel PT-BR)  |" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan
        Write-Host "[2]" -NoNewline -ForegroundColor Yellow
        Write-Host " Exportar Planilha Formatada Corporativa (.XLS Estilizado)      |" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan
        Write-Host "[3]" -NoNewline -ForegroundColor Magenta
        Write-Host " Exportar Relatorio Executivo Completo (.TXT Tecnico)           |" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan
        Write-Host "[4]" -NoNewline -ForegroundColor Cyan
        Write-Host " Exportar Apenas um Setor Especifico (Excel UTF-8)              |" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan
        Write-Host "[5]" -NoNewline -ForegroundColor White
        Write-Host " Abrir Planilha Principal no Excel agora                        |" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan
        Write-Host "[6]" -NoNewline -ForegroundColor DarkCyan
        Write-Host " Abrir Pasta de Exportacoes no Explorer                         |" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan
        Write-Host "[0]" -NoNewline -ForegroundColor Red
        Write-Host " Voltar ao Menu Principal                                       |" -ForegroundColor Gray
        Write-Host " |                                                                      |" -ForegroundColor Cyan
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host ""
        
        $subOpcao = Ler-OpcaoOuEsc -prompt " Digite o numero da opcao desejada (ou ESC para voltar)" -valorPadrao "1"
        
        switch ($subOpcao.Trim()) {
            "1" {
                Write-Host "`n Deseja abrir a planilha no Excel automaticamente apos gerar? (S/N) " -NoNewline -ForegroundColor Cyan
                $abrir = (Read-Host "[S]") -notmatch '^(?i)n|nao|no$'
                Exportar-InventarioExcel -registros $registros -tipo "CSV_EXCEL" -abrirAposExportar:$abrir
                Pausar-Tela
            }
            "2" {
                Write-Host "`n Deseja abrir no Excel automaticamente apos gerar? (S/N) " -NoNewline -ForegroundColor Cyan
                $abrir = (Read-Host "[S]") -notmatch '^(?i)n|nao|no$'
                Exportar-InventarioExcel -registros $registros -tipo "HTML_EXCEL" -abrirAposExportar:$abrir
                Pausar-Tela
            }
            "3" {
                Write-Host "`n Deseja abrir o relatorio no Bloco de Notas automaticamente? (S/N) " -NoNewline -ForegroundColor Cyan
                $abrir = (Read-Host "[S]") -notmatch '^(?i)n|nao|no$'
                Exportar-InventarioExcel -registros $registros -tipo "TXT_RELATORIO" -abrirAposExportar:$abrir
                Pausar-Tela
            }
            "4" {
                $setoresDisponiveis = @($registros | ForEach-Object { $_.Setor.Trim() } | Where-Object { $_ -ne "" } | Select-Object -Unique)
                Write-Host "`n Setores cadastrados no sistema:" -ForegroundColor Yellow
                for ($k = 0; $k -lt $setoresDisponiveis.Count; $k++) {
                    $nFormat = ($k + 1).ToString().PadLeft(2, ' ')
                    Write-Host "   [$nFormat] $($setoresDisponiveis[$k])" -ForegroundColor White
                }
                Write-Host ""
                $escolha = Read-Host " Escolha o numero do setor para exportar"
                $numSetor = 0
                if ([int]::TryParse($escolha.Trim(), [ref]$numSetor) -and $numSetor -ge 1 -and $numSetor -le $setoresDisponiveis.Count) {
                    $setorEscolhido = $setoresDisponiveis[$numSetor - 1]
                    Write-Host " Exportando equipamentos do setor: $setorEscolhido..." -ForegroundColor Green
                    Exportar-InventarioExcel -registros $registros -tipo "CSV_EXCEL" -filtroSetor $setorEscolhido -abrirAposExportar:$true
                } else {
                    Write-Host " Setor invalido selecionado." -ForegroundColor Red
                }
                Pausar-Tela
            }
            "5" {
                if (Test-Path $script:caminhoCsv) {
                    Write-Host "`n Abrindo base principal no Excel..." -ForegroundColor Green
                    try { Invoke-Item $script:caminhoCsv } catch { Write-Host " [!] Erro ao abrir arquivo: $($_.Exception.Message)" -ForegroundColor Red }
                } else {
                    Write-Host "`n [!] A planilha principal ainda nao foi criada." -ForegroundColor Yellow
                }
                Pausar-Tela
            }
            "6" {
                if (!(Test-Path $script:pastaExport)) {
                    New-Item -ItemType Directory -Path $script:pastaExport -Force | Out-Null
                }
                try { Invoke-Item $script:pastaExport } catch {}
            }
            "0" {
                $emExportacao = $false
            }
            default {
                Write-Host " Opcao invalida." -ForegroundColor Red
                Start-Sleep -Milliseconds 700
            }
        }
    }
}

# ==============================================================================
# MENU DE BUSCA E CONSULTA AVANCADA
# ==============================================================================
function Menu-ConsultarRegistros {
    param([array]$registros)
    
    if ($registros.Count -eq 0) {
        try { Clear-Host } catch {}
        Exibir-LogoPrincipal
        Write-Host " |                >>> CONSULTA DE CADASTROS <<<                         |" -ForegroundColor Yellow
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  Nenhum equipamento registrado na base InnovCore ainda." -ForegroundColor Yellow
        Write-Host "  Utilize a opcao [1] ou [2] para inventariar o primeiro computador." -ForegroundColor Gray
        Pausar-Tela
        return
    }
    
    $emConsulta = $true
    while ($emConsulta) {
        try { Clear-Host } catch {}
        Exibir-LogoPrincipal
        Write-Host " |                >>> CONSULTA DE CADASTROS <<<                         |" -ForegroundColor Yellow
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host "  Total de maquinas catalogadas no InnovCore: " -NoNewline -ForegroundColor White
        Write-Host "$($registros.Count)" -ForegroundColor Green
        Write-Host ""
        Write-Host "  [1] Listar todos os equipamentos em formato resumido" -ForegroundColor White
        Write-Host "  [2] Pesquisar por Colaborador / Nome" -ForegroundColor White
        Write-Host "  [3] Pesquisar por Setor" -ForegroundColor White
        Write-Host "  [4] Pesquisar por Patrimonio" -ForegroundColor White
        Write-Host "  [5] Pesquisar por Numero de Serie / Hostname" -ForegroundColor White
        Write-Host "  [6] Ver detalhes tecnicos completos de um equipamento" -ForegroundColor Cyan
        Write-Host "  [0] Voltar ao Menu Principal" -ForegroundColor Gray
        Write-Host ""
        
        $opcConsulta = Ler-OpcaoOuEsc -prompt " Digite o numero da opcao desejada (ou ESC para voltar)" -valorPadrao "1"
        
        switch ($opcConsulta.Trim()) {
            "1" {
                try { Clear-Host } catch {}
                Exibir-LogoPrincipal
                Write-Host " |              >>> LISTA GERAL DE EQUIPAMENTOS <<<                     |" -ForegroundColor Yellow
                Write-Host " +======================================================================+" -ForegroundColor Cyan
                Write-Host ""
                $idx = 1
                foreach ($r in $registros) {
                    Write-Host "  [$idx] " -NoNewline -ForegroundColor Cyan
                    Write-Host "PATRIMONIO: " -NoNewline -ForegroundColor Gray
                    Write-Host "$($r.Patrimonio)" -ForegroundColor Green
                    Write-Host "      Colaborador: " -NoNewline -ForegroundColor Gray
                    Write-Host "$($r.Usuario)" -NoNewline -ForegroundColor White
                    Write-Host " | Setor: " -NoNewline -ForegroundColor Gray
                    Write-Host "$($r.Setor)" -NoNewline -ForegroundColor Green
                    Write-Host " | Funcao: " -NoNewline -ForegroundColor Gray
                    Write-Host "$($r.Funcao)" -ForegroundColor White
                    Write-Host "      Equipamento: " -NoNewline -ForegroundColor Gray
                    Write-Host "$($r.Marca) $($r.Modelo)" -NoNewline -ForegroundColor Cyan
                    Write-Host " | Serial: " -NoNewline -ForegroundColor Gray
                    Write-Host "$($r.Numero_Serie)" -NoNewline -ForegroundColor Yellow
                    Write-Host " | Ano: " -NoNewline -ForegroundColor Gray
                    Write-Host "$($r.Ano_Fabricacao)" -ForegroundColor Yellow
                    Write-Host "      Host: " -NoNewline -ForegroundColor Gray
                    Write-Host "$($r.Nome_Computador)" -NoNewline -ForegroundColor DarkGray
                    Write-Host " | Registrado em: " -NoNewline -ForegroundColor Gray
                    Write-Host "$($r.Data_Registro)" -ForegroundColor DarkGray
                    Write-Host "  ----------------------------------------------------------------------" -ForegroundColor DarkCyan
                    $idx++
                }
                Pausar-Tela
            }
            
            { $_ -in @("2", "3", "4", "5") } {
                $tipoBusca = switch ($opcConsulta.Trim()) {
                    "2" { "Colaborador" }
                    "3" { "Setor" }
                    "4" { "Patrimonio" }
                    "5" { "Serial / Hostname" }
                }
                Write-Host "`n Digite o termo para buscar por ${tipoBusca}: " -NoNewline -ForegroundColor Yellow
                $termo = Read-Host
                if (![string]::IsNullOrWhiteSpace($termo)) {
                    $t = $termo.Trim()
                    $encontrados = switch ($opcConsulta.Trim()) {
                        "2" { @($registros | Where-Object { $_.Usuario -match "(?i)$t" }) }
                        "3" { @($registros | Where-Object { $_.Setor -match "(?i)$t" }) }
                        "4" { @($registros | Where-Object { $_.Patrimonio -match "(?i)$t" }) }
                        "5" { @($registros | Where-Object { $_.Numero_Serie -match "(?i)$t" -or $_.Nome_Computador -match "(?i)$t" }) }
                    }
                    
                    Write-Host "`n Resultados encontrados: " -NoNewline -ForegroundColor Gray
                    Write-Host "$($encontrados.Count)" -ForegroundColor Green
                    Write-Host " ----------------------------------------------------------------------" -ForegroundColor DarkCyan
                    
                    $i = 1
                    foreach ($r in $encontrados) {
                        Write-Host "  [$i] " -NoNewline -ForegroundColor Cyan
                        Write-Host "$($r.Patrimonio) " -NoNewline -ForegroundColor Green
                        Write-Host "- $($r.Usuario) " -NoNewline -ForegroundColor White
                        Write-Host "($($r.Setor)) " -NoNewline -ForegroundColor Yellow
                        Write-Host "[$($r.Marca) $($r.Modelo) | Serial: $($r.Numero_Serie)]" -ForegroundColor Gray
                        $i++
                    }
                }
                Pausar-Tela
            }
            
            "6" {
                Write-Host "`n Digite o numero da maquina (1 a $($registros.Count)) ou o codigo de Patrimonio: " -NoNewline -ForegroundColor Yellow
                $escolha = Read-Host
                $selecionado = $null
                $num = 0
                if ([int]::TryParse($escolha.Trim(), [ref]$num) -and $num -ge 1 -and $num -le $registros.Count) {
                    $selecionado = $registros[$num - 1]
                } else {
                    $selecionado = $registros | Where-Object { $_.Patrimonio -match "(?i)$($escolha.Trim())" } | Select-Object -First 1
                }
                
                if ($selecionado) {
                    try { Clear-Host } catch {}
                    Exibir-LogoPrincipal
                    Write-Host " |                >>> DETALHES COMPLETOS DO ATIVO <<<                   |" -ForegroundColor Yellow
                    Write-Host " +======================================================================+" -ForegroundColor Cyan
                    Write-Host "   EMPRESA:              " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Empresa)" -ForegroundColor White
                    Write-Host "   PATRIMONIO:           " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Patrimonio)" -ForegroundColor Green
                    Write-Host "   COLABORADOR:          " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Usuario)" -ForegroundColor White
                    Write-Host "   CARGO / FUNCAO:       " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Funcao)" -ForegroundColor White
                    Write-Host "   SETOR:                " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Setor)" -ForegroundColor Green
                    Write-Host "  ----------------------------------------------------------------------" -ForegroundColor DarkGray
                    Write-Host "   TIPO EQUIPAMENTO:     " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Tipo_Equipamento)" -ForegroundColor Cyan
                    Write-Host "   FABRICANTE / MARCA:   " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Marca)" -ForegroundColor Cyan
                    Write-Host "   MODELO DO PC:         " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Modelo)" -ForegroundColor Cyan
                    Write-Host "   NUMERO DE SERIE:      " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Numero_Serie)" -ForegroundColor Yellow
                    Write-Host "   ANO DE FABRICACAO:    " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Ano_Fabricacao)" -ForegroundColor Yellow
                    Write-Host "   NOME DO HOST (PC):    " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Nome_Computador)" -ForegroundColor White
                    Write-Host "  ----------------------------------------------------------------------" -ForegroundColor DarkGray
                    Write-Host "   PROCESSADOR:          " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Processador)" -ForegroundColor White
                    Write-Host "   MEMORIA RAM:          " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Memoria_RAM)" -ForegroundColor White
                    Write-Host "   ARMAZENAMENTO:        " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Armazenamento)" -ForegroundColor White
                    Write-Host "   PLACA DE VIDEO:       " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Placa_Video)" -ForegroundColor White
                    Write-Host "   SISTEMA OPERACIONAL:  " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Sistema_Operacional)" -ForegroundColor White
                    Write-Host "   ENDERECO IP / MAC:    " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.IP_Rede) | $($selecionado.MAC_Rede)" -ForegroundColor DarkGray
                    Write-Host "   DATA DO CADASTRO:     " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Data_Registro)" -ForegroundColor DarkGray
                    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
                } else {
                    Write-Host " [!] Equipamento nao encontrado." -ForegroundColor Red
                }
                Pausar-Tela
            }
            
            "0" {
                $emConsulta = $false
            }
            default {
                Write-Host " Opcao invalida." -ForegroundColor Red
                Start-Sleep -Milliseconds 700
            }
        }
    }
}

# ==============================================================================
# MENU DE EDICAO DE CADASTRO
# ==============================================================================
function Menu-EditarCadastro {
    param([array]$registros)
    
    if ($registros.Count -eq 0) {
        try { Clear-Host } catch {}
        Exibir-LogoPrincipal
        Write-Host " |                >>> EDICAO DE CADASTRO <<<                           |" -ForegroundColor Yellow
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  Nao ha registros cadastrados para editar." -ForegroundColor Yellow
        Pausar-Tela
        return
    }
    
    try { Clear-Host } catch {}
    Exibir-LogoPrincipal
    Write-Host " |                >>> EDICAO DE CADASTRO <<<                           |" -ForegroundColor Yellow
    Write-Host " +======================================================================+" -ForegroundColor Cyan
    Write-Host ""
    
    $idx = 1
    foreach ($r in $registros) {
        Write-Host "  [$idx] " -NoNewline -ForegroundColor Cyan
        Write-Host "$($r.Patrimonio) " -NoNewline -ForegroundColor Green
        Write-Host "- $($r.Usuario) " -NoNewline -ForegroundColor White
        Write-Host "($($r.Setor) - $($r.Marca) $($r.Modelo))" -ForegroundColor Gray
        $idx++
    }
    Write-Host ""
    Write-Host "  Digite o NUMERO do registro que deseja EDITAR (ou ESC/0 para voltar):" -ForegroundColor Yellow
    $resp = Ler-OpcaoOuEsc -prompt "  Registro a editar (ou ESC para voltar)" -valorPadrao "0"
    
    $numEdicao = 0
    if ([int]::TryParse($resp.Trim(), [ref]$numEdicao) -and $numEdicao -ge 1 -and $numEdicao -le $registros.Count) {
        $item = $registros[$numEdicao - 1]
        
        Write-Host "`n  Editando cadastro de: " -NoNewline -ForegroundColor Yellow
        Write-Host "$($item.Patrimonio) - $($item.Usuario)" -ForegroundColor Green
        Write-Host "  (Pressione ENTER para manter o valor atual)`n" -ForegroundColor Gray
        
        # Colaborador
        Write-Host "  Colaborador atual: " -NoNewline -ForegroundColor Gray
        Write-Host "$($item.Usuario)" -ForegroundColor White
        $novoUsuario = Read-Host "  Novo Colaborador"
        if (![string]::IsNullOrWhiteSpace($novoUsuario)) { $item.Usuario = $novoUsuario.Trim() }
        
        # Cargo / Funcao
        Write-Host "  Funcao atual: " -NoNewline -ForegroundColor Gray
        Write-Host "$($item.Funcao)" -ForegroundColor White
        $novaFuncao = Read-Host "  Nova Funcao"
        if (![string]::IsNullOrWhiteSpace($novoFuncao)) { $item.Funcao = $novaFuncao.Trim() }
        
        # Setor
        Write-Host "  Setor atual: " -NoNewline -ForegroundColor Gray
        Write-Host "$($item.Setor)" -ForegroundColor White
        $novoSetor = Read-Host "  Novo Setor"
        if (![string]::IsNullOrWhiteSpace($novoSetor)) { $item.Setor = $novoSetor.Trim() }
        
        # Patrimonio
        Write-Host "  Patrimonio atual: " -NoNewline -ForegroundColor Gray
        Write-Host "$($item.Patrimonio)" -ForegroundColor White
        $novoPat = Read-Host "  Novo Patrimonio"
        if (![string]::IsNullOrWhiteSpace($novoPat)) { $item.Patrimonio = $novoPat.Trim() }
        
        # Tipo Equipamento
        Write-Host "  Tipo atual: " -NoNewline -ForegroundColor Gray
        Write-Host "$($item.Tipo_Equipamento)" -ForegroundColor White
        $novoTipo = Read-Host "  Novo Tipo (Notebook, Desktop, etc)"
        if (![string]::IsNullOrWhiteSpace($novoTipo)) { $item.Tipo_Equipamento = $novoTipo.Trim() }
        
        # Marca e Modelo
        Write-Host "  Marca atual: " -NoNewline -ForegroundColor Gray
        Write-Host "$($item.Marca)" -ForegroundColor White
        $novaMarca = Read-Host "  Nova Marca"
        if (![string]::IsNullOrWhiteSpace($novaMarca)) { $item.Marca = $novaMarca.Trim() }
        
        Write-Host "  Modelo atual: " -NoNewline -ForegroundColor Gray
        Write-Host "$($item.Modelo)" -ForegroundColor White
        $novoModelo = Read-Host "  Novo Modelo"
        if (![string]::IsNullOrWhiteSpace($novoModelo)) { $item.Modelo = $novoModelo.Trim() }
        
        # Serial
        Write-Host "  Serial atual: " -NoNewline -ForegroundColor Gray
        Write-Host "$($item.Numero_Serie)" -ForegroundColor White
        $novoSerial = Read-Host "  Novo Serial"
        if (![string]::IsNullOrWhiteSpace($novoSerial)) { $item.Numero_Serie = $novoSerial.Trim() }
        
        # Salva alteracoes
        Salvar-RegistrosCsv -caminho $script:caminhoCsv -registros $registros
        Write-Host "`n  [OK] Cadastro atualizado com sucesso na base de dados!" -ForegroundColor Green
    } else {
        Write-Host "  Operacao cancelada." -ForegroundColor Gray
    }
    Pausar-Tela
}

# ==============================================================================
# DASHBOARD EXECUTIVO INTERATIVO NO TERMINAL (ASCII / ANSI)
# ==============================================================================
function Menu-DashboardExecutivo {
    param([array]$registros)
    
    $registrosValidos = @($registros | Where-Object { 
        $_ -and (![string]::IsNullOrWhiteSpace($_.Patrimonio) -or ![string]::IsNullOrWhiteSpace($_.Usuario))
    })
    
    $emDashboard = $true
    $abaAtual = "1"
    $filtroSetorAtivo = $null
    
    # Sub-funcao para ordenacao numerica estrita de patrimonio (PAT-0001, PAT-0002, etc.)
    $ordenarPorPatrimonio = {
        param([array]$lista)
        return @($lista | Sort-Object {
            if ($_.Patrimonio -match '(\d+)') {
                [int64]$matches[1]
            } else {
                [int64]::MaxValue
            }
        }, { if ($_.Patrimonio) { $_.Patrimonio } else { "" } })
    }
    
    while ($emDashboard) {
        try { Clear-Host } catch {}
        Exibir-LogoPrincipal
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host " |            >>> DASHBOARD EXECUTIVO DO PARQUE DE ATIVOS <<<           |" -ForegroundColor Yellow
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        
        $totalAtivos = $registrosValidos.Count
        if ($totalAtivos -eq 0) {
            Write-Host "`n  [!] A base de dados ainda nao possui equipamentos cadastrados." -ForegroundColor Yellow
            Write-Host "  Cadastre seu primeiro computador usando a opcao [1] do menu inicial.`n" -ForegroundColor Gray
            Pausar-Tela
            return
        }
        
        # 1. Calculo de Indicadores Chave (KPIs)
        $ramTotalGb = 0
        foreach ($r in $registrosValidos) {
            if ($r.Memoria_RAM -match '(\d+)') {
                $ramTotalGb += [int]$matches[1]
            }
        }
        
        $qtdNotebooks = @($registrosValidos | Where-Object { $_.Tipo_Equipamento -match '(?i)Notebook|Laptop' }).Count
        $qtdDesktops  = @($registrosValidos | Where-Object { $_.Tipo_Equipamento -match '(?i)Desktop|Computador|Gabinete' }).Count
        $qtdServidores= @($registrosValidos | Where-Object { $_.Tipo_Equipamento -match '(?i)Servidor|Server|All-in-One|Outro' }).Count
        if (($qtdNotebooks + $qtdDesktops + $qtdServidores) -lt $totalAtivos) {
            $qtdServidores = $totalAtivos - ($qtdNotebooks + $qtdDesktops)
        }
        
        $grpSetores = $registrosValidos | Group-Object { 
            if ($_.Setor -and $_.Setor.Trim() -ne "") { $_.Setor.Trim() } else { "Nao Informado" } 
        } | Sort-Object Count -Descending
        $topSetorNome = if ($grpSetores.Count -gt 0 -and $grpSetores[0].Name) { $grpSetores[0].Name } else { "N/A" }
        $topSetorQtd  = if ($grpSetores.Count -gt 0) { $grpSetores[0].Count } else { 0 }
        
        $grpMarcas = $registrosValidos | Group-Object Marca | Sort-Object Count -Descending
        $topMarcaNome = if ($grpMarcas.Count -gt 0 -and $grpMarcas[0].Name) { $grpMarcas[0].Name } else { "N/A" }
        $topMarcaQtd  = if ($grpMarcas.Count -gt 0) { $grpMarcas[0].Count } else { 0 }
        
        # 2. Renderizar Cartoes Executivos (KPI Cards)
        Write-Host ""
        Write-Host "  +----------------------+ +----------------------+ +----------------------+" -ForegroundColor Cyan
        Write-Host "  | TOTAL DE ATIVOS      | | MEMORIA RAM TOTAL    | | MAIOR DEPARTAMENTO   |" -ForegroundColor White
        
        $kpi1 = ("$totalAtivos ATIVOS").PadRight(18).Substring(0, 18)
        $kpi2 = ("$ramTotalGb GB TOTAL").PadRight(18).Substring(0, 18)
        $kpi3 = ("$topSetorNome ($topSetorQtd)").PadRight(18).Substring(0, 18)
        
        Write-Host "  | " -NoNewline -ForegroundColor Cyan; Write-Host ">> $kpi1" -NoNewline -ForegroundColor Green; Write-Host " | " -NoNewline -ForegroundColor Cyan
        Write-Host "| " -NoNewline -ForegroundColor Cyan; Write-Host ">> $kpi2" -NoNewline -ForegroundColor Yellow; Write-Host " | " -NoNewline -ForegroundColor Cyan
        Write-Host "| " -NoNewline -ForegroundColor Cyan; Write-Host ">> $kpi3" -NoNewline -ForegroundColor Magenta; Write-Host " |" -ForegroundColor Cyan
        Write-Host "  +----------------------+ +----------------------+ +----------------------+" -ForegroundColor Cyan
        
        Write-Host "  +----------------------+ +----------------------+ +----------------------+" -ForegroundColor Cyan
        Write-Host "  | NOTEBOOKS / LAPTOPS  | | DESKTOPS / PCS       | | PRINCIPAL FABRICANTE |" -ForegroundColor White
        
        $kpi4 = ("$qtdNotebooks UNIDADES").PadRight(18).Substring(0, 18)
        $kpi5 = ("$qtdDesktops UNIDADES").PadRight(18).Substring(0, 18)
        $kpi6 = ("$topMarcaNome ($topMarcaQtd)").PadRight(18).Substring(0, 18)
        
        Write-Host "  | " -NoNewline -ForegroundColor Cyan; Write-Host ">> $kpi4" -NoNewline -ForegroundColor Cyan; Write-Host " | " -NoNewline -ForegroundColor Cyan
        Write-Host "| " -NoNewline -ForegroundColor Cyan; Write-Host ">> $kpi5" -NoNewline -ForegroundColor White; Write-Host " | " -NoNewline -ForegroundColor Cyan
        Write-Host "| " -NoNewline -ForegroundColor Cyan; Write-Host ">> $kpi6" -NoNewline -ForegroundColor Yellow; Write-Host " |" -ForegroundColor Cyan
        Write-Host "  +----------------------+ +----------------------+ +----------------------+" -ForegroundColor Cyan
        
        # 3. Navegacao por Visoes do Dashboard
        $corAba1 = if ($abaAtual -eq "1") { "Green" } else { "DarkGray" }
        $corAba2 = if ($abaAtual -eq "2") { "Yellow" } else { "DarkGray" }
        $corAba3 = if ($abaAtual -eq "3") { "Cyan" } else { "DarkGray" }
        $corAba4 = if ($abaAtual -eq "4") { "Magenta" } else { "DarkGray" }
        
        Write-Host ""
        Write-Host "  VISOES DO DASHBOARD:  " -NoNewline -ForegroundColor Gray
        Write-Host "[1] Por Setor & Integrantes  " -NoNewline -ForegroundColor $corAba1
        Write-Host "[2] Graficos & Marcas  " -NoNewline -ForegroundColor $corAba2
        Write-Host "[3] Hardware & SO  " -NoNewline -ForegroundColor $corAba3
        Write-Host "[4] Tabela de Ativos  " -NoNewline -ForegroundColor $corAba4
        Write-Host "[0] Voltar" -ForegroundColor Gray
        Write-Host "  ----------------------------------------------------------------------" -ForegroundColor DarkGray
        
        # Sub-funcao para desenhar barras ASCII
        $desenharBarra = {
            param([string]$label, [int]$valor, [int]$total, [int]$tamMax = 20, [string]$cor = "Cyan")
            $pct = if ($total -gt 0) { [math]::Round(($valor / $total) * 100, 1) } else { 0 }
            $qtdBlocos = if ($total -gt 0) { [math]::Min($tamMax, [math]::Max(0, [int][math]::Round(($valor / $total) * $tamMax))) } else { 0 }
            $qtdEspacos = [math]::Max(0, $tamMax - $qtdBlocos)
            $barra = ("#" * $qtdBlocos) + ("-" * $qtdEspacos)
            
            $lblPad = $label.PadRight(18).Substring(0, [math]::Min(18, $label.Length)).PadRight(18)
            Write-Host "   $lblPad " -NoNewline -ForegroundColor White
            Write-Host "[$barra] " -NoNewline -ForegroundColor $cor
            Write-Host ("{0,5:N1}%" -f $pct) -NoNewline -ForegroundColor Green
            Write-Host " ($valor/$total)" -ForegroundColor Gray
        }
        
        switch ($abaAtual) {
            "1" {
                # ABA 1: DEMONSTRATIVO DE EQUIPAMENTOS POR SETOR & INTEGRANTES (ORDEM NUMERICA)
                Write-Host "`n  [+] DEMONSTRATIVO DE EQUIPAMENTOS POR SETOR & INTEGRANTES:" -ForegroundColor Yellow
                Write-Host "      (Equipamentos agrupados por setor e ordenados numericamente por patrimonio)`n" -ForegroundColor DarkGray
                
                $listaSetores = $grpSetores | Sort-Object Name
                if ($filtroSetorAtivo) {
                    $listaSetores = @($listaSetores | Where-Object { $_.Name -eq $filtroSetorAtivo })
                    Write-Host "  >> Filtrando apenas setor: [$filtroSetorAtivo]  (Digite [T] para ver todos)`n" -ForegroundColor Magenta
                }
                
                foreach ($grp in $listaSetores) {
                    $nomeSetor = if ($grp.Name) { $grp.Name.ToUpper() } else { "NAO INFORMADO" }
                    $qtd = $grp.Count
                    $txtQtd = if ($qtd -eq 1) { "1 Ativo" } else { "$qtd Ativos" }
                    
                    $tagSetor = " [ SETOR: $nomeSetor | $txtQtd ] "
                    $larguraTotal = 74
                    $tamHifens = [math]::Max(2, $larguraTotal - $tagSetor.Length - 5)
                    $linhaDecorada = " +---$tagSetor" + ("-" * $tamHifens) + "+"
                    if ($linhaDecorada.Length -gt $larguraTotal) { $linhaDecorada = $linhaDecorada.Substring(0, $larguraTotal) }
                    
                    Write-Host ""
                    Write-Host $linhaDecorada -ForegroundColor Cyan
                    Write-Host "  PATRIM.   | INTEGRANTE / USUARIO     | CARGO / FUNCAO       | MODELO / HARDWARE" -ForegroundColor DarkYellow
                    Write-Host "  ----------+--------------------------+----------------------+---------------------" -ForegroundColor DarkGray
                    
                    $itensDoSetor = & $ordenarPorPatrimonio $grp.Group
                    
                    foreach ($item in $itensDoSetor) {
                        $pat = if ($item.Patrimonio) { $item.Patrimonio.Trim() } else { "PAT-????" }
                        $usr = if ($item.Usuario) { $item.Usuario.Trim() } else { "Nao informado" }
                        $usrCortado = $usr.PadRight(24).Substring(0, [math]::Min(24, $usr.Length))
                        $cargo = if ($item.Funcao) { $item.Funcao.Trim() } else { "N/A" }
                        $cargoCortado = $cargo.PadRight(20).Substring(0, [math]::Min(20, $cargo.Length))
                        
                        $marcaModelo = ("$($item.Marca) $($item.Modelo)").Trim()
                        if ([string]::IsNullOrWhiteSpace($marcaModelo)) { $marcaModelo = $item.Tipo_Equipamento }
                        $equipCortado = $marcaModelo.PadRight(21).Substring(0, [math]::Min(21, $marcaModelo.Length))
                        
                        Write-Host "  " -NoNewline
                        Write-Host ("{0,-9} " -f $pat) -NoNewline -ForegroundColor Green
                        Write-Host "| " -NoNewline -ForegroundColor DarkGray
                        Write-Host ("{0,-24} " -f $usrCortado) -NoNewline -ForegroundColor White
                        Write-Host "| " -NoNewline -ForegroundColor DarkGray
                        Write-Host ("{0,-20} " -f $cargoCortado) -NoNewline -ForegroundColor Cyan
                        Write-Host "| " -NoNewline -ForegroundColor DarkGray
                        Write-Host ("{0,-21}" -f $equipCortado) -ForegroundColor Yellow
                    }
                }
            }
            
            "2" {
                # ABA 2: DISTRIBUICAO POR SETOR E FABRICANTE (GRAFICOS)
                Write-Host "`n  [+] DISTRIBUICAO POR SETOR / DEPARTAMENTO:" -ForegroundColor Yellow
                foreach ($grp in $grpSetores) {
                    $nome = if ($grp.Name) { $grp.Name } else { "Nao Informado" }
                    & $desenharBarra -label $nome -valor $grp.Count -total $totalAtivos -cor "Green"
                }
                
                Write-Host "`n  [+] DISTRIBUICAO POR FABRICANTE / MARCA:" -ForegroundColor Yellow
                foreach ($grp in $grpMarcas) {
                    $nome = if ($grp.Name) { $grp.Name } else { "N/A" }
                    & $desenharBarra -label $nome -valor $grp.Count -total $totalAtivos -cor "Cyan"
                }
            }
            
            "3" {
                # ABA 3: HARDWARE, PROCESSADORES E SISTEMAS OPERACIONAIS
                Write-Host "`n  [+] SISTEMAS OPERACIONAIS (SO):" -ForegroundColor Yellow
                $grpSo = $registrosValidos | Group-Object Sistema_Operacional | Sort-Object Count -Descending
                foreach ($grp in $grpSo) {
                    $nome = if ($grp.Name) { $grp.Name } else { "Nao Identificado" }
                    & $desenharBarra -label $nome -valor $grp.Count -total $totalAtivos -cor "Yellow"
                }
                
                Write-Host "`n  [+] TIPOS DE DISPOSITIVOS:" -ForegroundColor Yellow
                $grpTipo = $registrosValidos | Group-Object Tipo_Equipamento | Sort-Object Count -Descending
                foreach ($grp in $grpTipo) {
                    $nome = if ($grp.Name) { $grp.Name } else { "Computador" }
                    & $desenharBarra -label $nome -valor $grp.Count -total $totalAtivos -cor "Magenta"
                }
                
                Write-Host "`n  [+] FAMILIAS DE PROCESSADORES (CPU):" -ForegroundColor Yellow
                $grpCpu = $registrosValidos | ForEach-Object {
                    $cpu = $_.Processador
                    if ($cpu -match '(?i)i[3579]') { "Intel $($matches[0].ToUpper())" }
                    elseif ($cpu -match '(?i)Ryzen\s*[3579]') { "AMD $($matches[0])" }
                    elseif ($cpu -match '(?i)Apple\s*M\d') { "$($matches[0])" }
                    elseif ($cpu -match '(?i)Intel') { "Intel Geral" }
                    elseif ($cpu -match '(?i)AMD') { "AMD Geral" }
                    else { "Outro / N/A" }
                } | Group-Object | Sort-Object Count -Descending
                
                foreach ($grp in $grpCpu) {
                    & $desenharBarra -label $grp.Name -valor $grp.Count -total $totalAtivos -cor "White"
                }
            }
            
            "4" {
                # ABA 4: TABELA GERAL DE ATIVOS (ORDENADA POR PATRIMONIO)
                Write-Host "`n  [+] TABELA GERAL DE ATIVOS CADASTRADOS (ORDEM NUMERICA):" -ForegroundColor Yellow
                Write-Host "  " -NoNewline
                Write-Host ("{0,-9} | {1,-18} | {2,-14} | {3,-15} | {4,-6}" -f "PATRIM.", "COLABORADOR", "SETOR", "MARCA/MODELO", "RAM") -ForegroundColor Gray
                Write-Host "  ----------------------------------------------------------------------" -ForegroundColor DarkGray
                
                $todosOrdenados = & $ordenarPorPatrimonio $registrosValidos
                foreach ($item in $todosOrdenados) {
                    $pat = if ($item.Patrimonio) { $item.Patrimonio } else { "PAT-????" }
                    $usr = if ($item.Usuario) { $item.Usuario } else { "N/A" }
                    $usrCortado = $usr.PadRight(18).Substring(0, [math]::Min(18, $usr.Length))
                    $set = if ($item.Setor) { $item.Setor } else { "N/A" }
                    $setCortado = $set.PadRight(14).Substring(0, [math]::Min(14, $set.Length))
                    $mm = "$($item.Marca) $($item.Modelo)".Trim()
                    $mmCortado = $mm.PadRight(15).Substring(0, [math]::Min(15, $mm.Length))
                    $ram = if ($item.Memoria_RAM) { $item.Memoria_RAM } else { "N/A" }
                    
                    Write-Host "  " -NoNewline
                    Write-Host ("{0,-9} " -f $pat) -NoNewline -ForegroundColor Green
                    Write-Host "| " -NoNewline -ForegroundColor DarkGray
                    Write-Host ("{0,-18} " -f $usrCortado) -NoNewline -ForegroundColor White
                    Write-Host "| " -NoNewline -ForegroundColor DarkGray
                    Write-Host ("{0,-14} " -f $setCortado) -NoNewline -ForegroundColor Cyan
                    Write-Host "| " -NoNewline -ForegroundColor DarkGray
                    Write-Host ("{0,-15} " -f $mmCortado) -NoNewline -ForegroundColor Yellow
                    Write-Host "| " -NoNewline -ForegroundColor DarkGray
                    Write-Host "$ram" -ForegroundColor White
                }
            }
        }
        
        Write-Host ""
        $nav = Ler-OpcaoOuEsc -prompt " Digite o numero da visao [1, 2, 3, 4], [S] Filtrar Setor ou ESC/0 para voltar" -valorPadrao $abaAtual
        
        if ($nav -in @("1", "2", "3", "4")) {
            $abaAtual = $nav
        } elseif ($nav -eq "0") {
            $emDashboard = $false
        } elseif ($nav -match '^(?i)s|f$') {
            $setoresDisp = @($grpSetores | ForEach-Object { $_.Name } | Where-Object { $_ } | Sort-Object)
            Write-Host "`n  Selecione o setor para filtrar no Demonstrativo:" -ForegroundColor Yellow
            Write-Host "   [0] Ver Todos os Setores" -ForegroundColor White
            for ($i = 0; $i -lt $setoresDisp.Count; $i++) {
                Write-Host "   [$($i+1)] $($setoresDisp[$i])" -ForegroundColor Cyan
            }
            $opcSetor = Read-Host "  Digite o numero do setor desejado"
            $numS = 0
            if ([int]::TryParse($opcSetor, [ref]$numS) -and $numS -ge 1 -and $numS -le $setoresDisp.Count) {
                $filtroSetorAtivo = $setoresDisp[$numS - 1]
                $abaAtual = "1"
            } else {
                $filtroSetorAtivo = $null
                $abaAtual = "1"
            }
        } elseif ($nav -match '^(?i)t$') {
            $filtroSetorAtivo = $null
            $abaAtual = "1"
        } else {
            Write-Host " Opcao invalida." -ForegroundColor Red
            Start-Sleep -Milliseconds 600
        }
    }
}

function Menu-Estatisticas {
    param([array]$registros)
    Menu-DashboardExecutivo -registros $registros
}

# ==============================================================================
# MODULO MOBILE & TABLETS - CENTRAL DE GESTAO DE TABLETS ANDROID (ADB)
# ==============================================================================
function Menu-GestaoMobileTablets {
    param(
        [string]$caminhoCsvPrincipal = $script:caminhoCsv,
        [string]$scriptRootDir = $script:scriptRootDir
    )

    $caminhoCsvTablets = Join-Path $scriptRootDir "dados\mobile\relatorio_macs_tablets.csv"
    if (-not (Test-Path $caminhoCsvTablets)) {
        $caminhoCsvTabletsAlt = Join-Path $scriptRootDir "mobile\relatorio_macs_tablets.csv"
        if (Test-Path $caminhoCsvTabletsAlt) {
            $caminhoCsvTablets = $caminhoCsvTabletsAlt
        } else {
            $caminhoCsvTabletsLegacy = Join-Path $scriptRootDir "adb\relatorio_macs_tablets.csv"
            if (Test-Path $caminhoCsvTabletsLegacy) { $caminhoCsvTablets = $caminhoCsvTabletsLegacy }
        }
    }

    $obterTablets = {
        param([string]$caminho)
        $lista = @()
        if (Test-Path $caminho) {
            try {
                $raw = [System.IO.File]::ReadAllText($caminho, [System.Text.Encoding]::UTF8)
                $linhas = $raw -split "\r?\n" | Where-Object { ![string]::IsNullOrWhiteSpace($_) }
                if ($linhas.Count -gt 1) {
                    $cabecalhoLinha = ($linhas[0] -replace "^\uFEFF", "")
                    $cabecalhos = $cabecalhoLinha -split ';'
                    for ($i = 1; $i -lt $linhas.Count; $i++) {
                        $valores = $linhas[$i] -split ';'
                        $obj = [PSCustomObject]@{}
                        for ($k = 0; $k -lt $cabecalhos.Count; $k++) {
                            $prop = $cabecalhos[$k].Trim()
                            $val = if ($k -lt $valores.Count) { $valores[$k].Trim() } else { "" }
                            $obj | Add-Member -MemberType NoteProperty -Name $prop -Value $val -Force
                        }
                        if (-not $obj.PSObject.Properties["Data_Registro"]) { $obj | Add-Member NoteProperty "Data_Registro" "N/A" -Force }
                        if (-not $obj.PSObject.Properties["Serial"])        { $obj | Add-Member NoteProperty "Serial" "N/A" -Force }
                        if (-not $obj.PSObject.Properties["MAC_WiFi"])      { $obj | Add-Member NoteProperty "MAC_WiFi" "N/A" -Force }
                        if (-not $obj.PSObject.Properties["Fabricante"])    { $obj | Add-Member NoteProperty "Fabricante" "N/A" -Force }
                        if (-not $obj.PSObject.Properties["Modelo"])        { $obj | Add-Member NoteProperty "Modelo" "N/A" -Force }
                        if (-not $obj.PSObject.Properties["Android"])       { $obj | Add-Member NoteProperty "Android" "N/A" -Force }
                        if (-not $obj.PSObject.Properties["Bateria"])       { $obj | Add-Member NoteProperty "Bateria" "N/A" -Force }
                        $lista += $obj
                    }
                }
            } catch {
                Write-Host " [!] Erro ao carregar base de tablets: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        return $lista
    }

    $salvarTablets = {
        param([string]$caminho, [array]$registros)
        Fazer-BackupCsv -caminho $caminho
        $cabecalho = "Data_Registro;Serial;MAC_WiFi;Fabricante;Modelo;Android;Bateria"
        $linhas = @($cabecalho)
        foreach ($r in $registros) {
            $d = if ($r.Data_Registro) { ($r.Data_Registro.ToString() -replace ';', ',') } else { "N/A" }
            $s = if ($r.Serial)        { ($r.Serial.ToString() -replace ';', ',') } else { "N/A" }
            $m = if ($r.MAC_WiFi)      { ($r.MAC_WiFi.ToString() -replace ';', ',') } else { "N/A" }
            $f = if ($r.Fabricante)    { ($r.Fabricante.ToString() -replace ';', ',') } else { "N/A" }
            $mo = if ($r.Modelo)       { ($r.Modelo.ToString() -replace ';', ',') } else { "N/A" }
            $a = if ($r.Android)       { ($r.Android.ToString() -replace ';', ',') } else { "N/A" }
            $b = if ($r.Bateria)       { ($r.Bateria.ToString() -replace ';', ',') } else { "N/A" }
            $linhas += "$d;$s;$m;$f;$mo;$a;$b"
        }
        $conteudo = ($linhas -join "`r`n") + "`r`n"
        [System.IO.File]::WriteAllText($caminho, $conteudo, $script:utf8ComBOM)
    }

    $parseIndices = {
        param([string]$inputStr, [int]$maxIndex)
        $indices = [System.Collections.Generic.HashSet[int]]::new()
        if ([string]::IsNullOrWhiteSpace($inputStr)) { return @() }
        $inputClean = $inputStr.Trim()
        if ($inputClean -match '^(?i)todos|all$') {
            for ($i = 1; $i -le $maxIndex; $i++) { [void]$indices.Add($i) }
            return [array]($indices | Sort-Object)
        }
        $tokens = $inputClean -replace ';', ',' -replace '\s*,\s*', ',' -split ','
        foreach ($token in $tokens) {
            $t = $token.Trim()
            if ($t -match '^(\d+)\s*-\s*(\d+)$') {
                $start = [int]$Matches[1]
                $end   = [int]$Matches[2]
                if ($start -gt $end) { $tmp = $start; $start = $end; $end = $tmp }
                for ($k = $start; $k -le $end; $k++) {
                    if ($k -ge 1 -and $k -le $maxIndex) { [void]$indices.Add($k) }
                }
            } elseif ($t -match '^\d+$') {
                $val = [int]$t
                if ($val -ge 1 -and $val -le $maxIndex) { [void]$indices.Add($val) }
            }
        }
        return [array]($indices | Sort-Object)
    }

    $emMenuMobile = $true
    while ($emMenuMobile) {
        $tablets = & $obterTablets $caminhoCsvTablets
        $regsGerais = Obter-RegistrosCsv -caminho $caminhoCsvPrincipal

        try { Clear-Host } catch {}
        Exibir-LogoPrincipal
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host " |        >>> CENTRAL DE GESTAO MOBILE & TABLETS (ANDROID / ADB) <<<    |" -ForegroundColor Magenta
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host "  Base de Dados:    " -NoNewline -ForegroundColor Gray
        Write-Host "$caminhoCsvTablets" -ForegroundColor Yellow
        Write-Host "  Total Registrado: " -NoNewline -ForegroundColor Gray
        Write-Host "$($tablets.Count) tablet(s) cadastrado(s)" -ForegroundColor $(if ($tablets.Count -gt 0) { "Green" } else { "Yellow" })
        Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
        Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[1]" -NoNewline -ForegroundColor Green;  Write-Host " Auto-Scanner Bancada  (Plug & Play USB - Captura MAC wlan0)     |" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[2]" -NoNewline -ForegroundColor Cyan;   Write-Host " Cadastrar Tablet TI   (Cadastro formal com Patrimonio e Setor)  |" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[3]" -NoNewline -ForegroundColor Yellow; Write-Host " Listar Tablets        (Ver Seriais, MACs Reais e Baterias)      |" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[4]" -NoNewline -ForegroundColor Red;    Write-Host " Excluir Tablet(s)     (Excluir 1 ou mais registros com backup)  |" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[5]" -NoNewline -ForegroundColor Green;  Write-Host " Exportar Relatorio    (Gerar planilha Excel UTF-8 de Tablets)   |" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[0]" -NoNewline -ForegroundColor DarkGray;Write-Host " Voltar ao Menu Principal                                       |" -ForegroundColor Gray
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host ""
        $opcM = Ler-OpcaoOuEsc -prompt " Digite o numero da opcao desejada (ou ESC para voltar)" -valorPadrao "3"

        switch ($opcM.Trim().ToLower()) {
            "1" {
                # AUTO-SCANNER BANCADA (PLUG & PLAY)
                $scriptAutoScan = Join-Path $scriptRootDir "mobile\capturar_mac_tablet.ps1"
                if (-not (Test-Path $scriptAutoScan)) {
                    $scriptAutoScan = Join-Path $scriptRootDir "adb\capturar_mac_tablet.ps1"
                }
                if (Test-Path $scriptAutoScan) {
                    & $scriptAutoScan
                } else {
                    Write-Host " [!] Script de auto-scanner nao localizado em: $scriptAutoScan" -ForegroundColor Red
                    Pausar-Tela
                }
            }

            "2" {
                # CADASTRAR TABLET NO INVENTARIO GERAL TI
                Iniciar-CadastroAndroidTablet -registrosAtuais $regsGerais -caminhoCsv $caminhoCsvPrincipal
            }

            "3" {
                # LISTAR DISPOSITIVOS TABLETS REGISTRADOS
                try { Clear-Host } catch {}
                Exibir-LogoPrincipal
                Write-Host " |            >>> LISTA DE TABLETS ANDROID REGISTRADOS <<<              |" -ForegroundColor Yellow
                Write-Host " +======================================================================+" -ForegroundColor Cyan
                Write-Host ""

                if ($tablets.Count -eq 0) {
                    Write-Host "  Nenhum tablet registrado na base de dados no momento." -ForegroundColor Yellow
                    Write-Host "  Use a opcao [1] (Auto-Scanner) ou [2] (Cadastrar Tablet) para registrar." -ForegroundColor Gray
                    Pausar-Tela
                } else {
                    Write-Host "  " -NoNewline
                    Write-Host ("{0,-5} | {1,-19} | {2,-19} | {3,-17} | {4,-15} | {5,-5} | {6,-9}" -f "ID", "DATA REGISTRO", "NUMERO DE SERIE", "MAC WI-FI (REAL)", "MARCA/MODELO", "BAT.", "PATRIM.") -ForegroundColor Cyan
                    Write-Host "  ------+---------------------+---------------------+-------------------+-----------------+-------+----------" -ForegroundColor DarkGray

                    for ($i = 0; $i -lt $tablets.Count; $i++) {
                        $t = $tablets[$i]
                        $id = "[$($i + 1)]".PadRight(5)
                        $dt = if ($t.Data_Registro) { $t.Data_Registro } else { "N/A" }
                        $dt = $dt.PadRight(19).Substring(0, [math]::Min(19, $dt.Length))
                        
                        $sn = if ($t.Serial) { $t.Serial } else { "N/A" }
                        $sn = $sn.PadRight(19).Substring(0, [math]::Min(19, $sn.Length))

                        $mac = if ($t.MAC_WiFi) { $t.MAC_WiFi } else { "N/A" }
                        $mac = $mac.PadRight(17).Substring(0, [math]::Min(17, $mac.Length))

                        $mm = "$($t.Fabricante) $($t.Modelo)".Trim()
                        if ([string]::IsNullOrWhiteSpace($mm)) { $mm = "Android Tablet" }
                        $mm = $mm.PadRight(15).Substring(0, [math]::Min(15, $mm.Length))

                        $bat = if ($t.Bateria) { $t.Bateria } else { "N/A" }
                        $bat = $bat.PadRight(5).Substring(0, [math]::Min(5, $bat.Length))

                        # Procura se ja tem patrimonio no inventario geral
                        $patrimonio = "N/A"
                        $matchGeral = $regsGerais | Where-Object { 
                            ($_.Numero_Serie -and $_.Numero_Serie.Trim().ToUpper() -eq $t.Serial.Trim().ToUpper()) -or
                            ($_.MAC_Rede -and $_.MAC_Rede.Trim().ToUpper() -eq $t.MAC_WiFi.Trim().ToUpper())
                        } | Select-Object -First 1

                        if ($matchGeral -and $matchGeral.Patrimonio) {
                            $patrimonio = $matchGeral.Patrimonio
                        }

                        Write-Host "  " -NoNewline
                        Write-Host "$id " -NoNewline -ForegroundColor Cyan
                        Write-Host "| " -NoNewline -ForegroundColor DarkGray
                        Write-Host "$dt " -NoNewline -ForegroundColor Gray
                        Write-Host "| " -NoNewline -ForegroundColor DarkGray
                        Write-Host "$sn " -NoNewline -ForegroundColor Yellow
                        Write-Host "| " -NoNewline -ForegroundColor DarkGray
                        Write-Host "$mac " -NoNewline -ForegroundColor Green
                        Write-Host "| " -NoNewline -ForegroundColor DarkGray
                        Write-Host "$mm " -NoNewline -ForegroundColor White
                        Write-Host "| " -NoNewline -ForegroundColor DarkGray
                        Write-Host "$bat " -NoNewline -ForegroundColor Cyan
                        Write-Host "| " -NoNewline -ForegroundColor DarkGray
                        Write-Host "$patrimonio" -ForegroundColor $(if ($patrimonio -ne "N/A") { "Green" } else { "DarkGray" })
                    }

                    Write-Host "  ------+---------------------+---------------------+-------------------+-----------------+-------+----------" -ForegroundColor DarkGray
                    Write-Host "   Total de Tablets Registrados: " -NoNewline -ForegroundColor Gray
                    Write-Host "$($tablets.Count)" -ForegroundColor Green

                    Write-Host ""
                    $navRapida = Ler-OpcaoOuEsc -prompt "  Pressione ENTER ou ESC para voltar ou escolha: [E] Excluir, [C] Cadastrar TI, [X] Exportar" -valorPadrao "0"
                    if ($navRapida.Trim() -match '^(?i)e|excluir$') {
                        # Redireciona diretamente para exclusao
                        $opcM = "4"
                    } elseif ($navRapida.Trim() -match '^(?i)c|cadastrar$') {
                        Iniciar-CadastroAndroidTablet -registrosAtuais $regsGerais -caminhoCsv $caminhoCsvPrincipal
                    } elseif ($navRapida.Trim() -match '^(?i)x|exportar$') {
                        $opcM = "5"
                    }
                }
            }

            "4" {
                # EXCLUIR 1 OU MAIS REGISTROS DE TABLETS COM BACKUP
                try { Clear-Host } catch {}
                Exibir-LogoPrincipal
                Write-Host " |            >>> EXCLUSAO DE REGISTROS DE TABLETS <<<                  |" -ForegroundColor Red
                Write-Host " +======================================================================+" -ForegroundColor Cyan
                Write-Host ""

                if ($tablets.Count -eq 0) {
                    Write-Host "  Nao ha registros de tablets para excluir." -ForegroundColor Yellow
                    Pausar-Tela
                    continue
                }

                Write-Host "  Selecione o(s) registro(s) que deseja remover da base:`n" -ForegroundColor Yellow
                for ($i = 0; $i -lt $tablets.Count; $i++) {
                    $t = $tablets[$i]
                    $mm = "$($t.Fabricante) $($t.Modelo)".Trim()
                    Write-Host "  [$($i + 1)] " -NoNewline -ForegroundColor Cyan
                    Write-Host "Serial: " -NoNewline -ForegroundColor Gray
                    Write-Host ("{0,-18} " -f $t.Serial) -NoNewline -ForegroundColor Yellow
                    Write-Host "| MAC: " -NoNewline -ForegroundColor Gray
                    Write-Host ("{0,-17} " -f $t.MAC_WiFi) -NoNewline -ForegroundColor Green
                    Write-Host "| $mm ($($t.Data_Registro))" -ForegroundColor White
                }

                Write-Host ""
                Write-Host "  Como informar os registros para excluir:" -ForegroundColor Gray
                Write-Host "    - Um unico registro:         Ex: 2" -ForegroundColor White
                Write-Host "    - Multiplos registros:       Ex: 1, 3, 4" -ForegroundColor White
                Write-Host "    - Intervalo continuo:        Ex: 1-3" -ForegroundColor White
                Write-Host "    - Excluir todos:             Ex: TODOS" -ForegroundColor White
                Write-Host "    - Cancelar e voltar:         Ex: 0" -ForegroundColor Gray
                Write-Host ""
                $respExcluir = Ler-OpcaoOuEsc -prompt "  Digite o(s) numero(s) para excluir (ou ESC para cancelar)" -valorPadrao "0"

                if ([string]::IsNullOrWhiteSpace($respExcluir) -or $respExcluir.Trim() -eq "0") {
                    Write-Host "`n  Operacao cancelada. Nenhum registro foi alterado." -ForegroundColor Yellow
                    Pausar-Tela
                    continue
                }

                $indicesParaRemover = & $parseIndices $respExcluir $tablets.Count

                if ($indicesParaRemover.Count -eq 0) {
                    Write-Host "`n  [!] Nenhum registro valido foi identificado na selecao." -ForegroundColor Red
                    Pausar-Tela
                    continue
                }

                Write-Host ""
                Write-Host "  +----------------------------------------------------------------------+" -ForegroundColor Yellow
                Write-Host "  | [ATENCAO] Registros selecionados para exclusao:                      |" -ForegroundColor Yellow
                Write-Host "  +----------------------------------------------------------------------+" -ForegroundColor Yellow
                
                $seriaisRemover = @()
                foreach ($idx in $indicesParaRemover) {
                    $item = $tablets[$idx - 1]
                    $seriaisRemover += $item.Serial
                    Write-Host "   -> [$idx] Serial: $($item.Serial) | MAC: $($item.MAC_WiFi) | $($item.Fabricante) $($item.Modelo)" -ForegroundColor Yellow
                }
                Write-Host "  ------------------------------------------------------------------------" -ForegroundColor DarkGray
                Write-Host "  Total de registros selecionados: $($indicesParaRemover.Count)" -ForegroundColor Red
                Write-Host ""
                Write-Host "  Confirmar exclusao definitiva com geracao de backup? (S/N) " -NoNewline -ForegroundColor Red
                $confirma = Read-Host "[N]"

                if ($confirma.Trim() -match '^(?i)s|sim|y|yes$') {
                    # 1. Filtra registros mantidos
                    $novaListaTablets = @()
                    for ($k = 0; $k -lt $tablets.Count; $k++) {
                        $num1Based = $k + 1
                        if ($indicesParaRemover -notcontains $num1Based) {
                            $novaListaTablets += $tablets[$k]
                        }
                    }

                    # 2. Salva com backup
                    & $salvarTablets $caminhoCsvTablets $novaListaTablets
                    Write-Host ""
                    Write-Host "  [OK] $($indicesParaRemover.Count) registro(s) de tablet removido(s) com sucesso!" -ForegroundColor Green
                    Write-Host "  [OK] Backup de seguranca gerado automaticamente na pasta 'Backups'." -ForegroundColor Green

                    # 3. Verifica se algum dos seriais tambem consta no inventario geral
                    $temNoGeral = @($regsGerais | Where-Object { $seriaisRemover -contains $_.Numero_Serie })
                    if ($temNoGeral.Count -gt 0) {
                        Write-Host ""
                        Write-Host "  [INFO] $($temNoGeral.Count) dos tablets excluidos tambem possuem ficha no Inventario Geral TI:" -ForegroundColor Yellow
                        foreach ($tg in $temNoGeral) {
                            Write-Host "    -> $($tg.Patrimonio) - $($tg.Usuario) ($($tg.Numero_Serie))" -ForegroundColor Cyan
                        }
                        Write-Host "  Deseja tambem excluir estas fichas do Inventario Geral? (S/N) " -NoNewline -ForegroundColor Yellow
                        $confGeral = Read-Host "[N]"
                        if ($confGeral.Trim() -match '^(?i)s|sim|y|yes$') {
                            $novaListaGeral = @($regsGerais | Where-Object { $seriaisRemover -notcontains $_.Numero_Serie })
                            Salvar-RegistrosCsv -caminho $caminhoCsvPrincipal -registros $novaListaGeral
                            Write-Host "  [OK] Registros tambem removidos do Inventario Geral com backup!" -ForegroundColor Green
                        }
                    }
                } else {
                    Write-Host "`n  Exclusao cancelada pelo usuario. Nenhum registro foi alterado." -ForegroundColor Yellow
                }

                Pausar-Tela
            }

            "5" {
                # EXPORTAR RELATORIO DE TABLETS PARA EXCEL UTF-8 / CSV
                try { Clear-Host } catch {}
                Exibir-LogoPrincipal
                Write-Host " |            >>> EXPORTACAO DE RELATORIO DE TABLETS <<<                |" -ForegroundColor Green
                Write-Host " +======================================================================+" -ForegroundColor Cyan
                Write-Host ""

                if ($tablets.Count -eq 0) {
                    Write-Host "  Nenhum tablet registrado para exportar." -ForegroundColor Yellow
                    Pausar-Tela
                    continue
                }

                if (!(Test-Path $script:pastaExport)) {
                    New-Item -ItemType Directory -Path $script:pastaExport -Force | Out-Null
                }

                $dtStr = Get-Date -Format "yyyyMMdd_HHmmss"
                $arqExport = Join-Path $script:pastaExport "relatorio_tablets_android_$dtStr.csv"

                & $salvarTablets $arqExport $tablets

                Write-Host "  [OK] Relatorio de Tablets exportado com sucesso!" -ForegroundColor Green
                Write-Host "  Local: " -NoNewline -ForegroundColor Gray
                Write-Host "$arqExport" -ForegroundColor Yellow
                Write-Host ""
                Write-Host "  Deseja abrir o arquivo agora no Excel? (S/N) " -NoNewline -ForegroundColor Cyan
                $abrir = Read-Host "[S]"
                if ($abrir.Trim() -match '^(?i)s|sim|y|yes$|^$') {
                    try { Start-Process $arqExport } catch {
                        Write-Host "  [!] Nao foi possivel abrir o programa padrao." -ForegroundColor Yellow
                    }
                }
                Pausar-Tela
            }

            "0" {
                $emMenuMobile = $false
            }

            default {
                Write-Host " Opcao invalida." -ForegroundColor Red
                Start-Sleep -Milliseconds 600
            }
        }
    }
}

# ==============================================================================
# LOOP DO MENU PRINCIPAL (MODULAR: DESKTOP VS MOBILE)
# ==============================================================================
$executando = $true

while ($executando) {
    try { Clear-Host } catch {}
    Exibir-LogoPrincipal
    Write-Host " +======================================================================+" -ForegroundColor Cyan
    Write-Host " |      [1] MODULO COMPUTADORES & DESKTOPS (WINDOWS / MACOS)            |" -ForegroundColor Yellow
    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[1]" -NoNewline -ForegroundColor Green;  Write-Host " Novo Cadastro PC      (Coleta pecas e hardware de PCs/Laptops)  |" -ForegroundColor White
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[2]" -NoNewline -ForegroundColor Yellow; Write-Host " Consultar Computadores(Buscar por Patrimonio, Setor ou Usuario) |" -ForegroundColor White
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[3]" -NoNewline -ForegroundColor Cyan;   Write-Host " Editar Computador     (Alterar pecas, usuario ou departamento)  |" -ForegroundColor White
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[4]" -NoNewline -ForegroundColor Red;    Write-Host " Excluir Computador    (Remover PC com backup de seguranca)      |" -ForegroundColor White
    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
    Write-Host " |      [2] MODULO MOBILE & TABLETS (ANDROID VIA CABO USB / ADB)        |" -ForegroundColor Magenta
    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[5]" -NoNewline -ForegroundColor Magenta;Write-Host " Central Gestao Mobile (Lista de Tablets, Exclusao e Relatorios) |" -ForegroundColor White
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[6]" -NoNewline -ForegroundColor Green;  Write-Host " Auto-Scanner Bancada  (Plug & Play USB - Captura MAC wlan0)     |" -ForegroundColor White
    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
    Write-Host " |      [3] FERRAMENTAS GERAIS, EXPORTACAO & NUVEM                      |" -ForegroundColor Cyan
    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[7]" -NoNewline -ForegroundColor Green;  Write-Host " Central de Exportacao (Gerar planilhas Excel UTF-8 / CSV)       |" -ForegroundColor White
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[8]" -NoNewline -ForegroundColor Yellow; Write-Host " Dashboard & Metricas  (Graficos ASCII e resumo executivo)       |" -ForegroundColor White
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[9]" -NoNewline -ForegroundColor Green;  Write-Host " Sincronizar InnovStock(Enviar dados para nuvem corporativa)     |" -ForegroundColor White
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[0]" -NoNewline -ForegroundColor DarkGray;Write-Host " Sair do Sistema                                                 |" -ForegroundColor Gray
    Write-Host " +======================================================================+" -ForegroundColor Cyan
    Write-Host ""
    
    $opcao = Ler-OpcaoOuEsc -prompt " Digite o numero da opcao desejada (ou ESC para sair)" -valorPadrao "1"
    
    switch ($opcao.Trim()) {
        "1" {
            # ------------------------------------------------------------------
            # OPCAO 1: NOVO CADASTRO (AUTOMATICO OU MANUAL OU ANDROID ADB)
            # ------------------------------------------------------------------
            try { Clear-Host } catch {}
            Exibir-LogoPrincipal
            Write-Host " |                >>> INICIAR NOVO CADASTRO TI <<<                      |" -ForegroundColor Yellow
            Write-Host " +======================================================================+" -ForegroundColor Cyan
            Write-Host ""
            Write-Host " Como deseja cadastrar o equipamento?" -ForegroundColor Yellow
            Write-Host "   [1] MODO AUTOMATICO - Coletar pecas e hardware DESTE computador" -ForegroundColor Green
            Write-Host "   [2] MODO MANUAL     - Digitar manualmente os dados de QUALQUER equipamento" -ForegroundColor Cyan
            Write-Host "   [3] MODO ANDROID    - Scanner de Tablets Android via Cabo USB (ADB)" -ForegroundColor Magenta
            Write-Host "   [0] Voltar ao Menu Principal" -ForegroundColor Gray
            Write-Host ""
            $modoCad = Ler-OpcaoOuEsc -prompt " Escolha o modo de cadastro (ou ESC para voltar)" -valorPadrao "1"
            
            $registrosAtuais = Obter-RegistrosCsv -caminho $caminhoCsv
            
            if ($modoCad.Trim() -eq "3" -or $modoCad.Trim() -match '(?i)^android|tablet|adb$') {
                # MODO TABLET ANDROID ADB
                Iniciar-CadastroAndroidTablet -registrosAtuais $registrosAtuais -caminhoCsv $caminhoCsv
            } elseif ($modoCad.Trim() -eq "2" -or $modoCad.Trim() -match '(?i)^m|manual$') {
                # ENTRADA MANUAL COMPLETA
                Iniciar-CadastroManual -registrosAtuais $registrosAtuais -caminhoCsv $caminhoCsv
            } elseif ($modoCad.Trim() -eq "1" -or $modoCad.Trim() -match '(?i)^a|auto$') {
                # MODO AUTOMATICO (SCANNER)
                try { Clear-Host } catch {}
                Exibir-LogoPrincipal
                Write-Host " |            >>> NOVO CADASTRO - SCANNER AUTOMATICO <<<                |" -ForegroundColor Yellow
                Write-Host " +======================================================================+" -ForegroundColor Cyan
                Write-Host ""
                
                $patrimonioSugerido = Obter-ProximoPatrimonio -registros $registrosAtuais
                
                $setoresPadrao = @(
                    "Operacional", "Pedagogico", "Financeiro", "Marketing", "Comercial",
                    "Administrativo", "Administrativo / Financeiro", "Logistica",
                    "Motoristas", "Servicos Gerais / Limpeza", "TI", "RH"
                )
                $setoresHistorico = @()
                if ($registrosAtuais.Count -gt 0) {
                    $setoresHistorico = @($registrosAtuais | ForEach-Object { $_.Setor.Trim() } | Where-Object { $_ -ne "" -and $_ -notmatch '^\d+$' } | Select-Object -Unique)
                }
                $listaSetores = @()
                foreach ($s in $setoresHistorico) { if ($listaSetores -notcontains $s) { $listaSetores += $s } }
                foreach ($s in $setoresPadrao) { if ($listaSetores -notcontains $s) { $listaSetores += $s } }
                
                # [1] EMPRESA
                Write-Host " [1/5] IDENTIFICACAO DA EMPRESA" -ForegroundColor Yellow
                Write-Host "   Pressione ENTER para manter " -NoNewline -ForegroundColor Gray
                Write-Host "'InnovTech'" -ForegroundColor Green
                $empresaInput = Read-Host "   Empresa"
                $empresa = if ([string]::IsNullOrWhiteSpace($empresaInput)) { "InnovTech" } else { $empresaInput.Trim() }
                
                # [2] USUARIO
                Write-Host "`n [2/5] DADOS DO USUARIO / RESPONSAVEL" -ForegroundColor Yellow
                $usuario = ""
                while ([string]::IsNullOrWhiteSpace($usuario)) {
                    $usuario = Read-Host "   Nome Completo do Colaborador"
                    if ([string]::IsNullOrWhiteSpace($usuario)) {
                        Write-Host "   (!) Por favor, digite o nome do usuario." -ForegroundColor Red
                    }
                }
                
                # [3] FUNCAO
                Write-Host "`n [3/5] FUNCAO / CARGO" -ForegroundColor Yellow
                $funcao = ""
                while ([string]::IsNullOrWhiteSpace($funcao)) {
                    $funcao = Read-Host "   Cargo / Funcao (Ex: Professor, Operador, Motorista, Analista)"
                    if ([string]::IsNullOrWhiteSpace($funcao)) {
                        Write-Host "   (!) Por favor, digite a funcao." -ForegroundColor Red
                    }
                }
                
                # [4] SETOR
                Write-Host "`n [4/5] SETOR / DEPARTAMENTO" -ForegroundColor Yellow
                Write-Host "   Escolha um setor da lista ou digite um novo:" -ForegroundColor Cyan
                for ($i = 0; $i -lt $listaSetores.Count; $i++) {
                    $numFormatado = ($i + 1).ToString().PadLeft(2, ' ')
                    $ehHistorico = ($setoresHistorico -contains $listaSetores[$i])
                    $corSetor = if ($ehHistorico) { "Green" } else { "White" }
                    $qtdSetor = @($registrosAtuais | Where-Object { $_.Setor -eq $listaSetores[$i] }).Count
                    $tagQtd = if ($qtdSetor -gt 0) { " ($qtdSetor PCs)" } else { "" }
                    Write-Host "     [$numFormatado] " -NoNewline -ForegroundColor Cyan
                    Write-Host "$($listaSetores[$i])$tagQtd" -ForegroundColor $corSetor
                }
                Write-Host "     [ 0] Digitar um outro setor novo" -ForegroundColor Gray
                
                $setor = ""
                while ([string]::IsNullOrWhiteSpace($setor)) {
                    $setorInput = Read-Host "`n   Escolha o numero ou digite o nome do setor"
                    if ([string]::IsNullOrWhiteSpace($setorInput)) {
                        Write-Host "   (!) O setor nao pode ficar vazio." -ForegroundColor Red
                        continue
                    }
                    $numSetor = 0
                    if ([int]::TryParse($setorInput.Trim(), [ref]$numSetor)) {
                        if ($numSetor -ge 1 -and $numSetor -le $listaSetores.Count) {
                            $setor = $listaSetores[$numSetor - 1]
                        } elseif ($numSetor -eq 0) {
                            while ([string]::IsNullOrWhiteSpace($setor)) {
                                $setor = Read-Host "   Digite o nome do novo setor"
                            }
                        } else {
                            $setor = $setorInput.Trim()
                        }
                    } else {
                        $setor = $setorInput.Trim()
                    }
                }
                Write-Host "   -> Setor selecionado: " -NoNewline -ForegroundColor Gray
                Write-Host "$setor" -ForegroundColor Green
                
                # [5] PATRIMONIO
                Write-Host "`n [5/5] NUMERO DE PATRIMONIO" -ForegroundColor Yellow
                Write-Host "   Pressione ENTER para aceitar o sugerido: " -NoNewline -ForegroundColor Gray
                Write-Host "[$patrimonioSugerido]" -ForegroundColor Green
                $patrimonioInput = Read-Host "   Patrimonio"
                $patrimonio = if ([string]::IsNullOrWhiteSpace($patrimonioInput)) { $patrimonioSugerido } else { $patrimonioInput.Trim() }
                
                # COLETA DE HARDWARE
                Write-Host ""
                Write-Host " [*] Coletando especificacoes tecnicas completas do computador... Aguarde..." -ForegroundColor Green
                
                $hw = Coletar-HardwareCompleto
                
                $divisoria = "=" * 64
                $resumoTexto = @(
                    $divisoria,
                    "       INNOVCORE TI - RESUMO DO INVENTARIO ($empresa)",
                    $divisoria,
                    "  - Empresa:            $empresa",
                    "  - Patrimonio:         $patrimonio",
                    "  - Usuario:            $usuario",
                    "  - Funcao / Cargo:     $funcao",
                    "  - Setor:              $setor",
                    "",
                    "  - Tipo Equipamento:   $($hw.Tipo_Equipamento)",
                    "  - Marca:              $($hw.Marca)",
                    "  - Modelo:             $($hw.Modelo)",
                    "  - Numero de Serie:    $($hw.Numero_Serie)",
                    "  - Ano de Fabricacao:  $($hw.Ano_Fabricacao)",
                    "  - Computador (Host):  $($hw.Nome_Computador)",
                    "  - Processador (CPU):  $($hw.Processador)",
                    "  - Memoria RAM:        $($hw.Memoria_RAM)",
                    "  - Armazenamento:      $($hw.Armazenamento)",
                    "  - Placa de Video:     $($hw.Placa_Video)",
                    "  - Sistema Operacional:$($hw.Sistema_Operacional)",
                    "  - Endereco IP / MAC:  $($hw.IP_Rede) | $($hw.MAC_Rede)",
                    "  - Data do Registro:   $($hw.Data_Registro)",
                    $divisoria
                ) -join "`r`n"
                
                try { Clear-Host } catch {}
                Exibir-LogoPrincipal
                Write-Host ""
                Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
                Write-Host " |                 RESUMO DAS ESPECIFICACOES COLETADAS                  |" -ForegroundColor Yellow
                Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
                Write-Host "   EMPRESA:          " -NoNewline -ForegroundColor Gray; Write-Host "$empresa" -ForegroundColor White
                Write-Host "   PATRIMONIO:       " -NoNewline -ForegroundColor Gray; Write-Host "$patrimonio" -ForegroundColor Green
                Write-Host "   COLABORADOR:      " -NoNewline -ForegroundColor Gray; Write-Host "$usuario" -ForegroundColor White
                Write-Host "   CARGO / FUNCAO:   " -NoNewline -ForegroundColor Gray; Write-Host "$funcao" -ForegroundColor White
                Write-Host "   SETOR:            " -NoNewline -ForegroundColor Gray; Write-Host "$setor" -ForegroundColor Green
                Write-Host "  ----------------------------------------------------------------------" -ForegroundColor DarkGray
                Write-Host "   FABRICANTE/MARCA: " -NoNewline -ForegroundColor Gray; Write-Host "$($hw.Marca)" -ForegroundColor Cyan
                Write-Host "   MODELO DO PC:     " -NoNewline -ForegroundColor Gray; Write-Host "$($hw.Modelo)" -ForegroundColor Cyan
                Write-Host "   NUMERO DE SERIE:  " -NoNewline -ForegroundColor Gray; Write-Host "$($hw.Numero_Serie)" -ForegroundColor Yellow
                Write-Host "   ANO FABRICACAO:   " -NoNewline -ForegroundColor Gray; Write-Host "$($hw.Ano_Fabricacao)" -ForegroundColor Yellow
                Write-Host "   NOME DO PC (HOST):" -NoNewline -ForegroundColor Gray; Write-Host "$($hw.Nome_Computador)" -ForegroundColor White
                Write-Host "   PROCESSADOR:      " -NoNewline -ForegroundColor Gray; Write-Host "$($hw.Processador)" -ForegroundColor White
                Write-Host "   MEMORIA RAM:      " -NoNewline -ForegroundColor Gray; Write-Host "$($hw.Memoria_RAM)" -ForegroundColor White
                Write-Host "   DATA DA COLETA:   " -NoNewline -ForegroundColor Gray; Write-Host "$($hw.Data_Registro)" -ForegroundColor Gray
                Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
                
                try { Set-Clipboard -Value $resumoTexto } catch { try { $resumoTexto | clip.exe } catch {} }
                Write-Host ""
                Write-Host " [OK] Informacoes COPIADAS para a Area de Transferencia (Ctrl + V)!" -ForegroundColor Green
                
                $novoItem = [PSCustomObject]@{
                    Data_Registro       = $hw.Data_Registro
                    Empresa             = $empresa
                    Patrimonio          = $patrimonio
                    Usuario             = $usuario
                    Funcao              = $funcao
                    Setor               = $setor
                    Tipo_Equipamento    = $hw.Tipo_Equipamento
                    Marca               = $hw.Marca
                    Modelo              = $hw.Modelo
                    Numero_Serie        = $hw.Numero_Serie
                    Ano_Fabricacao      = $hw.Ano_Fabricacao
                    Nome_Computador     = $hw.Nome_Computador
                    Processador         = $hw.Processador
                    Memoria_RAM         = $hw.Memoria_RAM
                    Armazenamento       = $hw.Armazenamento
                    Placa_Video         = $hw.Placa_Video
                    Sistema_Operacional = $hw.Sistema_Operacional
                    IP_Rede             = $hw.IP_Rede
                    MAC_Rede            = $hw.MAC_Rede
                }
                
                # Verificacao de Duplicidade Inteligente
                $dup = Testar-DuplicidadeEquipamento -registros $registrosAtuais -patrimonio $novoItem.Patrimonio -numeroSerie $novoItem.Numero_Serie -macRede $novoItem.MAC_Rede -nomeComputador $novoItem.Nome_Computador
                
                if ($dup) {
                    $itemAntigo = $dup.RegistroExistente
                    Write-Host ""
                    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Yellow
                    Write-Host " |               [!] ALERTA: DUPLICIDADE DETECTADA                      |" -ForegroundColor Yellow
                    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Yellow
                    Write-Host "   Motivo:             " -NoNewline -ForegroundColor Gray; Write-Host "$($dup.Motivo)" -ForegroundColor Red
                    Write-Host "   Registro Anterior:  " -NoNewline -ForegroundColor Gray; Write-Host "$($itemAntigo.Patrimonio) - $($itemAntigo.Usuario) ($($itemAntigo.Setor))" -ForegroundColor Cyan
                    Write-Host "   Data Anterior:      " -NoNewline -ForegroundColor Gray; Write-Host "$($itemAntigo.Data_Registro)" -ForegroundColor DarkGray
                    Write-Host ""
                    Write-Host "   Como deseja proceder?" -ForegroundColor Yellow
                    Write-Host "     [1] ATUALIZAR cadastro existente (Sobrescrever dados com backup)" -ForegroundColor Green
                    Write-Host "     [2] MANTER AMBOS e gerar novo numero de patrimonio" -ForegroundColor Cyan
                    Write-Host "     [0] CANCELAR gravacao deste cadastro" -ForegroundColor Gray
                    Write-Host ""
                    $opcDup = Read-Host "   Escolha uma opcao [1]"
                    if ([string]::IsNullOrWhiteSpace($opcDup)) { $opcDup = "1" }
                    
                    if ($opcDup.Trim() -eq "1") {
                        $novoItem.Patrimonio = $itemAntigo.Patrimonio
                        $listaAtualizada = @($registrosAtuais | Where-Object { $_.Patrimonio -ne $itemAntigo.Patrimonio }) + $novoItem
                        Salvar-RegistrosCsv -caminho $caminhoCsv -registros $listaAtualizada
                        Write-Host "`n [OK] Cadastro $($novoItem.Patrimonio) ATUALIZADO com sucesso no Pen Drive!" -ForegroundColor Green
                        Pausar-Tela -mensagem " Pressione ENTER para voltar ao menu..."
                    } elseif ($opcDup.Trim() -eq "2") {
                        $novoPat = Obter-ProximoPatrimonio -registros $registrosAtuais
                        $novoItem.Patrimonio = $novoPat
                        $listaAtualizada = @($registrosAtuais) + $novoItem
                        Salvar-RegistrosCsv -caminho $caminhoCsv -registros $listaAtualizada
                        Write-Host "`n [OK] Gravado como NOVO equipamento com patrimonio: $novoPat!" -ForegroundColor Green
                        Pausar-Tela -mensagem " Pressione ENTER para voltar ao menu..."
                    } else {
                        Write-Host "`n [!] Operacao cancelada pelo usuario. Nenhum dado alterado." -ForegroundColor Yellow
                        Pausar-Tela -mensagem " Pressione ENTER para voltar ao menu..."
                    }
                } else {
                    $listaAtualizada = @($registrosAtuais) + $novoItem
                    Salvar-RegistrosCsv -caminho $caminhoCsv -registros $listaAtualizada
                    
                    Write-Host " [OK] Registrado com sucesso na planilha CSV do Pendrive!" -ForegroundColor Green
                    Write-Host "      Arquivo: " -NoNewline -ForegroundColor Gray
                    Write-Host "$caminhoCsv" -ForegroundColor Yellow
                    
                    Pausar-Tela -mensagem " Cadastro automatico finalizado com sucesso! Pressione ENTER para voltar ao menu..."
                }
            }
        }
        
        "2" {
            # ------------------------------------------------------------------
            # OPCAO 2: CONSULTAR / PESQUISAR CADASTROS
            # ------------------------------------------------------------------
            $registros = Obter-RegistrosCsv -caminho $caminhoCsv
            Menu-ConsultarRegistros -registros $registros
        }
        
        "3" {
            # ------------------------------------------------------------------
            # OPCAO 3: EDITAR CADASTRO
            # ------------------------------------------------------------------
            $registros = Obter-RegistrosCsv -caminho $caminhoCsv
            Menu-EditarCadastro -registros $registros
        }
        
        "4" {
            # ------------------------------------------------------------------
            # OPCAO 4: EXCLUIR CADASTRO
            # ------------------------------------------------------------------
            try { Clear-Host } catch {}
            Exibir-LogoPrincipal
            Write-Host " |                >>> EXCLUSAO DE CADASTRO <<<                          |" -ForegroundColor Red
            Write-Host " +======================================================================+" -ForegroundColor Cyan
            Write-Host ""
            
            $registros = Obter-RegistrosCsv -caminho $caminhoCsv
            
            if ($registros.Count -eq 0) {
                Write-Host "  Nao ha registros para excluir." -ForegroundColor Yellow
                Pausar-Tela
            } else {
                $indice = 1
                foreach ($r in $registros) {
                    Write-Host "  [$indice] " -NoNewline -ForegroundColor Cyan
                    Write-Host "$($r.Patrimonio) " -NoNewline -ForegroundColor Green
                    Write-Host "- $($r.Usuario) " -NoNewline -ForegroundColor White
                    Write-Host "($($r.Setor) - $($r.Marca) $($r.Modelo))" -ForegroundColor Gray
                    $indice++
                }
                Write-Host ""
                Write-Host "  Digite o NUMERO do registro que deseja EXCLUIR" -ForegroundColor Yellow
                Write-Host "  (Ou digite 0 para cancelar e voltar ao menu):" -ForegroundColor Gray
                $respExcluir = Ler-OpcaoOuEsc -prompt "  Registro a excluir (ou ESC para cancelar)" -valorPadrao "0"
                
                $numExcluir = 0
                if ([int]::TryParse($respExcluir.Trim(), [ref]$numExcluir)) {
                    if ($numExcluir -ge 1 -and $numExcluir -le $registros.Count) {
                        $itemRemovido = $registros[$numExcluir - 1]
                        Write-Host ""
                        Write-Host "  [ATENCAO] Deseja realmente excluir o cadastro de:" -ForegroundColor Red
                        Write-Host "    Patrimonio: " -NoNewline -ForegroundColor Gray
                        Write-Host "$($itemRemovido.Patrimonio)" -NoNewline -ForegroundColor Green
                        Write-Host " - $($itemRemovido.Usuario) ($($itemRemovido.Setor))" -ForegroundColor Yellow
                        Write-Host "  Confirmar exclusao? (S/N) " -NoNewline -ForegroundColor Red
                        $confirma = Read-Host "[N]"
                        
                        if ($confirma.Trim() -match '^(?i)s|sim|y|yes$') {
                            $novaLista = @()
                            for ($j = 0; $j -lt $registros.Count; $j++) {
                                if ($j -ne ($numExcluir - 1)) {
                                    $novaLista += $registros[$j]
                                }
                            }
                            Salvar-RegistrosCsv -caminho $caminhoCsv -registros $novaLista
                            Write-Host ""
                            Write-Host "  [OK] Registro excluido com sucesso! Backup de seguranca gerado." -ForegroundColor Green
                        } else {
                            Write-Host ""
                            Write-Host "  Exclusao cancelada pelo usuario." -ForegroundColor Yellow
                        }
                    } elseif ($numExcluir -eq 0) {
                        Write-Host "  Operacao cancelada." -ForegroundColor Gray
                    } else {
                        Write-Host "  Numero invalido." -ForegroundColor Red
                    }
                } else {
                    Write-Host "  Opcao invalida." -ForegroundColor Red
                }
                
                Pausar-Tela
            }
        }
        
        "5" {
            # ------------------------------------------------------------------
            # OPCAO 5: CENTRAL DE GESTAO MOBILE & TABLETS (ANDROID / ADB)
            # ------------------------------------------------------------------
            Menu-GestaoMobileTablets
        }

        "m" {
            Menu-GestaoMobileTablets
        }

        "mobile" {
            Menu-GestaoMobileTablets
        }

        "tablet" {
            Menu-GestaoMobileTablets
        }

        "tablets" {
            Menu-GestaoMobileTablets
        }
        
        "6" {
            # ------------------------------------------------------------------
            # OPCAO 6: AUTO-SCANNER BANCADA PLUG & PLAY (USB)
            # ------------------------------------------------------------------
            $scriptAutoScan = Join-Path $scriptRootDir "mobile\capturar_mac_tablet.ps1"
            if (-not (Test-Path $scriptAutoScan)) {
                $scriptAutoScan = Join-Path $scriptRootDir "adb\capturar_mac_tablet.ps1"
            }
            if (Test-Path $scriptAutoScan) {
                & $scriptAutoScan
            } else {
                Write-Host " [!] Script de auto-scanner nao localizado em: $scriptAutoScan" -ForegroundColor Red
                Pausar-Tela
            }
        }

        "scanner" {
            $scriptAutoScan = Join-Path $scriptRootDir "mobile\capturar_mac_tablet.ps1"
            if (-not (Test-Path $scriptAutoScan)) { $scriptAutoScan = Join-Path $scriptRootDir "adb\capturar_mac_tablet.ps1" }
            if (Test-Path $scriptAutoScan) { & $scriptAutoScan } else { Pausar-Tela }
        }

        "bancada" {
            $scriptAutoScan = Join-Path $scriptRootDir "mobile\capturar_mac_tablet.ps1"
            if (-not (Test-Path $scriptAutoScan)) { $scriptAutoScan = Join-Path $scriptRootDir "adb\capturar_mac_tablet.ps1" }
            if (Test-Path $scriptAutoScan) { & $scriptAutoScan } else { Pausar-Tela }
        }
        
        "7" {
            # ------------------------------------------------------------------
            # OPCAO 7: CENTRAL DE EXPORTACAO EXCEL UTF-8 / FORMATADOS
            # ------------------------------------------------------------------
            $registros = Obter-RegistrosCsv -caminho $caminhoCsv
            Menu-Exportacao -registros $registros
        }
        
        "8" {
            # ------------------------------------------------------------------
            # OPCAO 8: DASHBOARD EXECUTIVO E METRICAS
            # ------------------------------------------------------------------
            $registros = Obter-RegistrosCsv -caminho $caminhoCsv
            Menu-Estatisticas -registros $registros
        }

        "9" {
            # ------------------------------------------------------------------
            # OPCAO 9: SINCRONIZAR COM INNOVSTOCK (FIREBASE CLOUD)
            # ------------------------------------------------------------------
            $scriptSync = Join-Path $scriptRootDir "core\sincronizar_innovstock.ps1"
            if (-not (Test-Path $scriptSync)) {
                $scriptSync = Join-Path $scriptRootDir "sincronizar_innovstock.ps1"
            }
            if (Test-Path $scriptSync) {
                & $scriptSync
            } else {
                Write-Host " [!] Script de sincronizacao nao encontrado em: $scriptSync" -ForegroundColor Red
                Pausar-Tela
            }
        }
        
        "0" {
            # ------------------------------------------------------------------
            # SAIR
            # ------------------------------------------------------------------
            $executando = $false
            try { Clear-Host } catch {}
            Exibir-LogoPrincipal
            Write-Host ""
            Write-Host "  Programa encerrado. Bom trabalho, Jordan!" -ForegroundColor Green
            Write-Host ""
            Start-Sleep -Milliseconds 1200
        }
        
        default {
            Write-Host " Opcao invalida. Tente novamente." -ForegroundColor Red
            Start-Sleep -Milliseconds 800
        }
    }
}
