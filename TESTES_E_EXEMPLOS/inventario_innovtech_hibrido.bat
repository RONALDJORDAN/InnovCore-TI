@echo off
setlocal
chcp 65001 >nul 2>&1
title InnovTech - Inventario de Equipamentos
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$c = [System.IO.File]::ReadAllText('%~f0', [System.Text.Encoding]::UTF8); $c = $c.Substring($c.LastIndexOf('#---POWERSHELL---#') + 19); & ([ScriptBlock]::Create($c)) '%~dp0'"
exit /b %ERRORLEVEL%

#---POWERSHELL---#
param([string]$scriptDir = (Get-Location).Path)

# ----------------------------------------------------------------------------
# INNOVTECH - INVENTÁRIO AUTOMÁTICO DE HARDWARE E USUÁRIO
# Compatível com Windows 10 e Windows 11 (32 e 64 bits)
# ----------------------------------------------------------------------------

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try { $Host.UI.RawUI.WindowTitle = "InnovTech - Inventário de Equipamentos" } catch {}

try { Clear-Host } catch {}
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "                  INNOVTECH - INVENTÁRIO DE TI                          " -ForegroundColor Yellow
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  Este assistente coleta as especificações completas deste computador   " -ForegroundColor Gray
Write-Host "  e vincula ao colaborador e setor responsável na InnovTech.           " -ForegroundColor Gray
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

# ---------------------------------------------------------
# 1. PERGUNTAS INTERATIVAS
# ---------------------------------------------------------
Write-Host "[1/4] IDENTIFICAÇÃO DA EMPRESA" -ForegroundColor Yellow
$empresaInput = Read-Host "  Empresa [Pressione ENTER para 'InnovTech']"
$empresa = if ([string]::IsNullOrWhiteSpace($empresaInput)) { "InnovTech" } else { $empresaInput.Trim() }

Write-Host "`n[2/4] DADOS DO USUÁRIO / RESPONSÁVEL" -ForegroundColor Yellow
$colaborador = ""
while ([string]::IsNullOrWhiteSpace($colaborador)) {
    $colaborador = Read-Host "  De quem é esse computador (Nome Completo)"
    if ([string]::IsNullOrWhiteSpace($colaborador)) {
        Write-Host "  (!) O nome do colaborador não pode ficar vazio." -ForegroundColor Red
    }
}

Write-Host "`n[3/4] FUNÇÃO / CARGO" -ForegroundColor Yellow
$funcao = ""
while ([string]::IsNullOrWhiteSpace($funcao)) {
    $funcao = Read-Host "  Qual a função / cargo da pessoa (Ex: Analista, Gerente, etc.)"
    if ([string]::IsNullOrWhiteSpace($funcao)) {
        Write-Host "  (!) A função não pode ficar vazia." -ForegroundColor Red
    }
}

Write-Host "`n[4/4] SETOR / DEPARTAMENTO" -ForegroundColor Yellow
$setor = ""
while ([string]::IsNullOrWhiteSpace($setor)) {
    $setor = Read-Host "  Qual o setor da empresa (Ex: TI, Financeiro, RH, Comercial)"
    if ([string]::IsNullOrWhiteSpace($setor)) {
        Write-Host "  (!) O setor não pode ficar vazio." -ForegroundColor Red
    }
}

Write-Host "`n[OPCIONAL] NÚMERO DE PATRIMÔNIO / ETIQUETA" -ForegroundColor Yellow
$patrimonioInput = Read-Host "  Etiqueta de Patrimônio (ou ENTER se não houver)"
$patrimonio = if ([string]::IsNullOrWhiteSpace($patrimonioInput)) { "N/A" } else { $patrimonioInput.Trim() }

Write-Host ""
Write-Host "[*] Coletando dados de hardware e sistema... Por favor, aguarde..." -ForegroundColor Green

# ---------------------------------------------------------
# 2. COLETA DE HARDWARE E SISTEMA
# ---------------------------------------------------------
$dataColeta = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
$hostname = $env:COMPUTERNAME
$usuarioWindows = [System.Environment]::UserName

# Serial Number & BIOS
$bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
$serialBios = if ($bios.SerialNumber) { $bios.SerialNumber.Trim() } else { "" }
$biosVersao = if ($bios.SMBIOSBIOSVersion) { $bios.SMBIOSBIOSVersion.Trim() } else { "N/A" }

# Placa-Mãe
$bb = Get-CimInstance Win32_BaseBoard -ErrorAction SilentlyContinue
$serialPlaca = if ($bb.SerialNumber) { $bb.SerialNumber.Trim() } else { "" }
$placaMae = "$($bb.Manufacturer) $($bb.Product)".Trim()

# Serial Definitivo (com fallback se a BIOS for genérica)
$serialNumber = if ($serialBios -and $serialBios -notmatch '(?i)Default|To be filled|None|000000|System Serial Number') { 
    $serialBios 
} elseif ($serialPlaca -and $serialPlaca -notmatch '(?i)Default|To be filled|None|000000') { 
    $serialPlaca 
} elseif ($serialBios) {
    $serialBios
} else { 
    "N/A" 
}

# Computador (Fabricante / Modelo)
$cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
$fabricante = if ($cs.Manufacturer) { $cs.Manufacturer.Trim() } else { "N/A" }
$modelo = if ($cs.Model) { $cs.Model.Trim() } else { "N/A" }

# Tipo do Equipamento
$chassis = Get-CimInstance Win32_SystemEnclosure -ErrorAction SilentlyContinue
$tipoChassisNum = if ($chassis.ChassisTypes) { $chassis.ChassisTypes[0] } else { 0 }
$tipoEquipamento = switch ($tipoChassisNum) {
    { $_ -in 8, 9, 10, 11, 12, 14, 18, 21, 31, 32 } { "Notebook / Laptop" }
    { $_ -in 3, 4, 5, 6, 7, 15, 16 } { "Desktop / PC de Mesa" }
    { $_ -in 13 } { "All-in-One" }
    { $_ -in 23 } { "Servidor" }
    default { "Computador" }
}

# Processador (CPU)
$cpu = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
$cpuNome = if ($cpu.Name) { ($cpu.Name.Trim() -replace '\s+', ' ') } else { "N/A" }
$cpuCores = if ($cpu.NumberOfCores) { $cpu.NumberOfCores } else { "N/A" }
$cpuThreads = if ($cpu.NumberOfLogicalProcessors) { $cpu.NumberOfLogicalProcessors } else { "N/A" }

# Memória RAM
$ramModulos = Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue
$ramDetalhesList = @()
$ramTotalBytes = 0

if ($ramModulos) {
    foreach ($m in $ramModulos) {
        $ramTotalBytes += $m.Capacity
        $capGB = [math]::Round($m.Capacity / 1GB, 0)
        $speed = if ($m.Speed) { "$($m.Speed)MHz" } else { "" }
        $part = if ($m.PartNumber) { $m.PartNumber.Trim() } else { "" }
        $ramDetalhesList += "$capGB GB $speed $part".Trim()
    }
}

$ramTotalGB = if ($ramTotalBytes -gt 0) { 
    [math]::Round($ramTotalBytes / 1GB, 0) 
} elseif ($cs.TotalPhysicalMemory) { 
    [math]::Round($cs.TotalPhysicalMemory / 1GB, 0) 
} else { 
    0 
}

$ramResumo = if ($ramDetalhesList.Count -gt 0) {
    "$ramTotalGB GB RAM ($($ramDetalhesList -join ' + '))"
} else {
    "$ramTotalGB GB RAM"
}

# Armazenamento (Discos Físicos)
$discosFisicos = Get-CimInstance Win32_DiskDrive -ErrorAction SilentlyContinue
$listaDiscos = @()
if ($discosFisicos) {
    foreach ($d in $discosFisicos) {
        $tamGB = [math]::Round($d.Size / 1GB, 0)
        $tipoMedia = ""
        try {
            $pd = Get-PhysicalDisk | Where-Object { $_.DeviceId -eq $d.Index } -ErrorAction SilentlyContinue
            if ($pd.MediaType) { $tipoMedia = "[$($pd.MediaType)] " }
        } catch {}
        $listaDiscos += "$tipoMedia$($d.Model.Trim()) ($tamGB GB)"
    }
}
$armazenamentoDiscos = if ($listaDiscos.Count -gt 0) { $listaDiscos -join " | " } else { "N/A" }

# Partições de Disco (Letras, espaço livre e total)
$volumes = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue
$listaVolumes = @()
if ($volumes) {
    foreach ($v in $volumes) {
        $livre = [math]::Round($v.FreeSpace / 1GB, 1)
        $total = [math]::Round($v.Size / 1GB, 1)
        $listaVolumes += "$($v.DeviceID) ($livre GB livres de $total GB)"
    }
}
$particoesInfo = if ($listaVolumes.Count -gt 0) { $listaVolumes -join " | " } else { "N/A" }

# Placa de Vídeo (GPU)
$gpus = Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue
$gpuNomesList = ($gpus | Where-Object { $_.Name } | ForEach-Object { $_.Name.Trim() } | Select-Object -Unique)
$gpuInfo = if ($gpuNomesList) { $gpuNomesList -join " | " } else { "N/A" }

# Sistema Operacional
$os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
$osNome = if ($os.Caption) { $os.Caption.Trim() } else { "Windows" }
$osArch = if ($os.OSArchitecture) { $os.OSArchitecture } else { "64-bit" }
$osBuild = if ($os.BuildNumber) { $os.BuildNumber } else { "" }
$sistemaOperacional = "$osNome ($osArch) - Build $osBuild"

# Rede (Adaptador ativo com IPv4)
$rede = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" -ErrorAction SilentlyContinue | Select-Object -First 1
$ipV4 = if ($rede.IPAddress) { ($rede.IPAddress | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' } | Select-Object -First 1) } else { "N/A" }
$macAddr = if ($rede.MACAddress) { $rede.MACAddress } else { "N/A" }
$adaptadorRede = if ($rede.Description) { $rede.Description.Trim() } else { "N/A" }

# ---------------------------------------------------------
# 3. CRIAÇÃO DO RELATÓRIO FORMATADO
# ---------------------------------------------------------
$divisoria = "=" * 72
$relatorio = @"
$divisoria
           INVENTÁRIO DE EQUIPAMENTO - $empresa
$divisoria

[1. DADOS DO COLABORADOR E EMPRESA]
  - Empresa:             $empresa
  - Colaborador:         $colaborador
  - Cargo / Função:      $funcao
  - Setor:               $setor
  - Patrimônio / Tag:    $patrimonio
  - Data / Hora Coleta:  $dataColeta

[2. ESPECIFICAÇÕES DO COMPUTADOR]
  - Nome da Máquina:     $hostname
  - Tipo de Equipamento: $tipoEquipamento
  - Usuário Logado:      $usuarioWindows
  - Fabricante:          $fabricante
  - Modelo:              $modelo
  - Placa-Mãe:           $placaMae
  - NÚMERO DE SÉRIE:     $serialNumber
  - Versão da BIOS:      $biosVersao

[3. HARDWARE]
  - Processador (CPU):   $cpuNome
  - Núcleos / Threads:   $cpuCores Cores / $cpuThreads Threads
  - Memória RAM:         $ramResumo
  - Discos Físicos:      $armazenamentoDiscos
  - Espaço em Disco:     $particoesInfo
  - Placa de Vídeo (GPU):$gpuInfo

[4. SISTEMA E REDE]
  - Sistema Operacional: $sistemaOperacional
  - Endereço IPv4:       $ipV4
  - Endereço MAC:        $macAddr
  - Placa de Rede:       $adaptadorRede
$divisoria
"@

# ---------------------------------------------------------
# 4. EXIBIÇÃO NO CONSOLE
# ---------------------------------------------------------
try { Clear-Host } catch {}
Write-Host $relatorio -ForegroundColor Cyan

# ---------------------------------------------------------
# 5. COPIAR PARA A ÁREA DE TRANSFERÊNCIA (CLIPBOARD)
# ---------------------------------------------------------
$copiadoComSucesso = $false
try {
    Set-Clipboard -Value $relatorio
    $copiadoComSucesso = $true
} catch {
    try {
        $relatorio | clip.exe
        $copiadoComSucesso = $true
    } catch {}
}

Write-Host ""
if ($copiadoComSucesso) {
    Write-Host " [✓] AS ESPECIFICAÇÕES FORAM COPIADAS PARA A SUA ÁREA DE TRANSFERÊNCIA!" -ForegroundColor Green
    Write-Host "     -> Pressione Ctrl + V em qualquer lugar para colar o relatório." -ForegroundColor Yellow
} else {
    Write-Host " [!] Não foi possível copiar para a área de transferência automaticamente." -ForegroundColor Yellow
}

# ---------------------------------------------------------
# 6. SALVAR ARQUIVOS NO PENDRIVE
# ---------------------------------------------------------
$pastaDestino = Join-Path $scriptDir "Inventario_InnovTech"
if (!(Test-Path $pastaDestino)) {
    New-Item -ItemType Directory -Path $pastaDestino -Force | Out-Null
}

# 6.1 Salvar Arquivo de Texto Individual
$limparNome = { param($t) $t -replace '[\\/:*?""<>| ]', '_' }
$nomeArquivoTxt = "Inventario_$(& $limparNome $setor)_$(& $limparNome $colaborador)_$(& $limparNome $hostname).txt"
$caminhoTxt = Join-Path $pastaDestino $nomeArquivoTxt

[System.IO.File]::WriteAllText($caminhoTxt, $relatorio, [System.Text.Encoding]::UTF8)
Write-Host " [✓] Arquivo individual salvo: $nomeArquivoTxt" -ForegroundColor Green

# 6.2 Salvar / Atualizar Planilha Geral CSV (Compatível com Excel PT-BR)
$caminhoCsv = Join-Path $pastaDestino "Inventario_Geral_InnovTech.csv"
$cabecalhoCsv = "Data_Coleta;Empresa;Colaborador;Funcao;Setor;Patrimonio;Hostname;Tipo_Equipamento;Usuario_Windows;Serial_Number;Fabricante;Modelo;Placa_Mae;Processador;Cores;Threads;RAM_Total_GB;RAM_Detalhes;Discos;Volumes_Espaco;Placa_Video;Sistema_Operacional;IPv4;MAC"

if (!(Test-Path $caminhoCsv)) {
    # Cria o arquivo CSV com BOM UTF-8 para o Excel abrir com acentos perfeitos
    $utf8ComBOM = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($caminhoCsv, ($cabecalhoCsv + "`r`n"), $utf8ComBOM)
}

# Monta linha CSV escapando eventuais ponto-e-vírgulas
$limparCampo = { param($c) if ($c) { ($c.ToString() -replace ';', ',') -replace "[\r\n]+", " " } else { "" } }
$linhaCsv = @(
    (& $limparCampo $dataColeta),
    (& $limparCampo $empresa),
    (& $limparCampo $colaborador),
    (& $limparCampo $funcao),
    (& $limparCampo $setor),
    (& $limparCampo $patrimonio),
    (& $limparCampo $hostname),
    (& $limparCampo $tipoEquipamento),
    (& $limparCampo $usuarioWindows),
    (& $limparCampo $serialNumber),
    (& $limparCampo $fabricante),
    (& $limparCampo $modelo),
    (& $limparCampo $placaMae),
    (& $limparCampo $cpuNome),
    (& $limparCampo $cpuCores),
    (& $limparCampo $cpuThreads),
    (& $limparCampo $ramTotalGB),
    (& $limparCampo ($ramDetalhesList -join ' + ')),
    (& $limparCampo $armazenamentoDiscos),
    (& $limparCampo $particoesInfo),
    (& $limparCampo $gpuInfo),
    (& $limparCampo $sistemaOperacional),
    (& $limparCampo $ipV4),
    (& $limparCampo $macAddr)
) -join ';'

$utf8SemBOM = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::AppendAllText($caminhoCsv, ($linhaCsv + "`r`n"), [System.Text.Encoding]::UTF8)
Write-Host " [✓] Adicionado à planilha geral: Inventario_Geral_InnovTech.csv" -ForegroundColor Green

Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host " Coleta concluída com sucesso! Pode remover o pendrive ou ir ao próx PC." -ForegroundColor Green
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Pressione ENTER para encerrar..." -ForegroundColor Gray
try {
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
} catch {
    Read-Host | Out-Null
}
