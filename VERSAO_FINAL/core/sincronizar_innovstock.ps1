param(
    [switch]$Simulacao,
    [switch]$ForcarEnvio
)

# ==============================================================================
# INNOVCORE TI - SINCRONIZADOR EM LOTE COM INNOVSTOCK (FIREBASE CLOUD)
# Desenvolvido por Jordan | Integracao Oficial InnovCore TI -> InnovStock
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::InputEncoding = [System.Text.Encoding]::UTF8 } catch {}
try { $Host.UI.RawUI.WindowTitle = "InnovCore TI -> InnovStock Sincronizador de Ativos" } catch {}

$DB_URL = "https://matematica-em-foco-default-rtdb.firebaseio.com"

# Identifica diretorios
$scriptRootDir = if ($env:INNOVCORE_USB_DIR -and (Test-Path $env:INNOVCORE_USB_DIR)) { 
    $env:INNOVCORE_USB_DIR.TrimEnd('\') 
} elseif ($PSScriptRoot) { 
    if ($PSScriptRoot -match '(?i)[\\/]core$') {
        Split-Path $PSScriptRoot -Parent
    } else {
        $PSScriptRoot
    }
} else { 
    (Get-Location).Path 
}

# Localiza base CSV
$caminhoCsv = if (Test-Path (Join-Path $scriptRootDir "dados\desktop\inventario_innovtech.csv")) {
    Join-Path $scriptRootDir "dados\desktop\inventario_innovtech.csv"
} elseif (Test-Path (Join-Path $scriptRootDir "dados\desktop\inventario_innovcore.csv")) {
    Join-Path $scriptRootDir "dados\desktop\inventario_innovcore.csv"
} elseif (Test-Path (Join-Path $scriptRootDir "inventario_innovcore.csv")) {
    Join-Path $scriptRootDir "inventario_innovcore.csv"
} elseif (Test-Path (Join-Path $scriptRootDir "inventario_innovtech.csv")) {
    Join-Path $scriptRootDir "inventario_innovtech.csv"
} else {
    Join-Path $scriptRootDir "dados\desktop\inventario_innovtech.csv"
}

# Funcao para exibir banner
function Exibir-BannerSincronizacao {
    Write-Host " +======================================================================+" -ForegroundColor Cyan
    Write-Host " |  ___ _   _ _   _  _____     ______ _____ ____  _____   _____ ___     |" -ForegroundColor Yellow
    Write-Host " | |_ _| \ | | \ | |/ _ \ \   / / ___/ _ \|  _ \| ____| |_   _|_ _|    |" -ForegroundColor Yellow
    Write-Host " |  | ||  \| |  \| | | | \ \ / / |  | | | | |_) |  _|     | |  | |     |" -ForegroundColor Yellow
    Write-Host " |  | || |\  | |\  | |_| |\ V /| |__| |_| |  _ <| |___    | |  | |     |" -ForegroundColor Yellow
    Write-Host " | |___|_| \_|_| \_|\___/  \_/  \____\___/|_| \_\_____|   |_| |___|    |" -ForegroundColor Yellow
    Write-Host " |                                                                      |" -ForegroundColor Cyan
    Write-Host " |       SINCRONIZACAO DIRETA COM INNOVSTOCK (FIREBASE CLOUD)           |" -ForegroundColor Green
    Write-Host " |                 Desenvolvido por Jordan | Versao 2.0                 |" -ForegroundColor White
    Write-Host " +======================================================================+" -ForegroundColor Cyan
}

function Normalizar-Texto {
    param([string]$texto)
    if ([string]::IsNullOrWhiteSpace($texto)) { return "" }
    $t = $texto.ToLower().Trim()
    $t = $t -replace '[áàãâä]', 'a'
    $t = $t -replace '[éèêë]', 'e'
    $t = $t -replace '[íìîï]', 'i'
    $t = $t -replace '[óòõôö]', 'o'
    $t = $t -replace '[úùûü]', 'u'
    $t = $t -replace '[ç]', 'c'
    $t = $t -replace '\s+', ' '
    return $t
}

function Testar-NomesCompativeis {
    param([string]$nome1, [string]$nome2)
    $n1 = Normalizar-Texto $nome1
    $n2 = Normalizar-Texto $nome2
    if ($n1 -eq $n2) { return $true }
    
    $partes1 = @($n1 -split ' ' | Where-Object { $_.Length -gt 2 })
    $partes2 = @($n2 -split ' ' | Where-Object { $_.Length -gt 2 })
    
    if ($partes1.Count -gt 0 -and $partes2.Count -gt 0) {
        if ($partes1[0] -eq $partes2[0] -and $partes1[-1] -eq $partes2[-1]) {
            return $true
        }
    }
    return $false
}

function Obter-RegistrosCsvLocal {
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
                    if (![string]::IsNullOrWhiteSpace($obj.Patrimonio) -or ![string]::IsNullOrWhiteSpace($obj.Usuario)) {
                        $lista += $obj
                    }
                }
            }
        } catch {
            Write-Host " [!] Erro ao ler base CSV: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    return $lista
}

# Inicio da Execucao
try { Clear-Host } catch {}
Exibir-BannerSincronizacao
Write-Host ""

if (!(Test-Path $caminhoCsv)) {
    Write-Host " [!] Arquivo de banco de dados nao encontrado em:" -ForegroundColor Red
    Write-Host "     $caminhoCsv" -ForegroundColor Yellow
    Write-Host "`n Pressione ENTER para sair..." -ForegroundColor Gray
    Read-Host | Out-Null
    exit
}

$registros = @(Obter-RegistrosCsvLocal -caminho $caminhoCsv)
$totalEquipamentos = $registros.Count

Write-Host "  Base de dados local detectada:" -ForegroundColor Gray
Write-Host "  -> $caminhoCsv" -ForegroundColor Yellow
Write-Host "  -> Total de equipamentos para sincronizar: " -NoNewline -ForegroundColor Gray
Write-Host "$totalEquipamentos equipamento(s)" -ForegroundColor Green

if ($totalEquipamentos -eq 0) {
    Write-Host "`n  [!] Nao ha equipamentos cadastrados para enviar." -ForegroundColor Yellow
    Write-Host "  Cadastre suas maquinas primeiro com o InnovCore TI.`n" -ForegroundColor Gray
    Write-Host " Pressione ENTER para sair..." -ForegroundColor Gray
    Read-Host | Out-Null
    exit
}

# Teste de Conexao com a Nuvem
Write-Host "`n [*] Conectando ao InnovStock Cloud (Firebase)..." -ForegroundColor Cyan
try {
    $employeesRaw = Invoke-RestMethod -Uri "$DB_URL/employees.json" -Method Get -TimeoutSec 12
    $inventoryRaw = Invoke-RestMethod -Uri "$DB_URL/inventory.json" -Method Get -TimeoutSec 12
    Write-Host " [OK] Conexao estabelecida com sucesso!" -ForegroundColor Green
} catch {
    Write-Host " [X] Erro ao conectar com o InnovStock na nuvem:" -ForegroundColor Red
    Write-Host "     $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "     Verifique sua conexao com a internet e tente novamente." -ForegroundColor Gray
    Write-Host "`n Pressione ENTER para sair..." -ForegroundColor Gray
    Read-Host | Out-Null
    exit
}

# Processamento
Write-Host ""
Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
Write-Host " |  Deseja enviar TODOS os $totalEquipamentos equipamento(s) para o InnovStock?       |" -ForegroundColor Yellow
Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
Write-Host "   [S] Sim, Enviar Agora para o InnovStock (Nuvem)" -ForegroundColor Green
Write-Host "   [T] Modo Teste / Simulacao (Ver o que seria enviado)" -ForegroundColor Cyan
Write-Host "   [N] Cancelar" -ForegroundColor Gray
Write-Host ""

$respEnvio = "S"
if (-not $ForcarEnvio) {
    Write-Host " Escolha uma opcao " -NoNewline -ForegroundColor Cyan
    $respEnvio = Read-Host "[S]"
    if ([string]::IsNullOrWhiteSpace($respEnvio)) { $respEnvio = "S" }
}

$modoSimulacao = ($respEnvio.Trim() -match '^(?i)t|teste$') -or $Simulacao
if ($respEnvio.Trim() -match '^(?i)n|nao|cancelar$') {
    Write-Host "`n  Operacao cancelada pelo usuario." -ForegroundColor Yellow
    Start-Sleep -Milliseconds 1000
    exit
}

$textoModo = if ($modoSimulacao) { "[MODO TESTE / SIMULACAO]" } else { "[ENVIO REAL PARA A NUVEM]" }
Write-Host "`n === INICIANDO SINCRONIZACAO $textoModo ===`n" -ForegroundColor Yellow

$nowIso = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
$hojeIso = (Get-Date).ToString("yyyy-MM-dd")

$novosColaboradores = 0
$colaboradoresVinculados = 0
$equipamentosCriados = 0
$equipamentosAtualizados = 0

$itemIndex = 1
foreach ($item in $registros) {
    $pat = if ($item.Patrimonio) { $item.Patrimonio.Trim() } else { "PAT-" + $itemIndex.ToString("D4") }
    $colabNome = if ($item.Usuario) { $item.Usuario.Trim() } else { "Colaborador Nao Identificado" }
    $funcao = if ($item.Funcao) { $item.Funcao.Trim() } else { "Colaborador" }
    $setor = if ($item.Setor) { $item.Setor.Trim() } else { "Operacional" }
    $marca = if ($item.Marca) { $item.Marca.Trim() } else { "Generico" }
    $modelo = if ($item.Modelo) { $item.Modelo.Trim() } else { "Equipamento TI" }
    $serial = if ($item.Numero_Serie) { $item.Numero_Serie.Trim() } else { "N/A" }
    $tipo = if ($item.Tipo_Equipamento) { $item.Tipo_Equipamento.Trim() } else { "Notebook" }
    $hostName = if ($item.Nome_Computador) { $item.Nome_Computador.Trim() } else { "N/A" }
    $ano = if ($item.Ano_Fabricacao) { $item.Ano_Fabricacao.Trim() } else { (Get-Date).Year.ToString() }
    
    # 1. Matching ou Criacao do Colaborador
    $matchedEmpId = $null
    $matchedEmpName = $colabNome
    $matchedEmpRole = $funcao
    
    if ($employeesRaw) {
        foreach ($prop in $employeesRaw.PSObject.Properties) {
            $emp = $prop.Value
            if ($emp -and $emp.name) {
                if (Testar-NomesCompativeis -nome1 $emp.name -nome2 $colabNome) {
                    $matchedEmpId = $prop.Name
                    $matchedEmpName = $emp.name
                    if ($emp.role) { $matchedEmpRole = $emp.role }
                    break
                }
            }
        }
    }
    
    if (-not $matchedEmpId) {
        $newEmpId = "emp_" + ([System.Guid]::NewGuid().ToString("N").Substring(0, 12))
        $empPayload = @{
            id             = $newEmpId
            name           = $colabNome
            role           = $funcao
            cargo          = $funcao
            department     = $setor
            company        = "INNOV SOLUCOES EDUCATIVAS LTDA"
            type           = "interno_innovplay"
            cpf            = ""
            phone          = ""
            email          = ""
            city           = "Sao Miguel dos Campos"
            state          = "AL"
            status         = "active"
            pendingDetails = $true
            notes          = "Cadastrado via InnovCore TI ($pat). Host: $hostName."
            createdAt      = $nowIso
            updatedAt      = $nowIso
            origin         = "innovcore_sync"
        }
        
        if ($modoSimulacao) {
            Write-Host "  [+] [SIMULADO] Criar Colaborador: $colabNome ($funcao - $setor)" -ForegroundColor Cyan
        } else {
            $jsonEmp = $empPayload | ConvertTo-Json -Compress
            Invoke-RestMethod -Uri "$DB_URL/employees/$newEmpId.json" -Method Put -Body ([System.Text.Encoding]::UTF8.GetBytes($jsonEmp)) -ContentType "application/json; charset=utf-8" | Out-Null
            Write-Host "  [OK] Criado Colaborador: $colabNome ($setor)" -ForegroundColor Green
        }
        $matchedEmpId = $newEmpId
        $novosColaboradores++
    } else {
        Write-Host "  [LINK] Colaborador Vinculado: $matchedEmpName" -ForegroundColor DarkGray
        $colaboradoresVinculados++
    }
    
    # 2. Matching ou Criacao do Ativo no Inventario
    $existingEquipId = $null
    if ($inventoryRaw) {
        foreach ($invProp in $inventoryRaw.PSObject.Properties) {
            $eq = $invProp.Value
            if ($eq) {
                if ($eq.sku -eq $pat -or $eq.assetTag -eq $pat -or ($eq.serialNumber -and $eq.serialNumber -eq $serial -and $serial -ne "N/A")) {
                    $existingEquipId = $invProp.Name
                    break
                }
            }
        }
    }
    
    $nomeItem = "$tipo $marca $modelo".Trim()
    $specsResumo = "CPU: $($item.Processador) | RAM: $($item.Memoria_RAM) | Armazenamento: $($item.Armazenamento) | SO: $($item.Sistema_Operacional) | IP: $($item.IP_Rede)"
    $displayId = "{0:D3}-{1}" -f $itemIndex, ([System.Guid]::NewGuid().ToString("N").Substring(0, 8))
    
    $equipPayload = @{
        name                 = $nomeItem
        type                 = if ($tipo -match '(?i)notebook|laptop') { "Notebook" } elseif ($tipo -match '(?i)servidor') { "Servidor" } else { "Desktop" }
        equipmentType        = $tipo
        brand                = $marca
        model                = $modelo
        sku                  = $pat
        assetTag             = $pat
        serialNumber         = $serial
        serials              = @(
            @{ type = "Nº de Série"; value = $serial },
            @{ type = "Patrimônio"; value = $pat }
        )
        condition            = "otimo"
        qty                  = 1
        qtyDamaged           = 0
        damagedNotes         = ""
        isCountable          = $false
        isCompanyAsset       = $true
        ownershipType        = "uso_empresa_colaborador"
        category             = "Equipamentos de Colaboradores"
        assignedEmployeeId   = $matchedEmpId
        assignedEmployeeName = $matchedEmpName
        assignedEmployeeRole = $matchedEmpRole
        assignedDate         = $hojeIso
        location             = "Em posse do colaborador: $matchedEmpName"
        accessories          = "Host: $hostName | Ano: $ano | $specsResumo"
        supplier             = $marca
        price                = 0
        purchaseDate         = if ($ano) { "$ano-01-01" } else { "" }
        warrantyDate         = ""
        photo                = ""
        status               = @("disponivel")
        displayId            = $displayId
        lastModified         = $nowIso
        modifiedBy           = "InnovCore TI (Sincronizacao em Lote)"
        origin               = "innovcore_sync"
    }
    
    if ($existingEquipId) {
        if ($modoSimulacao) {
            Write-Host "  [*] [SIMULADO] Atualizar Ativo: $pat - $nomeItem ($matchedEmpName)" -ForegroundColor Yellow
        } else {
            $jsonEquip = $equipPayload | ConvertTo-Json -Compress
            Invoke-RestMethod -Uri "$DB_URL/inventory/$existingEquipId.json" -Method Patch -Body ([System.Text.Encoding]::UTF8.GetBytes($jsonEquip)) -ContentType "application/json; charset=utf-8" | Out-Null
            Write-Host "  [ATUALIZADO] Ativo: $pat - $nomeItem -> $matchedEmpName" -ForegroundColor Cyan
        }
        $equipamentosAtualizados++
    } else {
        $equipPayload["createdAt"] = $nowIso
        if ($modoSimulacao) {
            Write-Host "  [+] [SIMULADO] Novo Ativo: $pat - $nomeItem ($matchedEmpName)" -ForegroundColor Green
        } else {
            $jsonEquip = $equipPayload | ConvertTo-Json -Compress
            Invoke-RestMethod -Uri "$DB_URL/inventory.json" -Method Post -Body ([System.Text.Encoding]::UTF8.GetBytes($jsonEquip)) -ContentType "application/json; charset=utf-8" | Out-Null
            Write-Host "  [CRIADO] Novo Ativo: $pat - $nomeItem -> $matchedEmpName" -ForegroundColor Green
        }
        $equipamentosCriados++
    }
    
    $itemIndex++
}

# Resumo Final
Write-Host ""
Write-Host " +======================================================================+" -ForegroundColor Cyan
Write-Host " |               RESUMO DA SINCRONIZACAO COM INNOVSTOCK                 |" -ForegroundColor Yellow
Write-Host " +======================================================================+" -ForegroundColor Cyan
Write-Host "  - Total de Equipamentos Processados:  " -NoNewline -ForegroundColor Gray; Write-Host "$totalEquipamentos" -ForegroundColor Green
Write-Host "  - Novos Equipamentos Cadastrados:    " -NoNewline -ForegroundColor Gray; Write-Host "$equipamentosCriados" -ForegroundColor Green
Write-Host "  - Equipamentos Atualizados:          " -NoNewline -ForegroundColor Gray; Write-Host "$equipamentosAtualizados" -ForegroundColor Cyan
Write-Host "  - Novos Colaboradores Criados:       " -NoNewline -ForegroundColor Gray; Write-Host "$novosColaboradores" -ForegroundColor Yellow
Write-Host "  - Colaboradores Vinculados:          " -NoNewline -ForegroundColor Gray; Write-Host "$colaboradoresVinculados" -ForegroundColor White
Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan

if ($modoSimulacao) {
    Write-Host "`n  [*] Modo de Simulacao finalizado. Nenhum dado real foi alterado." -ForegroundColor Yellow
    Write-Host "  Para enviar de verdade, execute sem o parametro de simulacao.`n" -ForegroundColor Gray
} else {
    Write-Host "`n  [SUCESSO] Todos os ativos estao disponiveis no InnovStock Web!" -ForegroundColor Green
    Write-Host "  Voce ja pode abrir o InnovStock e gerar os Termos de Responsabilidade.`n" -ForegroundColor White
}

Write-Host " Pressione ENTER para voltar ao menu..." -ForegroundColor Gray
Read-Host | Out-Null
