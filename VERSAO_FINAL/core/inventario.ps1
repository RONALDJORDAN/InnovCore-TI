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

# ------------------------------------------------------------------------------
# MODULO 1 - REGISTRO DE ATIVOS E PATRIMONIO (BENS GERAIS: MOVEIS, EQUIPAMENTOS, INSUMOS)
# ------------------------------------------------------------------------------
# Base de dados EXCLUSIVA dos bens patrimoniais (separada do inventario TI de desktop)
$pastaDadosAtivos = Join-Path $scriptRootDir "dados\ativos"
$caminhoCsvAtivos = Join-Path $pastaDadosAtivos "inventario_ativos_innovcore.csv"

# Catalogo EDITAVEL de tipos de bens (Categoria;Tipo;Nome do Bem;Setor sugerido)
# O usuario pode editar este arquivo direto no Excel para incluir novos tipos.
$caminhoCatalogoAtivos = Join-Path $scriptRootDir "core\catalogos\catalogo_ativos.csv"
if (!(Test-Path $caminhoCatalogoAtivos)) {
    $caminhoCatalogoAtivosAlt = Join-Path $scriptRootDir "catalogos\catalogo_ativos.csv"
    if (Test-Path $caminhoCatalogoAtivosAlt) {
        $caminhoCatalogoAtivos = $caminhoCatalogoAtivosAlt
    }
}

# Cadastros MESTRES (editaveis no Excel e pelo proprio sistema)
$caminhoCatalogoColaboradores = Join-Path $scriptRootDir "core\catalogos\catalogo_colaboradores.csv"
$caminhoCatalogoFornecedores  = Join-Path $scriptRootDir "core\catalogos\catalogo_fornecedores.csv"

# Termos de responsabilidade e modelo de importacao em lote
$pastaTermos             = Join-Path $pastaDadosAtivos "termos"
$caminhoTermosEmitidos   = Join-Path $pastaDadosAtivos "termos_emitidos.csv"
$caminhoModeloImportacao = Join-Path $pastaDadosAtivos "modelo_importacao_ativos.csv"

# Setores oficiais para alocacao de bens (patrimonio)
$script:setoresAtivosPadrao = @(
    "Operacoes e Logistica", "Pedagogico", "TI", "Administrativo", "Financeiro",
    "Comercial", "Marketing", "Recursos Humanos", "Motoristas",
    "Servicos Gerais e Limpeza", "Manutencao", "Diretoria", "Almoxarifado",
    "Recepcao", "Convivencia"
)

# Empresas / Unidades cadastraveis no modulo de ativos
$script:empresasAtivosPadrao = @("InnovEdu", "InnovTech")

# Situacoes patrimoniais padrao
$script:situacoesAtivosPadrao = @(
    "Ativo/Operacional", "Em Manutencao", "Em Estoque",
    "Baixado/Inativo", "Emprestado/Cedido"
)

# Definicao de Cabecalho Completo Oficial InnovCore TI (Compativel com Excel PT-BR)
# Estrutura PATRIMONIAL COMPLETA para Gestao de Ativos Operacionais
$cabecalhoOficial = @(
    "Data_Registro",              # Data/hora do registro no sistema
    "Empresa_Unidade",            # Empresa/Unidade (ex: InnovEdu, InnovTech)
    "Codigo_Tombamento",          # PAT-XXXXX (Imutavel, sequencial global)
    "Numero_Nota_Fiscal",         # NF-XXXXX (Opcional)
    "CRM_Serial",                 # Serial/Chassi/RFID/IMEI (Opcional)
    "Categoria_Bem",              # Categoria principal (Tecnologia, Moveis, Eletrodomesticos, etc.)
    "Tipo_Especifico",            # Subtipo livre (ex: Cadeira Presidente, Mesa L, Notebook)
    "Nome_Bem_Produto",           # Nome completo do bem (ex: MacBook Pro M2 Max 16'', Ar Condicionado 12000 BTU)
    "Setor_Atribuido",            # Setor onde o bem esta alocado
    "Responsavel_Colaborador",    # Nome do responsavel/custodiante
    "Fornecedor",                 # Fornecedor da aquisicao (Opcional)
    "Situacao",                   # Ativo/Operacional, Em Manutencao, Em Estoque, Baixado/Inativo, Emprestado/Cedido
    "Data_Compra",                # DD/MM/AAAA (Obrigatorio para patrimônio)
    "Valor_Compra",               # R$ 0,00 (Opcional)
    "Localizacao_Fisica_Posto",   # Ex: Sala TI - Posto 03, Almoxarifado, Veiculo Placa ABC1234
    "Observacoes_Adicionais",     # Livre: garantia, condicoes especiais, etc.
    # Campos legados TI (mantidos para compatibilidade retroativa)
    "Nome_Computador",
    "Processador",
    "Memoria_RAM",
    "Armazenamento",
    "Placa_Video",
    "Sistema_Operacional",
    "IP_Rede",
    "MAC_Rede",
    "Bateria",
    "Android_Versao"
)
$cabecalhoCsvString = $cabecalhoOficial -join ';'

$utf8ComBOM = New-Object System.Text.UTF8Encoding($true)
$utf8SemBOM = New-Object System.Text.UTF8Encoding($false)

# ==============================================================================
# MOTOR DE LAYOUT ADAPTATIVO (RESOLUCAO REAL DO TERMINAL)
# ==============================================================================
# O sistema se ajusta ao terminal em que for aberto: detecta largura e altura,
# centraliza as caixas, expande as tabelas e compacta o banner quando o
# terminal e baixo (ex.: 80x24).
$script:terminalLargura = 100
$script:terminalAltura  = 30
$script:larguraCaixa    = 73
$script:recuoTela       = ""
$script:colunaLinha     = 0

function Atualizar-TamanhoTerminal {
    $largura = 0
    $altura = 0
    try { if ($env:COLUMNS) { [void][int]::TryParse($env:COLUMNS, [ref]$largura) } } catch {}
    try { if ($env:LINES) { [void][int]::TryParse($env:LINES, [ref]$altura) } } catch {}
    if ($largura -le 0) { try { $largura = $Host.UI.RawUI.WindowSize.Width } catch {} }
    if ($altura -le 0) { try { $altura = $Host.UI.RawUI.WindowSize.Height } catch {} }
    if ($largura -le 0) { $largura = 100 }
    if ($altura -le 0) { $altura = 30 }
    if ($largura -lt 40) { $largura = 40 }
    if ($largura -gt 220) { $largura = 220 }
    if ($altura -lt 8) { $altura = 8 }

    $script:terminalLargura = $largura
    $script:terminalAltura = $altura

    # Largura da caixa: ocupa o terminal com folga, sem nunca ficar menor que o padrao
    $caixa = $largura - 2
    if ($caixa -lt 73) { $caixa = 73 }
    if ($caixa -gt 120) { $caixa = 120 }
    $script:larguraCaixa = $caixa

    # Centraliza horizontalmente quando sobra espaco
    $sobra = $largura - $caixa
    if ($sobra -ge 4) { $script:recuoTela = " " * [int]($sobra / 2) } else { $script:recuoTela = "" }
}

# Quantos itens cabem em uma listagem sem estourar a altura do terminal
function Get-LimiteListagem {
    $limite = $script:terminalAltura - 14
    if ($limite -lt 8) { $limite = 8 }
    if ($limite -gt 500) { $limite = 500 }
    return $limite
}

# ------------------------------------------------------------------------------
# Write-Host ADAPTATIVO
# Centraliza a saida, redesenha os separadores de caixa e alinha a borda direita
# de todas as linhas "|" conforme a largura detectada do terminal.
# ------------------------------------------------------------------------------
function Write-Host {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true)]
        [object]$Object,
        [switch]$NoNewline,
        [string]$ForegroundColor
    )
    process {
        if ($null -eq $Object) {
            $texto = ""
        } elseif ($Object -is [array]) {
            $texto = ($Object -join " ")
        } else {
            $texto = [string]$Object
        }

        $partes = $texto -split "`r?`n"
        $saida = New-Object System.Collections.Generic.List[string]

        for ($i = 0; $i -lt $partes.Count; $i++) {
            $corpo = $partes[$i]
            $prefixo = ""
            if ($i -eq 0) {
                if ($script:colunaLinha -eq 0 -and $script:recuoTela -ne "") {
                    $prefixo = $script:recuoTela
                    $script:colunaLinha = $script:recuoTela.Length
                }
            } else {
                $script:colunaLinha = 0
                if ($script:recuoTela -ne "") {
                    $prefixo = $script:recuoTela
                    $script:colunaLinha = $script:recuoTela.Length
                }
            }

            # Separadores de caixa (+=====+ / +-----+) sao redesenhados na largura atual
            if ($corpo -match '^\s*\+([=\-])\1+\s*\+$') {
                $fill = $Matches[1]
                $corpo = " +" + ($fill * ($script:larguraCaixa - 3)) + "+"
            }

            # Borda direita das linhas de caixa: preenche ate a largura da caixa
            # (somente quando o trecho TERMINA em "|", sem espacos depois)
            if ($corpo.EndsWith("|")) {
                $usado = $script:colunaLinha + $corpo.Length
                $faltam = ($script:larguraCaixa - 1) - $usado
                if ($faltam -gt 0) {
                    $corpo = $corpo.Substring(0, $corpo.Length - 1) + (" " * $faltam) + "|"
                }
            }

            $saida.Add($prefixo + $corpo)
            $script:colunaLinha += $corpo.Length
        }

        $final = $saida -join "`n"
        if ($NoNewline) {
            if ($ForegroundColor -and $ForegroundColor -ne "") {
                Microsoft.PowerShell.Utility\Write-Host -Object $final -NoNewline -ForegroundColor $ForegroundColor
            } else {
                Microsoft.PowerShell.Utility\Write-Host -Object $final -NoNewline
            }
        } else {
            if ($ForegroundColor -and $ForegroundColor -ne "") {
                Microsoft.PowerShell.Utility\Write-Host -Object $final -ForegroundColor $ForegroundColor
            } else {
                Microsoft.PowerShell.Utility\Write-Host -Object $final
            }
            $script:colunaLinha = 0
        }
    }
}

# ------------------------------------------------------------------------------
# LIMPEZA DE TELA (sem historico)
# Limpa a area visivel E o buffer de rolagem (scrollback), para que o terminal
# apresente sempre apenas a interface atual, sem "restos" de telas anteriores.
# ------------------------------------------------------------------------------
function Limpar-Console {
    # 1) Limpa a area visivel (funciona em qualquer console)
    try {
        Microsoft.PowerShell.Utility\Clear-Host
    } catch {
        # 2) Fallback: sequencia ANSI (Windows Terminal e consoles modernos)
        try {
            $esc = [char]27
            [Console]::Write("$esc[2J$esc[3J$esc[H")
        } catch {}
        return
    }

    # 3) Limpa tambem o historico de rolagem (scrollback) quando o console suporta VT.
    #    No PowerShell 5.1 o escape de ESC e [char]27 (o "`e" so existe no 6+).
    try {
        if ($env:WT_SESSION -or $env:TERM_PROGRAM -or $env:ConEmuANSI -eq "ON") {
            $esc = [char]27
            [Console]::Write("$esc[2J$esc[3J$esc[H")
        }
    } catch {}
}

# ------------------------------------------------------------------------------
# Clear-Host ADAPTATIVO: zera o contador de coluna e limpa a tela sem historico
# ------------------------------------------------------------------------------
function Clear-Host {
    $script:colunaLinha = 0
    Limpar-Console
}

# ------------------------------------------------------------------------------
# Read-Host ADAPTATIVO (TECLA ESC / CANCELAMENTO)
# Durante o preenchimento de qualquer campo, a tecla ESC pergunta se o usuario
# deseja cancelar a operacao atual. Confirmando, a operacao e abortada com
# seguranca (nada e gravado pela metade) e o sistema volta ao menu anterior.
# ------------------------------------------------------------------------------
function Testar-CancelamentoOperacao {
    Write-Host ""
    Write-Host " [!] ESC pressionado. Deseja CANCELAR a operacao atual?" -ForegroundColor Yellow
    $resp = Ler-OpcaoOuEsc -prompt "   Confirmar cancelamento? (S/N)" -valorPadrao "N"
    return ($resp.Trim() -match '(?i)^s|sim|y|yes$')
}

function Read-Host {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true)]
        [object]$Prompt,
        [switch]$AsSecureString
    )
    process {
        $textoPrompt = if ($null -eq $Prompt) { "" } else { [string]$Prompt }

        if ($AsSecureString) {
            return Microsoft.PowerShell.Utility\Read-Host -Prompt $textoPrompt -AsSecureString
        }

        # Console real: leitura tecla a tecla para detectar o ESC
        $podeLerTeclas = $false
        try { if (-not [Console]::IsInputRedirected) { $podeLerTeclas = $true } } catch {}

        if ($podeLerTeclas) {
            Write-Host $textoPrompt -NoNewline
            $buffer = ""
            while ($true) {
                $key = $null
                try { $key = [Console]::ReadKey($true) } catch { $podeLerTeclas = $false; break }

                if ($key.Key -eq [System.ConsoleKey]::Escape) {
                    Write-Host ""
                    if (Testar-CancelamentoOperacao) {
                        throw [System.OperationCanceledException]::new("Operacao cancelada pelo usuario")
                    }
                    Write-Host $textoPrompt -NoNewline
                    continue
                }
                if ($key.Key -eq [System.ConsoleKey]::Enter) {
                    Write-Host ""
                    return $buffer.Trim()
                }
                if ($key.Key -eq [System.ConsoleKey]::Backspace) {
                    if ($buffer.Length -gt 0) {
                        $buffer = $buffer.Substring(0, $buffer.Length - 1)
                        [Console]::Write("`b `b")
                    }
                    continue
                }
                if ($key.KeyChar -and -not [char]::IsControl($key.KeyChar)) {
                    $buffer += $key.KeyChar
                    [Console]::Write($key.KeyChar)
                }
            }
        }

        # Fallback: entrada redirecionada ou console sem suporte a ReadKey
        # Le direto do Console.In para que o EOF seja detectado ($null) sem
        # consumir a linha de forma inconsistente.
        try {
            Write-Host $textoPrompt -NoNewline
            $resposta = [Console]::In.ReadLine()
            Write-Host ""
            if ($null -eq $resposta) { return $null }
            return $resposta
        } catch {
            if ($textoPrompt -eq "") { return Microsoft.PowerShell.Utility\Read-Host }
            return Microsoft.PowerShell.Utility\Read-Host -Prompt $textoPrompt
        }
    }
}

# ==============================================================================
# IDENTIDADE VISUAL E BANNER CORPORATIVO (ADAPTATIVO A ALTURA DO TERMINAL)
# ==============================================================================
function Exibir-LogoPrincipal {
    # Terminal baixo: banner compacto para nao empurrar o menu para fora da tela
    if ($script:terminalAltura -lt 28) {
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host " |        INNOVCORE TI - GESTAO DE ATIVOS E INVENTARIO  |  v2.0       |" -ForegroundColor Green
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        return
    }
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
        # Entrada redirecionada (execucao automatizada): nao usar ReadKey, que ficaria bloqueando
        if ([Console]::IsInputRedirected) {
            Read-Host | Out-Null
            return
        }
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
            # Entrada encerrada (EOF): encerra o fluxo em vez de repetir o menu em laco
            if ($null -eq $val) { return "0" }
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
    param(
        [string]$caminho,
        [string]$empresaPadrao = "InnovTech",
        [string]$setorPadrao = "TI"
    )
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
                    # Garante compatibilidade com todos os campos do novo cabecalho
                    foreach ($propOficial in $script:cabecalhoOficial) {
                        if (-not ($obj.PSObject.Properties[$propOficial])) {
                            $obj | Add-Member -MemberType NoteProperty -Name $propOficial -Value "N/A" -Force
                        }
                    }
                    # MIGRACAO RETROATIVA: Mapear campos antigos (legado TI) para novos campos patrimoniais
                    # Empresa
                    if ($obj.Empresa -and $obj.Empresa -ne "N/A" -and (-not $obj.Empresa_Unidade -or $obj.Empresa_Unidade -eq "N/A")) {
                        $obj.Empresa_Unidade = $obj.Empresa
                    }
                    # Codigo_Tombamento (Patrimonio)
                    if ($obj.Patrimonio -and $obj.Patrimonio -ne "N/A" -and (-not $obj.Codigo_Tombamento -or $obj.Codigo_Tombamento -eq "N/A")) {
                        $obj.Codigo_Tombamento = $obj.Patrimonio
                    }
                    # CRM_Serial (Numero_Serie)
                    if ($obj.Numero_Serie -and $obj.Numero_Serie -ne "N/A" -and (-not $obj.CRM_Serial -or $obj.CRM_Serial -eq "N/A")) {
                        $obj.CRM_Serial = $obj.Numero_Serie
                    }
                    # Categoria_Bem (Tipo_Equipamento -> Tecnologia)
                    if ($obj.Tipo_Equipamento -and $obj.Tipo_Equipamento -ne "N/A" -and (-not $obj.Categoria_Bem -or $obj.Categoria_Bem -eq "N/A")) {
                        $obj.Categoria_Bem = "Tecnologia"
                    }
                    # Tipo_Especifico (Tipo_Equipamento ou Tipo_Ativo)
                    if ($obj.Tipo_Ativo -and $obj.Tipo_Ativo -ne "N/A" -and (-not $obj.Tipo_Especifico -or $obj.Tipo_Especifico -eq "N/A")) {
                        $obj.Tipo_Especifico = $obj.Tipo_Ativo
                    } elseif ($obj.Tipo_Equipamento -and $obj.Tipo_Equipamento -ne "N/A" -and (-not $obj.Tipo_Especifico -or $obj.Tipo_Especifico -eq "N/A")) {
                        $obj.Tipo_Especifico = $obj.Tipo_Equipamento
                    }
                    # Nome_Bem_Produto (Marca + Modelo)
                    if (($obj.Marca -and $obj.Marca -ne "N/A") -or ($obj.Modelo -and $obj.Modelo -ne "N/A")) {
                        if (-not $obj.Nome_Bem_Produto -or $obj.Nome_Bem_Produto -eq "N/A") {
                            $marca = if ($obj.Marca -and $obj.Marca -ne "N/A") { $obj.Marca } else { "" }
                            $modelo = if ($obj.Modelo -and $obj.Modelo -ne "N/A") { $obj.Modelo } else { "" }
                            $obj.Nome_Bem_Produto = ($marca + " " + $modelo).Trim()
                        }
                    }
                    # Setor_Atribuido (Setor)
                    if ($obj.Setor -and $obj.Setor -ne "N/A" -and (-not $obj.Setor_Atribuido -or $obj.Setor_Atribuido -eq "N/A")) {
                        $obj.Setor_Atribuido = $obj.Setor
                    }
                    # Responsavel_Colaborador (Usuario)
                    if ($obj.Usuario -and $obj.Usuario -ne "N/A" -and (-not $obj.Responsavel_Colaborador -or $obj.Responsavel_Colaborador -eq "N/A")) {
                        $obj.Responsavel_Colaborador = $obj.Usuario
                    }
                    # Situacao (Status_Ativo)
                    if ($obj.Status_Ativo -and $obj.Status_Ativo -ne "N/A" -and (-not $obj.Situacao -or $obj.Situacao -eq "N/A")) {
                        $obj.Situacao = $obj.Status_Ativo
                    }
                    # Data_Compra (Ano_Fabricacao -> 01/01/ANO)
                    if ($obj.Ano_Fabricacao -and $obj.Ano_Fabricacao -ne "N/A" -and (-not $obj.Data_Compra -or $obj.Data_Compra -eq "N/A")) {
                        if ($obj.Ano_Fabricacao -match '^\d{4}$') {
                            $obj.Data_Compra = "01/01/" + $obj.Ano_Fabricacao
                        } else {
                            $obj.Data_Compra = $obj.Ano_Fabricacao
                        }
                    }
                    # Valor_Compra (Valor_Aquisicao)
                    if ($obj.Valor_Aquisicao -and $obj.Valor_Aquisicao -ne "N/A" -and (-not $obj.Valor_Compra -or $obj.Valor_Compra -eq "N/A")) {
                        $obj.Valor_Compra = $obj.Valor_Aquisicao
                    }
                    # Localizacao_Fisica_Posto (Localizacao ou Setor)
                    if ($obj.Localizacao -and $obj.Localizacao -ne "N/A" -and (-not $obj.Localizacao_Fisica_Posto -or $obj.Localizacao_Fisica_Posto -eq "N/A")) {
                        $obj.Localizacao_Fisica_Posto = $obj.Localizacao
                    } elseif ($obj.Setor -and $obj.Setor -ne "N/A" -and (-not $obj.Localizacao_Fisica_Posto -or $obj.Localizacao_Fisica_Posto -eq "N/A")) {
                        $obj.Localizacao_Fisica_Posto = $obj.Setor
                    }
                    # Observacoes_Adicionais (Observacoes)
                    if ($obj.Observacoes -and $obj.Observacoes -ne "N/A" -and (-not $obj.Observacoes_Adicionais -or $obj.Observacoes_Adicionais -eq "N/A")) {
                        $obj.Observacoes_Adicionais = $obj.Observacoes
                    }
                    # Defaults para campos novos obrigatorios
                    if (-not $obj.Empresa_Unidade -or $obj.Empresa_Unidade -eq "N/A") { $obj.Empresa_Unidade = $empresaPadrao }
                    if (-not $obj.Categoria_Bem -or $obj.Categoria_Bem -eq "N/A") { $obj.Categoria_Bem = "Tecnologia" }
                    if (-not $obj.Situacao -or $obj.Situacao -eq "N/A") { $obj.Situacao = "Ativo/Operacional" }
                    if (-not $obj.Setor_Atribuido -or $obj.Setor_Atribuido -eq "N/A") { $obj.Setor_Atribuido = $setorPadrao }
                    if (-not $obj.Responsavel_Colaborador -or $obj.Responsavel_Colaborador -eq "N/A") { $obj.Responsavel_Colaborador = $obj.Usuario }
                    if (-not $obj.Localizacao_Fisica_Posto -or $obj.Localizacao_Fisica_Posto -eq "N/A") { $obj.Localizacao_Fisica_Posto = $obj.Setor_Atribuido }

                    # ALIAS DE COMPATIBILIDADE (patrimonial -> legado)
                    # Cria os campos antigos quando a base usa apenas o cabecalho patrimonial novo,
                    # para que as telas antigas (consulta, dashboard, exclusao) continuem funcionando.
                    $anoCompra = if ($obj.Data_Compra -and $obj.Data_Compra -match '(\d{4})') { $Matches[1] } else { "" }
                    $aliases = [ordered]@{
                        "Empresa"          = $obj.Empresa_Unidade
                        "Patrimonio"       = $obj.Codigo_Tombamento
                        "Usuario"          = $obj.Responsavel_Colaborador
                        "Setor"            = $obj.Setor_Atribuido
                        "Tipo_Ativo"       = $obj.Tipo_Especifico
                        "Tipo_Equipamento" = $obj.Tipo_Especifico
                        "Categoria_Ativo"  = $obj.Categoria_Bem
                        "Numero_Serie"     = $obj.CRM_Serial
                        "Valor_Aquisicao"  = $obj.Valor_Compra
                        "Status_Ativo"     = $obj.Situacao
                        "Localizacao"      = $obj.Localizacao_Fisica_Posto
                        "Observacoes"      = $obj.Observacoes_Adicionais
                        "Ano_Aquisicao"    = $anoCompra
                        "Ano_Fabricacao"   = $anoCompra
                        "Funcao"           = "N/A"
                        "Marca"            = "N/A"
                        "Modelo"           = "N/A"
                    }
                    foreach ($chave in $aliases.Keys) {
                        if (-not $obj.PSObject.Properties[$chave]) {
                            $obj | Add-Member -MemberType NoteProperty -Name $chave -Value $aliases[$chave] -Force
                        }
                    }
                    $lista += $obj
                }
            }
        } catch {
            Write-Host " [!] Aviso ao ler CSV: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    # Retorna sempre como ARRAY (evita que um unico registro perca o .Count no PowerShell 5.1)
    return ,$lista
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
    param(
        [string]$caminho,
        [array]$registros,
        [switch]$BaseAtivos
    )
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
            (& $limparCampo $r.Empresa_Unidade),
            (& $limparCampo $r.Codigo_Tombamento),
            (& $limparCampo $r.Numero_Nota_Fiscal),
            (& $limparCampo $r.CRM_Serial),
            (& $limparCampo $r.Categoria_Bem),
            (& $limparCampo $r.Tipo_Especifico),
            (& $limparCampo $r.Nome_Bem_Produto),
            (& $limparCampo $r.Setor_Atribuido),
            (& $limparCampo $r.Responsavel_Colaborador),
            (& $limparCampo $r.Fornecedor),
            (& $limparCampo $r.Situacao),
            (& $limparCampo $r.Data_Compra),
            (& $limparCampo $r.Valor_Compra),
            (& $limparCampo $r.Localizacao_Fisica_Posto),
            (& $limparCampo $r.Observacoes_Adicionais),
            (& $limparCampo $r.Nome_Computador),
            (& $limparCampo $r.Processador),
            (& $limparCampo $r.Memoria_RAM),
            (& $limparCampo $r.Armazenamento),
            (& $limparCampo $r.Placa_Video),
            (& $limparCampo $r.Sistema_Operacional),
            (& $limparCampo $r.IP_Rede),
            (& $limparCampo $r.MAC_Rede),
            (& $limparCampo $r.Bateria),
            (& $limparCampo $r.Android_Versao)
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
            if ($BaseAtivos) {
                $script:caminhoCsvAtivos = $caminhoAlvo
            } else {
                $script:caminhoCsv = $caminhoAlvo
            }
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
                    $nomeResgate = if ($BaseAtivos) { "inventario_ativos_innovcore" } else { "inventario_innovcore" }
                    $caminhoAlvo = Join-Path $scriptRootDir "$nomeResgate`_$timestampNovo.csv"
                    Write-Host "   -> Tentando gravar em: $caminhoAlvo" -ForegroundColor Cyan
                }
                "E" {
                    $nomeResgate = if ($BaseAtivos) { "inventario_ativos_innovcore" } else { "inventario_innovcore" }
                    $caminhoAlvo = Join-Path $env:USERPROFILE "Desktop\$nomeResgate`_emergencia.csv"
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
# CATALOGO EDITAVEL DE TIPOS DE BENS (MODULO 1 - REGISTRO DE ATIVOS)
# ==============================================================================
# Formato do arquivo: Categoria;Tipo_Especifico;Nome_Bem_Produto;Setor_Sugerido
# O arquivo e livre para o usuario editar no Excel e adicionar novos tipos.
function Obter-CatalogoAtivos {
    $lista = @()
    $caminho = $script:caminhoCatalogoAtivos

    if (!(Test-Path $caminho)) {
        Write-Host " [!] Catalogo de tipos nao encontrado: $caminho" -ForegroundColor Yellow
        Write-Host "     Os tipos serao digitados livremente. Reinstale o pacote para restaurar o catalogo." -ForegroundColor Gray
        return ,$lista
    }

    try {
        $raw = [System.IO.File]::ReadAllText($caminho, [System.Text.Encoding]::UTF8)
        $linhas = $raw -split "\r?\n" | Where-Object { ![string]::IsNullOrWhiteSpace($_) }
        for ($i = 1; $i -lt $linhas.Count; $i++) {
            $v = $linhas[$i] -split ';'
            $categoria = if ($v.Count -gt 0) { $v[0].Trim() } else { "" }
            $tipo = if ($v.Count -gt 1) { $v[1].Trim() } else { "" }
            if ($tipo -eq "") { continue }
            $nomeBem = if ($v.Count -gt 2) { $v[2].Trim() } else { "" }
            $setorSugerido = if ($v.Count -gt 3) { $v[3].Trim() } else { "" }
            $lista += [PSCustomObject]@{
                Categoria        = $categoria
                Tipo_Especifico  = $tipo
                Nome_Bem_Produto = $(if ($nomeBem -eq "") { $tipo } else { $nomeBem })
                Setor_Sugerido   = $setorSugerido
            }
        }
    } catch {
        Write-Host " [!] Erro ao ler o catalogo de tipos: $($_.Exception.Message)" -ForegroundColor Red
    }

    return ,$lista
}

# Devolve a lista de categorias unicas do catalogo (ordem de aparicao no arquivo)
function Obter-CategoriasCatalogo {
    param([array]$catalogo)
    $lista = @()
    if (!$catalogo -or $catalogo.Count -eq 0) { return ,$lista }
    foreach ($item in $catalogo) {
        if ($item.Categoria -and ($lista -notcontains $item.Categoria)) {
            $lista += $item.Categoria
        }
    }
    return ,$lista
}

# Devolve os tipos de uma categoria (sem repeticao)
function Obter-TiposPorCategoria {
    param([array]$catalogo, [string]$categoria)
    $lista = @()
    if (!$catalogo -or $catalogo.Count -eq 0) { return ,$lista }
    foreach ($item in $catalogo) {
        if ($item.Categoria -eq $categoria -and $item.Tipo_Especifico) {
            if ($lista -notcontains $item.Tipo_Especifico) {
                $lista += $item.Tipo_Especifico
            }
        }
    }
    return ,$lista
}

# ==============================================================================
# AUDITORIA DE DUPLICIDADE PARA BENS PATRIMONIAIS
# ==============================================================================
# Observacao: a Nota Fiscal NAO e usada como criterio de duplicidade porque
# uma unica nota fiscal normalmente cobre varios bens do mesmo lote.
function Testar-DuplicidadeBem {
    param(
        [array]$registros,
        [string]$codigoTombamento,
        [string]$crmSerial
    )

    if (!$registros -or $registros.Count -eq 0) { return $null }

    foreach ($r in $registros) {
        $motivo = $null

        # 1. Codigo de Tombamento (PAT-XXXX) - identificador imutavel do bem
        $tombamentoExistente = if ($r.Codigo_Tombamento -and $r.Codigo_Tombamento -ne "N/A") { $r.Codigo_Tombamento } else { "" }
        if (![string]::IsNullOrWhiteSpace($codigoTombamento) -and ![string]::IsNullOrWhiteSpace($tombamentoExistente)) {
            if ($tombamentoExistente.Trim().ToUpper() -eq $codigoTombamento.Trim().ToUpper()) {
                $motivo = "Codigo de Tombamento repetido ($codigoTombamento)"
            }
        }

        # 2. CRM / Serial - identificacao fisica unica do bem
        $serialExistente = if ($r.CRM_Serial -and $r.CRM_Serial -ne "N/A") { $r.CRM_Serial } else { "" }
        if (-not $motivo -and ![string]::IsNullOrWhiteSpace($crmSerial) -and $crmSerial.Trim() -notmatch '(?i)^n/a|none|default|0000$') {
            if (![string]::IsNullOrWhiteSpace($serialExistente) -and $serialExistente.Trim().ToUpper() -eq $crmSerial.Trim().ToUpper()) {
                $motivo = "CRM/Serial identico ($crmSerial)"
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
# HELPERS DE COMPATIBILIDADE ENTRE O MODELO PATRIMONIAL E O MODELO LEGADO (TI)
# ==============================================================================
# Devolve o setor do registro, priorizando o campo patrimonial novo
function Obter-SetorDoRegistro {
    param($Registro)
    if ($Registro.Setor_Atribuido -and $Registro.Setor_Atribuido -ne "N/A") { return $Registro.Setor_Atribuido }
    if ($Registro.Setor -and $Registro.Setor -ne "N/A") { return $Registro.Setor }
    return ""
}

# ==============================================================================
# CALCULO INTELIGENTE DE CODIGO DE TOMBAMENTO SEQUENCIAL (PAT-XXXX)
# ==============================================================================
function Obter-ProximoPatrimonio {
    param([array]$registros)
    if (!$registros -or $registros.Count -eq 0) {
        return "PAT-0001"
    }
    
    $maiorNumero = 0
    foreach ($r in $registros) {
        # Verifica tanto Codigo_Tombamento (novo) quanto Patrimonio (legado)
        $campoVerificar = if ($r.Codigo_Tombamento -and $r.Codigo_Tombamento -ne "N/A") { $r.Codigo_Tombamento } elseif ($r.Patrimonio -and $r.Patrimonio -ne "N/A") { $r.Patrimonio } else { "" }
        if ($campoVerificar -match 'PAT-(\d+)') {
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
        [string]$codigoTombamento,
        [string]$crmSerial,
        [string]$numeroNotaFiscal,
        [string]$macRede,
        [string]$nomeComputador
    )
    
    if (!$registros -or $registros.Count -eq 0) { return $null }
    
    foreach ($r in $registros) {
        $motivo = $null
        
        # 1. Verifica Codigo de Tombamento (PAT-XXXX) - Campo principal imutavel
        $tombamentoExistente = if ($r.Codigo_Tombamento -and $r.Codigo_Tombamento -ne "N/A") { $r.Codigo_Tombamento } elseif ($r.Patrimonio -and $r.Patrimonio -ne "N/A") { $r.Patrimonio } else { "" }
        if (![string]::IsNullOrWhiteSpace($codigoTombamento) -and ![string]::IsNullOrWhiteSpace($tombamentoExistente)) {
            if ($tombamentoExistente.Trim().ToUpper() -eq $codigoTombamento.Trim().ToUpper()) {
                $motivo = "Codigo de Tombamento repetido ($codigoTombamento)"
            }
        }
        
        # 2. Verifica Numero da Nota Fiscal (se informado)
        if (-not $motivo -and ![string]::IsNullOrWhiteSpace($numeroNotaFiscal) -and $numeroNotaFiscal.Trim() -notmatch '(?i)^n/a|none|default|0000$') {
            $nfExistente = if ($r.Numero_Nota_Fiscal -and $r.Numero_Nota_Fiscal -ne "N/A") { $r.Numero_Nota_Fiscal } else { "" }
            if (![string]::IsNullOrWhiteSpace($nfExistente) -and $nfExistente.Trim().ToUpper() -eq $numeroNotaFiscal.Trim().ToUpper()) {
                $motivo = "Numero de Nota Fiscal identico ($numeroNotaFiscal)"
            }
        }
        
        # 3. Verifica CRM/Serial (se nao for N/A ou generico)
        $serialExistente = if ($r.CRM_Serial -and $r.CRM_Serial -ne "N/A") { $r.CRM_Serial } elseif ($r.Numero_Serie -and $r.Numero_Serie -ne "N/A") { $r.Numero_Serie } else { "" }
        if (-not $motivo -and ![string]::IsNullOrWhiteSpace($crmSerial) -and $crmSerial.Trim() -notmatch '(?i)^n/a|none|default|to be filled|0000$') {
            if (![string]::IsNullOrWhiteSpace($serialExistente) -and $serialExistente.Trim().ToUpper() -eq $crmSerial.Trim().ToUpper()) {
                $motivo = "CRM/Serial identico ($crmSerial)"
            }
        }
        
        # 4. Verifica Endereco MAC (se nao for N/A) - para equipamentos de rede
        if (-not $motivo -and ![string]::IsNullOrWhiteSpace($macRede) -and $macRede.Trim() -notmatch '(?i)^n/a$') {
            if (![string]::IsNullOrWhiteSpace($r.MAC_Rede) -and $r.MAC_Rede.Trim().ToUpper() -eq $macRede.Trim().ToUpper()) {
                $motivo = "MAC Address identico ($macRede)"
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
            $setoresHistorico = @($registrosAtuais | ForEach-Object { Obter-SetorDoRegistro $_ } | Where-Object { $_ -ne "" -and $_ -notmatch '^\d+$' } | Select-Object -Unique)
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
            $qtdSetor = @($registrosAtuais | Where-Object { (Obter-SetorDoRegistro $_) -eq $listaSetores[$i] }).Count
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
        $setoresHistorico = @($registrosAtuais | ForEach-Object { Obter-SetorDoRegistro $_ } | Where-Object { $_ -ne "" -and $_ -notmatch '^\d+$' } | Select-Object -Unique)
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
        $qtdSetor = @($registrosAtuais | Where-Object { (Obter-SetorDoRegistro $_) -eq $listaSetores[$i] }).Count
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
# MODULO 1 - REGISTRO DE ATIVOS E PATRIMONIO (BENS GERAIS) - v2.0
# ==============================================================================
# --------------------------------------------------------------------------
# HELPERS DE FORMATACAO E VALIDACAO
# --------------------------------------------------------------------------
function Normalizar-ValorMoeda {
    param([string]$entrada)
    if ([string]::IsNullOrWhiteSpace($entrada)) { return "0,00" }
    $v = $entrada.Trim() -replace '(?i)R\$', '' -replace '\s', ''
    if ($v -eq "") { return "0,00" }
    if ($v -match ',\d{1,2}$') { $v = ($v -replace '\.', '') -replace ',', '.' }
    $num = 0.0
    if ([double]::TryParse($v, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$num)) {
        return $num.ToString("0.00", [System.Globalization.CultureInfo]::InvariantCulture).Replace('.', ',')
    }
    return "0,00"
}

function Converter-ValorNumerico {
    param([string]$valor)
    if ([string]::IsNullOrWhiteSpace($valor)) { return 0.0 }
    $v = $valor.Trim() -replace '(?i)R\$', '' -replace '\s', ''
    if ($v -eq "" -or $v -eq "N/A") { return 0.0 }
    if ($v -match ',\d{1,2}$') { $v = ($v -replace '\.', '') -replace ',', '.' }
    $num = 0.0
    if ([double]::TryParse($v, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$num)) {
        return $num
    }
    return 0.0
}

function Formatar-Moeda {
    param([double]$valor)
    return $valor.ToString("N2", [System.Globalization.CultureInfo]::GetCultureInfo("pt-BR"))
}

function Validar-DataBR {
    param([string]$data)
    if ([string]::IsNullOrWhiteSpace($data)) { return $false }
    if ($data -match '^(\d{2})/(\d{2})/(\d{4})$') {
        $dia = [int]$Matches[1]; $mes = [int]$Matches[2]; $ano = [int]$Matches[3]
        if ($dia -ge 1 -and $dia -le 31 -and $mes -ge 1 -and $mes -le 12 -and $ano -ge 1900 -and $ano -le 2100) {
            try { [void](New-Object System.DateTime($ano, $mes, $dia)); return $true } catch { return $false }
        }
    }
    return $false
}

function Cortar-Texto {
    param([string]$texto, [int]$tamanho)
    if ($null -eq $texto) { return "" }
    if ($texto.Length -le $tamanho) { return $texto }
    if ($tamanho -le 1) { return $texto.Substring(0, [Math]::Max(0, $tamanho)) }
    return $texto.Substring(0, $tamanho - 1) + "."
}

function Obter-CorSituacao {
    param([string]$situacao)
    switch -Regex ($situacao) {
        'Ativo|Operacional' { return "Green" }
        'Manuten'            { return "Yellow" }
        'Estoque|Disponivel' { return "Cyan" }
        'Baixado|Inativo'     { return "Red" }
        'Emprestado|Cedido'  { return "Magenta" }
        default              { return "White" }
    }
}

# --------------------------------------------------------------------------
# SELETOR COM FILTRO INCREMENTAL
# Digite para filtrar (ex: "ar con") ou o numero da opcao. Muito mais rapido
# do que rolar uma lista com 50 itens.
# --------------------------------------------------------------------------
function Selecionar-Opcao {
    param(
        [string]$titulo,
        [array]$opcoes,
        [string]$padrao = "",
        [switch]$PermiteNovo,
        [string]$rotuloNovo = "Digitar outro",
        [int]$maximoVisiveis = 12
    )

    $lista = @($opcoes)
    if ($lista.Count -eq 0) {
        if ($PermiteNovo) {
            $novo = ""
            while ([string]::IsNullOrWhiteSpace($novo)) {
                $novo = Read-Host "   $titulo"
                if ($null -eq $novo) { return $null }
            }
            return $novo
        }
        return $null
    }

    $filtro = ""
    while ($true) {
        Write-Host ""
        Write-Host "   $titulo" -ForegroundColor Yellow
        if ($filtro -ne "") {
            Write-Host "   Filtro ativo: " -NoNewline -ForegroundColor DarkGray
            Write-Host "$filtro" -ForegroundColor Cyan
        }

        $exibidos = if ($filtro -ne "") { @($lista | Where-Object { $_ -match "(?i)$([regex]::Escape($filtro))" }) } else { $lista }
        if ($exibidos.Count -eq 0) {
            Write-Host "   [!] Nenhum item encontrado para '$filtro'. Limpando o filtro..." -ForegroundColor Red
            $filtro = ""
            continue
        }

        # Correspondencia exata: seleciona direto
        $exato = $exibidos | Where-Object { $_.ToUpper() -eq $filtro.ToUpper() } | Select-Object -First 1
        if ($exato) { return $exato }

        $mostrar = if ($exibidos.Count -gt $maximoVisiveis) { @($exibidos[0..($maximoVisiveis - 1)]) } else { $exibidos }
        for ($i = 0; $i -lt $mostrar.Count; $i++) {
            $n = ($i + 1).ToString().PadLeft(2, ' ')
            Write-Host "     [$n] " -NoNewline -ForegroundColor Cyan
            Write-Host "$($mostrar[$i])" -ForegroundColor White
        }
        if ($exibidos.Count -gt $mostrar.Count) {
            Write-Host "     ... mais $($exibidos.Count - $mostrar.Count) item(ns) - digite texto para filtrar" -ForegroundColor DarkGray
        }
        if ($PermiteNovo) {
            Write-Host "     [ 0] " -NoNewline -ForegroundColor DarkGray
            Write-Host "$rotuloNovo" -ForegroundColor DarkGray
        }

        $resposta = Read-Host "   Numero, texto para filtrar ou ENTER"
        if ($null -eq $resposta) { return $null }
        $resposta = $resposta.Trim()

        if ($resposta -eq "") {
            if ($padrao -ne "") { return $padrao }
            continue
        }

        $num = 0
        if ([int]::TryParse($resposta, [ref]$num)) {
            if ($num -eq 0 -and $PermiteNovo) { return "__NOVO__" }
            if ($num -ge 1 -and $num -le $exibidos.Count) { return $exibidos[$num - 1] }
            Write-Host "   [!] Opcao invalida." -ForegroundColor Red
            continue
        }

        $filtro = $resposta
    }
}

# --------------------------------------------------------------------------
# CADASTRO MESTRE DE COLABORADORES E FORNECEDORES
# --------------------------------------------------------------------------
function Obter-MestreCsv {
    param([string]$caminho)
    $lista = @()
    if (!(Test-Path $caminho)) { return ,$lista }
    try {
        $raw = [System.IO.File]::ReadAllText($caminho, [System.Text.Encoding]::UTF8)
        $linhas = $raw -split "\r?\n" | Where-Object { ![string]::IsNullOrWhiteSpace($_) }
        for ($i = 1; $i -lt $linhas.Count; $i++) {
            $v = $linhas[$i] -split ';'
            if ($v.Count -lt 2) { continue }
            $lista += [PSCustomObject]@{
                Codigo  = $v[0].Trim()
                Nome    = $v[1].Trim()
                Campo3  = $(if ($v.Count -gt 2) { $v[2].Trim() } else { "" })
                Campo4  = $(if ($v.Count -gt 3) { $v[3].Trim() } else { "" })
                Campo5  = $(if ($v.Count -gt 4) { $v[4].Trim() } else { "" })
                Campo6  = $(if ($v.Count -gt 5) { $v[5].Trim() } else { "" })
                Ativo   = $(if ($v.Count -gt 6) { $v[6].Trim() } else { "Sim" })
            }
        }
    } catch {
        Write-Host " [!] Erro ao ler o cadastro mestre: $($_.Exception.Message)" -ForegroundColor Red
    }
    return ,$lista
}

function Salvar-MestreCsv {
    param([string]$caminho, [array]$registros, [string[]]$cabecalho)
    $pasta = [System.IO.Path]::GetDirectoryName($caminho)
    if ($pasta -and !(Test-Path $pasta)) { try { New-Item -ItemType Directory -Path $pasta -Force | Out-Null } catch {} }
    $enc = New-Object System.Text.UTF8Encoding($true)
    $linhas = [System.Collections.Generic.List[string]]::new()
    $linhas.Add(($cabecalho -join ';'))
    foreach ($r in $registros) {
        $campos = @($r.Codigo, $r.Nome, $r.Campo3, $r.Campo4, $r.Campo5, $r.Campo6, $r.Ativo)
        $linhas.Add((($campos | ForEach-Object { ($_.ToString() -replace ';', ',') -replace "[\r\n]+", " " }) -join ';'))
    }
    [System.IO.File]::WriteAllText($caminho, (($linhas -join "`r`n") + "`r`n"), $enc)
}

function Obter-ProximoCodigoMestre {
    param([array]$registros, [string]$prefixo)
    $maior = 0
    foreach ($r in $registros) {
        if ($r.Codigo -match "$prefixo-(\d+)") {
            $n = [int]$Matches[1]
            if ($n -gt $maior) { $maior = $n }
        }
    }
    return ("{0}-{1}" -f $prefixo, ($maior + 1).ToString("D4"))
}

function Menu-GerenciarParceiros {
    param()

    $emGestao = $true
    while ($emGestao) {
        $colaboradores = Obter-MestreCsv -caminho $script:caminhoCatalogoColaboradores
        $fornecedores  = Obter-MestreCsv -caminho $script:caminhoCatalogoFornecedores

        try { Clear-Host } catch {}
        Exibir-LogoPrincipal
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host " |          >>> CADASTRO MESTRE - COLABORADORES E FORNECEDORES <<<      |" -ForegroundColor Yellow
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host "  Colaboradores cadastrados: " -NoNewline -ForegroundColor Gray; Write-Host "$($colaboradores.Count)" -ForegroundColor Green
        Write-Host "  Fornecedores cadastrados:  " -NoNewline -ForegroundColor Gray; Write-Host "$($fornecedores.Count)" -ForegroundColor Green
        Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
        Write-Host "  [1] Listar colaboradores" -ForegroundColor White
        Write-Host "  [2] Cadastrar colaborador" -ForegroundColor Green
        Write-Host "  [3] Editar colaborador" -ForegroundColor Cyan
        Write-Host "  [4] Desativar / Reativar colaborador" -ForegroundColor Yellow
        Write-Host "  [5] Listar fornecedores" -ForegroundColor White
        Write-Host "  [6] Cadastrar fornecedor" -ForegroundColor Green
        Write-Host "  [7] Editar fornecedor" -ForegroundColor Cyan
        Write-Host "  [8] Desativar / Reativar fornecedor" -ForegroundColor Yellow
        Write-Host "  [9] Abrir as planilhas mestres no Excel" -ForegroundColor DarkCyan
        Write-Host "  [0] Voltar ao Módulo 1" -ForegroundColor Gray
        Write-Host ""

        $opc = Ler-OpcaoOuEsc -prompt " Escolha a opcao" -valorPadrao "1"

        switch ($opc.Trim()) {
            "1" {
                try { Clear-Host } catch {}
                Exibir-LogoPrincipal
                Write-Host " +======================================================================+" -ForegroundColor Cyan
                Write-Host " |                  >>> COLABORADORES CADASTRADOS <<<                   |" -ForegroundColor Yellow
                Write-Host " +======================================================================+" -ForegroundColor Cyan
                if ($colaboradores.Count -eq 0) {
                    Write-Host "  Nenhum colaborador no cadastro mestre." -ForegroundColor Yellow
                } else {
                    Write-Host "  CODIGO     NOME                              SETOR                 FUNCAO" -ForegroundColor Cyan
                    Write-Host "  --------  --------------------------------  --------------------  --------------" -ForegroundColor DarkGray
                    foreach ($c in $colaboradores) {
                        $status = if ($c.Ativo -eq "Sim") { "" } else { "  (INATIVO)" }
                        Write-Host "  " -NoNewline
                        Write-Host "$($c.Codigo)" -NoNewline -ForegroundColor Green
                        Write-Host "  " -NoNewline
                        Write-Host "$(Cortar-Texto $c.Nome 32)".PadRight(32) -NoNewline -ForegroundColor White
                        Write-Host "  " -NoNewline
                        Write-Host "$(Cortar-Texto $c.Campo3 20)".PadRight(20) -NoNewline -ForegroundColor Cyan
                        Write-Host "  " -NoNewline
                        Write-Host "$(Cortar-Texto $c.Campo4 14)".PadRight(14) -NoNewline -ForegroundColor DarkGray
                        Write-Host "$status" -ForegroundColor Red
                    }
                }
                Pausar-Tela
            }
            "2" {
                Write-Host ""
                Write-Host " -- NOVO COLABORADOR --" -ForegroundColor Green
                $nome = ""
                while ([string]::IsNullOrWhiteSpace($nome)) {
                    $nome = Read-Host "   Nome completo"
                    if ($null -eq $nome) { $nome = ""; break }
                }
                if ([string]::IsNullOrWhiteSpace($nome)) { continue }
                $setor = Read-Host "   Setor"
                $funcao = Read-Host "   Funcao / Cargo"
                $email = Read-Host "   E-mail"
                $telefone = Read-Host "   Telefone"
                $novo = [PSCustomObject]@{
                    Codigo = (Obter-ProximoCodigoMestre -registros $colaboradores -prefixo "COL")
                    Nome = $nome.Trim()
                    Campo3 = $(if ($setor) { $setor.Trim() } else { "N/A" })
                    Campo4 = $(if ($funcao) { $funcao.Trim() } else { "N/A" })
                    Campo5 = $(if ($email) { $email.Trim() } else { "N/A" })
                    Campo6 = $(if ($telefone) { $telefone.Trim() } else { "N/A" })
                    Ativo = "Sim"
                }
                Salvar-MestreCsv -caminho $script:caminhoCatalogoColaboradores -registros (@($colaboradores) + $novo) -cabecalho @("Codigo", "Nome", "Setor", "Funcao", "Email", "Telefone", "Ativo")
                Write-Host "`n [OK] Colaborador $($novo.Codigo) cadastrado." -ForegroundColor Green
                Pausar-Tela
            }
            "3" {
                if ($colaboradores.Count -eq 0) { Write-Host " [!] Nenhum colaborador cadastrado." -ForegroundColor Yellow; Pausar-Tela; continue }
                $escolha = Read-Host " Digite o codigo (COL-0001) ou o nome do colaborador"
                $sel = $colaboradores | Where-Object { $_.Codigo -match "(?i)$($escolha.Trim())" -or $_.Nome -match "(?i)$($escolha.Trim())" } | Select-Object -First 1
                if (!$sel) { Write-Host " [!] Colaborador nao encontrado." -ForegroundColor Red; Pausar-Tela; continue }
                Write-Host ""
                Write-Host "   Editando: $($sel.Codigo) - $($sel.Nome)" -ForegroundColor Cyan
                $nome = Read-Host "   Nome [$($sel.Nome)]"
                $setor = Read-Host "   Setor [$($sel.Campo3)]"
                $funcao = Read-Host "   Funcao [$($sel.Campo4)]"
                $email = Read-Host "   E-mail [$($sel.Campo5)]"
                $telefone = Read-Host "   Telefone [$($sel.Campo6)]"
                if ($nome) { $sel.Nome = $nome.Trim() }
                if ($setor) { $sel.Campo3 = $setor.Trim() }
                if ($funcao) { $sel.Campo4 = $funcao.Trim() }
                if ($email) { $sel.Campo5 = $email.Trim() }
                if ($telefone) { $sel.Campo6 = $telefone.Trim() }
                Salvar-MestreCsv -caminho $script:caminhoCatalogoColaboradores -registros $colaboradores -cabecalho @("Codigo", "Nome", "Setor", "Funcao", "Email", "Telefone", "Ativo")
                Write-Host "`n [OK] Colaborador atualizado." -ForegroundColor Green
                Pausar-Tela
            }
            "4" {
                if ($colaboradores.Count -eq 0) { Write-Host " [!] Nenhum colaborador cadastrado." -ForegroundColor Yellow; Pausar-Tela; continue }
                $escolha = Read-Host " Digite o codigo (COL-0001) ou o nome do colaborador"
                $sel = $colaboradores | Where-Object { $_.Codigo -match "(?i)$($escolha.Trim())" -or $_.Nome -match "(?i)$($escolha.Trim())" } | Select-Object -First 1
                if (!$sel) { Write-Host " [!] Collaborador nao encontrado." -ForegroundColor Red; Pausar-Tela; continue }
                $novoStatus = if ($sel.Ativo -eq "Sim") { "Nao" } else { "Sim" }
                $sel.Ativo = $novoStatus
                Salvar-MestreCsv -caminho $script:caminhoCatalogoColaboradores -registros $colaboradores -cabecalho @("Codigo", "Nome", "Setor", "Funcao", "Email", "Telefone", "Ativo")
                Write-Host "`n [OK] Status de $($sel.Nome) alterado para '$novoStatus'." -ForegroundColor Green
                Pausar-Tela
            }
            "5" {
                try { Clear-Host } catch {}
                Exibir-LogoPrincipal
                Write-Host " +======================================================================+" -ForegroundColor Cyan
                Write-Host " |                    >>> FORNECEDORES CADASTRADOS <<<                  |" -ForegroundColor Yellow
                Write-Host " +======================================================================+" -ForegroundColor Cyan
                if ($fornecedores.Count -eq 0) {
                    Write-Host "  Nenhum fornecedor no cadastro mestre." -ForegroundColor Yellow
                } else {
                    Write-Host "  CODIGO     RAZAO SOCIAL                      CNPJ                 CONTATO" -ForegroundColor Cyan
                    Write-Host "  --------  --------------------------------  -------------------  --------------" -ForegroundColor DarkGray
                    foreach ($f in $fornecedores) {
                        $status = if ($f.Ativo -eq "Sim") { "" } else { "  (INATIVO)" }
                        Write-Host "  " -NoNewline
                        Write-Host "$($f.Codigo)" -NoNewline -ForegroundColor Green
                        Write-Host "  " -NoNewline
                        Write-Host "$(Cortar-Texto $f.Nome 32)".PadRight(32) -NoNewline -ForegroundColor White
                        Write-Host "  " -NoNewline
                        Write-Host "$(Cortar-Texto $f.Campo3 19)".PadRight(19) -NoNewline -ForegroundColor Cyan
                        Write-Host "  " -NoNewline
                        Write-Host "$(Cortar-Texto $f.Campo4 14)".PadRight(14) -NoNewline -ForegroundColor DarkGray
                        Write-Host "$status" -ForegroundColor Red
                    }
                }
                Pausar-Tela
            }
            "6" {
                Write-Host ""
                Write-Host " -- NOVO FORNECEDOR --" -ForegroundColor Green
                $razao = ""
                while ([string]::IsNullOrWhiteSpace($razao)) {
                    $razao = Read-Host "   Razao social / nome do fornecedor"
                    if ($null -eq $razao) { $razao = ""; break }
                }
                if ([string]::IsNullOrWhiteSpace($razao)) { continue }
                $cnpj = Read-Host "   CNPJ"
                $contato = Read-Host "   Contato"
                $email = Read-Host "   E-mail"
                $telefone = Read-Host "   Telefone"
                $novo = [PSCustomObject]@{
                    Codigo = (Obter-ProximoCodigoMestre -registros $fornecedores -prefixo "FOR")
                    Nome = $razao.Trim()
                    Campo3 = $(if ($cnpj) { $cnpj.Trim() } else { "N/A" })
                    Campo4 = $(if ($contato) { $contato.Trim() } else { "N/A" })
                    Campo5 = $(if ($email) { $email.Trim() } else { "N/A" })
                    Campo6 = $(if ($telefone) { $telefone.Trim() } else { "N/A" })
                    Ativo = "Sim"
                }
                Salvar-MestreCsv -caminho $script:caminhoCatalogoFornecedores -registros (@($fornecedores) + $novo) -cabecalho @("Codigo", "Razao_Social", "CNPJ", "Contato", "Email", "Telefone", "Ativo")
                Write-Host "`n [OK] Fornecedor $($novo.Codigo) cadastrado." -ForegroundColor Green
                Pausar-Tela
            }
            "7" {
                if ($fornecedores.Count -eq 0) { Write-Host " [!] Nenhum fornecedor cadastrado." -ForegroundColor Yellow; Pausar-Tela; continue }
                $escolha = Read-Host " Digite o codigo (FOR-0001) ou o nome do fornecedor"
                $sel = $fornecedores | Where-Object { $_.Codigo -match "(?i)$($escolha.Trim())" -or $_.Nome -match "(?i)$($escolha.Trim())" } | Select-Object -First 1
                if (!$sel) { Write-Host " [!] Fornecedor nao encontrado." -ForegroundColor Red; Pausar-Tela; continue }
                Write-Host ""
                Write-Host "   Editando: $($sel.Codigo) - $($sel.Nome)" -ForegroundColor Cyan
                $razao = Read-Host "   Razao social [$($sel.Nome)]"
                $cnpj = Read-Host "   CNPJ [$($sel.Campo3)]"
                $contato = Read-Host "   Contato [$($sel.Campo4)]"
                $email = Read-Host "   E-mail [$($sel.Campo5)]"
                $telefone = Read-Host "   Telefone [$($sel.Campo6)]"
                if ($razao) { $sel.Nome = $razao.Trim() }
                if ($cnpj) { $sel.Campo3 = $cnpj.Trim() }
                if ($contato) { $sel.Campo4 = $contato.Trim() }
                if ($email) { $sel.Campo5 = $email.Trim() }
                if ($telefone) { $sel.Campo6 = $telefone.Trim() }
                Salvar-MestreCsv -caminho $script:caminhoCatalogoFornecedores -registros $fornecedores -cabecalho @("Codigo", "Razao_Social", "CNPJ", "Contato", "Email", "Telefone", "Ativo")
                Write-Host "`n [OK] Fornecedor atualizado." -ForegroundColor Green
                Pausar-Tela
            }
            "8" {
                if ($fornecedores.Count -eq 0) { Write-Host " [!] Nenhum fornecedor cadastrado." -ForegroundColor Yellow; Pausar-Tela; continue }
                $escolha = Read-Host " Digite o codigo (FOR-0001) ou o nome do fornecedor"
                $sel = $fornecedores | Where-Object { $_.Codigo -match "(?i)$($escolha.Trim())" -or $_.Nome -match "(?i)$($escolha.Trim())" } | Select-Object -First 1
                if (!$sel) { Write-Host " [!] Fornecedor nao encontrado." -ForegroundColor Red; Pausar-Tela; continue }
                $novoStatus = if ($sel.Ativo -eq "Sim") { "Nao" } else { "Sim" }
                $sel.Ativo = $novoStatus
                Salvar-MestreCsv -caminho $script:caminhoCatalogoFornecedores -registros $fornecedores -cabecalho @("Codigo", "Razao_Social", "CNPJ", "Contato", "Email", "Telefone", "Ativo")
                Write-Host "`n [OK] Status de $($sel.Nome) alterado para '$novoStatus'." -ForegroundColor Green
                Pausar-Tela
            }
            "9" {
                Write-Host " [OK] Abrindo as planilhas mestres..." -ForegroundColor Green
                try { Invoke-Item $script:caminhoCatalogoColaboradores } catch {}
                Start-Sleep -Milliseconds 400
                try { Invoke-Item $script:caminhoCatalogoFornecedores } catch {}
                Pausar-Tela
            }
            "0" { $emGestao = $false }
            default {
                Write-Host " Opcao invalida." -ForegroundColor Red
                Start-Sleep -Milliseconds 700
            }
        }
    }
}
# --------------------------------------------------------------------------
# LEITURA DE CAMPO COM VALOR PADRAO (ENTER MANTEM O PADRAO)
# Retorna $null se a entrada for encerrada (permite cancelar em lote)
# --------------------------------------------------------------------------
function Ler-Campo {
    param(
        [string]$rotulo,
        [string]$padrao = "",
        [switch]$Obrigatorio
    )
    while ($true) {
        $prompt = if ($padrao -ne "") { "   $rotulo [$padrao]" } else { "   $rotulo" }
        $v = Read-Host $prompt
        if ($null -eq $v) { return $null }
        $v = $v.Trim()
        if ($v -eq "") {
            if ($padrao -ne "") { return $padrao }
            if ($Obrigatorio) {
                Write-Host "   (!) Campo obrigatorio. Digite um valor." -ForegroundColor Red
                continue
            }
            return ""
        }
        return $v
    }
}

# --------------------------------------------------------------------------
# NOVO CADASTRO DE BEM IMOBILIZADO
# --------------------------------------------------------------------------
function Iniciar-CadastroBemPatrimonial {
    param([array]$registrosAtuais, [string]$caminhoCsv)

    if (!$script:contextoAtivos) { $script:contextoAtivos = @{} }
    $ctx = $script:contextoAtivos

    $catalogo = Obter-CatalogoAtivos
    $categorias = Obter-CategoriasCatalogo -catalogo $catalogo

    $colaboradoresMestre = Obter-MestreCsv -caminho $script:caminhoCatalogoColaboradores
    $fornecedoresMestre  = Obter-MestreCsv -caminho $script:caminhoCatalogoFornecedores

    $listaResponsaveis = @()
    foreach ($c in $colaboradoresMestre) { if ($c.Ativo -eq "Sim" -and $c.Nome) { $listaResponsaveis += $c.Nome } }
    $listaFornecedores = @()
    foreach ($f in $fornecedoresMestre) { if ($f.Ativo -eq "Sim" -and $f.Nome) { $listaFornecedores += $f.Nome } }

    # ==========================================================================
    # ESCOLHA DO MODO DE CADASTRO
    # ==========================================================================
    try { Clear-Host } catch {}
    Exibir-LogoPrincipal
    Write-Host " +======================================================================+" -ForegroundColor Cyan
    Write-Host " |            >>> MODULO 1 - NOVO CADASTRO DE BEM <<<                   |" -ForegroundColor Yellow
    Write-Host " +======================================================================+" -ForegroundColor Cyan
    Write-Host "  [1] CADASTRO COMPLETO - preencher todos os 15 campos da ficha" -ForegroundColor Green
    Write-Host "  [2] CADASTRO RAPIDO   - herda setor/responsavel/situacao (lote)" -ForegroundColor Cyan
    Write-Host "  [0] Voltar ao Módulo 1" -ForegroundColor Gray
    Write-Host ""
    $modo = Ler-OpcaoOuEsc -prompt " Escolha o modo de cadastro" -valorPadrao "2"
    if ($modo.Trim() -eq "0") { return }
    $modoRapido = ($modo.Trim() -eq "2")

    $contador = 0
    while ($true) {
        $contador++
        try { Clear-Host } catch {}
        Exibir-LogoPrincipal
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        $tituloModo = if ($modoRapido) { "CADASTRO RAPIDO" } else { "CADASTRO COMPLETO" }
        Write-Host " |     >>> $tituloModo - BEM #$contador - PATRIMONIO <<<             |" -ForegroundColor Yellow
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host "  Proximo tombamento disponivel: " -NoNewline -ForegroundColor Gray
        $tombamentoSugerido = Obter-ProximoPatrimonio -registros $registrosAtuais
        Write-Host "$tombamentoSugerido" -ForegroundColor Green
        Write-Host ""

        $cancelar = $false

        # ------------------------------------------------------------------
        # [1] CATEGORIA DO BEM
        # ------------------------------------------------------------------
        $categoriaBem = $ctx["Categoria"]
        if ($categorias.Count -gt 0) {
            $opc = Selecionar-Opcao -titulo "Categoria do Bem" -opcoes $categorias -padrao $categoriaBem -PermiteNovo -rotuloNovo "Digitar outra categoria"
            if ($null -eq $opc) { $cancelar = $true }
            elseif ($opc -eq "__NOVO__") {
                $categoriaBem = Ler-Campo -rotulo "   Nome da nova categoria" -Obrigatorio
                if ($null -eq $categoriaBem) { $cancelar = $true }
            } else { $categoriaBem = $opc }
        } else {
            $categoriaBem = Ler-Campo -rotulo "Categoria do Bem" -padrao $categoriaBem
            if ($null -eq $categoriaBem) { $cancelar = $true }
        }
        if ($cancelar) { break }
        Write-Host "   -> Categoria: " -NoNewline -ForegroundColor Gray; Write-Host "$categoriaBem" -ForegroundColor Green

        # ------------------------------------------------------------------
        # [2] TIPO / SUBTIPO (catalogo)
        # ------------------------------------------------------------------
        $tiposDaCategoria = Obter-TiposPorCategoria -catalogo $catalogo -categoria $categoriaBem
        $tipoEspecifico = $ctx["Tipo"]
        if ($tiposDaCategoria.Count -gt 0) {
            $opc = Selecionar-Opcao -titulo "Tipo / Subtipo do Bem" -opcoes $tiposDaCategoria -padrao $tipoEspecifico -PermiteNovo -rotuloNovo "Digitar outro tipo"
            if ($null -eq $opc) { $cancelar = $true }
            elseif ($opc -eq "__NOVO__") {
                $tipoEspecifico = Ler-Campo -rotulo "   Novo tipo" -Obrigatorio
                if ($null -eq $tipoEspecifico) { $cancelar = $true }
            } else { $tipoEspecifico = $opc }
        } else {
            $tipoEspecifico = Ler-Campo -rotulo "Tipo / Subtipo do Bem" -padrao $tipoEspecifico
            if ($null -eq $tipoEspecifico) { $cancelar = $true }
        }
        if ($cancelar) { break }
        Write-Host "   -> Tipo: " -NoNewline -ForegroundColor Gray; Write-Host "$tipoEspecifico" -ForegroundColor Green

        # Sugestoes vindas do catalogo
        $itemCatalogo = $catalogo | Where-Object { $_.Categoria -eq $categoriaBem -and $_.Tipo_Especifico -eq $tipoEspecifico } | Select-Object -First 1
        $nomeSugerido = if ($itemCatalogo) { $itemCatalogo.Nome_Bem_Produto } else { "" }
        $setorSugerido = if ($itemCatalogo) { $itemCatalogo.Setor_Sugerido } else { "" }

        # ------------------------------------------------------------------
        # [3] NOME DO BEM IMOBILIZADO / PRODUTO
        # ------------------------------------------------------------------
        $nomeBem = Ler-Campo -rotulo "Nome do Bem Imobilizado / Produto" -padrao $nomeSugerido -Obrigatorio
        if ($null -eq $nomeBem) { $cancelar = $true; break }

        # ------------------------------------------------------------------
        # [4] CODIGO / TOMBAMENTO (imutavel) - validacao imediata
        # ------------------------------------------------------------------
        $codigoTombamento = Ler-Campo -rotulo "Codigo / Tombamento" -padrao $tombamentoSugerido -Obrigatorio
        if ($null -eq $codigoTombamento) { $cancelar = $true; break }

        $dupTombamento = Testar-DuplicidadeBem -registros $registrosAtuais -codigoTombamento $codigoTombamento -crmSerial ""
        if ($dupTombamento) {
            Write-Host ""
            Write-Host " [!] Este codigo de tombamento JA EXISTE: " -NoNewline -ForegroundColor Red
            Write-Host "$($dupTombamento.RegistroExistente.Codigo_Tombamento) - $($dupTombamento.RegistroExistente.Nome_Bem_Produto)" -ForegroundColor Yellow
            Write-Host "     [1] Usar o proximo tombamento livre ($tombamentoSugerido)" -ForegroundColor Green
            Write-Host "     [2] Manter este codigo (sera tratado como duplicidade no final)" -ForegroundColor Yellow
            Write-Host "     [0] Cancelar o cadastro" -ForegroundColor Gray
            $opcDup = Read-Host "     Escolha [1]"
            if ($null -eq $opcDup) { $cancelar = $true; break }
            if ($opcDup.Trim() -eq "0") { $cancelar = $true; break }
            if ($opcDup.Trim() -eq "1") {
                $codigoTombamento = $tombamentoSugerido
                Write-Host "   -> Tombamento ajustado para: $codigoTombamento" -ForegroundColor Green
            }
        }
        Write-Host "   -> Tombamento: " -NoNewline -ForegroundColor Gray; Write-Host "$codigoTombamento" -ForegroundColor Green

        # ------------------------------------------------------------------
        # [5] NOTA FISCAL
        # ------------------------------------------------------------------
        $numeroNotaFiscal = Ler-Campo -rotulo "Numero da Nota Fiscal (a mesma nota cobre varios bens)" -padrao $ctx["NotaFiscal"]
        if ($null -eq $numeroNotaFiscal) { $cancelar = $true; break }
        if ($numeroNotaFiscal -eq "") { $numeroNotaFiscal = "N/A" }

        # ------------------------------------------------------------------
        # [6] CRM / SERIAL
        # ------------------------------------------------------------------
        $crmSerial = Ler-Campo -rotulo "CRM / Serial (serial, RFID, codigo de barras)" -padrao $ctx["CRM"]
        if ($null -eq $crmSerial) { $cancelar = $true; break }
        if ($crmSerial -eq "") { $crmSerial = "N/A" }

        # ------------------------------------------------------------------
        # [7] EMPRESA / UNIDADE
        # ------------------------------------------------------------------
        $empresa = $ctx["Empresa"]
        if (!$empresa) { $empresa = $script:empresasAtivosPadrao[0] }
        if (!$modoRapido) {
            $opc = Selecionar-Opcao -titulo "Empresa / Unidade" -opcoes $script:empresasAtivosPadrao -padrao $empresa -PermiteNovo -rotuloNovo "Digitar outra empresa"
            if ($null -eq $opc) { $cancelar = $true; break }
            if ($opc -eq "__NOVO__") {
                $empresa = Ler-Campo -rotulo "   Nome da empresa / unidade" -Obrigatorio
                if ($null -eq $empresa) { $cancelar = $true; break }
            } else { $empresa = $opc }
        }

        # ------------------------------------------------------------------
        # [8] SETOR
        # ------------------------------------------------------------------
        $setorAtribuido = $ctx["Setor"]
        if (!$setorAtribuido -and $setorSugerido) { $setorAtribuido = $setorSugerido }
        if (!$modoRapido) {
            $opco = Selecionar-Opcao -titulo "Atribuido a: Setor" -opcoes $script:setoresAtivosPadrao -padrao $setorAtribuido -PermiteNovo -rotuloNovo "Digitar outro setor"
            if ($null -eq $opc) { $cancelar = $true; break }
            if ($opc -eq "__NOVO__") {
                $setorAtribuido = Ler-Campo -rotulo "   Nome do setor" -Obrigatorio
                if ($null -eq $setorAtribuido) { $cancelar = $true; break }
            } else { $setorAtribuido = $opc }
        }
        if (!$setorAtribuido) { $setorAtribuido = $script:setoresAtivosPadrao[0] }
        Write-Host "   -> Setor: " -NoNewline -ForegroundColor Gray; Write-Host "$setorAtribuido" -ForegroundColor Green

        # ------------------------------------------------------------------
        # [9] RESPONSAVEL / COLABORADOR
        # ------------------------------------------------------------------
        $responsavel = $ctx["Responsavel"]
        if (!$modoRapido) {
            $opc = Selecionar-Opcao -titulo "Responsavel / Colaborador" -opcoes $listaResponsaveis -padrao $responsavel -PermiteNovo -rotuloNovo "Digitar outro responsavel"
            if ($null -eq $opc) { $cancelar = $true; break }
            if ($opc -eq "__NOVO__") {
                $responsavel = Ler-Campo -rotulo "   Nome do responsavel" -Obrigatorio
                if ($null -eq $responsavel) { $cancelar = $true; break }
            } else { $responsavel = $opc }
        }
        if (!$responsavel) { $responsavel = Ler-Campo -rotulo "   Nome do responsavel" -Obrigatorio; if ($null -eq $responsavel) { $cancelar = $true; break } }
        Write-Host "   -> Responsavel: " -NoNewline -ForegroundColor Gray; Write-Host "$responsavel" -ForegroundColor White

        # ------------------------------------------------------------------
        # [10] FORNECEDOR
        # ------------------------------------------------------------------
        $fornecedor = $ctx["Fornecedor"]
        if (!$modoRapido) {
            $opc = Selecionar-Opcao -titulo "Fornecedor" -opcoes $listaFornecedores -padrao $fornecedor -PermiteNovo -rotuloNovo "Digitar outro fornecedor"
            if ($null -eq $opc) { $cancelar = $true; break }
            if ($opc -eq "__NOVO__") {
                $fornecedor = Ler-Campo -rotulo "   Nome do fornecedor"
                if ($null -eq $fornecedor) { $cancelar = $true; break }
            } else { $fornecedor = $opc }
        }
        if (!$fornecedor) { $fornecedor = "N/A" }
        if ($fornecedor -eq "") { $fornecedor = "N/A" }

        # ------------------------------------------------------------------
        # [11] SITUACAO
        # ------------------------------------------------------------------
        $situacao = $ctx["Situacao"]
        if (!$situacao) { $situacao = $script:situacoesAtivosPadrao[0] }
        if (!$modoRapido) {
            $opc = Selecionar-Opcao -titulo "Situacao" -opcoes $script:situacoesAtivosPadrao -padrao $situacao
            if ($null -eq $opc) { $cancelar = $true; break }
            $situacao = $opc
        }
        Write-Host "   -> Situacao: " -NoNewline -ForegroundColor Gray; Write-Host "$situacao" -ForegroundColor (Obter-CorSituacao $situacao)

        # ------------------------------------------------------------------
        # [12] DATA DA COMPRA
        # ------------------------------------------------------------------
        $dataPadrao = if ($ctx["DataCompra"]) { $ctx["DataCompra"] } else { Get-Date -Format "dd/MM/yyyy" }
        $dataCompra = ""
        while ($true) {
            $dataCompra = Ler-Campo -rotulo "Data da Compra (DD/MM/AAAA)" -padrao $dataPadrao -Obrigatorio
            if ($null -eq $dataCompra) { $cancelar = $true; break }
            if (Validar-DataBR $dataCompra) { break }
            Write-Host "   (!) Data invalida. Use o formato DD/MM/AAAA." -ForegroundColor Red
        }
        if ($cancelar) { break }

        # ------------------------------------------------------------------
        # [13] VALOR DA COMPRA
        # ------------------------------------------------------------------
        $valorPadrao = if ($ctx["Valor"]) { $ctx["Valor"] } else { "0,00" }
        $valorCompra = Normalizar-ValorMoeda (Ler-Campo -rotulo "Valor da Compra (R$)" -padrao $valorPadrao)
        if ($null -eq $valorCompra) { $cancelar = $true; break }

        # ------------------------------------------------------------------
        # [14] LOCALIZACAO FISICA / POSTO
        # ------------------------------------------------------------------
        $localizacaoPadrao = if ($ctx["Localizacao"]) { $ctx["Localizacao"] } else { $setorAtribuido }
        $localizacao = Ler-Campo -rotulo "Localizacao Fisica / Posto" -padrao $localizacaoPadrao
        if ($null -eq $localizacao) { $cancelar = $true; break }
        if ($localizacao -eq "") { $localizacao = $setorAtribuido }

        # ------------------------------------------------------------------
        # [15] OBSERVACOES
        # ------------------------------------------------------------------
        $observacoes = Ler-Campo -rotulo "Observacoes Adicionais" -padrao $ctx["Observacoes"]
        if ($null -eq $observacoes) { $cancelar = $true; break }
        if ($observacoes -eq "") { $observacoes = "N/A" }

        if ($cancelar) { break }

        $dataRegistro = Get-Date -Format "dd/MM/yyyy HH:mm:ss"

        # ======================================================================
        # RESUMO DA FICHA
        # ======================================================================
        $divisoria = "=" * 70
        $resumoTexto = @(
            $divisoria,
            "       INNOVCORE TI - FICHA PATRIMONIAL DO BEM ($empresa)",
            $divisoria,
            "  - Nome do Bem:              $nomeBem",
            "  - Codigo / Tombamento:      $codigoTombamento",
            "  - Numero da Nota Fiscal:    $numeroNotaFiscal",
            "  - CRM / Serial:             $crmSerial",
            "  - Empresa / Unidade:        $empresa",
            "  - Categoria do Bem:         $categoriaBem",
            "  - Tipo / Subtipo:           $tipoEspecifico",
            "  - Atribuido a (Setor):      $setorAtribuido",
            "  - Responsavel:              $responsavel",
            "  - Fornecedor:               $fornecedor",
            "  - Situacao:                 $situacao",
            "  - Data da Compra:           $dataCompra",
            "  - Valor da Compra:          R$ $valorCompra",
            "  - Localizacao Fisica:       $localizacao",
            "  - Observacoes:              $observacoes",
            "  - Data do Registro:         $dataRegistro",
            $divisoria
        ) -join "`r`n"

        try { Clear-Host } catch {}
        Exibir-LogoPrincipal
        Write-Host ""
        Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
        Write-Host " |               FICHA PATRIMONIAL - CONFIRMACAO                       |" -ForegroundColor Yellow
        Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
        Write-Host "   NOME DO BEM:          " -NoNewline -ForegroundColor Gray; Write-Host "$nomeBem" -ForegroundColor White
        Write-Host "   CODIGO / TOMBAMENTO:  " -NoNewline -ForegroundColor Gray; Write-Host "$codigoTombamento" -ForegroundColor Green
        Write-Host "   NOTA FISCAL:          " -NoNewline -ForegroundColor Gray; Write-Host "$numeroNotaFiscal" -ForegroundColor Yellow
        Write-Host "   CRM / SERIAL:         " -NoNewline -ForegroundColor Gray; Write-Host "$crmSerial" -ForegroundColor Yellow
        Write-Host "   EMPRESA / UNIDADE:    " -NoNewline -ForegroundColor Gray; Write-Host "$empresa" -ForegroundColor White
        Write-Host "   CATEGORIA:            " -NoNewline -ForegroundColor Gray; Write-Host "$categoriaBem" -ForegroundColor Cyan
        Write-Host "   TIPO:                 " -NoNewline -ForegroundColor Gray; Write-Host "$tipoEspecifico" -ForegroundColor Cyan
        Write-Host "   SETOR:                " -NoNewline -ForegroundColor Gray; Write-Host "$setorAtribuido" -ForegroundColor Green
        Write-Host "   RESPONSAVEL:          " -NoNewline -ForegroundColor Gray; Write-Host "$responsavel" -ForegroundColor White
        Write-Host "   FORNECEDOR:           " -NoNewline -ForegroundColor Gray; Write-Host "$fornecedor" -ForegroundColor White
        Write-Host "   SITUACAO:             " -NoNewline -ForegroundColor Gray; Write-Host "$situacao" -ForegroundColor (Obter-CorSituacao $situacao)
        Write-Host "   DATA DA COMPRA:       " -NoNewline -ForegroundColor Gray; Write-Host "$dataCompra" -ForegroundColor Yellow
        Write-Host "   VALOR DA COMPRA:      " -NoNewline -ForegroundColor Gray; Write-Host "R$ $valorCompra" -ForegroundColor Yellow
        Write-Host "   LOCALIZACAO FISICA:   " -NoNewline -ForegroundColor Gray; Write-Host "$localizacao" -ForegroundColor Green
        Write-Host "   OBSERVACOES:          " -NoNewline -ForegroundColor Gray; Write-Host "$observacoes" -ForegroundColor White
        Write-Host "   DATA DO REGISTRO:     " -NoNewline -ForegroundColor Gray; Write-Host "$dataRegistro" -ForegroundColor DarkGray
        Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan

        try { Set-Clipboard -Value $resumoTexto } catch { try { $resumoTexto | clip.exe } catch {} }
        Write-Host ""
        Write-Host " [OK] Ficha COPIADA para a Area de Transferencia (Ctrl + V)!" -ForegroundColor Green

        # ======================================================================
        # GRAVACAO
        # ======================================================================
        $novoItem = [PSCustomObject]@{
            Data_Registro              = $dataRegistro
            Empresa_Unidade           = $empresa
            Codigo_Tombamento          = $codigoTombamento
            Numero_Nota_Fiscal         = $numeroNotaFiscal
            CRM_Serial                 = $crmSerial
            Categoria_Bem              = $categoriaBem
            Tipo_Especifico            = $tipoEspecifico
            Nome_Bem_Produto           = $nomeBem
            Setor_Atribuido            = $setorAtribuido
            Responsavel_Colaborador    = $responsavel
            Fornecedor                 = $fornecedor
            Situacao                   = $situacao
            Data_Compra                = $dataCompra
            Valor_Compra               = $valorCompra
            Localizacao_Fisica_Posto   = $localizacao
            Observacoes_Adicionais     = $observacoes
            Nome_Computador            = "N/A"
            Processador                = "N/A"
            Memoria_RAM                = "N/A"
            Armazenamento              = "N/A"
            Placa_Video                = "N/A"
            Sistema_Operacional        = "N/A"
            IP_Rede                    = "N/A"
            MAC_Rede                   = "N/A"
            Bateria                    = "N/A"
            Android_Versao             = "N/A"
            Empresa                    = $empresa
            Patrimonio                 = $codigoTombamento
            Usuario                    = $responsavel
            Setor                      = $setorAtribuido
            Status_Ativo               = $situacao
            Localizacao                = $localizacao
            Observacoes                = $observacoes
        }

        $dupFinal = Testar-DuplicidadeBem -registros $registrosAtuais -codigoTombamento $codigoTombamento -crmSerial $crmSerial
        if ($dupFinal) {
            $itemAntigo = $dupFinal.RegistroExistente
            Write-Host ""
            Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Yellow
            Write-Host " |               [!] DUPLICIDADE DETECTADA                           |" -ForegroundColor Yellow
            Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Yellow
            Write-Host "   Motivo:             " -NoNewline -ForegroundColor Gray; Write-Host "$($dupFinal.Motivo)" -ForegroundColor Red
            Write-Host "   Registro Anterior:  " -NoNewline -ForegroundColor Gray; Write-Host "$($itemAntigo.Codigo_Tombamento) - $($itemAntigo.Nome_Bem_Produto) ($($itemAntigo.Setor_Atribuido))" -ForegroundColor Cyan
            Write-Host ""
            Write-Host "   [1] CANCELAR este cadastro (padrao)" -ForegroundColor Gray
            Write-Host "   [2] MANTER AMBOS e gerar um novo codigo de tombamento" -ForegroundColor Cyan
            Write-Host "   [3] SOBRESCREVER o cadastro existente" -ForegroundColor Red
            Write-Host ""
            $opcDup = Read-Host "   Escolha uma opcao [1]"
            if ($null -eq $opcDup) { break }
            if ($opcDup.Trim() -eq "2") {
                $codigoTombamento = Obter-ProximoPatrimonio -registros $registrosAtuais
                $novoItem.Codigo_Tombamento = $codigoTombamento
                $novoItem.Patrimonio = $codigoTombamento
                Write-Host "   -> Novo codigo de tombamento: $codigoTombamento" -ForegroundColor Green
            } elseif ($opcDup.Trim() -eq "3") {
                $listaAtualizada = @($registrosAtuais | Where-Object { $_.Codigo_Tombamento -ne $itemAntigo.Codigo_Tombamento }) + $novoItem
                Salvar-RegistrosCsv -caminho $caminhoCsv -registros $listaAtualizada -BaseAtivos
                $registrosAtuais = Obter-RegistrosCsv -caminho $caminhoCsv -empresaPadrao "InnovEdu" -setorPadrao "Operacoes e Logistica"
                Write-Host "`n [OK] Cadastro $codigoTombamento ATUALIZADO com sucesso!" -ForegroundColor Green
            } else {
                Write-Host "`n [!] Operacao cancelada. Nenhum dado foi alterado." -ForegroundColor Yellow
                break
            }
        } else {
            $listaAtualizada = @($registrosAtuais) + $novoItem
            Salvar-RegistrosCsv -caminho $caminhoCsv -registros $listaAtualizada -BaseAtivos
            $registrosAtuais = Obter-RegistrosCsv -caminho $caminhoCsv -empresaPadrao "InnovEdu" -setorPadrao "Operacoes e Logistica"
            Write-Host ""
            Write-Host " [OK] Bem patrimonial cadastrado com sucesso!" -ForegroundColor Green
            Write-Host "      Tombamento: " -NoNewline -ForegroundColor Gray; Write-Host "$codigoTombamento" -ForegroundColor Green
            Write-Host "      Arquivo:    " -NoNewline -ForegroundColor Gray; Write-Host "$caminhoCsv" -ForegroundColor Yellow
        }

        # ======================================================================
        # MEMORIZAR CONTEXTO PARA O PROXIMO ITEM (modo lote)
        # ======================================================================
        $ctx["Categoria"]  = $categoriaBem
        $ctx["Tipo"]       = $tipoEspecifico
        $ctx["Empresa"]    = $empresa
        $ctx["Setor"]      = $setorAtribuido
        $ctx["Responsavel"] = $responsavel
        $ctx["Fornecedor"] = $fornecedor
        $ctx["Situacao"]   = $situacao
        $ctx["DataCompra"] = $dataCompra
        $ctx["Valor"]      = $valorCompra
        $ctx["Localizacao"] = $localizacao
        $ctx["Observacoes"] = $observacoes
        $ctx["NotaFiscal"] = $numeroNotaFiscal
        $ctx["CRM"]        = $crmSerial

        Write-Host ""
        $continuar = Read-Host " Cadastrar OUTRO bem agora? (S/N) [N]"
        if ($null -eq $continuar) { break }
        if ($continuar.Trim() -notmatch '(?i)^s|sim|y|yes$') { break }
    }

    Write-Host ""
    Write-Host " [OK] Sessao de cadastro encerrada." -ForegroundColor Green
    Pausar-Tela
}
# --------------------------------------------------------------------------
# EXIBICAO EM TABELA (adapta a largura do console)
# --------------------------------------------------------------------------
function Exibir-TabelaAtivos {
    param([array]$registros, [int]$limite = 0)

    if (!$registros -or $registros.Count -eq 0) {
        Write-Host "  Nenhum bem encontrado." -ForegroundColor Yellow
        return
    }

    if ($limite -le 0) { $limite = Get-LimiteListagem }

    $total = $registros.Count
    if ($total -gt $limite) {
        Write-Host "  Mostrando $limite de $total bens (use a busca ou os filtros para ver o resto)." -ForegroundColor DarkGray
        $registros = @($registros[0..($limite - 1)])
    }

    # Calcula as colunas conforme a largura real do terminal
    $disp = $script:larguraCaixa - 2
    $wNum = 4; $wTomb = 13; $wNomeMin = 18; $wCat = 17; $wSetor = 19; $wSit = 18; $wVal = 13
    $usarVal = $true; $usarCat = $true; $usarSet = $true
    while ($true) {
        $totalCol = $wNum + $wTomb + $wNomeMin + $wSit + 1
        if ($usarVal) { $totalCol += $wVal + 1 }
        if ($usarCat) { $totalCol += $wCat + 1 }
        if ($usarSet) { $totalCol += $wSetor + 1 }
        if ($totalCol -le $disp) { break }
        if ($usarVal) { $usarVal = $false; continue }
        if ($usarSet) { $usarSet = $false; continue }
        if ($usarCat) { $usarCat = $false; continue }
        break
    }
    $usado = $wNum + $wTomb + $wNomeMin + $wSit + 1
    if ($usarVal) { $usado += $wVal + 1 }
    if ($usarCat) { $usado += $wCat + 1 }
    if ($usarSet) { $usado += $wSetor + 1 }
    $wNome = [Math]::Max($wNomeMin, $disp - $usado)

    # Definicao das colunas (a ordem se ajusta ao espaco disponivel)
    $cols = @()
    $cols += @{ Cab = "#";         Larg = $wNum;  Tipo = "Num" }
    $cols += @{ Cab = "TOMBAMENTO"; Larg = $wTomb; Tipo = "Tomb" }
    $cols += @{ Cab = "NOME DO BEM"; Larg = $wNome; Tipo = "Nome" }
    if ($usarCat) { $cols += @{ Cab = "CATEGORIA"; Larg = $wCat; Tipo = "Cat" } }
    if ($usarSet) { $cols += @{ Cab = "SETOR"; Larg = $wSetor; Tipo = "Setor" } }
    $cols += @{ Cab = "SITUACAO"; Larg = $wSit; Tipo = "Sit" }
    if ($usarVal) { $cols += @{ Cab = "VALOR"; Larg = $wVal; Tipo = "Val" } }

    # Cabecalho
    $partesCab = @()
    foreach ($c in $cols) { $partesCab += ("{0,-$($c.Larg)}" -f $c.Cab) }
    $cabecalho = ($partesCab -join " ").TrimEnd()

    Write-Host ""
    Write-Host "  $cabecalho" -ForegroundColor Cyan
    Write-Host ("  " + ("-" * [Math]::Min($disp, $cabecalho.Length))) -ForegroundColor DarkGray

    $i = 1
    foreach ($r in $registros) {
        foreach ($c in $cols) {
            $valor = switch ($c.Tipo) {
                "Num"  { "$i" }
                "Tomb" { $r.Codigo_Tombamento }
                "Nome" { Cortar-Texto $r.Nome_Bem_Produto $c.Larg }
                "Cat"  { Cortar-Texto $r.Categoria_Bem $c.Larg }
                "Setor"{ Cortar-Texto $r.Setor_Atribuido $c.Larg }
                "Sit"  { Cortar-Texto $r.Situacao $c.Larg }
                "Val"  { "R$ " + (Formatar-Moeda (Converter-ValorNumerico $r.Valor_Compra)) }
            }
            $cor = switch ($c.Tipo) {
                "Num"  { "Cyan" }
                "Tomb" { "Green" }
                "Nome" { "White" }
                "Cat"  { "Cyan" }
                "Setor"{ "Green" }
                "Sit"  { Obter-CorSituacao $r.Situacao }
                "Val"  { "Yellow" }
            }
            Write-Host $valor.PadRight($c.Larg) -NoNewline -ForegroundColor $cor
            Write-Host " " -NoNewline
        }
        Write-Host ""
        $i++
    }
    Write-Host ("  " + ("-" * [Math]::Min($disp, $cabecalho.Length))) -ForegroundColor DarkGray
    Write-Host "  Total exibido: " -NoNewline -ForegroundColor Gray
    Write-Host "$($registros.Count)" -ForegroundColor Green
    if ($total -gt $registros.Count) {
        Write-Host "  Total na base: " -NoNewline -ForegroundColor Gray
        Write-Host "$total" -ForegroundColor Green
    }
}

# --------------------------------------------------------------------------
# CONSULTA DE ATIVOS
# --------------------------------------------------------------------------
function Menu-ConsultarAtivos {
    param([array]$registros)

    if (!$registros -or $registros.Count -eq 0) {
        try { Clear-Host } catch {}
        Exibir-LogoPrincipal
        Write-Host " |                 >>> CONSULTA DE ATIVOS PATRIMONIAIS <<<              |" -ForegroundColor Yellow
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  Nenhum bem patrimonial cadastrado na base ainda." -ForegroundColor Yellow
        Write-Host "  Use a opcao [1] do Módulo 1 para cadastrar o primeiro bem." -ForegroundColor Gray
        Pausar-Tela
        return
    }

    $emConsulta = $true
    while ($emConsulta) {
        $totalValor = 0.0
        foreach ($r in $registros) { $totalValor += (Converter-ValorNumerico $r.Valor_Compra) }

        try { Clear-Host } catch {}
        Exibir-LogoPrincipal
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host " |                 >>> CONSULTA DE ATIVOS PATRIMONIAIS <<<              |" -ForegroundColor Yellow
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host "  Bens catalogados: " -NoNewline -ForegroundColor Gray; Write-Host "$($registros.Count)" -ForegroundColor Green
        Write-Host "  Valor total do patrimonio: " -NoNewline -ForegroundColor Gray; Write-Host ("R$ " + (Formatar-Moeda $totalValor)) -ForegroundColor Yellow
        Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
        Write-Host "  [1] Listar todos os bens" -ForegroundColor White
        Write-Host "  [2] Filtrar por Setor" -ForegroundColor White
        Write-Host "  [3] Filtrar por Categoria" -ForegroundColor White
        Write-Host "  [4] Filtrar por Situacao" -ForegroundColor White
        Write-Host "  [5] Buscar por tombamento / nome / serial" -ForegroundColor White
        Write-Host "  [6] Ver ficha completa de um bem" -ForegroundColor Cyan
        Write-Host "  [0] Voltar ao Módulo 1" -ForegroundColor Gray
        Write-Host ""

        $opcConsulta = Ler-OpcaoOuEsc -prompt " Escolha a opcao" -valorPadrao "1"

        switch ($opcConsulta.Trim()) {
            "1" {
                try { Clear-Host } catch {}
                Exibir-LogoPrincipal
                Write-Host " |                >>> LISTA GERAL DE BENS PATRIMONIAIS <<<             |" -ForegroundColor Yellow
                Write-Host " +======================================================================+" -ForegroundColor Cyan
                Exibir-TabelaAtivos -registros $registros
                Pausar-Tela
            }

            { $_ -in @("2", "3", "4") } {
                $campo = switch ($opcConsulta.Trim()) {
                    "2" { "Setor_Atribuido" }
                    "3" { "Categoria_Bem" }
                    "4" { "Situacao" }
                }
                $rotulo = switch ($opcConsulta.Trim()) {
                    "2" { "Setor" }
                    "3" { "Categoria" }
                    "4" { "Situacao" }
                }
                $valores = @()
                foreach ($r in $registros) {
                    $v = $r.$campo
                    if ($v -and $v -ne "N/A" -and $valores -notcontains $v) { $valores += $v }
                }

                try { Clear-Host } catch {}
                Exibir-LogoPrincipal
                Write-Host " |                >>> FILTRAR ATIVOS POR $($rotulo.ToUpper()) <<<              |" -ForegroundColor Yellow
                Write-Host " +======================================================================+" -ForegroundColor Cyan
                $opc = Selecionar-Opcao -titulo "Selecione o $rotulo" -opcoes $valores
                if ($null -eq $opc) { Pausar-Tela; break }
                $encontrados = @($registros | Where-Object { $_.$campo -eq $opc })
                Write-Host ""
                Write-Host " Resultado: " -NoNewline -ForegroundColor Gray
                Write-Host "$($encontrados.Count)" -ForegroundColor Green
                Write-Host " bem(ns) em '$opc'" -ForegroundColor Cyan
                Exibir-TabelaAtivos -registros $encontrados
                Pausar-Tela
            }

            "5" {
                Write-Host ""
                $termo = Read-Host " Digite o termo (tombamento, nome do bem, tipo ou serial)"
                if ($null -eq $termo) { Pausar-Tela; break }
                if (![string]::IsNullOrWhiteSpace($termo)) {
                    $t = [regex]::Escape($termo.Trim())
                    $encontrados = @($registros | Where-Object {
                        $_.Codigo_Tombamento -match "(?i)$t" -or
                        $_.Nome_Bem_Produto -match "(?i)$t" -or
                        $_.CRM_Serial -match "(?i)$t" -or
                        $_.Tipo_Especifico -match "(?i)$t" -or
                        $_.Responsavel_Colaborador -match "(?i)$t"
                    })
                    Write-Host ""
                    Write-Host " Resultados: " -NoNewline -ForegroundColor Gray
                    Write-Host "$($encontrados.Count)" -ForegroundColor Green
                    Exibir-TabelaAtivos -registros $encontrados
                }
                Pausar-Tela
            }

            "6" {
                Write-Host ""
                $escolha = Read-Host " Numero do bem (1 a $($registros.Count)) ou codigo de tombamento"
                if ($null -eq $escolha) { Pausar-Tela; break }
                $selecionado = $null
                $num = 0
                if ([int]::TryParse($escolha.Trim(), [ref]$num) -and $num -ge 1 -and $num -le $registros.Count) {
                    $selecionado = $registros[$num - 1]
                } elseif (![string]::IsNullOrWhiteSpace($escolha)) {
                    $selecionado = $registros | Where-Object { $_.Codigo_Tombamento -match "(?i)$($escolha.Trim())" } | Select-Object -First 1
                }

                if ($selecionado) {
                    try { Clear-Host } catch {}
                    Exibir-LogoPrincipal
                    Write-Host " |                >>> FICHA PATRIMONIAL COMPLETA <<<                   |" -ForegroundColor Yellow
                    Write-Host " +======================================================================+" -ForegroundColor Cyan
                    Write-Host "   NOME DO BEM:            " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Nome_Bem_Produto)" -ForegroundColor White
                    Write-Host "   CODIGO / TOMBAMENTO:    " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Codigo_Tombamento)" -ForegroundColor Green
                    Write-Host "   NOTA FISCAL:            " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Numero_Nota_Fiscal)" -ForegroundColor Yellow
                    Write-Host "   CRM / SERIAL:           " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.CRM_Serial)" -ForegroundColor Yellow
                    Write-Host "   EMPRESA / UNIDADE:      " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Empresa_Unidade)" -ForegroundColor White
                    Write-Host "   CATEGORIA DO BEM:       " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Categoria_Bem)" -ForegroundColor Cyan
                    Write-Host "   TIPO / SUBTIPO:         " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Tipo_Especifico)" -ForegroundColor Cyan
                    Write-Host "   ATRIBUIDO A (SETOR):    " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Setor_Atribuido)" -ForegroundColor Green
                    Write-Host "   RESPONSAVEL:            " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Responsavel_Colaborador)" -ForegroundColor White
                    Write-Host "   FORNECEDOR:             " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Fornecedor)" -ForegroundColor White
                    Write-Host "   SITUACAO:               " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Situacao)" -ForegroundColor (Obter-CorSituacao $selecionado.Situacao)
                    Write-Host "   DATA DA COMPRA:         " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Data_Compra)" -ForegroundColor Yellow
                    Write-Host "   VALOR DA COMPRA:        " -NoNewline -ForegroundColor Gray; Write-Host ("R$ " + (Formatar-Moeda (Converter-ValorNumerico $selecionado.Valor_Compra))) -ForegroundColor Yellow
                    Write-Host "   LOCALIZACAO FISICA:     " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Localizacao_Fisica_Posto)" -ForegroundColor Green
                    Write-Host "   OBSERVACOES:            " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Observacoes_Adicionais)" -ForegroundColor White
                    Write-Host "   DATA DO REGISTRO:       " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Data_Registro)" -ForegroundColor DarkGray
                    Write-Host " +======================================================================+" -ForegroundColor Cyan
                } else {
                    Write-Host " [!] Bem nao encontrado." -ForegroundColor Red
                }
                Pausar-Tela
            }

            "0" { $emConsulta = $false }
            default {
                Write-Host " Opcao invalida." -ForegroundColor Red
                Start-Sleep -Milliseconds 700
            }
        }
    }
}

# --------------------------------------------------------------------------
# EDICAO DE ATIVO
# --------------------------------------------------------------------------
function Menu-EditarAtivo {
    param([array]$registros, [string]$caminhoCsv)

    if (!$registros -or $registros.Count -eq 0) {
        Write-Host " [!] Nenhum bem cadastrado para editar." -ForegroundColor Yellow
        Pausar-Tela
        return
    }

    Write-Host ""
    $escolha = Read-Host " Numero do bem (1 a $($registros.Count)) ou codigo de tombamento"
    if ($null -eq $escolha) { return }
    $selecionado = $null
    $num = 0
    if ([int]::TryParse($escolha.Trim(), [ref]$num) -and $num -ge 1 -and $num -le $registros.Count) {
        $selecionado = $registros[$num - 1]
    } elseif (![string]::IsNullOrWhiteSpace($escolha)) {
        $selecionado = $registros | Where-Object { $_.Codigo_Tombamento -match "(?i)$($escolha.Trim())" } | Select-Object -First 1
    }

    if (!$selecionado) {
        Write-Host " [!] Bem nao encontrado." -ForegroundColor Red
        Pausar-Tela
        return
    }

    try { Clear-Host } catch {}
    Exibir-LogoPrincipal
    Write-Host " |                >>> ATUALIZACAO DE BEM PATRIMONIAL <<<                 |" -ForegroundColor Yellow
    Write-Host " +======================================================================+" -ForegroundColor Cyan
    Write-Host "   BEM: " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Codigo_Tombamento) - $($selecionado.Nome_Bem_Produto)" -ForegroundColor White
    Write-Host ""
    Write-Host "   [1] Situacao            (Atual: $($selecionado.Situacao))" -ForegroundColor White
    Write-Host "   [2] Setor               (Atual: $($selecionado.Setor_Atribuido))" -ForegroundColor White
    Write-Host "   [3] Responsavel         (Atual: $($selecionado.Responsavel_Colaborador))" -ForegroundColor White
    Write-Host "   [4] Localizacao Fisica  (Atual: $($selecionado.Localizacao_Fisica_Posto))" -ForegroundColor White
    Write-Host "   [5] Observacoes         (Atual: $($selecionado.Observacoes_Adicionais))" -ForegroundColor White
    Write-Host "   [6] Nome do Bem         (Atual: $($selecionado.Nome_Bem_Produto))" -ForegroundColor White
    Write-Host "   [7] Fornecedor          (Atual: $($selecionado.Fornecedor))" -ForegroundColor White
    Write-Host "   [0] Voltar sem alterar" -ForegroundColor Gray
    Write-Host ""

    $opc = Ler-OpcaoOuEsc -prompt " Escolha o campo que deseja alterar" -valorPadrao "0"
    if ($opc.Trim() -eq "0") { return }

    $novoValor = $null
    switch ($opc.Trim()) {
        "1" {
            $novoValor = Selecionar-Opcao -titulo "Nova situacao" -opcoes $script:situacoesAtivosPadrao -padrao $selecionado.Situacao
        }
        "2" {
            $novoValor = Selecionar-Opcao -titulo "Novo setor" -opcoes $script:setoresAtivosPadrao -padrao $selecionado.Setor_Atribuido -PermiteNovo -rotuloNovo "Digitar outro setor"
            if ($novoValor -eq "__NOVO__") { $novoValor = Ler-Campo -rotulo "   Nome do setor" -Obrigatorio }
        }
        "3" {
            $novoValor = Ler-Campo -rotulo "   Novo responsavel" -padrao $selecionado.Responsavel_Colaborador
        }
        "4" {
            $novoValor = Ler-Campo -rotulo "   Nova localizacao fisica" -padrao $selecionado.Localizacao_Fisica_Posto
        }
        "5" {
            $novoValor = Ler-Campo -rotulo "   Novas observacoes" -padrao $selecionado.Observacoes_Adicionais
        }
        "6" {
            $novoValor = Ler-Campo -rotulo "   Novo nome do bem" -padrao $selecionado.Nome_Bem_Produto -Obrigatorio
        }
        "7" {
            $novoValor = Ler-Campo -rotulo "   Novo fornecedor" -padrao $selecionado.Fornecedor
        }
        default { $novoValor = $null }
    }

    if ($null -eq $novoValor -or $novoValor -eq "") {
        Write-Host "`n [!] Nenhuma alteracao realizada." -ForegroundColor Yellow
        Pausar-Tela
        return
    }

    $listaAtualizada = @()
    foreach ($r in $registros) {
        if ($r.Codigo_Tombamento -eq $selecionado.Codigo_Tombamento) {
            switch ($opc.Trim()) {
                "1" { $r.Situacao = $novoValor; $r.Status_Ativo = $novoValor }
                "2" { $r.Setor_Atribuido = $novoValor; $r.Setor = $novoValor }
                "3" { $r.Responsavel_Colaborador = $novoValor; $r.Usuario = $novoValor }
                "4" { $r.Localizacao_Fisica_Posto = $novoValor; $r.Localizacao = $novoValor }
                "5" { $r.Observacoes_Adicionais = $novoValor; $r.Observacoes = $novoValor }
                "6" { $r.Nome_Bem_Produto = $novoValor }
                "7" { $r.Fornecedor = $novoValor }
            }
        }
        $listaAtualizada += $r
    }

    Salvar-RegistrosCsv -caminho $caminhoCsv -registros $listaAtualizada -BaseAtivos
    Write-Host "`n [OK] Bem $($selecionado.Codigo_Tombamento) atualizado com sucesso!" -ForegroundColor Green
    Pausar-Tela
}

# --------------------------------------------------------------------------
# EXCLUSAO DE ATIVOS (com backup automatico)
# --------------------------------------------------------------------------
function Excluir-BensAtivos {
    param([array]$registros, [string]$caminhoCsv)

    if (!$registros -or $registros.Count -eq 0) {
        Write-Host " [!] Nenhum bem cadastrado para excluir." -ForegroundColor Yellow
        Pausar-Tela
        return
    }

    Write-Host ""
    Write-Host " Digite o numero do bem (1 a $($registros.Count)), o codigo de tombamento ou 'T' para listar:" -NoNewline -ForegroundColor Yellow
    $escolha = Read-Host "   Numero, codigo do tombamento ou T para listar"
    if ($null -eq $escolha) { return }

    $escolha = $escolha.Trim()
    if ($escolha -eq "" -or $escolha -eq "T") {
        Exibir-TabelaAtivos -registros $registros
        Write-Host ""
        $escolha = Read-Host " Numero ou codigo do bem a excluir"
        if ($null -eq $escolha) { return }
        $escolha = $escolha.Trim()
    }

    $alvos = @()
    $num = 0
    if ([int]::TryParse($escolha, [ref]$num) -and $num -ge 1 -and $num -le $registros.Count) {
        $alvos = @($registros[$num - 1])
    } elseif ($escolha -ne "") {
        $alvos = @($registros | Where-Object { $_.Codigo_Tombamento -match "(?i)$([regex]::Escape($escolha))" })
    }

    if ($alvos.Count -eq 0) {
        Write-Host " [!] Bem nao encontrado." -ForegroundColor Red
        Pausar-Tela
        return
    }

    Write-Host ""
    Write-Host " BEM(S) QUE SERAO EXCLUIDO(S):" -ForegroundColor Red
    Exibir-TabelaAtivos -registros $alvos
    Write-Host ""
    Write-Host " Os valores e o historico serao perdidos da base (backup automatico em Backups\)." -ForegroundColor Yellow
    $confirma = Read-Host " Digite EXCLUIR para confirmar"
    if ($null -eq $confirma -or $confirma.Trim().ToUpper() -ne "EXCLUIR") {
        Write-Host "`n [!] Exclusao cancelada." -ForegroundColor Yellow
        Pausar-Tela
        return
    }

    $codigos = @($alvos | ForEach-Object { $_.Codigo_Tombamento })
    $listaRestante = @()
    foreach ($r in $registros) {
        if ($codigos -notcontains $r.Codigo_Tombamento) { $listaRestante += $r }
    }

    Salvar-RegistrosCsv -caminho $caminhoCsv -registros $listaRestante -BaseAtivos

    Write-Host ""
    Write-Host " [OK] $($alvos.Count) bem(ns) excluido(s) com sucesso!" -ForegroundColor Green
    foreach ($c in $codigos) { Write-Host "      - $c" -ForegroundColor DarkGray }
    Pausar-Tela
}

# --------------------------------------------------------------------------
# TERMO DE RESPONSABILIDADE E CUSTODIA
# --------------------------------------------------------------------------
function Gerar-TermoResponsabilidade {
    param([array]$registros, [string]$caminhoCsv)

    if (!$registros -or $registros.Count -eq 0) {
        Write-Host " [!] Nenhum bem cadastrado para gerar termo." -ForegroundColor Yellow
        Pausar-Tela
        return
    }

    Write-Host ""
    $escolha = Read-Host " Numero do bem (1 a $($registros.Count)) ou codigo de tombamento"
    if ($null -eq $escolha) { return }
    $selecionado = $null
    $num = 0
    if ([int]::TryParse($escolha.Trim(), [ref]$num) -and $num -ge 1 -and $num -le $registros.Count) {
        $selecionado = $registros[$num - 1]
    } elseif (![string]::IsNullOrWhiteSpace($escolha)) {
        $selecionado = $registros | Where-Object { $_.Codigo_Tombamento -match "(?i)$($escolha.Trim())" } | Select-Object -First 1
    }

    if (!$selecionado) {
        Write-Host " [!] Bem nao encontrado." -ForegroundColor Red
        Pausar-Tela
        return
    }

    $dataTermo = Get-Date -Format "dd/MM/yyyy"
    $empresa = $selecionado.Empresa_Unidade
    $nomeArquivo = "Termo_$($selecionado.Codigo_Tombamento)_$($selecionado.Nome_Bem_Produto)".Replace(" ", "_")
    $nomeArquivo = ($nomeArquivo -replace '[^\w\-]', '_')
    $caminhoTxt = Join-Path $script:pastaTermos "$nomeArquivo.txt"
    $caminhoHtml = Join-Path $script:pastaTermos "$nomeArquivo.html"

    if (!(Test-Path $script:pastaTermos)) { try { New-Item -ItemType Directory -Path $script:pastaTermos -Force | Out-Null } catch {} }

    # ---- CONTEUDO TXT ----
    $d = "=" * 78
    $l = "-" * 78
    $txt = @(
        $d,
        "            TERMO DE RESPONSABILIDADE E CUSTODIA DE BEM PATRIMONIAL",
        "             INNOVCORE TI - GESTAO DE ATIVOS E PATRIMONIO (v2.0)",
        $d,
        "",
        " Documento gerado automaticamente em $dataTermo.",
        " Empresa / Unidade: $empresa",
        "",
        $l,
        " 1. IDENTIFICACAO DO BEM",
        $l,
        "    Nome do Bem:              $($selecionado.Nome_Bem_Produto)",
        "    Codigo / Tombamento:      $($selecionado.Codigo_Tombamento)",
        "    Categoria do Bem:         $($selecionado.Categoria_Bem)",
        "    Tipo / Subtipo:           $($selecionado.Tipo_Especifico)",
        "    Nota Fiscal:              $($selecionado.Numero_Nota_Fiscal)",
        "    CRM / Serial:             $($selecionado.CRM_Serial)",
        "    Valor de Aquisicao:       R$ $($selecionado.Valor_Compra)",
        "    Data de Aquisicao:        $($selecionado.Data_Compra)",
        "    Fornecedor:               $($selecionado.Fornecedor)",
        "",
        $l,
        " 2. RESPONSABILIDADE E CUSTODIA",
        $l,
        "    Setor de Alocacao:        $($selecionado.Setor_Atribuido)",
        "    Responsavel / Custodiante:$($selecionado.Responsavel_Colaborador)",
        "    Localizacao Fisica:       $($selecionado.Localizacao_Fisica_Posto)",
        "    Situacao do Bem:          $($selecionado.Situacao)",
        "",
        $l,
        " 3. COMPROMISSOS DO CUSTODIANTE",
        $l,
        "    Declaro ter recebido o bem acima discriminado, em bom estado de",
        "    conservacao e para uso exclusivo da empresa, comprometendo-me a:",
        "",
        "    a) zelar pela integridade e conservacao do bem sob minha custodia;",
        "    b) nao permitir o uso por terceiros sem autorizacao da gestao;",
        "    c) comunicar a gestao qualquer avaria, perda ou desvio de finalidade;",
        "    d) devolver o bem a gesto no momento da solicitacao ou do fim da minha",
        "       vinculo com a empresa;",
        "    e) responder por perdas, danos ou avarias causadas por uso indevido.",
        "",
        $l,
        " 4. OBSERVACOES",
        $l,
        "    $($selecionado.Observacoes_Adicionais)",
        "",
        "======================================================================",
        "                        ASSINATURAS",
        "======================================================================",
        "",
        "",
        "    _____________________________________________     ____________________",
        "    Entregue por (responsavel pela entrega)        Data",
        "",
        "",
        "    _____________________________________________     ____________________",
        "    Recebido por (custodiante)                     Data",
        "",
        "======================================================================",
        " Documento gerado pelo InnovCore TI v2.0 - $dataTermo",
        "======================================================================"
    ) -join "`r`n"

    [System.IO.File]::WriteAllText($caminhoTxt, $txt, $script:utf8ComBOM)

    # ---- CONTEUDO HTML (imprimivel) ----
    $html = @(
        '<!DOCTYPE html>',
        '<html lang="pt-BR"><head><meta charset="utf-8">',
        '<title>Termo de Responsabilidade - ' + $selecionado.Codigo_Tombamento + '</title>',
        '<style>',
        ' body { font-family: Arial, Helvetica, sans-serif; color:#1e293b; font-size:11pt; margin: 24px; }',
        ' h1 { font-size:15pt; text-align:center; margin:0 0 4px 0; }',
        ' h2 { font-size:11.5pt; margin:20px 0 6px 0; color:#0369a1; border-left:4px solid #0284c7; padding-left:8px; }',
        ' .sub { text-align:center; color:#64748b; font-size:9.5pt; margin-bottom:18px; }',
        ' table { border-collapse:collapse; width:100%; }',
        ' td { border:1px solid #cbd5e1; padding:6px 8px; font-size:10pt; vertical-align:top; }',
        ' td.label { background:#f1f5f9; font-weight:bold; width:38%; }',
        ' .assinatura { margin-top:42px; }',
        ' .linha { border-top:1px solid #334155; margin:0 30px; padding-top:3px; font-size:9pt; color:#475569; }',
        ' .obs { font-size:10pt; text-align:justify; }',
        ' .rodape { margin-top:26px; font-size:8.5pt; color:#94a3b8; text-align:center; border-top:1px solid #e2e8f0; padding-top:8px; }',
        ' @media print { body { margin: 10mm; } .noprint { display:none; } }',
        '</style></head><body>',
        '<h1>TERMO DE RESPONSABILIDADE E CUSTODIA DE BEM PATRIMONIAL</h1>',
        '<div class="sub">INNOVCORE TI - GESTAO DE ATIVOS E PATRIMONIO &nbsp;|&nbsp; ' + $empresa + ' &nbsp;|&nbsp; Emitido em ' + $dataTermo + '</div>',
        '<h2>1. Identificacao do Bem</h2>',
        '<table>',
        '<tr><td class="label">Nome do Bem</td><td>' + $selecionado.Nome_Bem_Produto + '</td></tr>',
        '<tr><td class="label">Codigo / Tombamento</td><td>' + $selecionado.Codigo_Tombamento + '</td></tr>',
        '<tr><td class="label">Categoria / Tipo</td><td>' + $selecionado.Categoria_Bem + ' / ' + $selecionado.Tipo_Especifico + '</td></tr>',
        '<tr><td class="label">Nota Fiscal</td><td>' + $selecionado.Numero_Nota_Fiscal + '</td></tr>',
        '<tr><td class="label">CRM / Serial</td><td>' + $selecionado.CRM_Serial + '</td></tr>',
        '<tr><td class="label">Valor de Aquisicao</td><td>R$ ' + $selecionado.Valor_Compra + '</td></tr>',
        '<tr><td class="label">Data de Aquisicao</td><td>' + $selecionado.Data_Compra + '</td></tr>',
        '<tr><td class="label">Fornecedor</td><td>' + $selecionado.Fornecedor + '</td></tr>',
        '</table>',
        '<h2>2. Responsabilidade e Custodia</h2>',
        '<table>',
        '<tr><td class="label">Setor de Alocacao</td><td>' + $selecionado.Setor_Atribuido + '</td></tr>',
        '<tr><td class="label">Responsavel / Custodiante</td><td>' + $selecionado.Responsavel_Colaborador + '</td></tr>',
        '<tr><td class="label">Localizacao Fisica</td><td>' + $selecionado.Localizacao_Fisica_Posto + '</td></tr>',
        '<tr><td class="label">Situacao do Bem</td><td>' + $selecionado.Situacao + '</td></tr>',
        '</table>',
        '<h2>3. Compromissos do Custodiante</h2>',
        '<p class="obs">Declaro ter recebido o bem acima discriminado, em bom estado de conservacao e para uso exclusivo da empresa, comprometendo-me a: (a) zelar pela integridade e conservacao do bem sob minha custodia; (b) nao permitir o uso por terceiros sem autorizacao da gestao; (c) comunicar a gestao qualquer avaria, perda ou desvio de finalidade; (d) devolver o bem a gesto no momento da solicitacao ou do fim do meu vinculo com a empresa; (e) responder por perdas, danos ou avarias causadas por uso indevido.</p>',
        '<h2>4. Observacoes</h2>',
        '<p class="obs">' + $selecionado.Observacoes_Adicionais + '</p>',
        '<div class="assinatura">',
        '<div class="linha">&nbsp;</div>Entregue por (responsavel pela entrega) &nbsp;&nbsp;&nbsp; Data: ____/____/______',
        '<br><br>',
        '<div class="linha">&nbsp;</div>Recebido por (custodiante) &nbsp;&nbsp;&nbsp; Data: ____/____/______',
        '</div>',
        '<div class="rodape">Documento gerado automaticamente pelo InnovCore TI v2.0 em ' + $dataTermo + '.</div>',
        '</body></html>'
    ) -join "`r`n"

    [System.IO.File]::WriteAllText($caminhoHtml, $html, $script:utf8ComBOM)

    # ---- REGISTRO DE TERMOS EMITIDOS ----
    $linhaRegistro = "{0};{1};{2};{3};{4}" -f $selecionado.Codigo_Tombamento, $selecionado.Responsavel_Colaborador, $dataTermo, $selecionado.Empresa_Unidade, $nomeArquivo
    $cabecalhoTermos = "Codigo_Tombamento;Responsavel;Data_Termo;Empresa;Arquivo"
    $conteudoTermos = $cabecalhoTermos
    if (Test-Path $script:caminhoTermosEmitidos) {
        try { $conteudoTermos = [System.IO.File]::ReadAllText($script:caminhoTermosEmitidos, [System.Text.Encoding]::UTF8) } catch {}
    }
    if ($conteudoTermos -notmatch [regex]::Escape($selecionado.Codigo_Tombamento)) {
        [System.IO.File]::WriteAllText($script:caminhoTermosEmitidos, ($conteudoTermos + "`r`n" + $linhaRegistro + "`r`n"), $script:utf8ComBOM)
    }

    try { Clear-Host } catch {}
    Exibir-LogoPrincipal
    Write-Host " +======================================================================+" -ForegroundColor Cyan
    Write-Host " |              TERMO DE RESPONSABILIDADE GERADO COM SUCESSO            |" -ForegroundColor Yellow
    Write-Host " +======================================================================+" -ForegroundColor Cyan
    Write-Host "   Bem:          " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Codigo_Tombamento) - $($selecionado.Nome_Bem_Produto)" -ForegroundColor White
    Write-Host "   Responsavel:  " -NoNewline -ForegroundColor Gray; Write-Host "$($selecionado.Responsavel_Colaborador)" -ForegroundColor White
    Write-Host "   Data:         " -NoNewline -ForegroundColor Gray; Write-Host "$dataTermo" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   Arquivos gerados em: " -ForegroundColor Gray
    Write-Host "     $caminhoTxt" -ForegroundColor Cyan
    Write-Host "     $caminhoHtml" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "   [1] Abrir o termo em HTML (formatado, pronto para imprimir)" -ForegroundColor Green
    Write-Host "   [2] Abrir o termo em TXT" -ForegroundColor Cyan
    Write-Host "   [0] Voltar ao Módulo 1" -ForegroundColor Gray
    Write-Host ""
    $abrir = Read-Host " Escolha [1]"
    if ($null -ne $abrir) {
        if ($abrir.Trim() -eq "1") { try { Invoke-Item $caminhoHtml } catch {} }
        elseif ($abrir.Trim() -eq "2") { try { Invoke-Item $caminhoTxt } catch {} }
    }
    Pausar-Tela
}
# --------------------------------------------------------------------------
# IMPORTACAO DE PLANILHA (CADASTRO EM LOTE VIA CSV)
# --------------------------------------------------------------------------
function Gerar-ModeloImportacao {
    param([string]$caminho)
    $pasta = [System.IO.Path]::GetDirectoryName($caminho)
    if ($pasta -and !(Test-Path $pasta)) { try { New-Item -ItemType Directory -Path $pasta -Force | Out-Null } catch {} }
    $exemplo1 = "19/10/2026;InnovEdu;PAT-0001;NF-129480;N/A;Moveis e Escritorio;Cadeira;Cadeira Presidente;Operacoes e Logistica;Fulano de Tal;FOR-0001;Ativo/Operacional;19/10/2026;750,00;Sala 101 - Posto 01;EXEMPLO - apague esta linha;N/A;N/A;N/A;N/A;N/A;N/A;N/A;N/A;N/A;N/A"
    $exemplo2 = "19/10/2026;InnovEdu;PAT-0002;NF-129480;N/A;Moveis e Escritorio;Cadeira;Cadeira Presidente;Operacoes e Logistica;Fulano de Tal;FOR-0001;Ativo/Operacional;19/10/2026;750,00;Sala 101 - Posto 02;EXEMPLO - apague esta linha;N/A;N/A;N/A;N/A;N/A;N/A;N/A;N/A;N/A;N/A"
    [System.IO.File]::WriteAllText($caminho, ($script:cabecalhoCsvString + "`r`n" + $exemplo1 + "`r`n" + $exemplo2 + "`r`n"), $script:utf8ComBOM)
    return $caminho
}

function Importar-AtivosPlanilha {
    param([array]$registrosAtuais, [string]$caminhoCsv)

    Write-Host ""
    Write-Host " Digite o caminho da planilha CSV a importar (ENTER = usar o modelo oficial):" -NoNewline -ForegroundColor Yellow
    $entrada = Read-Host "   Caminho da planilha CSV"
    if ($null -eq $entrada) { return }

    $caminhoImport = $entrada.Trim()
    if ($caminhoImport -eq "") {
        $caminhoImport = $script:caminhoModeloImportacao
        if (!(Test-Path $caminhoImport)) {
            $caminhoImport = Gerar-ModeloImportacao -caminho $script:caminhoModeloImportacao
            Write-Host " [OK] Modelo criado em: $caminhoImport" -ForegroundColor Green
        }
    }
    if (!(Test-Path $caminhoImport)) {
        Write-Host " [!] Arquivo nao encontrado: $caminhoImport" -ForegroundColor Red
        Pausar-Tela
        return
    }

    # ---- LEITURA TOLERANTE (mapeia por nome de coluna) ----
    $novos = @()
    $ignorados = @()
    $tombamentosUsados = @()
    foreach ($r in $registrosAtuais) { $tombamentosUsados += $r.Codigo_Tombamento }
    $seriaisUsados = @()
    foreach ($r in $registrosAtuais) { if ($r.CRM_Serial -and $r.CRM_Serial -ne "N/A") { $seriaisUsados += $r.CRM_Serial } }

    try {
        $raw = [System.IO.File]::ReadAllText($caminhoImport, [System.Text.Encoding]::UTF8)
        $linhas = $raw -split "\r?\n" | Where-Object { ![string]::IsNullOrWhiteSpace($_) }
        if ($linhas.Count -lt 2) {
            Write-Host " [!] A planilha nao possui linhas de dados." -ForegroundColor Red
            Pausar-Tela
            return
        }

        $cabecalho = (($linhas[0] -replace "^\uFEFF", "") -split ';')
        for ($i = 1; $i -lt $linhas.Count; $i++) {
            $v = $linhas[$i] -split ';'
            $linhaNum = $i + 1
            $campos = @{}
            for ($k = 0; $k -lt $cabecalho.Count; $k++) {
                $nome = $cabecalho[$k].Trim()
                $campos[$nome] = $(if ($k -lt $v.Count) { $v[$k].Trim() } else { "" })
            }

            $nomeBem = $campos["Nome_Bem_Produto"]
            if ([string]::IsNullOrWhiteSpace($nomeBem)) {
                $ignorados += "Linha $linhaNum - Nome do Bem vazio"
                continue
            }
            if ($nomeBem -match '(?i)^exemplo' -or $campos["Observacoes_Adicionais"] -match '(?i)^exemplo') {
                $ignorados += "Linha $linhaNum - linha de exemplo"
                continue
            }

            $tombamento = $campos["Codigo_Tombamento"]
            if ([string]::IsNullOrWhiteSpace($tombamento)) {
                $tombamento = Obter-ProximoPatrimonio -registros (@($registrosAtuais) + $novos)
            }
            if ($tombamentosUsados -contains $tombamento) {
                $ignorados += "Linha $linhaNum - tombamento $tombamento ja cadastrado"
                continue
            }
            $serial = $campos["CRM_Serial"]
            if ($serial -and $serial -ne "N/A" -and $seriaisUsados -contains $serial) {
                $ignorados += "Linha $linhaNum - serial $serial ja cadastrado"
                continue
            }

            $tombamentosUsados += $tombamento
            if ($serial -and $serial -ne "N/A") { $seriaisUsados += $serial }

            $novos += [PSCustomObject]@{
                Data_Registro              = (Get-Date -Format "dd/MM/yyyy HH:mm:ss")
                Empresa_Unidade           = $(if ($campos["Empresa_Unidade"]) { $campos["Empresa_Unidade"] } else { "InnovEdu" })
                Codigo_Tombamento          = $tombamento
                Numero_Nota_Fiscal         = $(if ($campos["Numero_Nota_Fiscal"]) { $campos["Numero_Nota_Fiscal"] } else { "N/A" })
                CRM_Serial                 = $(if ($campos["CRM_Serial"]) { $campos["CRM_Serial"] } else { "N/A" })
                Categoria_Bem              = $(if ($campos["Categoria_Bem"]) { $campos["Categoria_Bem"] } else { "Outros" })
                Tipo_Especifico            = $(if ($campos["Tipo_Especifico"]) { $campos["Tipo_Especifico"] } else { "Generico" })
                Nome_Bem_Produto           = $nomeBem
                Setor_Atribuido            = $(if ($campos["Setor_Atribuido"]) { $campos["Setor_Atribuido"] } else { "Operacoes e Logistica" })
                Responsavel_Colaborador    = $(if ($campos["Responsavel_Colaborador"]) { $campos["Responsavel_Colaborador"] } else { "N/A" })
                Fornecedor                 = $(if ($campos["Fornecedor"]) { $campos["Fornecedor"] } else { "N/A" })
                Situacao                   = $(if ($campos["Situacao"]) { $campos["Situacao"] } else { "Ativo/Operacional" })
                Data_Compra                = $(if ($campos["Data_Compra"]) { $campos["Data_Compra"] } else { (Get-Date -Format "dd/MM/yyyy") })
                Valor_Compra               = (Normalizar-ValorMoeda $campos["Valor_Compra"])
                Localizacao_Fisica_Posto   = $(if ($campos["Localizacao_Fisica_Posto"]) { $campos["Localizacao_Fisica_Posto"] } else { "N/A" })
                Observacoes_Adicionais     = $(if ($campos["Observacoes_Adicionais"]) { $campos["Observacoes_Adicionais"] } else { "N/A" })
                Nome_Computador            = "N/A"
                Processador                = "N/A"
                Memoria_RAM                = "N/A"
                Armazenamento              = "N/A"
                Placa_Video                = "N/A"
                Sistema_Operacional        = "N/A"
                IP_Rede                    = "N/A"
                MAC_Rede                   = "N/A"
                Bateria                    = "N/A"
                Android_Versao             = "N/A"
                Empresa                    = $(if ($campos["Empresa_Unidade"]) { $campos["Empresa_Unidade"] } else { "InnovEdu" })
                Patrimonio                 = $tombamento
                Usuario                    = $(if ($campos["Responsavel_Colaborador"]) { $campos["Responsavel_Colaborador"] } else { "N/A" })
                Setor                      = $(if ($campos["Setor_Atribuido"]) { $campos["Setor_Atribuido"] } else { "Operacoes e Logistica" })
                Status_Ativo               = $(if ($campos["Situacao"]) { $campos["Situacao"] } else { "Ativo/Operacional" })
                Localizacao                = $(if ($campos["Localizacao_Fisica_Posto"]) { $campos["Localizacao_Fisica_Posto"] } else { "N/A" })
                Observacoes                = $(if ($campos["Observacoes_Adicionais"]) { $campos["Observacoes_Adicionais"] } else { "N/A" })
            }
        }
    } catch {
        Write-Host " [!] Erro ao ler a planilha: $($_.Exception.Message)" -ForegroundColor Red
        Pausar-Tela
        return
    }

    Write-Host ""
    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
    Write-Host " |                  PREVIA DA IMPORTACAO                               |" -ForegroundColor Yellow
    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
    Write-Host "   Arquivo:       " -NoNewline -ForegroundColor Gray; Write-Host "$caminhoImport" -ForegroundColor Yellow
    Write-Host "   A importar:    " -NoNewline -ForegroundColor Gray; Write-Host "$($novos.Count)" -ForegroundColor Green
    Write-Host "   Ignorados:     " -NoNewline -ForegroundColor Gray; Write-Host "$($ignorados.Count)" -ForegroundColor Yellow
    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan

    if ($novos.Count -gt 0) {
        Write-Host ""
        Write-Host "   Amostra do que sera cadastrado:" -ForegroundColor Cyan
        Exibir-TabelaAtivos -registros $novos -limite 8
    }

    if ($ignorados.Count -gt 0) {
        Write-Host ""
        Write-Host "   Linhas ignoradas:" -ForegroundColor Yellow
        foreach ($i in ($ignorados | Select-Object -First 10)) { Write-Host "     - $i" -ForegroundColor DarkGray }
        if ($ignorados.Count -gt 10) { Write-Host "     ... e mais $($ignorados.Count - 10)" -ForegroundColor DarkGray }
    }

    Write-Host ""
    if ($novos.Count -eq 0) {
        Write-Host " [!] Nada a importar." -ForegroundColor Red
        Pausar-Tela
        return
    }

    $confirma = Read-Host " Confirmar a importacao de $($novos.Count) bem(ns)? (S/N)"
    if ($null -eq $confirma -or $confirma.Trim() -notmatch '(?i)^s|sim|y|yes$') {
        Write-Host "`n [!] Importacao cancelada." -ForegroundColor Yellow
        Pausar-Tela
        return
    }

    $listaFinal = @($registrosAtuais) + $novos
    Salvar-RegistrosCsv -caminho $caminhoCsv -registros $listaFinal -BaseAtivos

    Write-Host ""
    Write-Host " [OK] Importacao concluida! $($novos.Count) bem(ns) cadastrados." -ForegroundColor Green
    Write-Host "      Arquivo: $caminhoCsv" -ForegroundColor Yellow
    Pausar-Tela
}

# --------------------------------------------------------------------------
# DASHBOARD DO PATRIMONIO
# --------------------------------------------------------------------------
function Exibir-Barras {
    param([array]$itens, [string]$titulo, [int]$top = 10)
    if (!$itens -or $itens.Count -eq 0) { return }
    Write-Host ""
    Write-Host "  $titulo" -ForegroundColor Yellow
    $maior = 0
    foreach ($i in $itens) { if ($i.Quantidade -gt $maior) { $maior = $i.Quantidade } }
    if ($maior -eq 0) { return }
    $exibidos = if ($itens.Count -gt $top) { @($itens[0..($top - 1)]) } else { $itens }
    foreach ($i in $exibidos) {
        $tamanho = [int][Math]::Round(($i.Quantidade / $maior) * 30)
        if ($tamanho -lt 1) { $tamanho = 1 }
        $barra = "#" * $tamanho
        Write-Host "   " -NoNewline
        Write-Host "$(Cortar-Texto $i.Rotulo 26)".PadRight(26) -NoNewline -ForegroundColor White
        Write-Host " " -NoNewline
        Write-Host $barra.PadRight(30) -NoNewline -ForegroundColor Cyan
        Write-Host " " -NoNewline
        Write-Host "$($i.Quantidade)" -NoNewline -ForegroundColor Green
        if ($i.Valor -gt 0) {
            Write-Host "  (R$ " -NoNewline -ForegroundColor DarkGray
            Write-Host (Formatar-Moeda $i.Valor) -NoNewline -ForegroundColor Yellow
            Write-Host ")" -ForegroundColor DarkGray
        } else {
            Write-Host ""
        }
    }
}

function Menu-DashboardAtivos {
    param([array]$registros)

    if (!$registros -or $registros.Count -eq 0) {
        Write-Host " [!] Nenhum bem cadastrado para compor o dashboard." -ForegroundColor Yellow
        Pausar-Tela
        return
    }

    $totalValor = 0.0
    $porSituacao = @()
    $porCategoria = @()
    $porSetor = @()

    foreach ($r in $registros) {
        $valor = Converter-ValorNumerico $r.Valor_Compra
        $totalValor += $valor

        $s = $r.Situacao
        $itemS = $porSituacao | Where-Object { $_.Rotulo -eq $s } | Select-Object -First 1
        if (!$itemS) { $itemS = [PSCustomObject]@{ Rotulo = $s; Quantidade = 0; Valor = 0.0 }; $porSituacao += $itemS }
        $itemS.Quantidade++; $itemS.Valor += $valor

        $c = $r.Categoria_Bem
        $itemC = $porCategoria | Where-Object { $_.Rotulo -eq $c } | Select-Object -First 1
        if (!$itemC) { $itemC = [PSCustomObject]@{ Rotulo = $c; Quantidade = 0; Valor = 0.0 }; $porCategoria += $itemC }
        $itemC.Quantidade++; $itemC.Valor += $valor

        $set = $r.Setor_Atribuido
        $itemSet = $porSetor | Where-Object { $_.Rotulo -eq $set } | Select-Object -First 1
        if (!$itemSet) { $itemSet = [PSCustomObject]@{ Rotulo = $set; Quantidade = 0; Valor = 0.0 }; $porSetor += $itemSet }
        $itemSet.Quantidade++; $itemSet.Valor += $valor
    }

    $porSituacao = @($porSituacao | Sort-Object -Property Quantidade -Descending)
    $porCategoria = @($porCategoria | Sort-Object -Property Valor -Descending)
    $porSetor = @($porSetor | Sort-Object -Property Valor -Descending)

    try { Clear-Host } catch {}
    Exibir-LogoPrincipal
    Write-Host " +======================================================================+" -ForegroundColor Cyan
    Write-Host " |                 >>> DASHBOARD DO PATRIMONIO <<<                       |" -ForegroundColor Yellow
    Write-Host " +======================================================================+" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "   TOTAL DE BENS:        " -NoNewline -ForegroundColor Gray; Write-Host "$($registros.Count)" -ForegroundColor Green
    Write-Host "   VALOR TOTAL:          " -NoNewline -ForegroundColor Gray; Write-Host ("R$ " + (Formatar-Moeda $totalValor)) -ForegroundColor Yellow
    Write-Host "   VALOR MEDIO POR BEM:  " -NoNewline -ForegroundColor Gray; Write-Host ("R$ " + (Formatar-Moeda ($totalValor / $registros.Count))) -ForegroundColor Yellow
    Write-Host "   EM OPERACAO:          " -NoNewline -ForegroundColor Gray
    $ativos = ($porSituacao | Where-Object { $_.Rotulo -match 'Ativo|Operacional' } | Select-Object -First 1)
    Write-Host "$(if ($ativos) { $ativos.Quantidade } else { 0 })" -ForegroundColor Green
    Write-Host "   TOMBAMENTO SEGUINTE:  " -NoNewline -ForegroundColor Gray; Write-Host "$(Obter-ProximoPatrimonio -registros $registros)" -ForegroundColor Cyan

    Exibir-Barras -itens $porSituacao -titulo "BENS POR SITUACAO" -top 10
    Exibir-Barras -itens $porCategoria -titulo "BENS POR CATEGORIA (maior valor)" -top 10
    Exibir-Barras -itens $porSetor -titulo "BENS POR SETOR (maior valor)" -top 10

    Write-Host ""
    Write-Host "  MAIORES VALORES" -ForegroundColor Yellow
    $top5 = @($registros | Sort-Object { -[double]((Converter-ValorNumerico $_.Valor_Compra) -replace ",", ".") } | Select-Object -First 5)
    foreach ($r in $top5) {
        Write-Host "   " -NoNewline
        Write-Host "$($r.Codigo_Tombamento)" -NoNewline -ForegroundColor Green
        Write-Host " " -NoNewline
        Write-Host (Cortar-Texto $r.Nome_Bem_Produto 38).PadRight(38) -NoNewline -ForegroundColor White
        Write-Host ("R$ " + (Formatar-Moeda (Converter-ValorNumerico $r.Valor_Compra))) -ForegroundColor Yellow
    }

    Write-Host " +======================================================================+" -ForegroundColor Cyan
    Pausar-Tela
}
function Exportar-AtivosPatrimoniais {
    param(
        [array]$registros,
        [string]$tipo = "CSV_EXCEL",
        [string]$filtroSetor = "",
        [string]$filtroCategoria = "",
        [string]$filtroSituacao = "",
        [switch]$abrirAposExportar
    )

    if (!(Test-Path $script:pastaExport)) {
        New-Item -ItemType Directory -Path $script:pastaExport -Force | Out-Null
    }

    $registrosFiltrados = @($registros)
    if (![string]::IsNullOrWhiteSpace($filtroSetor)) {
        $registrosFiltrados = @($registrosFiltrados | Where-Object { $_.Setor_Atribuido -eq $filtroSetor })
    }
    if (![string]::IsNullOrWhiteSpace($filtroCategoria)) {
        $registrosFiltrados = @($registrosFiltrados | Where-Object { $_.Categoria_Bem -eq $filtroCategoria })
    }
    if (![string]::IsNullOrWhiteSpace($filtroSituacao)) {
        $registrosFiltrados = @($registrosFiltrados | Where-Object { $_.Situacao -eq $filtroSituacao })
    }

    if ($registrosFiltrados.Count -eq 0) {
        Write-Host " [!] Nenhum bem encontrado para exportacao com os criterios definidos." -ForegroundColor Yellow
        return $null
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $sufixos = @()
    if ($filtroSetor) { $sufixos += ($filtroSetor -replace '[^\w]', '_') }
    if ($filtroCategoria) { $sufixos += ($filtroCategoria -replace '[^\w]', '_') }
    if ($filtroSituacao) { $sufixos += ($filtroSituacao -replace '[^\w]', '_') }
    $sufixoFiltro = if ($sufixos.Count -gt 0) { "_$($sufixos -join '_')" } else { "_Geral" }

    $limparCampo = { param($c) if ($null -ne $c) { ($c.ToString() -replace ';', ',') -replace "[\r\n]+", " " } else { "" } }

    switch ($tipo) {
        "CSV_EXCEL" {
            $nomeArquivo = "Patrimonio_InnovCore$($sufixoFiltro)_$($timestamp).csv"
            $caminhoArquivo = Join-Path $script:pastaExport $nomeArquivo

            $linhasExport = [System.Collections.Generic.List[string]]::new()
            $linhasExport.Add($script:cabecalhoCsvString)
            foreach ($r in $registrosFiltrados) {
                $linha = @(
                    (& $limparCampo $r.Data_Registro),
                    (& $limparCampo $r.Empresa_Unidade),
                    (& $limparCampo $r.Codigo_Tombamento),
                    (& $limparCampo $r.Numero_Nota_Fiscal),
                    (& $limparCampo $r.CRM_Serial),
                    (& $limparCampo $r.Categoria_Bem),
                    (& $limparCampo $r.Tipo_Especifico),
                    (& $limparCampo $r.Nome_Bem_Produto),
                    (& $limparCampo $r.Setor_Atribuido),
                    (& $limparCampo $r.Responsavel_Colaborador),
                    (& $limparCampo $r.Fornecedor),
                    (& $limparCampo $r.Situacao),
                    (& $limparCampo $r.Data_Compra),
                    (& $limparCampo $r.Valor_Compra),
                    (& $limparCampo $r.Localizacao_Fisica_Posto),
                    (& $limparCampo $r.Observacoes_Adicionais),
                    (& $limparCampo $r.Nome_Computador),
                    (& $limparCampo $r.Processador),
                    (& $limparCampo $r.Memoria_RAM),
                    (& $limparCampo $r.Armazenamento),
                    (& $limparCampo $r.Placa_Video),
                    (& $limparCampo $r.Sistema_Operacional),
                    (& $limparCampo $r.IP_Rede),
                    (& $limparCampo $r.MAC_Rede),
                    (& $limparCampo $r.Bateria),
                    (& $limparCampo $r.Android_Versao)
                ) -join ';'
                $linhasExport.Add($linha)
            }

            [System.IO.File]::WriteAllText($caminhoArquivo, (($linhasExport -join "`r`n") + "`r`n"), $script:utf8ComBOM)

            Write-Host "`n [OK] Planilha patrimonial Excel CSV UTF-8 gerada com sucesso!" -ForegroundColor Green
            Write-Host "      Arquivo: " -NoNewline -ForegroundColor Gray
            Write-Host "$caminhoArquivo" -ForegroundColor Yellow
            Write-Host "      Total de bens exportados: " -NoNewline -ForegroundColor Gray
            Write-Host "$($registrosFiltrados.Count)" -ForegroundColor Green

            if ($abrirAposExportar) { try { Invoke-Item $caminhoArquivo } catch {} }
            return $caminhoArquivo
        }

        "HTML_EXCEL" {
            $nomeArquivo = "Patrimonio_InnovCore_Formatado$($sufixoFiltro)_$($timestamp).xls"
            $caminhoArquivo = Join-Path $script:pastaExport $nomeArquivo

            $dataGeracao = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
            $totalAtivos = $registrosFiltrados.Count

            $linhasHtml = ""
            $num = 1
            foreach ($r in $registrosFiltrados) {
                $bgClass = if ($num % 2 -eq 0) { "even" } else { "odd" }
                $linhasHtml += "                <tr class=""$bgClass"">`r`n"
                $linhasHtml += "                    <td align=""center""><b>$num</b></td>`r`n"
                $linhasHtml += "                    <td class=""pat"">$($r.Codigo_Tombamento)</td>`r`n"
                $linhasHtml += "                    <td><b>$($r.Nome_Bem_Produto)</b></td>`r`n"
                $linhasHtml += "                    <td>$($r.Categoria_Bem)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Tipo_Especifico)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Setor_Atribuido)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Responsavel_Colaborador)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Fornecedor)</td>`r`n"
                $linhasHtml += "                    <td class=""serial"">$($r.Numero_Nota_Fiscal)</td>`r`n"
                $linhasHtml += "                    <td>$($r.CRM_Serial)</td>`r`n"
                $linhasHtml += "                    <td align=""center"">$($r.Data_Compra)</td>`r`n"
                $linhasHtml += "                    <td align=""right"">R$ $($r.Valor_Compra)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Situacao)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Localizacao_Fisica_Posto)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Observacoes_Adicionais)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Empresa_Unidade)</td>`r`n"
                $linhasHtml += "                </tr>`r`n"
                $num++
            }

            $escopoTxt = if ($filtroSetor -or $filtroCategoria -or $filtroSituacao) { "Filtros aplicados" } else { "Patrimonio Completo" }
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
                "        <h1 class=""header-title"">INNOVCORE TI - REGISTRO DE ATIVOS E PATRIMONIO</h1>",
                "        <div class=""header-subtitle"">Gerado em: $dataGeracao | Escopo: $escopoTxt</div>",
                '    </div>',
                '    <div class="stats-bar">',
                "        TOTAL DE BENS CATALOGADOS: $totalAtivos | SISTEMA: INNOVCORE TI v2.0",
                '    </div>',
                '    <table>',
                '        <thead>',
                '            <tr>',
                '                <th>#</th>',
                '                <th>Tombamento</th>',
                '                <th>Nome do Bem</th>',
                '                <th>Categoria</th>',
                '                <th>Tipo</th>',
                '                <th>Setor</th>',
                '                <th>Responsavel</th>',
                '                <th>Fornecedor</th>',
                '                <th>Nota Fiscal</th>',
                '                <th>Serial</th>',
                '                <th>Data Compra</th>',
                '                <th>Valor</th>',
                '                <th>Situacao</th>',
                '                <th>Localizacao</th>',
                '                <th>Observacoes</th>',
                '                <th>Empresa</th>',
                '            </tr>',
                '        </thead>',
                '        <tbody>',
                $linhasHtml,
                '        </tbody>',
                '    </table>',
                '    <div class="footer">',
                '        Relatorio gerado automaticamente por <b>InnovCore TI v2.0</b> - Gestao de Ativos e Patrimonio.',
                '    </div>',
                '</body>',
                '</html>'
            ) -join "`r`n"
            [System.IO.File]::WriteAllText($caminhoArquivo, $conteudoHtml, $script:utf8ComBOM)

            Write-Host "`n [OK] Planilha patrimonial formatada (.XLS/HTML) gerada com sucesso!" -ForegroundColor Green
            Write-Host "      Arquivo: " -NoNewline -ForegroundColor Gray
            Write-Host "$caminhoArquivo" -ForegroundColor Yellow

            if ($abrirAposExportar) { try { Invoke-Item $caminhoArquivo } catch {} }
            return $caminhoArquivo
        }

        "TXT_RELATORIO" {
            $nomeArquivo = "Patrimonio_Executivo_InnovCore$($sufixoFiltro)_$($timestamp).txt"
            $caminhoArquivo = Join-Path $script:pastaExport $nomeArquivo

            $dataGeracao = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
            $divisoria = "=" * 78
            $subdivisoria = "-" * 78
            $totalQtd = $registrosFiltrados.Count

            $txt = @(
                $divisoria,
                "        INNOVCORE TI - RELATORIO DE REGISTRO DE ATIVOS E PATRIMONIO",
                $divisoria,
                " Data de Emissao: $dataGeracao",
                " Total de Bens:   $totalQtd",
                $divisoria,
                ""
            ) -join "`r`n"

            $i = 1
            foreach ($r in $registrosFiltrados) {
                $txt += "[$i] TOMBAMENTO: $($r.Codigo_Tombamento) | BEM: $($r.Nome_Bem_Produto)`r`n"
                $txt += "$subdivisoria`r`n"
                $txt += "  - Empresa / Unidade:      $($r.Empresa_Unidade)`r`n"
                $txt += "  - Categoria do Bem:       $($r.Categoria_Bem)`r`n"
                $txt += "  - Tipo / Subtipo:         $($r.Tipo_Especifico)`r`n"
                $txt += "  - Atribuido a (Setor):    $($r.Setor_Atribuido)`r`n"
                $txt += "  - Responsavel:            $($r.Responsavel_Colaborador)`r`n"
                $txt += "  - Fornecedor:             $($r.Fornecedor)`r`n"
                $txt += "  - Nota Fiscal:            $($r.Numero_Nota_Fiscal)`r`n"
                $txt += "  - CRM / Serial:           $($r.CRM_Serial)`r`n"
                $txt += "  - Situacao:               $($r.Situacao)`r`n"
                $txt += "  - Data da Compra:         $($r.Data_Compra)`r`n"
                $txt += "  - Valor da Compra:        R$ $($r.Valor_Compra)`r`n"
                $txt += "  - Localizacao Fisica:     $($r.Localizacao_Fisica_Posto)`r`n"
                $txt += "  - Observacoes:            $($r.Observacoes_Adicionais)`r`n"
                $txt += "  - Registrado em:          $($r.Data_Registro)`r`n`r`n"
                $i++
            }

            $txt += "$divisoria`r`n INNOVCORE TI - Desenvolvido por Jordan para InnovTech`r`n$divisoria`r`n"
            [System.IO.File]::WriteAllText($caminhoArquivo, $txt, $script:utf8ComBOM)

            Write-Host "`n [OK] Relatorio executivo patrimonial gerado com sucesso!" -ForegroundColor Green
            Write-Host "      Arquivo: " -NoNewline -ForegroundColor Gray
            Write-Host "$caminhoArquivo" -ForegroundColor Yellow

            if ($abrirAposExportar) { try { Invoke-Item $caminhoArquivo } catch {} }
            return $caminhoArquivo
        }
    }
}
# ==============================================================================
# MENU DO MODULO 1 - REGISTRO DE ATIVOS E PATRIMONIO
# ==============================================================================
function Menu-ModuloAtivos {
    param([string]$caminhoCsvAtivos)

    $emModulo = $true
    while ($emModulo) {
        Atualizar-TamanhoTerminal
        $ativos = Obter-RegistrosCsv -caminho $caminhoCsvAtivos -empresaPadrao "InnovEdu" -setorPadrao "Operacoes e Logistica"
        $totalValor = 0.0
        foreach ($r in $ativos) { $totalValor += (Converter-ValorNumerico $r.Valor_Compra) }

        try { Clear-Host } catch {}
        Exibir-LogoPrincipal
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host " |        >>> MODULO 1 - REGISTRO DE ATIVOS E PATRIMONIO <<<           |" -ForegroundColor Yellow
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host "  Bens cadastrados: " -NoNewline -ForegroundColor Gray
        Write-Host "$($ativos.Count)" -ForegroundColor Green
        Write-Host "  Valor do patrimonio: " -NoNewline -ForegroundColor Gray
        Write-Host ("R$ " + (Formatar-Moeda $totalValor)) -ForegroundColor Yellow
        Write-Host "  Proximo tombamento: " -NoNewline -ForegroundColor Gray
        Write-Host "$(Obter-ProximoPatrimonio -registros $ativos)" -ForegroundColor Cyan
        Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
        Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[1]" -NoNewline -ForegroundColor Green;  Write-Host " Novo Cadastro de Bem     (Rapido ou completo, em lote)          |" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[2]" -NoNewline -ForegroundColor Green;  Write-Host " Importar Planilha         (Cadastrar muitos bens via CSV)       |" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[3]" -NoNewline -ForegroundColor Yellow; Write-Host " Consultar Ativos          (Buscar, filtrar e ver ficha)         |" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[4]" -NoNewline -ForegroundColor Cyan;   Write-Host " Editar Ativo              (Corrigir dados do bem)               |" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[5]" -NoNewline -ForegroundColor Red;    Write-Host " Excluir Ativo             (Remover bem com backup automatico)   |" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[6]" -NoNewline -ForegroundColor Magenta;Write-Host " Termo de Responsabilidade (Gerar termo de custodia do bem)      |" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[7]" -NoNewline -ForegroundColor Magenta;Write-Host " Exportar Ativos           (CSV Excel / XLS Formatado / TXT)     |" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[8]" -NoNewline -ForegroundColor Yellow; Write-Host " Dashboard do Patrimonio   (Totais por setor, categoria, valor)  |" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[9]" -NoNewline -ForegroundColor White; Write-Host " Colaboradores/Fornecedores (Cadastro mestre de pessoas e empresas)|" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[10]" -NoNewline -ForegroundColor DarkCyan;Write-Host " Catalogo e Arquivos       (Tipos, planilhas e pasta de exportacoes)|" -ForegroundColor White
        Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[0]" -NoNewline -ForegroundColor DarkGray;Write-Host " Voltar ao Menu Principal                                        |" -ForegroundColor Gray
        Write-Host " +======================================================================+" -ForegroundColor Cyan
        Write-Host ""

        $opcao = Ler-OpcaoOuEsc -prompt " Digite o numero da opcao desejada (ou ESC para voltar)" -valorPadrao "1"

        try {
        switch ($opcao.Trim()) {
            "1" {
                Iniciar-CadastroBemPatrimonial -registrosAtuais $ativos -caminhoCsv $caminhoCsvAtivos
            }
            "2" {
                Importar-AtivosPlanilha -registrosAtuais $ativos -caminhoCsv $caminhoCsvAtivos
            }
            "3" {
                Menu-ConsultarAtivos -registros $ativos
            }
            "4" {
                Menu-EditarAtivo -registros $ativos -caminhoCsv $caminhoCsvAtivos
            }
            "5" {
                Excluir-BensAtivos -registros $ativos -caminhoCsv $caminhoCsvAtivos
            }
            "6" {
                Gerar-TermoResponsabilidade -registros $ativos -caminhoCsv $caminhoCsvAtivos
            }
            "7" {
                try { Clear-Host } catch {}
                Exibir-LogoPrincipal
                Write-Host " |                 >>> EXPORTACAO DE ATIVOS PATRIMONIAIS <<<            |" -ForegroundColor Yellow
                Write-Host " +======================================================================+" -ForegroundColor Cyan
                Write-Host "  [1] Exportar Planilha CSV UTF-8 (Excel PT-BR)" -ForegroundColor White
                Write-Host "  [2] Exportar Planilha Formatada (.XLS Estilizado)" -ForegroundColor Yellow
                Write-Host "  [3] Exportar Relatorio Executivo (.TXT)" -ForegroundColor Magenta
                Write-Host "  [4] Exportar apenas um Setor especifico" -ForegroundColor Cyan
                Write-Host "  [5] Exportar por Categoria de Bem" -ForegroundColor Cyan
                Write-Host "  [6] Exportar por Situacao" -ForegroundColor Cyan
                Write-Host "  [0] Voltar ao Módulo 1" -ForegroundColor Gray
                Write-Host ""

                $subOpcao = Ler-OpcaoOuEsc -prompt " Escolha a opcao de exportacao" -valorPadrao "1"

                switch ($subOpcao.Trim()) {
                    "1" { Exportar-AtivosPatrimoniais -registros $ativos -tipo "CSV_EXCEL" -abrirAposExportar:$true }
                    "2" { Exportar-AtivosPatrimoniais -registros $ativos -tipo "HTML_EXCEL" -abrirAposExportar:$true }
                    "3" { Exportar-AtivosPatrimoniais -registros $ativos -tipo "TXT_RELATORIO" -abrirAposExportar:$false }
                    "4" {
                        $setores = @()
                        foreach ($r in $ativos) { if ($r.Setor_Atribuido -and $setores -notcontains $r.Setor_Atribuido) { $setores += $r.Setor_Atribuido } }
                        if ($setores.Count -eq 0) {
                            Write-Host " [!] Nenhum setor cadastrado." -ForegroundColor Yellow
                        } else {
                            $opc = Selecionar-Opcao -titulo "Setor para exportar" -opcoes $setores
                            if ($null -ne $opc) { Exportar-AtivosPatrimoniais -registros $ativos -tipo "CSV_EXCEL" -filtroSetor $opc -abrirAposExportar:$true }
                        }
                    }
                    "5" {
                        $categorias = @()
                        foreach ($r in $ativos) { if ($r.Categoria_Bem -and $categorias -notcontains $r.Categoria_Bem) { $categorias += $r.Categoria_Bem } }
                        if ($categorias.Count -eq 0) {
                            Write-Host " [!] Nenhuma categoria cadastrada." -ForegroundColor Yellow
                        } else {
                            $opc = Selecionar-Opcao -titulo "Categoria para exportar" -opcoes $categorias
                            if ($null -ne $opc) { Exportar-AtivosPatrimoniais -registros $ativos -tipo "CSV_EXCEL" -filtroCategoria $opc -abrirAposExportar:$true }
                        }
                    }
                    "6" {
                        $situacoes = @()
                        foreach ($r in $ativos) { if ($r.Situacao -and $situacoes -notcontains $r.Situacao) { $situacoes += $r.Situacao } }
                        if ($situacoes.Count -eq 0) {
                            Write-Host " [!] Nenhuma situacao cadastrada." -ForegroundColor Yellow
                        } else {
                            $opc = Selecionar-Opcao -titulo "Situacao para exportar" -opcoes $situacoes
                            if ($null -ne $opc) { Exportar-AtivosPatrimoniais -registros $ativos -tipo "CSV_EXCEL" -filtroSituacao $opc -abrirAposExportar:$true }
                        }
                    }
                    "0" { }
                    default { Write-Host " Opcao invalida." -ForegroundColor Red }
                }
                Pausar-Tela
            }
            "8" {
                Menu-DashboardAtivos -registros $ativos
            }
            "9" {
                Menu-GerenciarParceiros
            }
            "10" {
                try { Clear-Host } catch {}
                Exibir-LogoPrincipal
                Write-Host " |                 >>> CATALOGO E ARQUIVOS <<<                         |" -ForegroundColor Yellow
                Write-Host " +======================================================================+" -ForegroundColor Cyan
                Write-Host "  [1] Abrir Catalogo de Tipos de Bens (Excel)" -ForegroundColor White
                Write-Host "  [2] Abrir Cadastro Mestre de Colaboradores (Excel)" -ForegroundColor White
                Write-Host "  [3] Abrir Cadastro Mestre de Fornecedores (Excel)" -ForegroundColor White
                Write-Host "  [4] Abrir Planilha de Ativos (Excel)" -ForegroundColor White
                Write-Host "  [5] Abrir Modelo de Importacao (Excel)" -ForegroundColor White
                Write-Host "  [6] Abrir Pasta de Termos de Responsabilidade" -ForegroundColor DarkCyan
                Write-Host "  [7] Abrir Pasta de Exportacoes" -ForegroundColor DarkCyan
                Write-Host "  [0] Voltar ao Módulo 1" -ForegroundColor Gray
                Write-Host ""
                $sub = Ler-OpcaoOuEsc -prompt " Escolha o arquivo" -valorPadrao "1"
                switch ($sub.Trim()) {
                    "1" { try { Invoke-Item $script:caminhoCatalogoAtivos } catch { Write-Host " [!] Catalogo nao encontrado." -ForegroundColor Red } }
                    "2" { try { Invoke-Item $script:caminhoCatalogoColaboradores } catch { Write-Host " [!] Catalogo nao encontrado." -ForegroundColor Red } }
                    "3" { try { Invoke-Item $script:caminhoCatalogoFornecedores } catch { Write-Host " [!] Catalogo nao encontrado." -ForegroundColor Red } }
                    "4" { try { Invoke-Item $caminhoCsvAtivos } catch { Write-Host " [!] Planilha nao encontrada." -ForegroundColor Red } }
                    "5" {
                        if (!(Test-Path $script:caminhoModeloImportacao)) { [void](Gerar-ModeloImportacao -caminho $script:caminhoModeloImportacao) }
                        try { Invoke-Item $script:caminhoModeloImportacao } catch { Write-Host " [!] Modelo nao encontrado." -ForegroundColor Red }
                    }
                    "6" { if (!(Test-Path $script:pastaTermos)) { try { New-Item -ItemType Directory -Path $script:pastaTermos -Force | Out-Null } catch {} }; try { Invoke-Item $script:pastaTermos } catch {} }
                    "7" { if (!(Test-Path $script:pastaExport)) { try { New-Item -ItemType Directory -Path $script:pastaExport -Force | Out-Null } catch {} }; try { Invoke-Item $script:pastaExport } catch {} }
                    "0" { }
                    default { Write-Host " Opcao invalida." -ForegroundColor Red }
                }
                Pausar-Tela
            }
            "0" { $emModulo = $false }
            default {
                Write-Host " Opcao invalida." -ForegroundColor Red
                Start-Sleep -Milliseconds 700
            }
        }
        }
        catch [System.OperationCanceledException] {
            Write-Host ""
            Write-Host " [i] Operacao cancelada. Nada foi gravado." -ForegroundColor Yellow
            Start-Sleep -Milliseconds 400
        }
    }
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
        @($registros | Where-Object { ($_.Setor_Atribuido -eq $filtroSetor) -or ($_.Setor -eq $filtroSetor) })
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
                    (& $limparCampo $r.Empresa_Unidade),
                    (& $limparCampo $r.Codigo_Tombamento),
                    (& $limparCampo $r.Numero_Nota_Fiscal),
                    (& $limparCampo $r.CRM_Serial),
                    (& $limparCampo $r.Categoria_Bem),
                    (& $limparCampo $r.Tipo_Especifico),
                    (& $limparCampo $r.Nome_Bem_Produto),
                    (& $limparCampo $r.Setor_Atribuido),
                    (& $limparCampo $r.Responsavel_Colaborador),
                    (& $limparCampo $r.Fornecedor),
                    (& $limparCampo $r.Situacao),
                    (& $limparCampo $r.Data_Compra),
                    (& $limparCampo $r.Valor_Compra),
                    (& $limparCampo $r.Localizacao_Fisica_Posto),
                    (& $limparCampo $r.Observacoes_Adicionais),
                    (& $limparCampo $r.Nome_Computador),
                    (& $limparCampo $r.Processador),
                    (& $limparCampo $r.Memoria_RAM),
                    (& $limparCampo $r.Armazenamento),
                    (& $limparCampo $r.Placa_Video),
                    (& $limparCampo $r.Sistema_Operacional),
                    (& $limparCampo $r.IP_Rede),
                    (& $limparCampo $r.MAC_Rede),
                    (& $limparCampo $r.Bateria),
                    (& $limparCampo $r.Android_Versao)
                ) -join ';'
                $linhasExport.Add($linha)
            }
            
            [System.IO.File]::WriteAllText($caminhoArquivo, (($linhasExport -join "`r`n") + "`r`n"), $script:utf8ComBOM)
            
            Write-Host "`n [OK] Planilha Excel CSV UTF-8 gerada com sucesso!" -ForegroundColor Green
            Write-Host "      Arquivo: " -NoNewline -ForegroundColor Gray
            Write-Host "$caminhoArquivo" -ForegroundColor Yellow
            Write-Host "      Total de ativos exportados: " -NoNewline -ForegroundColor Gray
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
            $totalAtivos = $registrosFiltrados.Count
            
            $linhasHtml = ""
            $num = 1
            foreach ($r in $registrosFiltrados) {
                $bgClass = if ($num % 2 -eq 0) { "even" } else { "odd" }
                $linhasHtml += "                <tr class=""$bgClass"">`r`n"
                $linhasHtml += "                    <td align=""center""><b>$num</b></td>`r`n"
                $linhasHtml += "                    <td class=""pat"">$($r.Codigo_Tombamento)</td>`r`n"
                $linhasHtml += "                    <td><b>$($r.Nome_Bem_Produto)</b></td>`r`n"
                $linhasHtml += "                    <td>$($r.Categoria_Bem)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Tipo_Especifico)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Setor_Atribuido)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Responsavel_Colaborador)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Fornecedor)</td>`r`n"
                $linhasHtml += "                    <td class=""serial"">$($r.CRM_Serial)</td>`r`n"
                $linhasHtml += "                    <td align=""center"">$($r.Data_Compra)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Valor_Compra)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Situacao)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Localizacao_Fisica_Posto)</td>`r`n"
                $linhasHtml += "                    <td>$($r.Observacoes_Adicionais)</td>`r`n"
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
                "        TOTAL DE ATIVOS CATALOGADOS: $totalAtivos | SISTEMA: INNOVCORE TI v2.0",
                '    </div>',
                '    <table>',
                '        <thead>',
                '            <tr>',
                '                <th>#</th>',
                '                <th>Tombamento</th>',
                '                <th>Nome do Bem</th>',
                '                <th>Categoria</th>',
                '                <th>Tipo</th>',
                '                <th>Setor</th>',
                '                <th>Responsavel</th>',
                '                <th>Fornecedor</th>',
                '                <th>N Serie</th>',
                '                <th>Data Compra</th>',
                '                <th>Valor (R$)</th>',
                '                <th>Situacao</th>',
                '                <th>Localizacao</th>',
                '                <th>Observacoes</th>',
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
                " Total de Ativos: $totalQtd",
                " Escopo: $escopoTxt",
                $divisoria,
                ""
            ) -join "`r`n"
            
            $i = 1
            foreach ($r in $registrosFiltrados) {
                $txt += "[$i] TOMBAMENTO: $($r.Codigo_Tombamento) | BEM: $($r.Nome_Bem_Produto)`r`n"
                $txt += "$subdivisoria`r`n"
                $txt += "  - Empresa / Unidade:      $($r.Empresa_Unidade)`r`n"
                $txt += "  - Categoria do Bem:       $($r.Categoria_Bem)`r`n"
                $txt += "  - Tipo / Subtipo:         $($r.Tipo_Especifico)`r`n"
                $txt += "  - Atribuido a (Setor):    $($r.Setor_Atribuido)`r`n"
                $txt += "  - Responsavel:            $($r.Responsavel_Colaborador)`r`n"
                $txt += "  - Fornecedor:             $($r.Fornecedor)`r`n"
                $txt += "  - Numero da Nota Fiscal:  $($r.Numero_Nota_Fiscal)`r`n"
                $txt += "  - CRM / Serial:           $($r.CRM_Serial)`r`n"
                $txt += "  - Situacao:               $($r.Situacao)`r`n"
                $txt += "  - Data da Compra:         $($r.Data_Compra)`r`n"
                $txt += "  - Valor da Compra:        R$ $($r.Valor_Compra)`r`n"
                $txt += "  - Localizacao Fisica:     $($r.Localizacao_Fisica_Posto)`r`n"
                $txt += "  - Observacoes:            $($r.Observacoes_Adicionais)`r`n"
                $txt += "  - Registrado em:          $($r.Data_Registro)`r`n`r`n"
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
                $setoresDisponiveis = @($registros | ForEach-Object { Obter-SetorDoRegistro $_ } | Where-Object { $_ -ne "" } | Select-Object -Unique)
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
                        "3" { @($registros | Where-Object { (Obter-SetorDoRegistro $_) -match "(?i)$t" }) }
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
            if ((Obter-SetorDoRegistro $_) -ne "") { Obter-SetorDoRegistro $_ } else { "Nao Informado" } 
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
        Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[0]" -NoNewline -ForegroundColor DarkGray;Write-Host " Voltar ao Menu Principal                                        |" -ForegroundColor Gray
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
    try {
    Atualizar-TamanhoTerminal
    try { Clear-Host } catch {}
    Exibir-LogoPrincipal
    Write-Host " +======================================================================+" -ForegroundColor Cyan
    Write-Host " |      [M1] MODULO 1 - REGISTRO DE ATIVOS E PATRIMONIO (BENS GERAIS)  |" -ForegroundColor Green
    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[1]" -NoNewline -ForegroundColor Green;  Write-Host " Central de Ativos      (Cadastrar, consultar e exportar bens)   |" -ForegroundColor White
    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
    Write-Host " |      [M2] MODULO 2 - COMPUTADORES & DESKTOPS (WINDOWS / MACOS)       |" -ForegroundColor Yellow
    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[2]" -NoNewline -ForegroundColor Green;  Write-Host " Novo Cadastro PC      (Coleta pecas e hardware de PCs/Laptops)  |" -ForegroundColor White
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[3]" -NoNewline -ForegroundColor Yellow; Write-Host " Consultar Computadores(Buscar por Patrimonio, Setor ou Usuario) |" -ForegroundColor White
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[4]" -NoNewline -ForegroundColor Cyan;   Write-Host " Editar Computador     (Alterar pecas, usuario ou departamento)  |" -ForegroundColor White
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[5]" -NoNewline -ForegroundColor Red;    Write-Host " Excluir Computador    (Remover PC com backup de seguranca)      |" -ForegroundColor White
    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
    Write-Host " |      [M3] MODULO 3 - MOBILE & TABLETS (ANDROID VIA CABO USB / ADB)   |" -ForegroundColor Magenta
    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[6]" -NoNewline -ForegroundColor Magenta;Write-Host " Central Gestao Mobile (Lista de Tablets, Exclusao e Relatorios) |" -ForegroundColor White
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[7]" -NoNewline -ForegroundColor Green;  Write-Host " Auto-Scanner Bancada  (Plug & Play USB - Captura MAC wlan0)     |" -ForegroundColor White
    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
    Write-Host " |      [M4] FERRAMENTAS GERAIS, EXPORTACAO & NUVEM                     |" -ForegroundColor Cyan
    Write-Host " +----------------------------------------------------------------------+" -ForegroundColor Cyan
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[8]" -NoNewline -ForegroundColor Green;  Write-Host " Central de Exportacao (Gerar planilhas Excel UTF-8 / CSV)       |" -ForegroundColor White
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[9]" -NoNewline -ForegroundColor Yellow; Write-Host " Dashboard & Metricas  (Graficos ASCII e resumo executivo)       |" -ForegroundColor White
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[10]" -NoNewline -ForegroundColor Green; Write-Host " Sincronizar InnovStock(Enviar dados para nuvem corporativa)    |" -ForegroundColor White
    Write-Host " |  " -NoNewline -ForegroundColor Cyan; Write-Host "[0]" -NoNewline -ForegroundColor DarkGray;Write-Host " Sair do Sistema                                                 |" -ForegroundColor Gray
    Write-Host " +======================================================================+" -ForegroundColor Cyan
    Write-Host ""
    
    $opcao = Ler-OpcaoOuEsc -prompt " Digite o numero da opcao desejada (ou ESC para sair)" -valorPadrao "1"
    
    switch ($opcao.Trim()) {
        "2" {
            # ------------------------------------------------------------------
            # OPCAO 2: NOVO CADASTRO (AUTOMATICO OU MANUAL OU ANDROID ADB)
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
                    $setoresHistorico = @($registrosAtuais | ForEach-Object { Obter-SetorDoRegistro $_ } | Where-Object { $_ -ne "" -and $_ -notmatch '^\d+$' } | Select-Object -Unique)
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
                    $qtdSetor = @($registrosAtuais | Where-Object { (Obter-SetorDoRegistro $_) -eq $listaSetores[$i] }).Count
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
        
        "3" {
            # ------------------------------------------------------------------
            # OPCAO 3: CONSULTAR / PESQUISAR CADASTROS
            # ------------------------------------------------------------------
            $registros = Obter-RegistrosCsv -caminho $caminhoCsv
            Menu-ConsultarRegistros -registros $registros
        }
        
        "4" {
            # ------------------------------------------------------------------
            # OPCAO 4: EDITAR CADASTRO
            # ------------------------------------------------------------------
            $registros = Obter-RegistrosCsv -caminho $caminhoCsv
            Menu-EditarCadastro -registros $registros
        }
        
        "5" {
            # ------------------------------------------------------------------
            # OPCAO 5: EXCLUIR CADASTRO
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
        
        "6" {
            # ------------------------------------------------------------------
            # OPCAO 6: CENTRAL DE GESTAO MOBILE & TABLETS (ANDROID / ADB)
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
        
        "7" {
            # ------------------------------------------------------------------
            # OPCAO 7: AUTO-SCANNER BANCADA PLUG & PLAY (USB)
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
        
        "8" {
            # ------------------------------------------------------------------
            # OPCAO 8: CENTRAL DE EXPORTACAO EXCEL UTF-8 / FORMATADOS
            # ------------------------------------------------------------------
            $registros = Obter-RegistrosCsv -caminho $caminhoCsv
            Menu-Exportacao -registros $registros
        }
        
        "9" {
            # ------------------------------------------------------------------
            # OPCAO 9: DASHBOARD EXECUTIVO E METRICAS
            # ------------------------------------------------------------------
            $registros = Obter-RegistrosCsv -caminho $caminhoCsv
            Menu-Estatisticas -registros $registros
        }

        "10" {
            # ------------------------------------------------------------------
            # OPCAO 10: SINCRONIZAR COM INNOVSTOCK (FIREBASE CLOUD)
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
        
        "1" {
            # ------------------------------------------------------------------
            # OPCAO 1: MODULO 1 - REGISTRO DE ATIVOS E PATRIMONIO
            # ------------------------------------------------------------------
            Menu-ModuloAtivos -caminhoCsvAtivos $script:caminhoCsvAtivos
        }

        "a" { Menu-ModuloAtivos -caminhoCsvAtivos $script:caminhoCsvAtivos }
        "ativo" { Menu-ModuloAtivos -caminhoCsvAtivos $script:caminhoCsvAtivos }
        "ativos" { Menu-ModuloAtivos -caminhoCsvAtivos $script:caminhoCsvAtivos }
        "patrimonio" { Menu-ModuloAtivos -caminhoCsvAtivos $script:caminhoCsvAtivos }
        "bens" { Menu-ModuloAtivos -caminhoCsvAtivos $script:caminhoCsvAtivos }

        "0" {
            # ------------------------------------------------------------------
            # SAIR DO SISTEMA (COM CONFIRMACAO)
            # ------------------------------------------------------------------
            try { Clear-Host } catch {}
            Exibir-LogoPrincipal
            Write-Host " +======================================================================+" -ForegroundColor Cyan
            Write-Host " |                  >>> ENCERRAR O SISTEMA <<<                        |" -ForegroundColor Yellow
            Write-Host " +======================================================================+" -ForegroundColor Cyan
            Write-Host "  Tudo que voce cadastrou permanece salvo no Pen Drive." -ForegroundColor Gray
            Write-Host ""

            $ativosSaida = Obter-RegistrosCsv -caminho $script:caminhoCsvAtivos -empresaPadrao "InnovEdu"
            $equipSaida = Obter-RegistrosCsv -caminho $caminhoCsv
            Write-Host "  RESUMO DA BASE" -ForegroundColor Cyan
            Write-Host "    Bens patrimoniais (Módulo 1): " -NoNewline -ForegroundColor Gray
            Write-Host "$($ativosSaida.Count)" -ForegroundColor Green
            Write-Host "    Equipamentos de TI (Módulo 2): " -NoNewline -ForegroundColor Gray
            Write-Host "$($equipSaida.Count)" -ForegroundColor Green
            Write-Host "    Planilha de ativos: " -NoNewline -ForegroundColor Gray
            Write-Host "$script:caminhoCsvAtivos" -ForegroundColor DarkYellow
            Write-Host "    Planilha de equipamentos: " -NoNewline -ForegroundColor Gray
            Write-Host "$caminhoCsv" -ForegroundColor DarkYellow
            Write-Host ""
            Write-Host "  [S] SIM - encerrar o InnovCore TI agora" -ForegroundColor Green
            Write-Host "  [N] NAO  - voltar ao menu principal" -ForegroundColor Cyan
            Write-Host ""

            $confirmaSaida = Ler-OpcaoOuEsc -prompt "  Deseja realmente sair? (S/N)" -valorPadrao "N"

            if ($confirmaSaida.Trim() -match '(?i)^s|sim|y|yes$') {
                $executando = $false
                try { Clear-Host } catch {}
                Exibir-LogoPrincipal
                Write-Host ""
                Write-Host "  Programa encerrado. Bom trabalho, Jordan!" -ForegroundColor Green
                Write-Host "  Os dados cadastrados permanecem salvos no Pen Drive." -ForegroundColor DarkGray
                Write-Host ""
                Start-Sleep -Milliseconds 1200
            } else {
                Write-Host ""
                Write-Host " [i] Saida cancelada. Voltando ao menu principal..." -ForegroundColor Yellow
                Start-Sleep -Milliseconds 600
            }
        }
        
        default {
            Write-Host " Opcao invalida. Tente novamente." -ForegroundColor Red
            Start-Sleep -Milliseconds 800
        }
    }
    }
    catch [System.OperationCanceledException] {
        # ESC confirmado pelo usuario durante um preenchimento: aborta com seguranca
        Write-Host ""
        Write-Host " [i] Operacao cancelada. Nada foi gravado. Voltando ao menu..." -ForegroundColor Yellow
        Start-Sleep -Milliseconds 400
    }
}
