# ==============================================================================
# SKYNET TI - SISTEMA DE INVENTARIO OPERACIONAL DE EQUIPAMENTOS
# Cyberdyne Systems - Modulo de Reconhecimento de Hardware
# Desenvolvido por Jordan | Versao 1.0
# Compativel com Windows 10 e Windows 11
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try { $Host.UI.RawUI.WindowTitle = "SKYNET TI - Inventario de Equipamentos [v1.0]" } catch {}

# Identifica o diretorio onde o script esta rodando (Pendrive)
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$caminhoCsv = Join-Path $scriptDir "inventario_skynet.csv"
$cabecalhoCsv = "Data_Registro;Empresa;Patrimonio;Usuario;Funcao;Setor;Marca;Modelo;Numero_Serie;Ano_Fabricacao;Nome_Computador"

$utf8ComBOM = New-Object System.Text.UTF8Encoding($true)
$utf8SemBOM = New-Object System.Text.UTF8Encoding($false)

# Funcao para exibir o Banner Inicial Estilizado (Skynet / Cyberdyne Systems)
function Exibir-LogoPrincipal {
    Write-Host " +======================================================================+" -ForegroundColor Cyan
    Write-Host " |  ____  _  ____   ___   _ _____ _____                                 |" -ForegroundColor Red
    Write-Host " | / ___|| |/ /\ \ / / \ | | ____|_   _|                                |" -ForegroundColor Red
    Write-Host " | \___ \| ' /  \ V /|  \| |  _|   | |                                  |" -ForegroundColor Red
    Write-Host " |  ___) | . \   | | | |\  | |___  | |                                  |" -ForegroundColor Red
    Write-Host " | |____/|_|\_\  |_| |_| \_|_____| |_|                                  |" -ForegroundColor Red
    Write-Host " |                                                                      |" -ForegroundColor Cyan
    Write-Host " |             SKYNET TI - SISTEMA DE INVENTARIO OPERACIONAL            |" -ForegroundColor Yellow
    Write-Host " |         Desenvolvido por Jordan | Versao 1.0 (Cyberdyne Systems)     |" -ForegroundColor White
    Write-Host " +======================================================================+" -ForegroundColor Cyan
}

# Funcao para carregar registros do CSV de forma 100% confiavel (CRLF e LF, com/sem BOM)
function Obter-RegistrosCsv {
    param([string]$caminho)
    $lista = @()
    if (Test-Path $caminho) {
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
                    $obj | Add-Member -MemberType NoteProperty -Name $prop -Value $val
                }
                $lista += $obj
            }
        }
    }
    return $lista
}

# Funcao para salvar lista de registros no CSV
function Salvar-RegistrosCsv {
    param([string]$caminho, [array]$registros)
    $limparCampo = { param($c) if ($c) { ($c.ToString() -replace ';', ',') -replace "[\r\n]+", " " } else { "" } }
    
    [System.IO.File]::WriteAllText($caminho, ($cabecalhoCsv + "`r`n"), $utf8ComBOM)
    foreach ($r in $registros) {
        $linha = @(
            (& $limparCampo $r.Data_Registro),
            (& $limparCampo $r.Empresa),
            (& $limparCampo $r.Patrimonio),
            (& $limparCampo $r.Usuario),
            (& $limparCampo $r.Funcao),
            (& $limparCampo $r.Setor),
            (& $limparCampo $r.Marca),
            (& $limparCampo $r.Modelo),
            (& $limparCampo $r.Numero_Serie),
            (& $limparCampo $r.Ano_Fabricacao),
            (& $limparCampo $r.Nome_Computador)
        ) -join ';'
        [System.IO.File]::AppendAllText($caminho, ($linha + "`r`n"), $utf8SemBOM)
    }
}

# ==============================================================================
# MENU PRINCIPAL EM LOOP
# ==============================================================================
$executando = $true

while ($executando) {
    try { Clear-Host } catch {}
    Exibir-LogoPrincipal
    Write-Host " |                                                                      |" -ForegroundColor Cyan
    Write-Host " |  [1] Iniciar Novo Cadastro (Inventariar este computador)             |" -ForegroundColor White
    Write-Host " |  [2] Verificar Cadastros   (Listar registros salvos na planilha)     |" -ForegroundColor White
    Write-Host " |  [3] Excluir Cadastro      (Remover um registro especifico)          |" -ForegroundColor White
    Write-Host " |  [0] Sair                                                            |" -ForegroundColor Gray
    Write-Host " |                                                                      |" -ForegroundColor Cyan
    Write-Host " +======================================================================+" -ForegroundColor Cyan
    Write-Host ""
    
    $opcao = Read-Host " Escolha uma opcao [1]"
    if ([string]::IsNullOrWhiteSpace($opcao)) { $opcao = "1" }
    
    switch ($opcao.Trim()) {
        "1" {
            # ------------------------------------------------------------------
            # OPCAO 1: NOVO CADASTRO
            # ------------------------------------------------------------------
            try { Clear-Host } catch {}
            Exibir-LogoPrincipal
            Write-Host " |                >>> INICIAR NOVO CADASTRO TI <<<                      |" -ForegroundColor Yellow
            Write-Host " +======================================================================+" -ForegroundColor Cyan
            Write-Host ""
            
            # Carrega registros atuais para calcular proximo patrimonio e setores existentes
            $registrosAtuais = Obter-RegistrosCsv -caminho $caminhoCsv
            
            # 1. Proximo numero de patrimonio sequencial
            $proximoNumero = $registrosAtuais.Count + 1
            $patrimonioSugerido = "PAT-" + $proximoNumero.ToString("D4")
            
            # 2. Obter setores ja cadastrados anteriormente + setores comuns da InnovTech
            $setoresPadrao = @("Operacional", "TI", "Financeiro", "RH", "Comercial", "Marketing", "Administrativo", "Logistica")
            $setoresHistorico = @()
            if ($registrosAtuais.Count -gt 0) {
                $setoresHistorico = @($registrosAtuais | ForEach-Object { $_.Setor.Trim() } | Where-Object { $_ -ne "" -and $_ -notmatch '^\d+$' } | Select-Object -Unique)
            }
            
            $listaSetores = @()
            # Prioriza setores do historico
            foreach ($s in $setoresHistorico) {
                if ($listaSetores -notcontains $s) { $listaSetores += $s }
            }
            # Adiciona setores padrao
            foreach ($s in $setoresPadrao) {
                if ($listaSetores -notcontains $s) { $listaSetores += $s }
            }
            
            # Perguntas
            Write-Host " [1/5] IDENTIFICACAO DA EMPRESA" -ForegroundColor Yellow
            $empresaInput = Read-Host "   Empresa [Pressione ENTER para 'InnovTech']"
            $empresa = if ([string]::IsNullOrWhiteSpace($empresaInput)) { "InnovTech" } else { $empresaInput.Trim() }
            
            Write-Host "`n [2/5] DADOS DO USUARIO / RESPONSAVEL" -ForegroundColor Yellow
            $usuario = ""
            while ([string]::IsNullOrWhiteSpace($usuario)) {
                $usuario = Read-Host "   Nome do Usuario / Colaborador"
                if ([string]::IsNullOrWhiteSpace($usuario)) {
                    Write-Host "   (!) Por favor, digite o nome do usuario." -ForegroundColor Red
                }
            }
            
            Write-Host "`n [3/5] FUNCAO / CARGO" -ForegroundColor Yellow
            $funcao = ""
            while ([string]::IsNullOrWhiteSpace($funcao)) {
                $funcao = Read-Host "   Funcao / Cargo"
                if ([string]::IsNullOrWhiteSpace($funcao)) {
                    Write-Host "   (!) Por favor, digite a funcao." -ForegroundColor Red
                }
            }
            
            Write-Host "`n [4/5] SETOR / DEPARTAMENTO" -ForegroundColor Yellow
            Write-Host "   Setores encontrados / sugeridos:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $listaSetores.Count; $i++) {
                $marcador = if ($setoresHistorico -contains $listaSetores[$i]) { "(*Ja Cadastrado*)" } else { "" }
                Write-Host ("     [{0}] {1} {2}" -f ($i + 1), $listaSetores[$i], $marcador).TrimEnd() -ForegroundColor Green
            }
            Write-Host "     [0] Digitar um outro setor novo" -ForegroundColor Gray
            
            $setor = ""
            while ([string]::IsNullOrWhiteSpace($setor)) {
                $setorInput = Read-Host "   Escolha o numero ou digite o nome do setor"
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
                            if ([string]::IsNullOrWhiteSpace($setor)) {
                                Write-Host "   (!) O setor nao pode ficar vazio." -ForegroundColor Red
                            }
                        }
                    } else {
                        $setor = $setorInput.Trim()
                    }
                } else {
                    $setor = $setorInput.Trim()
                }
            }
            Write-Host "   -> Setor definido: $setor" -ForegroundColor Yellow
            
            Write-Host "`n [5/5] NUMERO DE PATRIMONIO" -ForegroundColor Yellow
            $patrimonioInput = Read-Host "   Patrimonio [Pressione ENTER para '$patrimonioSugerido']"
            $patrimonio = if ([string]::IsNullOrWhiteSpace($patrimonioInput)) { $patrimonioSugerido } else { $patrimonioInput.Trim() }
            
            Write-Host ""
            Write-Host " [*] Varrendo hardware do alvo... [Cyberdyne Scan]" -ForegroundColor Green
            
            # Coleta de Hardware
            $dataColeta = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
            $hostname = $env:COMPUTERNAME
            
            # BIOS (Serial e Ano de Fabricacao)
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
            
            # Computador (Marca e Modelo)
            $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
            $marca = if ($cs.Manufacturer) { $cs.Manufacturer.Trim() } else { "N/A" }
            $modelo = if ($cs.Model) { $cs.Model.Trim() } else { "N/A" }
            
            if ($marca -match '(?i)System manufacturer|To be filled|Default' -and $bb.Manufacturer) {
                $marca = $bb.Manufacturer.Trim()
            }
            if ($modelo -match '(?i)System Product Name|To be filled|Default' -and $bb.Product) {
                $modelo = $bb.Product.Trim()
            }
            
            # Resumo
            $divisoria = "=" * 60
            $resumo = @"
$divisoria
       SKYNET TI - INVENTARIO DE HARDWARE ($empresa)
$divisoria
  - Empresa:            $empresa
  - Patrimonio:         $patrimonio
  - Usuario:            $usuario
  - Funcao / Cargo:     $funcao
  - Setor:              $setor

  - Marca:              $marca
  - Modelo:             $modelo
  - Numero de Serie:    $numeroSerie
  - Ano de Fabricacao:  $anoFabricacao
  - Computador (Host):  $hostname
  - Data do Registro:   $dataColeta
$divisoria
"@
            try { Clear-Host } catch {}
            Exibir-LogoPrincipal
            Write-Host ""
            Write-Host $resumo -ForegroundColor Cyan
            
            # Copiar para Clipboard
            $copiado = $false
            try {
                Set-Clipboard -Value $resumo
                $copiado = $true
            } catch {
                try { $resumo | clip.exe; $copiado = $true } catch {}
            }
            
            Write-Host ""
            if ($copiado) {
                Write-Host " [OK] Informacoes COPIADAS para a Area de Transferencia (Ctrl + V)!" -ForegroundColor Green
            }
            
            # Salvar no CSV
            $novoItem = [PSCustomObject]@{
                Data_Registro   = $dataColeta
                Empresa         = $empresa
                Patrimonio      = $patrimonio
                Usuario         = $usuario
                Funcao          = $funcao
                Setor           = $setor
                Marca           = $marca
                Modelo          = $modelo
                Numero_Serie    = $numeroSerie
                Ano_Fabricacao  = $anoFabricacao
                Nome_Computador = $hostname
            }
            
            $listaAtualizada = @($registrosAtuais) + $novoItem
            Salvar-RegistrosCsv -caminho $caminhoCsv -registros $listaAtualizada
            
            Write-Host " [OK] Registrado com sucesso na base de dados Skynet CSV!" -ForegroundColor Green
            Write-Host "      Arquivo: $caminhoCsv" -ForegroundColor Yellow
            Write-Host ""
            Write-Host " Pressione ENTER para voltar ao menu principal..." -ForegroundColor Gray
            try { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") } catch { Read-Host | Out-Null }
        }
        
        "2" {
            # ------------------------------------------------------------------
            # OPCAO 2: VERIFICAR CADASTROS
            # ------------------------------------------------------------------
            try { Clear-Host } catch {}
            Exibir-LogoPrincipal
            Write-Host " |                >>> CADASTROS REGISTRADOS <<<                         |" -ForegroundColor Yellow
            Write-Host " +======================================================================+" -ForegroundColor Cyan
            Write-Host ""
            
            $registros = Obter-RegistrosCsv -caminho $caminhoCsv
            
            if ($registros.Count -eq 0) {
                Write-Host "  Nenhum equipamento registrado na base Skynet ainda." -ForegroundColor Yellow
                Write-Host "  Utilize a opcao [1] para inventariar o primeiro computador." -ForegroundColor Gray
            } else {
                Write-Host "  Total de maquinas catalogadas no Skynet: $($registros.Count)" -ForegroundColor Green
                Write-Host ""
                
                $indice = 1
                foreach ($r in $registros) {
                    Write-Host "  [$indice] PATRIMONIO: $($r.Patrimonio)" -ForegroundColor Yellow
                    Write-Host "      Colaborador: $($r.Usuario) | Setor: $($r.Setor) | Funcao: $($r.Funcao)" -ForegroundColor White
                    Write-Host "      Equipamento: $($r.Marca) $($r.Modelo) | Serial: $($r.Numero_Serie) | Ano: $($r.Ano_Fabricacao)" -ForegroundColor Cyan
                    Write-Host "      Host: $($r.Nome_Computador) | Registrado em: $($r.Data_Registro)" -ForegroundColor Gray
                    Write-Host "  ----------------------------------------------------------------------" -ForegroundColor DarkGray
                    $indice++
                }
                
                Write-Host ""
                Write-Host "  [E] Abrir planilha completa no Excel" -ForegroundColor Green
                Write-Host "  [ENTER] Voltar ao menu principal" -ForegroundColor Gray
                $acaoVer = Read-Host "  Opcao"
                if ($acaoVer.Trim() -match '^(?i)e$') {
                    try { Invoke-Item $caminhoCsv } catch {}
                }
            }
            
            if ($registros.Count -eq 0) {
                Write-Host ""
                Write-Host " Pressione ENTER para voltar ao menu principal..." -ForegroundColor Gray
                try { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") } catch { Read-Host | Out-Null }
            }
        }
        
        "3" {
            # ------------------------------------------------------------------
            # OPCAO 3: EXCLUIR CADASTRO
            # ------------------------------------------------------------------
            try { Clear-Host } catch {}
            Exibir-LogoPrincipal
            Write-Host " |                >>> EXCLUSAO DE CADASTRO <<<                          |" -ForegroundColor Red
            Write-Host " +======================================================================+" -ForegroundColor Cyan
            Write-Host ""
            
            $registros = Obter-RegistrosCsv -caminho $caminhoCsv
            
            if ($registros.Count -eq 0) {
                Write-Host "  Nao ha registros para excluir." -ForegroundColor Yellow
                Write-Host ""
                Write-Host " Pressione ENTER para voltar ao menu principal..." -ForegroundColor Gray
                try { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") } catch { Read-Host | Out-Null }
            } else {
                $indice = 1
                foreach ($r in $registros) {
                    Write-Host "  [$indice] $($r.Patrimonio) - $($r.Usuario) ($($r.Setor) - $($r.Marca) $($r.Modelo))" -ForegroundColor White
                    $indice++
                }
                Write-Host ""
                Write-Host "  Digite o NUMERO do registro que deseja EXCLUIR" -ForegroundColor Yellow
                Write-Host "  (Ou digite 0 para cancelar e voltar ao menu):" -ForegroundColor Gray
                $respExcluir = Read-Host "  Registro a excluir"
                
                $numExcluir = 0
                if ([int]::TryParse($respExcluir.Trim(), [ref]$numExcluir)) {
                    if ($numExcluir -ge 1 -and $numExcluir -le $registros.Count) {
                        $itemRemovido = $registros[$numExcluir - 1]
                        Write-Host ""
                        Write-Host "  ATENCAO: Deseja realmente excluir o cadastro de:" -ForegroundColor Red
                        Write-Host "  Patrimonio: $($itemRemovido.Patrimonio) - $($itemRemovido.Usuario) ($($itemRemovido.Setor))" -ForegroundColor Yellow
                        $confirma = Read-Host "  Confirmar exclusao? (S/N) [N]"
                        
                        if ($confirma.Trim() -match '^(?i)s|sim|y|yes$') {
                            $novaLista = @()
                            for ($j = 0; $j -lt $registros.Count; $j++) {
                                if ($j -ne ($numExcluir - 1)) {
                                    $novaLista += $registros[$j]
                                }
                            }
                            Salvar-RegistrosCsv -caminho $caminhoCsv -registros $novaLista
                            Write-Host ""
                            Write-Host "  [OK] Registro eliminado da base de dados com sucesso!" -ForegroundColor Green
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
                
                Write-Host ""
                Write-Host " Pressione ENTER para voltar ao menu principal..." -ForegroundColor Gray
                try { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") } catch { Read-Host | Out-Null }
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
            Write-Host "  Hasta la vista, baby! Programa encerrado. Bom trabalho, Jordan!" -ForegroundColor Green
            Write-Host ""
            Start-Sleep -Milliseconds 1200
        }
        
        default {
            Write-Host " Opcao invalida. Tente novamente." -ForegroundColor Red
            Start-Sleep -Milliseconds 1000
        }
    }
}
