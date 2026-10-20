# Memória Técnica: InnovCore TI
<!-- Developer Signature: [Jordan Araujo / Google DeepMind AAC Senior Full-Stack & CyberSecurity Architect] -->

**Data:** 2026-09-19 16:09 (BRT)  
**Projeto:** InnovCore TI — Automação de Auditoria de Hardware e Inventário  
**Repositório Oficial:** `https://github.com/RONALDJORDAN/InnovCore-TI.git`  
**Registro:** `INNOVCORE_TI_DEPLOY_v2026.09.19`

---
## 1. Visão Geral
Conjunto de ferramentas em PowerShell e Batch para diagnóstico, extração de relatórios de
hardware, auditoria de patrimônio e automação de extração de MAC/Série via Android Debug
Bridge (ADB) para tablets e computadores corporativos.

## 2. Estrutura
- `VERSAO_FINAL`: Scripts prontos para execução em campo e integração com InnovStock.
- `VERSAO_DESENVOLVIMENTO`: Scripts de testes de conformidade, benchmarks e cenários experimentais.

## 3. Módulo 1 — Registro de Ativos e Patrimônio (implementado em 19/10/2026)
Cadastro patrimonial completo de bens gerais (móveis, ar condicionado, impressoras, monitores,
chromebooks, tablets, carregadores, cabos, insumos, refrigeração, manutenção e lazer).

**Arquivos:**
- `VERSAO_FINAL/dados/ativos/inventario_ativos_innovcore.csv` — base de bens (cabeçalho patrimonial
  de 26 colunas, UTF-8 com BOM, delimitador `;`).
- `VERSAO_FINAL/dados/ativos/modelo_importacao_ativos.csv` — modelo de planilha para importação.
- `VERSAO_FINAL/dados/ativos/termos/` + `termos_emitidos.csv` — termos de responsabilidade.
- `VERSAO_FINAL/core/catalogos/catalogo_ativos.csv` — 95 tipos em 8 categorias.
- `VERSAO_FINAL/core/catalogos/catalogo_colaboradores.csv` — cadastro mestre (`COL-0001`).
- `VERSAO_FINAL/core/catalogos/catalogo_fornecedores.csv` — cadastro mestre (`FOR-0001`).

**Ficha patrimonial de 15 campos:** Categoria, Tipo, Nome do Bem, Código/Tombamento (`PAT-xxxxx`),
Nota Fiscal, CRM/Serial, Empresa/Unidade, Setor, Responsável, Fornecedor, Situação,
Data da Compra, Valor da Compra, Localização Física e Observações.

**Funcionalidades:** seletor com filtro incremental (digitar texto para achar item),
cadastro rápido em lote (herda setor/responsável/situação do item anterior), duplicidade
fail-fast no tombamento, importação em lote de planilha CSV com prévia, exclusão com backup,
termo de responsabilidade em TXT e HTML (imprimível) com histórico, dashboard do patrimônio
(totais por setor/categoria/situação + valor), cadastro mestre de colaboradores e fornecedores,
exportação CSV/XLS/TXT com filtros.

### Renumeração dos módulos (menu principal)
| Opção | Módulo |
|---|---|
| [1] | **Módulo 1** — Registro de Ativos e Patrimônio (novo) |
| [2] a [5] | **Módulo 2** — Computadores & Desktops |
| [6] e [7] | **Módulo 3** — Mobile & Tablets (ADB) |
| [8] a [10] | **Módulo 4** — Ferramentas, Exportação & Nuvem |

Atalhos de teclado para o Módulo 1: `a`, `ativo`, `ativos`, `patrimonio`, `bens`.

## 4. Interface adaptativa e tecla ESC (implementado em 19/10/2026)
- **Motor de layout adaptativo** (`Atualizar-TamanhoTerminal`): detecta a largura/altura real do
  terminal (`COLUMNS`/`LINES` → `RawUI`), expande as caixas de 73 a 120 colunas, centraliza na tela,
  compacta o banner em terminais com menos de 28 linhas e limita a listagem pela altura.
- **Proxies de console** (`Write-Host`, `Read-Host`, `Clear-Host`): o `Write-Host` redesenha os
  separadores de caixa e alinha a borda direita conforme a largura; o `Read-Host` detecta a tecla
  **ESC** e pergunta se o usuário deseja cancelar a operação (lança `OperationCanceledException`,
  capturada nos menus para abortar sem gravar nada pela metade); o `Clear-Host` apaga a área
  visível **e o scrollback** (sem "histórico" no terminal).
- Leitura de entrada redirecionada feita via `[Console]::In.ReadLine()` para que o EOF devolva
  `$null` (sinal de cancelamento) em vez de string vazia.
- **Armadilha do PowerShell 5.1:** o escape ANSI de ESC é `` `e `` **somente no PowerShell 6+**.
  No 5.1 ele vira a letra "e" (byte 101) e a sequência de limpeza vira lixo `e[2J` — a tela
  **não é limpa**. Correção: usar `[char]27` (`$esc = [char]27; [Console]::Write("$esc[2J$esc[3J$esc[H")`).

## 5. Escopo de entrega
O foco do projeto é o **sistema InnovCore TI em Windows** (PowerShell + Batch), rodando
100% portátil a partir do Pen Drive: auditoria de hardware (WMI), inventário de patrimônio,
exportação para Excel/CSV, dashboard no terminal e sincronização com o InnovStock.
A captura de tablets segue via **ADB no PC** (Módulo 3), sem aplicativo mobile nativo —
essa frente foi descartada e não faz parte da entrega.
