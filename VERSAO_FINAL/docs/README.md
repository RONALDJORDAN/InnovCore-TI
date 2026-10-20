# 💼 INNOVCORE TI - Sistema de Gestão de Ativos e Inventário de Hardware
> Sistema corporativo avançado de inventário, catalogação e auditoria de parques de TI.

Desenvolvido por **Jordan** | Versão 2.0 (Edição Corporativa Avançada)

Sistema 100% portátil, rápido, seguro e inteligente para inventariar computadores corporativos (Windows 10 e Windows 11) e dispositivos móveis/tablets (Android via ADB) diretamente através de um Pen Drive, com modos **Automático** (Scanner) e **Manual** (Digitação Completa), suíte de exportação para Excel UTF-8, Dashboard Executivo no Terminal e **Sincronização em Nuvem com o InnovStock**.

---

## ⚡ Identidade Visual

```text
 +======================================================================+
 |  ___ _   _ _   _  _____     ______ _____ ____  _____   _____ ___     |
 | |_ _| \ | | \ | |/ _ \ \   / / ___/ _ \|  _ \| ____| |_   _|_ _|    |
 |  | ||  \| |  \| | | | \ \ / / |  | | | | |_) |  _|     | |  | |     |
 |  | || |\  | |\  | |_| |\ V /| |__| |_| |  _ <| |___    | |  | |     |
 | |___|_| \_|_| \_|\___/  \_/  \____\___/|_| \_\_____|   |_| |___|    |
 |                                                                      |
 |             INNOVCORE TI - GESTAO DE ATIVOS E INVENTARIO             |
 |                 Desenvolvido por Jordan | Versao 2.0                 |
 +======================================================================+
```

---

## 🖥️ Interface adaptativa à resolução do terminal

O InnovCore TI se ajusta automaticamente à janela em que for aberto:

- **Largura:** as caixas se expandem para ocupar o terminal disponível (de 73 colunas até 120),
  centralizadas horizontalmente quando sobra espaço.
- **Altura:** em terminais baixos (menos de 28 linhas) o banner de 11 linhas vira uma versão
  compacta de 2 linhas, para o menu sempre caber na tela.
- **Tabelas:** ganham ou perdem colunas (Categoria, Setor, Valor) conforme a largura, e a
  listagem limita a quantidade de linhas exibidas conforme a altura.
- **Tela limpa:** a cada tela o sistema apaga a área visível **e o histórico de rolagem**
  (scrollback), de forma que o terminal mostre apenas a interface atual.

### Tecla ESC (cancelar e voltar)

| Situação | Comportamento |
|---|---|
| Em qualquer menu | Volta imediatamente para o menu anterior |
| Durante o preenchimento de um campo | Pergunta **"Deseja CANCELAR a operação atual?"** — confirmando, aborta com segurança (nada é gravado pela metade) e volta ao menu |

---

## 🚀 Menu Principal Modular do Sistema v2.0

```text
 +======================================================================+
 |      [M1] MODULO 1 - REGISTRO DE ATIVOS E PATRIMONIO (BENS GERAIS)  |
 +----------------------------------------------------------------------+
 |  [1] Central de Ativos      (Cadastrar, consultar e exportar bens)     |
 +----------------------------------------------------------------------+
 |      [M2] MODULO 2 - COMPUTADORES & DESKTOPS (WINDOWS / MACOS)       |
 +----------------------------------------------------------------------+
 |  [2] Novo Cadastro PC      (Coleta pecas e hardware de PCs/Laptops)  |
 |  [3] Consultar Computadores(Buscar por Patrimonio, Setor ou Usuario) |
 |  [4] Editar Computador     (Alterar pecas, usuario ou departamento)  |
 |  [5] Excluir Computador    (Remover PC com backup de seguranca)      |
 +----------------------------------------------------------------------+
 |      [M3] MODULO 3 - MOBILE & TABLETS (ANDROID VIA CABO USB / ADB)   |
 +----------------------------------------------------------------------+
 |  [6] Central Gestao Mobile (Lista de Tablets, Exclusao e Relatorios) |
 |  [7] Auto-Scanner Bancada  (Plug & Play USB - Captura MAC wlan0)     |
 +----------------------------------------------------------------------+
 |      [M4] FERRAMENTAS GERAIS, EXPORTACAO & NUVEM                     |
 +----------------------------------------------------------------------+
 |  [8] Central de Exportacao (Gerar planilhas Excel UTF-8 / CSV)       |
 |  [9] Dashboard & Metricas  (Graficos ASCII e resumo executivo)       |
 |  [10] Sincronizar InnovStock(Enviar dados para nuvem corporativa)    |
 |  [0] Sair do Sistema       (Ou pressione [ESC] em qualquer tela)     |
 +======================================================================+
```

> Atalhos de teclado: digite `a`, `ativo`, `patrimonio` ou `bens` para entrar direto no Módulo 1.

---

## 🏛️ Módulo 1 — Registro de Ativos e Patrimônio (Opção 1)

Cadastro **patrimonial completo** de bens que não são computadores: cadeiras, mesas, armários,
ar condicionado, impressoras, monitores, chromebooks, tablets, carregadores, cabos, insumos de
papelaria, equipamentos de manutenção, refrigeração, limpeza e lazer.

### Menu do Módulo 1

```text
 [1] Novo Cadastro de Bem     (Rápido ou completo, em lote)
 [2] Importar Planilha         (Cadastrar muitos bens via CSV)
 [3] Consultar Ativos          (Buscar, filtrar e ver ficha)
 [4] Editar Ativo              (Corrigir dados do bem)
 [5] Excluir Ativo             (Remover bem com backup automático)
 [6] Termo de Responsabilidade (Gerar termo de custódia do bem)
 [7] Exportar Ativos           (CSV Excel / XLS Formatado / TXT)
 [8] Dashboard do Patrimônio   (Totais por setor, categoria, valor)
 [9] Colaboradores/Fornecedores (Cadastro mestre de pessoas e empresas)
 [10] Catalogo e Arquivos       (Tipos, planilhas e pasta de exportações)
 [0] Voltar ao Menu Principal
```

### Ficha patrimonial (15 passos)

| # | Campo | Observação |
|---|-------|------------|
| 1 | Categoria do Bem | Vem do catálogo editável (8 categorias) |
| 2 | Tipo / Subtipo | Filtra automaticamente pela categoria escolhida |
| 3 | **Nome do Bem Imobilizado / Produto** | Obrigatório — sugere o nome do catálogo |
| 4 | **Código / Tombamento** | `PAT-00025` — imutável, sequencial automático |
| 5 | Nº da Nota Fiscal | `NF-129480` (opcional, a mesma NF cobre vários bens) |
| 6 | CRM / Serial | Serial, RFID, IMEI ou código de barras |
| 7 | **Empresa / Unidade** | `InnovEdu` (padrão) ou `InnovTech` |
| 8 | **Atribuído a: Setor** | Setores oficiais + histórico + livre |
| 9 | **Responsável / Colaborador** | Vem do cadastro mestre de colaboradores |
| 10 | **Fornecedor** | Vem do cadastro mestre de fornecedores |
| 11 | **Situação** | Ativo/Operacional, Em Manutenção, Em Estoque, Baixado/Inativo, Emprestado/Cedido |
| 12 | **Data da Compra** | `19/10/2026` (DD/MM/AAAA validado) |
| 13 | **Valor da Compra** | `R$ 0,00` (aceita `1500`, `1500,00`, `1.500,00`) |
| 14 | **Localização Física / Posto** | Ex.: `Sala TI - Posto 03`, `Almoxarifado` |
| 15 | **Observações Adicionais** | Garantia, condições, acessórios |

### Recursos do módulo

- **Seletor com filtro** — em qualquer lista (tipo, setor, responsável, fornecedor) você **digita
  texto para filtrar** em vez de rolar dezenas de itens. Ex.: digite `ar con` e o sistema mostra
  apenas “Ar condicionado”.
- **Cadastro rápido em lote** — o modo *Rápido* herda setor, responsável, fornecedor e situação do
  item anterior. Cadastrar 50 cadeiras = 50 × 4 perguntas, não 50 × 15. Ao final ele pergunta
  “Cadastrar OUTRO bem agora?”.
- **Duplicidade fail-fast** — o tombamento é validado **na hora em que é digitado**, antes das
  demais perguntas, com opção de trocar pelo próximo livre. A Nota Fiscal **não** é usada como
  critério (uma nota normalmente cobre um lote de bens).
- **Importação em lote via planilha** — preencha `dados/ativos/modelo_importacao_ativos.csv` no
  Excel e importe pela opção `[2]`. O sistema mostra uma **prévia** (quantos entram, quantos são
  ignorados e por quê) antes de gravar.
- **Exclusão com backup** — opção `[5]`, com confirmação digitada (`EXCLUIR`) e backup automático
  em `Backups/`.
- **Termo de Responsabilidade** — opção `[6]` gera o termo do bem em **TXT** e em **HTML**
  (pronto para imprimir, com campo de assinatura), salvo em `dados/ativos/termos/`. O histórico de
  quem recebeu, quando e qual arquivo fica em `dados/ativos/termos_emitidos.csv`.
- **Cadastro mestre** — opção `[9]` gerencia **colaboradores** (`COL-0001`) e **fornecedores**
  (`FOR-0001`) com código automático, setor, função, e-mail, telefone e status ativo/inativo.
- **Dashboard do patrimônio** — opção `[8]`: total de bens, valor total e médio, distribuição por
  situação, categoria e setor (com barras), e os 5 bens de maior valor.
- **Catálogo editável** (`core/catalogos/catalogo_ativos.csv`): 95 tipos em 8 categorias. Abra no
  Excel e adicione linhas no formato `Categoria;Tipo;Nome do Bem;Setor sugerido`.
- **Bases de dados isoladas**, no mesmo cabeçalho patrimonial de 26 colunas do restante do sistema
  (compatível com Excel PT-BR, UTF-8 com BOM).
- **Exportação** no mesmo padrão do restante do sistema: CSV Excel UTF-8, XLS formatado e TXT
  executivo, com filtros por Setor, Categoria e Situação.

### Arquivos do Módulo 1

```text
core/catalogos/catalogo_ativos.csv          # 95 tipos de bens (editável)
core/catalogos/catalogo_colaboradores.csv   # Cadastro mestre de colaboradores
core/catalogos/catalogo_fornecedores.csv    # Cadastro mestre de fornecedores
dados/ativos/inventario_ativos_innovcore.csv # Base de bens patrimoniais
dados/ativos/modelo_importacao_ativos.csv   # Modelo de planilha para importação
dados/ativos/termos_emitidos.csv            # Histórico de termos emitidos
dados/ativos/termos/                        # Termos gerados (.txt e .html)
```

---

## ☁️ Sincronização em Lote com o InnovStock (Opção 10)

Envio automatizado de todos os computadores coletados pelo Pen Drive diretamente para o Firebase Realtime Database do **InnovStock**:
- **Comando Direto:** Acionado via opção `[10]` no menu principal.
- **Resolução Automática de Colaboradores:** Vincula automaticamente a máquina ao colaborador existente ou cria o novo cadastro no InnovStock.
- **Auditoria de Duplicidade na Nuvem:** Identifica equipamentos existentes por patrimônio (`PAT-xxxx`) ou número de série, atualizando especificações sem duplicar registros no Firebase.
- **Auditoria de Duplicidade Local (Pen Drive):** Detecta em tempo real se o patrimônio, número de série da BIOS ou MAC Address já foram cadastrados e oferece opções de atualizar o cadastro existente com backup, gerar novo patrimônio ou cancelar.
- **Habilitação de Termos:** Permite emitir imediatamente o termo de responsabilidade no InnovStock Web.

---

## 📈 Dashboard Executivo no Terminal (Opção 9)

Interface rica em modo texto (ASCII/ANSI) com navegação interativa por abas:
- **Cartões de Indicadores (KPI Cards):** Total de ativos cadastrados, memória RAM global gerenciada (GB), maior departamento, contagem de notebooks vs desktops vs servidores, e fabricante predominante.
- **Visões Interativas do Dashboard:**
  - **Aba 1 (Por Setor & Integrantes):** Demonstrativo completo de equipamentos agrupados por departamento/setor, exibindo o nome de cada integrante, função/cargo, equipamento/modelo e o número de patrimônio em **ordem numérica rigorosa**. Inclui suporte a filtro por setor individual `[S]`.
  - **Aba 2 (Gráficos & Marcas):** Percentual e gráficos de barras ASCII por departamento e por fabricante.
  - **Aba 3 (Hardware & SO):** Distribuição por Sistema Operacional, tipos de máquinas e famílias de processadores.
  - **Aba 4 (Tabela Geral de Ativos):** Visão completa de todos os equipamentos cadastrados em ordem numérica de patrimônio.

---

## 🌟 Modos de Cadastro

### 🤖 1. Modo Automático (Scanner de Hardware de PC)
- Conecte o Pen Drive no computador que deseja inventariar.
- O sistema coleta automaticamente: Processador, Memória RAM (detalhes por slot), SSD/HDD, Placa de Vídeo, Sistema Operacional, IP, MAC, Serial da BIOS e Fabricante.

### ✍️ 2. Modo Manual (Digitação Completa de Qualquer Máquina)
- Ideal para cadastrar computadores desligados, notebooks na bancada ou na caixa, servidores ou fichas de papel.
- Permite preencher passo a passo: Empresa, Colaborador, Cargo, Setor, Patrimônio sugerido, Tipo de Equipamento (Notebook, Desktop, All-in-One, Servidor), Marca, Modelo, Serial, Ano, Hostname, CPU, RAM, SSD/HDD, GPU, SO, IP e MAC.

### 📱 3. Modo Android / Tablet (Scanner ADB via Cabo USB)
- Ideal para cadastrar tablets corporativos e educacionais (Positivo Tab Vision 7, Multilaser Tab-Multi, Nokia T20, Samsung, etc.).
- Conecte o tablet via cabo USB com a Depuração USB ativada.
- O sistema coleta em milissegundos: Número de Série de hardware (`ro.serialno`), **Endereço MAC Wi-Fi Físico de Fábrica** (via dumpsys wifi, NVRAM e sysfs mesmo com Wi-Fi desligado ou randomizado), Marca, Modelo, SoC/Processador, Memória RAM, Armazenamento Interno, Versão Android e Nível de Bateria.

---

## 📁 Estrutura Oficial do Projeto (VERSAO_FINAL)

```text
VERSAO_FINAL/
├── InnovCore_TI.bat               # ÚNICO EXECUTÁVEL NA RAIZ (Inicia o sistema com 2 cliques)
│
├── core/                          # MOTORES CENTRAIS DO SISTEMA
│   ├── inventario.ps1             # Motor principal do InnovCore TI v2.0
│   ├── sincronizar_innovstock.ps1 # Motor de sincronização com o Firebase
│   └── catalogos/                 # CATÁLOGOS MESTRES EDITÁVEIS (Excel)
│       ├── catalogo_ativos.csv        # 95 tipos de bens do Módulo 1
│       ├── catalogo_colaboradores.csv  # Cadastro mestre de colaboradores
│       └── catalogo_fornecedores.csv   # Cadastro mestre de fornecedores
│
├── dados/                         # BASES DE DADOS E REGISTROS EM CSV
│   ├── desktop/
│   │   └── inventario_innovtech.csv      # Base oficial de computadores/notebooks
│   ├── mobile/
│   │   └── relatorio_macs_tablets.csv    # Base oficial de tablets e dispositivos móveis
│   └── ativos/
│       ├── inventario_ativos_innovcore.csv # Base oficial de bens patrimoniais (Módulo 1)
│       ├── modelo_importacao_ativos.csv   # Modelo de planilha para importação em lote
│       ├── termos_emitidos.csv            # Histórico de termos de responsabilidade
│       └── termos/                        # Termos gerados (.txt e .html)
│
├── mobile/                        # FERRAMENTAS E SCRIPTS PARA ANDROID / ADB
│   ├── capturar_mac_tablet.ps1   # Script especialista em extração de MAC wlan0
│   └── platform-tools/           # Binários portáteis oficiais do Google ADB (offline)
│       └── adb.exe, etc.
│
├── docs/                          # DOCUMENTAÇÃO E GUIAS DE USO
│   ├── COMO_USAR.txt              # Guia rápido de consulta no Pen Drive
│   └── README.md                  # Documentação técnica e operacional
│
├── Exportacoes/                   # RELATÓRIOS E PLANILHAS EXCEL GERADAS
│   └── LEIA-ME.txt
│
└── Backups/                       # BACKUPS AUTOMÁTICOS COM TIMESTAMP
    └── LEIA-ME.txt
```

---

*Desenvolvido com excelencia por Jordan para InnovTech.*
