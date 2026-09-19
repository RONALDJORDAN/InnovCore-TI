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

## 🚀 Menu Principal Modular do Sistema v2.0

```text
 +======================================================================+
 |      [1] MODULO COMPUTADORES & DESKTOPS (WINDOWS / MACOS)            |
 +----------------------------------------------------------------------+
 |  [1] Novo Cadastro PC      (Coleta pecas e hardware de PCs/Laptops)  |
 |  [2] Consultar Computadores(Buscar por Patrimonio, Setor ou Usuario) |
 |  [3] Editar Computador     (Alterar pecas, usuario ou departamento)  |
 |  [4] Excluir Computador    (Remover PC com backup de seguranca)      |
 +----------------------------------------------------------------------+
 |      [2] MODULO MOBILE & TABLETS (ANDROID VIA CABO USB / ADB)        |
 +----------------------------------------------------------------------+
 |  [5] Central Gestao Mobile (Lista de Tablets, Exclusao e Relatorios) |
 |  [6] Auto-Scanner Bancada  (Plug & Play USB - Captura MAC wlan0)     |
 +----------------------------------------------------------------------+
 |      [3] FERRAMENTAS GERAIS, EXPORTACAO & NUVEM                      |
 +----------------------------------------------------------------------+
 |  [7] Central de Exportacao (Gerar planilhas Excel UTF-8 / CSV)       |
 |  [8] Dashboard & Metricas  (Graficos ASCII e resumo executivo)       |
 |  [9] Sincronizar InnovStock(Enviar dados para nuvem corporativa)     |
 |  [0] Sair do Sistema       (Ou pressione [ESC] em qualquer tela)     |
 +======================================================================+
```

---

## ☁️ Sincronização em Lote com o InnovStock (Opção 9)

Envio automatizado de todos os computadores coletados pelo Pen Drive diretamente para o Firebase Realtime Database do **InnovStock**:
- **Comando Direto:** Acionado via opção `[9]` no menu principal.
- **Resolução Automática de Colaboradores:** Vincula automaticamente a máquina ao colaborador existente ou cria o novo cadastro no InnovStock.
- **Auditoria de Duplicidade na Nuvem:** Identifica equipamentos existentes por patrimônio (`PAT-xxxx`) ou número de série, atualizando especificações sem duplicar registros no Firebase.
- **Auditoria de Duplicidade Local (Pen Drive):** Detecta em tempo real se o patrimônio, número de série da BIOS ou MAC Address já foram cadastrados e oferece opções de atualizar o cadastro existente com backup, gerar novo patrimônio ou cancelar.
- **Habilitação de Termos:** Permite emitir imediatamente o termo de responsabilidade no InnovStock Web.

---

## 📈 Dashboard Executivo no Terminal (Opção 8)

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
│   └── sincronizar_innovstock.ps1 # Motor de sincronização com o Firebase
│
├── dados/                         # BASES DE DADOS E REGISTROS EM CSV
│   ├── desktop/
│   │   └── inventario_innovtech.csv      # Base oficial de computadores/notebooks
│   └── mobile/
│       └── relatorio_macs_tablets.csv    # Base oficial de tablets e dispositivos móveis
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

*Desenvolvido com excelência por Jordan para InnovTech.*
