# 💼 INNOVCORE TI - Sistema de Gestão de Ativos e Inventário de Hardware
> Sistema corporativo de inventário e catalogação de equipamentos de TI.

Desenvolvido por **Jordan** | Versão 1.0

Sistema portátil, rápido e inteligente para inventariar computadores corporativos (Windows 10 e Windows 11) diretamente através de um Pen Drive.

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
 |                 Desenvolvido por Jordan | Versao 1.0                 |
 +======================================================================+
```

---

## 🚀 Funcionalidades Principais

- **🎨 Interface Profissional:** Banner ASCII art estilizado com cores e logotipo corporativo InnovCore TI.
- **📌 Menu Interativo Completo:**
  - `[1]` Iniciar Novo Cadastro (Inventariar computador)
  - `[2]` Verificar Cadastros (Visualizar todos os equipamentos e abrir planilha no Excel)
  - `[3]` Excluir Cadastro (Remover registros com confirmação de segurança)
  - `[0]` Sair
- **🧠 Memória de Setores Dinâmica:** Lembra todos os setores já cadastrados e coloca no topo como opções de dígito único `[1]`, `[2]`, `[3]`, etc.
- **🔢 Patrimônio Sequencial Automático:** Gera automaticamente `PAT-0001`, `PAT-0002`, `PAT-0003`...
- **🔍 Coleta Rápida de Hardware (CIM/WMI):**
  - Número de Série da BIOS (`Win32_BIOS` e fallback para `Win32_BaseBoard`)
  - Ano de Fabricação do Equipamento
  - Marca e Modelo (`Win32_ComputerSystem`)
  - Nome do Host e Data/Hora da coleta
- **📋 Cópia Automática:** Copia o relatório completo imediatamente para a Área de Transferência (`Ctrl + V`).
- **📊 Planilha CSV Portátil:** Registra e acumula tudo no arquivo `inventario_innovcore.csv` no próprio Pen Drive (formatado em UTF-8 com BOM e delimitador `;` para compatibilidade total com o Excel PT-BR).

---

## 📁 Estrutura do Projeto

```text
projetos/
├── VERSAO_FINAL_PENDRIVE/          # Pasta com os arquivos prontos para o Pen Drive
│   ├── executar_inventario.bat    # Executável (2 cliques) que abre o PowerShell
│   ├── inventario.ps1             # Motor principal do InnovCore TI
│   └── COMO_USAR.txt              # Guia rápido de utilização
├── executar_inventario.bat
├── inventario.ps1
├── README.md
└── .gitignore
```

---

## 🚀 Como Usar no Pen Drive

1. Copie o conteúdo da pasta `VERSAO_FINAL_PENDRIVE` para a raiz do seu **Pen Drive**.
2. Conecte o Pen Drive no computador que deseja inventariar (Windows 10 ou 11).
3. Dê **2 cliques** no arquivo `executar_inventario.bat`.
4. Responda às perguntas interativas no prompt (ou pressione Enter para usar os valores sugeridos).
5. O sistema salvará as informações na planilha `inventario_innovcore.csv` e copiará o resumo para o `Ctrl + V`.

---

## 🛡️ Compatibilidade
- Windows 10 (Todas as versões)
- Windows 11 (Todas as versões)
- PowerShell 5.1 e PowerShell 7+
- Não requer instalação de programas adicionais nos computadores dos clientes.

---

*Desenvolvido com excelência por Jordan para InnovTech.*
