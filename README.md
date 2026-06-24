# 🤖 Fiscal Service Automation (GissOnline Itu)

Uma ferramenta robusta de automação web projetada para realizar a emissão em massa de Notas Fiscais de Serviço Eletrônicas (NFS-e) no portal GissOnline da Prefeitura de Itu, eliminando a digitação manual de relatórios em planilhas Excel. 

Criado com foco em estabilidade, produtividade e fluidez.

---

## 🌟 Funcionalidades Principais

- **Automação Segura (Playwright):** Interação em tempo real com sites complexos e modais pesados, com simulação em nível de keystroke para driblar plugins anti-robô e validações rígidas do Angular.
- **Interface Gráfica Embutida (Web UI):** Um servidor backend em Python/Flask provém uma interface web em Light/Dark mode para acompanhamento visual do processo.
- **Batimento contra Planilhas Excel:** Lê arquivos `clientes.xlsx` e automaticamente processa CNDs, Valores, Descrições e Impostos.
- **Tratamento de Exceções "Anti-Travamento":** Sistema inteligente de ByPass para Pop-ups indesejáveis de comunicados da GISS.
- **Relatórios de Finalização:** Gera automaticamente um `.txt` tabular das emissões efetuadas exibido diretamente na Web Interface.
- **Janela Isolada do Navegador:** A interface abre em uma **nova janela** dedicada (Opera, Chrome ou Edge), sem interferir nas abas que você já está usando.
- **Preenchimento Automático dos Novos Campos Tributários:** Compatível com a atualização de Junho/2026 do GissOnline, preenchendo automaticamente os campos de Código Indicador da Operação, Classificação Tributária e CST-IBS/CBS.

## 🛠️ Tecnologias Utilizadas

- **Backend:** Python + Flask
- **Automação Web:** Playwright (sync_api)
- **Manipulação de Dados:** Pandas
- **Frontend UI:** HTML5, CSS3, e Vanilla JS dinâmico
- **Controle de Execução:** Multithreading com intercepção de inputs via interface gráfica.

## 🚀 Como Executar Localmente

### Pré-requisitos:
Assegure-se de que os pacotes do arquivo `requirements.txt` estão instalados. Os principais são:
```bash
pip install playwright flask pandas openpyxl python-dotenv
playwright install chromium
```

### Inicializando o Robô:
Para abrir a interface da automação com todas as opções:
1. Execute o arquivo `./START_AUTOMATION.bat` ou o atalho configurado.
2. O servidor local iniciará e abrirá automaticamente uma **nova janela** do navegador em `http://localhost:5000`.
3. Uma interface em Tabela será mostrada onde você poderá confirmar e preencher as Alíquotas globais caso haja necessidade de reavaliação.

> **Dica:** O robô detecta automaticamente Opera, Opera GX, Chrome e Edge, sempre abrindo em janela isolada para não atrapalhar sua navegação.

### Estrutura do Excel (`clientes.xlsx`):
Garante a importação das seguintes colunas obrigatórias:
* `Nome da empresa`
* `CNPJ`
* `VALOR`
* `DESCRIÇÃO DO SERVIÇO`

> _Qualquer variação de notação de valor `420`, `420.0`, `R$ 4.000,00` ou `4,20` é internamente tratada e convertida adequadamente antes de acoplar a injeção nas caixas de texto com máscaras PT-BR._

## 📋 Fluxo de Emissão (por nota)

| Passo | Campo | Valor Padrão |
|-------|-------|-------------|
| 1 | Serviço / Atividade | 17.19 - Atividades de Contabilidade |
| 2 | NBS | 1.1302.21.00 - Serviços de contabilidade |
| 2b | Código Indicador da Operação | 030101 - Estabelecimento do fornecedor |
| 2c | Classificação Tributária | 200052 - Prestação de serviços de profissões intelectuais |
| — | CST-IBS/CBS | *(preenchido automaticamente)* |
| 3 | Tomador | CNPJ da planilha |
| 4 | Valor do Serviço | Valor da planilha |
| 5 | Discriminação | Descrição da planilha |
| 6 | PIS/COFINS | 00 - Nenhum |
| 7 | **Próximo** | — |
| 8 | Alíquota ISS | Definida na interface |
| 9 | **Próximo** → **Concluir** | — |

## 🔧 Configuração Opcional: Domínio Customizado

Por padrão o robô roda em `http://localhost:5000`. Para usar o domínio `http://robo.itugiss:5000`:

1. Execute o script `update_hosts.ps1` **como Administrador**, ou
2. Adicione manualmente no arquivo `C:\Windows\System32\drivers\etc\hosts`:
```
127.0.0.1 robo.itugiss
```

> ⚠️ Alguns antivírus bloqueiam alterações no arquivo `hosts`. Caso não funcione, use `localhost` normalmente.

## 📌 Status
Emissão Batch **`Estável`** — Atualizado para o layout de Junho/2026 do GissOnline. Processando lotes com suporte a Bypass, Retry e Recovery state na extração dos dados (Pasta Evidências).
