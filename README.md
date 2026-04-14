# 🤖 Robô Emissor de NFS-e (Extendido - GissOnline Itu)

Uma ferramenta robusta de automação web projetada para realizar a emissão em massa de Notas Fiscais de Serviço Eletrônicas (NFS-e) no portal GissOnline da Prefeitura de Itu, eliminando a digitação manual de relatórios em planilhas Excel. 

Criado com foco em estabilidade, produtividade e fluidez.

---

## 🌟 Funcionalidades Principais

- **Automação Segura (Playwright):** Interação em tempo real com sites complexos e modais pesados, com simulação em nível de keystroke para driblar plugins anti-robô e validações rígidas do Angular.
- **Interface Gráfica Embutida (Web UI):** Um servidor backend em Python/Flask provém uma interface web em Light/Dark mode para acompanhamento visual do processo.
- **Batimento contra Planilhas Excel:** Lê arquivos `clientes.xlsx` e automaticamente processa CNDs, Valores, Descrições e Impostos.
- **Tratamento de Exceções "Anti-Travamento":** Sistema inteligente de ByPass para Pop-ups indesejáveis de comunicados da GISS.
- **Relatórios de Finalização:** Gera automaticamente um `.txt` tabular das emissões efetuadas exibido diretamente na Web Interface.
- **Servidor com Domínio Customizado:** Servido localmente simulando em host próprio (`http://robo.itugiss:5000`).

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
pip install playwright flask pandas openpyxl
playwright install chromium
```

### Inicializando o Robô:
Para abrir a interface da automação com todas as opções:
1. Execute o arquivo `./START_AUTOMATION.bat` ou o atalho configurado.
2. O servidor local fará o bind das rotas e abrirá automaticamente o seu navegador usando o domínio customizado (se ativado via `hosts`).
3. Uma interface em Tabela será mostrada onde você poderá confirmar e preencher as Alíquotas globais caso haja necessidade de reavaliação.

### Estrutura do Excel (`clientes.xlsx`):
Garante a importação das seguintes colunas obrigatórias:
* `Nome da empresa`
* `CNPJ`
* `VALOR`
* `DESCRIÇÃO DO SERVIÇO`

> _Qualquer variação de notação de valor `420`, `420.0`, `R$ 4.000,00` ou `4,20` é internamente tratada e convertida adequadamente antes de acoplar a injeção nas caixas de texto com máscaras PT-BR._

## 📌 Status
Emissão Batch **`Estável`**. Processando lotes com suporte a Bypass e Recovery state na extração dos dados (Pasta Evidências).
