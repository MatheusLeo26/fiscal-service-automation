import pandas as pd
import os

data = [
    ["42333077000123", "ARTIGOS ESPORTIVOS ITU LTDA", 420.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["39720442000448", "ARTIGOS ESPORTIVOS SOROCABA LTDA", 525.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["58974452000154", "CONDOMINIO TERRAS DE SANTA ROSA", 2700.00, "Serviços de Contabilidade e Departamento Pessoal Contagem de Cartão de Ponto. R$300,00"],
    ["33726493000109", "CPP SERVICOS EMPRESARIAIS LTDA", 360.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["22456111000140", "FATIMA ARAUJO DERMOFISIO LTDA", 700.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["03068171000140", "FLOR DA TERRA BRASIL COMERCIO E EVENTOS LTDA", 450.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["13878479000105", "G F IMPERMEABILIZACAO LTDA", 710.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["33133154000100", "HARAS CAPRICORNIO PECUARIA LTDA", 800.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["17571016000157", "IGREJA DE DEUS NO BRASIL EM ITU", 400.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["11278019000110", "JMX COMERCIO DE IMPORTACAO E EXPORTACAO", 440.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["18180336000149", "JSG REPRESENTACAO COMERCIAL LTDA", 250.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["41302048000131", "MERF COMERCIO IMPORTACAO E EXPORTACAO LTDA", 250.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["47416003000336", "MINAMO EMPREENDIMENTOS HOTELEIROS AGROP LTDA", 1000.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["18123181000109", "MORAES & SANTOS LTDA", 650.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["08313036000137", "PRADO CONSTRUCOES LTDA", 1300.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["28412741000107", "PRADO TECNOLOGIA EM CONSTRUCAO LTDA", 1290.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["48994164000108", "PRIMEIRA IGREJA BATISTA EM ITU", 400.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["51834618000198", "PSC EMPR SPE LTDA", 650.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["29867745000134", "RM DE SOUSA TURISMO", 400.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["09534864000168", "VANESSA ZIMBRES MARTINS", 500.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["67216689000167", "VERA COMERCIO DE TEHAS E TIJOLOS LTDA", 500.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["25309242000192", "VILTEMAR PEREIRA DE OLIVEIRA GESSO", 600.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["14272581000125", "VIRG FLORESTAL LTDA", 350.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["46853469000174", "W.R COMERCIO DE MATERIAIS NOVOS E USADOS EM GERAL LTDA", 650.00, "Serviços de Contabilidade e Departamento Pessoal"],
    ["07708532000127", "W.R SIQUEIRA & CIA LTDA", 980.00, "Serviços de Contabilidade e Departamento Pessoal"]
]

df = pd.DataFrame(data, columns=["CNPJ", "Nome da empresa", "VALOR", "Discriminação dos Serviços"])
current_dir = os.path.dirname(os.path.abspath(__file__))
excel_path = os.path.join(current_dir, "clientes.xlsx")
df.to_excel(excel_path, index=False)
print(f"Arquivo clientes.xlsx criado com sucesso em: {excel_path}")
