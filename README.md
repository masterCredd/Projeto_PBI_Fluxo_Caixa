# Projeto_PBI_Fluxo_Caixa

Este dashboard de fluxo de caixa utilizando o Power BI.
Exploraremos como importar e transformar dados financeiros para construir
visualizações que facilitem o acompanhamento e a análise do fluxo de caixa da
sua empresa. Este curso é ideal para contadores, analistas financeiros, gestores
e qualquer pessoa interessada em aprimorar suas habilidades em Business
 Intelligence e gestão financeira.

## importação de tabelas

1. ➡️ Setores
2. ➡️ Relatório CC P1 ao P5

### Código de importação de Dados

Param font ➡️caminho_fonte={caminha independente da maquina usada}

1.➡️ Dados_pasta

```powerquery
let
    Fonte = Folder.Files(font)
in
    Fonte
```

O código se conecta a uma pasta definida na variável `font`, cria uma tabela com
detalhes sobre cada arquivo encontrado nela (nome, extensão, data de criação,
conteúdo binário, etc.) e exibe essa tabela como saída. É o primeiro passo comum
para combinar múltiplos arquivos (como planilhas de Excel ou CSVs) de uma mesma
pasta.

2.➡️ Centro de Custo

```powerquery
let
    Fonte = Dados_pasta,
        Dados_Tb = Table.AddColumn(
        Fonte, "Personalizar",
        each Excel.Workbook([Content])
    ),
    Tb_CC = Dados_Tb{
        [#"Folder Path" = "D:\Projetos logistica pbi\Projeto_PBI_Fluxo_Caixa\Dados\",
        Name = "Setores.xlsx"]
    }[Personalizar],
    CC_Sheet = Tb_CC{
        [Item = "CC", Kind = "Sheet"]
    }[Data],
    #"Cabeçalhos Promovidos" = Table.PromoteHeaders(
        CC_Sheet, [PromoteAllScalars = true]
    ),
    #"Tipo Alterado" = Table.TransformColumnTypes(
        #"Cabeçalhos Promovidos", {
            {"Centro de Custo", Int64.Type}, 
            {"Setor", type text}
        }
    )
in
    #"Tipo Alterado"
```

O código automatiza a tarefa de:
Encontrar um arquivo Excel específico (Setores.xlsx) em uma pasta.
Abrir uma planilha específica (CC) dentro desse arquivo.
Formatar os dados, transformando a primeira linha em cabeçalho e corrigindo os
tipos das colunas.

3.➡️ Relatório

```powerquery
let
    Fonte = Dados_pasta,
    #"Linhas Filtradas" = Table.SelectRows(Fonte, 
        each ([Name] <> "Setores.xlsx")
    ),
    #"Personalização Adicionada" = Table.AddColumn(
        #"Linhas Filtradas", "Personalizar", each Excel.Workbook([Content])
    ),
    #"Outras Colunas Removidas" = Table.SelectColumns(
        #"Personalização Adicionada", {"Personalizar"}
    ),
    #"Personalizar Expandido" = Table.ExpandTableColumn(
        #"Outras Colunas Removidas", "Personalizar", {"Data"}, {"Data"}
    ),
    #"Data Expandido" = Table.ExpandTableColumn(
        #"Personalizar Expandido",
        "Data",
        {"Column1", "Column2", "Column3", "Column4", 
            "Column5", "Column6", "Column7", "Column8", "Column9"
        },
        {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", 
            "Column7", "Column8", "Column9"
        }
    ),
    #"Cabeçalhos Promovidos" = Table.PromoteHeaders(#"Data Expandido", 
        [PromoteAllScalars = true]
    ),
    #"Linhas Cabeçalho" = Table.SelectRows(
        #"Cabeçalhos Promovidos", each ([Data] <> "Data")
    ),
    #"Texto Extraído Após o Delimitador" = Table.TransformColumns(
        #"Linhas Cabeçalho", {{"Centro de Custo", 
        each Text.AfterDelimiter(_, "CC "), type text}}
    ),
    #"Tipo Alterado" = Table.TransformColumnTypes(
        #"Texto Extraído Após o Delimitador",
        {
            {"Data", type date},
            {"Conta Contábil", Int64.Type},
            {"Tipo Movimentação", type text},
            {"Classificação", type text},
            {"Documento Fiscal", Int64.Type},
            {"Centro de Custo", Int64.Type},
            {"Status", type text},
            {"Valor", Currency.Type},
            {"Saldo", Currency.Type}
        }
    )
in
    #"Tipo Alterado"
```

O script automatiza o processo de:
Ignorar um arquivo de "parâmetros" `(Setores.xlsx)`.
Pegar todos os outros arquivos Excel de uma pasta.
Juntar (empilhar) os dados de todos eles em uma tabela única.
Limpar a tabela resultante, removendo cabeçalhos repetidos e formatando colunas
específicas.
Garantir que todos os dados estejam no formato correto para análise.

### Visão de Relacionamento das tabelas

![alt text](Relacionamento%20das%20tabelas.png)

### Medidas em Dax

1. ➡️ Total valores

```dax
Total Valores = SUMX('Relatório','Relatório'[Valor])
```

Esta fórmula calcula a soma total da coluna `'Valor'` percorrendo a tabela
`'Relatório'` linha por linha.
2. ➡️ Total de Receita(Entradas)

```dax
Receita =
 CALCULATE(
    [Total Valores],
    'Relatório'[Tipo Movimentação] IN{"Entradas"}
)
```

Esta fórmula calcula o valor da medida `[Total Valores]`, mas aplicando um
filtro para que o cálculo considere apenas as linhas onde a
 `'Tipo Movimentação'` é `'Entradas'`.
3. ➡️ Despesas

```dax
Despesas = 
CALCULATE(
    [Total Valores],
    'Relatório'[Tipo Movimentação] IN{"Saídas"}
)
```

Esta fórmula calcula o valor da medida `[Total Valores]`, mas aplicando um
filtro para que o cálculo considere apenas as linhas onde a
`'Tipo Movimentação'` é `'Saídas'`.

4.➡️ Saldo

```dax
Saldo = SUMX('Relatório','Relatório'[Saldo])
```

Esta fórmula percorre a tabela `'Relatório'` linha por linha e soma todos os
valores que encontra na coluna `'Saldo'`.

