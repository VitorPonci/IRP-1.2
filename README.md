# IRP-1.2
Macro VBA utilizada para tratamento de dados do indicador 1.2 do PMCRP
README
IRP1.2 – Proporção da produção representada pelas Bacias de Santos, Campos e Espírito Santo
Macro VBA: CalcularProporcaoIndividualPorBacia
1. Objetivo do código

A macro CalcularProporcaoIndividualPorBacia automatiza o tratamento e a consolidação dos dados de produção por bacia sedimentar para fins de monitoramento da Questão 1 do PMCRP, calculando:

os valores anuais de produção (em boe ou Mboe/d, conforme a base utilizada) das bacias Campos, Santos e Espírito Santo;

o Total Geral anual;

as proporções individuais de cada bacia em relação ao Total Geral, gerando uma tabela pronta para gráficos na aba “Graficos”.

Embora o indicador IRP1.2 seja definido como a participação conjunta das três bacias em relação ao total nacional (IRP1.1), esta macro produz um desdobramento complementar (proporções por bacia), útil para análise e visualização da contribuição relativa de cada bacia dentro do agregado.

2. Relação com o IRP1.2 e com o IRP1.1

IRP1.2 (definição):

𝐼
𝑅
𝑃
1.2
=
𝑃
𝐵
𝑆
+
𝑃
𝐵
𝐶
+
𝑃
𝐵
𝐸
𝐼
𝑅
𝑃
1.1
×
100
IRP1.2=
IRP1.1
PBS+PBC+PBE
	​

×100

onde PBS, PBC e PBE são as produções anuais das bacias de Santos, Campos e Espírito Santo, e IRP1.1 é a produção nacional total.

O que este código faz:
Este código calcula e organiza PBS, PBC e PBE, além do Total Geral na base por bacia, e gera também as proporções individuais (Campos/Total, Santos/Total, ES/Total).
Para o cálculo do IRP1.2 “final” (agregado / nacional), o resultado desta macro pode ser combinado com o IRP1.1 calculado na etapa anterior.

3. Fonte dos dados

Os dados devem ser extraídos do Boletim Mensal da Produção de Petróleo e Gás Natural (ANP), preferencialmente a edição de dezembro (encarte consolidado anual). A tabela utilizada no boletim é a de produção por bacia (ex.: “Distribuição da produção de petróleo e gás natural por bacia”).

4. Estrutura esperada na planilha (pré-requisitos)

A macro pressupõe um arquivo Excel com as seguintes abas:

Aba “Produção Por Bacia” (entrada)

Cada linha representa um registro de produção associado a um ano e a uma bacia. A macro lê as seguintes colunas:

Coluna A (1): Ano (numérico)

Coluna B (2): Nome da bacia (texto)

Coluna E (5): Produção (numérica)

Observação importante:
A macro identifica explicitamente as seguintes categorias na Coluna B:

"Total Geral"

"Campos"

"Santos"

"Espírito Santo"

Qualquer divergência de grafia, acentuação ou espaços pode impedir o cálculo correto.

Aba “Graficos” (saída)

A macro escreve uma tabela consolidada a partir da célula A1 e limpa previamente intervalos específicos.

5. O que a macro faz (passo a passo)

Define as planilhas:

Fonte: "Produção Por Bacia"

Destino: "Graficos"

Limpa resultados antigos:

A2:H1000 (conteúdo)

Coluna I (conteúdo)

Percorre a base e alimenta dicionários por ano:

dictTotalGeral(ano) → produção do “Total Geral”

dictCampos(ano) → produção da bacia Campos

dictSantos(ano) → produção da bacia Santos

dictES(ano) → produção da bacia Espírito Santo

Cria o cabeçalho na aba “Graficos”:

Ano, Campos, Santos, Espírito Santo, Total Geral,
Prop. Campos, Prop. Santos, Prop. Espírito Santo

Para cada ano:

escreve os valores absolutos de produção por bacia;

escreve o Total Geral;

calcula as proporções individuais:

Campos / Total Geral

Santos / Total Geral

ES / Total Geral

formata as proporções como percentual (0,00%)

Aplica formatação alternada (copiando formatos das linhas 2 e 3 para as demais).

6. Saída gerada

Na aba “Graficos”, a macro cria uma tabela com as colunas:

Ano

Campos

Santos

Espírito Santo

Total Geral

Prop. Campos

Prop. Santos

Prop. Espírito Santo

Esses resultados podem ser usados diretamente para:

gráficos de participação por bacia;

validação de consistência temporal;

suporte ao cálculo consolidado do IRP1.2.

7. Validação e consistência recomendadas

Após executar a macro, recomenda-se:

conferir se todos os anos esperados aparecem;

verificar se o Total Geral é não-nulo para cada ano;

verificar se as proporções estão no intervalo [0%, 100%];

checar coerência: Campos + Santos + ES ≤ Total Geral (em bases em que existam outras bacias além dessas três).

8. Limitações conhecidas

Dependência da grafia exata dos nomes das bacias.

Dependência de estrutura fixa das colunas A, B e E na aba “Produção Por Bacia”.

O cálculo final do IRP1.2 (agregado sobre IRP1.1) depende de integração com o resultado do IRP1.1, que não é executada por esta macro.

9. Código – Macro utilizada

Macro: CalcularProporcaoIndividualPorBacia
(Lembrar de manter o código no apêndice/anexo ou no repositório, conforme padronização da Nota Metodológica.)
