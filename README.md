# Algoritmo Genetico - Médias Móveis Simples

## Resumo:
Implementação de algoritmo genético para otimização de estratégia de cruzamento de médias móveis simples.

A implementação foi feita em VBA de um algoritmo genético para otimização de estratégia de cruzamento de médias móveis simples (https://www.investopedia.com/terms/s/sma.asp), processando dados de um arquivo de cotações históricas de mercado à vista que pode foi baixado so site da B3 (http://www.bmfbovespa.com.br/pt_br/servicos/market-data/historico/mercado-a-vista/cotacoes-historicas/). Na mesma página, é possível, além dos arquivos de cotação histórica, baixar o template dos dados, que serviu de base para esta implementação.

Na versão, não foi implementada ainda a função de mutação, a ser corrigido em versões futuras.

## Objetivo
O objetivo de desenvolver este código é provar a eficiência do uso de algoritmos genéticos para otimização de estratégias de trading, cujos resultados devam vir a ser publicados posteriormente.

## Algoritmo implementado
O algoritmo funciona da seguinte maneira: 

### Geração de "população" inicial
Inicialmente é gerada uma população de <i>n</i> (um dos parâmetros de entrada da rotina) indivíduos de pares de médias móveis de forma aleatória; posteriormente, será posssível carregar essa população inicial a partir de um arquivo de populações já existente.

### Seleção
É feito um backtest (teste da estratégia contra o passado, com base nos dados do arquivo de cotações históricas utilizado) de cada par de médias móveis e feita uma seleção de x% (também passada como parâmetro na chamada da rotina), descartando-se também, qualquer indivíduo cujo resultado do backtest seja negativo; aos indivíduos "sobreviventes" à seleção, é aplicada uma função de aptidão, onde serão selecionados, com maior probabilidade, os indivíduos cujo P&L (resultado do backtest da estratégia) sejam maiores. A probabilidade de seleção de cada indivíduo é definida pela seguinte fórmula:

p_sel(i) = P&L(i) / soma(P&L_total),

<p>onde:</p>
<p><b>p_sel(i)</b> = probabilidade de seleção do indivíduo i</p>
<p><b>P&L(i)</b> = ganho / perda de i</p>
<p><b>soma(P&L_total)</b> = soma do ganho de todos os indivíduos</p>



Ou seja, a probabilidade de seleção é diretamente proporcional à representatividade de ganho de cada indivíduo sobre o total. Desta forma, por exemplo, um indivíduo que represente 30% do ganho total, de forma individual, terá 30% de chance de ser selecionado para a geração de novos indivíduos.

### Recombinação <i>(crossing over)</i>:
De acordo com a função de aptidão descrita acima, uma nova população é gerada, cruzando-se os <i>"genes"</i> de cada indivíduo (ou seja, a média móvel curta e a média móvel longa).

### Iterações
A partir disso, todo o processo é repetido por um número de vezes que é definido como um dos parâmetros da função.
