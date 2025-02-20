# Sistema Interativo de Clusterização de Dados em Excel

Este repositório contém um sistema interativo desenvolvido em Python para leitura, análise e clusterização de dados provenientes de planilhas Excel. O sistema foi projetado para ser amigável tanto para usuários técnicos quanto para leigos, oferecendo uma interface interativa com caixas de diálogo para seleção de arquivos e salvamento, além de prompts no terminal para configurar parâmetros.

---

## Sumário

- [Introdução](#introdução)
- [Funcionalidades](#funcionalidades)
- [Conceitos Básicos de Clusterização](#conceitos-básicos-de-clusterização)
- [Métodos de Clusterização Disponíveis](#métodos-de-clusterização-disponíveis)
  - [KMeans](#kmeans)
  - [DBSCAN](#dbscan)
  - [Agglomerative Clustering](#agglomerative-clustering)
- [Como o Sistema Funciona](#como-o-sistema-funciona)
  - [Seleção e Leitura do Arquivo Excel](#seleção-e-leitura-do-arquivo-excel)
  - [Identificação do Cabeçalho e Filtragem de Dados](#identificação-do-cabeçalho-e-filtragem-de-dados)
  - [Filtragem por Data](#filtragem-por-data)
  - [Seleção de Colunas](#seleção-de-colunas)
  - [Conversão e Preparação dos Dados](#conversão-e-preparação-dos-dados)
  - [Seleção e Execução dos Algoritmos de Clusterização](#seleção-e-execução-dos-algoritmos-de-clusterização)
  - [Visualização dos Resultados](#visualização-dos-resultados)
  - [Salvamento do Resultado](#salvamento-do-resultado)
- [Instalação e Dependências](#instalação-e-dependências)
- [Como Executar](#como-executar)
- [Licença](#licença)
- [Considerações Finais](#considerações-finais)

---

## Introdução

O sistema foi criado para facilitar a análise de dados empresariais presentes em arquivos Excel. Ele permite que o usuário:

- Selecione o arquivo e a planilha desejada utilizando caixas de diálogo (Tkinter).
- Escolha interativamente a linha que contém os nomes das colunas (cabeçalho) para que o pandas interprete os dados corretamente.
- Filtre os dados por um intervalo de datas, caso existam colunas que contenham datas.
- Selecione as colunas que serão utilizadas na análise de clusterização, podendo fazê-lo através dos nomes ou dos números das colunas.
- Converta os dados selecionados para valores numéricos, removendo linhas com valores inválidos.
- Escolha entre diferentes algoritmos de clusterização (KMeans, DBSCAN e Agglomerative Clustering) e configure seus parâmetros de forma interativa.
- Visualize os resultados da clusterização em gráficos interativos, onde os eixos são rotulados com os nomes originais das colunas quando possível.
- Salve os resultados (o DataFrame com os clusters atribuídos) utilizando uma caixa de diálogo para selecionar o local e o nome do arquivo.

---

## Funcionalidades

- **Interface Interativa:**  
  Utiliza caixas de diálogo (Tkinter) para facilitar a seleção do arquivo Excel e o salvamento dos resultados.

- **Leitura Personalizada do Excel:**  
  Permite a visualização prévia do conteúdo da planilha, escolha interativa da linha de cabeçalho e remoção automática de colunas indesejadas (como as "Unnamed").

- **Filtragem por Data:**  
  Detecta colunas que podem conter datas e permite que o usuário filtre o dataset por um intervalo específico (data inicial e data final).

- **Seleção Flexível de Colunas:**  
  Permite que o usuário escolha as colunas para análise tanto pelo nome quanto pelo número correspondente na listagem.

- **Clusterização com Múltiplos Algoritmos:**  
  Suporta três métodos de clusterização:
  - **KMeans:** Agrupa os dados em um número pré-definido de clusters.
  - **DBSCAN:** Agrupa os dados com base na densidade, identificando outliers.
  - **Agglomerative Clustering:** Realiza clusterização hierárquica, permitindo a escolha de diferentes métodos de ligação (linkage).

- **Visualização dos Gráficos:**  
  Os gráficos gerados exibem os nomes das colunas selecionadas ou os componentes resultantes do PCA, facilitando a interpretação dos clusters.

- **Salvamento dos Resultados:**  
  Utiliza uma caixa de diálogo para que o usuário escolha onde salvar o arquivo Excel contendo os resultados da clusterização.

---

## Conceitos Básicos de Clusterização

A **clusterização** é uma técnica de aprendizado não supervisionado que consiste em agrupar dados de forma que itens do mesmo grupo (cluster) sejam mais similares entre si do que com itens de outros grupos. Essa técnica é amplamente utilizada em análise de dados para:

- Identificar padrões ou grupos ocultos.
- Segmentar clientes ou produtos.
- Analisar grandes conjuntos de dados para insights estratégicos.

Em termos simples, o objetivo da clusterização é descobrir uma estrutura natural nos dados sem usar rótulos pré-definidos.

---

## Métodos de Clusterização Disponíveis

### KMeans

- **Como Funciona:**  
  O algoritmo KMeans divide os dados em *K* clusters, tentando minimizar a variância dentro de cada cluster. Cada ponto de dado é atribuído ao cluster com o centroide (média) mais próximo.
  
- **Parâmetros Importantes:**  
  - Número de clusters (K).

- **Vantagens:**  
  - Simples e eficiente para conjuntos de dados grandes.
  
- **Limitações:**  
  - Requer a definição prévia do número de clusters.
  - Sensível a outliers.

---

### DBSCAN

- **Como Funciona:**  
  O DBSCAN (Density-Based Spatial Clustering of Applications with Noise) agrupa pontos que estão densamente agrupados, definindo clusters com base na densidade dos pontos. Pontos que não se encaixam em nenhum cluster são considerados ruído.

- **Parâmetros Importantes:**  
  - `eps`: Distância máxima para considerar dois pontos como vizinhos.
  - `min_samples`: Número mínimo de pontos necessários para formar um cluster.

- **Vantagens:**  
  - Não requer a definição do número de clusters.
  - Capaz de identificar outliers.

- **Limitações:**  
  - Pode ser sensível à escolha dos parâmetros `eps` e `min_samples`.

---

### Agglomerative Clustering

- **Como Funciona:**  
  Trata-se de uma técnica de clusterização hierárquica que começa com cada ponto em seu próprio cluster e, em seguida, vai mesclando clusters de forma iterativa. O método de ligação (linkage) determina como a distância entre clusters é calculada.

- **Parâmetros Importantes:**  
  - Número de clusters desejado.
  - Método de linkage: `ward`, `complete`, `average` ou `single`.

- **Vantagens:**  
  - Não requer a especificação do número de clusters para construir a hierarquia (mas sim para “cortar” a árvore).
  - Fornece uma estrutura hierárquica que pode ser analisada de diferentes formas.

- **Limitações:**  
  - Pode ser computacionalmente custoso para grandes conjuntos de dados.
  - O método “ward” requer dados com distância euclidiana.

---

## Como o Sistema Funciona

### Seleção e Leitura do Arquivo Excel

- **Caixa de Diálogo para Seleção de Arquivo:**  
  Utilizando o Tkinter, o sistema abre uma janela para que o usuário selecione o arquivo Excel desejado, eliminando a necessidade de digitar o caminho completo.

- **Escolha da Planilha:**  
  O sistema lista todas as planilhas disponíveis no arquivo e permite que o usuário escolha a que deseja analisar.

---

### Identificação do Cabeçalho e Filtragem de Dados

- **Visualização Prévia:**  
  O sistema carrega uma pré-visualização (10 primeiras linhas) do arquivo sem cabeçalho para que o usuário identifique onde estão os nomes das colunas.

- **Definição do Cabeçalho:**  
  O usuário informa, de forma interativa, qual a linha (em índice 0-based) que contém os nomes das colunas. Se não houver cabeçalho, o sistema lê os dados sem essa informação.

- **Remoção de Colunas Indesejadas:**  
  Colunas geradas automaticamente (com nomes como "Unnamed") são removidas para limpar o DataFrame.

---

### Filtragem por Data

- **Detecção de Colunas com Datas:**  
  O sistema verifica se existem colunas com dados de data (seja porque o tipo já é datetime ou porque a conversão de um sample é bem-sucedida).

- **Escolha da Coluna de Data:**  
  Se houver mais de uma coluna candidata, o usuário pode escolher qual utilizar, digitando o nome ou o número correspondente à listagem.

- **Definição do Intervalo de Datas:**  
  O usuário informa a data inicial e final no formato `YYYY-MM-DD`, e o sistema filtra o dataset para trabalhar apenas com as linhas que se encontram dentro desse intervalo.

---

### Seleção de Colunas

- **Seleção Flexível:**  
  O sistema exibe uma lista numerada com as colunas disponíveis e permite que o usuário selecione as colunas para clusterização tanto digitando os números correspondentes quanto os nomes.

---

### Conversão e Preparação dos Dados

- **Conversão para Numérico:**  
  As colunas selecionadas são convertidas para valores numéricos, e linhas com dados inválidos (NaN) são removidas.

- **Escalonamento dos Dados:**  
  Os dados são normalizados usando `StandardScaler` do scikit-learn, garantindo que as diferentes escalas não afetem a análise de clusterização.

---

### Seleção e Execução dos Algoritmos de Clusterização

- **Escolha dos Algoritmos:**  
  O usuário pode escolher entre KMeans, DBSCAN e Agglomerative Clustering. As opções podem ser selecionadas por número ou por nome.

- **Configuração dos Parâmetros:**  
  Para cada algoritmo, o sistema solicita os parâmetros necessários:
  - Para **KMeans**: número de clusters.
  - Para **DBSCAN**: valores para `eps` e `min_samples`.
  - Para **Agglomerative Clustering**: número de clusters e o método de linkage (aceitando números ou nomes para as opções).

- **Execução dos Algoritmos:**  
  Os algoritmos são executados sobre os dados escalonados e os rótulos dos clusters são atribuídos ao DataFrame original.

---

### Visualização dos Resultados

- **Gráficos Interativos:**  
  Os resultados da clusterização são exibidos em gráficos utilizando o matplotlib:
  - **1D:** Se apenas uma coluna for utilizada, o gráfico exibe o nome da coluna no eixo Y.
  - **2D:** Se duas colunas forem utilizadas, os eixos X e Y exibem os nomes originais das colunas.
  - **Mais de 2 Colunas:** Caso seja necessário aplicar o PCA para redução para 2D, o título do gráfico inclui os nomes das colunas utilizadas.

---

### Salvamento do Resultado

- **Caixa de Diálogo para Salvamento:**  
  Ao final do processo, o sistema pergunta se o usuário deseja salvar o DataFrame com os clusters. Se sim, uma caixa de diálogo é aberta para que o usuário escolha o local e o nome do arquivo Excel de saída.

---

## Instalação e Dependências

### Requisitos

- **Python 3.6+**
- **Bibliotecas Python:**
  - pandas
  - numpy
  - matplotlib
  - scikit-learn
  - tkinter (geralmente incluído com a instalação padrão do Python)

### Instalação das Dependências

Use o pip para instalar as bibliotecas necessárias:

```bash
pip install pandas numpy matplotlib scikit-learn
