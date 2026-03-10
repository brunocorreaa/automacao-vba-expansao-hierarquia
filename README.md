# Automação VBA: Expansão de Hierarquia e Normalização de Clusters

## 📌 Sobre o Projeto

Este projeto consiste em uma ferramenta desenvolvida em **VBA (Visual Basic for Applications)** para automatizar a expansão de bases de dados matriciais no Excel. 

O objetivo principal é transformar um formato de entrada ("Input"), onde as quantidades de itens estão distribuídas por colunas de clusters (A, B, C, D, E, F), em uma lista normalizada ("Output"), onde cada item é repetido em linhas individuais de acordo com sua respectiva volumetria.

### 💡 Por que este projeto é importante?

Em análise de dados, bases matriciais (colunas para cada categoria) são fáceis de ler, mas difíceis de processar em ferramentas como Power BI ou SQL. Esta automação:
* **Padroniza os dados:** Converte formatos complexos em tabelas limpas e prontas para análise.
* **Elimina o trabalho manual:** O que levaria horas para ser feito via "copia e cola" é executado em segundos.
* **Garante Integridade:** O código conta com uma trava de segurança que impede a execução caso existam valores negativos nas colunas de cluster.

---

## 🛡️ Disclaimer e Privacidade (Dados Anonimizados)

* **Privacidade:** Todos os dados visíveis nas planilhas (Setores, Grupos, Classes e Códigos) foram totalmente **anonimizados**. As nomenclaturas foram alteradas para termos genéricos para proteger a propriedade intelectual e dados sensíveis.

---

## 🛠️ Detalhes das Visões (Excel)

### 1. Planilha de Entrada (Input)
A aba `Input` contém a hierarquia de produtos (Setor, Grupo, Classe, Subclasse) e as quantidades destinadas a cada Cluster (A até F).

<img width="1782" height="767" alt="Captura de tela 2026-03-10 202143" src="https://github.com/user-attachments/assets/f4d578d3-1bc4-4e2e-b58d-4040b0d45008" />

*Figura 1: Interface de entrada com os dados matriciais.*

### 2. Planilha de Saída (Output)
A aba `Output` é preenchida automaticamente pelo script. O código varre cada linha do Input, identifica a quantidade em cada cluster e gera o número correspondente de linhas, atribuindo a letra do cluster na última coluna.

<img width="865" height="757" alt="output" src="https://github.com/user-attachments/assets/48a62d6f-9d12-4506-8162-ce7e2a16fcec" />

*Figura 2: Resultado da expansão de dados pronta para uso.*

---

## 💻 Estrutura do Código VBA

O script utiliza loops aninhados para garantir que nenhuma informação seja perdida durante a expansão. Abaixo, uma visão geral da lógica:

1. **Limpeza:** O intervalo de saída é limpo antes de cada nova execução.
2. **Validação:** Um loop verifica o intervalo `L3:Q` em busca de números negativos para evitar erros lógicos.
3. **Processamento:** - Varre as linhas da planilha de entrada.
   - Para cada cluster (A-F) com valor maior que zero, ele inicia um loop interno.
   - Copia os dados da hierarquia e atribui a etiqueta do cluster correspondente.
