# Cadastro de Rede - Estruturação de Formulários

Script em Python para organizar automaticamente as respostas exportadas do Microsoft Forms e gerar uma planilha estruturada com os dados de responsáveis e servidores.

## Requisitos

* Python 3.12 ou superior
* Bibliotecas Python:

  * pandas
  * openpyxl
  * numpy

## Configuração do ambiente

1. Clone ou baixe o projeto.

2. Entre na pasta do projeto:

```
cd restruc
```

3. Crie o ambiente virtual:

```
python -m venv venv
```

4. Ative o ambiente virtual:

Windows (PowerShell):

```
venv\Scripts\activate
```

5. Instale as dependências:

```
pip install pandas openpyxl numpy
```

## Como usar

1. Exporte as respostas do Microsoft Forms para Excel.

2. Coloque o arquivo na pasta do projeto.

Exemplo:

```
respostas.xlsx
```

3. Execute o script:

```
python restruc.py
```

## Resultado

O script irá gerar automaticamente um novo arquivo Excel organizado:

```
cadastro_rede_organizado.xlsx
```

A planilha conterá:

* Responsável do setor
* Matrícula
* Login de rede
* Email institucional
* Coordenação
* Dados dos servidores

## Estrutura esperada do formulário

O formulário deve conter duas seções:

### Dados do Responsável

* Nome do responsável
* Matrícula
* Login de rede
* Email institucional
* Coordenação

### Dados dos Servidores

* Nome do servidor
* Matrícula do servidor
* Login de rede
* Email institucional
* Coordenação do servidor

## Tecnologias utilizadas

* Python
* pandas
* openpyxl
* Microsoft Forms
