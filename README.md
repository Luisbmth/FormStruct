# CADASTRO DE REDE - GUIA RÁPIDO

Script em Python para organizar automaticamente as respostas exportadas do Microsoft Forms 
e gerar uma planilha estruturada com os dados de responsáveis e servidores.

================================================================================
                              GUIA DE INSTALAÇÃO
================================================================================

1. REQUISITOS DO SISTEMA
--------------------------------------------------------------------------------
- Windows, Linux ou MacOS
- Python 3.8 ou superior instalado
- Conexão com internet (apenas para instalar as dependências)

2. VERIFICAR SE O PYTHON ESTÁ INSTALADO
--------------------------------------------------------------------------------
Abra o terminal (Prompt de Comando, PowerShell ou Terminal) e digite:

   python --version

Se aparecer algo como "Python 3.8.x" ou superior, está ok.
Se não estiver instalado, baixe em: https://www.python.org/downloads/

3. BAIXAR O PROJETO
--------------------------------------------------------------------------------
Crie uma pasta para o projeto, exemplo: C:\cadastro_rede
Coloque o arquivo do script (ex: cadastro.py) dentro desta pasta.

4. ABRIR O TERMINAL NA PASTA DO PROJETO
--------------------------------------------------------------------------------
Windows: Abra a pasta, clique no campo de endereço, digite "cmd" e pressione Enter
Linux/Mac: Abra o terminal e navegue até a pasta com o comando "cd"

5. CRIAR AMBIENTE VIRTUAL (recomendado)
--------------------------------------------------------------------------------
Digite no terminal:

   python -m venv venv

6. ATIVAR O AMBIENTE VIRTUAL
--------------------------------------------------------------------------------
Windows (Prompt de Comando):
   venv\Scripts\activate.bat

Windows (PowerShell):
   venv\Scripts\activate

Linux/Mac:
   source venv/bin/activate

Quando ativado, aparecerá (venv) no início da linha do terminal.

7. INSTALAR AS DEPENDÊNCIAS
--------------------------------------------------------------------------------
Com o ambiente virtual ativado, digite:

   pip install pandas openpyxl numpy xlsxwriter Pillow

Aguarde a instalação de todas as bibliotecas.

8. PREPARAR OS ARQUIVOS NECESSÁRIOS
--------------------------------------------------------------------------------
Na mesma pasta do script, você precisa ter:

   ✅ OBRIGATÓRIO:
      - Dados do Responsável do Setor.xlsx (exportado do Microsoft Forms)
   
   ❌ OPCIONAL:
      - niteroi.png (logo que aparece no cabeçalho da planilha)

================================================================================
                              COMO RODAR O PROJETO
================================================================================

1. EXPORTAR OS DADOS DO MICROSOFT FORMS
--------------------------------------------------------------------------------
- Acesse seu formulário no Microsoft Forms
- Clique em "Respostas" → "Abrir no Excel" (ou "Ver no Excel")
- Salve o arquivo com o nome: Dados do Responsável do Setor.xlsx
- Mova este arquivo para a mesma pasta do script

2. EXECUTAR O SCRIPT
--------------------------------------------------------------------------------
No terminal (com o ambiente virtual ativado), digite:

   python cadastro.py

(Substitua "cadastro.py" pelo nome real do seu arquivo)

3. ACOMPANHAR A EXECUÇÃO
--------------------------------------------------------------------------------
O script mostrará mensagens como:

   📁 Arquivos encontrados na pasta:
      - cadastro.py
      - Dados do Responsável do Setor.xlsx
   
   📋 Colunas encontradas:
     1. 'Id'
     2. 'Hora de início'
     ...
   
   🎯 Colunas de servidor encontradas:
     nome: 10 colunas
   
   🔄 Normalizando dados...
   ✅ Dados normalizados: XX servidores encontrados
   
   📊 Departamentos: ['infra', 'dti', ...]
   ✅ Criando aba: infra com X servidores
   ...
   
   ✅ Arquivo gerado: cadastro_rede.xlsx

4. RESULTADO FINAL
--------------------------------------------------------------------------------
Será criado um arquivo chamado:
   cadastro_rede.xlsx

Este arquivo contém:
- Uma aba para cada departamento
- Dados completos do responsável
- Lista de todos os servidores cadastrados

================================================================================
                          SOLUÇÃO DE PROBLEMAS COMUNS
================================================================================

PROBLEMA: "python não é reconhecido"
SOLUÇÃO: Instale o Python e marque a opção "Add Python to PATH"

PROBLEMA: "Arquivo não encontrado"
SOLUÇÃO: Verifique se o arquivo se chama exatamente "Dados do Responsável do Setor.xlsx"

PROBLEMA: Erro "xlsxwriter not found"
SOLUÇÃO: pip install xlsxwriter

PROBLEMA: Erro "PIL not found"
SOLUÇÃO: pip install Pillow

PROBLEMA: Apenas 1 servidor aparece
SOLUÇÃO: Verifique se as colunas no Excel estão como "Nome do servidor1", "Nome do servidor2", etc.

PROBLEMA: Erro com a logo
SOLUÇÃO: Remova o arquivo niteroi.png ou corrija a imagem

================================================================================
                          COMANDOS RÁPIDOS (RESUMO)
================================================================================

# Criar ambiente virtual
python -m venv venv

# Ativar (Windows PowerShell)
venv\Scripts\activate

# Instalar dependências
pip install pandas openpyxl numpy xlsxwriter Pillow

# Executar
python cadastro.py

================================================================================
                                FIM DO GUIA
================================================================================