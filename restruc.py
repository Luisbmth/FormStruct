import pandas as pd
import os
from PIL import Image as PILImage
import re

# ========== VERIFICAR ARQUIVOS NA PASTA ==========
print("📁 Arquivos encontrados na pasta:")
for arquivo in os.listdir('.'):
    print(f"   - {arquivo}")

# ========== NOME DO ARQUIVO (COM ACENTO) ==========
arquivo = "Dados do Responsável do Setor.xlsx"

if not os.path.exists(arquivo):
    print(f"\n❌ ERRO: Arquivo '{arquivo}' não encontrado!")
    exit()

# ========== LER PLANILHA ==========
print(f"\n✅ Arquivo encontrado: {arquivo}")
df = pd.read_excel(arquivo)

# MOSTRAR COLUNAS
print("\n📋 Colunas encontradas:")
for i, col in enumerate(df.columns):
    print(f"  {i+1}. '{col}'")

# ========== FUNÇÃO PARA ENCONTRAR TODAS AS COLUNAS DE SERVIDORES ==========
def encontrar_colunas_servidor(df):
    """Encontra todas as colunas relacionadas a servidores (até 15 servidores)"""
    
    # Padrões de colunas (primeiro servidor não tem número)
    # Agora suporta até 15 servidores (0 a 14, sendo 0 sem número)
    padroes = {
        'nome': ['Nome do servidor'] + [f'Nome do servidor{i}' for i in range(1, 15)],
        'matricula': ['Matrícula do servidor'] + [f'Matrícula do servidor{i}' for i in range(1, 15)],
        'login': ['Login de rede'] + [f'Login de rede{i}' for i in range(1, 15)],
        'email': ['E-mail institucional'] + [f'E-mail institucional{i}' for i in range(1, 15)],
        'coordenacao': ['Coordenação do servidor'] + [f'Coordenação do servidor{i}' for i in range(1, 15)]
    }
    
    # Verificar quais colunas realmente existem no DataFrame
    colunas_existentes = {}
    for tipo, lista_colunas in padroes.items():
        colunas_existentes[tipo] = [col for col in lista_colunas if col in df.columns]
    
    return colunas_existentes

# ========== FUNÇÃO PARA OBTER SUBSSECRETARIA ==========
def obter_subsecretaria(row):
    """Retorna o nome da subsecretaria se a resposta for Sim"""
    subordinado = row.get('Este departamento está subordinado à alguma Subsecretaria?')
    
    # Verificar se a resposta é "Sim"
    if pd.notna(subordinado) and str(subordinado).strip().lower() in ['sim', 'yes', 's', 'sim']:
        return row.get('Informe o nome da Subsecretaria')
    return None

# ========== ENCONTRAR COLUNAS DE SERVIDOR ==========
colunas_servidor = encontrar_colunas_servidor(df)

print("\n🎯 Colunas de servidor encontradas:")
for tipo, colunas in colunas_servidor.items():
    print(f"  {tipo}: {len(colunas)} colunas - {colunas}")

# ========== NORMALIZAR DADOS (CRIAR UMA LINHA POR SERVIDOR) ==========
print("\n🔄 Normalizando dados...")

linhas_normalizadas = []

for idx, row in df.iterrows():
    # Obter subsecretaria (se houver)
    subsecretaria = obter_subsecretaria(row)
    
    # Dados do responsável (constantes para esta linha)
    dados_base = {
        'Nome do responsável': row.get('Nome do responsável'),
        'Matrícula do responsável': row.get('Matrícula do responsável'),
        'Login de rede do responsável': row.get('Login de rede do responsável'),
        'E-mail institucional do responsável': row.get('E-mail institucional do responsável'),
        'Departamento': row.get('Departamento'),
        'Subordinado à Subsecretaria?': row.get('Este departamento está subordinado à alguma Subsecretaria?'),
        'Subsecretaria': subsecretaria
    }
    
    # Verificar se tem pelo menos um servidor
    if pd.notna(row.get('Nome do servidor')) and str(row.get('Nome do servidor')).strip():
        # Processar cada servidor (baseado no número de colunas de nome encontradas)
        for i in range(len(colunas_servidor['nome'])):
            # Nome da coluna para este servidor
            if i == 0:
                col_nome = 'Nome do servidor'
                col_matricula = 'Matrícula do servidor'
                col_login = 'Login de rede'
                col_email = 'E-mail institucional'
                col_coordenacao = 'Coordenação do servidor'
            else:
                # Para i=1, col_nome = 'Nome do servidor1', etc.
                col_nome = f'Nome do servidor{i}'
                col_matricula = f'Matrícula do servidor{i}'
                col_login = f'Login de rede{i}'
                col_email = f'E-mail institucional{i}'
                col_coordenacao = f'Coordenação do servidor{i}'
            
            # Verificar se a coluna existe no DataFrame
            if col_nome not in df.columns:
                continue
                
            nome = row.get(col_nome)
            
            # Pular se estiver vazio
            if pd.isna(nome) or not str(nome).strip():
                continue
            
            # Pegar os outros dados
            matricula = row.get(col_matricula) if col_matricula in df.columns else None
            login = row.get(col_login) if col_login in df.columns else None
            email = row.get(col_email) if col_email in df.columns else None
            coordenacao = row.get(col_coordenacao) if col_coordenacao in df.columns else None
            
            # Criar linha para este servidor
            linha_servidor = dados_base.copy()
            linha_servidor.update({
                'Nome do servidor': nome,
                'Matrícula do servidor': matricula,
                'Login de rede': login,
                'E-mail institucional': email,
                'Coordenação do servidor': coordenacao
            })
            
            linhas_normalizadas.append(linha_servidor)

# Criar DataFrame normalizado
df_normalizado = pd.DataFrame(linhas_normalizadas)

print(f"✅ Dados normalizados: {len(df_normalizado)} servidores encontrados")
print(f"   Distribuídos em {len(df)} respostas de formulário")

# Se não encontrou nenhum servidor, usar o DataFrame original
if len(df_normalizado) == 0:
    print("⚠️ Nenhum servidor encontrado na normalização. Usando DataFrame original.")
    df_normalizado = df
    # No DataFrame original, pegar apenas o primeiro servidor
    colunas_originais = ['Nome do responsável', 'Matrícula do responsável', 
                          'Login de rede do responsável', 'E-mail institucional do responsável',
                          'Departamento', 'Este departamento está subordinado à alguma Subsecretaria?',
                          'Informe o nome da Subsecretaria',
                          'Nome do servidor', 'Matrícula do servidor',
                          'Login de rede', 'E-mail institucional', 'Coordenação do servidor']
    
    # Filtrar apenas colunas que existem
    colunas_existentes = [col for col in colunas_originais if col in df_normalizado.columns]
    df_normalizado = df_normalizado[colunas_existentes].copy()
    df_normalizado = df_normalizado.dropna(subset=['Nome do servidor'])

# ========== CRIAR WORKBOOK COM XLSXWRITER ==========
writer = pd.ExcelWriter('cadastro_rede.xlsx', engine='xlsxwriter')
workbook = writer.book

# ========== ESTILOS ==========
titulo_format = workbook.add_format({
    'font_name': 'Calibri',
    'font_size': 24,
    'bold': True,
    'align': 'center',
    'valign': 'vcenter',
    'bg_color': '#F2F2F2',
    'border': 0
})

subtitulo_format = workbook.add_format({
    'font_name': 'Calibri',
    'font_size': 11,
    'bold': True,
    'align': 'center',
    'valign': 'vcenter',
    'bg_color': '#D9D9D9',
    'border': 1,
    'border_color': '#000000'
})

cabecalho_format = workbook.add_format({
    'font_name': 'Calibri',
    'font_size': 11,
    'bold': True,
    'align': 'center',
    'valign': 'vcenter',
    'bg_color': '#F2F2F2',
    'border': 1,
    'border_color': '#000000'
})

campo_format = workbook.add_format({
    'font_name': 'Calibri',
    'font_size': 11,
    'bold': True,
    'align': 'left',
    'valign': 'vcenter',
    'bg_color': '#FFFFFF',
    'border': 1,
    'border_color': '#000000'
})

valor_format = workbook.add_format({
    'font_name': 'Calibri',
    'font_size': 11,
    'align': 'left',
    'valign': 'vcenter',
    'bg_color': '#FFFFFF',
    'border': 1,
    'border_color': '#000000'
})

# ========== LOGO ==========
caminho_logo = "niteroi.png"
logo_existe = os.path.exists(caminho_logo)

# ========== CRIAR ABAS POR DEPARTAMENTO ==========
departamentos = df_normalizado['Departamento'].unique()
print(f"\n📊 Departamentos: {list(departamentos)}")

for depto in departamentos:
    dados_depto = df_normalizado[df_normalizado['Departamento'] == depto].copy()
    nome_aba = str(depto)[:31]
    
    worksheet = workbook.add_worksheet(nome_aba)
    
    print(f"✅ Criando aba: {nome_aba} com {len(dados_depto)} servidores")
    
    # ========== CONFIGURAÇÃO DAS COLUNAS ==========
    worksheet.set_column('A:A', 40)   # Campo
    worksheet.set_column('B:B', 35)   # Valor
    worksheet.set_column('C:C', 20)   # Vazio
    worksheet.set_column('D:D', 45)   # Vazio
    worksheet.set_column('E:E', 25)   # Vazio
    
    linha = 0
    
    # Agrupar por responsável
    responsaveis = dados_depto['Nome do responsável'].unique()
    
    for responsavel in responsaveis:
        dados_resp = dados_depto[dados_depto['Nome do responsável'] == responsavel]
        primeiro = dados_resp.iloc[0]
        
        # ========== LINHA DO TÍTULO ==========
        worksheet.set_row(linha, 70)
        
        # ========== TÍTULO CENTRALIZADO ==========
        worksheet.merge_range(linha, 0, linha, 4, "CADASTRO DE REDE", titulo_format)
        
        # ========== INSERIR LOGO COM TAMANHO FIXO ==========
        if logo_existe:
            try:
                with PILImage.open(caminho_logo) as img:
                    largura_original, altura_original = img.size
                
                largura_desejada = 720
                altura_desejada = 390
                
                escala_largura = largura_desejada / largura_original
                escala_altura = altura_desejada / altura_original
                escala = min(escala_largura, escala_altura)
                
                margin_left = 15
                margin_top = 10
                
                worksheet.insert_image(
                    linha, 0, 
                    caminho_logo,
                    {
                        'x_offset': margin_left,
                        'y_offset': margin_top,
                        'x_scale': escala,
                        'y_scale': escala,
                        'object_position': 1
                    }
                )
            
            except Exception as e:
                print(f"   ⚠️ Erro ao inserir logo: {e}")
        
        linha += 3
        
        # ========== DADOS DO RESPONSÁVEL ==========
        worksheet.merge_range(linha, 0, linha, 4, "DADOS DO RESPONSÁVEL DO SETOR", subtitulo_format)
        linha += 1
        
        # Montar lista de dados do responsável
        dados_responsavel = [
            ("Nome:", primeiro['Nome do responsável']),
            ("Matrícula:", primeiro['Matrícula do responsável']),
            ("Login de Rede:", primeiro['Login de rede do responsável']),
            ("E-mail institucional:", primeiro['E-mail institucional do responsável']),
            ("Departamento:", primeiro['Departamento'])
        ]
        
        # Adicionar subsecretaria se existir e for Sim
        if 'Subsecretaria' in primeiro and pd.notna(primeiro['Subsecretaria']):
            dados_responsavel.append(("Subsecretaria:", primeiro['Subsecretaria']))
        
        # Mostrar APENAS a quantidade de servidores cadastrados (que já estão na tabela)
        dados_responsavel.append(("Quantidade de servidores cadastrados:", len(dados_resp)))
        
        for campo, valor in dados_responsavel:
            worksheet.write(linha, 0, campo, campo_format)
            worksheet.write(linha, 1, valor, valor_format)
            worksheet.write(linha, 2, "", valor_format)
            worksheet.write(linha, 3, "", valor_format)
            worksheet.write(linha, 4, "", valor_format)
            linha += 1
        
        linha += 1
        
        # ========== TABELA DE SERVIDORES ==========
        worksheet.merge_range(linha, 0, linha, 4, "DADOS DOS SERVIDORES", subtitulo_format)
        linha += 1
        
        cabecalhos = ["Nome", "Matrícula", "Login de Rede", "E-mail Institucional", "Coordenação"]
        for col, cab in enumerate(cabecalhos):
            worksheet.write(linha, col, cab, cabecalho_format)
        linha += 1
        
        for _, row in dados_resp.iterrows():
            valores = [
                row['Nome do servidor'],
                row['Matrícula do servidor'],
                row['Login de rede'],
                row['E-mail institucional'],
                row['Coordenação do servidor']
            ]
            for col, val in enumerate(valores):
                worksheet.write(linha, col, val, valor_format)
            linha += 1
        
        linha += 3

# Salvar arquivo
writer.close()
print(f"\n✅ Arquivo gerado: cadastro_rede.xlsx")