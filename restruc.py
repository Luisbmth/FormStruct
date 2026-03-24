import pandas as pd
import os
from PIL import Image as PILImage

# ========== NOME DO ARQUIVO ==========
arquivo = "Dados do Responsável do Setor.xlsx"

if not os.path.exists(arquivo):
    print(f"❌ ERRO: Arquivo '{arquivo}' não encontrado!")
    exit()

# ========== LER PLANILHA ==========
df = pd.read_excel(arquivo)

# ========== FUNÇÃO PARA ENCONTRAR TODAS AS COLUNAS DE SERVIDORES ==========
def encontrar_colunas_servidor(df):
    """Encontra todas as colunas relacionadas a servidores (até 30 servidores)"""
    
    padroes = {
        'nome': ['Nome do servidor'] + [f'Nome do servidor{i}' for i in range(1, 30)],
        'matricula': ['Matrícula do servidor'] + [f'Matrícula do servidor{i}' for i in range(1, 30)],
        'login': ['Login de rede'] + [f'Login de rede{i}' for i in range(1, 30)],
        'email': ['E-mail institucional'] + [f'E-mail institucional{i}' for i in range(1, 30)],
        'coordenacao': ['Coordenação do servidor'] + [f'Coordenação do servidor{i}' for i in range(1, 30)]
    }
    
    colunas_existentes = {}
    for tipo, lista_colunas in padroes.items():
        colunas_existentes[tipo] = [col for col in lista_colunas if col in df.columns]
    
    return colunas_existentes

# ========== FUNÇÃO PARA OBTER SUBSSECRETARIA ==========
def obter_subsecretaria(row):
    """Retorna o nome da subsecretaria se a resposta for Sim"""
    subordinado = row.get('Este departamento está subordinado à alguma Subsecretaria?')
    
    if pd.notna(subordinado) and str(subordinado).strip().lower() in ['sim', 'yes', 's', 'sim']:
        return row.get('Informe o nome da Subsecretaria')
    return None

# ========== ENCONTRAR COLUNAS DE SERVIDOR ==========
colunas_servidor = encontrar_colunas_servidor(df)

# ========== NORMALIZAR DADOS (CRIAR UMA LINHA POR SERVIDOR) ==========
linhas_normalizadas = []

for idx, row in df.iterrows():
    subsecretaria = obter_subsecretaria(row)
    
    dados_base = {
        'Nome do responsável': row.get('Nome do responsável'),
        'Matrícula do responsável': row.get('Matrícula do responsável'),
        'Login de rede do responsável': row.get('Login de rede do responsável'),
        'E-mail institucional do responsável': row.get('E-mail institucional do responsável'),
        'Departamento': row.get('Departamento'),
        'Subordinado à Subsecretaria?': row.get('Este departamento está subordinado à alguma Subsecretaria?'),
        'Subsecretaria': subsecretaria
    }
    
    for i in range(len(colunas_servidor['nome'])):
        if i == 0:
            col_nome = 'Nome do servidor'
            col_matricula = 'Matrícula do servidor'
            col_login = 'Login de rede'
            col_email = 'E-mail institucional'
            col_coordenacao = 'Coordenação do servidor'
        else:
            col_nome = f'Nome do servidor{i}'
            col_matricula = f'Matrícula do servidor{i}'
            col_login = f'Login de rede{i}'
            col_email = f'E-mail institucional{i}'
            col_coordenacao = f'Coordenação do servidor{i}'
        
        if col_nome not in df.columns:
            continue
            
        nome = row.get(col_nome)
        
        if pd.isna(nome) or not str(nome).strip():
            continue
        
        matricula = row.get(col_matricula) if col_matricula in df.columns else None
        login = row.get(col_login) if col_login in df.columns else None
        email = row.get(col_email) if col_email in df.columns else None
        coordenacao = row.get(col_coordenacao) if col_coordenacao in df.columns else None
        
        linha_servidor = dados_base.copy()
        linha_servidor.update({
            'Nome do servidor': nome,
            'Matrícula do servidor': matricula,
            'Login de rede': login,
            'E-mail institucional': email,
            'Coordenação do servidor': coordenacao
        })
        
        linhas_normalizadas.append(linha_servidor)

df_normalizado = pd.DataFrame(linhas_normalizadas)

if len(df_normalizado) == 0:
    df_normalizado = df
    colunas_originais = ['Nome do responsável', 'Matrícula do responsável', 
                          'Login de rede do responsável', 'E-mail institucional do responsável',
                          'Departamento', 'Este departamento está subordinado à alguma Subsecretaria?',
                          'Informe o nome da Subsecretaria',
                          'Nome do servidor', 'Matrícula do servidor',
                          'Login de rede', 'E-mail institucional', 'Coordenação do servidor']
    
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

for depto in departamentos:
    dados_depto = df_normalizado[df_normalizado['Departamento'] == depto].copy()
    nome_aba = str(depto)[:31]
    
    worksheet = workbook.add_worksheet(nome_aba)
    
    worksheet.set_column('A:A', 40)
    worksheet.set_column('B:B', 35)
    worksheet.set_column('C:C', 20)
    worksheet.set_column('D:D', 45)
    worksheet.set_column('E:E', 25)
    
    linha = 0
    
    responsaveis = dados_depto['Nome do responsável'].unique()
    
    for responsavel in responsaveis:
        dados_resp = dados_depto[dados_depto['Nome do responsável'] == responsavel]
        primeiro = dados_resp.iloc[0]
        
        worksheet.set_row(linha, 70)
        worksheet.merge_range(linha, 0, linha, 4, "CADASTRO DE REDE", titulo_format)
        
        if logo_existe:
            try:
                with PILImage.open(caminho_logo) as img:
                    largura_original, altura_original = img.size
                
                largura_desejada = 720
                altura_desejada = 390
                
                escala_largura = largura_desejada / largura_original
                escala_altura = altura_desejada / altura_original
                escala = min(escala_largura, escala_altura)
                
                worksheet.insert_image(
                    linha, 0, 
                    caminho_logo,
                    {
                        'x_offset': 15,
                        'y_offset': 10,
                        'x_scale': escala,
                        'y_scale': escala,
                        'object_position': 1
                    }
                )
            except Exception:
                pass
        
        linha += 3
        
        worksheet.merge_range(linha, 0, linha, 4, "DADOS DO RESPONSÁVEL DO SETOR", subtitulo_format)
        linha += 1
        
        dados_responsavel = [
            ("Nome:", primeiro['Nome do responsável']),
            ("Matrícula:", primeiro['Matrícula do responsável']),
            ("Login de Rede:", primeiro['Login de rede do responsável']),
            ("E-mail institucional:", primeiro['E-mail institucional do responsável']),
            ("Departamento:", primeiro['Departamento']),
            ("Quantidade de servidores cadastrados:", len(dados_resp))
        ]
        
        if 'Subsecretaria' in primeiro and pd.notna(primeiro['Subsecretaria']):
            dados_responsavel.append(("Subsecretaria:", primeiro['Subsecretaria']))
        
        for campo, valor in dados_responsavel:
            worksheet.write(linha, 0, campo, campo_format)
            worksheet.write(linha, 1, valor, valor_format)
            worksheet.write(linha, 2, "", valor_format)
            worksheet.write(linha, 3, "", valor_format)
            worksheet.write(linha, 4, "", valor_format)
            linha += 1
        
        linha += 1
        
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

writer.close()
print(f"✅ Arquivo gerado: cadastro_rede.xlsx")