import pandas as pd
import os
from PIL import Image as PILImage

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
print("\nColunas encontradas:")
for col in df.columns:
    print(f"  - '{col}'")

# ========== CRIAR WORKBOOK COM XLSXWRITER ==========
writer = pd.ExcelWriter('cadastro_rede_tamanho_fixo.xlsx', engine='xlsxwriter')
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

if logo_existe:
    with PILImage.open(caminho_logo) as img:
        largura_original, altura_original = img.size
        print(f"\n📏 Tamanho original da imagem: {largura_original}x{altura_original}")

# ========== CRIAR ABAS ==========
departamentos = df['Departamento'].unique()
print(f"\n📊 Departamentos: {list(departamentos)}")

for depto in departamentos:
    dados_depto = df[df['Departamento'] == depto].copy()
    nome_aba = str(depto)[:31]
    
    worksheet = workbook.add_worksheet(nome_aba)
    
    print(f"✅ Criando aba: {nome_aba}")
    
    # ========== CONFIGURAÇÃO DAS COLUNAS ==========
    worksheet.set_column('A:A', 40)
    worksheet.set_column('B:B', 20)
    worksheet.set_column('C:C', 20)
    worksheet.set_column('D:D', 45)
    worksheet.set_column('E:E', 25)
    
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
                # ===== TAMANHO DESEJADO (220x70) =====
                largura_desejada = 220
                altura_desejada = 70
                
                # Calcular escala para manter proporção
                with PILImage.open(caminho_logo) as img:
                    largura_original, altura_original = img.size
                
                escala_largura = largura_desejada / largura_original
                escala_altura = altura_desejada / altura_original
                escala = min(escala_largura, escala_altura)
                
                # Margens
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
                
                print(f"   ✅ Logo inserida - Tamanho: ~{largura_desejada}x{altura_desejada}")
                print(f"   ✅ Margens - Left: {margin_left}px, Top: {margin_top}px")
                
            except Exception as e:
                print(f"   ⚠️ Erro ao inserir logo: {e}")
        
        linha += 3
        
        # ========== DADOS DO RESPONSÁVEL ==========
        worksheet.merge_range(linha, 0, linha, 4, "DADOS DO RESPONSÁVEL DO SETOR", subtitulo_format)
        linha += 1
        
        dados_responsavel = [
            ("Nome:", primeiro['Nome do responsável']),
            ("Matrícula:", primeiro['Matrícula do responsável']),
            ("Login de Rede:", primeiro['Login de rede do responsável']),
            ("E-mail institucional:", primeiro['E-mail institucional do responsável']),
            ("Departamento:", primeiro['Departamento']),
            ("Quantidade de servidores:", len(dados_resp))
        ]
        
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

print(f"\n{'='*50}")
print(f"✅ SUCESSO! Arquivo: Cadastro_de_rede.xlsx")
print(f"{'='*50}")
print(f"\n🎨 Configurações:")
print(f"   - Tamanho da imagem: ~220x70 pixels")
print(f"   - Margem Left: 15px")
print(f"   - Margem Top: 10px")