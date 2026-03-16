import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image
import os

# arquivo exportado do Forms
arquivo = "Dados do Responsável do Setor.xlsx"

# ler planilha
df = pd.read_excel(arquivo)

# MOSTRAR COLUNAS PARA CONFIRMAR
print("Colunas encontradas:")
for col in df.columns:
    print(f"  - '{col}'")

# Criar novo workbook
wb = Workbook()

# Remover a planilha padrão (será criada depois)
wb.remove(wb.active)

# ========== ESTILOS COM CALIBRI ==========
# Cores
cinza_claro = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
cinza_escuro = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
branco = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

# Fontes Calibri
titulo_font = Font(name='Calibri', size=24, bold=True)
subtitulo_font = Font(name='Calibri', size=11, bold=True)
bold_font = Font(name='Calibri', size=11, bold=True)
normal_font = Font(name='Calibri', size=11)

# Bordas completas
borda_completa = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

# Sem bordas
sem_borda = Border(
    left=Side(style=None),
    right=Side(style=None),
    top=Side(style=None),
    bottom=Side(style=None)
)

# ========== LOGO ==========
caminho_logo = "logo_prefeitura.png"
logo_existe = os.path.exists(caminho_logo)

if not logo_existe:
    print("⚠️ Arquivo da logo não encontrado. As abas serão criadas sem logo.")

# ========== CRIAR UMA ABA PARA CADA DEPARTAMENTO ==========
departamentos = df['Departamento'].unique()

print(f"\n📊 Departamentos encontrados: {list(departamentos)}")

for depto in departamentos:
    # Filtrar dados do departamento
    dados_depto = df[df['Departamento'] == depto].copy()
    
    # Criar nome da aba (limitar a 31 caracteres)
    nome_aba = str(depto)[:31]
    ws = wb.create_sheet(title=nome_aba)
    
    print(f"✅ Criando aba: {nome_aba} - {len(dados_depto)} registro(s)")
    
    # ========== AJUSTAR LARGURA DAS COLUNAS ==========
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 45
    ws.column_dimensions['E'].width = 25
    
    linha = 1  # COMEÇAR NA LINHA 1
    
    # ========== INSERIR LOGO (se existir) ==========
    if logo_existe:
        try:
            img = Image(caminho_logo)
            img.width = 60
            img.height = 60
            img.anchor = 'A1'
            ws.add_image(img)
            ws.row_dimensions[1].height = 50
        except Exception as e:
            print(f"   ⚠️ Erro ao inserir logo: {e}")
    
    # Se tem logo, o título começa na linha 2, senão na linha 1
    if logo_existe:
        linha_titulo = 2
    else:
        linha_titulo = 1
    
    # Agrupar por responsável dentro do departamento
    responsaveis = dados_depto['Nome do responsável'].unique()
    
    for responsavel in responsaveis:
        dados_resp = dados_depto[dados_depto['Nome do responsável'] == responsavel]
        primeiro = dados_resp.iloc[0]
        
        # ========== TÍTULO PRINCIPAL NA LINHA CORRETA ==========
        ws.merge_cells(f'A{linha_titulo}:E{linha_titulo}')
        celula_titulo = ws.cell(linha_titulo, 1, "CADASTRO DE REDE")
        celula_titulo.font = titulo_font
        celula_titulo.alignment = Alignment(horizontal='center', vertical='center')
        celula_titulo.fill = cinza_claro
        celula_titulo.border = sem_borda
        ws.row_dimensions[linha_titulo].height = 40
        
        linha = linha_titulo + 2  # Pular 2 linhas após o título
        
        # ========== DADOS DO RESPONSÁVEL ==========
        ws.merge_cells(f'A{linha}:E{linha}')
        celula_subtitulo = ws.cell(linha, 1, "DADOS DO RESPONSÁVEL DO SETOR")
        celula_subtitulo.font = subtitulo_font
        celula_subtitulo.alignment = Alignment(horizontal='center')
        celula_subtitulo.fill = cinza_escuro
        celula_subtitulo.border = borda_completa
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
            celula_campo = ws.cell(linha, 1, campo)
            celula_campo.font = bold_font
            celula_campo.border = borda_completa
            celula_campo.fill = branco
            
            celula_valor = ws.cell(linha, 2, valor)
            celula_valor.font = normal_font
            celula_valor.border = borda_completa
            celula_valor.fill = branco
            
            for col in range(3, 6):
                celula_vazia = ws.cell(linha, col, "")
                celula_vazia.border = borda_completa
                celula_vazia.fill = branco
            
            linha += 1
        
        linha += 1
        
        # ========== TABELA DE SERVIDORES ==========
        ws.merge_cells(f'A{linha}:E{linha}')
        celula_subtitulo = ws.cell(linha, 1, "DADOS DOS SERVIDORES")
        celula_subtitulo.font = subtitulo_font
        celula_subtitulo.alignment = Alignment(horizontal='center')
        celula_subtitulo.fill = cinza_escuro
        celula_subtitulo.border = borda_completa
        linha += 1
        
        # Cabeçalho
        cabecalhos = ["Nome", "Matrícula", "Login de Rede", "E-mail Institucional", "Coordenação"]
        for col, cab in enumerate(cabecalhos, 1):
            celula = ws.cell(linha, col, cab)
            celula.font = bold_font
            celula.alignment = Alignment(horizontal='center')
            celula.fill = cinza_claro
            celula.border = borda_completa
        
        linha += 1
        
        # Dados dos servidores
        for _, row in dados_resp.iterrows():
            valores = [
                row['Nome do servidor'],
                row['Matrícula do servidor'],
                row['Login de rede'],
                row['E-mail institucional'],
                row['Coordenação do servidor']
            ]
            
            for col, val in enumerate(valores, 1):
                celula = ws.cell(linha, col, val)
                celula.font = normal_font
                celula.border = borda_completa
                celula.fill = branco
            
            linha += 1
        
        # Espaço entre responsáveis
        linha_titulo = linha + 3
        linha = linha_titulo

# Salvar arquivo
wb.save("cadastro_rede_por_departamento.xlsx")
print("\n✅ Arquivo criado: cadastro_rede_por_departamento.xlsx")
print(f"📊 Total de abas criadas: {len(departamentos)}")
print("🎨 CADASTRO DE REDE na primeira linha (ou segunda se tiver logo)")