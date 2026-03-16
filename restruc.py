import pandas as pd

# arquivo exportado do Forms
arquivo = "teste.xlsx"

# ler planilha
df = pd.read_excel(arquivo)

# selecionar apenas as colunas importantes
df_limpo = df[[
    "nome1",
    "numero de matricula",
    "login de rede",
    "email institucional\n",
    "coordenacao"
]]

# renomear colunas para ficar organizado
df_limpo.columns = [
    "Nome",
    "Matricula",
    "Login de Rede",
    "Email Institucional",
    "Coordenacao"
]

# salvar nova planilha organizada
df_limpo.to_excel("usuarios_organizados.xlsx", index=False)

print("Tabela organizada criada: usuarios_organizados.xlsx")