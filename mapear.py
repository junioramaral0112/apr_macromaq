from docx import Document

# Carrega o seu template limpo
doc = Document("templates/T.SHE.046_Safe Job Analyses - OFICIAL PARA PREENCHIMENTO.docx")

# Vamos olhar a primeira tabela (Cabeçalho)
tabela_cabecalho = doc.tables[0]

print("--- MAPEAMENTO DA TABELA 0 (CABEÇALHO) ---")
for i, linha in enumerate(tabela_cabecalho.rows):
    for j, celula in enumerate(linha.cells):
        # Mostra o texto existente para sabermos o que tem lá dentro
        print(f"Linha [{i}], Coluna [{j}] -> Conteúdo atual: '{celula.text.strip()}'")