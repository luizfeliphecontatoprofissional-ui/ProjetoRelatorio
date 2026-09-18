import tkinter as tk
import zipfile
import os
from datetime import datetime


def gerar_docx_nativo():
    titulo = entry_titulo.get()
    autor = entry_autor.get()
    resumo = text_resumo.get("1.0", tk.END).strip()
    itens = lista_itens.get(0, tk.END)
    data_hora_atual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    #1.Monta a estrutura de texto em XML que o Word e o Docs entendem
    xml_document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <!-- Título do Relatório -->
        <w:p>
            <w:r>
                <w:rPr><w:b/><w:sz w:val="36"/><w:color w:val="1F4E79"/></w:rPr>
                <w:t>Título do Relatório: {titulo}</w:t>
            </w:r>
        </w:p>
        <w:p/>
        
        <w:p>
            <w:r>
                <w:rPr><w:i/><w:sz w:val="20"/><w:color w:val="595959"/></w:rPr>
                <w:t>Relatório Gerado às: {data_hora_atual}</w:t>
            </w:r>
        </w:p>
        
        <w:p>
            <w:r>
                <w:rPr><w:b/><w:sz w:val="22"/><w:color w:val="595959"/></w:rPr>
                <w:t>Autor: {autor}</w:t>
            </w:r>
        </w:p>
        <w:p/>
        
        <!-- Seção Resumo -->
        <w:p>
            <w:r>
                <w:rPr><w:b/><w:sz w:val="26"/><w:color w:val="595959"/></w:rPr>
                <w:t>Resumo</w:t>
            </w:r>
        </w:p>
'''
    for linha in resumo.split('\n'):
        if linha.strip():
            xml_document += f'''        <w:p>
            <w:r><w:t>{linha}</w:t></w:r>
        </w:p>
'''
    xml_document += '''        <w:p/>
        <w:p>
            <w:r>
                <w:rPr><w:b/><w:sz w:val="26"/><w:color w:val="595959"/></w:rPr>
                <w:t>Itens Adicionados</w:t>
            </w:r>
        </w:p>
'''
    for item in itens:
        xml_document += f'''        <w:p>
            <w:r><w:t>•  {item}</w:t></w:r>
        </w:p>
'''

    xml_document += '''    </w:body>
</w:document>'''

    xml_content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''

    xml_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

    nome_base = "relatorio"
    extensao = ".docx"
    nome_arquivo = f"{nome_base}{extensao}"
    contador = 1

    while os.path.exists(nome_arquivo):
        contador += 1
        nome_arquivo = f"{nome_base} ({contador}){extensao}"



    with zipfile.ZipFile(nome_arquivo, 'w', zipfile.ZIP_DEFLATED) as docx:
        docx.writestr('[Content_Types].xml', xml_content_types)
        docx.writestr('_rels/.rels', xml_rels)
        docx.writestr('word/document.xml', xml_document)

    print(f"Arquivo '{nome_arquivo}' criado com sucesso!")

    # Seleciona o arquivo na pasta do Windows automaticamente
    os.system(f'explorer /select,"{os.path.abspath(nome_arquivo)}"')


#Configuração da Janela (Tkinter NATIVO)
janela = tk.Tk()
janela.title("Gerador de Relatórios Word (.docx)")
janela.geometry("800x720")

tk.Label(janela, text="Título do Relatório:").pack(pady=2)
entry_titulo = tk.Entry(janela, width=80)
entry_titulo.pack(pady=2)

tk.Label(janela, text="Autor / Responsável:").pack(pady=2)
entry_autor = tk.Entry(janela, width=80)
entry_autor.pack(pady=2)

tk.Label(janela, text="Resumo das Atividades:").pack(pady=2)
text_resumo = tk.Text(janela, height=3, width=60)
text_resumo.pack(pady=2)

tk.Label(janela, text="Novo Item a Ser Adiconado:").pack(pady=2)
entry_item = tk.Entry(janela, width=80)
entry_item.pack(pady=2)


def add_item():
    if entry_item.get():
        lista_itens.insert(tk.END, entry_item.get())
        entry_item.delete(0, tk.END)


tk.Button(janela, text="Adicionar Item", command=add_item).pack(pady=4)
lista_itens = tk.Listbox(janela, width=35, height=10)
lista_itens.pack(pady=10)

def remove_item():
    selecao = lista_itens.curselection()
    if selecao:
        lista_itens.delete(selecao)


tk.Button(janela, text="Remover Item Selecionado", command=remove_item, bg="#B22222", fg="white").pack(pady=4)
tk.Button(janela, text="Gerar Documento Word(.docx)", command=gerar_docx_nativo, bg="#2b579a", fg="white").pack(
    pady=10)


janela.mainloop()
