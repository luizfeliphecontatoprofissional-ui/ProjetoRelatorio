import tkinter as tk
import zipfile
import os
import random
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
import pandas as pd

dados = []
setores = ['Tecnologia da Informação', 'Recursos Humanos', 'Financeiro', 'Operações', '']

for i in range(1, 1001):
    dados.append({
        'Nome': f'Colaborador {i}',
        'Setor': random.choice(setores),
        'Resumo': f'Relatório individual de desempenho e atividades referentes ao período do colaborador {i}.\nExecução de tarefas e rotinas diárias.'
    })

df = pd.DataFrame(dados)
nome_arquivo = 'planilha_teste_extensa.xlsx'
df.to_excel(nome_arquivo, index=False)
print(f"Planilha de teste '{nome_arquivo}' criada com sucesso contendo {len(df)} registros!")

PASTA_DESTINO = "relatorios_gerados"


def criar_documento_word(titulo, autor, resumo, itens=[], identificador=""):
    os.makedirs(PASTA_DESTINO, exist_ok=True)
    data_hora_atual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    xml_document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
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

        <w:p>
            <w:r>
                <w:rPr><w:b/><w:sz w:val="26"/><w:color w:val="595959"/></w:rPr>
                <w:t>Resumo</w:t>
            </w:r>
        </w:p>
'''
    for linha in str(resumo).split('\n'):
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

    nome_base = f"relatorio_{identificador}" if identificador else "relatorio"
    extensao = ".docx"
    caminho_arquivo = os.path.join(PASTA_DESTINO, f"{nome_base}{extensao}")
    contador = 1

    while os.path.exists(caminho_arquivo):
        contador += 1
        caminho_arquivo = os.path.join(PASTA_DESTINO, f"{nome_base}_({contador}){extensao}")

    with zipfile.ZipFile(caminho_arquivo, 'w', zipfile.ZIP_DEFLATED) as docx:
        docx.writestr('[Content_Types].xml', xml_content_types)
        docx.writestr('_rels/.rels', xml_rels)
        docx.writestr('word/document.xml', xml_document)

    return caminho_arquivo


def criar_documento_consolidado(dados):
    os.makedirs(PASTA_DESTINO, exist_ok=True)
    data_hora_atual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    total_registros = len(dados)

    xml_corpo = f'''
        <w:p>
            <w:r>
                <w:rPr><w:b/><w:sz w:val="36"/><w:color w:val="1F4E79"/></w:rPr>
                <w:t>Relatório Consolidado de Atividades</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:rPr><w:i/><w:sz w:val="20"/><w:color w:val="595959"/></w:rPr>
                <w:t>Gerado em: {data_hora_atual} | Total de Registros: {total_registros}</w:t>
            </w:r>
        </w:p>
        <w:p/>
'''

    for index, linha in dados.iterrows():
        titulo = (linha.get('nome') or linha.get('Nome') or linha.get('Titulo') or linha.get('Título') or f'Relatório_{index + 1}')
        autor = (linha.get('setor') or linha.get('Setor') or linha.get('Autor') or 'Autor Não Informado')
        resumo = (linha.get('Resumo') or linha.get('resumo') or f"Colaborador(a) {titulo} vinculado(a) ao setor {autor}.")

        xml_corpo += f'''
        <w:p>
            <w:r>
                <w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="1F4E79"/></w:rPr>
                <w:t>#{index + 1} - {titulo}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:rPr><w:b/><w:sz w:val="22"/><w:color w:val="595959"/></w:rPr>
                <w:t>Autor: {autor}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:rPr><w:b/><w:sz w:val="22"/><w:color w:val="595959"/></w:rPr>
                <w:t>Resumo:</w:t>
            </w:r>
        </w:p>
'''
        for l in resumo.split('\n'):
            if l.strip():
                xml_corpo += f'''        <w:p>
            <w:r><w:t>{l}</w:t></w:r>
        </w:p>
'''
        xml_corpo += '''        <w:p>
            <w:r><w:rPr><w:color w:val="D3D3D3"/></w:rPr><w:t>--------------------------------------------------------------------------------</w:t></w:r>
        </w:p>
        <w:p/>
'''

    xml_document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        {xml_corpo}
    </w:body>
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

    caminho_arquivo = os.path.join(PASTA_DESTINO, "Relatorio_Consolidado.docx")
    contador = 1
    while os.path.exists(caminho_arquivo):
        contador += 1
        caminho_arquivo = os.path.join(PASTA_DESTINO, f"Relatorio_Consolidado_({contador}).docx")

    with zipfile.ZipFile(caminho_arquivo, 'w', zipfile.ZIP_DEFLATED) as docx:
        docx.writestr('[Content_Types].xml', xml_content_types)
        docx.writestr('_rels/.rels', xml_rels)
        docx.writestr('word/document.xml', xml_document)

    return caminho_arquivo


def gerar_docx_nativo():
    titulo = entry_titulo.get()
    autor = entry_autor.get()
    resumo = text_resumo.get("1.0", tk.END).strip()
    itens = lista_itens.get(0, tk.END)

    caminho_gerado = criar_documento_word(titulo, autor, resumo, itens)
    print(f"Arquivo '{caminho_gerado}' criado na pasta '{PASTA_DESTINO}'!")


def importar_e_processar():
    global cancelamento_solicitado
    cancelamento_solicitado = False

    caminho = filedialog.askopenfilename(
        title="Selecione a planilha de dados",
        filetypes=[("Arquivos Excel/CSV", "*.xlsx *.csv")]
    )
    if not caminho:
        return

    lbl_status.config(text="Aguardando o processamento...", fg="#1F4E79")
    janela.update_idletasks()

    try:
        if caminho.endswith('.csv'):
            try:
                dados = pd.read_csv(caminho, sep=';', encoding='latin1')
            except Exception:
                dados = pd.read_csv(caminho)
        else:
            dados = pd.read_excel(caminho)

        dados = dados.fillna('')
        total_linhas = len(dados)

        resposta = messagebox.askyesnocancel(
            "Modo de Geração",
            f"Sua planilha tem {total_linhas} registros.\n\n"
            "• Clique em SIM para gerar 1 ARQUIVO INDIVIDUAL por linha.\n"
            "• Clique em NÃO para gerar 1 ÚNICO RELATÓRIO CONSOLIDADO.\n"
            "• Clique em CANCELAR para abortar."
        )

        if resposta is None:
            lbl_status.config(text="")
            return

        if resposta is True:
            print(f"\n--- Gerando {total_linhas} relatórios individuais ---")
            barra_progresso['maximum'] = total_linhas

            btn_cancelar.pack(before=lbl_status, pady=10)
            arquivo_gerados = 0

            for index, linha in dados.iterrows():
                if cancelamento_solicitado:
                    messagebox.showwarning(
                        "Processamento Interrompido",
                        f"O processamento foi cancelado pelo usuário.\n\n"
                        f"Foram gerados {arquivo_gerados} de {total_linhas} relatórios."
                    )
                    break

                titulo = (linha.get('nome') or linha.get('Nome') or linha.get('Titulo') or linha.get('Título') or f'Relatório_{index + 1}')
                autor = (linha.get('setor') or linha.get('Setor') or linha.get('Autor') or 'Autor Não Informado')
                resumo = (linha.get('Resumo') or linha.get('resumo') or f"Colaborador(a) {titulo} vinculado(a) ao setor {autor}.")

                criar_documento_word(
                    titulo=titulo,
                    autor=autor,
                    resumo=resumo,
                    itens=[],
                    identificador=f"linha_{index + 1}"
                )

                arquivo_gerados += 1
                barra_progresso['value'] = index + 1
                lbl_status.config(text=f"Gerando relatório {index + 1} de {total_linhas}...")

                janela.update()

            if not cancelamento_solicitado:
                messagebox.showinfo("Sucesso!", f"{total_linhas} arquivos individuais gerados com sucesso na pasta '{PASTA_DESTINO}'!")

        else:
            print("\n--- Gerando Relatório Consolidado Único ---")
            caminho_consolidado = criar_documento_consolidado(dados)
            messagebox.showinfo("Sucesso!", f"Relatório consolidado gerado com sucesso!\n\nSalvo em: {caminho_consolidado}")

        #os.system(f'explorer "{os.path.abspath(PASTA_DESTINO)}"')

    except Exception as e:
        messagebox.showerror("Erro de Processamento", f"Falha ao processar planilha:\n{str(e)}")

    finally:
        barra_progresso['value'] = 0
        lbl_status.config(text="")

        btn_cancelar.pack_forget()
        cancelamento_solicitado = False


janela = tk.Tk()
janela.title("Gerador de Relatórios")
janela.geometry("600x820")

btn_importar = tk.Button(
    janela,
    text="Importar Planilha (Individual ou Consolidado)",
    command=importar_e_processar,
    bg="#2E8B57", fg="white"
)
btn_importar.pack(pady=15)

tk.Label(janela, text="Título do Relatório:").pack(pady=2)
entry_titulo = tk.Entry(janela, width=80)
entry_titulo.pack(pady=2)

tk.Label(janela, text="Autor / Responsável:").pack(pady=2)
entry_autor = tk.Entry(janela, width=80)
entry_autor.pack(pady=2)

tk.Label(janela, text="Resumo das Atividades:").pack(pady=2)
text_resumo = tk.Text(janela, height=3, width=60)
text_resumo.pack(pady=2)

tk.Label(janela, text="Novo Item a Ser Adicionado:").pack(pady=2)
entry_item = tk.Entry(janela, width=80)
entry_item.pack(pady=2)


def add_item():
    if entry_item.get():
        lista_itens.insert(tk.END, entry_item.get())
        entry_item.delete(0, tk.END)


tk.Button(janela, text="Adicionar Item", command=add_item).pack(pady=4)
lista_itens = tk.Listbox(janela, width=35, height=8)
lista_itens.pack(pady=5)


def remove_item():
    selecao = lista_itens.curselection()
    if selecao:
        lista_itens.delete(selecao)


tk.Button(janela, text="Remover Item Selecionado", command=remove_item, bg="#B22222", fg="white").pack(pady=4)
tk.Button(janela, text="Gerar Documento Word Manual (.docx)", command=gerar_docx_nativo, bg="#2b579a", fg="white").pack(pady=10)

cancelamento_solicitado = False


def solicitar_cancelamento():
    global cancelamento_solicitado
    cancelamento_solicitado = True
    lbl_status.config(text="Cancelando... Aguarde a linha atual terminar.", fg="red")


btn_cancelar = tk.Button(
    janela,
    text="Cancelar Processamento",
    command=solicitar_cancelamento,
    bg='#de0600',
    fg='#ffffff',
    activebackground="#A94442",
    activeforeground="white",
    font=('Arial', 10, 'bold')
)

lbl_status = tk.Label(janela, text="", fg='#595959')
lbl_status.pack(pady=5)

tk.Label(janela, text="Progresso da Geração:").pack(pady=2)
barra_progresso = ttk.Progressbar(janela, orient='horizontal', length=400, mode='determinate')
barra_progresso.pack(pady=5)

janela.mainloop()