# 📄 Gerador de Relatórios

Aplicação em Python desenvolvida para automatizar a criação de relatórios individuais e consolidados em formato Word (`.docx`), integrando suporte a leitura de planilhas de dados e interface gráfica moderna.

---

## 🚀 Funcionalidades Principais

* **Geração Manual:** Formulário direto para criação de relatórios personalizados com títulos, autores, resumos e listas de tópicos.
* **Processamento em Lote:** Leitura de planilhas Excel (`.xlsx`) e arquivos CSV (`.csv`) para geração automática de centenas de documentos.
* **Relatórios Consolidados:** Agrupamento de múltiplos registos num único documento estruturado.

---

## 🛠️ Tecnologias Utilizadas

* **Linguagem:** Python 3.x
* **Interface Gráfica:** `tkinter` / `ttk`
* **Manipulação de Documentos:** `python-docx`
* **Processamento de Dados:** `pandas`, `openpyxl`

---

## 📊 Histórico e Evolução das Versões

A tabela abaixo resume a evolução do projeto ao longo do desenvolvimento:

| Versão | Interface | Formato de Saída | Leitura de Dados | Motor de Geração |
| :--- | :--- | :--- | :--- | :--- |
| **V1.0** | Linha de Comandos (CLI) | `.txt` | Manual (`input`) | Escrita direta de texto |
| **V2.0** | GUI Nativa (`tkinter`) | `.docx` | Formulario | Estruturação OpenXML + `zipfile` |
| **V3.0** | GUI com Progresso | `.docx` | Excel / CSV (`pandas`) | Estruturação OpenXML + `zipfile` |
| **V4.0** | GUI Dark Mode | `.docx` | Excel / CSV (`pandas`) | Biblioteca `python-docx` |

---

### 🔹 Versão 1.0 — Módulo Inicial em Linha de Comandos (CLI)
* **Descrição:** Primeira implementação funcional voltada para testes de lógica.
* **Formato:** Geração de ficheiros de texto simples (`.txt`).
* **Características:**
  * Entrada de dados interativa via terminal (`input()`).
  * Nomeação dinâmica de ficheiros com base na semana e na data atual.

---

### 🔹 Versão 2.0 — Interface Gráfica e Transição para `.docx`
* 💡 *Melhorias em relação à V1:*
  * **Evolução da Interface:** Substituição do terminal por uma janela gráfica desenvolvida com `tkinter`.
  * **Novo Formato:** Transição de ficheiros de texto puro (`.txt`) para documentos nativos do Word (`.docx`).
  * **Manipulação de XML:** Construção direta do pacote OpenXML (`document.xml`, `_rels`) compactado via `zipfile`, eliminando a necessidade de dependências externas de terceiros.
  * **Listas Dinâmicas:** Adição de campos para inserir e remover itens da lista em tempo real.
  * **Abertura Automática:** Integração com o Windows Explorer para destacar o ficheiro criado ao finalizar.

---

### 🔹 Versão 3.0 — Processamento em Lote e Automação com Pandas
* 💡 *Melhorias em relação à V2:*
  * **Importação de Dados:** Integração com a biblioteca `pandas` para leitura automatizada de tabelas Excel e CSV.
  * **Modos de Operação Duplos:**
    * *Individual:* Criação de 1 documento `.docx` separado para cada linha da planilha.
    * *Consolidado:* Compilação de todos os registos num único documento Word.
  * **Gestão de Tarefas Longas:** Implementação de barra de progresso (`ttk.Progressbar`), etiquetas de estado e botão para **cancelar** o processamento em lote a qualquer momento.
  * **Organização de Saída:** Criação automática do diretório `relatorios_gerados/` para armazenamento centralizado.

---

### 🔹 Versão 4.0 — Redesign Dark Mode e Integração `python-docx`
* 💡 *Melhorias em relação à V3:*
  * **Refatoração do Motor do Word:** Substituição da montagem manual em XML pela biblioteca oficial `python-docx`, garantindo maior estabilidade, formatação limpa e redução de erros de compatibilidade.
  * **Interface Dark Mode:** Redesign completo da interface com uma paleta escura(Dark Mode), fontes personalizadas e botões estilizados.
  * **Centralização Automática:** A janela da aplicação é calculada e posicionada automaticamente no centro do ecrã ao iniciar.
  * **Validação de Formulário:** Adição de alertas de aviso quando campos obrigatórios (como o Título) não são preenchidos.

---

## 💻 Como Executar o Projeto

1. **Clonar o repositório:**
   ```bash
   git clone [https://github.com/luizfeliphecontatoprofissional-ui/ProjetoRelatorio.git](https://github.com/luizfeliphecontatoprofissional-ui/ProjetoRelatorio.git)
   cd ProjetoRelatorio
   
---

###
2. **Entrar no Directório do Projeto**

A sigla `cd` significa *Change Directory* (Mudar de Diretório). Após o `git clone`, o computador cria uma nova pasta chamada `ProjetoRelatorio`. É necessário entrar dentro dessa pasta pelo terminal para que os comandos seguintes saibam onde encontrar os arquivos do programa.

```bash
cd ProjetoRelatorio
```

---
3. **Instalar as Dependências do Python**

* `python-docx`: Responsável pela criação e formatação dos documentos `.docx`.
* `pandas`: Utilizado na leitura e manipulação das planilhas de dados em lote.
* `openpyxl`: Permite que o `pandas` consiga interpretar os arquivos no formato `.xlsx`.

```bash
pip install python-docx pandas openpyxl
```
---

4. ### **Executar a Aplicação**

```bash
python ProjetoRelatorioV4.py
```
