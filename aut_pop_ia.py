# -*- coding: utf-8 -*-
"""
Created on Tue May 13 13:15:02 2025

# Desenvolvido por Thaina Braga – Projeto de Automação com IA (2025)
# -*- coding: utf-8 -*-
# Atualizado em 13/05/2025 - alterações prompt
# Script Unificado Compactado: Etapas 2, 3 e 4
"""

# === IMPORTS GERAIS ===

from openai import OpenAI
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2.service_account import Credentials
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
import json
import os
import io
import re
import unicodedata
from collections import defaultdict
from docx.shared import Inches

# === CONFIGURAÇÕES ===

from dotenv import load_dotenv
load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/documents"]
credenciais = ServiceAccountCredentials.from_json_keyfile_name("SUAS_CREDENCIAIS_AQUI", SCOPE)
cliente = gspread.authorize(credenciais)
planilha = cliente.open_by_url("URL_DO_SEU_FORMULÁRIO_AQUI")
aba = planilha.worksheet("Respostas ao formulário 1")
valores = aba.get_all_values()
cabecalhos = valores[0]
colunas = {h: idx + 1 for idx, h in enumerate(cabecalhos)}

# === FUNÇÕES GERAIS ===

def limpar_nome_arquivo(texto):
    return re.sub(r'[\\/*?:"<>|]', "-", texto.strip())

def extrair_file_id(link):
    padrao = r"(?:id=|/d/)([a-zA-Z0-9_-]{25,})"
    resultado = re.search(padrao, link)
    return resultado.group(1) if resultado else None

def baixar_arquivo_drive(file_id, nome_destino):
    credentials = Credentials.from_service_account_file("credentials.json", scopes=SCOPE)
    drive_service = build("drive", "v3", credentials=credentials)
    request = drive_service.files().get_media(fileId=file_id)
    fh = io.FileIO(nome_destino, "wb")
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()

# === ETAPA 2: Padronizar POPs com GPT-4o ===

def processar_pop(linha_idx, dados_formulario):
    link_arquivo = dados_formulario.get("Arquivo POP", "")
    file_id = extrair_file_id(link_arquivo)
    if not file_id:
        print(f"⚠️ Linha {linha_idx}: Nenhum ID válido no campo 'Arquivo POP'. Pulando...")
        return

    print(f"🔄 Linha {linha_idx}: Baixando arquivo (ID: {file_id})...")
    baixar_arquivo_drive(file_id, "POP_INPUT.docx")

    doc = Document("POP_INPUT.docx")
    output_folder = "imagens_pop"
    os.makedirs(output_folder, exist_ok=True)

    image_index = 1
    modified_text = []

    for para in doc.paragraphs:
        if para.text.strip():
            modified_text.append(para.text.strip())

    for rel in doc.part._rels:
        rel_obj = doc.part._rels[rel]
        if "image" in rel_obj.target_ref:
            image_name = f"IMAGEM_{image_index}.png"
            with open(os.path.join(output_folder, image_name), "wb") as f:
                f.write(rel_obj.target_part.blob)
            modified_text.append(f"[{image_name}]")
            image_index += 1

    texto_completo = "\n\n".join(modified_text)

    prompt = f"""
Você receberá o conteúdo bruto de um Procedimento Operacional Padrão (POP).

Sua tarefa é:

- NÃO OMITIR nenhuma informação existente (etapas, campos, observações, imagens).
- REESTRUTURAR o conteúdo de maneira formal, técnica e organizada.
- UTILIZAR verbos no infinitivo nas instruções (ex: iniciar, preencher, concluir).
- CONECTAR as ações usando transições claras (ex: "Após concluir...", "Em seguida...", "Retornar para...").
- EXPANDIR e DETALHAR as etapas: para cada ação, explicar o que deve ser feito, como fazer e qual é a finalidade.
- Sempre que identificar campos a serem preenchidos, apresentar cada item da lista de forma estruturada com:
  - "campo": o nome do campo (em negrito no Word)
  - "descricao": explicação do que é, como preencher e sua importância
- Sempre que identificar transações do SAP, apresentar no formato:
  - **Nome da Transação** explicação do que é, como preencher e sua importância.
- Se o procedimento envolver transações SAP, MENCIONAR o código da transação SAP no objetivo.
- AGRUPAR informações gerais que não sejam específicas de uma atividade em uma seção "Observações Gerais".
- CORRIGIR eventuais erros de ortografia, gramática e digitação.
- MELHORAR a fluidez, eliminando repetições desnecessárias e reorganizando frases, mantendo sempre o sentido original.

IMPORTANTE:
- Estruture cada etapa como uma atividade clara, separando por tópicos (Atividade 1, Atividade 2, etc.).
- Caso uma atividade tenha campos específicos, apresente-os como lista estruturada.
- Se existirem imagens, liste o nome das imagens associadas a cada atividade, mantendo a sequência lógica.
- NÃO invente informações que não existam no conteúdo enviado.
- Estamos utilizando o sheets para inputar essas informações então o número de atividades não deve exceder 10; una atividades compatíveis, se necessário.

Formato obrigatório de resposta, apenas em JSON puro:
{{
  "objetivo": "",
  "atividades": [
    {{
      "nome": "",
      "descricao_texto": "",
      "descricao_lista": [
        {{
          "campo": "",
          "descricao": ""
        }}
      ],
      "imagens": ["IMAGEM_1", "IMAGEM_2"]
    }}
  ],
  "observacoes": [""],
  "analise_melhorias": ""
}}

Conteúdo bruto:
{texto_completo}

⚠️ Retorne apenas o JSON, sem explicações adicionais, sem comentários, sem cabeçalhos extras.
"""

    resposta = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "Especialista em padronização de POPs."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.5
    )

    resposta_json = resposta.choices[0].message.content.strip()
    resposta_json = resposta_json[resposta_json.find('{'):resposta_json.rfind('}')+1]

    try:
        dados_chatgpt = json.loads(resposta_json)
    except json.JSONDecodeError:
        print(f"❌ Linha {linha_idx}: Erro ao interpretar JSON.")
        return

    atividades = dados_chatgpt.get("atividades", [])
    observacoes = dados_chatgpt.get("observacoes", [])

    dados_para_inserir = {
        "Objetivo": dados_chatgpt.get("objetivo", ""),
        "Análise de Melhorias": dados_chatgpt.get("analise_melhorias", "")
    }

    for i, atividade in enumerate(atividades):
        n = i + 1
        dados_para_inserir[f"Atividade {n}"] = atividade.get("nome", "")
        dados_para_inserir[f"Descrição {n} (texto)"] = atividade.get("descricao_texto", "")

        # Montar a lista formatada com campo em negrito
        lista_formatada = ""
        lista = atividade.get("descricao_lista", [])
        if isinstance(lista, list) and all(isinstance(item, dict) for item in lista):
            lista_formatada = "\n".join(
                [f"**{item['campo']} –** {item['descricao']}" for item in lista]
            )
        elif isinstance(lista, list):
            lista_formatada = "\n".join(lista)

        dados_para_inserir[f"Descrição {n} (lista)"] = lista_formatada
        dados_para_inserir[f"Imagens {n}"] = ", ".join(atividade.get("imagens", []))

    if observacoes:
        dados_para_inserir["Observações"] = "\n".join(observacoes)

    cabecalhos_atuais = aba.row_values(1)
    colunas = {h: idx + 1 for idx, h in enumerate(cabecalhos_atuais)}

    for cab in dados_para_inserir:
        if cab not in colunas:
            aba.update_cell(1, len(colunas) + 1, cab)
            colunas[cab] = len(colunas) + 1

    for cab, valor in dados_para_inserir.items():
        col = colunas[cab]
        aba.update_cell(linha_idx, col, valor)

    if "Status Padronização" not in colunas:
        aba.update_cell(1, len(colunas) + 1, "Status Padronização")
        colunas["Status Padronização"] = len(colunas) + 1

    aba.update_cell(linha_idx, colunas["Status Padronização"], "Não padronizado")
    print(f"✅ Linha {linha_idx} processada com sucesso.")



# === EXECUTAR ETAPA 2 ===

for i, linha in enumerate(valores[1:], start=2):
    dados = dict(zip(cabecalhos, linha))
    if not dados.get("Objetivo", "").strip() and dados.get("Arquivo POP", "").startswith("http"):
        processar_pop(i, dados)

# === ETAPA 3: Montagem do Documento com Modelo ===

# Funções auxiliares para montagem do Word

# Lista de campos fixos que você quer sempre em negrito

CAMPOS_NEGRITO = [
    "Qtd. Remessa", "Parc.", "Parceiro", "Detalhe Cabeçalho", "Clientes", "Enter", "Gravar"
]

def criar_paragrafo_apos(par_ref, texto, estilo):
    novo_par = OxmlElement("w:p")
    par_ref._element.addnext(novo_par)
    par = Paragraph(novo_par, par_ref._parent)

    try:
        par.style = estilo
    except KeyError:
        par.style = "Normal"

    if texto:
        # Primeiro, quebra o texto onde tiver **texto** ou 'texto'
        partes = re.split(r'(\*\*.*?\*\*|\'[^\']+\')', texto)

        for parte in partes:
            if parte.startswith('**') and parte.endswith('**'):
                # Negrito para **texto**
                run = par.add_run(parte[2:-2])
                run.bold = True
            elif parte.startswith("'") and parte.endswith("'"):
                # Negrito para 'texto'
                run = par.add_run(parte[1:-1])
                run.bold = True
            else:
                # Verificar se contém algum dos campos fixos
                palavras = parte.split(' ')
                for palavra in palavras:
                    if any(campo in palavra for campo in CAMPOS_NEGRITO):
                        run = par.add_run(palavra + ' ')
                        run.bold = True
                    else:
                        run = par.add_run(palavra + ' ')
                    run.font.name = "Calibri"
                    run.font.size = Pt(11)
                continue  # já processou as palavras
            run.font.name = "Calibri"
            run.font.size = Pt(11)

    return par




def criar_lista_personalizada(par_ref, texto):
    novo_par = OxmlElement("w:p")
    par_ref._element.addnext(novo_par)
    par = Paragraph(novo_par, par_ref._parent)
    par.style = "Modelomarcadores1"  # Estilo configurado no Word com alinhamento correto

    if texto:
        partes = re.split(r'(\*\*.*?\*\*|\'[^\']+\')', texto)
        for parte in partes:
            if parte.startswith('**') and parte.endswith('**'):
                run = par.add_run(parte[2:-2])
                run.bold = True
            elif parte.startswith("'") and parte.endswith("'"):
                run = par.add_run(parte[1:-1])
                run.bold = True
            else:
                palavras = parte.split(' ')
                for palavra in palavras:
                    if any(campo in palavra for campo in CAMPOS_NEGRITO):
                        run = par.add_run(palavra + ' ')
                        run.bold = True
                    else:
                        run = par.add_run(palavra + ' ')
                    run.font.name = "Calibri"
                    run.font.size = Pt(11)
                continue
            run.font.name = "Calibri"
            run.font.size = Pt(11)

    return par





def encontrar_imagem(caminho_diretorio, nome_base):
    for arquivo in os.listdir(caminho_diretorio):
        if arquivo.lower().startswith(nome_base.lower().split('.')[0]):
            return os.path.join(caminho_diretorio, arquivo)
    return None

def inserir_imagem_apos(par_ref, nome_imagem):
    caminho_diretorio = "imagens_pop"
    caminho_encontrado = encontrar_imagem(caminho_diretorio, nome_imagem)
    if not caminho_encontrado:
        print(f"⚠️ Imagem não encontrada: {nome_imagem}")
        return par_ref

    par_ref = criar_paragrafo_apos(par_ref, "", estilo="Normal")
    novo_par = OxmlElement("w:p")
    par_ref._element.addnext(novo_par)
    par = Paragraph(novo_par, par_ref._parent)
    par.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    par.add_run().add_picture(caminho_encontrado, width=Cm(16.47), height=Cm(9.26))
    par = criar_paragrafo_apos(par, "", estilo="Normal")
    return par

def limpar_nome_arquivo(texto):
    return re.sub(r'[\\/*?:"<>|]', "-", texto.strip())

# Conectar novamente na planilha se não estiver conectada
data_atualizada = aba.get_all_values()
cabecalhos = data_atualizada[0]

try:
    idx_status = cabecalhos.index("Status Padronização")
except ValueError:
    raise Exception("⚠️ Coluna 'Status Padronização' não encontrada.")

for i, linha in enumerate(data_atualizada[1:], start=2):
    status = linha[idx_status].strip().lower() if idx_status < len(linha) else ""
    if status != "não padronizado":
        continue

    dados_dict = dict(zip(cabecalhos, linha))

    atividades = []
    j = 1
    while f"Atividade {j}" in cabecalhos:
        idx_atividade = cabecalhos.index(f"Atividade {j}")
        if idx_atividade >= len(linha) or not linha[idx_atividade].strip():
            j += 1
            continue
        try:
            nome = linha[cabecalhos.index(f"Atividade {j}")].strip()
            descricao = linha[cabecalhos.index(f"Descrição {j} (texto)")].strip()
            lista_str = linha[cabecalhos.index(f"Descrição {j} (lista)")].strip()
            imagens_str = linha[cabecalhos.index(f"Imagens {j}")].strip()
        except IndexError:
            break

        lista = lista_str.split('\n') if lista_str else []
        imagens = [img.strip() for img in imagens_str.split(',') if img.strip()]

        atividades.append({
            "nome": nome,
            "descricao": descricao,
            "lista": lista,
            "imagens": imagens
        })
        j += 1

    observacoes = dados_dict.get("Observações", "").strip().split('\n') if "Observações" in dados_dict else []

    doc = Document("Modelo de Procedimento Metódico.docx")

    ancora = None
    for par in doc.paragraphs:
        if par.text.strip() == "=== ATIVIDADES AQUI ===":
            ancora = par
            break
    if not ancora:
        raise Exception("❌ Marcador '=== ATIVIDADES AQUI ===' não encontrado no modelo Word.")

    ancora.clear()
    par_ref = ancora

    for atividade in atividades:
        par_ref = criar_paragrafo_apos(par_ref, atividade['nome'], estilo="DescAtividade1")
        par_ref = criar_paragrafo_apos(par_ref, atividade['descricao'], estilo="Desatividade1")
        for item in atividade.get("lista", []):
            par_ref = criar_lista_personalizada(par_ref, item)
        for imagem in atividade.get("imagens", []):
            par_ref = inserir_imagem_apos(par_ref, imagem)

    if observacoes and observacoes[0]:
        par_ref = criar_paragrafo_apos(par_ref, "Observações Gerais", estilo="DescAtividade1")
        for obs in observacoes:
            par_ref = criar_lista_personalizada(par_ref, obs)

    codigo = limpar_nome_arquivo(dados_dict.get("Código", f"Linha{i}"))
    nome_proc = limpar_nome_arquivo(dados_dict.get("Nome do Procedimento", "Procedimento"))

    nome_doc = f"POP - {codigo} - {nome_proc}.docx"
    doc.save(nome_doc)
    print(f"✅ Documento gerado: {nome_doc}")
    aba.update_cell(i, idx_status + 1, "Pré-formatado")
    
    # === ETAPA 4: Substituição de Placeholders e Geração Final ===

def normalizar(texto):
    if not texto:
        return ""
    texto = unicodedata.normalize('NFD', texto)
    texto = texto.encode('ascii', 'ignore').decode('utf-8')
    return texto.lower().strip()

def substituir_em_paragrafos(paragrafos, dados, mapa_placeholders):
    substituidos = []
    for par in paragrafos:
        texto_original = ''.join(run.text for run in par.runs)
        novo_texto = texto_original
        for placeholder, chave_normalizada in mapa_placeholders.items():
            if placeholder in novo_texto and chave_normalizada in dados:
                novo_texto = novo_texto.replace(placeholder, dados[chave_normalizada])
                substituidos.append((placeholder, dados[chave_normalizada]))
        if novo_texto != texto_original and par.runs:
            par.runs[0].text = novo_texto
            for run in par.runs[1:]:
                run.text = ""
    return substituidos

def substituir_em_tabelas(tabelas, dados, mapa_placeholders):
    substituidos = []
    for tabela in tabelas:
        for linha in tabela.rows:
            for celula in linha.cells:
                substituidos += substituir_em_paragrafos(celula.paragraphs, dados, mapa_placeholders)
                substituidos += substituir_em_tabelas(celula.tables, dados, mapa_placeholders)
    return substituidos

def extrair_placeholders_doc(doc):
    encontrados = set()

    def extrair_em_paragrafos(paragrafos):
        for p in paragrafos:
            encontrados.update(re.findall(r"{{{.*?}}}|{{.*?}}", p.text))

    def extrair_em_tabelas(tabelas):
        for tabela in tabelas:
            for linha in tabela.rows:
                for celula in linha.cells:
                    extrair_em_paragrafos(celula.paragraphs)
                    extrair_em_tabelas(celula.tables)

    extrair_em_paragrafos(doc.paragraphs)
    extrair_em_tabelas(doc.tables)
    for section in doc.sections:
        extrair_em_paragrafos(section.header.paragraphs)
        extrair_em_tabelas(section.header.tables)
        extrair_em_paragrafos(section.footer.paragraphs)
        extrair_em_tabelas(section.footer.tables)

    return encontrados

# Recarregar planilha
valores = aba.get_all_values()
cabecalhos = valores[0]
idx_status = cabecalhos.index("Status Padronização")

# Encontrar linha "Pré-formatado"
linha_alvo = None
for i, linha in enumerate(valores[1:], start=2):
    status = linha[idx_status].strip().lower()
    if status == "pré-formatado":
        linha_alvo = linha
        linha_num = i
        break

if not linha_alvo:
    raise Exception("Nenhuma linha com status 'Pré-formatado' encontrada.")

dados_dict_original = dict(zip(cabecalhos, linha_alvo))
dados_dict = {normalizar(k): v for k, v in dados_dict_original.items()}

mapa_placeholders = {f"{{{{{k.strip()}}}}}": k for k in dados_dict_original}
mapa_placeholders.update({f"{{{{{{{k.strip()}}}}}}}": k for k in dados_dict_original})

codigo = dados_dict_original.get("Código") or dados_dict.get("codigo") or f"Linha{linha_num}"
nome_proc = dados_dict_original.get("Nome do Procedimento") or dados_dict.get("nome do procedimento") or "Procedimento"
versao = dados_dict_original.get("Versão") or dados_dict.get("versao") or "v1"

# Limpar para nome de arquivo
def limpar_nome_arquivo(texto):
    return re.sub(r'[\\/*?:"<>|]', "-", texto.strip())

codigo_limpo = limpar_nome_arquivo(codigo)
nome_proc_limpo = limpar_nome_arquivo(nome_proc)
versao_limpa = limpar_nome_arquivo(versao)

nome_doc = f"POP - {codigo_limpo} - {nome_proc_limpo}.docx"

if not os.path.exists(nome_doc):
    raise FileNotFoundError(f"""
Arquivo não encontrado: {nome_doc}
Dica: Verifique se os campos 'Código' e 'Nome do Procedimento' estão corretamente preenchidos
na aba 'Respostas ao formulário 1' da planilha.
""")


# Abrir documento
doc = Document(nome_doc)

# Substituir Placeholders
placeholders_encontrados = extrair_placeholders_doc(doc)
substituidos = substituir_em_paragrafos(doc.paragraphs, dados_dict_original, mapa_placeholders)
substituidos += substituir_em_tabelas(doc.tables, dados_dict_original, mapa_placeholders)

for section in doc.sections:
    substituidos += substituir_em_paragrafos(section.header.paragraphs, dados_dict_original, mapa_placeholders)
    substituidos += substituir_em_tabelas(section.header.tables, dados_dict_original, mapa_placeholders)
    substituidos += substituir_em_paragrafos(section.footer.paragraphs, dados_dict_original, mapa_placeholders)
    substituidos += substituir_em_tabelas(section.footer.tables, dados_dict_original, mapa_placeholders)

# Debug de placeholders
substituidos_set = {s[0] for s in substituidos}
nao_substituidos = placeholders_encontrados - substituidos_set

if nao_substituidos:
    print("\n⚠️ Placeholders NÃO substituídos:")
    for ph in nao_substituidos:
        print(f"   {ph}")
else:
    print("\n✅ Todos os placeholders foram substituídos com sucesso.")

# Remover trecho entre "=== EXCLUIR ==="
def remover_trecho_para_excluir(doc):
    excluir_inicio = None
    excluir_fim = None
    for i, par in enumerate(doc.paragraphs):
        if '=== EXCLUIR ===' in par.text:
            if excluir_inicio is None:
                excluir_inicio = i
            else:
                excluir_fim = i
                break
    if excluir_inicio is not None and excluir_fim is not None:
        for i in range(excluir_fim, excluir_inicio - 1, -1):
            p = doc.paragraphs[i]._element
            p.getparent().remove(p)
            p._p = p._element = None
        print("🗑️ Trecho entre '=== EXCLUIR ===' removido.")
    else:
        print("⚠️ Marcadores '=== EXCLUIR ===' não encontrados.")

remover_trecho_para_excluir(doc)

# Salvar novo documento final
nome_final = f"{codigo_limpo} - {nome_proc_limpo} - v.{versao_limpa}.docx"
caminho_destino = os.path.join("CAMINHO_PARA_SALVAR_POP", nome_final)
doc.save(caminho_destino)

print(f"\n📄 Documento final salvo em: {caminho_destino}")

# Atualizar status para "Padronizado"
aba.update_cell(linha_num, idx_status + 1, "Padronizado")
# === LIMPEZA DAS IMAGENS ===
imagens_dir = "imagens_pop"
if os.path.exists(imagens_dir):
    for arquivo in os.listdir(imagens_dir):
        caminho_arquivo = os.path.join(imagens_dir, arquivo)
        try:
            if os.path.isfile(caminho_arquivo):
                os.remove(caminho_arquivo)
        except Exception as e:
            print(f"⚠️ Erro ao apagar {arquivo}: {e}")
    print(f"🧹 Imagens da pasta '{imagens_dir}' apagadas com sucesso.")

# === LIMPEZA DE ARQUIVOS INTERMEDIÁRIOS ===
arquivos_intermediarios = ["POP_INPUT.docx", nome_doc]
for arquivo in arquivos_intermediarios:
    try:
        if os.path.exists(arquivo):
            os.remove(arquivo)
            print(f"🗑️ Arquivo intermediário removido: {arquivo}")
    except Exception as e:
        print(f"⚠️ Erro ao remover arquivo intermediário {arquivo}: {e}")
else:
    print(f"⚠️ Pasta '{imagens_dir}' não encontrada.")

