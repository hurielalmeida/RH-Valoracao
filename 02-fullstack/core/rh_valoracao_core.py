# ============================================================
# PROCESSADOR UNIVERSAL V11 — RH VALORAÇÃO / FOLHA MENSAL
# ============================================================
# Objetivo:
#   - Ler uma planilha-base já existente;
#   - Ler uma ou várias folhas mensais em PDF;
#   - Identificar automaticamente a competência;
#   - Extrair colaboradores, dados cadastrais e lançamentos;
#   - Cruzar matrícula e, quando necessário, nome normalizado;
#   - Criar/atualizar colunas para TODOS os códigos de proventos
#     encontrados no PDF (sem limitar-se às colunas existentes);
#   - Capturar o INSS 998 mesmo quando estiver como desconto;
#   - Preservar a estrutura e os meses anteriores;
#   - Criar um novo bloco mensal com população dinâmica;
#   - Calcular o Total Geral como soma de todos os valores
#     financeiros lançados no novo bloco.
#
# IMPORTANTE:
#   O código não inventa valores. Se uma informação não puder
#   ser identificada com segurança, ela fica em branco e aparece
#   no relatório de validação.
# ============================================================

# Se estiver no Google Colab, execute uma vez:
# !pip -q install pymupdf openpyxl pandas

import os
import re
import math
import shutil
import unicodedata
import zipfile
from copy import copy
from datetime import datetime, date

import pymupdf
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Alignment, Protection
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIGURAÇÕES
# ============================================================

ABA = "RH Valoração"

# O script procura automaticamente 1 Excel e todos os PDFs
# encontrados na pasta /content.
PASTA_TRABALHO = "/content/rh_valoracao_execucao"
_ZIPS_EXTRAIDOS = set()

# Se True, cria uma cópia de segurança antes de alterar a base.
CRIAR_BACKUP = True

# Se True, colaboradores novos encontrados no PDF e ausentes
# no último bloco também entram no novo mês.
ADICIONAR_NOVOS_DO_PDF = True

# Modo de população do novo mês:
#   "UNIAO"  = mantém o último modelo e acrescenta novos do PDF;
#   "PDF"    = usa exatamente os colaboradores encontrados no PDF;
#   "MODELO" = usa somente a população do último mês.
POPULACAO_MODO = "PDF"

# Se False, somente P entra como provento.
# O 998 é uma exceção: ele pode aparecer como D e é tratado
# separadamente como INSS.
INCLUIR_DESCONTOS_NAO_INSS = False

# Códigos de desconto que devem ser incorporados ao total,
# quando aplicável. Por padrão, somente 998.
CODIGOS_DESCONTO_INCLUIR = {"998"}

# O código 497 NÃO é convertido para 391.
# Cada código do PDF é tratado como seu próprio código.
MAPA_CODIGOS = {}
MAPEAR_RUBRICA_POR_DESCRICAO = False
EXCLUIR_COLUNAS_FINANCEIRAS_SEM_VALOR = True
PERMITIR_BASE_SEM_POPULACAO = True
ALIASES_CADASTRAIS = {
    "nome": {"NOME", "NOME COMPLETO", "COLABORADOR", "FUNCIONARIO", "FUNCIONÁRIO", "EMPREGADO"},
    "cpf": {"CPF", "CPF/CNPJ"},
    "registro": {"REGISTRO", "MATRICULA", "MATRÍCULA", "CODIGO FUNCIONARIO", "CÓDIGO FUNCIONÁRIO", "CHAPA"},
    "cc": {"CC", "CENTRO DE CUSTO", "CENTRO CUSTO", "CENTRO_CUSTO"},
    "cargo": {"CARGO", "CARGO/FUNÇÃO", "CARGO FUNCAO", "FUNÇÃO", "FUNCAO", "FUNÇÃO/CARGO"},
    "mes": {"MES", "MÊS", "COMPETENCIA", "COMPETÊNCIA", "PERIODO", "PERÍODO"},
    "total": {"TOTAL GERAL", "TOTAL", "TOTAL VALORAÇÃO", "TOTAL VALORACAO"}
}

# Tolerância para considerar duas linhas do PDF como a mesma
# linha visual.
TOLERANCIA_Y = 2.0

# Faixas de coordenadas do layout conhecido.
# O parser possui tolerância e procura candidatos nessas áreas.
FAIXA_CODIGO_ESQUERDA = (10, 60)
FAIXA_VALOR_PROVENTO = (160, 285)
FAIXA_INDICADOR = (245, 300)
FAIXA_CODIGO_DIREITA = (275, 340)
FAIXA_VALOR_DESCONTO = (480, 570)


# ============================================================
# FUNÇÕES BÁSICAS
# ============================================================

def normalizar_texto(valor):
    """Normaliza texto para comparação."""
    if valor is None:
        return ""

    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(
        c for c in texto
        if not unicodedata.combining(c)
    )
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()


# Normaliza os aliases depois que normalizar_texto já foi definida.
ALIASES_CADASTRAIS = {
    chave: {normalizar_texto(v) for v in valores}
    for chave, valores in ALIASES_CADASTRAIS.items()
}


def normalizar_registro(valor):
    """Normaliza matrícula/registro preservando apenas dígitos."""
    if valor is None:
        return ""

    texto = str(valor).strip()

    # Excel pode devolver matrícula numérica como 34711.0
    if re.fullmatch(r"\d+\.0", texto):
        texto = texto[:-2]

    return re.sub(r"\D", "", texto)


def eh_valor_br(texto):
    """Reconhece valores monetários brasileiros."""
    if texto is None:
        return False

    texto = str(texto).strip()

    return bool(
        re.fullmatch(
            r"-?(?:\d{1,3}(?:\.\d{3})+|\d+),\d{2}",
            texto
        )
    )


def converter_valor_br(texto):
    """Converte 10.122,30 para 10122.30."""
    texto = str(texto).strip()
    return float(
        texto.replace(".", "").replace(",", ".")
    )


def codigo_numerico(texto):
    """Normaliza código de rubrica para string numérica."""
    if texto is None:
        return None

    texto = str(texto).strip()

    if re.fullmatch(r"\d+\.0", texto):
        texto = texto[:-2]

    if re.fullmatch(r"\d{1,8}", texto):
        return texto

    return None


def copiar_estilo(origem, destino):
    """Copia estilo completo de uma célula para outra."""
    if origem.has_style:
        destino._style = copy(origem._style)

    if origem.number_format:
        destino.number_format = origem.number_format

    if origem.alignment:
        destino.alignment = copy(origem.alignment)

    if origem.protection:
        destino.protection = copy(origem.protection)


def primeira_data(valor):
    """Converte valores de data/datetime para date."""
    if isinstance(valor, datetime):
        return valor.date()

    if isinstance(valor, date):
        return valor

    return None


# ============================================================
# LOCALIZAÇÃO DOS ARQUIVOS
# ============================================================

def extrair_zips():
    """Extrai ZIPs de forma segura dentro da pasta da execução."""
    zip_invalidos = []
    for raiz, _, arquivos in os.walk(PASTA_TRABALHO):
        for arquivo in arquivos:
            if not arquivo.lower().endswith(".zip"):
                continue
            caminho_zip = os.path.join(raiz, arquivo)
            chave_zip = (os.path.realpath(caminho_zip), os.path.getsize(caminho_zip), os.path.getmtime(caminho_zip))
            if chave_zip in _ZIPS_EXTRAIDOS:
                continue
            pasta_saida = os.path.join(PASTA_TRABALHO, "_arquivos_extraidos")
            os.makedirs(pasta_saida, exist_ok=True)
            try:
                with zipfile.ZipFile(caminho_zip, "r") as z:
                    base = os.path.realpath(pasta_saida)
                    for info in z.infolist():
                        nome = info.filename.replace("\\", "/")
                        if not nome or nome.endswith("/"):
                            continue
                        destino = os.path.realpath(os.path.join(pasta_saida, nome))
                        if os.path.commonpath([base, destino]) != base:
                            raise ValueError(f"ZIP contém caminho inseguro: {info.filename}")
                    z.extractall(pasta_saida)
                _ZIPS_EXTRAIDOS.add(chave_zip)
                print(f"ZIP extraído: {arquivo}")
            except zipfile.BadZipFile:
                zip_invalidos.append(arquivo)
    if zip_invalidos:
        raise ValueError("ZIP inválido: " + ", ".join(zip_invalidos))


def localizar_arquivos():
    extrair_zips()

    excels = []
    pdfs = []

    for raiz, _, arquivos in os.walk(PASTA_TRABALHO):
        for arquivo in arquivos:
            caminho = os.path.join(raiz, arquivo)

            if arquivo.lower().endswith(
                (".xlsx", ".xlsm")
            ):
                excels.append(caminho)

            elif arquivo.lower().endswith(".pdf"):
                pdfs.append(caminho)

    # Remove eventuais arquivos de backup/output criados pelo script
    excels = [
        x for x in excels
        if "~$" not in os.path.basename(x)
        and "SAIDA_RH_VALORACAO" not in os.path.basename(x).upper()
        and "BACKUP_RH_VALORACAO" not in os.path.basename(x).upper()
        and "_V1_ATUALIZADA" not in os.path.basename(x).upper()
        and "_V2_ATUALIZADA" not in os.path.basename(x).upper()
        and "_V3_ATUALIZADA" not in os.path.basename(x).upper()
        and "_V5_0_ATUALIZADA" not in os.path.basename(x).upper()
        and re.search(r"_V\d+(?:_\d+)?_ATUALIZADA", os.path.basename(x).upper()) is None
    ]

    return sorted(excels), sorted(pdfs)


# ============================================================
# LEITURA DAS PALAVRAS DO PDF
# ============================================================

def agrupar_palavras_por_linha(page, tolerancia=TOLERANCIA_Y):
    """Agrupa palavras por linha usando ordenação espacial robusta."""
    palavras=sorted(page.get_text("words"), key=lambda p:(p[1],p[0]))
    linhas=[]
    for palavra in palavras:
        y0=palavra[1]
        alvo=None
        for linha in reversed(linhas[-4:]):
            if abs(linha["y"]-y0)<=tolerancia:
                alvo=linha; break
        if alvo is None:
            alvo={"y":y0,"palavras":[]}; linhas.append(alvo)
        alvo["palavras"].append(palavra)
    for linha in linhas:
        linha["palavras"].sort(key=lambda p:p[0])
    return linhas


def linha_eh_cabecalho_colaborador(palavras):
    textos = [
        str(p[4]).strip()
        for p in palavras
    ]

    return any(
        normalizar_texto(x) in {"EMPR.:", "CONTR.:"}
        for x in textos
    )



def extrair_cabecalho_colaborador(palavras):
    ordenadas = sorted(palavras, key=lambda p: p[0])
    textos = [str(p[4]).strip() for p in ordenadas]
    marcador_idx = next((i for i,t in enumerate(textos) if normalizar_texto(t) in {"EMPR.:","CONTR.:"}), None)
    if marcador_idx is None:
        return None
    reg_idx = next((j for j in range(marcador_idx+1, min(len(textos), marcador_idx+5))
                    if re.fullmatch(r"\d{4,6}", textos[j])), None)
    if reg_idx is None:
        return None
    partes=[]
    for t in textos[reg_idx+1:]:
        if normalizar_texto(t) in {"SITUACAO:","SITUAÇÃO:","CPF:","ADM:","VINCULO:","VÍNCULO:","CC:","DEPTO:","HORAS MÊS:","HORAS MES:"}:
            break
        if t: partes.append(t)
    nome=" ".join(partes).strip()
    if not nome: return None
    return {"matricula":normalizar_registro(textos[reg_idx]),"nome":nome,"y":ordenadas[reg_idx][1]}

def encontrar_valor_apos_rotulo(linhas, rotulos):
    """
    Procura um rótulo e retorna o primeiro valor útil que aparece
    depois dele. Como o PDF pode quebrar o layout em linhas,
    considera também palavras das linhas seguintes.
    """

    rotulos_norm = {
        normalizar_texto(x)
        for x in rotulos
    }

    for i, linha in enumerate(linhas):
        textos = [
            str(p[4]).strip()
            for p in linha["palavras"]
        ]

        for j, texto in enumerate(textos):

            if normalizar_texto(texto) not in rotulos_norm:
                continue

            # Primeiro tenta a mesma linha.
            for proximo in textos[j + 1:]:
                if proximo.strip():
                    return proximo.strip()

            # Depois procura nas próximas linhas.
            for linha2 in linhas[i + 1:i + 6]:
                for p in linha2["palavras"]:
                    candidato = str(p[4]).strip()

                    if not candidato:
                        continue

                    if normalizar_texto(candidato) in rotulos_norm:
                        continue

                    return candidato

    return None



def extrair_dados_cadastrais(linhas):
    cpf = encontrar_valor_apos_rotulo(linhas, {"CPF:"})
    cc = encontrar_valor_apos_rotulo(linhas, {"CC:"})
    cargo = None
    for linha in linhas:
        textos=[str(p[4]).strip() for p in linha["palavras"]]
        for j,t in enumerate(textos):
            if normalizar_texto(t)=="CARGO:":
                partes=[]
                for x in textos[j+1:]:
                    if normalizar_texto(x) in {"C.B.O:","CBO:","FILIAL:","SALARIO:","SALÁRIO:"}: break
                    if x: partes.append(x)
                cargo=" ".join(partes).strip() or None
                # No PDF da MG Info, o campo Cargo vem precedido pelo
                # código CBO (ex.: "25887 ANALISTA DE SISTEMAS III").
                # A planilha deve receber apenas a descrição do cargo.
                if cargo:
                    cargo = re.sub(r"^\d{3,6}\s+", "", cargo).strip()
                break
        if cargo: break
    return {"cpf_pdf":cpf,"cc_pdf":cc,"cargo_pdf":cargo}


def extrair_lancamentos_da_linha(linha):
    palavras=sorted(linha["palavras"],key=lambda p:p[0])
    resultados=[]
    for i,palavra in enumerate(palavras):
        x0,y0,x1,y1,texto=palavra[:5]; texto=str(texto).strip()
        if not texto: continue
        if FAIXA_CODIGO_ESQUERDA[0] <= x0 <= FAIXA_CODIGO_ESQUERDA[1] and re.fullmatch(r"\d{1,6}",texto):
            codigo=codigo_numerico(texto); indic_idx=None
            for j in range(i+1,min(len(palavras),i+12)):
                tj=str(palavras[j][4]).strip()
                if tj=="P" or tj.endswith("P"):
                    indic_idx=j; break
            if indic_idx is not None:
                valor=None
                for j in range(i+1,indic_idx):
                    tj=str(palavras[j][4]).strip()
                    m=re.fullmatch(r"(-?(?:\d{1,3}(?:\.\d{3})+|\d+),\d{2})P",tj)
                    if m: valor=converter_valor_br(m.group(1))
                    elif eh_valor_br(tj): valor=converter_valor_br(tj)
                if valor is not None:
                    desc=[]
                    for j in range(i+1,indic_idx):
                        tj=str(palavras[j][4]).strip()
                        if not eh_valor_br(tj) and not re.fullmatch(r"[\d:]+",tj): desc.append(tj)
                    resultados.append({"codigo":codigo,"indicador":"P","valor":valor,"descricao":" ".join(desc).strip(),"y":linha["y"]})
        if FAIXA_CODIGO_DIREITA[0] <= x0 <= FAIXA_CODIGO_DIREITA[1] and re.fullmatch(r"\d{1,6}",texto):
            codigo=codigo_numerico(texto); valor=None; indic_idx=None
            for j in range(i+1,len(palavras)):
                tj=str(palavras[j][4]).strip()
                if tj=="D" or tj.endswith("D"): indic_idx=j; break
                if FAIXA_VALOR_DESCONTO[0] <= palavras[j][0] <= FAIXA_VALOR_DESCONTO[1] and eh_valor_br(tj): valor=converter_valor_br(tj)
            if valor is not None and (codigo in CODIGOS_DESCONTO_INCLUIR or INCLUIR_DESCONTOS_NAO_INSS):
                desc=[]
                for j in range(i+1,indic_idx if indic_idx is not None else len(palavras)):
                    tj=str(palavras[j][4]).strip()
                    if not eh_valor_br(tj) and not re.fullmatch(r"[\d:]+",tj): desc.append(tj)
                resultados.append({"codigo":codigo,"indicador":"D","valor":valor,"descricao":" ".join(desc).strip(),"y":linha["y"]})
    return resultados

def detectar_competencia(texto):
    """
    Procura competência MM/AAAA no texto do PDF.
    """

    padroes = [
        r"EXTRATO\s+MENSAL\s*\n?\s*(\d{2}/\d{4})",
        r"Compet[eê]ncia:\s*\n?\s*(\d{2}/\d{4})",
        r"\b((?:0[1-9]|1[0-2])/\d{4})\b"
    ]

    for padrao in padroes:
        match = re.search(
            padrao,
            texto,
            flags=re.IGNORECASE
        )

        if match:
            grupo = match.group(1)

            # O terceiro padrão possui grupo do mês apenas.
            if re.fullmatch(r"\d{2}/\d{4}", grupo):
                return grupo

    return None



def _dinheiro_pdf(texto):
    return bool(re.fullmatch(r"-?(?:\d{1,3}(?:\.\d{3})+|\d+),\d{2}", str(texto).strip()))


def _valor_pdf(texto):
    return float(str(texto).strip().replace(".", "").replace(",", "."))


def _parse_linha_lancamentos_datasul(linha):
    """Lê uma linha Datasul com duas colunas laterais."""
    palavras=linha["palavras"]
    resultados=[]
    # O layout Datasul possui lado esquerdo e direito. O divisor fica
    # antes do segundo código de evento; usamos 360 como divisor espacial.
    for lo,hi in ((20,360),(360,730)):
        seg=[p for p in palavras if lo<=p[0]<hi]
        if not seg: continue
        ci=next((i for i,p in enumerate(seg) if re.fullmatch(r"\d{1,6}",str(p[4]).strip())),None)
        if ci is None: continue
        codigo=str(seg[ci][4]).strip()
        resto=[str(p[4]).strip() for p in seg[ci+1:]]
        indicador=resto[-1] if resto and resto[-1] in {"+","-"} else None
        if indicador: resto=resto[:-1]
        valores=[x for x in resto if _dinheiro_pdf(x)]
        if not valores: continue
        valor=_valor_pdf(valores[-1])
        descricao=[]
        for x in resto:
            if _dinheiro_pdf(x) or re.fullmatch(r"\d+(?:\.\d+)?",x): continue
            descricao.append(x)
        # No Datasul, + representa vencimento e - desconto. Eventos sem
        # sinal são bases/encargos e não entram no valorado automático.
        if indicador=="+" or (codigo=="998" and indicador=="-"):
            resultados.append({"codigo":codigo,"indicador":"P" if indicador=="+" else "D",
                               "valor":valor,"descricao":" ".join(descricao).strip(),"y":linha["y"]})
    return resultados


def _parse_datasul(caminho_pdf, paginas, texto_completo):
    colaboradores=[]
    padrao=re.compile(r"(\d{1,8}-\d+)\s*-\s*(.+?)\s+Admissao\s*:",re.I)
    for pagina in paginas:
        linhas=pagina["linhas"]
        cabecalhos=[]
        for idx,linha in enumerate(linhas):
            texto=" ".join(str(p[4]).strip() for p in linha["palavras"])
            m=padrao.search(texto)
            if m:
                cabecalhos.append({"indice":idx,"matricula":m.group(1).split("-")[0],"nome":m.group(2).strip()})
        for pos,cab in enumerate(cabecalhos):
            fim=cabecalhos[pos+1]["indice"] if pos+1<len(cabecalhos) else len(linhas)
            bloco=linhas[cab["indice"]:fim]
            cargo=""
            for linha in bloco[:5]:
                texto=" ".join(str(p[4]).strip() for p in linha["palavras"])
                m=re.search(r"Cargo:\s*(.+?)(?:\s+Sal\s+Cat:|$)",texto,re.I)
                if m: cargo=m.group(1).strip(); break
            lancamentos=[]
            for linha in bloco:
                texto=" ".join(str(p[4]).strip() for p in linha["palavras"])
                if "Total Vencimentos" in texto: break
                lancamentos.extend(_parse_linha_lancamentos_datasul(linha))
            colaboradores.append({"pagina":pagina["pagina"],"matricula":cab["matricula"],"nome":cab["nome"],
                                  "cpf_pdf":"","cc_pdf":"","cargo_pdf":cargo,"lancamentos":lancamentos})
    return colaboradores


def _parse_linha_log(linha):
    """Lê uma linha do relatório Log Planning: proventos à esquerda, descontos à direita."""
    palavras=linha["palavras"]; resultados=[]
    for lado,(lo,hi) in enumerate(((20,300),(300,590))):
        seg=[p for p in palavras if lo<=p[0]<hi]
        if not seg: continue
        ci=next((i for i,p in enumerate(seg) if re.fullmatch(r"\d{1,8}",str(p[4]).strip())),None)
        if ci is None: continue
        codigo=str(seg[ci][4]).strip()
        resto=[str(p[4]).strip() for p in seg[ci+1:]]
        valores=[x for x in resto if _dinheiro_pdf(x)]
        if not valores: continue
        valor=_valor_pdf(valores[-1])
        descricao=[x for x in resto if not _dinheiro_pdf(x) and not re.fullmatch(r"\d+(?:\.\d+)?",x)]
        if lado==0:
            resultados.append({"codigo":codigo,"indicador":"P","valor":valor,"descricao":" ".join(descricao).strip(),"y":linha["y"]})
        elif codigo=="998":
            resultados.append({"codigo":codigo,"indicador":"D","valor":valor,"descricao":" ".join(descricao).strip(),"y":linha["y"]})
    return resultados


def _parse_log_planning(caminho_pdf, paginas, texto_completo):
    colaboradores=[]
    for pagina in paginas:
        linhas=pagina["linhas"]
        for idx,linha in enumerate(linhas):
            texto=" ".join(str(p[4]).strip() for p in linha["palavras"])
            if not all(x in texto for x in ("Cód:","Nome:","Função:")): continue
            palavras=linha["palavras"]
            reg=next((str(p[4]).strip() for p in palavras if 40<=p[0]<80 and re.fullmatch(r"\d{1,8}",str(p[4]).strip())),None)
            nome=" ".join(str(p[4]).strip() for p in palavras if 135<=p[0]<330).strip()
            cargo=" ".join(str(p[4]).strip() for p in palavras if 365<=p[0]<513).strip()
            if not reg or not nome: continue
            lancamentos=[]
            for linha2 in linhas[idx+1:]:
                texto2=" ".join(str(p[4]).strip() for p in linha2["palavras"])
                if "Cód:" in texto2 or "Proventos:" in texto2: break
                lancamentos.extend(_parse_linha_log(linha2))
            colaboradores.append({"pagina":pagina["pagina"],"matricula":reg,"nome":nome,
                                  "cpf_pdf":"","cc_pdf":"","cargo_pdf":cargo,"lancamentos":lancamentos})
    return colaboradores


def extrair_pdf(caminho_pdf):
    """Parser universal: identifica o layout do PDF e normaliza a saída."""
    documento=pymupdf.open(caminho_pdf)
    if documento.page_count==0:
        documento.close(); raise ValueError(f"PDF sem páginas legíveis: {os.path.basename(caminho_pdf)}")
    paginas=[]; textos=[]
    for numero,pagina in enumerate(documento,start=1):
        paginas.append({"pagina":numero,"linhas":agrupar_palavras_por_linha(pagina)})
        textos.append(pagina.get_text("text"))
    texto_completo="\n".join(textos)
    # A competência deve vir do período de referência da folha, e não da
    # data de emissão do relatório (ex.: Datasul pode trazer Data: 02/01/2026
    # para uma folha Referente: 12/2025).
    m_ref=re.search(r"Referente\s*:\s*(0[1-9]|1[0-2])/(20\d{2})",texto_completo,re.I)
    m_periodo=re.search(r"Per[ií]odo\s+de:\s*(?:0[1-9]|1[0-2])/\d{2}/(20\d{2})",texto_completo,re.I)
    competencia=(f"{m_ref.group(1)}/{m_ref.group(2)}" if m_ref else None)
    if not competencia:
        m_periodo2=re.search(r"Per[ií]odo\s+de:\s*(\d{2})/(\d{2})/(20\d{2})",texto_completo,re.I)
        if m_periodo2: competencia=f"{m_periodo2.group(2)}/{m_periodo2.group(3)}"
    if not competencia:
        competencia=detectar_competencia(texto_completo)
    if not competencia:
        m=re.search(r"(?<!\d)(0[1-9]|1[0-2])\D{0,4}(20\d{2})(?!\d)",os.path.basename(caminho_pdf))
        if m: competencia=f"{m.group(1)}/{m.group(2)}"
    upper=normalizar_texto(texto_completo)
    if "DATASUL" in upper and "FP4000RP" in upper:
        colaboradores=_parse_datasul(caminho_pdf,paginas,texto_completo)
        formato="DATASUL"
    elif "COD:" in upper and "NOME:" in upper and "FUNCAO:" in upper and "PROVENTOS:" in upper:
        colaboradores=_parse_log_planning(caminho_pdf,paginas,texto_completo)
        formato="LOG_PLANNING"
    else:
        # fallback para o parser V6 original (layouts que já atendem ao padrão legado)
        colaboradores=[]
        for pagina in paginas:
            linhas=pagina["linhas"]; cabecalhos=[]
            for idx,linha in enumerate(linhas):
                cab=extrair_cabecalho_colaborador(linha["palavras"])
                if cab: cabecalhos.append({**cab,"indice":idx,"pagina":pagina["pagina"]})
            for idx_cab,cab in enumerate(cabecalhos):
                inicio=cab["indice"]; fim=cabecalhos[idx_cab+1]["indice"] if idx_cab+1<len(cabecalhos) else len(linhas)
                bloco=linhas[inicio:fim]; cadastro=extrair_dados_cadastrais(bloco); registros=[]
                for linha in bloco:
                    texto=" ".join(str(p[4]).strip() for p in linha["palavras"])
                    if "Resumo por Rubrica" in texto or "Total Geral" in texto: continue
                    registros.extend(extrair_lancamentos_da_linha(linha))
                colaboradores.append({"pagina":pagina["pagina"],"matricula":cab["matricula"],"nome":cab["nome"],
                                      "cpf_pdf":cadastro["cpf_pdf"],"cc_pdf":cadastro["cc_pdf"],"cargo_pdf":cadastro["cargo_pdf"],"lancamentos":registros})
        formato="LEGADO"
    documento.close()
    unicos={}
    for item in colaboradores:
        reg=item["matricula"]
        if reg not in unicos: unicos[reg]=item
        else:
            unicos[reg]["lancamentos"].extend(item["lancamentos"])
    return {"arquivo":caminho_pdf,"competencia":competencia,"formato":formato,
            "colaboradores":list(unicos.values())}


def localizar_aba(wb):
    if ABA in wb.sheetnames:
        return wb[ABA]

    # Fallback: procurar aba cujo nome normalizado seja equivalente.
    alvo = normalizar_texto(ABA)

    for nome in wb.sheetnames:
        if normalizar_texto(nome) == alvo:
            return wb[nome]

    raise KeyError(
        f"Aba '{ABA}' não encontrada."
    )


def localizar_colunas_de_rubricas(ws):
    """
    Procura códigos numéricos nos primeiros 6 níveis de cabeçalho.
    Retorna código -> coluna.
    """

    rubricas = {}

    for linha in range(1, 7):
        for coluna in range(
            1,
            ws.max_column + 1
        ):
            valor = ws.cell(
                linha,
                coluna
            ).value

            codigo = codigo_numerico(valor)

            if codigo is None:
                continue

            if codigo not in rubricas:
                rubricas[codigo] = coluna

    return rubricas


def localizar_coluna_por_texto(ws, texto_alvo):
    alvo = normalizar_texto(texto_alvo)

    # A linha 6 é o cabeçalho principal do modelo.
    for linha in (6, 5, 7, 8, 1, 2, 3):
        if linha > ws.max_row:
            continue
        for coluna in range(1, ws.max_column + 1):
            if normalizar_texto(ws.cell(linha, coluna).value) == alvo:
                return coluna
    return None


def identificar_colunas_principais(ws):
    resultado = {"nome": None, "cpf": None, "registro": None, "mes": None, "total": None}

    # Primeiro procura na linha 6, depois nas demais linhas de cabeçalho.
    linhas_prioritarias = [6, 5, 7, 8, 1, 2, 3]
    for linha in linhas_prioritarias:
        if linha > ws.max_row:
            continue
        for coluna in range(1, ws.max_column + 1):
            valor = normalizar_texto(ws.cell(linha, coluna).value)
            if not valor:
                continue
            for chave in resultado:
                if resultado[chave] is None and valor in ALIASES_CADASTRAIS.get(chave, set()):
                    resultado[chave] = coluna

    # Fallbacks compatíveis com a base histórica.
    resultado["nome"] = resultado["nome"] or 5
    resultado["registro"] = resultado["registro"] or 7
    resultado["mes"] = resultado["mes"] or 17

    if resultado["total"] is None:
        # Nunca usar a linha 4 como fonte primária: nela o modelo
        # pode ter subtítulos genéricos como "Total".
        resultado["total"] = localizar_coluna_por_texto(ws, "Total Geral")
        if resultado["total"] is None:
            resultado["total"] = localizar_coluna_por_texto(ws, "Total Valoração")

    if resultado["total"] is None:
        # Último fallback: procura apenas em linhas de cabeçalho,
        # excluindo explicitamente a linha 4.
        for linha in (6, 5, 7, 8, 1, 2, 3):
            if linha > ws.max_row:
                continue
            for coluna in range(1, ws.max_column + 1):
                valor = normalizar_texto(ws.cell(linha, coluna).value)
                if valor in {"TOTAL GERAL", "TOTAL VALORACAO"}:
                    resultado["total"] = coluna
                    break
            if resultado["total"] is not None:
                break

    if resultado["total"] is None:
        raise ValueError("Não encontrei a coluna Total Geral/Total Valoração no cabeçalho principal.")

    return resultado

def obter_linhas_com_mes(ws, col_mes):
    resultado = []

    for linha in range(1, ws.max_row + 1):
        valor = ws.cell(
            linha,
            col_mes
        ).value

        dt = primeira_data(valor)

        if dt is not None:
            resultado.append(
                (linha, dt)
            )

    return resultado



def localizar_ultimo_bloco(ws,col_mes,col_nome):
    linhas_mes=obter_linhas_com_mes(ws,col_mes)
    if not linhas_mes: return None,None,None
    ultima=max(dt for _,dt in linhas_mes)
    linhas=[r for r,dt in linhas_mes if dt==ultima and ws.cell(r,col_nome).value not in (None,"")]
    return (ultima,min(linhas),max(linhas)) if linhas else (ultima,None,None)

def ler_modelo_mensal(ws, inicio, fim, colunas):
    colaboradores = []

    for linha in range(inicio, fim + 1):

        nome = ws.cell(
            linha,
            colunas["nome"]
        ).value

        if nome in (None, ""):
            continue

        registro = ws.cell(
            linha,
            colunas["registro"]
        ).value

        cpf = (
            ws.cell(
                linha,
                colunas["cpf"]
            ).value
            if colunas["cpf"]
            else None
        )

        # Tentar descobrir C.C. e Cargo pelo cabeçalho
        # caso existam na planilha.
        colaboradores.append({
            "linha_origem": linha,
            "nome": str(nome).strip(),
            "registro": normalizar_registro(
                registro
            ),
            "cpf": cpf,
            "cc": None,
            "cargo": None
        })

    # Buscar cabeçalhos de CC e Cargo.
    cabecalhos = {}

    for linha in range(1, 7):
        for coluna in range(
            1,
            ws.max_column + 1
        ):
            valor = normalizar_texto(
                ws.cell(linha, coluna).value
            )

            if valor in {
                "CC",
                "CENTRO DE CUSTO",
                "CENTRO CUSTO"
            }:
                cabecalhos["cc"] = coluna

            elif valor in {
                "CARGO",
                "CARGO/FUNÇÃO",
                "CARGO FUNCAO"
            }:
                cabecalhos["cargo"] = coluna

    for item in colaboradores:
        linha = item["linha_origem"]

        if "cc" in cabecalhos:
            item["cc"] = ws.cell(
                linha,
                cabecalhos["cc"]
            ).value

        if "cargo" in cabecalhos:
            item["cargo"] = ws.cell(
                linha,
                cabecalhos["cargo"]
            ).value

    return colaboradores


# ============================================================
# CRUZAMENTO PLANILHA x PDF
# ============================================================

def construir_indice_pdf(colaboradores_pdf):
    por_registro = {}
    por_nome = {}

    for item in colaboradores_pdf:

        registro = normalizar_registro(
            item["matricula"]
        )

        nome = normalizar_texto(
            item["nome"]
        )

        if registro:
            por_registro[registro] = item

        if nome:
            por_nome.setdefault(
                nome,
                []
            ).append(item)

    return por_registro, por_nome


def fazer_match(modelo, por_registro, por_nome):
    """
    Prioridade:
      1. matrícula;
      2. nome normalizado.

    Retorna também o método utilizado.
    """

    registro = normalizar_registro(
        modelo.get("registro")
    )

    nome = normalizar_texto(
        modelo.get("nome")
    )

    if registro and registro in por_registro:
        return (
            por_registro[registro],
            "MATRÍCULA"
        )

    candidatos = por_nome.get(nome, [])

    if len(candidatos) == 1:
        return (
            candidatos[0],
            "NOME"
        )

    if len(candidatos) > 1:
        return (
            None,
            "AMBÍGUO"
        )

    return (
        None,
        "NÃO ENCONTRADO"
    )


# ============================================================
# COLUNAS DINÂMICAS DE RUBRICAS
# ============================================================

def obter_descricoes_por_codigo(pdf_resultados):
    descricoes = {}

    # O parser não depende da descrição para preencher valores,
    # mas podemos tentar obtê-la da linha visual.
    # Neste primeiro consolidado, usamos código como cabeçalho
    # caso não haja descrição disponível.
    for pdf in pdf_resultados:
        for colaborador in pdf["colaboradores"]:
            for lancamento in colaborador["lancamentos"]:
                codigo = lancamento["codigo"]

                if codigo not in descricoes:
                    descricoes[codigo] = ""

    return descricoes


def coluna_eh_financeira(coluna, colunas_principais):
    """
    Considera como financeiras as colunas de rubrica e a coluna
    Total Geral. Colunas cadastrais ficam fora.
    """

    if coluna in {
        colunas_principais["nome"],
        colunas_principais["cpf"],
        colunas_principais["registro"],
        colunas_principais["mes"],
        colunas_principais["total"]
    }:
        return False

    return True



def criar_colunas_novas(ws,codigos,rubricas_existentes,colunas_principais,descricoes_pdf=None):
    """
    Garante uma coluna exclusiva para cada código de rubrica.

    Regra importante para a MG Info:
      - 497 e 391 são códigos diferentes e nunca podem compartilhar coluna;
      - não usamos descrição para decidir que dois códigos são iguais;
      - após cada inserção, o mapa de colunas é reconstruído, porque inserir
        uma coluna antes do Total Geral desloca todas as colunas à direita.
    """
    descricoes_pdf=descricoes_pdf or {}

    # Recalcula o mapa atual antes de começar.
    rubricas_existentes=localizar_colunas_de_rubricas(ws)

    for codigo in sorted(codigos,key=lambda x:int(x)):
        if codigo in rubricas_existentes:
            continue

        col_total=identificar_colunas_principais(ws)["total"]
        origem=max(1,col_total-1)

        # Insere imediatamente antes do Total Geral.
        ws.insert_cols(col_total,1)

        largura=ws.column_dimensions[get_column_letter(origem)].width
        if largura is not None:
            ws.column_dimensions[get_column_letter(col_total)].width=largura

        for l in range(1,ws.max_row+1):
            copiar_estilo(ws.cell(l,origem),ws.cell(l,col_total))

        ws.cell(3,col_total).value=codigo
        ws.cell(6,col_total).value=descricoes_pdf.get(codigo) or codigo

        # Fundamental: a inserção deslocou as colunas existentes.
        # Reconstruir evita que dois códigos apontem para a mesma coluna.
        rubricas_existentes=localizar_colunas_de_rubricas(ws)

    return rubricas_existentes

def copiar_bloco_mensal(
    ws,
    inicio_origem,
    fim_origem,
    linha_destino,
    colunas
):
    quantidade = (
        fim_origem
        - inicio_origem
        + 1
    )

    for deslocamento in range(
        quantidade
    ):
        origem_linha = (
            inicio_origem
            + deslocamento
        )
        destino_linha = (
            linha_destino
            + deslocamento
        )

        # Copiar dimensões.
        altura = ws.row_dimensions[
            origem_linha
        ].height

        if altura is not None:
            ws.row_dimensions[
                destino_linha
            ].height = altura

        # Copiar células.
        for coluna in range(
            1,
            ws.max_column + 1
        ):
            origem = ws.cell(
                origem_linha,
                coluna
            )
            destino = ws.cell(
                destino_linha,
                coluna
            )

            if origem.value is not None:
                destino.value = origem.value
            else:
                destino.value = None

            copiar_estilo(
                origem,
                destino
            )

    return quantidade


# ============================================================
# LIMPAR CAMPOS MENSAIS DO NOVO BLOCO
# ============================================================

def limpar_bloco_novo(
    ws,
    inicio,
    fim,
    colunas,
    rubricas
):
    """
    Limpa somente as colunas financeiras do novo bloco.
    Dados cadastrais e colunas posteriores ao Total Geral são
    preservados.
    """

    col_mes = colunas["mes"]
    col_total = localizar_coluna_por_texto(
        ws,
        "Total Geral"
    )

    if col_total is None:
        col_total = colunas["total"]

    if rubricas:
        col_fin_inicio = min(
            rubricas.values()
        )
    else:
        col_fin_inicio = col_mes + 1

    for linha in range(
        inicio,
        fim + 1
    ):

        # Competência.
        # Os dados cadastrais já foram copiados.
        for coluna in range(
            col_fin_inicio,
            col_total + 1
        ):
            ws.cell(
                linha,
                coluna
            ).value = None

        # Garante a competência.
        # O valor é preenchido pelo chamador.


# ============================================================
# MAPEAR COLUNAS CADASTRAIS DO EXCEL
# ============================================================


def descobrir_colunas_cadastrais(ws):
    resultado = {}
    for linha in (6, 5, 7, 8, 1, 2, 3):
        if linha > ws.max_row:
            continue
        for coluna in range(1, ws.max_column + 1):
            valor = normalizar_texto(ws.cell(linha, coluna).value)
            if not valor:
                continue
            for chave in ("nome", "cpf", "registro", "cc", "cargo"):
                if chave not in resultado and valor in ALIASES_CADASTRAIS[chave]:
                    resultado[chave] = coluna
    resultado.setdefault("nome", 5)
    resultado.setdefault("registro", 7)
    return resultado

def preencher_cadastro(
    ws,
    linha,
    modelo,
    pdf_item,
    col_cad
):
    """
    Mantém como prioridade os dados da planilha-base.
    Quando um campo cadastral estiver vazio na planilha,
    utiliza o valor encontrado no PDF.
    """

    valores = {
        "nome": modelo.get("nome"),
        "registro": modelo.get("registro"),
        "cpf": modelo.get("cpf"),
        "cc": modelo.get("cc"),
        "cargo": modelo.get("cargo")
    }

    # Se o modelo estiver vazio, usar PDF.
    if pdf_item:

        if not valores["nome"]:
            valores["nome"] = pdf_item.get(
                "nome"
            )

        if not valores["registro"]:
            valores["registro"] = pdf_item.get(
                "matricula"
            )

        if not valores["cpf"]:
            valores["cpf"] = pdf_item.get(
                "cpf_pdf"
            )

        if not valores["cc"]:
            valores["cc"] = pdf_item.get(
                "cc_pdf"
            )

        if not valores["cargo"]:
            valores["cargo"] = pdf_item.get(
                "cargo_pdf"
            )

    if col_cad.get("nome"):
        ws.cell(
            linha,
            col_cad["nome"]
        ).value = valores["nome"]

    if col_cad.get("registro"):
        ws.cell(
            linha,
            col_cad["registro"]
        ).value = valores["registro"]

    if col_cad.get("cpf"):
        ws.cell(
            linha,
            col_cad["cpf"]
        ).value = valores["cpf"]

    if col_cad.get("cc"):
        ws.cell(
            linha,
            col_cad["cc"]
        ).value = valores["cc"]

    if col_cad.get("cargo"):
        ws.cell(
            linha,
            col_cad["cargo"]
        ).value = valores["cargo"]


# ============================================================
# PREENCHIMENTO FINANCEIRO
# ============================================================

def somar_lancamentos_por_codigo(lancamentos):
    resultado = {}

    for item in lancamentos:

        codigo = MAPA_CODIGOS.get(
            item["codigo"],
            item["codigo"]
        )

        # P: sempre entra.
        # D: somente os códigos configurados.
        if item["indicador"] == "P":
            resultado[codigo] = (
                resultado.get(codigo, 0)
                + item["valor"]
            )

        elif (
            item["indicador"] == "D"
            and (
                item["codigo"]
                in CODIGOS_DESCONTO_INCLUIR
                or INCLUIR_DESCONTOS_NAO_INSS
            )
        ):
            resultado[codigo] = (
                resultado.get(codigo, 0)
                + item["valor"]
            )

    return resultado


def preencher_financeiro(
    ws,
    linha,
    lancamentos,
    rubricas
):
    valores = somar_lancamentos_por_codigo(
        lancamentos
    )

    for codigo, valor in valores.items():

        if codigo not in rubricas:
            # Em teoria todas já foram criadas.
            # Segurança para evitar perda de informação.
            raise KeyError(
                f"Rubrica {codigo} não possui coluna."
            )

        coluna = rubricas[codigo]

        ws.cell(
            linha,
            coluna
        ).value = valor

        # Garantir formato monetário sem alterar estilo.
        if not ws.cell(
            linha,
            coluna
        ).number_format:
            ws.cell(
                linha,
                coluna
            ).number_format = '#,##0.00'

    return valores


# ============================================================
# TOTAL GERAL
# ============================================================

def atualizar_subtotais_cabecalho(ws, colunas_principais, rubricas):
    """Corrige os subtotais da linha 4 para apontarem para a própria coluna."""
    col_mes=colunas_principais["mes"]
    primeira=min(rubricas.values()) if rubricas else col_mes+1
    col_total=localizar_coluna_por_texto(ws,"Total Geral") or colunas_principais["total"]
    for col in range(primeira,col_total):
        letra=get_column_letter(col)
        ws.cell(4,col).value=f"=SUBTOTAL(9,{letra}7:{letra}1048576)"


def atualizar_formulas_total_geral(
    ws,
    colunas_principais,
    rubricas
):
    """
    Atualiza o Total Geral de todas as linhas de colaboradores
    já existentes na planilha.

    O total considera todas as colunas de rubricas criadas/
    existentes entre a primeira rubrica e a coluna Total Geral.
    Isso permite que novas rubricas sejam adicionadas no futuro
    sem deixar os meses anteriores com fórmulas desatualizadas.
    """

    col_total = localizar_coluna_por_texto(
        ws,
        "Total Geral"
    )

    if col_total is None:
        col_total = colunas_principais["total"]

    if not rubricas:
        return

    col_inicio = min(
        rubricas.values()
    )

    letra_inicio = get_column_letter(
        col_inicio
    )
    letra_fim = get_column_letter(
        col_total - 1
    )

    col_mes = colunas_principais["mes"]
    col_nome = colunas_principais["nome"]

    for linha in range(
        1,
        ws.max_row + 1
    ):

        nome = ws.cell(
            linha,
            col_nome
        ).value

        mes = primeira_data(
            ws.cell(
                linha,
                col_mes
            ).value
        )

        if nome in (None, "") or mes is None:
            continue

        ws.cell(
            linha,
            col_total
        ).value = (
            f"=SUM({letra_inicio}{linha}:"
            f"{letra_fim}{linha})"
        )



def remover_colunas_sem_valor(ws,colunas_principais,rubricas):
    if not EXCLUIR_COLUNAS_FINANCEIRAS_SEM_VALOR:
        return []

    # Desfaz temporariamente mesclagens que podem manter colunas
    # "fantasmas" além da última coluna efetivamente utilizada.
    merges=list(ws.merged_cells.ranges)
    for rng in merges:
        ws.unmerge_cells(str(rng))

    col_total=colunas_principais["total"]
    col_nome=colunas_principais["nome"]; col_mes=colunas_principais["mes"]
    linhas=[r for r in range(1,ws.max_row+1)
            if ws.cell(r,col_nome).value not in (None,"")
            and primeira_data(ws.cell(r,col_mes).value) is not None]

    preservar={col_nome,col_mes,col_total}
    for chave in ("cpf","registro"):
        if colunas_principais.get(chave): preservar.add(colunas_principais[chave])

    remover=[]
    for col in range(1,ws.max_column+1):
        if col in preservar: continue
        if not any(ws.cell(r,col).value not in (None,"") for r in linhas):
            remover.append((col,ws.cell(3,col).value or ws.cell(6,col).value or "SEM CABEÇALHO"))

    for col,_ in sorted(remover,reverse=True):
        ws.delete_cols(col,1)

    # Limpa colunas residuais no final, inclusive as que tinham somente estilo.
    while ws.max_column>1:
        c=ws.max_column
        if any(ws.cell(r,c).value not in (None,"") for r in range(1,ws.max_row+1)): break
        for r in range(1,ws.max_row+1): ws.cell(r,c)._style=None
        try: del ws.column_dimensions[get_column_letter(c)]
        except Exception: pass
        ws.delete_cols(c,1)

    # Recria somente mesclagens ainda válidas, limitadas à área utilizada.
    max_col=ws.max_column
    for rng in merges:
        if rng.min_col>max_col or rng.min_row>ws.max_row: continue
        fim=min(rng.max_col,max_col)
        if fim>rng.min_col:
            try: ws.merge_cells(start_row=rng.min_row,start_column=rng.min_col,end_row=rng.max_row,end_column=fim)
            except Exception: pass

    return [str(x) for _,x in remover]


def upload_arquivos_colab():
    """No Colab, abre o seletor para Excel + um ou vários PDFs."""
    try:
        from google.colab import files
        return files.upload()
    except ImportError:
        print("Execute em Google Colab ou coloque os arquivos em PASTA_TRABALHO.")
        return {}


def _criar_linhas_com_estilo(ws, inicio, quantidade, linha_molde):
    fim=inicio+quantidade-1
    if ws.max_row < fim:
        ws.insert_rows(ws.max_row+1, amount=fim-ws.max_row)
    for linha in range(inicio,fim+1):
        if linha==linha_molde: continue
        for c in range(1,ws.max_column+1):
            copiar_estilo(ws.cell(linha_molde,c),ws.cell(linha,c))
            ws.cell(linha,c).value=None
        ws.row_dimensions[linha].height=ws.row_dimensions[linha_molde].height
    return fim


def processar():
    excels,pdfs=localizar_arquivos()
    print("="*100); print("RH VALORAÇÃO — PROCESSADOR UNIVERSAL v5.0"); print("="*100)
    print("Excel:", [os.path.basename(x) for x in excels])
    print("PDFs:", [os.path.basename(x) for x in pdfs])
    if len(excels)!=1: raise ValueError("Deve existir exatamente 1 Excel-base na pasta.")
    if not pdfs: raise ValueError("Nenhum PDF encontrado.")

    caminho_excel=excels[0]
    manter_vba=caminho_excel.lower().endswith(".xlsm")
    wb=load_workbook(caminho_excel,keep_vba=manter_vba)
    ws=localizar_aba(wb)
    colunas=identificar_colunas_principais(ws); col_cad=descobrir_colunas_cadastrais(ws)
    ultima,inicio_base,fim_base=localizar_ultimo_bloco(ws,colunas["mes"],colunas["nome"])
    modelo_atual=ler_modelo_mensal(ws,inicio_base,fim_base,col_cad) if inicio_base else []
    if modelo_atual:
        print(f"Último bloco: {ultima.strftime('%m/%Y')} | {len(modelo_atual)} colaboradores")
    else:
        print("Base sem população mensal preenchida: o primeiro PDF definirá a população.")

    resultados=[extrair_pdf(p) for p in pdfs]
    resultados.sort(key=lambda x: datetime.strptime(x["competencia"],"%m/%Y") if x["competencia"] else datetime.max)

    competencias=[x["competencia"] for x in resultados if x.get("competencia")]
    duplicadas=sorted({c for c in competencias if competencias.count(c)>1})
    if duplicadas:
        raise ValueError(
            "Há mais de um PDF para a mesma competência: "
            + ", ".join(duplicadas)
            + ". Envie somente um PDF por mês."
        )

    codigos=set(); descricoes={}
    for res in resultados:
        for p in res["colaboradores"]:
            for item in p["lancamentos"]:
                cod=item["codigo"]
                if item["indicador"]=="P" or (item["indicador"]=="D" and (cod in CODIGOS_DESCONTO_INCLUIR or INCLUIR_DESCONTOS_NAO_INSS)):
                    final=MAPA_CODIGOS.get(cod,cod); codigos.add(final)
                    if item.get("descricao"): descricoes.setdefault(final,item["descricao"])

    rubricas=localizar_colunas_de_rubricas(ws)
    qtd_antes=len(rubricas)
    rubricas=criar_colunas_novas(ws,codigos,rubricas,colunas,descricoes)
    print(f"Rubricas necessárias: {len(codigos)} | existentes/criadas: {len(rubricas)} | novas: {len(rubricas)-qtd_antes}")

    relatorio=[]; processadas=[]
    for res in resultados:
        competencia=res["competencia"]
        if not competencia: raise ValueError(f"Competência não identificada: {res['arquivo']}")
        m,a=map(int,competencia.split("/")); data_mes=date(a,m,1)
        colunas=identificar_colunas_principais(ws); col_cad=descobrir_colunas_cadastrais(ws)
        if any(dt==data_mes for _,dt in obter_linhas_com_mes(ws,colunas["mes"])):
            print(f"{competencia}: já existe; ignorada para evitar sobrescrita.")
            continue

        por_registro,por_nome=construir_indice_pdf(res["colaboradores"])
        # População do mês.
        pdf_pop=[{"nome":p["nome"],"registro":p["matricula"],"cpf":p.get("cpf_pdf"),
                  "cc":p.get("cc_pdf"),"cargo":p.get("cargo_pdf"),"linha_origem":None}
                 for p in res["colaboradores"]]
        if not modelo_atual:
            populacao=pdf_pop
        elif POPULACAO_MODO.upper()=="PDF":
            populacao=pdf_pop
        elif POPULACAO_MODO.upper()=="MODELO":
            populacao=list(modelo_atual)
        else:
            populacao=list(modelo_atual)
            existentes={(normalizar_registro(x.get("registro")),normalizar_texto(x.get("nome"))) for x in populacao}
            for p in pdf_pop:
                chave=(normalizar_registro(p["registro"]),normalizar_texto(p["nome"]))
                if chave not in existentes and p["registro"] not in {x[0] for x in existentes} and p["nome"] not in {x[1] for x in existentes}:
                    populacao.append(p)

        # Cria novo bloco.
        if inicio_base and fim_base:
            inicio_novo=ws.max_row+1
            qtd_base=fim_base-inicio_base+1
            copiar_bloco_mensal(ws,inicio_base,fim_base,inicio_novo,colunas)
            fim_novo=inicio_novo+qtd_base-1
        else:
            # Base sem população: começa na primeira linha de dados
            # após o cabeçalho, reaproveitando linhas vazias já formatadas.
            inicio_novo=7
            molde=7 if ws.max_row>=7 else max(1,inicio_novo-1)
            fim_novo=_criar_linhas_com_estilo(ws,inicio_novo,len(populacao),molde)

        atual=fim_novo-inicio_novo+1
        desejada=len(populacao)
        if desejada<atual:
            ws.delete_rows(inicio_novo+desejada,amount=atual-desejada)
            fim_novo=inicio_novo+desejada-1
        elif desejada>atual:
            molde=fim_novo
            adicionar=desejada-atual
            ws.insert_rows(fim_novo+1,amount=adicionar)
            for linha in range(fim_novo+1,fim_novo+1+adicionar):
                for c in range(1,ws.max_column+1):
                    copiar_estilo(ws.cell(molde,c),ws.cell(linha,c)); ws.cell(linha,c).value=None
                ws.row_dimensions[linha].height=ws.row_dimensions[molde].height
            fim_novo+=adicionar

        limpar_bloco_novo(ws,inicio_novo,fim_novo,colunas,rubricas)
        for linha in range(inicio_novo,fim_novo+1): ws.cell(linha,colunas["mes"]).value=data_mes

        encontrados=0
        for off,item in enumerate(populacao):
            linha=inicio_novo+off
            pdf_item,metodo=fazer_match(item,por_registro,por_nome)
            if pdf_item: encontrados+=1
            preencher_cadastro(ws,linha,item,pdf_item,col_cad)
            valores=preencher_financeiro(ws,linha,pdf_item["lancamentos"],rubricas) if pdf_item else {}
            relatorio.append({"competencia":competencia,"nome_modelo":item.get("nome",""),
                              "registro_modelo":item.get("registro",""),"registro_pdf":pdf_item["matricula"] if pdf_item else "",
                              "metodo_match":metodo,"encontrado":bool(pdf_item),"lancamentos":len(pdf_item["lancamentos"]) if pdf_item else 0,
                              "valor_total":sum(valores.values()) if valores else 0})

        atualizar_formulas_total_geral(ws,colunas,rubricas)
        print(f"{competencia}: bloco {inicio_novo}:{fim_novo} | população {desejada} | encontrados {encontrados}/{desejada}")
        processadas.append(competencia)
        inicio_base,fim_base=inicio_novo,fim_novo
        modelo_atual=ler_modelo_mensal(ws,inicio_base,fim_base,col_cad)

    colunas=identificar_colunas_principais(ws); rubricas=localizar_colunas_de_rubricas(ws)
    removidas=remover_colunas_sem_valor(ws,colunas,rubricas)
    colunas=identificar_colunas_principais(ws); rubricas=localizar_colunas_de_rubricas(ws)
    atualizar_subtotais_cabecalho(ws,colunas,rubricas)
    atualizar_formulas_total_geral(ws,colunas,rubricas)

    # Não cria uma aba auxiliar: o resultado permanece exclusivamente na
    # aba "RH Valoração", preservando o modelo solicitado.

    try:
        wb.calculation.fullCalcOnLoad=True; wb.calculation.forceFullCalc=True; wb.calculation.calcMode="auto"
    except Exception: pass

    ext=".xlsm" if manter_vba else ".xlsx"
    base_nome=os.path.splitext(os.path.basename(caminho_excel))[0]
    saida=os.path.join(PASTA_TRABALHO,f"{base_nome}_V5_0_ATUALIZADA{ext}")
    if CRIAR_BACKUP:
        backup=os.path.join(PASTA_TRABALHO,f"{base_nome}_BACKUP_RH_VALORACAO{ext}")
        if not os.path.exists(backup): shutil.copy2(caminho_excel,backup)
    wb.save(saida)
    print("="*100); print("CONCLUÍDO — v5.0")
    print("Meses processados:",", ".join(processadas) or "nenhum")
    print("Colunas financeiras sem valor removidas:",len(removidas))
    print("Arquivo:",saida)
    return saida,relatorio


# ============================================================
# EXECUTAR
# ============================================================

# caminho_saida, relatorio = processar()

# ============================================================
# V6.0 — AJUSTES ESTRUTURAIS PARA MG INFO
# ============================================================

GRUPO_PAGAMENTOS = "PRINCIPAIS VERBAS DE PAGAMENTOS  (é recomendável incluir coluna para ampliar a abrangência e retratar a realidade da empresa)"
GRUPO_ENCARGOS = "PRINCIPAIS VERBAS DE ENCARGOS (é recomendável incluir coluna para ampliar a abrangência e retratar a realidade da empresa)"
GRUPO_BENEFICIOS = "PRINCIPAIS VERBAS DE BENEFÍCIOS LEGAIS/COMPLEMENTARES (é recomendável incluir coluna para ampliar a abrangência e retratar a realidade da empresa)"

# Colunas que devem permanecer no bloco de encargos mesmo sem valores.
# O FGTS entra aqui também, conforme a estrutura original da MG Info.
ENCARGOS_PROTEGIDOS = [
    ("FGTS", "B1"),
    ("FGTS RESCISORIO", "B2"),
    ("INSS EMP.", "B2"),
    ("INSS TERC.", "B4"),
    ("RAT", "B5"),
    ("I.N.S.S", "998"),
]

BENEFICIOS_CODIGOS = {"391", "497"}


def _eh_coluna_rubrica_v6(ws, col):
    """Identifica uma coluna financeira da planilha-base sem confundir cadastro com rubrica."""
    if col < 1 or col > ws.max_column:
        return False
    codigo = ws.cell(3, col).value
    cab = normalizar_texto(ws.cell(6, col).value)
    if codigo_numerico(codigo) is not None:
        return True
    if cab in {normalizar_texto(x[0]) for x in ENCARGOS_PROTEGIDOS}:
        return True
    # Códigos estruturais B1..B5 também são rubricas, mesmo não numéricos.
    if isinstance(codigo, str) and re.fullmatch(r"B\d+", codigo.strip().upper()):
        return True
    if cab in {"AUXILIO HOME OFFICE", "AUXÍLIO HOME OFFICE"}:
        return True
    return False


def _mapa_rubricas_v6(ws):
    """Mapeia códigos numéricos. Colunas protegidas por nome são tratadas separadamente."""
    mapa = {}
    for col in range(1, ws.max_column + 1):
        codigo = codigo_numerico(ws.cell(3, col).value)
        if codigo is not None and codigo not in mapa:
            mapa[codigo] = col
    return mapa


def _coluna_por_nome_v6(ws, nome):
    alvo = normalizar_texto(nome)
    for col in range(1, ws.max_column + 1):
        if normalizar_texto(ws.cell(6, col).value) == alvo:
            return col
    return None


def _desfazer_merges_linha5(ws):
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row <= 5 <= rng.max_row:
            ws.unmerge_cells(str(rng))


def _copiar_coluna_estilo_v6(ws, origem, destino):
    if origem < 1 or origem > ws.max_column or destino < 1 or destino > ws.max_column:
        return
    letra_o = get_column_letter(origem)
    letra_d = get_column_letter(destino)
    if ws.column_dimensions[letra_o].width is not None:
        ws.column_dimensions[letra_d].width = ws.column_dimensions[letra_o].width
    for r in range(1, ws.max_row + 1):
        copiar_estilo(ws.cell(r, origem), ws.cell(r, destino))


def _inserir_coluna_financeira_v6(ws, pos, codigo=None, descricao=None):
    """Insere uma coluna preservando a aparência do modelo."""
    origem = max(1, min(pos - 1, ws.max_column))
    ws.insert_cols(pos, 1)
    # Depois da inserção, origem continua sendo uma coluna válida de modelo.
    _copiar_coluna_estilo_v6(ws, origem, pos)
    if codigo is not None:
        ws.cell(3, pos).value = codigo
    if descricao is not None:
        ws.cell(6, pos).value = descricao
    return pos


def _encontrar_inicio_grupo_v6(ws, palavra):
    alvo = normalizar_texto(palavra)
    for col in range(1, ws.max_column + 1):
        v = normalizar_texto(ws.cell(5, col).value)
        if alvo in v:
            return col
    return None


def _preparar_estrutura_financeira_v6(ws, codigos, descricoes):
    """
    Organiza as colunas financeiras mantendo a estrutura visual da base:
      PAGAMENTOS -> ENCARGOS -> BENEFÍCIOS -> TOTAL GERAL -> REGIME.
    Novos proventos entram em PAGAMENTOS; 391/497/Home Office em BENEFÍCIOS;
    998 e as colunas manuais ficam em ENCARGOS.
    """
    _desfazer_merges_linha5(ws)

    # O Total Geral é sempre encontrado pelo cabeçalho da linha 6.
    col_total = localizar_coluna_por_texto(ws, "Total Geral") or localizar_coluna_por_texto(ws, "Total Valoração")
    if col_total is None:
        raise ValueError("Não encontrei a coluna Total Geral/Total Valoração.")

    # Captura os inícios atuais antes de inserir.
    inicio_enc = _encontrar_inicio_grupo_v6(ws, "ENCARGOS")
    inicio_ben = _encontrar_inicio_grupo_v6(ws, "BENEFÍCIOS")

    # Se a base estiver sem grupos, posiciona os grupos pelo primeiro bloco financeiro.
    if inicio_enc is None:
        rubs = [c for c in range(1, col_total) if _eh_coluna_rubrica_v6(ws, c)]
        inicio_enc = (max(rubs) + 1) if rubs else col_total
    if inicio_ben is None:
        inicio_ben = inicio_enc

    # Garantir 998 em encargos, se não existir.
    if _coluna_por_nome_v6(ws, "I.N.S.S") is None:
        # inserir antes do benefício, preservando o bloco de encargos
        _inserir_coluna_financeira_v6(ws, inicio_ben, "998", "I.N.S.S")
        inicio_ben += 1
        col_total += 1

    # Garantir as cinco colunas manuais de encargos.
    # Inserimos todas antes do benefício, na ordem original.
    for nome, codigo in ENCARGOS_PROTEGIDOS[:5]:
        if _coluna_por_nome_v6(ws, nome) is None:
            _inserir_coluna_financeira_v6(ws, inicio_ben, codigo, nome.title() if nome != "INSS EMP." and nome != "INSS TERC." else ("INSS Emp." if nome == "INSS EMP." else "INSS Terc."))
            inicio_ben += 1
            col_total += 1

    # Atualizar posições após inserções.
    inicio_enc = _encontrar_inicio_grupo_v6(ws, "ENCARGOS") or inicio_enc
    inicio_ben = _encontrar_inicio_grupo_v6(ws, "BENEFÍCIOS") or inicio_ben
    col_total = localizar_coluna_por_texto(ws, "Total Geral") or localizar_coluna_por_texto(ws, "Total Valoração")

    # Códigos existentes e novos: proventos comuns vão para pagamentos;
    # benefícios 391/497 vão para o bloco de benefícios.
    mapa = _mapa_rubricas_v6(ws)
    for codigo in sorted(codigos, key=lambda x: int(x)):
        if codigo in mapa:
            continue
        if codigo in BENEFICIOS_CODIGOS:
            col_total = localizar_coluna_por_texto(ws, "Total Geral") or localizar_coluna_por_texto(ws, "Total Valoração")
            _inserir_coluna_financeira_v6(ws, col_total, codigo, descricoes.get(codigo) or codigo)
        elif codigo == "998":
            # Já garantido acima; se não, cai no bloco de encargos.
            inicio_ben = _encontrar_inicio_grupo_v6(ws, "BENEFÍCIOS") or col_total
            _inserir_coluna_financeira_v6(ws, inicio_ben, codigo, "I.N.S.S")
        else:
            inicio_enc = _encontrar_inicio_grupo_v6(ws, "ENCARGOS")
            if inicio_enc is None:
                inicio_enc = _encontrar_inicio_grupo_v6(ws, "BENEFÍCIOS") or col_total
            _inserir_coluna_financeira_v6(ws, inicio_enc, codigo, descricoes.get(codigo) or codigo)
        mapa = _mapa_rubricas_v6(ws)

    # Reorganiza grupos pela natureza de cada coluna.
    _reconstruir_grupos_v6(ws)
    return _mapa_rubricas_v6(ws)


def _reconstruir_grupos_v6(ws):
    """Reconstrói a linha 5 e os merges dinamicamente após todas as inserções."""
    _desfazer_merges_linha5(ws)
    col_total = localizar_coluna_por_texto(ws, "Total Geral") or localizar_coluna_por_texto(ws, "Total Valoração")
    if col_total is None:
        return

    # Identifica a primeira rubrica financeira.
    primeira = None
    for c in range(1, col_total):
        if _eh_coluna_rubrica_v6(ws, c):
            primeira = c
            break
    if primeira is None:
        return

    enc = set(normalizar_texto(x[0]) for x in ENCARGOS_PROTEGIDOS)
    pagamentos, encargos, beneficios = [], [], []

    for c in range(primeira, col_total):
        cab = normalizar_texto(ws.cell(6, c).value)
        codigo = codigo_numerico(ws.cell(3, c).value)
        if cab in enc or codigo == "998":
            encargos.append(c)
        elif codigo in BENEFICIOS_CODIGOS or "HOME OFFICE" in cab:
            beneficios.append(c)
        else:
            pagamentos.append(c)

    def ranges_contiguos(cols):
        if not cols:
            return []
        cols = sorted(cols)
        saida = []
        inicio = anterior = cols[0]
        for c in cols[1:]:
            if c == anterior + 1:
                anterior = c
            else:
                saida.append((inicio, anterior))
                inicio = anterior = c
        saida.append((inicio, anterior))
        return saida

    # Como inserimos por bloco, os conjuntos devem ser contíguos.
    # Se algum layout excepcional quebrar a contiguidade, cada trecho recebe
    # o mesmo título, sem apagar informação.
    grupos = [
        (pagamentos, GRUPO_PAGAMENTOS),
        (encargos, GRUPO_ENCARGOS),
        (beneficios, GRUPO_BENEFICIOS),
    ]
    for cols, titulo in grupos:
        for ini, fim in ranges_contiguos(cols):
            ws.cell(5, ini).value = titulo
            if fim > ini:
                ws.merge_cells(start_row=5, start_column=ini, end_row=5, end_column=fim)

    # Limpa eventuais títulos antigos da linha 5 fora dos grupos financeiros.
    for c in range(primeira, col_total):
        if not any(c in range(a, b + 1) for cols, _ in grupos for a, b in ranges_contiguos(cols)):
            ws.cell(5, c).value = None


def _atualizar_subtotais_v6(ws):
    """Linha 4: subtotal de cada rubrica e, principalmente, do Total Geral na própria coluna."""
    col_total = localizar_coluna_por_texto(ws, "Total Geral") or localizar_coluna_por_texto(ws, "Total Valoração")
    if col_total is None:
        return
    primeira = next((c for c in range(1, col_total) if _eh_coluna_rubrica_v6(ws, c)), None)
    if primeira is None:
        return
    for c in range(primeira, col_total + 1):
        letra = get_column_letter(c)
        ws.cell(4, c).value = f"=SUBTOTAL(9,{letra}7:{letra}1048576)"
    # Total Geral explicitamente aponta para sua própria coluna.
    letra = get_column_letter(col_total)
    ws.cell(4, col_total).value = f"=SUBTOTAL(9,{letra}7:{letra}1048576)"


def _atualizar_totais_linhas_v6(ws):
    col_total = localizar_coluna_por_texto(ws, "Total Geral") or localizar_coluna_por_texto(ws, "Total Valoração")
    if col_total is None:
        return
    primeira = next((c for c in range(1, col_total) if _eh_coluna_rubrica_v6(ws, c)), None)
    if primeira is None:
        return
    # Soma todas as colunas financeiras antes do Total Geral, incluindo os campos
    # de encargos que serão preenchidos manualmente posteriormente.
    ini = get_column_letter(primeira)
    fim = get_column_letter(col_total - 1)
    col_nome = identificar_colunas_principais(ws)["nome"]
    col_mes = identificar_colunas_principais(ws)["mes"]
    for r in range(7, ws.max_row + 1):
        if ws.cell(r, col_nome).value in (None, ""):
            continue
        if primeira_data(ws.cell(r, col_mes).value) is None:
            continue
        ws.cell(r, col_total).value = f"=SUM({ini}{r}:{fim}{r})"


def _remover_rubricas_vazias_v6(ws):
    """
    Remove somente colunas financeiras/rubricas que estejam totalmente vazias.
    Não remove colunas cadastrais/estruturais da planilha-base.
    As colunas de encargos protegidas permanecem mesmo vazias.
    """
    if not EXCLUIR_COLUNAS_FINANCEIRAS_SEM_VALOR:
        return []
    _desfazer_merges_linha5(ws)
    col_total = localizar_coluna_por_texto(ws, "Total Geral") or localizar_coluna_por_texto(ws, "Total Valoração")
    if col_total is None:
        return []
    protegidas = set(normalizar_texto(x[0]) for x in ENCARGOS_PROTEGIDOS)
    removidas = []
    # Recalcula da direita para esquerda.
    for c in range(col_total - 1, 0, -1):
        if not _eh_coluna_rubrica_v6(ws, c):
            continue
        cab = normalizar_texto(ws.cell(6, c).value)
        codigo = codigo_numerico(ws.cell(3, c).value)
        if cab in protegidas:
            continue
        # 391/497 também podem ser vazios no template e devem ser mantidos
        # porque são benefícios estruturais da MG Info.
        if codigo in BENEFICIOS_CODIGOS or "HOME OFFICE" in cab:
            continue
        # Verifica apenas as linhas de dados (7 em diante).
        tem_valor = any(ws.cell(r, c).value not in (None, "") for r in range(7, ws.max_row + 1))
        if not tem_valor:
            removidas.append(ws.cell(6, c).value or ws.cell(3, c).value)
            ws.delete_cols(c, 1)
    _reconstruir_grupos_v6(ws)
    return removidas


def _ultima_linha_dados_v6(ws):
    """Última linha efetivamente ocupada por colaborador/mês, ignorando milhares de linhas apenas formatadas."""
    col = identificar_colunas_principais(ws)["nome"]
    mes = identificar_colunas_principais(ws)["mes"]
    ultima = 6
    for r in range(7, ws.max_row + 1):
        if ws.cell(r, col).value not in (None, "") and primeira_data(ws.cell(r, mes).value) is not None:
            ultima = r
    return ultima


def _limpar_linhas_vazias_apos_dados_v6(ws):
    ultima = _ultima_linha_dados_v6(ws)
    if ws.max_row > ultima:
        ws.delete_rows(ultima + 1, ws.max_row - ultima)
    return ultima


def processar():
    excels, pdfs = localizar_arquivos()
    print("=" * 100)
    print("RH VALORAÇÃO — PROCESSADOR UNIVERSAL v11.0")
    print("=" * 100)
    print("Excel:", [os.path.basename(x) for x in excels])
    print("PDFs:", [os.path.basename(x) for x in pdfs])
    if len(excels) != 1:
        raise ValueError("Deve existir exatamente 1 Excel-base na pasta.")
    if not pdfs:
        raise ValueError("Nenhum PDF encontrado.")

    caminho_excel = excels[0]
    manter_vba = caminho_excel.lower().endswith(".xlsm")
    wb = load_workbook(caminho_excel, keep_vba=manter_vba)
    ws = localizar_aba(wb)

    # Nunca usa um arquivo gerado anteriormente como base.
    colunas = identificar_colunas_principais(ws)
    col_cad = descobrir_colunas_cadastrais(ws)
    ultima, inicio_base, fim_base = localizar_ultimo_bloco(ws, colunas["mes"], colunas["nome"])
    modelo_atual = ler_modelo_mensal(ws, inicio_base, fim_base, col_cad) if inicio_base else []
    if inicio_base:
        print(f"Último bloco existente: {ultima.strftime('%m/%Y')} | linhas {inicio_base}:{fim_base} | {len(modelo_atual)} colaboradores")
    else:
        print("Base sem população mensal preenchida: o primeiro PDF definirá a população.")

    resultados = [extrair_pdf(p) for p in pdfs]
    resultados.sort(key=lambda x: datetime.strptime(x["competencia"], "%m/%Y") if x.get("competencia") else datetime.max)
    competencias = [x["competencia"] for x in resultados if x.get("competencia")]
    duplicadas = sorted({c for c in competencias if competencias.count(c) > 1})
    if duplicadas:
        raise ValueError("Há mais de um PDF para a mesma competência: " + ", ".join(duplicadas))

    codigos = set()
    descricoes = {}
    for res in resultados:
        for p in res["colaboradores"]:
            for item in p["lancamentos"]:
                cod = item["codigo"]
                if item["indicador"] == "P" or (item["indicador"] == "D" and (cod in CODIGOS_DESCONTO_INCLUIR or INCLUIR_DESCONTOS_NAO_INSS)):
                    final = MAPA_CODIGOS.get(cod, cod)
                    codigos.add(final)
                    if item.get("descricao"):
                        descricoes.setdefault(final, item["descricao"])

    rubricas = _preparar_estrutura_financeira_v6(ws, codigos, descricoes)
    _atualizar_subtotais_v6(ws)
    _atualizar_totais_linhas_v6(ws)

    processadas = []
    relatorio = []

    for res in resultados:
        competencia = res["competencia"]
        if not competencia:
            raise ValueError(f"Competência não identificada: {res['arquivo']}")
        m, a = map(int, competencia.split("/"))
        data_mes = date(a, m, 1)

        colunas = identificar_colunas_principais(ws)
        col_cad = descobrir_colunas_cadastrais(ws)
        if any(dt == data_mes for _, dt in obter_linhas_com_mes(ws, colunas["mes"])):
            print(f"{competencia}: já existe; ignorada para evitar sobrescrita.")
            continue

        por_registro, por_nome = construir_indice_pdf(res["colaboradores"])
        pdf_pop = [{"nome": p["nome"], "registro": p["matricula"], "cpf": p.get("cpf_pdf"),
                    "cc": p.get("cc_pdf"), "cargo": p.get("cargo_pdf"), "linha_origem": None}
                   for p in res["colaboradores"]]
        populacao = pdf_pop if POPULACAO_MODO.upper() == "PDF" or not modelo_atual else list(modelo_atual)

        # O novo bloco começa exatamente na linha seguinte ao último bloco real.
        if inicio_base and fim_base:
            inicio_novo = fim_base + 1
            qtd_base = fim_base - inicio_base + 1
            fim_novo = inicio_novo + qtd_base - 1
            if ws.max_row < fim_novo:
                ws.insert_rows(ws.max_row + 1, amount=fim_novo - ws.max_row)
            copiar_bloco_mensal(ws, inicio_base, fim_base, inicio_novo, colunas)
        else:
            inicio_novo = 7
            molde = 7 if ws.max_row >= 7 else 6
            fim_novo = _criar_linhas_com_estilo(ws, inicio_novo, len(populacao), molde)

        atual = fim_novo - inicio_novo + 1
        desejada = len(populacao)
        if desejada < atual:
            ws.delete_rows(inicio_novo + desejada, amount=atual - desejada)
            fim_novo = inicio_novo + desejada - 1
        elif desejada > atual:
            molde = fim_novo
            adicionar = desejada - atual
            ws.insert_rows(fim_novo + 1, amount=adicionar)
            for linha in range(fim_novo + 1, fim_novo + 1 + adicionar):
                for c in range(1, ws.max_column + 1):
                    copiar_estilo(ws.cell(molde, c), ws.cell(linha, c))
                    ws.cell(linha, c).value = None
                ws.row_dimensions[linha].height = ws.row_dimensions[molde].height
            fim_novo += adicionar

        # Limpa valores financeiros do bloco recém-criado.
        colunas = identificar_colunas_principais(ws)
        rubricas = _mapa_rubricas_v6(ws)
        col_total = localizar_coluna_por_texto(ws, "Total Geral") or localizar_coluna_por_texto(ws, "Total Valoração")
        primeira_rubrica = next((c for c in range(1, col_total) if _eh_coluna_rubrica_v6(ws, c)), None)
        if primeira_rubrica:
            for r in range(inicio_novo, fim_novo + 1):
                for c in range(primeira_rubrica, col_total + 1):
                    ws.cell(r, c).value = None
        for r in range(inicio_novo, fim_novo + 1):
            ws.cell(r, colunas["mes"]).value = data_mes

        encontrados = 0
        for off, item in enumerate(populacao):
            linha = inicio_novo + off
            pdf_item, metodo = fazer_match(item, por_registro, por_nome)
            if pdf_item:
                encontrados += 1
            preencher_cadastro(ws, linha, item, pdf_item, col_cad)
            valores = preencher_financeiro(ws, linha, pdf_item["lancamentos"], rubricas) if pdf_item else {}
            relatorio.append({"competencia": competencia, "arquivo": res.get("arquivo", ""), "formato": res.get("formato", ""), "nome_modelo": item.get("nome", ""),
                              "registro_modelo": item.get("registro", ""), "registro_pdf": pdf_item["matricula"] if pdf_item else "",
                              "metodo_match": metodo, "encontrado": bool(pdf_item),
                              "lancamentos": len(pdf_item["lancamentos"]) if pdf_item else 0,
                              "valor_total": sum(valores.values()) if valores else 0})

        _atualizar_totais_linhas_v6(ws)
        _atualizar_subtotais_v6(ws)
        print(f"{competencia}: bloco {inicio_novo}:{fim_novo} | população {desejada} | encontrados {encontrados}/{desejada}")
        processadas.append(competencia)
        inicio_base, fim_base = inicio_novo, fim_novo
        modelo_atual = ler_modelo_mensal(ws, inicio_base, fim_base, descobrir_colunas_cadastrais(ws))

    # Limpeza final: apenas rubricas financeiras vazias, nunca colunas cadastrais.
    removidas = _remover_rubricas_vazias_v6(ws)
    _atualizar_subtotais_v6(ws)
    _atualizar_totais_linhas_v6(ws)
    _limpar_linhas_vazias_apos_dados_v6(ws)
    _reconstruir_grupos_v6(ws)
    _atualizar_subtotais_v6(ws)
    _atualizar_totais_linhas_v6(ws)

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    ext = ".xlsm" if manter_vba else ".xlsx"
    base_nome = os.path.splitext(os.path.basename(caminho_excel))[0]
    saida = os.path.join(PASTA_TRABALHO, f"{base_nome}_V11_ATUALIZADA{ext}")
    if CRIAR_BACKUP:
        backup = os.path.join(PASTA_TRABALHO, f"{base_nome}_BACKUP_RH_VALORACAO{ext}")
        if not os.path.exists(backup):
            shutil.copy2(caminho_excel, backup)
    wb.save(saida)
    print("=" * 100)
    print("CONCLUÍDO — v11.0")
    print("Meses processados:", ", ".join(processadas) or "nenhum")
    print("Rubricas financeiras vazias removidas:", len(removidas))
    print("Arquivo:", saida)
    return saida, relatorio

# No Colab, a célula final chama:
# caminho_saida, relatorio = processar()

# ============================================================
# CAMADA COLAB V11 — ENTRADA ÚNICA + ZIP + DIAGNÓSTICO
# ============================================================
# A interface de uso permanece deliberadamente simples:
# uma única chamada files.upload(). O seletor do Colab aceita
# arquivos individuais e ZIPs. Pastas locais não são enviadas
# diretamente pelo seletor nativo do Colab; para uma pasta,
# compacte-a em ZIP ou use a opção de Google Drive separadamente.

def preparar_execucao_colab(uploaded):
    """Cria uma execução limpa e copia os arquivos enviados para ela."""
    global _ZIPS_EXTRAIDOS
    _ZIPS_EXTRAIDOS = set()
    if os.path.exists(PASTA_TRABALHO):
        shutil.rmtree(PASTA_TRABALHO)
    os.makedirs(PASTA_TRABALHO, exist_ok=True)

    for nome, conteudo in uploaded.items():
        destino = os.path.join(PASTA_TRABALHO, os.path.basename(nome))
        with open(destino, "wb") as f:
            f.write(conteudo)

    return PASTA_TRABALHO


def coletar_diagnostico_entrada():
    excels, pdfs = localizar_arquivos()
    return {
        "excel": [os.path.basename(x) for x in excels],
        "pdfs": [os.path.basename(x) for x in pdfs],
        "quantidade_excel": len(excels),
        "quantidade_pdf": len(pdfs),
    }


def gerar_relatorio_execucao(relatorio):
    """Consolida o retorno por competência/layout para inspeção rápida."""
    from collections import defaultdict
    grupos = defaultdict(lambda: {
        "competencia": "", "formato": "", "colaboradores_modelo": 0,
        "encontrados": 0, "nao_encontrados": 0, "lancamentos": 0,
        "valor_total_extraido": 0.0
    })
    for item in relatorio:
        chave = (item.get("competencia", ""), item.get("arquivo", ""))
        g = grupos[chave]
        g["competencia"] = item.get("competencia", "")
        g["formato"] = item.get("formato", "")
        g["arquivo"] = os.path.basename(item.get("arquivo", ""))
        g["colaboradores_modelo"] += 1
        if item.get("encontrado"):
            g["encontrados"] += 1
        else:
            g["nao_encontrados"] += 1
        g["lancamentos"] += int(item.get("lancamentos", 0) or 0)
        g["valor_total_extraido"] += float(item.get("valor_total", 0) or 0)
    saida = list(grupos.values())
    for g in saida:
        g["valor_total_extraido"] = round(g["valor_total_extraido"], 2)
    return saida


def executar_colab(uploaded=None):
    """Fluxo completo para uso no Colab.

    Se ``uploaded`` já foi fornecido pela célula de upload, nenhuma segunda
    janela é aberta. Se for omitido, o helper continua funcionando sozinho.
    """
    if uploaded is None:
        from google.colab import files
        print("Selecione o Excel-base, PDFs e/ou ZIPs em UMA ÚNICA janela.")
        uploaded = files.upload()
    if not uploaded:
        raise ValueError("Nenhum arquivo foi enviado.")

    preparar_execucao_colab(uploaded)
    diagnostico = coletar_diagnostico_entrada()
    print("\nArquivos identificados:")
    print("Excel:", diagnostico["excel"])
    print("PDFs:", diagnostico["pdfs"])

    if diagnostico["quantidade_excel"] != 1:
        raise ValueError(
            "A execução deve conter exatamente 1 Excel-base (.xlsx ou .xlsm). "
            f"Foram encontrados {diagnostico['quantidade_excel']}."
        )
    if diagnostico["quantidade_pdf"] == 0:
        raise ValueError("Nenhum PDF foi encontrado, inclusive após a extração dos ZIPs.")

    caminho_saida, relatorio = processar()
    resumo = gerar_relatorio_execucao(relatorio)

    print("\n" + "=" * 90)
    print("DIAGNÓSTICO DA EXECUÇÃO")
    print("=" * 90)
    for item in resumo:
        print(
            f"{item['competencia']} | {os.path.basename(item['arquivo'])} | "
            f"{item['colaboradores_modelo']} colaboradores | "
            f"encontrados {item['encontrados']} | "
            f"não encontrados {item['nao_encontrados']} | "
            f"lançamentos {item['lancamentos']}"
        )

    print("\nArquivo final:", caminho_saida)
    files.download(caminho_saida)
    return caminho_saida, resumo

