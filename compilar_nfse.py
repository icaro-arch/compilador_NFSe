#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Compilador de NFS-e (Padrão Nacional SPED nfse.gov.br) -> Excel (.xlsx) e CSV (.csv)
Possui Interface Gráfica (Tkinter) com pré-leitura inteligente da pasta de XMLs,
sugestão automática de nome de arquivo e geração de planilhas com SUBTOTAL dinâmico.
"""

import argparse
import csv
import os
import re
import sys
import threading
import unicodedata
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ==============================================================================
# UTILITÁRIOS DE FORMATAÇÃO E HIGIENIZAÇÃO
# ==============================================================================

CTRL_CHARS_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
MOJIBAKE_RE = re.compile(r"Ã[\x80-\xbf]|Â[\x80-\xbf]")


def decode_xml_bytes(raw_bytes: bytes) -> str:
    """Decodifica os bytes do XML garantindo preservação total de acentuação (UTF-8 ou CP1252/Latin-1)."""
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            text = raw_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = raw_bytes.decode("utf-8", errors="replace")
    # Remove eventual declaração de encoding conflitante caso o texto já tenha sido convertido para str Unicode
    text = re.sub(r'(<\?xml[^>]*encoding=[\'"])([^\'"]+)([\'"])', r"\1UTF-8\3", text, count=1, flags=re.IGNORECASE)
    return CTRL_CHARS_RE.sub("", text)


def clean_text(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    s = CTRL_CHARS_RE.sub("", s)
    # Corrige eventual mojibake (UTF-8 lido como Latin-1 na origem)
    if MOJIBAKE_RE.search(s):
        try:
            s = s.encode("latin-1").decode("utf-8")
        except Exception:
            pass
    # Normaliza espaços múltiplos e quebras de linha para evitar palavras quebradas nas células
    s = re.sub(r"\s+", " ", s).strip()
    return unicodedata.normalize("NFC", s)


def parse_float(val) -> float:
    if val is None or val == "":
        return 0.0
    s = str(val).strip().replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def format_cnpj_cpf(doc: str) -> str:
    d = re.sub(r"\D", "", clean_text(doc))
    if len(d) == 14:
        return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"
    if len(d) == 11:
        return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"
    return clean_text(doc)


def format_cep(cep: str) -> str:
    d = re.sub(r"\D", "", clean_text(cep))
    if len(d) == 8:
        return f"{d[:5]}-{d[5:]}"
    return clean_text(cep)


def format_date_str(d_str: str) -> str:
    s = clean_text(d_str)
    if not s:
        return ""
    # Formato simples YYYY-MM-DD
    if re.match(r"^\d{4}-\d{2}-\d{2}$", s):
        return f"{s[8:10]}/{s[5:7]}/{s[0:4]}"
    # Formato ISO com horário YYYY-MM-DDTHH:MM:SS...
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})", s)
    if m:
        yyyy, mm, dd, hh, mi, ss = m.groups()
        return f"{dd}/{mm}/{yyyy} {hh}:{mi}:{ss}"
    return s


def extract_competencia_ym(d_str: str) -> str:
    s = clean_text(d_str)
    if re.match(r"^\d{4}-\d{2}", s):
        return f"{s[5:7]}/{s[0:4]}"
    if re.match(r"^\d{2}/\d{2}/\d{4}", s):
        return f"{s[3:5]}/{s[6:10]}"
    return ""


def strip_ns(elem: ET.Element) -> None:
    """Remove namespaces de todas as tags da árvore XML para busca limpa."""
    for el in elem.iter():
        if isinstance(el.tag, str) and "}" in el.tag:
            el.tag = el.tag.split("}", 1)[1]


def find_text(elem: ET.Element, path: str, default: str = "") -> str:
    if elem is None:
        return default
    node = elem.find(path)
    if node is not None and node.text is not None:
        return clean_text(node.text)
    return default


def slugify_filename(text: str) -> str:
    """Converte razão social ou nome de pasta em identificador seguro para arquivo."""
    if not text:
        return "EMPRESA"
    norm = unicodedata.normalize("NFKD", text)
    ascii_str = norm.encode("ascii", "ignore").decode("ascii").upper()
    # Remove sufixos societários genéricos se ficarem muito longos, mas preserva clareza
    cleaned = re.sub(r"[^A-Z0-9]+", "_", ascii_str).strip("_")
    cleaned = re.sub(r"_+", "_", cleaned)
    return cleaned[:45].strip("_") or "EMPRESA"


# ==============================================================================
# LEITURA E EXTRAÇÃO DOS ARQUIVOS XML (NFS-e PADRÃO NACIONAL)
# ==============================================================================

SIMPLES_MAP = {
    "1": "1 - Não Optante",
    "2": "2 - Optante (ME/EPP)",
    "3": "3 - Optante (MEI)",
}

RET_ISS_MAP = {
    "1": "1 - Não Retido",
    "2": "2 - Retido na Fonte",
    "3": "3 - Retido Intermediário",
}

REG_ESP_MAP = {
    "0": "0 - Nenhum",
    "1": "1 - Ato Cooperado",
    "2": "2 - Estimativa",
    "3": "3 - Microempresa Municipal",
    "4": "4 - Notário ou Registrador",
    "5": "5 - Profissional Autônomo",
    "6": "6 - Sociedade de Profissionais",
}

# Aumenta o limite de tamanho de campo do módulo CSV para suportar XMLs completos dentro de células
try:
    csv.field_size_limit(min(sys.maxsize, 2147483647))
except OverflowError:
    csv.field_size_limit(2147483647)


def iter_xml_from_adn_csv(csv_path: Path):
    """
    Lê um arquivo CSV exportado do ADN (ex.: DADOS_ADN.csv delimitado por '|' ou ';')
    que contenha uma coluna com o XML da NFS-e (ex.: 'arquivo_xml_str').
    Produz tuplas: (identificador_origem, texto_xml, metadados_adn).
    """
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            with open(csv_path, mode="r", encoding=enc, newline="") as f:
                first_line = f.readline()
                if not first_line:
                    return
                second_line = f.readline()
            break
        except UnicodeDecodeError:
            continue
    else:
        return

    # Detecta delimitador predominante no cabeçalho
    delim = "|"
    for candidate_delim in ("|", ";", ",", "\t"):
        if candidate_delim in first_line:
            delim = candidate_delim
            break

    # Verifica rapidamente se o CSV aparenta conter XML de NFS-e
    combined_sample = (first_line + "\n" + second_line).lower()
    if "xml" not in combined_sample and "<nfse" not in combined_sample and "<infnfse" not in combined_sample:
        return

    with open(csv_path, mode="r", encoding=enc, errors="replace", newline="") as f:
        reader = csv.reader(f, delimiter=delim)
        try:
            header = [h.strip().lower() for h in next(reader)]
        except StopIteration:
            return

        xml_col_idx = None
        nsu_col_idx = None
        id_col_idx = None

        for i, col_name in enumerate(header):
            if "xml" in col_name:
                xml_col_idx = i
            elif col_name == "nsu":
                nsu_col_idx = i
            elif col_name == "id":
                id_col_idx = i

        for row_num, row in enumerate(reader, start=2):
            if not row:
                continue

            # Se a coluna XML não tinha 'xml' no nome do cabeçalho, detecta pelo conteúdo da célula
            if xml_col_idx is None:
                for i, cell_val in enumerate(row):
                    if "<NFSe" in cell_val or "<infNFSe" in cell_val or "<?xml" in cell_val:
                        xml_col_idx = i
                        break
                if xml_col_idx is None:
                    continue

            if xml_col_idx >= len(row):
                continue

            raw_xml_str = row[xml_col_idx].strip()
            if not raw_xml_str or ("<" not in raw_xml_str):
                continue

            nsu_val = row[nsu_col_idx].strip() if (nsu_col_idx is not None and nsu_col_idx < len(row)) else ""
            id_val = row[id_col_idx].strip() if (id_col_idx is not None and id_col_idx < len(row)) else ""

            if nsu_val:
                src_label = f"{csv_path.name} [NSU {nsu_val}]"
            elif id_val:
                src_label = f"{csv_path.name} [ID {id_val}]"
            else:
                src_label = f"{csv_path.name}#L{row_num}"

            text_xml = CTRL_CHARS_RE.sub("", raw_xml_str)
            text_xml = re.sub(
                r'(<\?xml[^>]*encoding=[\'"])([^\'"]+)([\'"])',
                r"\1UTF-8\3",
                text_xml,
                count=1,
                flags=re.IGNORECASE,
            )
            yield (src_label, text_xml, {"nsu": nsu_val, "id_adn": id_val})


def collect_xml_entries(source_path: str) -> tuple[list[tuple[str, str, dict]], str]:
    """
    Aceita tanto uma pasta (contendo arquivos .xml ou arquivos .csv do ADN)
    quanto o caminho direto de um arquivo .csv ou .xml.
    Retorna (lista_de_entradas, descricao_origem).
    """
    p = Path(source_path)
    if not p.exists():
        return [], "Caminho não encontrado."

    entries = []
    if p.is_file():
        ext = p.suffix.lower()
        if ext == ".csv":
            entries.extend(iter_xml_from_adn_csv(p))
            if not entries:
                return [], "O arquivo CSV selecionado não contém coluna com XML de NFS-e (padrão ADN)."
            return entries, "CSV ADN"
        elif ext == ".xml":
            try:
                entries.append((p.name, decode_xml_bytes(p.read_bytes()), {}))
                return entries, "Arquivo XML"
            except Exception as exc:
                return [], f"Erro ao ler XML: {exc}"
        else:
            return [], "Formato de arquivo não suportado (selecione uma pasta, .csv do ADN ou .xml)."

    # Se for pasta: procura primeiro arquivos .xml; se não houver (ou em conjunto), procura CSVs do ADN
    xml_files = sorted(p.glob("*.xml"))
    for f in xml_files:
        try:
            entries.append((f.name, decode_xml_bytes(f.read_bytes()), {}))
        except Exception:
            continue

    csv_files = sorted(p.glob("*.csv"))
    csv_adn_count = 0
    for cf in csv_files:
        csv_entries = list(iter_xml_from_adn_csv(cf))
        if csv_entries:
            csv_adn_count += 1
            entries.extend(csv_entries)

    if not entries:
        return [], "Nenhum arquivo .xml ou .csv do ADN encontrado na pasta selecionada."

    if xml_files and csv_adn_count:
        origem_tipo = "XMLs + CSV ADN"
    elif csv_adn_count:
        origem_tipo = "CSV ADN"
    else:
        origem_tipo = "XMLs"
    return entries, origem_tipo


def scan_folder_quick(source_path: str) -> dict:
    """
    Faz uma pré-leitura rápida da pasta ou arquivo (.csv ADN / .xml) selecionado
    para alimentar a interface gráfica e sugerir automaticamente o nome ideal para o Excel.
    """
    p = Path(source_path)
    if not p.exists():
        return {"valid": False, "count": 0, "error": "Caminho não encontrado."}

    entries, origem_info = collect_xml_entries(source_path)
    if not entries:
        return {"valid": False, "count": 0, "error": origem_info}

    tomadores = Counter()
    prestadores = Counter()
    competencias = Counter()
    municipios = Counter()
    total_servicos = 0.0
    valid_count = 0

    for _, text_xml, _ in entries:
        try:
            root = ET.fromstring(text_xml)
            strip_ns(root)

            inf = root if root.tag == "infNFSe" else root.find(".//infNFSe")
            if inf is None:
                continue

            valid_count += 1
            toma_nome = find_text(inf, "DPS/infDPS/toma/xNome")
            prest_nome = find_text(inf, "emit/xNome")
            d_comp = find_text(inf, "DPS/infDPS/dCompet") or find_text(inf, "DPS/infDPS/dhEmi")
            mun_prest = find_text(inf, "xLocPrestacao")
            v_serv = parse_float(find_text(inf, "DPS/infDPS/valores/vServPrest/vServ"))

            if toma_nome:
                tomadores[toma_nome] += 1
            if prest_nome:
                prestadores[prest_nome] += 1
            ym = extract_competencia_ym(d_comp)
            if ym:
                competencias[ym] += 1
            if mun_prest:
                municipios[mun_prest] += 1
            total_servicos += v_serv
        except Exception:
            continue

    if valid_count == 0:
        return {
            "valid": False,
            "count": 0,
            "error": "Os registros encontrados não seguem o padrão XML NFS-e esperado.",
        }

    # Determina empresa principal (Tomador predominante ou Prestador predominante)
    top_toma, toma_cnt = tomadores.most_common(1)[0] if tomadores else ("", 0)
    top_prest, prest_cnt = prestadores.most_common(1)[0] if prestadores else ("", 0)

    if toma_cnt >= prest_cnt and top_toma:
        empresa_principal = top_toma
        papel_empresa = "Tomador"
    else:
        empresa_principal = top_prest
        papel_empresa = "Prestador"

    # Período / Competência
    anos = sorted({ym.split("/")[1] for ym in competencias if "/" in ym})
    meses_anos = sorted(competencias.keys(), key=lambda x: (x.split("/")[1], x.split("/")[0]))
    if len(meses_anos) == 1:
        periodo_str = meses_anos[0].replace("/", "_")
        periodo_label = meses_anos[0]
    elif len(anos) == 1:
        periodo_str = anos[0]
        periodo_label = f"{meses_anos[0]} a {meses_anos[-1]} ({anos[0]})"
    elif anos:
        periodo_str = f"{anos[0]}_{anos[-1]}"
        periodo_label = f"{meses_anos[0]} a {meses_anos[-1]}"
    else:
        periodo_str = datetime.now().strftime("%Y")
        periodo_label = periodo_str

    # Sugestões inteligentes de nome (baseadas no nome do arquivo CSV ou da pasta)
    base_name_raw = p.stem.upper() if p.is_file() else p.name.upper()
    for prefix in ("XML NOTAS ", "XML NOTA ", "XML NFSE ", "XML_NOTA_", "XML_NOTAS_", "XML ", "NOTAS "):
        if base_name_raw.startswith(prefix):
            base_name_raw = base_name_raw[len(prefix):].strip()

    slug_folder = slugify_filename(base_name_raw)
    slug_empresa = slugify_filename(empresa_principal)

    sugestao_pasta = f"NOTAS_FISCAIS_{slug_folder}.xlsx" if slug_folder else f"NOTAS_FISCAIS_{slug_empresa}_{periodo_str}.xlsx"
    sugestao_empresa = f"NOTAS_FISCAIS_{slug_empresa}_{periodo_str}.xlsx"

    if slug_folder and slug_folder not in ("XML", "NOTAS", "NFSE", "SEFIN"):
        nome_sugerido = sugestao_pasta
    else:
        nome_sugerido = sugestao_empresa

    top_mun = municipios.most_common(1)[0][0] if municipios else "N/I"

    return {
        "valid": True,
        "count": valid_count,
        "origem_tipo": origem_info,
        "total_xml_files": len(entries),
        "empresa_principal": empresa_principal,
        "papel_empresa": papel_empresa,
        "qtd_fornecedores": len(prestadores),
        "qtd_tomadores": len(tomadores),
        "periodo_label": periodo_label,
        "municipio_principal": top_mun,
        "total_servicos": total_servicos,
        "nome_sugerido": nome_sugerido,
        "sugestao_empresa": sugestao_empresa,
        "sugestao_pasta": sugestao_pasta,
    }


def load_all_xml_records(source_path: str, progress_cb=None) -> list[dict]:
    """Lê todos os XMLs (de uma pasta ou de um CSV do ADN) e retorna a lista de dicionários estruturados."""
    entries, _ = collect_xml_entries(source_path)
    total_files = len(entries)
    step_notify = max(10, total_files // 40) if total_files > 0 else 10

    parsed_items = []
    chave_counts = Counter()

    for idx, (src_label, text_xml, meta) in enumerate(entries, start=1):
        if progress_cb and (idx % step_notify == 0 or idx == total_files):
            progress_cb(idx, total_files, f"Processando nota ({idx}/{total_files}): {src_label}")

        try:
            root = ET.fromstring(text_xml)
            strip_ns(root)
        except Exception as exc:
            print(f"[AVISO] Falha ao ler {src_label}: {exc}", file=sys.stderr)
            continue

        inf = root if root.tag == "infNFSe" else root.find(".//infNFSe")
        if inf is None:
            continue

        chave_id = clean_text(inf.attrib.get("Id", ""))
        chave_counts[chave_id] += 1
        parsed_items.append((src_label, inf, chave_id, meta))

    seen_chaves = set()
    records = []

    for src_label, inf, chave_id, meta in parsed_items:
        dps = inf.find("DPS/infDPS")
        serv = dps.find("serv") if dps is not None else None
        c_serv = serv.find("cServ") if serv is not None else None
        val = inf.find("valores")
        dps_val = dps.find("valores") if dps is not None else None
        tf = dps_val.find("trib/tribFed") if dps_val is not None else None
        tm = dps_val.find("trib/tribMun") if dps_val is not None else None
        piscof = tf.find("piscofins") if tf is not None else None
        ibscbs = inf.find("IBSCBS")

        if chave_counts[chave_id] > 1:
            if chave_id in seen_chaves:
                status_dup = "Duplicada"
            else:
                status_dup = "Original (Possui Cópia)"
        else:
            status_dup = "Única"
        seen_chaves.add(chave_id)

        op_simp_raw = find_text(dps, "prest/regTrib/opSimpNac")
        regime_simples = SIMPLES_MAP.get(op_simp_raw, op_simp_raw or "Não Informado")

        reg_esp_raw = find_text(dps, "prest/regTrib/regEspTrib")
        regime_especial = REG_ESP_MAP.get(reg_esp_raw, reg_esp_raw or "0 - Nenhum")

        tp_ret_raw = find_text(tm, "tpRetISSQN")
        tipo_ret_iss = RET_ISS_MAP.get(tp_ret_raw, tp_ret_raw or "1 - Não Retido")

        # Obra
        c_obra = find_text(serv, "obra/cObra")
        obra_end_node = serv.find("obra/end") if serv is not None else None
        if obra_end_node is not None:
            parts_obra = [
                find_text(obra_end_node, "xLgr"),
                find_text(obra_end_node, "nro"),
                find_text(obra_end_node, "xBairro"),
                format_cep(find_text(obra_end_node, "CEP")),
            ]
            end_obra = ", ".join(p for p in parts_obra if p)
        else:
            end_obra = ""

        # Endereço prestador
        ep_node = inf.find("emit/enderNac")
        if ep_node is not None:
            parts_ep = [
                find_text(ep_node, "xLgr"),
                find_text(ep_node, "nro"),
                find_text(ep_node, "xCpl"),
            ]
            end_prest = ", ".join(p for p in parts_ep if p)
        else:
            end_prest = ""

        # Endereço tomador
        et_node = dps.find("toma/end") if dps is not None else None
        if et_node is not None:
            parts_et = [
                find_text(et_node, "xLgr"),
                find_text(et_node, "nro"),
                find_text(et_node, "xCpl"),
            ]
            end_toma = ", ".join(p for p in parts_et if p)
        else:
            end_toma = ""

        # Valores monetários e alíquotas
        v_serv = parse_float(find_text(dps_val, "vServPrest/vServ"))
        v_dr = parse_float(find_text(dps_val, "vDedRed/vDR"))
        v_desc = parse_float(find_text(dps_val, "vDescCondIncond/vDescIncond"))
        v_bc_raw = find_text(val, "vBC")
        v_bc = parse_float(v_bc_raw) if v_bc_raw else max(0.0, v_serv - v_dr - v_desc)

        p_aliq_str = find_text(val, "pAliqAplic") or find_text(tm, "pAliq")
        p_aliq = parse_float(p_aliq_str)
        v_iss = parse_float(find_text(val, "vISSQN"))

        v_ret_cp = parse_float(find_text(tf, "vRetCP"))
        v_ret_irrf = parse_float(find_text(tf, "vRetIRRF"))
        v_ret_csll = parse_float(find_text(tf, "vRetCSLL"))
        v_bc_piscof = parse_float(find_text(piscof, "vBCPisCofins"))
        p_aliq_pis = parse_float(find_text(piscof, "pAliqPis"))
        v_pis = parse_float(find_text(piscof, "vPis"))
        p_aliq_cof = parse_float(find_text(piscof, "pAliqCofins"))
        v_cofins = parse_float(find_text(piscof, "vCofins"))

        v_tot_ret = parse_float(find_text(val, "vTotalRet"))
        v_liq_raw = find_text(val, "vLiq")
        v_liq = parse_float(v_liq_raw) if v_liq_raw else max(0.0, v_serv - v_tot_ret)

        # IBS / CBS
        p_ibs_uf = parse_float(find_text(ibscbs, "valores/uf/pIBSUF"))
        p_ibs_mun = parse_float(find_text(ibscbs, "valores/mun/pIBSMun"))
        p_cbs = parse_float(find_text(ibscbs, "valores/fed/pCBS"))
        v_ibs_tot = parse_float(find_text(ibscbs, "totCIBS/gIBS/vIBSTot"))
        v_cbs_tot = parse_float(find_text(ibscbs, "totCIBS/gCBS/vCBS"))

        # Resumo de Retenções Federais para filtro rápido
        fed_tags = []
        if v_ret_cp > 0:
            fed_tags.append("INSS")
        if v_ret_irrf > 0:
            fed_tags.append("IRRF")
        if v_ret_csll > 0:
            fed_tags.append("CSLL")
        if v_pis > 0 or v_cofins > 0:
            fed_tags.append("PIS/COFINS")
        ret_fed_status = " + ".join(fed_tags) if fed_tags else "Sem Retenção Federal"

        cnpj_prest_raw = find_text(inf, "emit/CNPJ") or find_text(inf, "emit/CPF") or find_text(dps, "prest/CNPJ")
        cnpj_toma_raw = find_text(dps, "toma/CNPJ") or find_text(dps, "toma/CPF")

        rec = {
            "Arquivo": src_label,
            "NSU": meta.get("nsu", ""),
            "StatusDuplicidade": status_dup,
            "ChaveNFSe": chave_id,
            "NumeroNFSe": find_text(inf, "nNFSe"),
            "NumeroDFSe": find_text(inf, "nDFSe"),
            "SerieDPS": find_text(dps, "serie"),
            "NumeroDPS": find_text(dps, "nDPS"),
            "DataEmissaoDPS": format_date_str(find_text(dps, "dhEmi")),
            "DataProcessamento": format_date_str(find_text(inf, "dhProc")),
            "DataCompetencia": format_date_str(find_text(dps, "dCompet")),
            "Status": find_text(inf, "cStat"),
            "CNPJPrestador": format_cnpj_cpf(cnpj_prest_raw),
            "RazaoSocialPrestador": find_text(inf, "emit/xNome"),
            "NomeFantasiaPrestador": find_text(inf, "emit/xFant"),
            "InscricaoMunPrestador": find_text(inf, "emit/IM") or find_text(dps, "prest/IM"),
            "RegimeSimples": regime_simples,
            "RegimeEspecial": regime_especial,
            "EnderecoPrestador": end_prest,
            "BairroPrestador": find_text(ep_node, "xBairro"),
            "MunicipioPrestador": find_text(inf, "xLocEmi"),
            "UFPrestador": find_text(ep_node, "UF"),
            "CEPPrestador": format_cep(find_text(ep_node, "CEP")),
            "TelefonePrestador": find_text(inf, "emit/fone"),
            "EmailPrestador": find_text(inf, "emit/email"),
            "CNPJTomador": format_cnpj_cpf(cnpj_toma_raw),
            "RazaoSocialTomador": find_text(dps, "toma/xNome"),
            "EnderecoTomador": end_toma,
            "BairroTomador": find_text(et_node, "xBairro"),
            "MunicipioTomador": find_text(et_node, "endNac/cMun"),
            "CEPTomador": format_cep(find_text(et_node, "endNac/CEP") or find_text(et_node, "CEP")),
            "TelefoneTomador": find_text(dps, "toma/fone"),
            "EmailTomador": find_text(dps, "toma/email"),
            "MunicipioPrestacao": find_text(inf, "xLocPrestacao") or "Não Informado",
            "CodMunPrestacao": find_text(serv, "locPrest/cLocPrestacao"),
            "MunicipioIncidencia": find_text(inf, "xLocIncid"),
            "CodMunIncidencia": find_text(inf, "cLocIncid") or find_text(tm, "cLocIncid"),
            "CodObra": c_obra,
            "EnderecoObra": end_obra,
            "CodTribNac": find_text(c_serv, "cTribNac"),
            "DescTribNac": find_text(inf, "xTribNac"),
            "CodTribMun": find_text(c_serv, "cTribMun"),
            "DescTribMun": find_text(inf, "xTribMun"),
            "CodNBS": find_text(c_serv, "cNBS"),
            "DescNBS": find_text(inf, "xNBS"),
            "DiscriminacaoServico": find_text(c_serv, "xDescServ"),
            "InfoComplementar": find_text(serv, "infoCompl/xInfComp"),
            "TipoRetencaoISS": tipo_ret_iss,
            "ValorServicos": v_serv,
            "DeducoesReducoes": v_dr,
            "DescontoIncond": v_desc,
            "BaseCalculoISS": v_bc,
            "AliquotaISS": p_aliq,
            "ValorISS": v_iss,
            "RetencoesFederaisAplicadas": ret_fed_status,
            "RetencaoINSS": v_ret_cp,
            "RetencaoIRRF": v_ret_irrf,
            "RetencaoCSLL": v_ret_csll,
            "CSTPisCofins": find_text(piscof, "CST"),
            "BasePisCofins": v_bc_piscof,
            "AliquotaPIS": p_aliq_pis,
            "ValorPIS": v_pis,
            "AliquotaCOFINS": p_aliq_cof,
            "ValorCOFINS": v_cofins,
            "TotalRetencoes": v_tot_ret,
            "ValorLiquido": v_liq,
            "AliquotaIBSUF": p_ibs_uf,
            "AliquotaIBSMun": p_ibs_mun,
            "AliquotaCBS": p_cbs,
            "ValorIBSTotal": v_ibs_tot,
            "ValorCBSTotal": v_cbs_tot,
            "HasIBSCBS": ibscbs is not None,
        }
        records.append(rec)

    # Ordena por Data de Competência e Número da NFS-e (numérico quando possível)
    def sort_key(r):
        num = r.get("NumeroNFSe", "")
        try:
            n_val = int(num)
        except ValueError:
            n_val = 0
        return (n_val, r.get("Arquivo", ""))

    records.sort(key=sort_key)
    return records


# ==============================================================================
# DEFINIÇÃO FUNCIONAL E ORGANIZADA DAS COLUNAS DA ABA 'NOTAS FISCAIS'
# ==============================================================================

# Grupos de cores de cabeçalho para facilitar localização visual imediata:
# - IDENT: Azul Marinho (#1F4E79)
# - PREST_LOCAL: Azul Corporativo (#2F5597)
# - ISS_VAL: Verde Petróleo Escuro (#1B6B53)
# - FED_LIQ: Âmbar / Bordô Corporativo (#833C0C)
# - IBSCBS: Roxo Executivo (#4A3B7C)
# - SERV: Azul Petróleo (#2B5868)
# - CAD_AUX: Cinza Grafite (#4A5568)

NOTAS_COL_DEFS = [
    # BLOCO 1: IDENTIFICAÇÃO ESSENCIAL (Fixado nas 3 primeiras colunas para rolagem)
    {"header": "Número NFS-e", "prop": "NumeroNFSe", "width": 15, "align": "center", "type": "text", "group_color": "1F4E79", "outline": 0},
    {"header": "Competência", "prop": "DataCompetencia", "width": 14, "align": "center", "type": "text", "group_color": "1F4E79", "outline": 0},
    {"header": "Razão Social Prestador", "prop": "RazaoSocialPrestador", "width": 38, "align": "left", "type": "text", "group_color": "1F4E79", "outline": 0},
    {"header": "CNPJ Prestador", "prop": "CNPJPrestador", "width": 20, "align": "center", "type": "text", "group_color": "1F4E79", "outline": 0},
    {"header": "Data Emissão DPS", "prop": "DataEmissaoDPS", "width": 20, "align": "center", "type": "text", "group_color": "1F4E79", "outline": 0},
    {"header": "Status Duplicidade", "prop": "StatusDuplicidade", "width": 18, "align": "center", "type": "text", "group_color": "1F4E79", "outline": 0},
    {"header": "Status (cStat)", "prop": "Status", "width": 13, "align": "center", "type": "text", "group_color": "1F4E79", "outline": 0},

    # BLOCO 2: ENQUADRAMENTO E LOCAL DE EXECUÇÃO
    {"header": "Regime Tributário", "prop": "RegimeSimples", "width": 22, "align": "left", "type": "text", "group_color": "2F5597", "outline": 0},
    {"header": "Regime Especial", "prop": "RegimeEspecial", "width": 22, "align": "left", "type": "text", "group_color": "2F5597", "outline": 1},
    {"header": "Município Prestação", "prop": "MunicipioPrestacao", "width": 22, "align": "left", "type": "text", "group_color": "2F5597", "outline": 0},
    {"header": "Município Incidência", "prop": "MunicipioIncidencia", "width": 22, "align": "left", "type": "text", "group_color": "2F5597", "outline": 1},
    {"header": "Cód. Obra", "prop": "CodObra", "width": 16, "align": "center", "type": "text", "group_color": "2F5597", "outline": 0},

    # BLOCO 3: VALORES E TRIBUTAÇÃO MUNICIPAL (ISS / ISSQN)
    {"header": "Tipo Retenção ISS", "prop": "TipoRetencaoISS", "width": 22, "align": "center", "type": "text", "group_color": "1B6B53", "outline": 0},
    {"header": "Valor Serviços (R$)", "prop": "ValorServicos", "width": 20, "align": "right", "type": "currency", "is_sum": True, "group_color": "1B6B53", "outline": 0},
    {"header": "Deduções/Reduções (R$)", "prop": "DeducoesReducoes", "width": 20, "align": "right", "type": "currency", "is_sum": True, "group_color": "1B6B53", "outline": 0},
    {"header": "Desconto Incond. (R$)", "prop": "DescontoIncond", "width": 20, "align": "right", "type": "currency", "is_sum": True, "group_color": "1B6B53", "outline": 0},
    {"header": "Base Cálculo ISS (R$)", "prop": "BaseCalculoISS", "width": 20, "align": "right", "type": "currency", "is_sum": True, "group_color": "1B6B53", "outline": 0},
    {"header": "Alíquota ISS (%)", "prop": "AliquotaISS", "width": 16, "align": "right", "type": "rate", "group_color": "1B6B53", "outline": 0},
    {"header": "Valor ISSQN (R$)", "prop": "ValorISS", "width": 19, "align": "right", "type": "currency", "is_sum": True, "group_color": "1B6B53", "outline": 0},

    # BLOCO 4: RETENÇÕES FEDERAIS E VALOR LÍQUIDO
    {"header": "Retenções Federais Aplicadas", "prop": "RetencoesFederaisAplicadas", "width": 26, "align": "center", "type": "text", "group_color": "833C0C", "outline": 0},
    {"header": "Retenção INSS (R$)", "prop": "RetencaoINSS", "width": 19, "align": "right", "type": "currency", "is_sum": True, "group_color": "833C0C", "outline": 0},
    {"header": "Retenção IRRF (R$)", "prop": "RetencaoIRRF", "width": 19, "align": "right", "type": "currency", "is_sum": True, "group_color": "833C0C", "outline": 0},
    {"header": "Retenção CSLL (R$)", "prop": "RetencaoCSLL", "width": 19, "align": "right", "type": "currency", "is_sum": True, "group_color": "833C0C", "outline": 0},
    {"header": "Valor PIS (R$)", "prop": "ValorPIS", "width": 18, "align": "right", "type": "currency", "is_sum": True, "group_color": "833C0C", "outline": 0},
    {"header": "Valor COFINS (R$)", "prop": "ValorCOFINS", "width": 18, "align": "right", "type": "currency", "is_sum": True, "group_color": "833C0C", "outline": 0},
    {"header": "CST PIS/COFINS", "prop": "CSTPisCofins", "width": 16, "align": "center", "type": "text", "group_color": "833C0C", "outline": 1},
    {"header": "Base PIS/COFINS (R$)", "prop": "BasePisCofins", "width": 20, "align": "right", "type": "currency", "is_sum": True, "group_color": "833C0C", "outline": 1},
    {"header": "Alíquota PIS (%)", "prop": "AliquotaPIS", "width": 16, "align": "right", "type": "rate", "group_color": "833C0C", "outline": 1},
    {"header": "Alíquota COFINS (%)", "prop": "AliquotaCOFINS", "width": 18, "align": "right", "type": "rate", "group_color": "833C0C", "outline": 1},
    {"header": "Total Retenções (R$)", "prop": "TotalRetencoes", "width": 20, "align": "right", "type": "currency", "is_sum": True, "group_color": "833C0C", "outline": 0},
    {"header": "Valor Líquido (R$)", "prop": "ValorLiquido", "width": 21, "align": "right", "type": "currency", "is_sum": True, "group_color": "833C0C", "outline": 0},

    # BLOCO 5: REFORMA TRIBUTÁRIA (IBS / CBS)
    {"header": "IBS Alíquota UF (%)", "prop": "AliquotaIBSUF", "width": 18, "align": "right", "type": "rate", "group_color": "4A3B7C", "outline": 1},
    {"header": "IBS Alíquota Mun (%)", "prop": "AliquotaIBSMun", "width": 18, "align": "right", "type": "rate", "group_color": "4A3B7C", "outline": 1},
    {"header": "CBS Alíquota (%)", "prop": "AliquotaCBS", "width": 18, "align": "right", "type": "rate", "group_color": "4A3B7C", "outline": 1},
    {"header": "IBS Total (R$)", "prop": "ValorIBSTotal", "width": 18, "align": "right", "type": "currency", "is_sum": True, "group_color": "4A3B7C", "outline": 0},
    {"header": "CBS Total (R$)", "prop": "ValorCBSTotal", "width": 18, "align": "right", "type": "currency", "is_sum": True, "group_color": "4A3B7C", "outline": 0},

    # BLOCO 6: DETALHAMENTO DO SERVIÇO E OBRA
    {"header": "Discriminação dos Serviços", "prop": "DiscriminacaoServico", "width": 55, "align": "left", "type": "text", "group_color": "2B5868", "outline": 0},
    {"header": "Desc. Trib. Nacional", "prop": "DescTribNac", "width": 38, "align": "left", "type": "text", "group_color": "2B5868", "outline": 0},
    {"header": "Cód. Trib. Nacional", "prop": "CodTribNac", "width": 18, "align": "center", "type": "text", "group_color": "2B5868", "outline": 1},
    {"header": "Desc. Trib. Municipal", "prop": "DescTribMun", "width": 36, "align": "left", "type": "text", "group_color": "2B5868", "outline": 1},
    {"header": "Cód. Trib. Municipal", "prop": "CodTribMun", "width": 18, "align": "center", "type": "text", "group_color": "2B5868", "outline": 1},
    {"header": "Desc. NBS", "prop": "DescNBS", "width": 32, "align": "left", "type": "text", "group_color": "2B5868", "outline": 1},
    {"header": "Cód. NBS", "prop": "CodNBS", "width": 16, "align": "center", "type": "text", "group_color": "2B5868", "outline": 1},
    {"header": "Info. Complementares", "prop": "InfoComplementar", "width": 45, "align": "left", "type": "text", "group_color": "2B5868", "outline": 0},
    {"header": "Endereço da Obra", "prop": "EnderecoObra", "width": 35, "align": "left", "type": "text", "group_color": "2B5868", "outline": 0},

    # BLOCO 7: DADOS CADASTRAIS COMPLETOS, TOMADOR E CHAVES (Agrupados para recolher/expandir com + / -)
    {"header": "Nome Fantasia Prestador", "prop": "NomeFantasiaPrestador", "width": 28, "align": "left", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "Inscr. Mun. Prestador", "prop": "InscricaoMunPrestador", "width": 18, "align": "center", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "Endereço Prestador", "prop": "EnderecoPrestador", "width": 34, "align": "left", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "Bairro Prestador", "prop": "BairroPrestador", "width": 22, "align": "left", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "Município Prestador", "prop": "MunicipioPrestador", "width": 22, "align": "left", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "UF Prestador", "prop": "UFPrestador", "width": 12, "align": "center", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "CEP Prestador", "prop": "CEPPrestador", "width": 14, "align": "center", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "Telefone Prestador", "prop": "TelefonePrestador", "width": 16, "align": "center", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "E-mail Prestador", "prop": "EmailPrestador", "width": 30, "align": "left", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "CNPJ Tomador", "prop": "CNPJTomador", "width": 20, "align": "center", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "Razão Social Tomador", "prop": "RazaoSocialTomador", "width": 32, "align": "left", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "Endereço Tomador", "prop": "EnderecoTomador", "width": 32, "align": "left", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "Bairro Tomador", "prop": "BairroTomador", "width": 20, "align": "left", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "Cód. Mun. Tomador", "prop": "MunicipioTomador", "width": 18, "align": "center", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "CEP Tomador", "prop": "CEPTomador", "width": 14, "align": "center", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "Telefone Tomador", "prop": "TelefoneTomador", "width": 16, "align": "center", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "E-mail Tomador", "prop": "EmailTomador", "width": 30, "align": "left", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "Cód. Mun. Prestação", "prop": "CodMunPrestacao", "width": 18, "align": "center", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "Cód. Mun. Incidência", "prop": "CodMunIncidencia", "width": 18, "align": "center", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "Data Processamento", "prop": "DataProcessamento", "width": 20, "align": "center", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "Número DFSe", "prop": "NumeroDFSe", "width": 14, "align": "center", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "Série DPS", "prop": "SerieDPS", "width": 12, "align": "center", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "Número DPS", "prop": "NumeroDPS", "width": 14, "align": "center", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "NSU (ADN)", "prop": "NSU", "width": 16, "align": "center", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "Chave NFS-e", "prop": "ChaveNFSe", "width": 52, "align": "center", "type": "text", "group_color": "4A5568", "outline": 1},
    {"header": "Arquivo / Origem", "prop": "Arquivo", "width": 32, "align": "left", "type": "text", "group_color": "4A5568", "outline": 1},
]


# ==============================================================================
# CONSTRUÇÃO DO ARQUIVO EXCEL (.XLSX) E (.CSV)
# ==============================================================================

def export_to_csv(records: list[dict], csv_path: str) -> None:
    """Exporta os registros em CSV UTF-8 com BOM (abre diretamente no Excel com acentos perfeitos)."""
    headers = [c["header"] for c in NOTAS_COL_DEFS]
    props = [c["prop"] for c in NOTAS_COL_DEFS]
    with open(csv_path, mode="w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(headers)
        for r in records:
            row = []
            for c_def, prop in zip(NOTAS_COL_DEFS, props):
                val = r.get(prop, "")
                if c_def["type"] in ("currency", "rate") and isinstance(val, (int, float)):
                    row.append(f"{val:.2f}".replace(".", ","))
                else:
                    row.append(val)
            writer.writerow(row)


def build_excel_workbook(records: list[dict], output_xlsx: str, progress_cb=None) -> dict:
    """Constrói a pasta de trabalho Excel completa com 4 abas, estilos e SUBTOTAL dinâmico."""
    if not records:
        raise ValueError("Nenhum registro de NFS-e encontrado para gerar a planilha.")

    wb = Workbook()
    # Remove aba padrão inicial
    wb.remove(wb.active)

    # Paleta de Estilos OpenPyXL
    font_family = "Calibri"
    thin_side = Side(style="thin", color="D9D9D9")
    med_navy = Side(style="medium", color="1F4E79")
    double_bottom = Side(style="double", color="1F4E79")
    top_thin_navy = Side(style="thin", color="1F4E79")

    border_data = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_total = Border(top=top_thin_navy, bottom=double_bottom, left=thin_side, right=thin_side)
    border_kpi = Border(left=med_navy, right=thin_side, top=thin_side, bottom=thin_side)

    fill_zebra_even = PatternFill("solid", fgColor="FFFFFF")
    fill_zebra_odd = PatternFill("solid", fgColor="F4F7FA")
    fill_total = PatternFill("solid", fgColor="D9E1F2")
    fill_subhdr = PatternFill("solid", fgColor="EDEDED")
    fill_sec_hdr = PatternFill("solid", fgColor="2F5597")
    fill_title = PatternFill("solid", fgColor="1F4E79")
    fill_kpi_lbl = PatternFill("solid", fgColor="F2F4F8")

    # Destaques semânticos para Retenção de ISS
    fill_retido = PatternFill("solid", fgColor="FFF2CC")
    font_retido = Font(name=font_family, size=10, bold=True, color="833C0C")
    fill_nao_retido = PatternFill("solid", fgColor="E2EFDA")
    font_nao_retido = Font(name=font_family, size=10, bold=False, color="274E13")
    fill_dup_alert = PatternFill("solid", fgColor="FCE4D6")
    font_dup_alert = Font(name=font_family, size=10, bold=True, color="C00000")

    font_data = Font(name=font_family, size=10, color="111111")
    font_data_bold = Font(name=font_family, size=10, bold=True, color="1F4E79")
    font_total = Font(name=font_family, size=10, bold=True, color="000000")
    font_sec_hdr = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    font_subhdr = Font(name=font_family, size=10, bold=True, color="1F4E79")

    num_fmt_currency = '#,##0.00'
    num_fmt_rate = '0.00"%"'
    num_fmt_pct = '0.00%'
    num_fmt_int = '#,##0'

    # Mapa de letra de coluna por propriedade na aba 'Notas Fiscais'
    prop_to_col_letter = {}
    for idx, c_def in enumerate(NOTAS_COL_DEFS, start=1):
        prop_to_col_letter[c_def["prop"]] = get_column_letter(idx)

    n_rows = len(records)
    data_start_row = 2
    data_end_row = n_rows + 1
    total_row_idx = n_rows + 2

    # Estatísticas agregadas para metadados e resumos
    total_servicos = sum(r["ValorServicos"] for r in records)
    total_deducoes = sum(r["DeducoesReducoes"] for r in records)
    total_desconto = sum(r["DescontoIncond"] for r in records)
    total_base_iss = sum(r["BaseCalculoISS"] for r in records)
    total_iss = sum(r["ValorISS"] for r in records)
    total_iss_retido = sum(r["ValorISS"] for r in records if r["TipoRetencaoISS"].startswith("2"))
    total_iss_nao_retido = sum(r["ValorISS"] for r in records if not r["TipoRetencaoISS"].startswith("2"))
    total_inss = sum(r["RetencaoINSS"] for r in records)
    total_irrf = sum(r["RetencaoIRRF"] for r in records)
    total_csll = sum(r["RetencaoCSLL"] for r in records)
    total_pis = sum(r["ValorPIS"] for r in records)
    total_cofins = sum(r["ValorCOFINS"] for r in records)
    total_retencoes = sum(r["TotalRetencoes"] for r in records)
    total_liquido = sum(r["ValorLiquido"] for r in records)
    total_ibs = sum(r["ValorIBSTotal"] for r in records)
    total_cbs = sum(r["ValorCBSTotal"] for r in records)

    count_obras = sum(1 for r in records if r["CodObra"])
    count_ibscbs = sum(1 for r in records if r["HasIBSCBS"])

    # Agrupamento por Fornecedor
    forn_map = defaultdict(lambda: {
        "CNPJ": "",
        "RazaoSocial": "",
        "RegimeSimples": "",
        "UF": "",
        "Municipio": "",
        "Qtd": 0,
        "TotalServ": 0.0,
        "ISSRetido": 0.0,
        "ISSNaoRetido": 0.0,
        "TotalISS": 0.0,
        "TotalINSS": 0.0,
        "TotalIRRF": 0.0,
        "TotalCSLL": 0.0,
        "TotalPIS": 0.0,
        "TotalCOF": 0.0,
        "TotalRet": 0.0,
        "TotalLiq": 0.0,
    })

    for r in records:
        key = r["CNPJPrestador"] or r["RazaoSocialPrestador"]
        fg = forn_map[key]
        if not fg["CNPJ"]:
            fg["CNPJ"] = r["CNPJPrestador"]
            fg["RazaoSocial"] = r["RazaoSocialPrestador"]
            fg["RegimeSimples"] = r["RegimeSimples"]
            fg["UF"] = r["UFPrestador"]
            fg["Municipio"] = r["MunicipioPrestador"]
        fg["Qtd"] += 1
        fg["TotalServ"] += r["ValorServicos"]
        if r["TipoRetencaoISS"].startswith("2"):
            fg["ISSRetido"] += r["ValorISS"]
        else:
            fg["ISSNaoRetido"] += r["ValorISS"]
        fg["TotalISS"] += r["ValorISS"]
        fg["TotalINSS"] += r["RetencaoINSS"]
        fg["TotalIRRF"] += r["RetencaoIRRF"]
        fg["TotalCSLL"] += r["RetencaoCSLL"]
        fg["TotalPIS"] += r["ValorPIS"]
        fg["TotalCOF"] += r["ValorCOFINS"]
        fg["TotalRet"] += r["TotalRetencoes"]
        fg["TotalLiq"] += r["ValorLiquido"]

    forn_groups = sorted(forn_map.values(), key=lambda x: x["TotalServ"], reverse=True)

    # Agrupamento por Município de Prestação
    mun_map = defaultdict(lambda: {
        "Municipio": "",
        "Qtd": 0,
        "TotalServ": 0.0,
        "ISSRetido": 0.0,
        "ISSNaoRetido": 0.0,
        "TotalISS": 0.0,
        "RetFederais": 0.0,
        "TotalRet": 0.0,
        "TotalLiq": 0.0,
    })
    for r in records:
        m_name = r["MunicipioPrestacao"] or "Não Informado"
        mg = mun_map[m_name]
        mg["Municipio"] = m_name
        mg["Qtd"] += 1
        mg["TotalServ"] += r["ValorServicos"]
        if r["TipoRetencaoISS"].startswith("2"):
            mg["ISSRetido"] += r["ValorISS"]
        else:
            mg["ISSNaoRetido"] += r["ValorISS"]
        mg["TotalISS"] += r["ValorISS"]
        mg["RetFederais"] += (
            r["RetencaoINSS"] + r["RetencaoIRRF"] + r["RetencaoCSLL"] + r["ValorPIS"] + r["ValorCOFINS"]
        )
        mg["TotalRet"] += r["TotalRetencoes"]
        mg["TotalLiq"] += r["ValorLiquido"]

    mun_groups = sorted(mun_map.values(), key=lambda x: x["TotalServ"], reverse=True)

    # Identifica Tomador e Município predominantes para o cabeçalho executivo
    toma_counter = Counter((r["RazaoSocialTomador"], r["CNPJTomador"]) for r in records if r["RazaoSocialTomador"])
    if toma_counter:
        (toma_nome_top, toma_cnpj_top), _ = toma_counter.most_common(1)[0]
    else:
        toma_nome_top, toma_cnpj_top = "Não Identificado", ""

    mun_principal = mun_groups[0]["Municipio"] if mun_groups else "N/I"
    anos_comp = sorted({r["DataCompetencia"][-4:] for r in records if len(r["DataCompetencia"]) >= 10})
    periodo_comp = "/".join(anos_comp) if anos_comp else "2026"

    # ==========================================================================
    # ABA 1: RESUMO FINANCEIRO CONSOLIDADO
    # ==========================================================================
    if progress_cb:
        progress_cb(1, 4, "Montando Aba 1: Resumo Financeiro Consolidado...")

    ws1 = wb.create_sheet(title="Resumo Financeiro Consolidado")
    ws1.views.sheetView[0].showGridLines = True

    ws1_col_widths = {1: 4, 2: 56, 3: 24, 4: 18, 5: 22, 6: 22, 7: 16}
    for col_idx, w in ws1_col_widths.items():
        ws1.column_dimensions[get_column_letter(col_idx)].width = w

    # Bloco de Título
    ws1.row_dimensions[2].height = 30
    ws1.merge_cells("B2:G2")
    c_title = ws1["B2"]
    c_title.value = f"RESUMO FINANCEIRO CONSOLIDADO — NFS-e ({toma_nome_top})"
    c_title.font = Font(name=font_family, size=14, bold=True, color="FFFFFF")
    c_title.fill = fill_title
    c_title.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws1.row_dimensions[3].height = 20
    ws1.merge_cells("B3:G3")
    ws1["B3"].value = "Consolidação Financeira, Tributária e de Retenções na Fonte (Padrão Nacional SPED NFS-e)"
    ws1["B3"].font = Font(name=font_family, size=10, italic=True, color="333333")

    ws1.row_dimensions[4].height = 20
    ws1.merge_cells("B4:G4")
    ws1["B4"].value = (
        f"Tomador Principal: {toma_nome_top} | CNPJ: {toma_cnpj_top} | "
        f"Município Principal de Execução: {mun_principal}"
    )
    ws1["B4"].font = Font(name=font_family, size=10, bold=True, color="1F4E79")

    # Seção 1: Indicadores Gerais
    ws1.row_dimensions[6].height = 24
    for col_let in ("B", "C", "D"):
        cell = ws1[f"{col_let}6"]
        cell.fill = fill_sec_hdr
        cell.font = font_sec_hdr
    ws1["B6"].value = "1. INDICADORES GERAIS DO LOTE DE NOTAS FISCAIS"

    kpis = [
        ("Total de Notas Fiscais Processadas", n_rows, "100% dos arquivos XML válidos do lote"),
        ("Total de Fornecedores / Prestadores Distintos", len(forn_groups), "Empresas prestadoras contratadas"),
        ("Município Principal de Execução", mun_principal, "Local predominante de prestação dos serviços"),
        ("Período / Exercício de Competência", periodo_comp, "Competência apurada nos documentos"),
        (
            "Notas com Registro de Obra / CNO",
            count_obras,
            f"{(count_obras / n_rows * 100):.1f}% do volume total de notas",
        ),
        (
            "Notas com Apuração de IBS / CBS (Novo SPED)",
            count_ibscbs,
            f"{(count_ibscbs / n_rows * 100):.1f}% do volume total de notas",
        ),
    ]
    for offset, (lbl, val, desc) in enumerate(kpis):
        r_idx = 7 + offset
        ws1.row_dimensions[r_idx].height = 20
        c_b = ws1[f"B{r_idx}"]
        c_c = ws1[f"C{r_idx}"]
        c_d = ws1[f"D{r_idx}"]

        c_b.value = lbl
        c_b.font = Font(name=font_family, size=10, bold=True, color="1F4E79")
        c_b.fill = fill_kpi_lbl
        c_b.border = border_kpi

        c_c.value = val
        c_c.font = Font(name=font_family, size=11, bold=True, color="000000")
        c_c.alignment = Alignment(horizontal="center", vertical="center")
        c_c.border = border_data

        c_d.value = desc
        c_d.font = Font(name=font_family, size=9, italic=True, color="595959")
        c_d.alignment = Alignment(horizontal="left", vertical="center")
        c_d.border = border_data

    # Seção 2: Demonstrativo Financeiro Consolidado (com fórmulas dinâmicas vinculadas à aba Notas Fiscais)
    ws1.row_dimensions[14].height = 24
    for col_let in ("B", "C", "D", "E"):
        ws1[f"{col_let}14"].fill = fill_sec_hdr
        ws1[f"{col_let}14"].font = font_sec_hdr
    ws1["B14"].value = "2. DEMONSTRATIVO FINANCEIRO CONSOLIDADO DOS SERVIÇOS E RETENÇÕES"

    ws1.row_dimensions[15].height = 24
    headers_s2 = [
        ("B", "Item / Descrição Financeira e Contábil"),
        ("C", "Valor Consolidado (R$)"),
        ("D", "% s/ Faturamento"),
        ("E", "Qtd. Notas c/ Incidência"),
    ]
    for col_let, text in headers_s2:
        c = ws1[f"{col_let}15"]
        c.value = text
        c.fill = fill_title
        c.font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center" if col_let != "B" else "left", vertical="center")
        c.border = border_data

    # Referências de colunas na aba 'Notas Fiscais'
    col_tp_ret = prop_to_col_letter["TipoRetencaoISS"]
    col_v_serv = prop_to_col_letter["ValorServicos"]
    col_v_ded = prop_to_col_letter["DeducoesReducoes"]
    col_v_desc = prop_to_col_letter["DescontoIncond"]
    col_v_bc = prop_to_col_letter["BaseCalculoISS"]
    col_v_iss = prop_to_col_letter["ValorISS"]
    col_v_inss = prop_to_col_letter["RetencaoINSS"]
    col_v_irrf = prop_to_col_letter["RetencaoIRRF"]
    col_v_csll = prop_to_col_letter["RetencaoCSLL"]
    col_v_pis = prop_to_col_letter["ValorPIS"]
    col_v_cof = prop_to_col_letter["ValorCOFINS"]
    col_v_tret = prop_to_col_letter["TotalRetencoes"]
    col_v_liq = prop_to_col_letter["ValorLiquido"]
    col_v_ibs = prop_to_col_letter["ValorIBSTotal"]
    col_v_cbs = prop_to_col_letter["ValorCBSTotal"]

    nf_range = lambda col: f"'Notas Fiscais'!${col}${data_start_row}:${col}${data_end_row}"

    def add_fin_formula_row(r_num: int, label: str, val_formula: str, qtd_formula: str, style_type: str = "normal"):
        ws1.row_dimensions[r_num].height = 20
        c_b = ws1[f"B{r_num}"]
        c_c = ws1[f"C{r_num}"]
        c_d = ws1[f"D{r_num}"]
        c_e = ws1[f"E{r_num}"]

        c_b.value = label
        c_c.value = val_formula
        c_c.number_format = num_fmt_currency
        c_d.value = f"=IF($C$16>0, C{r_num}/$C$16, 0)"
        c_d.number_format = num_fmt_pct
        c_e.value = qtd_formula
        c_e.number_format = num_fmt_int
        c_e.alignment = Alignment(horizontal="center", vertical="center")

        if style_type == "total":
            for cell in (c_b, c_c, c_d, c_e):
                cell.fill = fill_total
                cell.font = font_total
                cell.border = border_total
        elif style_type == "subhdr":
            for cell in (c_b, c_c, c_d, c_e):
                cell.fill = fill_subhdr
                cell.font = font_subhdr
                cell.border = border_data
        else:
            for cell in (c_b, c_c, c_d, c_e):
                cell.fill = fill_zebra_even
                cell.font = font_data
                cell.border = border_data

    add_fin_formula_row(
        16,
        "(+) VALOR TOTAL BRUTO DOS SERVIÇOS PRESTADOS",
        f"=SUM({nf_range(col_v_serv)})",
        f"=COUNTIF({nf_range(col_v_serv)}, \">0\")",
        "subhdr",
    )
    add_fin_formula_row(
        17,
        "(-) Deduções / Reduções da Base de Cálculo",
        f"=SUM({nf_range(col_v_ded)})",
        f"=COUNTIF({nf_range(col_v_ded)}, \">0\")",
        "normal",
    )
    add_fin_formula_row(
        18,
        "(-) Descontos Incondicionados Concedidos",
        f"=SUM({nf_range(col_v_desc)})",
        f"=COUNTIF({nf_range(col_v_desc)}, \">0\")",
        "normal",
    )
    add_fin_formula_row(
        19,
        "(=) BASE DE CÁLCULO EFETIVA DO ISSQN",
        f"=SUM({nf_range(col_v_bc)})",
        f"=COUNTIF({nf_range(col_v_bc)}, \">0\")",
        "subhdr",
    )

    # Subtítulo Tributação Municipal
    ws1.row_dimensions[20].height = 20
    for col_let in ("B", "C", "D", "E"):
        ws1[f"{col_let}20"].fill = fill_subhdr
        ws1[f"{col_let}20"].font = font_subhdr
        ws1[f"{col_let}20"].border = border_data
    ws1["B20"].value = "( - ) TRIBUTAÇÃO MUNICIPAL (ISS / ISSQN)"

    add_fin_formula_row(
        21,
        "    • ISSQN Retido na Fonte (a recolher pelo Tomador)",
        f"=SUMIF({nf_range(col_tp_ret)}, \"2 - Retido na Fonte\", {nf_range(col_v_iss)})",
        f"=COUNTIF({nf_range(col_tp_ret)}, \"2 - Retido na Fonte\")",
        "normal",
    )
    add_fin_formula_row(
        22,
        "    • ISSQN Não Retido (devido pelo Prestador)",
        f"=SUMIF({nf_range(col_tp_ret)}, \"1 - Não Retido\", {nf_range(col_v_iss)})",
        f"=COUNTIF({nf_range(col_tp_ret)}, \"1 - Não Retido\")",
        "normal",
    )
    add_fin_formula_row(
        23,
        "    • Total do ISSQN Apurado nas Notas",
        f"=SUM({nf_range(col_v_iss)})",
        f"=COUNTIF({nf_range(col_v_iss)}, \">0\")",
        "total",
    )

    # Subtítulo Retenções Federais
    ws1.row_dimensions[24].height = 20
    for col_let in ("B", "C", "D", "E"):
        ws1[f"{col_let}24"].fill = fill_subhdr
        ws1[f"{col_let}24"].font = font_subhdr
        ws1[f"{col_let}24"].border = border_data
    ws1["B24"].value = "( - ) RETENÇÕES FEDERAIS E PREVIDENCIÁRIAS"

    add_fin_formula_row(
        25,
        "    • Retenção Previdenciária (INSS / CP)",
        f"=SUM({nf_range(col_v_inss)})",
        f"=COUNTIF({nf_range(col_v_inss)}, \">0\")",
        "normal",
    )
    add_fin_formula_row(
        26,
        "    • Retenção de Imposto de Renda (IRRF)",
        f"=SUM({nf_range(col_v_irrf)})",
        f"=COUNTIF({nf_range(col_v_irrf)}, \">0\")",
        "normal",
    )
    add_fin_formula_row(
        27,
        "    • Retenção de Contribuição Social (CSLL)",
        f"=SUM({nf_range(col_v_csll)})",
        f"=COUNTIF({nf_range(col_v_csll)}, \">0\")",
        "normal",
    )
    add_fin_formula_row(
        28,
        "    • Contribuição para PIS / PASEP Retida",
        f"=SUM({nf_range(col_v_pis)})",
        f"=COUNTIF({nf_range(col_v_pis)}, \">0\")",
        "normal",
    )
    add_fin_formula_row(
        29,
        "    • Contribuição para COFINS Retida",
        f"=SUM({nf_range(col_v_cof)})",
        f"=COUNTIF({nf_range(col_v_cof)}, \">0\")",
        "normal",
    )
    add_fin_formula_row(
        30,
        "(=) TOTAL DAS RETENÇÕES TRIBUTÁRIAS NA FONTE",
        f"=SUM({nf_range(col_v_tret)})",
        f"=COUNTIF({nf_range(col_v_tret)}, \">0\")",
        "total",
    )
    add_fin_formula_row(
        31,
        "(=) VALOR TOTAL LÍQUIDO A PAGAR AOS FORNECEDORES",
        f"=SUM({nf_range(col_v_liq)})",
        f"=COUNTIF({nf_range(col_v_liq)}, \">0\")",
        "total",
    )

    # Seção 3: Reforma Tributária IBS / CBS
    ws1.row_dimensions[33].height = 24
    for col_let in ("B", "C", "D"):
        ws1[f"{col_let}33"].fill = fill_sec_hdr
        ws1[f"{col_let}33"].font = font_sec_hdr
    ws1["B33"].value = "3. APURAÇÃO DA REFORMA TRIBUTÁRIA (NOVO SPED: IBS / CBS - PROJETO PILOTO)"

    ws1.row_dimensions[34].height = 22
    for col_let, text in [("B", "Tributo / Esfera Governamental"), ("C", "Total Apurado (R$)"), ("D", "Alíquota / Referência")]:
        c = ws1[f"{col_let}34"]
        c.value = text
        c.fill = fill_title
        c.font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center" if col_let != "B" else "left", vertical="center")
        c.border = border_data

    ibs_rows = [
        (35, "Total IBS Apurado (Imposto sobre Bens e Serviços - Estadual/Municipal)", f"=SUM({nf_range(col_v_ibs)})", "0,10% Efetiva UF", False),
        (36, "Total CBS Apurado (Contribuição sobre Bens e Serviços - Federal)", f"=SUM({nf_range(col_v_cbs)})", "0,90% Efetiva Fed", False),
        (37, "Total Combinado IBS + CBS", "=C35+C36", "1,00% Total Piloto", True),
    ]
    for r_num, lbl, form, ref_txt, is_tot in ibs_rows:
        ws1.row_dimensions[r_num].height = 20
        ws1[f"B{r_num}"].value = lbl
        ws1[f"C{r_num}"].value = form
        ws1[f"C{r_num}"].number_format = num_fmt_currency
        ws1[f"D{r_num}"].value = ref_txt
        ws1[f"D{r_num}"].alignment = Alignment(horizontal="center", vertical="center")
        for col_let in ("B", "C", "D"):
            cell = ws1[f"{col_let}{r_num}"]
            cell.fill = fill_total if is_tot else fill_zebra_even
            cell.font = font_total if is_tot else font_data
            cell.border = border_total if is_tot else border_data

    # Seção 4: Top 10 Maiores Fornecedores
    ws1.row_dimensions[39].height = 24
    for col_let in ("B", "C", "D", "E", "F", "G"):
        ws1[f"{col_let}39"].fill = fill_sec_hdr
        ws1[f"{col_let}39"].font = font_sec_hdr
    ws1["B39"].value = "4. TOP 10 MAIORES FORNECEDORES CONTRATADOS (POR VOLUME FINANCEIRO)"

    ws1.row_dimensions[40].height = 24
    s4_headers = [
        ("B", "Razão Social do Fornecedor"),
        ("C", "CNPJ"),
        ("D", "Qtd. Notas"),
        ("E", "Total Serviços (R$)"),
        ("F", "Total Líquido (R$)"),
        ("G", "% do Total"),
    ]
    for col_let, text in s4_headers:
        c = ws1[f"{col_let}40"]
        c.value = text
        c.fill = fill_title
        c.font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center" if col_let != "B" else "left", vertical="center")
        c.border = border_data

    top10 = forn_groups[:10]
    others = forn_groups[10:]
    r_top = 41
    for fg in top10:
        ws1.row_dimensions[r_top].height = 20
        ws1[f"B{r_top}"].value = fg["RazaoSocial"]
        ws1[f"C{r_top}"].value = fg["CNPJ"]
        ws1[f"C{r_top}"].alignment = Alignment(horizontal="center", vertical="center")
        ws1[f"D{r_top}"].value = fg["Qtd"]
        ws1[f"D{r_top}"].alignment = Alignment(horizontal="center", vertical="center")
        ws1[f"E{r_top}"].value = fg["TotalServ"]
        ws1[f"E{r_top}"].number_format = num_fmt_currency
        ws1[f"F{r_top}"].value = fg["TotalLiq"]
        ws1[f"F{r_top}"].number_format = num_fmt_currency
        ws1[f"G{r_top}"].value = f"=IF($C$16>0, E{r_top}/$C$16, 0)"
        ws1[f"G{r_top}"].number_format = num_fmt_pct
        for col_let in ("B", "C", "D", "E", "F", "G"):
            ws1[f"{col_let}{r_top}"].font = font_data
            ws1[f"{col_let}{r_top}"].border = border_data
        r_top += 1

    if others:
        ws1.row_dimensions[r_top].height = 20
        ws1[f"B{r_top}"].value = f"DEMAIS FORNECEDORES ({len(others)} empresas)"
        ws1[f"C{r_top}"].value = "DIVERSOS"
        ws1[f"C{r_top}"].alignment = Alignment(horizontal="center", vertical="center")
        ws1[f"D{r_top}"].value = sum(x["Qtd"] for x in others)
        ws1[f"D{r_top}"].alignment = Alignment(horizontal="center", vertical="center")
        ws1[f"E{r_top}"].value = sum(x["TotalServ"] for x in others)
        ws1[f"E{r_top}"].number_format = num_fmt_currency
        ws1[f"F{r_top}"].value = sum(x["TotalLiq"] for x in others)
        ws1[f"F{r_top}"].number_format = num_fmt_currency
        ws1[f"G{r_top}"].value = f"=IF($C$16>0, E{r_top}/$C$16, 0)"
        ws1[f"G{r_top}"].number_format = num_fmt_pct
        for col_let in ("B", "C", "D", "E", "F", "G"):
            ws1[f"{col_let}{r_top}"].font = Font(name=font_family, size=10, italic=True, color="333333")
            ws1[f"{col_let}{r_top}"].border = border_data
        r_top += 1

    # Linha Total Top Fornecedores
    ws1.row_dimensions[r_top].height = 22
    ws1[f"B{r_top}"].value = "TOTAL GERAL FORNECEDORES"
    ws1[f"C{r_top}"].value = "-"
    ws1[f"C{r_top}"].alignment = Alignment(horizontal="center", vertical="center")
    ws1[f"D{r_top}"].value = f"=SUM(D41:D{r_top-1})"
    ws1[f"D{r_top}"].alignment = Alignment(horizontal="center", vertical="center")
    ws1[f"E{r_top}"].value = f"=SUM(E41:E{r_top-1})"
    ws1[f"E{r_top}"].number_format = num_fmt_currency
    ws1[f"F{r_top}"].value = f"=SUM(F41:F{r_top-1})"
    ws1[f"F{r_top}"].number_format = num_fmt_currency
    ws1[f"G{r_top}"].value = f"=SUM(G41:G{r_top-1})"
    ws1[f"G{r_top}"].number_format = num_fmt_pct
    for col_let in ("B", "C", "D", "E", "F", "G"):
        ws1[f"{col_let}{r_top}"].fill = fill_total
        ws1[f"{col_let}{r_top}"].font = font_total
        ws1[f"{col_let}{r_top}"].border = border_total

    # ==========================================================================
    # ABA 2: NOTAS FISCAIS (DETALHAMENTO FUNCIONAL + SUBTOTAL DINÂMICO)
    # ==========================================================================
    if progress_cb:
        progress_cb(2, 4, "Montando Aba 2: Notas Fiscais (Colunas Funcionais e SUBTOTAL)...")

    ws2 = wb.create_sheet(title="Notas Fiscais")
    ws2.views.sheetView[0].showGridLines = True
    # Congela a linha 1 (cabeçalho) e as 3 primeiras colunas (Número NFS-e, Competência, Razão Social)
    ws2.freeze_panes = "D2"
    ws2.sheet_properties.outlinePr.summaryRight = False

    # Configura cabeçalhos, larguras e agrupamento retrátil (outline_level)
    ws2.row_dimensions[1].height = 30
    for c_idx, c_def in enumerate(NOTAS_COL_DEFS, start=1):
        col_let = get_column_letter(c_idx)
        col_dim = ws2.column_dimensions[col_let]
        col_dim.width = c_def["width"]
        if c_def.get("outline", 0) > 0:
            col_dim.outline_level = c_def["outline"]

        hdr_cell = ws2.cell(row=1, column=c_idx, value=c_def["header"])
        hdr_cell.fill = PatternFill("solid", fgColor=c_def["group_color"])
        hdr_cell.font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
        hdr_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        hdr_cell.border = border_data

    # Preenche linhas de dados
    for row_offset, rec in enumerate(records):
        r_num = data_start_row + row_offset
        ws2.row_dimensions[r_num].height = 20
        row_fill = fill_zebra_even if row_offset % 2 == 0 else fill_zebra_odd

        for c_idx, c_def in enumerate(NOTAS_COL_DEFS, start=1):
            prop = c_def["prop"]
            val = rec.get(prop, "")
            cell = ws2.cell(row=r_num, column=c_idx, value=val)
            cell.border = border_data
            cell.fill = row_fill
            cell.font = font_data

            # Alinhamento e formato numérico
            c_type = c_def["type"]
            if c_type == "currency":
                cell.number_format = num_fmt_currency
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c_type == "rate":
                cell.number_format = num_fmt_rate
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal=c_def["align"], vertical="center")

            # Destaques visuais inteligentes para colunas-chave
            if prop == "TipoRetencaoISS":
                if str(val).startswith("2"):
                    cell.fill = fill_retido
                    cell.font = font_retido
                elif str(val).startswith("1"):
                    cell.fill = fill_nao_retido
                    cell.font = font_nao_retido
            elif prop == "RetencoesFederaisAplicadas" and val != "Sem Retenção Federal":
                cell.fill = fill_retido
                cell.font = font_retido
            elif prop == "StatusDuplicidade" and val != "Única":
                cell.fill = fill_dup_alert
                cell.font = font_dup_alert
            elif prop in ("ValorISS", "TotalRetencoes", "ValorLiquido") and isinstance(val, (int, float)) and val > 0:
                cell.font = font_data_bold

    # Linha de TOTAL GERAL DINÂMICO (usando SUBTOTAL(109, ...) para somar APENAS as linhas visíveis no filtro!)
    ws2.row_dimensions[total_row_idx].height = 25
    for c_idx, c_def in enumerate(NOTAS_COL_DEFS, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws2.cell(row=total_row_idx, column=c_idx)
        cell.fill = fill_total
        cell.font = font_total
        cell.border = border_total

        if c_idx == 1:
            cell.value = "TOTAL (FILTRO ATIVO)"
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif c_idx == 2:
            # Contagem dinâmica de notas visíveis no filtro
            cell.value = f"=SUBTOTAL(103, A{data_start_row}:A{data_end_row})"
            cell.number_format = '0" notas"'
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif c_def.get("is_sum"):
            # SUBTOTAL(109, ...) soma exclusivamente as linhas visíveis (ignora linhas ocultadas pelo filtro)
            cell.value = f"=SUBTOTAL(109, {col_let}{data_start_row}:{col_let}{data_end_row})"
            cell.number_format = num_fmt_currency
            cell.alignment = Alignment(horizontal="right", vertical="center")
        else:
            cell.value = ""

    last_col_let_ws2 = get_column_letter(len(NOTAS_COL_DEFS))
    ws2.auto_filter.ref = f"A1:{last_col_let_ws2}{data_end_row}"

    # ==========================================================================
    # ABA 3: RESUMO POR FORNECEDOR (COM QUEBRA DE ISS RETIDO X NÃO RETIDO + SUBTOTAL)
    # ==========================================================================
    if progress_cb:
        progress_cb(3, 4, "Montando Aba 3: Resumo por Fornecedor...")

    ws3 = wb.create_sheet(title="Resumo por Fornecedor")
    ws3.views.sheetView[0].showGridLines = True
    ws3.freeze_panes = "C2"

    s3_cols = [
        ("CNPJ Fornecedor", "CNPJ", 20, "center", "text", False),
        ("Razão Social", "RazaoSocial", 40, "left", "text", False),
        ("Regime Tributário", "RegimeSimples", 22, "left", "text", False),
        ("UF", "UF", 8, "center", "text", False),
        ("Município", "Municipio", 22, "left", "text", False),
        ("Qtd. Notas", "Qtd", 13, "center", "int", True),
        ("Total Serviços (R$)", "TotalServ", 20, "right", "currency", True),
        ("ISSQN Retido Fonte (R$)", "ISSRetido", 21, "right", "currency", True),
        ("ISSQN Não Retido (R$)", "ISSNaoRetido", 20, "right", "currency", True),
        ("Total ISSQN (R$)", "TotalISS", 18, "right", "currency", True),
        ("Total INSS (R$)", "TotalINSS", 18, "right", "currency", True),
        ("Total IRRF (R$)", "TotalIRRF", 18, "right", "currency", True),
        ("Total CSLL (R$)", "TotalCSLL", 18, "right", "currency", True),
        ("Total PIS (R$)", "TotalPIS", 18, "right", "currency", True),
        ("Total COFINS (R$)", "TotalCOF", 18, "right", "currency", True),
        ("Total Retenções (R$)", "TotalRet", 20, "right", "currency", True),
        ("Total Líquido (R$)", "TotalLiq", 21, "right", "currency", True),
    ]

    ws3.row_dimensions[1].height = 28
    for c_idx, (hdr, _, width, _, _, _) in enumerate(s3_cols, start=1):
        col_let = get_column_letter(c_idx)
        ws3.column_dimensions[col_let].width = width
        cell = ws3.cell(row=1, column=c_idx, value=hdr)
        cell.fill = fill_title
        cell.font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_data

    for idx_f, fg in enumerate(forn_groups):
        r_num = 2 + idx_f
        ws3.row_dimensions[r_num].height = 20
        r_fill = fill_zebra_even if idx_f % 2 == 0 else fill_zebra_odd
        for c_idx, (_, prop, _, align, c_type, _) in enumerate(s3_cols, start=1):
            cell = ws3.cell(row=r_num, column=c_idx, value=fg[prop])
            cell.fill = r_fill
            cell.font = font_data
            cell.border = border_data
            cell.alignment = Alignment(horizontal=align, vertical="center")
            if c_type == "currency":
                cell.number_format = num_fmt_currency
            elif c_type == "int":
                cell.number_format = num_fmt_int

    s3_end_row = 1 + len(forn_groups)
    s3_tot_row = s3_end_row + 1
    ws3.row_dimensions[s3_tot_row].height = 24
    for c_idx, (_, _, _, align, c_type, is_sum) in enumerate(s3_cols, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws3.cell(row=s3_tot_row, column=c_idx)
        cell.fill = fill_total
        cell.font = font_total
        cell.border = border_total
        if c_idx == 1:
            cell.value = "TOTAL (FILTRADO)"
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif c_idx == 2:
            cell.value = f"=SUBTOTAL(103, A2:A{s3_end_row})"
            cell.number_format = '0" fornecedores"'
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif is_sum:
            cell.value = f"=SUBTOTAL(109, {col_let}2:{col_let}{s3_end_row})"
            cell.number_format = num_fmt_currency if c_type == "currency" else num_fmt_int
            cell.alignment = Alignment(horizontal=align, vertical="center")

    ws3.auto_filter.ref = f"A1:{get_column_letter(len(s3_cols))}{s3_end_row}"

    # ==========================================================================
    # ABA 4: RESUMO POR MUNICÍPIO (COM QUEBRA DE ISS RETIDO X NÃO RETIDO + SUBTOTAL)
    # ==========================================================================
    if progress_cb:
        progress_cb(4, 4, "Montando Aba 4: Resumo por Município...")

    ws4 = wb.create_sheet(title="Resumo por Município")
    ws4.views.sheetView[0].showGridLines = True
    ws4.freeze_panes = "A2"

    s4_cols = [
        ("Município de Prestação", "Municipio", 30, "left", "text", False),
        ("Qtd. Notas", "Qtd", 14, "center", "int", True),
        ("Total Serviços (R$)", "TotalServ", 22, "right", "currency", True),
        ("ISSQN Retido Fonte (R$)", "ISSRetido", 22, "right", "currency", True),
        ("ISSQN Não Retido (R$)", "ISSNaoRetido", 22, "right", "currency", True),
        ("Total ISSQN (R$)", "TotalISS", 20, "right", "currency", True),
        ("Retenções Federais (R$)", "RetFederais", 22, "right", "currency", True),
        ("Total Retenções (R$)", "TotalRet", 22, "right", "currency", True),
        ("Total Líquido (R$)", "TotalLiq", 22, "right", "currency", True),
    ]

    ws4.row_dimensions[1].height = 28
    for c_idx, (hdr, _, width, _, _, _) in enumerate(s4_cols, start=1):
        col_let = get_column_letter(c_idx)
        ws4.column_dimensions[col_let].width = width
        cell = ws4.cell(row=1, column=c_idx, value=hdr)
        cell.fill = fill_title
        cell.font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_data

    for idx_m, mg in enumerate(mun_groups):
        r_num = 2 + idx_m
        ws4.row_dimensions[r_num].height = 20
        r_fill = fill_zebra_even if idx_m % 2 == 0 else fill_zebra_odd
        for c_idx, (_, prop, _, align, c_type, _) in enumerate(s4_cols, start=1):
            cell = ws4.cell(row=r_num, column=c_idx, value=mg[prop])
            cell.fill = r_fill
            cell.font = font_data
            cell.border = border_data
            cell.alignment = Alignment(horizontal=align, vertical="center")
            if c_type == "currency":
                cell.number_format = num_fmt_currency
            elif c_type == "int":
                cell.number_format = num_fmt_int

    s4_end_row = 1 + len(mun_groups)
    s4_tot_row = s4_end_row + 1
    ws4.row_dimensions[s4_tot_row].height = 24
    for c_idx, (_, _, _, align, c_type, is_sum) in enumerate(s4_cols, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws4.cell(row=s4_tot_row, column=c_idx)
        cell.fill = fill_total
        cell.font = font_total
        cell.border = border_total
        if c_idx == 1:
            cell.value = "TOTAL (FILTRADO)"
            cell.alignment = Alignment(horizontal="left", vertical="center")
        elif is_sum:
            cell.value = f"=SUBTOTAL(109, {col_let}2:{col_let}{s4_end_row})"
            cell.number_format = num_fmt_currency if c_type == "currency" else num_fmt_int
            cell.alignment = Alignment(horizontal=align, vertical="center")

    ws4.auto_filter.ref = f"A1:{get_column_letter(len(s4_cols))}{s4_end_row}"

    actual_xlsx = Path(output_xlsx)
    locked_fallback = False
    try:
        wb.save(str(actual_xlsx))
    except PermissionError:
        locked_fallback = True
        for attempt in range(1, 20):
            suffix_tag = "_NOVO" if attempt == 1 else f"_NOVO_{attempt}"
            candidate = actual_xlsx.with_name(f"{actual_xlsx.stem}{suffix_tag}{actual_xlsx.suffix}")
            try:
                wb.save(str(candidate))
                actual_xlsx = candidate
                break
            except PermissionError:
                continue
        else:
            raise

    return {
        "total_notas": n_rows,
        "total_fornecedores": len(forn_groups),
        "total_municipios": len(mun_groups),
        "total_servicos": total_servicos,
        "total_iss": total_iss,
        "total_iss_retido": total_iss_retido,
        "total_iss_nao_retido": total_iss_nao_retido,
        "total_retencoes": total_retencoes,
        "total_liquido": total_liquido,
        "output_xlsx": str(actual_xlsx.resolve()),
        "locked_fallback": locked_fallback,
    }


def compile_nfse_folder(input_folder: str, output_xlsx: str, generate_csv: bool = False, progress_cb=None) -> dict:
    """Pipeline completo: lê os XMLs da pasta, gera o XLSX dinâmico e opcionalmente o CSV."""
    records = load_all_xml_records(input_folder, progress_cb=progress_cb)
    if not records:
        raise RuntimeError("Nenhum arquivo XML válido de NFS-e foi encontrado na pasta selecionada.")

    summary = build_excel_workbook(records, output_xlsx, progress_cb=progress_cb)

    if generate_csv:
        csv_path = Path(summary["output_xlsx"]).with_suffix(".csv")
        try:
            export_to_csv(records, str(csv_path))
            summary["output_csv"] = str(csv_path.resolve())
        except PermissionError:
            alt_csv = csv_path.with_name(f"{csv_path.stem}_NOVO.csv")
            export_to_csv(records, str(alt_csv))
            summary["output_csv"] = str(alt_csv.resolve())

    return summary


# ==============================================================================
# INTERFACE GRÁFICA (GUI TKINTER) COM SUGESTÃO INTELIGENTE DE NOME
# ==============================================================================

def fmt_brl(val: float) -> str:
    return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def launch_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("Compilador de NFS-e (XML -> Excel Dinâmico) - SEFIN")

    win_w, win_h = 780, 520
    scr_w = root.winfo_screenwidth()
    scr_h = root.winfo_screenheight()
    pos_x = max(0, (scr_w - win_w) // 2)
    pos_y = max(0, (scr_h - win_h) // 2 - 20)
    root.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")
    root.minsize(740, 500)
    root.configure(bg="#EEF2F6")

    style = ttk.Style()
    if "vista" in style.theme_names():
        style.theme_use("vista")
    elif "clam" in style.theme_names():
        style.theme_use("clam")

    # Estado reativo da interface (compatível tanto com .py quanto com .exe PyInstaller)
    if getattr(sys, "frozen", False):
        base_dir = Path(sys.executable).resolve().parent
    else:
        base_dir = Path(__file__).resolve().parent
    default_xml_dir = base_dir / "XML NOTA OMNIA DATACENTER"

    var_input_dir = tk.StringVar(value=str(default_xml_dir if default_xml_dir.exists() else base_dir))
    var_output_dir = tk.StringVar(value=str(base_dir))
    var_filename = tk.StringVar(value="NOTAS_FISCAIS_OMNIA_DATACENTER.xlsx")
    var_gen_csv = tk.BooleanVar(value=False)
    var_status = tk.StringVar(value="Pronto para compilar. Verifique a pasta e clique no botão verde abaixo.")

    # Variáveis dos Mini-Cards de Pré-Leitura
    var_card_notas = tk.StringVar(value="-")
    var_card_empresa = tk.StringVar(value="-")
    var_card_empresa_lbl = tk.StringVar(value="EMPRESA PRINCIPAL")
    var_card_periodo = tk.StringVar(value="-")
    var_card_valor = tk.StringVar(value="-")

    current_scan = {}

    # Cabeçalho Visual
    header_frame = tk.Frame(root, bg="#1F4E79", padx=22, pady=14)
    header_frame.pack(fill="x")
    tk.Label(
        header_frame,
        text="Compilador de Notas Fiscais Eletrônicas (NFS-e Nacional) - Feito por Icaro",
        font=("Segoe UI", 14, "bold"),
        fg="#FFFFFF",
        bg="#1F4E79",
    ).pack(anchor="w")
    tk.Label(
        header_frame,
        text="Converte lotes de XML do nfse.gov.br em planilha Excel (.xlsx) organizada e com somas dinâmicas por filtro",
        font=("Segoe UI", 9),
        fg="#D9E1F2",
        bg="#1F4E79",
    ).pack(anchor="w", pady=(2, 0))

    main_frame = tk.Frame(root, bg="#EEF2F6", padx=20, pady=14)
    main_frame.pack(fill="both", expand=True)

    # ==========================================================================
    # PASSO 1: SELEÇÃO DA PASTA DE XMLs OU ARQUIVO CSV ADN + CARDS DE DIAGNÓSTICO
    # ==========================================================================
    step1 = tk.LabelFrame(
        main_frame,
        text=" 1. Origem das Notas Fiscais (Pasta de XMLs ou Arquivo .CSV) ",
        font=("Segoe UI", 10, "bold"),
        fg="#1F4E79",
        bg="#FFFFFF",
        bd=1,
        relief="solid",
        padx=14,
        pady=10,
    )
    step1.pack(fill="x", pady=(0, 12))

    row1 = tk.Frame(step1, bg="#FFFFFF")
    row1.pack(fill="x")
    entry_input = ttk.Entry(row1, textvariable=var_input_dir, font=("Segoe UI", 10))
    entry_input.pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=2)

    # Grid de 4 Mini-Cards para mostrar a pré-leitura sem cortar texto
    cards_frame = tk.Frame(step1, bg="#FFFFFF")
    cards_frame.pack(fill="x", pady=(10, 2))
    for i, weight in enumerate((1, 2, 1, 1)):
        cards_frame.columnconfigure(i, weight=weight)

    def make_info_card(parent, col, title_var_or_str, val_var, bg_color="#F4F7FA", accent="#1F4E79"):
        card = tk.Frame(parent, bg=bg_color, highlightbackground="#D9E1F2", highlightthickness=1, padx=10, pady=6)
        card.grid(row=0, column=col, sticky="nsew", padx=(0 if col == 0 else 6, 0))
        if isinstance(title_var_or_str, tk.StringVar):
            lbl_t = tk.Label(card, textvariable=title_var_or_str, font=("Segoe UI", 7, "bold"), fg="#595959", bg=bg_color, anchor="w")
        else:
            lbl_t = tk.Label(card, text=title_var_or_str, font=("Segoe UI", 7, "bold"), fg="#595959", bg=bg_color, anchor="w")
        lbl_t.pack(fill="x")
        lbl_v = tk.Label(card, textvariable=val_var, font=("Segoe UI", 9, "bold"), fg=accent, bg=bg_color, anchor="w")
        lbl_v.pack(fill="x", pady=(1, 0))
        return card, lbl_v

    make_info_card(cards_frame, 0, "NOTAS VÁLIDAS", var_card_notas, "#F2F7F5", "#1B6B53")
    make_info_card(cards_frame, 1, var_card_empresa_lbl, var_card_empresa, "#F4F7FA", "#1F4E79")
    make_info_card(cards_frame, 2, "FORNECEDORES / PERÍODO", var_card_periodo, "#F4F7FA", "#2F5597")
    make_info_card(cards_frame, 3, "TOTAL BRUTO SERVIÇOS", var_card_valor, "#F2F7F5", "#1B6B53")

    def on_scan_folder(source_str: str):
        nonlocal current_scan
        info = scan_folder_quick(source_str)
        current_scan = info
        if not info.get("valid"):
            var_card_notas.set("0 notas válidas")
            var_card_empresa_lbl.set("STATUS DA ORIGEM")
            var_card_empresa.set(info.get("error", "Origem inválida"))
            var_card_periodo.set("-")
            var_card_valor.set("R$ 0,00")
            var_status.set(f"Atenção: {info.get('error', 'Selecione uma pasta ou CSV válido.')}")
            return

        origem_tag = info.get("origem_tipo", "XML")
        var_card_notas.set(f"✔ {info['count']} ({origem_tag})")
        var_card_empresa_lbl.set(f"{info['papel_empresa'].upper()} PRINCIPAL IDENTIFICADO")
        emp_txt = info["empresa_principal"]
        if len(emp_txt) > 34:
            emp_txt = emp_txt[:32] + "..."
        var_card_empresa.set(emp_txt)
        var_card_periodo.set(f"{info['qtd_fornecedores']} forn. | {info['periodo_label']}")
        var_card_valor.set(fmt_brl(info["total_servicos"]))

        var_filename.set(info["nome_sugerido"])
        p_src = Path(source_str).resolve()
        parent_dir = str(p_src.parent)
        var_output_dir.set(parent_dir)
        var_status.set(f"Origem validada ({info['count']} notas via {origem_tag}).")
        update_open_btn_state()

    def browse_input_folder():
        cur = var_input_dir.get().strip()
        initial = cur if os.path.isdir(cur) else (str(Path(cur).parent) if os.path.exists(cur) else str(base_dir))
        chosen = filedialog.askdirectory(title="Selecione a pasta contendo os XMLs ou CSV do ADN", initialdir=initial)
        if chosen:
            var_input_dir.set(chosen)
            on_scan_folder(chosen)

    def browse_input_file():
        cur = var_input_dir.get().strip()
        initial = cur if os.path.isdir(cur) else (str(Path(cur).parent) if os.path.exists(cur) else str(base_dir))
        chosen = filedialog.askopenfilename(
            title="Selecione o arquivo CSV (ADN) ou XML de Notas Fiscais",
            initialdir=initial,
            filetypes=[
                ("Arquivos Fiscais (CSV ADN ou XML)", "*.csv *.xml"),
                ("Arquivo CSV do ADN (*.csv)", "*.csv"),
                ("Arquivo XML NFS-e (*.xml)", "*.xml"),
                ("Todos os arquivos (*.*)", "*.*"),
            ],
        )
        if chosen:
            var_input_dir.set(chosen)
            on_scan_folder(chosen)

    entry_input.bind("<Return>", lambda _: on_scan_folder(var_input_dir.get().strip()))
    entry_input.bind("<FocusOut>", lambda _: on_scan_folder(var_input_dir.get().strip()))

    btn_browse_file = ttk.Button(row1, text="Selecionar CSV ou XML...", command=browse_input_file)
    btn_browse_file.pack(side="right")

    btn_browse_in = ttk.Button(row1, text="Procurar Pasta...", command=browse_input_folder)
    btn_browse_in.pack(side="right", padx=(0, 6))

    # ==========================================================================
    # PASSO 2: NOME SUGERIDO E PASTA DE SAÍDA
    # ==========================================================================
    step2 = tk.LabelFrame(
        main_frame,
        text=" 2. Arquivo Excel de Saída (Sugerido Automaticamente pelas Notas) ",
        font=("Segoe UI", 10, "bold"),
        fg="#1F4E79",
        bg="#FFFFFF",
        bd=1,
        relief="solid",
        padx=14,
        pady=10,
    )
    step2.pack(fill="x", pady=(0, 14))

    tk.Label(step2, text="Nome do Arquivo (.xlsx):", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", fg="#333333").pack(anchor="w")

    row_name = tk.Frame(step2, bg="#FFFFFF")
    row_name.pack(fill="x", pady=(2, 8))
    entry_name = ttk.Entry(row_name, textvariable=var_filename, font=("Segoe UI", 10, "bold"))
    entry_name.pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=2)

    def use_empresa_name():
        if current_scan.get("sugestao_empresa"):
            var_filename.set(current_scan["sugestao_empresa"])
            update_open_btn_state()

    def use_folder_name():
        if current_scan.get("sugestao_pasta"):
            var_filename.set(current_scan["sugestao_pasta"])
            update_open_btn_state()

    ttk.Button(row_name, text="Usar Nome da Empresa", command=use_empresa_name).pack(side="left", padx=(0, 4))
    ttk.Button(row_name, text="Usar Nome da Pasta", command=use_folder_name).pack(side="left")

    tk.Label(step2, text="Salvar na Pasta:", font=("Segoe UI", 9), bg="#FFFFFF", fg="#333333").pack(anchor="w")
    row_out = tk.Frame(step2, bg="#FFFFFF")
    row_out.pack(fill="x", pady=(2, 6))
    entry_out = ttk.Entry(row_out, textvariable=var_output_dir, font=("Segoe UI", 9))
    entry_out.pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=2)

    def browse_output():
        initial = var_output_dir.get() if os.path.isdir(var_output_dir.get()) else str(base_dir)
        chosen = filedialog.askdirectory(title="Selecione onde salvar o arquivo Excel", initialdir=initial)
        if chosen:
            var_output_dir.set(chosen)
            update_open_btn_state()

    ttk.Button(row_out, text="Alterar Destino...", command=browse_output).pack(side="left")

    chk_csv = ttk.Checkbutton(
        step2,
        text="Gerar também arquivo complementar .CSV (UTF-8 compatível com Excel)",
        variable=var_gen_csv,
    )
    chk_csv.pack(anchor="w", pady=(4, 0))

    # ==========================================================================
    # PASSO 3: AÇÕES E PAINEL DE STATUS (SEM CAIXA DE DIGITAÇÃO)
    # ==========================================================================
    step3 = tk.Frame(main_frame, bg="#EEF2F6")
    step3.pack(fill="x")

    btn_frame = tk.Frame(step3, bg="#EEF2F6")
    btn_frame.pack(fill="x", pady=(0, 10))

    last_generated_xlsx = {"path": None}

    def get_target_xlsx_path() -> str:
        if last_generated_xlsx.get("path") and os.path.exists(last_generated_xlsx["path"]):
            return last_generated_xlsx["path"]
        fname = var_filename.get().strip()
        if not fname.lower().endswith(".xlsx"):
            fname += ".xlsx"
        candidate = str(Path(var_output_dir.get().strip()) / fname)
        return candidate if os.path.exists(candidate) else ""

    def set_btn_open_enabled(enabled: bool):
        if enabled:
            btn_open_excel.configure(state="normal", bg="#2F5597", fg="#FFFFFF", cursor="hand2")
        else:
            btn_open_excel.configure(state="disabled", bg="#CBD5E1", fg="#64748B", cursor="arrow")

    def update_open_btn_state(*_):
        set_btn_open_enabled(bool(get_target_xlsx_path()))

    def open_generated_xlsx():
        p = get_target_xlsx_path()
        if p and os.path.exists(p):
            os.startfile(p)
        else:
            messagebox.showinfo("Abrir Planilha", "Gere a planilha primeiro para abri-la.")

    def open_output_folder():
        out_d = var_output_dir.get().strip()
        if out_d and os.path.isdir(out_d):
            os.startfile(out_d)

    btn_compile = tk.Button(
        btn_frame,
        text="▶  COMPILAR NOTAS E GERAR EXCEL",
        font=("Segoe UI", 11, "bold"),
        bg="#1B6B53",
        fg="#FFFFFF",
        activebackground="#14523F",
        activeforeground="#FFFFFF",
        relief="flat",
        bd=0,
        padx=22,
        pady=10,
        cursor="hand2",
    )
    btn_compile.pack(side="left", padx=(0, 10))

    btn_open_excel = tk.Button(
        btn_frame,
        text="📊  Abrir Planilha no Excel",
        font=("Segoe UI", 10, "bold"),
        bg="#CBD5E1",
        fg="#64748B",
        activebackground="#1F4E79",
        activeforeground="#FFFFFF",
        disabledforeground="#64748B",
        relief="flat",
        bd=0,
        padx=16,
        pady=10,
        state="disabled",
        command=open_generated_xlsx,
    )
    btn_open_excel.pack(side="left", padx=(0, 8))

    btn_open_folder = tk.Button(
        btn_frame,
        text="📁  Abrir Pasta",
        font=("Segoe UI", 10),
        bg="#E2E8F0",
        fg="#1E293B",
        activebackground="#CBD5E1",
        activeforeground="#0F172A",
        relief="flat",
        bd=0,
        padx=14,
        pady=10,
        cursor="hand2",
        command=open_output_folder,
    )
    btn_open_folder.pack(side="left")

    # Status direto no fundo (sem retângulo branco embaixo dos botões)
    lbl_status = tk.Label(
        step3,
        textvariable=var_status,
        font=("Segoe UI", 9, "bold"),
        fg="#1F4E79",
        bg="#EEF2F6",
        anchor="w",
    )
    lbl_status.pack(fill="x", pady=(2, 4))

    progress = ttk.Progressbar(step3, orient="horizontal", mode="determinate")

    def run_compilation():
        in_dir = var_input_dir.get().strip()
        out_dir = var_output_dir.get().strip()
        fname = var_filename.get().strip()

        if not in_dir or not os.path.exists(in_dir):
            messagebox.showerror("Erro", "Selecione uma pasta ou arquivo CSV/XML válido.")
            return
        if not fname:
            fname = "NOTAS_FISCAIS_COMPILADAS.xlsx"
        if not fname.lower().endswith(".xlsx"):
            fname += ".xlsx"
            var_filename.set(fname)

        out_path = str(Path(out_dir) / fname)
        btn_compile.configure(state="disabled", bg="#94A3B8")
        set_btn_open_enabled(False)
        progress["value"] = 0
        progress.pack(fill="x", pady=(2, 0))
        lbl_status.configure(fg="#1F4E79")

        def progress_callback(curr, tot, text):
            pct = int((curr / max(tot, 1)) * 100)
            root.after(0, lambda: progress.configure(value=pct))
            root.after(0, lambda: var_status.set(text))

        def worker():
            try:
                res = compile_nfse_folder(
                    input_folder=in_dir,
                    output_xlsx=out_path,
                    generate_csv=var_gen_csv.get(),
                    progress_cb=progress_callback,
                )

                def on_done():
                    progress.pack_forget()
                    last_generated_xlsx["path"] = res["output_xlsx"]
                    btn_compile.configure(state="normal", bg="#1B6B53")
                    set_btn_open_enabled(True)

                    saved_name = Path(res["output_xlsx"]).name
                    if res.get("locked_fallback"):
                        var_status.set(
                            f"✔ Concluído! Original aberto no Excel; salvo como: {saved_name}  |  "
                            f"Líquido: {fmt_brl(res['total_liquido'])}"
                        )
                    else:
                        var_status.set(
                            f"✔ Planilha gerada: {saved_name}  |  "
                            f"{res['total_notas']} notas  |  "
                            f"ISS Retido: {fmt_brl(res['total_iss_retido'])}  |  "
                            f"Líquido: {fmt_brl(res['total_liquido'])}"
                        )
                    lbl_status.configure(fg="#1B6B53")

                root.after(0, on_done)
            except PermissionError:
                def on_perm_err():
                    progress.pack_forget()
                    btn_compile.configure(state="normal", bg="#1B6B53")
                    lbl_status.configure(fg="#C00000")
                    var_status.set("❌ Erro: O arquivo Excel está aberto em outro programa.")
                    messagebox.showerror(
                        "Arquivo em Uso",
                        f"Não foi possível salvar '{out_path}' porque ele está aberto no Excel.\n\n"
                        "Feche a planilha no Excel ou escolha outro nome e tente novamente.",
                    )
                root.after(0, on_perm_err)
            except Exception as e:
                def on_err():
                    progress.pack_forget()
                    btn_compile.configure(state="normal", bg="#1B6B53")
                    lbl_status.configure(fg="#C00000")
                    var_status.set(f"❌ Erro ao compilar: {e}")
                    messagebox.showerror("Erro durante a compilação", str(e))
                root.after(0, on_err)

        threading.Thread(target=worker, daemon=True).start()

    btn_compile.configure(command=run_compilation)

    # Faz o scan automático inicial se a pasta ou arquivo padrão já existir
    if os.path.exists(var_input_dir.get()):
        on_scan_folder(var_input_dir.get())

    # Ajusta automaticamente a altura da janela ao conteúdo para não sobrar nem cortar nada
    root.update_idletasks()
    req_h = max(470, root.winfo_reqheight() + 10)
    root.geometry(f"{win_w}x{req_h}+{pos_x}+{pos_y}")
    root.minsize(740, req_h)

    root.mainloop()


# ==============================================================================
# PONTO DE ENTRADA (CLI OU GUI)
# ==============================================================================

def main():
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")

    parser = argparse.ArgumentParser(description="Compilador de NFS-e (XML -> XLSX Dinâmico)")
    parser.add_argument("--cli", action="store_true", help="Executa em modo linha de comando (sem abrir janela GUI)")
    parser.add_argument("--input", "-i", type=str, default="XML NOTA OMNIA DATACENTER", help="Pasta de entrada com os XMLs")
    parser.add_argument("--output", "-o", type=str, default="", help="Arquivo .xlsx de saída (opcional: sugere automaticamente se omitido)")
    parser.add_argument("--csv", action="store_true", help="Gera também o arquivo CSV complementar")
    args = parser.parse_args()

    if args.cli:
        in_folder = Path(args.input)
        if not in_folder.exists():
            print(f"[ERRO] Pasta não encontrada: {in_folder}", file=sys.stderr)
            sys.exit(1)

        scan = scan_folder_quick(str(in_folder))
        if not scan.get("valid"):
            print(f"[ERRO] {scan.get('error')}", file=sys.stderr)
            sys.exit(1)

        out_file = args.output.strip() if args.output else scan["nome_sugerido"]
        print(f"Pasta analisada: {scan['count']} XMLs | Empresa: {scan['empresa_principal']} | Nome sugerido: {out_file}")

        res = compile_nfse_folder(
            input_folder=str(in_folder),
            output_xlsx=out_file,
            generate_csv=args.csv,
            progress_cb=lambda c, t, msg: print(f"[{c}/{t}] {msg}"),
        )
        print(f"Concluído com sucesso! Arquivo salvo em: {res['output_xlsx']}")
    else:
        launch_gui()


if __name__ == "__main__":
    main()
