"""
Atualizar Dashboard DTO — script único
=======================================

O que faz, em ordem:
  1. Obtém a planilha exportada do Microsoft Forms (.xlsx)
     - No GitHub Actions (nuvem): baixa via link do SharePoint (variável de
       ambiente DTO_SHAREPOINT_URL), sem precisar de login.
     - No seu PC: usa o arquivo local sincronizado do OneDrive
       (DEFAULT_XLSX_PATH), como sempre.
  2. Obtém a base de crachás do RH (Visão.xlsx) da mesma forma (SharePoint
     no Actions, local no PC) e monta um dicionário crachá -> nome oficial.
  3. Extrai e limpa os dados das colunas B, F, G, J, L, M, O, P, R, Q:AK,
     AL, AM, AN, AO
     - Nome do avaliador: corrigido pelo crachá (coluna G) usando a base
       do RH; se o crachá não bater com ninguém, usa o nome digitado no
       Forms (coluna F)
     - Nome do colaborador avaliado: mesma lógica, usando o crachá da
       coluna AO (nome digitado como reserva na coluna AN)
     - Ambos sempre em CAIXA ALTA
  4. Substitui os dados dentro do próprio index.html (a linha
     "const dtoRaw = [...]")
  5. Atualiza a data em "Atualizado em ..."
  6. Faz commit e push para o GitHub (o GitHub Pages publica sozinho)

USO LOCAL (sem mudar nada):
  python atualizar_DTO.py
  python atualizar_DTO.py "caminho\\para\\export_do_forms.xlsx"

USO NO GITHUB ACTIONS:
  Defina os secrets do repositório:
    DTO_SHAREPOINT_URL    -> link "Alguém com o link" da planilha do Forms
    CRACHA_SHAREPOINT_URL -> link "Alguém com o link" do Visão.xlsx
  O workflow passa essas duas variáveis de ambiente e roda:
    python atualizar_DTO.py
  (sem argumento — o script detecta sozinho que deve baixar do SharePoint)

Pré-requisito: rode este script de dentro da pasta do repositório git clonado
de https://github.com/raulgribeiro/App-DTO (remoto 'origin' apontando pra
lá, branch 'main').
"""

import os
import re
import sys
import glob
import json
import subprocess
from pathlib import Path
from datetime import datetime

import openpyxl
import requests

# ------------------ CONFIGURAÇÃO ------------------
REPO_DIR = Path(__file__).parent.resolve()
INDEX_FILE = REPO_DIR / "index.html"
GIT_REMOTE_NAME = "origin"
GIT_BRANCH = "main"
COMMIT_MSG_PREFIX = "Atualizacao automatica DTO"

# Caminho local (OneDrive sincronizado no seu PC) — usado quando o script
# roda na sua máquina e nenhum link de SharePoint está configurado.
DEFAULT_XLSX_PATH = r"C:\Users\raulribeiro\OneDrive - CLEALCO AÇÚCAR E ÁLCOOL S.A\DTO Excel\Teste\Diagnóstico de Trabalho Operacional (DTO) SF 25_26.xlsx"
RH_XLSX_PATH = r"C:\Users\raulribeiro\OneDrive - CLEALCO AÇÚCAR E ÁLCOOL S.A\DTO Excel\Teste\Visão.xlsx"

# Links de compartilhamento "Alguém com o link" do SharePoint — usados
# automaticamente quando essas variáveis de ambiente existem (ex: rodando
# no GitHub Actions, onde não há OneDrive sincronizado).
DTO_SHAREPOINT_URL = os.environ.get("DTO_SHAREPOINT_URL")
CRACHA_SHAREPOINT_URL = os.environ.get("CRACHA_SHAREPOINT_URL")
# ---------------------------------------------------

# Colunas com o nome da(s) IT(s) treinada(s), dependendo da área escolhida no Forms
IT_COLS = ['S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC',
           'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AK']

USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")


def baixar_sharepoint(url, destino):
    """Baixa um arquivo de um link de compartilhamento 'Alguém com o link'
    do SharePoint, sem autenticação nenhuma.

    Precisa de 2 requisições com a MESMA sessão (cookies):
      1a: no link puro -> o SharePoint libera uma sessão anônima (cookie)
      2a: no mesmo link + '&download=1' -> devolve o arquivo .xlsx puro
    """
    sessao = requests.Session()
    sessao.headers.update({"User-Agent": USER_AGENT})

    sessao.get(url, timeout=60, allow_redirects=True)

    sep = "&" if "?" in url else "?"
    resp = sessao.get(f"{url}{sep}download=1", timeout=180, allow_redirects=True)
    resp.raise_for_status()

    ctype = resp.headers.get("Content-Type", "")
    if "spreadsheetml" not in ctype and "excel" not in ctype.lower():
        raise RuntimeError(
            f"Download de '{destino.name}' nao retornou um Excel valido "
            f"(Content-Type recebido: '{ctype}'). O link do SharePoint pode "
            f"ter expirado ou mudado de permissao — gere um novo link "
            f"'Alguem com o link' e atualize o secret no GitHub."
        )

    destino.write_bytes(resp.content)
    print(f"Baixado do SharePoint: {destino.name} ({len(resp.content)} bytes)")
    return destino


def clean(v):
    if v is None:
        return None
    if isinstance(v, str):
        v = v.replace('\xa0', ' ').strip()
        return v if v else None
    return v


def cracha_para_int(v):
    """Normaliza um crachá (número ou texto com zeros à esquerda) para int, para poder comparar."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        try:
            return int(v)
        except (ValueError, OverflowError):
            return None
    texto = str(v).strip()
    digitos = re.sub(r'\D', '', texto)
    if not digitos:
        return None
    try:
        return int(digitos)
    except ValueError:
        return None


def carregar_nomes_rh(caminho_rh):
    """Le o Visão.xlsx do RH e monta um dicionario {cracha_int: NOME EM CAIXA ALTA}."""
    if not caminho_rh or not Path(caminho_rh).exists():
        print(f"Aviso: base de crachás do RH não encontrada em '{caminho_rh}'. "
              f"Os nomes serão usados como digitados no Forms (em caixa alta).")
        return {}

    wb = openpyxl.load_workbook(caminho_rh, data_only=True)
    ws = wb.worksheets[0]
    mapa = {}
    for r in range(2, ws.max_row + 1):
        chapa = cracha_para_int(ws.cell(row=r, column=1).value)
        nome = clean(ws.cell(row=r, column=2).value)
        if chapa is not None and nome:
            mapa[chapa] = nome.upper()
    print(f"Base de crachás do RH carregada: {len(mapa)} colaboradores.")
    return mapa


def nome_corrigido(mapa_rh, cracha_raw, nome_digitado):
    """Prioriza o nome oficial do RH (pelo crachá); se não achar, usa o nome digitado no Forms.
    Sempre retorna em CAIXA ALTA."""
    chapa = cracha_para_int(cracha_raw)
    if chapa is not None and chapa in mapa_rh:
        return mapa_rh[chapa]
    if nome_digitado:
        return nome_digitado.upper()
    return None


def encontrar_planilha_dto(caminho_informado):
    if caminho_informado:
        p = Path(caminho_informado)
        if not p.exists():
            raise FileNotFoundError(f"Arquivo nao encontrado: {p}")
        return p

    if DTO_SHAREPOINT_URL:
        return baixar_sharepoint(DTO_SHAREPOINT_URL, REPO_DIR / "_dto_forms_tmp.xlsx")

    if DEFAULT_XLSX_PATH:
        p_default = Path(DEFAULT_XLSX_PATH)
        if p_default.exists():
            return p_default

    candidatos = sorted(glob.glob(str(REPO_DIR / "*.xlsx")), key=lambda p: Path(p).stat().st_mtime, reverse=True)
    if not candidatos:
        raise FileNotFoundError(
            "Nenhuma planilha do Forms encontrada (nem SharePoint, nem caminho "
            "padrao, nem na pasta do script).\n"
            "Informe o caminho como argumento:\n"
            "  python atualizar_DTO.py caminho/do/arquivo.xlsx"
        )
    return Path(candidatos[0])


def encontrar_planilha_rh():
    if CRACHA_SHAREPOINT_URL:
        return baixar_sharepoint(CRACHA_SHAREPOINT_URL, REPO_DIR / "_visao_rh_tmp.xlsx")
    if RH_XLSX_PATH and Path(RH_XLSX_PATH).exists():
        return Path(RH_XLSX_PATH)
    return None


def extrair_dados(caminho_xlsx, mapa_rh):
    wb = openpyxl.load_workbook(caminho_xlsx, data_only=True)
    ws = wb['Sheet1']
    registros = []
    for r in range(2, ws.max_row + 1):
        get = lambda c: clean(ws[f'{c}{r}'].value)
        data_raw = get('B')
        if data_raw is None:
            continue
        try:
            data_str = data_raw.strftime('%Y-%m-%d') if hasattr(data_raw, 'strftime') else str(data_raw)[:10]
        except Exception:
            data_str = str(data_raw)[:10]

        area_agricola = get('M')
        area_ind_cle = get('O')
        area_ind_qrz = get('P')
        area_manut = get('R')
        area_detalhe = area_agricola or area_ind_cle or area_ind_qrz or area_manut or None

        its = []
        for c in IT_COLS:
            v = get(c)
            if v:
                its.extend([x.strip() for x in v.split(';') if x.strip()])

        avaliador = nome_corrigido(mapa_rh, get('G'), get('F'))
        colaborador = nome_corrigido(mapa_rh, get('AO'), get('AN'))

        registros.append({
            'id': get('A'),
            'data': data_str,
            'avaliador': avaliador,
            'unidade': get('J'),
            'area': get('L'),
            'areaDetalhe': area_detalhe,
            'turno': get('AL'),
            'cargo': get('AM'),
            'colaborador': colaborador,
            'its': its,
            'revisaoIT': get('BD'),
            'pontosRevisao': get('BE'),
        })
    return registros


def atualizar_index_html(registros):
    html = INDEX_FILE.read_text(encoding='utf-8')
    data_json = json.dumps(registros, ensure_ascii=False)

    novo_html, n = re.subn(
        r'const dtoRaw = \[.*?\];',
        f'const dtoRaw = {data_json};',
        html,
        count=1,
        flags=re.DOTALL,
    )
    if n == 0:
        raise RuntimeError("Nao encontrei 'const dtoRaw = [...]' no index.html — verifique o arquivo.")

    agora = datetime.now().strftime('%d/%m/%Y %H:%M')
    novo_html, n2 = re.subn(
        r"document\.getElementById\('lastUpdate'\)\.textContent = 'Atualizado em .*?';",
        f"document.getElementById('lastUpdate').textContent = 'Atualizado em {agora}';",
        novo_html,
        count=1,
    )

    INDEX_FILE.write_text(novo_html, encoding='utf-8')
    print(f"index.html atualizado: {len(registros)} registros, {agora}")


def rodar_git(cmd):
    print(f"$ git {' '.join(cmd)}")
    result = subprocess.run(["git"] + cmd, cwd=REPO_DIR, capture_output=True, text=True)
    if result.stdout.strip():
        print(result.stdout.strip())
    if result.returncode != 0:
        print(result.stderr.strip())
    return result


def main():
    caminho_arg = sys.argv[1] if len(sys.argv) > 1 else None
    planilha = encontrar_planilha_dto(caminho_arg)
    print(f"Usando planilha: {planilha.name}")

    caminho_rh = encontrar_planilha_rh()
    mapa_rh = carregar_nomes_rh(caminho_rh) if caminho_rh else {}

    registros = extrair_dados(planilha, mapa_rh)
    atualizar_index_html(registros)

    # REMOVIDO: As linhas de git add, commit e push foram retiradas daqui
    # para evitar o bloqueio de segurança (Exit Code 129) no GitHub Actions.
    print("\nArquivo index.html atualizado localmente com sucesso!")



if __name__ == "__main__":
    main()
