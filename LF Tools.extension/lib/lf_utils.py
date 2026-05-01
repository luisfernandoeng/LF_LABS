# -*- coding: utf-8 -*-
"""
lf_utils.py — Biblioteca compartilhada LF Tools
================================================
Utilitários PUROS para todos os scripts pyRevit da extensão.
Compatível com IronPython 2, IronPython 3 e CPython.

SEM dependências de Revit API, .NET ou pyRevit neste módulo.
Para utilitários que precisam da Revit API, use lf_revit.py

Uso nos scripts:
    from lf_utils import DebugLogger, get_script_config, save_script_config
    from lf_utils import slugify, truncate, flatten, chunk, deep_merge
    from lf_utils import format_meters, parse_bool, safe_int, safe_float
    from lf_utils import now_str, elapsed_str, Timer
"""

import io
import os
import re
import time
import json
import math
import traceback


# =============================================================================
#  DEBUG LOGGER
# =============================================================================

class _DebugCtx(object):
    """Context manager retornado por DebugLogger.ctx(). Não instanciar diretamente."""

    def __init__(self, logger, label):
        self._logger = logger
        self._label  = label
        self._t0     = None

    def __enter__(self):
        self._t0 = time.time()
        if self._logger.enabled:
            self._logger._print("", u"  {} >> {}".format(self._logger._SUB[:20], self._label))
        return self

    def __exit__(self, exc_type, _exc_val, _exc_tb):
        elapsed = time.time() - (self._t0 or time.time())
        if self._logger.enabled:
            status = u"ERRO" if exc_type else u"ok"
            self._logger._print(
                "[TIMER]",
                u"<< {} — {:.3f}s [{}]".format(self._label, elapsed, status)
            )
        return False  # não suprime exceções


class DebugLogger(object):
    """
    Logger de debug para scripts pyRevit.
    Controlado pela flag DEBUG_MODE no topo de cada script ou por checkbox na UI.

    Uso básico:
        dbg = DebugLogger(DEBUG_MODE)
        dbg.section("Iniciando")
        dbg.info("Elemento: {}".format(el.Id))
        dbg.warn("Sem conector")
        dbg.exc("Falhou ao criar", e)   # traceback completo
        dbg.var("elemento", el)         # nome + repr + tipo
        with dbg.ctx("Busca de pontos"): # timing automático
            ...

    Métodos:
      dbg.section(titulo)      — separador visual de seção
      dbg.sub(titulo)          — sub-separador
      dbg.info(msg)            — informação geral
      dbg.debug(msg)           — detalhe técnico
      dbg.warn(msg)            — aviso (sempre impresso)
      dbg.error(msg)           — erro (sempre impresso)
      dbg.exc(msg, e=None)     — exceção com traceback completo (sempre impresso)
      dbg.var(name, value)     — imprime nome = repr(value) [tipo]
      dbg.dump(label, obj)     — despeja repr de um objeto
      dbg.xyz(label, pt)       — coordenadas XYZ (ponto Revit ou tupla)
      dbg.table(headers, rows) — tabela formatada
      dbg.timer_start(label)   — inicia cronômetro nomeado
      dbg.timer_end(label)     — encerra e imprime tempo
      dbg.ctx(label)           — context manager com timing automático
      dbg.result(ok, msg)      — OK/FAIL visual
      dbg.sep()                — linha separadora simples
    """

    _SECTION = "=" * 60
    _SUB     = "-" * 40

    def __init__(self, enabled=False, log_func=None, timestamps=False):
        self.enabled     = bool(enabled)
        self._timers     = {}
        self.log_func    = log_func
        self._timestamps = bool(timestamps)

    def _print(self, prefix, msg):
        import sys
        ts = time.strftime("[%H:%M:%S] ") if self._timestamps else ""
        line = u"{}{}{}".format(ts, prefix + " " if prefix else "", msg)

        if self.log_func:
            try:
                self.log_func(line)
            except:
                pass

        try:
            print(line)
        except Exception:
            try:
                sys.__stdout__.write(line + "\n")
                sys.__stdout__.flush()
            except Exception:
                pass

    def section(self, titulo):
        if not self.enabled:
            return
        self._print("", "")
        self._print("", self._SECTION)
        self._print("", "  {}".format(titulo.upper()))
        self._print("", self._SECTION)

    def sub(self, titulo):
        if not self.enabled:
            return
        self._print("", "  {} {}".format(self._SUB, titulo))

    def sep(self):
        if self.enabled:
            self._print("", self._SUB)

    def info(self, msg):
        if self.enabled:
            self._print("[INFO ]", msg)

    def debug(self, msg):
        if self.enabled:
            self._print("[DEBUG]", msg)

    def warn(self, msg):
        if self.enabled:
            self._print("[WARN ]", msg)

    def error(self, msg):
        """Sempre impresso — independente de enabled."""
        self._print("[ERROR]", msg)

    def dump(self, label, obj):
        if self.enabled:
            self._print("[DUMP ]", "{} = {!r}".format(label, obj))

    def xyz(self, label, pt):
        """Aceita objeto com .X/.Y/.Z (Revit XYZ) ou tupla/lista (x, y, z)."""
        if not self.enabled:
            return
        if pt is None:
            self._print("[XYZ  ]", "{} = None".format(label))
        elif hasattr(pt, "X") and hasattr(pt, "Y") and hasattr(pt, "Z"):
            self._print("[XYZ  ]", "{} = ({:.4f}, {:.4f}, {:.4f}) ft".format(
                label, pt.X, pt.Y, pt.Z))
        else:
            try:
                x, y, z = pt[0], pt[1], pt[2]
                self._print("[XYZ  ]", "{} = ({:.4f}, {:.4f}, {:.4f})".format(label, x, y, z))
            except Exception:
                self._print("[XYZ  ]", "{} = {!r}".format(label, pt))

    def table(self, headers, rows):
        """Imprime uma tabela simples formatada no console."""
        if not self.enabled:
            return
        col_widths = [len(str(h)) for h in headers]
        for row in rows:
            for i, cell in enumerate(row):
                col_widths[i] = max(col_widths[i], len(str(cell)))
        fmt = "  ".join("{{:<{}}}".format(w) for w in col_widths)
        self._print("[TABLE]", fmt.format(*[str(h) for h in headers]))
        self._print("[TABLE]", "-" * sum(col_widths + [2 * (len(col_widths) - 1)]))
        for row in rows:
            self._print("[TABLE]", fmt.format(*[str(c) for c in row]))

    def timer_start(self, label):
        self._timers[label] = time.time()
        if self.enabled:
            self._print("[TIMER]", "START  '{}'".format(label))

    def timer_end(self, label):
        if label in self._timers:
            elapsed = time.time() - self._timers.pop(label)
            if self.enabled:
                self._print("[TIMER]", "END    '{}' — {:.3f}s".format(label, elapsed))
        elif self.enabled:
            self._print("[TIMER]", "END    '{}' (sem start registrado)".format(label))

    def result(self, ok, msg):
        if self.enabled:
            tag = "[  OK ]" if ok else "[FAIL ]"
            self._print(tag, msg)

    def exc(self, msg, e=None):
        """Sempre impresso. Mostra mensagem + traceback completo da exceção atual ou de `e`."""
        self._print("[EXCEP]", msg)
        tb_str = ""
        try:
            tb_str = traceback.format_exc()
        except Exception:
            pass
        if not tb_str or tb_str.strip() == "None":
            tb_str = repr(e) if e is not None else ""
        if tb_str:
            for line in tb_str.splitlines():
                if line.strip():
                    self._print("      |", line)

    def var(self, name, value):
        """Imprime: name = repr(value)  [tipo]. Útil para inspecionar variáveis."""
        if not self.enabled:
            return
        try:
            type_name = type(value).__name__
            self._print("[VAR  ]", u"{} = {!r}  [{}]".format(name, value, type_name))
        except Exception:
            self._print("[VAR  ]", u"{} = <erro ao formatar>".format(name))

    def ctx(self, label):
        """Retorna context manager que imprime sub-seção + tempo decorrido ao sair."""
        return _DebugCtx(self, label)


# =============================================================================
#  TIMER UTILITÁRIO (context manager)
# =============================================================================

class Timer(object):
    """
    Cronômetro simples como context manager.

    Uso:
        with Timer("busca de elementos") as t:
            # código demorado
            pass
        print("Demorou:", t.elapsed_str)

    Ou sem context manager:
        t = Timer()
        t.start()
        # código
        print(t.stop())  # retorna string "1.234s"
    """
    def __init__(self, label=""):
        self.label   = label
        self._start  = None
        self.elapsed = 0.0

    def start(self):
        self._start = time.time()
        return self

    def stop(self):
        self.elapsed = time.time() - (self._start or time.time())
        return self.elapsed_str

    @property
    def elapsed_str(self):
        return "{:.3f}s".format(self.elapsed)

    def __enter__(self):
        self.start()
        return self

    def __exit__(self, *args):
        self.stop()
        if self.label:
            try:
                print("[TIMER] {} — {}".format(self.label, self.elapsed_str))
            except Exception:
                pass


# =============================================================================
#  CONFIG JSON
# =============================================================================

def get_script_config(script_path, defaults=None):
    """
    Lê configurações de um arquivo JSON ao lado do script.
    Substitui script.get_config() sem importar pyrevit.script.

    Parâmetros:
        script_path — __commandpath__ do script
        defaults    — dict com valores padrão (usados se a chave não existir no JSON)

    Retorna:
        dict com as configurações (mesclado: defaults + valores salvos)

    Uso:
        settings = get_script_config(__commandpath__, defaults={
            'modo': 'auto',
            'distancia_min': 0.1,
        })
        modo = settings.get('modo')
    """
    config_path = os.path.join(os.path.dirname(str(script_path)), "config.json")
    data = dict(defaults or {})
    if os.path.isfile(config_path):
        try:
            with io.open(config_path, "r", encoding="utf-8") as f:
                saved = json.load(f)
            data.update(saved)
        except Exception:
            pass
    return data


def save_script_config(script_path, settings):
    """
    Salva configurações em JSON ao lado do script.
    Substitui script.save_config() sem importar pyrevit.script.

    Uso:
        save_script_config(__commandpath__, {'modo': 'manual', 'distancia_min': 0.2})
    """
    config_path = os.path.join(os.path.dirname(str(script_path)), "config.json")
    try:
        with io.open(config_path, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# =============================================================================
#  UTILITÁRIOS DE STRING
# =============================================================================

def slugify(text, separator="-"):
    """
    Converte texto em slug: minúsculas, sem acentos, separador entre palavras.
    Útil para gerar nomes de arquivo, chaves de config, IDs.

    Exemplo:
        slugify("Painel Elétrico 01") -> "painel-eletrico-01"
        slugify("Quadro Geral", "_") -> "quadro_geral"
    """
    # Remove acentos manualmente (sem unicodedata — compatível com IronPython 2)
    _accents = {
        u"á": "a", u"à": "a", u"ã": "a", u"â": "a", u"ä": "a",
        u"é": "e", u"è": "e", u"ê": "e", u"ë": "e",
        u"í": "i", u"ì": "i", u"î": "i", u"ï": "i",
        u"ó": "o", u"ò": "o", u"õ": "o", u"ô": "o", u"ö": "o",
        u"ú": "u", u"ù": "u", u"û": "u", u"ü": "u",
        u"ç": "c", u"ñ": "n",
        u"Á": "a", u"À": "a", u"Ã": "a", u"Â": "a", u"Ä": "a",
        u"É": "e", u"È": "e", u"Ê": "e", u"Ë": "e",
        u"Í": "i", u"Ì": "i", u"Î": "i", u"Ï": "i",
        u"Ó": "o", u"Ò": "o", u"Õ": "o", u"Ô": "o", u"Ö": "o",
        u"Ú": "u", u"Ù": "u", u"Û": "u", u"Ü": "u",
        u"Ç": "c", u"Ñ": "n",
    }
    result = u""
    for ch in text:
        result += _accents.get(ch, ch)
    result = result.lower()
    result = re.sub(r"[^a-z0-9]+", separator, result)
    result = result.strip(separator)
    return result


def truncate(text, max_len=50, suffix="..."):
    """
    Trunca texto ao comprimento máximo, adicionando sufixo se necessário.

    Exemplo:
        truncate("Quadro Geral de Distribuição", 20) -> "Quadro Geral de D..."
    """
    text = str(text)
    if len(text) <= max_len:
        return text
    return text[: max_len - len(suffix)] + suffix


def pad_number(n, digits=2):
    """
    Formata número com zeros à esquerda.

    Exemplo:
        pad_number(3, 2)  -> "03"
        pad_number(42, 4) -> "0042"
    """
    return str(int(n)).zfill(digits)


def clean_name(text):
    """
    Remove caracteres inválidos para nomes de elementos/vistas no Revit.
    Mantém: letras, números, espaços, hífens, underscores, pontos, parênteses.

    Exemplo:
        clean_name("Vista/Planta [01]") -> "Vista-Planta (01)"
    """
    text = re.sub(r'[\\/:*?"<>|{}]', '-', text)
    text = re.sub(r"\[", "(", text)
    text = re.sub(r"\]", ")", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_spaces(text):
    """Remove espaços extras e normaliza o texto."""
    return re.sub(r"\s+", " ", str(text)).strip()


def extract_numbers(text):
    """
    Extrai todos os números (inteiros e decimais) de uma string.

    Exemplo:
        extract_numbers("Cabo 2x4mm²") -> [2.0, 4.0]
        extract_numbers("3ª fase - 127V") -> [3.0, 127.0]
    """
    return [float(x) for x in re.findall(r"\d+(?:\.\d+)?", str(text))]


def parse_bool(value, default=False):
    """
    Converte valor para bool de forma robusta.
    Aceita: True/False, 1/0, "true"/"false", "sim"/"não", "yes"/"no", "1"/"0".

    Exemplo:
        parse_bool("sim")   -> True
        parse_bool("nao")   -> False
        parse_bool("yes")   -> True
        parse_bool("1")     -> True
    """
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    s = str(value).strip().lower()
    if s in ("true", "1", "sim", "yes", "s", "y", "on"):
        return True
    if s in ("false", "0", "nao", "não", "no", "n", "off"):
        return False
    return default


def safe_int(value, default=0):
    """Converte para int sem lançar exceção. Retorna default se falhar."""
    try:
        return int(value)
    except (ValueError, TypeError):
        return default


def safe_float(value, default=0.0):
    """Converte para float sem lançar exceção. Retorna default se falhar."""
    try:
        return float(value)
    except (ValueError, TypeError):
        return default


# =============================================================================
#  UTILITÁRIOS DE LISTAS / COLEÇÕES
# =============================================================================

def flatten(lst):
    """
    Achata listas aninhadas (1 nível).

    Exemplo:
        flatten([[1, 2], [3, 4], [5]]) -> [1, 2, 3, 4, 5]
    """
    result = []
    for item in lst:
        if isinstance(item, (list, tuple)):
            result.extend(item)
        else:
            result.append(item)
    return result


def chunk(lst, size):
    """
    Divide uma lista em sublistas de tamanho `size`.

    Exemplo:
        chunk([1,2,3,4,5], 2) -> [[1,2], [3,4], [5]]
    """
    return [lst[i: i + size] for i in range(0, len(lst), size)]


def unique(lst, key=None):
    """
    Remove duplicatas preservando ordem.

    Parâmetros:
        key — função opcional para extrair a chave de comparação

    Exemplo:
        unique([3, 1, 2, 1, 3]) -> [3, 1, 2]
        unique(elements, key=lambda e: e.Id.IntegerValue)
    """
    seen = set()
    result = []
    for item in lst:
        k = key(item) if key else item
        if k not in seen:
            seen.add(k)
            result.append(item)
    return result


def group_by(lst, key):
    """
    Agrupa itens de uma lista por uma chave.

    Retorna:
        dict { chave: [itens] }

    Exemplo:
        group_by(elements, key=lambda e: e.Category.Name)
        # -> {"Equipamentos Elétricos": [...], "Luminárias": [...]}
    """
    result = {}
    for item in lst:
        k = key(item)
        if k not in result:
            result[k] = []
        result[k].append(item)
    return result


def deep_merge(base, override):
    """
    Mescla dois dicts recursivamente (override sobrescreve base).
    Útil para mesclar configs padrão com configs do usuário.

    Exemplo:
        base     = {"dist": 1.0, "modo": "auto", "cores": {"fundo": "#fff"}}
        override = {"dist": 2.0, "cores": {"texto": "#000"}}
        deep_merge(base, override)
        # -> {"dist": 2.0, "modo": "auto", "cores": {"fundo": "#fff", "texto": "#000"}}
    """
    result = dict(base)
    for k, v in override.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = deep_merge(result[k], v)
        else:
            result[k] = v
    return result


# =============================================================================
#  UTILITÁRIOS DE MEDIDAS / CONVERSÃO
# =============================================================================

# Fatores de conversão: 1 pé = 304.8 mm
_FT_TO_MM  = 304.8
_FT_TO_CM  = 30.48
_FT_TO_M   = 0.3048
_MM_TO_FT  = 1.0 / 304.8
_M_TO_FT   = 1.0 / 0.3048


def ft_to_mm(feet):
    """Converte pés (unidade interna Revit) para milímetros."""
    return feet * _FT_TO_MM

def ft_to_m(feet):
    """Converte pés (unidade interna Revit) para metros."""
    return feet * _FT_TO_M

def mm_to_ft(mm):
    """Converte milímetros para pés (unidade interna Revit)."""
    return mm * _MM_TO_FT

def m_to_ft(meters):
    """Converte metros para pés (unidade interna Revit)."""
    return meters * _M_TO_FT


def format_length(feet, unit="mm", decimals=0):
    """
    Formata comprimento em pés para string legível.

    Parâmetros:
        feet     — valor em pés (unidade interna Revit)
        unit     — "mm", "cm" ou "m"
        decimals — casas decimais

    Exemplo:
        format_length(3.28084, unit="m", decimals=2) -> "1.00 m"
        format_length(0.328084, unit="mm")            -> "100 mm"
    """
    if unit == "m":
        val = feet * _FT_TO_M
        fmt = "{:.{d}f} m".format(val, d=decimals)
    elif unit == "cm":
        val = feet * _FT_TO_CM
        fmt = "{:.{d}f} cm".format(val, d=decimals)
    else:
        val = feet * _FT_TO_MM
        fmt = "{:.{d}f} mm".format(val, d=decimals)
    return fmt


def format_area(sq_feet, unit="m2", decimals=2):
    """
    Formata área em pés² para string legível.

    Exemplo:
        format_area(10.764, unit="m2") -> "1.00 m²"
    """
    if unit == "m2":
        val = sq_feet * (_FT_TO_M ** 2)
        return "{:.{d}f} m²".format(val, d=decimals)
    else:
        val = sq_feet * (_FT_TO_MM ** 2)
        return "{:.{d}f} mm²".format(val, d=decimals)


def round_to(value, step):
    """
    Arredonda valor para o múltiplo mais próximo de step.

    Exemplo:
        round_to(37.3, 5)   -> 35
        round_to(0.123, 0.05) -> 0.10
    """
    return round(float(value) / step) * step


def clamp(value, min_val, max_val):
    """Limita value entre min_val e max_val."""
    return max(min_val, min(max_val, value))


# =============================================================================
#  UTILITÁRIOS DE DATA / HORA
# =============================================================================

def now_str(fmt="%Y-%m-%d %H:%M:%S"):
    """
    Retorna data/hora atual como string formatada.

    Exemplo:
        now_str()               -> "2026-04-17 21:00:00"
        now_str("%d/%m/%Y")     -> "17/04/2026"
    """
    return time.strftime(fmt)


def elapsed_str(seconds):
    """
    Converte segundos em string legível: "2h 03m 15s", "45m 10s", "8s".

    Exemplo:
        elapsed_str(7395) -> "2h 03m 15s"
        elapsed_str(90)   -> "1m 30s"
        elapsed_str(8)    -> "8s"
    """
    s = int(seconds)
    h = s // 3600
    m = (s % 3600) // 60
    sec = s % 60
    if h:
        return "{}h {:02d}m {:02d}s".format(h, m, sec)
    elif m:
        return "{}m {:02d}s".format(m, sec)
    else:
        return "{}s".format(sec)


# =============================================================================
#  UTILITÁRIOS DE ARQUIVOS
# =============================================================================

def ensure_dir(path):
    """
    Cria o diretório e todos os pais se não existirem.
    Equivalente a os.makedirs com exist_ok=True (Python 2 compatible).

    Retorna:
        O próprio path.
    """
    if not os.path.isdir(path):
        try:
            os.makedirs(path)
        except OSError:
            pass
    return path


def safe_filename(name, replacement="_"):
    """
    Remove/substitui caracteres inválidos para nome de arquivo.

    Exemplo:
        safe_filename("Vista: Piso 1/2") -> "Vista_ Piso 1_2"
    """
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', replacement, str(name)).strip()


def read_json(path, default=None):
    """
    Lê um arquivo JSON com tratamento de erros.

    Retorna:
        Conteúdo do JSON ou default se falhar.
    """
    try:
        with io.open(str(path), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default if default is not None else {}


def write_json(path, data, indent=2):
    """
    Escreve dados em JSON com tratamento de erros.

    Retorna:
        True se salvou, False se falhou.
    """
    try:
        ensure_dir(os.path.dirname(str(path)))
        with io.open(str(path), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=indent)
        return True
    except Exception:
        return False


def clear_pyrevit_cache(silent=False):
    """
    Remove o cache do pyRevit para o usuário atual.
    Exibe resultado via print (compatível com qualquer engine).

    Parâmetros:
        silent — se True, não imprime resultado

    Retorna:
        True se limpou com sucesso, False se houve erro.

    ATENÇÃO: Recomenda-se reiniciar o Revit após executar.
    """
    import shutil
    appdata = os.environ.get("APPDATA", "")
    if not appdata:
        if not silent:
            print("[CACHE] Variável APPDATA não encontrada.")
        return False

    cache_paths = [
        os.path.join(appdata, "pyRevit", "cache"),
        os.path.join(appdata, "pyRevit-Master", "cache"),
    ]
    cleaned, errors = [], []

    for path in cache_paths:
        if os.path.isdir(path):
            try:
                shutil.rmtree(path)
                cleaned.append(path)
            except Exception as e:
                errors.append("{}: {}".format(path, e))

    if not silent:
        if cleaned:
            print("[CACHE] Removido: " + ", ".join(cleaned))
        if errors:
            print("[CACHE] Erro: " + ", ".join(errors))
        if not cleaned and not errors:
            print("[CACHE] Nenhuma pasta de cache encontrada.")

    return len(errors) == 0


# =============================================================================
#  REGRAS DE TRANSFORMAÇÃO DE NOME (engine de renomeação LF Tools)
# =============================================================================

def apply_rename_rules(
    original,
    name_pattern="{Original}",
    regex_find="",
    regex_replace="",
    find_simple="",
    replace_simple="",
    numbering=None,
    case_converter=None,
    use_pure_regex=False,
    param_resolver=None,
):
    """
    Engine de renomeação reutilizável da LF Tools.
    Aplica sequência de regras a uma string original.

    Parâmetros:
        original       — string original do elemento
        name_pattern   — padrão de composição com placeholders {Original}, {Param}
        regex_find     — padrão de busca (smart pattern ou regex puro)
        regex_replace  — substituição para o regex
        find_simple    — busca simples (case-insensitive)
        replace_simple — substituição para a busca simples
        numbering      — dict com keys: separator, current_val, padding, position (0=fim, 1=início)
        case_converter — "upper", "lower" ou "title"
        use_pure_regex — se True, usa regex_find como regex puro (não smart pattern)
        param_resolver — callable(param_name) -> str, para resolver {Param} dinâmico

    Retorna:
        string com o novo nome após todas as regras

    Exemplo:
        novo = apply_rename_rules(
            "QUADRO 01",
            name_pattern="{Original}",
            find_simple="QUADRO",
            replace_simple="QD",
            case_converter="title",
        )
        # -> "Qd 01"
    """
    txt = str(original)

    # 1. Composição com padrão e parâmetros dinâmicos {NomeParam}
    txt = name_pattern.replace("{Original}", txt)

    def _replace_dynamic(match):
        param_name = match.group(1)
        if param_name == "Original":
            return str(original)
        if param_resolver is not None:
            resolved = param_resolver(param_name)
            if resolved is not None:
                return str(resolved)
        return match.group(0)

    txt = re.sub(r"\{(.*?)\}", _replace_dynamic, txt)

    # 2. Padrão inteligente ou Regex puro
    if regex_find:
        try:
            if use_pure_regex:
                txt = re.sub(regex_find, regex_replace or "", txt)
            else:
                pattern = ""
                for char in regex_find:
                    if char == "#":
                        pattern += r"\d"
                    elif char == "@":
                        pattern += r"[a-zA-Z]"
                    elif char == "?":
                        pattern += r"."
                    elif char in "()|":
                        pattern += char
                    elif char in ".*+^{}\\$[]":
                        pattern += "\\" + char
                    else:
                        pattern += char
                repl = re.sub(r"\$(\d+)", r"\\g<\1>", regex_replace or "")
                txt = re.sub(pattern, repl, txt)
        except Exception:
            pass

    # 3. Substituição simples (case-insensitive)
    if find_simple:
        try:
            txt = re.sub(re.escape(find_simple), replace_simple or "", txt, flags=re.IGNORECASE)
        except Exception:
            pass

    # 4. Numeração sequencial
    if numbering:
        sep = numbering.get("separator", "-")
        val = numbering.get("current_val", 1)
        pad = numbering.get("padding", 2)
        pos = numbering.get("position", 0)
        num_str = str(int(val)).zfill(pad)
        if pos == 1:
            txt = "{}{}{}".format(num_str, sep, txt)
        else:
            txt = "{}{}{}".format(txt, sep, num_str)

    # 5. Conversão de caixa
    if case_converter == "upper":
        txt = txt.upper()
    elif case_converter == "lower":
        txt = txt.lower()
    elif case_converter == "title":
        txt = txt.title()

    return txt

# =============================================================================
#  REVIT API HELPERS (FAILURE PREPROCESSORS)
# =============================================================================

def make_warning_swallower():
    """Retorna uma instância de WarningSwallower para usar em transações."""
    try:
        import clr
        clr.AddReference('RevitAPI')
        from Autodesk.Revit.DB import (
            IFailuresPreprocessor, FailureProcessingResult, FailureSeverity)

        class WarningSwallower(IFailuresPreprocessor):
            def PreprocessFailures(self, failuresAccessor):
                try:
                    failuresAccessor.DeleteAllWarnings()
                except Exception:
                    pass
                try:
                    for f in list(failuresAccessor.GetFailureMessages()):
                        try:
                            sev = f.GetSeverity()
                            if sev == FailureSeverity.Warning:
                                failuresAccessor.DeleteWarning(f)
                            elif sev == FailureSeverity.Error:
                                try:
                                    failuresAccessor.ResolveFailure(f)
                                except Exception:
                                    failuresAccessor.DeleteWarning(f)
                        except Exception:
                            pass
                except Exception:
                    pass
                return FailureProcessingResult.ProceedWithCommit

        return WarningSwallower()
    except Exception:
        return None
