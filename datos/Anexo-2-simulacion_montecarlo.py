"""Simulación de Monte Carlo del modelo financiero de upcycling de palas eólicas.

Un solo archivo que hace todo:
  1. Lee el Excel congelado y arma una réplica en Python: evalúa cada fórmula del libro,
     en el mismo orden que Excel, con un vector de N valores por celda (una iteración por valor).
  2. Verifica que la réplica reproduce el VAN del Excel antes de simular.
  3. Sortea las siete variables y corre las N iteraciones de una vez.
  4. Guarda seis gráficos PNG y un Excel de resultados de dos hojas (Resultados e Iteraciones).

Uso:
    python simulacion_montecarlo.py "Modelo_financiero_palas_eolicas_v5_15_Opcion_2_10_a_os.xlsx"

Requiere Python 3 con numpy, openpyxl y matplotlib (pip install numpy openpyxl matplotlib).
Los resultados quedan en la carpeta "resultados_montecarlo", junto al Excel del modelo.
"""
import os

# ---------------- Configuración ----------------
MODELO = 'Modelo_financiero_palas_eolicas_v5_15_Opcion_2_10_a_os.xlsx'
N = 10000                # iteraciones
SEMILLA = 20261017       # semilla fija: misma semilla, mismos resultados
ASK = 748025             # ask con reserva de liquidez del modelo (USD)
GATE_VAN_CERO = 517.76   # gate fee de VAN cero (USD/t), hoja Riesgos y sensibilidad del modelo

# ==========================================================================
# 1. RÉPLICA DEL MODELO
# ==========================================================================
import re
import numpy as np
import openpyxl
from openpyxl.formula.tokenizer import Tokenizer, Token
from openpyxl.utils import column_index_from_string, get_column_letter

# --------------------------------------------------------------------------
# Lectura del libro
# --------------------------------------------------------------------------
REF_RE = re.compile(r"^(?:(?:'((?:[^']|'')+)'|([^!:']+))!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$")


class Rango:
    """Un rango de celdas: lista 2D de claves (hoja, fila, col)."""
    def __init__(self, hoja, r1, c1, r2, c2):
        self.hoja, self.r1, self.c1, self.r2, self.c2 = hoja, r1, c1, r2, c2

    def claves(self):
        return [[(self.hoja, r, c) for c in range(self.c1, self.c2 + 1)]
                for r in range(self.r1, self.r2 + 1)]

    def plano(self):
        return [k for fila in self.claves() for k in fila]


class Crit:
    """Criterio de SUMIF/COUNTIF armado con '"<="&X'."""
    def __init__(self, op, val):
        self.op, self.val = op, val


def _ref(texto, hoja_actual):
    m = REF_RE.match(texto)
    if not m:
        raise ValueError('Referencia no reconocida: ' + texto)
    hoja = (m.group(1).replace("''", "'") if m.group(1) else m.group(2)) or hoja_actual
    c1, r1 = column_index_from_string(m.group(3)), int(m.group(4))
    if m.group(5):
        c2, r2 = column_index_from_string(m.group(5)), int(m.group(6))
        return Rango(hoja, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))
    return (hoja, r1, c1)


# --------------------------------------------------------------------------
# Parser de fórmulas a árbol
# --------------------------------------------------------------------------
PREC = {'=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3,
        '*': 4, '/': 4, '^': 5}


def parsear(formula, hoja):
    toks = [t for t in Tokenizer(formula).items if t.type != Token.WSPACE]
    pos = [0]

    def ver():
        return toks[pos[0]] if pos[0] < len(toks) else None

    def tomar():
        t = toks[pos[0]]
        pos[0] += 1
        return t

    def primario():
        t = tomar()
        if t.type == Token.OP_PRE:
            arg = expr(6)
            return ('neg', arg) if t.value == '-' else arg
        if t.type == Token.PAREN and t.subtype == Token.OPEN:
            e = expr(0)
            tomar()  # ')'
            return e
        if t.type == Token.FUNC and t.subtype == Token.OPEN:
            nombre = t.value[:-1].upper()
            args = []
            if not (ver().type == Token.FUNC and ver().subtype == Token.CLOSE):
                while True:
                    args.append(expr(0))
                    s = tomar()
                    if s.type == Token.FUNC and s.subtype == Token.CLOSE:
                        break
            else:
                tomar()
            return ('fn', nombre, args)
        if t.type == Token.OPERAND:
            if t.subtype == Token.NUMBER:
                return ('num', float(t.value))
            if t.subtype == Token.TEXT:
                return ('str', t.value[1:-1].replace('""', '"'))
            if t.subtype == Token.LOGICAL:
                return ('num', 1.0 if t.value.upper() == 'TRUE' else 0.0)
            if t.subtype == Token.RANGE:
                r = _ref(t.value, hoja)
                return ('rng', r) if isinstance(r, Rango) else ('ref', r)
        raise ValueError('Token inesperado %r en %s' % (t.value, formula))

    def expr(pmin):
        izq = primario()
        while True:
            t = ver()
            if t is not None and t.type == Token.OP_POST and t.value == '%':
                tomar()
                izq = ('op', '/', izq, ('num', 100.0))
                continue
            if t is None or t.type != Token.OP_IN:
                break
            p = PREC[t.value]
            if p < pmin:
                break
            tomar()
            der = expr(p + 1)
            izq = ('op', t.value, izq, der)
        return izq

    arbol = expr(0)
    if pos[0] != len(toks):
        raise ValueError('Sobran tokens en ' + formula)
    return arbol


def dependencias(nodo, out):
    tipo = nodo[0]
    if tipo == 'ref':
        out.add(nodo[1])
    elif tipo == 'rng':
        out.update(nodo[1].plano())
    elif tipo == 'neg':
        dependencias(nodo[1], out)
    elif tipo == 'op':
        dependencias(nodo[2], out)
        dependencias(nodo[3], out)
    elif tipo == 'fn':
        for a in nodo[2]:
            dependencias(a, out)


# --------------------------------------------------------------------------
# Evaluación vectorizada
# --------------------------------------------------------------------------
def _num(v):
    """Convierte a número como Excel: vacío = 0, booleano = 1/0."""
    if v is None:
        return 0.0
    if isinstance(v, np.ndarray):
        return v.astype(float) if v.dtype != object else v
    if isinstance(v, (bool, np.bool_)):
        return float(v)
    if isinstance(v, str):
        return np.nan
    return v


def _bool(v):
    v = _num(v)
    if isinstance(v, np.ndarray):
        return v != 0
    return bool(v)


def _roundup(x, d):
    f = 10.0 ** d
    a = np.abs(x) * f
    return np.sign(x) * np.ceil(np.round(a, 9)) / f


def _cmp(op, a, b):
    if isinstance(a, str) or isinstance(b, str):
        a = '' if a is None else a
        b = '' if b is None else b
        if isinstance(a, str) and isinstance(b, str):
            a, b = a.lower(), b.lower()
    else:
        a, b = _num(a), _num(b)
    return {'=': lambda: a == b, '<>': lambda: a != b, '<': lambda: a < b,
            '>': lambda: a > b, '<=': lambda: a <= b, '>=': lambda: a >= b}[op]()


def _cumple(valor, crit):
    if isinstance(crit, Crit):
        return _cmp(crit.op, valor, crit.val)
    if isinstance(crit, str):
        m = re.match(r'^(<=|>=|<>|<|>|=)?(.*)$', crit)
        op, resto = m.group(1) or '=', m.group(2)
        try:
            return _cmp(op, valor, float(resto))
        except ValueError:
            return _cmp(op, valor, resto)
    return _cmp('=', valor, crit)


class Modelo:
    def __init__(self, ruta):
        self.ruta = ruta
        wf = openpyxl.load_workbook(ruta)
        wv = openpyxl.load_workbook(ruta, data_only=True)
        self.formulas, self.constantes, self.cache_excel = {}, {}, {}
        for ws in wf:
            for fila in ws.iter_rows():
                for c in fila:
                    k = (ws.title, c.row, c.column)
                    v = c.value
                    self.cache_excel[k] = wv[ws.title].cell(c.row, c.column).value
                    if isinstance(v, str) and v.startswith('='):
                        self.formulas[k] = v
                    elif v is not None:
                        self.constantes[k] = v
        # Códigos de parámetros (columna A de Parámetros) -> celda de valor (columna C)
        self.codigos = {}
        for (h, r, c), v in self.constantes.items():
            if h == 'Parámetros' and c == 1 and isinstance(v, str):
                self.codigos[v.strip()] = ('Parámetros', r, 3)
        self.arboles = {}

    # -- árbol y orden de cálculo de lo que hace falta para las salidas ------
    def _arbol(self, k):
        if k not in self.arboles:
            self.arboles[k] = parsear(self.formulas[k], k[0])
        return self.arboles[k]

    def orden(self, objetivos):
        visto, orden = set(), []
        pila = [(k, False) for k in objetivos]
        while pila:
            k, listo = pila.pop()
            if listo:
                orden.append(k)
                continue
            if k in visto:
                continue
            visto.add(k)
            if k not in self.formulas:
                continue
            pila.append((k, True))
            deps = set()
            dependencias(self._arbol(k), deps)
            for d in deps:
                if d not in visto:
                    pila.append((d, False))
        return orden

    # -- evaluación ----------------------------------------------------------
    def _valor(self, k):
        if k in self.V:
            return self.V[k]
        return self.constantes.get(k)

    def _ev(self, nodo):
        t = nodo[0]
        if t == 'num' or t == 'str':
            return nodo[1]
        if t == 'ref':
            return self._valor(nodo[1])
        if t == 'rng':
            return nodo[1]
        if t == 'neg':
            return -_num(self._ev(nodo[1]))
        if t == 'op':
            op = nodo[1]
            a, b = self._ev(nodo[2]), self._ev(nodo[3])
            if op == '&':
                if isinstance(a, str) and a in ('<', '<=', '>', '>=', '=', '<>'):
                    return Crit(a, _num(b))
                if isinstance(a, np.ndarray) or isinstance(b, np.ndarray):
                    raise ValueError('Concatenación vectorial no soportada')
                return ('' if a is None else str(a)) + ('' if b is None else str(b))
            if op in ('=', '<>', '<', '>', '<=', '>='):
                return _cmp(op, a, b)
            a, b = _num(a), _num(b)
            if op == '+':
                return a + b
            if op == '-':
                return a - b
            if op == '*':
                return a * b
            if op == '/':
                return a / b
            if op == '^':
                return np.power(a, b)
        if t == 'fn':
            return self._fn(nodo[1], nodo[2])
        raise ValueError(nodo)

    def _valores(self, args, saltar_vacios=True):
        """Aplana argumentos (rangos y escalares) a una lista de valores."""
        out = []
        for a in args:
            v = self._ev(a)
            if isinstance(v, Rango):
                for k in v.plano():
                    x = self._valor(k)
                    if x is None and saltar_vacios:
                        continue
                    if isinstance(x, str):
                        continue
                    out.append(x)
            else:
                out.append(v)
        return out

    def _fn(self, f, args):
        if f == 'IF':
            c = _bool(self._ev(args[0]))
            si = args[1] if len(args) > 1 else ('num', 1.0)
            no = args[2] if len(args) > 2 else ('num', 0.0)
            if not isinstance(c, np.ndarray):
                return self._ev(si) if c else self._ev(no)
            if c.all():
                return self._ev(si)
            if not c.any():
                return self._ev(no)
            a, b = self._ev(si), self._ev(no)
            if isinstance(a, str) or isinstance(b, str):
                return np.where(c, np.asarray(a, dtype=object), np.asarray(b, dtype=object))
            return np.where(c, _num(a), _num(b))
        if f == 'IFERROR':
            with np.errstate(all='ignore'):
                try:
                    a = self._ev(args[0])
                except (IndexError, ValueError, ZeroDivisionError):
                    return self._ev(args[1])
            if isinstance(a, np.ndarray) and a.dtype != object:
                malo = ~np.isfinite(a)
                if malo.any():
                    b = self._ev(args[1])
                    return np.where(malo, _num(b), a)
                return a
            if isinstance(a, float) and not np.isfinite(a):
                return self._ev(args[1])
            return a
        if f == 'AND':
            vals = [_bool(v) for v in self._valores(args)]
            r = vals[0]
            for v in vals[1:]:
                r = np.logical_and(r, v)
            return r
        if f == 'SUM':
            return sum(_num(v) for v in self._valores(args))
        if f in ('MIN', 'MAX'):
            vals = [_num(v) for v in self._valores(args)]
            r = vals[0]
            for v in vals[1:]:
                r = np.minimum(r, v) if f == 'MIN' else np.maximum(r, v)
            return r
        if f == 'AVERAGE':
            vals = [_num(v) for v in self._valores(args)]
            return sum(vals) / len(vals)
        if f == 'ROUNDUP':
            return _roundup(_num(self._ev(args[0])), int(_num(self._ev(args[1]))))
        if f == 'INDEX':
            rng = self._ev(args[0])
            cl = rng.claves()
            i = _num(self._ev(args[1]))
            j = _num(self._ev(args[2])) if len(args) > 2 else None
            if len(cl) == 1 and j is None:      # vector fila
                lista = cl[0]
            elif len(cl[0]) == 1 and j is None:  # vector columna
                lista = [f_[0] for f_ in cl]
            else:
                raise ValueError('INDEX 2D no soportado')
            if not isinstance(i, np.ndarray):
                return self._valor(lista[int(i) - 1])
            idx = i.astype(int) - 1
            out = np.full(self.n, np.nan)
            for u in np.unique(idx):
                if 0 <= u < len(lista):
                    out = np.where(idx == u, _num(self._valor(lista[u])), out)
            return out
        if f in ('SUMIF', 'COUNTIF'):
            rng = self._ev(args[0]).plano()
            crit = self._ev(args[1])
            sumr = self._ev(args[2]).plano() if (f == 'SUMIF' and len(args) > 2) else rng
            tot = 0.0
            for kc, ks in zip(rng, sumr):
                ok = _cumple(self._valor(kc), crit)
                if f == 'COUNTIF':
                    tot = tot + np.where(ok, 1.0, 0.0) if isinstance(ok, np.ndarray) else tot + float(ok)
                else:
                    tot = tot + np.where(ok, _num(self._valor(ks)), 0.0) if isinstance(ok, np.ndarray) \
                        else tot + (_num(self._valor(ks)) if ok else 0.0)
            return tot
        if f == 'SUMIFS':
            sumr = self._ev(args[0]).plano()
            pares = [(self._ev(args[i]).plano(), self._ev(args[i + 1])) for i in range(1, len(args), 2)]
            tot = 0.0
            for idx, ks in enumerate(sumr):
                ok = True
                for rr, cr in pares:
                    ok = np.logical_and(ok, _cumple(self._valor(rr[idx]), cr))
                tot = tot + np.where(ok, _num(self._valor(ks)), 0.0)
            return tot
        if f == 'NPV':
            r = _num(self._ev(args[0]))
            vals = [_num(v) for v in self._valores(args[1:], saltar_vacios=False)]
            return sum(v / (1 + r) ** (t + 1) for t, v in enumerate(vals))
        if f == 'IRR':
            flujos = np.array([np.broadcast_to(_num(v), (self.n,)) for v in self._valores(args[:1], saltar_vacios=False)])
            return tir(flujos)
        if f == 'MATCH':
            x = _num(self._ev(args[0]))
            lista = [_num(self._valor(k)) for k in self._ev(args[1]).plano()]
            out = np.full(self.n, np.nan)
            for pos in range(len(lista) - 1, -1, -1):
                out = np.where(np.asarray(lista[pos]) == x, pos + 1, out)
            return out
        raise ValueError('Función no soportada: ' + f)

    def evaluar(self, cambios, n, objetivos):
        """cambios: {codigo de parámetro: valor o arreglo de n}. Devuelve {celda: arreglo}."""
        self.n = n
        self.V = {}
        for cod, val in cambios.items():
            k = self.codigos[cod]
            if k in self.formulas:
                raise ValueError(cod + ' es una fórmula, no un parámetro editable')
            self.V[k] = np.broadcast_to(np.asarray(val, dtype=float), (n,)).copy()
        orden = self.orden(objetivos)
        with np.errstate(all='ignore'):
            for k in orden:
                if k in self.V and k not in self.formulas:
                    continue
                v = self._ev(self._arbol(k))
                if isinstance(v, (int, float, bool, np.bool_)):
                    v = float(v)
                self.V[k] = v
        return {k: self.V.get(k, self.constantes.get(k)) for k in objetivos}


def tir(flujos, guess=0.02, it=200):
    """TIR mensual de cada columna de flujos (meses x iteraciones), como IRR de Excel."""
    n_t = flujos.shape[0]
    t = np.arange(n_t)[:, None]
    r = np.full(flujos.shape[1], guess)
    for _ in range(it):
        d = (1 + r) ** t
        f = (flujos / d).sum(0)
        df = (-t * flujos / (d * (1 + r))).sum(0)
        paso = f / df
        r = np.clip(r - paso, -0.99, 10)
        if np.all(np.abs(paso) < 1e-12):
            break
    f = (flujos / (1 + r) ** t).sum(0)
    # Donde Newton no converge (flujos con TIR negativa o baja), se busca la raíz por bisección
    # entre -20% y 100% mensual. Solo queda sin TIR la iteración que no cambia de signo en ese tramo.
    malo = np.abs(f) > 1e-3
    if malo.any():
        F = flujos[:, malo]
        van = lambda x: (F / (1 + x) ** t).sum(0)
        lo, hi = np.full(F.shape[1], -0.2), np.full(F.shape[1], 1.0)
        flo, fhi = van(lo), van(hi)
        ok = np.sign(flo) != np.sign(fhi)
        for _ in range(200):
            mid = (lo + hi) / 2
            fm = van(mid)
            s = np.sign(fm) == np.sign(flo)
            lo, flo, hi = np.where(s, mid, lo), np.where(s, fm, flo), np.where(s, hi, mid)
        r[malo] = np.where(ok, lo, np.nan)
    return r


def celda(ref):
    """'Hoja!A1' -> clave."""
    h, a = ref.split('!')
    m = re.match(r'([A-Z]+)(\d+)', a)
    return (h, int(m.group(2)), column_index_from_string(m.group(1)))


# ==========================================================================
# 2. SIMULACIÓN
# ==========================================================================
NOMBRES = {'GATE': 'Gate fee (USD/t)', 'DIAS': 'Días operativos al mes', 'P_VID': 'Precio fibra de vidrio (USD/t)',
           'PCAR': 'Fracción de carbono', 'VIDA': 'Vida útil de los parques (años)', 'MVAL': 'Meses de validación',
           'C_CAM': 'Costo por camión (USD)'}


def simular(m):
    P = lambda cod: float(m.constantes[m.codigos[cod]])
    g = np.random.default_rng(SEMILLA)
    eur = P('GATE') / 650                      # el gate fee base son 650 EUR/t
    VARS = {  # código: (distribución, mínimo, moda, máximo, de dónde sale)
        'GATE':  ('Triangular', 400 * eur, P('GATE'), 900 * eur, 'Pronk (2022): 400 / 650 / 900 EUR/t. Fuente extranjera'),
        'DIAS':  ('Triangular', 22, P('DIAS'), 30, 'Lunes a viernes / calendario de la fuente del CAPEX / máximo físico'),
        'P_VID': ('Triangular', 290, P('P_VID'), 833, 'Fibra reciclada a granel (Yousef, 2026) / fibra virgen (ChemAnalyst, 2026)'),
        'PCAR':  ('Triangular', 0, P('PCAR'), 0.15, 'Pasaporte de la V90 e inventario. Máximo supuesto del autor'),
        'VIDA':  ('Uniforme, años enteros', 20, None, 27, 'Diseño IEC 61400-1 / promedio de las RCA (Barra y Cisternas, 2022)'),
        'MVAL':  ('Triangular, meses enteros', 12, 12, 24, 'Plazo de un FONDEF IDeA I+D (ANID, 2026)'),
        'C_CAM': ('Triangular', P('C_CAM'), P('C_CAM'), 1446, 'Carga parcial de retorno / camión vacío de regreso'),
    }
    X = {}
    for cod, (dist, a, mo, b, _) in VARS.items():
        if dist.startswith('Uniforme'):
            X[cod] = g.integers(int(a), int(b) + 1, N).astype(float)
        else:
            x = g.triangular(a, mo, b, N) if b > a else np.full(N, a)
            X[cod] = np.round(x) if 'enteros' in dist else x
    OUT = {'VAN': 'Evaluación!C7', 'TIR': 'Evaluación!C8', 'MINCAJA': 'Evaluación!C10', 'EBITDA_2032': 'Resumen anual!I21',
           'LEAS_A1': 'Parámetros!C185', 'CAPEX3': 'Parámetros!C93'}
    objs = {k: celda(v) for k, v in OUT.items()}
    r = m.evaluar(dict(X, SW_OFE=1.0), N, list(objs.values()))
    R = {k: np.broadcast_to(np.asarray(r[c], float), (N,)).copy() for k, c in objs.items()}
    R['CAPITAL'] = -R.pop('MINCAJA')
    R['AMPLIACION'] = np.where(R['CAPEX3'] == 0, 'Sin ampliación', np.where(R['LEAS_A1'] == 1, 'Leasing', 'Ronda'))
    # caso base y umbral del gate fee, con la misma réplica
    esc = {}
    for nom, ch in [('Caso base', {}), ('Gate fee de VAN cero', {'GATE': GATE_VAN_CERO})]:
        rr = m.evaluar(dict(ch, SW_OFE=1.0), 1, [objs['VAN'], objs['TIR'], objs['MINCAJA'], objs['EBITDA_2032']])
        v = [float(np.asarray(x).ravel()[0]) for x in rr.values()]
        esc[nom] = {'GATE': ch.get('GATE', P('GATE')), 'VAN': v[0], 'TIR': v[1], 'CAPITAL': -v[2], 'EBITDA_2032': v[3]}
    return VARS, X, R, esc


# ==========================================================================
# 3. GRÁFICOS
# ==========================================================================
AZUL, NAR, TINTA, GRIS, GRILLA = '#2a78d6', '#eb6834', '#0b0b0b', '#52514e', '#e4e3df'


def cl(x, dec=0):
    """Número en formato chileno: punto de miles y coma decimal."""
    return f'{x:,.{dec}f}'.replace(',', '§').replace('.', ',').replace('§', '.')


def graficos(X, R, esc, carpeta):
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter, PercentFormatter
    plt.rcParams.update({'font.size': 10, 'axes.edgecolor': GRIS, 'axes.labelcolor': TINTA, 'xtick.color': GRIS,
                         'ytick.color': GRIS, 'axes.spines.top': False, 'axes.spines.right': False,
                         'axes.titleweight': 'bold', 'axes.titlesize': 12, 'figure.dpi': 200})
    VAN, CAP, E32, VIDA = R['VAN'], R['CAPITAL'], R['EBITDA_2032'], X['VIDA']
    base = esc['Caso base']
    P = lambda x, q: np.percentile(x, q)
    mill = lambda x, _=None: cl(x / 1e6, 1)
    pct = lambda x, dec=1: cl(100 * x, dec) + '%'
    rutas = []

    def nuevo(alto=4.4):
        fig, ax = plt.subplots(figsize=(8, alto))
        ax.grid(axis='y', color=GRILLA, lw=0.8)
        ax.set_axisbelow(True)
        return fig, ax

    def guardar(fig, nombre):
        fig.tight_layout()
        ruta = os.path.join(carpeta, nombre)
        fig.savefig(ruta, facecolor='white')
        plt.close(fig)
        rutas.append(ruta)

    # 1. Distribución del VAN
    fig, ax = nuevo()
    n, b = np.histogram(VAN, np.arange(-2.0e6, 3.1e6, 1e5))
    c = (b[:-1] + b[1:]) / 2
    ax.bar(c[b[:-1] < 0], n[b[:-1] < 0], 1e5 * 0.92, color=NAR, label='VAN negativo (%s de las iteraciones)' % pct((VAN < 0).mean()))
    ax.bar(c[b[:-1] >= 0], n[b[:-1] >= 0], 1e5 * 0.92, color=AZUL, label='VAN positivo')
    ax.set_ylim(0, n.max() * 1.28)
    for x, txt, ls in [(P(VAN, 50), 'Mediana: %s M' % cl(P(VAN, 50) / 1e6, 2), '-'),
                       (base['VAN'], 'Caso base: %s M' % cl(base['VAN'] / 1e6, 2), '--')]:
        ax.axvline(x, color=TINTA, lw=1.2, ls=ls)
        ax.text(x + 4e4, n.max() * 1.24, txt, va='top', fontsize=9, color=TINTA, bbox=dict(fc='white', ec='none', pad=1))
    ax.xaxis.set_major_formatter(FuncFormatter(mill))
    ax.set_xlabel('VAN al 10% real (millones de USD)')
    ax.set_ylabel('Iteraciones')
    ax.set_title('Distribución del VAN en %s iteraciones' % cl(N), loc='left')
    ax.legend(frameon=False, loc='upper left', fontsize=9, bbox_to_anchor=(0, 0.80))
    guardar(fig, 'fig1_distribucion_VAN.png')

    # 2. Probabilidad acumulada del VAN
    fig, ax = nuevo()
    ax.grid(axis='x', color=GRILLA, lw=0.8)
    xs = np.sort(VAN)
    ax.plot(xs, np.arange(1, N + 1) / N, color=AZUL, lw=2)
    ax.axvline(0, color=GRIS, lw=1)
    p0 = (VAN < 0).mean()
    ax.plot([0], [p0], 'o', ms=8, color=NAR, mec='white', mew=2, zorder=5)
    ax.annotate('Probabilidad de VAN negativo: %s' % pct(p0), (0, p0), (2e5, p0 - 0.12), fontsize=9, color=TINTA,
                arrowprops=dict(arrowstyle='-', color=GRIS, lw=0.8))
    for q in (10, 50, 90):
        x = P(VAN, q)
        ax.plot([x], [q / 100], 'o', ms=8, color=AZUL, mec='white', mew=2, zorder=5)
        ax.text(x - 6e4, q / 100 + 0.03, 'P%d: %s M' % (q, cl(x / 1e6, 2)), ha='right', fontsize=9, color=TINTA)
    ax.xaxis.set_major_formatter(FuncFormatter(mill))
    ax.yaxis.set_major_formatter(PercentFormatter(1, decimals=0))
    ax.set_xlabel('VAN al 10% real (millones de USD)')
    ax.set_ylabel('Probabilidad de que el VAN sea menor')
    ax.set_title('Probabilidad acumulada del VAN', loc='left')
    ax.set_ylim(0, 1.02)
    guardar(fig, 'fig2_curva_acumulada_VAN.png')

    # 3. Qué variable mueve el VAN
    r = {k: np.corrcoef(X[k], VAN)[0, 1] for k in X}
    orden = sorted(r, key=lambda k: abs(r[k]))
    fig, ax = plt.subplots(figsize=(8, 4.2))
    ax.grid(axis='x', color=GRILLA, lw=0.8)
    ax.set_axisbelow(True)
    etiquetas = [NOMBRES[k].split(' (')[0] for k in orden]
    ax.barh(etiquetas, [r[k] for k in orden], height=0.62,
            color=[GRIS if abs(r[k]) < 0.05 else (AZUL if r[k] >= 0 else NAR) for k in orden])
    for i, k in enumerate(orden):
        v = r[k]
        txt = cl(abs(v) if abs(v) < 0.005 else v, 2) + (' (casi nula)' if abs(v) < 0.05 else '')
        ax.text(v + (0.015 if v >= 0 else -0.015), i, txt, va='center', ha='left' if v >= 0 else 'right', fontsize=9, color=TINTA)
    ax.axvline(0, color=GRIS, lw=1)
    ax.set_xlim(-0.6, 0.8)
    ax.xaxis.set_major_formatter(FuncFormatter(lambda x, _: cl(x, 1)))
    ax.set_xlabel('Correlación con el VAN (azul: sube el VAN, naranjo: lo baja)')
    ax.set_title('Qué variable mueve el VAN', loc='left')
    ax.spines['left'].set_visible(False)
    ax.tick_params(axis='y', length=0)
    guardar(fig, 'fig3_sensibilidad_VAN.png')

    # 4. VAN negativo según la vida útil de los parques
    fig, ax = nuevo(4.0)
    anios = np.arange(20, 28)
    pr = [(VAN[VIDA == a] < 0).mean() for a in anios]
    ax.bar(anios, pr, 0.62, color=AZUL)
    for a, p in zip(anios, pr):
        ax.text(a, p + 0.01, pct(p, 0), ha='center', fontsize=9, color=TINTA)
    ax.set_xticks(anios)
    ax.yaxis.set_major_formatter(PercentFormatter(1, decimals=0))
    ax.set_ylim(0, max(pr) * 1.2)
    ax.set_xlabel('Vida útil de los parques eólicos (años)')
    ax.set_ylabel('Probabilidad de VAN negativo')
    ax.set_title('Mientras más duran los parques, más tarde llega la ola de palas', loc='left')
    guardar(fig, 'fig4_VAN_negativo_por_vida_util.png')

    # 5. Capital a levantar frente al ask
    fig, ax = nuevo()
    b = np.arange(3.0e5, 2.6e6, 5e4)
    n, b = np.histogram(np.clip(CAP, 3.0e5, 2.55e6), b)
    ax.bar((b[:-1] + b[1:]) / 2, n, 5e4 * 0.92, color=[NAR if x > ASK else AZUL for x in b[:-1]])
    for x, txt, ls in [(ASK, 'Ask con reserva\n%s M' % cl(ASK / 1e6, 2), '-'),
                       (P(CAP, 90), 'P90\n%s M' % cl(P(CAP, 90) / 1e6, 2), '--')]:
        ax.axvline(x, color=TINTA, lw=1.2, ls=ls)
        ax.text(x + 2.5e4, n.max() * 0.97, txt, va='top', fontsize=9, color=TINTA)
    ax.text(0.98, 0.62, 'Necesitan más que el ask:\n%s de las iteraciones' % pct((CAP > ASK).mean()),
            transform=ax.transAxes, ha='right', fontsize=9, color=TINTA)
    ax.xaxis.set_major_formatter(FuncFormatter(mill))
    ax.set_xlabel('Capital a levantar: máxima necesidad de caja (millones de USD, la última barra junta 2,5 o más)')
    ax.set_ylabel('Iteraciones')
    ax.set_title('Capital a levantar frente al ask con reserva', loc='left')
    guardar(fig, 'fig5_capital_vs_ask.png')

    # 6. EBITDA de 2032
    fig, ax = nuevo(4.2)
    n, b = np.histogram(E32, np.arange(-1e5, 1.15e6, 5e4))
    ax.bar((b[:-1] + b[1:]) / 2, n, 5e4 * 0.92, color=AZUL)
    ax.set_ylim(0, n.max() * 1.3)
    for q, ls, ha in [(10, ':', 'right'), (50, '-', 'right'), (90, ':', 'left')]:
        x = P(E32, q)
        ax.axvline(x, color=TINTA, lw=1.1, ls=ls)
        ax.text(x + (1.2e4 if ha == 'left' else -1.2e4), n.max() * 1.26, 'P%d: %s M' % (q, cl(x / 1e6, 2)),
                va='top', ha=ha, fontsize=9, color=TINTA, bbox=dict(fc='white', ec='none', pad=1))
    ax.axvline(base['EBITDA_2032'], color=GRIS, lw=1.1, ls='--')
    ax.text(base['EBITDA_2032'] + 1.2e4, n.max() * 1.15, 'Caso base: %s M' % cl(base['EBITDA_2032'] / 1e6, 2),
            va='top', fontsize=9, color=GRIS, bbox=dict(fc='white', ec='none', pad=1))
    ax.xaxis.set_major_formatter(FuncFormatter(lambda x, _: cl(x / 1e6, 2)))
    ax.set_xlabel('EBITDA de 2032, año de salida del capital de riesgo (millones de USD)')
    ax.set_ylabel('Iteraciones')
    ax.set_title('EBITDA de 2032 que usa la valorización', loc='left')
    guardar(fig, 'fig6_EBITDA_2032.png')
    return rutas


# ==========================================================================
# 4. EXCEL DE RESULTADOS (dos hojas: Resultados e Iteraciones)
# ==========================================================================
def excel(VARS, X, R, esc, figs, ruta):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.drawing.image import Image as XLImage
    F = 'Arial'
    f_t, f_n, f_b = Font(name=F, bold=True, size=14, color='1F3864'), Font(name=F, size=10), Font(name=F, size=10, bold=True)
    f_h, f_in, f_nt = Font(name=F, bold=True, color='FFFFFF'), Font(name=F, size=10, color='0000FF'), Font(name=F, size=9, italic=True, color='595959')
    f_kpi = Font(name=F, bold=True, size=16, color='1F3864')
    fill_h, fill_k = PatternFill('solid', fgColor='1F3864'), PatternFill('solid', fgColor='EEF3FA')
    s = Side(style='thin', color='BFBFBF')
    box = Border(top=s, bottom=s, left=s, right=s)
    USD, PCT = '#,##0;(#,##0);-', '0.0%'
    wb = openpyxl.Workbook()
    wb.calculation.fullCalcOnLoad = True

    # --- Iteraciones (los datos de la simulación) ---
    it = wb.active
    it.title = 'Iteraciones'
    cols = [('Iteración', None, '0')] + [(NOMBRES[k], X[k], PCT if k == 'PCAR' else ('0.0' if k == 'DIAS' else ('0' if k in ('VIDA', 'MVAL') else USD))) for k in VARS] + \
           [('VAN (USD)', R['VAN'], USD), ('TIR anual', R['TIR'], PCT), ('Capital a levantar (USD)', R['CAPITAL'], USD),
            ('EBITDA 2032 (USD)', R['EBITDA_2032'], USD), ('Primera ampliación', R['AMPLIACION'], None)]
    for j, (lab, _, _) in enumerate(cols, 1):
        c = it.cell(1, j, lab)
        c.font, c.fill, c.alignment = f_h, fill_h, Alignment(wrap_text=True, vertical='center')
        it.column_dimensions[get_column_letter(j)].width = 15
    it.row_dimensions[1].height = 40
    for i in range(N):
        for j, (lab, arr, fm) in enumerate(cols, 1):
            v = i + 1 if arr is None else arr[i]
            if isinstance(v, np.generic):
                v = v.item()
            if isinstance(v, float) and not np.isfinite(v):
                v = 'n/d'
            c = it.cell(i + 2, j, v)
            c.font = f_n
            if fm:
                c.number_format = fm
    it.freeze_panes = 'B2'
    L = {lab: get_column_letter(j) for j, (lab, _, _) in enumerate(cols, 1)}
    rg = lambda lab: "Iteraciones!$%s$2:$%s$%d" % (L[lab], L[lab], N + 1)

    # --- Resultados ---
    re_ = wb.create_sheet('Resultados', 0)
    re_.sheet_view.showGridLines = False
    re_['B2'] = 'Simulación de Monte Carlo del modelo financiero'
    re_['B2'].font = f_t
    re_['B3'] = ('%s iteraciones, semilla %d, sobre el modelo congelado. Todo en USD. Siete variables sorteadas a la vez, '
                 'independientes, con la oferta de palas activada. Las cifras se calculan con fórmulas sobre la hoja Iteraciones.' % (cl(N), SEMILLA))
    re_['B3'].font = f_nt
    # cifras clave
    kpis = [('Probabilidad de VAN negativo', '=COUNTIF(%s,"<0")/COUNT(%s)' % (rg('VAN (USD)'), rg('VAN (USD)')), PCT),
            ('VAN mediano (USD)', '=MEDIAN(%s)' % rg('VAN (USD)'), USD),
            ('VAN del caso base (USD)', esc['Caso base']['VAN'], USD),
            ('Gate fee de VAN cero (USD/t)', esc['Gate fee de VAN cero']['GATE'], USD)]
    for i, (lab, v, fm) in enumerate(kpis):
        col = 2 + 2 * i
        c = re_.cell(5, col, lab)
        c.font, c.fill = f_b, fill_k
        c.alignment = Alignment(wrap_text=True, vertical='top')
        c = re_.cell(6, col, v)
        c.font, c.number_format, c.fill = f_kpi, fm, fill_k
        re_.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col + 1)
        re_.merge_cells(start_row=6, start_column=col, end_row=6, end_column=col + 1)
    re_.row_dimensions[5].height = 30
    re_.row_dimensions[6].height = 26
    # tabla de percentiles
    fila = 8
    for j, t in enumerate(['Indicador', 'Caso base', 'P10', 'Mediana', 'P90'], 2):
        c = re_.cell(fila, j, t)
        c.font, c.fill, c.border = f_h, fill_h, box
    base = esc['Caso base']
    for i, (lab, col_it, b, fm) in enumerate([('VAN', 'VAN (USD)', base['VAN'], USD),
                                              ('Capital a levantar', 'Capital a levantar (USD)', base['CAPITAL'], USD),
                                              ('EBITDA 2032', 'EBITDA 2032 (USD)', base['EBITDA_2032'], USD),
                                              ('TIR', 'TIR anual', base['TIR'], PCT)]):
        r = fila + 1 + i
        vals = [lab, b] + ['=PERCENTILE(%s,%s)' % (rg(col_it), q) for q in ('0.1', '0.5', '0.9')]
        for j, v in enumerate(vals, 2):
            c = re_.cell(r, j, v)
            c.border = box
            c.font = f_b if j == 2 else (f_in if j == 3 else f_n)
            if j > 2:
                c.number_format = fm
    # probabilidades
    fila = 14
    for j, t in enumerate(['Qué tan seguido', '% de iteraciones'], 2):
        c = re_.cell(fila, j, t)
        c.font, c.fill, c.border = f_h, fill_h, box
    re_.cell(fila, 6, 'Ask con reserva (USD)').font = f_b
    c = re_.cell(fila, 7, ASK)
    c.font, c.number_format, c.fill = f_in, USD, PatternFill('solid', fgColor='FFFF00')
    amp = "Iteraciones!$%s$2:$%s$%d" % (L['Primera ampliación'], L['Primera ampliación'], N + 1)
    for i, (lab, f) in enumerate([
            ('VAN negativo', '=COUNTIF(%s,"<0")/COUNT(%s)' % (rg('VAN (USD)'), rg('VAN (USD)'))),
            ('Capital mayor que el ask con reserva', '=COUNTIF(%s,">"&$G$14)/COUNT(%s)' % (rg('Capital a levantar (USD)'), rg('Capital a levantar (USD)'))),
            ('La planta no alcanza a ampliarse antes de 2036', '=COUNTIF(%s,"Sin ampliación")/%d' % (amp, N)),
            ('La primera ampliación necesita una ronda', '=COUNTIF(%s,"Ronda")/%d' % (amp, N))]):
        r = fila + 1 + i
        c = re_.cell(r, 2, lab)
        c.font, c.border = f_b, box
        c = re_.cell(r, 3, f)
        c.font, c.number_format, c.border = f_n, PCT, box
    # supuestos, compactos
    fila = 20
    for j, t in enumerate(['Variable sorteada', 'Distribución', 'Mínimo', 'Moda', 'Máximo', 'Correlación con el VAN'], 2):
        c = re_.cell(fila, j, t)
        c.font, c.fill, c.border = f_h, fill_h, box
        c.alignment = Alignment(wrap_text=True)
    re_.row_dimensions[fila].height = 28
    for i, (cod, (dist, a, mo, b, fu)) in enumerate(VARS.items()):
        r = fila + 1 + i
        fm = PCT if cod == 'PCAR' else ('0.0' if cod == 'DIAS' else USD)
        vals = [NOMBRES[cod], dist, a, '-' if mo is None else mo, b,
                '=CORREL(%s,%s)' % (rg(NOMBRES[cod]), rg('VAN (USD)'))]
        for j, v in enumerate(vals, 2):
            c = re_.cell(r, j, v)
            c.border = box
            c.font = f_b if j == 2 else (f_in if 4 <= j <= 6 else f_n)
            if 4 <= j <= 6 and v != '-':
                c.number_format = fm
            if j == 7:
                c.number_format = '0.00'
    r = fila + len(VARS) + 1
    re_.cell(r, 2, 'Fuentes de los rangos y supuestos en la metodología de la memoria. El gate fee se sortea en el rango europeo: '
                   'como no hay precio chileno, se informa además el gate fee de VAN cero.').font = f_nt
    for col, w in zip('ABCDEFGHIJ', [2, 44, 22, 14, 14, 14, 16, 14, 14, 14]):
        re_.column_dimensions[col].width = w
    # gráficos dentro de la hoja
    fila_img = r + 2
    for k, ruta_fig in enumerate(figs[:4]):
        img = XLImage(ruta_fig)
        img.width, img.height = 640, 352
        re_.add_image(img, 'B%d' % (fila_img + 19 * (k // 2)) if k % 2 == 0 else 'F%d' % (fila_img + 19 * (k // 2)))
    wb.save(ruta)


# ==========================================================================
# 5. EJECUCIÓN
# ==========================================================================
if __name__ == '__main__':
    import sys, time
    t0 = time.time()
    ruta_modelo = sys.argv[1] if len(sys.argv) > 1 else MODELO
    carpeta = os.path.join(os.path.dirname(os.path.abspath(ruta_modelo)), 'resultados_montecarlo')
    os.makedirs(carpeta, exist_ok=True)
    print('Leyendo el modelo:', ruta_modelo)
    m = Modelo(ruta_modelo)
    print('Verificando la réplica contra el Excel...')
    chk = m.evaluar({}, 1, [celda('Evaluación!C7')])[celda('Evaluación!C7')]
    van_excel = m.cache_excel[celda('Evaluación!C7')]
    if van_excel is not None and abs(float(chk) - van_excel) > 1:
        raise SystemExit('La réplica no reproduce el VAN del Excel (%.0f contra %.0f). No se simula.' % (float(chk), van_excel))
    print('  VAN del Excel %s = VAN de la réplica %s' % (cl(van_excel or 0), cl(float(chk))))
    print('Simulando %s iteraciones...' % cl(N))
    VARS, X, R, esc = simular(m)
    print('Haciendo los gráficos...')
    figs = graficos(X, R, esc, carpeta)
    ruta_xlsx = os.path.join(carpeta, 'Resultados Monte Carlo palas eolicas.xlsx')
    print('Escribiendo el Excel de resultados...')
    excel(VARS, X, R, esc, figs, ruta_xlsx)
    van = R['VAN']
    print('\nListo en %.0f segundos. Archivos en %s' % (time.time() - t0, carpeta))
    print('  VAN mediano %s USD, P10 %s, P90 %s' % (cl(np.median(van)), cl(np.percentile(van, 10)), cl(np.percentile(van, 90))))
    print('  Probabilidad de VAN negativo: %s%%' % cl(100 * (van < 0).mean(), 1))
    print('  Capital mayor que el ask: %s%%' % cl(100 * (R['CAPITAL'] > ASK).mean(), 1))
