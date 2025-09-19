import flet as ft
import pandas as pd
from datetime import datetime, date, time, timedelta
import logging, asyncio
from typing import Optional, Tuple, List, Dict

# ============================== Logger ==============================
logging.basicConfig(
    filename="app.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# ============================== Config ==============================
MODO_TV = True                           # otimizações para exibir em TV
SHOW_DROPDOWN = True                     # mostra o dropdown de semanas
CAMINHO_EXCEL = r"C:\Users\Administrador\Downloads\Agendamentos_Rollout_2025_SO.xlsx"

INTERVALO_ATUALIZA_TEMPORIZADOR = 30     # s (re-render para contar tempo restante)
INTERVALO_RELOAD_PLANILHA = 300          # s (recarregar Excel)

# ============================== Paleta ==============================
def paleta():
    return dict(
        BG=ft.Colors.GREY_50, SURFACE=ft.Colors.WHITE, BORDA=ft.Colors.GREY_300,
        TEXTO=ft.Colors.GREY_900, TEXTO_SUAVE=ft.Colors.GREY_700,
        PRIMARIA=ft.Colors.BLUE_600, PRIMARIA_50=ft.Colors.BLUE_50,
        OK=ft.Colors.GREEN_700, OK_BG=ft.Colors.GREEN_50,
        HEADER_DIA_BG=ft.Colors.GREY_100, HEADER_HOJE_BG=ft.Colors.BLUE_50,
        HOJE_BADGE_BG=ft.Colors.BLUE_100
    )

# ===================== Escala TRAVADA (não recalcula) =====================
def init_fixed_scale(page):
    base = page.width or 1920
    return max(0.95, min(1.10, base / 1920))  # faixa curta evita “encolher” com scrollbar

def make_sz(fixed_scale: float):
    def sz(v): return max(10, int(v * fixed_scale))
    return sz

# ============================== Datas / Horas ==============================
def to_date_safe(v) -> Optional[date]:
    try:
        if isinstance(v, (pd.Timestamp, datetime)): return v.date()
        if isinstance(v, date): return v
        if isinstance(v, str):
            for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
                try: return datetime.strptime(v, fmt).date()
                except: pass
            return pd.to_datetime(v, dayfirst=True, errors="coerce").date()
        if pd.isna(v): return None
        return pd.to_datetime(v, errors="coerce").date()
    except Exception as e:
        logging.error(f"to_date_safe: {e}"); return None

def to_time_safe(v) -> Optional[time]:
    try:
        if isinstance(v, time): return v.replace(second=0, microsecond=0)
        if isinstance(v, datetime): return v.time().replace(second=0, microsecond=0)
        if isinstance(v, pd.Timestamp): return v.to_pydatetime().time().replace(second=0, microsecond=0)
        if isinstance(v, (int, float)):
            h = int(v); m = int(round((float(v) - h) * 60)); return time(h, m)
        if isinstance(v, str):
            s = v.strip().replace("h", ":").replace("H", ":")
            if ":" not in s and s.isdigit(): s = f"{s}:00"
            for fmt in ("%H:%M", "%H:%M:%S"):
                try: return datetime.strptime(s, fmt).time().replace(second=0, microsecond=0)
                except: pass
            ts = pd.to_datetime(s, errors="coerce")
            if ts is not pd.NaT: return ts.time().replace(second=0, microsecond=0)
        return None
    except Exception as e:
        logging.error(f"to_time_safe: {e}"); return None

def calcular_temporizador(h: Optional[time], d: Optional[date] = None) -> timedelta:
    try:
        if not h: return timedelta(0)
        d = d or datetime.now().date()
        fim = datetime.combine(d, h) + timedelta(hours=3, minutes=30)
        return max(fim - datetime.now(), timedelta(0))
    except Exception as e:
        logging.error(f"calcular_temporizador: {e}"); return timedelta(0)

def semana_iso_de(d: date) -> Tuple[int, int]:
    iso = d.isocalendar()
    return iso[0], iso[1]

def monday_friday_from_iso(y: int, w: int) -> Tuple[date, date]:
    return date.fromisocalendar(y, w, 1), date.fromisocalendar(y, w, 5)

# ============================== Leitura Excel ==============================
COL_DATA = ["Data","DATA","data","Dia","Dia agendado"]
COL_HORA = ["Hora formatada","Hora","HORA","Agendamento","Horário"]
COL_NOME = ["Nome","NOME","Colaborador","Usuário","Pessoa"]
COL_DIR  = ["Diretoria","DIRETORIA","Dir"]
COL_GER  = ["Gerencia","Gerência","GERENCIA","GERÊNCIA","Ger"]

def achar_col(df: pd.DataFrame, alts: List[str]) -> Optional[str]:
    for c in alts:
        if c in df.columns: return c
    low = {c.lower(): c for c in df.columns}
    for c in alts:
        if c.lower() in low: return low[c.lower()]
    return None

def ler_planilha(path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name="Planilha SO")
        cd = achar_col(df, COL_DATA); ch = achar_col(df, COL_HORA); cn = achar_col(df, COL_NOME)
        if not (cd and cn):
            logging.error(f"Faltam colunas. Data:{cd} Nome:{cn}")
            return pd.DataFrame()
        df["_Data"] = df[cd].apply(to_date_safe)
        df["_Hora"] = df[ch].apply(to_time_safe) if ch else None
        df["_Nome"] = df[cn].astype(str).fillna("Sem nome")
        cd2 = achar_col(df, COL_DIR); cg2 = achar_col(df, COL_GER)
        df["_Diretoria"] = df[cd2].astype(str).fillna("") if cd2 else ""
        df["_Gerencia"]  = df[cg2].astype(str).fillna("") if cg2 else ""
        df = df.dropna(subset=["_Data"]).copy()
        df["_DiaSemana"] = df["_Data"].apply(lambda d: d.weekday())
        df = df[df["_DiaSemana"] <= 4].copy()  # seg..sex
        df[["_AnoISO","_SemanaISO"]] = df["_Data"].apply(lambda d: pd.Series(semana_iso_de(d)))
        df.sort_values(by=["_Data","_Hora","_Nome"], inplace=True, kind="stable")
        df.reset_index(drop=True, inplace=True)
        return df
    except Exception as e:
        logging.error(f"ler_planilha: {e}")
        return pd.DataFrame()

# ============================== App ==============================
PT_DIAS = ["Segunda","Terça","Quarta","Quinta","Sexta"]

def main(page: ft.Page):
    # janela
    page.window_full_screen = MODO_TV
    page.window_maximized = True
    page.theme_mode = ft.ThemeMode.LIGHT
    page.bgcolor = paleta()["BG"]
    page.padding = 16

    # ESCALA TRAVADA
    FIXED_SCALE = init_fixed_scale(page)
    sz = make_sz(FIXED_SCALE)
    P = paleta()

    # estado
    state = {"df": pd.DataFrame(), "y": None, "w": None, "view": "week"}  # "week" ou "day"

    # helpers dados
    def dados_semana(df: pd.DataFrame, y: int, w: int) -> Dict[int, List[Dict]]:
        seg, sex = monday_friday_from_iso(y, w)
        mask = (df["_Data"] >= seg) & (df["_Data"] <= sex)
        sem = df.loc[mask].copy()
        por = {i: [] for i in range(5)}
        for _, r in sem.iterrows():
            i = r["_Data"].weekday()
            if i <= 4: por[i].append(r.to_dict())
        for k in por:
            por[k].sort(key=lambda r: (r.get("_Hora") is None, r.get("_Hora") or time(23,59), r.get("_Nome") or ""))
        return por

    # UI atoms
    def chip(texto: str) -> ft.Container:
        if not str(texto).strip(): return ft.Container()
        return ft.Container(
            content=ft.Row(
                [ft.Icon(ft.Icons.LABEL_OUTLINED, size=sz(16), color=P["TEXTO_SUAVE"]),
                 ft.Text(str(texto), size=sz(16), color=P["TEXTO_SUAVE"])],
                spacing=sz(6), tight=True),
            bgcolor=ft.Colors.GREY_100, border=ft.border.all(1, P["BORDA"]),
            border_radius=999, padding=ft.Padding(sz(10), sz(6), sz(10), sz(6))
        )

    def abrevia_nome(nome: str, tam: int = 2) -> str:
        p = str(nome).split()
        return " ".join(p[:tam]) if p else nome

    def card_compromisso(reg: Dict, hoje: date, booster: float = 1.0) -> ft.Card:
        # booster: 1.0 (normal) | 1.1~1.2 (maior no "Hoje")
        def szz(v): return sz(int(v * booster))
        nome = abrevia_nome(reg.get("_Nome",""))
        d = reg.get("_Data"); h = reg.get("_Hora")
        diret = reg.get("_Diretoria",""); ger = reg.get("_Gerencia","")
        htxt = h.strftime("%H:%M") if isinstance(h, time) else "--:--"
        tempo_txt = str(calcular_temporizador(h, d)).split(".")[0] if (isinstance(d, date) and d == hoje and isinstance(h, time)) else ""

        head = ft.Row(
            [ft.Icon(ft.Icons.PERSON_OUTLINE, size=szz(22), color=P["TEXTO_SUAVE"]),
             ft.Text(nome, size=szz(22), weight=ft.FontWeight.W_600, color=P["TEXTO"]),
             ft.Container(expand=True),
             ft.Row([ft.Icon(ft.Icons.ACCESS_TIME, size=szz(20), color=P["TEXTO_SUAVE"]),
                     ft.Text(htxt, size=szz(20), color=P["TEXTO_SUAVE"])], spacing=szz(6), tight=True)],
            spacing=szz(10), vertical_alignment=ft.CrossAxisAlignment.CENTER
        )

        rodape = ft.Container()
        if tempo_txt:
            rodape = ft.Container(
                bgcolor=P["OK_BG"], border=ft.border.all(1, ft.Colors.GREEN_200),
                border_radius=szz(10), padding=szz(10),
                content=ft.Row(
                    [ft.Icon(ft.Icons.SCHEDULE, size=szz(20), color=P["OK"]),
                     ft.Text(f"Tempo até entrega: {tempo_txt}", size=szz(18), color=P["OK"], weight=ft.FontWeight.W_600)],
                    spacing=szz(8), tight=True)
            )

        return ft.Card(
            elevation=1, surface_tint_color=P["SURFACE"], margin=ft.margin.only(bottom=szz(10)),
            content=ft.Container(
                bgcolor=P["SURFACE"], border=ft.border.all(1, P["BORDA"]),
                border_radius=szz(14), padding=szz(12),
                content=ft.Column(
                    [head, ft.Row([chip(diret), chip(ger)], spacing=szz(8)), rodape],
                    spacing=szz(10), tight=True)
            )
        )

    def coluna_dia(titulo: str, data_dia: date, registros: List[Dict], destaque_hoje: bool, altura_scroll: int) -> ft.Container:
        header_bg = P["HEADER_HOJE_BG"] if destaque_hoje else P["HEADER_DIA_BG"]
        header_text = P["PRIMARIA"] if destaque_hoje else P["TEXTO"]

        cards = [card_compromisso(r, hoje=datetime.now().date()) for r in registros] or [
            ft.Container(
                padding=sz(12),
                content=ft.Row(
                    [ft.Icon(ft.Icons.INBOX, size=sz(20), color=P["TEXTO_SUAVE"]),
                     ft.Text("Sem compromissos", size=sz(18), color=P["TEXTO_SUAVE"])],
                    spacing=sz(8)))
        ]

        header = ft.Container(
            bgcolor=header_bg, border=ft.border.all(1, P["BORDA"]),
            border_radius=sz(12), padding=sz(10),
            content=ft.Row(
                [ft.Text(titulo, size=sz(22), weight=ft.FontWeight.W_700, color=header_text),
                 ft.Container(expand=True),
                 ft.Text(data_dia.strftime("%d/%m"), size=sz(18), color=P["TEXTO_SUAVE"])],
                alignment=ft.MainAxisAlignment.START),
            height=sz(52)
        )

        corpo = ft.Container(
            height=altura_scroll,
            content=ft.Column(controls=cards, spacing=sz(10), scroll=ft.ScrollMode.ALWAYS, expand=True)
        )

        return ft.Container(
            padding=ft.Padding(sz(8), 0, sz(8), 0),
            width=0,
            content=ft.Column(controls=[header, ft.Container(height=sz(10)), corpo], spacing=sz(10), expand=True)
        )

    # ======= UI Topo (agora com alternador Semana/Hoje) =======
    titulo = ft.Text("Planner Semanal", size=sz(48 if MODO_TV else 32),
                     weight=ft.FontWeight.W_700, color=P["PRIMARIA"])
    label_semana_txt = ft.Text("", size=sz(22 if MODO_TV else 16),
                               color=P["TEXTO_SUAVE"], weight=ft.FontWeight.W_500)
    relogio_txt = ft.Text("", size=sz(26 if MODO_TV else 18),
                          color=P["TEXTO"], weight=ft.FontWeight.W_600)

    btn_prev = ft.IconButton(ft.Icons.CHEVRON_LEFT, tooltip="Semana anterior")
    btn_next = ft.IconButton(ft.Icons.CHEVRON_RIGHT, tooltip="Próxima semana")

    date_picker = ft.DatePicker(first_date=date(2020,1,1), last_date=date(2035,12,31))
    btn_pick = ft.IconButton(ft.Icons.EVENT, tooltip="Ir para data", on_click=lambda e: date_picker.pick_date())
    page.overlay.append(date_picker)

    # Alternador de visão
    def make_view_button(texto: str, is_active: bool) -> ft.Container:
        return ft.Container(
            bgcolor=(P["PRIMARIA_50"] if is_active else ft.Colors.GREY_100),
            border=ft.border.all(1, P["BORDA"]),
            border_radius=sz(999),
            padding=ft.Padding(sz(12), sz(6), sz(12), sz(6)),
            content=ft.Text(texto, size=sz(16), weight=ft.FontWeight.W_600,
                            color=(P["PRIMARIA"] if is_active else P["TEXTO_SUAVE"]))
        )

    btn_view_semana = ft.GestureDetector(
        content=make_view_button("Semana", True),
        on_tap=lambda e: set_view("week")
    )
    btn_view_hoje = ft.GestureDetector(
        content=make_view_button("Hoje", False),
        on_tap=lambda e: set_view("day")
    )

    dd_semana = ft.Dropdown(label="Semana (ISO)", width=sz(420), visible=SHOW_DROPDOWN, border_color=P["BORDA"])

    header = ft.Container(
        bgcolor=P["SURFACE"], border=ft.border.all(1, P["BORDA"]),
        border_radius=sz(16), padding=sz(12),
        content=ft.Row(
            [
                ft.Row([ft.Icon(ft.Icons.CALENDAR_MONTH, color=P["PRIMARIA"], size=sz(32)), titulo], spacing=sz(10)),
                ft.Container(expand=True),
                ft.Row([btn_prev, btn_pick, btn_next], spacing=sz(6)),
                ft.Container(width=sz(10)),
                ft.Row([
                    ft.Icon(ft.Icons.SCHEDULE, color=P["TEXTO_SUAVE"], size=sz(22)),
                    relogio_txt
                ], spacing=sz(8)),
            ],
            alignment=ft.MainAxisAlignment.START)
    )

    # Barra com info + controles de visão
    view_switch = ft.Container(
        bgcolor=P["SURFACE"], border=ft.border.all(1, P["BORDA"]),
        border_radius=sz(16), padding=sz(8),
        content=ft.Row(
            [
                ft.Row([btn_view_semana, btn_view_hoje], spacing=sz(8)),
                ft.Container(width=sz(12)),
                ft.Icon(ft.Icons.VIEW_WEEK, color=P["TEXTO_SUAVE"], size=sz(22)),
                label_semana_txt,
                ft.Container(expand=True),
                dd_semana
            ],
            alignment=ft.MainAxisAlignment.START, spacing=sz(12))
    )

    grid_container = ft.Container(expand=True)

    # seleção de semana / view
    def update_view_buttons():
        btn_view_semana.content = make_view_button("Semana", state["view"] == "week")
        btn_view_hoje.content   = make_view_button("Hoje",   state["view"] == "day")

    def set_view(mode: str):
        state["view"] = mode
        update_view_buttons()
        if state["y"] and state["w"]:
            render()

    def set_week(y: int, w: int):
        state["y"], state["w"] = y, w
        if SHOW_DROPDOWN:
            dd_semana.value = f"{y}-{w}"
        render()

    def offset_week(delta: int):
        y, w = state["y"], state["w"]
        if y is None or w is None: return
        monday = date.fromisocalendar(y, w, 1) + timedelta(days=7*delta)
        yy, ww = semana_iso_de(monday)
        set_week(yy, ww)

    btn_prev.on_click = lambda e: offset_week(-1)
    btn_next.on_click = lambda e: offset_week(+1)

    def on_dd_change(e):
        if dd_semana.value:
            y, w = map(int, dd_semana.value.split("-"))
            set_week(y, w)
    dd_semana.on_change = on_dd_change

    def on_date_change(e):
        d = to_date_safe(date_picker.value)
        if d:
            y, w = semana_iso_de(d)
            set_week(y, w)
    date_picker.on_change = on_date_change

    # atalhos de teclado
    def on_key(e: ft.KeyboardEvent):
        if e.key == "Arrow Left": offset_week(-1)
        elif e.key == "Arrow Right": offset_week(+1)
        elif e.key in ("W", "w"): set_view("week")
        elif e.key in ("H", "h"): set_view("day")
    page.on_keyboard_event = on_key

    # ======= Renders =======
    def render_semana(y: int, w: int, df: pd.DataFrame):
        seg, sex = monday_friday_from_iso(y, w)
        label_semana_txt.value = f"{seg.strftime('%d/%m/%Y')} – {sex.strftime('%d/%m/%Y')}  •  Sem {w:02d}/{y}"
        dados = dados_semana(df, y, w)
        hoje = datetime.now().date()

        altura_total = page.height or 1080
        altura_scroll = max(sz(380), int(altura_total - sz(320)))  # reserva header+view bar

        cols = []
        for i, nd in enumerate(PT_DIAS):
            dia = seg + timedelta(days=i)
            cols.append(
                ft.Container(
                    expand=True, width=0,
                    content=coluna_dia(nd, dia, dados.get(i, []), destaque_hoje=(dia == hoje), altura_scroll=altura_scroll)
                )
            )

        grid_container.content = ft.Container(
            padding=sz(12), bgcolor=P["SURFACE"], border=ft.border.all(1, P["BORDA"]), border_radius=sz(16),
            content=ft.Row(cols, spacing=sz(10), alignment=ft.MainAxisAlignment.SPACE_BETWEEN, tight=True)
        )

    def render_hoje(y: int, w: int, df: pd.DataFrame):
        # Mostra apenas o dia atual (se não estiver na semana escolhida, ainda usa a semana escolhida, mas foca na data "hoje")
        hoje = datetime.now().date()
        yy, ww = semana_iso_de(hoje)
        seg, sex = monday_friday_from_iso(y, w)  # mantém barra de semana escolhida
        label_semana_txt.value = f"Hoje: {hoje.strftime('%A, %d/%m/%Y').title()}  •  Sem {w:02d}/{y}"
        dados = dados_semana(df, y, w)
        idx = min(max((hoje - seg).days, 0), 4)  # 0..4 dentro da semana corrente
        dia_data = seg + timedelta(days=idx)
        registros = dados.get(idx, [])

        altura_total = page.height or 1080
        altura_scroll = max(sz(500), int(altura_total - sz(300)))  # maior para Hoje

        # Cabeçalho do HOJE com badge
        badge_hoje = ft.Container(
            bgcolor=P["HOJE_BADGE_BG"], border_radius=sz(999),
            padding=ft.Padding(sz(10), sz(6), sz(10), sz(6)),
            content=ft.Row(
                [ft.Icon(ft.Icons.TODAY, size=sz(18), color=P["PRIMARIA"]),
                 ft.Text("Destaque do Dia", size=sz(16), weight=ft.FontWeight.W_600, color=P["PRIMARIA"])],
                spacing=sz(6), tight=True
            )
        )

        cards = [card_compromisso(r, hoje=hoje, booster=1.12) for r in registros] or [
            ft.Container(
                padding=sz(12),
                content=ft.Row(
                    [ft.Icon(ft.Icons.INBOX, size=sz(22), color=P["TEXTO_SUAVE"]),
                     ft.Text("Sem compromissos para hoje.", size=sz(20), color=P["TEXTO_SUAVE"])],
                    spacing=sz(10)))
        ]

        coluna_unica = ft.Container(
            padding=sz(12),
            content=ft.Column(
                [
                    ft.Row([
                        ft.Text(dia_data.strftime("%A").title(), size=sz(30), weight=ft.FontWeight.W_800, color=P["PRIMARIA"]),
                        ft.Container(width=sz(10)),
                        badge_hoje,
                        ft.Container(expand=True),
                        ft.Text(dia_data.strftime("%d/%m/%Y"), size=sz(22), color=P["TEXTO_SUAVE"])
                    ], spacing=sz(8)),
                    ft.Container(height=sz(12)),
                    ft.Container(
                        height=altura_scroll,
                        content=ft.Column(controls=cards, spacing=sz(12), scroll=ft.ScrollMode.ALWAYS, expand=True)
                    )
                ],
                spacing=sz(8),
                expand=True
            )
        )

        grid_container.content = ft.Container(
            padding=sz(12), bgcolor=P["SURFACE"], border=ft.border.all(1, P["BORDA"]), border_radius=sz(16),
            content=coluna_unica
        )

    def render():
        if state["view"] == "day":
            render_hoje(state["y"], state["w"], state["df"])
        else:
            render_semana(state["y"], state["w"], state["df"])
        page.update()

    # carga inicial
    def carregar_e_montar():
        df = ler_planilha(CAMINHO_EXCEL)
        state["df"] = df
        if df.empty:
            grid_container.content = ft.Container(
                padding=sz(20), bgcolor=ft.Colors.RED_50, border=ft.border.all(1, ft.Colors.RED_400), border_radius=sz(12),
                content=ft.Row(
                    [ft.Icon(ft.Icons.ERROR, color=ft.Colors.RED_700, size=sz(26)),
                     ft.Text("Erro ao carregar planilha/aba 'Planilha SO'.", size=sz(22), color=ft.Colors.RED_700)],
                    spacing=sz(10)))
            page.update(); return

        semanas = df[["_AnoISO","_SemanaISO"]].drop_duplicates().sort_values(by=["_AnoISO","_SemanaISO"])
        opcoes = [(int(a), int(s)) for a, s in semanas.itertuples(index=False, name=None)]
        y0, w0 = semana_iso_de(datetime.now().date())
        y, w = (y0, w0) if (y0, w0) in opcoes else (opcoes[-1] if opcoes else (y0, w0))

        if SHOW_DROPDOWN:
            dd_semana.options = [
                ft.dropdown.Option(
                    key=f"{a}-{s}",
                    text=f"{date.fromisocalendar(a,s,1).strftime('%d/%m')} a {date.fromisocalendar(a,s,5).strftime('%d/%m')} • {a} (Sem {s:02d})"
                ) for a, s in opcoes
            ]
            dd_semana.value = f"{y}-{w}"

        set_week(y, w)
        update_view_buttons()

    # tarefas periódicas
    async def tick_relogio():
        while True:
            relogio_txt.value = datetime.now().strftime("%a, %d/%m • %H:%M:%S")
            page.update(); await asyncio.sleep(1)

    async def tick_temporizador():
        while True:
            if not state["df"].empty and state["y"] and state["w"]:
                render()
            await asyncio.sleep(INTERVALO_ATUALIZA_TEMPORIZADOR)

    async def tick_reload_planilha():
        while True:
            await asyncio.sleep(INTERVALO_RELOAD_PLANILHA)
            carregar_e_montar()

    # resize (sem recalcular escala!)
    def on_resize(e):
        if state["y"] and state["w"] and not state["df"].empty:
            render()
    page.on_resized = on_resize

    # layout
    page.add(
        header,
        ft.Container(height=sz(10)),
        view_switch,
        ft.Container(height=sz(10)),
        grid_container
    )

    # start
    carregar_e_montar()
    page.run_task(tick_relogio)
    page.run_task(tick_temporizador)
    page.run_task(tick_reload_planilha)

if __name__ == "__main__":
    ft.app(target=main)
