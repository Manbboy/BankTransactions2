"""Transactions — application Kivy 100 % hors ligne (Android + bureau pour les tests)."""
import datetime as dt
import json
import math
import os
import shutil
import threading
import traceback

from kivy.app import App
from kivy.base import ExceptionHandler, ExceptionManager
from kivy.clock import Clock
from kivy.core.text import Label as CoreLabel
from kivy.core.window import Window
from kivy.factory import Factory
from kivy.graphics import Color, Line, Rectangle
from kivy.lang import Builder
from kivy.metrics import dp, sp
from kivy.properties import BooleanProperty, ListProperty, NumericProperty, ObjectProperty, StringProperty
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.uix.textinput import TextInput
from kivy.uix.togglebutton import ToggleButton
from kivy.uix.widget import Widget
from kivy.utils import escape_markup, platform

import logic

MONO = "data/fonts/RobotoMono-Regular.ttf"
MAX_DISPLAY_ROWS = 20000
MAX_PICKER_ROWS = 3000
MAX_BARS = 60
PICK_REQUEST = 4242
NORMAL_BG = (0.16, 0.17, 0.20, 1)
ERROR_BG = (0.55, 0.15, 0.15, 1)
GREEN, RED = "7fd18b", "ff7b7b"
HINTS = {"text": "mot  a|b  !x", "num": ">50  10..20", "date": "2026-03  >=2026-01"}
QUICK = "Période rapide"
NO_LEVEL = "—"
PERIOD_LEVEL = "Période"
# Colonnes affichées par défaut (celles demandées) ; les autres se (ré)affichent via « Colonnes ».
DEFAULT_VISIBLE = [
    "Date d'exécution", "Date valeur", "Montant", "Numéro de compte", "Type de transaction",
    "Contrepartie", "Communication", "Détails", "Origine", "Année", "Période",
    "Nom de la contrepartie NEW", "Clé unique", "Catégorie", "Sous Catégorie",
]
AGG_COLS = (24, 6, 14, 14, 14)   # largeurs (caractères) : groupe, Nb, Net, Entrées, Sorties

KV = """
#:import dp kivy.metrics.dp
#:import sp kivy.metrics.sp
#:set MONO 'data/fonts/RobotoMono-Regular.ttf'
#:set GOLD (1, 0.8, 0.3, 1)

<Row>:
    font_name: MONO
    font_size: sp(13)
    markup: True
    text_size: self.size
    halign: 'left'
    valign: 'middle'
    canvas.before:
        Color:
            rgba: (0.17, 0.18, 0.21, 1) if self.odd else (0.12, 0.13, 0.15, 1)
        Rectangle:
            pos: self.pos
            size: self.size

<ChoiceRow>:
    background_normal: ''
    background_down: ''
    background_color: (0.15, 0.45, 0.85, 1) if self.selected else (0.17, 0.18, 0.21, 1)
    halign: 'left'
    valign: 'middle'
    text_size: self.width - dp(16), self.height
    shorten: True
    shorten_from: 'right'
    font_size: sp(14)
    on_release: app.picker_toggle(self.val)

<FilterBtn>:
    background_normal: ''
    background_down: ''
    background_color: (0.35, 0.24, 0.05, 1) if self.active else (0.22, 0.23, 0.27, 1)
    canvas.after:
        Color:
            rgba: (1, 0.8, 0.3, 1) if self.active else (0.8, 0.8, 0.85, 1)
        Triangle:
            points: [self.center_x - dp(5), self.center_y + dp(3), self.center_x + dp(5), self.center_y + dp(3), self.center_x, self.center_y - dp(4)]

<PickerBox@BoxLayout>:
    orientation: 'vertical'
    spacing: dp(6)
    padding: dp(6)
    TextInput:
        id: search
        hint_text: 'Rechercher...'
        multiline: False
        size_hint_y: None
        height: dp(40)
    Label:
        id: info
        size_hint_y: None
        height: dp(22)
        font_size: sp(12)
    RecycleView:
        id: rv
        viewclass: 'ChoiceRow'
        bar_width: dp(8)
        scroll_type: ['bars', 'content']
        RecycleBoxLayout:
            default_size: None, dp(40)
            default_size_hint: 1, None
            size_hint_y: None
            height: self.minimum_height
            orientation: 'vertical'
            spacing: dp(1)
    BoxLayout:
        size_hint_y: None
        height: dp(42)
        spacing: dp(6)
        Button:
            text: 'Cocher les affichés'
            on_release: app.picker_check_all()
        Button:
            text: 'Vider'
            on_release: app.picker_clear()
    BoxLayout:
        size_hint_y: None
        height: dp(44)
        spacing: dp(6)
        Button:
            text: 'Annuler'
            on_release: app.picker_close(False)
        Button:
            text: 'Appliquer'
            on_release: app.picker_close(True)

<ColumnsBox@BoxLayout>:
    orientation: 'vertical'
    spacing: dp(6)
    padding: dp(6)
    ScrollView:
        GridLayout:
            id: grid
            cols: 1
            size_hint_y: None
            height: self.minimum_height
            spacing: dp(2)
    BoxLayout:
        size_hint_y: None
        height: dp(44)
        spacing: dp(6)
        Button:
            text: 'Tout'
            on_release: app.columns_all(True)
        Button:
            text: 'Aucune'
            on_release: app.columns_all(False)
        Button:
            text: 'Annuler'
            on_release: app.columns_close(False)
        Button:
            text: 'OK'
            on_release: app.columns_close(True)

BoxLayout:
    orientation: 'vertical'
    canvas.before:
        Color:
            rgba: 0.08, 0.08, 0.10, 1
        Rectangle:
            pos: self.pos
            size: self.size
    BoxLayout:
        size_hint_y: None
        height: dp(42)
        padding: dp(4), dp(3)
        spacing: dp(4)
        Button:
            text: 'Ouvrir Excel'
            on_release: app.open_file()
        Button:
            text: 'Colonnes'
            on_release: app.open_columns()
        Button:
            text: 'Effacer filtres'
            on_release: app.clear_filters()
    BoxLayout:
        size_hint_y: None
        height: dp(38)
        padding: dp(4), dp(2)
        spacing: dp(4)
        Spinner:
            id: sp_sheet
            text: ''
            on_text: app.on_sheet_change(self.text)
        Spinner:
            id: sp_quick
            text: 'Période rapide'
            values: ['Ce mois', 'Mois dernier', 'Cette année', 'Tout']
            on_text: app.quick_period(self.text)
    Label:
        text: app.status
        size_hint_y: None
        height: dp(22)
        text_size: self.size
        halign: 'left'
        valign: 'middle'
        shorten: True
        font_size: sp(12)
        padding: dp(6), 0
    TabbedPanel:
        id: tabs
        do_default_tab: False
        tab_width: dp(110)
        tab_height: dp(36)
        TabbedPanelItem:
            id: tab_tx
            text: 'Transactions'
            BoxLayout:
                orientation: 'vertical'
                ScrollView:
                    do_scroll_x: True
                    do_scroll_y: False
                    bar_width: dp(6)
                    BoxLayout:
                        id: table
                        orientation: 'vertical'
                        size_hint_x: None
                        width: dp(600)
                        BoxLayout:
                            id: hdr
                            size_hint_y: None
                            height: dp(26)
                        BoxLayout:
                            id: flt
                            size_hint_y: None
                            height: dp(40)
                        RecycleView:
                            id: rv
                            viewclass: 'Row'
                            bar_width: dp(8)
                            scroll_type: ['bars', 'content']
                            RecycleBoxLayout:
                                default_size: None, dp(30)
                                default_size_hint: 1, None
                                size_hint_y: None
                                height: self.minimum_height
                                orientation: 'vertical'
                BoxLayout:
                    size_hint_y: None
                    height: dp(34)
                    padding: dp(4), dp(2)
                    spacing: dp(4)
                    Label:
                        text: app.count_text
                        text_size: self.size
                        halign: 'left'
                        valign: 'middle'
                        font_size: sp(12)
                    Button:
                        text: app.sort_text
                        size_hint_x: None
                        width: dp(130)
                        font_size: sp(12)
                        on_release: app.toggle_sort()
        TabbedPanelItem:
            id: tab_tot
            text: 'Totaux'
            BoxLayout:
                orientation: 'vertical'
                padding: dp(4)
                spacing: dp(4)
                BoxLayout:
                    size_hint_y: None
                    height: dp(34)
                    spacing: dp(4)
                    Button:
                        text: 'Réglages'
                        on_release: app.ctrl_hidden = not app.ctrl_hidden
                    Button:
                        text: 'Tout déplier'
                        on_release: app.expand_all(True)
                    Button:
                        text: 'Tout replier'
                        on_release: app.expand_all(False)
                GridLayout:
                    id: agg_ctrl
                    cols: 4
                    size_hint_y: None
                    height: 0 if app.ctrl_hidden else self.minimum_height
                    opacity: 0 if app.ctrl_hidden else 1
                    disabled: app.ctrl_hidden
                    row_default_height: dp(34)
                    row_force_default: True
                    spacing: dp(4)
                    Label:
                        text: 'Date'
                        size_hint_x: 0.24
                        font_size: sp(12)
                    Spinner:
                        id: sp_date
                        on_text: app.recompute()
                    Label:
                        text: 'Montant'
                        size_hint_x: 0.24
                        font_size: sp(12)
                    Spinner:
                        id: sp_amount
                        on_text: app.recompute()
                    Label:
                        text: 'Période'
                        size_hint_x: 0.24
                        font_size: sp(12)
                    Spinner:
                        id: sp_period
                        text: app.period
                        values: ['Jour', 'Semaine', 'Mois', 'Année']
                        on_text: app.period = self.text
                    Label:
                        text: 'Niv. 1'
                        size_hint_x: 0.24
                        font_size: sp(12)
                    Spinner:
                        id: sp_l1
                        on_text: app.on_levels()
                    Label:
                        text: 'Niv. 2'
                        size_hint_x: 0.24
                        font_size: sp(12)
                    Spinner:
                        id: sp_l2
                        on_text: app.on_levels()
                    Label:
                        text: 'Niv. 3'
                        size_hint_x: 0.24
                        font_size: sp(12)
                    Spinner:
                        id: sp_l3
                        on_text: app.on_levels()
                Label:
                    text: app.agg_info
                    size_hint_y: None
                    height: dp(34)
                    text_size: self.width, None
                    halign: 'left'
                    valign: 'middle'
                    font_size: sp(12)
                ScrollView:
                    do_scroll_x: True
                    do_scroll_y: False
                    bar_width: dp(6)
                    BoxLayout:
                        id: agg_table
                        orientation: 'vertical'
                        size_hint_x: None
                        width: dp(560)
                        Label:
                            text: app.agg_header
                            font_name: MONO
                            font_size: sp(13)
                            color: GOLD
                            size_hint_y: None
                            height: dp(26)
                            text_size: self.size
                            halign: 'left'
                            valign: 'middle'
                        RecycleView:
                            id: rv_agg
                            viewclass: 'Row'
                            bar_width: dp(8)
                            scroll_type: ['bars', 'content']
                            RecycleBoxLayout:
                                default_size: None, dp(32)
                                default_size_hint: 1, None
                                size_hint_y: None
                                height: self.minimum_height
                                orientation: 'vertical'
        TabbedPanelItem:
            id: tab_chart
            text: 'Graphique'
            BoxLayout:
                orientation: 'vertical'
                padding: dp(4)
                spacing: dp(4)
                BoxLayout:
                    size_hint_y: None
                    height: dp(36)
                    spacing: dp(4)
                    Label:
                        text: 'Période'
                        size_hint_x: 0.3
                    Spinner:
                        text: app.period
                        values: ['Jour', 'Semaine', 'Mois', 'Année']
                        on_text: app.period = self.text
                Label:
                    text: app.chart_info
                    size_hint_y: None
                    height: dp(40)
                    text_size: self.width, None
                    halign: 'left'
                    valign: 'middle'
                    font_size: sp(12)
                BarChart:
                    id: chart
"""


# --------------------------------------------------------------------- widgets
class Row(Label):
    odd = BooleanProperty(False)
    idx = NumericProperty(-1)
    kind = StringProperty("tx")
    path = ObjectProperty(())

    def on_touch_down(self, touch):
        if getattr(touch, "is_mouse_scrolling", False) or not self.collide_point(*touch.pos):
            return super().on_touch_down(touch)
        app = App.get_running_app()
        if self.kind == "tx" and self.idx >= 0:
            app.show_detail(self.idx)
            return True
        if self.kind == "agg":
            app.toggle_node(self.path)
            return True
        return super().on_touch_down(touch)


class ChoiceRow(Button):
    selected = BooleanProperty(False)
    val = StringProperty("")


class FilterBtn(Button):
    active = BooleanProperty(False)


class BarChart(Widget):
    """Histogramme : entrées (vert, vers le haut) et sorties (rouge, vers le bas) par période."""
    data = ListProperty([])
    sel = NumericProperty(-1)

    def __init__(self, **kw):
        super().__init__(**kw)
        self.bind(pos=self.redraw, size=self.redraw, data=self.redraw, sel=self.redraw)

    @staticmethod
    def _tex(txt, size=10):
        lab = CoreLabel(text=txt, font_size=sp(size), color=(0.85, 0.85, 0.9, 1))
        lab.refresh()
        return lab.texture

    def _geom(self):
        left, right, top, bottom = dp(8), dp(8), dp(24), dp(28)
        w = max(self.width - left - right, 1)
        return self.x + left, w, self.y + bottom, max(self.height - bottom - top, 1)

    def redraw(self, *_):
        self.canvas.clear()
        with self.canvas:
            Color(0.1, 0.1, 0.12, 1)
            Rectangle(pos=self.pos, size=self.size)
        data = self.data
        if not data:
            t = self._tex("Aucune donnée (vérifiez date, montant et filtres)", 12)
            with self.canvas:
                Rectangle(texture=t, size=t.size, pos=(self.x + dp(10), self.center_y))
            return
        x0, w, yb, h = self._geom()
        maxp = max(max(p for _, p, _n in data), 0.0)
        maxn = max(max(-n for _, _p, n in data), 0.0)
        tot = (maxp + maxn) or 1.0
        y0 = yb + h * maxn / tot
        scale = h / tot
        slot = w / len(data)
        bw = max(1.0, slot * 0.7)
        with self.canvas:
            if 0 <= self.sel < len(data):
                Color(1, 1, 1, 0.10)
                Rectangle(pos=(x0 + self.sel * slot, yb), size=(slot, h))
            for i, (_k, p, n) in enumerate(data):
                bx = x0 + i * slot + (slot - bw) / 2
                if p > 0:
                    Color(0.30, 0.75, 0.40, 1)
                    Rectangle(pos=(bx, y0), size=(bw, p * scale))
                if n < 0:
                    Color(0.90, 0.35, 0.35, 1)
                    Rectangle(pos=(bx, y0 + n * scale), size=(bw, -n * scale))
            Color(0.6, 0.6, 0.65, 1)
            Line(points=[x0, y0, x0 + w, y0], width=1)
            top_t = self._tex(f"max entrées {logic.fmt_money(maxp)}   max sorties {logic.fmt_money(-maxn)}", 10)
            Color(1, 1, 1, 1)
            Rectangle(texture=top_t, size=top_t.size, pos=(x0, self.top - dp(20)))
            first = self._tex(data[0][0], 10)
            step = max(1, int(math.ceil((first.width * 1.3) / slot)))
            for i in range(0, len(data), step):
                t = self._tex(data[i][0], 10)
                cx = x0 + i * slot + slot / 2
                px = min(max(cx - t.width / 2, self.x + dp(2)), self.right - t.width - dp(2))
                Rectangle(texture=t, size=t.size, pos=(px, self.y + dp(6)))

    def on_touch_down(self, touch):
        if self.data and self.collide_point(*touch.pos):
            x0, w, _yb, _h = self._geom()
            slot = w / len(self.data)
            i = min(max(int((touch.x - x0) / slot), 0), len(self.data) - 1)
            self.sel = i
            k, p, n = self.data[i]
            App.get_running_app().chart_info = (
                f"{k}   Entrées {logic.fmt_money(p)}   Sorties {logic.fmt_money(n)}   "
                f"Net {logic.fmt_money(p + n)}"
            )
            return True
        return super().on_touch_down(touch)


class ErrorCatcher(ExceptionHandler):
    """Affiche l'erreur à l'écran (et dans error.log) au lieu de fermer l'appli."""

    def handle_exception(self, inst):
        try:
            tb = "".join(traceback.format_exception(type(inst), inst, inst.__traceback__))
            app = App.get_running_app()
            if app is not None:
                app.report_error(tb)
        except Exception:  # noqa: BLE001
            pass
        return ExceptionManager.PASS


# ------------------------------------------------------------------------- app
class TransactionsApp(App):
    status = StringProperty("Aucun fichier ouvert")
    count_text = StringProperty("")
    agg_info = StringProperty("")
    agg_header = StringProperty("")
    chart_info = StringProperty("Touchez une barre pour voir le détail")
    period = StringProperty("Mois")
    sort_text = StringProperty("Récents en haut")
    ctrl_hidden = BooleanProperty(False)

    # ------------------------------------------------------------ démarrage
    def build(self):
        self.title = "Transactions"
        self.headers, self.data, self.kinds, self.money, self.strs = [], [], [], [], []
        self.visible, self.exprs, self.sels = [], [], []
        self.filtered = []
        self.inputs, self.fbtns, self.col_chars = {}, {}, {}
        self.path, self.file_name, self.sheet = None, "", None
        self._setup = False
        self._ev = None
        self._err_shown = 0
        self._pk = None
        self._cols_pop = None
        self.sort_desc = True
        self.levels = []
        self.agg_root = None
        self.expanded = set()
        self.all_open = False
        self.cfg = self._read_cfg()
        ExceptionManager.add_handler(ErrorCatcher())
        Window.softinput_mode = "below_target"

        self.root_widget = Builder.load_string(KV)
        probe = CoreLabel(text="M" * 50, font_name=MONO, font_size=sp(13))
        probe.refresh()
        self.charw = probe.texture.width / 50.0
        self.root_widget.ids.agg_table.width = max(sum(AGG_COLS) * self.charw + dp(6), dp(360))
        if platform == "android":
            from android import activity  # type: ignore
            activity.bind(on_activity_result=self._on_activity_result)
        return self.root_widget

    def on_start(self):
        last = os.path.join(self.user_data_dir, "last.xlsx")
        if os.path.exists(last):
            try:
                with open(os.path.join(self.user_data_dir, "last_name.txt"), encoding="utf-8") as f:
                    name = f.read().strip()
            except OSError:
                name = "dernier fichier"
            self.load_path(last, name)

    def on_pause(self):
        return True  # indispensable : le sélecteur de fichiers met l'app en pause

    # ----------------------------------------------------------- réglages
    def _cfg_file(self):
        return os.path.join(self.user_data_dir, "settings.json")

    def _read_cfg(self):
        try:
            with open(self._cfg_file(), encoding="utf-8") as f:
                return json.load(f)
        except (OSError, ValueError):
            return {}

    def _save_cfg(self):
        try:
            with open(self._cfg_file(), "w", encoding="utf-8") as f:
                json.dump(self.cfg, f, ensure_ascii=False)
        except OSError:
            pass

    # -------------------------------------------------------------- erreurs
    def report_error(self, tb):
        try:
            with open(os.path.join(self.user_data_dir, "error.log"), "a", encoding="utf-8") as f:
                f.write(f"--- {dt.datetime.now():%Y-%m-%d %H:%M:%S}\n{tb}\n")
        except OSError:
            pass
        self._err_shown += 1
        if self._err_shown > 3:
            return
        short = tb.strip().splitlines()[-1] if tb.strip() else "erreur"
        self.status = f"Erreur : {short}"[:200]
        Clock.schedule_once(lambda _dt: self._error_popup(tb), 0)

    def _error_popup(self, tb):
        lab = Label(text=escape_markup(tb[-1800:]), markup=True, font_size=sp(11), size_hint_y=None,
                    halign="left", valign="top")
        lab.bind(width=lambda i, w: setattr(i, "text_size", (w, None)),
                 texture_size=lambda i, s: setattr(i, "height", s[1]))
        sv = ScrollView()
        sv.add_widget(lab)
        box = BoxLayout(orientation="vertical", spacing=dp(6), padding=dp(6))
        box.add_widget(Label(text="Une erreur est survenue. Faites une capture d'écran et envoyez-la.",
                             size_hint_y=None, height=dp(44), text_size=(Window.width * 0.85, None),
                             font_size=sp(13)))
        box.add_widget(sv)
        b = Button(text="Fermer", size_hint_y=None, height=dp(44))
        box.add_widget(b)
        pop = Popup(title="Erreur", content=box, size_hint=(0.95, 0.85))
        b.bind(on_release=pop.dismiss)
        pop.open()

    # ---------------------------------------------------- choix du fichier
    def open_file(self):
        if platform == "android":
            self._android_pick()
        else:
            self._desktop_pick()

    def _android_pick(self):
        try:
            from jnius import autoclass  # type: ignore
            Intent = autoclass("android.content.Intent")
            PythonActivity = autoclass("org.kivy.android.PythonActivity")
            intent = Intent(Intent.ACTION_OPEN_DOCUMENT)
            intent.addCategory(Intent.CATEGORY_OPENABLE)
            intent.setType("*/*")
            PythonActivity.mActivity.startActivityForResult(intent, PICK_REQUEST)
        except Exception as e:  # noqa: BLE001
            self.status = f"Sélecteur indisponible : {e}"

    def _on_activity_result(self, request_code, result_code, intent):
        if request_code != PICK_REQUEST or result_code != -1 or intent is None:  # -1 = RESULT_OK
            return
        try:
            from jnius import autoclass  # type: ignore
            PythonActivity = autoclass("org.kivy.android.PythonActivity")
            OpenableColumns = autoclass("android.provider.OpenableColumns")
            uri = intent.getData()
            resolver = PythonActivity.mActivity.getContentResolver()
            name = "fichier.xlsx"
            cursor = resolver.query(uri, None, None, None, None)
            if cursor is not None:
                try:
                    if cursor.moveToFirst():
                        i = cursor.getColumnIndex(OpenableColumns.DISPLAY_NAME)
                        if i >= 0:
                            name = cursor.getString(i) or name
                finally:
                    cursor.close()
            # Copie dans le stockage privé de l'app : fonctionne hors ligne, sans permission.
            pfd = resolver.openFileDescriptor(uri, "r")
            fd = pfd.detachFd()
            dest = os.path.join(self.user_data_dir, "last.xlsx")
            with os.fdopen(fd, "rb") as src, open(dest, "wb") as out:
                shutil.copyfileobj(src, out)
            with open(os.path.join(self.user_data_dir, "last_name.txt"), "w", encoding="utf-8") as f:
                f.write(name)
        except Exception as e:  # noqa: BLE001
            msg = f"Copie impossible : {e}"
            Clock.schedule_once(lambda _dt: setattr(self, "status", msg))
            return
        Clock.schedule_once(lambda _dt: self.load_path(dest, name))

    def _desktop_pick(self):
        chooser = FileChooserListView(path=os.path.expanduser("~"), filters=["*.xlsx", "*.xlsm"])
        buttons = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(6))
        ok, cancel = Button(text="Ouvrir"), Button(text="Annuler")
        buttons.add_widget(ok)
        buttons.add_widget(cancel)
        box = BoxLayout(orientation="vertical")
        box.add_widget(chooser)
        box.add_widget(buttons)
        popup = Popup(title="Choisir un fichier Excel", content=box, size_hint=(0.95, 0.9))

        def _ok(*_):
            if chooser.selection:
                path = chooser.selection[0]
                popup.dismiss()
                self.load_path(path, os.path.basename(path))

        ok.bind(on_release=_ok)
        cancel.bind(on_release=popup.dismiss)
        popup.open()

    # -------------------------------------------------------------- lecture
    def load_path(self, path, name, sheet=None):
        self.status = f"Chargement de {name}…"

        def work():
            try:
                names = logic.sheet_names(path)
                wanted = logic.pick_sheet(names, sheet, self.cfg.get("sheet"))
                headers, data = logic.load_xlsx(path, wanted)
                kinds = logic.detect_kinds(headers, data)
                money = logic.money_flags(headers, kinds)
                strs = [[logic.fmt(c, m) for c, m in zip(r, money)] for r in data]
            except Exception as e:  # noqa: BLE001
                msg = f"Erreur de lecture : {e}"
                Clock.schedule_once(lambda _dt: setattr(self, "status", msg))
                return
            Clock.schedule_once(
                lambda _dt: self._on_loaded(path, name, names, wanted, headers, data, kinds, money, strs))

        threading.Thread(target=work, daemon=True).start()

    def on_sheet_change(self, text):
        if self._setup or not self.path or not text or text == self.sheet:
            return
        self.load_path(self.path, self.file_name, text)

    def _on_loaded(self, path, name, names, sheet, headers, data, kinds, money, strs):
        self.path, self.file_name, self.sheet = path, name, sheet
        self.cfg["sheet"] = sheet
        self._save_cfg()
        ids = self.root_widget.ids
        self._setup = True
        ids.sp_sheet.values = names
        ids.sp_sheet.text = sheet
        self._setup = False
        self.headers, self.data, self.kinds, self.money, self.strs = headers, data, kinds, money, strs
        self.exprs = [""] * len(headers)
        self.sels = [None] * len(headers)
        self.expanded, self.all_open = set(), False
        if not headers:
            self.filtered = []
            ids.hdr.clear_widgets()
            ids.flt.clear_widgets()
            ids.rv.data = []
            ids.rv_agg.data = []
            ids.chart.data = []
            self.count_text = self.agg_info = self.agg_header = ""
            self.status = f"Feuille « {sheet} » vide : choisissez une autre feuille"
            return
        saved = self.cfg.get("visible") or []
        self.visible = [i for i, h in enumerate(headers) if h in saved] \
            or [i for i, h in enumerate(headers) if h in DEFAULT_VISIBLE] \
            or list(range(len(headers)))
        self.status = f"{name} · {sheet} · {len(data)} lignes, {len(headers)} colonnes"
        self._setup_spinners()
        self._build_table()
        self.refilter()

    # ------------------------------------------------------------- tableau
    def _build_table(self):
        ids = self.root_widget.ids
        ids.hdr.clear_widgets()
        ids.flt.clear_widgets()
        self.inputs, self.fbtns, self.col_chars = {}, {}, {}
        total = 0.0
        for c in self.visible:
            h = self.headers[c]
            longest = max((len(s[c]) for s in self.strs[:400]), default=0)
            n = max(12, min(24, max(len(h), longest)))
            self.col_chars[c] = n
            w = (n + 3) * self.charw
            total += w
            lab = Label(text=h[:n], font_name=MONO, font_size=sp(13), color=(1, 0.8, 0.3, 1),
                        size_hint=(None, 1), width=w, halign="left", valign="middle")
            lab.bind(size=lambda inst, s: setattr(inst, "text_size", s))
            ids.hdr.add_widget(lab)

            cell = BoxLayout(size_hint=(None, 1), width=w, padding=(dp(1), dp(2)), spacing=dp(2))
            ti = TextInput(multiline=False, text=self.exprs[c], hint_text=HINTS[self.kinds[c]],
                           font_size=sp(13), background_normal="", background_active="",
                           background_color=NORMAL_BG, foreground_color=(1, 1, 1, 1),
                           hint_text_color=(0.5, 0.5, 0.55, 1), write_tab=False, padding=(dp(5), dp(9)))
            ti.bind(text=lambda inst, val, col=c: self._on_expr(col, val))
            fb = FilterBtn(size_hint=(None, 1), width=dp(28), active=bool(self.sels[c]))
            fb.bind(on_release=lambda inst, col=c: self.open_picker(col))
            cell.add_widget(ti)
            cell.add_widget(fb)
            ids.flt.add_widget(cell)
            self.inputs[c], self.fbtns[c] = ti, fb
        ids.table.width = max(total + dp(4), Window.width)

    def _format_row(self, i):
        sr = self.strs[i]
        parts = []
        for c in self.visible:
            n = self.col_chars[c]
            s = sr[c].replace("\n", " ")[:n]
            parts.append(s.rjust(n) if self.kinds[c] == "num" else s.ljust(n))
        return " | ".join(parts) + " | "

    def _on_expr(self, col, val):
        self.exprs[col] = val
        if getattr(self, "_clearing", False):
            return
        if self._ev is not None:
            self._ev.cancel()
        self._ev = Clock.schedule_once(lambda _dt: self.refilter(), 0.35)

    def date_idx(self):
        ids = self.root_widget.ids
        t = ids.sp_date.text
        if t in self.headers:
            return self.headers.index(t)
        for i, k in enumerate(self.kinds):
            if k == "date":
                return i
        return None

    def refilter(self):
        if not self.headers:
            return
        idx, errors = logic.apply_filters(self.data, self.kinds, self.exprs, self.sels, self.strs)
        for c, ti in self.inputs.items():
            ti.background_color = ERROR_BG if c in errors else NORMAL_BG
        for c, fb in self.fbtns.items():
            fb.active = bool(self.sels[c])
        di = self.date_idx()
        if di is not None:
            idx.sort(key=lambda i: logic.to_date(self.data[i][di]) or dt.date.min)
            if self.sort_desc:
                idx.reverse()
        self.filtered = idx
        shown = idx[:MAX_DISPLAY_ROWS]
        self.root_widget.ids.rv.data = [
            {"text": escape_markup(self._format_row(i)), "odd": bool(k % 2), "idx": i, "kind": "tx",
             "path": (), "color": (1, 1, 1, 1)}
            for k, i in enumerate(shown)
        ]
        txt = f"{len(idx)} / {len(self.data)} lignes"
        if len(idx) > len(shown):
            txt += f" (affichage limité à {MAX_DISPLAY_ROWS})"
        if errors:
            txt += " · filtre invalide : " + ", ".join(self.headers[c] for c in errors)
        self.count_text = txt
        self.recompute()

    def clear_filters(self):
        self._clearing = True        # évite un refiltrage par champ vidé
        try:
            self.exprs = [""] * len(self.headers)
            self.sels = [None] * len(self.headers)
            for ti in self.inputs.values():
                ti.text = ""
        finally:
            self._clearing = False
        self.refilter()

    def toggle_sort(self):
        self.sort_desc = not self.sort_desc
        self.sort_text = "Récents en haut" if self.sort_desc else "Anciens en haut"
        self.refilter()

    def quick_period(self, name):
        if name == QUICK or not self.headers:
            return
        col = self.date_idx()
        if col is not None:
            self._ensure_visible(col)
            expr = logic.date_shortcut(name)
            self.exprs[col] = expr
            self.sels[col] = None
            if col in self.inputs:
                self.inputs[col].text = expr   # déclenche le refiltrage
            self.refilter()
        self.root_widget.ids.sp_quick.text = QUICK

    def _ensure_visible(self, col):
        if col not in self.visible:
            self.visible = sorted(self.visible + [col])
            self._build_table()

    # -------------------------------------------------------------- détail
    def show_detail(self, i):
        row = self.data[i]
        parts = []
        for h, v, m in zip(self.headers, row, self.money):
            s = logic.fmt(v, m) or "-"
            parts.append(f"[b]{escape_markup(h)}[/b]\n{escape_markup(s)}")
        lab = Label(text="\n\n".join(parts), markup=True, font_size=sp(14), size_hint_y=None,
                    halign="left", valign="top")
        lab.bind(width=lambda inst, w: setattr(inst, "text_size", (w - dp(8), None)),
                 texture_size=lambda inst, s: setattr(inst, "height", s[1] + dp(8)))
        sv = ScrollView()
        sv.add_widget(lab)
        box = BoxLayout(orientation="vertical", spacing=dp(6), padding=dp(6))
        box.add_widget(sv)
        b = Button(text="Fermer", size_hint_y=None, height=dp(44))
        box.add_widget(b)
        pop = Popup(title=f"Transaction (ligne {i + 2} de la feuille)", content=box, size_hint=(0.95, 0.9))
        b.bind(on_release=pop.dismiss)
        pop.open()

    # ------------------------------------------------ colonnes à afficher
    def open_columns(self):
        if not self.headers:
            return
        box = Factory.ColumnsBox()
        for i, h in enumerate(self.headers):
            tb = ToggleButton(text=h, state="down" if i in self.visible else "normal",
                              size_hint_y=None, height=dp(40))
            box.ids.grid.add_widget(tb)
        self._cols_pop = Popup(title="Colonnes affichées", content=box, size_hint=(0.95, 0.9))
        self._cols_box = box
        self._cols_pop.open()

    def _col_buttons(self):
        # add_widget insère en tête : on remet dans l'ordre des colonnes
        return list(reversed(self._cols_box.ids.grid.children))

    def columns_all(self, on):
        for tb in self._col_buttons():
            tb.state = "down" if on else "normal"

    def columns_close(self, ok):
        pop, self._cols_pop = self._cols_pop, None
        if ok:
            chosen = [i for i, tb in enumerate(self._col_buttons()) if tb.state == "down"] or [0]
            for c in range(len(self.headers)):          # un filtre caché serait invisible : on l'efface
                if c not in chosen:
                    self.exprs[c] = ""
                    self.sels[c] = None
            self.visible = chosen
            self.cfg["visible"] = [self.headers[c] for c in chosen]
            self._save_cfg()
            self._build_table()
            self.refilter()
        if pop:
            pop.dismiss()

    # ----------------------------------------- sélection de valeurs (listes)
    def open_picker(self, col):
        idx, _ = logic.apply_filters(self.data, self.kinds, self.exprs, self.sels, self.strs, skip=col)
        vals = logic.distinct_values(idx, self.data, self.strs, col, self.kinds[col])
        box = Factory.PickerBox()
        pop = Popup(title=f"Filtrer : {self.headers[col]}", content=box, size_hint=(0.95, 0.92))
        self._pk = {"col": col, "vals": vals, "chosen": set(self.sels[col] or ()), "box": box,
                    "pop": pop, "shown": []}
        box.ids.search.bind(text=lambda *_: self._pk_refresh())
        self._pk_refresh()
        pop.open()

    def _pk_refresh(self):
        pk = self._pk
        if not pk:
            return
        q = logic.norm(pk["box"].ids.search.text.strip())
        matches = [(s, n) for s, n in pk["vals"] if not q or q in logic.norm(s)]
        pk["shown"] = [s for s, _ in matches]
        view = matches[:MAX_PICKER_ROWS]
        pk["box"].ids.rv.data = [
            {"text": f"{'[x]' if s in pk['chosen'] else '[   ]'}  {s if s else '(vide)'}   ({n})",
             "val": s, "selected": s in pk["chosen"]}
            for s, n in view
        ]
        more = f" (affichage limité à {MAX_PICKER_ROWS}, utilisez la recherche)" if len(matches) > len(view) else ""
        pk["box"].ids.info.text = f"{len(pk['chosen'])} coché(s) · {len(matches)} valeur(s){more}"

    def picker_toggle(self, val):
        pk = self._pk
        if not pk:
            return
        if val in pk["chosen"]:
            pk["chosen"].discard(val)
        else:
            pk["chosen"].add(val)
        self._pk_refresh()

    def picker_check_all(self):
        if self._pk:
            self._pk["chosen"].update(self._pk["shown"])
            self._pk_refresh()

    def picker_clear(self):
        if self._pk:
            self._pk["chosen"].clear()
            self._pk_refresh()

    def picker_close(self, ok):
        pk, self._pk = self._pk, None
        if not pk:
            return
        if ok:
            self.sels[pk["col"]] = set(pk["chosen"]) or None
            self.refilter()
        pk["pop"].dismiss()

    # -------------------------------------------------------------- totaux
    def _setup_spinners(self):
        ids = self.root_widget.ids
        hs, ks = self.headers, self.kinds
        likely_d = [h for h, k in zip(hs, ks) if k == "date"]
        likely_n = [h for h, k in zip(hs, ks) if k == "num"]
        date_cols = likely_d + [h for h in hs if h not in likely_d]
        num_cols = likely_n + [h for h in hs if h not in likely_n]
        pref_d = "Date d'exécution" if "Date d'exécution" in date_cols else date_cols[0]
        pref_n = [h for h in likely_n if logic.norm(h).startswith(logic.NUM_NAMES)]
        levels = [NO_LEVEL, PERIOD_LEVEL] + list(hs)
        self._setup = True
        ids.sp_date.values, ids.sp_amount.values = date_cols, num_cols
        ids.sp_date.text = pref_d
        ids.sp_amount.text = pref_n[0] if pref_n else num_cols[0]
        for sp_ in (ids.sp_l1, ids.sp_l2, ids.sp_l3):
            sp_.values = levels
        ids.sp_l1.text = PERIOD_LEVEL
        ids.sp_l2.text = "Catégorie" if "Catégorie" in hs else NO_LEVEL
        ids.sp_l3.text = "Sous Catégorie" if "Sous Catégorie" in hs else NO_LEVEL
        self._setup = False

    def on_period(self, *_):
        self.expanded, self.all_open = set(), False
        self.recompute()

    def on_levels(self):
        if self._setup:
            return
        self.expanded, self.all_open = set(), False
        self.recompute()

    def expand_all(self, on):
        self.all_open = on
        self.expanded = set()
        self._render_agg()

    def toggle_node(self, path):
        if self.all_open:           # on repart de l'état « tout ouvert » pour replier un seul nœud
            self.all_open = False
            self.expanded = self._all_paths()
        if path in self.expanded:
            self.expanded.discard(path)
        else:
            self.expanded.add(path)
        self._render_agg()

    def _all_paths(self):
        out = set()

        def walk(n):
            for c in n["ch"].values():
                if c["ch"]:
                    out.add(c["path"])
                walk(c)

        if self.agg_root:
            walk(self.agg_root)
        return out

    def recompute(self):
        if not self.headers or self._setup:
            return
        ids = self.root_widget.ids
        if ids.sp_date.text not in self.headers or ids.sp_amount.text not in self.headers:
            return
        d_idx = self.headers.index(ids.sp_date.text)
        a_idx = self.headers.index(ids.sp_amount.text)
        period = logic.PERIODS.get(self.period, "month")
        levels = []
        for sp_ in (ids.sp_l1, ids.sp_l2, ids.sp_l3):
            t = sp_.text
            if t == PERIOD_LEVEL:
                levels.append(("period",))
            elif t in self.headers:
                levels.append(("col", self.headers.index(t)))
        rows = [self.data[i] for i in self.filtered]
        self.levels = levels
        self.agg_root, skipped = logic.build_tree(rows, levels, d_idx, a_idx, period)
        info = (f"{len(rows)} ligne(s) filtrée(s) · Net {logic.fmt_money(self.agg_root['pos'] + self.agg_root['neg'])}"
                f"  (entrées {logic.fmt_money(self.agg_root['pos'])}, sorties {logic.fmt_money(self.agg_root['neg'])})")
        if skipped:
            info += f"\n{skipped} ligne(s) ignorée(s) (date ou montant illisible)"
        self.agg_info = info
        w = AGG_COLS
        self.agg_header = f"{'Groupe':<{w[0]}}{'Nb':>{w[1]}}{'Net':>{w[2]}}{'Entrées':>{w[3]}}{'Sorties':>{w[4]}}"
        self._render_agg()
        series, _ = logic.period_series(rows, d_idx, a_idx, period)
        ids.chart.sel = -1
        ids.chart.data = series[-MAX_BARS:]
        self.chart_info = (f"{len(series)} période(s)" + (f", les {MAX_BARS} dernières affichées" if len(series) > MAX_BARS else "")
                           + " · touchez une barre pour le détail")

    def _render_agg(self):
        if not self.agg_root:
            return
        w = AGG_COLS
        flat = logic.flatten(self.agg_root, self.levels, self.expanded, self.all_open)
        data = []
        for k, r in enumerate(flat):
            if r["depth"] == 0:
                label = "TOTAL"
            else:
                marker = ("- " if r["open"] else "+ ") if r["has_children"] else "  "
                label = ("  " * (r["depth"] - 1) + marker + r["label"])
            label = label[:w[0] - 1].ljust(w[0])
            net = logic.fmt_money(r["net"]).rjust(w[2])
            net_col = GREEN if r["net"] > 0 else RED if r["net"] < 0 else "cccccc"
            text = (escape_markup(label) + str(r["n"]).rjust(w[1])
                    + f"[color={net_col}]{net}[/color]"
                    + logic.fmt_money(r["pos"]).rjust(w[3]) + logic.fmt_money(r["neg"]).rjust(w[4]))
            data.append({"text": text, "odd": bool(k % 2), "idx": -1, "kind": "agg", "path": r["path"],
                         "color": (1, 0.85, 0.4, 1) if r["depth"] == 0 else (1, 1, 1, 1)})
        self.root_widget.ids.rv_agg.data = data


if __name__ == "__main__":
    TransactionsApp().run()
