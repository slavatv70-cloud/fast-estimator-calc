#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Калькулятор Сметчика
Версия 5.4 (редактирование пользовательских материалов прямо в диалоге)
"""

import sys
import math
import random
import re
import ast
import json
import logging
from datetime import datetime
from typing import Dict, List, Tuple, Optional, Any

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget,
    QVBoxLayout, QHBoxLayout, QGridLayout, QLabel,
    QLineEdit, QComboBox, QPushButton, QGroupBox,
    QMessageBox, QScrollArea, QTableWidget,
    QHeaderView, QAbstractItemView, QSizePolicy,
    QRadioButton, QButtonGroup, QFileDialog, QTableWidgetItem,
    QStatusBar, QListWidget, QListWidgetItem, QDialog, QSpinBox,
    QFormLayout, QDialogButtonBox, QStackedWidget, QFrame,
    QInputDialog, QDoubleSpinBox,
    QMenu
)
from PyQt6.QtCore import Qt, QTimer, QRect, QPoint, QSettings, QRectF, QPointF, QLocale, QSize
from PyQt6.QtGui import (
    QFont, QBrush, QPainter, QColor, QPen,
    QDoubleValidator, QIntValidator, QPainterPath, QKeyEvent, QValidator
)

try:
    import openpyxl
except ImportError:
    openpyxl = None

# ---------- Логирование ----------
logger = logging.getLogger("MetalCalculator")
logger.setLevel(logging.INFO)
try:
    _file_handler = logging.FileHandler("calculator.log", encoding="utf-8")
    _file_handler.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(_file_handler)
except Exception:
    _stderr_handler = logging.StreamHandler()
    _stderr_handler.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(_stderr_handler)

# ============================================================================
#  КОНСТАНТЫ
# ============================================================================

STEEL_DENSITY = 7850

BUILTIN_MATERIALS = {
    "Сталь (углеродистая)": 7850,
    "Нержавеющая сталь (аустенитная)": 7930,
    "Алюминий": 2700,
    "Медь": 8900,
    "Бронза": 8800,
    "Чугун": 7200,
    "Латунь": 8500,
}

# Коэффициенты изменения массы крепежа в зависимости от покрытия.
COATING_FACTOR = {
    "Без покрытия": 1.00,
    "Оцинкованный": 1.02,             # +2 % на слой цинка
    "Черный фосфатированный": 1.00,
    "—": 1.00,
}

# Заголовки столбцов 3/4/5 в «Крепеже» зависят от типа метиза.
HEADERS_BY_TYPE = {
    "Болт":          ("Диаметр, мм",          "Длина, мм",           "—"),
    "Винт":          ("Диаметр, мм",          "Длина, мм",           "—"),
    "Гайка":         ("Диаметр, мм",          "—",                   "—"),
    "Гвоздь":        ("Диаметр, мм",          "Длина, мм",           "—"),
    "Дюбели":        ("Диаметр, мм",          "Длина, мм",           "—"),
    "Заклепка":      ("Диаметр, мм",          "Длина, мм",           "—"),
    "Круг отрезной": ("Диаметр диска, мм",    "Толщина диска, мм",   "—"),
    "Перфорация":    ("Ширина/Полка, мм",     "Длина, мм",           "Толщина, мм"),
    "Саморез":       ("Диаметр, мм",          "Длина, мм",           "—"),
}

# ============================================================================
#  ПЕСКОСТРУЙНАЯ ОЧИСТКА
# ============================================================================

SANDBLAST_DEGREES = {
    "Sa 1 (лёгкая)": {
        "abrasive_min": 5, "abrasive_max": 15,
        "time_per_m2": 0.075,
        "description": "Лёгкая очистка: удаление рыхлой ржавчины, окалины, старой краски",
    },
    "Sa 2 (тщательная)": {
        "abrasive_min": 15, "abrasive_max": 30,
        "time_per_m2": 0.125,
        "description": "Тщательная очистка: удаление окалины, ржавчины, остатков покрытий",
    },
    "Sa 2½ (очень тщательная)": {
        "abrasive_min": 30, "abrasive_max": 60,
        "time_per_m2": 0.200,
        "description": "Очень тщательная: до металлического блеска, подготовка под АКЗ",
    },
    "Sa 3 (белая)": {
        "abrasive_min": 60, "abrasive_max": 100,
        "time_per_m2": 0.325,
        "description": "Белая очистка: 100% удаление загрязнений, максимальная адгезия",
    },
}

SANDBLAST_ABRASIVES = {
    "Кварцевый песок": {"k_consumption": 1.00, "k_time": 1.00, "k_reuse": 1.0, "color": "#f5d76e"},
    "Купершлак":       {"k_consumption": 0.85, "k_time": 0.90, "k_reuse": 1.0, "color": "#a0522d"},
    "Электрокорунд":   {"k_consumption": 0.70, "k_time": 0.85, "k_reuse": 3.0, "color": "#6d6875"},
    "Стальная дробь":  {"k_consumption": 0.30, "k_time": 0.80, "k_reuse": 10.0, "color": "#95a5a6"},
}

# ============================================================================
#  СОРТАМЕНТ ГОСТ
# ============================================================================

SORTAMENT_DATA = {
    "Труба ГОСТ 8732-78 (бесшовная горячедеформированная)": {
        "type": "pipe", "params": ["D", "s"],
        "data": [
            (20, 2.5, 1.08), (20, 2.8, 1.19), (20, 3.0, 1.26), (20, 3.5, 1.42),
            (25, 2.5, 1.39), (25, 3.0, 1.63), (25, 3.5, 1.86), (25, 4.0, 2.07),
            (32, 3.0, 2.15), (32, 3.5, 2.46), (32, 4.0, 2.76), (32, 5.0, 3.33),
            (38, 3.0, 2.59), (38, 3.5, 2.98), (38, 4.0, 3.35), (38, 5.0, 4.07),
            (45, 3.0, 3.11), (45, 3.5, 3.58), (45, 4.0, 4.04), (45, 5.0, 4.93),
            (57, 3.0, 4.00), (57, 3.5, 4.62), (57, 4.0, 5.23), (57, 5.0, 6.41),
            (76, 3.5, 6.26), (76, 4.0, 7.10), (76, 5.0, 8.75), (76, 6.0, 10.36),
            (89, 4.0, 8.38), (89, 5.0, 10.36), (89, 6.0, 12.28), (89, 8.0, 15.98),
            (108, 4.0, 10.26), (108, 5.0, 12.70), (108, 6.0, 15.09), (108, 8.0, 19.73),
            (133, 4.0, 12.73), (133, 5.0, 15.78), (133, 6.0, 18.79), (133, 8.0, 24.66),
            (159, 5.0, 18.99), (159, 6.0, 22.64), (159, 8.0, 29.79), (159, 10.0, 36.75),
            (219, 6.0, 31.52), (219, 8.0, 41.63), (219, 10.0, 51.54), (219, 12.0, 61.26),
            (273, 8.0, 52.28), (273, 10.0, 64.86), (273, 12.0, 77.24), (273, 14.0, 89.42),
            (325, 10.0, 77.68), (325, 12.0, 92.63), (325, 14.0, 107.38), (325, 16.0, 121.93),
            (377, 12.0, 108.02), (377, 14.0, 125.36), (377, 16.0, 142.50), (377, 18.0, 159.44),
            (426, 14.0, 142.26), (426, 16.0, 161.78), (426, 18.0, 181.11), (426, 20.0, 200.24),
        ],
    },
    "Труба ГОСТ 10704-91 (электросварная прямошовная)": {
        "type": "pipe", "params": ["D", "s"],
        "data": [
            (20, 1.5, 0.684), (20, 2.0, 0.888), (25, 1.5, 0.869),
            (25, 2.0, 1.134), (32, 2.0, 1.480), (32, 2.5, 1.819),
            (40, 2.0, 1.875), (40, 3.0, 2.737), (50, 2.0, 2.366),
            (50, 3.0, 3.478), (57, 3.0, 3.994), (76, 3.0, 5.401),
            (89, 3.0, 6.363), (108, 3.5, 9.020), (114, 4.0, 10.85),
            (133, 4.0, 12.73), (159, 4.0, 15.29), (219, 5.0, 26.39),
            (273, 6.0, 39.51), (325, 6.0, 47.20), (426, 6.0, 62.15),
        ],
    },
    "Швеллер ГОСТ 8240-97 (с уклоном полок, серия У)": {
        "type": "channel", "params": ["h", "b", "s", "t"],
        "data": [
            (50, 32, 4.4, 7.0, 4.84), (65, 36, 4.4, 7.2, 5.90),
            (80, 40, 4.5, 7.4, 7.05), (100, 46, 4.5, 7.6, 8.59),
            (120, 52, 4.8, 7.8, 10.40), (140, 58, 4.9, 8.1, 12.30),
            (160, 64, 5.0, 8.4, 14.20), (180, 70, 5.1, 8.7, 16.30),
            (200, 76, 5.2, 9.0, 18.40), (220, 82, 5.4, 9.5, 21.00),
            (240, 90, 5.6, 10.0, 24.00), (270, 95, 6.0, 10.5, 27.70),
            (300, 100, 6.5, 11.0, 31.80), (330, 105, 7.0, 11.7, 36.50),
            (360, 110, 7.5, 12.6, 41.90), (400, 115, 8.0, 13.5, 48.30),
        ],
    },
    "Швеллер ГОСТ 8240-97 (параллельные полки, серия П)": {
        "type": "channel", "params": ["h", "b", "s", "t"],
        "data": [
            (50, 32, 4.4, 7.0, 4.84), (65, 36, 4.4, 7.2, 5.90),
            (80, 40, 4.5, 7.4, 7.05), (100, 46, 4.5, 7.6, 8.59),
            (120, 52, 4.8, 7.8, 10.40), (140, 58, 4.9, 8.1, 12.30),
            (160, 64, 5.0, 8.4, 14.20), (180, 70, 5.1, 8.7, 16.30),
            (200, 76, 5.2, 9.0, 18.40), (220, 82, 5.4, 9.5, 21.00),
            (240, 90, 5.6, 10.0, 24.00), (270, 95, 6.0, 10.5, 27.70),
            (300, 100, 6.5, 11.0, 31.80), (330, 105, 7.0, 11.7, 36.50),
            (360, 110, 7.5, 12.6, 41.90), (400, 115, 8.0, 13.5, 48.30),
        ],
    },
    "Двутавр ГОСТ 26020-83 (нормальные, серия Б)": {
        "type": "beam", "params": ["h", "b", "s", "t"],
        "data": [
            (100, 55, 4.1, 5.7, 8.10), (120, 64, 4.4, 6.2, 10.40),
            (140, 73, 4.7, 6.9, 12.90), (160, 81, 5.0, 7.4, 15.80),
            (180, 90, 5.1, 8.0, 18.40), (200, 100, 5.2, 8.4, 21.00),
            (220, 110, 5.4, 8.7, 24.00), (240, 115, 5.6, 9.5, 27.30),
            (270, 125, 6.0, 9.8, 31.50), (300, 135, 6.5, 10.2, 36.50),
            (330, 140, 7.0, 11.0, 42.20), (360, 145, 7.5, 11.8, 48.60),
            (400, 155, 8.3, 13.0, 57.00), (450, 160, 9.0, 14.2, 66.50),
            (500, 170, 10.0, 15.2, 78.50), (550, 180, 11.0, 16.5, 92.60),
            (600, 190, 12.0, 17.8, 108.00),
        ],
    },
    "Двутавр ГОСТ 26020-83 (широкополочные, серия Ш)": {
        "type": "beam", "params": ["h", "b", "s", "t"],
        "data": [
            (193, 150, 6.0, 9.0, 30.60), (226, 155, 6.5, 10.0, 36.20),
            (251, 180, 7.0, 10.0, 42.70), (291, 200, 8.0, 11.0, 53.60),
            (338, 250, 9.5, 12.5, 75.10), (388, 300, 9.5, 14.0, 96.10),
            (484, 300, 11.0, 15.0, 114.40), (580, 320, 12.0, 17.0, 142.10),
        ],
    },
    "Уголок ГОСТ 8509-93 (равнополочный)": {
        "type": "angle", "params": ["A", "t"],
        "data": [
            (20, 3, 0.892), (20, 4, 1.150), (25, 3, 1.120), (25, 4, 1.460),
            (30, 3, 1.360), (30, 4, 1.780), (35, 3, 1.600), (35, 4, 2.090),
            (35, 5, 2.570), (40, 3, 1.840), (40, 4, 2.420), (40, 5, 2.970),
            (45, 4, 2.740), (45, 5, 3.380), (50, 4, 3.060), (50, 5, 3.770),
            (50, 6, 4.470), (60, 5, 4.570), (60, 6, 5.420), (60, 8, 7.090),
            (63, 5, 4.810), (63, 6, 5.710), (63, 8, 7.470), (70, 6, 6.390),
            (70, 8, 8.370), (75, 6, 6.890), (75, 8, 9.030), (80, 6, 7.360),
            (80, 8, 9.660), (90, 6, 8.330), (90, 8, 10.90), (90, 9, 12.20),
            (100, 6, 9.300), (100, 8, 12.20), (100, 10, 15.10),
            (110, 8, 13.50), (125, 8, 15.50), (125, 10, 19.10),
            (140, 10, 21.50), (160, 12, 29.40), (180, 12, 33.70),
            (200, 12, 37.30), (200, 16, 49.50),
        ],
    },
    "Круг ГОСТ 2590-2006 (горячекатаный)": {
        "type": "circle", "params": ["d"],
        "data": [
            (5.0, 0.154), (5.5, 0.187), (6.0, 0.222), (6.3, 0.245),
            (7.0, 0.302), (8.0, 0.395), (9.0, 0.499), (10.0, 0.617),
            (11.0, 0.746), (12.0, 0.888), (13.0, 1.042), (14.0, 1.208),
            (15.0, 1.387), (16.0, 1.578), (17.0, 1.782), (18.0, 1.998),
            (19.0, 2.226), (20.0, 2.466), (21.0, 2.719), (22.0, 2.984),
            (23.0, 3.262), (24.0, 3.551), (25.0, 3.853), (26.0, 4.168),
            (27.0, 4.495), (28.0, 4.834), (29.0, 5.185), (30.0, 5.549),
            (32.0, 6.313), (34.0, 7.127), (36.0, 7.990), (38.0, 8.903),
            (40.0, 9.865), (42.0, 10.876), (45.0, 12.485), (48.0, 14.205),
            (50.0, 15.413), (55.0, 18.650), (60.0, 22.190), (65.0, 26.050),
            (70.0, 30.210), (75.0, 34.680), (80.0, 39.460), (85.0, 44.540),
            (90.0, 49.940), (100.0, 61.650),
        ],
    },
}

# ============================================================================
#  НОРМЫ АКЗ
# ============================================================================

AKZ_STANDARDS = [
    {"name": "Эпоксидная грунтовка", "category": "Грунт",
     "consumption_min": 0.180, "consumption_max": 0.250, "unit": "кг/м²",
     "dry_thickness": "60–80 мкм", "layers": 2, "interval": "8–24 ч"},
    {"name": "Полиуретановая эмаль", "category": "Эмаль",
     "consumption_min": 0.120, "consumption_max": 0.160, "unit": "кг/м²",
     "dry_thickness": "30–50 мкм", "layers": 2, "interval": "12–36 ч"},
    {"name": "Цинконаполненный грунт", "category": "Грунт",
     "consumption_min": 0.200, "consumption_max": 0.300, "unit": "кг/м²",
     "dry_thickness": "70–100 мкм", "layers": 1, "interval": "12–24 ч"},
    {"name": "Акриловая эмаль", "category": "Эмаль",
     "consumption_min": 0.100, "consumption_max": 0.140, "unit": "кг/м²",
     "dry_thickness": "25–40 мкм", "layers": 2, "interval": "6–12 ч"},
    {"name": "Алкидная эмаль ПФ-115", "category": "Эмаль",
     "consumption_min": 0.100, "consumption_max": 0.150, "unit": "кг/м²",
     "dry_thickness": "25–35 мкм", "layers": 2, "interval": "24 ч"},
    {"name": "Грунтовка ГФ-021", "category": "Грунт",
     "consumption_min": 0.060, "consumption_max": 0.100, "unit": "кг/м²",
     "dry_thickness": "15–20 мкм", "layers": 1, "interval": "24 ч"},
    {"name": "Грунт-эмаль 3 в 1", "category": "Грунт-эмаль",
     "consumption_min": 0.070, "consumption_max": 0.120, "unit": "кг/м²",
     "dry_thickness": "40–60 мкм", "layers": 2, "interval": "24 ч"},
    {"name": "ХС-720 (хлоркаучуковая)", "category": "Эмаль",
     "consumption_min": 0.130, "consumption_max": 0.180, "unit": "кг/м²",
     "dry_thickness": "30–40 мкм", "layers": 3, "interval": "2–4 ч"},
    {"name": "ЭП-0010 (эпоксидная шпатлёвка)", "category": "Шпатлёвка",
     "consumption_min": 0.300, "consumption_max": 0.500, "unit": "кг/м²",
     "dry_thickness": "100–200 мкм", "layers": 1, "interval": "24 ч"},
    {"name": "Кремнийорганическая эмаль КО-811", "category": "Эмаль",
     "consumption_min": 0.120, "consumption_max": 0.170, "unit": "кг/м²",
     "dry_thickness": "25–40 мкм", "layers": 2, "interval": "2–6 ч"},
]

# ============================================================================
#  ВАЛИДАТОР
# ============================================================================

class FlexibleDoubleValidator(QValidator):
    _PARTIAL_RE = re.compile(r'^[+-]?\d*[.,]?\d*([eE][+-]?\d*)?$')

    def __init__(self, bottom: float = -1e12, top: float = 1e12,
                 decimals: int = 6, parent=None):
        super().__init__(parent)
        self.bottom = bottom
        self.top = top
        self.decimals = decimals

    @staticmethod
    def _normalize(text: str) -> str:
        return text.replace(',', '.')

    def validate(self, text: str, pos: int):
        if text == "":
            return QValidator.State.Intermediate, text, pos
        if not self._PARTIAL_RE.match(text):
            return QValidator.State.Invalid, text, pos
        normalized = self._normalize(text)
        if normalized in ('+', '-', '.', '+.', '-.'):
            return QValidator.State.Intermediate, text, pos
        try:
            value = float(normalized)
        except ValueError:
            return QValidator.State.Intermediate, text, pos
        if value < self.bottom - 1e-12 or value > self.top + 1e-12:
            return QValidator.State.Invalid, text, pos
        return QValidator.State.Acceptable, text, pos

    def fixup(self, text: str) -> str:
        return self._normalize(text)


# ============================================================================
#  УТИЛИТЫ
# ============================================================================

def mark_invalid(edit: QLineEdit, invalid: bool = True) -> None:
    if edit is None or not hasattr(edit, "property"):
        return
    if edit.property("invalid") == invalid:
        return
    edit.setProperty("invalid", invalid)
    edit.style().unpolish(edit)
    edit.style().polish(edit)


def get_current_material_name(widget: QWidget) -> str:
    try:
        window = widget.window()
        if hasattr(window, "material_combo"):
            return window.material_combo.currentText()
    except Exception:
        pass
    return "—"


def safe_eval(expr: str) -> float:
    expr = expr.strip().replace(',', '.')
    if not expr:
        raise ValueError("Пустое выражение")
    if not re.fullmatch(r'[\d+\-*/(). \teE]+', expr):
        raise ValueError("Недопустимые символы в выражении")
    try:
        tree = ast.parse(expr, mode="eval")
    except SyntaxError as exc:
        raise ValueError("Некорректное выражение") from exc

    allowed_binops = (ast.Add, ast.Sub, ast.Mult, ast.Div)
    allowed_unary = (ast.UAdd, ast.USub)

    def visit(node):
        if isinstance(node, ast.Expression):
            return visit(node.body)
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
            return float(node.value)
        if isinstance(node, ast.BinOp) and isinstance(node.op, allowed_binops):
            left, right = visit(node.left), visit(node.right)
            if isinstance(node.op, ast.Div) and right == 0:
                raise ValueError("Деление на ноль")
            return {
                ast.Add: left + right, ast.Sub: left - right,
                ast.Mult: left * right, ast.Div: left / right
            }[type(node.op)]
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, allowed_unary):
            value = visit(node.operand)
            return value if isinstance(node.op, ast.UAdd) else -value
        raise ValueError("Разрешены только числа и операции + - * / ( )")

    result = float(visit(tree))
    if not math.isfinite(result):
        raise ValueError("Результат не является конечным числом")
    return result


def format_number(value: float, default_decimals: int = 3,
                  widget: Optional[QWidget] = None) -> str:
    decimals = default_decimals
    if widget is not None:
        try:
            window = widget.window()
            decimals = int(getattr(window, "rounding_decimals", default_decimals))
        except Exception:
            pass
    return f"{value:.{max(0, min(8, decimals))}f}"


def create_c_locale_double_validator(bottom: float = 0.0,
                                     top: float = 1e12,
                                     decimals: int = 6) -> FlexibleDoubleValidator:
    return FlexibleDoubleValidator(bottom, top, decimals)


def create_input_row(label_text: str, unit_text: str = "мм", placeholder: str = "",
                     validator: Optional[QValidator] = None,
                     tooltip: str = ""
                     ) -> Tuple[QHBoxLayout, QLineEdit]:
    row = QHBoxLayout()
    row.setSpacing(10)
    label = QLabel(label_text)
    label.setMinimumWidth(200)
    label.setStyleSheet("font-weight: 500;")
    if tooltip:
        label.setToolTip(tooltip)
    line_edit = QLineEdit()
    line_edit.setPlaceholderText(placeholder)
    line_edit.setMinimumHeight(28)
    if validator is None:
        validator = create_c_locale_double_validator()
    line_edit.setValidator(validator)
    if tooltip:
        line_edit.setToolTip(tooltip)
    line_edit.textChanged.connect(lambda: mark_invalid(line_edit, False))
    unit_label = QLabel(unit_text)
    unit_label.setObjectName("unitLabel")
    unit_label.setMinimumWidth(40)
    row.addWidget(label)
    row.addWidget(line_edit, stretch=1)
    row.addWidget(unit_label)
    return row, line_edit


def get_float_from_edit(edit: QLineEdit, default: float = 0.0) -> float:
    try:
        text = edit.text().strip().replace(',', '.')
        return float(text) if text else default
    except ValueError:
        return default


def get_int_from_edit(edit: QLineEdit, default: int = 0,
                      warn: bool = False, parent: Optional[QWidget] = None) -> int:
    text = edit.text().strip().replace(',', '.')
    if not text:
        return default
    try:
        value = float(text)
        result = int(value)
        if warn and result != value and parent is not None:
            QMessageBox.information(
                parent, "Внимание",
                f"Значение «{text}» усечено до целого: {result}.")
        return result
    except ValueError:
        return default


def clear_layout(layout) -> None:
    while layout.count():
        item = layout.takeAt(0)
        widget = item.widget()
        child_layout = item.layout()
        if widget is not None:
            widget.deleteLater()
        elif child_layout is not None:
            clear_layout(child_layout)


# ============================================================================
#  ХЕЛПЕР: КОПИРОВАНИЕ ЗНАЧЕНИЙ ИЗ ПОЛЕЙ РЕЗУЛЬТАТА
# ============================================================================

class CopyableResultHelper:
    """Делает QLabel с objectName='resultLabel' копируемым."""

    _TOOLTIP = "Двойной клик — копировать всё\nПКМ — копировать только число"

    @staticmethod
    def make_copyable(label: QLabel, parent: QWidget) -> None:
        label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse |
            Qt.TextInteractionFlag.TextSelectableByKeyboard
        )
        label.setCursor(Qt.CursorShape.IBeamCursor)
        label.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        label.customContextMenuRequested.connect(
            lambda pos, l=label, p=parent: CopyableResultHelper._show_menu(l, p, pos))
        original_dbl = label.mouseDoubleClickEvent

        def _dbl(event, l=label, orig=original_dbl):
            QApplication.clipboard().setText(l.text())
            l.setToolTip("✅ Скопировано")
            QTimer.singleShot(1500, lambda: l.setToolTip(CopyableResultHelper._TOOLTIP))
            try:
                orig(event)
            except Exception:
                pass

        label.mouseDoubleClickEvent = _dbl
        label.setToolTip(CopyableResultHelper._TOOLTIP)

    @staticmethod
    def _show_menu(label: QLabel, parent: QWidget, pos: QPoint) -> None:
        menu = QMenu(parent)
        act_all = menu.addAction("📋 Копировать всё")
        act_num = menu.addAction("🔢 Копировать только число")
        chosen = menu.exec(label.mapToGlobal(pos))
        if chosen == act_all:
            QApplication.clipboard().setText(label.text())
        elif chosen == act_num:
            text = label.text().replace(',', '.')
            m = re.search(r'-?\d+\.?\d*', text)
            QApplication.clipboard().setText(m.group() if m else label.text())

    @staticmethod
    def enable_for_all(root_widget: QWidget) -> None:
        for lbl in root_widget.findChildren(QLabel):
            if lbl.objectName() == "resultLabel":
                CopyableResultHelper.make_copyable(lbl, root_widget)


# ---------- Работа с историей и материалами через QSettings ----------

def save_history(widget: QWidget, key: str, list_widget: QListWidget,
                 limit: int = 100) -> None:
    try:
        window = widget.window()
        if not hasattr(window, "settings"):
            return
        items = [list_widget.item(i).text() for i in range(list_widget.count())]
        items = items[-limit:]
        window.settings.setValue(key, json.dumps(items, ensure_ascii=False))
    except Exception:
        logger.exception("save_history failed for key %s", key)


def load_history(widget: QWidget, key: str, list_widget: QListWidget) -> None:
    try:
        window = widget.window()
        if not hasattr(window, "settings"):
            return
        raw = window.settings.value(key, "")
        if not raw:
            return
        items = json.loads(raw) if isinstance(raw, str) else raw
        if not isinstance(items, list):
            return
        for it in items:
            list_widget.addItem(str(it))
        if list_widget.count():
            list_widget.scrollToBottom()
    except Exception:
        logger.exception("load_history failed for key %s", key)


def clear_history(widget: QWidget, key: str, list_widget: QListWidget) -> None:
    list_widget.clear()
    save_history(widget, key, list_widget)


def load_custom_materials(settings: QSettings) -> Dict[str, float]:
    raw = settings.value("custom_materials", "")
    if not raw:
        return {}
    try:
        data = json.loads(raw) if isinstance(raw, str) else raw
        if isinstance(data, dict):
            return {str(k): float(v) for k, v in data.items()}
    except Exception:
        logger.exception("load_custom_materials failed")
    return {}


def save_custom_materials(settings: QSettings, materials: Dict[str, float]) -> None:
    settings.setValue("custom_materials", json.dumps(materials, ensure_ascii=False))


# ============================================================================
#  РАСЧЁТНЫЕ ФУНКЦИИ
# ============================================================================

def calculate_pipe_mass(outer_diameter, wall_thickness, length, density=STEEL_DENSITY):
    D = outer_diameter / 1000; s = wall_thickness / 1000; L = length
    return math.pi * (D - s) * s * L * density

def calculate_circle_mass(diameter, length, density=STEEL_DENSITY):
    d = diameter / 1000; L = length
    return (math.pi * d ** 2 / 4) * L * density

def calculate_sheet_mass(width, length, thickness, density=STEEL_DENSITY):
    w = width / 1000; L = length / 1000; t = thickness / 1000
    return w * L * t * density

def calculate_hexagon_mass(size_across_flats, length, density=STEEL_DENSITY):
    s = size_across_flats / 1000; L = length
    area = (math.sqrt(3) / 2) * s ** 2
    return area * L * density

def calculate_angle_mass(side_a, side_b, thickness, length, density=STEEL_DENSITY):
    a = side_a / 1000; b = side_b / 1000; t = thickness / 1000; L = length
    area = (a + b - t) * t
    return area * L * density

def calculate_channel_mass(height, flange_width, wall_thickness, length, density=STEEL_DENSITY):
    h = height / 1000; b = flange_width / 1000; s = wall_thickness / 1000; L = length
    area = (h - 2 * s) * s + 2 * b * s
    return area * L * density

def calculate_beam_mass(height, flange_width, web_thickness, flange_thickness, length,
                        density=STEEL_DENSITY):
    h = height / 1000; b = flange_width / 1000
    s = web_thickness / 1000; t = flange_thickness / 1000; L = length
    area = (h - 2 * t) * s + 2 * b * t
    return area * L * density

def calculate_single_pipe_insulation(D, t, L):
    Sr = math.pi * D * L
    Spi = math.pi * (D + 2 * t) * L
    Vi = math.pi * t * (D + t) * L
    return {'Sr': Sr, 'Spi': Spi, 'Vi': Vi}

def calculate_multiple_pipes_insulation_variant1(D1, D2, t, p, L):
    M = D2 + 2 * p
    Sr = (math.pi * D1 + 2 * M) * L
    Spi = (math.pi * (D1 + 2 * t) + 2 * M) * L
    S_outer = (math.pi / 4) * (D1 + 2 * t) ** 2 + M * (D1 + 2 * t)
    S_inner = (math.pi / 4) * D1 ** 2 + M * D1
    Vi = (S_outer - S_inner) * L
    return {'Sr': Sr, 'Spi': Spi, 'Vi': Vi}

def calculate_multiple_pipes_insulation_variant2(D1, M, t, L):
    # Sr - честная площадь поверхности двух металлических труб
    Sr = (math.pi * D1 * 2) * L
    
    # Полные внешние габариты изоляционного пучка (овальное сечение)
    W_outer = M + D1 + 2 * t  # полная ширина с изоляцией
    H_outer = D1 + 2 * t      # полная высота с изоляцией
    
    # Площадь внешнего покровного слоя (периметр овала * длина)
    Spi = (2 * M + math.pi * H_outer) * L
    
    # Честный расчет объемов через сечение овального контура:
    # 1. Полная площадь сечения внешнего кожуха (прямоугольник + круг)
    S_outer_total = (M * H_outer) + (math.pi / 4.0) * (H_outer ** 2)
    
    # 2. Фактическая площадь сечения двух внутренних металлических труб
    S_pipes_total = 2 * ((math.pi / 4.0) * (D1 ** 2))
    
    # 3. Объем чистой изоляции (Внешний овал минус металл труб)
    Vi = (S_outer_total - S_pipes_total) * L
    
    return {'Sr': Sr, 'Spi': Spi, 'Vi': Vi}

def calculate_bolt_weight_formula(diameter_mm, length_mm, density=STEEL_DENSITY):
    d = int(diameter_mm); L = int(length_mm)
    if d not in BOLT_PARAMETERS: return None
    params = BOLT_PARAMETERS[d]
    S = params["S"]; k = params["k"]; d1 = params["d1"]
    d1_m = d1 / 1000; L_m = L / 1000
    stem_mass = math.pi * (d1_m ** 2) / 4 * L_m * density
    S_m = S / 1000; k_m = k / 1000
    head_area = 0.866 * S_m ** 2
    head_mass = head_area * k_m * density
    return stem_mass + head_mass

def calculate_nut_weight_formula(diameter_mm, density=STEEL_DENSITY):
    d = int(diameter_mm)
    nut_params = {
        5: {"S": 8, "m": 4.7, "d1": 4.2}, 6: {"S": 10, "m": 5.2, "d1": 5.0},
        8: {"S": 13, "m": 6.5, "d1": 6.8}, 10: {"S": 17, "m": 8.4, "d1": 8.4},
        12: {"S": 19, "m": 10.4, "d1": 10.2}, 14: {"S": 22, "m": 11.5, "d1": 12.0},
        16: {"S": 24, "m": 13.0, "d1": 13.8}, 18: {"S": 27, "m": 15.0, "d1": 15.5},
        20: {"S": 30, "m": 16.0, "d1": 17.3}, 22: {"S": 32, "m": 18.0, "d1": 19.0},
        24: {"S": 36, "m": 19.0, "d1": 20.7}, 27: {"S": 41, "m": 22.0, "d1": 23.5},
        30: {"S": 46, "m": 24.0, "d1": 26.2}, 36: {"S": 55, "m": 29.0, "d1": 31.5},
    }
    if d not in nut_params: return None
    params = nut_params[d]
    S = params["S"]; m = params["m"]; d1 = params["d1"]
    S_m = S / 1000; m_m = m / 1000; d1_m = d1 / 1000
    hex_area = 0.866 * S_m ** 2
    hole_area = math.pi * (d1_m ** 2) / 4
    return (hex_area - hole_area) * m_m * density

def calculate_nail_weight(diameter_mm, length_mm, density=STEEL_DENSITY):
    d = diameter_mm / 1000; L = length_mm / 1000
    stem_mass = math.pi * (d ** 2) / 4 * L * density
    head_d = 2.5 * d; head_h = 0.1 * d
    head_mass = math.pi * (head_d ** 2) / 4 * head_h * density
    return stem_mass + head_mass

def paint_area_pipe(outer_diameter_mm, length_m):
    D = outer_diameter_mm / 1000.0
    return math.pi * D * length_m

def paint_area_circle(diameter_mm, length_m):
    d = diameter_mm / 1000.0
    return math.pi * d * length_m

def paint_area_channel(height_mm, flange_width_mm, wall_thickness_mm, length_m):
    h = height_mm / 1000.0; b = flange_width_mm / 1000.0; s = wall_thickness_mm / 1000.0
    P = 2 * h + 4 * b - 2 * s
    return P * length_m

def paint_area_beam(height_mm, flange_width_mm, web_thickness_mm, flange_thickness_mm, length_m):
    h = height_mm / 1000.0; b = flange_width_mm / 1000.0; s = web_thickness_mm / 1000.0
    P = 2 * h + 4 * b - 2 * s
    return P * length_m

def paint_area_sheet(width_mm, length_mm):
    w = width_mm / 1000.0; L = length_mm / 1000.0
    return 2 * w * L

def paint_area_hexagon(size_across_flats_mm, length_m):
    S = size_across_flats_mm / 1000.0
    side = S / math.sqrt(3.0)
    P = 6 * side
    return P * length_m

def paint_area_angle(side_a_mm, side_b_mm, thickness_mm, length_m):
    a = side_a_mm / 1000.0; b = side_b_mm / 1000.0
    P = 2 * a + 2 * b
    return P * length_m

def calculate_sandblasting(degree: str, abrasive: str, area: float):
    d = SANDBLAST_DEGREES[degree]
    a = SANDBLAST_ABRASIVES[abrasive]
    avg_abrasive = (d["abrasive_min"] + d["abrasive_max"]) / 2
    abrasive_per_m2 = avg_abrasive * a["k_consumption"]
    time_per_m2 = d["time_per_m2"] * a["k_time"]
    abrasive_kg = abrasive_per_m2 * area
    time_h = time_per_m2 * area
    return {
        "abrasive_per_m2": abrasive_per_m2,
        "time_per_m2": time_per_m2,
        "abrasive_kg": abrasive_kg,
        "time_h": time_h,
    }


# ============================================================================
#  КРЕПЁЖ (данные)
# ============================================================================

BOLT_MASS_GOST7798 = {
    (6, 10): 4.712, (6, 12): 4.924, (6, 16): 5.345, (6, 20): 5.767,
    (6, 25): 6.293, (6, 30): 6.820, (6, 35): 7.346, (6, 40): 7.872,
    (6, 45): 8.398, (6, 50): 8.925, (6, 55): 9.451, (6, 60): 9.977,
    (8, 12): 8.282, (8, 16): 9.065, (8, 20): 9.847, (8, 25): 11.012,
    (8, 30): 12.177, (8, 35): 13.342, (8, 40): 14.507, (8, 45): 15.672,
    (8, 50): 16.837, (8, 55): 18.002, (8, 60): 19.167, (8, 65): 20.332,
    (8, 70): 21.497, (8, 80): 23.827,
    (10, 16): 13.280, (10, 20): 14.570, (10, 25): 16.150, (10, 30): 17.730,
    (10, 35): 19.310, (10, 40): 20.890, (10, 45): 22.470, (10, 50): 24.050,
    (10, 55): 25.630, (10, 60): 27.210, (10, 65): 28.790, (10, 70): 30.370,
    (10, 80): 33.530, (10, 90): 36.690, (10, 100): 39.850,
    (12, 20): 19.270, (12, 25): 21.230, (12, 30): 23.190, (12, 35): 25.150,
    (12, 40): 27.110, (12, 45): 29.070, (12, 50): 31.030, (12, 55): 32.990,
    (12, 60): 34.950, (12, 65): 36.910, (12, 70): 38.870, (12, 80): 42.790,
    (12, 90): 46.710, (12, 100): 50.630, (12, 110): 54.550, (12, 120): 58.470,
    (16, 25): 34.330, (16, 30): 37.070, (16, 35): 39.810, (16, 40): 42.550,
    (16, 45): 45.290, (16, 50): 48.030, (16, 55): 50.770, (16, 60): 53.510,
    (16, 65): 56.250, (16, 70): 58.990, (16, 80): 64.470, (16, 90): 69.950,
    (16, 100): 75.430, (16, 110): 80.910, (16, 120): 86.390,
    (20, 30): 54.450, (20, 35): 58.170, (20, 40): 61.890, (20, 45): 65.610,
    (20, 50): 69.330, (20, 55): 73.050, (20, 60): 76.770, (20, 65): 80.490,
    (20, 70): 84.210, (20, 80): 91.650, (20, 90): 99.090, (20, 100): 106.530,
    (20, 110): 113.970, (20, 120): 121.410,
    (24, 35): 78.660, (24, 40): 83.220, (24, 45): 87.780, (24, 50): 92.340,
    (24, 55): 96.900, (24, 60): 101.460, (24, 65): 106.020, (24, 70): 110.580,
    (24, 80): 119.700, (24, 90): 128.820, (24, 100): 137.940, (24, 110): 147.060,
    (24, 120): 156.180,
    (30, 45): 123.200, (30, 50): 129.400, (30, 55): 135.600, (30, 60): 141.800,
    (30, 65): 148.000, (30, 70): 154.200, (30, 80): 166.600, (30, 90): 179.000,
    (30, 100): 191.400, (30, 110): 203.800, (30, 120): 216.200,
}

NUT_MASS_GOST5927 = {
    5: 1.44, 6: 2.20, 8: 5.10, 10: 8.40,
    12: 14.0, 14: 20.0, 16: 29.0, 18: 38.0,
    20: 51.0, 22: 64.0, 24: 79.0, 27: 110.0,
    30: 140.0, 36: 230.0,
}

BOLT_PARAMETERS = {
    3: {"S": 5.5, "k": 2.0, "d1": 2.5}, 4: {"S": 7, "k": 2.8, "d1": 3.4},
    5: {"S": 8, "k": 3.5, "d1": 4.2}, 6: {"S": 10, "k": 4.0, "d1": 5.0},
    8: {"S": 13, "k": 5.3, "d1": 6.8}, 10: {"S": 17, "k": 6.4, "d1": 8.4},
    12: {"S": 19, "k": 7.5, "d1": 10.2}, 16: {"S": 24, "k": 10.0, "d1": 13.8},
    20: {"S": 30, "k": 12.5, "d1": 17.3}, 24: {"S": 36, "k": 15.0, "d1": 20.7},
    30: {"S": 46, "k": 18.7, "d1": 26.2},
}

STANDARDS_DATA = {
    "Болт": {
        "ГОСТ 7798 шестигр.головка полная резьба": {
            "diameters": [6, 8, 10, 12, 16, 20, 24, 30],
            "lengths": [10, 12, 16, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 80, 90, 100, 110, 120],
            "coatings": ["Без покрытия", "Оцинкованный"],
            "strength_classes": ["5.8", "8.8", "10.9"],
        },
        "DIN 931 шестигр. головка неполная резьба": {
            "diameters": [6, 8, 10, 12, 16, 20, 24, 30],
            "lengths": [20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 80, 90, 100, 110, 120, 130, 140, 150],
            "coatings": ["Без покрытия", "Оцинкованный"],
            "strength_classes": ["5.8", "8.8", "10.9"],
        },
        "DIN 933 шестигр. головка полная резьба": {
            "diameters": [6, 8, 10, 12, 16, 20, 24, 30],
            "lengths": [10, 12, 16, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 80, 90, 100],
            "coatings": ["Без покрытия", "Оцинкованный"],
            "strength_classes": ["5.8", "8.8", "10.9"],
        },
        "ГОСТ 7795 с напр. подголовником": {
            "diameters": [6, 8, 10, 12, 16, 20],
            "lengths": [20, 25, 30, 35, 40, 45, 50, 55, 60, 70, 80, 90, 100],
            "coatings": ["Без покрытия", "Оцинкованный"],
            "strength_classes": ["5.8", "8.8"],
        },
        "ГОСТ 7796 с уменьшенной головой": {
            "diameters": [6, 8, 10, 12, 16, 20],
            "lengths": [16, 20, 25, 30, 35, 40, 45, 50, 55, 60, 70, 80],
            "coatings": ["Без покрытия", "Оцинкованный"],
            "strength_classes": ["5.8", "8.8"],
        },
        "DIN 607 с усом": {
            "diameters": [6, 8, 10, 12, 16, 20],
            "lengths": [20, 25, 30, 35, 40, 45, 50, 55, 60, 70, 80, 90, 100],
            "coatings": ["Без покрытия", "Оцинкованный"],
            "strength_classes": ["4.6", "5.8"],
        },
        "DIN 603 мебельный с квад.подголов.": {
            "diameters": [6, 8, 10, 12, 16, 20],
            "lengths": [20, 25, 30, 35, 40, 45, 50, 55, 60, 70, 80, 90, 100, 120, 140, 160],
            "coatings": ["Без покрытия", "Оцинкованный"],
            "strength_classes": ["4.6", "5.8"],
        },
        "DIN 6921 болт с фланцем": {
            "diameters": [6, 8, 10, 12, 16, 20],
            "lengths": [16, 20, 25, 30, 35, 40, 45, 50, 55, 60, 70, 80, 90, 100],
            "coatings": ["Без покрытия", "Оцинкованный"],
            "strength_classes": ["8.8", "10.9"],
        },
    },
    "Винт": {
        "DIN 912 с цилиндрической головкой": {
            "diameters": [3, 4, 5, 6, 8, 10, 12, 16, 20],
            "lengths": [6, 8, 10, 12, 16, 20, 25, 30, 35, 40, 45, 50, 55, 60, 70, 80, 90, 100],
            "coatings": ["Без покрытия", "Оцинкованный"],
            "strength_classes": ["8.8", "10.9", "12.9"],
        },
        "DIN 7991 с потайной головкой": {
            "diameters": [3, 4, 5, 6, 8, 10, 12, 16, 20],
            "lengths": [6, 8, 10, 12, 16, 20, 25, 30, 35, 40, 45, 50, 55, 60, 70, 80, 90, 100],
            "coatings": ["Без покрытия", "Оцинкованный"],
            "strength_classes": ["8.8", "10.9"],
        },
    },
    "Гайка": {
        "ГОСТ 5927 шестигранная": {
            "diameters": [5, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 27, 30, 36],
            "lengths": [], "coatings": ["Без покрытия", "Оцинкованный"],
            "strength_classes": ["5", "6", "8", "10"],
        },
        "DIN 934 шестигранная": {
            "diameters": [6, 8, 10, 12, 16, 20, 24, 30],
            "lengths": [], "coatings": ["Без покрытия", "Оцинкованный"],
            "strength_classes": ["6", "8", "10"],
        },
    },
    "Гвоздь": {
        "ГОСТ 4028 строительный": {
            "diameters": [2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 6.0],
            "lengths": [25, 30, 40, 50, 60, 70, 80, 90, 100, 120, 150],
            "coatings": ["Без покрытия", "Оцинкованный"],
            "strength_classes": ["—"],
        },
    },
    "Дюбели": {
        "Распорный пластиковый": {
            "diameters": [5, 6, 8, 10, 12, 14],
            "lengths": [25, 30, 40, 50, 60, 80, 100, 120],
            "coatings": ["—"], "strength_classes": ["—"],
        },
    },
    "Заклепка": {
        "Вытяжная алюминиевая": {
            "diameters": [3.2, 4.0, 4.8, 6.4],
            "lengths": [6, 8, 10, 12, 16, 20, 25, 30],
            "coatings": ["—"], "strength_classes": ["—"],
        },
        "Вытяжная стальная": {
            "diameters": [3.2, 4.0, 4.8, 6.4],
            "lengths": [6, 8, 10, 12, 16, 20, 25, 30],
            "coatings": ["Без покрытия", "Оцинкованный"],
            "strength_classes": ["—"],
        },
    },
    "Круг отрезной": {
        "По металлу": {
            "diameters": [115, 125, 180, 230],
            "lengths": [1.0, 1.6, 2.0, 2.5, 3.0, 3.2],
            "coatings": ["—"], "strength_classes": ["—"],
        },
    },
    "Перфорация": {
        "Уголок перфорированный": {
            "diameters": [30, 40, 50],
            "lengths": [1000, 2000, 3000],
            "thicknesses": [1.0, 1.2, 1.5, 2.0],
            "coatings": ["Без покрытия", "Оцинкованный"],
            "strength_classes": ["—"],
        },
        "Пластина перфорированная": {
            "diameters": [50, 75, 100],
            "lengths": [100, 150, 200, 250, 300],
            "thicknesses": [1.5, 2.0, 2.5, 3.0],
            "coatings": ["Без покрытия", "Оцинкованный"],
            "strength_classes": ["—"],
        },
    },
    "Саморез": {
        "По дереву черный": {
            "diameters": [3.5, 3.9, 4.2, 4.8, 6.3],
            "lengths": [16, 19, 25, 32, 35, 41, 45, 51, 55, 64, 70, 76, 89, 102],
            "coatings": ["Черный фосфатированный"],
            "strength_classes": ["—"],
        },
        "По металлу с буром": {
            "diameters": [3.5, 3.9, 4.2, 4.8, 5.5, 6.3],
            "lengths": [13, 16, 19, 25, 32, 38, 45, 51, 64, 76, 89, 102, 127],
            "coatings": ["Без покрытия", "Оцинкованный"],
            "strength_classes": ["—"],
        },
    },
}

# ============================================================================
#  ПРОФИЛИ ДЛЯ ГКЛ И ПЕРФОРИРОВАННЫЕ ИЗДЕЛИЯ
# ============================================================================

PROFILES_DATA = {
    "Профиль направляющий ПН (UD)": {
        "description": "Направляющий профиль для каркаса перегородок и облицовок (UD).",
        "sizes": [
            ("ПН-50×40×0.5",  50,  40, 0.5, 0.51),
            ("ПН-50×40×0.6",  50,  40, 0.6, 0.61),
            ("ПН-75×40×0.5",  75,  40, 0.5, 0.62),
            ("ПН-75×40×0.6",  75,  40, 0.6, 0.74),
            ("ПН-100×40×0.5", 100, 40, 0.5, 0.72),
            ("ПН-100×40×0.6", 100, 40, 0.6, 0.86),
        ],
    },
    "Профиль стоечный ПС (CD)": {
        "description": "Стоечный профиль для вертикальных стоек каркаса (CD).",
        "sizes": [
            ("ПС-50×50×0.5",  50,  50, 0.5, 0.58),
            ("ПС-50×50×0.6",  50,  50, 0.6, 0.70),
            ("ПС-75×50×0.5",  75,  50, 0.5, 0.68),
            ("ПС-75×50×0.6",  75,  50, 0.6, 0.82),
            ("ПС-100×50×0.5", 100, 50, 0.5, 0.78),
            ("ПС-100×50×0.6", 100, 50, 0.6, 0.94),
        ],
    },
    "Профиль потолочный ПП 60×27 (CD)": {
        "description": "Потолочный профиль для каркаса подвесных потолков.",
        "sizes": [
            ("ПП-60×27×0.5", 60, 27, 0.5, 0.38),
            ("ПП-60×27×0.6", 60, 27, 0.6, 0.46),
        ],
    },
    "Профиль направляющий потолочный ПН 28×27 (UD)": {
        "description": "Направляющий потолочный профиль для подвесных потолков.",
        "sizes": [
            ("ПН-28×27×0.5", 28, 27, 0.5, 0.26),
            ("ПН-28×27×0.6", 28, 27, 0.6, 0.31),
        ],
    },
    "Уголок защитный ПУ 31×31": {
        "description": "Уголок защитный для отделки внешних углов ГКЛ.",
        "sizes": [
            ("ПУ-31×31×0.4", 31, 31, 0.4, 0.15),
            ("ПУ-31×31×0.5", 31, 31, 0.5, 0.18),
        ],
    },
    "Профиль маячковый ПМ 20×20": {
        "description": "Профиль маячковый для штукатурных работ.",
        "sizes": [
            ("ПМ-20×20×0.3", 20, 20, 0.3, 0.09),
            ("ПМ-20×20×0.4", 20, 20, 0.4, 0.12),
            ("ПМ-20×20×0.5", 20, 20, 0.5, 0.15),
        ],
    },
    "Уголок штукатурный перфорированный": {
        "description": "Перфорированный уголок для защиты углов при штукатурных работах.",
        "sizes": [
            ("25×25×0.3", 25, 25, 0.3, 0.08),
            ("25×25×0.4", 25, 25, 0.4, 0.11),
            ("30×30×0.3", 30, 30, 0.3, 0.10),
            ("30×30×0.4", 30, 30, 0.4, 0.13),
            ("35×35×0.3", 35, 35, 0.3, 0.12),
            ("35×35×0.4", 35, 35, 0.4, 0.15),
        ],
    },
    "Пластина перфорированная (лента)": {
        "description": "Перфорированная монтажная лента для крепежа.",
        "sizes": [
            ("12×0.5", 12, 0, 0.5, 0.033),
            ("20×0.5", 20, 0, 0.5, 0.055),
            ("20×0.7", 20, 0, 0.7, 0.077),
            ("30×0.7", 30, 0, 0.7, 0.115),
            ("40×0.7", 40, 0, 0.7, 0.154),
            ("60×0.7", 60, 0, 0.7, 0.231),
        ],
    },
}

# ============================================================================
#  СТИЛЬ
# ============================================================================

DARK_THEME_STYLESHEET = """
QMainWindow { background-color: #1e1e1e; }
QWidget { background-color: #1e1e1e; color: #e0e0e0; font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px; }
QTabWidget::pane { border: 1px solid #3c3c3c; background-color: #252526; border-radius: 4px; }
QTabBar::tab { background-color: #2d2d30; color: #cccccc; padding: 10px 20px; margin-right: 2px; border-top-left-radius: 4px; border-top-right-radius: 4px; min-width: 120px; }
QTabBar::tab:selected { background-color: #007acc; color: #ffffff; font-weight: bold; }
QTabBar::tab:hover { background-color: #3e3e42; }
QGroupBox { border: 1px solid #3c3c3c; border-radius: 6px; margin-top: 12px; padding-top: 16px; font-weight: bold; background-color: #252526; }
QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; left: 12px; padding: 0 6px; color: #4ec9b0; }
QLineEdit { background-color: #3c3c3c; border: 1px solid #555555; border-radius: 4px; padding: 6px 10px; color: #ffffff; selection-background-color: #007acc; min-height: 28px; }
QLineEdit:focus { border: 1px solid #007acc; }
QLineEdit[invalid="true"] { border: 2px solid #da3633; background-color: #3a1e1e; }
QComboBox { background-color: #3c3c3c; border: 1px solid #555555; border-radius: 4px; padding: 4px 6px; color: #ffffff; font-size: 11px; min-height: 28px; }
QComboBox:hover { border: 1px solid #007acc; }
QComboBox::drop-down { border: none; width: 20px; }
QComboBox::down-arrow { image: none; border-left: 4px solid transparent; border-right: 4px solid transparent; border-top: 5px solid #cccccc; margin-right: 4px; }
QComboBox QAbstractItemView { background-color: #2d2d30; border: 1px solid #555555; selection-background-color: #007acc; color: #ffffff; }
QPushButton { background-color: #007acc; color: #ffffff; border: none; border-radius: 4px; padding: 8px 16px; font-weight: bold; font-size: 12px; min-width: 100px; min-height: 32px; }
QPushButton:hover { background-color: #0098ff; }
QPushButton:pressed { background-color: #005a9e; }
QPushButton:disabled { background-color: #3c3c3c; color: #888888; }
QPushButton#addButton { background-color: #2ea043; min-width: 40px; min-height: 40px; font-size: 18px; border-radius: 20px; }
QPushButton#addButton:hover { background-color: #3fb950; }
QPushButton#deleteButton { background-color: #da3633; min-width: 30px; padding: 4px 8px; }
QPushButton#deleteButton:hover { background-color: #f85149; }
QLabel { color: #e0e0e0; background-color: transparent; }
QLabel#resultLabel { color: #4ec9b0; font-size: 15px; font-weight: bold; padding: 8px; background-color: #1e3a2e; border: 1px solid #4ec9b0; border-radius: 6px; }
QLabel#titleLabel { color: #569cd6; font-size: 20px; font-weight: bold; padding: 10px 0; }
QLabel#unitLabel { color: #9cdcfe; font-size: 11px; }
QLabel#infoLabel { color: #8b949e; font-size: 11px; font-style: italic; }
QScrollArea { border: none; background-color: transparent; }
QTableWidget {
    background-color: #252526;
    alternate-background-color: #2d2d30;
    color: #e0e0e0;
    gridline-color: #3c3c3c;
    border: 1px solid #3c3c3c;
    border-radius: 4px;
    font-size: 11px;
}
QTableWidget::item {
    padding: 3px;
    border: none;
    background-color: #252526;
    color: #e0e0e0;
}
QTableWidget::item:alternate {
    background-color: #2d2d30;
    color: #e0e0e0;
}
QTableWidget::item:selected {
    background-color: #007acc;
    color: #ffffff;
}
QTableCornerButton::section {
    background-color: #2d2d30;
    border: 1px solid #3c3c3c;
}
QHeaderView::section {
    background-color: #2d2d30;
    color: #cccccc;
    padding: 6px;
    border: 1px solid #3c3c3c;
    font-weight: bold;
    font-size: 11px;
}
QRadioButton { color: #e0e0e0; spacing: 8px; }
QRadioButton::indicator { width: 16px; height: 16px; }
QRadioButton::indicator:checked { background-color: #007acc; border: 2px solid #007acc; border-radius: 8px; }
QRadioButton::indicator:unchecked { background-color: #3c3c3c; border: 2px solid #555555; border-radius: 8px; }
QPushButton#calcButton { min-width: 50px; min-height: 45px; font-size: 16px; background-color: #3c3c3c; color: #ffffff; border: 1px solid #555555; border-radius: 4px; }
QPushButton#calcButton:hover { background-color: #555555; }
QPushButton#calcButton:pressed { background-color: #007acc; }
QPushButton#calcButton[operator="true"] { background-color: #2d2d30; color: #4ec9b0; }
QPushButton#calcButton[operator="true"]:hover { background-color: #3c3c3c; }
QPushButton#calcButton[equals="true"] { background-color: #007acc; color: #ffffff; }
QPushButton#calcButton[equals="true"]:hover { background-color: #0098ff; }
QPushButton#calcButton[clear="true"] { background-color: #da3633; }
QPushButton#calcButton[clear="true"]:hover { background-color: #f85149; }
QListWidget { background-color: #252526; border: 1px solid #3c3c3c; border-radius: 4px; color: #e0e0e0; font-size: 11px; }
QListWidget::item { padding: 4px; }
QListWidget::item:selected { background-color: #007acc; }
#tetrisWidget { background-color: #0a0812; }
#game2048Widget { background-color: #0a0812; }
QDialog { background-color: #1e1e1e; }
QMenu { background-color: #2d2d30; color: #e0e0e0; border: 1px solid #555555; }
QMenu::item:selected { background-color: #007acc; color: #ffffff; }
"""

# ============================================================================
#  ДИАЛОГ: МАТЕРИАЛЫ (с редактированием пользовательских прямо в таблице)
# ============================================================================

class MaterialsDialog(QDialog):
    def __init__(self, parent, custom_materials: Dict[str, float]):
        super().__init__(parent)
        self.setWindowTitle("Библиотека материалов")
        self.setMinimumSize(560, 500)
        self.setModal(True)
        self.custom_materials = dict(custom_materials)
        self._updating_table = False  # защита от рекурсивного itemChanged

        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(12)

        info = QLabel(
            "Встроенные материалы (серые) изменить нельзя.\n"
            "Пользовательские (белые) можно добавлять, удалять "
            "и редактировать прямо в таблице: двойной клик по названию "
            "или плотности → правка → Enter.\n"
            "Изменения сохраняются при нажатии «ОК»."
        )
        info.setWordWrap(True)
        info.setObjectName("infoLabel")
        layout.addWidget(info)

        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Материал", "Плотность, кг/м³", "Источник"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.itemChanged.connect(self._on_item_changed)
        layout.addWidget(self.table, stretch=1)

        btn_row = QHBoxLayout()
        self.add_btn = QPushButton("➕ Добавить материал")
        self.add_btn.setObjectName("addButton")
        self.add_btn.clicked.connect(self._add_material)
        btn_row.addWidget(self.add_btn)

        self.delete_btn = QPushButton("🗑 Удалить выбранный")
        self.delete_btn.setObjectName("deleteButton")
        self.delete_btn.clicked.connect(self._delete_selected)
        btn_row.addWidget(self.delete_btn)

        btn_row.addStretch()

        reset_btn = QPushButton("↺ Сбросить пользовательские")
        reset_btn.clicked.connect(self._reset_custom)
        btn_row.addWidget(reset_btn)
        layout.addLayout(btn_row)

        dlg_btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        dlg_btns.accepted.connect(self.accept)
        dlg_btns.rejected.connect(self.reject)
        layout.addWidget(dlg_btns)

        self._refresh_table()

    # ---------------------------------------------------------------
    #  Заполнение таблицы
    # ---------------------------------------------------------------
    def _refresh_table(self):
        self._updating_table = True
        try:
            self.table.setRowCount(0)

            # Встроенные — только просмотр
            for name, density in BUILTIN_MATERIALS.items():
                row = self.table.rowCount()
                self.table.insertRow(row)

                item_name = QTableWidgetItem(name)
                item_name.setFlags(item_name.flags() & ~Qt.ItemFlag.ItemIsEditable)

                item_density = QTableWidgetItem(f"{density}")
                item_density.setFlags(item_density.flags() & ~Qt.ItemFlag.ItemIsEditable)

                item_src = QTableWidgetItem("встроенный")
                item_src.setFlags(item_src.flags() & ~Qt.ItemFlag.ItemIsEditable)

                for it in (item_name, item_density, item_src):
                    it.setForeground(QBrush(QColor("#8b949e")))

                self.table.setItem(row, 0, item_name)
                self.table.setItem(row, 1, item_density)
                self.table.setItem(row, 2, item_src)

            # Пользовательские — редактируемые столбцы 0 и 1
            for name, density in self.custom_materials.items():
                row = self.table.rowCount()
                self.table.insertRow(row)

                item_name = QTableWidgetItem(name)
                item_name.setData(Qt.ItemDataRole.UserRole, name)

                item_density = QTableWidgetItem(f"{density}")
                item_density.setData(Qt.ItemDataRole.UserRole, name)

                item_src = QTableWidgetItem("пользовательский")
                item_src.setFlags(item_src.flags() & ~Qt.ItemFlag.ItemIsEditable)

                self.table.setItem(row, 0, item_name)
                self.table.setItem(row, 1, item_density)
                self.table.setItem(row, 2, item_src)
        finally:
            self._updating_table = False

    # ---------------------------------------------------------------
    #  Обработка правок
    # ---------------------------------------------------------------
    def _on_item_changed(self, item: QTableWidgetItem):
        if self._updating_table:
            return
        row, col = item.row(), item.column()
        if col not in (0, 1):
            return

        src_item = self.table.item(row, 2)
        if src_item is None or src_item.text() != "пользовательский":
            return

        name_item = self.table.item(row, 0)
        density_item = self.table.item(row, 1)
        if name_item is None or density_item is None:
            return

        original_name = name_item.data(Qt.ItemDataRole.UserRole)
        if not original_name:
            return

        new_name = name_item.text().strip()
        new_density_text = density_item.text().strip().replace(',', '.')

        # --- Валидация названия ---
        if not new_name:
            QMessageBox.warning(self, "Ошибка",
                                "Название материала не может быть пустым.")
            self._restore_row(row, original_name)
            return
        if new_name in BUILTIN_MATERIALS:
            QMessageBox.warning(
                self, "Ошибка",
                f"Название «{new_name}» совпадает со встроенным материалом.")
            self._restore_row(row, original_name)
            return
        if new_name != original_name and new_name in self.custom_materials:
            QMessageBox.warning(
                self, "Ошибка",
                f"Материал «{new_name}» уже существует в пользовательских.")
            self._restore_row(row, original_name)
            return

        # --- Валидация плотности ---
        try:
            new_density = float(new_density_text)
            if not (0.0 < new_density <= 25000.0):
                raise ValueError
        except (ValueError, TypeError):
            QMessageBox.warning(
                self, "Ошибка",
                "Плотность должна быть числом в диапазоне 1…25000 кг/м³.")
            self._restore_row(row, original_name)
            return

        # --- Применение изменений ---
        if new_name != original_name:
            self.custom_materials.pop(original_name, None)
        self.custom_materials[new_name] = new_density

        self._updating_table = True
        try:
            name_item.setText(new_name)
            name_item.setData(Qt.ItemDataRole.UserRole, new_name)
            density_item.setText(f"{new_density:g}")
            density_item.setData(Qt.ItemDataRole.UserRole, new_name)
        finally:
            self._updating_table = False

    def _restore_row(self, row: int, original_name: str) -> None:
        """Откатить содержимое строки к текущему состоянию словаря."""
        self._updating_table = True
        try:
            name_item = self.table.item(row, 0)
            density_item = self.table.item(row, 1)
            if name_item is not None:
                name_item.setText(original_name)
            if density_item is not None:
                density = self.custom_materials.get(original_name)
                if density is not None:
                    density_item.setText(f"{density:g}")
        finally:
            self._updating_table = False

    # ---------------------------------------------------------------
    #  Кнопки
    # ---------------------------------------------------------------
    def _add_material(self):
        name, ok = QInputDialog.getText(self, "Новый материал", "Название:")
        if not ok or not name.strip():
            return
        name = name.strip()
        if name in BUILTIN_MATERIALS or name in self.custom_materials:
            QMessageBox.warning(self, "Ошибка",
                                "Материал с таким названием уже существует")
            return
        density, ok = QInputDialog.getDouble(
            self, "Плотность", f"Плотность «{name}», кг/м³:",
            1000.0, 1.0, 25000.0, 1)
        if not ok:
            return
        self.custom_materials[name] = float(density)
        self._refresh_table()

    def _delete_selected(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            QMessageBox.information(self, "Информация",
                                    "Выберите строку для удаления")
            return
        for idx in sorted([r.row() for r in rows], reverse=True):
            name_item = self.table.item(idx, 0)
            src_item = self.table.item(idx, 2)
            if not name_item or not src_item:
                continue
            if src_item.text() != "пользовательский":
                QMessageBox.information(
                    self, "Информация",
                    "Встроенные материалы удалить нельзя.")
                continue
            name = name_item.text()
            self.custom_materials.pop(name, None)
        self._refresh_table()

    def _reset_custom(self):
        if not self.custom_materials:
            return
        reply = QMessageBox.question(
            self, "Сбросить?",
            "Удалить все пользовательские материалы?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.custom_materials.clear()
            self._refresh_table()

    def get_custom_materials(self) -> Dict[str, float]:
        return self.custom_materials


# ============================================================================
#  ТЕТРИС
# ============================================================================

class Particle:
    def __init__(self, x, y, color, speed_mult=1.0, gravity=0.05):
        self.x = x; self.y = y; self.color = color
        self.vx = random.uniform(-2, 2) * speed_mult
        self.vy = random.uniform(-4, 0) * speed_mult
        self.gravity = gravity
        self.size = random.uniform(3, 8)
        self.life = random.uniform(30, 60)
        self.max_life = self.life

    def update(self):
        self.x += self.vx; self.y += self.vy
        self.vy += self.gravity; self.life -= 1


class FloatingText:
    def __init__(self, x, y, text, color):
        self.x = x; self.y = y; self.text = text; self.color = color
        self.life = 60; self.max_life = 60; self.vy = -1.5

    def update(self):
        self.y += self.vy; self.life -= 1


TETRIS_SHAPES = [
    [[(0, 0), (1, 0), (2, 0), (3, 0)], [(2, 0), (2, 1), (2, -1), (2, -2)]],
    [[(0, 0), (1, 0), (0, 1), (1, 1)]],
    [[(1, 0), (0, 1), (1, 1), (2, 1)], [(1, 0), (1, 1), (2, 1), (1, 2)],
     [(0, 0), (1, 0), (2, 0), (1, 1)], [(1, 0), (0, 1), (1, 1), (1, 2)]],
    [[(1, 0), (2, 0), (0, 1), (1, 1)], [(0, 0), (0, 1), (1, 1), (1, 2)]],
    [[(0, 0), (1, 0), (1, 1), (2, 1)], [(1, 0), (0, 1), (1, 1), (0, 2)]],
    [[(0, 0), (0, 1), (1, 1), (2, 1)], [(0, 0), (1, 0), (0, 1), (0, 2)],
     [(0, 0), (1, 0), (2, 0), (2, 1)], [(2, 0), (2, 1), (1, 2), (2, 2)]],
    [[(0, 1), (1, 1), (2, 1), (2, 0)], [(0, 0), (0, 1), (0, 2), (1, 2)],
     [(0, 0), (1, 0), (2, 0), (0, 1)], [(0, 0), (1, 0), (1, 1), (1, 2)]],
]

TETRIS_COLORS = [
    QColor(0, 240, 255), QColor(255, 230, 0), QColor(180, 0, 255),
    QColor(0, 255, 80), QColor(255, 30, 60), QColor(30, 120, 255), QColor(255, 140, 0)
]


class TetrisPiece:
    def __init__(self):
        self.shape_idx = random.randint(0, len(TETRIS_SHAPES) - 1)
        self.rotation = 0
        self.color = random.choice(TETRIS_COLORS)
        self.x = 10 // 2 - 2; self.y = 0

    def get_blocks(self):
        variants = TETRIS_SHAPES[self.shape_idx]
        return variants[self.rotation % len(variants)]

    def rotate(self):
        self.rotation = (self.rotation + 1) % len(TETRIS_SHAPES[self.shape_idx])


class TetrisWidget(QWidget):
    BG_COLOR = QColor(12, 10, 20)
    NEON_PURPLE = QColor(180, 50, 255); NEON_CYAN = QColor(0, 240, 255)
    NEON_PINK = QColor(255, 30, 140); NEON_GREEN = QColor(50, 255, 150)
    TEXT_WHITE = QColor(245, 245, 255); TEXT_MUTED = QColor(130, 115, 165)
    GRID_LINE = QColor(30, 22, 50)
    PIECE_COLORS = TETRIS_COLORS; SHAPES = TETRIS_SHAPES
    BLOCK_SIZE = 30; GRID_WIDTH, GRID_HEIGHT = 10, 20; GRID_X, GRID_Y = 40, 80

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("tetrisWidget")
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setMinimumSize(820, 720)
        self.grid = [[None for _ in range(self.GRID_WIDTH)] for _ in range(self.GRID_HEIGHT)]
        self.current_piece = self.new_piece(); self.next_piece = self.new_piece()
        self.score = 0; self.lines_cleared = 0; self.level = 1
        self.game_over = False; self.paused = False
        self.fall_speed = 500; self.bomb_charges = 1; self.pulse_timer = 0
        self.bomb_cursor_x = self.GRID_WIDTH // 2; self.bomb_cursor_y = self.GRID_HEIGHT // 2
        self.placing_bomb = False; self.particles = []; self.floating_texts = []
        self.shake_timer = 0; self.flash_rows = []; self.flash_timer = 0
        self.timer = QTimer(self); self.timer.timeout.connect(self.game_loop); self.timer.start(16)
        self.fall_timer = QTimer(self); self.fall_timer.timeout.connect(self.fall_step)
        self.fall_timer.start(self.fall_speed)
        self.pause_button = QPushButton("⏸ ПАУЗА", self)
        self.pause_button.setObjectName("tetrisPauseButton")
        self.pause_button.setFixedSize(125, 38)
        self.pause_button.clicked.connect(self.toggle_pause); self.pause_button.raise_()
        self.font_title = self.font(); self.font_title.setPointSize(16); self.font_title.setBold(True)
        self.font_large = self.font(); self.font_large.setPointSize(18); self.font_large.setBold(True)
        self.font_medium = self.font(); self.font_medium.setPointSize(12); self.font_medium.setBold(True)
        self.font_small = self.font(); self.font_small.setPointSize(10); self.font_small.setBold(True)
        self.shake_offset = QPoint(0, 0)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.pause_button.move(max(10, self.width() - 155), 28)

    def toggle_pause(self):
        if self.game_over: return
        self.paused = not self.paused
        self.pause_button.setText("▶ ПРОДОЛЖИТЬ" if self.paused else "⏸ ПАУЗА")
        if self.paused: self.timer.stop(); self.fall_timer.stop()
        else: self.timer.start(16); self.fall_timer.start(self.fall_speed)
        self.setFocus(); self.update()

    def new_piece(self): return TetrisPiece()

    def check_collision(self, piece, offset_x=0, offset_y=0):
        for b in piece.get_blocks():
            x = piece.x + b[0] + offset_x; y = piece.y + b[1] + offset_y
            if x < 0 or x >= self.GRID_WIDTH or y >= self.GRID_HEIGHT: return True
            if y >= 0 and self.grid[y][x] is not None: return True
        return False

    def lock_piece(self):
        for b in self.current_piece.get_blocks():
            x = self.current_piece.x + b[0]; y = self.current_piece.y + b[1]
            if y >= 0:
                self.grid[y][x] = self.current_piece.color
                if y == self.GRID_HEIGHT - 1 or self.grid[y + 1][x] is not None:
                    px = self.GRID_X + x * self.BLOCK_SIZE + self.BLOCK_SIZE // 2
                    py = self.GRID_Y + (y + 1) * self.BLOCK_SIZE
                    for _ in range(2):
                        self.particles.append(Particle(px, py, self.current_piece.color, 0.6, 0.1))
        cleared_rows = [y for y in range(self.GRID_HEIGHT) if all(self.grid[y][x] is not None for x in range(self.GRID_WIDTH))]
        if cleared_rows:
            self.flash_rows = cleared_rows; self.flash_timer = 6
            self.shake_timer = 8 + len(cleared_rows) * 2
            for row in cleared_rows:
                for x in range(self.GRID_WIDTH):
                    color = self.grid[row][x]
                    px = self.GRID_X + x * self.BLOCK_SIZE + self.BLOCK_SIZE // 2
                    py = self.GRID_Y + row * self.BLOCK_SIZE + self.BLOCK_SIZE // 2
                    for _ in range(4):
                        self.particles.append(Particle(px, py, color))
            self.grid = [row for i, row in enumerate(self.grid) if i not in cleared_rows]
            for _ in range(len(cleared_rows)):
                self.grid.insert(0, [None for _ in range(self.GRID_WIDTH)])
            cleared_now = len(cleared_rows)
            earned_score = 100 * cleared_now * self.level
            if cleared_now == 4:
                earned_score += 300
                self.floating_texts.append(FloatingText(self.GRID_X + self.GRID_WIDTH * self.BLOCK_SIZE // 2 - 40,
                                                         self.GRID_Y + self.GRID_HEIGHT * self.BLOCK_SIZE // 2,
                                                         "TETRIS!", self.NEON_CYAN))
            self.score += earned_score
            self.floating_texts.append(FloatingText(self.GRID_X + self.GRID_WIDTH * self.BLOCK_SIZE // 2 - 30,
                                                     self.GRID_Y + 100, f"+{earned_score}", self.NEON_GREEN))
            old_lines = self.lines_cleared; self.lines_cleared += cleared_now
            if self.lines_cleared // 6 > old_lines // 6:
                self.bomb_charges += 1
                self.floating_texts.append(FloatingText(self.GRID_X + 20, self.GRID_Y + 50, "+1 BOMB!", self.NEON_PINK))
            new_level = self.lines_cleared // 6 + 1
            if new_level > self.level:
                self.level = new_level
                self.fall_speed = max(150, 500 - (self.level - 1) * 40)
                self.fall_timer.setInterval(self.fall_speed)
        self.current_piece = self.next_piece; self.next_piece = self.new_piece()
        if self.check_collision(self.current_piece): self.game_over = True

    def explode_bomb(self, cx, cy):
        self.shake_timer = 20
        center_px = self.GRID_X + cx * self.BLOCK_SIZE + self.BLOCK_SIZE // 2
        center_py = self.GRID_Y + cy * self.BLOCK_SIZE + self.BLOCK_SIZE // 2
        for _ in range(70):
            color = random.choice([self.NEON_PINK, self.NEON_CYAN, self.TEXT_WHITE, QColor(255, 160, 0)])
            self.particles.append(Particle(center_px, center_py, color, 1.5))
        self.floating_texts.append(FloatingText(center_px - 20, center_py - 10, "BOOM!", QColor(255, 50, 50)))
        for y in range(max(0, cy - 1), min(self.GRID_HEIGHT, cy + 2)):
            for x in range(max(0, cx - 1), min(self.GRID_WIDTH, cx + 2)):
                self.grid[y][x] = None
        self.apply_gravity()

    def apply_gravity(self):
        for x in range(self.GRID_WIDTH):
            column = [self.grid[y][x] for y in range(self.GRID_HEIGHT) if self.grid[y][x] is not None]
            new_col = [None] * (self.GRID_HEIGHT - len(column)) + column
            for y in range(self.GRID_HEIGHT): self.grid[y][x] = new_col[y]

    def fall_step(self):
        if not self.game_over and not self.placing_bomb:
            if not self.check_collision(self.current_piece, 0, 1): self.current_piece.y += 1
            else: self.lock_piece()

    def game_loop(self):
        if self.paused: return
        self.pulse_timer += 0.02 + self.level * 0.002
        for p in self.particles[:]:
            p.update()
            if p.life <= 0: self.particles.remove(p)
        for ft in self.floating_texts[:]:
            ft.update()
            if ft.life <= 0: self.floating_texts.remove(ft)
        if self.flash_timer > 0: self.flash_timer -= 1
        if self.shake_timer > 0:
            self.shake_offset = QPoint(random.randint(-5, 5), random.randint(-5, 5))
            self.shake_timer -= 1
        else: self.shake_offset = QPoint(0, 0)
        self.update()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_R and self.game_over: self.reset_game(); return
        if self.game_over: return
        if event.key() == Qt.Key.Key_P: self.toggle_pause(); return
        if self.placing_bomb:
            if event.key() == Qt.Key.Key_Left: self.bomb_cursor_x = max(0, self.bomb_cursor_x - 1)
            elif event.key() == Qt.Key.Key_Right: self.bomb_cursor_x = min(self.GRID_WIDTH - 1, self.bomb_cursor_x + 1)
            elif event.key() == Qt.Key.Key_Up: self.bomb_cursor_y = max(0, self.bomb_cursor_y - 1)
            elif event.key() == Qt.Key.Key_Down: self.bomb_cursor_y = min(self.GRID_HEIGHT - 1, self.bomb_cursor_y + 1)
            elif event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Space):
                self.explode_bomb(self.bomb_cursor_x, self.bomb_cursor_y)
                self.bomb_charges -= 1; self.placing_bomb = False
            elif event.key() in (Qt.Key.Key_Escape, Qt.Key.Key_B): self.placing_bomb = False
            return
        if event.key() == Qt.Key.Key_Left and not self.check_collision(self.current_piece, -1, 0): self.current_piece.x -= 1
        elif event.key() == Qt.Key.Key_Right and not self.check_collision(self.current_piece, 1, 0): self.current_piece.x += 1
        elif event.key() == Qt.Key.Key_Down and not self.check_collision(self.current_piece, 0, 1):
            self.current_piece.y += 1; self.score += 1
        elif event.key() == Qt.Key.Key_Up:
            old_rotation = self.current_piece.rotation
            old_x, old_y = self.current_piece.x, self.current_piece.y
            self.current_piece.rotate(); kicked = False
            for dx, dy in ((0, 0), (-1, 0), (1, 0), (-2, 0), (2, 0), (0, -1), (-1, -1), (1, -1)):
                if not self.check_collision(self.current_piece, dx, dy):
                    self.current_piece.x += dx; self.current_piece.y += dy; kicked = True; break
            if not kicked:
                self.current_piece.rotation = old_rotation
                self.current_piece.x, self.current_piece.y = old_x, old_y
        elif event.key() == Qt.Key.Key_Space:
            while not self.check_collision(self.current_piece, 0, 1):
                self.current_piece.y += 1; self.score += 2
            self.lock_piece()
        elif event.key() == Qt.Key.Key_B and self.bomb_charges > 0: self.placing_bomb = True

    def reset_game(self):
        self.grid = [[None for _ in range(self.GRID_WIDTH)] for _ in range(self.GRID_HEIGHT)]
        self.current_piece = self.new_piece(); self.next_piece = self.new_piece()
        self.score = 0; self.lines_cleared = 0; self.level = 1
        self.game_over = False; self.paused = False
        self.pause_button.setText("⏸ ПАУЗА")
        self.timer.start(16); self.fall_timer.start(self.fall_speed)
        self.fall_speed = 500; self.bomb_charges = 1; self.pulse_timer = 0
        self.placing_bomb = False; self.particles = []; self.floating_texts = []
        self.shake_timer = 0; self.flash_rows = []; self.flash_timer = 0
        self.shake_offset = QPoint(0, 0); self.fall_timer.setInterval(self.fall_speed)

    def draw_block(self, painter, px, py, color, size=BLOCK_SIZE, alpha=255):
        if color is None: color = self.NEON_CYAN
        if alpha < 255: painter.setOpacity(alpha / 255.0)
        painter.setBrush(color); painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(px + 1, py + 1, size - 2, size - 2, 4, 4)
        painter.setBrush(QColor(15, 12, 25))
        painter.drawRoundedRect(px + 5, py + 5, size - 10, size - 10, 2, 2)
        painter.setBrush(color)
        painter.drawRoundedRect(px + 9, py + 9, size - 18, size - 18, 1, 1)
        painter.setOpacity(1.0)

    def draw_panel_box(self, painter, x, y, width, height, title):
        painter.setBrush(QColor(22, 16, 35, 230)); painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(x, y, width, height, 6, 6)
        painter.setPen(QPen(self.NEON_PURPLE, 1)); painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawRoundedRect(x, y, width, height, 6, 6)
        painter.setPen(QPen(self.NEON_CYAN, 2)); painter.drawLine(x, y, x + 20, y)
        painter.setPen(self.NEON_CYAN); painter.setFont(self.font_small)
        painter.drawText(x + 15, y + 22, title)

    def paintEvent(self, event):
        painter = QPainter(self); painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.translate(self.shake_offset)
        painter.fillRect(self.rect(), self.BG_COLOR)
        for x_line in range(0, self.width(), 40):
            painter.setPen(QColor(18, 14, 30)); painter.drawLine(x_line, 0, x_line, self.height())
        for y_line in range(0, self.height(), 40):
            painter.setPen(QColor(18, 14, 30)); painter.drawLine(0, y_line, self.width(), y_line)
        painter.setPen(self.NEON_CYAN); painter.setFont(self.font_title)
        painter.drawText(self.GRID_X, 30, "ROGUE-TETRIS // BOMB PROTOCOL")
        painter.setPen(self.NEON_PINK)
        painter.drawLine(self.GRID_X, 60, self.GRID_X + self.GRID_WIDTH * self.BLOCK_SIZE, 60)
        play_width = self.GRID_WIDTH * self.BLOCK_SIZE; play_height = self.GRID_HEIGHT * self.BLOCK_SIZE
        painter.fillRect(self.GRID_X, self.GRID_Y, play_width, play_height, QColor(10, 8, 18, 220))
        painter.setPen(self.GRID_LINE)
        for x in range(self.GRID_WIDTH + 1):
            painter.drawLine(self.GRID_X + x * self.BLOCK_SIZE, self.GRID_Y,
                             self.GRID_X + x * self.BLOCK_SIZE, self.GRID_Y + play_height)
        for y in range(self.GRID_HEIGHT + 1):
            painter.drawLine(self.GRID_X, self.GRID_Y + y * self.BLOCK_SIZE,
                             self.GRID_X + play_width, self.GRID_Y + y * self.BLOCK_SIZE)
        base_pulse_r = int(140 + 60 * math.sin(self.pulse_timer))
        base_pulse_g = int(30 + 30 * self.level)
        pulse_color = QColor(min(255, base_pulse_r), min(255, base_pulse_g), 255)
        painter.setPen(QPen(pulse_color, 2)); painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawRoundedRect(self.GRID_X, self.GRID_Y, play_width, play_height, 4, 4)
        for y in range(self.GRID_HEIGHT):
            for x in range(self.GRID_WIDTH):
                if self.grid[y][x] is not None:
                    self.draw_block(painter, self.GRID_X + x * self.BLOCK_SIZE,
                                    self.GRID_Y + y * self.BLOCK_SIZE, self.grid[y][x])
        if self.flash_timer > 0:
            painter.setBrush(QColor(255, 255, 255, 200)); painter.setPen(Qt.PenStyle.NoPen)
            for row in self.flash_rows:
                painter.fillRect(self.GRID_X, self.GRID_Y + row * self.BLOCK_SIZE,
                                 play_width, self.BLOCK_SIZE, QColor(255, 255, 255, 200))
        if not self.game_over and not self.placing_bomb:
            ghost_y = self.current_piece.y
            while not self.check_collision(self.current_piece, 0, ghost_y - self.current_piece.y + 1): ghost_y += 1
            for b in self.current_piece.get_blocks():
                gx = self.GRID_X + (self.current_piece.x + b[0]) * self.BLOCK_SIZE
                gy = self.GRID_Y + (ghost_y + b[1]) * self.BLOCK_SIZE
                if gy >= self.GRID_Y:
                    painter.setBrush(Qt.BrushStyle.NoBrush)
                    painter.setPen(QPen(self.current_piece.color, 1))
                    painter.drawRoundedRect(gx + 2, gy + 2, self.BLOCK_SIZE - 4, self.BLOCK_SIZE - 4, 4, 4)
        if not self.game_over and not self.placing_bomb:
            for b in self.current_piece.get_blocks():
                px = self.GRID_X + (self.current_piece.x + b[0]) * self.BLOCK_SIZE
                py = self.GRID_Y + (self.current_piece.y + b[1]) * self.BLOCK_SIZE
                if py >= self.GRID_Y:
                    self.draw_block(painter, px, py, self.current_piece.color)
                    painter.setOpacity(0.2); painter.setBrush(self.current_piece.color)
                    painter.setPen(Qt.PenStyle.NoPen)
                    painter.drawRoundedRect(px - 2, py - 2, self.BLOCK_SIZE + 4, self.BLOCK_SIZE + 4, 4, 4)
                    painter.setOpacity(1.0)
        for p in self.particles:
            painter.setOpacity(p.life / p.max_life); painter.setBrush(p.color)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawEllipse(int(p.x), int(p.y), int(p.size), int(p.size))
        painter.setOpacity(1.0)
        for ft in self.floating_texts:
            painter.setOpacity(ft.life / ft.max_life); painter.setPen(ft.color)
            painter.setFont(self.font_medium); painter.drawText(int(ft.x), int(ft.y), ft.text)
        painter.setOpacity(1.0)
        if self.placing_bomb:
            bx = self.GRID_X + self.bomb_cursor_x * self.BLOCK_SIZE
            by = self.GRID_Y + self.bomb_cursor_y * self.BLOCK_SIZE
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.setPen(QPen(QColor(255, 50, 50), 2))
            painter.drawRoundedRect(bx - self.BLOCK_SIZE, by - self.BLOCK_SIZE,
                                    self.BLOCK_SIZE * 3, self.BLOCK_SIZE * 3, 4, 4)
            painter.setPen(self.NEON_PINK); painter.setFont(self.font_small)
            painter.drawText(self.GRID_X, self.GRID_Y + play_height + 25,
                             "СТРЕЛКИ — ВЫБОР ЗОНЫ, ENTER — ВЗРЫВ")
        panel_width = 350; panel_x = self.GRID_X + play_width + 40
        stats_y = 80; stats_h = 100
        self.draw_panel_box(painter, panel_x, stats_y, panel_width, stats_h, "СТАТИСТИКА")
        painter.setPen(self.TEXT_WHITE); painter.setFont(self.font_large)
        painter.drawText(panel_x + 15, stats_y + 70, f"СЧЁТ: {self.score}")
        next_y = 200; next_h = 110
        self.draw_panel_box(painter, panel_x, next_y, panel_width, next_h, "СЛЕДУЮЩИЙ МОДУЛЬ")
        blocks = self.next_piece.get_blocks()
        min_x = min(b[0] for b in blocks); max_x = max(b[0] for b in blocks)
        min_y = min(b[1] for b in blocks); max_y = max(b[1] for b in blocks)
        block_s = 22; fig_w = (max_x - min_x + 1) * block_s; fig_h = (max_y - min_y + 1) * block_s
        center_x = panel_x + panel_width // 2; center_y = next_y + 22 + (next_h - 22) // 2 + 8
        start_x = center_x - fig_w // 2 - min_x * block_s; start_y = center_y - fig_h // 2 - min_y * block_s
        for b in blocks:
            bx = start_x + b[0] * block_s; by = start_y + b[1] * block_s
            self.draw_block(painter, bx, by, self.next_piece.color, block_s)
        arsenal_y = 330; arsenal_h = 130
        self.draw_panel_box(painter, panel_x, arsenal_y, panel_width, arsenal_h, "АРСЕНАЛ")
        painter.setPen(self.NEON_GREEN); painter.setFont(self.font_medium)
        painter.drawText(panel_x + 15, arsenal_y + 60, f"УРОВЕНЬ: {self.level}")
        painter.setPen(self.TEXT_MUTED)
        painter.drawText(panel_x + 15, arsenal_y + 85, f"ЛИНИИ: {self.lines_cleared}")
        painter.setPen(self.NEON_PINK)
        painter.drawText(panel_x + 15, arsenal_y + 110, f"БОМБЫ (B): {self.bomb_charges}")
        ctrl_y = 480; ctrl_h = 210
        self.draw_panel_box(painter, panel_x, ctrl_y, panel_width, ctrl_h, "УПРАВЛЕНИЕ")
        controls = ["← →   Движение фигуры", "↑     Поворот", "↓     Ускорить падение",
                    "ПРОБЕЛ - Сброс вниз", "B     Активировать бомбу"]
        painter.setPen(self.TEXT_MUTED); painter.setFont(self.font_small)
        for i, c_text in enumerate(controls):
            painter.drawText(panel_x + 15, ctrl_y + 60 + i * 28, c_text)
        if self.paused and not self.game_over:
            overlay = QColor(12, 10, 20, 210); painter.fillRect(self.rect(), overlay)
            painter.setPen(self.NEON_CYAN); painter.setFont(self.font_large)
            text_pause = "ПАУЗА"; r = painter.fontMetrics().boundingRect(text_pause)
            painter.drawText((self.width() - r.width()) // 2, self.height() // 2, text_pause)
        if self.game_over:
            overlay = QColor(12, 10, 20, 230); painter.fillRect(self.rect(), overlay)
            painter.setPen(QColor(255, 40, 80)); painter.setFont(self.font_large)
            go_text = "СБОЙ СИСТЕМЫ: ЗАБЕГ ЗАВЕРШЕН"
            rect = painter.fontMetrics().boundingRect(go_text)
            painter.drawText((self.width() - rect.width()) // 2, self.height() // 2 - 10, go_text)
            painter.setPen(self.TEXT_MUTED); painter.setFont(self.font_medium)
            restart_text = "Нажмите 'R' для перезагрузки"
            painter.drawText((self.width() - painter.fontMetrics().boundingRect(restart_text).width()) // 2,
                             self.height() // 2 + 30, restart_text)
        painter.end()


# ============================================================================
#  ИГРА 2048
# ============================================================================

class Game2048Widget(QWidget):
    BG_COLOR = QColor(12, 10, 20)
    TILE_COLORS = {
        0: QColor(30, 25, 45), 2: QColor(60, 50, 80), 4: QColor(80, 65, 105),
        8: QColor(120, 60, 160), 16: QColor(180, 50, 200), 32: QColor(230, 40, 170),
        64: QColor(255, 60, 100), 128: QColor(255, 120, 40), 256: QColor(255, 180, 30),
        512: QColor(255, 230, 50), 1024: QColor(100, 255, 150), 2048: QColor(0, 240, 255),
    }
    NEON_CYAN = QColor(0, 240, 255); NEON_PINK = QColor(255, 30, 140)
    TEXT_WHITE = QColor(245, 245, 255); TEXT_MUTED = QColor(130, 115, 165)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("game2048Widget")
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setMinimumSize(600, 700)
        self.grid = [[0] * 4 for _ in range(4)]
        self.score = 0; self.best_score = 0; self.game_over = False; self.won = False
        self.tile_size = 100; self.tile_margin = 12; self.grid_origin_x = 40; self.grid_origin_y = 140
        self.spawn_tile(); self.spawn_tile()
        self.font_tile = self.font(); self.font_tile.setPointSize(28); self.font_tile.setBold(True)
        self.font_tile_small = self.font(); self.font_tile_small.setPointSize(20); self.font_tile_small.setBold(True)
        self.font_title = self.font(); self.font_title.setPointSize(22); self.font_title.setBold(True)
        self.font_score = self.font(); self.font_score.setPointSize(16); self.font_score.setBold(True)

    def spawn_tile(self):
        empty = [(r, c) for r in range(4) for c in range(4) if self.grid[r][c] == 0]
        if not empty: return
        r, c = random.choice(empty)
        self.grid[r][c] = 4 if random.random() < 0.1 else 2

    def can_move(self):
        for r in range(4):
            for c in range(4):
                if self.grid[r][c] == 0: return True
                if c < 3 and self.grid[r][c] == self.grid[r][c + 1]: return True
                if r < 3 and self.grid[r][c] == self.grid[r + 1][c]: return True
        return False

    def move(self, direction):
        if self.game_over: return
        moved = False
        new_grid = [row[:] for row in self.grid]
        if direction == "left":
            for r in range(4):
                row = [x for x in new_grid[r] if x != 0]
                merged = []; i = 0
                while i < len(row):
                    if i + 1 < len(row) and row[i] == row[i + 1]:
                        merged.append(row[i] * 2); self.score += row[i] * 2; i += 2
                    else:
                        merged.append(row[i]); i += 1
                new_grid[r] = merged + [0] * (4 - len(merged))
        elif direction == "right":
            for r in range(4):
                row = [x for x in new_grid[r] if x != 0]
                merged = []; i = len(row) - 1
                while i >= 0:
                    if i - 1 >= 0 and row[i] == row[i - 1]:
                        merged.insert(0, row[i] * 2); self.score += row[i] * 2; i -= 2
                    else:
                        merged.insert(0, row[i]); i -= 1
                new_grid[r] = [0] * (4 - len(merged)) + merged
        elif direction == "up":
            for c in range(4):
                col = [new_grid[r][c] for r in range(4) if new_grid[r][c] != 0]
                merged = []; i = 0
                while i < len(col):
                    if i + 1 < len(col) and col[i] == col[i + 1]:
                        merged.append(col[i] * 2); self.score += col[i] * 2; i += 2
                    else:
                        merged.append(col[i]); i += 1
                merged += [0] * (4 - len(merged))
                for r in range(4): new_grid[r][c] = merged[r]
        elif direction == "down":
            for c in range(4):
                col = [new_grid[r][c] for r in range(4) if new_grid[r][c] != 0]
                merged = []; i = len(col) - 1
                while i >= 0:
                    if i - 1 >= 0 and col[i] == col[i - 1]:
                        merged.insert(0, col[i] * 2); self.score += col[i] * 2; i -= 2
                    else:
                        merged.insert(0, col[i]); i -= 1
                merged = [0] * (4 - len(merged)) + merged
                for r in range(4): new_grid[r][c] = merged[r]
        if new_grid != self.grid: moved = True
        if moved:
            self.grid = new_grid
            self.spawn_tile()
            if not self.can_move(): self.game_over = True
            if any(self.grid[r][c] >= 2048 for r in range(4) for c in range(4)): self.won = True
            if self.score > self.best_score: self.best_score = self.score
        self.update()

    def keyPressEvent(self, event):
        key = event.key()
        if key in (Qt.Key.Key_Left, Qt.Key.Key_A): self.move("left")
        elif key in (Qt.Key.Key_Right, Qt.Key.Key_D): self.move("right")
        elif key in (Qt.Key.Key_Up, Qt.Key.Key_W): self.move("up")
        elif key in (Qt.Key.Key_Down, Qt.Key.Key_S): self.move("down")
        elif key == Qt.Key.Key_R: self.reset_game()
        self.setFocus()

    def reset_game(self):
        self.grid = [[0] * 4 for _ in range(4)]; self.score = 0
        self.game_over = False; self.won = False
        self.spawn_tile(); self.spawn_tile(); self.update()

    def paintEvent(self, event):
        painter = QPainter(self); painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.fillRect(self.rect(), self.BG_COLOR)
        for x_line in range(0, self.width(), 40):
            painter.setPen(QColor(18, 14, 30)); painter.drawLine(x_line, 0, x_line, self.height())
        for y_line in range(0, self.height(), 40):
            painter.setPen(QColor(18, 14, 30)); painter.drawLine(0, y_line, self.width(), y_line)
        painter.setPen(self.NEON_CYAN); painter.setFont(self.font_title)
        painter.drawText(40, 45, "2048 // NEON EDITION")
        painter.setPen(self.NEON_PINK); painter.drawLine(40, 60, 40 + 4 * (self.tile_size + self.tile_margin), 60)
        painter.setPen(self.TEXT_WHITE); painter.setFont(self.font_score)
        painter.drawText(40, 100, f"СЧЁТ: {self.score}")
        painter.setPen(self.TEXT_MUTED)
        painter.drawText(280, 100, f"РЕКОРД: {self.best_score}")
        painter.setPen(Qt.PenStyle.NoPen)
        for r in range(4):
            for c in range(4):
                x = self.grid_origin_x + c * (self.tile_size + self.tile_margin)
                y = self.grid_origin_y + r * (self.tile_size + self.tile_margin)
                painter.setBrush(QColor(25, 20, 40))
                painter.drawRoundedRect(x, y, self.tile_size, self.tile_size, 8, 8)
                val = self.grid[r][c]
                if val > 0:
                    color = self.TILE_COLORS.get(val, QColor(0, 240, 255))
                    painter.setBrush(color)
                    painter.drawRoundedRect(x + 2, y + 2, self.tile_size - 4, self.tile_size - 4, 6, 6)
                    text_color = QColor(15, 12, 25) if val >= 128 else QColor(245, 245, 255)
                    painter.setPen(text_color)
                    font = self.font_tile_small if val >= 128 else self.font_tile
                    painter.setFont(font)
                    text = str(val)
                    fm = painter.fontMetrics()
                    tw = fm.horizontalAdvance(text); th = fm.height()
                    painter.drawText(int(x + (self.tile_size - tw) / 2),
                                     int(y + (self.tile_size + th) / 2 - 4), text)
        if self.game_over:
            overlay = QColor(12, 10, 20, 220); painter.fillRect(self.rect(), overlay)
            painter.setPen(QColor(255, 40, 80)); painter.setFont(self.font_title)
            go_text = "ИГРА ОКОНЧЕНА"
            rect = painter.fontMetrics().boundingRect(go_text)
            painter.drawText((self.width() - rect.width()) // 2, self.height() // 2 - 20, go_text)
            painter.setPen(self.TEXT_MUTED); painter.setFont(self.font_score)
            restart = "Нажмите R для новой игры"
            painter.drawText((self.width() - painter.fontMetrics().boundingRect(restart).width()) // 2,
                             self.height() // 2 + 20, restart)
        elif self.won:
            painter.setPen(self.NEON_CYAN); painter.setFont(self.font_score)
            painter.drawText(40, 130, "🎉 ВЫ ДОСТИГЛИ 2048! Нажмите R для новой игры")
        painter.end()


# ============================================================================
#  КАЛЬКУЛЯТОР
# ============================================================================

class CalculatorTab(QWidget):
    def __init__(self):
        super().__init__()
        self.memory = 0.0; self.history = []; self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15); layout.setSpacing(10)
        self.display = QLineEdit()
        self.display.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.display.setReadOnly(True)
        self.display.setStyleSheet("QLineEdit { background-color: #1e1e1e; border: 2px solid #3c3c3c; border-radius: 6px; font-size: 24px; padding: 8px 12px; min-height: 50px; color: #ffffff; }")
        layout.addWidget(self.display)
        grid = QGridLayout(); grid.setSpacing(8)
        buttons = [
            ('C', 0, 0, 1, 1, {'clear': 'true'}), ('CE', 0, 1, 1, 1, {'clear': 'true'}),
            ('⌫', 0, 2, 1, 1, {}), ('÷', 0, 3, 1, 1, {'operator': 'true'}),
            ('7', 1, 0, 1, 1, {}), ('8', 1, 1, 1, 1, {}), ('9', 1, 2, 1, 1, {}),
            ('×', 1, 3, 1, 1, {'operator': 'true'}),
            ('4', 2, 0, 1, 1, {}), ('5', 2, 1, 1, 1, {}), ('6', 2, 2, 1, 1, {}),
            ('-', 2, 3, 1, 1, {'operator': 'true'}),
            ('1', 3, 0, 1, 1, {}), ('2', 3, 1, 1, 1, {}), ('3', 3, 2, 1, 1, {}),
            ('+', 3, 3, 1, 1, {'operator': 'true'}),
            ('0', 4, 0, 1, 2, {}), ('.', 4, 2, 1, 1, {}),
            ('=', 4, 3, 1, 1, {'equals': 'true'}),
        ]
        for btn in buttons:
            text, row, col, rowspan, colspan = btn[0], btn[1], btn[2], btn[3], btn[4]
            props = btn[5] if len(btn) > 5 else {}
            button = QPushButton(text); button.setObjectName("calcButton")
            for key, value in props.items(): button.setProperty(key, value)
            button.clicked.connect(self.on_button_clicked)
            grid.addWidget(button, row, col, rowspan, colspan)
        memory_buttons = [('MC', 5, 0), ('MR', 5, 1), ('M+', 5, 2), ('M-', 5, 3)]
        for text, row, col in memory_buttons:
            button = QPushButton(text); button.setObjectName("calcButton")
            button.clicked.connect(self.on_memory_clicked)
            grid.addWidget(button, row, col)
        layout.addLayout(grid)
        hist_group = QGroupBox("История вычислений"); hist_layout = QVBoxLayout(hist_group)
        self.history_list = QListWidget(); hist_layout.addWidget(self.history_list)
        hist_buttons = QHBoxLayout()
        self.clear_hist_btn = QPushButton("Очистить историю")
        self.clear_hist_btn.clicked.connect(self.clear_history)
        hist_buttons.addStretch(); hist_buttons.addWidget(self.clear_hist_btn)
        hist_layout.addLayout(hist_buttons); layout.addWidget(hist_group)
        self.memory_label = QLabel("Память: 0"); self.memory_label.setObjectName("infoLabel")
        self.memory_label.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(self.memory_label)
        load_history(self, "history/calculator", self.history_list)

    def on_button_clicked(self):
        button = self.sender(); text = button.text()
        if text == 'C': self.display.clear()
        elif text == 'CE':
            current = self.display.text(); operators = ['+', '-', '×', '÷']
            last_op_pos = -1
            for op in operators:
                pos = current.rfind(op)
                if pos > last_op_pos: last_op_pos = pos
            if last_op_pos != -1: self.display.setText(current[:last_op_pos + 1])
            else: self.display.clear()
        elif text == '⌫':
            current = self.display.text()
            if current: self.display.setText(current[:-1])
        elif text == '=': self.calculate_result()
        else:
            current = self.display.text()
            if current.endswith('='): self.display.clear()
            if text == '×': text = '*'
            elif text == '÷': text = '/'
            self.display.setText(self.display.text() + text)

    def on_memory_clicked(self):
        button = self.sender(); text = button.text(); current_text = self.display.text()
        try:
            if text == 'MC': self.memory = 0.0; self.update_memory_label()
            elif text == 'MR': self.display.setText(current_text + str(self.memory))
            elif text == 'M+':
                result = self.evaluate_safe(current_text)
                if result is not None: self.memory += result; self.update_memory_label()
            elif text == 'M-':
                result = self.evaluate_safe(current_text)
                if result is not None: self.memory -= result; self.update_memory_label()
        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Ошибка при работе с памятью: {str(e)}")

    def evaluate_safe(self, expr):
        if not expr: return None
        try:
            expr_eval = expr.replace('×', '*').replace('÷', '/')
            return safe_eval(expr_eval)
        except Exception: return None

    def calculate_result(self):
        current = self.display.text()
        if not current: return
        if current.endswith('='): current = current[:-1]
        expr_display = current
        try:
            result = self.evaluate_safe(current)
            if result is None: raise ValueError("Некорректное выражение")
            if isinstance(result, float) and result.is_integer(): result_str = str(int(result))
            else: result_str = format_number(result, 10, self).rstrip('0').rstrip('.')
            self.display.setText(result_str)
            self.history.append(f"{expr_display} = {result_str}")
            self.history_list.addItem(f"{expr_display} = {result_str}")
            save_history(self, "history/calculator", self.history_list)
        except Exception as e:
            QMessageBox.warning(self, "Ошибка", str(e))

    def clear_history(self):
        self.history.clear()
        clear_history(self, "history/calculator", self.history_list)

    def update_memory_label(self):
        self.memory_label.setText(f"Память: {self.memory}")


# ============================================================================
#  СОРТАМЕНТ ГОСТ
# ============================================================================

class SortamentTab(QWidget):
    def __init__(self):
        super().__init__()
        self.current_data = None; self.param_fields = {}
        self.current_mass_per_m = None
        self.init_ui()

    def init_ui(self):
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15); main_layout.setSpacing(15)
        left_widget = QWidget(); layout = QVBoxLayout(left_widget)
        layout.setContentsMargins(0, 0, 0, 0); layout.setSpacing(10)
        info_group = QGroupBox("Встроенный сортамент ГОСТ")
        info_layout = QVBoxLayout(info_group)
        info_text = QLabel("Выберите тип проката и размер из встроенной таблицы. "
                          "Масса 1 м и параметры подставятся автоматически.")
        info_text.setObjectName("infoLabel"); info_text.setWordWrap(True)
        info_layout.addWidget(info_text); layout.addWidget(info_group)
        selection_group = QGroupBox("Тип проката"); selection_layout = QHBoxLayout(selection_group)
        self.product_combo = QComboBox()
        self.product_combo.addItems(list(SORTAMENT_DATA.keys()))
        self.product_combo.currentIndexChanged.connect(self.on_product_changed)
        selection_layout.addWidget(QLabel("Сортамент:"))
        selection_layout.addWidget(self.product_combo, stretch=1)
        layout.addWidget(selection_group)
        size_group = QGroupBox("Размер"); self.size_layout = QGridLayout(size_group)
        self.size_layout.setHorizontalSpacing(8); self.size_layout.setVerticalSpacing(6)
        layout.addWidget(size_group)
        result_group = QGroupBox("Результат"); result_layout = QVBoxLayout(result_group)
        self.mass_per_m_label = QLabel("Масса 1 м: — кг"); self.mass_per_m_label.setObjectName("resultLabel")
        result_layout.addWidget(self.mass_per_m_label)
        self.params_label = QLabel("Параметры: —"); self.params_label.setObjectName("resultLabel")
        result_layout.addWidget(self.params_label); layout.addWidget(result_group)
        len_group = QGroupBox("Длина проката"); len_layout = QHBoxLayout(len_group)
        len_layout.addWidget(QLabel("Длина, м:")); self.length_edit = QLineEdit("6")
        self.length_edit.setValidator(create_c_locale_double_validator(0.0, 1e6, 3))
        self.length_edit.setMinimumHeight(28)
        self.length_edit.setToolTip("Длина проката, м. Общая масса = масса 1 м × длина.")
        self.length_edit.textChanged.connect(lambda: mark_invalid(self.length_edit, False))
        len_layout.addWidget(self.length_edit, stretch=1)
        layout.addWidget(len_group)
        btn_layout = QHBoxLayout(); btn_layout.addStretch()
        self.calculate_btn = QPushButton("🧮 Рассчитать массу"); self.calculate_btn.clicked.connect(self.calculate)
        btn_layout.addWidget(self.calculate_btn); btn_layout.addStretch()
        layout.addLayout(btn_layout)
        self.total_mass_label = QLabel("Общая масса: — кг"); self.total_mass_label.setObjectName("resultLabel")
        self.total_mass_label.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(self.total_mass_label)
        layout.addStretch(); main_layout.addWidget(left_widget, stretch=3)
        history_group = QGroupBox("История расчетов"); history_layout = QVBoxLayout(history_group)
        self.history_list = QListWidget(); self.history_list.setWordWrap(True)
        history_layout.addWidget(self.history_list, stretch=1)
        self.clear_history_btn = QPushButton("🗑 Очистить историю")
        self.clear_history_btn.clicked.connect(self.clear_history)
        history_layout.addWidget(self.clear_history_btn)
        history_group.setMinimumWidth(280); history_group.setMaximumWidth(380)
        main_layout.addWidget(history_group, stretch=1)
        self.on_product_changed(0)
        load_history(self, "history/sortament", self.history_list)

    def on_product_changed(self, index):
        key = list(SORTAMENT_DATA.keys())[index]
        self.current_data = SORTAMENT_DATA[key]
        while self.size_layout.count():
            item = self.size_layout.takeAt(0); w = item.widget()
            if w is not None: w.deleteLater()
        self.param_fields = {}
        for attr in ("D_combo", "s_combo", "h_combo", "A_combo", "t_combo", "d_combo", "b_label"):
            if hasattr(self, attr):
                delattr(self, attr)
        if self.current_data["type"] == "pipe":
            self.size_layout.addWidget(QLabel("Наружный диаметр D, мм:"), 0, 0)
            self.D_combo = QComboBox()
            diameters = sorted(set(row[0] for row in self.current_data["data"]))
            self.D_combo.addItems([str(d) for d in diameters])
            self.D_combo.currentIndexChanged.connect(self.on_D_changed)
            self.size_layout.addWidget(self.D_combo, 0, 1)
            self.size_layout.addWidget(QLabel("Толщина стенки s, мм:"), 0, 2)
            self.s_combo = QComboBox(); self.size_layout.addWidget(self.s_combo, 0, 3)
            self.s_combo.currentIndexChanged.connect(self.update_info)
            self.on_D_changed(0)
        elif self.current_data["type"] in ("channel", "beam"):
            self.size_layout.addWidget(QLabel("Высота h, мм:"), 0, 0)
            self.h_combo = QComboBox()
            heights = sorted(set(row[0] for row in self.current_data["data"]))
            self.h_combo.addItems([str(h) for h in heights])
            self.h_combo.currentIndexChanged.connect(self.update_info)
            self.size_layout.addWidget(self.h_combo, 0, 1)
            self.size_layout.addWidget(QLabel("Ширина полки b, мм:"), 0, 2)
            self.b_label = QLabel("—"); self.b_label.setObjectName("unitLabel")
            self.size_layout.addWidget(self.b_label, 0, 3)
        elif self.current_data["type"] == "angle":
            self.size_layout.addWidget(QLabel("Полка A, мм:"), 0, 0)
            self.A_combo = QComboBox()
            sides = sorted(set(row[0] for row in self.current_data["data"]))
            self.A_combo.addItems([str(a) for a in sides])
            self.A_combo.currentIndexChanged.connect(self.on_A_changed)
            self.size_layout.addWidget(self.A_combo, 0, 1)
            self.size_layout.addWidget(QLabel("Толщина t, мм:"), 0, 2)
            self.t_combo = QComboBox(); self.size_layout.addWidget(self.t_combo, 0, 3)
            self.t_combo.currentIndexChanged.connect(self.update_info)
            self.on_A_changed(0)
        elif self.current_data["type"] == "circle":
            self.size_layout.addWidget(QLabel("Диаметр d, мм:"), 0, 0)
            self.d_combo = QComboBox()
            diams = sorted(set(row[0] for row in self.current_data["data"]))
            self.d_combo.addItems([str(d) for d in diams])
            self.d_combo.currentIndexChanged.connect(self.update_info)
            self.size_layout.addWidget(self.d_combo, 0, 1)
        self.update_info()

    def on_D_changed(self, idx):
        if not hasattr(self, 'D_combo'): return
        d = float(self.D_combo.currentText())
        thicknesses = sorted(set(row[1] for row in self.current_data["data"] if row[0] == d))
        self.s_combo.blockSignals(True); self.s_combo.clear()
        self.s_combo.addItems([str(s) for s in thicknesses])
        self.s_combo.blockSignals(False); self.update_info()

    def on_A_changed(self, idx):
        if not hasattr(self, 'A_combo'): return
        a = float(self.A_combo.currentText())
        thicknesses = sorted(set(row[1] for row in self.current_data["data"] if row[0] == a))
        self.t_combo.blockSignals(True); self.t_combo.clear()
        self.t_combo.addItems([str(t) for t in thicknesses])
        self.t_combo.blockSignals(False); self.update_info()

    def update_info(self):
        if self.current_data is None: return
        self.current_mass_per_m = None
        try:
            if self.current_data["type"] == "pipe":
                d = float(self.D_combo.currentText()); s = float(self.s_combo.currentText())
                for row in self.current_data["data"]:
                    if row[0] == d and row[1] == s:
                        self.mass_per_m_label.setText(f"Масса 1 м: {row[2]:.3f} кг")
                        self.params_label.setText(f"D={d} мм, s={s} мм")
                        self.current_mass_per_m = row[2]; break
            elif self.current_data["type"] in ("channel", "beam"):
                h = float(self.h_combo.currentText())
                for row in self.current_data["data"]:
                    if row[0] == h:
                        self.b_label.setText(f"{row[1]} мм")
                        self.mass_per_m_label.setText(f"Масса 1 м: {row[4]:.3f} кг")
                        self.params_label.setText(f"h={h} мм, b={row[1]} мм, s={row[2]} мм, t={row[3]} мм")
                        self.current_mass_per_m = row[4]; break
            elif self.current_data["type"] == "angle":
                a = float(self.A_combo.currentText()); t = float(self.t_combo.currentText())
                for row in self.current_data["data"]:
                    if row[0] == a and row[1] == t:
                        self.mass_per_m_label.setText(f"Масса 1 м: {row[2]:.3f} кг")
                        self.params_label.setText(f"A={a} мм, t={t} мм")
                        self.current_mass_per_m = row[2]; break
            elif self.current_data["type"] == "circle":
                d = float(self.d_combo.currentText())
                for row in self.current_data["data"]:
                    if row[0] == d:
                        self.mass_per_m_label.setText(f"Масса 1 м: {row[1]:.3f} кг")
                        self.params_label.setText(f"d={d} мм")
                        self.current_mass_per_m = row[1]; break
        except (ValueError, AttributeError):
            pass

    def calculate(self):
        if self.current_mass_per_m is None:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите размер из сортамента")
            return
        if not self.length_edit.text().strip():
            mark_invalid(self.length_edit, True)
            QMessageBox.warning(self, "Ошибка ввода", "Введите длину")
            return
        try:
            length = float(self.length_edit.text().strip().replace(',', '.'))
            if length <= 0: raise ValueError
        except ValueError:
            mark_invalid(self.length_edit, True)
            QMessageBox.warning(self, "Ошибка ввода", "Введите положительную длину")
            return
        total = self.current_mass_per_m * length
        self.total_mass_label.setText(f"Общая масса: {format_number(total, 3, self)} кг")
        product = self.product_combo.currentText()
        ts = datetime.now().strftime("%H:%M:%S")
        self.history_list.addItem(f"[{ts}] {product}\n"
                                  f"  {self.params_label.text()}, L={length} м\n"
                                  f"  → {format_number(total, 3, self)} кг")
        self.history_list.scrollToBottom()
        save_history(self, "history/sortament", self.history_list)

    def clear_history(self):
        clear_history(self, "history/sortament", self.history_list)


# ============================================================================
#  ПРОФИЛИ ГКЛ И ПЕРФОРАЦИЯ
# ============================================================================

class ProfilesTab(QWidget):
    def __init__(self):
        super().__init__()
        self.current_sizes = []
        self.current_mass_per_m = None
        self.init_ui()

    def init_ui(self):
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)

        left_widget = QWidget()
        layout = QVBoxLayout(left_widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        info_group = QGroupBox("Профили для ГКЛ и перфорированные изделия")
        info_layout = QVBoxLayout(info_group)
        info_text = QLabel(
            "Расчёт массы профилей для гипсокартона, защитных уголков "
            "и перфорированных изделий. Выберите тип и типоразмер, укажите "
            "длину одной штуки и количество — программа посчитает общую массу.\n"
            "Массы 1 м — справочные, по данным производителей (Кнауф/Gyproc)."
        )
        info_text.setObjectName("infoLabel")
        info_text.setWordWrap(True)
        info_layout.addWidget(info_text)
        layout.addWidget(info_group)

        type_group = QGroupBox("Тип профиля / изделия")
        type_layout = QVBoxLayout(type_group)
        self.profile_combo = QComboBox()
        self.profile_combo.addItems(list(PROFILES_DATA.keys()))
        self.profile_combo.currentIndexChanged.connect(self.on_profile_changed)
        type_layout.addWidget(self.profile_combo)
        layout.addWidget(type_group)

        size_group = QGroupBox("Типоразмер")
        size_layout = QGridLayout(size_group)
        size_layout.addWidget(QLabel("Размер:"), 0, 0)
        self.size_combo = QComboBox()
        self.size_combo.currentIndexChanged.connect(self.on_size_changed)
        size_layout.addWidget(self.size_combo, 0, 1)
        self.desc_label = QLabel("—")
        self.desc_label.setObjectName("infoLabel")
        self.desc_label.setWordWrap(True)
        size_layout.addWidget(self.desc_label, 1, 0, 1, 2)
        layout.addWidget(size_group)

        mass_group = QGroupBox("Справочные данные")
        mass_layout = QVBoxLayout(mass_group)
        self.mass_per_m_label = QLabel("Масса 1 м: — кг")
        self.mass_per_m_label.setObjectName("resultLabel")
        mass_layout.addWidget(self.mass_per_m_label)
        self.geom_label = QLabel("Параметры: —")
        self.geom_label.setObjectName("infoLabel")
        mass_layout.addWidget(self.geom_label)
        layout.addWidget(mass_group)

        input_group = QGroupBox("Длина и количество")
        input_layout = QGridLayout(input_group)
        input_layout.addWidget(QLabel("Длина одной штуки, м:"), 0, 0)
        self.length_edit = QLineEdit("3")
        self.length_edit.setValidator(create_c_locale_double_validator(0.0, 1e6, 3))
        self.length_edit.setMinimumHeight(28)
        self.length_edit.setToolTip("Стандартная длина хлыста, м (обычно 3 или 4).")
        self.length_edit.textChanged.connect(lambda: mark_invalid(self.length_edit, False))
        input_layout.addWidget(self.length_edit, 0, 1)
        input_layout.addWidget(QLabel("Количество, шт:"), 1, 0)
        self.qty_edit = QLineEdit("1")
        self.qty_edit.setValidator(QIntValidator(1, 999999))
        self.qty_edit.setMinimumHeight(28)
        self.qty_edit.setToolTip("Количество штук (хлыстов).")
        self.qty_edit.textChanged.connect(lambda: mark_invalid(self.qty_edit, False))
        input_layout.addWidget(self.qty_edit, 1, 1)
        input_layout.addWidget(QLabel("Общая длина, м:"), 2, 0)
        self.total_length_label = QLabel("—")
        self.total_length_label.setObjectName("resultLabel")
        input_layout.addWidget(self.total_length_label, 2, 1)
        layout.addWidget(input_group)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self.calculate_btn = QPushButton("🧮 Рассчитать массу")
        self.calculate_btn.clicked.connect(self.calculate)
        btn_layout.addWidget(self.calculate_btn)
        self.clear_btn = QPushButton("🗑 Очистить")
        self.clear_btn.clicked.connect(self.clear_fields)
        btn_layout.addWidget(self.clear_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.total_mass_label = QLabel("Общая масса: — кг")
        self.total_mass_label.setObjectName("resultLabel")
        self.total_mass_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.total_mass_label)

        layout.addStretch()
        main_layout.addWidget(left_widget, stretch=3)

        history_group = QGroupBox("История расчетов")
        history_layout = QVBoxLayout(history_group)
        self.history_list = QListWidget()
        self.history_list.setWordWrap(True)
        history_layout.addWidget(self.history_list, stretch=1)
        self.clear_history_btn = QPushButton("🗑 Очистить историю")
        self.clear_history_btn.clicked.connect(self.clear_history)
        history_layout.addWidget(self.clear_history_btn)
        history_group.setMinimumWidth(280)
        history_group.setMaximumWidth(380)
        main_layout.addWidget(history_group, stretch=1)

        self.on_profile_changed(0)
        load_history(self, "history/profiles", self.history_list)

    def on_profile_changed(self, index):
        if index < 0:
            return
        key = list(PROFILES_DATA.keys())[index]
        self.current_sizes = PROFILES_DATA[key]["sizes"]
        self.desc_label.setText(PROFILES_DATA[key]["description"])
        self.size_combo.blockSignals(True)
        self.size_combo.clear()
        for item in self.current_sizes:
            self.size_combo.addItem(item[0])
        self.size_combo.blockSignals(False)
        if self.current_sizes:
            self.size_combo.setCurrentIndex(0)
            self.on_size_changed(0)

    def on_size_changed(self, index):
        if index < 0 or index >= len(self.current_sizes):
            return
        label, b, h, s, mass = self.current_sizes[index]
        self.current_mass_per_m = mass
        self.mass_per_m_label.setText(f"Масса 1 м: {mass:.3f} кг")
        if h > 0:
            self.geom_label.setText(
                f"Ширина b={b} мм, высота h={h} мм, толщина s={s} мм")
        else:
            self.geom_label.setText(
                f"Ширина b={b} мм, толщина s={s} мм (плоская перфолента)")

    def calculate(self):
        if self.current_mass_per_m is None:
            QMessageBox.warning(self, "Ошибка", "Выберите типоразмер")
            return
        if not self.length_edit.text().strip():
            mark_invalid(self.length_edit, True)
            QMessageBox.warning(self, "Ошибка ввода", "Введите длину")
            return
        try:
            length = float(self.length_edit.text().strip().replace(',', '.'))
            if length <= 0:
                raise ValueError
        except ValueError:
            mark_invalid(self.length_edit, True)
            QMessageBox.warning(self, "Ошибка ввода", "Введите положительную длину")
            return
        try:
            qty = int(float(self.qty_edit.text().strip().replace(',', '.') or "0"))
            if qty <= 0:
                raise ValueError
        except ValueError:
            mark_invalid(self.qty_edit, True)
            QMessageBox.warning(self, "Ошибка ввода", "Введите положительное количество")
            return

        total_length = length * qty
        total_mass = self.current_mass_per_m * total_length

        self.total_length_label.setText(f"{total_length:.3f} м")
        self.total_mass_label.setText(
            f"Общая масса: {format_number(total_mass, 3, self)} кг")

        profile_name = self.profile_combo.currentText()
        size_name = self.size_combo.currentText()
        ts = datetime.now().strftime("%H:%M:%S")
        self.history_list.addItem(
            f"[{ts}] {profile_name}\n"
            f"  {size_name}, L={length} м × {qty} шт = {total_length:.2f} м\n"
            f"  → {format_number(total_mass, 3, self)} кг")
        self.history_list.scrollToBottom()
        save_history(self, "history/profiles", self.history_list)

    def clear_fields(self):
        self.length_edit.setText("3")
        self.qty_edit.setText("1")
        self.total_length_label.setText("—")
        self.total_mass_label.setText("Общая масса: — кг")
        mark_invalid(self.length_edit, False)
        mark_invalid(self.qty_edit, False)

    def clear_history(self):
        clear_history(self, "history/profiles", self.history_list)


# ============================================================================
#  НОРМЫ АКЗ (библиотека)
# ============================================================================

class AkzLibraryTab(QWidget):
    def __init__(self):
        super().__init__(); self.init_ui()

    def init_ui(self):
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15); main_layout.setSpacing(15)
        left_widget = QWidget(); layout = QVBoxLayout(left_widget)
        layout.setContentsMargins(0, 0, 0, 0); layout.setSpacing(10)
        info_group = QGroupBox("Встроенная библиотека норм ЛКМ")
        info_layout = QVBoxLayout(info_group)
        info_text = QLabel("Нормы расхода лакокрасочных материалов для антикоррозионной защиты. "
                          "Выберите материал и укажите площадь окраски.")
        info_text.setObjectName("infoLabel"); info_text.setWordWrap(True)
        info_layout.addWidget(info_text); layout.addWidget(info_group)
        self.table = QTableWidget(); self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["Материал", "Категория", "Расход, кг/м²", "ТСП", "Слоёв", "Интервал"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        for row, mat in enumerate(AKZ_STANDARDS):
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(mat["name"]))
            self.table.setItem(row, 1, QTableWidgetItem(mat["category"]))
            self.table.setItem(row, 2, QTableWidgetItem(f'{mat["consumption_min"]:.3f}–{mat["consumption_max"]:.3f}'))
            self.table.setItem(row, 3, QTableWidgetItem(mat["dry_thickness"]))
            self.table.setItem(row, 4, QTableWidgetItem(str(mat["layers"])))
            self.table.setItem(row, 5, QTableWidgetItem(mat["interval"]))
        self.table.selectRow(0)
        layout.addWidget(self.table, stretch=1)
        param_group = QGroupBox("Параметры расчёта"); param_layout = QGridLayout(param_group)
        param_layout.addWidget(QLabel("Площадь окраски, м²:"), 0, 0)
        self.area_edit = QLineEdit("100")
        self.area_edit.setValidator(create_c_locale_double_validator(0.0, 1e9, 6))
        self.area_edit.setMinimumHeight(28)
        self.area_edit.setToolTip("Площадь окраски в м².\nРасход = S × норма_ср × слои")
        self.area_edit.textChanged.connect(lambda: mark_invalid(self.area_edit, False))
        param_layout.addWidget(self.area_edit, 0, 1)
        param_layout.addWidget(QLabel("Количество слоёв:"), 1, 0)
        self.layers_edit = QLineEdit(""); self.layers_edit.setPlaceholderText("по умолчанию")
        self.layers_edit.setValidator(QIntValidator(1, 20)); self.layers_edit.setMinimumHeight(28)
        self.layers_edit.setToolTip("Оставьте пустым — возьмётся из таблицы ЛКМ.")
        param_layout.addWidget(self.layers_edit, 1, 1); layout.addWidget(param_group)
        btn_layout = QHBoxLayout(); btn_layout.addStretch()
        self.calc_btn = QPushButton("🧮 Рассчитать"); self.calc_btn.clicked.connect(self.calculate)
        btn_layout.addWidget(self.calc_btn); btn_layout.addStretch()
        layout.addLayout(btn_layout)
        self.result_label = QLabel("Расход: — кг"); self.result_label.setObjectName("resultLabel")
        self.result_label.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(self.result_label)
        layout.addStretch(); main_layout.addWidget(left_widget, stretch=3)
        history_group = QGroupBox("История расчетов"); history_layout = QVBoxLayout(history_group)
        self.history_list = QListWidget(); self.history_list.setWordWrap(True)
        history_layout.addWidget(self.history_list, stretch=1)
        self.clear_history_btn = QPushButton("🗑 Очистить историю")
        self.clear_history_btn.clicked.connect(self.clear_history)
        history_layout.addWidget(self.clear_history_btn)
        history_group.setMinimumWidth(280); history_group.setMaximumWidth(380)
        main_layout.addWidget(history_group, stretch=1)
        load_history(self, "history/akz_lib", self.history_list)

    def calculate(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Ошибка", "Выберите материал из таблицы"); return
        mat = AKZ_STANDARDS[row]
        if not self.area_edit.text().strip():
            mark_invalid(self.area_edit, True)
            QMessageBox.warning(self, "Ошибка ввода", "Введите площадь"); return
        try:
            area = float(self.area_edit.text().strip().replace(',', '.') or "0")
            if area <= 0: raise ValueError
        except ValueError:
            mark_invalid(self.area_edit, True)
            QMessageBox.warning(self, "Ошибка ввода", "Введите положительную площадь"); return
        layers_text = self.layers_edit.text().strip()
        if layers_text:
            layers = int(float(layers_text.replace(',', '.')))
            if layers <= 0: layers = mat["layers"]
        else: layers = mat["layers"]
        avg_consumption = (mat["consumption_min"] + mat["consumption_max"]) / 2
        total = avg_consumption * area * layers
        self.result_label.setText(f"Расход: {format_number(total, 3, self)} кг")
        ts = datetime.now().strftime("%H:%M:%S")
        self.history_list.addItem(f"[{ts}] ЛКМ: {mat['name']}\n"
                                  f"  S={area} м², слоёв={layers}, норма={avg_consumption:.3f} кг/м²\n"
                                  f"  → {format_number(total, 3, self)} кг")
        self.history_list.scrollToBottom()
        save_history(self, "history/akz_lib", self.history_list)

    def clear_history(self):
        clear_history(self, "history/akz_lib", self.history_list)


# ============================================================================
#  ПЕСКОСТРУЙ
# ============================================================================

class SandblastTab(QWidget):
    def __init__(self):
        super().__init__(); self.init_ui()

    def init_ui(self):
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15); main_layout.setSpacing(15)
        left_widget = QWidget(); layout = QVBoxLayout(left_widget)
        layout.setContentsMargins(0, 0, 0, 0); layout.setSpacing(10)

        info_group = QGroupBox("Пескоструйная (абразивоструйная) очистка")
        info_layout = QVBoxLayout(info_group)
        info_text = QLabel(
            "Расчёт расхода абразива и трудоёмкости очистки поверхности.\n"
            "Степени очистки по ISO 8501-1: Sa 1 … Sa 3.\n"
            "Расход = S × средняя_норма × коэффициент_абразива.\n"
            "Трудоёмкость = S × норма_времени × коэффициент_абразива."
        )
        info_text.setObjectName("infoLabel"); info_text.setWordWrap(True)
        info_layout.addWidget(info_text); layout.addWidget(info_group)

        param_group = QGroupBox("Параметры очистки")
        param_grid = QGridLayout(param_group)
        param_grid.setHorizontalSpacing(10); param_grid.setVerticalSpacing(8)

        param_grid.addWidget(QLabel("Степень очистки:"), 0, 0)
        self.degree_combo = QComboBox()
        for key, val in SANDBLAST_DEGREES.items():
            self.degree_combo.addItem(key, val)
        self.degree_combo.currentIndexChanged.connect(self.update_degree_info)
        self.degree_combo.setToolTip(
            "Степени очистки по ISO 8501-1:\n"
            "Sa 1 — лёгкая\n"
            "Sa 2 — тщательная\n"
            "Sa 2½ — очень тщательная (типовая для АКЗ)\n"
            "Sa 3 — белая (максимальная)"
        )
        param_grid.addWidget(self.degree_combo, 0, 1, 1, 3)

        param_grid.addWidget(QLabel("Тип абразива:"), 1, 0)
        self.abrasive_combo = QComboBox()
        for name, data in SANDBLAST_ABRASIVES.items():
            self.abrasive_combo.addItem(name, data)
        self.abrasive_combo.setToolTip(
            "Разные абразивы имеют разный расход, время и число оборотов.\n"
            "Стальная дробь — самая экономичная (многоразовая)."
        )
        self.abrasive_combo.currentIndexChanged.connect(self.update_degree_info)
        param_grid.addWidget(self.abrasive_combo, 1, 1, 1, 3)

        param_grid.addWidget(QLabel("Площадь очистки, м²:"), 2, 0)
        self.area_edit = QLineEdit("100")
        self.area_edit.setValidator(create_c_locale_double_validator(0.0, 1e9, 6))
        self.area_edit.setMinimumHeight(28)
        self.area_edit.setToolTip("Площадь поверхности в м².")
        self.area_edit.textChanged.connect(lambda: mark_invalid(self.area_edit, False))
        param_grid.addWidget(self.area_edit, 2, 1)
        param_grid.addWidget(QLabel("Кол-во проходов:"), 2, 2)
        self.passes_edit = QLineEdit("1")
        self.passes_edit.setValidator(QIntValidator(1, 10))
        self.passes_edit.setMinimumHeight(28)
        self.passes_edit.setToolTip("Сколько проходов по одной поверхности. Обычно 1.")
        self.passes_edit.textChanged.connect(lambda: mark_invalid(self.passes_edit, False))
        param_grid.addWidget(self.passes_edit, 2, 3)

        layout.addWidget(param_group)

        self.degree_info = QLabel("—")
        self.degree_info.setObjectName("infoLabel")
        self.degree_info.setWordWrap(True)
        layout.addWidget(self.degree_info)

        btn_layout = QHBoxLayout(); btn_layout.addStretch()
        self.calculate_btn = QPushButton("🧮 Рассчитать")
        self.calculate_btn.clicked.connect(self.calculate)
        btn_layout.addWidget(self.calculate_btn)
        self.clear_btn = QPushButton("🗑 Очистить")
        self.clear_btn.clicked.connect(self.clear_fields)
        btn_layout.addWidget(self.clear_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        result_group = QGroupBox("Результаты расчёта")
        result_layout = QVBoxLayout(result_group)
        result_layout.setContentsMargins(8, 8, 8, 8); result_layout.setSpacing(4)
        self.abrasive_per_m2_label = QLabel("Норма расхода абразива: — кг/м²")
        self.abrasive_per_m2_label.setObjectName("resultLabel")
        result_layout.addWidget(self.abrasive_per_m2_label)
        self.abrasive_total_label = QLabel("Расход абразива: — кг")
        self.abrasive_total_label.setObjectName("resultLabel")
        result_layout.addWidget(self.abrasive_total_label)
        self.time_per_m2_label = QLabel("Норма времени: — н/ч на м²")
        self.time_per_m2_label.setObjectName("resultLabel")
        result_layout.addWidget(self.time_per_m2_label)
        self.time_total_label = QLabel("Трудоёмкость: — н/ч")
        self.time_total_label.setObjectName("resultLabel")
        result_layout.addWidget(self.time_total_label)
        layout.addWidget(result_group)

        layout.addStretch()
        main_layout.addWidget(left_widget, stretch=3)

        history_group = QGroupBox("История расчетов")
        history_layout = QVBoxLayout(history_group)
        self.history_list = QListWidget(); self.history_list.setWordWrap(True)
        history_layout.addWidget(self.history_list, stretch=1)
        self.clear_history_btn = QPushButton("🗑 Очистить историю")
        self.clear_history_btn.clicked.connect(self.clear_history)
        history_layout.addWidget(self.clear_history_btn)
        history_group.setMinimumWidth(280); history_group.setMaximumWidth(380)
        main_layout.addWidget(history_group, stretch=1)

        load_history(self, "history/sandblast", self.history_list)
        self.update_degree_info()

    def update_degree_info(self):
        key = self.degree_combo.currentText()
        data = SANDBLAST_DEGREES.get(key, {})
        abras = self.abrasive_combo.currentText()
        a = SANDBLAST_ABRASIVES.get(abras, {})
        if data:
            self.degree_info.setText(
                f"ℹ️ {data['description']}.\n"
                f"Расход абразива (справочно): {data['abrasive_min']}–{data['abrasive_max']} кг/м²; "
                f"время: {data['time_per_m2']:.3f} н/ч на м².\n"
                f"Поправка на «{abras}»: расход ×{a.get('k_consumption', 1.0):.2f}, "
                f"время ×{a.get('k_time', 1.0):.2f}."
            )

    def calculate(self):
        invalid = []
        if not self.area_edit.text().strip():
            mark_invalid(self.area_edit, True); invalid.append(self.area_edit)
        if not self.passes_edit.text().strip():
            mark_invalid(self.passes_edit, True); invalid.append(self.passes_edit)
        if invalid:
            QMessageBox.warning(self, "Ошибка ввода", "Заполните выделенные поля")
            return
        try:
            area = float(self.area_edit.text().strip().replace(',', '.'))
            if area <= 0: raise ValueError
        except ValueError:
            mark_invalid(self.area_edit, True)
            QMessageBox.warning(self, "Ошибка ввода", "Введите положительную площадь")
            return
        try:
            passes = int(float(self.passes_edit.text().strip().replace(',', '.')))
            if passes < 1: passes = 1
        except ValueError:
            mark_invalid(self.passes_edit, True)
            QMessageBox.warning(self, "Ошибка ввода", "Количество проходов — целое число")
            return

        degree = self.degree_combo.currentText()
        abrasive = self.abrasive_combo.currentText()
        result = calculate_sandblasting(degree, abrasive, area)
        result["abrasive_kg"] *= passes
        result["time_h"] *= passes

        self.abrasive_per_m2_label.setText(
            f"Норма расхода абразива: {format_number(result['abrasive_per_m2'], 3, self)} кг/м²")
        self.abrasive_total_label.setText(
            f"Расход абразива: {format_number(result['abrasive_kg'], 3, self)} кг")
        self.time_per_m2_label.setText(
            f"Норма времени: {format_number(result['time_per_m2'], 4, self)} н/ч на м²")
        self.time_total_label.setText(
            f"Трудоёмкость: {format_number(result['time_h'], 3, self)} н/ч")

        ts = datetime.now().strftime("%H:%M:%S")
        self.history_list.addItem(
            f"[{ts}] {degree} | {abrasive}\n"
            f"  S={area} м², проходов={passes}\n"
            f"  → абразив {format_number(result['abrasive_kg'], 2, self)} кг; "
            f"труд {format_number(result['time_h'], 2, self)} н/ч")
        self.history_list.scrollToBottom()
        save_history(self, "history/sandblast", self.history_list)

    def clear_fields(self):
        self.area_edit.setText("100")
        self.passes_edit.setText("1")
        self.abrasive_per_m2_label.setText("Норма расхода абразива: — кг/м²")
        self.abrasive_total_label.setText("Расход абразива: — кг")
        self.time_per_m2_label.setText("Норма времени: — н/ч на м²")
        self.time_total_label.setText("Трудоёмкость: — н/ч")

    def clear_history(self):
        clear_history(self, "history/sandblast", self.history_list)


# ============================================================================
#  ИЗОЛЯЦИЯ
# ============================================================================

class InsulationTab(QWidget):
    def __init__(self):
        super().__init__(); self.init_ui()

    def init_ui(self):
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15); main_layout.setSpacing(15)
        left_widget = QWidget(); layout = QVBoxLayout(left_widget)
        layout.setContentsMargins(0, 0, 0, 0); layout.setSpacing(15)
        info_group = QGroupBox("Расчет объема тепловой изоляции труб")
        info_layout = QVBoxLayout(info_group)
        info_text = QLabel("Калькулятор позволяет рассчитать объем изоляции, площадь покровного слоя "
                          "и площадь обертывания для одной или нескольких труб.")
        info_text.setObjectName("infoLabel"); info_text.setWordWrap(True)
        info_layout.addWidget(info_text); layout.addWidget(info_group)
        type_group = QGroupBox("Тип расчета"); type_layout = QHBoxLayout(type_group)
        self.radio_single = QRadioButton("Одна труба")
        self.radio_multiple_v1 = QRadioButton("Несколько труб (Вариант 1 - три трубы)")
        self.radio_multiple_v2 = QRadioButton("Несколько труб (Вариант 2 - 2 и более труб)")
        self.radio_single.setChecked(True)
        type_layout.addWidget(self.radio_single); type_layout.addWidget(self.radio_multiple_v1)
        type_layout.addWidget(self.radio_multiple_v2); type_layout.addStretch()
        self.radio_single.toggled.connect(self.on_calc_type_changed)
        self.radio_multiple_v1.toggled.connect(self.on_calc_type_changed)
        self.radio_multiple_v2.toggled.connect(self.on_calc_type_changed)
        layout.addWidget(type_group)
        self.input_group = QGroupBox("Параметры")
        self.input_layout = QVBoxLayout(self.input_group)
        self.input_widgets = {}; self.create_input_fields()
        layout.addWidget(self.input_group)
        button_layout = QHBoxLayout(); button_layout.addStretch()
        self.calculate_btn = QPushButton("🧮 Рассчитать"); self.calculate_btn.clicked.connect(self.calculate)
        button_layout.addWidget(self.calculate_btn)
        self.clear_btn = QPushButton("Очистить"); self.clear_btn.clicked.connect(self.clear_fields)
        button_layout.addWidget(self.clear_btn); button_layout.addStretch()
        layout.addLayout(button_layout)
        result_group = QGroupBox("Результаты расчета"); result_layout = QVBoxLayout(result_group)
        self.Sr_label = QLabel("Площадь обертывания труб (Sr): — м²")
        self.Sr_label.setObjectName("resultLabel"); result_layout.addWidget(self.Sr_label)
        self.Spi_label = QLabel("Площадь покровного слоя изоляции (Spi): — м²")
        self.Spi_label.setObjectName("resultLabel"); result_layout.addWidget(self.Spi_label)
        self.Vi_label = QLabel("Объем изоляции (Vi): — м³")
        self.Vi_label.setObjectName("resultLabel"); result_layout.addWidget(self.Vi_label)
        layout.addWidget(result_group); layout.addStretch()
        main_layout.addWidget(left_widget, stretch=3)
        history_group = QGroupBox("История расчетов"); history_layout = QVBoxLayout(history_group)
        self.history_list = QListWidget(); self.history_list.setWordWrap(True)
        history_layout.addWidget(self.history_list, stretch=1)
        self.clear_history_btn = QPushButton("🗑 Очистить историю")
        self.clear_history_btn.clicked.connect(self.clear_history)
        history_layout.addWidget(self.clear_history_btn)
        history_group.setMinimumWidth(280); history_group.setMaximumWidth(380)
        main_layout.addWidget(history_group, stretch=1)
        load_history(self, "history/insulation", self.history_list)

    def create_input_fields(self):
        clear_layout(self.input_layout); self.input_widgets.clear()
        if self.radio_single.isChecked():
            fields = [("Диаметр трубы (D)", "м", "например, 0.125", "D",
                       "Наружный диаметр трубы, м.\nSr = π · D · L"),
                      ("Толщина изоляции (t)", "м", "например, 0.08", "t",
                       "Толщина слоя изоляции, м.\nVi = π · t · (D + t) · L"),
                      ("Длина участка изоляции (L)", "м", "например, 130", "L",
                       "Длина изолируемого участка, м.")]
        elif self.radio_multiple_v1.isChecked():
            fields = [("Диаметр крайних труб (D1)", "м", "например, 0.108", "D1",
                       "Диаметр крайних труб, м."),
                      ("Диаметр средней трубы (D2)", "м", "например, 0.076", "D2",
                       "Диаметр средней трубы, м."),
                      ("Толщина изоляции (t)", "м", "например, 0.1", "t",
                       "Толщина слоя изоляции, м."),
                      ("Расстояние между трубами (p)", "м", "например, 0.1", "p",
                       "Зазор между трубами, м."),
                      ("Длина участка изоляции (L)", "м", "например, 65", "L",
                       "Длина участка, м.")]
        else:
            fields = [("Диаметр крайних труб (D1)", "м", "например, 0.108", "D1",
                       "Диаметр крайних труб, м."),
                      ("Расстояние между осями крайних труб (M)", "м", "например, 0.384", "M",
                       "Расстояние между осями крайних труб, м."),
                      ("Толщина изоляции (t)", "м", "например, 0.1", "t",
                       "Толщина слоя изоляции, м."),
                      ("Длина участка изоляции (L)", "м", "например, 65", "L",
                       "Длина участка, м.")]
        for label, unit, placeholder, key, tooltip in fields:
            row, edit = create_input_row(label, unit, placeholder, tooltip=tooltip)
            self.input_layout.addLayout(row); self.input_widgets[key] = edit

    def on_calc_type_changed(self):
        self.create_input_fields(); self.clear_results()

    def get_float(self, key, default=0.0):
        edit = self.input_widgets.get(key)
        return get_float_from_edit(edit, default) if edit else default

    def _current_type_name(self):
        if self.radio_single.isChecked(): return "Одна труба"
        if self.radio_multiple_v1.isChecked(): return "Пучок (3 трубы)"
        return "Пучок (2+)"

    def calculate(self):
        missing = [e for e in self.input_widgets.values() if not e.text().strip()]
        for e in self.input_widgets.values():
            mark_invalid(e, not e.text().strip())
        if missing:
            QMessageBox.warning(self, "Ошибка ввода", "Заполните выделенные поля")
            return
        try:
            if self.radio_single.isChecked():
                D = self.get_float("D"); t = self.get_float("t"); L = self.get_float("L")
                if D <= 0 or t <= 0 or L <= 0: raise ValueError("Все значения должны быть положительными")
                result = calculate_single_pipe_insulation(D, t, L)
            elif self.radio_multiple_v1.isChecked():
                D1 = self.get_float("D1"); D2 = self.get_float("D2"); t = self.get_float("t")
                p = self.get_float("p"); L = self.get_float("L")
                if any(v <= 0 for v in (D1, D2, t, p, L)): raise ValueError("Все значения должны быть положительными")
                result = calculate_multiple_pipes_insulation_variant1(D1, D2, t, p, L)
            else:
                D1 = self.get_float("D1"); M = self.get_float("M"); t = self.get_float("t"); L = self.get_float("L")
                if any(v <= 0 for v in (D1, M, t, L)): raise ValueError("Все значения должны быть положительными")
                result = calculate_multiple_pipes_insulation_variant2(D1, M, t, L)
            self.Sr_label.setText(f"Площадь обертывания труб (Sr): {format_number(result['Sr'], 6, self)} м²")
            self.Spi_label.setText(f"Площадь покровного слоя изоляции (Spi): {format_number(result['Spi'], 6, self)} м²")
            self.Vi_label.setText(f"Объем изоляции (Vi): {format_number(result['Vi'], 6, self)} м³")
            ts = datetime.now().strftime("%H:%M:%S")
            params = ", ".join(f"{k}={v.text()}" for k, v in self.input_widgets.items() if v.text().strip())
            self.history_list.addItem(f"[{ts}] {self._current_type_name()}\n"
                                      f"  {params}\n"
                                      f"  → Sr={format_number(result['Sr'], 4, self)} м²; "
                                      f"Vi={format_number(result['Vi'], 4, self)} м³")
            self.history_list.scrollToBottom()
            save_history(self, "history/insulation", self.history_list)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка расчета", str(e))

    def clear_fields(self):
        for edit in self.input_widgets.values():
            edit.clear(); mark_invalid(edit, False)
        self.clear_results()

    def clear_results(self):
        self.Sr_label.setText("Площадь обертывания труб (Sr): — м²")
        self.Spi_label.setText("Площадь покровного слоя изоляции (Spi): — м²")
        self.Vi_label.setText("Объем изоляции (Vi): — м³")

    def clear_history(self):
        clear_history(self, "history/insulation", self.history_list)


# ============================================================================
#  КРЕПЁЖ
# ============================================================================

class FastenerTab(QWidget):
    COLUMNS = ['number', 'type', 'std', 'diam', 'length', 'thick',
               'coat', 'strength', 'qty', 'weight_one', 'weight_total']

    def __init__(self):
        super().__init__(); self.row_widgets = {}; self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15); layout.setSpacing(15)
        info_group = QGroupBox("Коэффициенты материалов")
        info_layout = QVBoxLayout(info_group)
        info_text = QLabel("Данные приведены в основном для изделий из стали.\n"
                          "0,35 – алюминий | 1,08 – латунь | 0,97 – бронза | 1,13 – медь")
        info_text.setObjectName("infoLabel"); info_text.setWordWrap(True)
        info_layout.addWidget(info_text); layout.addWidget(info_group)
        self.table = QTableWidget()

        self.table.setColumnCount(11)
        self.table.setHorizontalHeaderLabels([
            "№", "Тип метизов", "Стандарт / Наименование",
            "Размер 1, мм", "Размер 2, мм", "Размер 3, мм",
            "Покрытие", "Класс прочности", "Кол-во, шт",
            "Вес 1 шт, кг", "Вес, кг"
        ])
        for i, w in enumerate([45, 110, 320, 120, 110, 105, 120, 105, 90, 105, 105]):
            self.table.setColumnWidth(i, w)
        for i in range(11):
            self.table.horizontalHeader().setSectionResizeMode(i, QHeaderView.ResizeMode.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True); self.table.setWordWrap(False)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.currentCellChanged.connect(self._update_headers_for_current_row)
        layout.addWidget(self.table, stretch=1)
        button_layout = QHBoxLayout()
        self.add_row_btn = QPushButton("➕ Добавить"); self.add_row_btn.setObjectName("addButton")
        self.add_row_btn.clicked.connect(self.add_row); button_layout.addWidget(self.add_row_btn)
        self.delete_row_btn = QPushButton("🗑 Удалить"); self.delete_row_btn.setObjectName("deleteButton")
        self.delete_row_btn.clicked.connect(self.delete_selected_row)
        button_layout.addWidget(self.delete_row_btn); button_layout.addStretch()
        self.calculate_btn = QPushButton("Рассчитать"); self.calculate_btn.clicked.connect(self.calculate_all)
        button_layout.addWidget(self.calculate_btn)
        self.clear_btn = QPushButton("Очистить всё"); self.clear_btn.clicked.connect(self.clear_all)
        button_layout.addWidget(self.clear_btn); layout.addLayout(button_layout)
        result_group = QGroupBox("Итоговый результат"); result_layout = QVBoxLayout(result_group)
        self.total_weight_label = QLabel("Общий вес: — кг"); self.total_weight_label.setObjectName("resultLabel")
        self.total_weight_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        result_layout.addWidget(self.total_weight_label); layout.addWidget(result_group)
        self.add_row()

    def _apply_headers_for_type(self, type_name: str) -> None:
        h1, h2, h3 = HEADERS_BY_TYPE.get(
            type_name, ("Размер 1, мм", "Размер 2, мм", "Размер 3, мм"))
        for col, text in ((3, h1), (4, h2), (5, h3)):
            item = self.table.horizontalHeaderItem(col)
            if item is None:
                item = QTableWidgetItem(text)
                self.table.setHorizontalHeaderItem(col, item)
            else:
                item.setText(text)

    def _update_headers_for_current_row(self, *_) -> None:
        row = self.table.currentRow()
        if row < 0 or row not in self.row_widgets:
            return
        type_widget = self.row_widgets[row].get('type')
        if type_widget is not None:
            type_name = type_widget.currentText()
            if type_name:
                self._apply_headers_for_type(type_name)

    def get_density(self):
        return STEEL_DENSITY

    def create_combo(self, items, min_width=80):
        combo = QComboBox(); combo.addItems(items); combo.setMinimumWidth(min_width)
        combo.setMinimumHeight(28); combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        combo.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        combo.setStyleSheet(f"QComboBox {{ background-color: #3c3c3c; border: 1px solid #555555; border-radius: 4px; padding: 4px 6px; color: #ffffff; font-size: 11px; min-width: {min_width}px; min-height: 28px; }} QComboBox:hover {{ border: 1px solid #007acc; }} QComboBox::drop-down {{ border: none; width: 20px; }} QComboBox::down-arrow {{ image: none; border-left: 4px solid transparent; border-right: 4px solid transparent; border-top: 5px solid #cccccc; margin-right: 4px; }} QComboBox QAbstractItemView {{ background-color: #2d2d30; border: 1px solid #555555; selection-background-color: #007acc; color: #ffffff; }}")
        return combo

    def create_line_edit(self, placeholder="", min_width=80, validator=None):
        edit = QLineEdit(); edit.setPlaceholderText(placeholder); edit.setMinimumWidth(min_width)
        edit.setMinimumHeight(28); edit.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        edit.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        if validator: edit.setValidator(validator)
        return edit

    def create_label(self, text="—", min_width=80):
        label = QLabel(text); label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setMinimumWidth(min_width); label.setMinimumHeight(28)
        label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        label.setStyleSheet("background-color: #2d2d30; padding: 4px; border-radius: 2px; color: #e0e0e0;")
        return label

    def add_row(self):
        row = self.table.rowCount(); self.table.insertRow(row); self.table.setRowHeight(row, 35)
        number_label = QLabel(str(row + 1)); number_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        number_label.setMinimumHeight(28); number_label.setStyleSheet("background-color: #2d2d30; padding: 4px; color: #e0e0e0;")
        type_combo = self.create_combo(list(STANDARDS_DATA.keys()), 100)
        std_combo = self.create_combo([], 300); diam_combo = self.create_combo([], 80)
        length_combo = self.create_combo([], 80); thick_combo = self.create_combo([], 80)
        coat_combo = self.create_combo([], 120); strength_combo = self.create_combo([], 100)
        int_validator = QIntValidator(1, 999999)
        qty_edit = self.create_line_edit("шт", 80, int_validator)
        weight_one_label = self.create_label("—", 100); weight_total_label = self.create_label("—", 100)

        widgets = {'number': number_label, 'type': type_combo, 'std': std_combo,
                   'diam': diam_combo, 'length': length_combo, 'thick': thick_combo,
                   'coat': coat_combo, 'strength': strength_combo, 'qty': qty_edit,
                   'weight_one': weight_one_label, 'weight_total': weight_total_label}
        self.row_widgets[row] = widgets
        type_combo.currentIndexChanged.connect(lambda idx, w=widgets: self.on_type_changed(w, idx))
        std_combo.currentIndexChanged.connect(lambda idx, w=widgets: self.on_standard_changed(w, idx))
        for col, key in enumerate(self.COLUMNS):
            self.table.setCellWidget(row, col, widgets[key])
        type_combo.setCurrentIndex(0)
        self._apply_headers_for_type(type_combo.currentText())

    def on_type_changed(self, widgets, type_index):
        if type_index < 0: return
        type_name = list(STANDARDS_DATA.keys())[type_index]
        standards = list(STANDARDS_DATA[type_name].keys())
        std_combo = widgets['std']; std_combo.blockSignals(True); std_combo.clear()
        std_combo.addItems(standards)
        if standards: std_combo.setCurrentIndex(0)
        std_combo.blockSignals(False); self.on_standard_changed(widgets, 0)
        self._apply_headers_for_type(type_name)

    def on_standard_changed(self, widgets, std_index):
        if std_index < 0: return
        type_name = widgets['type'].currentText()
        if type_name not in STANDARDS_DATA: return
        standards = list(STANDARDS_DATA[type_name].keys())
        if std_index >= len(standards): return
        std_name = standards[std_index]; data = STANDARDS_DATA[type_name][std_name]
        for key, values in [('diam', data["diameters"]),
                            ('length', data["lengths"]),
                            ('thick', data.get("thicknesses", [])),
                            ('coat', data["coatings"]),
                            ('strength', data["strength_classes"])]:
            combo = widgets[key]; combo.blockSignals(True); combo.clear()
            if values: combo.addItems([str(v) for v in values]); combo.setCurrentIndex(0)
            else: combo.addItem("—")
            combo.blockSignals(False)

    def delete_selected_row(self):
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.information(self, "Информация", "Выберите строку для удаления"); return
        rows_to_delete = sorted([index.row() for index in selected_rows], reverse=True)
        for row in rows_to_delete: self.table.removeRow(row)
        self.rebuild_row_widgets(); self.update_row_numbers(); self.rebind_signals()

    def rebuild_row_widgets(self):
        new_row_widgets = {}
        for row in range(self.table.rowCount()):
            widgets = {}
            for col, key in enumerate(self.COLUMNS):
                widgets[key] = self.table.cellWidget(row, col)
            new_row_widgets[row] = widgets
        self.row_widgets = new_row_widgets

    def rebind_signals(self):
        for widgets in self.row_widgets.values():
            for combo_key in ('type', 'std'):
                combo = widgets.get(combo_key)
                if combo is None: continue
                try: combo.currentIndexChanged.disconnect()
                except TypeError: pass
            widgets['type'].currentIndexChanged.connect(lambda idx, w=widgets: self.on_type_changed(w, idx))
            widgets['std'].currentIndexChanged.connect(lambda idx, w=widgets: self.on_standard_changed(w, idx))

    def update_row_numbers(self):
        for row in range(self.table.rowCount()):
            if row in self.row_widgets:
                self.row_widgets[row]['number'].setText(str(row + 1))

    def get_widget_text(self, widgets, key):
        widget = widgets.get(key)
        if isinstance(widget, QComboBox): return widget.currentText()
        if isinstance(widget, QLineEdit): return widget.text().strip()
        if isinstance(widget, QLabel): return widget.text()
        return ""

    def calculate_weight_one(self, type_name, std_name, diameter, length,
                             thickness="—", density=STEEL_DENSITY):
        try:
            diam_float = float(diameter)
            length_float = float(length) if length not in ("", "—") else 0
            thick_float = float(thickness) if thickness not in ("", "—") else 0
            if type_name == "Болт" and "ГОСТ 7798" in std_name:
                key = (int(diam_float), int(length_float))
                if key in BOLT_MASS_GOST7798:
                    return (BOLT_MASS_GOST7798[key] / 1000) * density / STEEL_DENSITY
            if type_name == "Болт":
                fw = calculate_bolt_weight_formula(int(diam_float), int(length_float), density)
                if fw: return fw
            if type_name == "Винт":
                fw = calculate_bolt_weight_formula(int(diam_float), int(length_float), density)
                if fw: return fw * 0.9
            if type_name == "Гайка":
                di = int(diam_float)
                if di in NUT_MASS_GOST5927:
                    return (NUT_MASS_GOST5927[di] / 1000) * density / STEEL_DENSITY
                fw = calculate_nut_weight_formula(di, density)
                if fw: return fw
            if type_name == "Гвоздь": return calculate_nail_weight(diam_float, length_float, density)
            if type_name == "Саморез":
                d = diam_float / 1000; L = length_float / 1000
                return math.pi * (d ** 2) / 4 * L * density * 0.85
            if type_name == "Заклепка":
                d = diam_float / 1000; L = length_float / 1000
                body = math.pi * (d ** 2) / 4 * L * density
                head_d = 2.0 * d; head_h = 0.6 * d
                head = math.pi * (head_d ** 2) / 4 * head_h * density
                return body + head
            if type_name == "Дюбели":
                d = diam_float / 1000; L = length_float / 1000
                return math.pi * (d ** 2) / 4 * L * 1200
            if type_name == "Круг отрезной":
                od = diam_float / 1000; th = length_float / 1000; idd = od * 0.4
                area = math.pi / 4 * (od ** 2 - idd ** 2)
                return area * th * 2700
            if type_name == "Перфорация":
                if thick_float <= 0:
                    thick_float = 2.0
                w = diam_float / 1000.0
                L = length_float / 1000.0
                t = thick_float / 1000.0
                k_perf = 0.7
                if "Уголок" in std_name:
                    area_side = w * L
                    area_corner = t * L
                    area = 2 * area_side - area_corner
                else:
                    area = w * L
                return area * t * density * k_perf
        except (ValueError, TypeError, ZeroDivisionError): return 0.0
        return 0.0

    def calculate_all(self):
        density = self.get_density(); total_weight = 0.0
        for row, widgets in list(self.row_widgets.items()):
            try:
                type_name = self.get_widget_text(widgets, 'type')
                std_name = self.get_widget_text(widgets, 'std')
                diameter = self.get_widget_text(widgets, 'diam')
                length = self.get_widget_text(widgets, 'length')
                thickness = self.get_widget_text(widgets, 'thick')
                coat_name = self.get_widget_text(widgets, 'coat')
                qty_text = self.get_widget_text(widgets, 'qty')
                if not qty_text or not diameter: continue
                qty = int(float(qty_text.replace(',', '.')))
                weight_one = self.calculate_weight_one(
                    type_name, std_name, diameter, length, thickness, density)
                if weight_one == 0: continue
                if self.get_density() == STEEL_DENSITY:
                    weight_one *= COATING_FACTOR.get(coat_name, 1.0)
                weight_total = weight_one * qty
                if isinstance(widgets.get('weight_one'), QLabel):
                    widgets['weight_one'].setText(format_number(weight_one, 6, self))
                if isinstance(widgets.get('weight_total'), QLabel):
                    widgets['weight_total'].setText(format_number(weight_total, 3, self))
                total_weight += weight_total
            except Exception: logger.exception("FastenerTab: ошибка при обработке строки %s", row)
        self.total_weight_label.setText(f"Общий вес: {format_number(total_weight, 3, self)} кг")

    def clear_all(self):
        self.table.setRowCount(0); self.row_widgets.clear()
        self.total_weight_label.setText("Общий вес: — кг"); self.add_row()


# ============================================================================
#  ПРЕДПРОСМОТР МЕТАЛЛОПРОКАТА
# ============================================================================

class SectionPreview(QWidget):
    def __init__(self, parent=None, min_height=190):
        super().__init__(parent); self.product_index = 0
        self.setMinimumHeight(min_height)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

    def set_product(self, index):
        self.product_index = index; self.update()

    def paintEvent(self, event):
        painter = QPainter(self); painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.fillRect(self.rect(), QColor("#202124"))
        painter.setPen(QPen(QColor("#4ec9b0"), 2))
        painter.drawRoundedRect(self.rect().adjusted(1, 1, -1, -1), 8, 8)
        cx, cy = self.width() // 2, self.height() // 2 + 5
        scale = max(1, min(self.width(), self.height()) * 0.58)
        p = painter; p.setBrush(QColor("#3f444a")); p.setPen(QPen(QColor("#9cdcfe"), 2))
        i = self.product_index
        if i == 0:
            r = scale * 0.45
            p.drawEllipse(int(cx - r), int(cy - r), int(2 * r), int(2 * r))
            p.setBrush(QColor("#202124")); r2 = r * 0.65
            p.drawEllipse(int(cx - r2), int(cy - r2), int(2 * r2), int(2 * r2))
            self._label(p, "D", cx + r + 12, cy - 5); self._label(p, "s", cx + r2 - 4, cy + 20)
        elif i == 1:
            r = scale * 0.42
            p.drawEllipse(int(cx - r), int(cy - r), int(2 * r), int(2 * r))
            self._label(p, "d", cx + r + 12, cy)
        elif i == 2:
            w, h, t = scale * .72, scale * .9, scale * .13
            x, y = cx - w / 2, cy - h / 2
            p.drawRect(int(x), int(y), int(t), int(h))
            p.drawRect(int(x), int(y), int(w), int(t))
            p.drawRect(int(x), int(y + h - t), int(w), int(t))
            self._label(p, "h", x - 30, cy); self._label(p, "b", cx, y - 10)
        elif i == 3:
            w, h, t, tw = scale * .72, scale * .9, scale * .14, scale * .12
            x, y = cx - w / 2, cy - h / 2
            p.drawRect(int(x), int(y), int(w), int(t))
            p.drawRect(int(cx - tw / 2), int(y), int(tw), int(h))
            p.drawRect(int(x), int(y + h - t), int(w), int(t))
            self._label(p, "h", x - 30, cy); self._label(p, "b", cx, y - 10)
        elif i == 4:
            w, h = scale * .85, scale * .45
            p.drawRect(int(cx - w / 2), int(cy - h / 2), int(w), int(h))
            self._label(p, "b × L × t", cx - 45, cy + 8)
        elif i == 5:
            r = scale * .43
            pts = [QPoint(int(cx + r * math.cos(math.radians(60 * k + 30))),
                          int(cy + r * math.sin(math.radians(60 * k + 30)))) for k in range(6)]
            p.drawPolygon(pts); self._label(p, "S", cx + r + 12, cy)
        else:
            w, h, t = scale * .72, scale * .72, scale * .14
            x, y = cx - w / 2, cy - h / 2
            p.drawRect(int(x), int(y), int(t), int(h))
            p.drawRect(int(x), int(y + h - t), int(w), int(t))
            self._label(p, "A", x - 25, cy); self._label(p, "B", cx, y + h + 20)
        p.setPen(QColor("#d7ba7d"))
        p.drawText(15, 25, ["Круглая труба", "Круг / пруток", "Швеллер", "Двутавр",
                            "Лист", "Шестигранник", "Уголок"][i])
        p.end()

    def _label(self, painter, text, x, y):
        painter.setPen(QColor("#dcdcaa"))
        painter.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        painter.drawText(int(x), int(y), text)


# ============================================================================
#  МЕТАЛЛОПРОКАТ
# ============================================================================

class RolledMetalTab(QWidget):
    def __init__(self):
        super().__init__(); self.param_fields = {}; self.init_ui()

    def init_ui(self):
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15); main_layout.setSpacing(15)
        left_widget = QWidget(); layout = QVBoxLayout(left_widget)
        layout.setContentsMargins(0, 0, 0, 0); layout.setSpacing(15)
        selection_group = QGroupBox("Выбор типа металлопроката")
        selection_layout = QHBoxLayout(selection_group)
        self.product_combo = QComboBox()
        self.product_combo.addItems(["Труба круглая", "Круг (пруток)", "Швеллер", "Двутавр",
                                     "Лист", "Шестигранник", "Уголок"])
        self.product_combo.currentIndexChanged.connect(self.on_product_changed)
        selection_layout.addWidget(QLabel("Тип изделия:"))
        selection_layout.addWidget(self.product_combo); selection_layout.addStretch()
        layout.addWidget(selection_group)
        self.preview = SectionPreview(); layout.addWidget(self.preview)
        self.params_container = QVBoxLayout(); self.params_container.setSpacing(10)
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        scroll_widget = QWidget(); self.params_layout = QVBoxLayout(scroll_widget)
        self.params_layout.addLayout(self.params_container); self.params_layout.addStretch()
        scroll.setWidget(scroll_widget); layout.addWidget(scroll, stretch=1)
        result_group = QGroupBox("Результат расчета"); result_layout = QVBoxLayout(result_group)
        self.result_label = QLabel("Масса: — кг"); self.result_label.setObjectName("resultLabel")
        self.result_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        result_layout.addWidget(self.result_label); layout.addWidget(result_group)
        button_layout = QHBoxLayout(); button_layout.addStretch()
        self.calculate_btn = QPushButton("🧮 Рассчитать массу"); self.calculate_btn.clicked.connect(self.calculate)
        button_layout.addWidget(self.calculate_btn)
        self.clear_btn = QPushButton("🗑 Очистить"); self.clear_btn.clicked.connect(self.clear_fields)
        button_layout.addWidget(self.clear_btn); button_layout.addStretch()
        layout.addLayout(button_layout)
        main_layout.addWidget(left_widget, stretch=3)
        history_group = QGroupBox("История расчетов"); history_layout = QVBoxLayout(history_group)
        self.history_list = QListWidget(); self.history_list.setWordWrap(True)
        history_layout.addWidget(self.history_list, stretch=1)
        self.clear_history_btn = QPushButton("🗑 Очистить историю")
        self.clear_history_btn.clicked.connect(self.clear_history)
        history_layout.addWidget(self.clear_history_btn)
        history_group.setMinimumWidth(280); history_group.setMaximumWidth(380)
        main_layout.addWidget(history_group, stretch=1)
        self.on_product_changed(0)
        load_history(self, "history/rolled_metal", self.history_list)

    def get_density(self):
        main_window = self.window()
        if hasattr(main_window, 'current_density'): return main_window.current_density
        return STEEL_DENSITY

    def on_product_changed(self, index):
        self.preview.set_product(index); clear_layout(self.params_container); self.param_fields = {}
        configs = {
            0: [("Наружный диаметр", "мм", "например, 57", "D",
                 "m = π · (D − s) · s · L · ρ"),
                ("Толщина стенки", "мм", "например, 3.5", "s",
                 "m = π · (D − s) · s · L · ρ"),
                ("Длина", "м", "например, 6", "L",
                 "m = π · (D − s) · s · L · ρ")],
            1: [("Диаметр", "мм", "например, 20", "d",
                 "m = π · d² / 4 · L · ρ"),
                ("Длина", "м", "например, 6", "L",
                 "m = π · d² / 4 · L · ρ")],
            2: [("Высота h", "мм", "например, 100", "h",
                 "A = (h − 2s)·s + 2·b·s; m = A · L · ρ"),
                ("Ширина полки b", "мм", "например, 46", "b", "A = (h − 2s)·s + 2·b·s"),
                ("Толщина стенки s", "мм", "например, 4.5", "s", "A = (h − 2s)·s + 2·b·s"),
                ("Длина", "м", "например, 6", "L", "m = A · L · ρ")],
            3: [("Высота h", "мм", "например, 200", "h", "A = (h − 2t)·s + 2·b·t"),
                ("Ширина полки b", "мм", "например, 100", "b", "A = (h − 2t)·s + 2·b·t"),
                ("Толщина стенки s", "мм", "например, 5.5", "s", "A = (h − 2t)·s + 2·b·t"),
                ("Толщина полки t", "мм", "например, 8.5", "t", "A = (h − 2t)·s + 2·b·t"),
                ("Длина", "м", "например, 6", "L", "m = A · L · ρ")],
            4: [("Ширина", "мм", "например, 1000", "w", "m = w·L·t·ρ"),
                ("Длина", "мм", "например, 2000", "L", "m = w·L·t·ρ"),
                ("Толщина", "мм", "например, 3", "t", "m = w·L·t·ρ")],
            5: [("Размер под ключ (S)", "мм", "например, 17", "S",
                 "A = (√3/2) · S²; m = A · L · ρ"),
                ("Длина", "м", "например, 6", "L", "m = A · L · ρ")],
            6: [("Полка A", "мм", "например, 50", "A", "A_сеч = (a + b − t) · t"),
                ("Полка B", "мм", "например, 50", "B", "A_сеч = (a + b − t) · t"),
                ("Толщина полки", "мм", "например, 5", "t", "A_сеч = (a + b − t) · t"),
                ("Длина", "м", "например, 6", "L", "m = A_сеч · L · ρ")],
        }
        config = configs.get(index, [])
        for label_text, unit, placeholder, key, tooltip in config:
            row, edit = create_input_row(label_text, unit, placeholder, tooltip=tooltip)
            self.params_container.addLayout(row); self.param_fields[key] = edit

    def get_float(self, key, default=0.0):
        edit = self.param_fields.get(key)
        return get_float_from_edit(edit, default) if edit else default

    def clear_history(self):
        clear_history(self, "history/rolled_metal", self.history_list)

    def calculate(self):
        index = self.product_combo.currentIndex(); density = self.get_density()
        missing = [e for e in self.param_fields.values() if not e.text().strip()]
        for e in self.param_fields.values():
            mark_invalid(e, not e.text().strip())
        if missing:
            QMessageBox.warning(self, "Ошибка ввода", "Заполните выделенные поля"); return
        for field in self.param_fields.values():
            try:
                val = float(field.text().strip().replace(',', '.'))
                if val <= 0: raise ValueError
            except ValueError:
                mark_invalid(field, True)
                QMessageBox.warning(self, "Ошибка ввода", "Введите положительное число"); return
        try:
            if index == 0:
                D = self.get_float("D"); s = self.get_float("s"); L = self.get_float("L")
                if s * 2 >= D: raise ValueError("Толщина стенки должна быть меньше половины диаметра")
                mass = calculate_pipe_mass(D, s, L, density)
            elif index == 1:
                d = self.get_float("d"); L = self.get_float("L")
                mass = calculate_circle_mass(d, L, density)
            elif index == 2:
                h = self.get_float("h"); b = self.get_float("b"); s = self.get_float("s"); L = self.get_float("L")
                if s * 2 >= h: raise ValueError("2*s должно быть меньше h")
                mass = calculate_channel_mass(h, b, s, L, density)
            elif index == 3:
                h = self.get_float("h"); b = self.get_float("b"); s = self.get_float("s"); t = self.get_float("t"); L = self.get_float("L")
                if s * 2 >= h or t * 2 >= h: raise ValueError("2*s и 2*t должны быть меньше h")
                mass = calculate_beam_mass(h, b, s, t, L, density)
            elif index == 4:
                w = self.get_float("w"); L = self.get_float("L"); t = self.get_float("t")
                mass = calculate_sheet_mass(w, L, t, density)
            elif index == 5:
                S = self.get_float("S"); L = self.get_float("L")
                mass = calculate_hexagon_mass(S, L, density)
            elif index == 6:
                A = self.get_float("A"); B = self.get_float("B"); t = self.get_float("t"); L = self.get_float("L")
                if t >= min(A, B): raise ValueError("Толщина полки должна быть меньше размеров полок")
                mass = calculate_angle_mass(A, B, t, L, density)
            else: mass = 0
            self.result_label.setText(f"Масса: {format_number(mass, 3, self)} кг")
            material = get_current_material_name(self)
            ts = datetime.now().strftime("%H:%M:%S")
            params = ", ".join(f"{k}={v.text()}" for k, v in self.param_fields.items() if v.text().strip())
            self.history_list.addItem(f"[{ts}] Материал: {material}\n"
                                      f"  {self.product_combo.currentText()}\n"
                                      f"  {params}\n"
                                      f"  → {format_number(mass, 3, self)} кг")
            self.history_list.scrollToBottom()
            save_history(self, "history/rolled_metal", self.history_list)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка расчета: {str(e)}")

    def clear_fields(self):
        for field in self.param_fields.values():
            field.clear(); mark_invalid(field, False)
        self.result_label.setText("Масса: — кг")


# ============================================================================
#  ПОКРАСКА
# ============================================================================

class PaintingTab(QWidget):
    def __init__(self):
        super().__init__(); self.param_fields = {}; self.init_ui()

    def init_ui(self):
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(15, 10, 15, 10); main_layout.setSpacing(15)
        left_widget = QWidget(); layout = QVBoxLayout(left_widget)
        layout.setContentsMargins(0, 0, 0, 0); layout.setSpacing(8)
        info_group = QGroupBox("Расчёт площади окраски металлопроката")
        info_layout = QVBoxLayout(info_group); info_layout.setContentsMargins(8, 8, 8, 8)
        info_text = QLabel("Калькулятор рассчитывает площадь окрашиваемой поверхности и расход краски. "
                          "Учитываются: слои, норма расхода краски (кг/м²), запас на потери.")
        info_text.setObjectName("infoLabel"); info_text.setWordWrap(True)
        info_layout.addWidget(info_text); layout.addWidget(info_group)
        top_row = QHBoxLayout(); top_row.setSpacing(10)
        selection_group = QGroupBox("Тип металлопроката")
        selection_layout = QHBoxLayout(selection_group); selection_layout.setContentsMargins(8, 8, 8, 8)
        self.product_combo = QComboBox()
        self.product_combo.addItems(["Труба круглая", "Круг (пруток)", "Швеллер", "Двутавр",
                                     "Лист", "Шестигранник", "Уголок"])
        self.product_combo.currentIndexChanged.connect(self.on_product_changed)
        selection_layout.addWidget(QLabel("Изделие:"))
        selection_layout.addWidget(self.product_combo, stretch=1)
        top_row.addWidget(selection_group, stretch=2)
        self.preview = SectionPreview(min_height=120); top_row.addWidget(self.preview, stretch=3)
        layout.addLayout(top_row)
        param_row = QHBoxLayout(); param_row.setSpacing(15)
        self.geom_group = QGroupBox("Геометрические параметры")
        self.geom_container = QGridLayout(self.geom_group)
        self.geom_container.setContentsMargins(10, 10, 10, 10)
        self.geom_container.setHorizontalSpacing(8); self.geom_container.setVerticalSpacing(6)
        param_row.addWidget(self.geom_group, stretch=3)
        paint_group = QGroupBox("Параметры покраски"); paint_layout = QGridLayout(paint_group)
        paint_layout.setContentsMargins(10, 10, 10, 10)
        paint_layout.setHorizontalSpacing(8); paint_layout.setVerticalSpacing(6)
        paint_layout.addWidget(QLabel("Количество слоёв:"), 0, 0)
        self.layers_edit = QLineEdit("2"); self.layers_edit.setValidator(QIntValidator(1, 50))
        self.layers_edit.setMinimumHeight(28)
        self.layers_edit.setToolTip("Сколько слоёв краски наносится.")
        paint_layout.addWidget(self.layers_edit, 0, 1)
        paint_layout.addWidget(QLabel("Расход на 1 слой, кг/м²:"), 1, 0)
        self.consumption_edit = QLineEdit("0.15")
        self.consumption_edit.setValidator(create_c_locale_double_validator(0.0, 1e6, 6))
        self.consumption_edit.setMinimumHeight(28)
        self.consumption_edit.setToolTip("Норма расхода краски на 1 слой, кг/м².")
        paint_layout.addWidget(self.consumption_edit, 1, 1)
        paint_layout.addWidget(QLabel("Запас на потери, %:"), 2, 0)
        self.loss_edit = QLineEdit("5")
        self.loss_edit.setValidator(create_c_locale_double_validator(0.0, 100.0, 3))
        self.loss_edit.setMinimumHeight(28)
        self.loss_edit.setToolTip("Процент технологических потерь.")
        paint_layout.addWidget(self.loss_edit, 2, 1)
        paint_layout.setRowStretch(3, 1); param_row.addWidget(paint_group, stretch=2)
        layout.addLayout(param_row)
        button_layout = QHBoxLayout(); button_layout.addStretch()
        self.calculate_btn = QPushButton("🧮 Рассчитать"); self.calculate_btn.clicked.connect(self.calculate)
        button_layout.addWidget(self.calculate_btn)
        self.clear_btn = QPushButton("🗑 Очистить"); self.clear_btn.clicked.connect(self.clear_fields)
        button_layout.addWidget(self.clear_btn); button_layout.addStretch()
        layout.addLayout(button_layout)
        result_group = QGroupBox("Результаты расчёта"); result_layout = QVBoxLayout(result_group)
        result_layout.setContentsMargins(8, 8, 8, 8); result_layout.setSpacing(4)
        self.area_label = QLabel("Площадь окраски: — м²"); self.area_label.setObjectName("resultLabel")
        result_layout.addWidget(self.area_label)
        self.paint_one_layer_label = QLabel("Расход краски на 1 слой: — кг")
        self.paint_one_layer_label.setObjectName("resultLabel"); result_layout.addWidget(self.paint_one_layer_label)
        self.paint_total_label = QLabel("Полный расход краски: — кг")
        self.paint_total_label.setObjectName("resultLabel"); result_layout.addWidget(self.paint_total_label)
        layout.addWidget(result_group); layout.addStretch()
        main_layout.addWidget(left_widget, stretch=3)
        history_group = QGroupBox("История расчетов"); history_layout = QVBoxLayout(history_group)
        self.history_list = QListWidget(); self.history_list.setWordWrap(True)
        history_layout.addWidget(self.history_list, stretch=1)
        self.clear_history_btn = QPushButton("🗑 Очистить историю")
        self.clear_history_btn.clicked.connect(self.clear_history)
        history_layout.addWidget(self.clear_history_btn)
        history_group.setMinimumWidth(280); history_group.setMaximumWidth(380)
        main_layout.addWidget(history_group, stretch=1)
        self.on_product_changed(0)
        load_history(self, "history/painting", self.history_list)

    def on_product_changed(self, index):
        self.preview.set_product(index)
        while self.geom_container.count():
            item = self.geom_container.takeAt(0); w = item.widget()
            if w is not None: w.deleteLater()
        self.param_fields = {}
        configs = {
            0: [("Наружный диаметр", "мм", "например, 57", "D", "S = π · D · L"),
                ("Длина", "м", "например, 6", "L", "S = π · D · L")],
            1: [("Диаметр", "мм", "например, 20", "d", "S = π · d · L"),
                ("Длина", "м", "например, 6", "L", "S = π · d · L")],
            2: [("Высота h", "мм", "например, 100", "h", "S = (2h + 4b − 2s) · L"),
                ("Ширина полки b", "мм", "например, 46", "b", "S = (2h + 4b − 2s) · L"),
                ("Толщина стенки s", "мм", "например, 4.5", "s", "S = (2h + 4b − 2s) · L"),
                ("Длина", "м", "например, 6", "L", "S = (2h + 4b − 2s) · L")],
            3: [("Высота h", "мм", "например, 200", "h", "S = (2h + 4b − 2s) · L"),
                ("Ширина полки b", "мм", "например, 100", "b", "S = (2h + 4b − 2s) · L"),
                ("Толщина стенки s", "мм", "например, 5.5", "s", "S = (2h + 4b − 2s) · L"),
                ("Толщина полки t", "мм", "например, 8.5", "t", "Учитывается в периметре"),
                ("Длина", "м", "например, 6", "L", "S = (2h + 4b − 2s) · L")],
            4: [("Ширина", "мм", "например, 1000", "w", "S = 2 · w · L (обе стороны)"),
                ("Длина", "мм", "например, 2000", "L", "S = 2 · w · L"),
                ("Толщина", "мм", "например, 3", "t", "Толщина на площадь не влияет")],
            5: [("Размер под ключ (S)", "мм", "например, 17", "S", "сторона = S/√3; P = 6·сторона"),
                ("Длина", "м", "например, 6", "L", "S = P · L")],
            6: [("Полка A", "мм", "например, 50", "A", "P ≈ 2A + 2B"),
                ("Полка B", "мм", "например, 50", "B", "P ≈ 2A + 2B"),
                ("Толщина полки", "мм", "например, 5", "t", "Слабо влияет"),
                ("Длина", "м", "например, 6", "L", "S = P · L")],
        }
        config = configs.get(index, [])
        for r, (label_text, unit, placeholder, key, tooltip) in enumerate(config):
            lbl = QLabel(label_text); lbl.setStyleSheet("font-weight: 500;"); lbl.setMinimumWidth(150)
            if tooltip: lbl.setToolTip(tooltip)
            edit = QLineEdit(); edit.setPlaceholderText(placeholder); edit.setMinimumHeight(28)
            edit.setValidator(create_c_locale_double_validator())
            if tooltip: edit.setToolTip(tooltip)
            edit.textChanged.connect(lambda *_, e=edit: mark_invalid(e, False))
            unit_lbl = QLabel(unit); unit_lbl.setObjectName("unitLabel"); unit_lbl.setMinimumWidth(30)
            self.geom_container.addWidget(lbl, r, 0); self.geom_container.addWidget(edit, r, 1)
            self.geom_container.addWidget(unit_lbl, r, 2); self.param_fields[key] = edit
        self.geom_container.setColumnStretch(1, 1); self.geom_container.setRowStretch(len(config), 1)

    def get_float(self, key, default=0.0):
        edit = self.param_fields.get(key)
        return get_float_from_edit(edit, default) if edit else default

    def clear_history(self):
        clear_history(self, "history/painting", self.history_list)

    def calculate(self):
        index = self.product_combo.currentIndex()
        missing = [e for e in self.param_fields.values() if not e.text().strip()]
        for e in self.param_fields.values():
            mark_invalid(e, not e.text().strip())
        if missing:
            QMessageBox.warning(self, "Ошибка ввода", "Заполните выделенные поля"); return
        for field in self.param_fields.values():
            try:
                val = float(field.text().strip().replace(',', '.'))
                if val <= 0: raise ValueError
            except ValueError:
                mark_invalid(field, True)
                QMessageBox.warning(self, "Ошибка ввода", "Введите положительное число"); return
        try:
            layers = int(float(self.layers_edit.text().strip().replace(',', '.') or "0"))
            if layers <= 0: raise ValueError
        except ValueError:
            mark_invalid(self.layers_edit, True)
            QMessageBox.warning(self, "Ошибка ввода", "Количество слоёв — целое положительное"); return
        try:
            consumption = float(self.consumption_edit.text().strip().replace(',', '.') or "0")
            if consumption <= 0: raise ValueError
        except ValueError:
            mark_invalid(self.consumption_edit, True)
            QMessageBox.warning(self, "Ошибка ввода", "Расход краски — положительное число"); return
        try:
            loss = float(self.loss_edit.text().strip().replace(',', '.') or "0")
            if loss < 0: raise ValueError
        except ValueError:
            mark_invalid(self.loss_edit, True)
            QMessageBox.warning(self, "Ошибка ввода", "Запас на потери — неотрицательное число"); return
        try:
            if index == 0:
                D = self.get_float("D"); L = self.get_float("L"); area = paint_area_pipe(D, L)
            elif index == 1:
                d = self.get_float("d"); L = self.get_float("L"); area = paint_area_circle(d, L)
            elif index == 2:
                h = self.get_float("h"); b = self.get_float("b"); s = self.get_float("s"); L = self.get_float("L")
                area = paint_area_channel(h, b, s, L)
            elif index == 3:
                h = self.get_float("h"); b = self.get_float("b"); s = self.get_float("s"); t = self.get_float("t"); L = self.get_float("L")
                area = paint_area_beam(h, b, s, t, L)
            elif index == 4:
                w = self.get_float("w"); L_mm = self.get_float("L"); area = paint_area_sheet(w, L_mm)
            elif index == 5:
                S = self.get_float("S"); L = self.get_float("L"); area = paint_area_hexagon(S, L)
            elif index == 6:
                A = self.get_float("A"); B = self.get_float("B"); t = self.get_float("t"); L = self.get_float("L")
                area = paint_area_angle(A, B, t, L)
            else: area = 0.0
        except Exception as e:
            QMessageBox.critical(self, "Ошибка расчёта", str(e)); return
        paint_per_layer = area * consumption
        paint_total = paint_per_layer * layers * (1 + loss / 100.0)
        self.area_label.setText(f"Площадь окраски: {format_number(area, 4, self)} м²")
        self.paint_one_layer_label.setText(f"Расход краски на 1 слой: {format_number(paint_per_layer, 4, self)} кг")
        self.paint_total_label.setText(f"Полный расход краски: {format_number(paint_total, 4, self)} кг")
        ts = datetime.now().strftime("%H:%M:%S")
        params = ", ".join(f"{k}={v.text()}" for k, v in self.param_fields.items() if v.text().strip())
        self.history_list.addItem(f"[{ts}] {self.product_combo.currentText()}\n"
                                  f"  {params}, слоёв={layers}\n"
                                  f"  → S={format_number(area, 4, self)} м²; "
                                  f"краски={format_number(paint_total, 4, self)} кг")
        self.history_list.scrollToBottom()
        save_history(self, "history/painting", self.history_list)

    def clear_fields(self):
        for field in self.param_fields.values():
            field.clear(); mark_invalid(field, False)
        self.layers_edit.setText("2"); self.consumption_edit.setText("0.15"); self.loss_edit.setText("5")
        self.area_label.setText("Площадь окраски: — м²")
        self.paint_one_layer_label.setText("Расход краски на 1 слой: — кг")
        self.paint_total_label.setText("Полный расход краски: — кг")


# ============================================================================
#  АКЗ (ручной, из Excel)
# ============================================================================

class AnticorrTab(QWidget):
    def __init__(self):
        super().__init__(); self.data = {}; self.current_file = None; self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15); layout.setSpacing(15)
        load_btn = QPushButton("📂 Загрузить файл Excel с нормами")
        load_btn.clicked.connect(self.load_excel); layout.addWidget(load_btn)
        param_group = QGroupBox("Параметры расчета"); param_layout = QGridLayout(param_group)
        param_layout.addWidget(QLabel("Раздел работ:"), 0, 0)
        self.cbo_section = QComboBox(); self.cbo_section.currentIndexChanged.connect(self.on_section_changed)
        param_layout.addWidget(self.cbo_section, 0, 1)
        param_layout.addWidget(QLabel("Ведомость:"), 1, 0)
        self.cbo_vest = QComboBox(); param_layout.addWidget(self.cbo_vest, 1, 1)
        param_layout.addWidget(QLabel("Площадь (м²):"), 0, 2)
        self.edit_area = QLineEdit(); self.edit_area.setText("100")
        self.edit_area.setValidator(create_c_locale_double_validator(0.0, 1e9, 6))
        self.edit_area.textChanged.connect(lambda: mark_invalid(self.edit_area, False))
        param_layout.addWidget(self.edit_area, 0, 3)
        param_layout.addWidget(QLabel("Количество слоёв:"), 1, 2)
        self.cbo_layers = QComboBox(); self.cbo_layers.addItems(["По умолчанию", "1", "2", "3", "4", "5"])
        param_layout.addWidget(self.cbo_layers, 1, 3)
        layout.addWidget(param_group)
        btn_layout = QHBoxLayout(); self.calc_btn = QPushButton("🧮 Рассчитать")
        self.calc_btn.clicked.connect(self.calculate)
        btn_layout.addStretch(); btn_layout.addWidget(self.calc_btn); btn_layout.addStretch()
        layout.addLayout(btn_layout)
        self.table = QTableWidget(); self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["Материал", "Марка", "Слоёв", "Норма (кг/м²)", "Расход (кг)"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        layout.addWidget(self.table, stretch=1)
        self.total_label = QLabel("Итого: — кг"); self.total_label.setObjectName("resultLabel")
        self.total_label.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(self.total_label)

    def load_excel(self):
        if openpyxl is None:
            QMessageBox.critical(self, "Ошибка", "Библиотека openpyxl не установлена. Установите её: pip install openpyxl"); return
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите файл Excel с нормами", "", "Excel files (*.xlsx *.xls)")
        if not file_path: return
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True); data = {}
            for sheet_name in wb.sheetnames:
                if sheet_name in ["Оглавление", "литература"]: continue
                sheet = wb[sheet_name]; vest_list = []; current_vest = None; last_vest_num = None
                for row in sheet.iter_rows(min_row=2, values_only=True):
                    vest_num = str(row[0]).strip() if row[0] else ""
                    purpose = str(row[1]).strip() if row[1] else ""
                    mat_name = str(row[2]).strip() if row[2] else ""
                    mark = str(row[3]).strip() if row[3] else ""
                    norm_1_layer = float(row[7]) if row[7] and isinstance(row[7], (int, float)) else 0
                    layers = int(row[8]) if row[8] and isinstance(row[8], (int, float)) else 0
                    total_norm = float(row[10]) if row[10] and isinstance(row[10], (int, float)) else 0
                    if total_norm == 0 and norm_1_layer > 0 and layers > 0: total_norm = norm_1_layer * layers
                    if vest_num and vest_num != "NaN":
                        if (current_vest and vest_num != last_vest_num and current_vest["materials"]):
                            vest_list.append(current_vest)
                        if not current_vest or vest_num != last_vest_num:
                            current_vest = {"number": vest_num, "purpose": purpose, "materials": []}
                            last_vest_num = vest_num
                    if mat_name and mat_name != "NaN" and total_norm > 0 and current_vest:
                        current_vest["materials"].append({"name": mat_name, "mark": mark, "layers": layers,
                                                          "total_norm": total_norm, "norm_1_layer": norm_1_layer})
                if current_vest and current_vest["materials"]: vest_list.append(current_vest)
                if vest_list: data[sheet_name] = vest_list
            self.data = data; self.current_file = file_path; self.update_sections()
            QMessageBox.information(self, "Успех", f"Загружено {len(data)} разделов")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка загрузки", str(e))

    def update_sections(self):
        self.cbo_section.clear(); self.cbo_section.addItems(list(self.data.keys()))
        if self.cbo_section.count(): self.cbo_section.setCurrentIndex(0)

    def on_section_changed(self):
        section = self.cbo_section.currentText()
        if not section or section not in self.data: self.cbo_vest.clear(); return
        self.cbo_vest.clear()
        for vest in self.data[section]: self.cbo_vest.addItem(vest["number"], vest)
        if self.cbo_vest.count(): self.cbo_vest.setCurrentIndex(0)

    def get_selected_vest(self):
        idx = self.cbo_vest.currentIndex()
        return self.cbo_vest.itemData(idx) if idx >= 0 else None

    def calculate(self):
        if not self.data:
            QMessageBox.warning(self, "Нет данных", "Сначала загрузите файл Excel с нормами"); return
        vest = self.get_selected_vest()
        if not vest:
            QMessageBox.warning(self, "Ошибка", "Выберите ведомость"); return
        if not self.edit_area.text().strip():
            mark_invalid(self.edit_area, True)
            QMessageBox.warning(self, "Ошибка ввода", "Введите площадь"); return
        try:
            area = float(self.edit_area.text().strip().replace(',', '.'))
            if area <= 0: raise ValueError
        except ValueError:
            mark_invalid(self.edit_area, True)
            QMessageBox.warning(self, "Ошибка ввода", "Введите положительное число"); return
        layers_text = self.cbo_layers.currentText()
        user_layers = None if layers_text == "По умолчанию" else int(layers_text)
        self.table.setRowCount(0); total_sum = 0.0
        for mat in vest["materials"]:
            total_norm = mat["norm_1_layer"] * user_layers if user_layers else mat["total_norm"]
            layers_display = user_layers if user_layers else mat["layers"]
            if total_norm > 0:
                total = total_norm * area; total_sum += total
                row = self.table.rowCount(); self.table.insertRow(row)
                self.table.setItem(row, 0, QTableWidgetItem(str(mat["name"])))
                self.table.setItem(row, 1, QTableWidgetItem(str(mat["mark"])))
                self.table.setItem(row, 2, QTableWidgetItem(str(layers_display)))
                self.table.setItem(row, 3, QTableWidgetItem(f"{total_norm:.4f}"))
                self.table.setItem(row, 4, QTableWidgetItem(f"{total:.4f}"))
        self.total_label.setText(f"Итого: {total_sum:.4f} кг")


# ============================================================================
#  ПРЕДПРОСМОТР СВАРНЫХ ШВОВ
# ============================================================================

class WeldPreviewWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent); self.joint_key = "C2"
        self.setMinimumHeight(120); self.setMinimumWidth(200)

    def setJoint(self, key):
        self.joint_key = key; self.update()

    def paintEvent(self, event):
        painter = QPainter(self); painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        logic_w = 220; logic_h = 90
        widget_w = self.width(); widget_h = self.height()
        if widget_w <= 0 or widget_h <= 0: painter.end(); return
        scale_x = widget_w / logic_w; scale_y = widget_h / logic_h
        scale = min(scale_x, scale_y)
        if scale <= 0: painter.end(); return
        offset_x = (widget_w - logic_w * scale) / 2; offset_y = (widget_h - logic_h * scale) / 2
        painter.translate(offset_x, offset_y); painter.scale(scale, scale)
        dispatch = {"C2": self.draw_C2, "C8": self.draw_C8, "C17": self.draw_C17,
                    "U5": self.draw_U5, "U7": self.draw_U7, "U8": self.draw_U8,
                    "H1": self.draw_H1, "H2": self.draw_H2, "T1": self.draw_T1,
                    "T3": self.draw_T3, "U4": self.draw_U4, "U5_S": self.draw_U5_S,
                    "C25": self.draw_C25}
        method = dispatch.get(self.joint_key)
        if method: method(painter, logic_w, logic_h)
        else: painter.drawText(QRectF(0, 0, logic_w, logic_h), Qt.AlignmentFlag.AlignCenter, "Схема не определена")
        painter.end()

    def _draw_text(self, painter, text, x, y, color="#94a3b8"):
        painter.setPen(QColor(color)); painter.setFont(QFont("Segoe UI", 8))
        painter.drawText(QPointF(x, y), text)

    def draw_C2(self, painter, w, h):
        painter.setPen(QPen(QColor("#64748b"), 1.5)); painter.setBrush(QBrush(QColor("#475569")))
        painter.drawRect(10, 25, 85, 40); painter.drawRect(w - 95, 25, 85, 40)
        painter.setBrush(QBrush(QColor("#38bdf8"), Qt.BrushStyle.SolidPattern))
        path = QPainterPath(); path.moveTo(95, 25); path.quadTo(110, 12, w - 95, 25)
        path.lineTo(w - 95, 65); path.quadTo(110, 78, 95, 65); path.closeSubpath()
        painter.drawPath(path); self._draw_text(painter, "Стык С2 (Зазор)", 60, 82)

    def draw_C8(self, painter, w, h):
        painter.setPen(QPen(QColor("#64748b"), 1.5)); painter.setBrush(QBrush(QColor("#475569")))
        painter.drawRect(10, 25, 85, 40)
        points = [QPointF(w - 95, 25), QPointF(w - 10, 25), QPointF(w - 10, 65), QPointF(w - 65, 65)]
        painter.drawPolygon(*points)
        painter.setBrush(QBrush(QColor("#38bdf8"), Qt.BrushStyle.SolidPattern))
        path = QPainterPath(); path.moveTo(95, 25); path.quadTo(110, 12, w - 95, 25)
        path.lineTo(w - 65, 65); path.lineTo(95, 65); path.closeSubpath()
        painter.drawPath(path); self._draw_text(painter, "Стык С8 (1 скос)", 60, 82)

    def draw_C17(self, painter, w, h):
        painter.setPen(QPen(QColor("#64748b"), 1.5)); painter.setBrush(QBrush(QColor("#475569")))
        points = [QPointF(10, 25), QPointF(80, 25), QPointF(95, 65), QPointF(10, 65)]
        painter.drawPolygon(*points)
        points2 = [QPointF(w - 80, 25), QPointF(w - 10, 25), QPointF(w - 10, 65), QPointF(w - 95, 65)]
        painter.drawPolygon(*points2)
        painter.setBrush(QBrush(QColor("#38bdf8"), Qt.BrushStyle.SolidPattern))
        path = QPainterPath(); path.moveTo(80, 25); path.quadTo(w // 2, 8, w - 80, 25)
        path.lineTo(w - 95, 65); path.lineTo(95, 65); path.closeSubpath()
        painter.drawPath(path); self._draw_text(painter, "Стык С17 (V-разделка)", 60, 82)

    def draw_U5(self, painter, w, h):
        painter.setPen(QPen(QColor("#64748b"), 1.5)); painter.setBrush(QBrush(QColor("#475569")))
        painter.drawRect(20, 55, w - 40, 20); painter.drawRect(w // 2 - 20, 5, 40, 50)
        painter.setBrush(QBrush(QColor("#38bdf8"), Qt.BrushStyle.SolidPattern))
        path = QPainterPath(); path.moveTo(w // 2 - 20, 55); path.lineTo(w // 2 - 40, 55)
        path.lineTo(w // 2 - 20, 35); path.closeSubpath()
        painter.drawPath(path); self._draw_text(painter, "Угловое У5 (Фланец)", 60, 82)

    def draw_U7(self, painter, w, h):
        painter.setPen(QPen(QColor("#64748b"), 1.5)); painter.setBrush(QBrush(QColor("#475569")))
        painter.drawRect(20, 60, w - 40, 18)
        points = [QPointF(w // 2 - 20, 10), QPointF(w // 2 + 20, 10), QPointF(w // 2 + 20, 60), QPointF(w // 2 - 20, 45)]
        painter.drawPolygon(*points)
        painter.setBrush(QBrush(QColor("#38bdf8"), Qt.BrushStyle.SolidPattern))
        path = QPainterPath(); path.moveTo(w // 2 - 20, 45); path.lineTo(w // 2 + 20, 60)
        path.lineTo(w // 2 - 20, 60); path.closeSubpath()
        painter.drawPath(path); self._draw_text(painter, "Угловое У7 (Скос)", 60, 82)

    def draw_U8(self, painter, w, h):
        painter.setPen(QPen(QColor("#64748b"), 1.5)); painter.setBrush(QBrush(QColor("#475569")))
        painter.drawRect(20, 55, w - 40, 20); painter.drawRect(w // 2 - 20, 5, 40, 50)
        painter.setBrush(QBrush(QColor("#38bdf8"), Qt.BrushStyle.SolidPattern))
        for dx in (-1, 1):
            path = QPainterPath(); path.moveTo(w // 2 + dx * 20, 55)
            path.lineTo(w // 2 + dx * 40, 55); path.lineTo(w // 2 + dx * 20, 35); path.closeSubpath()
            painter.drawPath(path)
        self._draw_text(painter, "У8 (Двустороннее)", 60, 82)

    def draw_H1(self, painter, w, h):
        painter.setPen(QPen(QColor("#64748b"), 1.5)); painter.setBrush(QBrush(QColor("#475569")))
        painter.drawRect(10, 20, 120, 20); painter.drawRect(70, 45, 140, 20)
        painter.setBrush(QBrush(QColor("#38bdf8"), Qt.BrushStyle.SolidPattern))
        path = QPainterPath(); path.moveTo(130, 45); path.lineTo(130, 20); path.lineTo(150, 45); path.closeSubpath()
        painter.drawPath(path); self._draw_text(painter, "Нахлест Н1 (1 шов)", 60, 82)

    def draw_H2(self, painter, w, h):
        painter.setPen(QPen(QColor("#64748b"), 1.5)); painter.setBrush(QBrush(QColor("#475569")))
        painter.drawRect(10, 20, 120, 20); painter.drawRect(70, 45, 140, 20)
        painter.setBrush(QBrush(QColor("#38bdf8"), Qt.BrushStyle.SolidPattern))
        path1 = QPainterPath(); path1.moveTo(130, 45); path1.lineTo(130, 20); path1.lineTo(150, 45); path1.closeSubpath()
        painter.drawPath(path1)
        path2 = QPainterPath(); path2.moveTo(70, 45); path2.lineTo(70, 65); path2.lineTo(50, 45); path2.closeSubpath()
        painter.drawPath(path2); self._draw_text(painter, "Нахлест Н2 (2 шва)", 60, 82)

    def draw_T1(self, painter, w, h):
        painter.setPen(QPen(QColor("#64748b"), 1.5)); painter.setBrush(QBrush(QColor("#475569")))
        painter.drawRect(10, 55, w - 20, 18); painter.drawRect(w // 2 - 15, 10, 30, 45)
        painter.setBrush(QBrush(QColor("#38bdf8"), Qt.BrushStyle.SolidPattern))
        path = QPainterPath(); path.moveTo(w // 2 - 15, 55); path.lineTo(w // 2 - 35, 55)
        path.lineTo(w // 2 - 15, 35); path.closeSubpath()
        painter.drawPath(path); self._draw_text(painter, "Тавровое Т1", 80, 82)

    def draw_T3(self, painter, w, h):
        painter.setPen(QPen(QColor("#64748b"), 1.5)); painter.setBrush(QBrush(QColor("#475569")))
        painter.drawRect(10, 55, w - 20, 18); painter.drawRect(w // 2 - 15, 10, 30, 45)
        painter.setBrush(QBrush(QColor("#38bdf8"), Qt.BrushStyle.SolidPattern))
        for dx in (-1, 1):
            path = QPainterPath(); path.moveTo(w // 2 + dx * 15, 55)
            path.lineTo(w // 2 + dx * 35, 55); path.lineTo(w // 2 + dx * 15, 35); path.closeSubpath()
            painter.drawPath(path)
        self._draw_text(painter, "Тавровое Т3 (Двустороннее)", 60, 82)

    def draw_U4(self, painter, w, h):
        painter.setPen(QPen(QColor("#64748b"), 1.5)); painter.setBrush(QBrush(QColor("#475569")))
        painter.drawRect(30, 55, 140, 20); painter.drawRect(30, 10, 20, 45)
        painter.setBrush(QBrush(QColor("#38bdf8"), Qt.BrushStyle.SolidPattern))
        path = QPainterPath(); path.moveTo(50, 55); path.lineTo(65, 55); path.lineTo(50, 35); path.closeSubpath()
        painter.drawPath(path); self._draw_text(painter, "Угловое У4", 80, 82)

    def draw_U5_S(self, painter, w, h):
        painter.setPen(QPen(QColor("#64748b"), 1.5)); painter.setBrush(QBrush(QColor("#475569")))
        painter.drawRect(30, 55, 140, 20); painter.drawRect(30, 10, 20, 45)
        painter.setBrush(QBrush(QColor("#38bdf8"), Qt.BrushStyle.SolidPattern))
        path1 = QPainterPath(); path1.moveTo(50, 55); path1.lineTo(65, 55); path1.lineTo(50, 35); path1.closeSubpath()
        painter.drawPath(path1)
        path2 = QPainterPath(); path2.moveTo(30, 55); path2.lineTo(15, 55); path2.lineTo(30, 35); path2.closeSubpath()
        painter.drawPath(path2); self._draw_text(painter, "Угловое У5 (Двустороннее)", 60, 82)

    def draw_C25(self, painter, w, h):
        painter.setPen(QPen(QColor("#64748b"), 1.5)); painter.setBrush(QBrush(QColor("#475569")))
        points = [QPointF(20, 15), QPointF(100, 15), QPointF(70, 55), QPointF(20, 55)]
        painter.drawPolygon(*points)
        points2 = [QPointF(w - 100, 15), QPointF(w - 20, 15), QPointF(w - 20, 55), QPointF(w - 70, 55)]
        painter.drawPolygon(*points2)
        painter.setBrush(QBrush(QColor("#38bdf8"), Qt.BrushStyle.SolidPattern))
        path = QPainterPath(); path.moveTo(100, 15); path.lineTo(w - 100, 15)
        path.lineTo(w - 70, 55); path.lineTo(70, 55); path.closeSubpath()
        painter.drawPath(path); self._draw_text(painter, "Пруток С25 (Косой рез)", 60, 82)


# ============================================================================
#  СВАРКА
# ============================================================================

class WeldingTab(QWidget):
    def __init__(self):
        super().__init__(); self.setMinimumSize(750, 700); self.init_ui()

    def init_ui(self):
        self.main_layout = QVBoxLayout(self); self.main_layout.setSpacing(15)
        top_grid = QGridLayout(); top_grid.setSpacing(15)
        self.category_combo = QComboBox()
        self.category_combo.addItems(["Трубы (ГОСТ 16037-80)", "Листы (ГОСТ 5264-80)", "Прутки (ГОСТ 5264-80)"])
        self.category_combo.currentIndexChanged.connect(self.on_category_changed)
        top_grid.addWidget(QLabel("Тип конструкции:"), 0, 0); top_grid.addWidget(self.category_combo, 0, 1)
        self.method_combo = QComboBox()
        self.method_combo.addItems(["Ручная дуговая (РДС)", "Полуавтомат (MIG/MAG в CO₂)",
                                    "Аргонодуговая (TIG)", "Газовая сварка"])
        self.method_combo.currentIndexChanged.connect(self.calculate)
        top_grid.addWidget(QLabel("Способ сварки:"), 0, 2); top_grid.addWidget(self.method_combo, 0, 3)
        self.main_layout.addLayout(top_grid)
        mid_grid = QGridLayout(); mid_grid.setSpacing(15)
        self.position_combo = QComboBox()
        self.position_combo.addItems(["Нижнее", "Горизонтальное", "Вертикальное", "Потолочное", "Наклонное"])
        self.position_combo.setCurrentIndex(0); self.position_combo.currentIndexChanged.connect(self.calculate)
        mid_grid.addWidget(QLabel("Положение шва:"), 0, 0); mid_grid.addWidget(self.position_combo, 0, 1)
        self.qual_combo = QComboBox()
        self.qual_combo.addItems(["Высокая (минимальный угар)", "Средняя (+8% к потерям)"])
        self.qual_combo.currentIndexChanged.connect(self.calculate)
        mid_grid.addWidget(QLabel("Квалификация сварщика:"), 0, 2); mid_grid.addWidget(self.qual_combo, 0, 3)
        self.main_layout.addLayout(mid_grid)
        self.stack = QStackedWidget()
        self.pipe_widget = self.create_pipe_widget(); self.sheet_widget = self.create_sheet_widget()
        self.bar_widget = self.create_bar_widget()
        self.stack.addWidget(self.pipe_widget); self.stack.addWidget(self.sheet_widget); self.stack.addWidget(self.bar_widget)
        self.main_layout.addWidget(self.stack)
        preview_group = QGroupBox("Схема разделки кромок и шва по ГОСТ")
        preview_layout = QVBoxLayout(); self.preview_widget = WeldPreviewWidget()
        self.preview_widget.setMinimumHeight(120); preview_layout.addWidget(self.preview_widget)
        preview_group.setLayout(preview_layout); self.main_layout.addWidget(preview_group)
        results_group = QGroupBox("Результаты расчёта"); results_layout = QGridLayout(); results_layout.setSpacing(10)
        self.res_len_label = QLabel("0.00 м"); self.res_len_label.setStyleSheet("color: #f8fafc; font-weight: 700; font-size: 16px;")
        results_layout.addWidget(QLabel("Протяженность шва (всего):"), 0, 0); results_layout.addWidget(self.res_len_label, 0, 1)
        self.res_F_label = QLabel("0.00 мм²"); self.res_F_label.setStyleSheet("color: #f8fafc; font-weight: 700; font-size: 16px;")
        results_layout.addWidget(QLabel("Расчетная площадь сечения шва:"), 1, 0); results_layout.addWidget(self.res_F_label, 1, 1)
        self.res_pure_label = QLabel("0.000 кг"); self.res_pure_label.setStyleSheet("color: #f8fafc; font-weight: 700; font-size: 16px;")
        results_layout.addWidget(QLabel("Чистый вес наплавленного металла:"), 2, 0); results_layout.addWidget(self.res_pure_label, 2, 1)
        self.mat_label = QLabel("Расход сварочных материалов:"); self.res_total_label = QLabel("0.000 кг")
        self.res_total_label.setStyleSheet("color: #fb923c; font-weight: 700; font-size: 19px;")
        results_layout.addWidget(self.mat_label, 3, 0); results_layout.addWidget(self.res_total_label, 3, 1)
        self.gas_co2_row_label = QLabel("Расход защитного газа CO₂:")
        self.gas_co2_label = QLabel("0.00 кг"); self.gas_co2_label.setStyleSheet("color: #34d399; font-weight: 700;")
        self.gas_co2_row_label.hide(); self.gas_co2_label.hide()
        results_layout.addWidget(self.gas_co2_row_label, 4, 0); results_layout.addWidget(self.gas_co2_label, 4, 1)
        self.gas_ar_row_label = QLabel("Расход защитного газа Аргон (Ar):")
        self.gas_ar_label = QLabel("0.00 л (0.00 баллонов)"); self.gas_ar_label.setStyleSheet("color: #34d399; font-weight: 700;")
        self.gas_ar_row_label.hide(); self.gas_ar_label.hide()
        results_layout.addWidget(self.gas_ar_row_label, 5, 0); results_layout.addWidget(self.gas_ar_label, 5, 1)
        results_group.setLayout(results_layout); self.main_layout.addWidget(results_group)
        btn_layout = QHBoxLayout(); btn_layout.addStretch()
        reset_btn = QPushButton("🔄 Сбросить"); reset_btn.setObjectName("reset"); reset_btn.clicked.connect(self.reset_all)
        btn_layout.addWidget(reset_btn)
        copy_btn = QPushButton("📋 Копировать результаты"); copy_btn.setObjectName("primary"); copy_btn.clicked.connect(self.export_results)
        btn_layout.addWidget(copy_btn); self.main_layout.addLayout(btn_layout)
        self.on_category_changed()

    def create_pipe_widget(self):
        widget = QWidget(); layout = QVBoxLayout(widget); layout.setSpacing(10)
        self.pipe_joint_combo = QComboBox()
        self.pipe_joint_combo.addItems(["С2 — Стыковое одностороннее без скоса", "С8 — Стыковое со скосом одной кромки",
                                        "С17 — Стыковое с V-образной разделкой", "У5 — Угловое без скоса кромок (с фланцем)",
                                        "У7 — Угловое со скосом одной кромки", "У8 — Угловое двустороннее без скоса"])
        self.pipe_joint_combo.currentIndexChanged.connect(self.update_preview_and_calc)
        layout.addWidget(QLabel("Тип соединения (ГОСТ 16037-80):")); layout.addWidget(self.pipe_joint_combo)
        grid = QGridLayout(); grid.setSpacing(10); validator = create_c_locale_double_validator()
        self.pipe_D_edit = QLineEdit("219"); self.pipe_D_edit.setPlaceholderText("мм"); self.pipe_D_edit.setValidator(validator)
        self.pipe_D_edit.setToolTip("Наружный диаметр трубы, мм.")
        self.pipe_S_edit = QLineEdit("6"); self.pipe_S_edit.setPlaceholderText("мм"); self.pipe_S_edit.setValidator(validator)
        self.pipe_S_edit.setToolTip("Толщина стенки, мм.")
        self.pipe_count_edit = QLineEdit("4"); self.pipe_count_edit.setPlaceholderText("шт")
        self.pipe_count_edit.setToolTip("Количество стыков.")
        for edit in [self.pipe_D_edit, self.pipe_S_edit, self.pipe_count_edit]: edit.textChanged.connect(self.calculate)
        grid.addWidget(QLabel("Наружный диаметр (D), мм:"), 0, 0); grid.addWidget(self.pipe_D_edit, 0, 1)
        grid.addWidget(QLabel("Толщина стенки (S), мм:"), 0, 2); grid.addWidget(self.pipe_S_edit, 0, 3)
        grid.addWidget(QLabel("Количество стыков, шт:"), 1, 0); grid.addWidget(self.pipe_count_edit, 1, 1)
        layout.addLayout(grid); return widget

    def create_sheet_widget(self):
        widget = QWidget(); layout = QVBoxLayout(widget); layout.setSpacing(10)
        self.sheet_joint_combo = QComboBox()
        self.sheet_joint_combo.addItems(["Н1 — Нахлесточное одностороннее", "Н2 — Нахлесточное двустороннее",
                                         "Т1 — Тавровое одностороннее без скоса", "Т3 — Тавровое двустороннее без скоса",
                                         "У4 — Угловое одностороннее", "У5 — Угловое двустороннее"])
        self.sheet_joint_combo.currentIndexChanged.connect(self.update_preview_and_calc)
        layout.addWidget(QLabel("Тип соединения (ГОСТ 5264-80):")); layout.addWidget(self.sheet_joint_combo)
        grid = QGridLayout(); grid.setSpacing(10); validator = create_c_locale_double_validator()
        self.sheet_S_edit = QLineEdit("8"); self.sheet_S_edit.setPlaceholderText("мм"); self.sheet_S_edit.setValidator(validator)
        self.sheet_S_edit.setToolTip("Толщина листа, мм.")
        self.sheet_L_edit = QLineEdit("10"); self.sheet_L_edit.setPlaceholderText("м"); self.sheet_L_edit.setValidator(validator)
        self.sheet_L_edit.setToolTip("Суммарная длина швов, м.")
        for edit in [self.sheet_S_edit, self.sheet_L_edit]: edit.textChanged.connect(self.calculate)
        grid.addWidget(QLabel("Толщина листа (S), мм:"), 0, 0); grid.addWidget(self.sheet_S_edit, 0, 1)
        grid.addWidget(QLabel("Общая длина швов, м:"), 0, 2); grid.addWidget(self.sheet_L_edit, 0, 3)
        layout.addLayout(grid); return widget

    def create_bar_widget(self):
        widget = QWidget(); layout = QVBoxLayout(widget); layout.setSpacing(10)
        self.bar_joint_combo = QComboBox()
        self.bar_joint_combo.addItems(["С25 — Стыковое двух круглых прутков с косым резом"])
        self.bar_joint_combo.currentIndexChanged.connect(self.update_preview_and_calc)
        layout.addWidget(QLabel("Тип соединения (ГОСТ 5264-80):")); layout.addWidget(self.bar_joint_combo)
        grid = QGridLayout(); grid.setSpacing(10); validator = create_c_locale_double_validator()
        self.bar_d_edit = QLineEdit("25"); self.bar_d_edit.setPlaceholderText("мм"); self.bar_d_edit.setValidator(validator)
        self.bar_d_edit.setToolTip("Диаметр прутка, мм.")
        self.bar_count_edit = QLineEdit("10"); self.bar_count_edit.setPlaceholderText("шт")
        self.bar_count_edit.setToolTip("Количество стыков.")
        for edit in [self.bar_d_edit, self.bar_count_edit]: edit.textChanged.connect(self.calculate)
        grid.addWidget(QLabel("Диаметр прутка (d), мм:"), 0, 0); grid.addWidget(self.bar_d_edit, 0, 1)
        grid.addWidget(QLabel("Количество стыков, шт:"), 0, 2); grid.addWidget(self.bar_count_edit, 0, 3)
        layout.addLayout(grid); return widget

    def on_category_changed(self):
        idx = self.category_combo.currentIndex(); self.stack.setCurrentIndex(idx)
        self.update_preview_and_calc()

    def update_preview_and_calc(self): self.update_preview(); self.calculate()

    def update_preview(self):
        mapping = {"С2": "C2", "С8": "C8", "С17": "C17", "У5": "U5", "У7": "U7", "У8": "U8",
                   "Н1": "H1", "Н2": "H2", "Т1": "T1", "Т3": "T3", "У4": "U4", "С25": "C25"}
        cat_idx = self.category_combo.currentIndex()
        if cat_idx == 0:
            joint_text = self.pipe_joint_combo.currentText()
            raw_key = joint_text.split("—")[0].strip(); key = mapping.get(raw_key, raw_key)
        elif cat_idx == 1:
            joint_text = self.sheet_joint_combo.currentText()
            raw_key = joint_text.split("—")[0].strip(); key = mapping.get(raw_key, raw_key)
            if key == "U5": key = "U5_S"
        else: key = "C25"
        self.preview_widget.setJoint(key)

    def get_value(self, edit, default=0.0): return get_float_from_edit(edit, default)
    def get_int_value(self, edit, default=0): return get_int_from_edit(edit, default)

    def calculate(self):
        cat_idx = self.category_combo.currentIndex(); method = self.method_combo.currentIndex()
        pos_idx = self.position_combo.currentIndex(); qual_idx = self.qual_combo.currentIndex()
        k_pos = [1.0, 1.05, 1.10, 1.20, 1.07][pos_idx]
        k_loss = [1.0, 1.08][qual_idx]
        k_method = [1.50, 1.12, 1.05, 1.03][method]
        mat_text = ["Расход электродов:", "Расход св. проволоки:", "Расход присадочных прутков:",
                    "Расход присадочной проволоки:"][method]
        self.mat_label.setText(mat_text)
        DENSITY = 7850; F = 0.0; length_m = 0.0
        if cat_idx == 0:
            D = self.get_value(self.pipe_D_edit); S = self.get_value(self.pipe_S_edit)
            count = self.get_int_value(self.pipe_count_edit)
            joint_text = self.pipe_joint_combo.currentText(); joint = joint_text.split("-")[0].strip()
            if not (D > 0 and S > 0 and count > 0): return
            length_m = (math.pi * (D - S) / 1000.0) * count
            if joint == "С2":
            # Динамический подбор параметров шва С2 по ГОСТ 16037-80 в зависимости от стенки
            if S <= 3.0:
                b = 1.0  # зазор, мм
                g = 1.0  # высота усиления, мм
                e = S + b + 1.0  # ширина шва, мм
            elif S <= 5.0:
                b = 2.0
                g = 1.5
                e = S + b + 2.0
            else:  # Для S = 6 мм и более
                b = 3.0  # нормативный зазор по стандарту
                g = 2.0  # нормальное усиление шва
                e = 10.0 # ширина шва для стенки 6-8 мм
                
                # F = площадь зазора (прямоугольник) + площадь усиления шва (парабола 2/3 * e * g)
                F = (S * b) + ((2.0 / 3.0) * e * g)
            elif joint == "С8":
                b = 1.5; c = 1.5; angle = 30
                h_bevel = S - c; b_bevel = h_bevel * math.tan(math.radians(angle))
                g = 2.0; e = b + b_bevel + 3
                F = (S * b) + (0.5 * b_bevel * h_bevel) + (0.75 * e * g)
            elif joint == "С17":
                b = 2.0; c = 1.5; angle = 60
                h_bevel = S - c; b_bevel = 2 * (h_bevel * math.tan(math.radians(angle / 2)))
                g = 2.0; e = b + b_bevel + 2
                F = (S * b) + (0.5 * b_bevel * h_bevel) + (0.75 * e * g)

            K = S; F_single = (0.5 * K * K) + (0.2 * K * K)
            F = F_single * 2 if joint == "У8" else F_single
            elif joint == "У7":
                K = S; angle = 45
                F_bevel = 0.5 * K * (K * math.tan(math.radians(angle)))
                F = F_bevel + (0.2 * K * K)
        elif cat_idx == 1:
            S = self.get_value(self.sheet_S_edit); L = self.get_value(self.sheet_L_edit)
            joint_text = self.sheet_joint_combo.currentText(); joint = joint_text.split("-")[0].strip()
            if S > 0 and L > 0:
                length_m = L; K = S; F_single = (0.5 * K * K) * 1.2
                if joint in ("Н1", "Т1", "У4"): F = F_single
                elif joint in ("Н2", "Т3", "У5"): F = F_single * 2
        elif cat_idx == 2:
            d = self.get_value(self.bar_d_edit); count = self.get_int_value(self.bar_count_edit)
            if d > 0 and count > 0:
                length_m = (math.pi * d / 1000.0) * count
                F = (math.pi * d * d / 4.0) * 0.45
        mass_pure = (F * 1e-6) * length_m * DENSITY
        mass_total = mass_pure * k_method * k_pos * k_loss
        self.res_len_label.setText(f"{length_m:.2f} м"); self.res_F_label.setText(f"{F:.2f} мм²")
        self.res_pure_label.setText(f"{mass_pure:.3f} кг"); self.res_total_label.setText(f"{mass_total:.3f} кг")
        self.gas_co2_row_label.hide(); self.gas_co2_label.hide()
        self.gas_ar_row_label.hide(); self.gas_ar_label.hide()
        if method == 1:
            gas_co2 = mass_pure * 1.25; self.gas_co2_label.setText(f"{gas_co2:.2f} кг")
            self.gas_co2_row_label.show(); self.gas_co2_label.show()
        elif method == 2:
            gas_ar_liters = length_m * 120; cylinders = gas_ar_liters / 6000.0
            self.gas_ar_label.setText(f"{gas_ar_liters:.0f} л ({cylinders:.2f} баллонов)")
            self.gas_ar_row_label.show(); self.gas_ar_label.show()

    def reset_all(self):
        self.category_combo.setCurrentIndex(0); self.method_combo.setCurrentIndex(0)
        self.position_combo.setCurrentIndex(0); self.qual_combo.setCurrentIndex(0)
        self.pipe_D_edit.setText("219"); self.pipe_S_edit.setText("6")
        self.pipe_count_edit.setText("4"); self.pipe_joint_combo.setCurrentIndex(0)
        self.sheet_S_edit.setText("8"); self.sheet_L_edit.setText("10"); self.sheet_joint_combo.setCurrentIndex(0)
        self.bar_d_edit.setText("25"); self.bar_count_edit.setText("10"); self.on_category_changed()

    def export_results(self):
        lines = ["📊 Результаты расчёта сварочного калькулятора", "----------------------------------------",
                 f"Протяженность шва: {self.res_len_label.text()}",
                 f"Площадь сечения шва: {self.res_F_label.text()}",
                 f"Чистый вес наплавки: {self.res_pure_label.text()}",
                 f"Расход материалов: {self.res_total_label.text()}"]
        if not self.gas_co2_row_label.isHidden(): lines.append(f"Расход CO₂: {self.gas_co2_label.text()}")
        if not self.gas_ar_row_label.isHidden(): lines.append(f"Расход Ar: {self.gas_ar_label.text()}")
        lines.append("----------------------------------------"); lines.append("Разработчик: Тищенко В.В., ООО СГК")
        QApplication.clipboard().setText("\n".join(lines))
        QMessageBox.information(self, "Готово", "✅ Результаты скопированы в буфер обмена!")


# ============================================================================
#  СПРАВКА
# ============================================================================

class HelpTab(QWidget):
    def __init__(self):
        super().__init__(); self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20); layout.setSpacing(15)
        title = QLabel("📘 Справочная информация"); title.setObjectName("titleLabel")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(title)
        info_text = QLabel("Данная программа предназначена для инженерно-технических расчётов "
                          "в области металлообработки, крепёжных изделий, тепловой изоляции "
                          "трубопроводов и антикоррозионных покрытий.\n\n"
                          "Модули программы:\n"
                          "• Расчёт массы металлопроката\n"
                          "• Сортамент ГОСТ\n"
                          "• Профили ГКЛ и перфорация\n"
                          "• Расчёт площади окраски и расхода краски\n"
                          "• Библиотека норм АКЗ (ЛКМ)\n"
                          "• Расчёт массы крепёжных изделий\n"
                          "• Расчёт объёма тепловой изоляции\n"
                          "• Расчёт материалов на АКЗ по Excel\n"
                          "• Расчёт норм расхода сварочных материалов\n"
                          "• Пескоструйная очистка (Sa 1 … Sa 3)\n"
                          "• Обычный калькулятор\n"
                          "• Игры: Тетрис, 2048\n\n"
                          "Все числовые поля принимают и точку, и запятую.\n"
                          "При наведении на поле — подсказка с формулой.\n"
                          "Пустые поля подсвечиваются красным при попытке расчёта.\n"
                          "Поля результатов копируются: двойной клик — всё, ПКМ — только число.\n"
                          "История расчётов и пользовательские материалы сохраняются автоматически.\n"
                          "Во вкладке «Крепеж» заголовки столбцов 3/4/5 меняются\n"
                          "в зависимости от типа выбранного метиза.\n"
                          "В диалоге «Материалы» пользовательские записи можно править\n"
                          "прямо в таблице: двойной клик → правка → Enter.")
        info_text.setWordWrap(True); info_text.setStyleSheet("font-size: 13px; color: #e0e0e0;")
        layout.addWidget(info_text)
        dev_label = QLabel("Разработчик: Тищенко Вячеслав Владимирович | Сметный отдел ООО «СГК»")
        dev_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        dev_label.setStyleSheet("font-size: 12px; color: #9cdcfe; background-color: #2d2d30; padding: 10px; border-radius: 4px;")
        layout.addWidget(dev_label)
        version = QLabel("Версия 5.4 (редактирование пользовательских материалов в диалоге)")
        version.setAlignment(Qt.AlignmentFlag.AlignCenter)
        version.setStyleSheet("font-size: 11px; color: #8b949e;"); layout.addWidget(version)
        layout.addStretch()


# ============================================================================
#  ГЛАВНОЕ ОКНО
# ============================================================================

class MetalCalculator(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Калькулятор Сметчика")
        self.setMinimumSize(1200, 800); self.resize(1500, 950)
        self.current_density = 7850
        self.settings = QSettings("SlavaTV", "MetalCalculator")
        self.rounding_decimals = int(self.settings.value("rounding_decimals", 3))
        self.custom_materials: Dict[str, float] = load_custom_materials(self.settings)
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget(); self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20); main_layout.setSpacing(15)
        title_label = QLabel("🔧 Калькулятор Сметчика"); title_label.setObjectName("titleLabel")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        material_label = QLabel("Материал:"); self.material_combo = QComboBox()
        self._refresh_material_combo(initial=True)
        self.material_combo.currentIndexChanged.connect(self.on_material_changed)
        header_layout = QHBoxLayout()
        header_layout.addWidget(title_label, stretch=1); header_layout.addWidget(material_label)
        header_layout.addWidget(self.material_combo)
        self.materials_btn = QPushButton("📚 Материалы")
        self.materials_btn.setToolTip("Редактировать библиотеку материалов (добавить свои)")
        self.materials_btn.clicked.connect(self.open_materials_dialog)
        header_layout.addWidget(self.materials_btn)
        settings_btn = QPushButton("⚙ Настройки"); settings_btn.clicked.connect(self.open_settings)
        header_layout.addWidget(settings_btn); main_layout.addLayout(header_layout)
        self.tabs = QTabWidget(); main_layout.addWidget(self.tabs)

        self.tab_rolled = RolledMetalTab(); self.tabs.addTab(self.tab_rolled, "📐 Металлопрокат")
        self.tab_sortament = SortamentTab(); self.tabs.addTab(self.tab_sortament, "📋 Сортамент ГОСТ")
        self.tab_profiles = ProfilesTab(); self.tabs.addTab(self.tab_profiles, "📏 Профили")
        self.tab_painting = PaintingTab(); self.tabs.addTab(self.tab_painting, "🎨 Покраска")
        self.tab_sandblast = SandblastTab(); self.tabs.addTab(self.tab_sandblast, "🪨 Пескоструй")
        self.tab_akz_lib = AkzLibraryTab(); self.tabs.addTab(self.tab_akz_lib, "🎨 Нормы АКЗ")
        self.tab_fastener = FastenerTab(); self.tabs.addTab(self.tab_fastener, "🔩 Крепеж")
        self.tab_insulation = InsulationTab(); self.tabs.addTab(self.tab_insulation, "🌡️ Изоляция")
        self.tab_anticorr = AnticorrTab(); self.tabs.addTab(self.tab_anticorr, "🧪 Расчет АКЗ")
        self.tab_welding = WeldingTab(); self.tabs.addTab(self.tab_welding, "⚡ Сварка")
        self.tab_calculator = CalculatorTab(); self.tabs.addTab(self.tab_calculator, "🧮 Калькулятор")
        self.tab_tetris = TetrisWidget(); self.tabs.addTab(self.tab_tetris, "🎮 Тетрис")
        self.tab_game2048 = Game2048Widget(); self.tabs.addTab(self.tab_game2048, "🔢 2048")
        self.tab_help = HelpTab(); self.tabs.addTab(self.tab_help, "📖 Справки")

        self.statusBar().showMessage("Разработал Тищенко Вячеслав Владимирович | Сметный отдел ООО «СГК»")

        CopyableResultHelper.enable_for_all(self)

    def _refresh_material_combo(self, initial=False):
        prev = self.material_combo.currentText() if not initial else None
        self.material_combo.blockSignals(True)
        self.material_combo.clear()
        for name, density in BUILTIN_MATERIALS.items():
            self.material_combo.addItem(name, density)
        for name, density in self.custom_materials.items():
            self.material_combo.addItem(name, density)
        if prev:
            idx = self.material_combo.findText(prev)
            if idx >= 0:
                self.material_combo.setCurrentIndex(idx)
        self.material_combo.blockSignals(False)
        self.current_density = self.material_combo.currentData() or STEEL_DENSITY

    def open_materials_dialog(self):
        dialog = MaterialsDialog(self, self.custom_materials)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.custom_materials = dialog.get_custom_materials()
            save_custom_materials(self.settings, self.custom_materials)
            self._refresh_material_combo()
            self.statusBar().showMessage(
                f"Материалов: {len(BUILTIN_MATERIALS)} встроенных + "
                f"{len(self.custom_materials)} пользовательских", 4000)

    def open_settings(self):
        dialog = QDialog(self); dialog.setWindowTitle("Настройки калькулятора")
        dialog.setModal(True); form = QFormLayout(dialog)
        spin = QSpinBox(dialog); spin.setRange(0, 8); spin.setValue(self.rounding_decimals); spin.setSuffix(" знаков")
        form.addRow("Округление результатов:", spin)
        note = QLabel("Настройка применяется к новым результатам во всех вкладках.")
        note.setWordWrap(True); form.addRow(note)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept); buttons.rejected.connect(dialog.reject); form.addRow(buttons)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.rounding_decimals = spin.value(); self.settings.setValue("rounding_decimals", self.rounding_decimals)
            self.statusBar().showMessage(f"Округление: {self.rounding_decimals} знаков", 3000)

    def on_material_changed(self, index):
        if index >= 0:
            self.current_density = self.material_combo.currentData()
            name = self.material_combo.currentText()
            self.statusBar().showMessage(f"Материал: {name} (ρ = {self.current_density} кг/м³). "
                                         f"Влияет только на вкладку «Металлопрокат».", 4000)


# ============================================================================
#  ТОЧКА ВХОДА
# ============================================================================

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(DARK_THEME_STYLESHEET)
    font = QFont("Segoe UI", 11); app.setFont(font)
    window = MetalCalculator(); window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
