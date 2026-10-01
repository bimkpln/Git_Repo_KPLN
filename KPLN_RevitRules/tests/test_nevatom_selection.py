import importlib.util
from pathlib import Path
import unittest


MODULE_PATH = Path(__file__).resolve().parents[1] / "src" / "nevantom_selection.py"
SPEC = importlib.util.spec_from_file_location("nevantom_selection", MODULE_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)


SAMPLE = """Установка П6 (ID установки 5690505, ID расчета 11663824) KT
2,2-250911903-02.02-K-О-P-А
Серия KT Длина установки 3210 мм
Типоразмер KT_ 2,2 Тип панели Zn / Zn
Вес 205 кг
Количество 1 шт
Сторона обслуживания Правая
1 - Вставка гибкая ВГп-500-300-ОП-ш20.ш20
8 - Вставка гибкая ВГп-500-300-ОП-ш20.ш20
Приточный воздух
Свободный напор 300 Па
Производительность 1985 м3/ч
Температура -3 ◦C
4. Водяной нагреватель
Объемный расход жидкости 0.74 (0.86) м3/ч Мощность 16.71 (19.46) кВт
Суммарные потери давления по
1.94 (2.49) кПа Запас по поверхности теплообмена 14.13 %
жидкости
Диаметр подключения (вход/выход) 25х3.2-N/25х3.2-N
6. Вентилятор
Расход фактический 1985 м3/ч Напор свободный 300 Па
Мощность двигателя 0.75 кВт
Обороты фактические 2810 об/мин Номинальный ток двигателя 1.77 A
Рабочая частота 48 Гц
7. Водяной охладитель
Суммарные потери давления по
38.79 (42.99) кПа Мощность 15.58 (16.42) кВт
жидкости
Объемный расход жидкости 2.67 (2.82) м3/ч
Диаметр подключения (вход/выход) 25х3.2-OW/25х3.2-OW
8. Гибкая вставка
"""


class NevatomSelectionTests(unittest.TestCase):
    def test_normalizes_cyrillic_and_latin_marks(self):
        self.assertEqual(
            MODULE.normalize_mark("КТ 2,2-250911903-02.02-K-О-P-А"),
            MODULE.normalize_mark("KT 2,2-250911903-02.02-K-O-P-A"),
        )

    def test_parses_selected_and_reference_values_separately(self):
        parsed = MODULE.parse_selections([{"page": 1, "text": SAMPLE}])[0]
        self.assertEqual("П6", parsed["system"])
        self.assertEqual("5690505", parsed["installation_id"])
        self.assertEqual(3210.0, parsed["overall_length_mm"])
        self.assertEqual([500.0, 300.0], parsed["inlet_connector_mm"])
        self.assertEqual((695.0, 480.0), (parsed["body_width_mm"], parsed["body_height_mm"]))
        self.assertEqual(
            {"selected": 16.71, "reference": 19.46}, parsed["heater"]["power_kw"]
        )
        self.assertEqual(
            {"selected": 38.79, "reference": 42.99},
            parsed["cooler"]["fluid_pressure_loss_kpa"],
        )
        self.assertEqual(2810.0, parsed["fan"]["actual_speed_rpm"])
        self.assertEqual([], MODULE.validate_selection(parsed))

    def test_blocks_summary_and_fan_disagreement(self):
        parsed = MODULE.parse_selections([{"page": 1, "text": SAMPLE}])[0]
        parsed["fan"]["free_pressure_pa"] = 299.0
        issues = MODULE.validate_selection(parsed)
        self.assertEqual("summary_fan_mismatch", issues[0]["code"])

    def test_marks_sl_offer_as_outside_kt_family(self):
        page = {
            "page": 9,
            "text": "Установка SL_ 5,8 250911903.24.04-K-O-P\nОфис Система П8",
        }
        result = MODULE.find_unmatched_offers([page])
        self.assertEqual("П8", result[0]["system"])
        self.assertEqual("series_not_supported_by_kt_family", result[0]["reason"])

    def test_does_not_attach_sl_offer_to_previous_kt_selection(self):
        pages = [
            {"page": 1, "text": SAMPLE},
            {
                "page": 2,
                "text": "Установка SL_ 5,8 250911903.24.04-K-O-P\nОфис Система П8",
            },
        ]
        parsed = MODULE.parse_selections(pages)[0]
        self.assertEqual([1, 1], parsed["source_pages"])


if __name__ == "__main__":
    unittest.main()
