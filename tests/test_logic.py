import unittest
from logic import analyze_data

class TestAnalyzeData(unittest.TestCase):
    def setUp(self):
        self.classes = [
            {"Sınıf Adı": "A1.01", "Seviye": "A1", "Zaman Kodu": 0},
            {"Sınıf Adı": "A2.01", "Seviye": "A2", "Zaman Kodu": 0}
        ]

    def test_analyze_data_valid(self):
        teachers = [
            {
                "Ad Soyad": "Ahmet Hoca",
                "Rol": "Danışman",
                "Sabit Sınıf": "A1.01",
                "Yasaklı Günler": "Salı"
            }
        ]
        errors, warnings = analyze_data(teachers, self.classes, allow_native_advisor=False)
        self.assertEqual(len(errors), 0)
        self.assertEqual(len(warnings), 0)

    def test_analyze_data_native_with_fixed_class_error(self):
        teachers = [
            {
                "Ad Soyad": "Sarah Native",
                "Rol": "Native",
                "Sabit Sınıf": "A1.01",
                "Yasaklı Günler": ""
            }
        ]
        # Case 1: allow_native_advisor = False (Error expected)
        errors, warnings = analyze_data(teachers, self.classes, allow_native_advisor=False)
        self.assertTrue(any("Native hocaya sabit sınıf verilmesi engellendi" in e for e in errors))

        # Case 2: allow_native_advisor = True (No error expected)
        errors, warnings = analyze_data(teachers, self.classes, allow_native_advisor=True)
        self.assertFalse(any("Native hocaya sabit sınıf verilmesi engellendi" in e for e in errors))

    def test_analyze_data_missing_class_error(self):
        teachers = [
            {
                "Ad Soyad": "Mehmet Hoca",
                "Rol": "Danışman",
                "Sabit Sınıf": "NON_EXISTENT_CLASS",
                "Yasaklı Günler": ""
            }
        ]
        errors, warnings = analyze_data(teachers, self.classes, allow_native_advisor=False)
        self.assertTrue(any("sistemde yok" in e for e in errors))

    def test_analyze_data_monday_forbidden_warning(self):
        teachers = [
            {
                "Ad Soyad": "Ayşe Hoca",
                "Rol": "Danışman",
                "Sabit Sınıf": "A1.01",
                "Yasaklı Günler": "Pazartesi"
            }
        ]
        errors, warnings = analyze_data(teachers, self.classes, allow_native_advisor=False)
        self.assertEqual(len(errors), 0)
        self.assertTrue(any("Pazartesi yasaklı" in w for w in warnings))

if __name__ == "__main__":
    unittest.main()
