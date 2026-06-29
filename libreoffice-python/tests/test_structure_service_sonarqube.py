import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.append(str(ROOT))

from services.structure_service import build_slide_structure


class SonarQubeStructureServiceTest(unittest.TestCase):
    def test_sonarqube_vulnerabilities_chart_uses_ratings_and_trims_summary(self):
        chart_definitions = {
            "vulnerabilidades": {
                "Seguridad": "Seguridad",
                "Security Review": "Security review",
            },
            "soporte": {
                "Consultas Respondidas": "Consultas Respondidas",
                "Monitoreos Realizados": "Monitoreos Realizados",
                "Capacitaciones/ Asistencias": "Capacitaciones/ Asistencias",
            },
        }

        product_data = [
            {"Semana": "Semana 1", "Consultas Respondidas": 0, "Monitoreos Realizados": 1, "Capacitaciones/ Asistencias": 2, "Rating": "A", "Security review": 74, "Seguridad": 76, "product": "sonarqube"},
            {"Semana": "Semana 2", "Consultas Respondidas": 1, "Monitoreos Realizados": 1, "Capacitaciones/ Asistencias": 1, "Rating": "B", "Security review": 0, "Seguridad": 1},
            {"Semana": "Semana 3", "Consultas Respondidas": 1, "Monitoreos Realizados": 1, "Capacitaciones/ Asistencias": 3, "Rating": "C", "Security review": 1, "Seguridad": 1},
            {"Semana": "Semana 4", "Consultas Respondidas": 0, "Monitoreos Realizados": 1, "Capacitaciones/ Asistencias": 1, "Rating": "D", "Security review": 3, "Seguridad": 0},
            {"Semana": "Rating E", "Consultas Respondidas": "", "Monitoreos Realizados": "", "Capacitaciones/ Asistencias": "", "Rating": "E", "Security review": 0, "Seguridad": 0},
            {"Semana": "resumen", "Consultas Respondidas": "Versión actual: v2026.1.2"},
            {"Semana": "sugerencia_1", "Consultas Respondidas": "Le recomendamos utilizar las siguientes funciones que tiene a su alcance:\n\nUtilización de aperturas automáticas y elevación de privilegio."},
            {"Semana": "sugerencia_version", "Consultas Respondidas": "Ultima versión del producto: v2026.3"},
            {"Semana": "desc", "Consultas Respondidas": "Durante el período reportado se llevaron a cabo dos actividades de alto impacto sobre la plataforma SonarQube, herramienta central para el análisis de calidad y seguridad del código fuente."},
        ]

        slide = build_slide_structure(
            product_data,
            "sonarqube",
            chart_definitions,
            "Consultas Respondidas",
            "",
        )

        self.assertEqual(slide["resumen"], "Versión actual: v2026.1.2")
        self.assertEqual(slide["sugerencia_1"], "Le recomendamos utilizar las siguientes funciones que tiene a su alcance:\n\nUtilización de aperturas automáticas y elevación de privilegio.")
        self.assertEqual(slide["sugerencia_version"], "Ultima versión del producto: v2026.3")
        self.assertEqual(slide["desc"], "Durante el período reportado se llevaron a cabo dos actividades de alto impacto sobre la plataforma SonarQube, herramienta central para el análisis de calidad y seguridad del código fuente.")
        self.assertEqual(slide["charts"]["vulnerabilidades"]["labels"], ["A", "B", "C", "D", "E"])
        self.assertEqual(slide["charts"]["vulnerabilidades"]["Seguridad"], [76, 1, 1, 0, 0])
        self.assertEqual(slide["charts"]["vulnerabilidades"]["Security Review"], [74, 0, 1, 3, 0])
        self.assertEqual(slide["kpis_1"], "Seguridad: 78\nSecurity review: 78\n")
        self.assertEqual(slide["title_1"], "Vulnerabilidades")


if __name__ == "__main__":
    unittest.main()
