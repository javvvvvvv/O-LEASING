import os
import sys
import unittest
import openpyxl

# Agregar directorio raíz
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _ROOT)

from reports.auditoria_balanza import (
    parsear_balanza_contpaqi,
    obtener_datos_sistema,
    conciliar_balanza_vs_sistema,
    generar_excel_auditoria_balanza,
    _obtener_db_path
)


class TestAuditoriaBalanza(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.balanza_path = r"C:\Users\cynti\Downloads\BALANZA 26\BALANZZA SEPTIEMBRE.xlsx"
        cls.db_path = _obtener_db_path()

    def test_parsear_balanza(self):
        """Verifica la lectura correcta del archivo de Balanza y metadatos."""
        meta, df_bal = parsear_balanza_contpaqi(self.balanza_path)
        self.assertIn("LEASY", meta['empresa'].upper())
        self.assertEqual(meta['fecha_corte'], "2026-09-30")
        self.assertGreater(meta['total_cuentas'], 3000)

        # Verificar cuenta padre 102-01 y cuenta padre 115-00
        padre_102 = df_bal[df_bal['cuenta'] == '102-01-00-0000-0000-00']
        self.assertEqual(len(padre_102), 1)
        self.assertAlmostEqual(padre_102['s_act_d'].values[0], -1434443.04, places=2)

        padre_115 = df_bal[df_bal['cuenta'] == '115-00-00-0000-0000-00']
        self.assertEqual(len(padre_115), 1)
        self.assertAlmostEqual(padre_115['s_act_d'].values[0], 63958832.87, places=2)

    def test_clasificacion_contratos(self):
        """Verifica la separación cronológica entre activos y liquidados al 30/Sep/2026."""
        datos_sis = obtener_datos_sistema(self.db_path, "2026-09-30")
        self.assertEqual(len(datos_sis['df_activos']), 275)
        self.assertEqual(len(datos_sis['df_liquidados']), 440)

    def test_conciliacion_y_omision_liquidados(self):
        """Verifica que la cartera activa de 102-01 omita los liquidados y estos se auditen aparte."""
        meta, df_bal = parsear_balanza_contpaqi(self.balanza_path)
        datos_sis = obtener_datos_sistema(self.db_path, meta['fecha_corte'])
        tablas = conciliar_balanza_vs_sistema(df_bal, datos_sis, meta['fecha_corte'])

        # Cartera activa (102-01) debe tener exactamente los 275 activos
        df_cxc = tablas['df_concil_cxc']
        self.assertEqual(len(df_cxc), 275)

        # Contratos liquidados debe tener exactamente los 440 liquidados
        df_liq = tablas['df_concil_liquidados']
        self.assertEqual(len(df_liq), 440)

        # Comprobar que hay contratos liquidados alertados con saldo pendiente
        con_saldo = df_liq[df_liq['saldo_cxc_balanza'] > 0.01]
        self.assertGreater(len(con_saldo), 0)

    def test_generacion_excel(self):
        """Verifica la creación del libro oficial con las 6 pestañas requeridas."""
        out_test = os.path.join(_ROOT, "scratch", "test_auditoria_output.xlsx")
        res_path = generar_excel_auditoria_balanza(self.balanza_path, output_excel_path=out_test, db_path=self.db_path)
        self.assertTrue(os.path.exists(res_path))

        wb = openpyxl.load_workbook(res_path, data_only=False)
        expected_sheets = [
            'Caratula_Ejecutiva',
            'Conciliacion_CxC_102',
            'Contratos_Liquidados',
            'Ingresos_403_Rentas',
            'Activos_150_Vehiculos',
            'Discrepancias_y_Alertas'
        ]
        for s in expected_sheets:
            self.assertIn(s, wb.sheetnames)
        wb.close()


if __name__ == '__main__':
    unittest.main()
