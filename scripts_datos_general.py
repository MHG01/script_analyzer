import pandas as pd
import logging
from pathlib import Path
from typing import Dict, List, Union, Optional
from datetime import datetime


class ExcelAnalyzer:
    """
    Clase mejorada para analizar archivos Excel con funcionalidades específicas para análisis financiero y no financiero.
    """
    AUTOR = "M.H.G."
    COLUMNAS_FINANCIERAS = [
        'monto', 'subtotal', 'iva', 'total', 'descuento',
        'impuesto', 'factura', 'precio'
    ]

    def __init__(self, archivo_excel: Union[str, Path], configuracion: Optional[Dict] = None):
        self.logger = None
        self.df = None
        self.directorio_salida = None
        self.columnas_monetarias = []
        self.columnas_no_monetarias = []
        self.configuracion = configuracion or {}

        self._configurar_logging()
        self.archivo_excel = self._validar_archivo(archivo_excel)
        self.df = self._cargar_archivo()
        self.reporte_fecha = datetime.now().strftime("%Y-%m-%d")
        self._crear_directorio_salida()
        self._identificar_columnas()

    def _configurar_logging(self) -> None:
        """Configura el sistema de logging."""
        log_dir = Path('logs')
        log_dir.mkdir(exist_ok=True)
        log_file = log_dir / 'excel_analyzer.log'

        self.logger = logging.getLogger('ExcelAnalyzer')
        self.logger.setLevel(logging.INFO)

        if not self.logger.handlers:
            formatter = logging.Formatter(
                '%(asctime)s - %(levelname)s - %(message)s')
            file_handler = logging.FileHandler(log_file)
            file_handler.setFormatter(formatter)
            self.logger.addHandler(file_handler)

            console_handler = logging.StreamHandler()
            console_handler.setFormatter(formatter)
            self.logger.addHandler(console_handler)

    def _validar_archivo(self, archivo_excel: Union[str, Path]) -> Path:
        """Valida la existencia del archivo Excel y lo retorna como Path."""
        try:
            archivo_path = Path(archivo_excel)
            if archivo_path.is_file():
                return archivo_path

            if archivo_excel.isdigit():
                num = int(archivo_excel)
                archivos_excel = list(Path().glob("**/*.xls*"))
                if 1 <= num <= len(archivos_excel):
                    return archivos_excel[num - 1]

            posibles_rutas = list(Path().glob(f"**/{archivo_excel}"))
            if posibles_rutas:
                if len(posibles_rutas) == 1:
                    return posibles_rutas[0]

                print("\nSe encontraron varios archivos posibles:")
                for i, ruta in enumerate(posibles_rutas, 1):
                    print(f"{i}. {ruta}")

                while True:
                    try:
                        seleccion = int(
                            input("\nSeleccione el número del archivo correcto: "))
                        if 1 <= seleccion <= len(posibles_rutas):
                            return posibles_rutas[seleccion - 1]
                    except ValueError:
                        print("Por favor, ingrese un número válido.")

            self._mostrar_archivos_disponibles()
            raise FileNotFoundError(
                f"No se encontró el archivo Excel: {archivo_excel}")

        except Exception as e:
            self.logger.error(f"Error al validar archivo: {str(e)}")
            raise

    def _mostrar_archivos_disponibles(self) -> None:
        """Muestra los archivos Excel disponibles en el directorio."""
        print("\nArchivos Excel disponibles en el directorio:")
        archivos_excel = list(Path().glob("**/*.xls*"))

        if not archivos_excel:
            print("No se encontraron archivos Excel en el directorio.")
        else:
            for i, archivo in enumerate(archivos_excel, 1):
                print(f"{i}. {archivo}")

    def _cargar_archivo(self) -> pd.DataFrame:
        """Carga el archivo Excel permitiendo selección de hoja."""
        try:
            xls = pd.ExcelFile(self.archivo_excel)
            self._mostrar_hojas_disponibles(xls.sheet_names)

            while True:
                try:
                    seleccion = int(
                        input("Seleccione el número de la hoja a analizar: ")) - 1
                    if 0 <= seleccion < len(xls.sheet_names):
                        break
                    print("Selección no válida. Intente de nuevo.")
                except ValueError:
                    print("Por favor, ingrese un número válido.")

            hoja_seleccionada = xls.sheet_names[seleccion]
            df = pd.read_excel(xls, sheet_name=hoja_seleccionada)
            self._mostrar_info_columnas(df)

            # Imprimir información de depuración
            self.logger.info(f"Tipos de datos de las columnas:\n{df.dtypes}")
            self.logger.info(f"Primeras filas del DataFrame:\n{df.head()}")

            return self._preparar_columnas_financieras(df)
        except Exception as e:
            self.logger.error(f"Error al leer el archivo Excel: {str(e)}")
            self.logger.error(f"Detalles del error:", exc_info=True)
            raise

    def _preparar_columnas_financieras(self, df: pd.DataFrame) -> pd.DataFrame:
        """Prepara las columnas financieras del DataFrame."""
        for columna in df.columns:
            columna_lower = columna.lower()
            if any(termino in columna_lower for termino in self.COLUMNAS_FINANCIERAS):
                df[columna] = pd.to_numeric(df[columna], errors='coerce')
                df[columna] = df[columna].fillna(0)
        return df

    def _mostrar_hojas_disponibles(self, hojas: List[str]) -> None:
        """Muestra las hojas disponibles en el archivo Excel."""
        print("\nHojas disponibles:")
        for i, hoja in enumerate(hojas, 1):
            print(f"{i}. {hoja}")

    def _mostrar_info_columnas(self, df: pd.DataFrame) -> None:
        """Muestra información sobre las columnas del DataFrame."""
        print("\nColumnas encontradas:")
        for col in df.columns:
            print(f"- {col}")

    def _crear_directorio_salida(self) -> None:
        """Crea el directorio de salida para los reportes."""
        self.directorio_salida = Path(f'reportes_{self.reporte_fecha}')
        self.directorio_salida.mkdir(exist_ok=True)
        self.logger.info(f"Directorio de salida creado: {
                         self.directorio_salida}")

    def _preparar_columnas_financieras(self, df: pd.DataFrame) -> pd.DataFrame:
        """Prepara las columnas financieras del DataFrame."""
        for columna in df.columns:
            if isinstance(columna, str):  # Verificar si la columna es una cadena
                columna_lower = columna.lower()
                if any(termino in columna_lower for termino in self.COLUMNAS_FINANCIERAS):
                    df[columna] = pd.to_numeric(df[columna], errors='coerce')
                    df[columna] = df[columna].fillna(0)
        return df

    def _identificar_columnas(self) -> None:
        """Identifica y clasifica las columnas del DataFrame."""
        self.columnas_monetarias = [col for col in self.df.columns if
                                    # Verificar si la columna es una cadena
                                    isinstance(col, str) and
                                    any(termino in col.lower() for termino in self.COLUMNAS_FINANCIERAS)]
        self.columnas_no_monetarias = [
            col for col in self.df.columns if col not in self.columnas_monetarias]

        self.logger.info(f"Columnas monetarias identificadas: {
                         self.columnas_monetarias}")
        self.logger.info(f"Columnas no monetarias identificadas: {
                         self.columnas_no_monetarias}")

    def calcular_totales_financieros(self) -> Dict[str, float]:
        """Calcula los totales financieros del DataFrame de manera adaptable."""
        totales = {}
        for columna in self.columnas_monetarias:
            total_columna = self.df[columna].sum()
            totales[f'Total {columna}'] = total_columna

            # Cálculo de IVA configurable
            if self.configuracion.get('calcular_iva', True) and 'iva' not in columna.lower():
                iva_rate = self.configuracion.get('iva_rate', 0.16)
                iva = total_columna * iva_rate
                totales[f'IVA de {columna}'] = iva
                totales[f'Total con IVA de {columna}'] = total_columna + iva

        # Cálculo del total de factura
        totales['Total Factura'] = sum(
            valor for clave, valor in totales.items() if 'Total con IVA' in clave)

        return totales

    def generar_reporte_html(self) -> Path:
        """Genera un reporte HTML con análisis financiero detallado en formato profesional."""
        try:
            totales_financieros = self.calcular_totales_financieros()

            # Preparar los datos de la tabla principal
            tabla_datos = self.df.to_html(classes='styled-table', index=False,
                                          float_format=lambda x: f"${x:,.2f}" if isinstance(x, (int, float)) else x)

            # Preparar el resumen financiero en formato texto
            resumen_financiero = ""
            for concepto, valor in totales_financieros.items():
                if concepto != 'Total Factura':  # Excluimos el total de la factura del resumen
                    resumen_financiero += f"<div class='resumen-item'><span>{
                        concepto}:</span> <strong>${valor:,.2f}</strong></div>"

            html_content = f"""
            <!DOCTYPE html>
            <html lang="es">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Reporte Financiero - {self.reporte_fecha}</title>
                <style>
                    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;700&display=swap');

                    body {{
                        font-family: 'Roboto', sans-serif;
                        line-height: 1.6;
                        color: #333;
                        margin: 0;
                        padding: 0;
                        background-color: #f5f5f5;
                    }}
                    .container {{
                        max-width: 100%;
                        margin: 0 auto;
                        padding: 20px;
                        background-color: #ffffff;
                        box-shadow: 0 0 20px rgba(0, 0, 0, 0.1);
                        overflow-x: auto;
                    }}
                    .date {{
                        text-align: right;
                        font-size: 0.9em;
                        color: #777;
                        margin-bottom: 20px;
                    }}
                    h1 {{
                        color: #2c3e50;
                        text-align: center;
                        font-size: 2.5em;
                        margin-bottom: 30px;
                        border-bottom: 2px solid #3498db;
                        padding-bottom: 10px;
                    }}
                    .styled-table {{
                        width: 100%;
                        border-collapse: collapse;
                        margin: 25px 0;
                        font-size: 0.9em;
                        box-shadow: 0 0 20px rgba(0, 0, 0, 0.15);
                    }}
                    .styled-table thead tr {{
                        background-color: #3498db;
                        color: #ffffff;
                        text-align: left;
                    }}
                    .styled-table th,
                    .styled-table td {{
                        padding: 12px 15px;
                        white-space: nowrap;
                    }}
                    .styled-table tbody tr {{
                        border-bottom: 1px solid #dddddd;
                    }}
                    .styled-table tbody tr:nth-of-type(even) {{
                        background-color: #f3f3f3;
                    }}
                    .styled-table tbody tr:last-of-type {{
                        border-bottom: 2px solid #3498db;
                    }}
                    .financial-footer {{
                        display: flex;
                        justify-content: space-between;
                        margin-top: 40px;
                        padding-top: 20px;
                        border-top: 2px solid #eee;
                    }}
                    .resumen-financiero {{
                        background-color: #f8f9fa;
                        padding: 20px;
                        border-radius: 6px;
                        width: 48%;
                        box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
                    }}
                    .resumen-item {{
                        display: flex;
                        justify-content: space-between;
                        margin-bottom: 10px;
                    }}
                    .total-factura {{
                        font-size: 1.2em;
                        color: #2c3e50;
                        padding: 20px;
                        background-color: #e8f4f8;
                        border-radius: 6px;
                        width: 48%;
                        box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
                        display: flex;
                        justify-content: space-between;
                        align-items: center;
                    }}
                    .total-factura strong {{
                        font-size: 1.5em;
                        color: #3498db;
                    }}
                    footer {{
                        margin-top: 40px;
                        text-align: center;
                        font-size: 0.8em;
                        color: #777;
                    }}
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="date">
                        {self.reporte_fecha}
                    </div>
                    <h1>Reporte Financiero Detallado</h1>

                    <div style="overflow-x: auto;">
                        {tabla_datos}
                    </div>

                    <div class="financial-footer">
                        <div class="resumen-financiero">
                            <h3>Resumen Financiero</h3>
                            {resumen_financiero}
                        </div>
                        <div class="total-factura">
                            <span>Total Factura:</span>
                            <strong>${totales_financieros['Total Factura']:,.2f}</strong>
                        </div>
                    </div>

                    <footer>
                        <p>Generado por: {self.AUTOR}</p>
                    </footer>
                </div>
            </body>
            </html>
            """

            nombre_archivo = self.directorio_salida / \
                f'reporte_financiero_{self.reporte_fecha}.html'
            nombre_archivo.write_text(html_content, encoding='utf-8')

            self.logger.info(f"Reporte HTML generado: {nombre_archivo}")
            return nombre_archivo

        except Exception as e:
            self.logger.error(f"Error al generar reporte HTML: {str(e)}")
            raise


def main():
    print("="*50)
    print(f"Analizador de Datos Excel - {ExcelAnalyzer.AUTOR}")
    print("="*50)

    try:
        archivos_excel = list(Path().glob("**/*.xls*"))

        if archivos_excel:
            print("\nArchivos Excel encontrados:")
            for i, archivo in enumerate(archivos_excel, 1):
                print(f"{i}. {archivo}")
            print("\nPuede escribir el nombre del archivo o su número de la lista.")
        else:
            print("No se encontraron archivos Excel en el directorio actual.")

        archivo_excel = input(
            "\nIngrese el nombre o número del archivo Excel a analizar: ").strip()

        analizador = ExcelAnalyzer(archivo_excel)
        print("Columnas del DataFrame:",
              analizador.df.columns.tolist())  # Depuración
        reporte = analizador.generar_reporte_html()

        print("\n¡Análisis completado exitosamente!")
        print(f"Reporte guardado en: {reporte}")
        print("\nDesarrollado por: Michael Haring García")

    except FileNotFoundError:
        print("\nError: No se pudo encontrar el archivo.")
        print("Asegúrese de que el archivo existe y tiene permisos de lectura.")
    except ValueError as e:
        print(f"\nError: {e}")
    except Exception as e:
        print(f"\nError inesperado: {str(e)}")
        print("Consulte el archivo de log para más detalles.")


if __name__ == "__main__":
    main()
