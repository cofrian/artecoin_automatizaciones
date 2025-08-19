from __future__ import annotations

from pathlib import Path
import pandas as pd
from typing import Dict, List
from dataclasses import dataclass
import requests
from urllib.parse import urlparse
import re

"""
Obtener Fotos - Cargar datos desde Excel
----------------------------------------
Script que carga todas las hojas de Excel desde "CENTRO" hasta "CERRAMIENTOS"
en DataFrames separados para su posterior procesamiento.
"""

# --------------------------------------------------------------------
# CONFIGURACIÓN
# --------------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent

# Ruta del archivo Excel
EXCEL_PATH = (
    BASE_DIR.parent / "excel/inventario_20250814_1414.xlsx"
)

# Directorio para guardar las fotos descargadas
PHOTOS_DIR = BASE_DIR / "downloaded_photos"
PHOTOS_DIR.mkdir(exist_ok=True)


# --------------------------------------------------------------------
# FUNCIONES
# --------------------------------------------------------------------

@dataclass
class FotoExtractorExcel:
    excel_file: Path
    initial_sheet: str = "CENTRO"
    last_sheet: str = "CERRAMIENTOS"


    # ------------------- Cargar datos de Excel -------------------
    
    @staticmethod
    def _get_sheet_range(
        excel_file: Path, start_sheet: str, end_sheet: str
    ) -> List[str]:
        """
        Obtiene la lista de nombres de hojas desde start_sheet hasta end_sheet (inclusive).
        """
        with pd.ExcelFile(excel_file) as xls:
            all_sheets = xls.sheet_names

            start_idx = all_sheets.index(start_sheet)
            end_idx = all_sheets.index(end_sheet)

            return all_sheets[start_idx : end_idx + 1]


    @staticmethod
    def _load_excel_sheets(
        excel_file: Path, sheet_names: List[str]
    ) -> Dict[str, pd.DataFrame]:
        """
        Carga las hojas especificadas del archivo Excel en DataFrames.
        """
        dataframes = {}

        with pd.ExcelFile(excel_file) as xls:
            for sheet_name in sheet_names:
                print(f"-> Procesando hoja: {sheet_name}")

                try:
                    # Leer la hoja
                    df = pd.read_excel(
                        xls,
                        sheet_name=sheet_name,
                        header=0,
                        skiprows=None,
                        dtype=str,  # Leer todo como string para evitar conversiones automáticas
                    )

                    # Limpiar datos básicos
                    df = df.fillna("")  # Reemplazar NaN con strings vacíos

                    # Guardar en el diccionario
                    dataframes[sheet_name] = df

                except Exception as e:
                    print(f"     ! Error cargando hoja '{sheet_name}': {e}")
                    continue

        return dataframes


    def get_sheets_data(self) -> Dict[str, pd.DataFrame]:
        """
        Obtiene los datos de las hojas del archivo Excel entre START_SHEET y END_SHEET.
        """
        sheet_names = self._get_sheet_range(
            self.excel_file, self.initial_sheet, self.last_sheet
        )
        dataframes = self._load_excel_sheets(self.excel_file, sheet_names)
        return dataframes
    

    # ------------------- Crear directorios -------------------


    def _create_directories(self, dataframes: Dict[str, pd.DataFrame]) -> None:
        """
        Crea los directorios con la estructura:
        fotos / {id_centro}_{nombre_centro} / Referencias / {nombre_hoja} 
        """
        centro = dataframes['CENTRO']
        ids_and_names = centro['Etiqueta'].unique()
        
        for id_name in ids_and_names:
            # Limpiar el nombre para evitar problemas con espacios extra y caracteres especiales
            clean_id_name = str(id_name).strip()  # Remover espacios al inicio y final
            clean_id_name = re.sub(r'\s+', ' ', clean_id_name)  # Reemplazar múltiples espacios por uno solo
            clean_id_name = re.sub(r'[<>:"|?*\\]', '_', clean_id_name)  # Reemplazar caracteres no válidos en Windows
            
            # Crear un directorio para cada ID y nombre único
            dir_path = PHOTOS_DIR / clean_id_name / "Referencias"
            dir_path.mkdir(parents=True, exist_ok=True)
            
            for sheet_name in dataframes:
                # Crear subdirectorios para cada hoja
                sheet_dir = dir_path / sheet_name
                sheet_dir.mkdir(exist_ok=True)

    def _get_center_ids(self, dataframes: Dict[str, pd.DataFrame]) -> List[str]:
        """
        Obtiene todos los IDs únicos de centros desde la hoja CENTRO.
        """
        centro_df = dataframes.get('CENTRO')
        if centro_df is None:
            print("Error: No se encontró la hoja CENTRO")
            return []
        
        id_column = 'Etiqueta'
        
        if id_column not in centro_df.columns:
            print(f"Error: No se encontró la columna '{id_column}' en la hoja CENTRO")
            return []
        
        center_ids = centro_df[id_column].dropna().unique().tolist()
        # Limpiar los IDs también para mantener consistencia
        center_ids = [str(id_val).strip() for id_val in center_ids if str(id_val).strip()]
        return center_ids

    # ------------------- Descargar fotos -------------------

    @staticmethod
    def _extract_photo_columns(df: pd.DataFrame) -> List[str]:
        """
        Extrae las columnas que contienen fotos de un DataFrame.
        """
        photo_columns = []
        
        for col in df.columns:
            if "foto" in col.lower() and not df[col].empty:
                photo_columns.append(col)
                
        return photo_columns


    @staticmethod
    def _get_id_column_name(sheet_name: str) -> str:
        """
        Determina el nombre de la columna ID basado en el nombre de la hoja.
        
        Args:
            sheet_name: Nombre de la hoja de Excel
            
        Returns:
            str: Nombre de la columna ID correspondiente
        """
        # Casos especiales
        if sheet_name == "EDIFICIO":
            return "ID EDIFICACION"
        elif sheet_name == "CERRAMIENTOS":
            return "ID CERRAMIENTO"
        elif sheet_name == "DATOS_ELECTRICOS_EDIFICIOS":
            return "ID DATOS ELECTRICOS EDIFICIOS"
        elif sheet_name == "CENTRO":
            return "ID CENTRO"
        else:
            # Para el resto de hojas: remover acentos, convertir guiones bajos a espacios
            # y agregar "ID " al principio
            clean_name = sheet_name
            
            # Remover acentos
            clean_name = clean_name.replace('Á', 'A').replace('É', 'E').replace('Í', 'I')
            clean_name = clean_name.replace('Ó', 'O').replace('Ú', 'U').replace('Ñ', 'N')
            clean_name = clean_name.replace('á', 'a').replace('é', 'e').replace('í', 'i')
            clean_name = clean_name.replace('ó', 'o').replace('ú', 'u').replace('ñ', 'n')
            
            # Convertir guiones bajos a espacios
            clean_name = clean_name.replace('_', ' ')
            
            return f"ID {clean_name}"

    def _extract_filename_from_url(self, url, sheet_name, row_index, column_name, df):
        """
        Nombra los archivos con la estructura:
        {ID_ELEMENTO}_{Fxxx}.ext donde xxx es el número de foto
        """
        # Obtener el nombre de la columna ID para esta hoja
        id_column = self._get_id_column_name(sheet_name)
        
        # Obtener el ID del elemento de la fila actual
        elemento_id = "UNKNOWN"
        if id_column in df.columns:
            try:
                elemento_id = str(df.loc[row_index, id_column]).strip()
                if not elemento_id: # Esa fila no tiene ID ELEMENTO
                    raise ValueError(
                        f"Fila {row_index} en hoja '{sheet_name}' no tiene un ID válido en '{id_column}'"
                    )
                    # elemento_id = f"ROW_{row_index}"
            except (KeyError, IndexError):
                raise ValueError(
                    f"Error al obtener el ID del elemento en la fila {row_index} de la hoja '{sheet_name}'. "
                    f"Columna '{id_column}' no encontrada o índice fuera de rango."
                )
                # elemento_id = f"ROW_{row_index}"
        
        # Determinar el número de foto basado en el nombre de la columna
        foto_num = "001"  # Valor por defecto
        
        # Buscar número en el nombre de la columna (ej: "FOTO_1", "Foto 2", etc.)
        foto_match = re.search(r'(\d+)', column_name)
        if foto_match:
            foto_num = foto_match.group(1).zfill(3)  # Rellenar con ceros a la izquierda
        
        # Obtener extensión del archivo original
        parsed_url = urlparse(url)
        original_filename = ""
        
        if "resource_name=" in url:
            original_filename = url.split("resource_name=")[-1]
        else:
            original_filename = parsed_url.path.split("/")[-1]
        
        # Extraer extensión
        if original_filename and "." in original_filename:
            extension = original_filename.split(".")[-1].lower()
            # Validar que sea una extensión de imagen común
            if extension not in ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp']:
                extension = 'jpg'
        else:
            extension = 'jpg'
        
        # Limpiar el elemento_id para que sea seguro como nombre de archivo
        elemento_id_clean = re.sub(r'[<>:"|?*\\/]', '_', elemento_id)
        
        return f"{elemento_id_clean}_F{foto_num}.{extension}"

    def _download_photo_from_url(
        self, url: str, center_id: str, sheet_name: str, row_index: int, column_name: str, df: pd.DataFrame, matching_rows: pd.DataFrame
    ) -> str:
        """
        Descarga una foto desde una URL y la guarda en el directorio local.

        Args:
            url: URL de la foto a descargar
            center_id: ID del centro
            sheet_name: Nombre de la hoja de Excel
            row_index: Índice de la fila
            column_name: Nombre de la columna
            df: DataFrame completo
            matching_rows: Filas que coinciden con el centro

        Returns:
            str: Ruta del archivo descargado o mensaje de error
        """
        try:
            # Hacer la petición HTTP
            response = requests.get(url, timeout=30)
            response.raise_for_status()

            # Extraer el nombre del archivo con el nuevo formato
            filename = self._extract_filename_from_url(url, sheet_name, row_index, column_name, df, matching_rows)

            # Limpiar el center_id para usar el mismo formato que en _create_directories
            clean_center_id = str(center_id).strip()
            clean_center_id = re.sub(r'\s+', ' ', clean_center_id)
            clean_center_id = re.sub(r'[<>:"|?*\\]', '_', clean_center_id)

            # Usar directorio ya creado por _create_directories
            center_sheet_dir = PHOTOS_DIR / clean_center_id / "Referencias" / sheet_name

            # Crear ruta completa del archivo
            file_path = center_sheet_dir / filename

            # Guardar el archivo
            with open(file_path, "wb") as f:
                f.write(response.content)

            print(f"   ✓ Descargada: {filename}")
            return str(file_path)

        except requests.exceptions.RequestException as e:
            error_msg = f"   ✗ Error descargando {url}: {str(e)}"
            print(error_msg)
            return error_msg
        except Exception as e:
            error_msg = f"   ✗ Error inesperado: {str(e)}"
            print(error_msg)
            return error_msg

    
    @staticmethod
    def _extract_center_id_from_nombre_centro(nombre_centro: str) -> str:
        """
        Extrae el ID del centro desde el formato Cxxx_{nombre del centro}.
        
        Args:
            nombre_centro: String con formato Cxxx_{nombre}
            
        Returns:
            str: ID del centro (ej: "C001") o string vacío si no coincide el patrón
        """
        if not nombre_centro or not isinstance(nombre_centro, str):
            return ""
        
        # Buscar patrón Cxxx_ al inicio del string
        match = re.match(r'^(C\d+)_', nombre_centro.strip())
        if match:
            return match.group(1)
        
        return ""
    
    
    def _download_photos_for_center_in_sheet(
        self, center_id: str, sheet_name: str, df: pd.DataFrame
    ) -> Dict[str, List[str]]:
        """
        Descarga todas las fotos de un centro específico en una hoja específica.

        Args:
            center_id: ID del centro
            sheet_name: Nombre de la hoja
            df: DataFrame con los datos

        Returns:
            Dict con los resultados de descarga por columna
        """
        center_id_extracted = self._extract_center_id_from_nombre_centro(center_id)
        
        photo_columns: List[str] = self._extract_photo_columns(df)
        results = {}

        if not photo_columns:
            return results

        # Encontrar el ID del centro en la hoja
        if sheet_name == 'CENTRO':
            id_column = 'ID CENTRO'
            if id_column not in df.columns:
                print(f"   - No se encontró columna '{id_column}' en {sheet_name}")
                return results
            # Filtrar filas que coinciden con el center_id
            matching_rows = df[df[id_column].astype(str) == str(center_id_extracted)]
        else:
            # Para otras hojas, usar NOMBRE_CENTRO y extraer el ID
            id_column = 'NOMBRE_CENTRO'
            if id_column not in df.columns:
                print(f"   - No se encontró columna '{id_column}' en {sheet_name}")
                return results
            
            # Filtrar filas donde el ID extraído coincida con center_id
            mask = df[id_column].apply(self._extract_center_id_from_nombre_centro) == str(center_id_extracted)
            matching_rows = df[mask]
        
        if matching_rows.empty:
            return results

        print(f"   → Procesando centro {center_id} en hoja {sheet_name} ({len(matching_rows)} filas)")

        for col in photo_columns:
            results[col] = []
            
            for idx in matching_rows.index:
                url = df.loc[idx, col]
                if url and str(url).strip():  # Solo procesar URLs no vacías
                    result = self._download_photo_from_url(url, center_id, sheet_name, idx, col, df, matching_rows)
                    results[col].append(result)

        return results


    def download_all_photos(
        self, dataframes: Dict[str, pd.DataFrame]
    ) -> Dict[str, Dict[str, Dict[str, List[str]]]]:
        """
        Descarga todas las fotos organizadas por centro ID.

        Args:
            dataframes: Diccionario con los DataFrames de cada hoja

        Returns:
            Dict con todos los resultados organizados por centro, hoja y columna
        """
        all_results = {}

        print("Iniciando descarga de fotos...")

        # Obtener todos los IDs de centro
        center_ids = self._get_center_ids(dataframes)
        
        if not center_ids:
            print("No se encontraron IDs de centro para procesar")
            return all_results

        print(f"Se procesarán {len(center_ids)} centros")

        # Crear directorios
        self._create_directories(dataframes)

        # Para cada centro ID
        for center_id in center_ids:
            print(f"\nProcesando centro: {center_id}")
            all_results[center_id] = {}
            
            # Para cada hoja
            for sheet_name, df in dataframes.items():
                results = self._download_photos_for_center_in_sheet(center_id, sheet_name, df)
                if results:
                    all_results[center_id][sheet_name] = results

        print("\nDescarga completada")
        return all_results


if __name__ == "__main__":
    # Crear instancia del extractor
    extractor = FotoExtractorExcel(EXCEL_PATH)

    # Cargar datos de Excel
    dataframes = extractor.get_sheets_data()

    # Descargar fotos
    all_photos = extractor.download_all_photos(dataframes)

    # Mostrar resumen de fotos descargadas
    for center_id, center_results in all_photos.items():
        print(f"\nFotos descargadas del centro '{center_id}':")
        for sheet, results in center_results.items():
            for col, files in results.items():
                successful_downloads = len([f for f in files if not f.startswith("   ✗")])
                print(f"  Hoja '{sheet}' - Columna '{col}': {successful_downloads} fotos descargadas")