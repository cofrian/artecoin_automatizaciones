#!/usr/bin/env python3
"""
Debug específico para fotos sin usar
"""

import re

def debug_fotos_sin_usar():
    """Debug de cada patrón problemático"""
    
    casos_problematicos = [
        # B001 - Bombas
        {"centro": "C0011", "foto": "B001_FB002", "tipos_posibles": ["EQHORIZ", "SISTCC", "ACOM"]},
        {"centro": "C0014", "foto": "B001_FB002", "tipos_posibles": ["EQHORIZ", "SISTCC", "ACOM"]},
        {"centro": "C0015", "foto": "B001_FB002", "tipos_posibles": ["EQHORIZ", "SISTCC", "ACOM"]},
        
        # B002 - Bombas
        {"centro": "C0011", "foto": "B002_FB002", "tipos_posibles": ["EQHORIZ", "SISTCC", "ACOM"]},
        {"centro": "C0015", "foto": "B002_FB002", "tipos_posibles": ["EQHORIZ", "SISTCC", "ACOM"]},
        {"centro": "C0020", "foto": "B002_FB002", "tipos_posibles": ["EQHORIZ", "SISTCC", "ACOM"]},
        
        # CDRO - Cuadros eléctricos
        {"centro": "C0011", "foto": "CDRO_SECUND_FD0007", "tipos_posibles": ["ACOM"]},
        {"centro": "C0011", "foto": "CDRO_SECUND_FD0008", "tipos_posibles": ["ACOM"]},
        {"centro": "C0049", "foto": "CDRO_SECUND_FD0008", "tipos_posibles": ["ACOM"]},
        
        # QH - Equipos horizontales
        {"centro": "C0011", "foto": "QH00010_FD0001", "tipos_posibles": ["EQHORIZ"]},
        {"centro": "C0011", "foto": "QH00011_FD0001", "tipos_posibles": ["EQHORIZ"]},
        {"centro": "C0011", "foto": "QH0008_FD0002", "tipos_posibles": ["EQHORIZ"]},
        
        # CR - Cerramientos/Envolventes  
        {"centro": "C0020", "foto": "CR00001_FF0002", "tipos_posibles": ["ENVOL"]},
        {"centro": "C0020", "foto": "CR00016_FF0002", "tipos_posibles": ["ENVOL"]},
        {"centro": "C0021", "foto": "CR00025_FVN0001", "tipos_posibles": ["ENVOL"]},
    ]
    
    print("🔍 DEBUG DE FOTOS SIN USAR")
    print("=" * 80)
    
    PATRON_MAPEO_TIPOS = {
        "ACOM": ["CDRO", "CUADRO", "ELECT", "ACOM"],
        "EQHORIZ": ["QH", "BOMB", "B00", "BOMBA", "EQUIPO"],
        "SISTCC": ["QG", "CALEF", "BOMB", "SIST", "SISTEMA", "B00"],
        "CLIMA": ["QE", "QI", "CLIMA", "CLIM", "AC", "HVAC"],
        "ILUM": ["I00", "ILUM", "LUZ", "LAMP", "LED"],
        "ENVOL": ["CR", "ENVOL", "CERR", "FACH", "VENT"],
        "ELEVA": ["QV", "ELEV", "ASCEN", "MONTAC"],
        "OTROSEQ": ["OTROS", "EQUIP", "MAQUIN"],
        "DEPENDENCIA": ["D00", "DEP", "SALA", "AULA"],
        "EDIFICIO": ["E00", "EDIF", "BLOQ", "NAVE"],
        "CENTRO": ["C00", "CENT", "FC"]
    }
    
    for i, caso in enumerate(casos_problematicos, 1):
        print(f"\n🎯 CASO {i}: {caso['centro']} - {caso['foto']}")
        print("-" * 50)
        
        foto_name = caso['foto']
        foto_clean = foto_name.upper()
        foto_parts = foto_clean.split('_')
        foto_id_part = foto_parts[0] if foto_parts else foto_clean
        
        print(f"📷 Foto limpia: {foto_clean}")
        print(f"🏷️  ID part: {foto_id_part}")
        print(f"🎭 Tipos posibles: {caso['tipos_posibles']}")
        
        for tipo in caso['tipos_posibles']:
            print(f"\n   🧪 PROBANDO TIPO: {tipo}")
            incluir_foto = False
            metodo_usado = None
            
            # Simular entidad ID típica
            if tipo == "ACOM":
                ident = f"{caso['centro']}E001"  # Ejemplo: C0011E001
            elif tipo == "EQHORIZ":
                ident = f"{caso['centro']}E001D001{foto_id_part}001"  # Con QH incluido
            elif tipo == "ENVOL":
                ident = f"{caso['centro']}E001{foto_id_part}"  # Con CR incluido
            else:
                ident = f"{caso['centro']}E001D001"  # Genérico
            
            print(f"   🆔 ID entidad simulado: {ident}")
            
            # MÉTODO 1: Exacto
            if foto_id_part in ident.upper():
                incluir_foto = True
                metodo_usado = "MÉTODO 1 (exacto)"
            
            # MÉTODO 2: Patrones por tipo
            if not incluir_foto and tipo in PATRON_MAPEO_TIPOS:
                patrones_tipo = PATRON_MAPEO_TIPOS[tipo]
                for patron in patrones_tipo:
                    if foto_clean.startswith(patron) or patron in foto_clean:
                        incluir_foto = True
                        metodo_usado = f"MÉTODO 2 (patrón {patron})"
                        break
            
            # MÉTODO 3: Equipos con numeración
            if not incluir_foto:
                if re.match(r'^[A-Z]{1,2}\d+', foto_id_part):
                    letra_match = re.match(r'^([A-Z]{1,2})', foto_id_part)
                    if letra_match:
                        prefijo_foto = letra_match.group(1)
                        if prefijo_foto in ident.upper():
                            incluir_foto = True
                            metodo_usado = f"MÉTODO 3 (prefijo {prefijo_foto})"
                
                # Bombas específicas
                if foto_id_part.startswith('B00') or foto_id_part.startswith('BOMB'):
                    if tipo in ["EQHORIZ", "SISTCC", "ACOM"]:
                        incluir_foto = True
                        metodo_usado = f"MÉTODO 3 (bomba para {tipo})"
            
            # MÉTODO 6: Proximidad equipos
            if not incluir_foto and tipo in ["EQHORIZ", "SISTCC", "CLIMA"]:
                equipos_generales = ["BOMBA", "BOMB", "EQUIP", "MAQUIN", "MOTOR", "PANEL"]
                if any(eq in foto_clean for eq in equipos_generales):
                    incluir_foto = True
                    metodo_usado = "MÉTODO 6 (proximidad equipos)"
            
            # MÉTODO 7: Cuadros ACOM
            if not incluir_foto and tipo == "ACOM":
                cuadros_electricos = ["CDRO", "CUADRO", "ELECT", "PANEL", "ARMARIO"]
                if any(cuadro in foto_clean for cuadro in cuadros_electricos):
                    incluir_foto = True
                    metodo_usado = "MÉTODO 7 (cuadros ACOM)"
            
            # Resultado
            if incluir_foto:
                print(f"   ✅ DEBERÍA INCLUIRSE ({metodo_usado})")
            else:
                print(f"   ❌ NO SE INCLUYE")
                
                # Análisis de por qué no
                print(f"   🔍 ANÁLISIS:")
                if tipo in PATRON_MAPEO_TIPOS:
                    patrones = PATRON_MAPEO_TIPOS[tipo]
                    print(f"      - Patrones de {tipo}: {patrones}")
                    matches = [p for p in patrones if p in foto_clean or foto_clean.startswith(p)]
                    if matches:
                        print(f"      - ⚠️  ENCONTRÉ MATCHES: {matches} (debería haber funcionado)")
                    else:
                        print(f"      - No hay matches en patrones")
                
                # Verificar si es bomba
                if foto_id_part.startswith('B00'):
                    print(f"      - Es bomba B00X, debería aplicar a EQHORIZ/SISTCC/ACOM")
                
                # Verificar prefijo
                if re.match(r'^[A-Z]{1,2}\d+', foto_id_part):
                    letra_match = re.match(r'^([A-Z]{1,2})', foto_id_part)
                    if letra_match:
                        prefijo = letra_match.group(1)
                        print(f"      - Prefijo '{prefijo}' debería buscarse en entidades con {prefijo}")

if __name__ == "__main__":
    debug_fotos_sin_usar()
