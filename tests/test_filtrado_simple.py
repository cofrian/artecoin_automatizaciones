#!/usr/bin/env python3
"""
Script para verificar si nuestras correcciones funcionan aplicando el filtro directamente
a las rutas de fotos que vimos en el HTML problemático.
"""

import re

def verificar_filtro_centro(ident, rutas_fotos, tipo="CLIMA"):
    """
    Simula la lógica de filtrado corregida para verificar si funciona.
    """
    print(f"🔍 Verificando filtro para ID: {ident} (tipo: {tipo})")
    print(f"📸 Fotos a evaluar: {len(rutas_fotos)}")
    
    # Extraer el código de centro del ID
    centro_match = re.match(r'(C\d+)', ident.upper())
    if not centro_match:
        print("❌ No se pudo extraer el código de centro")
        return []
    
    centro_codigo = centro_match.group(1)
    print(f"🏢 Centro esperado: {centro_codigo}")
    
    fotos_validas = []
    
    for ruta in rutas_fotos:
        # Extraer el nombre del archivo de la ruta
        nombre_archivo = ruta.split('/')[-1]
        stem_upper = nombre_archivo.split('.')[0].upper()
        
        print(f"\n📝 Evaluando: {nombre_archivo}")
        print(f"   Ruta: {ruta}")
        print(f"   Stem: {stem_upper}")
        
        es_valida = False
        
        if tipo in ["CLIMA", "EQHORIZ", "ELEVA", "OTROSEQ", "ILUM", "ENVOL", "SISTCC"]:
            # EQUIPOS: solo fotos cuyo nombre esté contenido en el ID de la entidad
            if 'Q' in stem_upper:
                # Extraer la parte del nombre de la foto (ej: QI007_FQI0001 -> QI007)
                foto_parts = stem_upper.split('_')
                foto_id_part = foto_parts[0] if foto_parts else stem_upper
                
                print(f"   ID parte foto: {foto_id_part}")
                print(f"   ¿ID en entidad?: {foto_id_part in ident.upper()}")
                
                # Verificar que la parte del nombre de foto está contenida en el ID de la entidad
                if foto_id_part in ident.upper():
                    # NUEVA VERIFICACIÓN: la foto debe estar en la carpeta del centro correcto
                    foto_path_str = ruta.upper()
                    centro_en_ruta = centro_codigo in foto_path_str
                    
                    print(f"   ¿Centro en ruta?: {centro_en_ruta}")
                    
                    if centro_en_ruta:
                        es_valida = True
                        print(f"   ✅ VÁLIDA")
                    else:
                        print(f"   ❌ RECHAZADA (centro incorrecto)")
                else:
                    print(f"   ❌ RECHAZADA (ID no coincide)")
            else:
                print(f"   ❌ RECHAZADA (no es equipo Q)")
        
        if es_valida:
            fotos_validas.append(ruta)
    
    print(f"\n📊 RESUMEN:")
    print(f"   Fotos evaluadas: {len(rutas_fotos)}")
    print(f"   Fotos válidas: {len(fotos_validas)}")
    print(f"   Centro objetivo: {centro_codigo}")
    
    for foto in fotos_validas:
        print(f"   ✅ {foto}")
    
    return fotos_validas

# Datos del caso problemático que vimos en el diagnóstico
ident_problema = "C0007E001D0007QI007"
rutas_problematicas = [
    "file:///C:/Users/IGP/Desktop/02_ENTREGA%20SONINGEO/1_CONSULTA%201/C0026_CENTRO%20ATENCI%C3%93N%20A%20DROGODEPENDENCIAS/Referencias/QI007_FQI0001.jpg",
    "file:///C:/Users/IGP/Desktop/02_ENTREGA%20SONINGEO/2_CONSULTA%202/C0001_COMPLEJO%20DEPORTIVO%20%E2%80%9CJOSE%20ANTONIO%20SAMARANCH%E2%80%9D/Referencias/QI007_FQI0001.jpg",
    "file:///C:/Users/IGP/Desktop/02_ENTREGA%20SONINGEO/2_CONSULTA%202/C0016_COLEGIO%20%E2%80%9CVIRGEN%20DE%20LOS%20REMEDIOS%E2%80%9D/Referencias/QI007_FQI0001.jpg",
    "file:///C:/Users/IGP/Desktop/02_ENTREGA%20SONINGEO/1_CONSULTA%201/C0049_COLEGIO%20%E2%80%9CFEDERICO%20GARC%C3%8DA%20LORCA%E2%80%9D/Referencias/QI007_FQI0001.jpg",
    "file:///C:/Users/IGP/Desktop/02_ENTREGA%20SONINGEO/1_CONSULTA%201/C0046_CUARTEL%20DE%20LA%20GUARDIA%20CIVIL/Referencias/QI007_FQI0001.jpg",
    "file:///C:/Users/IGP/Desktop/02_ENTREGA%20SONINGEO/1_CONSULTA%201/C0023_CENTRO%20CULTURAL%20%E2%80%9CPABLO%20NERUDA%E2%80%9D/Referencias/QI007_FQI0001.jpg"
]

print("🧪 PRUEBA DE CORRECCIÓN DE FILTRADO POR CENTRO")
print("=" * 60)

fotos_correctas = verificar_filtro_centro(ident_problema, rutas_problematicas)

print(f"\n🎯 RESULTADO FINAL:")
if len(fotos_correctas) == 0:
    print("✅ ¡CORRECCIÓN EXITOSA! Se rechazaron todas las fotos de centros incorrectos.")
    print("   (Esto es correcto porque ninguna foto es del centro C0007)")
else:
    print(f"⚠️  Se aceptaron {len(fotos_correctas)} fotos:")
    for foto in fotos_correctas:
        print(f"   📸 {foto}")
