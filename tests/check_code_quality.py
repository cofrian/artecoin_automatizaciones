#!/usr/bin/env python3
"""
Script de verificación automática de calidad de código para el proyecto artecoin_automatizaciones.
"""

import ast
import os
import sys
from pathlib import Path
from typing import List, Tuple

def check_python_syntax(file_path: Path) -> List[str]:
    """Verifica la sintaxis de un archivo Python."""
    errors = []
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            source = f.read()
        
        # Verificar sintaxis
        ast.parse(source)
        
        # Verificar importaciones redundantes
        lines = source.split('\n')
        imports_found = set()
        local_imports = []
        
        for i, line in enumerate(lines, 1):
            stripped = line.strip()
            
            # Detectar importaciones globales
            if stripped.startswith('import ') and ' as ' not in stripped:
                module = stripped.replace('import ', '').split(',')[0].strip()
                if module in imports_found:
                    errors.append(f"Línea {i}: Importación duplicada de '{module}'")
                imports_found.add(module)
            
            # Detectar importaciones locales redundantes
            if '    import ' in line or '        import ' in line:
                module = stripped.replace('import ', '').split(',')[0].strip()
                if module in imports_found:
                    local_imports.append(f"Línea {i}: Importación local redundante de '{module}' (ya importado globalmente)")
        
        errors.extend(local_imports[:10])  # Limitar a 10 para no saturar
        
    except SyntaxError as e:
        errors.append(f"Error de sintaxis en línea {e.lineno}: {e.msg}")
    except Exception as e:
        errors.append(f"Error al procesar archivo: {e}")
    
    return errors

def check_code_quality(file_path: Path) -> List[str]:
    """Verifica problemas de calidad de código."""
    warnings = []
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
            lines = content.split('\n')
        
        # Verificar líneas muy largas
        for i, line in enumerate(lines, 1):
            if len(line) > 120:
                warnings.append(f"Línea {i}: Línea muy larga ({len(line)} caracteres)")
                if len(warnings) >= 5:  # Limitar warnings de líneas largas
                    break
        
        # Verificar funciones duplicadas
        function_names = []
        for i, line in enumerate(lines, 1):
            if line.strip().startswith('def '):
                func_name = line.split('(')[0].replace('def ', '').strip()
                if func_name in function_names:
                    warnings.append(f"Línea {i}: Función duplicada '{func_name}'")
                function_names.append(func_name)
        
    except Exception as e:
        warnings.append(f"Error al verificar calidad: {e}")
    
    return warnings

def scan_directory(directory: Path, extensions: List[str] = ['.py']) -> List[Tuple[Path, List[str], List[str]]]:
    """Escanea un directorio en busca de archivos con problemas."""
    results = []
    
    for ext in extensions:
        for file_path in directory.rglob(f'*{ext}'):
            if file_path.is_file():
                errors = check_python_syntax(file_path)
                warnings = check_code_quality(file_path)
                
                if errors or warnings:
                    results.append((file_path, errors, warnings))
    
    return results

def main():
    project_root = Path(__file__).parent
    print(f"Verificando calidad de codigo en: {project_root}")
    print("=" * 60)
    
    # Escanear archivos Python
    results = scan_directory(project_root, ['.py'])
    
    if not results:
        print("✅ ¡Excelente! No se encontraron problemas de código.")
        return
    
    total_errors = 0
    total_warnings = 0
    
    for file_path, errors, warnings in results:
        rel_path = file_path.relative_to(project_root)
        
        if errors:
            print(f"\n[!] ERRORES en {rel_path}:")
            for error in errors:
                print(f"   • {error}")
            total_errors += len(errors)
        
        if warnings:
            print(f"\n⚠️  ADVERTENCIAS en {rel_path}:")
            for warning in warnings[:5]:  # Limitar a 5 warnings por archivo
                print(f"   • {warning}")
            if len(warnings) > 5:
                print(f"   ... y {len(warnings) - 5} advertencias más")
            total_warnings += len(warnings)
    
    print(f"\nRESUMEN:")
    print(f"   Archivos revisados: {len([f for f in project_root.rglob('*.py') if f.is_file()])}")
    print(f"   Archivos con problemas: {len(results)}")
    print(f"   Total errores: {total_errors}")
    print(f"   Total advertencias: {total_warnings}")
    
    if total_errors > 0:
        print(f"\n🚨 Se encontraron {total_errors} errores que necesitan corrección.")
        sys.exit(1)
    else:
        print(f"\n✅ No se encontraron errores críticos. Solo {total_warnings} advertencias.")

if __name__ == "__main__":
    main()
