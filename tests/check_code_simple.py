#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de verificacion de calidad de codigo
Verifica imports redundantes, funciones duplicadas y problemas comunes
"""

import ast
import os
from pathlib import Path
from collections import defaultdict


def analyze_file(filepath):
    """Analiza un archivo Python en busca de problemas de calidad"""
    errors = []
    warnings = []
    
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Parse AST
        tree = ast.parse(content)
        
        # Recopilar imports y funciones
        global_imports = set()
        local_imports = defaultdict(list)  # {module: [line_numbers]}
        functions = defaultdict(list)  # {function_name: [line_numbers]}
        
        for node in ast.walk(tree):
            # Imports globales
            if isinstance(node, (ast.Import, ast.ImportFrom)):
                if hasattr(node, 'col_offset') and node.col_offset == 0:  # Import global
                    if isinstance(node, ast.Import):
                        for alias in node.names:
                            global_imports.add(alias.name)
                    elif isinstance(node, ast.ImportFrom) and node.module:
                        global_imports.add(node.module)
            
            # Imports locales (dentro de funciones)
            elif isinstance(node, ast.FunctionDef):
                # Buscar imports dentro de la función
                for child in ast.walk(node):
                    if isinstance(child, (ast.Import, ast.ImportFrom)):
                        if isinstance(child, ast.Import):
                            for alias in child.names:
                                if alias.name in global_imports:
                                    local_imports[alias.name].append(child.lineno)
                        elif isinstance(child, ast.ImportFrom) and child.module:
                            if child.module in global_imports:
                                local_imports[child.module].append(child.lineno)
            
            # Funciones duplicadas
            if isinstance(node, ast.FunctionDef):
                functions[node.name].append(node.lineno)
        
        # Detectar problemas
        # 1. Imports redundantes
        for module, lines in local_imports.items():
            if lines:
                warnings.append(f"Import redundante '{module}' en lineas {lines} - ya esta importado globalmente")
        
        # 2. Funciones duplicadas
        for func_name, lines in functions.items():
            if len(lines) > 1:
                errors.append(f"Funcion duplicada '{func_name}' en lineas {lines}")
        
        # 3. Verificar sintaxis basica
        try:
            compile(content, filepath, 'exec')
        except SyntaxError as e:
            errors.append(f"Error de sintaxis en linea {e.lineno}: {e.msg}")
        
    except Exception as e:
        errors.append(f"Error al analizar archivo: {e}")
    
    return errors, warnings


def main():
    project_root = Path(__file__).parent
    print(f"Verificando calidad de codigo en: {project_root}")
    print("=" * 60)
    print()
    
    total_files = 0
    files_with_issues = 0
    total_errors = 0
    total_warnings = 0
    
    for root, dirs, files in os.walk(project_root):
        # Ignorar carpetas de cache
        dirs[:] = [d for d in dirs if not d.startswith('.') and d != '__pycache__']
        
        for file in files:
            if file.endswith('.py'):
                filepath = os.path.join(root, file)
                rel_path = os.path.relpath(filepath, project_root)
                
                total_files += 1
                
                errors, warnings = analyze_file(filepath)
                
                if errors:
                    files_with_issues += 1
                    print(f"\n[ERROR] {rel_path}:")
                    for error in errors:
                        total_errors += 1
                        print(f"  -> {error}")
                
                if warnings:
                    if not errors:
                        files_with_issues += 1
                    print(f"\n[AVISO] {rel_path}:")
                    for warning in warnings:
                        total_warnings += 1
                        print(f"  -> {warning}")
    
    # Resumen final
    print("\n" + "=" * 40)
    print("RESUMEN:")
    print("=" * 40)
    print(f"Total de archivos: {total_files}")
    print(f"Archivos con problemas: {files_with_issues}")
    print(f"Errores criticos: {total_errors}")
    print(f"Advertencias: {total_warnings}")
    print()
    
    if files_with_issues > 0:
        print("Recomendacion: Corregir los problemas identificados.")
    else:
        print("Todo el codigo esta limpio!")


if __name__ == "__main__":
    main()
