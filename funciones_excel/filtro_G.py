import xlwings as xw

def aplicar_filtro(filtro_deseado):
    wb = xw.Book.caller()

    slicercaches_objetivo = [
        "SegmentaciónDeDatos_CENTRO",
        "SegmentaciónDeDatos_CENTRO_8",
        "SegmentaciónDeDatos_CENTRO_10",
        "SegmentaciónDeDatos_CENTRO_11",
        "SegmentaciónDeDatos_CENTRO_15",
        "SegmentaciónDeDatos_CENTRO_7",
        "SegmentaciónDeDatos_CENTRO_1",
        "SegmentaciónDeDatos_CENTRO_2",
        "SegmentaciónDeDatos_CENTRO_3",
        "SegmentaciónDeDatos_CENTRO_4",
        "SegmentaciónDeDatos_CENTRO_5",
        "SegmentaciónDeDatos_CENTRO_6"
    ]

    print(f"\n🔍 Aplicando filtro: '{filtro_deseado}' en los slicers de Filtros_G")

    filtro_normalizado = filtro_deseado.strip().lower()

    for cache in wb.api.SlicerCaches:
        if cache.Name not in slicercaches_objetivo:
            continue

        slicers_en_filtros_g = [
            s for s in cache.Slicers
            if s.Shape.TopLeftCell.Worksheet.Name == "Filtros_G"
        ]
        if not slicers_en_filtros_g:
            continue

        nombres_items = [item.Name.strip().lower() for item in cache.SlicerItems]

        if filtro_normalizado in nombres_items:
            print(f"✅ '{filtro_deseado}' encontrado en {cache.Name} → aplicando selección")
            for item in cache.SlicerItems:
                item.Selected = (item.Name.strip().lower() == filtro_normalizado)
        else:
            print(f"⚠️ '{filtro_deseado}' no está en {cache.Name} → se omite")

    print("\n✅ Filtro aplicado a todos los slicers válidos.")
