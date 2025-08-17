
Plantillas A3 (ACOM-style) unificadas
====================================

- Misma estructura (layout, tipografías, tabla .kv, etc.).
- Bloque de fotos MAX-FILL sin recuadro: siempre ocupa el máximo espacio posible.
- Placeholders con {{...}} son literales para sustitución por string (no Jinja).

Cómo usar con tu script:
------------------------
1) Carga el HTML como texto y sustituye todas las llaves "{{campo}}" por
   los valores del JSON (claves planas: id_centro, id_edificio, ...).

2) Reemplaza el marcador de fotos de cada plantilla por el bloque generado
   por tu script (ejemplo):

   Marcador en la plantilla (depende del tipo):
     [[FOTOS_ACOM]]        [[FOTOS_CC]]            [[FOTOS_CENTRO]]
     [[FOTOS_EDIFICIOS]]   [[FOTOS_ELEVA]]         [[FOTOS_ENVOL_CUBIERTA]]
     [[FOTOS_ENVOL_FACHADA]] [[FOTOS_ENVOL_PUERTAS]] [[FOTOS_ENVOL_VENTANAS]]
     [[FOTOS_EQH]]         [[FOTOS_ILUM]]          [[FOTOS_OTROS]]

   Sustitúyelo por algo así (elige la clase photos-1/2/.../many según nº fotos):

     <div class="ph-grid photos-4">
       <figure class="ph-card">
         <div class="ph-imgwrap"><img class="ph-img" src="ruta/a/foto1.jpg" alt="foto1"></div>
         <figcaption class="ph-cap">foto1</figcaption>
       </figure>
       <figure class="ph-card">
         <div class="ph-imgwrap"><img class="ph-img" src="ruta/a/foto2.jpg" alt="foto2"></div>
         <figcaption class="ph-cap">foto2</figcaption>
       </figure>
       <!-- ... -->
     </div>

   Reglas de columnas (max-fill):
     - 1–2 fotos  -> 1 columna (photos-1/2)
     - 3–6 fotos  -> 2 columnas (photos-3 .. photos-6)
     - >6 fotos   -> auto-fill (photos-many)

3) El fondo SVG debe estar al lado del HTML: A3_FOTOS_AUDITORIA_SIN_ICONO.svg

Notas:
- Envolvente está separada por tipo: cubierta, fachada, puertas y ventanas (archivos distintos).
