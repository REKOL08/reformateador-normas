# Reformateador de documentos académicos

Convierte un `.docx` o `.pdf` en un `.docx` nuevo con el formato de **APA 7, IEEE, MLA 9, Chicago 17, Vancouver o ICONTEC (NTC 1486)**.
Solo cambia el formato; el texto no se modifica (el script lo verifica al final).

## Instalación
```bash
python -m venv .venv && source .venv/bin/activate     # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```
Python 3.9+. Tkinter viene con Python en Windows/macOS; en Linux: `sudo apt install python3-tk`.
LibreOffice es opcional (para `.doc/.odt/.rtf` y para exportar a PDF).

## Uso
```bash
python reformateador.py                         # ventana con arrastrar y soltar
python reformateador.py tesis.docx -n apa7
python reformateador.py tesis.docx -n mla9 --apellido Labrada
python reformateador.py articulo.docx -n ieee --columnas 2 --numerar-refs
python reformateador.py tesis.docx -n icontec --ordenar-refs --pdf
python reformateador.py --listar-normas
```
Opciones: `--papel carta|a4|original`, `--encabezado "TÍTULO CORTO"`, `--ordenar-refs`, `--numerar-refs`,
`--sin-detectar-titulos`, `--columnas 1|2`, `--pdf`.

## Qué hace
Márgenes y papel · fuente y tamaño · interlineado · sangría de párrafo · títulos por nivel · citas en bloque ·
listas · leyendas de tablas/figuras · referencias con sangría francesa (orden y numeración opcionales) ·
numeración de páginas (campo PAGE) · encabezado corrido.

Para `.docx` edita una copia del original: se conservan imágenes, tablas, notas al pie, hipervínculos y tabla de contenido.

## Límites (importante)
- **No convierte las citas dentro del texto** entre estilos (p. ej. `(García, 2020)` ↔ `[1]`). Avisa si detecta un estilo distinto al de la norma.
- **Desde PDF solo se recupera el texto** (títulos por tamaño/negrita, párrafos, cursivas). Se pierden imágenes y tablas, y los PDF de varias columnas pueden salir desordenados. Si tiene el Word original, úselo. Los PDF escaneados requieren OCR previo.
- No cambia los bordes de las tablas ni genera portada.
- Los títulos sin estilo de Word se detectan por su aspecto: revise el resultado.
- Los valores de cada norma están en el diccionario `NORMAS` y son una aproximación de los manuales. Verifique con la guía de su institución (sobre todo ICONTEC, que muchas universidades adaptan) y edite ahí lo que difiera.
- Al deshacer guiones de fin de línea en PDF, una palabra compuesta partida justo en el guion (`socio-económico`) puede quedar unida.
