# Reformateador de documentos académicos

Programa gratuito que toma tu trabajo (**Word o PDF**) y te entrega un **Word nuevo con el formato de la norma que elijas**:
APA 7 · IEEE · MLA 9 · Chicago 17 · Vancouver · ICONTEC.

- Cambia **solo el formato** (márgenes, letra, interlineado, títulos, sangrías, referencias, numeración de páginas).
- **No cambia tu texto.** Al terminar, el programa compara el texto de entrada con el de salida y te avisa si hay alguna diferencia.
- Tu archivo original **no se toca**: siempre se crea una copia nueva.

> No necesitas saber programar. Sigue los pasos en orden.

---

## Paso 1. Instalar Python (solo la primera vez)

1. Entra a <https://www.python.org/downloads/> y pulsa el botón amarillo **Download Python**.
2. Abre el instalador descargado.
3. **Muy importante:** en la primera pantalla, marca la casilla **"Add python.exe to PATH"** (abajo) y luego pulsa **Install Now**.
4. Espera a que termine y pulsa **Close**.

**Comprobar que quedó bien:** pulsa la tecla Windows, escribe `cmd`, abre **Símbolo del sistema** y escribe:

```
python --version
```

Debe aparecer algo como `Python 3.12.x`. Si dice que `python` no se reconoce, prueba con `py --version`. Si tampoco, vuelve a instalar y revisa la casilla del punto 3.

## Paso 2. Descargar el programa

1. En esta página de GitHub pulsa el botón verde **Code** y luego **Download ZIP**.
2. Ve a tu carpeta de Descargas, haz **clic derecho** sobre el ZIP y elige **Extraer todo… → Extraer**.
3. Mueve la carpeta resultante (por ejemplo al Escritorio). Debe contener, entre otros, `reformateador.py` y `ejecutar.bat`.

> No ejecutes el programa desde dentro del ZIP: primero hay que extraerlo.

## Paso 3. Abrir el programa

1. Dentro de la carpeta, haz **doble clic en `ejecutar.bat`**.
2. Se abre una ventana negra. **La primera vez** instala lo necesario (1 a 2 minutos, necesita internet). No la cierres.
3. Después aparece la ventana del programa.

Si Windows muestra *"Windows protegió su PC"*: pulsa **Más información → Ejecutar de todas formas**. (Es un archivo de texto con 4 líneas; puedes abrirlo con el Bloc de notas para verlo.)

Las siguientes veces, el mismo doble clic abre el programa directamente.

## Paso 4. Reformatear tu documento

1. **Arrastra** tu archivo `.docx` o `.pdf` al recuadro azul (o haz clic en el recuadro para buscarlo).
2. En **Norma**, elige la que te piden (ver tabla más abajo).
3. Opciones (todas son opcionales):

   | Opción | Cuándo usarla |
   |---|---|
   | **Encabezado corrido** | Si tu norma/institución pide un texto corto arriba en cada página (p. ej. APA profesional). Déjalo vacío si no te lo piden. |
   | **Apellido (solo MLA)** | Para el encabezado MLA: aparecerá "Apellido 1", "Apellido 2"… |
   | **Ordenar referencias alfabéticamente** | APA, MLA, Chicago e ICONTEC piden la lista en orden alfabético. |
   | **Numerar referencias si faltan** | Solo IEEE y Vancouver (referencias `[1]`, `[2]`…). |
   | **Detectar títulos automáticamente** | Déjalo marcado si tu Word no usa los estilos "Título 1, 2, 3". |
   | **Generar también un PDF** | Solo funciona si tienes LibreOffice instalado. |

4. Pulsa **Reformatear documento**.
5. Al terminar, el programa te avisa y te ofrece abrir la carpeta. El resultado queda **junto a tu archivo original** con el nombre de la norma, por ejemplo:
   `MiTrabajo.docx` → `MiTrabajo_apa7.docx`

## Paso 5. Revisar el resultado (no te lo saltes)

Abre el archivo nuevo en Word y comprueba:

- [ ] **Títulos:** que cada uno tenga el nivel correcto (sobre todo si el programa avisó que los detectó por su aspecto).
- [ ] **Tabla de contenido:** si tu documento tiene una, Word preguntará si actualiza campos: acepta. Si no pregunta, clic derecho sobre el índice → *Actualizar campo*.
- [ ] **Citas dentro del texto:** el programa **no** las convierte. Si pasas de APA `(García, 2020)` a IEEE `[1]`, debes reescribirlas tú. El programa te avisa si detecta que no coinciden con la norma.
- [ ] **Tablas y figuras:** revisa que queden bien ubicadas y con su leyenda. Los bordes de las tablas no se modifican.
- [ ] **Portada:** el programa no la crea; si tu norma pide una, hay que hacerla aparte.

---

## Qué norma elegir y qué aplica

| Norma | Letra | Interlineado | Márgenes | Papel |
|---|---|---|---|---|
| APA 7 | Times New Roman 12 | Doble | 2,54 cm | Carta |
| IEEE | Times New Roman 10 | Sencillo | Arriba 1,9 · Abajo 2,54 · Lados 1,59 cm | Carta |
| MLA 9 | Times New Roman 12 | Doble | 2,54 cm | Carta |
| Chicago 17 | Times New Roman 12 | Doble | 2,54 cm | Carta |
| Vancouver | Arial 12 | Doble | 2,5 cm | A4 |
| ICONTEC (NTC 1486) | Arial 12 | 1,5 | Sup. 3 · Inf. 3 · Izq. 3 · Der. 2 cm | Carta |

> **Importante:** estos valores son una aproximación de los manuales. Muchas universidades tienen su propia guía (especialmente con ICONTEC). **Confirma con la guía de tu institución** antes de entregar.
> Si algo difiere, abre `reformateador.py` con el Bloc de notas, busca `NORMAS` y cambia el valor (por ejemplo el margen o la fuente).

---

## Problemas frecuentes

| Qué ves | Qué hacer |
|---|---|
| `'python' no se reconoce como un comando` | Python no se instaló con PATH. Reinstálalo y marca **Add python.exe to PATH** (Paso 1). |
| La ventana negra se cierra sola | Abre `cmd`, entra a la carpeta (`cd` + ruta) y ejecuta `python reformateador.py` para ver el mensaje de error. |
| Arrastrar y soltar no funciona | Haz clic en el recuadro para buscar el archivo. Para activar arrastrar: `pip install tkinterdnd2`. |
| *"ciérrelo si lo tiene abierto en Word"* | Cierra el documento de salida en Word y vuelve a intentarlo. |
| *"El PDF no contiene texto seleccionable"* | Tu PDF es un escaneo. Conviértelo con OCR primero (por ejemplo con `ocrmypdf` o Adobe Acrobat). |
| *"No se pudo abrir el .docx"* | El archivo está dañado o protegido con contraseña. Ábrelo en Word y guárdalo como nuevo `.docx`. |
| *"No se encontró el archivo"* (línea de comandos) | La ruta o el nombre están mal. Usa comillas si tiene espacios y la ruta completa. |
| Los títulos salen mal desde un PDF | Es normal: un PDF no guarda la estructura. Usa el Word original si lo tienes. |

---

## Uso desde la línea de comandos (opcional)

Abre `cmd` dentro de la carpeta del programa y ejecuta:

```
python reformateador.py "C:\ruta\MiTrabajo.docx" -n apa7
python reformateador.py "C:\ruta\articulo.docx" -n ieee --columnas 2 --numerar-refs
python reformateador.py "C:\ruta\tesis.docx" -n icontec --ordenar-refs --pdf
python reformateador.py --listar-normas
```

Claves de norma: `apa7`, `ieee`, `mla9`, `chicago17`, `vancouver`, `icontec`.
Otras opciones: `--papel carta|a4|original`, `--encabezado "TÍTULO CORTO"`, `--apellido Perez`, `--sin-detectar-titulos`, `-o salida.docx`.

**Mac / Linux:** en la terminal, dentro de la carpeta:

```
pip3 install -r requirements.txt
python3 reformateador.py
```
(En Linux, si falta la ventana: `sudo apt install python3-tk`.)

---

## Límites

- **No convierte las citas dentro del texto** entre estilos. Solo te avisa.
- **Desde PDF solo se recupera el texto** (sin imágenes ni tablas; los PDF de varias columnas pueden salir desordenados). Con el Word original el resultado es mucho mejor.
- No cambia los bordes de las tablas ni crea portada.
- Los títulos sin estilo de Word se detectan por su aspecto y pueden fallar: revísalos.

## Licencia

MIT. Ver el archivo `LICENSE`.
