# Reporte de producción · Kosher (USA e Israel)

Aplicación web que genera el *Production report – DEVESA* a partir de la bajada de stock. Se elige el informe
(**USA**, **Israel** o **Ambos**), se sube el Excel y se obtiene un archivo por cada fecha de producción, con el formato
exacto del modelo. No depende de Claude ni de ningún servicio externo, y no guarda los archivos que se suben.

## Uso
1. Elegir **USA**, **Israel** o **Ambos** y seleccionar uno o varios archivos de bajada (.xlsx).
2. Presionar **Generar reportes**.
3. Revisar cada reporte: tilde verde si está completo, alarma roja si quedan datos pendientes. Descargar cada archivo o todos en un .zip.

Con **Ambos**, el mismo archivo sirve para los dos informes siempre que traiga las columnas de cada uno. Si un archivo no
sirve para un informe, la pantalla lo avisa y sigue con el resto.

## Informe USA
- Filtra el destino indicado (por defecto `KOSHER USA`), agrupa por código y fecha, y suma cajas, peso bruto y peso neto.
- Reconoce las variantes de columnas (`Unidad/Peso/PesoBruto` y `cajas/kg/kg bruto`) y corrige pesos multiplicados por 100.000.
- Barcodes en `barcodes.csv`. Si falta uno, el producto se incluye al final con alarma roja y los códigos a completar.
- Modelo: `modelo/modelo.xlsx`. La estructura se detecta sola (columnas por encabezado, totales por etiqueta); si hay más
  productos que filas, agrega las necesarias y ajusta las fórmulas.

## Informe Israel
- Una fila por cartón, con código de caja (CODBAR CAJA) y fecha de vencimiento (BEST BEFORE DATE). Un archivo por día de producción.
- Entran las descripciones que contienen NETO, OZUL, T BONE VEAL o GRASA SALADA. Las demás se listan en el control.
- Estados: se excluye S; se incluyen E, PA, D y 104 (los 104 sin CodBarS quedan con la celda de código vacía).
- Orden: primero los códigos FD y luego los CD; dentro de cada grupo, por código, fecha de faena y CodBarS.
- Índice NETO y barcode salen de `indices_israel.csv`. Si un código no está, las celdas quedan vacías y se marca con alarma roja
  (no se inventan valores). Los códigos marcados `inferido` usan el índice y el barcode de su par GF AA; se avisan en pantalla.
- La fecha de vencimiento se toma del stock tal cual. Si faltan, se avisa con alarma roja y no se estima ninguna fecha.
- Corrige unidades o pesos multiplicados por 100.000, y cuenta una sola vez los cartones repetidos entre archivos (mismo CodBarS).
  Si dos archivos tienen contenido idéntico, usa uno solo y lo avisa.
- Además de los reportes, genera **Control producción Israel.xlsx**: resumen por día (cartones, unidades, bruto, neto) y las listas de
  pendientes (códigos sin índice, cartones sin vencimiento, inferidos, descripciones excluidas, cartones en estado S, días hábiles sin producción).
- Modelo: `modelo/modelo_israel.xlsx`.

## Tablas editables (en GitHub)
- `barcodes.csv`: `codigo,barcode,descripcion` (informe USA).
- `indices_israel.csv`: `codigo,indice,barcode,descripcion,inferido` (informe Israel; `SI` en la última columna si es inferido).

Para agregar un código: abrir el archivo en GitHub, presionar el lápiz, agregar la línea y *Commit*. Render vuelve a desplegar solo.
Cuando falta un dato, la pantalla ofrece las líneas listas para copiar. En *Opciones avanzadas* se pueden pegar datos para una
corrida puntual (no se guardan).

## Cambios en los modelos
La aplicación lee la estructura de los modelos en cada uso. Se pueden cambiar formatos, anchos o la posición de los totales y
alcanza con reemplazar el archivo en GitHub. Debe conservarse el texto de los encabezados y de las etiquetas de totales (USA:
BARCODE, MercDescripcion, Codigo, TOTAL, Spiced, No Spiced, % Spiced; Israel: PRODUCER en A2 y la fila TOTAL).

## Publicación en GitHub y Render
1. **GitHub:** crear un repositorio (privado recomendado) y subir todos los archivos. Las carpetas son `templates/` (con `index.html`)
   y `modelo/` (con `modelo.xlsx` y `modelo_israel.xlsx`). La app también funciona si los archivos están todos sueltos en la raíz.
2. **Render:** *New +* → *Blueprint* → conectar el repositorio. Render lee `render.yaml` y pide solo el valor de `APP_PASSWORD`.
3. **Contraseña (recomendado):** variable `APP_PASSWORD`. El navegador la pedirá al entrar (el usuario puede ser cualquiera).
4. Cada cambio confirmado en GitHub se despliega automáticamente.

Notas: en el plan gratuito de Render el servicio se suspende tras un rato sin uso y la primera carga puede demorar cerca de un minuto.
Debe usarse un solo worker (ya configurado), porque las descargas se sirven desde memoria. Un día de Israel con unos 2.000 cartones
se genera en pocos segundos.

## Prueba local
```
pip install -r requirements.txt
python app.py        # http://localhost:5000
```
