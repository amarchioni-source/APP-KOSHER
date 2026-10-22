# Reporte de producción · Kosher USA

Aplicación web que genera el *Production report – DEVESA (USA)* a partir de la bajada de stock.
Se sube el Excel de la bajada y se obtiene un archivo por cada fecha de producción, con el formato exacto del modelo
(`modelo/modelo.xlsx`). No depende de Claude ni de ningún servicio externo, y no guarda los archivos que se suben.

## Uso
1. Abrir la aplicación y seleccionar uno o varios archivos de bajada (.xlsx).
2. Mantener el destino `KOSHER USA` (puede cambiarse) y presionar **Generar reportes**.
3. Revisar el resumen (cajas, kg, % Spiced, productos sin barcode) y descargar cada archivo, o todos en un .zip.

## Cómo opera
- Filtra los registros del destino indicado y los agrupa por código y por fecha (columna `Fecha`).
- Suma cajas, peso bruto y peso neto por código; la fecha de faena es la máxima. Ordena por barcode.
- Reconoce las dos variantes de columnas (`Unidad/Peso/PesoBruto` y `cajas/kg/kg bruto`).
- Si los pesos vienen multiplicados por 100.000, los divide automáticamente y lo informa.
- Si hay más productos que filas del modelo, agrega filas y ajusta los totales y las fórmulas.
- Los productos sin barcode se incluyen igual, al final de la lista, y se avisan en pantalla.

## Barcodes
Los barcodes están en `barcodes.csv` (columnas `codigo,barcode,descripcion`). Para agregar uno nuevo, editar ese
archivo en GitHub (lápiz → agregar línea → *Commit*); Render vuelve a desplegar solo. Para una prueba puntual, se pueden
pegar en el recuadro «Barcodes adicionales» de la pantalla (no se guardan).

Si cambia el modelo, reemplazar `modelo/modelo.xlsx` conservando la misma estructura (hoja visible en segundo lugar,
datos desde la fila 3, totales en la fila 37).

## Publicación en GitHub y Render
1. **GitHub:** crear un repositorio nuevo (privado recomendado) y subir todo el contenido de esta carpeta.
2. **Render:** *New +* → *Web Service* → conectar el repositorio. Render detecta `render.yaml`; si se completa a mano:
   - Build command: `pip install -r requirements.txt`
   - Start command: `gunicorn app:app --workers 1 --threads 4 --timeout 120`
3. **Contraseña (recomendado):** en *Environment* crear la variable `APP_PASSWORD` con la clave elegida. El navegador la
   pedirá al entrar (el usuario puede ser cualquiera). Sin esa variable, la dirección queda abierta a cualquiera que la conozca.
4. Cada cambio que se confirme en GitHub se despliega automáticamente.

Notas: en el plan gratuito de Render el servicio se suspende tras un rato sin uso y la primera carga puede demorar
alrededor de un minuto. Debe usarse un solo worker (ya configurado), porque las descargas se sirven desde memoria.

## Prueba local
```
pip install -r requirements.txt
python app.py        # http://localhost:5000
```
