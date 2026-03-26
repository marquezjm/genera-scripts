# generar-script-xlsx

> Convierte archivos Excel en scripts SQL listos para ejecutar en SQL Server.

![Node.js Version](https://img.shields.io/badge/node-%3E%3D18-brightgreen)
![npm Version](https://img.shields.io/badge/npm-latest-blue)

---

## Que es esto?

**generar-script-xlsx** es una herramienta de linea de comandos que lee archivos Excel y genera scripts SQL (para SQL Server) con transacciones `INSERT` y `UPDATE`.

Imagina que tienes un archivo Excel como este:

| ACCION | NUMERO_INSTITUCION | TRANSACCION_EXTERNA | MONTO |
|--------|--------------------|---------------------|-------|
| INSERT | 001                | 12345               | 100.50|
| UPDATE | 001                | 12345               | 200.00|

La herramienta detectara automaticamente:
- **INSERT**: Generara un script para insertar todos los valores de la fila.
- **UPDATE**: Generara un script para actualizar solo las celdas que tengan estilo (negrita o color de fondo), usando las columnas clave en el `WHERE`.

El resultado sera algo como:

```sql
BEGIN TRAN TRAN_NombreHoja

BEGIN TRY
    INSERT INTO NombreHoja (NUMERO_INSTITUCION, TRANSACCION_EXTERNA, MONTO) VALUES ('001', '12345', '100.50');
    UPDATE NombreHoja SET MONTO = '200.00' WHERE NUMERO_INSTITUCION = '001' AND TRANSACCION_EXTERNA = '12345';

COMMIT TRAN TRAN_NombreHoja;
PRINT 'PROCESO EJECUTADO CORRECTAMENTE';
END TRY
BEGIN CATCH
    SELECT 'LINEA ERROR - ' + CAST(ERROR_LINE() AS VARCHAR(5)) + ': ' + ERROR_MESSAGE();
    ROLLBACK TRAN TRAN_NombreHoja;
    PRINT 'OCURRIO UN ERROR EN EL PROCESO';
END CATCH
```

---

## Caracteristicas principales

- **Lectura de multiples hojas**: Procesa todas las hojas del archivo Excel.
- **Deteccion automatica de acciones**: Identifica columnas con valores `INSERT` o `UPDATE`.
- **Estilos como indicadores**: Solo actualiza celdas que tengan **negrita** o **color de fondo**.
- **Llaves configurables**: Define que columnas se usaran para el `WHERE` en los `UPDATE`.
- **Modo interactivo**: Si detecta `UPDATE`s, te pregunta interactivamente que columnas usar como llaves.
- **Soporte por hoja**: Cada hoja puede tener su propia configuracion de llaves.
- **Manejo de valores especiales**: Detecta `NULL`, `GETDATE()` y fechas automaticamente.
- **Esquema de transacciones**: Incluye `BEGIN TRAN`, `TRY/CATCH` para ejecucion segura.

---

## Requisitos previos

### 1. Node.js

Necesitas tener **Node.js** instalado en tu computadora.

**Como verificar si lo tienes:**
```bash
node --version
```

**Si no lo tienes, instalalo:**

- **Windows**: Descarga el instalador desde [nodejs.org](https://nodejs.org/)
- **Mac**: `brew install node`
- **Linux**: `sudo apt install nodejs npm` (Ubuntu/Debian)

Se recomienda usar **Node.js 18 o superior**.

### 2. npm

Normalmente viene incluido con Node.js. Verifica con:

```bash
npm --version
```

---

## Instalacion

Tienes varias opciones:

### Opcion A: Instalar globalmente con npm (Recomendado)

Esta opcion instala la herramienta para que puedas usarla desde cualquier carpeta.

```bash
npm install -g generar-script-xlsx
```

### Opcion B: Vincular el proyecto localmente (para desarrollo)

Si clonaste o descargaste el codigo fuente:

```bash
# Entra a la carpeta del proyecto
cd ruta/al/proyecto

# Instala las dependencias
npm install

# Vincula como comando global
npm link
```

### Opcion C: Usar con npx (sin instalar)

Tambien puedes ejecutarlo sin instalar:

```bash
npx generar-script-xlsx "archivo.xlsx"
```

---

## Uso basico

### Sintaxis general

```bash
generar-script-xlsx <archivo_excel> [opciones]
```

### Ejemplo mas simple

```bash
# Supongamos que tienes un archivo llamado "datos.xlsx" en tu escritorio
generar-script-xlsx "C:\Users\TuUsuario\Desktop\datos.xlsx"
```

La herramienta:
1. Leera el archivo `datos.xlsx`
2. Procesara cada hoja
3. Generara archivos `.sql` en la misma carpeta donde esta el archivo Excel
4. Los archivos se llamaran igual que las hojas (ej: `Hoja1.sql`, `Hoja2.sql`)

---

## Opciones disponibles

### `-k, --key-columns <config>`

Define las columnas que se usaran en la clausula `WHERE` para los `UPDATE`.

**Formatos:**

#### 1. Llaves globales (aplica a todas las hojas)

```bash
generar-script-xlsx archivo.xlsx -k "COLUMNA1,COLUMNA2"
```

Ejemplo:
```bash
generar-script-xlsx archivo.xlsx -k "ID_CLIENTE,FECHA"
```

#### 2. Llaves por hoja (diferentes llaves para cada hoja)

Usa el formato `NombreHoja:Columnas`:

```bash
generar-script-xlsx archivo.xlsx -k "Hoja1:COL1,COL2;Hoja2:COL3"
```

Ejemplo completo:
```bash
generar-script-xlsx archivo.xlsx -k "Clientes:ID_CLIENTE;Ventas:ID_VENTA,FECHA;Transacciones:ID"
```

| Hoja        | Llaves usadas en WHERE              |
|-------------|-------------------------------------|
| Clientes    | `ID_CLIENTE`                        |
| Ventas      | `ID_VENTA, FECHA`                   |
| Transacciones | `ID`                             |

#### 3. Llaves por hoja + valor por defecto

Si una hoja no esta definida, usara las llaves por defecto:

```bash
generar-script-xlsx archivo.xlsx -k "HojaEspecial:ID_UNICO;COL_DEFAULT1,COL_DEFAULT2"
```

### `-V, --version`

Muestra la version de la herramienta.

```bash
generar-script-xlsx --version
```

### `-h, --help`

Muestra la ayuda con todas las opciones disponibles.

```bash
generar-script-xlsx --help
```

---

## Modo interactivo

Una de las caracteristicas mas utiles es el **modo interactivo**.

### Cuando se activa?

Cuando el script detecta filas con `UPDATE` y **no** le has proporcionado las llaves mediante la opcion `-k`.

### Como funciona?

1. La herramienta detecta que hay `UPDATE`s en una hoja.
2. Te muestra la lista de columnas disponibles numeradas.
3. Puedes seleccionar las columnas usando el **checkbox interactivo**.
4. Presiona `Espacio` para marcar/desmarcar columnas.
5. Presiona `Enter` para confirmar tu seleccion.

### Ejemplo en terminal:

```
Processing sheet: Transacciones
  -> Stats: 150 rows (approx).
  -> Action Column: 'ACCION' (Index 0)

[!] UPDATE operations detected in 'Transacciones'
? Select the Key Columns for the WHERE clause: (Press <space> to select, <Enter> to confirm)
  - NUMERO_INSTITUCION
  X TRANSACCION_EXTERNA
  - MONTO
  - FECHA
  - ESTADO
```

En este ejemplo:
- `NUMERO_INSTITUCION` esta desmarcado
- `TRANSACCION_EXTERNA` esta seleccionado (marcado con `X`)

Presiona `Enter` para confirmar.

---

## Ejemplos practicos

### Ejemplo 1: Uso basico (archivo en el escritorio)

```bash
generar-script-xlsx "C:\Users\MiUsuario\Desktop\actualizacion.xlsx"
```

**Resultado**: Se generara `actualizacion.sql` en el escritorio.

---

### Ejemplo 2: Definir llaves globales

```bash
generar-script-xlsx "datos.xlsx" -k "ID_CLIENTE,FECHA_OPERACION"
```

Todas las hojas usaran `ID_CLIENTE` y `FECHA_OPERACION` en sus `WHERE`.

---

### Ejemplo 3: Llaves diferentes por hoja

```bash
generar-script-xlsx "reportes.xlsx" -k "Clientes:ID_CLIENTE;Productos:CODIGO_PRODUCTO;Ventas:ID_VENTA,FECHA"
```

---

### Ejemplo 4: Ruta con espacios

```bash
generar-script-xlsx "C:\Users\Mi Usuario\Documentos\Mi Carpeta\datos.xlsx"
```

O usando comillas simples:

```bash
generar-script-xlsx 'C:\Users\Mi Usuario\Documentos\Mi Carpeta\datos.xlsx'
```

---

## Formato del archivo Excel

### Estructura esperada

| Columna A (ACCION) | Columna B | Columna C | ... |
|--------------------|-----------|-----------|-----|
| INSERT o UPDATE    | valor1    | valor2    | ... |

### Reglas importantes

1. **Columna de accion**: La primera columna (A) debe contener `INSERT` o `UPDATE` (no distingue mayusculas/minusculas).

2. **Encabezados**: La primera fila debe contener los nombres de las columnas de tu tabla en SQL.

3. **Valores especiales**:
   - Escribe `NULL` para valores nulos
   - Escribe `GETDATE()` para fechas actuales del servidor

### Ejemplo de Excel:

| ACCION | ID_PRODUCTO | NOMBRE      | PRECIO | FECHA_CREACION |
|--------|------------|-------------|--------|----------------|
| INSERT | 1001       | Producto A  | 25.50  | GETDATE()      |
| UPDATE | 1001       | Producto A  | 29.99  |                |
| INSERT | 1002       | Producto B  | 15.00  | NULL          

> **Nota**: En el UPDATE, solo `PRECIO` se actualizara si la celda tiene estilo (negrita o color). `ID_PRODUCTO` y `NOMBRE` se usan en el WHERE pero no se actualizan.

---

## Interpretacion de estilos en Excel

La herramienta detecta cambios basandose en el **estilo** de la celda:

| Estilo detectado | Comportamiento                              |
|-----------------|---------------------------------------------|
| **Negrita**     | La celda se incluira en el `SET` del `UPDATE` |
| **Color fondo** | La celda se incluira en el `SET` del `UPDATE` |

### Como marcar celdas para actualizar?

1. Abre tu archivo Excel
2. Identifica las filas con `UPDATE` en la columna de accion
3. Aplica **negrita** o **color de fondo** a las celdas que quieres que se actualicen
4. Guarda el archivo
5. Ejecuta la herramienta

---

## Solucion de problemas

### Error: "Command not found"

Si ves este error despues de instalar:

```bash
generar-script-xlsx : El termino 'generar-script-xlsx' no se reconoce...
```

**Solucion:**
1. Verifica que la instalacion termino correctamente
2. Cierra y abre una nueva terminal
3. O prueba con la ruta completa de npm:
   ```bash
   npx generar-script-xlsx archivo.xlsx
   ```

---

### Error: "Input file not found"

```bash
Error: Input file not found at C:\ruta\archivo.xlsx
```

**Solucion:**
- Verifica que la ruta este completa y correcta
- Verifica que el archivo realmente exista
- Usa comillas si la ruta tiene espacios

---

### No se genera ningun SQL (archivo vacio)

Posibles causas:

1. **La columna de accion no se detecto correctamente**
   - Verifica que la primera columna se llame `ACCION`, `ACTION`, `TIPO`, `OPERACION` o `MOVIMIENTO`
   - O que el texto `INSERT`/`UPDATE` este en la primera columna

2. **Los valores no coinciden**
   - Verifica que el texto sea exactamente `INSERT` o `UPDATE`
   - No debe haber espacios extra (` INSERT` o `UPDATE `)

3. **Hoja vacia**
   - Verifica que la hoja tenga filas con datos

---

### Error de permisos al escribir

```bash
Error: EACCES: permission denied, open 'archivo.sql'
```

**Solucion:**
- Verifica que tengas permisos de escritura en la carpeta destino
- Intenta ejecutar la terminal como Administrador
- O guarda el archivo Excel en una carpeta diferente

---

### Problemas con fechas

Si las fechas se generan con formato incorrecto:

La herramienta convierte fechas de Excel al formato SQL: `'2026-03-26 16:30:00'`

Si necesitas un formato especifico, edita manualmente el archivo `.sql` generado despues.

---

## Desinstalar

Si deseas quitar la herramienta de tu sistema:

```bash
npm uninstall -g generar-script-xlsx
```

---

## Tecnologias utilizadas

- [Node.js](https://nodejs.org/) - Entorno de ejecucion
- [ExcelJS](https://github.com/exceljs/exceljs) - Lectura de archivos Excel
- [Commander.js](https://github.com/tj/commander.js) - CLI framework
- [Inquirer.js](https://github.com/SBoudrias/Inquirer.js) - Preguntas interactivas
- [Chalk](https://github.com/chalk/chalk) - Colores en terminal

---

## Licencia

MIT License - Libre para usar, modificar y distribuir.

---

## Contribuir

Las contribuciones son bienvenidas! Si encuentras un bug o tienes una mejora:

1. Haz un fork del repositorio
2. Crea una rama para tu feature: `git checkout -b mi-nueva-funcion`
3. Haz commit de tus cambios: `git commit -m 'Agregar nueva funcion'`
4. Haz push a la rama: `git push origin mi-nueva-funcion`
5. Abre un Pull Request

---

## Contacto y soporte

Si tienes dudas o problemas:

1. Revisa la seccion de [Solucion de problemas](#solucion-de-problemas)
2. Busca en los [issues del repositorio](https://github.com/marquezjm/genera-scripts/issues)
3. Abre un nuevo issue si no encuentras la solucion

---

Gracias por usar **generar-script-xlsx**!
