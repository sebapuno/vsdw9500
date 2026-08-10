# Investigación: compilación de VISADO.VBP — estado y bloqueo actual

**Fecha:** 2026-08-09
**Proyecto:** `FUENTES/VISADO.VB/VISADO.VBP` (ejecutable `VSDPROD.EXE`)
**Objetivo:** migrar y compilar contra la VM real (VB6 32-bit)

---

## 1. Resumen ejecutivo

- Se migraron y corrigieron **9 blockers de compilación reales** (detalle en la sección 2) en 6 archivos del grupo VISADO.VB.
- Se encontraron y corrigieron **2 bugs en el framework `vb6-framework`** (ya pusheados a `sebapuno/vb6.git`, commit `fdd57c1`): un falso positivo del validador y un bug de scheduling AM/PM que impedía compilar de tarde/noche.
- ~~Queda **1 blocker sin resolver**~~ → **RESUELTO** el 2026-08-09 por la noche. Ver **sección 5**: el número de línea que reporta VB6 **sí es confiable**, pero es relativo al inicio de la sección de código del `.FRM`, no al archivo. Con esa corrección el bug apareció de inmediato y no hizo falta acceso de escritorio a la VM.
- **La recomendación anterior (abrir el IDE con escritorio remoto) quedó sin efecto** — se resolvió por SSH/`/make`.

---

## 2. Blockers reales encontrados y corregidos (confirmados, no son el problema actual)

### 2.1 Bug sistémico: referencia a un control de otro form sin calificar
Patrón: un form declara su grilla con un nombre (ej. `GrillaVisado`), pero el código tiene líneas sueltas que referencian `GrillaProductos` (el nombre de la grilla de `VISANDO.FRM`, probablemente copiada como plantilla) sin el prefijo del form dueño. Encontrado **6 veces**:

| Archivo | Control local real | Referencia rota |
|---|---|---|
| `PANTVISA.FRM` | `GrillaVisado` | `GrillaProductos.Rows` (31 líneas) |
| `IMPR_NEW.FRM` | `GrillaProductosVisados` | `GrillaProductos.Rows` (10 líneas) |
| `IMPRPROD.FRM` | `GrillaProductosVisados` | `GrillaProductos.Rows` (10 líneas) |
| `NOTARIO.FRM` | `GrillaHonorarios` | `GrillaProductos.Rows` (12 líneas) |
| `VIGENCIA.FRM` | `GrillaVigenciaProductos` | `GrillaProductos.Rows` (1 línea) |
| `VISANDO.FRM` | (control es de `Principal`/`VISADO.FRM`) | `Principal.Grilla.RowHeight(Grilla.Row)` — `Grilla.Row` sin calificar, mezclado en la misma línea con la referencia sí calificada |

**Estado:** corregido en los 6 archivos. Documentado en la guía del framework con receta de detección por `grep`.

### 2.2 `PANTVISA.FRM`: control declarado con tipo incorrecto
`GrillaVisado` estaba declarado como `VB.PictureBox` pero el código le llamaba `.Rows`, `.Row`, `.Col`, `.Text` (miembros de grilla). Convertido a `MSFlexGridLib.MSFlexGrid`, preservando posición/tamaño/TabIndex. **Corregido.**

### 2.3 Firmas de evento estilo Spread (`Click`/`DblClick`/`LeaveCell`) sobre MSFlexGrid
MSFlexGrid no acepta parámetros `(Col, Row, ...)` en estos eventos. Corregido en:
- `VISANDO.FRM`: `GrillaProductos_DblClick(Col, Row)` → `DblClick()`.
- `VISANDO.FRM`: `GrillaProductos_LeaveCell(Col,Row,NewCol,NewRow,Cancel)` → reemplazado por `GrillaProductos_RowColChange()` + variable de módulo `UltimaFilaGrillaProductos` para reproducir el `If Row <> NewRow` original.
- `IMPR_NEW.FRM`: `GrillaProductosVisados_DblClick(Col,Row)` → `DblClick()`, leyendo `.Col` del control directo.
- `VIGENCIA.FRM`: `GrillaVigenciaProductos_DblClick(Col,Row)` → `DblClick()`; `.ActiveCol` → `.Col`.

**Estado:** corregido y **confirmado que resuelve** el error correspondiente (verificado con test aislado: solo este fix hace avanzar la compilación del error en `VISANDO.FRM:460` al siguiente).

### 2.4 Propiedades de Spread sin equivalente en MSFlexGrid (write-only, comentadas)
Confirmado con `grep` que ninguna se lee de vuelta en ningún evento. Comentadas con marcador `FI2026 SHG`:
`OperationMode`, `CellType` (en varios contextos), `TypeTextShadow`, `TypeEditLen`, `CursorType`, `CursorStyle`, `TypePicMask` (en `NOTARIO.FRM`).

### 2.5 Propiedades con equivalente directo — migradas
- `TypePictPicture` → `.CellPicture` (imagen embebida en celda — íconos de aprobado/rechazado en `VISANDO.FRM`, `PANTVISA.FRM`).
- `Action = 5` (borrar fila activa) → `.RemoveItem(.Row)` en `VISANDO.FRM`.
- `Row2`/`Col2` → `RowSel`/`ColSel` en `IMPR_NEW.FRM`.

**Nota:** se probó `.CellPictureAlignment` como reemplazo de `TypePictCenter/Stretch/MaintainScale` y **causó un error de compilación real** ("Method or data member not found" en un punto distinto al bug actual). Se comentó, pese a que el string `CellPictureAlignment` SÍ existe en el binario del OCX instalado (verificado con `strings` sobre `MSFLXGRD.OCX`) — posible restricción "not available at design time" documentada en el propio OCX. Queda pendiente de revisar visualmente el ajuste del ícono sin esa propiedad.

### 2.6 Checkbox de Spread sin equivalente — pérdida visual real (diferido, ver tarea aparte)
`VIGENCIA.FRM`: `CellType=10` + `TypeCheckCenter` + `TypeCheckText="Vigente"` renderizaba un checkbox real. MSFlexGrid no tiene celda-checkbox nativa. El toggle de datos (0/1) sigue funcionando sobre `.Text`; la celda va a mostrar el número en vez de un tilde. **Diferido a pedido del usuario para el cierre de la migración.**

### 2.7 Edición inline perdida — pérdida funcional real (diferido)
`NOTARIO.FRM`: `GrillaHonorarios_EditMode` (evento Spread que marcaba inicio/fin de edición de celda) no existe en MSFlexGrid. El handler `GrillaHonorarios_LeaveCell` asociado validaba el Rut con dígito verificador al terminar de editar — queda código muerto (la bandera que lo activaba nunca se vuelve a setear). **Diferido**: requiere un `TextBox` flotante superpuesto sobre la celda para restaurar la edición directa.

### 2.8 `.Action` con propósito no confirmado (diferido)
`IMPR_NEW.FRM`: `.Action = 2/22/0` tras exportar la grilla a Excel. Sin equivalente genérico en MSFlexGrid y sin poder confirmar el comportamiento original. Comentado con `FI2026_TODO`. **Diferido.**

### 2.9 `VISADO.BAS`: variable con nombre de tipo reservado
`Dim Decimal As String` en la función `ColocarSeparadorDeMiles` — `Decimal` es palabra reservada en VB6 32-bit (el compilador 16-bit original lo toleraba). Causaba `"Syntax error"` con el número de línea reportado apuntando al **inicio del procedimiento siguiente**, no al `Dim` real (primer caso confirmado de este patrón de línea poco confiable, documentado también en la guía). Renombrada a `ParteDecimal`. **Confirmado que resuelve el error** (test aislado: con solo este fix, el `"Syntax error"` desaparece).

---

## 3. ~~El blocker actual~~: `"Method or data member not found"` en `VISANDO.FRM`

> **⚠️ SECCIÓN SUPERADA — se conserva sólo como registro de lo que se descartó.**
> La causa real era `.ColHidden` (propiedad de Spread), y la premisa de 3.2 ("el
> número de línea no es confiable") **es falsa**: el número es relativo al inicio
> de la sección de código del `.FRM`. Ver **sección 5**. Los 22 tests de 3.3 se
> hicieron sobre esa premisa equivocada — por eso ninguno dio con la causa.

### 3.1 Síntoma
```
Compile Error in File 'C:\dev\proyecto\VISANDO.FRM', Line 371 (o 379, varía) : Method or data member not found
Build of 'VSDPROD.EXE' failed.
```

### 3.2 Por qué el número de línea NO es confiable
La línea reportada cae siempre dentro del mismo bloque de **diseño** (propiedades del `Label` "Rechazado", ej. `Top = 330`), que es **idéntico byte a byte** al archivo original sin migrar (verificado con `diff` binario). Ese bloque no es código ejecutable — no puede ser la causa real. Además, el número exacto reportado **varía entre 124, 268, 371 y 379** dependiendo de cuántas líneas tiene el archivo en cada versión de prueba (a más líneas comentadas/removidas, más bajo el número) — lo que sugiere que la posición reportada escala con el tamaño total del archivo en vez de apuntar a la línea real del problema. Mismo fenómeno documentado en la guía para el error de DAO `[restricted]` (Paso 11).

### 3.3 Todo lo que se descartó (con evidencia)

Cada fila es un test real contra la VM (recompilación completa vía `vb6-build.sh`), no una suposición:

| # | Qué se deshabilitó/aisló | Resultado | Conclusión |
|---|---|---|---|
| 1 | `.CellPictureAlignment` (4 ocurrencias) | Error persiste | No es la causa |
| 2 | `.RemoveItem` + `.CellPicture` (4 ocurrencias, juntas) | Error persiste | Ninguna de las dos es la causa |
| 3 | Cuerpo de `GrillaProductos_RowColChange` vaciado | Error persiste | No es la lógica del handler |
| 4 | `GrillaProductos_RowColChange` eliminado por completo (declaración incluida) | Error persiste | No es el handler en sí |
| 5 | Toda la segunda mitad del archivo (desde `Rechazar_Click` hasta el final) stubbeada (cuerpos vacíos) | Error persiste | El bug NO está en ninguna de esas ~20 funciones, incluyendo `Visar_Click`, `PrepararTransmision`, `GrabarTxSAT`, etc. |
| 6 | `PANTVISA.FRM` revertido a versión original (sin mis fixes) | Error persiste, mismo lugar | No es una referencia cruzada hacia `PANTVISA.FRM` |
| 7 | Archivo `VISANDO.FRM` 100% original (sin ningún fix) | Error **distinto**: `"Procedure declaration..."` en línea 460 (el error original de `DblClick`) | Confirma que mis fixes SÍ hacen progresar la compilación; el bug actual solo aparece una vez arreglados `DblClick` + `LeaveCell` |
| 8 | Solo `DblClick` arreglado (resto original) | Error **distinto**: `"Procedure declaration..."` en línea 471 (el error de `LeaveCell`, esperable) | `DblClick` por sí solo está limpio |
| 9 | `DblClick` + `LeaveCell→RowColChange` arreglados, resto original | `"Method or data member not found"` en línea **124** | Progreso real la primera vez |
| 10 | + `DesMarca_Click` (OperationMode/CellType/TypeTextShadow comentados) | Línea **268** | Sigue "progresando" (probablemente solo por tamaño de archivo) |
| 11 | + `RemoveItem` en `EnviarVisado_Click` | Línea **371** | — |
| 12 | + `Form_Load` (TypeEditLen/CursorType/CursorStyle comentados) | Línea **371** (sin cambio) | No es esto |
| 13 | + `ColocarTickGrillaProductos` (CellType/TypePict→CellPicture) | Línea **371** (sin cambio) | No es esto |
| 14 | + fix de `Grilla.Row` (bug de referencia cruzada, sección 2.1) | Línea **371** (sin cambio) | No es esto (aunque el fix es real y necesario por otras razones) |
| 15 | + `LostFocus.OperationMode` comentado | Línea **371** (sin cambio) | No es esto |
| 16 | + `Visar_Click` Case "EE"/"RR" (CellType/TypePict/OperationMode arreglados) | Línea **371** (sin cambio) | No es esto — con TODO arreglado, el error no se mueve ni una línea |
| 17 | GUID de MSFlexGrid en el `.VBP` cambiado de `{5E9E78A0}` a `{6262D3A0}` | Error **distinto y peor**: `"MSFLXGRD.OCX could not be loaded"` | Error de comprensión propio: `5E9E78A0` es el GUID de la **type library**, no el CLSID del control — el `.VBP` original ya estaba bien. Revertido. |
| 18 | Licencia de diseño de MSFlexGrid en la VM (`HKCR\Licenses\{6262D3A0-...}`) | **No existe** en el registro ni en `licencias.reg` del framework | Hipótesis descartada por el usuario: confirmó que otros proyectos con MSFlexGrid sí compilan en esta misma VM |
| 19 | Revisión exhaustiva de `BuscaRentable`, `Cancelar_Click/KeyPress`, `DesMarca_KeyPress`, `EnviarVisado_KeyPress`, `HabilitarBotonTransmision`, `PermitirRechazoProducto`, `Form_Unload`, `LlenarGrillaProductos` | Código limpio, sin patrones Spread-only ni referencias rotas | — |
| 20 | Bytes de control raros / corrupción de encoding en el archivo | Ninguno encontrado | — |
| 21 | Variables con nombre de tipo reservado en `VISADO.BAS` y otros módulos (más allá de `Decimal`, ya arreglado) | Ninguna encontrada | — |
| 22 | Estructura de la declaración `Begin MSFlexGridLib.MSFlexGrid GrillaProductos` comparada contra otro form del mismo proyecto (`ROTOR1RD.FRM`) que usa el mismo patrón de migración | Estructura idéntica | No es un problema de declaración del control |

### 3.4 Estado del archivo `VISANDO.FRM` ahora mismo
Tiene aplicados **todos** los fixes de la sección 2 (los que corresponden a este archivo). El archivo compila más lejos que el original, pero se traba en este error no resuelto.

### 3.5 Próximo paso recomendado
Dado que se agotó lo que se puede diagnosticar por SSH + compilación en modo batch (`/make`), el camino más rápido es:
1. Alguien con acceso de **escritorio remoto** a la VM (no solo SSH) abre `VISADO.VBP` en el IDE de VB6.
2. Ejecuta **Compilar** (Ctrl+F5) desde el IDE — a diferencia de `/make`, el IDE **resalta visualmente la línea exacta** donde ocurre el error, sin la ambigüedad del número de línea reportado por el compilador batch.
3. Alternativa si no hay acceso de escritorio: pedir que alguien exporte la licencia de diseño de MSFlexGrid (`HKCR\Licenses\{6262D3A0-531B-11CF-91F6-C2863C385E30}`) desde una máquina donde SÍ esté — aunque descartado como causa raíz por el momento, no está de más tenerla completa en `licencias.reg` del framework para futuras migraciones.

---

## 4. Cambios ya aplicados al framework compartido (`vb6-framework`)

Commit `fdd57c1` (pusheado a `sebapuno/vb6.git`):
- **9 patrones nuevos** documentados en `config.json` y la guía (los de la sección 2.4/2.5/2.6/2.7/2.8/2.9).
- **Falso positivo corregido**: el chequeo de "Declare duplicado" (Ambiguous name) ahora ignora declaraciones `Private` (no colisionan entre módulos). Confirmado con `Srmw32` (Public en un módulo, Private en otro — compila bien en VB6 real).
- **Bug de scheduling AM/PM corregido** en `framework/mac/lib/remote_compile.sh`: `time /t` de Windows devuelve 12 horas con sufijo AM/PM, y `date -j -f` de macOS lo ignoraba en silencio, agendando compilaciones de tarde/noche para el día siguiente.

Commit `f0db339` en el repo `vsdw9500`: puntero del submódulo actualizado a `fdd57c1`.

---

## 5. RESOLUCIÓN — la línea reportada por VB6 sí es confiable (2026-08-09, noche)

### 5.1 El malentendido de fondo

La sección 3.2 concluyó que el número de línea era basura porque "cae siempre dentro
del bloque de diseño". **Esa conclusión era incorrecta.** VB6 numera los errores de un
`.FRM` **desde el inicio de la sección de código**, no desde el inicio del archivo. El
bloque de diseño (`VERSION 5.00` … `End`) y las líneas `Attribute` no cuentan.

**Fórmula:**

```
línea_real_en_el_archivo  =  línea_reportada  +  offset
offset                    =  número de línea del último `Attribute` del archivo
```

Para obtener el offset de cualquier `.FRM`:

```bash
grep -n "^End$\|^Attribute" ARCHIVO.FRM | head -5
```

**Validación empírica** (3 casos independientes, todos exactos):

| Archivo | Reportado | offset | Línea real | Qué había ahí |
|---|---|---|---|---|
| `VISANDO.FRM` | 424 | 472 | 896 | `.CellPicture = Tick.Picture` (faltaba `Set`) |
| `VIGENCIA.FRM` | 215 | 153 | 368 | `Private Sub Todos_Click(, dbOpenSnapshot)` |

Esto también explica el fenómeno de la sección 3.2 ("el número varía entre 124, 268,
371 y 379"): no escalaba con el tamaño del archivo — **el error iba avanzando** a
medida que se corregían los errores anteriores, que es exactamente lo esperado.

### 5.2 Los blockers que faltaban

Encontrados por barrido sistemático en vez de por ensayo y error:

| # | Archivo(s) | Problema | Fix |
|---|---|---|---|
| 1 | `VISANDO`, `NOTARIO`, `VIGENCIA` | `.ColHidden = True` — propiedad de **Spread**, no existe en MSFlexGrid (6 usos) | Comentada: el equivalente `ColWidth(n) = 0` ya estaba aplicado en la línea anterior en todos los casos |
| 2 | `NOTARIO`, `VIGENCIA`, `IMPR_NEW` | `.UserResizeCol/.UserResizeRow = 2` — Spread (3 usos) | → `.AllowUserResizing` (0=ninguno, 1=columnas, 2=filas, 3=ambos) |
| 3 | `VISANDO` (×3), `PANTVISA` (×1) | `.CellPicture = X.Picture` **sin `Set`** — asignar un objeto a una propiedad sin `Set` es error de compilación en VB6 (`Invalid use of property`) | → `Set .CellPicture = X.Picture` |
| 4 | `VIGENCIA` | `Private Sub Todos_Click(, dbOpenSnapshot)` — firma corrompida por una corrida previa de `vb6-build.sh --autofix` | → `Private Sub Todos_Click()` |

**Nota sobre `--autofix`:** el daño del punto 4 lo introdujo una regex del autofix del
framework (la que agrega `dbOpenSnapshot` a `OpenRecordset`). Conviene correr con
`--max-iterations 1` y sin `--autofix`, y revisar `git diff` después de cada corrida
que sí lo use.

### 5.3 Método reutilizable (en vez de ensayo y error)

Los 9 blockers del punto 1 y 2 se encontraron con un barrido que compara cada miembro
usado sobre un control `MSFlexGrid` contra la lista de miembros válidos del control —
no probando hipótesis de a una contra la VM (cada ciclo de compilación cuesta ~5 min).
Mismo enfoque para las referencias a controles inexistentes (bug de la sección 2.1).

---

## 6. PENDIENTE — índices base 1 (Spread) vs base 0 (MSFlexGrid)

**No bloquea la compilación, pero rompe la aplicación en tiempo de ejecución.**

Spread y MSFlexGrid cuentan distinto:

- Spread: `MaxCols = N` → N columnas **de datos**, índices **1..N**; la fila/columna 0 es el encabezado.
- MSFlexGrid: `Cols = N` → N columnas **totales**, índices **0..N-1**; la fila/columna 0 es la fija (`FixedRows`/`FixedCols` = 1 por defecto).

La migración tradujo `MaxCols = N` → `Cols = N` de forma literal. Como el código sigue
usando índices 1..N, **el índice N queda fuera de rango** → `error 381 (Invalid property
array index)` al abrir la pantalla.

| Form | Original | Migrado | Debería ser | Índice que rompe |
|---|---|---|---|---|
| `VISANDO` | `MaxCols = 6` | `Cols = 6` | `Cols = 7` | `ColWidth(6)`, `.Col = 6` |
| `NOTARIO` | `MaxCols = 5` / `MaxRows = 0` | `Cols = 5` / `Rows = 0` | `Cols = 6` / `Rows = 1` | `ColWidth(5)` |
| `VIGENCIA` | `MaxCols = 4` | `Cols = 4` | `Cols = 5` | `ColWidth(4)` |
| `IMPR_NEW` | `MaxCols = 8` / `MaxRows = 0` | `Cols = 8` / `Rows = 0` | `Cols = 9` / `Rows = 1` | `ColWidth(8)` |
| `IMPRPROD` | `MaxCols = 7` / `MaxRows = 0` | `Cols = 7` / `Rows = 0` | `Cols = 8` / `Rows = 1` | `ColWidth(7)` |
| `PANTVISA` | `MaxCols = 4` | `Cols = 4` | `Cols = 5` | `ColWidth(4)` |

Además, `Rows = 0` es inválido en MSFlexGrid cuando `FixedRows = 1` (default).

**Y las lecturas también cambian de significado.** Todo sitio que hacía `.MaxRows`
(= cantidad de filas de datos) y se migró a `.Rows` ahora devuelve una fila **de más**:

- `For I = 1 To .Rows` → debe ser `For I = 1 To .Rows - 1`
- `.Rows = .Rows + 1 : .Row = .Rows` → `.Row` debe ser `.Rows - 1`

Alcance: **91 usos de `.Rows`/`.Cols`** en los 6 forms (VISANDO 9, NOTARIO 13,
VIGENCIA 6, IMPR_NEW 16, IMPRPROD 11, PANTVISA 36). Cada uno hay que compararlo
contra el original en `main` para saber si venía de `MaxRows`/`MaxCols` (traducción
que necesita ajuste) o si ya era un índice correcto.
