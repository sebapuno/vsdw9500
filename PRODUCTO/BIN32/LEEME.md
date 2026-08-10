# Ejecutables de 32 bits (migración VB6 16→32)

Binarios compilados desde `FUENTES/` contra la VM de VB6 32-bit, en la branch
`migracion/vb6-32bit`. **No reemplazan todavía a `PRODUCTO/BIN/S/`**, que sigue
siendo el entregable de 16 bits en producción.

## Qué hay acá

- **155 ejecutables** de 32 bits, uno por `.VBP` compilado con éxito.
- 138 de ellos tienen su equivalente en el entregable actual (`PRODUCTO/BIN/S/`)
  y son los que lo reemplazarían.
- 16 no figuran en el entregable actual: se compilan desde el repo pero su `.EXE`
  no está en el paquete de producción (proyectos internos o en desuso).

## Qué FALTA

- **227 `.EXE` del entregable actual no tienen todavía versión de 32 bits.** La
  mayoría no tiene código fuente en este repo: `PRODUCTO/BIN/S` corresponde a
  `C:\Bin\S\` de la estación, que junta ejecutables de varios sistemas del banco.
  Confirmar con el responsable del paquete cuáles deben salir de acá.
- **13 `.VBP` de `FUENTES/` todavía no compilan** (de 177 compilables). Se agrupan
  en pocas causas: 4 comparten `RUTINAS.BAS`, 3 son de `VALCTDLC`, 2 dan
  `Wrong number of arguments`, 2 son `PTACRN04`/`PTACRS04` y 2 `VSDVL3CR`.
- Aparte, 2 `.VBP` (`GES_1C2`, `GES_3GR`) son **huérfanos sin código fuente** en
  el repo y su `.EXE` no está en el entregable: quedan fuera de alcance.

## RESERVA IMPORTANTE sobre estos binarios

Se compilaron contra un directorio remoto (`C:\dev\proyecto`) que **no se limpia
entre proyectos**: al terminar la tanda tenía ~723 archivos acumulados de todos
los grupos. En el repo hay 21 nombres de archivo con varias versiones distintas
(`BAJADA.BAS` tiene 21, `ERRORCOM.FRM` 10, `RUTINAS.BAS` 4).

Mientras cada `.VBP` traiga en su carpeta todos los archivos que referencia, el
deploy los sobrescribe y la compilación usa los correctos. El riesgo está en un
`.VBP` que referencie un archivo que NO está en su carpeta: ahí VB6 tomaría el
residual de otro grupo, y el binario saldría mal sin que la compilación falle.

**Antes de dar estos binarios por definitivos**: limpiar `C:\dev\proyecto` y
recompilar, al menos una muestra de los proyectos que comparten nombres de
archivo. Si el resultado no cambia, el conjunto es confiable.

## Además: ninguno fue probado ejecutándose

Estos `.EXE` compilan, pero **no se ejecutó ninguno**. Los índices de grilla
(Spread base 1 → MSFlexGrid base 0), los anchos de columna y los textos con
acentos sólo se validan de verdad abriendo las pantallas.
