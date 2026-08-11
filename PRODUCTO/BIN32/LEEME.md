# Ejecutables de 32 bits (migración VB6 16→32)

Binarios compilados desde `FUENTES/` contra la VM de VB6 32-bit, en la branch
`migracion/vb6-32bit`. **No reemplazan todavía a `PRODUCTO/BIN/S/`**, que sigue
siendo el entregable de 16 bits en producción.

Última recompilación completa: **2026-08-11**, los 179 `.VBP` del repo en una
sola tanda (44 minutos, con `framework/mac/vb6-build-batch.sh`).

## Qué hay acá

- **159 ejecutables** de 32 bits, uno por `.VBP` compilado con éxito
  (169 proyectos compilan; 10 de ellos generan un `.EXE` de nombre repetido).
- Todos fueron regenerados el 2026-08-11: **155 reemplazan** a la versión
  anterior de esta misma carpeta y **4 son nuevos**.
- Ningún `.EXE` que estuviera acá quedó sin recompilar.

## Qué FALTA

- **227 `.EXE` del entregable actual no tienen todavía versión de 32 bits.** La
  mayoría no tiene código fuente en este repo: `PRODUCTO/BIN/S` corresponde a
  `C:\Bin\S\` de la estación, que junta ejecutables de varios sistemas del banco.
  Confirmar con el responsable del paquete cuáles deben salir de acá.
- **10 `.VBP` de `FUENTES/` todavía no compilan** (de 179). Por causa:
  - 3 comparten `RUTINAS.BAS` (`EXECRE00`, `EXECRN00`, `EXECRS00`)
  - 3 dan `Wrong number of arguments` (`VSDVL2PH`, `EXE2F000`, `EXEAE000`)
  - 2 dan `Errors during load` (`PTACRN04`, `PTACRS04`)
  - `GES_1C2` — `Sub or Function not defined` en `BAJADA.BAS`
  - `GES_3GR` — referencia `GESLCD01.FRM`, que **no existe en el repo ni en la
    rama original `main`**: falta desde el origen, no lo borró la migración.

## Nombres de `.EXE` repetidos entre proyectos — DECISIÓN PENDIENTE

Siete `.EXE` los genera más de un `.VBP`, desde carpetas distintas. Sólo puede
quedar uno en esta carpeta plana. Se resolvió con el criterio **"carpeta base"**
(se descarta la variante con prefijo `pyme2000_`/`sybase_` o el subdirectorio
`*_par`), porque es el único que tiene respaldo: en `VSDVL2FC.EXE` el binario
que ya estaba en el entregable coincide en tamaño con la variante base
(110.592 bytes) y no con la de PYME2000 (114.688).

| `.EXE` | se tomó de | se descartó |
|---|---|---|
| `BAJCTDLC.EXE` | `valctdlc` | `valctdlc_baja_par` |
| `EXEGTC00.EXE` | `vsdlc2.vb` | `vsdlcd.vb` |
| `GES_2CD.EXE` | `vsdlc2.vb` | `vsdlcd.vb` |
| `VLDVSCTD.EXE` | `valctdlc` | `valctdlc_vald_par` |
| `VSDVL2FC.EXE` | `vsdcce.vb` | `pyme2000_vsdcce.vb` |
| `VSDVL3CR.EXE` | `vsdrth.vb` | `sybase_vsdrth.vb` |
| `VSDVL3FC.EXE` | `vsdcce.vb` | `pyme2000_vsdcce.vb` |

**Los binarios descartados NO son iguales al elegido** (mismo tamaño en varios
casos, pero distinto MD5). Si para alguno de estos siete el paquete de
producción debe llevar la variante PYME2000/SYBASE, hay que reemplazarlo a mano.
Confirmar con el responsable del paquete.

Aparte, 3 pares de `.VBP` de una misma carpeta generan el mismo `.EXE`
(`EXEACR00` en `vsdrce.vb`, `ROTOR2RL` y `ROTOR3RL` en `vsdcrl.vb`): ahí el
último compilado pisa al anterior. Viene así del repo original.

## Estado de prueba: uno solo se ejecutó

- **`VSDPROD.EXE` (VISADO)** es el único que se probó corriendo, contra la VM y
  el simulador SRM. Llega a la pantalla principal, autentica el login contra el
  SRM, abre `Visado.MDB` y despliega la grilla de "Seleccionar Productos" con
  los datos correctos. Ese recorrido destapó **seis** fallas que la compilación
  no detecta (errores 75, 6, 424, 30009 ×2, 3019); todas corregidas y aplicadas
  también al resto del repo donde correspondía.
- **Los otros 158 no se ejecutaron nunca.** Compilar valida sintaxis, no
  comportamiento: de las seis fallas de VISADO, la mitad las había introducido
  la propia migración y ninguna se veía compilando. Hay que asumir que otros
  proyectos esconden equivalentes.

## Paridad visual pendiente

En las grillas migradas de Spread a MSFlexGrid, las celdas que en el original
eran **checkbox** (`CellType=10` + `TypeCheckText`) hoy muestran el valor `0`/`1`
como texto: MSFlexGrid no tiene celda-checkbox nativa. El comportamiento
(marcar/desmarcar con doble clic) funciona igual. Diferido de común acuerdo para
el cierre; si hace falta, se resuelve superponiendo un `CheckBox` sobre la celda.
