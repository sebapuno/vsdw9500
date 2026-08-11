# VSDW9500 — Entregable de 32 bits

Paquete de instalación del aplicativo **Visado de Productos** migrado de VB6
16 bits a 32 bits. Armado según
[`docs/empaquetado-y-despliegue.md`](../vb6-framework/docs/empaquetado-y-despliegue.md)
del framework.

Generado el **2026-08-11** desde la branch `migracion/vb6-32bit`.
**No reemplaza todavía a `PRODUCTO/BIN/S/`**, que sigue siendo el entregable de
16 bits en producción.

## Contenido

| Ruta | Qué es | Destino en la estación |
|---|---|---|
| `BIN/S/*.EXE` | **166 ejecutables** de 32 bits | `C:\Bin\S\` |
| `BIN/F/SRMW32.DLL` | cliente SRM **original** (49.152 bytes) | `C:\Bin\F\` |
| `DATA/VISADO/*.MDB` | bases Access (DAO 3.6) | `C:\Data\Visado\` |
| `DATA/VISADO/*.INI` | config de runtime | `C:\Data\Visado\` |
| `ETC/BIN/*.OCX` | controles + `licencias.reg` | se registran, `C:\Etc\Bin\` |
| `VSDW9500.LIS` | **manifiesto** (packing list) | no se instala |
| `instalador_vsdw9500.bat` | instalador de estación | no se instala |

**Ojo con el SRM:** acá va la librería **original de 49.152 bytes**
(md5 `9bbd5c4819894bd306e48b5b1828a4f7`). En el repo también existe el *shim*
del simulador (`PRODUCTO/srm/srm_shim/SRMW32.DLL`, 17.408 bytes) — ese es sólo
para pruebas contra el mock y **no** debe viajar. Verificar el tamaño antes de
entregar: 49.152 = original, 17.408 = shim.

Los OCX (`THREED32`, `MSFLXGRD`, `COMDLG32`, `MSCOMCT2`) salen de lo que
declaran los `.VBP`: THREED32 en 82 proyectos, MSFLXGRD en 48, COMDLG32 en 10 y
MSCOMCT2 en 1. Ya no viaja nada de Spread/VBX ni Crystal Reports.

El **manifiesto se regenera** en cada reempaquetado; lista todo el paquete menos
sí mismo. Formato por línea: `RUTA\RELATIVA TAMANO DD/MM/YYYY HH:MM:SS`.

## Instalación

En una estación ya configurada del banco (el caso normal):

```bat
instalador_vsdw9500.bat /OFICINA 715
```

Requiere **Administrador**. Opciones: `/SOURCE` (ruta del paquete si no se
ejecuta desde su carpeta), `/OFICINA` (variable de ambiente), `/CHECKONLY` para
verificar una estación **sin instalar nada**, y `/SRMHOST` + `/SRMPORT` —
que **sólo se usan si la estación no tiene `Srmw.ini`** (ver abajo).

### `C:\Windows\Srmw.ini` NO se pisa

Es un archivo **delicado y compartido** con otros aplicativos del banco: tiene
la configuración real de nodos de la estación. El instalador **nunca lo
sobrescribe si ya existe**, aunque se le pase `/SRMHOST`; sólo informa el
`Host=` que encontró.

Únicamente cuando **no existe** (estación nueva) lo crea, y ahí sí hace falta
`/SRMHOST` con el host del ambiente destino. Si no existe y no se indicó
`/SRMHOST`, el instalador avisa y sigue, sin inventar un valor.

> El `/SRMHOST 192.168.x.x` que aparece en ejemplos de prueba corresponde al
> **ambiente de desarrollo** (VM + simulador SRM). **No usar ese valor en el
> banco.**

### Config de la estación

El instalador **preserva el `Visado.INI` de la estación** en un upgrade: lo
respalda antes de copiar y lo restaura después. Sólo en una instalación limpia
queda el `Visado.INI` del paquete, que hay que ajustar con los nodos reales.

Al final corre una verificación que informa archivos faltantes, OCX sin
registrar y el `Host=` del SRM. Los controles se verifican **por TypeLib**, no
por CLSID (el CLSID da falso "no registrado" aunque el control funcione).

## Prerrequisitos de la estación (no vienen en el paquete)

- **Runtime VB6**: `MSVBVM60.DLL` en `C:\Windows\system32`
- **DAO 3.6**: `DAO360.DLL` en `C:\Program Files\Common Files\Microsoft Shared\DAO`
- Gateway SRM alcanzable y los nodos resolubles (`hosts` o DNS)

El instalador los **verifica** pero no los instala.

## Estado real de lo que se entrega — LEER

### Compilación: 177 de 179 proyectos

Los 2 que faltan **no son trabajo pendiente**: sus `.VBP` vienen incompletos del
repositorio original.

- `GES_1C2` — su `.VBP` en la rama original sólo incluye `BAJADA.BAS` y
  `ERRORCOM.FRM`, y ninguno define la `SacaValorINI` que `BAJADA.BAS` llama.
  **Nunca compiló, tampoco en 16 bits.**
- `GES_3GR` — referencia `GESLCD01.FRM`, que **no existe** en el repo, ni en la
  rama original, ni en ninguna carpeta.

Para cerrarlos hay que conseguir los fuentes originales del banco.

### Prueba de ejecución: uno solo

**`VSDPROD.EXE` (VISADO) es el único que se probó corriendo**, contra la VM y un
simulador SRM. Llega a la pantalla principal, autentica el login contra el SRM,
abre `Visado.MDB` y despliega la grilla de "Seleccionar Productos" con los datos
correctos.

Ese único recorrido destapó **seis** fallas que la compilación no detecta
(errores 75, 6, 424, 30009 ×2 y 3019), y **la mitad las había introducido la
propia migración**. Todas corregidas y replicadas al resto del repo.

**Los otros 165 ejecutables no se ejecutaron nunca.** Compilar valida sintaxis,
no comportamiento: hay que asumir que otros proyectos esconden fallas
equivalentes, y presupuestar una fase de prueba en ejecución.

### Nombres de `.EXE` repetidos — DECISIÓN PENDIENTE

Ocho `.EXE` los genera más de un `.VBP`, desde carpetas distintas, y en esta
carpeta plana sólo puede quedar uno. Se resolvió con el criterio **"carpeta
base"** (se descarta la variante con prefijo `pyme2000_`/`sybase_` o el
subdirectorio `*_par`), porque es el único con respaldo: en `VSDVL2FC.EXE` el
binario que ya estaba en el entregable coincide en tamaño con la variante base
y no con la de PYME2000.

| `.EXE` | se tomó de | se descartó |
|---|---|---|
| `BAJCTDLC.EXE` | `valctdlc` | `valctdlc_baja_par` |
| `EXEGTC00.EXE` | `vsdlc2.vb` | `vsdlcd.vb` |
| `GES_2CD.EXE` | `vsdlc2.vb` | `vsdlcd.vb` |
| `VLDVSCTD.EXE` | `valctdlc` | `valctdlc_vald_par` |
| `VSDVL2FC.EXE` | `vsdcce.vb` | `pyme2000_vsdcce.vb` |
| `VSDVL2PH.EXE` | `vsdcph.vb` | `sybase_vsdcph.vb` |
| `VSDVL3CR.EXE` | `vsdrth.vb` | `sybase_vsdrth.vb` |
| `VSDVL3FC.EXE` | `vsdcce.vb` | `pyme2000_vsdcce.vb` |

**Los descartados NO son iguales al elegido** (mismo tamaño en varios casos,
pero distinto MD5). Si para alguno el paquete debe llevar la variante
PYME2000/SYBASE, hay que reemplazarlo a mano. **Confirmar con el responsable
del paquete.**

Aparte, 3 pares de `.VBP` de una misma carpeta generan el mismo `.EXE`
(`EXEACR00` en `vsdrce.vb`, `ROTOR2RL` y `ROTOR3RL` en `vsdcrl.vb`): ahí el
último compilado pisa al anterior. Viene así del repo original.

### Alcance del paquete de producción

`PRODUCTO/BIN/S` (el entregable de 16 bits) tiene **227 `.EXE` que no tienen
fuente en este repo**: corresponde a `C:\Bin\S\` de la estación, que junta
ejecutables de varios sistemas del banco. **Definir con el responsable del
paquete** cuáles deben salir de acá y cuáles los provee otro equipo.

### Paridad visual pendiente

En las grillas migradas de Spread a MSFlexGrid, las celdas que en el original
eran **checkbox** (`CellType=10` + `TypeCheckText`) hoy muestran el valor `0`/`1`
como texto: MSFlexGrid no tiene celda-checkbox nativa. El comportamiento
(marcar/desmarcar con doble clic) funciona igual. Diferido para el cierre; si
hace falta, se resuelve superponiendo un `CheckBox` sobre la celda.
