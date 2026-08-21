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

En una estación ya configurada del banco (el caso normal, sin argumentos):

```bat
instalador_vsdw9500.bat
```

Requiere **Administrador**. Opciones: `/SOURCE` (ruta del paquete si no se
ejecuta desde su carpeta), `/CHECKONLY` para verificar una estación **sin
instalar nada**, y `/SRMHOST` + `/SRMPORT` — que **sólo se usan si la estación
no tiene `Srmw.ini`** (ver abajo).

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

### La variable `OFICINA` no la pone este instalador

La asigna el **software básico** de la estación, que no forma parte de este
entregable. El instalador **sólo informa** si está definida y avisa cuando
falta; nunca la escribe. Ponerla a mano desde acá podría dejarla en desacuerdo
con la oficina real y hacer que las consultas salgan mal.

Si falta, hay que reclamarla al área que administra la estación.

> Cómo la usa el aplicativo: lee primero `[Visado] Oficina` del `Visado.INI`; si
> está vacía, toma la variable de ambiente y la persiste en el INI. O sea que el
> valor correcto tiene que estar puesto **antes** del primer arranque.

### El `Visado.INI` del paquete — dos cosas que hay que saber

**1. Secciones nuevas obligatorias.** El aplicativo migrado carga los
parámetros de acceso a la base desde `[BASE_DATOS]`, que **no existía** en el
INI anterior. Si falta `NODO`, `APLI` o `BDATOS`, VISADO muestra
`Faltan parametros [BASE_DATOS]` y **termina**: no llega ni al login.

| Sección | Claves | Para qué |
|---|---|---|
| `[BASE_DATOS]` | `NODO`, `APLI`, `BDATOS` | acceso a la base en el login — **sin esto no arranca** |
| `[FUNCIONALIDAD]` | `FUNCIONALIDAD` | código con el que el login valida al usuario |
| `[VISADO]` | `LogLlamadas` | reemplaza la ruta `d:\` que estaba hardcodeada |
| `[Logea]` | `Flag` | activa el log de llamadas al SRM |

Los valores del paquete son los **de referencia** (los mismos que usa
WCRE9500): `NODO=u104`, `APLI=GS35`, `BDATOS=mcambio`, `FUNCIONALIDAD=88`.
**Verificarlos contra el ambiente destino antes de instalar.**

**2. `[PARAMETROS]` va al final a propósito.** El aplicativo reescribe ahí
`LlaveVisad` con el contenido de la grilla, que puede pasar los **12.000
caracteres**. `GetPrivateProfileString` deja de leer todo lo que venga
**después** de una línea así de larga — y lo hace **en silencio**, devolviendo
cadena vacía sin ningún error. El INI que venía en el repo tenía esa línea en
el medio, con `[SIEBEL]`, `[SYBASE]` y el resto detrás: todas ilegibles.

Por eso `[PARAMETROS]` quedó al final y `LlaveVisad` se entrega vacío. Si en
algún momento hay que reordenar el archivo, mantener esa sección última.

### Config de la estación

El instalador **preserva el `Visado.INI` de la estación** en un upgrade: lo
respalda antes de copiar y lo restaura después. Sólo en una instalación limpia
queda el `Visado.INI` del paquete, que hay que ajustar con los nodos reales.

Al final corre una verificación que informa archivos faltantes, OCX sin
registrar y el `Host=` del SRM. Los controles se verifican **por TypeLib**, no
por CLSID (el CLSID da falso "no registrado" aunque el control funcione).

## Prerrequisitos de la estación (no vienen en el paquete)

| Componente | Windows XP | Windows de 64 bits (10/11) |
|---|---|---|
| Runtime VB6 `MSVBVM60.DLL` | `C:\Windows\system32` | `%SystemRoot%\SysWOW64` |
| DAO 3.6 `DAO360.DLL` | `…\Program Files\Common Files\…\DAO` | `…\Program Files (x86)\Common Files\…\DAO` |

- Gateway SRM alcanzable y los nodos resolubles (`hosts` o DNS)
- **Software básico** de la estación instalado: es el que asigna la variable
  `OFICINA` (y `C:\Windows\Srmw.ini`)

El instalador los **verifica** pero no los instala, y prueba **las dos
ubicaciones** según la arquitectura.

### Instalación real probada en Windows 11 24H2 (2026-08-20)

El instalador se corrió **completo** (no sólo `/CHECKONLY`) en una VM
Windows 11 24H2 (build 26100), elevado, y terminó en **`RESULTADO: OK`** con
todos los chequeos en verde. Quedaron instalados los 166 `.EXE` en `C:\Bin\S`,
la `SRMW32.DLL` de 49.152 bytes en `C:\Bin\F` y las bases y config en
`C:\Data\Visado`.

Las dos protecciones se comportaron como corresponde:

```
[6/7] C:\Windows\Srmw.ini ya existe: NO se toca (archivo compartido)
      Host actual = 192.168.64.1
[7/7] OFICINA=000 (la asigna el software basico, no el instalador)
```

> **Nota sobre elevación:** por SSH, una cuenta administradora llega con el
> token filtrado por UAC y `net session` falla, así que el instalador aborta
> pidiendo Administrador. Para probarlo remoto hay que lanzarlo elevado — por
> ejemplo con `schtasks /create ... /rl HIGHEST` y `schtasks /run`. En la
> estación, ejecutándolo desde una consola "como Administrador", no aplica.

Los dos prerrequisitos **ya vienen instalados de fábrica**, no hay que agregar
nada:

```
[OK]    MSVBVM60.DLL (runtime VB6, SysWOW64)
[OK]    DAO360.DLL   (DAO 3.6, Program Files x86)
```

**DAO 3.6 funciona bajo WOW64**: se abrió el `Visado.MDB` del paquete desde el
`cscript` de 32 bits y listó sus **26 TableDefs** (ADODB+Jet también funciona,
por si alguna vez hiciera falta). **No hay que migrar el acceso a datos a
ADODB.**

> Si alguna vez aparece un `80040154` al crear `DAO.DBEngine.36`, ese código es
> `REGDB_E_CLASSNOTREG` — significa "DAO no está registrado en esta máquina",
> **no** "DAO no anda en 64 bits". Se arregla instalando DAO, no rehaciendo el
> acceso a datos.

Dos detalles del `.bat` que sólo se ven en 64 bits, y que ya están resueltos:
`regsvr32` se toma de `SysWOW64` (el de `system32` es de 64 bits y falla con
OCX de 32), y las rutinas de chequeo **no usan bloques `( … )`** — `%~1` quita
las comillas y el `)` de `Program Files (x86)` cerraría el bloque, abortando
`cmd` con "No se esperaba \Common en este momento". En XP no se nota porque
ahí la ruta no tiene paréntesis.

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
