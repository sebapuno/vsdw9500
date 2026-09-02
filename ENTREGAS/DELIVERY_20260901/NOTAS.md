# Entrega 20260901 — VSDW9500 (VB6 32 bits)

| | |
|---|---|
| Paquete | `DELIVERY_20260901.zip` |
| sha256 | `762b06ce13bad697f77cda01c1dc42f97164b015e958331441d3ca293fab46a4` |
| Archivos | 184 |
| Tamaño | 5.4M |
| Commit de origen | `318f53e06883b812200ba14667a0cc2512a612ab` |
| Payload | `PRODUCTO/ENTREGABLE32` |

## Cómo reproducirla

```bash
git checkout 318f53e06883b812200ba14667a0cc2512a612ab
./empaquetar.sh --fecha 20260901
shasum -a256 ENTREGAS/DELIVERY_20260901/DELIVERY_20260901.zip   # debe dar 762b06ce13bad697f77cda01c1dc42f97164b015e958331441d3ca293fab46a4
```

El `.zip` es reproducible: las fechas del árbol se normalizan a
`202609010000` antes de comprimir y se usa `zip -X`, así que dos corridas
sobre el mismo commit dan el **mismo sha256**.

## Qué contiene

Ver `../../PRODUCTO/ENTREGABLE32/LEEME.md` (viaja dentro del paquete) para el
detalle, los prerrequisitos y las decisiones pendientes.

- **166 ejecutables** de 32 bits — 177 de 179 proyectos compilan
- `SRMW32.DLL` **oficial** de 49.152 bytes (no el shim del simulador)
- 7 bases Access + `Visado.INI` con `[BASE_DATOS]` y `[FUNCIONALIDAD]`
- 4 OCX + `licencias.reg`
- `instalador_vsdw9500.bat`

## Estado de validación

- Instalación **ejecutada de punta a punta en Windows 11**: 166/166 archivos de
  `BIN/S` actualizados, verificado por fecha contra el paquete (no por el
  `RESULTADO: OK`, que antes mentía — ver abajo)
- El instalador **aborta** si hay aplicativos del paquete corriendo: `xcopy` sin
  `/C` cortaba el árbol entero y la instalación reportaba OK sin copiar nada
- DAO 3.6 abre `Visado.MDB` bajo WOW64 (26 TableDefs) — no hace falta ADODB
- **Sin autenticación**: VISADO entra directo, no pide clave. El `Visado.INI` no
  trae sección `[LOGIN]`; `requiereLogin=1` la reactiva. Verificado corriendo en
  XP (contra el simulador SRM, llega a la pantalla principal) y en Windows 11
  (ya instalado por el instalador, no aparece `ctrclave`)
- Sin login el RUT que se graba en las transacciones es fijo (`12345678`):
  **no hay trazabilidad por usuario**
- `Srmw.ini` y la variable `OFICINA` **no se tocan** si ya existen
- `SRM_NS` la pone el software básico; si falta, el aplicativo abre con un modal
  de advertencia que hay que aceptar
- Sólo `VSDPROD.EXE` se probó ejecutando; los otros 165 nunca corrieron
