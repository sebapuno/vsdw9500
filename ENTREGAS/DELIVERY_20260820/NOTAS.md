# Entrega 20260820 — VSDW9500 (VB6 32 bits)

| | |
|---|---|
| Paquete | `DELIVERY_20260820.zip` |
| sha256 | `cfb79d586bb7b31709f5f33f1bdd932a2ac9d98425c6239db38d2c6ff422c7d0` |
| Archivos | 184 |
| Tamaño | 5.5M |
| Commit de origen | `9e3dda912a34244810c1e2eec761f2bf2c2aef76` |
| Payload | `PRODUCTO/ENTREGABLE32` |

## Cómo reproducirla

```bash
git checkout 9e3dda912a34244810c1e2eec761f2bf2c2aef76
./empaquetar.sh --fecha 20260820
shasum -a256 ENTREGAS/DELIVERY_20260820/DELIVERY_20260820.zip   # debe dar cfb79d586bb7b31709f5f33f1bdd932a2ac9d98425c6239db38d2c6ff422c7d0
```

El `.zip` es reproducible: las fechas del árbol se normalizan a
`202608200000` antes de comprimir y se usa `zip -X`, así que dos corridas
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

- Instalación **ejecutada de punta a punta en Windows 11 24H2**: `RESULTADO: OK`
- DAO 3.6 abre `Visado.MDB` bajo WOW64 (26 TableDefs) — no hace falta ADODB
- **Login verificado**: llega al SRM (`apli=GS35 nodo=u104 fun=01`) y autentica
- `Srmw.ini` y la variable `OFICINA` **no se tocan** si ya existen
- Sólo `VSDPROD.EXE` se probó ejecutando; los otros 165 nunca corrieron
