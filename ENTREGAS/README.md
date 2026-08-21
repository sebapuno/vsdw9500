# Entregas de VSDW9500

Cada entrega que sale al cliente se versiona acá, con la metadata que permite
auditarla y **reproducirla** después. Convención tomada de
[`docs/empaquetado-y-despliegue.md`](../PRODUCTO/vb6-framework/docs/empaquetado-y-despliegue.md)
(sección 9) del submódulo `vb6-framework`.

## Estructura

```
ENTREGAS/
  README.md
  DELIVERY_<YYYYMMDD>/
    DELIVERY_<fecha>.zip      el artefacto exacto que se entregó
    MANIFEST_SHA256.txt       sha256 de cada archivo del paquete
    VSDW9500.LIS              packing list del payload
    NOTAS.md                  sha256 del zip, commit de origen, cómo reproducir
```

Los dos manifiestos son **texto**: el `git diff` entre dos entregas muestra
exactamente qué archivos cambiaron, sin abrir un solo binario.

## Generar una entrega

```bash
./empaquetar.sh                    # con la fecha de hoy
./empaquetar.sh --fecha 20260820   # con una fecha concreta
./empaquetar.sh -o ~/Downloads     # además deja una copia suelta del .zip
```

El script arma el paquete desde `PRODUCTO/ENTREGABLE32`, corre seis chequeos y
**aborta si alguno falla**:

| Chequeo | Por qué |
|---|---|
| `SRMW32.DLL` de 49.152 bytes | que no viaje el *shim* del simulador (17.408) |
| `Visado.MDB` tiene `Vsd_Prd` | detectar un `.mdb` renombrado al que le falta la tabla |
| `Visado.INI` con `[BASE_DATOS]` y `[FUNCIONALIDAD]` | sin eso el aplicativo **no arranca** |
| Ninguna línea de más de 1000 chars en el `.INI` | una línea larga deja ilegible todo lo que sigue |
| `.bat` en CRLF y sin bytes no-ASCII | `cmd.exe` y la consola CP850 |
| Los `.EXE` coinciden con `PRODUCTO/BIN32/S` | no entregar un ejecutable viejo |

## Reproducibilidad

El `.zip` es reproducible: **volver al commit de una entrega y regenerarla da el
mismo `sha256`**. Para eso el script copia preservando fechas (`cp -Rp`),
normaliza los timestamps del árbol a `<fecha>0000` antes de comprimir y usa
`zip -X`. Sin eso, dos corridas seguidas darían hashes distintos aunque el
contenido fuera idéntico (un `.zip` guarda la mtime de cada archivo).

Verificado: dos corridas consecutivas → mismo `sha256`.

## Una entrega commiteada es un registro cerrado

Cuando el `NOTAS.md` de una entrega ya está commiteado, el script **no la pisa**:

| Situación | Qué hace |
|---|---|
| Entrega nueva | La escribe |
| Cerrada, paquete **idéntico** | No toca nada — confirma la reproducibilidad |
| Cerrada, paquete **difiere** | Avisa con ambos `sha256` y **no escribe** |
| `--forzar` | La reemplaza a propósito |

El caso "difiere" es el valioso: significa que el código o el producto cambiaron
desde que se entregó. Eso amerita una **entrega nueva**, no pisar la vieja.

## Cerrar una entrega

```bash
git add ENTREGAS/DELIVERY_<fecha> && git commit -m "Entrega <fecha>"
git tag -a entrega-<fecha> -m "Entrega <fecha>" && git push --tags
```
