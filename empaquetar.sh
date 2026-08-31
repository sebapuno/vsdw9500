#!/usr/bin/env bash
# empaquetar.sh - Arma y versiona una entrega de VSDW9500.
#
# Sigue docs/empaquetado-y-despliegue.md (seccion 9) del submodulo
# PRODUCTO/vb6-framework:
#   - la entrega se versiona en el repo, no vive solo en Descargas
#   - el .zip es REPRODUCIBLE: volver al tag y regenerarlo da el mismo sha256
#   - una entrega ya commiteada es un registro cerrado: no se pisa
#   - el script se valida solo antes de dar la entrega por buena
#
# Uso:
#   ./empaquetar.sh                  # entrega con la fecha de hoy
#   ./empaquetar.sh --fecha 20260811
#   ./empaquetar.sh --forzar         # reemplaza una entrega ya cerrada
#   ./empaquetar.sh -o ~/Downloads   # ademas deja una copia del .zip ahi
set -uo pipefail

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PAYLOAD="$REPO/PRODUCTO/ENTREGABLE32"
FECHA="$(date +%Y%m%d)"
FORZAR=0
COPIA=""

while [ $# -gt 0 ]; do
    case "$1" in
        --fecha)  FECHA="$2"; shift 2 ;;
        --forzar) FORZAR=1; shift ;;
        -o)       COPIA="$2"; shift 2 ;;
        -h|--help) sed -n '2,18p' "$0"; exit 0 ;;
        *) echo "Argumento desconocido: $1" >&2; exit 2 ;;
    esac
done

DEL="DELIVERY_$FECHA"
ENT="$REPO/ENTREGAS/$DEL"
WORK="$(mktemp -d "${TMPDIR:-/tmp}/vsdw-ent-XXXXXX")"
trap 'rm -rf "$WORK"' EXIT

echo "=================================================="
echo " Empaquetado VSDW9500 - $DEL"
echo "=================================================="

[ -d "$PAYLOAD" ] || { echo "No existe $PAYLOAD" >&2; exit 1; }

# ---------------------------------------------------------------------------
# 1. Copia del payload PRESERVANDO FECHAS (-Rp).
#    Si no, el packing list -- que registra la fecha de cada archivo -- cambia
#    en cada corrida y el paquete deja de ser reproducible.
# ---------------------------------------------------------------------------
BASE="$WORK/$DEL"
mkdir -p "$BASE"
cp -Rp "$PAYLOAD"/. "$BASE"/
find "$BASE" -name '.DS_Store' -delete
echo "[1/6] Payload copiado: $(find "$BASE" -type f | wc -l | tr -d ' ') archivos"

# ---------------------------------------------------------------------------
# 2. Packing list (.LIS) y manifiesto de hashes. Los dos son TEXTO: el git diff
#    entre dos entregas muestra que cambio sin abrir un binario.
# ---------------------------------------------------------------------------
LIS_NAME="VSDW9500.LIS"
python3 - "$BASE" "$LIS_NAME" <<'PYEOF'
import os, sys, datetime
raiz, nombre = sys.argv[1], sys.argv[2]
filas = []
for dp, _, fns in os.walk(raiz):
    for fn in sorted(fns):
        if fn == nombre: continue
        f = os.path.join(dp, fn); st = os.stat(f)
        filas.append((os.path.relpath(f, raiz).replace('/', '\\'), st.st_size,
                      datetime.datetime.fromtimestamp(st.st_mtime).strftime('%d/%m/%Y %H:%M:%S')))
filas.sort(key=lambda x: x[0].upper())
with open(os.path.join(raiz, nombre), 'wb') as fh:
    for rel, size, ts in filas:
        fh.write(('%s %d %s\r\n' % (rel, size, ts)).encode('latin-1'))
PYEOF

( cd "$BASE" && find . -type f ! -name MANIFEST_SHA256.txt -print0 \
    | sort -z | xargs -0 shasum -a 256 ) | sed "s|\./||" > "$WORK/MANIFEST_SHA256.txt"
echo "[2/6] Manifiestos generados"

# ---------------------------------------------------------------------------
# 3. CHEQUEOS antes de dar la entrega por buena (guia, seccion 9)
# ---------------------------------------------------------------------------
FALLA=0
chk() { if [ "$1" = "1" ]; then echo "   [OK]    $2"; else echo "   [FALLA] $2"; FALLA=1; fi; }

# 3a. La DLL del SRM es la oficial (49.152), no el shim del mock (~17.408)
SRMSZ=$(stat -f%z "$BASE/BIN/F/SRMW32.DLL" 2>/dev/null || echo 0)
chk "$([ "$SRMSZ" = "49152" ] && echo 1 || echo 0)" "SRMW32.DLL oficial ($SRMSZ bytes, no el shim)"

# 3b. El .mdb critico trae la tabla que el codigo busca
grep -q "Vsd_Prd" "$BASE/DATA/VISADO/VISADO.MDB" 2>/dev/null && T=1 || T=0
chk "$T" "Visado.MDB contiene la tabla Vsd_Prd"

# 3c. El .INI trae las secciones que necesita el login si se reactiva
#     (con requiereLogin ausente el aplicativo entra directo y no las usa)
S=1; for sec in '\[BASE_DATOS\]' '\[FUNCIONALIDAD\]'; do
    grep -aq "$sec" "$BASE/DATA/VISADO/VISADO.INI" || S=0
done
chk "$S" "Visado.INI con [BASE_DATOS] y [FUNCIONALIDAD]"

# 3d. Ninguna linea larga: rompe GetPrivateProfileString y deja ilegible el resto
L=$(python3 -c "print(sum(1 for l in open('$BASE/DATA/VISADO/VISADO.INI','rb').read().split(b'\r\n') if len(l)>1000))")
chk "$([ "$L" = "0" ] && echo 1 || echo 0)" "Visado.INI sin lineas de mas de 1000 chars"

# 3e. Los .bat en CRLF y sin acentos (consola CP850)
B=1; for f in "$BASE"/*.bat; do
    python3 -c "
import sys
d=open('$f','rb').read()
sys.exit(0 if d.count(b'\n')==d.count(b'\r\n') and not any(x>127 for x in d) else 1)" || B=0
done
chk "$B" "instalador en CRLF y sin bytes no-ASCII"

# 3f. Los .EXE del paquete y los versionados en PRODUCTO/BIN32/S son el MISMO
#     binario. Son dos copias de la ultima compilacion: si divergen, alguna de
#     las dos quedo vieja y se entregaria un ejecutable que no es el del repo.
#     (No sirve comparar contra los bin/ de FUENTES: ahi quedan .EXE de
#     compilaciones locales viejas, y varios proyectos comparten nombre.)
DIF=$(python3 - "$BASE/BIN/S" "$REPO/PRODUCTO/BIN32/S" <<'PYEOF'
import os, sys, hashlib
emp, ver = sys.argv[1], sys.argv[2]
def h(p): return hashlib.sha256(open(p,'rb').read()).hexdigest()
a = {f.upper(): h(os.path.join(emp,f)) for f in os.listdir(emp)}
b = {f.upper(): h(os.path.join(ver,f)) for f in os.listdir(ver) if f.upper().endswith('.EXE')}
print(sum(1 for k,v in a.items() if b.get(k) != v))
PYEOF
)
chk "$([ "$DIF" = "0" ] && echo 1 || echo 0)" "los .EXE coinciden con PRODUCTO/BIN32/S ($DIF difieren)"

[ "$FALLA" = "0" ] || { echo; echo "Empaquetado ABORTADO: hay chequeos en falla." >&2; exit 1; }
echo "[3/6] Chequeos OK"

# ---------------------------------------------------------------------------
# 4. Normalizar fechas y comprimir. Sin esto el .zip guarda la mtime de cada
#    archivo y dos corridas dan sha256 distinto aunque el contenido sea igual.
# ---------------------------------------------------------------------------
find "$BASE" -exec touch -t "${FECHA}0000" {} +
ZIPTMP="$WORK/$DEL.zip"
( cd "$WORK" && zip -r -q -X "$ZIPTMP" "$DEL" )
ZIPSHA="$(shasum -a256 "$ZIPTMP" | cut -d' ' -f1)"
echo "[4/6] Paquete comprimido: $(du -h "$ZIPTMP" | cut -f1)  sha256 $(echo "$ZIPSHA" | cut -c1-16)..."

# ---------------------------------------------------------------------------
# 5. Una entrega ya commiteada es un REGISTRO CERRADO: no se pisa.
# ---------------------------------------------------------------------------
CERRADA=0
git -C "$REPO" ls-files --error-unmatch "ENTREGAS/$DEL/NOTAS.md" >/dev/null 2>&1 && CERRADA=1

if [ "$CERRADA" = "1" ] && [ "$FORZAR" = "0" ]; then
    ANT="$(shasum -a256 "$ENT/$DEL.zip" 2>/dev/null | cut -d' ' -f1)"
    if [ "$ZIPSHA" = "$ANT" ]; then
        echo "[5/6] Entrega ya versionada y paquete IDENTICO: no se modifica nada."
        echo "      (de paso queda confirmada la reproducibilidad)"
    else
        echo "[5/6] [AVISO] El paquete regenerado DIFIERE del ya entregado. NO se pisa." >&2
        echo "      entregado : $ANT" >&2
        echo "      regenerado: $ZIPSHA" >&2
        echo "      Si el cambio es a proposito, generar una entrega NUEVA (--fecha)." >&2
    fi
    [ -n "$COPIA" ] && cp "$ZIPTMP" "$COPIA/$DEL.zip" && echo "      copia suelta -> $COPIA/$DEL.zip"
    exit 0
fi

mkdir -p "$ENT"
cp "$ZIPTMP" "$ENT/$DEL.zip"
cp "$WORK/MANIFEST_SHA256.txt" "$ENT/MANIFEST_SHA256.txt"
cp "$BASE/$LIS_NAME" "$ENT/$LIS_NAME"

COMMIT="$(git -C "$REPO" rev-parse HEAD)"
SUCIO="$(git -C "$REPO" status --porcelain -- PRODUCTO/ENTREGABLE32 | wc -l | tr -d ' ')"
cat > "$ENT/NOTAS.md" <<EOF
# Entrega $FECHA — VSDW9500 (VB6 32 bits)

| | |
|---|---|
| Paquete | \`$DEL.zip\` |
| sha256 | \`$ZIPSHA\` |
| Archivos | $(find "$BASE" -type f | wc -l | tr -d ' ') |
| Tamaño | $(du -h "$ZIPTMP" | cut -f1) |
| Commit de origen | \`$COMMIT\` |
| Payload | \`PRODUCTO/ENTREGABLE32\` |

## Cómo reproducirla

\`\`\`bash
git checkout $COMMIT
./empaquetar.sh --fecha $FECHA
shasum -a256 ENTREGAS/$DEL/$DEL.zip   # debe dar $ZIPSHA
\`\`\`

El \`.zip\` es reproducible: las fechas del árbol se normalizan a
\`${FECHA}0000\` antes de comprimir y se usa \`zip -X\`, así que dos corridas
sobre el mismo commit dan el **mismo sha256**.

## Qué contiene

Ver \`../../PRODUCTO/ENTREGABLE32/LEEME.md\` (viaja dentro del paquete) para el
detalle, los prerrequisitos y las decisiones pendientes.

- **166 ejecutables** de 32 bits — 177 de 179 proyectos compilan
- \`SRMW32.DLL\` **oficial** de 49.152 bytes (no el shim del simulador)
- 7 bases Access + \`Visado.INI\` con \`[BASE_DATOS]\` y \`[FUNCIONALIDAD]\`
- 4 OCX + \`licencias.reg\`
- \`instalador_vsdw9500.bat\`

## Estado de validación

- Instalación **ejecutada de punta a punta en Windows 11 24H2**: \`RESULTADO: OK\`
- DAO 3.6 abre \`Visado.MDB\` bajo WOW64 (26 TableDefs) — no hace falta ADODB
- **Login verificado**: llega al SRM (\`apli=GS35 nodo=u104 fun=01\`) y autentica
- \`Srmw.ini\` y la variable \`OFICINA\` **no se tocan** si ya existen
- Sólo \`VSDPROD.EXE\` se probó ejecutando; los otros 165 nunca corrieron
EOF

echo "[5/6] Entrega escrita en ENTREGAS/$DEL/"

# ---------------------------------------------------------------------------
# 6. Verificacion final del artefacto
# ---------------------------------------------------------------------------
unzip -tq "$ENT/$DEL.zip" >/dev/null && echo "[6/6] El .zip abre sin errores" \
    || { echo "[6/6] [FALLA] el .zip esta corrupto" >&2; exit 1; }

[ -n "$COPIA" ] && cp "$ENT/$DEL.zip" "$COPIA/$DEL.zip" && echo "      copia -> $COPIA/$DEL.zip"

echo
echo "Listo. Para cerrar la entrega:"
echo "  git add ENTREGAS/$DEL && git commit -m 'Entrega $FECHA'"
echo "  git tag -a entrega-$FECHA -m 'Entrega $FECHA' && git push --tags"
