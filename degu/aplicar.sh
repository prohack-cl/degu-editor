#!/usr/bin/env bash
# Aplica la capa de Degú a un árbol de Euro-Office (el que `compilar.sh` compila y el que
# `publicar-fork.sh` publica): el editor servido también en el puerto 8080 bajo /editor/.
#
#   degu/aplicar.sh <raíz del código de Euro-Office>
#
# Agrega un archivo de nginx y tres líneas al final de la receta de la imagen standalone. La marca
# de la interfaz y el «Acerca de» no van acá: los pone la configuración que firma el motor.
set -euo pipefail
AQUI="$(cd "$(dirname "$0")" && pwd)"
RAIZ="${1:?Uso: aplicar.sh <raíz del código de Euro-Office>}"
RECETA="$RAIZ/build/.docker/standalone.bake.Dockerfile"
[[ -f "$RECETA" ]] || { echo "no encontré $RECETA" >&2; exit 1; }
grep -q 'degu-editor.conf' "$RECETA" && { echo "la capa de Degú ya está aplicada"; exit 0; }

# Las imágenes oficiales de Docker, del espejo de AWS (CLAUDE.md de Degú, «Imágenes base»): el
# límite de Docker Hub por IP ya tumbó un deploy. Solo las que llevan tag y no tienen registro
# (`ubuntu:24.04`, `node:20`); las etapas internas (`FROM core`) no llevan tag y no se tocan.
find "$RAIZ" -path '*/.docker/*' -name '*Dockerfile' -print0 | xargs -0 sed -E -i.orig \
  's#^(FROM[[:space:]]+(--platform=[^[:space:]]+[[:space:]]+)?)([a-z0-9._-]+:[^[:space:]]+)#\1public.ecr.aws/docker/library/\3#'
find "$RAIZ" -path '*/.docker/*' -name '*.orig' -delete

cp "$AQUI/nginx-editor.conf" "$RAIZ/build/configs/standalone/degu-editor.conf"
cat >> "$RECETA" <<'RECETA'

# Degú (https://github.com/prohack-cl/degu-editor, carpeta degu/): el editor bajo /editor/, detrás
# de un balanceador. El nginx de la imagen tiene que incluir conf.d: si deja de hacerlo, mejor que
# el build falle a que el editor arranque sin esa ruta.
RUN grep -Eq 'include +/etc/nginx/conf\.d/\*\.conf' /etc/nginx/nginx.conf
COPY build/configs/standalone/degu-editor.conf /etc/nginx/conf.d/degu-editor.conf
EXPOSE 8080
RECETA
echo "capa de Degú aplicada en $RAIZ"
