#!/usr/bin/env bash
# Aplica la capa de Degú a un árbol de Euro-Office (el que `compilar.sh` compila y el que
# `publicar-fork.sh` publica): el editor servido también en el puerto 8080 bajo /editor/.
#
#   degu/aplicar.sh <raíz del código de Euro-Office>
#
# Agrega la marca de Degú en web-apps (`theme/degu`), un archivo de nginx, el plugin de Degú y cuatro líneas al final
# de la receta de la imagen standalone. El nombre y el «Acerca de» que ve la persona los completa la configuración que
# firma el motor.
set -euo pipefail
AQUI="$(cd "$(dirname "$0")" && pwd)"
RAIZ="${1:?Uso: aplicar.sh <raíz del código de Euro-Office>}"
RECETA="$RAIZ/build/.docker/standalone.bake.Dockerfile"
[[ -f "$RECETA" ]] || { echo "no encontré $RECETA" >&2; exit 1; }
# LA MARCA DE DEGÚ (`theme/degu`, el fundador, 9-nov-2026): la cara de Degú sobre el tema «white» de Euro-Office. Va
# ANTES de «ya está aplicada» y se puede repetir: vive en el submódulo web-apps, que `publicar-fork.sh` no trae, así que
# quien compile el fork corre esto de nuevo después de `git submodule update` (ver degu/README.md). Se parte de la marca
# de Euro-Office (logos y «Acerca de») y se pone la nuestra encima; la receta de web-apps la elige con `THEME`, que viene
# fija en `euro-office`. Con web-apps presente, si Euro-Office cambió cómo se elige, el build falla en vez de salir sin
# la marca.
TEMAS="$RAIZ/web-apps/theme"
RECETA_WEB="$RAIZ/web-apps/.docker/web-apps.bake.Dockerfile"
if [[ -n "$(ls -A "$RAIZ/web-apps" 2>/dev/null)" ]]; then
  [[ -d "$TEMAS/euro-office" ]] || { echo "no encontré $TEMAS/euro-office" >&2; exit 1; }
  rm -rf "$TEMAS/degu"
  cp -R "$TEMAS/euro-office" "$TEMAS/degu"
  cp -R "$AQUI/theme/degu/." "$TEMAS/degu/"
  # `THEME` y el nombre, como palabras enteras: ni `OLD_THEME=euro-office` ni `THEME=euro-office-dark` son esta marca.
  sed -E -i.orig 's/(^|[^A-Za-z0-9_])THEME=euro-office([^A-Za-z0-9_-]|$)/\1THEME=degu\2/g' "$RECETA_WEB"
  rm -f "$RECETA_WEB.orig"
  # Se mira el RESULTADO: que no quede ninguna elección de la marca de Euro-Office y que haya una de la de Degú (`terra`).
  if grep -Eq '(^|[^A-Za-z0-9_])THEME=euro-office([^A-Za-z0-9_-]|$)' "$RECETA_WEB" \
    || ! grep -Eq '^([^#]*[^A-Za-z0-9_#])?THEME=degu([^A-Za-z0-9_-]|$)' "$RECETA_WEB"; then
    echo "la receta de web-apps no quedó eligiendo la marca de Degú con THEME" >&2; exit 1
  fi
  echo "marca de Degú puesta en web-apps"
else
  echo "sin web-apps (falta git submodule update): la marca de Degú se pone al compilar"
fi

# Las imágenes oficiales de Docker, del espejo de AWS (CLAUDE.md de Degú, «Imágenes base»): el
# límite de Docker Hub por IP ya tumbó un deploy. Las que no tienen registro: con tag
# (`ubuntu:24.04`, `node:20`) y también SIN tag (`FROM alpine`, que es `alpine:latest`: el 9-nov-2026
# tumbó la imagen del editor con un 429). Sin tag también se escriben las etapas internas, y esas no
# se tocan: las de `contexts` de bake (`FROM core`, `FROM finalubuntu`), y las `AS <nombre>` de ESA
# receta declaradas más arriba (en `FROM alpine AS alpine` la imagen es la de Docker Hub). Docker
# compara los nombres de etapa sin mayúsculas. Va antes de «ya está aplicada» y se puede repetir (lo
# que ya apunta al espejo no se toca): quien compile el fork lo corre de nuevo después de
# `git submodule update`, y así también llegan al espejo las recetas de los submódulos.
recetas=()
while IFS= read -r -d '' f; do recetas+=("$f"); done < <(find "$RAIZ" -path '*/.docker/*' -name '*Dockerfile' -print0)
CONTEXTOS="$(find "$RAIZ" -name '*.hcl' -exec grep -hoiE '"?[a-z0-9._-]+"?[[:space:]]*=[[:space:]]*"target:' {} + \
  | sed -E 's/[[:space:]]*=.*//; s/"//g'; echo scratch)"
export CONTEXTOS
for receta in "${recetas[@]}"; do
  awk '
    BEGIN { n = split(tolower(ENVIRON["CONTEXTOS"]), c, "\n"); for (i = 1; i <= n; i++) etapa[c[i]] = 1 }
    /^FROM[[:space:]]/ {
      k = ($2 ~ /^--platform=/) ? 3 : 2
      if ($k !~ /[:\/@$]/ && !(tolower($k) in etapa)) $k = $k ":latest"
      for (i = k + 1; i < NF; i++) if (toupper($i) == "AS") etapa[tolower($(i + 1))] = 1
    }
    { print }' "$receta" > "$receta.orig"
  cat "$receta.orig" > "$receta"
done
sed -E -i.orig \
  's#^(FROM[[:space:]]+(--platform=[^[:space:]]+[[:space:]]+)?)([a-z0-9._-]+:[^[:space:]/]+)([[:space:]]|$)#\1public.ecr.aws/docker/library/\3\4#' \
  "${recetas[@]}"
find "$RAIZ" -path '*/.docker/*' -name '*.orig' -delete

grep -q 'degu-editor.conf' "$RECETA" && { echo "la capa de Degú ya está aplicada"; exit 0; }

cp "$AQUI/nginx-editor.conf" "$RAIZ/build/configs/standalone/degu-editor.conf"
# El plugin de Degú (`plugin/`): lee el documento abierto y aplica lo que Degú cambia. El motor lo
# nombra en la configuración que firma (`editor_oficina.PLUGIN_DE_DEGU`).
rm -rf "$RAIZ/build/configs/standalone/degu-plugin"
cp -R "$AQUI/plugin" "$RAIZ/build/configs/standalone/degu-plugin"
cat >> "$RECETA" <<'RECETA'

# Degú (https://github.com/prohack-cl/degu-editor, carpeta degu/): el editor bajo /editor/, detrás
# de un balanceador. El nginx de la imagen tiene que incluir conf.d: si deja de hacerlo, mejor que
# el build falle a que el editor arranque sin esa ruta.
RUN grep -Eq 'include +/etc/nginx/conf\.d/\*\.conf' /etc/nginx/nginx.conf
COPY build/configs/standalone/degu-editor.conf /etc/nginx/conf.d/degu-editor.conf
COPY build/configs/standalone/degu-plugin ${EO_ROOT}/sdkjs-plugins/degu-al-lado
EXPOSE 8080
RECETA
echo "capa de Degú aplicada en $RAIZ"
