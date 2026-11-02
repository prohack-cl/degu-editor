# Degú Editor

El editor de documentos y planillas de [Degú](https://getdegu.com). Es
[Euro-Office](https://github.com/Euro-Office/DocumentServer), a su vez basado en ONLYOFFICE, con
licencia **GNU AGPLv3**. Este repositorio es el código que corre en Degú: el commit de Euro-Office
de esta rama más la capa de `degu/`, que solo agrega cómo se sirve el editor bajo `/editor/`: un
archivo de nginx y tres líneas al final de `build/.docker/standalone.bake.Dockerfile`, que deja
`degu/aplicar.sh`.

## Cómo se compila

```sh
git submodule update --init --recursive
cd build && docker buildx bake standalone --load --set "*.platform=linux/amd64" \
  --set "standalone.tags=degu-editor:local"
```

## Licencia

AGPLv3, como Euro-Office y ONLYOFFICE (ver `LICENSE` y `ATTRIBUTION`). Los avisos legales de la
interfaz se mantienen, y el «Acerca de» del editor en Degú dice la licencia y enlaza a este código.
