"""Arma el editor de documentos SIN CONEXIÓN que viaja dentro de Degú Desktop (`oficina_local.py`).

Es el MISMO editor del servidor (`apps/editor`): Euro-Office 9.3.7, el commit de `apps/editor/FORK`,
sacado de su paquete publicado (`fuentes.json`, con su huella). De ahí sale solo lo que corre en
la ventana: `sdkjs` y `web-apps` de documentos y planillas, los diccionarios de español e inglés
y las fuentes. Lo único que no es de Euro-Office es el conversor de .docx/.xlsx compilado para
correr en la ventana (x2t en WebAssembly, el de CryptPad, también con su huella) y nuestra capa
(`sin_conexion.js`, `editor.html`).

Las fuentes se generan con `allfontsgen`, la herramienta del mismo paquete, y es un binario de
Linux: por eso esto corre en Linux (el CI, `instaladores.yml`) y deja un .zip que los instaladores
de Mac y Windows solo desempacan.

    python3 armar.py --zip editor-local.zip      # en Linux
"""
from __future__ import annotations

import argparse
import hashlib
import json
import lzma
import os
import shutil
import subprocess
import sys
import tarfile
import tempfile
import time
import urllib.request
import zipfile
from pathlib import Path

AQUI = Path(__file__).resolve().parent
RAIZ_DEB = "./var/www/euro-office/documentserver/"
IDIOMAS = ("es", "en")
DICCIONARIOS = ("es_ES", "en_US")
# Las familias para oficinas en español, inglés y portugués, con las equivalentes métricas de
# Arial, Times, Courier, Calibri y Cambria. Sin las CJK, que son la mitad de los 157 MB.
FAMILIAS = ("liberation", "crosextra", "caladea", "dejavu", "opensans", "ubuntu-font-family",
            "asana", "openoffice", "droid")
EDITORES = ("documenteditor", "spreadsheeteditor")


def bajar(url: str, cache: Path, algoritmo: str, huella: str) -> Path:
    """Baja `url` una vez a `cache` y comprueba su huella; una huella distinta corta todo."""
    destino = cache / f"{huella[:16]}-{Path(url).name}"
    if not destino.exists():
        cache.mkdir(parents=True, exist_ok=True)
        tmp = destino.with_suffix(".parcial")
        with urllib.request.urlopen(url, timeout=600) as r, open(tmp, "wb") as f:
            shutil.copyfileobj(r, f, 1 << 20)
        tmp.rename(destino)
    h = hashlib.new(algoritmo)
    with open(destino, "rb") as f:
        for trozo in iter(lambda: f.read(1 << 20), b""):
            h.update(trozo)
    if h.hexdigest() != huella:
        destino.unlink()
        sys.exit(f"la huella de {url} no calza: {h.hexdigest()}")
    return destino


def datos_del_deb(deb: Path, destino: Path) -> Path:
    """Copia el `data.tar.xz` de un .deb (un archivo `ar`) a `destino`, sin `ar` ni `dpkg`."""
    with open(deb, "rb") as f:
        if f.read(8) != b"!<arch>\n":
            sys.exit(f"{deb} no es un .deb")
        while cabecera := f.read(60):
            nombre = cabecera[:16].decode().strip().rstrip("/")
            largo = int(cabecera[48:58])
            if nombre.startswith("data.tar"):
                if not nombre.endswith(".xz"):
                    sys.exit(f"compresión no prevista: {nombre}")
                with open(destino, "wb") as out:
                    while largo:
                        trozo = f.read(min(largo, 1 << 20))
                        out.write(trozo)
                        largo -= len(trozo)
                return destino
            f.seek(largo + (largo % 2), 1)
    sys.exit("el .deb no trae data.tar")


def se_queda(ruta: str) -> bool:
    """Lo que el editor necesita en la ventana, más lo que hace falta para generar las fuentes."""
    partes = ruta.split("/")
    cima = partes[0]
    # `LICENSE` no: es el mismo texto que `LICENSE.txt` y en Mac choca con la carpeta `license`.
    if cima in ("LICENSE.txt", "3rd-Party.txt", "license"):
        return True
    if cima == "sdkjs":
        return len(partes) > 1 and partes[1] in ("common", "word", "cell", "vendor")
    if cima == "web-apps":
        if partes[1:2] == ["vendor"]:
            return True
        if partes[1:2] != ["apps"] or len(partes) < 3:
            return False
        app = partes[2]
        if app in ("api", "common"):
            return "mobile" not in partes
        if app not in EDITORES or partes[3:4] != ["main"]:
            return False
        if "help" in partes[4:6]:
            return False
        if "locale" in partes[4:5] or "formula-lang" in partes[4:6]:
            hoja = partes[-1].split(".")[0].split("_")[0]
            return hoja in IDIOMAS
        return True
    if cima == "dictionaries":
        return len(partes) > 1 and partes[1] in DICCIONARIOS
    if cima == "core-fonts":
        return len(partes) > 1 and partes[1] in FAMILIAS
    # Solo para generar las fuentes: no viajan en el paquete.
    if cima == "server":
        return ruta.startswith(("server/tools/allfontsgen", "server/FileConverter/bin/"))
    return False


def extraer(deb: Path, destino: Path) -> None:
    xz = datos_del_deb(deb, destino / "data.tar.xz")
    # En flujo (`r|`): el tar descomprimido pasa de 2 GB y no tiene por qué caber en memoria.
    with lzma.open(xz) as datos, tarfile.open(fileobj=datos, mode="r|") as tar:
        for m in tar:
            # Los enlaces simbólicos también: las bibliotecas de `allfontsgen` vienen así.
            if not m.name.startswith(RAIZ_DEB) or not (m.isfile() or m.isdir() or m.issym()):
                continue
            ruta = m.name[len(RAIZ_DEB):]
            if not ruta or not se_queda(ruta):
                continue
            m.name = ruta
            tar.extract(m, destino, filter="data")
    xz.unlink()


def generar_fuentes(trabajo: Path) -> None:
    """Lo mismo que `documentserver-generate-allfonts.sh`, sin las fuentes del sistema."""
    bin_ = trabajo / "server/FileConverter/bin"
    entorno = {**os.environ, "LD_LIBRARY_PATH": str(bin_)}
    (trabajo / "fonts").mkdir(exist_ok=True)
    subprocess.run([
        str(trabajo / "server/tools/allfontsgen"),
        f"--input={trabajo / 'core-fonts'}",
        f"--allfonts-web={trabajo / 'sdkjs/common/AllFonts.js'}",
        f"--allfonts={bin_ / 'AllFonts.js'}",
        f"--images={trabajo / 'sdkjs/common/Images'}",
        f"--selection={bin_ / 'font_selection.bin'}",
        f"--output-web={trabajo / 'fonts'}",
        "--use-system=false",
        "--use-system-user-fonts=false",
    ], check=True, env=entorno)
    if not (trabajo / "sdkjs/common/AllFonts.js").is_file():
        sys.exit("allfontsgen no dejó AllFonts.js")


def armar(salida: Path, cache: Path) -> dict:
    fuentes = json.loads((AQUI / "fuentes.json").read_text())
    eo, x2t = fuentes["euro_office"], fuentes["x2t"]
    deb = bajar(eo["url"], cache, "sha256", eo["sha256"])
    zip_x2t = bajar(x2t["url"], cache, "sha512", x2t["sha512"])

    with tempfile.TemporaryDirectory() as t:
        trabajo = Path(t)
        extraer(deb, trabajo)
        generar_fuentes(trabajo)
        if salida.exists():
            shutil.rmtree(salida)
        salida.mkdir(parents=True)
        for cosa in ("sdkjs", "web-apps", "dictionaries", "fonts", "license", "LICENSE.txt",
                     "3rd-Party.txt"):
            origen = trabajo / cosa
            if origen.is_dir():
                shutil.copytree(origen, salida / cosa)
            elif origen.is_file():
                shutil.copy2(origen, salida / cosa)
    # `api.js` viene como plantilla: el servidor le pone su versión en `{{HASH_POSTFIX}}` al arrancar.
    # Sin completar, el editor no reescribe sus rutas con esa versión, que es lo que hace falta acá.
    api = salida / "web-apps/apps/api/documents"
    shutil.copy2(api / "api.js.tpl", api / "api.js")
    with zipfile.ZipFile(zip_x2t) as z:
        for nombre in ("x2t.js", "x2t.wasm"):
            (salida / "sdkjs/common/x2t").mkdir(parents=True, exist_ok=True)
            (salida / "sdkjs/common/x2t" / nombre).write_bytes(z.read(nombre))
    degu = salida / "degu"
    degu.mkdir()
    for nombre in ("sin_conexion.js", "editor.html", "biblioteca.html"):
        shutil.copy2(AQUI / nombre, degu / nombre)
    version = {"euro_office": eo["version"], "commit": eo["commit"], "x2t": x2t["version"],
               "armado": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())}
    (salida / "VERSION.json").write_text(json.dumps(version, indent=2))
    return version


def comprimir(carpeta: Path, archivo: Path) -> None:
    with zipfile.ZipFile(archivo, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for ruta in sorted(carpeta.rglob("*")):
            if ruta.is_file():
                z.write(ruta, ruta.relative_to(carpeta).as_posix())


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--salida", type=Path, default=AQUI / "paquete")
    p.add_argument("--zip", type=Path, help="además, el paquete comprimido (lo que suben los instaladores)")
    p.add_argument("--cache", type=Path, default=Path.home() / ".cache/degu-editor-local")
    a = p.parse_args()
    if sys.platform != "linux":
        sys.exit("allfontsgen es de Linux: esto corre en el CI (instaladores.yml)")
    version = armar(a.salida, a.cache)
    peso = sum(f.stat().st_size for f in a.salida.rglob("*") if f.is_file())
    print(f"paquete en {a.salida}: {peso / 1048576:.1f} MB · {version}")
    if a.zip:
        comprimir(a.salida, a.zip)
        print(f"comprimido en {a.zip}: {a.zip.stat().st_size / 1048576:.1f} MB")


if __name__ == "__main__":
    main()
