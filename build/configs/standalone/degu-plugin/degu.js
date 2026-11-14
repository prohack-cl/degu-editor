/*
 * Degú al lado del documento (Degú, AGPLv3 como el resto del editor).
 *
 * Un plugin de sistema: el editor lo arranca solo, en un marco oculto, con cada Word o Excel. Hace
 * dos cosas, y solo cuando se las pide la app de Degú que muestra el editor:
 *   - `leer`: el documento ABIERTO, también lo que se acaba de escribir y todavía no se guarda;
 *   - `aplicar`: las operaciones que devolvió Degú, todas en un solo paso del editor. Las ven al
 *     tiro quienes editan a la vez, y se deshacen con Ctrl+Z como cualquier cambio a mano.
 *
 * Habla solo con la app de su mismo nombre sin `editor.` (como el nginx de `degu/`): revisa el
 * origen de lo que recibe y no le contesta a nadie más.
 *
 * Las funciones `leerWord`, `aplicarWord`, `leerCell` y `aplicarCell` corren DENTRO del editor
 * (`Asc.plugin.callCommand` las manda como texto): no pueden usar nada de afuera salvo `Api` y
 * `Asc.scope`.
 */
(function (window) {
	'use strict';

	var APP = window.location.protocol + '//' + window.location.host.replace(/^editor\./, '');

	function leerWord() {
		var doc = Api.GetDocument();
		// El texto SIN el número o la viñeta de una lista, que los dibuja el editor. Con ellos, Degú los
		// devolvía como texto y la lista quedaba «1. 1. Revisar…» (dev, 10-nov-2026).
		function sinNumero(p) {
			return p.GetText({ Numbering: false });
		}
		var n = Math.min(doc.GetElementsCount(), 1500);
		var elementos = [];
		var total = 0;
		// Las celdas, con sus números: las tablas son de solo lectura, y así Degú sabe cuál es el punto 7.
		function textoDe(contenido) {
			var partes = [];
			for (var k = 0; k < contenido.GetElementsCount(); k++) {
				var e = contenido.GetElement(k);
				if (e.GetClassType() === 'paragraph') partes.push(e.GetText());
			}
			return partes.join('\n');
		}
		for (var i = 0; i < n && total < 120000; i++) {
			var el = doc.GetElement(i);
			var clase = el.GetClassType();
			if (clase === 'paragraph') {
				var texto = sinNumero(el).replace(/\r$/, '');
				var estilo = el.GetStyle();
				// El largo real: uno más largo que lo que se lee no se reescribe entero (`editor_degu`).
				var p = { clase: 'parrafo', texto: texto.slice(0, 4000), largo: texto.length, estilo: estilo ? estilo.GetName() : '' };
				// El número o la viñeta de la lista va aparte: así Degú sabe que está y no lo escribe.
				var conNumero = el.GetText().replace(/\r$/, '');
				if (conNumero !== texto && conNumero.slice(conNumero.length - texto.length) === texto) {
					var numero = conNumero.slice(0, conNumero.length - texto.length).trim();
					if (numero) p.numero = numero.slice(0, 60).trim();
				}
				elementos.push(p);
				total += texto.length;
			} else if (clase === 'table') {
				var filas = [];
				for (var r = 0; r < Math.min(el.GetRowsCount(), 300) && total < 120000; r++) {
					var fila = el.GetRow(r);
					var celdas = [];
					for (var c = 0; c < Math.min(fila.GetCellsCount(), 40); c++) {
						var t = textoDe(fila.GetCell(c).GetContent()).slice(0, 300);
						total += t.length;
						celdas.push(t);
					}
					filas.push(celdas);
				}
				elementos.push({ clase: 'tabla', filas: filas });
			} else {
				// Otra cosa (un control de contenido, un índice): ocupa su número y no se toca.
				elementos.push({ clase: 'otro' });
			}
		}
		var rango = doc.GetRangeBySelect();
		var seleccion = rango ? rango.GetText() : '';
		return JSON.stringify({ tipo: 'word', elementos: elementos, seleccion: (seleccion || '').slice(0, 4000) });
	}

	function aplicarWord() {
		var doc = Api.GetDocument();
		// Sin el número de la lista, igual que al leer (`leerWord`).
		function sinNumero(p) {
			return p.GetText({ Numbering: false });
		}
		// El número o la viñeta, como al leer: un ítem con el mismo texto y otro número no es el que leyó
		// Degú (panel). Dos idénticos en texto y número no se distinguen, igual que dos párrafos iguales.
		function numeroDe(p) {
			var texto = sinNumero(p).replace(/\r$/, '');
			var con = p.GetText().replace(/\r$/, '');
			return con !== texto && con.slice(con.length - texto.length) === texto ? con.slice(0, con.length - texto.length).trim().slice(0, 60).trim() : '';
		}
		var estilos = { normal: 'Normal', titulo1: 'Heading 1', titulo2: 'Heading 2', titulo3: 'Heading 3', lista: 'List Paragraph' };
		// Los números son los de la lectura: se aplican de abajo hacia arriba, así lo que se agrega
		// o se borra no corre los de más arriba.
		// «Reemplazar en todo» va al final: antes cambiaría párrafos que los otros cambios del mismo
		// pedido esperan como Degú los leyó.
		function lugar(o) {
			if (o.op === 'insertar_parrafo') return o.despues_de + 0.5;
			return typeof o.indice === 'number' ? o.indice : -Infinity;
		}
		var ops = Asc.scope.ops.map(function (o, k) { return { o: o, k: k }; });
		// Dos párrafos nuevos en el mismo lugar: el último va primero, así quedan en el orden pedido.
		ops.sort(function (a, b) {
			return lugar(b.o) - lugar(a.o) || (a.o.op === 'insertar_parrafo' ? b.k - a.k : a.k - b.k);
		});
		// Lo que tenía cada párrafo cuando Degú lo leyó: si cambió mientras pensaba (alguien escribió
		// arriba y los números se corrieron), ese cambio se salta en vez de caer en otro párrafo.
		function igual(el, texto, numero) {
			return !!el && el.GetClassType() === 'paragraph' && sinNumero(el).replace(/\r$/, '') === texto && numeroDe(el) === (numero || '');
		}
		// El elemento junto al que va un párrafo nuevo, como lo leyó Degú (`editor_degu._huella`).
		function coincide(el, h) {
			if (!el || !h) return !h;
			var clase = el.GetClassType();
			if (h.clase === 'parrafo') {
				if (clase !== 'paragraph') return false;
				var t = sinNumero(el).replace(/\r$/, '');
				if (numeroDe(el) !== (h.numero || '')) return false;
				return h.entero === false ? t.slice(0, 4000) === h.texto : t === h.texto;
			}
			if (h.clase === 'tabla') {
				if (clase !== 'table') return false;
				var celda = el.GetRow(0).GetCell(0).GetContent().GetElement(0);
				return !h.primera || (celda && celda.GetText && celda.GetText().slice(0, 300) === h.primera);
			}
			return clase !== 'paragraph' && clase !== 'table';
		}
		var saltadas = 0;
		function llenar(p, texto) {
			var lineas = String(texto).split('\n');
			for (var j = 0; j < lineas.length; j++) {
				if (j) p.AddLineBreak();
				p.AddText(lineas[j]);
			}
		}
		var hechas = 0;
		// Todo se verifica contra el documento ANTES de cambiar nada: si no, un párrafo nuevo al
		// inicio haría fallar la verificación del siguiente que va en el mismo lugar.
		function vale(o) {
			if (o.op === 'reemplazar_parrafo' || o.op === 'borrar_parrafo') return igual(doc.GetElement(o.indice), o.antes, o.numero);
			if (o.op === 'insertar_parrafo')
				return (!o.ancla || coincide(doc.GetElement(o.despues_de), o.ancla)) && (!o.primero || coincide(doc.GetElement(0), o.primero));
			return o.op === 'reemplazar_en_todo';
		}
		var validas = ops.map(function (x) {
			try {
				return vale(x.o);
			} catch (e) {
				return false;
			}
		});
		// Cada cambio se comprueba releyendo el documento: solo cuenta como hecho lo que quedó. El texto se
		// compara tal cual, salvo cómo marca el editor los saltos de línea y el fin del párrafo.
		function plano(t) {
			return String(t == null ? '' : t).replace(/\r\n?|\u000b/g, '\n').replace(/\s+$/, '');
		}
		function quedo(el, texto) {
			return !!el && el.GetClassType() === 'paragraph' && plano(sinNumero(el)) === plano(texto);
		}
		function nuevo(texto, estilo) {
			var p = Api.CreateParagraph();
			llenar(p, texto);
			if (estilo) p.SetStyle(estilo);
			return p;
		}
		// Pone el párrafo `p` en `lugar` y comprueba que entró con su texto; si no, lo quita y avisa.
		function poner(lugar, p, texto) {
			var n = doc.GetElementsCount();
			doc.AddElement(lugar, p);
			// Sin contar, un párrafo igual que ya estaba ahí haría pasar por hecho uno que no entró (panel).
			if (doc.GetElementsCount() !== n + 1) throw new Error('el párrafo nuevo no entró');
			if (!quedo(doc.GetElement(lugar), texto)) {
				doc.RemoveElement(lugar);
				throw new Error('el párrafo nuevo no quedó con su texto');
			}
		}
		// Todo el texto del documento, también el de las tablas: para saber si un cambio cambió algo.
		function todo(contenido) {
			var partes = [];
			for (var k = 0; k < contenido.GetElementsCount(); k++) {
				var e = contenido.GetElement(k);
				var clase = e.GetClassType();
				if (clase === 'paragraph') partes.push(e.GetText());
				else if (clase === 'table')
					for (var r = 0; r < e.GetRowsCount(); r++)
						for (var c = 0; c < e.GetRow(r).GetCellsCount(); c++) partes.push(todo(e.GetRow(r).GetCell(c).GetContent()));
			}
			return partes.join('\n');
		}
		function aplicar(o) {
			if (o.op === 'reemplazar_en_todo') {
				if (o.buscar === o.reemplazo) return null;
				// Se mira si el texto del documento cambió: sirve también para «Ana» → «Ana María», donde lo
				// buscado sigue apareciendo (panel).
				var antes = todo(doc);
				doc.SearchAndReplace({ searchString: o.buscar, replaceString: o.reemplazo, matchCase: true });
				if (todo(doc) !== antes) return true;
				// Si no aparecía, no había nada que cambiar: ni hecho ni saltado.
				if (antes.indexOf(o.buscar) < 0) return null;
				throw new Error('el texto buscado sigue en el documento');
			}
			if (o.op === 'borrar_parrafo') {
				var n = doc.GetElementsCount();
				doc.RemoveElement(o.indice);
				if (doc.GetElementsCount() !== n - 1) throw new Error('el párrafo no se borró');
				return true;
			}
			if (o.op === 'reemplazar_parrafo') {
				var el = doc.GetElement(o.indice);
				var cuantos = doc.GetElementsCount();
				// 1. Corregir el texto conservando el formato de cada tramo (negritas, enlaces), como los
				// asistentes de ONLYOFFICE. Reemplaza lo SELECCIONADO: en un párrafo vacío no hay nada que
				// seleccionar, y la primera línea de un Word en blanco no se escribía (dev, 8-nov-2026). Con
				// saltos de línea tampoco: podría partirlo en varios párrafos (panel).
				if (typeof Api.ReplaceTextSmart === 'function' && plano(sinNumero(el)) && o.texto.indexOf('\n') < 0) {
					(el.Select ? el : el.GetRange()).Select();
					Api.ReplaceTextSmart([o.texto], '\t', '\n');
					if (doc.GetElementsCount() !== cuantos) throw new Error('el reemplazo partió el párrafo');
					if (quedo(doc.GetElement(o.indice), o.texto)) return true;
				}
				// 2. Si no: una copia del párrafo (misma numeración, sangría y alineación) con el texto pedido, en
				// su lugar. Es el camino de «insertar», que sí escribe en el editor real; escribir DENTRO del
				// párrafo vacío no dejaba nada. El de antes no se quita hasta que la copia quedó, así que nunca
				// se borra el de abajo (panel).
				el = doc.GetElement(o.indice);
				var copia = typeof el.Copy === 'function' ? el.Copy() : null;
				if (copia) {
					copia.RemoveAllElements();
					llenar(copia, o.texto);
				}
				poner(o.indice, copia || nuevo(o.texto, el.GetStyle()), o.texto);
				doc.RemoveElement(o.indice + 1);
				if (doc.GetElementsCount() !== cuantos) {
					doc.RemoveElement(o.indice);
					throw new Error('el párrafo de antes no se quitó');
				}
				return true;
			}
			if (o.op === 'insertar_parrafo') {
				poner(o.despues_de + 1, nuevo(o.texto, doc.GetStyle(estilos[o.estilo] || 'Normal')), o.texto);
				return true;
			}
			throw new Error('operación desconocida');
		}
		// Saltada: el documento cambió mientras Degú pensaba y no se tocó. Fallida: se intentó y no quedó.
		var fallos = [];
		for (var i = 0; i < ops.length; i++) {
			if (!validas[i]) {
				saltadas++;
				continue;
			}
			try {
				if (aplicar(ops[i].o)) hechas++;
			} catch (e) {
				fallos.push(ops[i].o.op + ': ' + String((e && e.message) || e));
			}
		}
		return JSON.stringify({ hechas: hechas, saltadas: saltadas, fallidas: fallos.length, fallos: fallos });
	}

	function leerCell() {
		function columna(n) {
			var s = '';
			n += 1;
			while (n > 0) {
				var r = (n - 1) % 26;
				s = String.fromCharCode(65 + r) + s;
				n = Math.floor((n - 1) / 26);
			}
			return s;
		}
		var hoja = Api.GetActiveSheet();
		var hojas = Api.GetSheets().map(function (h) { return h.GetName(); });
		var usado = hoja.GetUsedRange();
		var direccion = usado ? usado.GetAddress(false, false, 'xlA1', false) : 'A1';
		var m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(direccion) || [null, 'A', '1'];
		function numero(letras) {
			var n = 0;
			for (var k = 0; k < letras.length; k++) n = n * 26 + (letras.charCodeAt(k) - 64);
			return n - 1;
		}
		var c0 = numero(m[1]);
		var f0 = parseInt(m[2], 10) - 1;
		var c1 = m[3] ? numero(m[3]) : c0;
		var f1 = m[4] ? parseInt(m[4], 10) - 1 : f0;
		// Una celda larga va cortada, con la huella de su contenido entero: así se nota si cambió
		// más allá de lo que se lee (la misma en `leerCell` y `aplicarCell`).
		function corto(t) {
			if (t.length <= 300) return t;
			var h = 0x811c9dc5;
			for (var k = 0; k < t.length; k++) h = Math.imul(h ^ t.charCodeAt(k), 16777619) >>> 0;
			return t.slice(0, 300) + '…#' + ('0000000' + h.toString(16)).slice(-8);
		}
		var celdas = [];
		var total = 0;
		for (var f = f0; f <= Math.min(f1, f0 + 299) && total < 120000; f++) {
			var fila = [];
			for (var c = c0; c <= Math.min(c1, c0 + 39); c++) {
				var celda = hoja.GetRangeByNumber(f, c);
				var formula = celda.GetFormula();
				var valor = celda.GetValue();
				var texto = corto(String(formula && formula.charAt(0) === '=' ? formula : valor == null ? '' : valor));
				total += texto.length;
				fila.push(texto);
			}
			celdas.push(fila);
		}
		var sel = Api.GetSelection();
		return JSON.stringify({
			tipo: 'cell',
			hojas: hojas,
			hoja: hoja.GetName(),
			desde: columna(c0) + (f0 + 1),
			// Todo lo usado, aunque se lea menos: así Degú sabe que hay datos que no vio.
			usado: direccion,
			celdas: celdas,
			seleccion: sel ? sel.GetAddress(false, false, 'xlA1', false) : ''
		});
	}

	function aplicarCell() {
		var activa = Api.GetActiveSheet().GetName();
		// La hoja que nombra el cambio (el motor la pone siempre: la que leyó Degú o una que crea este
		// pedido), o ninguna: nunca otra en su lugar. La activa de ahora no cuenta, porque la persona
		// pudo cambiar de pestaña mientras Degú pensaba. Si otra persona creó primero una hoja que este
		// pedido iba a crear, esa no es de Degú y no se toca.
		var ajenas = {};
		function hoja(nombre) {
			if (!nombre || ajenas[nombre]) return null;
			return Api.GetSheet(nombre);
		}
		var saltadas = 0;
		function numero(letras) {
			var n = 0;
			for (var k = 0; k < letras.length; k++) n = n * 26 + (letras.charCodeAt(k) - 64);
			return n - 1;
		}
		var hechas = 0;
		// Una celda larga va cortada, con la huella de su contenido entero: así se nota si cambió
		// más allá de lo que se lee (la misma en `leerCell` y `aplicarCell`).
		function corto(t) {
			if (t.length <= 300) return t;
			var h = 0x811c9dc5;
			for (var k = 0; k < t.length; k++) h = Math.imul(h ^ t.charCodeAt(k), 16777619) >>> 0;
			return t.slice(0, 300) + '…#' + ('0000000' + h.toString(16)).slice(-8);
		}
		// Lo que muestra una celda, igual que al leerla (`leerCell`).
		function vale(h, f, c) {
			var celda = h.GetRangeByNumber(f, c);
			var formula = celda.GetFormula();
			var valor = celda.GetValue();
			return corto(String(formula && formula.charAt(0) === '=' ? formula : valor == null ? '' : valor));
		}
		// ¿Siguen como Degú las leyó? Si alguien las cambió mientras pensaba, el cambio se salta en
		// vez de pisar lo nuevo.
		function iguales(h, f0, c0, antes) {
			if (!antes) return true;
			for (var f = 0; f < antes.length; f++)
				for (var c = 0; c < antes[f].length; c++)
					if (vale(h, f0 + f, c0 + c) !== antes[f][c]) return false;
			return true;
		}
		function aplicar(o) {
			if (o.op === 'nueva_hoja') {
				if (Api.GetSheet(o.nombre)) {
					ajenas[o.nombre] = true;
					return false;
				}
				Api.AddSheet(o.nombre);
				// Agregar una hoja la deja activa: se vuelve a la que estaba mirando la persona.
				var antes = Api.GetSheet(activa);
				if (antes) antes.SetActive();
				return true;
			}
			var h = hoja(o.hoja);
			if (!h) return false;
			if (o.op === 'escribir') {
				var m = /^\$?([A-Z]+)\$?(\d+)$/.exec(o.desde);
				if (!m) return false;
				var c0 = numero(m[1]);
				var f0 = parseInt(m[2], 10) - 1;
				for (var f = 0; f < o.valores.length; f++)
					for (var c = 0; c < o.valores[f].length; c++)
						h.GetRangeByNumber(f0 + f, c0 + c).SetValue(o.valores[f][c]);
				return true;
			}
			if (o.op === 'limpiar') {
				h.GetRange(o.rango).Clear();
				return true;
			}
			if (o.op === 'formato') {
				var rango = h.GetRange(o.rango);
				if (typeof o.negrita === 'boolean') rango.SetBold(o.negrita);
				if (o.formato_numero) rango.SetNumberFormat(o.formato_numero);
				if (o.relleno) {
					var x = parseInt(o.relleno.slice(1), 16);
					rango.SetFillColor(Api.CreateColorFromRGB((x >> 16) & 255, (x >> 8) & 255, x & 255));
				}
				return true;
			}
			return false;
		}
		// La esquina de arriba a la izquierda de lo que toca un cambio, aunque el rango venga al revés.
		function esquina(o) {
			if (o.op === 'escribir') {
				var m = /^\$?([A-Z]+)\$?(\d+)$/.exec(o.desde);
				return m ? [parseInt(m[2], 10) - 1, numero(m[1])] : null;
			}
			var puntas = String(o.rango).replace(/\$/g, '').split(':').map(function (x) {
				return /^([A-Z]+)(\d+)$/.exec(x);
			});
			if (puntas.some(function (x) { return !x; })) return null;
			return [
				Math.min.apply(null, puntas.map(function (x) { return parseInt(x[2], 10) - 1; })),
				Math.min.apply(null, puntas.map(function (x) { return numero(x[1]); }))
			];
		}
		var ops = Asc.scope.ops;
		// Lo que había en las celdas se compara contra lo que leyó Degú ANTES de cambiar nada: si no,
		// «limpiar A1 y escribir en A1» haría fallar la escritura por la limpieza del mismo pedido.
		var validas = ops.map(function (o) {
			if (!o.antes) return true;
			try {
				var h = hoja(o.hoja);
				var e = esquina(o);
				return !!h && !!e && iguales(h, e[0], e[1], o.antes);
			} catch (x) {
				return false;
			}
		});
		// Saltada: lo que leyó Degú ya no está (otra hoja, otras celdas). Fallida: el editor no lo aceptó,
		// y puede haber quedado a medias, así que no se cuenta como «el documento cambió» (panel, Terra).
		var fallos = [];
		for (var i = 0; i < ops.length; i++) {
			if (!validas[i]) {
				saltadas++;
				continue;
			}
			try {
				if (aplicar(ops[i])) hechas++;
				else saltadas++;
			} catch (e) {
				fallos.push(ops[i].op + ': ' + String((e && e.message) || e));
			}
		}
		return JSON.stringify({ hechas: hechas, saltadas: saltadas, fallidas: fallos.length, fallos: fallos });
	}

	// Lo que se prueba sin el editor (`apps/web/tests/unit/el-plugin-de-degu-en-el-editor…`).
	window.DeguOficina = { leerWord: leerWord, aplicarWord: aplicarWord, leerCell: leerCell, aplicarCell: aplicarCell, APP: APP };

	// ── El plugin ───────────────────────────────────────────────────────────────────────────
	// El editor le pone al marco `iframe_<guid>` y, cuando el plugin dice `initialize`, le manda su
	// API (`plugin_init`) para evaluarla: es como arrancan sus plugins, sin un archivo de afuera.
	if (!window.name || window.name.indexOf('iframe_') !== 0) return;
	window.Asc = window.Asc || {};
	window.Asc.plugin = { guid: window.name.slice(7) };

	var cola = Promise.resolve();
	var tipo = '';

	function contestar(id, datos) {
		window.top.postMessage(Object.assign({ degu: 'oficina', id: id }, datos), APP);
	}

	function ejecutar(fn, ops) {
		return new Promise(function (resolver, fallar) {
			// Con plazo: si el editor no contesta, la cola no se queda trabada para siempre. Largo a
			// propósito: antes de escribir, el editor baja las fuentes del texto («Cargando datos»), y con una
			// conexión lenta eso pasaba de los 50 s de antes; rendirse ahí decía «no se aplicó» y el texto
			// aparecía después (dev, 8-nov-2026). Lo que importa es no quedar trabado, no apurar al editor.
			var plazo = window.setTimeout(function () {
				fallar(new Error('el editor no contestó'));
			}, 180000);
			window.Asc.scope = { ops: ops || [] };
			window.Asc.plugin.callCommand(fn, false, true, function (r) {
				window.clearTimeout(plazo);
				resolver(typeof r === 'string' ? JSON.parse(r) : r);
			});
		});
	}

	function atender(e) {
		var d = e.data;
		if (typeof d === 'string') {
			// El arranque y la API del plugin: solo del editor que lo contiene, nunca de otra ventana.
			if (window.parent === window || e.source !== window.parent || e.origin !== window.location.origin) return;
			var m;
			try {
				m = JSON.parse(d);
			} catch (x) {
				return;
			}
			if (m && m.type === 'plugin_init') {
				// eslint-disable-next-line no-eval
				eval(m.data);
				return;
			}
			if (window.plugin_onMessage) window.plugin_onMessage(e);
			return;
		}
		if (e.origin !== APP || !d || d.degu !== 'oficina' || !d.id) return;
		// ¿Ya arrancó? La app pregunta hasta que el editor le pasó al plugin su API.
		if (d.accion === 'hola') return contestar(d.id, { ok: !!tipo, tipo: tipo });
		// De a un pedido por vez: el editor contesta los comandos por un solo canal.
		cola = cola.then(function () {
			if (d.accion === 'leer')
				return ejecutar(tipo === 'cell' ? leerCell : leerWord).then(function (doc) {
					contestar(d.id, { ok: true, documento: doc });
				});
			if (d.accion === 'aplicar')
				return ejecutar(tipo === 'cell' ? aplicarCell : aplicarWord, d.operaciones).then(function (r) {
					// Sin la cuenta del editor (el comando falló entero), nada se da por hecho.
					if (!r) return contestar(d.id, { ok: false, error: 'el editor no aplicó los cambios' });
					contestar(d.id, { ok: true, hechas: r.hechas, saltadas: r.saltadas, fallidas: r.fallidas || 0, fallos: r.fallos || [] });
				});
			contestar(d.id, { ok: false, error: 'accion' });
		}).catch(function (x) {
			contestar(d.id, { ok: false, error: String(x && x.message || x) });
		});
	}

	window.Asc.plugin.init = function () {
		tipo = (window.Asc.plugin.info && window.Asc.plugin.info.editorType) === 'cell' ? 'cell' : 'word';
	};
	window.Asc.plugin.button = function () {};

	window.addEventListener('message', atender, false);
	window.parent.postMessage(JSON.stringify({ type: 'initialize', guid: window.Asc.plugin.guid }), '*');
})(window);
