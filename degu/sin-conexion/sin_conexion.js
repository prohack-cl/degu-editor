/*
 * El servidor del editor, DENTRO de la ventana: así Euro-Office abre y guarda sin internet.
 *
 * Degú Desktop sirve el editor desde el equipo (`oficina_local.py`) y mete este archivo al
 * comienzo de la página del editor (`web-apps/apps/<editor>/main/index.html`). No se toca ni una
 * línea de Euro-Office: el editor habla con su servidor por dos puertas, y las dos se atienden acá.
 *
 *  1. La conexión. El editor la pide a `window.io` si existe (`AscCommon.getSocketIO` en sdkjs),
 *     así que basta con poner aquí un `io` que contesta como el servidor (`DocsCoServer.js` de
 *     Euro-Office 9.3.7): la licencia, la autorización y el documento abierto.
 *  2. Lo que va por HTTP: guardar (`../../../../downloadas/<id>`, el documento entero en binario)
 *     e insertar imágenes (`../../../../upload/<id>`). Se atienden en `XMLHttpRequest`.
 *
 * Convertir .docx/.xlsx ↔ binario del editor lo hace la página de afuera (`editor.html`), con el
 * conversor de Euro-Office en WebAssembly: `parent.deguLocal`. Una sola persona edita, así que no
 * hay bloqueos que repartir: todo se concede.
 */
(function () {
	'use strict';
	var degu;
	try {
		degu = window.parent && window.parent !== window && window.parent.deguLocal;
	} catch (e) {
		degu = null;
	}
	if (!degu) return;

	var VERSION = '9.3.7';
	var sesion = 'local-' + Date.now().toString(36);

	// ── 1. La conexión ─────────────────────────────────────────────────────────────────────────

	function Conexion(opciones) {
		this.opciones = opciones || {};
		this.oyentes = {};
		this.connected = false;
		var yo = this;
		this.io = {
			opts: {},
			on: function () {},
			reconnectionAttempts: function () {},
			reconnectionDelay: function () {},
			reconnectionDelayMax: function () {}
		};
		// Como el servidor: al conectar manda la licencia, y recién con ella el editor pide su
		// autorización (`auth`, con la orden de abrir). Conecta después de que el editor puso sus
		// oyentes, como un socket de verdad.
		setTimeout(function () {
			yo.connected = true;
			yo._emitir('connect');
			yo._licencia();
		}, 0);
	}
	Conexion.prototype.on = function (evento, fn) {
		(this.oyentes[evento] = this.oyentes[evento] || []).push(fn);
		return this;
	};
	Conexion.prototype.off = function () {
		return this;
	};
	Conexion.prototype._emitir = function (evento, datos) {
		(this.oyentes[evento] || []).forEach(function (fn) {
			try {
				fn(datos);
			} catch (e) {
				console.error('[degu] el editor falló al recibir', evento, e);
			}
		});
	};
	Conexion.prototype._mensaje = function (datos) {
		var yo = this;
		setTimeout(function () {
			yo._emitir('message', datos);
		}, 0);
	};
	Conexion.prototype.disconnect = Conexion.prototype.close = function () {
		this.connected = false;
		return this;
	};
	Conexion.prototype.emit = function (evento, datos) {
		if (evento === 'message' && datos) this._recibir(datos);
		return this;
	};

	Conexion.prototype._licencia = function () {
		this._mensaje({
			type: 'license',
			license: {
				type: 3, // c_oLicenseResult.Success
				light: false,
				mode: 0,
				rights: degu.soloLectura ? 4 : 1, // RIGHTS del servidor: 1 = editar, 4 = ver
				buildVersion: VERSION,
				buildNumber: 0,
				branding: false,
				customization: false,
				advancedApi: true
			}
		});
	};

	Conexion.prototype._autorizar = function (auth) {
		if (this.autorizada) {
			this._mensaje({ type: 'auth', result: 1, sessionId: sesion, participants: [], locks: [] });
			return;
		}
		this.autorizada = true;
		var usuario = auth.user || {};
		this.usuario = usuario.id || 'local';
		this.indice = 1;
		var participante = {
			id: this.usuario + this.indice,
			idOriginal: this.usuario,
			username: usuario.username || '',
			indexUser: this.indice,
			view: degu.soloLectura,
			connectionId: sesion,
			isCloseCoAuthoring: false
		};
		this._mensaje({
			type: 'auth',
			result: 1,
			sessionId: sesion,
			sessionTimeConnect: Date.now(),
			participants: [participante],
			locks: [],
			indexUser: this.indice,
			buildVersion: VERSION,
			buildNumber: 0,
			licenseType: 3,
			settings: {}
		});
		// La orden de abrir viaja con la autorización (`openCmd`), como en el servidor.
		if (auth.openCmd) this._abrir(auth.openCmd);
	};

	Conexion.prototype._abrir = function () {
		var yo = this;
		degu
			.abrir()
			.then(function (urls) {
				yo._mensaje({ type: 'documentOpen', data: { type: 'open', status: 'ok', data: urls } });
			})
			.catch(function (e) {
				console.error('[degu] no se pudo abrir', e);
				degu.avisar('No se pudo abrir el documento: ' + (e && e.message ? e.message : e));
				yo._mensaje({ type: 'documentOpen', data: { type: 'open', status: 'err', data: -80 } });
			});
	};

	// Lo que el editor pide por el canal de abrir, según su orden (`c`), como `canvasservice.js`.
	Conexion.prototype._orden = function (orden) {
		var yo = this;
		var c = orden.c;
		if (c === 'open' || c === 'reopen' || !c) return this._abrir(orden);
		if (c === 'imgurls') {
			// Una imagen pegada o insertada por dirección: queda entre los medios del documento.
			var urls = [].concat(orden.data || []);
			Promise.all(
				urls.map(function (u) {
					return fetch(u)
						.then(function (r) {
							if (!r.ok) throw new Error(String(r.status));
							return r.blob();
						})
						.then(function (b) {
							return b.arrayBuffer().then(function (buf) {
								var medio = degu.agregarImagen(new Uint8Array(buf), b.type);
								var ruta = Object.keys(medio)[0];
								return { path: ruta, url: medio[ruta] };
							});
						})
						.catch(function () {
							// Sin internet, una imagen de otro sitio no se puede traer.
							return { path: 'error', url: 'error' };
						});
				})
			).then(function (salida) {
				yo._mensaje({
					type: 'documentOpen',
					data: { type: 'imgurls', status: 'ok', data: { error: 0, urls: salida } }
				});
			});
			return;
		}
		if (c === 'pathurl' || c === 'pathurls') {
			var rutas = [].concat(orden.data || []).map(function (r) {
				return degu.urlDeMedio(r);
			});
			this._mensaje({
				type: 'documentOpen',
				data: { type: c, status: 'ok', data: c === 'pathurl' ? rutas[0] : rutas }
			});
			return;
		}
		console.info('[degu] orden del editor sin atender:', c);
	};

	Conexion.prototype._recibir = function (datos) {
		switch (datos.type) {
			case 'auth':
				this._autorizar(datos);
				break;
			case 'openDocument':
				this._orden(datos.message || {});
				break;
			case 'isSaveLock':
				this._mensaje({ type: 'saveLock', saveLock: false });
				break;
			case 'saveChanges':
				// Los cambios no se guardan acá: guardar es el documento entero (`downloadas`).
				// Solo se confirman, para que el editor no espere a un servidor que no existe.
				if (datos.endSaveChanges) {
					this._mensaje({ type: 'unSaveLock', index: -1, time: Date.now(), syncChangesIndex: -1 });
				} else {
					this._mensaje({ type: 'savePartChanges', changesIndex: -1, syncChangesIndex: -1 });
				}
				break;
			case 'getLock':
				var bloqueos = {};
				var yo = this;
				(datos.block || []).forEach(function (b) {
					bloqueos[(b && b.guid) || b] = { time: Date.now(), user: yo.usuario + yo.indice, block: b };
				});
				this._mensaje({ type: 'getLock', locks: bloqueos });
				break;
			case 'authChangesAck':
			case 'cursor':
			case 'clientLog':
			case 'extendSession':
			case 'unLockDocument':
			case 'close':
			case 'getMessages':
			case 'message':
				break;
			default:
				console.info('[degu] mensaje del editor sin atender:', datos.type);
		}
	};

	window.io = function (opciones) {
		return (degu.conexion = new Conexion(opciones));
	};

	// ── 2. Lo que va por HTTP: guardar e insertar imágenes ────────────────────────────────────────

	var XHR = window.XMLHttpRequest;
	var partes = {};

	function respuestaLocal(xhr, estado, cuerpo) {
		Object.defineProperty(xhr, 'readyState', { value: 4, configurable: true });
		Object.defineProperty(xhr, 'status', { value: estado, configurable: true });
		Object.defineProperty(xhr, 'responseText', { value: cuerpo, configurable: true });
		Object.defineProperty(xhr, 'response', { value: cuerpo, configurable: true });
		setTimeout(function () {
			if (xhr.onreadystatechange) xhr.onreadystatechange();
			if (xhr.onload) xhr.onload();
		}, 0);
	}

	function bytesDe(cuerpo) {
		if (cuerpo instanceof Blob) {
			return cuerpo.arrayBuffer().then(function (b) {
				return new Uint8Array(b);
			});
		}
		if (cuerpo instanceof ArrayBuffer) return Promise.resolve(new Uint8Array(cuerpo));
		if (ArrayBuffer.isView(cuerpo)) {
			return Promise.resolve(new Uint8Array(cuerpo.buffer, cuerpo.byteOffset, cuerpo.byteLength));
		}
		return Promise.resolve(new TextEncoder().encode(String(cuerpo || '')));
	}

	// `savetype` de sdkjs: 0 = primera parte, 1 = una del medio, 2 = la última, 3 = todo de una vez.
	function guardarParte(xhr, url, cuerpo) {
		var cmd = {};
		try {
			cmd = JSON.parse(decodeURIComponent((/[?&]cmd=([^&]*)/.exec(url) || [])[1] || '%7B%7D'));
		} catch (e) {
			cmd = {};
		}
		var llave = cmd.savekey || cmd.saveKey || 'guardar';
		bytesDe(cuerpo).then(function (bytes) {
			var tipo = Number(cmd.savetype);
			if (tipo === 0 || tipo === 3 || !partes[llave]) partes[llave] = [];
			partes[llave].push(bytes);
			respuestaLocal(xhr, 200, JSON.stringify({ type: 'save', status: 'ok', data: llave }));
			if (tipo !== 2 && tipo !== 3) return;
			var todo = partes[llave];
			delete partes[llave];
			degu
				.guardarBinario(todo)
				.then(function (url) {
					degu.conexion._mensaje({
						type: 'documentOpen',
						data: { type: 'save', status: 'ok', data: url, filetype: cmd.outputformat }
					});
				})
				.catch(function (e) {
					console.error('[degu] no se pudo guardar', e);
					degu.conexion._mensaje({
						type: 'documentOpen',
						data: { type: 'save', status: 'err', data: -81 }
					});
				});
		});
	}

	function subirImagen(xhr, cuerpo, tipo) {
		bytesDe(cuerpo).then(function (bytes) {
			respuestaLocal(xhr, 200, JSON.stringify(degu.agregarImagen(bytes, tipo)));
		});
	}

	function XhrLocal() {
		var xhr = new XHR();
		var abrir = xhr.open;
		var enviar = xhr.send;
		var cabecera = xhr.setRequestHeader;
		var ruta = '';
		var tipo = '';
		xhr.open = function (metodo, url) {
			ruta = String(url || '');
			return abrir.apply(xhr, arguments);
		};
		xhr.setRequestHeader = function (nombre, valor) {
			if (String(nombre).toLowerCase() === 'content-type') tipo = valor;
			return cabecera.apply(xhr, arguments);
		};
		xhr.send = function (cuerpo) {
			if (/\/downloadas\//.test(ruta)) return guardarParte(xhr, ruta, cuerpo);
			if (/\/upload\//.test(ruta)) return subirImagen(xhr, cuerpo, tipo);
			return enviar.apply(xhr, arguments);
		};
		return xhr;
	}
	XhrLocal.prototype = XHR.prototype;
	['UNSENT', 'OPENED', 'HEADERS_RECEIVED', 'LOADING', 'DONE'].forEach(function (k) {
		XhrLocal[k] = XHR[k];
	});
	window.XMLHttpRequest = XhrLocal;

	// Ctrl+S / Cmd+S guarda en el equipo, como el botón «Guardar» de arriba. Sin esto, el editor lo
	// toma para su propio guardado de coautoría, que acá no escribe nada.
	window.addEventListener(
		'keydown',
		function (e) {
			if ((e.ctrlKey || e.metaKey) && !e.altKey && String(e.key).toLowerCase() === 's') {
				e.preventDefault();
				e.stopImmediatePropagation();
				degu.guardar();
			}
		},
		true
	);
})();
