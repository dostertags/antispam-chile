"""E2E con Chromium real (headless, navegador incluido en Playwright) contra un portal FALSO que
imita la estructura REAL grabada el 2026-10-23 (WALK.md):

- login del portal con #btnClaveUnica y el badge reCAPTCHA v3 (iframe size=invisible de 256x60);
- ClaveÚnica con #uname/#pword y #login-submit deshabilitado hasta un keyup en ambos campos;
- home con button[onclick='logoutCU()'];
- página No Molestar con #aceptoNoMolestarCheck, modal #sessionTimeoutModal e
  <iframe id="frameNoMolestar" src="...?p=<JWT>">; dentro, la app Rails: listado con
  «Ingresar nuevo…», formulario (radios, #tipo-telefono, #telefono-sub con teléfono precargado,
  #mail-pre/#mail-sub, #ingresar-bloqueo), un select2 parecido al real (#empresas oculto,
  .select2-selection, buscador AJAX con espera de 250 ms, opciones que se eligen con mouseup),
  #ingresar-empresa, #fake-continuar (GET configuración) -> «Revisar solicitud» (canales marcados,
  #pills-telefonos button#confirmar-5, pestaña correo) -> #finalizar (POST) -> listado con N°.

Todo se sirve con context.route("https://portal.test/**"); cualquier otra URL se aborta (sin red).
Las URLs del portal en no_molestar se reemplazan por las de portal.test con monkeypatch.
"""
from __future__ import annotations

import html
import json
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

import pytest

import no_molestar as nm
from conftest import (
    CORREO_PERSONA,
    ENTRADAS_CLARO,
    JWT_PORTAL,
    SECRETO_CU,
    PreguntarGuionado,
    SleepGrabador,
    buscar_en_directorio,
    fila_real,
    tiempos_prueba,
)

pytestmark = pytest.mark.e2e

BASE = "https://portal.test"
URL_PORTAL_E2E = f"{BASE}/app/consumidor/"
URL_HOME_E2E = f"{BASE}/app/consumidor/index.php?c=home&a=home"
URL_NM_E2E = f"{BASE}/app/consumidor/index.php?c=home&a=nomolestar"
URL_CU_E2E = f"{BASE}/claveunica/accounts/login/?next=%2Fopenid%2Fauthorize%3Fclient_id%3Dsernac"
URL_BADGE_E2E = f"{BASE}/recaptcha/api2/anchor?ar=1&k=6LdVkKAaPRUEBA&co=aHR0cHM6&hl=es&v=x&size=invisible&cb=abc"
URL_DESAFIO_E2E = f"{BASE}/recaptcha/api2/anchor?ar=1&k=prueba&co=aHR0cHM6&hl=es&size=normal"
MARCADOR_SCRIPT = "MARCADOR_SCRIPT_E2E_7f3a"
CLAVE_INCORRECTA = "clave-incorrecta"
TEL_PRECARGADO = "+56987654321"

BADGE = (
    '<div class="grecaptcha-badge" data-style="bottomright" style="position:fixed;right:0;bottom:14px;'
    'width:256px;height:60px"><div class="grecaptcha-logo"><iframe title="reCAPTCHA" width="256" height="60" '
    f'src="{URL_BADGE_E2E}" frameborder="0" scrolling="no"></iframe></div></div>'
)

HTML_LOGIN = """<!doctype html>
<html lang="es"><head><meta charset="utf-8"><title>Portal del Consumidor</title></head>
<body>
  <h1>Portal del Consumidor</h1>
  <form id="formLogin" onsubmit="return false;">
    <input id="rutConsumidor"> <input id="passwdConsumidor" type="password"> <input type="hidden" id="recaptcha">
    <button id="btnRegistrado" type="button">Clave SERNAC</button>
  </form>
  <button id="btnClaveUnica" class="boton-clave-unica" type="button"
          onclick="location.href='__URL_CU__'">ClaveÚnica</button>
  __BADGE__
</body></html>
"""

HTML_CU = """<!doctype html>
<html lang="es"><head><meta charset="utf-8"><title>ClaveÚnica</title></head>
<body>
  <h1>Ingresa con tu ClaveÚnica</h1>
  <form id="login-form" data-show-recaptcha="true" onsubmit="return false;">
    <input id="uname" name="run" class="form-control rut" aria-label="Ingresa tu RUN" autocomplete="off">
    <input id="pword" name="password" type="password" aria-label="Ingresa tu ClaveÚnica">
    __EXTRA__
    <button id="login-submit" type="button" class="disabled_butt" disabled>INGRESA</button>
  </form>
  <div id="errores"></div>
  __DESAFIO__
  __BADGE__
  <script>
    // __MARCADOR__
    const u = document.getElementById('uname'), p = document.getElementById('pword');
    const b = document.getElementById('login-submit');
    function checkInputAndDisableSubmitButton() {
      const ok = u.value.trim() !== '' && p.value !== '';
      b.disabled = !ok; b.classList.toggle('disabled_butt', !ok);
    }
    u.addEventListener('keyup', checkInputAndDisableSubmitButton);
    p.addEventListener('keyup', checkInputAndDisableSubmitButton);
    b.addEventListener('click', async () => {
      if (document.querySelector('iframe[data-desafio]')) { return; }
      const r = await fetch('/claveunica/autenticar', {method: 'POST', headers: {'Content-Type': 'application/json'},
                                                      body: JSON.stringify({run: u.value, clave: p.value})});
      const d = await r.json();
      if (d.ok) { location.href = '__URL_HOME__'; }
      else { document.getElementById('errores').innerHTML = '<div class="alert alert-danger">RUN o ClaveÚnica incorrectos</div>'; }
    });
  </script>
</body></html>
"""

DESAFIO = f'<iframe data-desafio title="reCAPTCHA" width="304" height="78" src="{URL_DESAFIO_E2E}"></iframe>'

CABECERA = """<header><span>Hola Ana</span>
  <button class="btn btn-link" onclick="logoutCU()">Cerrar sesión</button></header>
  <script>function logoutCU() {}</script>"""

HTML_HOME = """<!doctype html>
<html lang="es"><head><meta charset="utf-8"><title>Portal del Consumidor — Inicio</title></head>
<body>__CABECERA__
  <a href="index.php?c=home&a=nomolestar">No Molestar</a>
</body></html>
"""

HTML_NM = """<!doctype html>
<html lang="es"><head><meta charset="utf-8"><title>No Molestar</title></head>
<body>__CABECERA__
  <h2>Registro No Molestar</h2>
  <div class="form-check"><input type="checkbox" class="form-check-input" id="aceptoNoMolestarCheck" checked>
    <label for="aceptoNoMolestarCheck">Acepto que el SERNAC envíe mis datos personales a la(s) empresa(s)…</label></div>
  <iframe id="frameNoMolestar" src="/no-molestar/solicitudes?p=__P__" width="1100" height="900"></iframe>
  <a href="/no-molestar/solicitudes?p=__P__">Abrir en otra pestaña</a>
  <div class="modal" id="sessionTimeoutModal" style="display:none"><p>Tu sesión está por expirar</p>
    <button type="button" id="btnKeepAlive">Permanecer conectado</button></div>
</body></html>
"""

HTML_LISTADO = """<!doctype html>
<html lang="es"><head><meta charset="utf-8"><title>No Molestar</title></head>
<body>
  __ALERTA__
  <ul class="nav nav-pills"><li><a id="pills-telefonos-tab" class="nav-link active">Teléfonos</a></li>
    <li><a id="pills-mail-tab" class="nav-link">Correos electrónicos</a></li></ul>
  <a class="btn btn-primary" href="/no-molestar/solicitudes/new?p=__P__">Ingresar nuevo teléfono y/o correo electrónico</a>
  <div id="bloqueos">__BLOQUEOS__</div>
</body></html>
"""

HTML_FORMULARIO = r"""<!doctype html>
<html lang="es"><head><meta charset="utf-8"><title>Nueva solicitud</title>
<style>.d-none{display:none!important} #panel-mail{display:none} .select2-dropdown{display:block;border:1px solid #888}
/* Bootstrap 4.4 custom-control: el <input> queda bajo su <label> (la etiqueta intercepta el clic) */
.custom-control{position:relative;z-index:1;display:block;min-height:1.5rem;padding-left:1.5rem}
.custom-control-inline{display:inline-flex;margin-right:1rem}
.custom-control-input{position:absolute;left:0;z-index:-1;width:1rem;height:1.25rem;opacity:0}
.custom-control-label{position:relative;margin-bottom:0;vertical-align:top}
.custom-control-label::before{position:absolute;top:.25rem;left:-1.5rem;display:block;width:1rem;height:1rem;pointer-events:none;content:"";background-color:#fff;border:#adb5bd solid 1px;border-radius:50%}
.custom-control-label::after{position:absolute;top:.25rem;left:-1.5rem;display:block;width:1rem;height:1rem;content:"";background:no-repeat 50%/50% 50%}
</style>
</head>
<body>
<div id="contenido">
  <h2>Ingresar nuevo teléfono y/o correo electrónico</h2>
  <div>
    <div class="custom-control custom-radio custom-control-inline">
      <input type="radio" id="customRadioInline1" name="tipo-contacto" value="telefono" class="custom-control-input" checked>
      <label class="custom-control-label" for="customRadioInline1">Teléfono</label>
    </div>
    <div class="custom-control custom-radio custom-control-inline">
      <input type="radio" id="customRadioInline2" name="tipo-contacto" value="mail" class="custom-control-input">
      <label class="custom-control-label" for="customRadioInline2">Correo electrónico</label>
    </div>
  </div>
  <div id="panel-telefono">
    <!-- como el portal real: el contenedor y el <select> comparten id="tipo-telefono" -->
    <div id="tipo-telefono"><label>Tipo</label>
      <select id="tipo-telefono"><option value="movil">Móvil</option><option value="fijo">Fijo</option></select>
    </div>
    <input id="telefono-pre" value="+56" disabled size="3">
    <input id="telefono-sub" maxlength="9">
  </div>
  <div id="panel-mail"><input id="mail-pre"> @ <input id="mail-sub"></div>
  <button id="ingresar-bloqueo" type="button">Ingresar</button>
  <div id="mensajes-error-telefonos"><div class="alert alert-danger d-none"></div></div>
  <ul id="lista-telefonos-bloqueos"></ul>
  <ul id="lista-mails-bloqueos"></ul>
  <h3>Empresas</h3>
  <select id="empresas" style="display:none"></select><span class="select2 select2-container select2-container--default"><span class="select2-selection select2-selection--single" role="combobox" aria-labelledby="select2-empresas-container" tabindex="0"><span id="select2-empresas-container">Ingrese el nombre de la empresa</span></span></span>
  <button id="ingresar-empresa" type="button">Ingresar</button>
  <div id="error-empresas" class="alert alert-danger d-none"></div>
  <ul id="lista-empresas"></ul>
  <a class="btn btn-success" href="/no-molestar/solicitudes?p=__P__">Volver</a>
  <a id="fake-continuar" href="#">Continuar</a>
</div>
<script>
const P = '__P__';
const $ = (s) => document.querySelector(s);
const datos = {telefonos: __TELEFONOS__, mails: [], empresas: []};
function li(clase, texto) { const e = document.createElement('li'); e.className = clase; e.textContent = texto; return e; }
function pintar() {
  const t = $('#lista-telefonos-bloqueos'); t.innerHTML = ''; datos.telefonos.forEach(x => t.appendChild(li('telefonos', x.numero)));
  const m = $('#lista-mails-bloqueos'); m.innerHTML = ''; datos.mails.forEach(x => m.appendChild(li('correos', x)));
  const e = $('#lista-empresas'); e.innerHTML = ''; datos.empresas.forEach(x => e.appendChild(li('empresas', x.text)));
}
function cambiarTipo() {
  const tel = $('#customRadioInline1').checked;
  $('#panel-telefono').style.display = tel ? '' : 'none';
  $('#panel-mail').style.display = tel ? 'none' : 'block';
}
$('#customRadioInline1').addEventListener('change', cambiarTipo);
$('#customRadioInline2').addEventListener('change', cambiarTipo);
function errorContacto(msg) { const a = $('#mensajes-error-telefonos > .alert'); a.textContent = msg; a.classList.remove('d-none'); }
$('#ingresar-bloqueo').addEventListener('click', (ev) => {
  ev.preventDefault(); $('#mensajes-error-telefonos > .alert').classList.add('d-none');
  if ($('#customRadioInline1').checked) {
    const n = '+56' + $('#telefono-sub').value.trim();
    if (!/^\+56[2-9]\d{8}$/.test(n)) { errorContacto('Ingrese un número de teléfono válido'); return; }
    if (datos.telefonos.some(x => x.numero === n)) { errorContacto('El teléfono ya fue ingresado'); return; }
    datos.telefonos.push({numero: n, tipo: $('select#tipo-telefono').value}); $('#telefono-sub').value = '';
  } else {
    const c = $('#mail-pre').value.trim() + '@' + $('#mail-sub').value.trim();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(c)) { errorContacto('Ingrese un correo electrónico válido'); return; }
    if (datos.mails.some(x => x.toLowerCase() === c.toLowerCase())) { errorContacto('El correo ya fue ingresado'); return; }
    datos.mails.push(c); $('#mail-pre').value = ''; $('#mail-sub').value = '';
  }
  pintar();
});
let desplegable = null, espera = null;
function cerrar() { if (desplegable) { desplegable.remove(); desplegable = null; } }
function abrir() {
  cerrar();
  desplegable = document.createElement('span');
  desplegable.className = 'select2-container select2-container--default select2-container--open';
  desplegable.innerHTML = '<span class="select2-dropdown"><span class="select2-search select2-search--dropdown">' +
    '<input class="select2-search__field" type="search" autocomplete="off"></span><span class="select2-results">' +
    '<ul class="select2-results__options" role="listbox" id="select2-empresas-results"></ul></span></span>';
  document.body.appendChild(desplegable);
  const campo = desplegable.querySelector('.select2-search__field');
  campo.addEventListener('input', () => {
    $('#select2-empresas-results').innerHTML = '<li class="select2-results__option loading-results" aria-disabled="true">Searching…</li>';
    clearTimeout(espera); espera = setTimeout(() => buscar(campo.value), 250);
  });
  campo.focus();
}
async function buscar(q) {
  const r = await fetch('/no-molestar/empresas/buscador?search=' + encodeURIComponent(q) + '&p=' + P);
  const d = await r.json();
  if (!desplegable) { return; }
  const ul = $('#select2-empresas-results'); ul.innerHTML = '';
  if (!d.results.length) { ul.innerHTML = '<li class="select2-results__option select2-results__message" role="alert">No se encontraron resultados</li>'; return; }
  d.results.forEach(it => {
    const e = document.createElement('li'); e.className = 'select2-results__option'; e.setAttribute('role', 'option'); e.textContent = it.text;
    e.addEventListener('mouseup', () => {
      const s = $('#empresas'); s.innerHTML = '';
      const o = document.createElement('option'); o.value = String(it.id); o.textContent = it.text; o.selected = true; s.appendChild(o);
      $('#select2-empresas-container').textContent = it.text; cerrar();
    });
    ul.appendChild(e);
  });
}
$('#empresas + .select2 .select2-selection').addEventListener('click', abrir);
$('#ingresar-empresa').addEventListener('click', (ev) => {
  ev.preventDefault(); const s = $('#empresas'), err = $('#error-empresas');
  if (!s.value) { err.textContent = 'Debe seleccionar una empresa'; err.classList.remove('d-none'); return; }
  if (datos.empresas.some(x => String(x.id) === s.value)) { err.textContent = 'La empresa ya fue ingresada'; err.classList.remove('d-none'); return; }
  err.classList.add('d-none');
  datos.empresas.push({id: Number(s.value), text: s.options[s.selectedIndex].textContent});
  s.innerHTML = ''; $('#select2-empresas-container').textContent = 'Ingrese el nombre de la empresa'; pintar();
});
$('#fake-continuar').addEventListener('click', async (ev) => {
  ev.preventDefault();
  if (!(datos.telefonos.length || datos.mails.length) || !datos.empresas.length) { return; }
  const r = await fetch('/no-molestar/solicitudes/configuracion?data_value=' + encodeURIComponent(JSON.stringify(datos)) + '&p=' + P);
  $('#contenido').innerHTML = await r.text();
  $('#confirmar-5').addEventListener('click', () => { $('#pills-telefonos').style.display = 'none'; $('#pills-mail').style.display = 'block'; });
  $('#cancelar-6').addEventListener('click', () => { $('#pills-mail').style.display = 'none'; $('#pills-telefonos').style.display = 'block'; });
});
pintar();
</script>
</body></html>
"""


def fragmento_revisar(datos: dict) -> str:
    """HTML que el servidor real devuelve para «Revisar solicitud» (canales marcados por defecto)."""
    filas_tel, filas_mail = [], []
    for emp in datos["empresas"]:
        texto = html.escape(emp["text"])
        for tel in datos["telefonos"]:
            numero, tipo = tel["numero"], tel.get("tipo", "movil")
            casillas = "".join(
                f'<div class="form-check form-check-inline"><input class="form-check-input" type="checkbox" checked '
                f'id="solicitud__1761200000_{tipo}_{emp["id"]}-empresa_canal_ids_{canal}" '
                f'name="solicitud[{numero}][telefono][{emp["id"]}-empresa][canal_ids][]" value="{canal}">'
                f'<label class="form-check-label">{etiqueta}</label></div>'
                for canal, etiqueta in ((10000, "Llamada"), (10001, "Sms"), (10002, "Whatsapp"))
            )
            filas_tel.append(f'<tr class="lista_agregar_empresa"><td>{texto}</td><td>{numero}</td><td>{casillas}</td></tr>')
        for correo in datos["mails"]:
            nombre = f"solicitud[{correo}][mail][{emp['id']}-empresa]"
            filas_mail.append(
                f'<tr><td>{texto}</td><td>{html.escape(correo)}</td><td><div class="form-check">'
                f'<input class="form-check-input" type="checkbox" checked id="{nombre}" name="{nombre}" value="1">'
                f'<label class="form-check-label">Bloquear</label></div></td></tr>'
            )
    return f"""<h2>Revisar solicitud</h2><p>Revise información y seleccione, al menos, un canal a bloquear.</p>
<ul class="nav nav-pills"><li><a id="pills-telefonos-tab" class="nav-link active">Teléfonos</a></li>
<li><a id="pills-mail-tab" class="nav-link">Correos electrónicos</a></li></ul>
<form id="new_solicitud" class="simple_form new_solicitud" action="/no-molestar/solicitudes" method="post" accept-charset="UTF-8">
  <input type="hidden" name="utf8" value="✓"><input type="hidden" name="authenticity_token" value="tokenFalso==">
  <input type="hidden" name="p" value="{JWT_PORTAL}">
  <div class="tab-pane active" id="pills-telefonos"><table id="configuracion-table">{''.join(filas_tel)}</table>
    <button type="button" class="btn" id="cancelar-5">Volver</button>
    <button type="button" class="btn btn-primary" id="confirmar-5">Confirmar</button></div>
  <div class="tab-pane" id="pills-mail" style="display:none"><table>{''.join(filas_mail)}</table>
    <button type="button" class="btn" id="cancelar-6">Volver</button>
    <input type="submit" name="commit" value="Enviar" id="finalizar" class="btn btn-primary" data-disable-with="Enviar"></div>
</form>"""


# Páginas con el token p= fuera de href/src/action (onclick, data-*, formaction) y con & -> &amp; al
# serializar: --explorar no debe guardarlo en ningún archivo (revisión security-1).
HTML_TOKEN_TOP = f"""<!doctype html>
<html lang="es"><head><meta charset="utf-8"><title>Portal con token</title></head>
<body><h1>Hola Ana</h1>
  <button id="card" onclick="window.location='index.php?c=home&a=nomolestar&p={JWT_PORTAL}'">Ingresar</button>
  <div id="d" data-url="/no-molestar/solicitudes/new?locale=es&p={JWT_PORTAL}">x</div>
  <iframe id="frameNoMolestar" src="/token/marco?p={JWT_PORTAL}" width="600" height="300"></iframe>
</body></html>
"""
HTML_TOKEN_MARCO = f"""<!doctype html>
<html lang="es"><head><meta charset="utf-8"><title>Marco con token</title></head>
<body><form action="/no-molestar/solicitudes?x=1&p={JWT_PORTAL}">
  <input type="hidden" name="p" value="{JWT_PORTAL}">
  <button formaction="/no-molestar/solicitudes?x=1&p={JWT_PORTAL}">Enviar</button>
  <a data-remote="/x?a=1&p={JWT_PORTAL}" href="#">enlace</a></form>
</body></html>
"""


class PortalFalsoE2E:
    """Servidor falso (vía context.route) con estado observable."""

    def __init__(self) -> None:
        self.status_portal = 200
        self.desafio_cu = False
        self.extra_cu = ""
        self.autenticado = False
        self.logins: list[dict] = []
        self.busquedas: list[str] = []
        self.configuraciones: list[dict] = []
        self.posts_finalizar: list[dict] = []
        self.telefonos_precargados = [{"numero": TEL_PRECARGADO, "tipo": "movil"}]
        self.numero = "7654321"

    def _html(self, route, cuerpo: str, status: int = 200) -> None:
        route.fulfill(status=status, content_type="text/html; charset=utf-8", body=cuerpo)

    def _json(self, route, datos: dict) -> None:
        route.fulfill(status=200, content_type="application/json; charset=utf-8", body=json.dumps(datos))

    def _comun(self, texto: str) -> str:
        return (
            texto.replace("__CABECERA__", CABECERA).replace("__BADGE__", BADGE).replace("__P__", JWT_PORTAL)
            .replace("__URL_CU__", URL_CU_E2E).replace("__URL_HOME__", URL_HOME_E2E)
        )

    def _listado(self, alerta: str = "", bloques: list[str] | None = None) -> str:
        cuerpo = "".join(f"<div class='bloqueo'>{html.escape(b)}</div>" for b in bloques or []) or "No posee bloqueos."
        return self._comun(HTML_LISTADO.replace("__ALERTA__", alerta).replace("__BLOQUEOS__", cuerpo))

    def manejar(self, route) -> None:
        req = route.request
        partes = urlsplit(req.url)
        ruta, consulta = partes.path, parse_qs(partes.query)
        if ruta.startswith("/app/consumidor"):
            if self.status_portal != 200:
                self._html(route, "<html><body><h1>403 Forbidden</h1>nginx</body></html>", self.status_portal)
            elif not self.autenticado:
                self._html(route, self._comun(HTML_LOGIN))
            elif consulta.get("a", [""])[0] == "nomolestar":
                self._html(route, self._comun(HTML_NM))
            else:
                self._html(route, self._comun(HTML_HOME))
        elif ruta.startswith("/claveunica/accounts/login"):
            cuerpo = (
                HTML_CU.replace("__EXTRA__", self.extra_cu)
                .replace("__DESAFIO__", DESAFIO if self.desafio_cu else "")
                .replace("__MARCADOR__", MARCADOR_SCRIPT)
            )
            self._html(route, self._comun(cuerpo))
        elif ruta == "/claveunica/autenticar":
            datos = json.loads(req.post_data or "{}")
            self.logins.append(datos)
            ok = bool(datos.get("run")) and datos.get("clave") not in ("", CLAVE_INCORRECTA)
            self.autenticado = self.autenticado or ok
            self._json(route, {"ok": ok})
        elif ruta.startswith("/recaptcha/"):
            self._html(route, "<html><body><label><input type=checkbox> No soy un robot</label></body></html>")
        elif ruta == "/no-molestar/solicitudes" and req.method == "POST":
            campos = parse_qs(req.post_data or "")
            self.posts_finalizar.append(campos)
            ultima = self.configuraciones[-1] if self.configuraciones else {"empresas": []}
            alerta = f'<div class="alert alert-success">Solicitud ingresada correctamente. N° de solicitud: {self.numero}</div>'
            self._html(route, self._listado(alerta, [e["text"] for e in ultima["empresas"]]))
        elif ruta == "/no-molestar/solicitudes":
            self._html(route, self._listado())
        elif ruta == "/no-molestar/solicitudes/new":
            cuerpo = HTML_FORMULARIO.replace("__TELEFONOS__", json.dumps(self.telefonos_precargados))
            self._html(route, self._comun(cuerpo))
        elif ruta == "/no-molestar/solicitudes/configuracion":
            datos = json.loads(consulta.get("data_value", ["{}"])[0])
            self.configuraciones.append(datos)
            self._html(route, fragmento_revisar(datos))
        elif ruta == "/no-molestar/empresas/buscador":
            termino = consulta.get("search", [""])[0]
            self.busquedas.append(termino)
            self._json(route, {"results": [{"id": i, "text": t, "rubro_id": 1} for i, t in buscar_en_directorio(termino)]})
        elif ruta == "/token/portal":
            self._html(route, HTML_TOKEN_TOP)
        elif ruta == "/token/marco":
            self._html(route, HTML_TOKEN_MARCO)
        else:
            self._html(route, "<html><body>404</body></html>", 404)


@pytest.fixture(scope="module")
def navegador():
    try:
        from playwright.sync_api import sync_playwright

        pw = sync_playwright().start()
    except Exception as exc:  # pragma: no cover - depende del entorno
        pytest.skip(f"Playwright no disponible: {exc}")
    try:
        browser = pw.chromium.launch(headless=True)
    except Exception as exc:  # pragma: no cover - depende del entorno
        pw.stop()
        pytest.skip(f"No se pudo lanzar Chromium headless: {exc}")
    yield browser
    browser.close()
    pw.stop()


@pytest.fixture
def portal() -> PortalFalsoE2E:
    return PortalFalsoE2E()


@pytest.fixture
def pagina(navegador, portal):
    contexto = navegador.new_context(locale="es-CL", viewport={"width": 1280, "height": 900})
    contexto.route("**/*", lambda route: route.abort())  # nada sale a la red
    contexto.route(f"{BASE}/**", portal.manejar)  # registrado después: tiene prioridad
    page = contexto.new_page()
    page.set_default_timeout(5000)
    yield page
    contexto.close()


@pytest.fixture
def urls_e2e(monkeypatch):
    monkeypatch.setattr(nm, "PORTAL_URL", URL_PORTAL_E2E)
    monkeypatch.setattr(nm, "URL_NO_MOLESTAR", URL_NM_E2E)
    monkeypatch.setattr(nm, "HOST_CLAVEUNICA", urlsplit(URL_CU_E2E).hostname)


def crear_bot_e2e(page, persona, tmp_path: Path, preguntar=None):
    return nm.BotSernac(
        page,
        persona,
        selectores=nm.SELECTORES,
        preguntar=preguntar if preguntar is not None else PreguntarGuionado(),
        sleep=SleepGrabador(),
        confirmar=False,
        capturas_dir=tmp_path / "salida" / "capturas",
        tiempos=tiempos_prueba(
            timeout_elemento_ms=5000,
            timeout_login_s=15,
            timeout_confirmacion_ms=8000,
            pausa_entre_acciones=(0.0, 0.0),
            pausa_entre_busquedas=(0.0, 0.0),
            retardo_tecleo_ms=5,
        ),
    )


def test_e2e_login_solicitud_real_con_numero_capturas_y_un_solo_post(pagina, portal, persona, tmp_path, urls_e2e):
    bot = crear_bot_e2e(pagina, persona, tmp_path)  # preguntar vacío: el badge invisible NO debe pausar

    bot.abrir_portal()
    bot.iniciar_sesion(nm.Secreto(SECRETO_CU))
    r = bot.enviar_solicitud(fila_real(1), ENTRADAS_CLARO)

    assert portal.logins == [{"run": "123456785", "clave": SECRETO_CU}]
    assert (r.estado, r.numero) == ("ok", "7654321")
    assert len(portal.posts_finalizar) == 1
    conf = portal.configuraciones[-1]
    assert [t["numero"] for t in conf["telefonos"]] == [TEL_PRECARGADO, "+56912345678"]
    assert conf["mails"] == [CORREO_PERSONA]
    assert [e["id"] for e in conf["empresas"]] == [178, 146420]
    marcadas = [k for k in portal.posts_finalizar[0] if k.startswith("solicitud[")]
    assert len(marcadas) == 2 * 2 + 2  # 2 empresas x 2 teléfonos (canal_ids[] agrupa 3) + 2 correos
    assert sum(len(v) for k, v in portal.posts_finalizar[0].items() if k.endswith("[canal_ids][]")) == 2 * 2 * 3
    capturas = sorted((tmp_path / "salida" / "capturas").glob("*.png"))
    assert len(capturas) == 2 and all(c.stat().st_size > 0 for c in capturas)
    assert sum(c.name.endswith("_marco.png") for c in capturas) == 1


def test_e2e_desafio_captcha_visible_pausa_hasta_que_la_persona_lo_resuelva(pagina, portal, persona, tmp_path, urls_e2e):
    portal.desafio_cu = True

    def persona_resuelve(prompt):
        pagina.evaluate("document.querySelector('iframe[data-desafio]').remove()")
        return ""

    preguntar = PreguntarGuionado([persona_resuelve])
    bot = crear_bot_e2e(pagina, persona, tmp_path, preguntar=preguntar)
    bot.abrir_portal()
    bot.iniciar_sesion(nm.Secreto(SECRETO_CU))

    assert preguntar.respuestas == []
    assert len(portal.logins) == 1
    assert portal.autenticado


def test_e2e_login_con_clave_incorrecta_no_reintenta(pagina, portal, persona, tmp_path, urls_e2e):
    bot = crear_bot_e2e(pagina, persona, tmp_path)
    bot.abrir_portal()
    with pytest.raises(nm.ErrorLogin):
        bot.iniciar_sesion(nm.Secreto(CLAVE_INCORRECTA))
    assert len(portal.logins) == 1


def test_e2e_portal_con_403_lanza_error_portal(pagina, portal, persona, tmp_path, urls_e2e):
    portal.status_portal = 403
    with pytest.raises(nm.ErrorPortal):
        crear_bot_e2e(pagina, persona, tmp_path).abrir_portal()


def _archivos_de(carpeta: Path) -> dict[str, list[Path]]:
    return {ext: sorted(carpeta.rglob(f"*.{ext}")) for ext in ("json", "html", "png")}


def test_e2e_capturar_pagina_no_guarda_valores_tipeados(pagina, portal, tmp_path):
    portal.extra_cu = (
        '<label for="notas">Notas</label><textarea id="notas" name="notas"></textarea>'
        '<input id="pre" name="pre" value="valor-precargado-123">'
    )
    pagina.goto(URL_CU_E2E)
    pagina.locator("#uname").fill("12345678-5")
    pagina.locator("#pword").fill(SECRETO_CU)
    pagina.locator("#notas").fill("comentario-privado-xyz")
    salida = tmp_path / "exploracion"

    rutas = nm.capturar_pagina(pagina, salida, 1)

    archivos = _archivos_de(salida)
    assert all(Path(r).exists() for r in rutas)
    assert archivos["json"] and archivos["html"] and archivos["png"]
    assert all(p.name.startswith("01_") for lista in archivos.values() for p in lista)
    datos = [json.loads(p.read_text(encoding="utf-8")) for p in archivos["json"]]
    principal = next(d for d in datos if "ClaveÚnica" in str(d.get("title", "")))
    assert principal["url"].endswith("?<omitido>")
    textos = [p.read_text(encoding="utf-8") for p in archivos["json"] + archivos["html"]]
    assert any("uname" in t for t in textos) and any('id="pword"' in t for t in textos)
    for prohibido in (SECRETO_CU, "12345678-5", "comentario-privado-xyz", "valor-precargado-123", MARCADOR_SCRIPT):
        assert not any(prohibido in t for t in textos), prohibido


def test_e2e_capturar_pagina_incluye_el_iframe_y_censura_el_token_p(pagina, portal, tmp_path):
    portal.autenticado = True
    pagina.goto(URL_NM_E2E)
    marco = pagina.frame_locator("#frameNoMolestar")
    marco.locator("a.btn-primary:has-text('Ingresar nuevo')").click()
    marco.locator("#telefono-sub").fill("987650000")
    salida = tmp_path / "exploracion"

    nm.capturar_pagina(pagina, salida, 2)

    archivos = _archivos_de(salida)
    assert archivos["json"] and archivos["html"] and archivos["png"]
    assert all(p.name.startswith("02_") for lista in archivos.values() for p in lista)
    textos = [p.read_text(encoding="utf-8") for p in archivos["json"] + archivos["html"]]
    assert any("telefono-sub" in t for t in textos)  # inventario/HTML del documento del iframe
    assert any("p=<omitido>" in t or "?<omitido>" in t for t in textos)
    for prohibido in (JWT_PORTAL, "987650000"):
        assert not any(prohibido in t for t in textos), prohibido


def test_e2e_capturar_pagina_censura_el_token_p_en_cualquier_atributo(pagina, portal, tmp_path):
    """security-1: onclick/data-url/data-remote/formaction con «&p=<JWT>» (serializado como &amp;p=)."""
    pagina.goto(f"{BASE}/token/portal?p={JWT_PORTAL}")
    pagina.frame_locator("#frameNoMolestar").locator("a").wait_for(state="attached")
    salida = tmp_path / "exploracion"

    nm.capturar_pagina(pagina, salida, 1)

    archivos = _archivos_de(salida)
    assert len(archivos["html"]) >= 2  # página + iframe
    textos = [p.read_text(encoding="utf-8") for p in archivos["json"] + archivos["html"]]
    assert any("data-remote" in t for t in textos) and any("onclick" in t for t in textos)
    for p, t in zip(archivos["json"] + archivos["html"], textos):
        assert JWT_PORTAL not in t, p.name


def test_e2e_capturar_pagina_png_no_muestra_lo_tipeado_en_el_iframe(pagina, portal, tmp_path):
    """security-3: el PNG enmascara también los campos del iframe (mismo PNG con y sin valor tipeado)."""
    portal.autenticado = True
    pagina.goto(URL_NM_E2E)
    marco = pagina.frame_locator("#frameNoMolestar")
    marco.locator("a.btn-primary:has-text('Ingresar nuevo')").click()
    campo = marco.locator("#telefono-sub")

    campo.fill("987650000")
    con_valor = nm.capturar_pagina(pagina, tmp_path / "con", 1)
    campo.fill("")
    sin_valor = nm.capturar_pagina(pagina, tmp_path / "sin", 1)

    def png(rutas):
        return next(r for r in rutas if r.suffix == ".png").read_bytes()

    assert png(con_valor) == png(sin_valor)
