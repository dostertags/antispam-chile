"""Regresiones de la revisión de no_molestar.py (hallazgos correctness-*, security-*, platform-*).

Cada prueba reproduce el escenario de falla del hallazgo con los dobles de conftest.py (FakePage /
PortalFalso, FabricaSMTP, PreguntarGuionado y COPIAS del tracker y del mapeo reales). Sin red.
Las pruebas E2E de security-1, security-3 y platform-1 están en test_e2e_portal_falso.py.
"""
from __future__ import annotations

import builtins
import contextlib
import csv
import getpass
import io
import logging
import os
import re
import smtplib
import time
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pytest
from playwright._impl._errors import TargetClosedError
from playwright.sync_api import Error as PWError
from playwright.sync_api import TimeoutError as PWTimeoutError

import no_molestar as nm
from conftest import (
    CONFIG_VALIDA,
    CORREO_PERSONA,
    ENTRADAS_CLARO,
    JWT_PORTAL,
    SECRETO_CU,
    SECRETO_GMAIL,
    TEL_PERSONA_E164,
    TEXTO_POR_ID,
    Elemento,
    FabricaSMTP,
    PortalFalso,
    PreguntarGuionado,
    S,
    es_fecha,
    fila_de,
    fila_real,
    tiempos_prueba,
    vacio,
    valor_celda,
)


# ----------------------------------------------------------------------------- utilidades
def _listo(**opciones) -> PortalFalso:
    p = PortalFalso(**opciones)
    p.entrar()
    return p


def _al_finalizar(p: PortalFalso, accion) -> None:
    """Reemplaza lo que hace el portal falso al hacer clic en #finalizar (el envío sí se cuenta)."""

    def finalizar(dom, el):
        p.eventos.append("finalizar")
        p.envios += 1
        accion(dom)

    p._finalizar = finalizar


def _bloques_claro() -> list[str]:
    return [e.texto for e in ENTRADAS_CLARO]


def _lanzador(pagina):
    @contextlib.contextmanager
    def falso(*args, **kwargs):
        yield pagina

    return falso


def _claves(monkeypatch, clave_gmail: str = SECRETO_GMAIL) -> list[str]:
    pedidos: list[str] = []

    def falso(prompt="", stream=None):
        pedidos.append(prompt)
        return clave_gmail if "gmail" in prompt.lower() else SECRETO_CU

    monkeypatch.setattr(getpass, "getpass", falso)
    return pedidos


def _sin_lineas(path: Path, *prefijos: str) -> None:
    lineas = path.read_text(encoding="utf-8-sig").splitlines()
    path.write_text("\n".join(ln for ln in lineas if not ln.startswith(prefijos)) + "\n", encoding="utf-8-sig")


def _resumen_csv(salida: Path) -> str:
    return next(salida.glob("resumen_*.csv")).read_text(encoding="utf-8-sig")


def _tracker_bloqueado(monkeypatch, tracker: Path) -> None:
    """os.replace sobre el tracker falla como si estuviera abierto en Excel."""
    real = os.replace

    def replace(origen, destino):
        if Path(destino) == Path(tracker):
            raise PermissionError("El proceso no tiene acceso al archivo (abierto en Excel)")
        return real(origen, destino)

    monkeypatch.setattr(os, "replace", replace)


# =============================================================================
# correctness-1: confirmación con el formulario aún en pantalla / mensajes negativos
# =============================================================================
def test_confirmacion_no_cuenta_si_el_formulario_sigue_en_pantalla(crear_bot):
    p = _listo()
    _al_finalizar(p, lambda dom: dom.agregar(S("confirmacion"), texto="Su solicitud fue ingresada correctamente"))
    preguntar = PreguntarGuionado(["o"])
    r = crear_bot(p.page, preguntar=preguntar).enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert r.estado == "incierta"
    assert p.envios == 1
    assert any("Se envió" in t for t in preguntar.prompts)


@pytest.mark.parametrize(
    "mensaje",
    ["La solicitud no pudo ser enviada", "Solicitud no ingresada", "Error: la solicitud no fue registrada"],
)
def test_confirmacion_con_mensaje_negativo_no_cuenta_como_exito(crear_bot, mensaje):
    p = _listo()

    def rechazo(dom):
        p.formulario()  # Rails vuelve a mostrar el formulario (ya sin #finalizar) con un aviso
        dom.agregar(S("confirmacion", 1), texto=mensaje)

    _al_finalizar(p, rechazo)
    preguntar = PreguntarGuionado(["o"])
    r = crear_bot(p.page, preguntar=preguntar).enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert r.estado == "incierta"
    assert p.envios == 1
    assert preguntar.respuestas == []


def test_confirmacion_con_no_molestar_en_el_texto_sigue_siendo_exito(crear_bot):
    p = _listo()
    _al_finalizar(p, lambda dom: p.listado(
        alerta="Su solicitud No Molestar fue ingresada. N° de solicitud: 7654321", bloques=_bloques_claro()
    ))
    r = crear_bot(p.page).enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert (r.estado, r.numero) == ("ok", "7654321")


# =============================================================================
# correctness-2: error de Playwright después del clic en «Enviar»
# =============================================================================
def _marco_roto(p: PortalFalso, errores: int | None = None) -> None:
    """Desde ahora el documento del iframe lanza errores de Playwright (navega o murió el navegador);
    con `errores`, solo las primeras N lecturas fallan."""
    todos, texto = p.marco.todos, p.marco.texto_visible
    cuenta = {"n": 0}

    def quizas_fallar():
        cuenta["n"] += 1
        if errores is None or cuenta["n"] <= errores:
            raise PWError("Execution context was destroyed, most likely because of a navigation")

    def todos_rotos(selector):
        quizas_fallar()
        return todos(selector)

    def texto_roto():
        quizas_fallar()
        return texto()

    p.marco.todos = todos_rotos
    p.marco.texto_visible = texto_roto


def test_error_de_playwright_tras_enviar_queda_incierta_y_pregunta(crear_bot):
    p = _listo()

    def navega(dom):
        p.listado(alerta="Su solicitud fue ingresada correctamente.", bloques=_bloques_claro())
        _marco_roto(p)

    _al_finalizar(p, navega)
    preguntar = PreguntarGuionado(["o"])
    r = crear_bot(p.page, preguntar=preguntar).enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert r.estado == "incierta"
    assert p.envios == 1
    assert preguntar.respuestas == []


def test_error_transitorio_mientras_el_marco_navega_no_impide_confirmar(crear_bot):
    p = _listo()

    def navega(dom):
        p.listado(alerta="Su solicitud fue ingresada correctamente. N° de solicitud: 1234567", bloques=_bloques_claro())
        _marco_roto(p, errores=3)

    _al_finalizar(p, navega)
    bot = crear_bot(p.page, tiempos=tiempos_prueba(timeout_confirmacion_ms=2000))
    r = bot.enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert (r.estado, r.numero) == ("ok", "1234567")
    assert p.envios == 1


# =============================================================================
# correctness-3 / platform-2: N° de solicitud con palabras comunes
# =============================================================================
@pytest.mark.parametrize(
    "texto",
    [
        "N° Solicitud\tFecha\tEmpresa",
        "Folio\tEmpresa",
        "Número de solicitud\nFecha de ingreso",
        "Solicitud no ingresada",
        "Tu solicitud No Molestar fue ingresada correctamente.",
        "N° Solicitud | Empresa | Fecha",
    ],
)
def test_extraer_numero_solicitud_no_devuelve_palabras(texto):
    assert nm.extraer_numero_solicitud(texto) is None


def test_extraer_numero_solicitud_salta_no_molestar_y_encuentra_el_numero():
    texto = "Tu solicitud No Molestar fue ingresada. N° de solicitud: 7654321"
    assert nm.extraer_numero_solicitud(texto) == "7654321"


# =============================================================================
# correctness-4 / platform-4: Ctrl+C o aborto después del clic en «Enviar»
# =============================================================================
def test_ctrl_c_al_pedir_el_numero_tras_responder_s_conserva_el_ok(crear_bot):
    p = _listo(resultados_envio=(("nada",),))
    with pytest.raises((nm.AbortoUsuario, KeyboardInterrupt)) as e:
        crear_bot(p.page, preguntar=PreguntarGuionado(["s", KeyboardInterrupt])).enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert isinstance(e.value, nm.AbortoUsuario)
    assert e.value.resultado is not None
    assert (e.value.resultado.estado, e.value.resultado.numero) == ("ok", None)
    assert p.envios == 1


def test_ctrl_c_durante_la_espera_de_confirmacion_queda_incierta_con_captura(crear_bot, tmp_path):
    p = _listo()

    def sin_respuesta(dom):
        def interrumpir(timeout):
            raise KeyboardInterrupt

        p.page.wait_for_timeout = interrumpir

    _al_finalizar(p, sin_respuesta)
    capturas = tmp_path / "capturas"
    bot = crear_bot(p.page, capturas_dir=capturas, tiempos=tiempos_prueba(timeout_confirmacion_ms=2000))
    with pytest.raises((nm.AbortoUsuario, KeyboardInterrupt)) as e:
        bot.enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert isinstance(e.value, nm.AbortoUsuario)
    assert e.value.resultado is not None and e.value.resultado.estado == "incierta"
    assert e.value.resultado.captura and Path(e.value.resultado.captura).exists()
    assert p.envios == 1


def test_main_ctrl_c_tras_confirmar_s_anota_el_tracker_y_sale_con_130(archivos, monkeypatch, tiempos_rapidos, capsys):
    portal = PortalFalso(resultados_envio=(("nada",),))
    monkeypatch.setattr(nm, "lanzar_navegador", _lanzador(portal.page))
    _claves(monkeypatch)
    monkeypatch.setattr(builtins, "input", PreguntarGuionado(["s", KeyboardInterrupt]))
    codigo = nm.main(archivos.argv("--solo", "sernac", "--empresa", "entel", "--sin-confirmar"))
    assert codigo == 130
    assert portal.envios == 1
    assert es_fecha(valor_celda(archivos.tracker, 5, "No Molestar ingresado"), date.today())
    assert vacio(valor_celda(archivos.tracker, 5, "N° solicitud SERNAC"))
    assert ";sernac;ok;" in _resumen_csv(archivos.salida)


def test_main_ctrl_c_al_cerrar_excel_conserva_el_ok_sernac_en_el_resumen(archivos, monkeypatch, tiempos_rapidos, capsys):
    portal = PortalFalso()
    monkeypatch.setattr(nm, "lanzar_navegador", _lanzador(portal.page))
    _claves(monkeypatch)
    _tracker_bloqueado(monkeypatch, archivos.tracker)
    monkeypatch.setattr(builtins, "input", PreguntarGuionado([KeyboardInterrupt]))
    codigo = nm.main(archivos.argv("--solo", "sernac", "--empresa", "entel", "--sin-confirmar"))
    assert codigo == 130
    assert portal.envios == 1
    assert ";sernac;ok;" in _resumen_csv(archivos.salida)
    assert archivos.tracker.with_name(archivos.tracker.stem + ".pendiente.xlsx").exists()


# =============================================================================
# correctness-5: resultados de correo perdidos al interrumpir la fase
# =============================================================================
def test_main_ctrl_c_en_la_pausa_entre_correos_conserva_los_enviados(archivos, monkeypatch, tiempos_rapidos, capsys):
    fab = FabricaSMTP()
    monkeypatch.setattr(smtplib, "SMTP_SSL", fab)
    _claves(monkeypatch)
    monkeypatch.setattr(builtins, "input", PreguntarGuionado())
    pausas: list[float] = []

    def dormir(segundos):
        if segundos == 70:  # pausa_entre_correos (tiempos_rapidos)
            pausas.append(segundos)
            if len(pausas) == 2:
                raise KeyboardInterrupt

    monkeypatch.setattr(time, "sleep", dormir)
    codigo = nm.main(archivos.argv("--solo", "correo", "--limite", "3"))
    assert codigo == 130
    assert len(fab.enviados) == 2
    assert _resumen_csv(archivos.salida).count(";correo;ok;") == 2


def test_main_ctrl_c_al_guardar_tras_un_correo_lo_cuenta_como_enviado(archivos, monkeypatch, tiempos_rapidos, capsys):
    fab = FabricaSMTP()
    monkeypatch.setattr(smtplib, "SMTP_SSL", fab)
    _claves(monkeypatch)
    _tracker_bloqueado(monkeypatch, archivos.tracker)
    monkeypatch.setattr(builtins, "input", PreguntarGuionado([KeyboardInterrupt]))
    codigo = nm.main(archivos.argv("--solo", "correo", "--limite", "2"))
    assert codigo == 130
    assert len(fab.enviados) == 1
    assert _resumen_csv(archivos.salida).count(";correo;ok;") == 1


# =============================================================================
# correctness-6: --limite gastado en empresas «sin entrada en SERNAC»
# =============================================================================
def _sernac_hecho(path: Path, numeros) -> None:
    wb = openpyxl.load_workbook(path)
    ws = wb["Empresas"]
    for n in numeros:
        ws.cell(row=n + 1, column=15).value = datetime(2026, 10, 20)  # O = «No Molestar ingresado»
    wb.save(path)


def test_dry_run_limite_no_se_gasta_en_empresas_sin_entrada_en_sernac(archivos, sin_efectos, capsys):
    _sernac_hecho(archivos.tracker, range(1, 18))
    assert nm.main(archivos.argv("--dry-run", "--solo", "sernac", "--limite", "1")) == 0
    salida = capsys.readouterr().out
    assert "SERNAC: 1 solicitudes pendientes" in salida
    linea_19 = next(ln for ln in salida.splitlines() if re.match(r"\s*20\s+19\s", ln))
    linea_18 = next(ln for ln in salida.splitlines() if re.match(r"\s*19\s+18\s", ln))
    assert "pendiente con 1 entrada" in linea_19
    assert nm.DETALLE_SIN_ENTRADA in linea_18


def test_planificar_limite_cuenta_solo_las_empresas_que_van_al_portal(archivos):
    _sernac_hecho(archivos.tracker, range(1, 18))
    args = nm.parse_args(archivos.argv("--solo", "sernac", "--limite", "1"))
    plan = nm._planificar(args, nm._validar(args, con_plantilla=False))
    assert [f.numero for f in plan.sernac] == [19]
    assert [f.numero for f in plan.sin_entrada] == [18]


# =============================================================================
# correctness-7: error transitorio al mapear una empresa detenía toda la fase
# =============================================================================
def test_main_error_transitorio_al_mapear_una_empresa_no_detiene_la_fase(archivos, monkeypatch, tiempos_rapidos, capsys):
    _sin_lineas(archivos.mapeo, "4;")
    portal = PortalFalso()
    monkeypatch.setattr(nm, "lanzar_navegador", _lanzador(portal.page))
    _claves(monkeypatch)
    monkeypatch.setattr(builtins, "input", PreguntarGuionado())

    def portal_caido(self, termino):
        raise PWTimeoutError("Timeout 30000ms exceeded.")

    monkeypatch.setattr(nm.BotSernac, "buscar_empresas", portal_caido)
    codigo = nm.main(archivos.argv("--solo", "sernac", "--empresa", "entel", "--empresa", "banco de chile", "--sin-confirmar"))
    salida = capsys.readouterr()
    assert codigo == 1
    assert portal.envios == 1
    assert portal.solicitudes[0]["empresas"] == [(27498, TEXTO_POR_ID[27498])]
    assert "detenida" not in salida.out + salida.err
    resumen = _resumen_csv(archivos.salida)
    assert ";Entel;sernac;fallida;" in resumen and ";Banco de Chile;sernac;ok;" in resumen


# =============================================================================
# correctness-8: contacto precargado parecido (subcadena) ocultaba el del usuario
# =============================================================================
def test_correo_precargado_que_contiene_el_del_usuario_no_lo_reemplaza(crear_bot):
    p = _listo(correos_precargados=("j" + CORREO_PERSONA,))
    r = crear_bot(p.page).enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert r.estado == "ok"
    assert ("correo", CORREO_PERSONA, None) in p.contactos_agregados
    assert CORREO_PERSONA in p.solicitudes[0]["correos"]


def test_telefono_precargado_mas_largo_no_cuenta_como_el_del_usuario(crear_bot):
    p = _listo(telefonos_precargados=(TEL_PERSONA_E164 + "9",))
    r = crear_bot(p.page).enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert r.estado == "ok"
    assert ("telefono", TEL_PERSONA_E164, "Móvil") in p.contactos_agregados


# =============================================================================
# security-1: _censurar_urls con formas codificadas del token p=
# =============================================================================
@pytest.mark.parametrize(
    "plantilla",
    [
        "index.php?c=home&amp;a=nomolestar&amp;p={t}",
        "/x?a=1&#38;p={t}",
        "/x?a=1&#x26;p={t}",
        "/x%3Fa%3D1%26p%3D{t}",
        '<a href="/no-molestar/solicitudes?locale=es&amp;p={t}">',
    ],
)
def test_censurar_urls_cubre_las_formas_codificadas_del_token(plantilla):
    assert JWT_PORTAL not in nm._censurar_urls(plantilla.format(t=JWT_PORTAL))


def test_sanear_html_quita_la_query_de_urls_relativas():
    html = f'<a href="/no-molestar/solicitudes/new?correo=ana@example.cl&amp;p={JWT_PORTAL}">x</a>'
    limpio = nm._censurar_urls(nm._sanear_html(html))
    assert "ana@example.cl" not in limpio and JWT_PORTAL not in limpio


# =============================================================================
# security-2: la ClaveÚnica solo se escribe en accounts.claveunica.gob.cl (https)
# =============================================================================
@pytest.mark.parametrize(
    "url",
    [
        "https://claveunica-gob.example/accounts/login/?next=/openid",
        "http://accounts.claveunica.gob.cl/accounts/login/?next=/openid",
        "https://accounts.claveunica.gob.cl.example.com/accounts/login/",
    ],
)
def test_iniciar_sesion_no_escribe_la_clave_fuera_de_claveunica(crear_bot, capsys, url):
    p = PortalFalso()
    original = p.pantalla_cu

    def cu_en_otro_sitio():
        original()
        p.top.url = url

    p.pantalla_cu = cu_en_otro_sitio
    p.page.goto(nm.PORTAL_URL)

    def persona_ingresa_a_mano(prompt):
        p.entrar()
        return ""

    preguntar = PreguntarGuionado([persona_ingresa_a_mano])
    crear_bot(p.page, preguntar=preguntar).iniciar_sesion(nm.Secreto(SECRETO_CU))
    assert p.top_cu["run"].escrito == [] and p.top_cu["clave"].escrito == []
    assert p.envios_login == []
    assert p.logueado and preguntar.respuestas == []
    assert SECRETO_CU not in capsys.readouterr().out


# =============================================================================
# security-4: mapeo_sernac.csv sin fórmulas de Excel
# =============================================================================
def test_mapeo_neutraliza_formulas_al_guardar_y_las_recupera_al_cargar(tmp_path):
    ruta = tmp_path / "mapeo.csv"
    entradas = [
        nm.EntradaSernac(None, '=HYPERLINK("http://x.example") ENTEL PCS'),
        nm.EntradaSernac(183, "-ENTEL MOVIL"),
        nm.EntradaSernac(None, "@ENTEL"),
        nm.EntradaSernac(None, "+ENTEL"),
    ]
    nm.Mapeo(ruta).registrar(fila_de("=Entel", numero=4), entradas, nota="-nota")
    filas = list(csv.reader(io.StringIO(ruta.read_text(encoding="utf-8-sig")), delimiter=";"))[1:]
    assert filas
    for fila in filas:
        assert not any(c[:1] in ("=", "+", "-", "@") for c in fila), fila
    cargado = nm.Mapeo.cargar(ruta)
    assert cargado.entradas(fila_de("x", numero=4)) == entradas


def test_mapeo_conserva_textos_que_empiezan_con_apostrofo_comun(tmp_path):
    ruta = tmp_path / "mapeo.csv"
    entradas = [nm.EntradaSernac(None, "'S GRAVENHAGE IMPORTADORA")]
    nm.Mapeo(ruta).registrar(fila_de("Otra", numero=7), entradas)
    assert nm.Mapeo.cargar(ruta).entradas(fila_de("x", numero=7)) == entradas


# =============================================================================
# security-5: claves con aspecto de secreto (otros nombres, otras secciones)
# =============================================================================
@pytest.mark.parametrize("clave", ["secreto", "secret", "pwd", "app_pwd", "passwd", "api_key", "apiKey", "pin", "llave"])
def test_es_clave_secreta_reconoce_otros_nombres(clave):
    assert nm._es_clave_secreta(clave) is True


@pytest.mark.parametrize("clave", ["nombre", "rut", "telefono", "celular", "correo", "usuario", "email", "spinner", "opinion"])
def test_es_clave_secreta_no_marca_datos_normales(clave):
    assert nm._es_clave_secreta(clave) is False


def test_cargar_config_avisa_secretos_con_otros_nombres_y_en_otras_secciones(escribir, caplog, capsys):
    texto = (
        CONFIG_VALIDA
        + 'secreto = "s1-valor"\npwd = "s2-valor"\napi_key = "s3-valor"\npin = "9911"\n\n'
        + '[claveunica]\npassword = "s4-valor"\n\n[smtp.cuenta]\ntoken = "s5-valor"\n'
    )
    caplog.set_level(logging.WARNING)
    cfg = nm.cargar_config(escribir("config.toml", texto))
    assert cfg.persona.nombre == "Ana Pérez Soto" and cfg.gmail_usuario == "ana.optout@gmail.com"
    for clave in ("gmail.secreto", "gmail.pwd", "gmail.api_key", "gmail.pin", "claveunica.password", "smtp.cuenta.token"):
        assert clave in caplog.text, clave
    salida = capsys.readouterr()
    for valor in ("s1-valor", "s2-valor", "s3-valor", "9911", "s4-valor", "s5-valor"):
        assert valor not in caplog.text + salida.out + salida.err


# =============================================================================
# platform-1: radios/casillas «custom-control» de Bootstrap 4 (la etiqueta intercepta el clic)
# =============================================================================
def test_radio_custom_control_que_intercepta_el_clic_se_marca_igual(crear_bot):
    p = _listo(acepto_marcado=False)
    original = p.formulario

    def formulario_bs4():
        original()
        for clave in ("radio_telefono", "radio_correo"):
            p.marco.el(S(clave)).datos["intercepta"] = True

    p.formulario = formulario_bs4
    original_ir = p.pantalla_no_molestar

    def no_molestar_bs4():
        original_ir()
        p.top.el(S("acepto_datos")).datos["intercepta"] = True

    p.pantalla_no_molestar = no_molestar_bs4
    r = crear_bot(p.page).enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert r.estado == "ok"
    assert ("correo", CORREO_PERSONA, None) in p.contactos_agregados
    assert p.solicitudes[0]["acepto"] is True


def test_canal_custom_control_desmarcado_se_marca_igual(crear_bot):
    p = _listo(canales_tel_desmarcados=(0,))
    original = p.revisar

    def revisar_bs4():
        original()
        for el in p.marco.todos(S("canales", 0)):
            el.datos["intercepta"] = True

    p.revisar = revisar_bs4
    r = crear_bot(p.page).enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert r.estado == "ok"
    assert all(p.solicitudes[0]["canales_tel"])


# =============================================================================
# platform-3: re-ingreso con la sesión de ClaveÚnica aún vigente
# =============================================================================
def test_reingreso_con_sesion_de_claveunica_vigente_sigue_sin_pedir_la_clave(crear_bot):
    p = PortalFalso()
    p.page.goto(nm.PORTAL_URL)
    bot = crear_bot(p.page)
    bot.iniciar_sesion(nm.Secreto(SECRETO_CU))
    p.pantalla_cu = p.entrar  # ClaveÚnica recuerda la sesión y devuelve al portal sin formulario
    p.expirar_en_goto_no_molestar = p.gotos_no_molestar + 1
    r = bot.enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert r.estado == "ok" and p.envios == 1
    assert len(p.envios_login) == 1


# =============================================================================
# platform-5: navegador cerrado o caído
# =============================================================================
def test_es_transitorio_playwright_con_navegador_cerrado_es_falso():
    assert nm.es_transitorio_playwright(TargetClosedError()) is False
    assert nm.es_transitorio_playwright(PWError("Target page, context or browser has been closed")) is False
    assert nm.es_transitorio_playwright(PWError("Target crashed")) is False
    assert nm.es_transitorio_playwright(PWTimeoutError("Timeout 15000ms exceeded.")) is True


def test_main_navegador_cerrado_detiene_la_fase_sin_reintentar_cada_empresa(archivos, monkeypatch, tiempos_rapidos, capsys):
    portal = PortalFalso()
    monkeypatch.setattr(nm, "lanzar_navegador", _lanzador(portal.page))
    _claves(monkeypatch)
    monkeypatch.setattr(builtins, "input", PreguntarGuionado())
    original = portal.top.al_goto
    intentos: list[str] = []

    def goto(dom, url):
        if url == nm.URL_NO_MOLESTAR:
            intentos.append(url)
            raise TargetClosedError()
        return original(dom, url)

    portal.top.al_goto = goto
    codigo = nm.main(archivos.argv("--solo", "sernac", "--empresa", "entel", "--empresa", "banco de chile", "--sin-confirmar"))
    salida = capsys.readouterr()
    assert codigo == 1
    assert len(intentos) == 1
    assert tiempos_rapidos.llamadas.count(50) == 0  # sin pausas entre empresas: la fase se detuvo
    resumen = _resumen_csv(archivos.salida)
    assert resumen.count("fase SERNAC detenida") == 2


# =============================================================================
# platform-6: --mapear mantiene la sesión y permite un re-ingreso por empresa
# =============================================================================
def test_buscar_empresas_con_formulario_abierto_pulsa_mantener_sesion(crear_bot):
    p = _listo()
    bot = crear_bot(p.page)
    assert bot.buscar_empresas("ENTEL")
    p._modal()  # aparece «Tu sesión está por expirar» mientras el buscador sigue abierto
    assert bot.buscar_empresas("CLARO")
    assert p.clicks_mantener == 1


def test_main_mapear_con_sesion_que_expira_reingresa_una_vez_por_empresa(archivos, monkeypatch, tiempos_rapidos, capsys):
    _sin_lineas(archivos.mapeo, "3;", "4;", "11;")
    portal = PortalFalso()
    monkeypatch.setattr(nm, "lanzar_navegador", _lanzador(portal.page))
    _claves(monkeypatch)

    def acepta_y_expira(prompt):
        portal.logueado = False
        portal.pantalla_login()  # el portal cerró la sesión mientras la persona respondía
        return ""

    monkeypatch.setattr(builtins, "input", PreguntarGuionado([acepta_y_expira, acepta_y_expira, ""]))
    codigo = nm.main(archivos.argv("--mapear"))
    assert codigo == 0
    assert len(portal.envios_login) == 3  # inicial + un re-ingreso para cada una de las 2 empresas siguientes
    m = nm.Mapeo.cargar(archivos.mapeo)
    for numero in (3, 4, 11):
        assert m.entradas(fila_de("x", numero=numero)), numero


# =============================================================================
# platform-7: contraseña de Gmail con caracteres no ASCII
# =============================================================================
def test_enviador_conectar_con_clave_no_ascii_es_error_de_autenticacion(persona):
    clave = "ñandú-Secreto-91"
    fab = FabricaSMTP({"login": [UnicodeEncodeError("ascii", clave, 0, 1, "ordinal not in range(128)")]})
    enviador = nm.EnviadorCorreo("ana.optout@gmail.com", nm.Secreto(clave), persona, smtp_factory=fab,
                                 sleep=lambda s: None, tiempos=tiempos_prueba())
    with pytest.raises(smtplib.SMTPAuthenticationError) as e:
        enviador.conectar()
    assert clave not in str(e.value) and e.value.__cause__ is None and e.value.__suppress_context__


def test_main_clave_gmail_no_ascii_explica_y_la_fase_sernac_se_ejecuta(archivos, monkeypatch, tiempos_rapidos, capsys):
    portal = PortalFalso()
    monkeypatch.setattr(nm, "lanzar_navegador", _lanzador(portal.page))
    fab = FabricaSMTP({"login": [UnicodeEncodeError("ascii", "mi contraseña", 9, 10, "ordinal not in range(128)")]})
    monkeypatch.setattr(smtplib, "SMTP_SSL", fab)
    _claves(monkeypatch, clave_gmail="mi contraseña")
    monkeypatch.setattr(builtins, "input", PreguntarGuionado())
    codigo = nm.main(archivos.argv("--empresa", "entel", "--sin-confirmar"))
    salida = capsys.readouterr()
    texto = salida.out + salida.err
    assert codigo == 1
    assert portal.envios == 1
    assert "aplicaci" in texto.lower()
    assert "Error inesperado" not in texto
    assert ";correo;fallida;" in _resumen_csv(archivos.salida)


# =============================================================================
# platform-8: «The results could not be loaded.» no es «sin resultados»
# =============================================================================
def test_buscar_empresas_con_error_de_carga_del_buscador_no_devuelve_lista_vacia(crear_bot):
    p = _listo()

    def falla(dom, termino):
        p.busquedas.append(termino)
        dom.quitar(S("opciones_empresa"))
        dom.poner(nm.SELECTOR_SIN_RESULTADOS, [Elemento(texto="The results could not be loaded.")])

    p._buscar = falla
    with pytest.raises(PWTimeoutError):
        crear_bot(p.page).buscar_empresas("ENTEL")
    assert len(p.busquedas) == tiempos_prueba()["reintentos"]


@pytest.mark.parametrize("texto", ["No se encontraron resultados", "No results found"])
def test_buscar_empresas_con_mensaje_sin_resultados_devuelve_lista_vacia(crear_bot, texto):
    p = _listo()

    def sin_resultados(dom, termino):
        dom.quitar(S("opciones_empresa"))
        dom.poner(nm.SELECTOR_SIN_RESULTADOS, [Elemento(texto=texto)])

    p._buscar = sin_resultados
    assert crear_bot(p.page).buscar_empresas("ZZZ") == []


# =============================================================================
# platform-9: lista de teléfonos sin el id de contenedor supuesto
# =============================================================================
def test_telefono_precargado_en_lista_sin_id_de_contenedor_se_reconoce(crear_bot):
    p = _listo()
    original = p.formulario

    def formulario():
        original()
        p.marco.poner("li.telefonos", p.marco.todos(S("lista_telefonos")))
        p.marco.quitar(S("lista_telefonos"))

    p.formulario = formulario
    r = crear_bot(p.page).enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert r.estado == "ok"
    assert not any(c[0] == "telefono" for c in p.contactos_agregados)


def test_telefono_precargado_en_items_telefonos_que_no_son_li_se_reconoce(crear_bot):
    """WALK.md registró los ítems como «.telefonos» sin la etiqueta: también deben reconocerse."""
    p = _listo()
    original = p.formulario

    def formulario():
        original()
        p.marco.poner(".telefonos", p.marco.todos(S("lista_telefonos")))
        p.marco.quitar(S("lista_telefonos"))

    p.formulario = formulario
    r = crear_bot(p.page).enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert r.estado == "ok"
    assert not any(c[0] == "telefono" for c in p.contactos_agregados)


# =============================================================================
# Ramas adicionales de los arreglos (después del clic, origen de ClaveÚnica, correo)
# =============================================================================
def test_navegador_cerrado_tras_enviar_deja_de_esperar_y_queda_incierta(crear_bot):
    p = _listo()

    def se_cierra(dom):
        p.top.esperas.clear()

        def cerrado(*args, **kwargs):
            raise TargetClosedError()

        p.marco.todos = cerrado
        p.marco.texto_visible = cerrado

    _al_finalizar(p, se_cierra)
    bot = crear_bot(p.page, preguntar=PreguntarGuionado(["o"]), tiempos=tiempos_prueba(timeout_confirmacion_ms=2000))
    r = bot.enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert r.estado == "incierta" and p.envios == 1
    assert p.top.esperas == []  # no siguió sondeando un navegador cerrado


def test_listado_con_aviso_de_error_no_cuenta_como_exito(crear_bot):
    p = _listo()

    def listado_con_error(dom):
        p.listado(bloques=_bloques_claro())
        dom.agregar(nm.SELECTOR_ALERTA_ERROR, texto="Ocurrió un problema al guardar la solicitud")

    _al_finalizar(p, listado_con_error)
    r = crear_bot(p.page, preguntar=PreguntarGuionado(["o"])).enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert r.estado == "incierta" and p.envios == 1


def test_formulario_que_desaparece_sin_listado_ni_aviso_sigue_sin_confirmar(crear_bot):
    p = _listo()
    _al_finalizar(p, lambda dom: dom.reiniciar(texto="Cargando…"))
    r = crear_bot(p.page, preguntar=PreguntarGuionado(["o"])).enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert r.estado == "incierta" and p.envios == 1


def test_error_inesperado_despues_del_clic_queda_incierta_no_fallida(crear_bot):
    p = _listo()
    bot = crear_bot(p.page)

    def disco_lleno(fila, sufijo=""):
        raise RuntimeError("disco lleno")

    bot._capturar = disco_lleno
    r = bot.enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert r.estado == "incierta" and "RuntimeError" in r.detalle
    assert r.captura is None and p.envios == 1


def test_marcar_que_no_queda_marcado_falla_sin_enviar(crear_bot):
    p = _listo(acepto_marcado=False)
    original = p.pantalla_no_molestar

    def acepto_que_se_desmarca():
        original()
        el = p.top.el(S("acepto_datos"))
        el.al_marcar = lambda dom, e: setattr(e, "marcado", False)

    p.pantalla_no_molestar = acepto_que_se_desmarca
    r = crear_bot(p.page).enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert r.estado == "fallida" and "marcar" in r.detalle
    assert p.envios == 0


def test_iniciar_sesion_no_escribe_la_clave_si_el_sitio_cambia_tras_el_run(crear_bot):
    p = PortalFalso()
    original = p.pantalla_cu

    def cu_que_redirige():
        original()

        def redirigir(dom, el, texto):
            dom.url = "https://claveunica-gob.example/accounts/login/"

        p.top_cu["run"].al_escribir = redirigir

    p.pantalla_cu = cu_que_redirige
    p.page.goto(nm.PORTAL_URL)
    with pytest.raises(nm.ErrorLogin, match="no la escribí"):
        crear_bot(p.page).iniciar_sesion(nm.Secreto(SECRETO_CU))
    assert p.top_cu["clave"].escrito == []


def test_main_ctrl_c_durante_el_envio_de_un_correo_lo_deja_incierto(archivos, monkeypatch, tiempos_rapidos, capsys):
    fab = FabricaSMTP({"send_message": [None, KeyboardInterrupt()]})
    monkeypatch.setattr(smtplib, "SMTP_SSL", fab)
    _claves(monkeypatch)
    monkeypatch.setattr(builtins, "input", PreguntarGuionado())
    assert nm.main(archivos.argv("--solo", "correo", "--limite", "3")) == 130
    resumen = _resumen_csv(archivos.salida)
    assert resumen.count(";correo;ok;") == 1 and resumen.count(";correo;incierta;") == 1


def test_main_sernac_sesion_que_expira_antes_de_cada_empresa_reingresa_una_vez_por_empresa(
    archivos, monkeypatch, tiempos_rapidos, capsys
):
    portal = PortalFalso(resultados_envio=(("ok", "1234567"), ("ok", "7654321")))
    monkeypatch.setattr(nm, "lanzar_navegador", _lanzador(portal.page))
    _claves(monkeypatch)
    monkeypatch.setattr(builtins, "input", PreguntarGuionado())
    expira = {"pendiente": True}
    finalizar_original, goto_original = portal._finalizar, portal.top.al_goto

    def finalizar(dom, el):
        finalizar_original(dom, el)
        expira["pendiente"] = True  # la sesión vence mientras corre la pausa entre empresas

    def goto(dom, url):
        if url == nm.URL_NO_MOLESTAR and expira["pendiente"]:
            expira["pendiente"] = False
            portal.logueado = False
        return goto_original(dom, url)

    portal._finalizar, portal.top.al_goto = finalizar, goto
    codigo = nm.main(archivos.argv("--solo", "sernac", "--empresa", "entel", "--empresa", "banco de chile", "--sin-confirmar"))
    assert codigo == 0
    assert portal.envios == 2
    assert len(portal.envios_login) == 3  # inicial + un re-ingreso en cada empresa
