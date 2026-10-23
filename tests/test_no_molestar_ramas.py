"""Ramas de error y de borde de no_molestar.py que la suite principal no recorre (sin red).

Mismos dobles que la suite principal (conftest.py): FakePage/PortalFalso, FabricaSMTP,
PreguntarGuionado y COPIAS del tracker y del mapeo reales.
"""
from __future__ import annotations

import builtins
import contextlib
import getpass
import json
import logging
import os
import pickle
import smtplib
import time
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pytest
from playwright.sync_api import Error as PWError

import no_molestar as nm
from conftest import (
    CONFIG_VALIDA,
    BuscadorFalso,
    ENTRADAS_CLARO,
    JWT_PORTAL,
    SECRETO_CU,
    SECRETO_GMAIL,
    TEXTO_POR_ID,
    URL_ANCHOR_V2,
    DOMFalso,
    FabricaSMTP,
    FakePage,
    PortalFalso,
    PreguntarGuionado,
    SleepGrabador,
    fila_de,
    fila_real,
    filas_reales,
    tiempos_prueba,
    valor_celda,
)

HOY = date(2026, 10, 22)


def _listo(**opciones) -> PortalFalso:
    p = PortalFalso(**opciones)
    p.entrar()
    return p


# =============================================================================
# Secretos, config y plantilla
# =============================================================================
def test_secreto_no_se_puede_serializar():
    with pytest.raises(TypeError):
        pickle.dumps(nm.Secreto(SECRETO_CU))


def _registro(msg, args=(), **extra) -> logging.LogRecord:
    r = logging.LogRecord(nm.NOMBRE_LOGGER, logging.INFO, __file__, 1, msg, args, None)
    for k, v in extra.items():
        setattr(r, k, v)
    return r


def test_filtro_secretos_censura_stack_info_y_exc_text_y_tolera_args_rotos():
    filtro = nm.FiltroSecretos([nm.Secreto(SECRETO_CU)])
    r = _registro("sin formato", stack_info=f"pila con {SECRETO_CU}", exc_text=f"traza con {SECRETO_CU}")
    assert filtro.filter(r) is True
    assert SECRETO_CU not in r.stack_info and SECRETO_CU not in r.exc_text
    roto = _registro("dos %s %s", (SECRETO_CU,))  # faltan argumentos: no debe romper el logging
    filtro.filter(roto)
    assert SECRETO_CU not in roto.getMessage()


def test_contexto_para_formatea_fechas_y_enteros(persona):
    fila = fila_de("Entel", numero=4, valores={"#": 4, "Enviado (fecha)": datetime(2026, 10, 1), "Monto": 3.0})
    ctx = nm.contexto_para(fila, persona, HOY)
    assert ctx["enviado_fecha"] == "01-10-2026"
    assert ctx["monto"] == "3"


def test_cargar_config_con_gmail_usuario_invalido(escribir):
    texto = CONFIG_VALIDA.replace('usuario = "ana.optout@gmail.com"', 'usuario = "no-es-correo"')
    with pytest.raises(nm.ErrorConfig, match="gmail"):
        nm.cargar_config(escribir("config.toml", texto))


def test_cargar_config_y_plantilla_ilegibles(tmp_path):
    malo = tmp_path / "malo.txt"
    malo.write_bytes(b"\xff\xfe\x00 no es utf-8 \x81")
    with pytest.raises(nm.ErrorConfig):
        nm.cargar_config(malo)
    with pytest.raises(nm.ErrorPlantilla):
        nm.cargar_plantilla(malo)


# =============================================================================
# Tracker
# =============================================================================
def test_tracker_corrupto_lanza_error_tracker(tmp_path):
    malo = tmp_path / "tracker.xlsx"
    malo.write_bytes(b"esto no es un xlsx")
    with pytest.raises(nm.ErrorTracker):
        nm.Tracker(malo)


def test_tracker_numero_que_parece_formula_se_guarda_como_texto(fabrica_tracker):
    path = fabrica_tracker(filas_reales(1))
    t = nm.Tracker(path)
    t.marcar_sernac(t.filas()[0], HOY, "=1+1")
    t.guardar(preguntar=PreguntarGuionado())
    assert valor_celda(path, 2, "N° solicitud SERNAC") == "'=1+1"


def test_tracker_guardar_con_ctrl_c_deja_pendiente_y_aborta(fabrica_tracker, monkeypatch):
    path = fabrica_tracker(filas_reales(1))
    antes = path.read_bytes()
    real = os.replace

    def falla(src, dst, *a, **k):
        if Path(dst).resolve() == path.resolve():
            raise PermissionError(13, "en uso")
        return real(src, dst, *a, **k)

    monkeypatch.setattr(os, "replace", falla)
    t = nm.Tracker(path)
    t.marcar_correo(t.filas()[0], HOY)
    with pytest.raises(nm.AbortoUsuario):
        t.guardar(preguntar=PreguntarGuionado([KeyboardInterrupt]))
    assert path.with_name(f"{path.stem}.pendiente.xlsx").exists()
    assert path.read_bytes() == antes


def test_tracker_respaldar_dos_veces_no_pisa_la_copia(fabrica_tracker, tmp_path):
    t = nm.Tracker(fabrica_tracker(filas_reales(1)))
    a, b = t.respaldar(tmp_path / "r"), t.respaldar(tmp_path / "r")
    assert a != b and a.exists() and b.exists()


def test_tracker_informa_formulas_en_las_celdas_de_salida(fabrica_tracker):
    filas = filas_reales(1)
    filas[0]["No Molestar ingresado"] = "=TODAY()"
    t = nm.Tracker(fabrica_tracker(filas))
    assert t.formulas_en_salida(t.filas()[0]) == ["No Molestar ingresado"]


# =============================================================================
# Mapeo
# =============================================================================
ENCABEZADO = "#;Marca;ID SERNAC;Entrada SERNAC;Nota"


@pytest.mark.parametrize(
    "fila",
    ["x;Claro Chile;178;CLARO MOVIL;", "1;Claro Chile;178;;", ";;178;CLARO MOVIL;"],
    ids=["numero-no-entero", "id-sin-entrada", "sin-numero-ni-marca"],
)
def test_mapeo_mal_formado_variantes(escribir, fila):
    with pytest.raises(nm.ErrorMapeo):
        nm.Mapeo.cargar(escribir("mapeo.csv", f"{ENCABEZADO}\n{fila}\n"))


def test_mapeo_vacio_queda_sin_empresas(escribir):
    assert nm.Mapeo.cargar(escribir("mapeo.csv", "   \n")).entradas(fila_de("Claro", numero=1)) is None


def test_mapeo_ilegible_lanza_error_mapeo(tmp_path):
    malo = tmp_path / "mapeo.csv"
    malo.write_bytes(b"\xff\xfe\x81\x00")
    with pytest.raises(nm.ErrorMapeo):
        nm.Mapeo.cargar(malo)


def test_mapeo_nota_con_salto_de_linea_entre_comillas(escribir):
    path = escribir("mapeo.csv", f'{ENCABEZADO}\n1;Claro Chile;178;CLARO MOVIL;"línea 1\nlínea 2"\n2;VTR;;;\n')
    m = nm.Mapeo.cargar(path)
    assert m.entradas(fila_de("Claro Chile", numero=1)) == [nm.EntradaSernac(178, "CLARO MOVIL")]
    assert m.entradas(fila_de("VTR", numero=2)) == []


def test_mapeo_registrar_empresa_nueva_se_inserta_en_orden_y_sin_numero_al_final(escribir):
    path = escribir("mapeo.csv", f"{ENCABEZADO}\n1;Claro Chile;178;CLARO MOVIL;\n3;Movistar;175;MOVISTAR;\n", bom=True)
    m = nm.Mapeo.cargar(path)
    m.registrar(fila_de("VTR", numero=2), [nm.EntradaSernac(None, "VTR TELEFONIA")])
    m.registrar(fila_de("Sin Número", numero=None), [])
    numeros = [ln.split(";")[0] for ln in path.read_text(encoding="utf-8-sig").splitlines()[1:]]
    assert numeros == ["1", "2", "3", ""]
    assert nm.Mapeo.cargar(path).entradas(fila_de("sin numero", numero=None)) == []


def test_mapeo_registrar_con_el_csv_bloqueado_lanza_error_mapeo(tmp_path, monkeypatch):
    path = tmp_path / "mapeo.csv"
    monkeypatch.setattr(os, "replace", lambda *a, **k: (_ for _ in ()).throw(PermissionError(13, "en uso")))
    with pytest.raises(nm.ErrorMapeo, match="Excel"):
        nm.Mapeo(path).registrar(fila_de("Entel", numero=4), [])
    assert not list(tmp_path.iterdir())


# =============================================================================
# resolver_entradas
# =============================================================================
class Buscador:
    def __init__(self, resultados):
        self.resultados = resultados

    def buscar_empresas(self, termino):
        return list(self.resultados.get(nm.normalizar(termino), []))


def test_resolver_entradas_reintenta_tras_respuesta_invalida_y_sin_sugerencias(tmp_path, capsys):
    m = nm.Mapeo.cargar(tmp_path / "m.csv")
    bot = Buscador({"zeta": ["OTRA EMPRESA", "MAS RUIDO"]})
    r = nm.resolver_entradas(fila_de("Zeta", numero=9), bot, m, PreguntarGuionado(["", "99", "x", "2"]))
    assert r == [nm.EntradaSernac(None, "MAS RUIDO")]
    salida = capsys.readouterr().out
    assert "No hay sugerencias" in salida and "no válida" in salida


def test_resolver_entradas_si_no_puede_guardar_sigue_en_memoria(tmp_path, monkeypatch, caplog):
    m = nm.Mapeo.cargar(tmp_path / "m.csv")
    monkeypatch.setattr(m, "registrar", lambda *a, **k: (_ for _ in ()).throw(nm.ErrorMapeo("bloqueado")))
    caplog.set_level(logging.WARNING)
    r = nm.resolver_entradas(fila_real(11), Buscador({"banco de chile": ["BANCO DE CHILE - BANCO EDWARDS - CITI"]}), m, PreguntarGuionado([""]))
    assert r == [nm.EntradaSernac(None, "BANCO DE CHILE - BANCO EDWARDS - CITI")]
    assert "memoria" in caplog.text


def test_resolver_entradas_eof_aborta(tmp_path):
    with pytest.raises(nm.AbortoUsuario):
        nm.resolver_entradas(fila_real(11), Buscador({}), nm.Mapeo(tmp_path / "m.csv"), PreguntarGuionado([EOFError]))


# =============================================================================
# EnviadorCorreo
# =============================================================================
def _enviador(persona, fabrica):
    return nm.EnviadorCorreo("ana.optout@gmail.com", nm.Secreto("clave"), persona, smtp_factory=fabrica,
                             sleep=SleepGrabador(), tiempos=tiempos_prueba())


def test_enviador_procesar_marca_ok_aunque_el_tracker_tenga_formula(persona, fabrica_tracker):
    path = fabrica_tracker([{"Marca": "Paris", "Vía de envío": "Correo", "Correo datos personales": "a@paris.cl",
                             "Enviado (fecha)": "=TODAY()"}])
    t = nm.Tracker(path)
    (r,) = _enviador(persona, FabricaSMTP()).procesar(t.filas(), nm.Plantilla("A", "B"), t, hoy=HOY, preguntar=PreguntarGuionado())
    assert r.estado == "ok" and "no se anotó" in r.detalle
    assert valor_celda(path, 2, "Enviado (fecha)") == "=TODAY()"


def test_enviador_procesar_con_marcador_desconocido_falla_esa_fila(persona, fabrica_tracker):
    path = fabrica_tracker([{"Marca": "Paris", "Vía de envío": "Correo", "Correo datos personales": "a@paris.cl"}])
    t = nm.Tracker(path)
    fab = FabricaSMTP()
    (r,) = _enviador(persona, fab).procesar(t.filas(), nm.Plantilla("A {nada}", "B"), t, hoy=HOY, preguntar=PreguntarGuionado())
    assert r.estado == "fallida" and "nada" in r.detalle
    assert fab.enviados == []


def test_enviador_login_fallido_cierra_la_conexion(persona):
    fab = FabricaSMTP({"login": [smtplib.SMTPServerDisconnected("cortado")]})
    with pytest.raises(smtplib.SMTPServerDisconnected):
        _enviador(persona, fab).conectar()


# =============================================================================
# BotSernac (FakePage)
# =============================================================================
def test_abrir_portal_sin_red_lanza_error_portal(crear_bot):
    dom = DOMFalso()
    dom.al_goto = lambda d, url: (_ for _ in ()).throw(PWError("net::ERR_NAME_NOT_RESOLVED"))
    with pytest.raises(nm.ErrorPortal, match="mantenci"):
        crear_bot(FakePage(dom)).abrir_portal()


def _con_status_no_molestar(p: PortalFalso, estados: list[int]) -> None:
    original = p.top.al_goto

    def goto(dom, url):
        devuelto = original(dom, url)
        return estados.pop(0) if url == nm.URL_NO_MOLESTAR and estados else devuelto

    p.top.al_goto = goto


def test_enviar_solicitud_con_403_en_no_molestar_detiene_la_fase(crear_bot):
    p = _listo()
    _con_status_no_molestar(p, [403])
    with pytest.raises(nm.ErrorPortal):
        crear_bot(p.page).enviar_solicitud(fila_real(1), ENTRADAS_CLARO)


def test_enviar_solicitud_con_503_transitorio_reintenta(crear_bot):
    p = _listo()
    _con_status_no_molestar(p, [503])
    assert crear_bot(p.page).enviar_solicitud(fila_real(1), ENTRADAS_CLARO).estado == "ok"
    assert p.envios == 1


def test_iniciar_sesion_sin_boton_claveunica_lanza_error_login(crear_bot):
    p = PortalFalso()
    p.page.goto(nm.PORTAL_URL)
    with pytest.raises(nm.ErrorLogin, match="ClaveÚnica"):
        crear_bot(p.page, selectores=dict(nm.SELECTORES, login_boton_claveunica=["#otro"])).iniciar_sesion(nm.Secreto(SECRETO_CU))


def test_iniciar_sesion_sin_formulario_claveunica_lanza_error_login(crear_bot):
    p = PortalFalso()
    p.page.goto(nm.PORTAL_URL)
    with pytest.raises(nm.ErrorLogin, match="formulario"):
        crear_bot(p.page, selectores=dict(nm.SELECTORES, cu_run=["#otro-run"])).iniciar_sesion(nm.Secreto(SECRETO_CU))


def test_iniciar_sesion_con_boton_ingresa_deshabilitado_lanza_error_login(crear_bot):
    p = PortalFalso()
    original = p.pantalla_cu

    def sin_keyup():
        original()
        p.top_cu["run"].al_escribir = p.top_cu["clave"].al_escribir = None

    p.pantalla_cu = sin_keyup
    p.page.goto(nm.PORTAL_URL)
    with pytest.raises(nm.ErrorLogin, match="deshabilitado"):
        crear_bot(p.page).iniciar_sesion(nm.Secreto(SECRETO_CU))
    assert p.envios_login == []


def test_enviar_solicitud_con_sesion_expirada_y_sin_clave_lanza_error_login(crear_bot):
    p = _listo()
    p.expirar_en_goto_no_molestar = 1
    with pytest.raises(nm.ErrorLogin, match="expiró"):
        crear_bot(p.page).enviar_solicitud(fila_real(1), ENTRADAS_CLARO)


def test_enviar_solicitud_sin_coincidencia_acepta_abortar_tras_respuesta_invalida(crear_bot, capsys):
    p = _listo()
    with pytest.raises(nm.AbortoUsuario):
        crear_bot(p.page, preguntar=PreguntarGuionado(["99", "abortar"])).enviar_solicitud(
            fila_de("Falabella", numero=28), [nm.EntradaSernac(None, "FALABELLA")]
        )
    assert "no válida" in capsys.readouterr().out
    assert p.envios == 0


def test_enviar_solicitud_m_sin_elegir_en_el_navegador_queda_fallida(crear_bot):
    p = _listo()
    r = crear_bot(p.page, preguntar=PreguntarGuionado(["m"])).enviar_solicitud(
        fila_de("Falabella", numero=28), [nm.EntradaSernac(None, "FALABELLA")]
    )
    assert r.estado == "fallida" and "Debe seleccionar una empresa" in r.detalle
    assert p.envios == 0


def test_enviar_solicitud_contacto_que_no_aparece_se_reintenta_y_falla(crear_bot):
    p = _listo(telefonos_precargados=())
    p._agregar_contacto = lambda dom, el: None  # el portal no responde al «Ingresar»
    r = crear_bot(p.page).enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert r.estado == "fallida" and "No se pudo preparar" in r.detalle
    assert p.clicks_nueva == tiempos_prueba()["reintentos"] and p.envios == 0


def test_enviar_solicitud_ctrl_c_tras_enviar_conserva_el_resultado_incierto(crear_bot):
    p = _listo(resultados_envio=(("nada",),))
    with pytest.raises(nm.AbortoUsuario) as e:
        crear_bot(p.page, preguntar=PreguntarGuionado(["quizás", KeyboardInterrupt])).enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert e.value.resultado is not None and e.value.resultado.estado == "incierta"
    assert p.envios == 1


def test_enviar_solicitud_sin_carpeta_de_capturas_no_se_cae(crear_bot, tmp_path):
    archivo = tmp_path / "no-es-carpeta"
    archivo.write_text("x")
    p = _listo(resultados_envio=(("nada",),))
    r = crear_bot(p.page, preguntar=PreguntarGuionado(["o"]), capturas_dir=archivo / "capturas").enviar_solicitud(fila_real(1), ENTRADAS_CLARO)
    assert (r.estado, r.captura) == ("incierta", None)


def test_hay_captcha_por_url_de_iframe_anidado_pero_no_hcaptcha_invisible(crear_bot):
    p = _listo()
    bot = crear_bot(p.page)
    p.top.frames.append("https://newassets.hcaptcha.com/captcha/v1/x/static/hcaptcha.html#frame=checkbox-invisible")
    assert bot.hay_captcha() is False
    p.top.frames.append(URL_ANCHOR_V2)  # sin locator: p. ej. dentro del iframe de No Molestar
    assert bot.hay_captcha() is True


@pytest.mark.parametrize("texto", ["No se encontraron resultados", "Searching…", "Buscando…"])
def test_buscar_empresas_no_devuelve_mensajes_de_select2_como_opciones(crear_bot, texto):
    p = _listo(directorio=[(1, texto)])
    assert crear_bot(p.page).buscar_empresas(texto[:5]) == []


def test_buscar_empresas_reingresa_si_la_sesion_expiro(crear_bot):
    p = PortalFalso()
    p.page.goto(nm.PORTAL_URL)
    bot = crear_bot(p.page)
    bot.iniciar_sesion(nm.Secreto(SECRETO_CU))
    p.expirar_en_goto_no_molestar = p.gotos_no_molestar + 1
    assert bot.buscar_empresas("MOVISTAR") == [TEXTO_POR_ID[175], TEXTO_POR_ID[1192505]]
    assert len(p.envios_login) == 2


# =============================================================================
# Exploración
# =============================================================================
def _pagina_con_formulario() -> FakePage:
    dom = DOMFalso(url=f"https://www.sernac.cl/no-molestar/solicitudes?p={JWT_PORTAL}", titulo="No Molestar",
                   texto="Formulario con valor-tipeado-123")
    dom.frames.append(f"https://www.sernac.cl/no-molestar/solicitudes/new?p={JWT_PORTAL}")
    return FakePage(dom)


def test_capturar_pagina_respaldo_sin_inventario_js_censura_urls_y_valores(tmp_path):
    page = _pagina_con_formulario()
    rutas = nm.capturar_pagina(page, tmp_path / "exp", 3)
    assert {p.suffix for p in rutas} == {".json", ".html", ".png"}
    assert all(p.name.startswith("03_") for p in rutas)
    datos = json.loads(next(p for p in rutas if p.suffix == ".json").read_text(encoding="utf-8"))
    assert datos["url"].endswith("?<omitido>") and all("p=" not in f for f in datos["frames"])
    for p in rutas:
        if p.suffix != ".png":
            assert JWT_PORTAL not in p.read_text(encoding="utf-8")


def test_sanear_html_quita_valores_textareas_y_scripts():
    html = '<input value="secreto"><textarea>texto</textarea><script>const P="x";</script><a href="/a?p=TOKEN">a</a>'
    limpio = nm._censurar_urls(nm._sanear_html(html))
    for prohibido in ("secreto", "texto", 'const P', "TOKEN"):
        assert prohibido not in limpio
    assert "p=<omitido>" in limpio or "?<omitido>" in limpio  # la query de URLs relativas se quita entera


def test_explorar_toma_instantaneas_al_cambiar_y_con_c(tmp_path):
    page = _pagina_con_formulario()
    ordenes = iter(["c", "q"])

    def preguntar(prompt=""):
        time.sleep(0.05)
        return next(ordenes)

    n = nm.explorar(page, tmp_path / "exp", preguntar, lambda s: time.sleep(0.01))
    assert n >= 2
    assert page.url.startswith(nm.PORTAL_URL)  # abrió el portal
    assert len(list((tmp_path / "exp").glob("*.json"))) == n


def test_explorar_termina_con_eof(tmp_path):
    page = FakePage(DOMFalso(url=nm.PORTAL_URL, titulo="Portal"))

    def eof(prompt=""):
        raise EOFError

    assert nm.explorar(page, tmp_path / "exp", eof, lambda s: time.sleep(0.01)) >= 0


# =============================================================================
# main(): modos y errores de fase
# =============================================================================
def _lanzador(pagina, registro=None):
    @contextlib.contextmanager
    def falso(*args, **kwargs):
        if registro is not None:
            registro.append((args, kwargs))
        yield pagina

    return falso


def _claves(monkeypatch, pedidos):
    def falso(prompt="", stream=None):
        pedidos.append(prompt)
        return SECRETO_GMAIL if "gmail" in prompt.lower() else SECRETO_CU

    monkeypatch.setattr(getpass, "getpass", falso)


def _sin_lineas(path: Path, *prefijos: str) -> None:
    lineas = path.read_text(encoding="utf-8-sig").splitlines()
    path.write_text("\n".join(ln for ln in lineas if not ln.startswith(prefijos)) + "\n", encoding="utf-8-sig")


def test_main_explorar_abre_el_navegador_sin_pedir_claves(archivos, monkeypatch, tiempos_rapidos):
    page = _pagina_con_formulario()
    lanzamientos: list = []
    monkeypatch.setattr(nm, "lanzar_navegador", _lanzador(page, lanzamientos))
    monkeypatch.setattr(getpass, "getpass", lambda *a, **k: pytest.fail("--explorar no pide claves"))
    monkeypatch.setattr(builtins, "input", lambda prompt="": "q")
    assert nm.main(archivos.argv("--explorar")) == 0
    assert len(lanzamientos) == 1 and page.url == nm.PORTAL_URL


def test_main_sernac_con_empresa_sin_mapeo_pregunta_y_registra(archivos, monkeypatch, tiempos_rapidos, capsys):
    _sin_lineas(archivos.mapeo, "4;")
    portal = PortalFalso()
    monkeypatch.setattr(nm, "lanzar_navegador", _lanzador(portal.page))
    pedidos: list[str] = []
    _claves(monkeypatch, pedidos)
    monkeypatch.setattr(builtins, "input", PreguntarGuionado([""]))
    codigo = nm.main(archivos.argv("--solo", "sernac", "--empresa", "entel", "--sin-confirmar"))
    assert codigo == 0
    assert nm.Mapeo.cargar(archivos.mapeo).entradas(fila_de("x", numero=4)) == [nm.EntradaSernac(None, TEXTO_POR_ID[183])]
    assert portal.solicitudes[0]["empresas"] == [(183, TEXTO_POR_ID[183])]
    assert valor_celda(archivos.tracker, 5, "N° solicitud SERNAC") == "1234567"


@pytest.mark.parametrize("respuesta, estado", [("s", "omitida"), ("n", "sin_coincidencia")])
def test_main_sernac_empresa_sin_mapeo_omitida_o_sin_entrada(archivos, monkeypatch, tiempos_rapidos, capsys, respuesta, estado):
    _sin_lineas(archivos.mapeo, "4;")
    antes = archivos.tracker.read_bytes()
    portal = PortalFalso()
    monkeypatch.setattr(nm, "lanzar_navegador", _lanzador(portal.page))
    _claves(monkeypatch, [])
    monkeypatch.setattr(builtins, "input", PreguntarGuionado([respuesta]))
    assert nm.main(archivos.argv("--solo", "sernac", "--empresa", "entel", "--sin-confirmar")) == 0
    assert portal.envios == 0
    assert archivos.tracker.read_bytes() == antes
    resumen = next(archivos.salida.glob("resumen_*.csv")).read_text(encoding="utf-8-sig")
    assert f";sernac;{estado};" in resumen


def test_main_login_rechazado_falla_sernac_y_sigue_con_el_correo(archivos, monkeypatch, tiempos_rapidos, capsys):
    portal = PortalFalso(login="error")
    monkeypatch.setattr(nm, "lanzar_navegador", _lanzador(portal.page))
    fab = FabricaSMTP()
    monkeypatch.setattr(smtplib, "SMTP_SSL", fab)
    pedidos: list[str] = []
    _claves(monkeypatch, pedidos)
    monkeypatch.setattr(builtins, "input", PreguntarGuionado())
    codigo = nm.main(archivos.argv("--empresa", "entel", "--sin-confirmar"))
    salida = capsys.readouterr()
    assert codigo == 1
    assert len(pedidos) == 2
    assert [str(m["To"]) for m in fab.enviados] == ["protecciondedatos@entel.cl"]
    assert "fase SERNAC detenida" in salida.out
    assert SECRETO_CU not in salida.out + salida.err


def test_main_portal_en_mantencion_reintenta_cuando_la_persona_lo_pide(archivos, monkeypatch, tiempos_rapidos, capsys):
    portal = PortalFalso(status_portal=403)
    monkeypatch.setattr(nm, "lanzar_navegador", _lanzador(portal.page))
    _claves(monkeypatch, [])

    def reintentar(prompt):
        portal.status_portal = 200
        return ""

    monkeypatch.setattr(builtins, "input", PreguntarGuionado([reintentar]))
    assert nm.main(archivos.argv("--solo", "sernac", "--empresa", "entel", "--sin-confirmar")) == 0
    assert portal.envios == 1


def test_main_portal_en_mantencion_y_abortar_falla_la_fase(archivos, monkeypatch, tiempos_rapidos, capsys):
    portal = PortalFalso(status_portal=503)
    monkeypatch.setattr(nm, "lanzar_navegador", _lanzador(portal.page))
    _claves(monkeypatch, [])
    monkeypatch.setattr(builtins, "input", PreguntarGuionado(["abortar"]))
    assert nm.main(archivos.argv("--solo", "sernac", "--empresa", "entel", "--sin-confirmar")) == 1
    assert "mantención" in capsys.readouterr().out


def test_main_error_inesperado_en_una_empresa_no_detiene_las_demas(archivos, monkeypatch, tiempos_rapidos, capsys):
    portal = PortalFalso()
    monkeypatch.setattr(nm, "lanzar_navegador", _lanzador(portal.page))
    _claves(monkeypatch, [])
    monkeypatch.setattr(builtins, "input", PreguntarGuionado())
    original = nm.BotSernac.enviar_solicitud
    llamadas = []

    def a_veces(self, fila, entradas):
        llamadas.append(fila.empresa)
        if len(llamadas) == 1:
            raise RuntimeError("fallo raro")
        return original(self, fila, entradas)

    monkeypatch.setattr(nm.BotSernac, "enviar_solicitud", a_veces)
    codigo = nm.main(archivos.argv("--solo", "sernac", "--empresa", "entel", "--empresa", "banco de chile", "--sin-confirmar"))
    assert codigo == 1
    assert llamadas == ["Entel", "Banco de Chile"] and portal.envios == 1
    assert "error inesperado" in capsys.readouterr().out


def test_main_ctrl_c_al_confirmar_sale_con_130_y_resume(archivos, monkeypatch, tiempos_rapidos, capsys):
    portal = PortalFalso()
    monkeypatch.setattr(nm, "lanzar_navegador", _lanzador(portal.page))
    _claves(monkeypatch, [])
    monkeypatch.setattr(builtins, "input", PreguntarGuionado([KeyboardInterrupt]))
    assert nm.main(archivos.argv("--solo", "sernac", "--empresa", "entel", "--confirmar")) == 130
    assert portal.envios == 0
    assert "RESUMEN" in capsys.readouterr().out


def test_main_gmail_inalcanzable_marca_los_correos_fallidos(archivos, monkeypatch, tiempos_rapidos, capsys):
    fab = FabricaSMTP({"conectar": [ConnectionRefusedError("sin red")] * 5})
    monkeypatch.setattr(smtplib, "SMTP_SSL", fab)
    _claves(monkeypatch, [])
    monkeypatch.setattr(nm, "lanzar_navegador", lambda *a, **k: pytest.fail("no debía abrir el navegador"))
    monkeypatch.setattr(builtins, "input", PreguntarGuionado())
    assert nm.main(archivos.argv("--solo", "correo", "--limite", "2")) == 1
    assert len(fab.conexiones) == nm.TIEMPOS["reintentos"]
    assert "No se pudo conectar a Gmail" in next(archivos.salida.glob("resumen_*.csv")).read_text(encoding="utf-8-sig")


def test_main_gmail_rechaza_la_clave_a_mitad_de_la_fase(archivos, monkeypatch, tiempos_rapidos, capsys):
    auth = smtplib.SMTPAuthenticationError(535, b"5.7.8 revoked")
    fab = FabricaSMTP({"login": [None, auth], "send_message": [None, smtplib.SMTPServerDisconnected("cortado")]})
    monkeypatch.setattr(smtplib, "SMTP_SSL", fab)
    _claves(monkeypatch, [])
    monkeypatch.setattr(builtins, "input", PreguntarGuionado())
    assert nm.main(archivos.argv("--solo", "correo", "--limite", "3")) == 1
    assert len(fab.enviados) == 1
    csv_texto = next(archivos.salida.glob("resumen_*.csv")).read_text(encoding="utf-8-sig")
    assert csv_texto.count(";correo;ok;") == 1 and csv_texto.count(";correo;fallida;") == 2


def test_main_sin_pendientes_no_pide_claves(archivos, sin_efectos, capsys):
    assert nm.main(archivos.argv("--empresa", "empresa que no existe")) == 0
    assert "No hay nada pendiente" in capsys.readouterr().out


def test_main_mapear_sin_pendientes_no_pide_clave(archivos, sin_efectos, capsys):
    assert nm.main(archivos.argv("--mapear")) == 0
    assert "ya están en el mapeo" in capsys.readouterr().out


def test_main_mapear_con_limite_y_omitida(archivos, monkeypatch, tiempos_rapidos, capsys):
    _sin_lineas(archivos.mapeo, "3;", "11;")
    monkeypatch.setattr(nm, "lanzar_navegador", _lanzador(PortalFalso().page))
    monkeypatch.setattr(nm, "BotSernac", BuscadorFalso)
    _claves(monkeypatch, [])
    monkeypatch.setattr(builtins, "input", PreguntarGuionado(["s"]))
    assert nm.main(archivos.argv("--mapear", "--limite", "1")) == 0
    assert "1 omitidas" in capsys.readouterr().out
    assert nm.Mapeo.cargar(archivos.mapeo).entradas(fila_de("x", numero=3)) is None


@pytest.mark.parametrize(
    "extra, esperado, no_esperado",
    [
        (("--solo", "correo"), "no se procesa (--solo correo)", None),
        (("--solo", "sernac"), "no se procesa (--solo sernac)", "Vista previa"),
        (("--limite", "1"), "fuera de --limite", None),
    ],
)
def test_main_dry_run_variantes_de_filtro(archivos, sin_efectos, capsys, extra, esperado, no_esperado):
    assert nm.main(archivos.argv("--dry-run", *extra)) == 0
    salida = capsys.readouterr().out
    assert esperado in salida
    if no_esperado:
        assert no_esperado not in salida


def test_main_dry_run_informa_problemas_de_fila(archivos, sin_efectos, capsys):
    wb = openpyxl.load_workbook(archivos.tracker)
    wb["Empresas"]["F2"] = "no-es-correo"
    wb["Empresas"]["O3"] = "=TODAY()"
    wb.save(archivos.tracker)
    assert nm.main(archivos.argv("--dry-run")) == 0
    salida = capsys.readouterr().out
    assert "falta correo válido de la empresa" in salida
    assert "fórmula en No Molestar ingresado" in salida


def test_main_salida_que_es_un_archivo_sale_con_2(archivos, sin_efectos, capsys, tmp_path):
    archivo = tmp_path / "ocupado"
    archivo.write_text("x")
    argv = archivos.argv("--dry-run")
    argv[argv.index("--salida") + 1] = str(archivo)
    assert nm.main(argv) == 2


def test_main_error_inesperado_sale_con_1_y_lo_registra(archivos, sin_efectos, monkeypatch, capsys):
    monkeypatch.setattr(nm, "_dry_run", lambda *a, **k: (_ for _ in ()).throw(RuntimeError("bug")))
    assert nm.main(archivos.argv("--dry-run")) == 1
    assert "RuntimeError" in next(archivos.salida.glob("*.log")).read_text(encoding="utf-8")


def test_main_sin_playwright_instalado_sale_con_2(archivos, monkeypatch, capsys):
    def sin_playwright(*a, **k):
        raise ImportError("No module named 'playwright'")

    monkeypatch.setattr(nm, "lanzar_navegador", sin_playwright)
    assert nm.main(archivos.argv("--explorar")) == 2
    assert "playwright install" in capsys.readouterr().err
