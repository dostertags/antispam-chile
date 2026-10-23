#!/usr/bin/env python3
"""no_molestar.py — Solicitudes «No Molestar» en el Portal del Consumidor (SERNAC) y correos de
oposición a empresas, a partir del tracker Excel `tracker_no_molestar_50_empresas.xlsx`.

Uso rápido:
    python no_molestar.py --dry-run      # valida todo y muestra el plan, sin enviar nada
    python no_molestar.py --mapear       # asocia cada empresa con sus entradas del directorio SERNAC
    python no_molestar.py --explorar     # recorre el portal contigo para grabar selectores
    python no_molestar.py                # ejecuta (pide ClaveÚnica y/o clave de aplicación Gmail)

Ver README.md.
"""
import argparse
import contextlib
import csv
import getpass
import io
import json
import logging
import math
import os
import queue
import random
import re
import shutil
import smtplib
import socket
import ssl
import sys
import threading
import time
import tomllib
import unicodedata
import warnings
from dataclasses import dataclass
from datetime import date, datetime
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid
from functools import lru_cache
from pathlib import Path
from typing import Any, Callable, Iterable, Iterator, Mapping, Sequence
from urllib.parse import urlsplit

import openpyxl
from openpyxl.styles.numbers import is_date_format

# ====================== CONFIGURACIÓN (edita solo este bloque) ======================
PORTAL_URL = "https://www.sernac.cl/app/consumidor/"
URL_NO_MOLESTAR = "https://www.sernac.cl/app/consumidor/index.php?c=home&a=nomolestar"
HOST_CLAVEUNICA = "accounts.claveunica.gob.cl"  # la clave solo se escribe en https://<este host>
SELECTORES_VERIFICADOS = True    # grabados en el portal real y confirmados con un envío real el 2026-10-23
                                 # (salvo "cu_error"); en False pregunta antes de cada envío y guarda diagnóstico
FORMATO_RUN = "solo_digitos"     # ClaveÚnica formatea sola: se tipea 123456785

# Cada clave: lista de selectores Playwright alternativos, probados en orden.
# Ámbito: P = página del portal; CU = página de ClaveÚnica; F = dentro del iframe "marco_no_molestar".
SELECTORES: dict[str, list[str]] = {
    "login_boton_claveunica": ["#btnClaveUnica"],  # P
    "cu_run": ["#uname"],  # CU
    "cu_clave": ["#pword"],  # CU
    "cu_ingresar": ["#login-submit"],  # CU: parte deshabilitado; se habilita con keyup en #uname/#pword
    "cu_error": [  # CU  # SIN VERIFICAR
        ".alert-danger",  # SIN VERIFICAR
        ".errorlist",  # SIN VERIFICAR
        "text=/(RUN|ClaveÚnica|clave).{0,40}(incorrect|inv[aá]lid|no coincide)|bloquead/i",  # SIN VERIFICAR
    ],
    "sesion_iniciada": ["button[onclick='logoutCU()']"],  # P: basta que exista (hay copias ocultas)
    "boton_mantener_sesion": ["#sessionTimeoutModal #btnKeepAlive"],  # P: clic solo si está visible
    "acepto_datos": ["#aceptoNoMolestarCheck"],  # P: debe quedar marcado
    "marco_no_molestar": ["#frameNoMolestar"],  # P: iframe de la app No Molestar
    "boton_nueva_solicitud": ["a.btn-primary:has-text('Ingresar nuevo')"],  # F
    "radio_telefono": ["#customRadioInline1"],  # F
    "radio_correo": ["#customRadioInline2"],  # F
    "tipo_telefono": ["select#tipo-telefono"],  # F: etiquetas "Móvil" / "Fijo"
    "campo_telefono": ["input#telefono-sub"],  # F: 9 dígitos nacionales (el +56 es fijo)
    "campo_correo_usuario": ["input#mail-pre"],  # F: parte antes de "@"
    "campo_correo_dominio": ["input#mail-sub"],  # F: parte después de "@"
    "boton_agregar_contacto": ["#ingresar-bloqueo"],  # F
    "lista_telefonos": [  # F: texto "+56XXXXXXXXX"
        "#lista-telefonos-bloqueos li.telefonos",
        "li.telefonos",  # respaldo si cambia el id del contenedor
        ".telefonos",  # respaldo si cambia también la etiqueta
    ],
    "lista_correos": ["#lista-mails-bloqueos li.correos"],  # F
    "error_contacto": ["#mensajes-error-telefonos > .alert:not(.d-none)"],  # F
    "select_empresa": ["select#empresas"],  # F: <select> oculto; input_value() = id SERNAC elegido
    "abrir_buscador_empresa": [  # F
        "select#empresas + .select2 .select2-selection",
        "[aria-labelledby='select2-empresas-container']",
    ],
    "campo_buscar_empresa": [".select2-container--open .select2-search__field"],  # F
    "opciones_empresa": ["#select2-empresas-results li.select2-results__option[role='option']"],  # F
    "boton_agregar_empresa": ["#ingresar-empresa"],  # F
    "lista_empresas": ["#lista-empresas li.empresas"],  # F
    "error_empresa": ["#error-empresas:not(.d-none)"],  # F
    "boton_continuar": ["#fake-continuar"],  # F: carga «Revisar solicitud» (GET, NO envía)
    "canales": [  # F: Llamada/Sms/Whatsapp por teléfono y "Bloquear" por correo (vienen marcadas)
        "input.form-check-input[id*='-empresa_canal_ids_']",
        "input.form-check-input[id$='-empresa]']",
    ],
    "boton_confirmar_telefonos": ["#pills-telefonos button[id^='confirmar-']"],  # F: solo cambia a la pestaña correo
    "boton_enviar": ["#finalizar"],  # F: POST /no-molestar/solicitudes — PUNTO SIN RETORNO
    "confirmacion": [  # F: «Su solicitud ha sido ingresada y enviada a la(s) empresa(s)…» (sin N°)
        "#messages-remote-block .alert-success",
        ".alert-success",
        "text=/solicitud ha sido ingresada/i",
    ],
}

# N° de solicitud / folio en el texto de confirmación (se aplica sin distinguir mayúsculas). El N° debe
# tener al menos un dígito: así «solicitud No Molestar» o un encabezado «N° Solicitud | Fecha» no cuentan.
REGEX_NUMERO_SOLICITUD = (
    r"(?:N[°ºo]\.?\s*(?:de\s+)?(?:solicitud|folio)|n[uú]mero\s+de\s+(?:solicitud|folio)"
    r"|c[oó]digo\s+de\s+solicitud|solicitud\s+N[°ºo]\.?|folio)"
    r"\s*[:#]?\s*((?=[A-Z0-9-]*\d)[A-Z0-9][A-Z0-9-]{3,})"
)

# Clave lógica -> encabezados aceptados en la hoja "Empresas" (comparación normalizada).
COLUMNAS: dict[str, list[str]] = {
    # Entrada
    "numero": ["#", "N°", "Nro"],
    "empresa": ["Marca", "Empresa", "Nombre empresa"],
    "razon_social": ["Razón social"],
    "rut_empresa": ["RUT", "RUT empresa"],
    "correo_empresa": ["Correo datos personales", "Correo empresa", "Email empresa"],
    "via": ["Vía de envío", "Via de envio"],
    "formulario_web": ["Formulario web"],
    "canal_alternativo": ["Canal alternativo"],
    # Salida (deben existir; el script nunca crea columnas)
    "fecha_correo": ["Enviado (fecha)", "Fecha correo"],
    "fecha_sernac": ["No Molestar ingresado", "Fecha SERNAC"],
    "numero_sernac": ["N° solicitud SERNAC", "Numero solicitud SERNAC"],
}

# Pausas (min, max) en segundos, timeouts y reintentos
TIEMPOS: dict[str, Any] = {
    "pausa_entre_empresas": (8, 20),
    "pausa_entre_correos": (15, 40),
    "pausa_entre_acciones": (0.4, 1.2),
    "pausa_entre_busquedas": (1.0, 2.0),
    "timeout_elemento_ms": 15000,
    "timeout_login_s": 300,
    "timeout_confirmacion_ms": 30000,
    "retardo_tecleo_ms": 70,
    "reintentos": 3,
    "backoff_base_s": 4,
}

SMTP_HOST, SMTP_PUERTO = "smtp.gmail.com", 465
HOJA = "Empresas"
# ====================================================================================

# ---------------------------------------------------------------- constantes internas
NOMBRE_LOGGER = "no_molestar"
DIR_SCRIPT = Path(__file__).resolve().parent
ARCHIVO_TRACKER = "tracker_no_molestar_50_empresas.xlsx"
ARCHIVO_CONFIG = "config.toml"
ARCHIVO_PLANTILLA = "plantilla.txt"
ARCHIVO_MAPEO = "mapeo_sernac.csv"
ASUNTO_POR_DEFECTO = "Solicitud de suspensión de comunicaciones publicitarias – {nombre}"
FASES = ("sernac", "correo")
ESTADOS = frozenset({"ok", "fallida", "incierta", "sin_coincidencia", "omitida", "ya_hecha"})
VIAS_CORREO = frozenset({"correo", "email", "e mail", "mail"})
VIAS_MANUALES = frozenset({"formulario web", "canal alternativo"})
COLUMNAS_REQUERIDAS = ("empresa", "via")
COLUMNAS_SALIDA = ("fecha_correo", "fecha_sernac", "numero_sernac")
FORMATO_FECHA_EXCEL = "dd-mm-yyyy"
TRACKER_PLANTILLA = "tracker_plantilla_50_empresas.xlsx"  # tracker limpio que se publica con el script
DETALLE_SIN_NUMERO = "SERNAC no muestra N°; llega en su correo de respaldo"
DETALLE_SIN_ENTRADA = "sin entrada en SERNAC"
DETALLE_SIN_CORREO = "Falta correo válido de la empresa"
ENCABEZADOS_MAPEO = ("#", "Marca", "ID SERNAC", "Entrada SERNAC", "Nota")
STOPWORDS_MARCA = frozenset({"de", "del", "la", "las", "los", "el", "y", "chile", "s", "a", "sa", "spa", "ltda", "cia"})
MAX_OPCIONES_MAPEO = 25
MAX_OPCIONES_SIN_COINCIDENCIA = 50
PALABRAS_SECRETAS = (  # aparecen dentro del nombre de la clave
    "clave", "password", "contraseña", "contrasena", "pass", "app_password", "token", "secret", "pwd",
    "api_key", "apikey", "llave",
)
PARTES_SECRETAS = frozenset({"pin", "key"})  # solo como palabra completa o parte separada por "_"
RESERVADOS_WINDOWS = frozenset(
    {"con", "prn", "aux", "nul"} | {f"com{i}" for i in range(1, 10)} | {f"lpt{i}" for i in range(1, 10)}
)

# Detección de CAPTCHA VISIBLE (el script nunca lo resuelve: pausa y espera a la persona).
# NO incluir ".grecaptcha-badge" ni iframes con size=invisible (reCAPTCHA v3/Enterprise invisible:
# en el portal real ese iframe mide 256x60 y está visible).
SELECTORES_CAPTCHA: list[str] = [
    "iframe[src*='recaptcha'][src*='anchor']:not([src*='size=invisible'])",  # reCAPTCHA v2 (casilla)
    "iframe[src*='recaptcha'][src*='bframe']",  # desafío de imágenes (cuenta solo si tiene tamaño)
    ".h-captcha",
    "iframe[src*='hcaptcha']",
    "iframe[src*='challenges.cloudflare.com']",  # Cloudflare Turnstile
    "text=/No soy un robot/i",
]
# Detección de segundo factor (2FA)
SELECTORES_2FA: list[str] = ["input[autocomplete='one-time-code']"]
REGEX_2FA = (
    r"c[oó]digo de (verificaci[oó]n|seguridad)|segundo factor|doble factor"
    r"|autenticaci[oó]n de dos|ingresa el c[oó]digo"
)
# Token JWT del portal en URLs (codifica nombre y correo de la persona)
REGEX_TOKEN_PORTAL = r"([?&]p=)[A-Za-z0-9._-]+"

# select2: indicadores internos de «buscando…» y «sin resultados» (no son parte de la config)
SELECTOR_CARGANDO_EMPRESAS = "#select2-empresas-results li.loading-results"
SELECTOR_SIN_RESULTADOS = "#select2-empresas-results li.select2-results__message"
REGEX_TEXTO_CARGANDO = re.compile(r"^(?:searching|buscando|cargando|loading)\b", re.I)
REGEX_TEXTO_SIN_RESULTADOS = re.compile(
    r"^(?:no se encontraron resultados|no results found|no hay resultados|sin resultados)\.?$", re.I
)
# select2 «errorLoading» (falló el AJAX del buscador): NO es «sin resultados»
REGEX_TEXTO_ERROR_BUSCADOR = re.compile(r"could not be loaded|no se pudieron cargar|no se pudo cargar|\berror\b", re.I)
# Confirmación del envío: un aviso con estas palabras NO es un éxito (se pregunta a la persona)
REGEX_CONFIRMACION_NEGATIVA = re.compile(
    r"\bno\b(?!\s*(?:molestar\b|[.:°º#]|\d))|\berror|rechaz|inv[aá]lid|\bfall[oóa]|incorrect|problema", re.I
)
SELECTOR_ALERTA_ERROR = ".alert-danger"  # aviso de error en el iframe (impide dar el envío por exitoso)
INTERVALO_SONDEO_MS = 250
MAX_INSTANTANEAS = 500

MENSAJE_AUTH_GMAIL = (
    "Gmail rechazó el usuario o la contraseña de aplicación. Usa una «contraseña de aplicación» de 16 letras "
    "(https://myaccount.google.com/apppasswords; requiere verificación en 2 pasos), no tu contraseña normal."
)

SALIDA_OK = 0
SALIDA_FALLAS = 1
SALIDA_VALIDACION = 2
SALIDA_ABORTO = 130

Pagina = Any  # playwright.sync_api.Page (playwright se importa de forma diferida)
Preguntar = Callable[[str], str]
Dormir = Callable[[float], Any]

log = logging.getLogger(NOMBRE_LOGGER)


# ----------------------------------------------------------------------- excepciones
class ErrorNoMolestar(Exception):
    """Base de los errores propios del script."""


class ErrorConfig(ErrorNoMolestar):
    """config.toml ausente, ilegible o con datos inválidos (el mensaje lista todos los problemas)."""


class ErrorPlantilla(ErrorNoMolestar):
    """plantilla.txt ausente o con marcadores desconocidos (el mensaje lista los nombres)."""


class ErrorTracker(ErrorNoMolestar):
    """Tracker ausente, sin la hoja, sin encabezados/columnas requeridas o celda de salida con fórmula."""


class ErrorMapeo(ErrorNoMolestar):
    """mapeo_sernac.csv mal formado (columnas faltantes, # o ID SERNAC no numéricos)."""


class ErrorPortal(ErrorNoMolestar):
    """El portal respondió 403/5xx o no se pudo abrir (mantención o bloqueo de red)."""


class ErrorLogin(ErrorNoMolestar):
    """El inicio de sesión con ClaveÚnica falló. Nunca se reintenta automáticamente."""


class ErrorFormulario(ErrorNoMolestar):
    """El formulario No Molestar rechazó un dato (se informa el mensaje del portal; no se reintenta)."""


class AbortoUsuario(Exception):
    """La persona pidió abortar (escribió «abortar», Ctrl+C o fin de entrada).
    `resultado`: el de la empresa en curso si ya hubo un envío sin confirmar (va al resumen)."""

    resultado: "Resultado | None" = None


class EmpresaOmitida(Exception):
    """La persona omitió todas las entradas SERNAC de una empresa (sin coincidencia)."""


@lru_cache(maxsize=1)
def _errores_playwright() -> tuple[type[BaseException], ...]:
    """Clases de error de Playwright (importación diferida; vacío si no está instalado)."""
    try:
        from playwright.sync_api import Error as ErrorPlaywright
    except ImportError:  # pragma: no cover - depende del entorno
        return ()
    return (ErrorPlaywright,)


def _timeout_playwright(mensaje: str) -> BaseException:
    """Un TimeoutError de Playwright (transitorio: lo reintenta con_reintentos)."""
    try:
        from playwright.sync_api import TimeoutError as TimeoutPlaywright
    except ImportError:  # pragma: no cover - depende del entorno
        return TimeoutError(mensaje)
    return TimeoutPlaywright(mensaje)


# ------------------------------------------------------------------------- secretos
class Secreto:
    """Envuelve una contraseña. repr/str -> "Secreto('***')"; el valor solo sale con revelar().

    Igualdad por identidad (no compara valores); bool() indica si no está vacío.
    """

    __slots__ = ("_valor",)

    def __init__(self, valor: str) -> None:
        self._valor = str(valor)

    def revelar(self) -> str:
        """Devuelve el valor real (usar solo al tipear el campo o hacer login SMTP)."""
        return self._valor

    def __repr__(self) -> str:
        return "Secreto('***')"

    def __str__(self) -> str:
        return "Secreto('***')"

    def __bool__(self) -> bool:
        return bool(self._valor)

    def __reduce__(self) -> Any:
        raise TypeError("Un Secreto no se puede serializar")


def _sin_query(url: str) -> str:
    """URL sin su query string (para registrar URLs del portal sin el token p=)."""
    url = str(url or "")
    return url.split("?", 1)[0] + "?<omitido>" if "?" in url else url


class FiltroSecretos(logging.Filter):
    """Reemplaza por "***" cada secreto (y su variante sin espacios, si tiene ≥4 caracteres) y el
    token del portal (REGEX_TOKEN_PORTAL -> "\\1***") en record.msg, record.args y el texto de la
    excepción (deja record.exc_text ya censurado). Nunca descarta registros: filter() -> True."""

    def __init__(self, secretos: Iterable[Secreto] = ()) -> None:
        super().__init__()
        self._valores: list[str] = []
        for secreto in secretos:
            self.agregar(secreto)

    def agregar(self, secreto: Secreto) -> None:
        """Agrega un secreto a censurar (p. ej. después de pedirlo con getpass)."""
        valor = secreto.revelar() if isinstance(secreto, Secreto) else str(secreto)
        sin_espacios = re.sub(r"\s+", "", valor)
        candidatos = [valor] + ([sin_espacios] if len(sin_espacios) >= 4 else [])
        for candidato in candidatos:
            if candidato and candidato not in self._valores:
                self._valores.append(candidato)
        self._valores.sort(key=len, reverse=True)

    def censurar(self, texto: str) -> str:
        """Texto con los secretos y los tokens p= del portal reemplazados por "***"."""
        texto = str(texto)
        for valor in self._valores:
            texto = texto.replace(valor, "***")
        return re.sub(REGEX_TOKEN_PORTAL, r"\1***", texto)

    def filter(self, record: logging.LogRecord) -> bool:
        """Censura el registro en el lugar y devuelve True."""
        try:
            mensaje = record.getMessage()
        except Exception:  # argumentos que no calzan con el formato
            mensaje = f"{record.msg} {record.args}"
        record.msg = self.censurar(mensaje)
        record.args = None
        if record.exc_info:
            record.exc_text = self.censurar(logging.Formatter().formatException(record.exc_info))
        elif record.exc_text:
            record.exc_text = self.censurar(record.exc_text)
        if record.stack_info:
            record.stack_info = self.censurar(record.stack_info)
        return True


# ---------------------------------------------------------------------- datos (puros)
@dataclass(frozen=True)
class Persona:
    """Datos personales validados. rut se guarda como "12345678-5" (sin puntos, DV en
    mayúscula) y telefono como 9 dígitos nacionales "912345678"."""

    nombre: str
    rut: str
    telefono: str
    correo: str

    @property
    def rut_formateado(self) -> str:
        """RUT con puntos: "12.345.678-5"."""
        return formatear_rut(self.rut, "puntos")

    @property
    def telefono_e164(self) -> str:
        """Teléfono E.164: "+56912345678"."""
        return f"+56{self.telefono}"

    @property
    def telefono_legible(self) -> str:
        """Teléfono legible: "+56 9 1234 5678"."""
        t = self.telefono
        return f"+56 {t[0]} {t[1:5]} {t[5:]}"

    @property
    def correo_usuario(self) -> str:
        """Parte del correo antes de "@" (campo #mail-pre del portal)."""
        return self.correo.rsplit("@", 1)[0]

    @property
    def correo_dominio(self) -> str:
        """Parte del correo después de "@" (campo #mail-sub del portal)."""
        return self.correo.rsplit("@", 1)[1]


@dataclass(frozen=True)
class Config:
    """Configuración personal. gmail_usuario = [gmail] usuario o, si falta, persona.correo."""

    persona: Persona
    gmail_usuario: str


@dataclass
class Fila:
    """Una empresa de la hoja "Empresas". numero_fila = fila real de Excel; numero = columna "#"
    (None si no existe). Los campos de texto son str ("" si la celda está vacía o tiene fórmula).
    valores: encabezado ORIGINAL -> valor de celda (fórmulas como ""). sernac_hecho: «No Molestar
    ingresado» o «N° solicitud SERNAC» no vacíos; correo_hecho: «Enviado (fecha)» no vacío."""

    numero_fila: int
    numero: int | None
    empresa: str
    razon_social: str
    rut_empresa: str
    correo_empresa: str
    via: str
    formulario_web: str
    canal_alternativo: str
    valores: dict[str, Any]
    sernac_hecho: bool = False
    correo_hecho: bool = False

    @property
    def es_correo(self) -> bool:
        """True si «Vía de envío» es correo (es_via_correo)."""
        return es_via_correo(self.via)

    @property
    def etiqueta(self) -> str:
        """Número corto para nombres de archivo y mensajes: "#04" (o la fila de Excel)."""
        return f"{self.numero:02d}" if self.numero is not None else f"{self.numero_fila:02d}"


@dataclass(frozen=True)
class Plantilla:
    """Plantilla de correo: asunto y cuerpo aún con marcadores {clave}."""

    asunto: str
    cuerpo: str


@dataclass(frozen=True)
class EntradaSernac:
    """Entrada del directorio de empresas de SERNAC. id None = desconocido (se elige por texto)."""

    id: int | None
    texto: str


@dataclass
class Resultado:
    """Resultado de una solicitud/correo. fase ∈ FASES; estado ∈ ESTADOS;
    captura = ruta (str) de la captura de pantalla de la página, si la hay."""

    numero_fila: int
    empresa: str
    fase: str
    estado: str
    numero: str | None = None
    detalle: str = ""
    captura: str | None = None


def _resultado(fila: Fila, fase: str, estado: str, **kwargs: Any) -> Resultado:
    return Resultado(fila.numero_fila, fila.empresa, fase, estado, **kwargs)


# ---------------------------------------------------------------- normalización/validación
def normalizar(texto: object) -> str:
    """Minúsculas, sin tildes (ñ->n), todo carácter no alfanumérico -> un espacio, espacios
    colapsados y sin espacios extremos. None -> "". Ej.: "  Vía de  Envío " -> "via de envio"."""
    if texto is None:
        return ""
    descompuesto = unicodedata.normalize("NFKD", str(texto))
    sin_tildes = "".join(c for c in descompuesto if not unicodedata.combining(c)).lower()
    return re.sub(r"[^a-z0-9]+", " ", sin_tildes).strip()


def normalizar_encabezado(texto: object) -> str:
    """Clave de comparación de encabezados: insensible a mayúsculas, tildes, espacios y
    puntuación; "N°", "Nº", "No", "No." equivalen entre sí; "#" da una clave no vacía.
    None/"" -> ""."""
    if texto is None:
        return ""
    s = str(texto).strip()
    if not s:
        return ""
    s = s.replace("#", " numero ")
    s = re.sub(r"(?i)\bn(?:\s*[°º]|o\b\.?|ro\b\.?)", " numero ", s)
    return normalizar(s).replace(" ", "")


def es_via_correo(valor: object) -> bool:
    """True si normalizar(valor) ∈ VIAS_CORREO ("correo", "email", "e mail", "mail")."""
    return normalizar(valor) in VIAS_CORREO


def _partes_rut(rut: str) -> tuple[str, str]:
    """("12345678", "5") desde cualquier formato de RUT; ValueError si es ilegible."""
    limpio = re.sub(r"[.\s-]", "", str(rut or "")).upper()
    if not re.fullmatch(r"\d{1,8}[\dK]", limpio):
        raise ValueError(f"RUT ilegible: {rut!r}")
    return limpio[:-1], limpio[-1]


def _digito_verificador(cuerpo: str) -> str:
    suma, factor = 0, 2
    for digito in reversed(cuerpo):
        suma += int(digito) * factor
        factor = 2 if factor == 7 else factor + 1
    resto = 11 - suma % 11
    return {10: "K", 11: "0"}.get(resto, str(resto))


def validar_rut(rut: str) -> bool:
    """Valida RUT chileno por módulo 11. Acepta "12.345.678-5", "12345678-5", "123456785", k/K."""
    try:
        cuerpo, dv = _partes_rut(rut)
    except ValueError:
        return False
    return _digito_verificador(cuerpo) == dv


def formatear_rut(rut: str, estilo: str = "puntos") -> str:
    """estilo "puntos" -> "12.345.678-5"; "sin_puntos" -> "12345678-5"; "solo_digitos" ->
    "123456785". DV siempre en mayúscula. Estilo desconocido o RUT ilegible -> ValueError."""
    cuerpo, dv = _partes_rut(rut)
    if estilo == "puntos":
        return f"{int(cuerpo):,}".replace(",", ".") + f"-{dv}"
    if estilo == "sin_puntos":
        return f"{cuerpo}-{dv}"
    if estilo == "solo_digitos":
        return f"{cuerpo}{dv}"
    raise ValueError(f"Estilo de RUT desconocido: {estilo!r} (usa puntos, sin_puntos o solo_digitos)")


def normalizar_telefono(texto: str) -> str:
    """Teléfono chileno -> 9 dígitos nacionales ("912345678", "223456789"). Quita espacios,
    guiones, paréntesis, "+" y el prefijo 56. Si no quedan exactamente 9 dígitos -> ValueError."""
    limpio = re.sub(r"[\s()+.-]", "", str(texto or ""))
    if not limpio.isdigit():
        raise ValueError(f"Teléfono con caracteres no válidos: {texto!r}")
    if len(limpio) == 11 and limpio.startswith("56"):
        limpio = limpio[2:]
    if len(limpio) != 9 or limpio[0] not in "23456789":
        raise ValueError(f"Teléfono chileno inválido (se esperan 9 dígitos, p. ej. 9 1234 5678): {texto!r}")
    return limpio


def validar_correo(texto: str) -> bool:
    """Validación básica (usuario@dominio.tld, sin espacios ni saltos de línea en ninguna parte)."""
    return bool(re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s.]+", str(texto or "")))


# ------------------------------------------------------------------------ entradas
ALIAS_PERSONA: dict[str, tuple[str, ...]] = {
    "nombre": ("nombre", "nombre completo", "name", "full name"),
    "rut": ("rut", "run"),
    "telefono": ("telefono", "fono", "celular", "movil", "phone", "tel"),
    "correo": ("correo", "email", "e mail", "mail", "correo electronico"),
}
ALIAS_GMAIL_USUARIO = ("usuario", "user", "correo", "email", "mail")


def _es_clave_secreta(clave: str) -> bool:
    """True si el nombre parece de un secreto: contiene PALABRAS_SECRETAS (clave, password, secret,
    pwd, token, api_key…) o tiene una parte PARTES_SECRETAS (pin, key) como palabra completa."""
    normal = normalizar(clave).replace(" ", "_")
    if any(normalizar(p).replace(" ", "_") in normal for p in PALABRAS_SECRETAS):
        return True
    return any(parte in PARTES_SECRETAS for parte in normalizar(clave).split())


def _quitar_secretos(tabla: Mapping[str, Any], ruta: str) -> dict[str, Any]:
    """Copia de la tabla sin las claves con aspecto de secreto (avisa sin mostrar el valor)."""
    limpia = {}
    for clave, valor in tabla.items():
        if _es_clave_secreta(clave):
            log.warning(
                "config.toml: se ignoró la clave «%s%s»; las claves se piden al ejecutar y nunca se guardan.",
                ruta, clave,
            )
            continue
        limpia[clave] = valor
    return limpia


def _avisar_secretos(tabla: Mapping[str, Any], ruta: str) -> None:
    """Avisa (sin mostrar valores) de claves con aspecto de secreto en una sección que el script no
    usa, incluidas sus subtablas. No devuelve nada: esas secciones se ignoran completas."""
    for clave, valor in _quitar_secretos(tabla, ruta).items():
        if isinstance(valor, dict):
            _avisar_secretos(valor, f"{ruta}{clave}.")


def _buscar_alias(datos: Mapping[str, Any], alias: Iterable[str]) -> str:
    buscados = {normalizar(a) for a in alias}
    for clave, valor in datos.items():
        if not isinstance(valor, dict) and normalizar(clave) in buscados:
            return str(valor).strip()
    return ""


# Datos ficticios de config.example.toml: válidos a propósito, así que se rechazan explícitamente
RUT_DE_EJEMPLO = "11111111-1"
CORREO_DE_EJEMPLO = "maria.perez@example.cl"


def cargar_config(path: Path) -> Config:
    """Lee config.toml (valores en [persona] o en la raíz, con alias; [gmail] usuario opcional).
    Claves con aspecto de secreto se ignoran con advertencia (sin mostrar su valor).
    Problemas -> ErrorConfig con TODOS los problemas listados."""
    path = Path(path)
    try:
        crudo = tomllib.loads(path.read_text(encoding="utf-8-sig"))
    except FileNotFoundError:
        raise ErrorConfig(f"No existe {path}. Copia config.example.toml como config.toml y completa tus datos.") from None
    except (OSError, UnicodeDecodeError) as exc:
        raise ErrorConfig(f"No pude leer {path}: {exc}") from None
    except tomllib.TOMLDecodeError as exc:
        raise ErrorConfig(f"{path.name} no es TOML válido: {exc}") from None

    raiz = _quitar_secretos({k: v for k, v in crudo.items() if not isinstance(v, dict)}, "")
    secciones = {normalizar(k): v for k, v in crudo.items() if isinstance(v, dict)}
    datos = dict(raiz)
    if "persona" in secciones:
        datos.update(_quitar_secretos(secciones["persona"], "persona."))
    gmail = _quitar_secretos(secciones.get("gmail", {}), "gmail.")
    for nombre, tabla in crudo.items():  # secciones que el script no usa: solo se avisa
        if isinstance(tabla, dict) and normalizar(nombre) not in ("persona", "gmail"):
            _avisar_secretos(tabla, f"{nombre}.")

    problemas: list[str] = []
    nombre = _buscar_alias(datos, ALIAS_PERSONA["nombre"])
    if not nombre:
        problemas.append("nombre: falta o está vacío")
    rut_texto = _buscar_alias(datos, ALIAS_PERSONA["rut"])
    rut = ""
    if not rut_texto:
        problemas.append("rut: falta")
    elif not validar_rut(rut_texto):
        problemas.append(f"rut: «{rut_texto}» no es un RUT válido (revisa el dígito verificador)")
    else:
        rut = formatear_rut(rut_texto, "sin_puntos")
    telefono = ""
    telefono_texto = _buscar_alias(datos, ALIAS_PERSONA["telefono"])
    try:
        telefono = normalizar_telefono(telefono_texto)
    except ValueError:
        problemas.append(f"teléfono: «{telefono_texto}» no es un teléfono chileno de 9 dígitos (ej.: +56 9 1234 5678)")
    correo = _buscar_alias(datos, ALIAS_PERSONA["correo"])
    if not validar_correo(correo):
        problemas.append(f"correo: «{correo}» no es un correo válido")
    gmail_usuario = _buscar_alias(gmail, ALIAS_GMAIL_USUARIO)
    if gmail_usuario and not validar_correo(gmail_usuario):
        problemas.append(f"gmail.usuario: «{gmail_usuario}» no es un correo válido")
    if rut == RUT_DE_EJEMPLO or str(correo or "").strip().lower() == CORREO_DE_EJEMPLO:
        problemas.append("todavía tiene el RUT o el correo de ejemplo de config.example.toml: pon tus datos")
    if problemas:
        raise ErrorConfig(f"{path.name} tiene problemas:\n  - " + "\n  - ".join(problemas))
    persona = Persona(nombre=nombre, rut=rut, telefono=telefono, correo=correo)
    return Config(persona=persona, gmail_usuario=gmail_usuario or correo)


def cargar_plantilla(path: Path) -> Plantilla:
    """Lee plantilla.txt (UTF-8, acepta BOM). Primera línea no vacía "Asunto:"/"Subject:" ->
    asunto (sin espacios extremos); el resto, sin líneas vacías iniciales -> cuerpo.
    Sin línea de asunto -> ASUNTO_POR_DEFECTO + advertencia. Archivo ausente -> ErrorPlantilla."""
    path = Path(path)
    try:
        texto = path.read_text(encoding="utf-8-sig")
    except FileNotFoundError:
        raise ErrorPlantilla(f"No existe la plantilla {path}.") from None
    except (OSError, UnicodeDecodeError) as exc:
        raise ErrorPlantilla(f"No pude leer la plantilla {path}: {exc}") from None
    lineas = texto.splitlines()
    while lineas and not lineas[0].strip():
        lineas.pop(0)
    asunto = None
    if lineas:
        m = re.match(r"\s*(?:asunto|subject)\s*:\s*(.*)$", lineas[0], re.I)
        if m:
            asunto = m.group(1).strip()
            lineas.pop(0)
    if asunto is None:
        log.warning("La plantilla no tiene línea «Asunto:»; se usa el asunto por defecto: %s", ASUNTO_POR_DEFECTO)
        asunto = ASUNTO_POR_DEFECTO
    while lineas and not lineas[0].strip():
        lineas.pop(0)
    return Plantilla(asunto=asunto, cuerpo="\n".join(lineas).rstrip() + "\n")


_REGEX_MARCADOR = re.compile(r"\{\{|\}\}|\{([^{}]*)\}")


def placeholders(texto: str) -> set[str]:
    """Nombres de marcadores {clave} tal como están escritos (sin espacios extremos).
    "{{" y "}}" son llaves literales y no cuentan."""
    return {m.group(1).strip() for m in _REGEX_MARCADOR.finditer(texto) if m.group(1) is not None}


def _clave_contexto(nombre: object) -> str:
    return normalizar(nombre).replace(" ", "_")


def _texto_valor(valor: Any) -> str:
    """Valor de celda como texto legible (fechas dd-mm-aaaa, enteros sin ".0")."""
    if valor is None:
        return ""
    if isinstance(valor, datetime):
        return valor.strftime("%d-%m-%Y")
    if isinstance(valor, date):
        return valor.strftime("%d-%m-%Y")
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()


def contexto_para(fila: Fila, persona: Persona, hoy: date) -> dict[str, str]:
    """Contexto de la plantilla con claves normalizar(...) y espacios -> "_". Prioridad (de menor a
    mayor): encabezados del tracker de la fila (claves vacías se omiten; None -> "") < datos de la
    empresa (empresa, razon_social, rut_empresa, correo_empresa) < datos de la persona (nombre,
    rut "12.345.678-5", telefono "+56 9 1234 5678", correo, fecha "dd-mm-aaaa"). Valores str."""
    contexto: dict[str, str] = {}
    for encabezado, valor in fila.valores.items():
        clave = _clave_contexto(encabezado)
        if clave:
            contexto[clave] = _texto_valor(valor)
    contexto.update(
        empresa=fila.empresa,
        razon_social=fila.razon_social,
        rut_empresa=fila.rut_empresa,
        correo_empresa=fila.correo_empresa,
    )
    contexto.update(
        nombre=persona.nombre,
        rut=persona.rut_formateado,
        telefono=persona.telefono_legible,
        correo=persona.correo,
        fecha=hoy.strftime("%d-%m-%Y"),
    )
    return contexto


def renderizar(plantilla: str, contexto: Mapping[str, str]) -> str:
    """Sustituye {clave} con búsqueda normalizada, en una sola pasada con regex (nunca
    str.format: sin acceso a atributos ni índices). "{{"/"}}" -> llaves literales.
    Marcadores desconocidos -> ErrorPlantilla que lista sus nombres."""
    desconocidos: list[str] = []

    def sustituir(m: re.Match) -> str:
        if m.group(0) == "{{":
            return "{"
        if m.group(0) == "}}":
            return "}"
        nombre = m.group(1).strip()
        clave = _clave_contexto(nombre)
        if clave and clave in contexto:
            return str(contexto[clave])
        if nombre not in desconocidos:
            desconocidos.append(nombre)
        return m.group(0)

    resultado = _REGEX_MARCADOR.sub(sustituir, plantilla)
    if desconocidos:
        disponibles = ", ".join(sorted(contexto))
        raise ErrorPlantilla(
            "La plantilla usa marcadores desconocidos: "
            + ", ".join("{" + n + "}" for n in desconocidos)
            + f". Disponibles: {disponibles}"
        )
    return resultado


# ----------------------------------------------------------- directorio SERNAC / números
def _sin_duplicados(textos: Iterable[str]) -> list[str]:
    vistos: set[str] = set()
    salida = []
    for texto in textos:
        texto = re.sub(r"\s+", " ", texto).strip()
        clave = normalizar(texto)
        if clave and clave not in vistos:
            vistos.add(clave)
            salida.append(texto)
    return salida


def candidatos_busqueda(fila: Fila) -> list[str]:
    """Textos de búsqueda para la empresa, sin duplicados (comparación normalizada), en orden:
    Marca sin la parte entre paréntesis (si tiene "/", cada parte por separado), cada parte entre
    paréntesis, y las partes de la Razón social (separadas por "/" o ";"). Sin textos vacíos."""
    marca = fila.empresa or ""
    sin_parentesis = re.sub(r"\([^)]*\)", " ", marca)
    partes = sin_parentesis.split("/")
    for entre_parentesis in re.findall(r"\(([^)]*)\)", marca):
        partes.extend(entre_parentesis.split("/"))
    partes.extend(re.split(r"[/;]", fila.razon_social or ""))
    return _sin_duplicados(partes)


def _terminos_mapeo(fila: Fila) -> list[str]:
    """candidatos_busqueda + la primera palabra (≥4 caracteres) de cada candidato."""
    candidatos = candidatos_busqueda(fila)
    primeras = [c.split()[0] for c in candidatos if c.split() and len(c.split()[0]) >= 4]
    return _sin_duplicados(candidatos + primeras)


def termino_busqueda(texto: str) -> str:
    """Término para el buscador SERNAC: el texto hasta el primer " (" o " - " (sin espacios
    extremos) y luego máximo 40 caracteres (sin espacios extremos). Si queda con < 3 caracteres,
    se usan los primeros 40 caracteres del texto. "CLARO MOVIL (CELULAR…" -> "CLARO MOVIL"."""
    texto = str(texto or "")
    cortes = [i for i in (texto.find(" ("), texto.find(" - ")) if i >= 0]
    base = texto[: min(cortes)] if cortes else texto
    base = base.strip()[:40].strip()
    if len(base) < 3:
        base = texto[:40].strip()
    return base


def puntuar_opciones(fila: Fila, opciones: Sequence[str]) -> list[tuple[int, float]]:
    """[(índice, puntaje)] para TODAS las opciones, ordenado por puntaje descendente (empates por
    índice). Candidatos = candidatos_busqueda(fila) + la primera palabra de cada uno si tiene ≥4
    caracteres, normalizados; tokens de marca = sus palabras menos STOPWORDS_MARCA. Puntaje de la
    opción normalizada: +3 si empieza con algún candidato (prefijo de texto); por cada token: +2 si
    aparece como palabra completa, −2 si solo aparece dentro de una palabra más larga.
    Sugerencia = puntaje ≥ 3."""
    candidatos = [normalizar(c) for c in _terminos_mapeo(fila)]
    tokens: list[str] = []
    for candidato in candidatos:
        for palabra in candidato.split():
            if palabra not in STOPWORDS_MARCA and palabra not in tokens:
                tokens.append(palabra)
    puntajes = []
    for indice, opcion in enumerate(opciones):
        texto = normalizar(opcion)
        palabras = set(texto.split())
        puntaje = 3.0 if any(c and texto.startswith(c) for c in candidatos) else 0.0
        for token in tokens:
            if token in palabras:
                puntaje += 2
            elif token in texto:
                puntaje -= 2
        puntajes.append((indice, puntaje))
    return sorted(puntajes, key=lambda par: (-par[1], par[0]))


def extraer_numero_solicitud(texto: str) -> str | None:
    """N° de solicitud/folio en el texto (REGEX_NUMERO_SOLICITUD, sin distinguir mayúsculas);
    token de ≥4 caracteres [A-Z0-9-] con al menos un dígito. None si no aparece."""
    m = re.search(REGEX_NUMERO_SOLICITUD, str(texto or ""), re.I)
    return m.group(1).upper() if m else None


def slug(texto: str, largo: int = 60) -> str:
    """Nombre de archivo seguro: minúsculas ASCII [a-z0-9-], guiones colapsados, máximo `largo`
    caracteres, nunca vacío ("sin-nombre") ni un nombre reservado de Windows (con, nul, com1...)."""
    s = re.sub(r"-+", "-", normalizar(texto).replace(" ", "-")).strip("-")
    s = s[: max(1, largo)].strip("-") or "sin-nombre"
    if s in RESERVADOS_WINDOWS:
        s = f"{s}-archivo"[:largo] if largo >= len(s) + 2 else "sin-nombre"[:largo]
    return s


# -------------------------------------------------------------- reintentos y pausas
def con_reintentos(
    fn: Callable[[], Any],
    *,
    intentos: int,
    base: float,
    es_transitorio: Callable[[BaseException], bool],
    sleep: Dormir,
    al_reintentar: Callable[[int, BaseException], Any] | None = None,
) -> Any:
    """Llama fn() hasta `intentos` veces. Error no transitorio -> se relanza de inmediato.
    Error transitorio con intentos restantes -> al_reintentar(n_intento_fallido, exc) (n desde 1),
    luego sleep(espera) con espera creciente ≥ base (backoff exponencial); agotados -> relanza."""
    total = max(1, int(intentos))
    for numero in range(1, total + 1):
        try:
            return fn()
        except Exception as exc:
            if numero >= total or not es_transitorio(exc):
                raise
            if al_reintentar is not None:
                al_reintentar(numero, exc)
            espera = base * 2 ** (numero - 1) + (random.uniform(0, base / 4) if base > 0 else 0)
            sleep(espera)
    raise AssertionError("inalcanzable")  # pragma: no cover


def es_transitorio_playwright(exc: BaseException) -> bool:
    """True para TimeoutError/Error de Playwright; False para AbortoUsuario, EmpresaOmitida,
    ErrorLogin, navegador cerrado o caído (_navegador_cerrado) y cualquier otra excepción."""
    if isinstance(exc, (AbortoUsuario, EmpresaOmitida, ErrorNoMolestar)) or _navegador_cerrado(exc):
        return False
    errores = _errores_playwright()
    return bool(errores) and isinstance(exc, errores)


_TEXTOS_NAVEGADOR_CERRADO = (
    "target page, context or browser has been closed",
    "browser has been closed",
    "target closed",
    "target crashed",
)


def _navegador_cerrado(exc: BaseException) -> bool:
    """True si el error de Playwright indica que la ventana/navegador se cerró o se cayó
    (TargetClosedError o su mensaje): no tiene sentido reintentar ni seguir con otras empresas."""
    if any(clase.__name__ == "TargetClosedError" for clase in type(exc).__mro__):
        return True
    errores = _errores_playwright()
    if not (errores and isinstance(exc, errores)):
        return False
    texto = str(exc).lower()
    return any(t in texto for t in _TEXTOS_NAVEGADOR_CERRADO)


def es_transitorio_smtp(exc: BaseException) -> bool:
    """True: SMTPServerDisconnected, SMTPConnectError, TimeoutError/socket.timeout,
    ConnectionError, SMTPResponseException 4xx. False: SMTPAuthenticationError,
    SMTPRecipientsRefused, respuestas 5xx y lo demás."""
    if isinstance(exc, (smtplib.SMTPAuthenticationError, smtplib.SMTPRecipientsRefused)):
        return False
    if isinstance(exc, (smtplib.SMTPServerDisconnected, smtplib.SMTPConnectError)):
        return True
    if isinstance(exc, smtplib.SMTPResponseException):
        return 400 <= int(exc.smtp_code) < 500
    return isinstance(exc, (TimeoutError, socket.timeout, ConnectionError))


def pausa_cortes(rango: tuple[float, float], sleep: Dormir) -> float:
    """Pausa cortés: sleep(random.uniform(min, max)) una vez; devuelve los segundos usados."""
    minimo, maximo = rango
    segundos = random.uniform(minimo, maximo)
    sleep(segundos)
    return segundos


def pausar(motivo: str, preguntar: Preguntar) -> None:
    """Imprime un aviso con "\\a" e instrucciones en español y espera Enter vía preguntar().
    Respuesta "abortar" (sin importar mayúsculas/espacios), EOFError o KeyboardInterrupt ->
    AbortoUsuario."""
    print(f"\a\n{'=' * 70}\nPAUSA: {motivo}\n{'=' * 70}", flush=True)
    try:
        respuesta = preguntar("Presiona Enter para continuar (o escribe «abortar» para terminar): ")
    except (EOFError, KeyboardInterrupt):
        raise AbortoUsuario("Ejecución interrumpida durante una pausa") from None
    if str(respuesta).strip().lower() == "abortar":
        raise AbortoUsuario("La persona escribió «abortar»")


def _pedir(preguntar: Preguntar, prompt: str) -> str:
    """preguntar() con EOF/Ctrl+C convertidos en AbortoUsuario; devuelve la respuesta sin espacios."""
    try:
        return str(preguntar(prompt)).strip()
    except (EOFError, KeyboardInterrupt):
        raise AbortoUsuario("Entrada cerrada o Ctrl+C") from None


# -------------------------------------------------------------------------- tracker
def _es_formula(valor: Any) -> bool:
    return isinstance(valor, str) and valor.startswith("=")


def _vacio(valor: Any) -> bool:
    return valor is None or (isinstance(valor, str) and not valor.strip())


def _texto_celda(valor: Any) -> str:
    return "" if _vacio(valor) or _es_formula(valor) else _texto_valor(valor)


def _entero(valor: Any) -> int | None:
    """Entero si el valor «parece» entero (4, 4.0, " 4 "); None si no."""
    if isinstance(valor, bool):
        return None
    if isinstance(valor, int):
        return valor
    if isinstance(valor, float) and valor.is_integer():
        return int(valor)
    if isinstance(valor, str) and re.fullmatch(r"\s*\d+\s*", valor):
        return int(valor)
    return None


class Tracker:
    """Tracker Excel (openpyxl sin data_only: conserva fórmulas, formatos y validaciones).
    Encabezados = primera de las filas 1–10 con un alias de "empresa". Requeridas: empresa, vía y
    las 3 columnas de salida (nunca se crean). Nada se escribe en disco hasta guardar()."""

    def __init__(self, path: Path, hoja: str = HOJA) -> None:
        """Carga y valida. Errores -> ErrorTracker (lista hojas o encabezados encontrados y nombra
        la columna faltante)."""
        self.path = Path(path)
        if not self.path.exists():
            raise ErrorTracker(
                f"No existe el tracker {self.path}. Copia {TRACKER_PLANTILLA} (está junto al script) "
                f"como {self.path.name} y úsalo."
            )
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                self.wb = openpyxl.load_workbook(self.path)
        except Exception as exc:
            raise ErrorTracker(f"No pude abrir {self.path.name}: {exc}") from None
        if hoja not in self.wb.sheetnames:
            raise ErrorTracker(
                f"{self.path.name} no tiene la hoja «{hoja}». Hojas encontradas: {', '.join(self.wb.sheetnames)}"
            )
        self.ws = self.wb[hoja]
        self.fila_encabezado, self.encabezados = self._detectar_encabezados()
        self.columnas = self._mapear_columnas()

    def _detectar_encabezados(self) -> tuple[int, dict[int, str]]:
        alias_empresa = {normalizar_encabezado(a) for a in COLUMNAS["empresa"]}
        for numero in range(1, min(10, self.ws.max_row) + 1):
            celdas = {c.column: c.value for c in self.ws[numero] if isinstance(c.value, str) and c.value.strip()}
            if any(normalizar_encabezado(v) in alias_empresa for v in celdas.values()):
                return numero, {col: v.strip() for col, v in celdas.items()}
        raise ErrorTracker(
            f"No encontré la fila de encabezados en la hoja «{self.ws.title}» (una celda «Marca» o «Empresa» "
            "en las filas 1 a 10)."
        )

    def _mapear_columnas(self) -> dict[str, int]:
        por_clave = {}
        for col, encabezado in self.encabezados.items():
            por_clave.setdefault(normalizar_encabezado(encabezado), col)
        columnas: dict[str, int] = {}
        for clave, alias in COLUMNAS.items():
            for a in alias:
                col = por_clave.get(normalizar_encabezado(a))
                if col is not None:
                    columnas[clave] = col
                    break
        faltantes = [COLUMNAS[c][0] for c in COLUMNAS_REQUERIDAS + COLUMNAS_SALIDA if c not in columnas]
        if faltantes:
            raise ErrorTracker(
                f"A la hoja «{self.ws.title}» le faltan columnas: "
                + ", ".join(f"«{f}»" for f in faltantes)
                + " (el script nunca crea columnas). Encabezados encontrados: "
                + ", ".join(self.encabezados.values())
            )
        return columnas

    def _valor(self, numero_fila: int, clave: str) -> Any:
        col = self.columnas.get(clave)
        return None if col is None else self.ws.cell(row=numero_fila, column=col).value

    def filas(self) -> list[Fila]:
        """Filas de datos bajo los encabezados: se detiene en la primera fila sin empresa (aunque su
        "#" venga prenumerado); omite filas cuyo "#" no sea entero (el bloque «Resumen» nunca se lee)."""
        salida = []
        tiene_numero = "numero" in self.columnas
        for r in range(self.fila_encabezado + 1, self.ws.max_row + 1):
            empresa = _texto_celda(self._valor(r, "empresa"))
            if not empresa:
                break
            numero = _entero(self._valor(r, "numero")) if tiene_numero else None
            if tiene_numero and numero is None:
                continue
            texto = {clave: _texto_celda(self._valor(r, clave)) for clave in COLUMNAS}
            valores = {}
            for col, encabezado in self.encabezados.items():
                v = self.ws.cell(row=r, column=col).value
                valores[encabezado] = "" if _es_formula(v) else v
            salida.append(
                Fila(
                    numero_fila=r,
                    numero=numero,
                    empresa=empresa,
                    razon_social=texto["razon_social"],
                    rut_empresa=texto["rut_empresa"],
                    correo_empresa=texto["correo_empresa"],
                    via=texto["via"],
                    formulario_web=texto["formulario_web"],
                    canal_alternativo=texto["canal_alternativo"],
                    valores=valores,
                    sernac_hecho=bool(texto["fecha_sernac"] or texto["numero_sernac"]),
                    correo_hecho=bool(texto["fecha_correo"]),
                )
            )
        return salida

    def formulas_en_salida(self, fila: Fila) -> list[str]:
        """Encabezados de las celdas de salida de la fila que contienen fórmulas (no se pisan)."""
        return [
            COLUMNAS[clave][0]
            for clave in COLUMNAS_SALIDA
            if _es_formula(self._valor(fila.numero_fila, clave))
            or self.ws.cell(row=fila.numero_fila, column=self.columnas[clave]).data_type == "f"
        ]

    def _celdas_escribibles(self, fila: Fila, claves: Sequence[str]) -> list[Any]:
        celdas = [self.ws.cell(row=fila.numero_fila, column=self.columnas[c]) for c in claves]
        for clave, c in zip(claves, celdas):
            if _es_formula(c.value) or c.data_type == "f":
                raise ErrorTracker(
                    f"La celda {c.coordinate} («{COLUMNAS[clave][0]}», fila de {fila.empresa}) tiene una fórmula; "
                    "no la sobrescribo. Bórrala a mano si quieres que el script anote ahí."
                )
        return celdas

    @staticmethod
    def _escribir_fecha(celda: Any, fecha: date) -> None:
        formato = celda.number_format or "General"  # leerlo ANTES: openpyxl lo cambia al asignar una fecha
        celda.value = fecha
        celda.number_format = formato if is_date_format(formato) else FORMATO_FECHA_EXCEL

    def marcar_sernac(self, fila: Fila, fecha: date, numero: str | None) -> None:
        """Escribe SOLO «No Molestar ingresado» (datetime.date; conserva el formato de fecha de la
        celda o usa FORMATO_FECHA_EXCEL) y, si hay número, «N° solicitud SERNAC» como texto. El
        portal no muestra N° (llega en el correo de respaldo de SERNAC): sin número, esa celda no se
        toca. Celda de salida con fórmula -> ErrorTracker sin escribir nada."""
        celda_fecha, celda_numero = self._celdas_escribibles(fila, ("fecha_sernac", "numero_sernac"))
        self._escribir_fecha(celda_fecha, fecha)
        texto = str(numero).strip() if numero else ""
        if texto:
            celda_numero.value = "'" + texto if texto[:1] in "=+-@" else texto

    def marcar_correo(self, fila: Fila, fecha: date) -> None:
        """Escribe SOLO «Enviado (fecha)» (datetime.date, mismo criterio de formato). Celda con
        fórmula -> ErrorTracker."""
        (celda,) = self._celdas_escribibles(fila, ("fecha_correo",))
        self._escribir_fecha(celda, fecha)

    def guardar(self, preguntar: Preguntar = input) -> Path:
        """wb.save(tmp) + os.replace(tmp, path). PermissionError -> pide "Cierra el archivo en
        Excel…" vía preguntar y reintenta; con EOF/abortar guarda en <stem>.pendiente.xlsx y avisa.
        Devuelve la ruta realmente escrita."""
        temporal = self.path.with_name(f"{self.path.stem}.guardando.tmp{self.path.suffix}")
        while True:
            self.wb.save(temporal)
            try:
                os.replace(temporal, self.path)
                return self.path
            except PermissionError:
                temporal.unlink(missing_ok=True)
            interrumpido = False
            try:
                respuesta = str(
                    preguntar(
                        f"No pude guardar {self.path.name}: cierra el archivo en Excel y presiona Enter "
                        "(o escribe «abortar» para guardar una copia pendiente): "
                    )
                )
            except EOFError:
                respuesta = "abortar"
            except KeyboardInterrupt:
                respuesta, interrumpido = "abortar", True
            if respuesta.strip().lower() == "abortar":
                pendiente = self.path.with_name(f"{self.path.stem}.pendiente.xlsx")
                self.wb.save(pendiente)
                log.warning(
                    "No se pudo guardar %s; los cambios quedaron en %s. Cópialos al tracker cuando lo cierres.",
                    self.path.name, pendiente.name,
                )
                if interrumpido:
                    raise AbortoUsuario("Interrumpido al guardar el tracker")
                return pendiente

    def respaldar(self, directorio: Path) -> Path:
        """Copia el archivo original a directorio/<stem>_AAAAMMDD-HHMMSS.xlsx."""
        directorio = Path(directorio)
        directorio.mkdir(parents=True, exist_ok=True)
        marca = datetime.now().strftime("%Y%m%d-%H%M%S")
        destino = directorio / f"{self.path.stem}_{marca}{self.path.suffix}"
        n = 1
        while destino.exists():
            n += 1
            destino = directorio / f"{self.path.stem}_{marca}-{n}{self.path.suffix}"
        shutil.copy2(self.path, destino)
        return destino


def filtrar_pendientes(
    filas: Sequence[Fila],
    fase: str,
    *,
    forzar: bool = False,
    empresas: Sequence[str] | None = None,
    limite: int | None = None,
) -> tuple[list[Fila], list[Resultado]]:
    """Fase "sernac": todas las filas; fase "correo": solo fila.es_correo. Filtra por `empresas`
    (empresa normalizada contiene alguno de los textos normalizados). Devuelve (pendientes,
    ya_hechas) con ya_hechas = Resultado(estado="ya_hecha") salvo con forzar. `limite` acota solo
    las pendientes."""
    candidatas = [f for f in filas if fase == "sernac" or f.es_correo]
    candidatas = _filtrar_por_empresa(candidatas, empresas)
    pendientes: list[Fila] = []
    hechas: list[Resultado] = []
    for fila in candidatas:
        hecho = fila.sernac_hecho if fase == "sernac" else fila.correo_hecho
        if hecho and not forzar:
            hechas.append(_resultado(fila, fase, "ya_hecha"))
        else:
            pendientes.append(fila)
    if limite is not None:
        pendientes = pendientes[: max(0, limite)]
    return pendientes, hechas


def _filtrar_por_empresa(filas: Sequence[Fila], empresas: Sequence[str] | None) -> list[Fila]:
    textos = [normalizar(e) for e in (empresas or []) if normalizar(e)]
    if not textos:
        return list(filas)
    return [f for f in filas if any(t in normalizar(f.empresa) for t in textos)]


# -------------------------------------------------------------------------- mapeo SERNAC
class Mapeo:
    """mapeo_sernac.csv: UTF-8 con BOM, delimitador ";" (al leer se detecta ";" o ","), encabezado
    ENCABEZADOS_MAPEO, una fila por entrada. Fila con ID y Entrada vacíos = «sin entrada en SERNAC».
    Clave: "#" (respaldo: Marca normalizada)."""

    def __init__(self, path: Path, registros: Sequence[Mapping[str, str]] = ()) -> None:
        """Mapeo en memoria asociado a `path` (donde registrar() guarda)."""
        self.path = Path(path)
        self._filas: list[dict[str, str]] = [
            {h: str(r.get(h) or "").strip() for h in ENCABEZADOS_MAPEO} for r in registros
        ]
        self._salto = "\r\n"

    @classmethod
    def cargar(cls, path: Path) -> "Mapeo":
        """Archivo inexistente -> mapeo vacío + advertencia. Mal formado -> ErrorMapeo."""
        path = Path(path)
        if not path.exists():
            log.warning(
                "No existe %s: todas las empresas quedan «sin mapeo» (se preguntará; usa --mapear).", path
            )
            return cls(path)
        try:
            texto = path.read_text(encoding="utf-8-sig")
        except (OSError, UnicodeDecodeError) as exc:
            raise ErrorMapeo(f"No pude leer {path.name}: {exc}") from None
        if not texto.strip():
            return cls(path)
        primera = texto.splitlines()[0]
        delimitador = ";" if primera.count(";") >= primera.count(",") and ";" in primera else ","
        lector = csv.reader(io.StringIO(texto, newline=""), delimiter=delimitador)
        encabezado = next(lector)
        canon = {normalizar_encabezado(h): h for h in ENCABEZADOS_MAPEO}
        indices: dict[str, int] = {}
        for i, h in enumerate(encabezado):
            nombre = canon.get(normalizar_encabezado(h))
            if nombre and nombre not in indices:
                indices[nombre] = i
        faltan = [h for h in ENCABEZADOS_MAPEO[:4] if h not in indices]
        if faltan:
            raise ErrorMapeo(
                f"{path.name}: faltan columnas {', '.join(faltan)} (encabezado esperado: {';'.join(ENCABEZADOS_MAPEO)})"
            )
        registros = []
        for numero_linea, campos in enumerate(lector, start=2):
            fila = {h: _sin_neutralizar(campos[i].strip()) if i < len(campos) else "" for h, i in indices.items()}
            fila.setdefault("Nota", "")
            if not any(fila.values()):
                continue
            if fila["#"] and _entero(fila["#"]) is None:
                raise ErrorMapeo(f"{path.name}, línea {numero_linea}: «#» no es un número ({fila['#']!r}).")
            if fila["ID SERNAC"] and _entero(fila["ID SERNAC"]) is None:
                raise ErrorMapeo(
                    f"{path.name}, línea {numero_linea}: «ID SERNAC» no es un número ({fila['ID SERNAC']!r})."
                )
            if fila["ID SERNAC"] and not fila["Entrada SERNAC"]:
                raise ErrorMapeo(f"{path.name}, línea {numero_linea}: hay «ID SERNAC» pero falta «Entrada SERNAC».")
            if not fila["#"] and not fila["Marca"]:
                raise ErrorMapeo(f"{path.name}, línea {numero_linea}: falta «#» y «Marca».")
            registros.append(fila)
        mapeo = cls(path, registros)
        mapeo._salto = "\r\n" if "\r\n" in texto else "\n"
        return mapeo

    def _filas_de(self, fila: Fila) -> list[int]:
        """Índices de las filas del CSV que corresponden a la empresa (por "#", luego por Marca)."""
        if fila.numero is not None:
            por_numero = [i for i, r in enumerate(self._filas) if _entero(r["#"]) == fila.numero]
            if por_numero:
                return por_numero
        marca = normalizar(fila.empresa)
        return [i for i, r in enumerate(self._filas) if marca and normalizar(r["Marca"]) == marca]

    def entradas(self, fila: Fila) -> list[EntradaSernac] | None:
        """Entradas de la empresa (por "#"; si no hay, por Marca normalizada). None = la empresa no
        está en el mapeo (hay que preguntar); [] = explícitamente sin entrada en SERNAC."""
        indices = self._filas_de(fila)
        if not indices:
            return None
        salida = []
        for i in indices:
            r = self._filas[i]
            if r["Entrada SERNAC"]:
                salida.append(EntradaSernac(_entero(r["ID SERNAC"]), r["Entrada SERNAC"]))
        return salida

    def registrar(self, fila: Fila, entradas: Sequence[EntradaSernac], nota: str = "") -> None:
        """Reemplaza las filas de esa empresa ([] -> una fila sin ID ni entrada) y guarda de forma
        atómica (tmp + os.replace), conservando las demás filas y sus notas."""
        numero = "" if fila.numero is None else str(fila.numero)
        marca = normalizar(fila.empresa)
        quitar = {
            i for i, r in enumerate(self._filas)
            if (numero and _entero(r["#"]) == fila.numero) or (not r["#"] and normalizar(r["Marca"]) == marca)
        }
        nuevas = [
            {"#": numero, "Marca": fila.empresa, "ID SERNAC": "" if e.id is None else str(e.id),
             "Entrada SERNAC": e.texto, "Nota": nota if k == 0 else ""}
            for k, e in enumerate(entradas)
        ] or [{"#": numero, "Marca": fila.empresa, "ID SERNAC": "", "Entrada SERNAC": "", "Nota": nota}]
        if quitar:
            posicion = min(quitar)
        elif fila.numero is not None:
            posicion = next(
                (i for i, r in enumerate(self._filas) if (_entero(r["#"]) or 0) > fila.numero), len(self._filas)
            )
        else:
            posicion = len(self._filas)
        restantes = [r for i, r in enumerate(self._filas) if i not in quitar]
        posicion -= sum(1 for i in quitar if i < posicion)
        self._filas = restantes[:posicion] + nuevas + restantes[posicion:]
        self._guardar()

    def _guardar(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporal = self.path.with_name(self.path.name + ".tmp")
        try:
            with open(temporal, "w", encoding="utf-8-sig", newline="") as fh:
                escritor = csv.writer(fh, delimiter=";", lineterminator=self._salto)
                escritor.writerow(ENCABEZADOS_MAPEO)
                for r in self._filas:
                    escritor.writerow([_celda_csv(r.get(h, "")) for h in ENCABEZADOS_MAPEO])
            os.replace(temporal, self.path)
        except PermissionError as exc:
            temporal.unlink(missing_ok=True)
            raise ErrorMapeo(f"No pude guardar {self.path.name} (¿está abierto en Excel?): {exc}") from None


def _sin_neutralizar(valor: str) -> str:
    """Inverso de _celda_csv al leer el mapeo: quita el «'» que se antepuso a un texto que empieza
    con = + - @ (para que Excel no lo trate como fórmula). Otros apóstrofos iniciales se conservan."""
    return valor[1:] if valor[:1] == "'" and valor[1:2] in ("=", "+", "-", "@", "\t", "\r") else valor


def resolver_entradas(
    fila: Fila, bot: "BotSernac", mapeo: Mapeo, preguntar: Preguntar
) -> list[EntradaSernac] | None:
    """Mapeo interactivo de una empresa usando solo bot.buscar_empresas. Términos =
    candidatos_busqueda(fila) + primera palabra (≥4 caracteres) de cada uno; junta las opciones (sin
    duplicados, en orden), las ordena con puntuar_opciones y muestra una lista numerada 1..N (máx.
    MAX_OPCIONES_MAPEO) con "*" en las sugerencias. Respuestas: Enter -> acepta las sugerencias;
    "1,3" -> esas; "b TEXTO" -> busca TEXTO y vuelve a mostrar; "n" -> sin entrada en SERNAC
    (registra []); "s" -> omitir por ahora (no registra, devuelve None). Registra con id=None vía
    mapeo.registrar y devuelve las entradas registradas."""
    print(f"\n=== Mapeo SERNAC de #{fila.etiqueta} «{fila.empresa}» ===")
    opciones: list[str] = []
    for termino in _terminos_mapeo(fila):
        opciones = _sin_duplicados(opciones + bot.buscar_empresas(termino))
    while True:
        mostradas = puntuar_opciones(fila, opciones)[:MAX_OPCIONES_MAPEO]
        if mostradas:
            for posicion, (indice, puntaje) in enumerate(mostradas, 1):
                marca = "*" if puntaje >= 3 else " "
                print(f" {marca} {posicion:2d}. {opciones[indice]}")
            if len(opciones) > len(mostradas):
                print(f"   (se muestran {len(mostradas)} de {len(opciones)}; usa «b TEXTO» para acotar)")
        else:
            print("   (la búsqueda no devolvió resultados)")
        sugeridas = [indice for indice, puntaje in mostradas if puntaje >= 3]
        respuesta = _pedir(
            preguntar,
            "Enter = aceptar las sugeridas (*) · números (ej. 1,3) · «b TEXTO» = buscar otro texto · "
            "«n» = no está en SERNAC · «s» = omitir por ahora: ",
        )
        minuscula = respuesta.lower()
        if minuscula == "s":
            print("   Omitida por ahora (no se guarda en el mapeo).")
            return None
        if minuscula == "n":
            return _registrar_mapeo(mapeo, fila, [], f"{DETALLE_SIN_ENTRADA} (confirmado {date.today():%d-%m-%Y})")
        if minuscula.startswith("b ") and respuesta[2:].strip():
            nuevas = bot.buscar_empresas(respuesta[2:].strip())
            opciones = _sin_duplicados(nuevas + opciones)
            continue
        if not respuesta:
            elegidas = sugeridas
            if not elegidas:
                print("   No hay sugerencias: escribe números, «b TEXTO», «n» o «s».")
                continue
        else:
            numeros = [_entero(x) for x in re.split(r"[\s,;]+", respuesta) if x]
            if not numeros or any(n is None or not 1 <= n <= len(mostradas) for n in numeros):
                print("   Respuesta no válida.")
                continue
            elegidas = list(dict.fromkeys(mostradas[n - 1][0] for n in numeros))
        entradas = [EntradaSernac(None, opciones[i]) for i in elegidas]
        return _registrar_mapeo(mapeo, fila, entradas, f"confirmado {date.today():%d-%m-%Y}")


def _registrar_mapeo(mapeo: Mapeo, fila: Fila, entradas: list[EntradaSernac], nota: str) -> list[EntradaSernac]:
    try:
        mapeo.registrar(fila, entradas, nota=nota)
    except ErrorMapeo as exc:
        log.warning("%s — el mapeo de «%s» queda solo en memoria por esta ejecución.", exc, fila.empresa)
    print(f"   Guardado en el mapeo: {len(entradas)} entrada(s).")
    return entradas


# -------------------------------------------------------------------------- correo
class EnviadorCorreo:
    """Envío por Gmail SMTP (SSL, puerto 465). smtp_factory/sleep None -> smtplib.SMTP_SSL /
    time.sleep resueltos al momento de usarse; tiempos None -> TIEMPOS."""

    def __init__(
        self,
        usuario: str,
        clave: Secreto,
        persona: Persona,
        smtp_factory: Callable[..., Any] | None = None,
        sleep: Dormir | None = None,
        tiempos: Mapping[str, Any] | None = None,
    ) -> None:
        self.usuario = usuario
        self._clave = clave
        self.persona = persona
        self._smtp_factory = smtp_factory
        self._dormir = sleep
        self.tiempos = dict(tiempos or TIEMPOS)
        self._conexion: Any = None

    def _sleep(self, segundos: float) -> None:
        (self._dormir or time.sleep)(segundos)

    def conectar(self) -> None:
        """smtp_factory(SMTP_HOST, SMTP_PUERTO, context=ssl.create_default_context(), timeout=30)
        y login(usuario, clave.revelar()). SMTPAuthenticationError se propaga; un usuario o clave con
        caracteres no ASCII (smtplib lanza UnicodeEncodeError) también se informa como
        SMTPAuthenticationError, sin el texto de la clave."""
        fabrica = self._smtp_factory or smtplib.SMTP_SSL
        conexion = fabrica(SMTP_HOST, SMTP_PUERTO, context=ssl.create_default_context(), timeout=30)
        try:
            conexion.login(self.usuario, self._clave.revelar())
        except UnicodeEncodeError:
            with contextlib.suppress(Exception):
                conexion.close()
            raise smtplib.SMTPAuthenticationError(
                535, "El usuario o la contraseña tienen caracteres no ASCII (ñ, tildes): no es una contraseña de aplicación."
            ) from None
        except BaseException:
            with contextlib.suppress(Exception):
                conexion.close()
            raise
        self._conexion = conexion

    def construir_mensaje(self, destinatario: str, asunto: str, cuerpo: str) -> EmailMessage:
        """From formataddr((nombre, usuario)), To, Reply-To (persona.correo si difiere del
        usuario), Subject, Date, Message-ID, text/plain UTF-8. CR/LF en cabeceras -> ValueError."""
        for nombre, valor in (("destinatario", destinatario), ("asunto", asunto), ("nombre", self.persona.nombre)):
            if "\r" in valor or "\n" in valor:
                raise ValueError(f"Salto de línea no permitido en la cabecera ({nombre})")
        msg = EmailMessage()
        msg["From"] = formataddr((self.persona.nombre, self.usuario))
        msg["To"] = destinatario
        if self.persona.correo.strip().lower() != self.usuario.strip().lower():
            msg["Reply-To"] = self.persona.correo
        msg["Subject"] = asunto
        msg["Date"] = formatdate(localtime=True)
        msg["Message-ID"] = make_msgid(domain=self.usuario.rsplit("@", 1)[-1])
        msg.set_content(cuerpo, charset="utf-8")
        return msg

    def _cerrar_conexion(self) -> None:
        conexion, self._conexion = self._conexion, None
        if conexion is not None:
            with contextlib.suppress(Exception):
                conexion.close()

    def enviar(self, destinatario: str, asunto: str, cuerpo: str) -> None:
        """Envía un correo reutilizando la conexión; reconecta si se cortó. Reintenta solo errores
        transitorios (es_transitorio_smtp); los permanentes se propagan sin reintentar."""
        mensaje = self.construir_mensaje(destinatario, asunto, cuerpo)

        def intento() -> None:
            if self._conexion is None:
                self.conectar()
            try:
                self._conexion.send_message(mensaje)
            except (smtplib.SMTPResponseException, smtplib.SMTPRecipientsRefused):
                raise  # el servidor respondió: la conexión sigue sirviendo
            except OSError:  # desconexión, timeout o red (SMTPServerDisconnected incluido)
                self._cerrar_conexion()
                raise

        con_reintentos(
            intento,
            intentos=self.tiempos["reintentos"],
            base=self.tiempos["backoff_base_s"],
            es_transitorio=es_transitorio_smtp,
            sleep=self._sleep,
            al_reintentar=lambda n, exc: log.warning(
                "Correo a %s: error transitorio (%s); reintento %d.", destinatario, type(exc).__name__, n + 1
            ),
        )

    def procesar(
        self,
        filas: Sequence[Fila],
        plantilla: Plantilla,
        tracker: Tracker,
        hoy: date | None = None,
        preguntar: Preguntar = input,
    ) -> list[Resultado]:
        """Un correo por fila. Correo de empresa ausente/inválido -> fallida "Falta correo válido de
        la empresa" (tracker intacto). Éxito -> marcar_correo + guardar; fallo -> tracker intacto.
        pausa_entre_correos solo entre envíos. SMTPAuthenticationError se propaga (aborta la fase).
        Cualquier excepción que corte la fase (también Ctrl+C o AbortoUsuario al guardar) sale con
        exc.resultados_parciales = los resultados ya obtenidos (incluido el correo recién enviado)."""
        hoy = hoy or date.today()
        resultados: list[Resultado] = []
        try:
            self._procesar(filas, plantilla, tracker, hoy, preguntar, resultados)
        except BaseException as exc:
            with contextlib.suppress(AttributeError, TypeError):
                exc.resultados_parciales = resultados  # type: ignore[attr-defined]
            raise
        return resultados

    def _procesar(
        self,
        filas: Sequence[Fila],
        plantilla: Plantilla,
        tracker: Tracker,
        hoy: date,
        preguntar: Preguntar,
        resultados: list[Resultado],
    ) -> None:
        intentados = 0
        for fila in filas:
            if not validar_correo(fila.correo_empresa):
                resultados.append(_resultado(fila, "correo", "fallida", detalle=DETALLE_SIN_CORREO))
                continue
            try:
                contexto = contexto_para(fila, self.persona, hoy)
                asunto = renderizar(plantilla.asunto, contexto)
                cuerpo = renderizar(plantilla.cuerpo, contexto)
            except ErrorPlantilla as exc:
                resultados.append(_resultado(fila, "correo", "fallida", detalle=str(exc)))
                continue
            if intentados:
                pausa_cortes(self.tiempos["pausa_entre_correos"], self._sleep)
            intentados += 1
            log.info("Correo a %s (%s)…", fila.empresa, fila.correo_empresa)
            try:
                self.enviar(fila.correo_empresa, asunto, cuerpo)
            except smtplib.SMTPAuthenticationError:
                raise
            except (smtplib.SMTPException, OSError, ValueError) as exc:
                log.warning("No se pudo enviar el correo a %s: %s", fila.empresa, exc)
                resultados.append(_resultado(fila, "correo", "fallida", detalle=f"{type(exc).__name__}: {exc}"))
                continue
            except KeyboardInterrupt:
                resultados.append(_resultado(
                    fila, "correo", "incierta",
                    detalle="Ctrl+C durante el envío: revisa «Enviados» en Gmail antes de reintentar",
                ))
                raise
            detalle = ""
            try:
                tracker.marcar_correo(fila, hoy)
                tracker.guardar(preguntar)
            except ErrorTracker as exc:
                detalle = f"Enviado, pero no se anotó en el tracker: {exc}"
                log.warning(detalle)
            except AbortoUsuario:
                resultados.append(_resultado(
                    fila, "correo", "ok", detalle="Enviado; la fecha quedó en la copia .pendiente.xlsx del tracker"
                ))
                raise
            resultados.append(_resultado(fila, "correo", "ok", detalle=detalle))

    def cerrar(self) -> None:
        """Cierra la conexión SMTP si existe (ignora errores)."""
        conexion, self._conexion = self._conexion, None
        if conexion is not None:
            try:
                conexion.quit()
            except Exception:
                with contextlib.suppress(Exception):
                    conexion.close()


# ------------------------------------------------------------------------- SERNAC
@contextlib.contextmanager
def lanzar_navegador(headless: bool = False) -> Iterator[Pagina]:  # pragma: no cover
    """Abre Chrome (channel="chrome", con respaldo a Chromium incluido), contexto es-CL 1280x900,
    sin tracing/video/HAR/perfil persistente/storage_state, y entrega una página. Importa
    playwright aquí y cierra todo en finally."""
    from playwright.sync_api import Error as ErrorPlaywright
    from playwright.sync_api import sync_playwright

    with sync_playwright() as pw:
        try:
            navegador = pw.chromium.launch(headless=headless, channel="chrome")
        except ErrorPlaywright:
            log.info("Google Chrome no está disponible; uso el Chromium incluido en Playwright.")
            navegador = pw.chromium.launch(headless=headless)
        contexto = None
        try:
            contexto = navegador.new_context(locale="es-CL", viewport={"width": 1280, "height": 900})
            pagina = contexto.new_page()
            pagina.set_default_timeout(TIEMPOS["timeout_elemento_ms"])
            yield pagina
        finally:
            if contexto is not None:
                with contextlib.suppress(Exception):
                    contexto.close()
            with contextlib.suppress(Exception):
                navegador.close()


def _colapsar(texto: str) -> str:
    return re.sub(r"\s+", " ", str(texto or "")).strip()


class BotSernac:
    """Automatiza «No Molestar» en el Portal del Consumidor sobre una página Playwright.
    Selectores P/CU vía page; selectores F vía page.frame_locator(selectores["marco_no_molestar"][0]).
    Usa PORTAL_URL / URL_NO_MOLESTAR / FORMATO_RUN del módulo (leídas al usarse)."""

    def __init__(
        self,
        page: Pagina,
        persona: Persona,
        *,
        selectores: Mapping[str, Sequence[str]],
        preguntar: Preguntar,
        sleep: Dormir,
        confirmar: bool,
        capturas_dir: Path,
        tiempos: Mapping[str, Any],
    ) -> None:
        self.page = page
        self.persona = persona
        self.sel = {k: list(v) for k, v in selectores.items()}
        self.preguntar = preguntar
        self.sleep = sleep
        self.confirmar = confirmar
        self.capturas_dir = Path(capturas_dir)
        self.tiempos = dict(tiempos)
        self._f = page.frame_locator(self.sel["marco_no_molestar"][0])
        self._clave: Secreto | None = None
        self._relogin_usado = False

    # ------------------------------------------------------------ utilidades internas
    @property
    def _t_elemento(self) -> float:
        return float(self.tiempos["timeout_elemento_ms"])

    @property
    def _retardo(self) -> float:
        return float(self.tiempos["retardo_tecleo_ms"])

    @staticmethod
    def _ciclos(timeout_ms: float, minimo: int = 1) -> int:
        return max(minimo, math.ceil(float(timeout_ms) / INTERVALO_SONDEO_MS))

    def _esperar_un_poco(self) -> None:
        self.page.wait_for_timeout(INTERVALO_SONDEO_MS)

    def _pausa(self) -> None:
        pausa_cortes(self.tiempos["pausa_entre_acciones"], self.sleep)

    def _raiz(self, marco: bool) -> Any:
        return self._f if marco else self.page

    def _visible(self, clave: str, marco: bool = False) -> Any:
        """Primer elemento VISIBLE AHORA entre las coincidencias de las alternativas de `clave`, o None.
        Revisa todas las coincidencias, no solo .first: el portal repite botones ocultos (p. ej. hay
        dos «logoutCU()» y el primero, del menú móvil, está oculto)."""
        for selector in self.sel[clave]:
            locator = self._raiz(marco).locator(selector)
            try:
                for i in range(min(locator.count(), 10)):
                    elemento = locator.nth(i)
                    if elemento.is_visible():
                        return elemento
            except _errores_playwright():
                continue
        return None

    def _esperar(self, clave: str, marco: bool = False, timeout_ms: float | None = None) -> Any:
        """Espera (sondeando) a que alguna alternativa de `clave` esté visible; None si no."""
        ciclos = self._ciclos(self._t_elemento if timeout_ms is None else timeout_ms)
        for ciclo in range(ciclos):
            locator = self._visible(clave, marco)
            if locator is not None:
                return locator
            if ciclo < ciclos - 1:
                self._esperar_un_poco()
        return None

    def _requerir(self, clave: str, marco: bool = False, timeout_ms: float | None = None) -> Any:
        """Como _esperar, pero lanza un TimeoutError de Playwright (transitorio) si no aparece."""
        locator = self._esperar(clave, marco, timeout_ms)
        if locator is None:
            raise _timeout_playwright(f"No apareció «{clave}» en el portal")
        return locator

    def _locator_con_elementos(self, clave: str, marco: bool = False) -> Any:
        """Locator (sin .first) de la primera alternativa que tenga elementos (o la primera)."""
        raiz = self._raiz(marco)
        for selector in self.sel[clave]:
            locator = raiz.locator(selector)
            if locator.count():
                return locator
        return raiz.locator(self.sel[clave][0])

    def _textos(self, clave: str, marco: bool = False) -> list[str]:
        return [_colapsar(t) for t in self._locator_con_elementos(clave, marco).all_inner_texts()]

    def _contar(self, clave: str, marco: bool = False) -> int:
        return max(self._raiz(marco).locator(s).count() for s in self.sel[clave])

    def _texto_de(self, raiz: Any) -> str:
        """Texto visible del documento (página o iframe); "" si está navegando."""
        try:
            return str(raiz.locator("body").inner_text(timeout=min(self._t_elemento, 3000.0)))
        except _errores_playwright():
            return ""

    def _esperar_existente(self, clave: str, marco: bool = False) -> Any:
        """Primer elemento de `clave` cuando exista (visible o no); TimeoutError si no aparece."""
        ciclos = self._ciclos(self._t_elemento)
        for ciclo in range(ciclos):
            for selector in self.sel[clave]:
                locator = self._raiz(marco).locator(selector)
                if locator.count():
                    return locator.first
            if ciclo < ciclos - 1:
                self._esperar_un_poco()
        raise _timeout_playwright(f"No apareció «{clave}» en el portal")

    def _ir(self, url: str) -> None:
        respuesta = self.page.goto(url, timeout=max(30000.0, 2 * self._t_elemento))
        estado = getattr(respuesta, "status", None)
        if estado == 403:
            raise ErrorPortal(
                f"El portal respondió HTTP 403 en {_sin_query(url)}: portal en mantención o bloqueo de red; "
                "prueba más tarde."
            )
        if estado is not None and estado >= 500:
            raise _timeout_playwright(f"El portal respondió HTTP {estado} en {_sin_query(url)}")
        self._pausa()

    def _escribir(self, locator: Any, texto: str, limpiar: bool = True) -> None:
        if limpiar:
            locator.fill("")
        locator.press_sequentially(texto, delay=self._retardo)

    def _marcar(self, locator: Any) -> None:
        """Marca un radio/casilla. Los «custom-control» de Bootstrap 4 dejan el <input> bajo su
        <label> y Playwright aborta el check() («label … intercepts pointer events»): en ese caso se
        repite con force=True (el clic cae en la etiqueta, que marca el input) y se verifica."""
        if locator.is_checked():
            return
        try:
            locator.check(timeout=min(self._t_elemento, 3000.0))
        except _errores_playwright() as exc:
            log.debug("check() normal falló (%s); repito con force=True.", _resumen_error(exc))
            locator.check(force=True, timeout=self._t_elemento)
        if not locator.is_checked():
            raise _timeout_playwright("No se pudo marcar una opción del formulario del portal")

    # ------------------------------------------------------------ sesión y bloqueos
    def abrir_portal(self) -> None:
        """goto(PORTAL_URL). Respuesta 403 o 5xx -> ErrorPortal ("portal en mantención o bloqueo de
        red; prueba más tarde"); no pregunta nada (la pausa reintentar/abortar la hace main)."""
        url = PORTAL_URL
        try:
            respuesta = self.page.goto(url, timeout=max(30000.0, 2 * self._t_elemento))
        except _errores_playwright() as exc:
            raise ErrorPortal(
                f"No pude abrir {_sin_query(url)} ({type(exc).__name__}): portal en mantención o bloqueo de red; "
                "prueba más tarde."
            ) from None
        estado = getattr(respuesta, "status", None)
        if estado is not None and estado >= 400:
            raise ErrorPortal(
                f"El portal respondió HTTP {estado}: portal en mantención o bloqueo de red; prueba más tarde."
            )
        log.info("Portal del Consumidor abierto.")

    def _atender_bloqueos(self, incluir_2fa: bool = True) -> bool:
        """Pausa (hasta 3 veces) mientras haya un CAPTCHA o 2FA visible. True si pausó."""
        pauso = False
        for _ in range(3):
            if self.hay_captcha():
                pausar(
                    "Hay un CAPTCHA en el navegador. Resuélvelo tú (el script nunca resuelve CAPTCHAs) "
                    "y luego vuelve aquí.",
                    self.preguntar,
                )
            elif incluir_2fa and self.hay_2fa():
                pausar("ClaveÚnica pide un código de verificación (2FA). Ingrésalo en el navegador.", self.preguntar)
            else:
                break
            pauso = True
        return pauso

    def iniciar_sesion(self, clave: Secreto) -> None:
        """Si la sesión ya está activa (_sesion_activa), vuelve. Si no: login_boton_claveunica y espera cu_run
        o sesion_iniciada (ClaveÚnica con sesión vigente devuelve al portal sin formulario -> listo).
        CAPTCHA/2FA -> pausar. Solo si la página es https://HOST_CLAVEUNICA: press_sequentially del
        RUN (FORMATO_RUN) y de la clave UNA vez (nunca fill); en otro sitio no escribe nada y pausa
        para que la persona ingrese a mano. Espera cu_ingresar habilitado y hace clic UNA vez; luego
        hasta timeout_login_s: cu_error -> ErrorLogin (sin reintentar); CAPTCHA/2FA -> pausar y
        sigue; sesion_iniciada -> listo; tiempo agotado -> ErrorLogin. Guarda la clave para un
        re-login."""
        self._clave = clave
        if self._sesion_activa():
            log.info("La sesión del portal ya estaba iniciada.")
            return
        if self._visible("cu_run") is None:
            boton = self._esperar("login_boton_claveunica")
            if boton is None:
                raise ErrorLogin("No encontré el botón «ClaveÚnica» en el portal (revisa SELECTORES).")
            log.info("Ingresando con ClaveÚnica…")
            boton.click(timeout=self._t_elemento)
        if self._esperar_formulario_o_sesion() is None:
            log.info("ClaveÚnica devolvió la sesión sin pedir la clave (sesión aún vigente).")
            return
        self._atender_bloqueos()
        campo_run = self._requerir("cu_run")
        if not self._en_claveunica():
            self._ingreso_manual()
            return
        campo_run.press_sequentially(formatear_rut(self.persona.rut, FORMATO_RUN), delay=self._retardo)
        self._pausa()
        if not self._en_claveunica():
            raise ErrorLogin("La página de ingreso cambió de sitio antes de escribir la clave; no la escribí.")
        self._requerir("cu_clave").press_sequentially(clave.revelar(), delay=self._retardo)
        self._pausa()
        boton = self._requerir("cu_ingresar")
        for _ in range(self._ciclos(self._t_elemento)):  # se habilita con el keyup de los dos campos
            if boton.is_enabled():
                break
            self._esperar_un_poco()
        else:
            raise ErrorLogin("El botón «INGRESA» de ClaveÚnica siguió deshabilitado.")
        boton.click(timeout=self._t_elemento)
        self._esperar_inicio_sesion()

    def _esperar_formulario_o_sesion(self) -> Any:
        """Tras el botón ClaveÚnica: locator de cu_run, o None si el portal volvió con la sesión ya
        iniciada (sesion_iniciada visible). Ninguno de los dos a tiempo -> ErrorLogin."""
        ciclos = self._ciclos(self._t_elemento)
        for ciclo in range(ciclos):
            if self._sesion_activa():
                return None
            campo_run = self._visible("cu_run")
            if campo_run is not None:
                return campo_run
            if ciclo < ciclos - 1:
                self._esperar_un_poco()
        raise ErrorLogin("No apareció el formulario de ClaveÚnica (revisa SELECTORES o el portal).")

    def _en_claveunica(self) -> bool:
        """True si la página actual es https://HOST_CLAVEUNICA (el único sitio donde se escribe la clave)."""
        partes = urlsplit(str(self.page.url or ""))
        return partes.scheme == "https" and (partes.hostname or "").lower() == HOST_CLAVEUNICA.lower()

    def _ingreso_manual(self) -> None:
        """El formulario de ingreso está en otro sitio: no se escribe nada; la persona decide."""
        host = urlsplit(str(self.page.url or "")).hostname or "?"
        log.warning("El formulario de ingreso está en «%s», no en %s: no escribo la ClaveÚnica.", host, HOST_CLAVEUNICA)
        pausar(
            f"El formulario de ingreso está en «{host}», no en https://{HOST_CLAVEUNICA}. Por seguridad NO escribo "
            "tu ClaveÚnica ahí. Si confías en ese sitio, ingresa tú en el navegador y presiona Enter; si no, "
            "escribe «abortar».",
            self.preguntar,
        )
        self._esperar_inicio_sesion()

    def _esperar_inicio_sesion(self) -> None:
        log.info("Esperando que el portal confirme el ingreso (hasta %s s)…", self.tiempos["timeout_login_s"])
        total = self._ciclos(float(self.tiempos["timeout_login_s"]) * 1000)
        restantes = total
        while restantes > 0:
            restantes -= 1
            if self._sesion_activa():
                log.info("Sesión iniciada con ClaveÚnica.")
                return
            error = self._visible("cu_error")
            if error is not None:
                texto = ""
                with contextlib.suppress(*_errores_playwright()):
                    texto = _colapsar(error.inner_text(timeout=self._t_elemento))[:200]
                self._capturar_error_login()
                raise ErrorLogin(f"ClaveÚnica rechazó el ingreso: {texto or 'error sin texto'}. No se reintenta.")
            if self._atender_bloqueos():
                restantes = max(restantes, total)
                continue
            if restantes:
                self._esperar_un_poco()
        raise ErrorLogin(
            f"No se confirmó el inicio de sesión en {self.tiempos['timeout_login_s']} s. No se reintenta."
        )

    def _capturar_error_login(self) -> None:
        with contextlib.suppress(Exception):
            self.capturas_dir.mkdir(parents=True, exist_ok=True)
            destino = self.capturas_dir / f"{datetime.now():%Y%m%d-%H%M%S}_login-error.png"
            self.page.screenshot(path=str(destino), full_page=True)

    def mantener_sesion(self) -> None:
        """Si boton_mantener_sesion está visible, le hace clic (si no, no hace nada)."""
        boton = self._visible("boton_mantener_sesion")
        if boton is not None:
            boton.click(timeout=self._t_elemento)
            log.info("Se pulsó «Permanecer conectado».")

    def nueva_empresa(self) -> None:
        """Empieza el trabajo de otra empresa (--mapear o fase SERNAC): vuelve a permitir UN
        re-ingreso automático si la sesión expira y pulsa «Permanecer conectado» si está a la vista."""
        self._relogin_usado = False
        self.mantener_sesion()

    def hay_captcha(self) -> bool:
        """True solo si hay un CAPTCHA VISIBLE (SELECTORES_CAPTCHA visibles con caja no nula, y
        URLs de page.frames: anchor v2 sin size=invisible, hcaptcha, challenges.cloudflare.com).
        Ignora el badge invisible de reCAPTCHA v3/Enterprise y bframes sin tamaño."""
        for selector in SELECTORES_CAPTCHA:
            locator = self.page.locator(selector)
            try:
                for i in range(min(locator.count(), 5)):
                    elemento = locator.nth(i)
                    if elemento.is_visible():
                        caja = elemento.bounding_box()
                        if caja and caja.get("width", 0) > 0 and caja.get("height", 0) > 0:
                            return True
            except _errores_playwright():
                continue
        try:
            urls = [str(getattr(marco, "url", "") or "") for marco in self.page.frames]
        except _errores_playwright():
            urls = []
        return any(_url_de_captcha(u) for u in urls)

    def hay_2fa(self) -> bool:
        """True si hay un campo de código visible (SELECTORES_2FA) o el texto de la página calza
        con REGEX_2FA."""
        for selector in SELECTORES_2FA:
            try:
                if self.page.locator(selector).first.is_visible():
                    return True
            except _errores_playwright():
                continue
        return bool(re.search(REGEX_2FA, self._texto_de(self.page), re.I))

    def _sesion_activa(self) -> bool:
        """True si algún elemento de `sesion_iniciada` está en la página, visible o no: el botón
        «logoutCU()» solo existe con la sesión iniciada y puede estar oculto en un menú."""
        for selector in self.sel["sesion_iniciada"]:
            try:
                if self.page.locator(selector).count() > 0:
                    return True
            except _errores_playwright():
                continue
        return False

    def _necesita_login(self) -> bool:
        if self._sesion_activa():
            return False
        return self._visible("login_boton_claveunica") is not None or self._visible("cu_run") is not None

    def _reingresar_si_hace_falta(self) -> None:
        if not self._necesita_login():
            return
        if self._clave is None or self._relogin_usado:
            raise ErrorLogin("La sesión del portal expiró y no puedo volver a ingresar automáticamente.")
        self._relogin_usado = True
        log.warning("La sesión del portal expiró; vuelvo a ingresar con ClaveÚnica.")
        self.iniciar_sesion(self._clave)
        self._ir(URL_NO_MOLESTAR)
        self.mantener_sesion()

    # ------------------------------------------------------------ buscador de empresas
    def buscar_empresas(self, termino: str) -> list[str]:
        """Pulsa «Permanecer conectado» si aparece. Si el buscador no está a la vista, abre
        URL_NO_MOLESTAR y «Ingresar nuevo». Clic en abrir_buscador_empresa, limpia y
        press_sequentially(termino) en campo_buscar_empresa, espera opciones_empresa (sin resultados /
        tiempo agotado -> []) y devuelve sus textos con espacios colapsados, en el orden del buscador.
        Luego pausa_entre_busquedas. Si el buscador informa un error de carga, reintenta la búsqueda
        (con_reintentos) y, agotada, lanza un TimeoutError de Playwright."""
        self.mantener_sesion()
        if self._visible("abrir_buscador_empresa", True) is None and self._visible("campo_buscar_empresa", True) is None:
            self._ir(URL_NO_MOLESTAR)
            self.mantener_sesion()
            self._reingresar_si_hace_falta()
            self._requerir("boton_nueva_solicitud", True).click(timeout=self._t_elemento)
            self._requerir("abrir_buscador_empresa", True)
        return con_reintentos(
            lambda: self._buscar(termino),
            intentos=self.tiempos["reintentos"],
            base=self.tiempos["backoff_base_s"],
            es_transitorio=es_transitorio_playwright,
            sleep=self.sleep,
            al_reintentar=lambda n, exc: log.warning(
                "Búsqueda «%s» en SERNAC falló (%s); reintento %d.", termino, _resumen_error(exc), n + 1
            ),
        )

    def _buscar(self, termino: str) -> list[str]:
        campo = self._visible("campo_buscar_empresa", True)
        if campo is None:
            self._requerir("abrir_buscador_empresa", True).click(timeout=self._t_elemento)
            campo = self._requerir("campo_buscar_empresa", True)
        self._escribir(campo, termino)
        textos = self._esperar_opciones()
        log.debug("Búsqueda «%s»: %d opciones.", termino, len(textos))
        pausa_cortes(self.tiempos["pausa_entre_busquedas"], self.sleep)
        return textos

    def _esperar_opciones(self) -> list[str]:
        """Opciones del buscador cuando dejan de cambiar (dos lecturas iguales sin «buscando…»).
        Mensaje «sin resultados» -> []; mensaje de error de carga de select2 («The results could not
        be loaded.») -> TimeoutError de Playwright (transitorio), para no confundirlo con «no está»."""
        anterior: list[str] | None = None
        ciclos = self._ciclos(self._t_elemento, minimo=3)
        for ciclo in range(ciclos):
            textos = [t for t in self._textos("opciones_empresa", True) if t]
            mensaje = None if textos else self._mensaje_buscador()
            cargando = (
                self._f.locator(SELECTOR_CARGANDO_EMPRESAS).first.is_visible()
                or any(REGEX_TEXTO_CARGANDO.search(t) for t in textos)
                or bool(mensaje and REGEX_TEXTO_CARGANDO.search(mensaje))
            )
            if not cargando:
                if any(REGEX_TEXTO_SIN_RESULTADOS.search(t) for t in textos):
                    return []
                if mensaje and REGEX_TEXTO_ERROR_BUSCADOR.search(mensaje):
                    raise _timeout_playwright(f"El buscador de empresas de SERNAC no pudo cargar resultados: {mensaje}")
                if mensaje is not None:  # «sin resultados» (u otro aviso que no es error)
                    return []
                if textos and textos == anterior:
                    return textos
                anterior = textos
            if ciclo < ciclos - 1:
                self._esperar_un_poco()
        return []

    def _mensaje_buscador(self) -> str | None:
        """Texto del mensaje de select2 visible (sin resultados, error, buscando…); None si no hay."""
        mensaje = self._f.locator(SELECTOR_SIN_RESULTADOS).first
        if not mensaje.is_visible():
            return None
        with contextlib.suppress(*_errores_playwright()):
            return _colapsar(mensaje.inner_text(timeout=min(self._t_elemento, 3000.0)))
        return ""

    def _clic_opcion(self, texto: str) -> None:
        locator = self._locator_con_elementos("opciones_empresa", True)
        for i, actual in enumerate(locator.all_inner_texts()):
            if _colapsar(actual) == texto:
                locator.nth(i).click(timeout=self._t_elemento)
                return
        raise _timeout_playwright(f"La opción «{texto}» ya no está en el buscador")

    def _id_elegido(self) -> str:
        locator = self._locator_con_elementos("select_empresa", True).first
        valor = ""
        for ciclo in range(self._ciclos(self._t_elemento)):
            valor = str(locator.input_value(timeout=self._t_elemento) or "")
            if valor:
                break
            self._esperar_un_poco()
        return valor

    # ------------------------------------------------------------ preparación (reintentable)
    def _preparar(self, fila: Fila, entradas: Sequence[EntradaSernac]) -> list[EntradaSernac]:
        """Parte REINTENTABLE: URL_NO_MOLESTAR, mantener_sesion, re-login (una vez por empresa) si
        aparece el login, acepto_datos marcado, nueva solicitud, teléfono y correo de la persona (sin
        duplicar los precargados ni quitar ninguno), cada entrada por texto exacto (+ id si se
        conoce; si no calza: lista numerada 1..N en el orden del buscador; número -> esa opción;
        m -> la persona la eligió en el navegador y el bot sigue con boton_agregar_empresa; s -> se
        omite esa entrada), y paso «Revisar solicitud» con todos los canales marcados hasta dejar
        boton_enviar visible.
        Devuelve las entradas agregadas; todas omitidas -> EmpresaOmitida."""
        self._ir(URL_NO_MOLESTAR)
        self.mantener_sesion()
        self._reingresar_si_hace_falta()
        self._marcar(self._esperar_existente("acepto_datos"))
        self._requerir("boton_nueva_solicitud", True).click(timeout=self._t_elemento)
        self._requerir("campo_telefono", True)
        self._pausa()
        self._asegurar_telefono()
        self._asegurar_correo()
        agregadas: list[EntradaSernac] = []
        textos_agregados: set[str] = set()
        for entrada in entradas:
            if normalizar(entrada.texto) in textos_agregados:
                continue
            if self._agregar_entrada(fila, entrada):
                agregadas.append(entrada)
                textos_agregados.add(normalizar(entrada.texto))
        if not agregadas:
            raise EmpresaOmitida(f"Se omitieron todas las entradas SERNAC de «{fila.empresa}»")
        self.mantener_sesion()
        self._requerir("boton_continuar", True).click(timeout=self._t_elemento)
        confirmar_telefonos = self._requerir("boton_confirmar_telefonos", True)
        self._marcar_canales()
        confirmar_telefonos.click(timeout=self._t_elemento)
        self._requerir("boton_enviar", True)
        self._marcar_canales()
        self.mantener_sesion()
        self._atender_bloqueos(incluir_2fa=False)
        return agregadas

    def _esperar_contacto(self, clave_lista: str, presente: Callable[[list[str]], bool]) -> None:
        ciclos = self._ciclos(self._t_elemento)
        for ciclo in range(ciclos):
            if presente(self._textos(clave_lista, True)):
                return
            error = self._visible("error_contacto", True)
            if error is not None:
                raise ErrorFormulario(f"El portal rechazó el contacto: {_colapsar(error.inner_text())}")
            if ciclo < ciclos - 1:
                self._esperar_un_poco()
        raise _timeout_playwright("El contacto no apareció en la lista del formulario")

    def _asegurar_telefono(self) -> None:
        telefono = self.persona.telefono

        def presente(textos: list[str]) -> bool:  # igualdad del número, no subcadena
            return any(telefono in _telefonos_en(t) for t in textos)

        if presente(self._textos("lista_telefonos", True)):
            log.debug("El teléfono ya estaba en la solicitud (precargado).")
            return
        self._marcar(self._requerir("radio_telefono", True))
        tipo = "Móvil" if self.persona.telefono.startswith("9") else "Fijo"
        self._requerir("tipo_telefono", True).select_option(label=tipo, timeout=self._t_elemento)
        self._escribir(self._requerir("campo_telefono", True), self.persona.telefono)
        self._requerir("boton_agregar_contacto", True).click(timeout=self._t_elemento)
        self._esperar_contacto("lista_telefonos", presente)
        self._pausa()

    def _asegurar_correo(self) -> None:
        correo = self.persona.correo.lower()

        def presente(textos: list[str]) -> bool:  # igualdad de la dirección, no subcadena
            return any(correo in _correos_en(t) for t in textos)

        if presente(self._textos("lista_correos", True)):
            log.debug("El correo ya estaba en la solicitud (precargado).")
            return
        self._marcar(self._requerir("radio_correo", True))
        self._escribir(self._requerir("campo_correo_usuario", True), self.persona.correo_usuario)
        self._escribir(self._requerir("campo_correo_dominio", True), self.persona.correo_dominio)
        self._requerir("boton_agregar_contacto", True).click(timeout=self._t_elemento)
        self._esperar_contacto("lista_correos", presente)
        self._pausa()

    def _agregar_entrada(self, fila: Fila, entrada: EntradaSernac) -> bool:
        """Elige la entrada en el buscador y la agrega a la lista. False si la persona la omite."""
        objetivo = normalizar(entrada.texto)
        termino = termino_busqueda(entrada.texto)
        usado, opciones = termino, self._buscar(termino)
        elegida = next((o for o in opciones if normalizar(o) == objetivo), None)
        if elegida is None and normalizar(termino) != objetivo:
            usado, opciones = entrada.texto, self._buscar(entrada.texto)
            elegida = next((o for o in opciones if normalizar(o) == objetivo), None)
        aviso = ""
        if elegida is not None:
            self._clic_opcion(elegida)
            if entrada.id is not None:
                elegido = self._id_elegido()
                if elegido != str(entrada.id):
                    aviso = f"La opción calza por texto, pero su ID ({elegido or '?'}) no es el del mapeo ({entrada.id})."
                    log.warning("«%s»: %s", entrada.texto, aviso)
                    elegida = None
        if elegida is None:
            decision = self._preguntar_sin_coincidencia(fila, entrada, opciones, aviso)
            if decision == "s":
                log.info("Entrada «%s» omitida por la persona.", entrada.texto)
                return False
            if isinstance(decision, int):
                self._buscar(usado)
                self._clic_opcion(opciones[decision])
        antes = self._contar("lista_empresas", True)
        self._requerir("boton_agregar_empresa", True).click(timeout=self._t_elemento)
        ciclos = self._ciclos(self._t_elemento)
        for ciclo in range(ciclos):
            if self._contar("lista_empresas", True) > antes:
                self._pausa()
                return True
            error = self._visible("error_empresa", True)
            if error is not None:
                raise ErrorFormulario(f"El portal no aceptó la empresa «{entrada.texto}»: {_colapsar(error.inner_text())}")
            if ciclo < ciclos - 1:
                self._esperar_un_poco()
        raise _timeout_playwright(f"La empresa «{entrada.texto}» no apareció en la lista")

    def _preguntar_sin_coincidencia(
        self, fila: Fila, entrada: EntradaSernac, opciones: Sequence[str], aviso: str
    ) -> int | str:
        print(f"\a\nNo encontré automáticamente la entrada SERNAC «{entrada.texto}» de «{fila.empresa}».")
        if aviso:
            print(aviso)
        mostradas = list(opciones)[:MAX_OPCIONES_SIN_COINCIDENCIA]
        if mostradas:
            print("Opciones del buscador:")
            for i, opcion in enumerate(mostradas, 1):
                print(f"  {i:2d}. {opcion}")
        else:
            print("  (la búsqueda no devolvió resultados)")
        while True:
            respuesta = _pedir(
                self.preguntar,
                "Número de la opción correcta, «m» si ya la elegiste en el navegador o «s» para omitir esta entrada: ",
            ).lower()
            if respuesta == "abortar":
                raise AbortoUsuario("La persona escribió «abortar»")
            if respuesta in ("m", "s"):
                return respuesta
            numero = _entero(respuesta)
            if numero is not None and 1 <= numero <= len(mostradas):
                return numero - 1
            print("Respuesta no válida.")

    def _marcar_canales(self) -> None:
        """Marca cada casilla de canal VISIBLE que esté desmarcada (vienen marcadas por defecto)."""
        for selector in self.sel["canales"]:
            casillas = self._f.locator(selector)
            for i in range(casillas.count()):
                casilla = casillas.nth(i)
                if casilla.is_visible():
                    self._marcar(casilla)

    # ------------------------------------------------------------ envío (punto sin retorno)
    def enviar_solicitud(self, fila: Fila, entradas: Sequence[EntradaSernac]) -> Resultado:
        """UNA solicitud por empresa: _preparar con reintentos, confirmación opcional
        ("¿Enviar solicitud para «Marca» (N entradas)? Revisa el navegador. [s/N]"; no "s" ->
        omitida), clic en boton_enviar (PUNTO SIN RETORNO: nada posterior se reintenta solo).
        Confirmación con N° -> ok; sin N° -> ok "sin N° — revisar captura"; no determinable ->
        pregunta s (enviada; luego N°, Enter = sin N°) / n (no enviada: UN intento completo más) /
        o o EOF (incierta). Dos capturas en capturas_dir: AAAAMMDD-HHMMSS_<NN>_<slug>.png (página
        completa, va en Resultado.captura) y ..._marco.png (iframe). El re-ingreso automático (uno
        por empresa, también si hubo que mapearla) se rehabilita con nueva_empresa()."""
        for intento in (1, 2):
            sufijo = "" if intento == 1 else "-intento2"
            try:
                agregadas = con_reintentos(
                    lambda: self._preparar(fila, entradas),
                    intentos=self.tiempos["reintentos"],
                    base=self.tiempos["backoff_base_s"],
                    es_transitorio=es_transitorio_playwright,
                    sleep=self.sleep,
                    al_reintentar=lambda n, exc: log.warning(
                        "«%s»: el portal falló al preparar la solicitud (%s); reintento %d desde cero.",
                        fila.empresa, _resumen_error(exc), n + 1,
                    ),
                )
            except EmpresaOmitida:
                return _resultado(fila, "sernac", "sin_coincidencia", detalle="ninguna entrada SERNAC elegida")
            except ErrorFormulario as exc:
                return _resultado(fila, "sernac", "fallida", detalle=str(exc), captura=self._capturar(fila, "-fallida"))
            except Exception as exc:
                if not es_transitorio_playwright(exc):
                    raise
                return _resultado(
                    fila, "sernac", "fallida",
                    detalle=f"No se pudo preparar la solicitud: {_resumen_error(exc)}",
                    captura=self._capturar(fila, "-fallida"),
                )
            if self.confirmar:
                respuesta = _pedir(
                    self.preguntar,
                    f"¿Enviar solicitud para «{fila.empresa}» ({len(agregadas)} entradas)? Revisa el navegador. [s/N] ",
                )
                if respuesta.lower() != "s":
                    return _resultado(fila, "sernac", "omitida", detalle="no confirmada por la persona")
            resultado = self._enviar_y_confirmar(fila, agregadas, sufijo, intento)
            if resultado is not None:
                return resultado
            log.warning("«%s»: la persona indica que no se envió; un intento completo más.", fila.empresa)
        raise AssertionError("inalcanzable")  # pragma: no cover

    def _enviar_y_confirmar(
        self, fila: Fila, agregadas: Sequence[EntradaSernac], sufijo: str, intento: int
    ) -> Resultado | None:
        """Desde el clic en «Enviar» (PUNTO SIN RETORNO). Devuelve el Resultado, o None si la persona
        dijo «n» en el primer intento (se hace UN intento completo más). Todo lo que ocurra después
        del clic queda registrado: un error inesperado -> incierta; Ctrl+C -> AbortoUsuario con
        .resultado (incierta, u ok si la persona ya respondió «s»)."""
        incierta: Resultado | None = None
        try:
            exito, texto = self._finalizar(fila, agregadas)
            numero = extraer_numero_solicitud(texto) or (
                extraer_numero_solicitud(self._texto_de(self.page)) if exito else None
            )
            captura = self._capturar(fila, sufijo)
            self._diagnostico(fila)
            if exito:
                log.info("«%s»: solicitud ingresada%s.", fila.empresa, f" (N° {numero})" if numero else "")
                return _resultado(
                    fila, "sernac", "ok", numero=numero, detalle="" if numero else DETALLE_SIN_NUMERO, captura=captura
                )
            incierta = self._incierta(fila, captura)
            decision = self._preguntar_envio_incierto(fila, captura)
            if decision == "s":
                return self._confirmada_por_la_persona(fila, captura)
        except AbortoUsuario as exc:
            if exc.resultado is None:
                exc.resultado = incierta or self._incierta(fila, self._capturar_sin_cortes(fila, sufijo))
            raise
        except KeyboardInterrupt:
            aborto = AbortoUsuario("Interrumpido después de hacer clic en «Enviar»")
            aborto.resultado = incierta or self._incierta(fila, self._capturar_sin_cortes(fila, sufijo))
            raise aborto from None
        except Exception as exc:  # p. ej. el navegador se cerró: el envío pudo haber salido igual
            log.warning("«%s»: error después de hacer clic en «Enviar» (%s).", fila.empresa, _resumen_error(exc))
            if incierta is None:
                incierta = self._incierta(fila, self._capturar_sin_cortes(fila, sufijo))
            incierta.detalle = f"{incierta.detalle} ({_resumen_error(exc)})"
            return incierta
        if decision == "n" and intento == 1:
            return None
        if decision == "n":
            return _resultado(
                fila, "sernac", "fallida", detalle="la persona indicó que no se envió (2 intentos)", captura=incierta.captura
            )
        return incierta

    def _diagnostico(self, fila: Fila) -> None:
        """Mientras SELECTORES_VERIFICADOS sea False: instantánea saneada (JSON + HTML sin valores ni
        token p= + PNG, también del iframe) de lo que mostró el portal tras «Enviar», para grabar los
        selectores de `confirmacion`. Nunca interrumpe el envío."""
        if SELECTORES_VERIFICADOS:
            return
        try:
            rutas = capturar_pagina(self.page, self.capturas_dir / "diagnostico", fila.numero or fila.numero_fila)
            log.info("Diagnóstico tras «Enviar» guardado: %s", ", ".join(Path(r).name for r in rutas))
        except Exception as exc:
            log.warning("No pude guardar el diagnóstico tras «Enviar»: %s", _resumen_error(exc))

    @staticmethod
    def _incierta(fila: Fila, captura: str | None) -> Resultado:
        return _resultado(
            fila, "sernac", "incierta",
            detalle="no se pudo confirmar el envío; revisa la captura y el portal antes de reintentar",
            captura=captura,
        )

    def _capturar_sin_cortes(self, fila: Fila, sufijo: str) -> str | None:
        """_capturar que nunca lanza (ni con un segundo Ctrl+C): se usa al registrar un aborto."""
        try:
            return self._capturar(fila, sufijo)
        except BaseException:
            return None

    def _confirmada_por_la_persona(self, fila: Fila, captura: str | None) -> Resultado:
        """La persona dijo «s»: pide el N° (Enter/EOF = sin N°). Ctrl+C aquí -> AbortoUsuario con el ok."""
        try:
            numero_manual = str(self.preguntar("Si ves el N° de solicitud, escríbelo (Enter = sin N°): ")).strip()
        except EOFError:
            numero_manual = ""
        except KeyboardInterrupt:
            aborto = AbortoUsuario("Interrumpido al pedir el N° de solicitud")
            aborto.resultado = _resultado(
                fila, "sernac", "ok", detalle=f"confirmada por la persona; {DETALLE_SIN_NUMERO}", captura=captura
            )
            raise aborto from None
        detalle = "confirmada por la persona" + ("" if numero_manual else f"; {DETALLE_SIN_NUMERO}")
        return _resultado(fila, "sernac", "ok", numero=numero_manual or None, detalle=detalle, captura=captura)

    def _finalizar(self, fila: Fila, agregadas: Sequence[EntradaSernac]) -> tuple[bool, str]:
        """Clic ÚNICO en boton_enviar y espera de la confirmación. (éxito, texto del marco).
        Un error de Playwright mientras el iframe navega no corta la espera (se vuelve a leer); si el
        navegador se cerró, se deja de esperar y el envío queda sin confirmar."""
        log.info("«%s»: enviando la solicitud (%d entradas)…", fila.empresa, len(agregadas))
        try:
            self.mantener_sesion()  # la confirmación pudo tardar: que el modal no tape «Enviar»
            self._requerir("boton_enviar", True).click(timeout=self._t_elemento)
        except _errores_playwright() as exc:
            log.warning("«%s»: el clic en «Enviar» falló (%s); NO se reintenta.", fila.empresa, _resumen_error(exc))
        textos = [normalizar(e.texto) for e in agregadas]
        ciclos = self._ciclos(float(self.tiempos["timeout_confirmacion_ms"]))
        for ciclo in range(ciclos):
            try:
                estado, leido = self._leer_confirmacion(textos)
                if estado is not None:
                    return estado, leido
                if ciclo < ciclos - 1:
                    self._esperar_un_poco()
            except _errores_playwright() as exc:
                log.debug("Lectura de la confirmación falló (%s).", _resumen_error(exc))
                if _navegador_cerrado(exc):
                    break
        return False, self._texto_de(self._f)

    def _leer_confirmacion(self, textos_entradas: Sequence[str]) -> tuple[bool | None, str]:
        """Una lectura del resultado del envío: (True, texto) éxito; (False, texto) el portal mostró un
        aviso negativo o de error; (None, "") todavía no se sabe. Solo hay éxito si el formulario ya
        no está (boton_enviar ausente) y: aparece `confirmacion` sin negación/error, o vuelve el
        listado (boton_nueva_solicitud) con alguna de las entradas y sin aviso de error."""
        if self._contar("boton_enviar", True) > 0:
            return None, ""
        alerta = self._visible("confirmacion", True)
        if alerta is not None:
            aviso = _colapsar(alerta.inner_text(timeout=min(self._t_elemento, 3000.0)))
            if _confirmacion_negativa(aviso):
                log.warning("El portal respondió: «%s».", aviso[:200])
                return False, self._texto_de(self._f)
            return True, self._texto_de(self._f)
        if self._visible("boton_nueva_solicitud", True) is not None:
            if self._f.locator(SELECTOR_ALERTA_ERROR).first.is_visible():
                return False, self._texto_de(self._f)
            texto = self._texto_de(self._f)
            if any(t and t in normalizar(texto) for t in textos_entradas):
                return True, texto
        return None, ""

    def _preguntar_envio_incierto(self, fila: Fila, captura: str | None) -> str:
        print(
            f"\a\nNo pude confirmar si la solicitud de «{fila.empresa}» se envió."
            + (f" Captura: {captura}" if captura else "")
        )
        while True:
            try:
                respuesta = str(
                    self.preguntar(
                        "Revisa el navegador. ¿Se envió? «s» = sí · «n» = no (un intento más) · «o» = no sé: "
                    )
                ).strip().lower()
            except EOFError:
                return "o"
            except KeyboardInterrupt:
                raise AbortoUsuario("Interrumpido tras enviar una solicitud sin confirmar") from None
            if respuesta in ("s", "n", "o"):
                return respuesta
            print("Responde «s», «n» u «o».")

    def _capturar(self, fila: Fila, sufijo: str = "") -> str | None:
        """Captura de página completa + del iframe; devuelve la ruta de la de página (o None)."""
        if "claveunica" in str(self.page.url).lower():
            return None
        try:
            self.capturas_dir.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            log.warning("No pude crear %s: %s", self.capturas_dir, exc)
            return None
        base = f"{datetime.now():%Y%m%d-%H%M%S}_{fila.etiqueta}_{slug(fila.empresa + sufijo)}"
        pagina: Path | None = self.capturas_dir / f"{base}.png"
        try:
            self.page.screenshot(path=str(pagina), full_page=True)
        except (*_errores_playwright(), OSError) as exc:
            log.warning("No pude guardar la captura de la página: %s", _resumen_error(exc))
            pagina = None
        try:
            self.page.locator(self.sel["marco_no_molestar"][0]).screenshot(
                path=str(self.capturas_dir / f"{base}_marco.png"), timeout=self._t_elemento
            )
        except (*_errores_playwright(), OSError) as exc:
            log.warning("No pude guardar la captura del formulario: %s", _resumen_error(exc))
        return str(pagina) if pagina else None


def _telefonos_en(texto: str) -> set[str]:
    """Teléfonos chilenos (9 dígitos nacionales) escritos en el texto: "+56 9 1234 5678" -> {"912345678"}.
    Un número más largo no contiene al de 9 dígitos (se compara igualdad, no subcadena)."""
    limpio = re.sub(r"[\s().-]", "", str(texto or ""))
    return set(re.findall(r"(?<!\d)(?:\+?56)?([2-9]\d{8})(?!\d)", limpio))


def _correos_en(texto: str) -> set[str]:
    """Direcciones de correo escritas en el texto, en minúsculas (para comparar por igualdad)."""
    return {
        c.lower().rstrip(".")
        for c in re.findall(r"[^\s@<>()\[\],;:\"'«»]+@[^\s@<>()\[\],;:\"'«»]+", str(texto or ""))
    }


def _confirmacion_negativa(texto: str) -> bool:
    """True si el aviso del portal niega o informa un error («no pudo», «no ingresada», «error»…).
    «No Molestar» (el nombre del servicio) y «No.»/«N°» (número) no cuentan como negación."""
    return bool(REGEX_CONFIRMACION_NEGATIVA.search(str(texto or "")))


def _url_de_captcha(url: str) -> bool:
    """URL de un iframe de CAPTCHA que requiere a la persona (nunca el reCAPTCHA invisible)."""
    u = url.lower()
    if "invisible" in u:  # reCAPTCHA v3/Enterprise (size=invisible) o hCaptcha invisible
        return False
    if "recaptcha" in u and "/anchor" in u:
        return True
    return "hcaptcha" in u or "challenges.cloudflare.com" in u


def _resumen_error(exc: BaseException) -> str:
    """Primera línea del error, sin URLs con query (para logs y el resumen)."""
    primera = (str(exc).strip().splitlines() or [""])[0][:200]
    primera = re.sub(r"(https?://[^\s\"'<>?#]+)\?\S*", r"\1?<omitido>", primera)
    return f"{type(exc).__name__}: {primera}" if primera else type(exc).__name__


# ------------------------------------------------------------------- exploración
_JS_DOCUMENTOS = r"""() => {
  const sinQuery = (u) => { if (!u) return u; const i = u.indexOf('?'); return i < 0 ? u : u.slice(0, i) + '?<omitido>'; };
  const docs = [];
  const visitar = (doc, ruta) => {
    docs.push([doc, ruta]);
    const marcos = doc.querySelectorAll('iframe, frame');
    for (let i = 0; i < marcos.length; i++) {
      let interno = null;
      try { interno = marcos[i].contentDocument; } catch (e) { interno = null; }
      if (interno && interno.documentElement) { visitar(interno, ruta.concat([i])); }
    }
  };
  visitar(document, []);
  const CAMPOS = 'input, select, textarea, button, a, iframe, frame, form, label, [role=button], [onclick]';
  const SIN_TEXTO = ['input', 'select', 'textarea', 'iframe', 'frame', 'form'];
  return docs.map(([doc, ruta]) => {
    const inventario = Array.from(doc.querySelectorAll(CAMPOS)).slice(0, 3000).map((el) => {
      const tag = el.tagName.toLowerCase();
      const caja = el.getBoundingClientRect();
      return {
        tag, id: el.id || null, name: el.getAttribute('name'), type: el.getAttribute('type'),
        clase: el.getAttribute('class'), aria_label: el.getAttribute('aria-label'),
        placeholder: el.getAttribute('placeholder'), role: el.getAttribute('role'),
        href: sinQuery(el.getAttribute('href')), src: sinQuery(el.getAttribute('src')),
        action: sinQuery(el.getAttribute('action')), onclick: el.getAttribute('onclick'),
        opciones: tag === 'select' ? Array.from(el.options).map((o) => o.label) : null,
        texto: SIN_TEXTO.includes(tag) ? null : (el.innerText || el.textContent || '').trim().replace(/\s+/g, ' ').slice(0, 160),
        visible: caja.width > 0 && caja.height > 0, deshabilitado: el.disabled === true,
      };
    });
    const clon = doc.documentElement.cloneNode(true);
    clon.querySelectorAll('[value]').forEach((e) => e.removeAttribute('value'));
    clon.querySelectorAll('textarea').forEach((e) => { e.textContent = ''; });
    clon.querySelectorAll('script').forEach((e) => { e.textContent = ''; });
    // Todos los atributos: URLs (href, src, action, formaction, data-*) sin query y el token p= censurado
    // en cualquier valor (onclick, data-url, data-remote…), antes de que outerHTML escriba & como &amp;.
    const ES_URL = /^(href|src|action|formaction|poster|data-.*)$/i;
    const TOKEN = /([?&]p=)[^&"'\s]+/g;
    clon.querySelectorAll('*').forEach((e) => {
      Array.from(e.attributes).forEach((a) => {
        let v = a.value;
        if (ES_URL.test(a.name) && v.indexOf('?') >= 0 && !/\s/.test(v.trim())) { v = sinQuery(v); }
        v = v.replace(TOKEN, '$1<omitido>');
        if (v !== a.value) { e.setAttribute(a.name, v); }
      });
    });
    return {ruta, url: sinQuery(doc.location ? doc.location.href : ''), titulo: doc.title || '', inventario,
            html: '<!doctype html>\n' + clon.outerHTML};
  });
}"""


# Token p= en cualquier forma: ?p= &p= y sus versiones HTML (&amp; &#38; &#x26;) o URL (%26p%3D)
_REGEX_TOKEN_P = re.compile(r"((?:[?&]|&amp;|&#0*38;|&#x0*26;|%26|%3F)p(?:=|%3D))(?!<omitido>)[^&\"'\s\\<>]+", re.I)


def _censurar_urls(texto: str) -> str:
    """Quita query strings de URLs absolutas y censura tokens p= (también en URLs relativas y en
    sus formas codificadas &amp;p=, &#38;p=, %26p%3D…)."""
    texto = re.sub(r"(https?://[^\s\"'<>?#\\]+)\?(?!<omitido>)[^\s\"'<>#\\]*", r"\1?<omitido>", texto)
    return _REGEX_TOKEN_P.sub(r"\1<omitido>", texto)


def _sanear_html(html: str) -> str:
    """Respaldo en Python: sin atributos value, textareas vacíos, <script> sin contenido y URLs
    (relativas o absolutas) de href/src/action/formaction sin query string."""
    html = re.sub(r"(<script\b[^>]*>).*?(</script>)", r"\1\2", html, flags=re.I | re.S)
    html = re.sub(r"(<textarea\b[^>]*>).*?(</textarea>)", r"\1\2", html, flags=re.I | re.S)
    html = re.sub(
        r"(\s(?:href|src|action|formaction)\s*=\s*([\"']))([^\"'?]*)\?(?!<omitido>)[^\"']*(\2)",
        r"\1\3?<omitido>\4", html, flags=re.I,
    )
    return re.sub(r"\svalue\s*=\s*(\"[^\"]*\"|'[^']*'|[^\s>]+)", "", html, flags=re.I)


def _documentos(page: Pagina, titulo: str) -> list[dict[str, Any]]:
    """Inventario + HTML saneado de la página y sus iframes del mismo origen."""
    try:
        datos = page.evaluate(_JS_DOCUMENTOS)
    except _errores_playwright() as exc:
        log.debug("No pude inventariar los documentos: %s", _resumen_error(exc))
        datos = None
    if isinstance(datos, list) and datos and all(isinstance(d, dict) for d in datos):
        return [dict(d, html=_sanear_html(str(d.get("html", "")))) for d in datos]
    try:
        html = page.content()
    except _errores_playwright():
        html = ""
    return [{"ruta": [], "url": _sin_query(page.url), "titulo": titulo, "inventario": [], "html": _sanear_html(html)}]


def capturar_pagina(page: Pagina, salida_dir: Path, n: int) -> list[Path]:
    """Escribe NN_<slug>.json (url, title, timestamp, frames, inventario de elementos SIN valores,
    también de los iframes), .png y .html (página e iframes: sin atributos value, textareas vacíos,
    <script> sin contenido). URLs con query -> "?<omitido>"; tokens p= -> "p=<omitido>"."""
    salida_dir = Path(salida_dir)
    salida_dir.mkdir(parents=True, exist_ok=True)
    try:
        titulo = str(page.title() or "")
    except _errores_playwright():
        titulo = ""
    url = str(page.url)
    base = f"{n:02d}_{slug(titulo or url.split('?')[0].rstrip('/').rsplit('/', 1)[-1], 40)}"
    documentos = _documentos(page, titulo)
    datos = {
        "url": _sin_query(url),
        "title": titulo,
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "frames": [_sin_query(getattr(f, "url", "")) for f in page.frames],
        "documentos": [{k: v for k, v in d.items() if k != "html"} for d in documentos],
    }
    rutas: list[Path] = []
    ruta_json = salida_dir / f"{base}.json"
    ruta_json.write_text(_censurar_urls(json.dumps(datos, ensure_ascii=False, indent=2)), encoding="utf-8")
    rutas.append(ruta_json)
    for i, documento in enumerate(documentos):
        ruta_html = salida_dir / (f"{base}.html" if i == 0 else f"{base}_marco{i}.html")
        ruta_html.write_text(_censurar_urls(str(documento.get("html", ""))), encoding="utf-8")
        rutas.append(ruta_html)
    ruta_png = salida_dir / f"{base}.png"
    try:
        page.screenshot(path=str(ruta_png), full_page=True, mask=_mascaras(page), style=_ESTILO_SIN_VALORES)
        rutas.append(ruta_png)
    except _errores_playwright() as exc:
        log.warning("No pude guardar la captura %s: %s", ruta_png.name, _resumen_error(exc))
    return rutas


_CAMPOS_CON_VALOR = "input:not([type=hidden]), textarea"
# El style de page.screenshot se aplica también dentro de los iframes (y de los anidados)
_ESTILO_SIN_VALORES = f"{_CAMPOS_CON_VALOR} {{ color: transparent !important; text-shadow: none !important; }}"


def _mascaras(page: Pagina) -> list[Any]:
    """Locators a enmascarar en el PNG: campos de la página y de cada iframe de primer nivel (un
    page.locator no entra en los iframes)."""
    mascaras = [page.locator(_CAMPOS_CON_VALOR)]
    try:
        marcos = page.locator("iframe").count()
    except _errores_playwright():
        marcos = 0
    for i in range(min(marcos, 20)):
        mascaras.append(page.frame_locator(f"iframe >> nth={i}").locator(_CAMPOS_CON_VALOR))
    return mascaras


_JS_CONTAR_ELEMENTOS = "() => document.querySelectorAll('input, select, textarea, button, a, iframe').length"


def _firma_pagina(page: Pagina) -> tuple[Any, ...] | None:
    """Firma barata de la página para detectar cambios de URL/DOM (None si está navegando)."""
    try:
        return (
            _sin_query(page.url),
            page.title(),
            tuple(_sin_query(getattr(f, "url", "")) for f in page.frames),
            page.evaluate(_JS_CONTAR_ELEMENTOS),
        )
    except _errores_playwright():
        return None


def explorar(page: Pagina, salida_dir: Path, preguntar: Preguntar, sleep: Dormir) -> int:
    """Recorrido para grabar selectores: la persona navega e inicia sesión a mano; se toma una
    instantánea al cambiar URL/DOM o con la tecla c (q = salir). Devuelve cuántas se tomaron."""
    salida_dir = Path(salida_dir)
    print(
        "\n=== Exploración del portal ===\n"
        "Navega e inicia sesión TÚ en el navegador. Tomo una instantánea (sin valores tipeados) cada vez que\n"
        f"cambia la página. Escribe «c» + Enter para forzar una, «q» + Enter para salir. Carpeta: {salida_dir}\n"
    )
    try:
        page.goto(PORTAL_URL)
    except _errores_playwright() as exc:
        log.warning("No pude abrir el portal: %s", _resumen_error(exc))
    ordenes: "queue.Queue[str]" = queue.Queue()

    def leer_ordenes() -> None:
        while True:
            try:
                orden = str(preguntar("") or "").strip().lower()
            except BaseException:  # EOF, Ctrl+C o el fin de las respuestas: salir
                ordenes.put("q")
                return
            ordenes.put(orden)
            if orden in ("q", "salir"):
                return

    threading.Thread(target=leer_ordenes, name="explorar-teclado", daemon=True).start()
    n = 0
    capturada: tuple[Any, ...] | None = None
    previa: tuple[Any, ...] | None = None

    def instantanea(firma: tuple[Any, ...] | None) -> None:
        nonlocal n, capturada
        n += 1
        try:
            rutas = capturar_pagina(page, salida_dir, n)
            print(f"  [{n:02d}] {_sin_query(page.url)} -> {len(rutas)} archivos")
        except (*_errores_playwright(), OSError) as exc:
            log.warning("No pude tomar la instantánea %d: %s", n, _resumen_error(exc))
        capturada = firma

    try:
        while n < MAX_INSTANTANEAS:
            firma = _firma_pagina(page)
            if firma is not None and firma == previa and firma != capturada:
                instantanea(firma)
            previa = firma
            try:
                orden = ordenes.get_nowait()
            except queue.Empty:
                orden = ""
            if orden in ("q", "salir"):
                break
            if orden == "c":
                instantanea(firma)
            sleep(1.0)
    except KeyboardInterrupt:
        pass
    print(f"Exploración terminada: {n} instantáneas en {salida_dir}")
    return n


# ------------------------------------------------------------------ logging/resumen
def configurar_logging(
    directorio: Path, secretos: Iterable[Secreto] = (), verbose: bool = False
) -> tuple[Path, FiltroSecretos]:
    """Logger NOMBRE_LOGGER con consola + archivo directorio/no_molestar_AAAAMMDD-HHMMSS.log,
    ambos con el MISMO FiltroSecretos (devuelto para agregar secretos después)."""
    directorio = Path(directorio)
    directorio.mkdir(parents=True, exist_ok=True)
    ruta = directorio / f"no_molestar_{datetime.now():%Y%m%d-%H%M%S}.log"
    filtro = FiltroSecretos(secretos)
    _cerrar_logging()
    consola = logging.StreamHandler(sys.stderr)
    consola.setLevel(logging.DEBUG if verbose else logging.INFO)
    consola.setFormatter(logging.Formatter("%(message)s"))
    archivo = logging.FileHandler(ruta, encoding="utf-8")
    archivo.setLevel(logging.DEBUG)
    archivo.setFormatter(logging.Formatter("%(asctime)s %(levelname)-7s %(message)s"))
    for handler in (consola, archivo):
        handler.addFilter(filtro)
        handler._no_molestar = True  # type: ignore[attr-defined]
        log.addHandler(handler)
    log.setLevel(logging.DEBUG)
    log.propagate = False
    return ruta, filtro


def _cerrar_logging() -> None:
    for handler in list(log.handlers):
        if getattr(handler, "_no_molestar", False):
            log.removeHandler(handler)
            with contextlib.suppress(Exception):
                handler.close()


def _celda_csv(valor: Any) -> str:
    texto = "" if valor is None else str(valor)
    return "'" + texto if texto[:1] in ("=", "+", "-", "@", "\t", "\r") else texto


def escribir_resumen_csv(resultados: Sequence[Resultado], ruta: Path) -> Path:
    """CSV utf-8-sig con delimitador ";" (fila, empresa, fase, estado, numero, detalle, captura).
    Celdas que empiezan con = + - @ se neutralizan anteponiendo "'". Devuelve la ruta."""
    ruta = Path(ruta)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    with open(ruta, "w", encoding="utf-8-sig", newline="") as fh:
        escritor = csv.writer(fh, delimiter=";")
        escritor.writerow(["fila", "empresa", "fase", "estado", "numero", "detalle", "captura"])
        for r in resultados:
            escritor.writerow(
                [_celda_csv(v) for v in (r.numero_fila, r.empresa, r.fase, r.estado, r.numero, r.detalle, r.captura)]
            )
    return ruta


ETIQUETAS_ESTADO = (
    ("ok", "ok"),
    ("fallida", "fallidas"),
    ("incierta", "inciertas"),
    ("sin_coincidencia", "sin coincidencia"),
    ("omitida", "omitidas"),
    ("ya_hecha", "ya hechas"),
)


def _accion_manual(filas: Sequence[Fila]) -> list[str]:
    """Líneas «Acción manual pendiente» (vía Formulario web o Canal alternativo)."""
    lineas = []
    for fila in filas:
        via = normalizar(fila.via)
        if via not in VIAS_MANUALES:
            continue
        if via == "formulario web":
            destino = fila.formulario_web or fila.canal_alternativo
        else:
            destino = fila.canal_alternativo or fila.formulario_web
        lineas.append(f"  - #{fila.etiqueta} {fila.empresa} ({fila.via}): {destino or 'sin enlace en el tracker'}")
    return lineas


def resumen_texto(resultados: Sequence[Resultado], filas: Sequence[Fila] = ()) -> str:
    """Tabla por fase (ok, fallidas, inciertas, sin coincidencia, omitidas, ya hechas), lista de
    fallas (empresa, fase, motivo), «Acción manual pendiente» (filas de `filas` con vía Formulario
    web o Canal alternativo, con su enlace o texto) y «Empresas sin entrada en SERNAC»."""
    lineas = ["", "=" * 70, "RESUMEN", "=" * 70]
    lineas.append("Fase    | " + " | ".join(etiqueta for _, etiqueta in ETIQUETAS_ESTADO))
    for fase in FASES:
        conteo = [sum(1 for r in resultados if r.fase == fase and r.estado == e) for e, _ in ETIQUETAS_ESTADO]
        lineas.append(f"{'SERNAC' if fase == 'sernac' else 'Correo':<7} | " + " | ".join(
            str(c).rjust(len(etiqueta)) for c, (_, etiqueta) in zip(conteo, ETIQUETAS_ESTADO)
        ))
    fallas = [r for r in resultados if r.estado in ("fallida", "incierta")]
    lineas += ["", f"Fallas e inciertas ({len(fallas)}):"]
    lineas += [
        f"  - {r.empresa} [{'SERNAC' if r.fase == 'sernac' else 'correo'}] {r.estado}: {r.detalle or 'sin detalle'}"
        for r in fallas
    ] or ["  (ninguna)"]
    manuales = _accion_manual(filas)
    lineas += ["", f"Acción manual pendiente ({len(manuales)}):"] + (manuales or ["  (ninguna)"])
    sin_entrada = [r for r in resultados if r.fase == "sernac" and r.estado == "sin_coincidencia"]
    lineas += ["", f"Empresas sin entrada en SERNAC ({len(sin_entrada)}):"]
    lineas += [f"  - {r.empresa}: {r.detalle or DETALLE_SIN_ENTRADA}" for r in sin_entrada] or ["  (ninguna)"]
    return "\n".join(lineas)


# ------------------------------------------------------------------------------ CLI
class _FormatoAyuda(argparse.HelpFormatter):
    """Ayuda de argparse con encabezados en español."""

    def add_usage(self, usage: Any, actions: Any, groups: Any, prefix: str | None = None) -> None:
        super().add_usage(usage, actions, groups, prefix="uso: " if prefix is None else prefix)


class _Parser(argparse.ArgumentParser):
    def error(self, message: str) -> Any:  # mismo código de salida 2 que argparse
        self.print_usage(sys.stderr)
        self.exit(SALIDA_VALIDACION, f"{self.prog}: error: {message}\nUsa --help para ver las opciones.\n")


def _entero_positivo(texto: str) -> int:
    valor = _entero(texto)
    if valor is None or valor < 1:
        raise argparse.ArgumentTypeError(f"se esperaba un entero ≥ 1, no {texto!r}")
    return valor


def parse_args(argv: Sequence[str] | None) -> argparse.Namespace:
    """Atributos: dry_run, mapear, solo, limite, empresa (lista), forzar, confirmar (por defecto
    not SELECTORES_VERIFICADOS; --confirmar/--sin-confirmar), explorar, tracker, config,
    plantilla, mapeo, salida (Path; por defecto relativos a DIR_SCRIPT), verbose.
    --mapear excluye --dry-run y --explorar; --explorar excluye --dry-run."""
    p = _Parser(
        prog="no_molestar.py",
        description="Solicitudes «No Molestar» en el Portal del Consumidor (SERNAC) y correos de oposición "
        "a partir del tracker Excel. Empieza siempre con --dry-run.",
        add_help=False,
        formatter_class=_FormatoAyuda,
    )
    p._optionals.title = "opciones"
    p.add_argument("-h", "--help", action="help", help="muestra esta ayuda y termina")
    modo = p.add_mutually_exclusive_group()
    modo.add_argument("--dry-run", action="store_true", help="valida todo y muestra el plan; no pide claves ni envía nada")
    modo.add_argument("--mapear", action="store_true", help="asocia cada empresa con sus entradas del directorio SERNAC")
    modo.add_argument("--explorar", action="store_true", help="recorre el portal contigo y guarda instantáneas para grabar selectores")
    p.add_argument("--solo", choices=FASES, help="ejecuta solo una fase: sernac o correo")
    p.add_argument("--limite", type=_entero_positivo, metavar="N", help="procesa como máximo N empresas pendientes por fase")
    p.add_argument("--empresa", action="extend", nargs="+", metavar="TEXTO", help="solo empresas cuya Marca contiene TEXTO (repetible)")
    p.add_argument("--forzar", action="store_true", help="incluye empresas ya hechas (y en --mapear revisa todas)")
    confirmar = p.add_mutually_exclusive_group()
    confirmar.add_argument("--confirmar", dest="confirmar", action="store_true", help="pide confirmación antes de cada envío a SERNAC (por defecto, mientras SELECTORES_VERIFICADOS = False)")
    confirmar.add_argument("--sin-confirmar", dest="confirmar", action="store_false", help="no pide confirmación antes de cada envío")
    p.set_defaults(confirmar=not SELECTORES_VERIFICADOS)
    p.add_argument("--tracker", type=Path, default=DIR_SCRIPT / ARCHIVO_TRACKER, metavar="P", help=f"tracker Excel (por defecto: {ARCHIVO_TRACKER} junto al script)")
    p.add_argument("--config", type=Path, default=DIR_SCRIPT / ARCHIVO_CONFIG, metavar="P", help=f"tus datos (por defecto: {ARCHIVO_CONFIG} junto al script)")
    p.add_argument("--plantilla", type=Path, default=DIR_SCRIPT / ARCHIVO_PLANTILLA, metavar="P", help=f"plantilla del correo (por defecto: {ARCHIVO_PLANTILLA} junto al script)")
    p.add_argument("--mapeo", type=Path, default=DIR_SCRIPT / ARCHIVO_MAPEO, metavar="P", help=f"mapeo empresa -> entradas SERNAC (por defecto: {ARCHIVO_MAPEO} junto al script)")
    p.add_argument("--salida", type=Path, default=DIR_SCRIPT / "salida", metavar="DIR", help="carpeta de logs, resúmenes y capturas (por defecto: salida junto al script)")
    p.add_argument("-v", "--verbose", action="store_true", help="muestra más detalle en la consola")
    return p.parse_args(list(argv) if argv is not None else None)


def _configurar_consola() -> None:
    for flujo in (sys.stdout, sys.stderr):
        with contextlib.suppress(Exception):
            flujo.reconfigure(encoding="utf-8", errors="replace")


@dataclass
class _Entradas:
    """Entradas validadas para un modo de ejecución."""

    config: Config
    tracker: Tracker
    filas: list[Fila]
    mapeo: Mapeo
    plantilla: Plantilla | None


def _validar(args: argparse.Namespace, con_plantilla: bool) -> _Entradas:
    """Carga y valida config, tracker, mapeo y (si corresponde) la plantilla. Errores -> excepción."""
    config = cargar_config(args.config)
    tracker = Tracker(args.tracker)
    filas = tracker.filas()
    mapeo = Mapeo.cargar(args.mapeo)
    plantilla = None
    if con_plantilla:
        plantilla = cargar_plantilla(args.plantilla)
        hoy = date.today()
        for fila in filas:
            if fila.es_correo:
                contexto = contexto_para(fila, config.persona, hoy)
                renderizar(plantilla.asunto, contexto)
                renderizar(plantilla.cuerpo, contexto)
    return _Entradas(config, tracker, filas, mapeo, plantilla)


def _pedir_clave(prompt: str) -> str:
    """getpass hasta 3 veces (vacía -> repetir); Ctrl+C/EOF o 3 vacías -> AbortoUsuario."""
    for _ in range(3):
        try:
            valor = getpass.getpass(prompt)
        except (KeyboardInterrupt, EOFError):
            raise AbortoUsuario("Se canceló el ingreso de la clave") from None
        if valor:
            return valor
        print("La clave no puede estar vacía.")
    raise AbortoUsuario("No se ingresó la clave (3 intentos)")


def _estado_sernac(fila: Fila, mapeo: Mapeo, pendientes: set[int], hechas: set[int], solo: str | None) -> str:
    if solo == "correo":
        return "no se procesa (--solo correo)"
    if fila.numero_fila in hechas:
        return "ya hecho"
    if fila.numero_fila not in pendientes:
        return "fuera de --limite"
    entradas = mapeo.entradas(fila)
    if entradas is None:
        return "falta mapeo (se preguntará)"
    if not entradas:
        return DETALLE_SIN_ENTRADA
    return f"pendiente con {len(entradas)} entrada{'s' if len(entradas) != 1 else ''}"


def _estado_correo(fila: Fila, pendientes: set[int], hechas: set[int], solo: str | None) -> str:
    if not fila.es_correo:
        return "no aplica"
    if solo == "sernac":
        return "no se procesa (--solo sernac)"
    if fila.numero_fila in hechas:
        return "ya hecho"
    return "pendiente" if fila.numero_fila in pendientes else "fuera de --limite"


def _escribir_vista_previa(datos: _Entradas, filas: Sequence[Fila], carpeta: Path) -> int:
    carpeta.mkdir(parents=True, exist_ok=True)
    for viejo in carpeta.glob("*.txt"):
        viejo.unlink()
    persona, hoy, n = datos.config.persona, date.today(), 0
    for fila in filas:
        if not validar_correo(fila.correo_empresa) or datos.plantilla is None:
            continue
        contexto = contexto_para(fila, persona, hoy)
        cabecera = [f"Para: {fila.correo_empresa}", f"De: {persona.nombre} <{datos.config.gmail_usuario}>"]
        if persona.correo.lower() != datos.config.gmail_usuario.lower():
            cabecera.append(f"Responder a: {persona.correo}")
        cabecera.append(f"Asunto: {renderizar(datos.plantilla.asunto, contexto)}")
        texto = "\n".join(cabecera) + "\n\n" + renderizar(datos.plantilla.cuerpo, contexto)
        (carpeta / f"{fila.etiqueta}_{slug(fila.empresa)}.txt").write_text(texto, encoding="utf-8")
        n += 1
    return n


def _dry_run(args: argparse.Namespace, datos: _Entradas) -> int:
    filas = _filtrar_por_empresa(datos.filas, args.empresa)
    s_pend, s_sin_entrada, s_hechas = _pendientes_sernac(
        filas, datos.mapeo, forzar=args.forzar, empresas=None, limite=args.limite
    )
    c_pend, c_hechas = filtrar_pendientes(filas, "correo", forzar=args.forzar, limite=args.limite)
    s_ids, s_hechas_ids = {f.numero_fila for f in s_pend + s_sin_entrada}, {r.numero_fila for r in s_hechas}
    c_ids, c_hechas_ids = {f.numero_fila for f in c_pend}, {r.numero_fila for r in c_hechas}
    print("\nPLAN (--dry-run: no se pide ninguna clave, no se abre el navegador y no se envía ni modifica nada)")
    print(f"{'Fila':>4}  {'#':>3}  {'Empresa':<36}  {'SERNAC':<30}  {'Correo':<12}  Problemas")
    conteo: dict[str, int] = {}
    for fila in filas:
        sernac = _estado_sernac(fila, datos.mapeo, s_ids, s_hechas_ids, args.solo)
        correo = _estado_correo(fila, c_ids, c_hechas_ids, args.solo)
        clave = "pendiente" if sernac.startswith("pendiente") else sernac
        conteo[clave] = conteo.get(clave, 0) + 1
        problemas = []
        if fila.es_correo and not validar_correo(fila.correo_empresa):
            problemas.append("falta correo válido de la empresa")
        if formulas := datos.tracker.formulas_en_salida(fila):
            problemas.append("fórmula en " + ", ".join(formulas))
        numero = "" if fila.numero is None else str(fila.numero)
        print(f"{fila.numero_fila:>4}  {numero:>3}  {fila.empresa:<36}  {sernac:<30}  {correo:<12}  {'; '.join(problemas)}")
    entradas_total = sum(len(datos.mapeo.entradas(f) or []) for f in s_pend)
    print(
        f"\nSERNAC: {conteo.get('pendiente', 0)} solicitudes pendientes ({entradas_total} entradas), "
        f"{conteo.get(DETALLE_SIN_ENTRADA, 0)} sin entrada en SERNAC, "
        f"{conteo.get('falta mapeo (se preguntará)', 0)} con falta de mapeo, {len(s_hechas)} ya hechas."
    )
    print(f"Correo: {len(c_pend) if args.solo != 'sernac' else 0} pendientes, {len(c_hechas)} ya hechos, "
          f"{sum(1 for f in filas if not f.es_correo)} no aplica.")
    manuales = _accion_manual(filas)
    print(f"\nAcción manual pendiente ({len(manuales)}):")
    print("\n".join(manuales) or "  (ninguna)")
    if args.solo != "sernac":
        carpeta = Path(args.salida) / "vista_previa"
        n = _escribir_vista_previa(datos, c_pend, carpeta)
        print(f"\nVista previa de {n} correos en {carpeta}")
    print("\nTodo validado. Para ejecutar de verdad, quita --dry-run (empieza con --limite 1).")
    return SALIDA_OK


def _abrir_y_entrar(bot: Any, clave: Secreto, preguntar: Preguntar) -> None:
    """abrir_portal con pausa reintentar/abortar ante ErrorPortal, luego iniciar_sesion."""
    while True:
        try:
            bot.abrir_portal()
            break
        except ErrorPortal as exc:
            print(f"\a{exc}")
            respuesta = _pedir(preguntar, "Enter = reintentar · «abortar» = saltar la fase SERNAC: ")
            if respuesta.lower() == "abortar":
                raise
    bot.iniciar_sesion(clave)


def _nuevo_bot(page: Pagina, persona: Persona, args: argparse.Namespace, preguntar: Preguntar, dormir: Dormir) -> Any:
    return BotSernac(
        page,
        persona,
        selectores=SELECTORES,
        preguntar=preguntar,
        sleep=dormir,
        confirmar=args.confirmar,
        capturas_dir=Path(args.salida) / "capturas",
        tiempos=TIEMPOS,
    )


def _fase_sernac(
    args: argparse.Namespace, datos: _Entradas, trabajo: list[Fila], clave: Secreto,
    preguntar: Preguntar, dormir: Dormir, resultados: list[Resultado],
) -> None:
    """Una solicitud por empresa. Errores de fase (login, portal, navegador) fallan las filas
    restantes; la fase de correo se ejecuta igual."""
    hechas: set[int] = set()
    try:
        with lanzar_navegador() as page:
            bot = _nuevo_bot(page, datos.config.persona, args, preguntar, dormir)
            _abrir_y_entrar(bot, clave, preguntar)
            for k, fila in enumerate(trabajo):
                if k:
                    pausa_cortes(TIEMPOS["pausa_entre_empresas"], dormir)
                log.info("[%d/%d] SERNAC: %s", k + 1, len(trabajo), fila.empresa)
                resultados.append(_procesar_empresa(bot, datos, fila, preguntar))
                hechas.add(fila.numero_fila)
    except AbortoUsuario as exc:
        if exc.resultado is not None:
            resultados.append(exc.resultado)
        raise
    except (ErrorLogin, ErrorPortal, *_errores_playwright()) as exc:
        if isinstance(exc, ErrorNoMolestar):
            motivo = str(exc)
        elif _navegador_cerrado(exc):
            motivo = f"el navegador se cerró o falló ({_resumen_error(exc)})"
        else:
            motivo = _resumen_error(exc)
        log.error("Fase SERNAC detenida: %s", motivo)
        resultados.extend(
            _resultado(f, "sernac", "fallida", detalle=f"fase SERNAC detenida: {motivo}")
            for f in trabajo if f.numero_fila not in hechas
        )


def _procesar_empresa(bot: Any, datos: _Entradas, fila: Fila, preguntar: Preguntar) -> Resultado:
    """Mapeo (interactivo si falta) + envío de UNA empresa; anota el tracker si quedó ok. Los errores
    de la empresa (también un error transitorio del portal al buscarla) solo fallan esa empresa;
    ErrorLogin, ErrorPortal y un navegador cerrado detienen la fase. Un AbortoUsuario que trae un
    resultado ok (la persona confirmó el envío) anota el tracker antes de propagarse."""
    try:
        bot.nueva_empresa()
        entradas = datos.mapeo.entradas(fila)
        if entradas is None:
            entradas = resolver_entradas(fila, bot, datos.mapeo, preguntar)
            if entradas is None:
                return _resultado(fila, "sernac", "omitida", detalle="sin mapeo SERNAC (omitida por ahora)")
            if not entradas:
                return _resultado(fila, "sernac", "sin_coincidencia", detalle=DETALLE_SIN_ENTRADA)
        resultado = bot.enviar_solicitud(fila, entradas)
    except AbortoUsuario as exc:
        if exc.resultado is not None and exc.resultado.estado == "ok":
            _anotar_sernac(datos, fila, exc.resultado, preguntar)
        raise
    except (ErrorLogin, ErrorPortal):
        raise
    except Exception as exc:
        if _navegador_cerrado(exc):
            raise
        if es_transitorio_playwright(exc):
            log.warning("%s: el portal falló (%s); sigo con la siguiente empresa.", fila.empresa, _resumen_error(exc))
            return _resultado(fila, "sernac", "fallida", detalle=f"el portal falló: {_resumen_error(exc)}")
        log.exception("Error inesperado con %s", fila.empresa)
        return _resultado(fila, "sernac", "fallida", detalle=f"error inesperado: {_resumen_error(exc)}")
    if resultado.estado == "ok":
        _anotar_sernac(datos, fila, resultado, preguntar)
    return resultado


def _anotar_sernac(datos: _Entradas, fila: Fila, resultado: Resultado, preguntar: Preguntar) -> None:
    """Fecha y N° en el tracker + guardar. Un aborto al guardar (Ctrl+C con el Excel abierto) se
    propaga con este resultado adjunto, para que el envío quede en el resumen."""
    try:
        datos.tracker.marcar_sernac(fila, date.today(), resultado.numero)
        datos.tracker.guardar(preguntar)
    except ErrorTracker as exc:
        resultado.detalle = f"{resultado.detalle}; no se anotó en el tracker: {exc}".lstrip("; ")
        log.warning("%s: %s", fila.empresa, exc)
    except AbortoUsuario as exc:
        resultado.detalle = f"{resultado.detalle}; la fecha quedó en la copia .pendiente.xlsx del tracker".lstrip("; ")
        exc.resultado = resultado
        raise


def _fase_correo(
    enviador: EnviadorCorreo, datos: _Entradas, pendientes: list[Fila], preguntar: Preguntar,
    resultados: list[Resultado],
) -> None:
    assert datos.plantilla is not None
    try:
        resultados.extend(enviador.procesar(pendientes, datos.plantilla, datos.tracker, preguntar=preguntar))
    except smtplib.SMTPAuthenticationError as exc:
        print(f"\a{MENSAJE_AUTH_GMAIL}")
        log.error(MENSAJE_AUTH_GMAIL)
        parciales = list(getattr(exc, "resultados_parciales", []))
        resultados.extend(parciales)
        listos = {r.numero_fila for r in parciales}
        resultados.extend(
            _resultado(f, "correo", "fallida", detalle=MENSAJE_AUTH_GMAIL) for f in pendientes if f.numero_fila not in listos
        )
    except BaseException as exc:  # Ctrl+C, aborto al guardar, error inesperado: lo ya hecho va al resumen
        resultados.extend(getattr(exc, "resultados_parciales", []))
        raise
    finally:
        enviador.cerrar()


def _probar_gmail(enviador: EnviadorCorreo, dormir: Dormir) -> str:
    """Login de prueba en Gmail. Devuelve "" si funciona o el motivo del fallo."""
    try:
        con_reintentos(
            enviador.conectar, intentos=TIEMPOS["reintentos"], base=TIEMPOS["backoff_base_s"],
            es_transitorio=es_transitorio_smtp, sleep=dormir,
        )
    except smtplib.SMTPAuthenticationError:
        print(f"\a{MENSAJE_AUTH_GMAIL}")
        log.error(MENSAJE_AUTH_GMAIL)
        return MENSAJE_AUTH_GMAIL
    except (smtplib.SMTPException, OSError) as exc:
        motivo = f"No se pudo conectar a Gmail: {_resumen_error(exc)}"
        log.error(motivo)
        return motivo
    log.info("Gmail: inicio de sesión correcto.")
    return ""


@dataclass
class _Plan:
    """Qué hay que hacer en una ejecución real, por fase."""

    sernac: list[Fila]  # empresas que van al portal (tienen entradas o falta mapeo)
    sin_entrada: list[Fila]  # mapeo explícito []: no se abre el portal por ellas
    correo: list[Fila]
    ya_hechas: list[Resultado]

    @property
    def con_correo_valido(self) -> list[Fila]:
        return [f for f in self.correo if validar_correo(f.correo_empresa)]


def _pendientes_sernac(
    filas: Sequence[Fila], mapeo: Mapeo, *, forzar: bool, empresas: Sequence[str] | None, limite: int | None
) -> tuple[list[Fila], list[Fila], list[Resultado]]:
    """(van al portal, sin entrada en SERNAC, ya hechas). Las empresas con mapeo explícito []
    nunca se marcan como hechas: se separan ANTES de aplicar `limite`, que acota solo las que van
    al portal (si no, «--limite 1» se quedaría para siempre en la primera de ellas)."""
    pendientes, hechas = filtrar_pendientes(filas, "sernac", forzar=forzar, empresas=empresas)
    sin_entrada = [f for f in pendientes if mapeo.entradas(f) == []]
    ids_sin_entrada = {f.numero_fila for f in sin_entrada}
    al_portal = [f for f in pendientes if f.numero_fila not in ids_sin_entrada]
    if limite is not None:
        al_portal = al_portal[: max(0, limite)]
    return al_portal, sin_entrada, hechas


def _planificar(args: argparse.Namespace, datos: _Entradas) -> _Plan:
    s_pend: list[Fila] = []
    sin_entrada: list[Fila] = []
    s_hechas: list[Resultado] = []
    if args.solo != "correo":
        s_pend, sin_entrada, s_hechas = _pendientes_sernac(
            datos.filas, datos.mapeo, forzar=args.forzar, empresas=args.empresa, limite=args.limite
        )
    c_pend, c_hechas = (
        filtrar_pendientes(datos.filas, "correo", forzar=args.forzar, empresas=args.empresa, limite=args.limite)
        if args.solo != "sernac" else ([], [])
    )
    return _Plan(sernac=s_pend, sin_entrada=sin_entrada, correo=c_pend, ya_hechas=s_hechas + c_hechas)


def _ejecutar(
    args: argparse.Namespace, datos: _Entradas, filtro: FiltroSecretos, preguntar: Preguntar, dormir: Dormir,
    resultados: list[Resultado],
) -> None:
    """Ejecución real: claves (solo las necesarias) -> respaldo -> prueba de Gmail -> SERNAC -> correo."""
    plan = _planificar(args, datos)
    resultados.extend(plan.ya_hechas)
    resultados.extend(_resultado(f, "sernac", "sin_coincidencia", detalle=DETALLE_SIN_ENTRADA) for f in plan.sin_entrada)

    clave_cu = clave_gmail = None
    if plan.sernac:
        clave_cu = Secreto(_pedir_clave("Clave ClaveÚnica (no se muestra ni se guarda): "))
        filtro.agregar(clave_cu)
    if plan.con_correo_valido:
        crudo = _pedir_clave(
            f"Contraseña de aplicación de Gmail para {datos.config.gmail_usuario} (no se muestra ni se guarda): "
        )
        filtro.agregar(Secreto(crudo))
        clave_gmail = Secreto(crudo.replace(" ", ""))  # Google la muestra en grupos de 4 letras
        filtro.agregar(clave_gmail)
    if not (plan.sernac or plan.correo):
        print("No hay nada pendiente con estos filtros.")
        return
    log.info("Respaldo del tracker: %s", datos.tracker.respaldar(Path(args.salida) / "respaldos"))

    enviador, motivo_gmail = None, ""
    if clave_gmail is not None:
        enviador = EnviadorCorreo(datos.config.gmail_usuario, clave_gmail, datos.config.persona, sleep=dormir)
        motivo_gmail = _probar_gmail(enviador, dormir)
        if not motivo_gmail and plan.sernac:
            enviador.cerrar()  # la fase SERNAC puede tardar: se reconecta al enviar
    if clave_cu is not None:
        _fase_sernac(args, datos, plan.sernac, clave_cu, preguntar, dormir, resultados)
    if enviador is not None and not motivo_gmail:
        _fase_correo(enviador, datos, plan.correo, preguntar, resultados)
        return
    for fila in plan.correo:  # sin conexión a Gmail (o sin ningún correo válido): nada se envía
        motivo = motivo_gmail if motivo_gmail and validar_correo(fila.correo_empresa) else DETALLE_SIN_CORREO
        resultados.append(_resultado(fila, "correo", "fallida", detalle=motivo))


def _modo_mapear(args: argparse.Namespace, datos: _Entradas, filtro: FiltroSecretos, preguntar: Preguntar, dormir: Dormir) -> int:
    filas = _filtrar_por_empresa(datos.filas, args.empresa)
    pendientes = [f for f in filas if args.forzar or datos.mapeo.entradas(f) is None]
    if args.limite is not None:
        pendientes = pendientes[: args.limite]
    if not pendientes:
        print("Todas las empresas ya están en el mapeo (usa --forzar para revisarlas igual).")
        return SALIDA_OK
    print(f"Empresas por mapear: {len(pendientes)}")
    clave = Secreto(_pedir_clave("Clave ClaveÚnica (no se muestra ni se guarda): "))
    filtro.agregar(clave)
    registradas = omitidas = 0
    codigo = SALIDA_OK
    try:
        with lanzar_navegador() as page:
            bot = _nuevo_bot(page, datos.config.persona, args, preguntar, dormir)
            _abrir_y_entrar(bot, clave, preguntar)
            for fila in pendientes:
                try:
                    bot.nueva_empresa()  # un re-ingreso automático por empresa; «Permanecer conectado»
                    entradas = resolver_entradas(fila, bot, datos.mapeo, preguntar)
                except Exception as exc:
                    if not es_transitorio_playwright(exc):
                        raise
                    log.warning("%s: el portal falló al buscar (%s); queda sin mapear.", fila.empresa, _resumen_error(exc))
                    entradas = None
                if entradas is None:
                    omitidas += 1
                else:
                    registradas += 1
    except (ErrorLogin, ErrorPortal, *_errores_playwright()) as exc:
        motivo = str(exc) if isinstance(exc, ErrorNoMolestar) else _resumen_error(exc)
        log.error("Mapeo detenido: %s (lo ya mapeado quedó guardado).", motivo)
        codigo = SALIDA_FALLAS
    print(f"\nMapeo: {registradas} empresas registradas, {omitidas} omitidas. Archivo: {datos.mapeo.path}")
    return codigo


def main(argv: Sequence[str] | None = None) -> int:
    """Punto de entrada. Códigos: 0 todo ok; 1 fallas/inciertas; 2 error de configuración o
    validación; 130 Ctrl+C o aborto."""
    _configurar_consola()
    args = parse_args(argv)
    salida = Path(args.salida)
    try:
        ruta_log, filtro = configurar_logging(salida, verbose=args.verbose)
    except OSError as exc:
        print(f"No pude crear la carpeta de salida {salida}: {exc}", file=sys.stderr)
        return SALIDA_VALIDACION

    def preguntar(prompt: str = "") -> str:
        return input(prompt)

    def dormir(segundos: float) -> None:
        time.sleep(segundos)

    resultados: list[Resultado] = []
    datos: _Entradas | None = None
    abortado = False
    try:
        if args.explorar:
            with lanzar_navegador() as page:
                explorar(page, salida / "exploracion", preguntar, dormir)
            return SALIDA_OK
        try:
            datos = _validar(args, con_plantilla=not (args.mapear or args.solo == "sernac"))
        except (ErrorConfig, ErrorPlantilla, ErrorTracker, ErrorMapeo) as exc:
            log.error("Error de validación: %s", exc)
            return SALIDA_VALIDACION
        if args.dry_run:
            return _dry_run(args, datos)
        if args.mapear:
            return _modo_mapear(args, datos, filtro, preguntar, dormir)
        _ejecutar(args, datos, filtro, preguntar, dormir, resultados)
    except (AbortoUsuario, KeyboardInterrupt) as exc:
        abortado = True
        log.warning("Ejecución interrumpida: %s", exc or "Ctrl+C")
    except ImportError as exc:
        log.error("Falta una dependencia (%s). Instala: pip install -r requirements.txt && playwright install chromium", exc)
        return SALIDA_VALIDACION
    except Exception:
        log.exception("Error inesperado")
        return SALIDA_FALLAS
    finally:
        if datos is not None and not args.dry_run and not args.mapear and (resultados or abortado):
            _cerrar_ejecucion(resultados, datos, args, filtro)
        log.info("Log: %s", ruta_log)
        _cerrar_logging()
    if abortado:
        return SALIDA_ABORTO
    return SALIDA_FALLAS if any(r.estado in ("fallida", "incierta") for r in resultados) else SALIDA_OK


def _cerrar_ejecucion(resultados: list[Resultado], datos: _Entradas, args: argparse.Namespace, filtro: FiltroSecretos) -> None:
    """Resumen en pantalla + CSV (con detalles censurados)."""
    for r in resultados:
        r.detalle = filtro.censurar(r.detalle)
    print(resumen_texto(resultados, _filtrar_por_empresa(datos.filas, args.empresa)))
    ruta = escribir_resumen_csv(resultados, Path(args.salida) / f"resumen_{datetime.now():%Y%m%d-%H%M%S}.csv")
    print(f"\nResumen CSV: {ruta}")


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
