# antispam-chile — «No Molestar» en el SERNAC y correos de oposición

Herramienta gratuita y de código abierto (licencia MIT) para dejar de recibir llamadas, SMS, WhatsApp y correos publicitarios de 50 empresas chilenas: automatiza las solicitudes «No Molestar» del SERNAC y los correos de oposición y eliminación de datos.

## Qué hace

- **SERNAC (las 50 empresas):** entra al Portal del Consumidor con tu ClaveÚnica (navegador visible) y crea **una solicitud «No Molestar» por empresa** con tu teléfono, tu correo y las entradas del directorio SERNAC de esa empresa (según `mapeo_sernac.csv`).
- **Correo (solo «Vía de envío» = `Correo`, 30 empresas):** envía `plantilla.txt` por Gmail, un correo por empresa.
- Anota fechas y N° de solicitud en el tracker, guarda capturas y un resumen de fallas.
- **Nunca** resuelve CAPTCHAs (te espera), **nunca** guarda contraseñas y **nunca** toca las columnas con fórmulas.

## Requisitos e instalación

Necesitas Python 3.12 (mínimo 3.11, por `tomllib`). En la carpeta del proyecto:

```powershell
py -3.12 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m playwright install chromium
```

Si tienes Google Chrome instalado, el script lo usa; si no, usa el Chromium de Playwright.

## Archivos de entrada

**`config.toml`** — cópialo desde el ejemplo y pon tus datos (`nombre`, `rut`, `telefono`, `correo`; `[gmail] usuario` solo si envías desde otra cuenta). Sin contraseñas.

```powershell
Copy-Item config.example.toml config.toml
```

**`tracker_no_molestar_50_empresas.xlsx`** (hoja «Empresas»): tu copia de trabajo de la plantilla limpia que viene en el repositorio. Ciérralo en Excel antes de ejecutar.

```powershell
Copy-Item tracker_plantilla_50_empresas.xlsx tracker_no_molestar_50_empresas.xlsx
```

| | Columnas |
|---|---|
| Se LEEN | `#`, `Marca`, `Razón social`, `RUT`, `Correo datos personales`, `Vía de envío`, `Formulario web`, `Canal alternativo` |
| Se ESCRIBEN | `Enviado (fecha)` (correo) · `No Molestar ingresado` (SERNAC) · `N° solicitud SERNAC` solo si el portal muestra uno: hoy no lo muestra (solo «Su solicitud ha sido ingresada…») y lo envía en un correo de respaldo; anótalo tú si lo quieres |
| NUNCA se tocan | Todo lo demás: fórmulas (`Plazo…`, `Estado`), `¿Respondió?`, `¿Siguen llamando?`, `Notas`, el bloque «Resumen» y las otras hojas |

El script no crea columnas y se niega a escribir sobre una celda con fórmula. Una empresa con fecha o N° ya puesto se salta en la siguiente ejecución (`--forzar` la repite).

**`mapeo_sernac.csv`** — empresa del tracker → entradas del directorio SERNAC. Existe porque el directorio tiene nombres largos y **varias entradas por marca** (Movistar tiene «MOVISTAR TELEFONIA FIJA…» y «MOVISTAR MOVIL…»; Entel tiene tres), y el buscador devuelve ruido («entel» → HUENTELAUQUEN). Ya viene armado con búsquedas reales (50 empresas, 67 entradas).

- Columnas: `#;Marca;ID SERNAC;Entrada SERNAC;Nota`. **Una fila por entrada** (Movistar = 2 filas con el mismo `#`).
- `Entrada SERNAC` debe ser el texto exacto que muestra el buscador; el `ID SERNAC` es opcional (si está, se verifica).
- Fila con `ID SERNAC` y `Entrada SERNAC` vacíos = «no está en SERNAC» (ej.: #18 Banco Security): se salta sin preguntar.
- Empresa ausente del archivo = el script te preguntará (o usa `--mapear`).
- Editarlo en Excel: ábrelo con doble clic (separador `;`) y guárdalo como **«CSV UTF-8»**. El script acepta `;` o `,`. Un texto que empieza con `=`, `+`, `-` o `@` se guarda con `'` delante (para que Excel no lo tome como fórmula) y el script lo quita al leer.

**`plantilla.txt`** — primera línea `Asunto: …`; el resto es el cuerpo. Marcadores:

| Marcador | Valor |
|---|---|
| `{nombre}`, `{correo}`, `{telefono}` | tus datos (`+56 9 1234 5678`) |
| `{rut}` | **tu** RUT, con puntos |
| `{fecha}` | fecha de envío (dd-mm-aaaa) |
| `{empresa}`, `{razon_social}` | `Marca` y `Razón social` (esta puede estar vacía) |
| `{rut_empresa}`, `{correo_empresa}` | `RUT` y `Correo datos personales` de la empresa |

Otras columnas también sirven, en minúsculas, sin tildes y con `_` (ej. `{sector}`). Un marcador desconocido detiene la validación. `{{` y `}}` = llaves literales.

## Contraseña de aplicación de Gmail

Gmail no acepta tu contraseña normal por SMTP. Activa la **Verificación en 2 pasos** de tu cuenta Google y crea una contraseña de aplicación en <https://myaccount.google.com/apppasswords>. El script te la pide al ejecutar; revócala cuando termines. (En cuentas institucionales de Google Workspace el administrador puede tenerlas deshabilitadas.)

## Uso

1. **Ensayo:** `python no_molestar.py --dry-run` — valida todo, muestra el plan y deja vistas previas de los correos en `salida\vista_previa\`. No pide contraseñas ni envía nada.
2. **Mapeo:** revisa `mapeo_sernac.csv`, o `python no_molestar.py --mapear` (inicia sesión y te pregunta por las empresas que faltan; con `--forzar` revisa todas). Nunca envía solicitudes.
3. **Una sola empresa (opcional):** `python no_molestar.py --solo sernac --limite 1`, y revisa la captura en `salida\capturas\`. El flujo ya se verificó con un envío real (23-10-2026), por eso `SELECTORES_VERIFICADOS = True`. Si el portal cambia, vuelve a ponerlo en `False`: el script preguntará antes de cada «Enviar» y guardará un diagnóstico en `salida\capturas\diagnostico\`.
4. **Todo:** `python no_molestar.py` — pide solo las contraseñas que necesita (ClaveÚnica y/o la de aplicación de Gmail) y primero prueba el login de Gmail.

| Opción | Efecto |
|---|---|
| `--solo sernac` / `--solo correo` | solo esa fase (la de correo pide solo la clave de Gmail) |
| `--empresa TEXTO` | solo empresas cuya Marca contiene TEXTO (repetible: `--empresa entel --empresa wom`) |
| `--limite N` | máximo N empresas pendientes por fase (en SERNAC no cuentan las «sin entrada en SERNAC») |
| `--forzar` | repite empresas ya hechas |
| `--confirmar` / `--sin-confirmar` | preguntar o no antes de cada «Enviar» (por defecto solo pregunta si `SELECTORES_VERIFICADOS = False`) |
| `--salida DIR` | carpeta de resultados (por defecto `salida\`) |
| `--tracker P`, `--config P`, `--plantilla P`, `--mapeo P` | usar otros archivos (por defecto, los de la carpeta del script) |
| `-v`, `--verbose` | más detalle en la consola (el archivo de log siempre es detallado) |

Entre empresas y entre correos hay pausas aleatorias corteses; los errores pasajeros se reintentan (nunca después de hacer clic en «Enviar»).

## Pausas: qué responder

| Situación | Qué haces |
|---|---|
| CAPTCHA visible | resuélvelo tú en el navegador y presiona Enter |
| Código de verificación (2FA) | ingrésalo en el navegador y presiona Enter |
| Entrada no encontrada en el buscador | número de la opción · `m` = la elegiste tú en el navegador · `s` = omitir esa entrada |
| Empresa sin fila en el mapeo | Enter = aceptar sugerencias `*` · `1,3` = esas · `b TEXTO` = buscar otra cosa · `n` = no está en SERNAC · `s` = omitir por ahora |
| Confirmación antes de «Enviar» | revisa el navegador; `s` envía, cualquier otra cosa la deja omitida |
| Envío incierto | `s` = sí se envió (luego el N°, o Enter) · `n` = no se envió (un intento más) · `o` = dejarla incierta |
| Tracker abierto en Excel / portal caído | ciérralo o espera, y presiona Enter |
| Ingreso fuera de `accounts.claveunica.gob.cl` | el script no escribe tu clave: ingresa tú en el navegador solo si confías en el sitio y presiona Enter, o escribe `abortar` |

En cualquier pausa, `abortar` o Ctrl+C detiene todo (igual se escribe el resumen; si ya se hizo clic en «Enviar», esa empresa queda en él como incierta, u ok si respondiste `s`).

El portal precarga el teléfono de tu perfil: se mantiene y, si es distinto, se agrega el de `config.toml`.

## Grabar selectores si el portal cambia

Todos los selectores están en **un solo bloque** al inicio de `no_molestar.py` (`SELECTORES`). Si el portal cambió:

- `python no_molestar.py --explorar` abre el portal, inicias sesión a mano y guarda instantáneas (`c` = capturar, `q` = salir) en `salida\exploracion\`: inventario de elementos sin valores, HTML sin scripts y PNG con los campos tapados (también dentro del iframe), con URLs y tokens censurados.
- `python -m playwright codegen https://www.sernac.cl/app/consumidor/` también sirve, pero **ATENCIÓN: escribe en el código generado todo lo que tipeas, incluida tu ClaveÚnica**. No la uses para iniciar sesión, o borra ese código.

## Resultados (en `salida\`)

- **Tracker:** se guarda tras cada éxito; antes se respalda en `respaldos\<tracker>_AAAAMMDD-HHMMSS.xlsx`.
- **`capturas\`:** `AAAAMMDD-HHMMSS_NN_empresa.png` (página) y `…_marco.png` (formulario), también en fallas e inciertas.
- **`resumen_*.csv`** (separador `;`, abre directo en Excel): estado por empresa y fase — `ok`, `fallida`, `incierta`, `sin_coincidencia`, `omitida`, `ya_hecha`. En pantalla: conteos, fallas, **«Acción manual pendiente»** (empresas con vía Formulario web o Canal alternativo, con su enlace o texto) y empresas sin entrada en SERNAC.
- **`no_molestar_AAAAMMDD-HHMMSS.log`** (sin contraseñas ni tokens).
- Una **incierta** no se anota en el tracker: revisa el listado del portal antes de reintentarla, para no duplicarla.
- Códigos de salida: 0 todo ok · 1 fallas o inciertas · 2 error de validación · 130 abortado.

## Seguridad y privacidad

- Las contraseñas se piden con `getpass` y solo viven en memoria: nunca van a disco, log ni `config.toml`. Un error de ClaveÚnica no se reintenta (para no bloquear tu cuenta).
- La ClaveÚnica solo se escribe en `https://accounts.claveunica.gob.cl` (`HOST_CLAVEUNICA` en el bloque de configuración); en cualquier otro sitio el script se detiene y te pregunta.
- El navegador abre con un perfil limpio: no se guardan cookies, sesiones, trazas ni video.
- Los enlaces del portal llevan un token `p=` que contiene tu nombre y correo: el script lo censura en el log y en `--explorar`.
- Capturas y tracker contienen tus datos: `config.toml`, `salida/` y tu `tracker_no_molestar_50_empresas.xlsx` están en `.gitignore`; no los compartas.

## Solución de problemas

- **Portal 403 / «portal en mantención»:** el SERNAC lo cierra por mantención (pasó el 22-10-2026). Prueba más tarde.
- **«Cierra el archivo en Excel…»:** cierra el tracker y presiona Enter. Si abortas, los cambios quedan en `<tracker>.pendiente.xlsx` junto al original.
- **`No module named tomllib`:** tu Python es anterior a 3.11; activa el entorno `.venv` creado con Python 3.12.
- **La clave no avanza o la ventana se cierra:** ejecútalo en una ventana normal de PowerShell o Windows Terminal (abierta desde el menú Inicio), no en terminales integradas ni con «Ejecutar». Ahí `getpass` recibe lo que escribes y la ventana queda abierta al terminar.
- **Falla el login de Gmail:** usa una contraseña de aplicación (no la normal) y revisa `[gmail] usuario`.
- **Entrada no encontrada siempre:** el SERNAC renombró la entrada; corrige `mapeo_sernac.csv` o `--mapear --forzar --empresa X`.

## Pruebas

```powershell
python -m pip install pytest pytest-cov
python -m pytest                 # todas (usan copias, nunca la red)
python -m pytest -m "not e2e"    # sin las pruebas con Chromium
```

## Licencia y créditos

Licencia [MIT](LICENSE): úsalo, modifícalo y compártelo libremente. Sin garantía: revisa lo que envías.

Desarrollado por [TurboDevs](https://turbodevs.web.app). ¿Necesitas ayuda o un desarrollo como este? Escríbenos desde el sitio.
