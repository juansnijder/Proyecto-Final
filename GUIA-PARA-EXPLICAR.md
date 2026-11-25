# Cómo entender y presentar FatToFit

Abrí esta guía en VS Code y presioná **Ctrl+Shift+V** para verla con formato. El código que se ejecuta está en los archivos `.py`, `.html` y `.css`; esta guía es material para estudiar.

## 1. La idea del proyecto

FatToFit permite planificar rutinas semanales y registrar entrenamientos, comidas y peso. Cada cuenta tiene sus propios registros. Flask ejecuta Python, los HTML muestran el contenido, CSS define el diseño y Excel conserva la información.

Una frase para presentar el trabajo:

> «Cuando envío un formulario, Flask recibe los datos, los valida y los guarda en Excel. La Raspberry también puede leer una tarjeta RC522 y relacionar su identificador con mi cuenta. Para entrar con la tarjeta también tengo que escribir mi PIN».

No hay JavaScript. La lógica se ejecuta en Python; las acciones del navegador usan formularios, enlaces y recargas HTML.

## 2. Cada carpeta y archivo

```text
FatToFit-Final/
├── app.py                    Servidor, rutas, Excel, cuentas y lector
├── fitness.xlsx              Base de datos REAL
├── requirements.txt          Bibliotecas de la web y Excel
├── requirements-rfid.txt     Bibliotecas del RC522 para Raspberry Pi 3
├── probar_lector.py          Prueba física del lector, sin Flask ni Excel
├── pruebas.py                Pruebas automáticas sin hardware
├── LEEME.txt                 Instalación y pruebas paso a paso
├── LEEME-MOVIL.txt           Notas para el celular
├── GUIA-PARA-EXPLICAR.md      Esta guía
├── COMPROBACIONES.txt        Resultado de la revisión de esta entrega
├── templates/               Plantillas HTML
│   ├── base.html            Estructura común y navegación
│   ├── componentes.html     Fragmentos reutilizables: íconos y campos
│   ├── index.html           Inicio después de ingresar
│   ├── ingresar.html        Ingreso y creación de cuentas
│   ├── ajustes.html         Perfil, PIN y asociación de tarjetas
│   ├── rutinas.html         Plan de cada día
│   ├── entrenar.html        Series y registro de entrenamiento
│   ├── progreso.html        Peso, gráfico e historial
│   ├── alimentacion.html    Registro de comidas
│   ├── asistente.html       Asistente guiado e IA opcional
│   ├── reloj.html           Documento del temporizador
│   └── estado_lector.html   Mensajes de lector ocupado/no disponible
└── static/                  Archivos enviados directamente al navegador
    ├── estilos.css          Diseño general
    ├── movil.css            Adaptación del diseño a celular
    ├── manifest.json        Nombre, colores e íconos para el navegador
    └── icons/               Imágenes de los íconos
```

No renombres `templates` a `template` ni `static` a `statics`: el programa usa esos nombres. Cada página tiene su HTML y aprovecha la estructura de `base.html`.

Carpetas que se crean al usarlo:

| Carpeta/archivo | Para qué sirve |
| --- | --- |
| `copias_excel` | Hasta cinco respaldos de las escrituras anteriores. |
| `fotos` | Fotos de perfil, junto al Excel en uso. |
| `__pycache__` | Caché automática de Python. No se edita ni se necesita entregar. |
| `venv` | Entorno virtual con bibliotecas. Se crea de nuevo en cada equipo. |
| Archivos `.lock` | Coordinan escrituras y uso del lector. No son la base de datos. |

## 3. Navegador, servidor, ruta y formulario

El **servidor** es el programa Python en ejecución. En el colegio estará en la Raspberry. El **navegador** pide páginas al servidor y las muestra; puede estar en otra PC de la misma red.

Una **ruta** es una dirección que atiende Flask, como `/ajustes`. Una URL completa combina equipo, puerto y ruta. Por ejemplo: `http://192.168.1.50:5001/ajustes`; esa IP es solo un ejemplo.

**GET** pide una pantalla. **POST** envía un formulario para hacer una acción. Mirá este fragmento real de `app.py`:

```python
@app.post('/ajustes/pin')
def cambiar_pin():
    profile, _, _ = datos()
    pin = validar_pin('pin')
    actualizar_perfil(profile['id'], pin_hash=generate_password_hash(pin))
    flash('Tu PIN personal fue actualizado y guardado de forma segura.', 'success')
    return redirect(url_for('ajustes') + '#seguridad')
```

Se lee así: «Cuando llega un POST a esa dirección, busco la cuenta actual, valido el PIN, genero su hash, actualizo el perfil, muestro un mensaje y vuelvo a Ajustes». Los guiones bajos reciben valores que esa función no necesita usar.

`Flask(__name__, template_folder='templates', static_folder='static')` configura la aplicación y sus carpetas. `render_template()` prepara un HTML, `redirect()` envía al navegador a otra dirección y `url_for()` genera la dirección de una función. `request.form` permite leer los campos enviados por el formulario.

## 4. Cómo se completa un HTML

Las plantillas usan **Jinja**, que se ejecuta en Flask antes de enviar el HTML. No es JavaScript.

| Expresión | Significado |
| --- | --- |
| `{{ perfil.nombre }}` | Insertar el nombre de la cuenta. |
| `{% if ... %}` | Mostrar una parte si se cumple una condición. |
| `{% for ... %}` | Repetir HTML para una lista, como los ejercicios. |
| `{% extends 'base.html' %}` | Reutilizar la estructura común. |
| `{% block contenido %}` | Completar la zona de contenido de esa página. |
| `{{ csrf() }}` | Agregar el campo oculto de protección del formulario. |

Las funciones de `componentes.html` son **macros**, piezas reutilizables de HTML. `base.html` evita repetir toda la navegación en cada página. Abrir los archivos con `file://` o Live Server no ejecuta Flask ni Jinja: para verlos hay que iniciar `app.py` y abrir la URL HTTP.

## 5. CSS y archivos estáticos

HTML define el contenido y CSS define cómo se ve. `class="card"` conecta un elemento HTML con la regla `.card` de CSS. Las reglas eligen colores, tamaño, espacios y posición. `estilos.css` tiene la base; `movil.css` se carga después y adapta algunas reglas al celular. `@media` aplica reglas según el tamaño de pantalla.

Los SVG son dibujos. `manifest.json` es una descripción de datos para el navegador: nombre, colores, íconos y dirección de inicio. Ninguno de ellos necesita JavaScript. Para practicar, buscá una clase HTML y después su regla en CSS; cambiá temporalmente un espacio o un color y observá el resultado.

## 6. Excel como base de datos

Cada hoja es una tabla. La primera fila tiene las columnas y las demás, los registros.

| Hoja | Qué guarda |
| --- | --- |
| `Perfiles` | Cuenta, objetivo, altura, meta, tema, foto, hashes de clave/PIN, UID asociado y duración de descanso. |
| `Dias` | Días activos y título de cada rutina. |
| `Ejercicios` | Ejercicios del perfil y día con series y repeticiones. La columna de descanso antigua se conserva por compatibilidad. |
| `Sesiones` | Entrenamientos realizados: nombre de rutina, fecha, duración y notas. |
| `Series` | Carga, repeticiones y técnica de las series realizadas. |
| `Peso` | Registros corporales por fecha. |
| `Comidas` | Comidas, cantidades, kcal y proteínas. |
| `Alimentos` | Valores de referencia por 100 gramos. |
| `Mensajes` | Consultas y respuestas del asistente. |

`id` distingue un registro. `perfil_id` lo relaciona con una cuenta; `sesion_id` relaciona una serie con su entrenamiento. Excel no obliga a respetar esas relaciones: las comprueba el programa.

`SCHEMA` describe los encabezados esperados. `openpyxl` abre y modifica el libro. El programa usa diccionarios como `{'dia': 0, 'nombre': 'Press banca'}` y los convierte en filas.

Funciones para reconocer:

- `filas_excel()`: convierte filas en diccionarios.
- `agregar_excel()`: agrega un registro en el orden de las columnas.
- `reemplazar_excel()`: actualiza una tabla reescribiendo sus registros.
- `validar_libro()`: comprueba hojas y columnas; admite agregar al final las columnas de perfil previstas en versiones anteriores.
- `leer_excel()`: lee con un bloqueo y cierra el libro al terminar.
- `editar_excel()`: bloquea, modifica, guarda un temporal, respalda el anterior y reemplaza el original cuando terminó de guardar.
- `texto_excel()`: evita que un texto del usuario se ejecute como fórmula.

Tu `fitness.xlsx` original se conservó. Las comprobaciones automáticas usan copias temporales. La versión final utiliza `fitness.xlsx`. La descarga desde la web contiene solo los registros de la cuenta actual sin credenciales; no se debe usar para reemplazar la base completa.

## 7. Cuenta, sesión, clave, PIN y UID

La **cuenta** es un registro permanente en Excel. La **sesión del navegador** identifica qué cuenta está usando ese navegador. Flask guarda su identificador en una cookie firmada. La firma detecta modificaciones, pero no vuelve secreto el contenido: no ponemos contraseñas ni PIN dentro de la cookie.

La **contraseña** permite entrar con usuario y clave. El **UID** lo lee el RC522 de la tarjeta. El **PIN** se escribe en la página después de presentar la tarjeta. No se graba el PIN en el llavero.

La asociación consiste en guardar el UID en `Perfiles → rfid_uid`, en la misma fila de la cuenta. Cuando se vuelve a presentar esa tarjeta, se encuentra su perfil y se verifica el PIN.

`generate_password_hash()` genera una representación protegida. `check_password_hash()` comprueba si el valor escrito coincide. Un hash no se descifra para recuperar la clave original. No reemplaces hashes por contraseñas visibles en Excel.

`SECRET_KEY` firma las sesiones. Si no se configura una fija, se genera al iniciar: las sesiones dejan de ser válidas después de reiniciar. Una clave fija debe ser larga, privada y no publicarse en GitHub.

## 8. Cómo se vincula una tarjeta

```text
Cuenta con sesión iniciada
        ↓
Ajustes → Asociar una tarjeta
        ↓
Flask reserva el lector para ese navegador y para "asociar"
        ↓
El RC522 detecta el UID y Python lo conserva temporalmente
        ↓
El usuario confirma en la página
        ↓
Flask verifica dueño, propósito y vencimiento de la lectura
        ↓
Comprueba que otra cuenta no tenga esa tarjeta
        ↓
Guarda el UID en la cuenta dentro de Excel
```

Hay 45 segundos para presentar la tarjeta y otros 45 desde su detección para confirmar. Una lectura usada no puede confirmarse dos veces. Si ya tenés una tarjeta, la interfaz permite desvincularla antes de asociar otra.

El campo oculto `lectura` contiene un comprobante temporal, no el UID. El servidor comprueba que corresponda a una lectura real de esa operación. Escribir un UID inventado en un formulario no sirve para entrar.

## 9. Cómo se entra con RFID y PIN

```text
Ingresar → Tarjeta RFID y PIN
        ↓
Reserva para ese navegador y para "ingresar"
        ↓
Lectura del UID → búsqueda de la cuenta
        ↓
Formulario del PIN
        ↓
Verificación de la lectura y comparación contra pin_hash
        ↓
Se consume la lectura y se guarda el perfil en la sesión
        ↓
index.html muestra los datos de esa cuenta
```

Un PIN incorrecto no inicia sesión. Después de cinco fallos la credencial se bloquea hasta que se completen cinco minutos desde el primer fallo. El contador está en memoria: se pierde al detener el programa.

Un solo lector físico atiende un navegador a la vez. Otro navegador ve ocupado y no recibe el UID del primero. Cancelar libera la reserva; también vence sola. Una pestaña abierta no mantiene una reserva indefinidamente por recargarse.

El UID puede clonarse. El PIN añade una verificación, pero este trabajo escolar no es un sistema de acceso certificado.

## 10. El hilo del lector y las funciones nuevas

Un **hilo** permite consultar el lector mientras Flask atiende páginas. `_hilo_lector_rfid()` usa `read_id_no_block()`: devuelve un UID si detecta una tarjeta o `None` si no. Solo necesitamos el identificador; no leemos ni escribimos la memoria de la tarjeta.

`RFID_STATE` es un diccionario en memoria: disponibilidad, navegador propietario, motivo, comprobantes, UID y vencimiento. `RFID_LOCK` impide que dos hilos lo modifiquen al mismo tiempo. `time.monotonic()` mide los plazos sin depender de cambios de hora del sistema.

| Función | Responsabilidad |
| --- | --- |
| `estado_lector()` | Reservar/consultar una lectura del navegador y propósito. |
| `registrar_lectura()` | Recibir un UID del hardware físico. Conserva la primera tarjeta de la operación. |
| `uid_confirmado()` | Comprobar dueño, motivo, comprobante y vigencia; consumir si corresponde. |
| `limpiar_uid_reciente()` | Cancelar solo la lectura del navegador actual. |
| `_vencer_lectura()` / `_vaciar_lectura()` | Limpiar el estado dentro del bloqueo. |
| `iniciar_lector_rfid()` | Iniciar una sola tarea física en el modo real. |
| `comprobar_intentos()` / `registrar_fallo()` | Controlar intentos fallidos. |

El hilo empieza al ejecutar `python app.py`, no al importar el archivo desde las pruebas. No se usa recargador ni depuración. `.lector_rfid.lock` coordina nuestros dos scripts. No se deben iniciar varios procesos o workers: cada proceso tendría su propia memoria y no compartiría la operación RFID.

## 11. Actualización sin JavaScript

Mientras se espera una tarjeta real, el HTML incluye:

```html
<meta http-equiv="refresh" content="2">
```

Cada dos segundos pide otra versión de la página. Flask comprueba la lectura. Cuando detecta tarjeta, deja de incluir la recarga para permitir escribir el PIN. 

El temporizador muestra un documento aparte dentro de un `iframe`. Solo se recarga ese marco, no el formulario del entrenamiento. Python calcula cuánto falta a partir del instante de finalización.

## 12. El resto de app.py por grupos

| Grupo | Funciones y tarea |
| --- | --- |
| Validación | `cadena`, `numero`, `fecha_form`, `validar_pin`: comprobar los campos. |
| Preparación | `datos`, `vista`, `helpers`, `volver`, `nuevo_id`, `nueva_fila`: filtrar por cuenta, preparar HTML, redirigir y crear identificadores. |
| Protección | `proteger_formularios` verifica CSRF; `exigir_cuenta` exige sesión. Los manejadores `error_*`, `archivo_grande`, `sesion_vencida` muestran errores esperados. |
| Cuentas | `crear_cuenta`, `entrar`, `entrar_rfid`, `salir`: registro e ingreso/cierre de sesión. |
| Ajustes | `ajustes`, `guardar_ajustes`, `actualizar_perfil`, `cambiar_pin`: preferencias del perfil. |
| Tarjetas | `iniciar_asociacion_rfid`, `asociar_rfid`, `desvincular_rfid` y las funciones `cancelar_*`: administrar tarjeta y espera. |
| Pantallas | `inicio`, `rutinas`, `entrenar`, `progreso`, `alimentacion`, `pagina_asistente`: preparar cada página. |
| Planificación | `guardar_dia`, `guardar_ejercicio`, `eliminar_ejercicio`: editar el plan semanal. |
| Entrenamiento | `estado_entreno`, `marcar_serie`, `tecnica_serie`, `guardar_sesion`: conservar borrador y guardar las series hechas. |
| Progreso | `registrar_peso`, `grafico`: registrar el peso y producir puntos del gráfico SVG. |
| Comidas | `buscar_alimento`, `registrar_comida`: buscar valores y guardar registros. |
| Asistente | `responder_guiado`, `preguntar`: responder con reglas y guardar mensajes. |
| IA opcional | `modelo_local`, `consultar_ollama`, `analizar_foto`, `estimar_foto`: conectarse a un modelo local instalado aparte. |
| Apariencia | `carpeta_fotos`, `subir_foto`, `quitar_foto`, `ver_foto`, `cambiar_tema`: foto y tema de la cuenta. |
| Reloj | `estado_reloj`, `estado_temporizador`, `reloj_descanso`, `descanso`, `controlar_descanso`: gestionar descanso. |
| Exportación | `descargar_excel`: descargar registros propios sin credenciales. |

Las acciones de editar/eliminar comprueban también `perfil_id`. Al cambiar de cuenta se limpia el borrador anterior del entrenamiento.

**CSRF:** cada formulario manda un token que debe coincidir con la sesión. Ayuda a impedir que otra página envíe acciones por el usuario. Ese token no demuestra una lectura de tarjeta: la validación RFID es otra comprobación.

## 13. Dos ejemplos del gimnasio

**Rutina:** lunes es `dia=0` y domingo `dia=6`. Los ejercicios se guardan con día y perfil. Por eso el lunes siguiente aparece el mismo plan. Una sesión realizada se guarda por fecha; cambiar el plan no borra su historial.

**Calorías:** si la tabla indica 130 kcal por 100 g y cargás 200 g, el cálculo es `200 × 130 / 100 = 260 kcal`. Es orientativo según cantidades y preparación. El análisis por foto requiere una IA local instalada y es aproximado.

El asistente guiado usa palabras de tu consulta y tus registros mediante reglas: no es IA generativa. La integración con Ollama es opcional y necesita recursos y un modelo compatible; no se probó su rendimiento en Pi 3.

## 14. Ensayo para presentar el trabajo

1. Mostrá la carpeta y explicá `app.py`, `templates`, `static` y Excel.
2. Abrí un formulario en `ingresar.html` y buscá su ruta `action` en `app.py`.
3. Creá una cuenta de prueba y mostrá que el PIN se guarda como hash.
4. Asociá tu llavero al RC522 de la Raspberry, cerrá sesión y entrá con llavero y PIN.
5. Guardá un ejercicio del lunes y volvé a la página para demostrar persistencia.
6. Explicá que el lector detecta el UID, Excel relaciona ese UID con la cuenta y el servidor comprueba el PIN.

Preguntas para practicar:

| Pregunta | Respuesta |
| --- | --- |
| ¿Dónde están las rutinas? | En `Dias` y `Ejercicios`, relacionadas con el perfil. |
| ¿Qué lee el RC522? | El UID de una tarjeta compatible; el PIN se escribe en la página. |
| ¿Qué grabamos en la tarjeta? | Nada. La relación tarjeta/cuenta se guarda en Excel. |
| ¿Por qué Flask? | Procesa plantillas, formularios, Excel y el lector. |
| ¿Por qué una tarjeta no basta? | Su UID puede copiarse; se exige PIN. |
| ¿Un hash se descifra? | No; se compara el valor escrito mediante una función de verificación. |
| ¿GET y POST? | GET pide una pantalla; POST envía una acción o datos. |
| ¿Para qué el hilo? | Para consultar el lector sin detener la atención de páginas. |
| ¿La revisión del código confirma los cables? | No. La lectura física se comprueba en el colegio. |
| ¿Qué es localhost? | El equipo del propio navegador; otra PC usa la IP de la Raspberry. |
| ¿Por qué no Live Server? | El proyecto necesita ejecutar Python, no solo servir HTML estático. |

## 15. Comprobación pendiente en el colegio

El hardware real necesita Raspberry Pi 3, SPI, pines, alimentación, bibliotecas y tarjetas compatibles. Primero ejecutá `probar_lector.py`, detenelo y después iniciá `app.py`. Las pruebas automáticas no pueden garantizar esas condiciones.

HTTP no cifra claves en tránsito: usá cuentas del proyecto y una red de confianza. No expongas esta entrega directamente a Internet. No hay recuperación automática de claves, múltiples lectores ni memoria compartida entre varios procesos. Para uso permanente faltaría preparar HTTPS y un servidor adecuado.

Fuentes: [SPI en Raspberry Pi](https://www.raspberrypi.com/documentation/computers/configuration.html#enable-or-disable-spi), [biblioteca RC522](https://github.com/pimylifeup/MFRC522-python/blob/master/mfrc522/SimpleMFRC522.py), [guía oficial de Flask](https://flask.palletsprojects.com/en/stable/quickstart/).


## 16. Cambios de organización y contador configurable

**Desplegables:** Entrenar y Ajustes usan las etiquetas HTML `details` y `summary`. El navegador abre o cierra el contenido al tocar el título. El atributo `open` indica si empieza abierto. CSS gira la flecha con el selector `[open]`; no se ejecuta JavaScript. Al marcar una serie, Flask redirige a `Entrenar?abrir=ID` y la plantilla vuelve a abrir ese ejercicio.

**Registro real:** el formulario pregunta qué rutina hiciste, en qué fecha, cuánto duraste y una nota opcional. `guardar_sesion()` escribe el nombre en la nueva columna `Sesiones → rutina`. Es una copia del nombre de ese momento: cambiar el plan después no cambia el historial. Las sesiones antiguas conservan sus registros; si no tenían título, muestran «Entrenamiento».

**Descanso por cuenta:** ya no se pide descanso al agregar o editar un ejercicio. Se configura en Entrenar, arriba, con minutos y segundos. `duracion_descanso()` lee `Perfiles → descanso_s`; si todavía no hay una preferencia, usa 180 segundos. Se admite entre 1 segundo y 30 minutos.

**Contador en otras páginas:** `base.html` muestra un iframe compartido cuando entrás a Entrenar o ya tenés un reloj iniciado. Se fija arriba, debajo de la barra de navegación. La sesión guarda `duracion`, `restante` y `fin`. `fin` es un instante calculado con `time.time() + segundos`. Abrir otra página no lo cambia: el siguiente documento vuelve a mostrar `fin - hora_actual`. Aunque cambies de página o cierres la pestaña y la abras con la misma sesión, el instante final es el mismo.

Pausar conserva lo restante. Reanudar continúa ese mismo descanso. Reiniciar prepara otro. Cambiar la preferencia no reinicia un descanso en marcha: se aplica al siguiente. Marcar Hecha usa la duración elegida y guarda el borrador de los campos; Guardar entrenamiento escribe la sesión en Excel. Los cambios de campos que todavía no se enviaron mediante un formulario no se guardan automáticamente.

**Migración de Excel:** no se borraron hojas ni columnas. Se agregan al final `descanso_s` de Perfiles y `rutina` de Sesiones. `validar_libro()` comprueba que los encabezados anteriores sean el inicio del esquema nuevo. `guardar_migracion()` guarda un respaldo y un archivo temporal antes de reemplazar el libro. La metadata de las tablas de Excel se amplía también. Las pruebas revisan que se conserven todas las celdas anteriores.

**Volver:** `registrar_navegacion()` conserva hasta doce direcciones de páginas principales en la sesión. No registra las recargas del iframe. `/volver` quita la página actual y vuelve a la anterior. Si todavía no hay historial, abre Inicio. Así el botón de la esquina superior izquierda funciona sin `history.back()` ni JavaScript.

Para mostrar estos cambios: configurá un descanso de 1 min 15 s, marcá una serie, abrí Progreso y observá que el contador sigue. Volvé a Entrenar, completá «Pecho y bíceps», fecha, duración y nota. Guardá el entrenamiento y abrí su entrada en el historial. Después cambiá el título del plan: el registro guardado debe conservar su nombre original.


Contador: la × lo oculta y libera el espacio superior. Sigue oculto al cambiar de página. En Entrenar, Mostrar contador de descanso permite recuperarlo. Marcar una nueva serie como Hecha también vuelve a mostrarlo. Ocultarlo no borra la duración configurada ni reinicia el reloj. Es un formulario POST con destino _top para actualizar la página que contiene el iframe, sin JavaScript.


## Entrega final en Linux y servidor para varias computadoras

Se utiliza Waitress para servir Flask con un solo proceso y ocho hilos. `iniciar_lector_rfid()` prepara el RC522 al iniciar `app.py`. El bucle no consulta tarjetas mientras no haya una solicitud. Pulsar Asociar verifica que la cuenta tenga un PIN, reserva el lector y permite que el hilo consulte `read_id_no_block()`. La confirmación vuelve a validar propietario, operación y vigencia antes de guardar el UID junto a la cuenta en Excel.

Linux ejecuta el mismo código Python. GPIO/SPI se importan solamente dentro de la tarea del lector. En una PC sin ese hardware se informa que el lector no está disponible y sigue funcionando el acceso con contraseña. No hay tarjetas ficticias ni una ruta web para inventar lecturas.

La Raspberry guarda una única base central. Las otras computadoras solo necesitan un navegador y la IP de la Raspberry. Cada navegador tiene su sesión; una cuenta no obtiene datos de otra. El lector se comparte de a una operación por vez. El procedimiento completo está en INSTALAR-EN-RASPBERRY.md.
