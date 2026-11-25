# FatToFit: instalación en Raspberry Pi y acceso desde otras computadoras

Esta entrega está preparada para **Raspberry Pi 3 con Raspberry Pi OS, lector RC522 por SPI y clientes Linux, Windows o celulares en la misma red**. No hace falta instalar Python en las computadoras que solo abrirán la página.

Se conservaron la cuenta, las contraseñas/PIN protegidos, rutinas, registros y fotos de la vista de Windows. No hay modo de tarjetas ficticias en esta versión. Las asociaciones se hacen con el lector físico.

## 1. Copiar el proyecto completo

1. Descargá y extraé `FatToFit-Final.zip`.
2. Con FileZilla o un pendrive, copiá la carpeta **FatToFit-Final** a la carpeta personal del usuario de la Raspberry. Debe quedar `~/FatToFit-Final/app.py`.
3. Si la Raspberry tiene una versión anterior en ejecución, detenela con **Ctrl+C** en su terminal antes de actualizar. Si ya tiene datos más nuevos que los de Windows, respaldá su `fitness.xlsx` y `fotos` y conservá esos archivos al actualizar.
4. No copies un entorno `venv` de Windows. El código sirve en Linux; los paquetes se instalan de nuevo para Linux.

En el escritorio de la Raspberry podés abrir una terminal directamente o entrar mediante VNC. FileZilla transfiere los archivos; VNC muestra el escritorio. Se necesita acceso a Internet en la Raspberry durante la instalación de paquetes.

## 2. Revisar la conexión del RC522

**Apagá la Raspberry antes de conectar o cambiar cables.** Los números de la tabla son posiciones físicas del conector de 40 pines, no números BCM.

| RC522 | Pin físico de Raspberry Pi 3 | Función |
| --- | --- | --- |
| VCC / 3.3V | 1 | 3.3 V. Nunca conectarlo a 5 V. |
| GND | 6 | Tierra |
| RST | 22 | GPIO25 |
| MISO | 21 | GPIO9 |
| MOSI | 19 | GPIO10 |
| SCK | 23 | GPIO11 |
| SDA / SS | 24 | GPIO8 / CE0 |
| IRQ | Sin conectar | No se utiliza |

En este módulo, SDA/SS es la selección de SPI; no se conecta al bus I2C. El llavero tiene que ser compatible con RC522: su color azul por sí solo no asegura la compatibilidad.

## 3. Habilitar SPI

Encendé la Raspberry y ejecutá:

```bash
sudo raspi-config
```

Elegí **Interface Options → SPI → Yes → Finish**. Si solicita reiniciar, aceptá. También podés reiniciar con:

```bash
sudo reboot
```

Después de reiniciar, abrí otra terminal y comprobá:

```bash
ls /dev/spidev*
```

Debe existir **`/dev/spidev0.0`**, que usa el proyecto. Puede aparecer también `/dev/spidev0.1`. Si no aparece el primero, revisá la activación de SPI antes de continuar.

## 4. Instalar los paquetes de Linux y el entorno de Python

Estos comandos se ejecutan en la **Raspberry**, dentro de una terminal. El primero puede pedir la contraseña del usuario de Linux:

```bash
sudo apt update
sudo apt install python3-venv python3-dev build-essential python3-rpi.gpio python3-spidev
cd ~/FatToFit-Final
python3 -m venv venv --system-site-packages
source venv/bin/activate
python -m pip install -r requirements.txt
python -m pip install -r requirements-rfid.txt
```

`venv` es un entorno local de Python. `source` lo activa. `--system-site-packages` permite usar los controladores de GPIO instalados por el sistema. El archivo `requirements.txt` contiene Flask, Excel, imágenes, bloqueo de archivos y Waitress. `requirements-rfid.txt` contiene las bibliotecas del RC522 para Raspberry Pi 3.

No instales las bibliotecas de GPIO en las computadoras cliente. Si aparece un error de permisos para GPIO o SPI, revisá con el docente los grupos del usuario y los permisos de `/dev/spidev0.0`. En Raspberry Pi OS suelen utilizarse los grupos `gpio` y `spi`; si el docente confirma que faltan, puede agregar al usuario con:

```bash
sudo usermod -aG gpio,spi "$USER"
```

Después hay que cerrar sesión y entrar nuevamente o reiniciar. No se necesita ejecutar toda la aplicación con `sudo`.

## 5. Comprobar el lector por separado

Con la aplicación detenida y el entorno activado:

```bash
cd ~/FatToFit-Final
source venv/bin/activate
python probar_lector.py
```

Acercá tu llavero azul al RC522. Debe aparecer **`UID detectado: ...`**. Este programa solo lee el identificador: no escribe la tarjeta ni guarda cuentas. Retirá el llavero y terminá con **Ctrl+C**.

Si no aparece ningún UID, revisá alimentación, cables, SPI y compatibilidad del llavero. Que el LED encienda o que la biblioteca arranque no demuestra una lectura correcta. No continúes atribuyendo el problema a la página hasta comprobar este paso.

No ejecutes `probar_lector.py` al mismo tiempo que `app.py`: ambos usan el mismo dispositivo.

## 6. Iniciar la aplicación final

```bash
cd ~/FatToFit-Final
source venv/bin/activate
python app.py
```

La terminal debe mostrar algo parecido a:

```text
[WEB] FatToFit iniciado en el puerto 5001. Para salir: Ctrl+C.
[RFID] RC522 iniciado. Solicitá una lectura desde la página.
```

El orden de esas líneas puede variar. Si aparece `[RFID] No disponible`, el acceso con usuario y contraseña sigue funcionando, pero hay que corregir la instalación/conexión antes de asociar tarjetas.

La web usa **Waitress en un solo proceso con ocho hilos**. No hace falta lanzar otro servidor, activar debug ni usar Live Server. No inicies varias copias de `app.py`, ni lo ejecutes mediante varios workers: el único lector debe tener un solo proceso propietario.

En el navegador de la propia Raspberry abrí:

```text
http://127.0.0.1:5001
```

Mantené la terminal abierta y la Raspberry encendida mientras lo usen. Cerrar esa terminal o pulsar Ctrl+C detiene la web. Para inicio automático, consultá el apartado opcional al final.

## 7. Asociar el llavero a tu cuenta

1. Ingresá con el usuario y contraseña que ya usabas. Las personas nuevas pueden crear su propia cuenta con nombre, contraseña y PIN.
2. Abrí **Ajustes → Llavero azul y PIN**.
3. Elegí un PIN numérico de **4 a 8 cifras** y pulsá **Guardar PIN**. Si tu cuenta ya tiene uno y lo recordás, podés conservarlo.
4. Pulsá **Asociar una tarjeta o llavero**. En ese momento se reserva la lectura para TU navegador durante 45 segundos. El hilo del lector comienza a consultar las tarjetas para esa operación.
5. Acercá el llavero al **lector conectado a la Raspberry**. La página se actualiza cada dos segundos mientras espera.
6. Cuando aparezca detectado, pulsá **Confirmar asociación**. Tenés 45 segundos desde la detección para confirmar.
7. La app guarda el UID en la fila de tu cuenta en `fitness.xlsx`, junto al hash de tu PIN. No graba el PIN en el llavero y no hace falta otro comando de terminal para cada persona.
8. Cerrá sesión. Elegí **Tarjeta RFID y PIN**, acercá ese llavero al RC522 y escribí tu PIN. Debe abrir TU cuenta.

El RC522 lee el **UID**, no un PIN del usuario. El PIN se escribe en la web y se comprueba contra su hash. Un mismo llavero no puede asociarse a dos cuentas distintas. Si querés cambiarlo, desvinculá el anterior desde Ajustes y repetí el procedimiento.

Si otro navegador está usando el lector, aparece ocupado: esperá, cancelá desde ese navegador o dejá vencer la reserva. Una lectura no puede usarse desde otro navegador ni repetirse después de consumida.

## 8. Entrar desde las computadoras del colegio

En otra terminal de la Raspberry, sin detener el servidor, ejecutá:

```bash
hostname -I
```

Usá la IP de su conexión a la red del colegio. Si aparece, por ejemplo, `192.168.1.50`, las otras computadoras deben abrir:

```text
http://192.168.1.50:5001
```

Reemplazá la IP del ejemplo por la real. **`127.0.0.1` no sirve desde otra computadora**: significa el equipo en el que está ese navegador, no la Raspberry.

- Raspberry y computadoras deben estar en una red que permita comunicarse entre ellas. La misma WiFi puede tener aislamiento entre clientes; en ese caso pedí al docente o administrador que permita la conexión local al puerto TCP 5001.
- Los clientes solo necesitan Firefox, Chrome u otro navegador. No se instala Flask, Python ni Excel en cada PC, y no se copia la base a cada una.
- Cada persona crea o usa SU cuenta. La base central permanece en la Raspberry y filtra los registros por usuario.
- Para usar RFID desde cualquier PC, la persona tiene que acercar el llavero al lector FÍSICO de la Raspberry. El navegador de otra PC no incorpora un lector propio.
- Varias personas pueden usar la web y guardar registros; el único lector se utiliza de a una operación por vez.
- Si la IP cambia al reiniciar, hay que consultar la nueva. Para mantener una dirección fija, el administrador de la red puede reservar una IP para la Raspberry.

HTTP no cifra las credenciales en tránsito: usá claves específicas del proyecto y una red escolar de confianza. No abras este puerto directamente a Internet. Acceso desde fuera del colegio y HTTPS requieren una configuración aparte.

## 9. Si algo no funciona

| Síntoma | Qué revisar |
| --- | --- |
| No existe `/dev/spidev0.0` | SPI en raspi-config y reinicio. |
| `No module named RPi` o `mfrc522` | Activar el venv correcto e instalar requirements-rfid.txt. |
| Permiso denegado sobre GPIO/SPI | Grupos y permisos del usuario con el docente; volver a entrar tras agregar grupos. |
| El lector está en uso | Detener otro app.py o probar_lector.py. No basta con cerrar el navegador. |
| Enciende el LED pero no detecta | Cables, alimentación 3.3 V y llavero compatible. Ejecutar la comprobación física del paso 5. |
| Tarjeta no asociada | Ingresar con contraseña y vincularla desde Ajustes. |
| PIN incorrecto | Es el PIN elegido en tu cuenta, no los dígitos impresos del llavero. Tras 5 fallos hay espera temporal. |
| Lectura vencida | Iniciar otra lectura y presentar nuevamente el llavero. |
| La Pi abre la web, otra PC no | IP real, puerto 5001, red compartida, aislamiento WiFi y firewall de la red. |
| Puerto ya ocupado | Detener la versión anterior; no lanzar dos servidores. |
| No se guardan cambios | Cerrar fitness.xlsx en Excel/LibreOffice antes de guardar desde la web y revisar permisos de la carpeta. |

## 10. Revisión que conviene mostrar al docente

1. Un llavero detectado por `probar_lector.py`.
2. Asociación a una cuenta, cierre de sesión e ingreso con llavero + PIN.
3. PIN incorrecto rechazado y otro llavero sin asociar rechazado.
4. Dos navegadores con cuentas distintas: cada uno ve sus registros.
5. Una rutina creada desde una computadora y recuperada tras volver a ingresar.
6. Un entrenamiento guardado con nombre, fecha, duración y nota.
7. Contador configurable, visible al navegar y ocultable con ×.
8. Un respaldo de `fitness.xlsx` y las fotos con la aplicación detenida.

El código fue comprobado con clientes de Flask y un controlador sustituto. Esa revisión no confirma por sí sola cables, SPI ni el RC522 físico; los pasos 5 y 7 completan esa verificación en el colegio.

## Opcional: inicio automático en Linux

Este apartado es para hacerlo con el docente cuando el funcionamiento manual ya esté comprobado. Antes, detené el servidor manual con Ctrl+C. El servicio y la ejecución manual no deben estar activos a la vez.

```bash
mkdir -p ~/.config/systemd/user
nano ~/.config/systemd/user/fattofit.service
```

Pegá lo siguiente y guardá el archivo:

```ini
[Unit]
Description=FatToFit

[Service]
WorkingDirectory=%h/FatToFit-Final
ExecStart=%h/FatToFit-Final/venv/bin/python %h/FatToFit-Final/app.py
Restart=on-failure
RestartSec=5

[Install]
WantedBy=default.target
```

Luego:

```bash
systemctl --user daemon-reload
systemctl --user enable --now fattofit
systemctl --user status fattofit
```

Para que siga disponible aunque ese usuario no haya iniciado sesión, el docente puede habilitar:

```bash
sudo loginctl enable-linger "$USER"
```

Para ver mensajes, detener o reiniciar:

```bash
journalctl --user -u fattofit -f
systemctl --user stop fattofit
systemctl --user restart fattofit
```

Ctrl+C sale de la lectura de mensajes, no detiene el servicio. Si se quiere ejecutar `probar_lector.py`, primero hay que detener el servicio. Al reiniciar la app las sesiones de navegador pueden vencer y habrá que volver a ingresar; los registros de Excel permanecen.

Fuentes: [SPI en Raspberry Pi](https://www.raspberrypi.com/documentation/computers/configuration.html#enable-or-disable-spi), [biblioteca RC522](https://github.com/pimylifeup/MFRC522-python/blob/master/mfrc522/SimpleMFRC522.py), [Waitress](https://docs.pylonsproject.org/projects/waitress/en/stable/usage.html).
