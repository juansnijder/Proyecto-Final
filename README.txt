FATTOFIT — ENTREGA FINAL

Carpetas principales: app.py, templates, static y fitness.xlsx.
Conserva la cuenta, rutinas, registros y fotos que cargaste en Windows.
No incluye tarjetas ficticias ni carteles de demostración. Sin JavaScript.

GUÍAS
INSTALAR-EN-RASPBERRY.md: instalación Linux, cables, SPI, lector, cuentas,
acceso desde otras computadoras y arranque automático opcional.
GUIA-PARA-EXPLICAR.md: explicación del código y las carpetas para estudiar.
COMPROBACIONES.txt: resultados y límites de la comprobación.

RASPBERRY PI 3
Copiar la carpeta a ~/FatToFit-Final. Luego, por terminal:
    sudo apt update
    sudo apt install python3-venv python3-dev build-essential python3-rpi.gpio python3-spidev
    cd ~/FatToFit-Final
    python3 -m venv venv --system-site-packages
    source venv/bin/activate
    python -m pip install -r requirements.txt
    python -m pip install -r requirements-rfid.txt

Habilitar SPI con sudo raspi-config > Interface Options > SPI > Yes.
Tras reiniciar si se solicita, comprobar que existe /dev/spidev0.0.
La guía explica las conexiones: siempre VCC a 3.3 V, nunca 5 V.

PRIMERO EL LECTOR
Con app.py detenido:
    python probar_lector.py
Acercar el llavero y verificar UID detectado. Terminar con Ctrl+C.

INICIAR LA WEB
    python app.py
No usar Live Server, file://, Flask debug ni varios procesos.
Waitress atiende las páginas; el hilo de Python atiende el único lector.
En la propia Raspberry: http://127.0.0.1:5001
Desde otras PCs: http://IP_DE_LA_RASPBERRY:5001
La IP se consulta con hostname -I. Todos deben poder comunicarse por la red.
Solo la Raspberry ejecuta la aplicación y conserva la base de Excel.

ASOCIAR LLAVERO
Entrar con usuario y contraseña > Ajustes > Llavero azul y PIN > Guardar PIN >
Asociar una tarjeta o llavero > acercarlo al RC522 de la Raspberry >
Confirmar asociación. Después, cerrar sesión y entrar con llavero + PIN.
El PIN se escribe en la web; el lector lee el UID. No escribe la tarjeta.
El lector atiende una operación a la vez, reservada al navegador que la pidió.

WINDOWS PARA VER LAS PÁGINAS
    python -m pip install -r requirements.txt
    python app.py
En el navegador: http://127.0.0.1:5001
El lector no estará disponible sin la Raspberry, pero usuario y contraseña
permiten entrar. Los avisos de conexión real del lector se conservan.

DATOS
Cerrar Excel/LibreOffice antes de guardar desde la web.
Respaldos automáticos: copias_excel. Fotos: fotos. Ambos junto a fitness.xlsx.
Si la Raspberry tiene datos más nuevos, respaldar y conservar SU Excel y fotos
al actualizar. La descarga de Excel desde la página es personal, sin claves;
no sustituye el respaldo completo del servidor.
No exponer directamente a Internet: HTTP no cifra claves en tránsito.

COMPROBACIONES DEL CÓDIGO
    python pruebas.py
Trabajan sobre copias temporales. No requieren conectar el RC522 ni cambian
la base original. La comprobación física se completa en el colegio.
