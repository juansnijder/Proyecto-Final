"""Prueba física del RC522, sin iniciar Flask ni abrir Excel.

En la Raspberry: detener app.py y ejecutar python probar_lector.py.
Ctrl+C termina la prueba. No escribe nada en las tarjetas.
"""
from pathlib import Path
from time import sleep
from filelock import FileLock, Timeout


def main():
    gpio = None
    bloqueo = FileLock(str(Path(__file__).resolve().parent / '.lector_rfid.lock'))
    adquirido = False
    try:
        bloqueo.acquire(timeout=0)
        adquirido = True
        import RPi.GPIO as GPIO
        from mfrc522 import SimpleMFRC522
        gpio = GPIO
        lector = SimpleMFRC522()
        print('Acercá una tarjeta. Ctrl+C para salir. No se modifica la tarjeta.')
        anterior = None
        while True:
            uid = lector.read_id_no_block()
            if uid is not None and uid != anterior:
                print(f'UID detectado: {uid}', flush=True)
            anterior = uid
            sleep(0.2)
    except Timeout:
        print('El lector ya está en uso. Detené app.py u otra prueba primero.')
        return 1
    except KeyboardInterrupt:
        print('\nPrueba terminada.')
        return 0
    except Exception as exc:
        print(f'No se pudo probar el RC522: {exc}')
        print('Revisá SPI, dependencias y conexiones indicadas en LEEME.txt.')
        return 1
    finally:
        try:
            if gpio is not None:
                gpio.cleanup()
        finally:
            if adquirido:
                bloqueo.release()


if __name__ == '__main__':
    raise SystemExit(main())
