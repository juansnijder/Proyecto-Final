"""Pruebas sin Raspberry: python pruebas.py. Solo modifican copias temporales.

El cliente de Flask hace peticiones como un navegador. registrar_lectura simula
la entrada del hardware; no simula que la conexión SPI haya sido comprobada.
"""
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from datetime import date
from html.parser import HTMLParser
import hashlib
import io
import shutil
import types
import unittest
import time
import app as modulo


class AnalizarHTML(HTMLParser):
    def __init__(self):
        super().__init__()
        self.campos = {}
        self.scripts = []
        self.detalles = []

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == 'input' and attrs.get('name'):
            self.campos[attrs['name']] = attrs.get('value', '')
        if tag == 'details':
            self.detalles.append(attrs)
        if tag == 'script' or any(k.startswith('on') for k in attrs):
            self.scripts.append((tag, attrs))
        if any(str(v).lower().startswith('javascript:') for v in attrs.values()):
            self.scripts.append((tag, attrs))


class PruebasFatToFit(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.original = Path(__file__).resolve().parent / 'fitness.xlsx'
        cls.huella = hashlib.sha256(cls.original.read_bytes()).hexdigest()

    @classmethod
    def tearDownClass(cls):
        assert hashlib.sha256(cls.original.read_bytes()).hexdigest() == cls.huella, 'Se cambió el Excel original'

    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.raiz = Path(self.temp.name)
        self.excel = self.raiz / 'fitness.xlsx'
        shutil.copy2(self.original, self.excel)
        self.nombre_base = 'Alumno-' + modulo.nuevo_id()
        modulo.app.config.update(TESTING=True, SECRET_KEY='solo-para-pruebas-automaticas',
            DATA_FILE=str(self.excel), MODO_PRUEBA_RFID=False)
        with modulo.RFID_LOCK:
            modulo._vaciar_lectura()
            modulo.RFID_STATE.update(disponible=True, error=None, iniciado=False)
        with modulo.INTENTOS_LOCK:
            modulo.INTENTOS.clear()
        self.cliente = self.navegador()

    def navegador(self):
        c = modulo.app.test_client()
        c.get('/ingresar')
        return c

    def post(self, ruta, datos=None, cliente=None, seguir=True, **kwargs):
        c = cliente or self.cliente
        with c.session_transaction() as ses:
            csrf = ses['csrf']
        return c.post(ruta, data={'csrf': csrf, **(datos or {})}, follow_redirects=seguir, **kwargs)

    def cuenta(self, nombre=None, cliente=None, pin='2468'):
        c = cliente or self.cliente
        nombre = self.nombre_base if nombre is None else nombre + '-' + self.nombre_base[-6:]
        response = self.post('/cuentas', {'nombre': nombre, 'clave': 'clave-prueba', 'pin': pin,
            'objetivo': 'Proyecto escolar', 'altura_cm': '', 'meta_peso_kg': ''}, c)
        self.assertEqual(response.status_code, 200)
        with c.session_transaction() as ses:
            return ses['perfil']

    def datos(self):
        return modulo.leer_excel(self.excel)

    def campos(self, response):
        parser = AnalizarHTML()
        parser.feed(response.get_data(as_text=True))
        self.assertFalse(parser.scripts)
        return parser.campos

    def detectar(self, uid='1234567890'):
        with modulo.RFID_LOCK:
            solicitud = modulo.RFID_STATE['solicitud']
        self.assertTrue(modulo.registrar_lectura(uid, solicitud))

    def asociar(self, uid='1234567890', cliente=None):
        c = cliente or self.cliente
        self.post('/ajustes/rfid/iniciar', cliente=c)
        self.detectar(uid)
        fields = self.campos(c.get('/ajustes?escanear=1'))
        self.assertIn('lectura', fields)
        return self.post('/ajustes/rfid/asociar', {'lectura': fields['lectura']}, c)

    def lectura_ingreso(self, cliente=None, uid='1234567890'):
        c = cliente or self.cliente
        c.get('/ingresar?modo=rfid')
        self.detectar(uid)
        return self.campos(c.get('/ingresar?modo=rfid'))['lectura']

    def test_01_cuenta_hashes_y_password(self):
        identifier = self.cuenta()
        p = next(p for p in self.datos()['Perfiles'] if p['id'] == identifier)
        self.assertNotEqual(p['clave_hash'], 'clave-prueba')
        self.assertNotEqual(p['pin_hash'], '2468')
        self.assertTrue(modulo.check_password_hash(p['pin_hash'], '2468'))
        self.post('/salir')
        self.post('/entrar', {'nombre': self.nombre_base, 'clave': 'incorrecta'})
        with self.cliente.session_transaction() as ses:
            self.assertNotIn('perfil', ses)
        self.post('/entrar', {'nombre': self.nombre_base, 'clave': 'clave-prueba'})
        with self.cliente.session_transaction() as ses:
            self.assertEqual(ses['perfil'], identifier)

    def test_02_vincular_cerrar_y_entrar_rfid(self):
        identifier = self.cuenta()
        self.asociar()
        self.post('/salir')
        lectura = self.lectura_ingreso()
        self.post('/entrar/rfid', {'lectura': lectura, 'pin': '2468'})
        with self.cliente.session_transaction() as ses:
            self.assertEqual(ses['perfil'], identifier)
        with modulo.RFID_LOCK:
            self.assertIsNone(modulo.RFID_STATE['uid'])

    def test_03_uid_escrito_en_formulario_no_sirve(self):
        identifier = self.cuenta()
        self.post('/ajustes/rfid/asociar', {'uid': 'inventado'})
        p = next(p for p in self.datos()['Perfiles'] if p['id'] == identifier)
        self.assertFalse(p['rfid_uid'])
        self.asociar()
        self.post('/salir')
        self.post('/entrar/rfid', {'uid': '1234567890', 'pin': '2468'})
        with self.cliente.session_transaction() as ses:
            self.assertNotIn('perfil', ses)

    def test_04_uid_formulario_no_reemplaza_lectura(self):
        self.cuenta()
        self.asociar()
        self.post('/salir')
        lectura = self.lectura_ingreso()
        self.post('/entrar/rfid', {'lectura': lectura, 'uid': 'otra-tarjeta', 'pin': '2468'})
        with self.cliente.session_transaction() as ses:
            self.assertIn('perfil', ses)

    def test_05_lectura_vencida_no_permite_entrar(self):
        self.cuenta()
        self.asociar()
        self.post('/salir')
        lectura = self.lectura_ingreso()
        with modulo.RFID_LOCK:
            modulo.RFID_STATE['vence'] = 0
        self.post('/entrar/rfid', {'lectura': lectura, 'pin': '2468'})
        with self.cliente.session_transaction() as ses:
            self.assertNotIn('perfil', ses)

    def test_06_lectura_consumida_no_se_reutiliza(self):
        self.cuenta()
        self.asociar()
        self.post('/salir')
        lectura = self.lectura_ingreso()
        self.post('/entrar/rfid', {'lectura': lectura, 'pin': '2468'})
        self.post('/salir')
        self.post('/entrar/rfid', {'lectura': lectura, 'pin': '2468'})
        with self.cliente.session_transaction() as ses:
            self.assertNotIn('perfil', ses)

    def test_07_otro_navegador_no_roba_lectura(self):
        self.cuenta()
        self.asociar()
        self.post('/salir')
        lectura = self.lectura_ingreso()
        otro = self.navegador()
        page = otro.get('/ingresar?modo=rfid').get_data(as_text=True)
        self.assertIn('otro navegador', page)
        self.assertNotIn('name="lectura"', page)
        self.post('/entrar/rfid', {'lectura': lectura, 'pin': '2468'}, otro)
        with otro.session_transaction() as ses:
            self.assertNotIn('perfil', ses)
        self.post('/entrar/rfid', {'lectura': lectura, 'pin': '2468'})
        with self.cliente.session_transaction() as ses:
            self.assertIn('perfil', ses)

    def test_08_otro_navegador_no_cancela(self):
        self.cliente.get('/ingresar?modo=rfid')
        with modulo.RFID_LOCK:
            antes = modulo.RFID_STATE['solicitud']
        otro = self.navegador()
        self.post('/ingresar/rfid/cancelar', cliente=otro)
        with modulo.RFID_LOCK:
            self.assertEqual(modulo.RFID_STATE['solicitud'], antes)

    def test_09_tarjeta_no_se_asocia_dos_cuentas(self):
        self.cuenta()
        self.asociar()
        otro = self.navegador()
        identifier = self.cuenta('Segundo alumno', otro)
        response = self.asociar(cliente=otro)
        self.assertIn('ya está asociada', response.get_data(as_text=True))
        p = next(p for p in self.datos()['Perfiles'] if p['id'] == identifier)
        self.assertFalse(p['rfid_uid'])

    def test_10_pin_incorrecto_y_limite(self):
        self.cuenta()
        self.asociar()
        self.post('/salir')
        lectura = self.lectura_ingreso()
        for _ in range(5):
            self.post('/entrar/rfid', {'lectura': lectura, 'pin': '0000'})
        response = self.post('/entrar/rfid', {'lectura': lectura, 'pin': '2468'})
        self.assertIn('Demasiados intentos', response.get_data(as_text=True))
        with self.cliente.session_transaction() as ses:
            self.assertNotIn('perfil', ses)
        with modulo.INTENTOS_LOCK:
            modulo.INTENTOS['pin:1234567890']['vence'] = 0
        lectura = self.lectura_ingreso()
        self.post('/entrar/rfid', {'lectura': lectura, 'pin': '2468'})
        with self.cliente.session_transaction() as ses:
            self.assertIn('perfil', ses)

    def test_11_actualizar_pin_y_desvincular(self):
        identifier = self.cuenta()
        self.asociar()
        self.post('/ajustes/pin', {'pin': '9753'})
        self.post('/salir')
        lectura = self.lectura_ingreso()
        self.post('/entrar/rfid', {'lectura': lectura, 'pin': '2468'})
        with self.cliente.session_transaction() as ses:
            self.assertNotIn('perfil', ses)
        self.post('/entrar/rfid', {'lectura': lectura, 'pin': '9753'})
        self.post('/ajustes/rfid/desvincular')
        self.assertFalse(next(p for p in self.datos()['Perfiles'] if p['id'] == identifier)['rfid_uid'])

    def test_12_simulador_real_desactivado(self):
        response = self.post('/rfid/simular', {'motivo': 'ingresar', 'tarjeta': 'A'})
        self.assertEqual(response.status_code, 404)

    def test_13_version_final_no_expone_simulacion(self):
        self.cuenta()
        rutas = {rule.rule for rule in modulo.app.url_map.iter_rules()}
        self.assertNotIn('/rfid/simular', rutas)
        self.assertFalse(hasattr(modulo, 'simular_rfid'))
        self.assertFalse(hasattr(modulo, 'preparar_simulacion'))

    def test_14_paginas_sin_mensajes_de_demostracion(self):
        self.cuenta()
        textos = ['PRUEBA EN CASA', 'VISTA DE EJEMPLO', 'Tarjeta de prueba',
            'Acercar tarjeta de prueba', 'Los registros de Alex son ficticios']
        for url in ['/ingresar', '/', '/rutinas', '/entrenar', '/progreso',
                '/alimentacion', '/asistente', '/ajustes']:
            html = self.cliente.get(url).get_data(as_text=True)
            for texto in textos:
                self.assertNotIn(texto, html)

    def test_15_paginas_css_y_sin_javascript(self):
        self.cuenta()
        for ruta in ['/', '/rutinas?dia=0', '/entrenar?dia=0', '/progreso', '/alimentacion',
                '/asistente', '/ajustes', '/descanso/reloj', '/ingresar?modo=rfid']:
            with self.subTest(ruta=ruta):
                r = self.cliente.get(ruta)
                self.assertEqual(r.status_code, 200)
                self.campos(r)
        for ruta in ['/static/estilos.css', '/static/movil.css', '/static/manifest.json',
                '/static/icons/icon-192.png']:
            response = self.cliente.get(ruta)
            self.assertEqual(response.status_code, 200)
            response.close()
        root = Path(__file__).resolve().parent
        self.assertFalse(list((root / 'static').rglob('*.js')))

    def test_16_sin_lector_permite_password(self):
        with modulo.RFID_LOCK:
            modulo.RFID_STATE['disponible'] = False
        r = self.cliente.get('/ingresar?modo=rfid')
        self.assertIn('lector no está disponible', r.get_data(as_text=True))
        self.cuenta()
        self.assertEqual(self.cliente.get('/').status_code, 200)

    def test_17_csrf_y_paginas_privadas(self):
        self.assertEqual(self.cliente.get('/ajustes').status_code, 302)
        self.assertEqual(self.cliente.post('/cuentas', data={}).status_code, 400)
        self.assertEqual(self.cliente.get('/descargar-excel').status_code, 302)

    def test_18_reserva_no_se_renueva_sola(self):
        self.cliente.get('/ingresar?modo=rfid')
        with modulo.RFID_LOCK:
            modulo.RFID_STATE['vence'] = 0
        r = self.cliente.get('/ingresar?modo=rfid')
        self.assertIn('espera terminó', r.get_data(as_text=True))
        self.assertIsNone(modulo.RFID_STATE['propietario'])
        self.post('/ingresar/rfid/cancelar')
        self.assertIsNotNone(modulo.RFID_STATE['propietario'])

    def test_19_no_acepta_lectura_anterior_ni_cambia_primera(self):
        self.cliente.get('/ingresar?modo=rfid')
        solicitud = modulo.RFID_STATE['solicitud']
        self.post('/ingresar/rfid/cancelar')
        self.assertFalse(modulo.registrar_lectura('111', solicitud))
        self.detectar('222')
        self.assertFalse(modulo.registrar_lectura('333', modulo.RFID_STATE['solicitud']))
        self.assertEqual(modulo.RFID_STATE['uid'], '222')

    def test_20_tipo_operacion_no_se_puede_mezclar(self):
        self.cuenta()
        self.post('/ajustes/rfid/iniciar')
        self.detectar()
        lectura = self.campos(self.cliente.get('/ajustes?escanear=1'))['lectura']
        self.post('/entrar/rfid', {'lectura': lectura, 'pin': '2468'}, seguir=False)
        self.assertEqual(modulo.RFID_STATE['motivo'], 'asociar')

    def test_21_rutinas_peso_comidas_y_sesion(self):
        identifier = self.cuenta()
        self.post('/rutinas/dia', {'dia': '0', 'titulo': 'Pecho y bíceps', 'activo': 'si'})
        self.post('/rutinas/ejercicio', {'dia': '0', 'nombre': 'Press banca', 'grupo': 'Pecho',
            'series': '3', 'reps': '10', 'descanso_s': '90'})
        ejercicio = next(e for e in self.datos()['Ejercicios'] if e['perfil_id'] == identifier)
        prefijo = ejercicio['id'] + '_1'
        valores = {'dia': '0', 'serie': prefijo, 'reps_' + prefijo: '10',
            'peso_' + prefijo: '20', 'tecnica_' + prefijo: 'bien'}
        self.post('/entrenar/serie', valores)
        self.post('/entrenar/guardar', {**valores, 'fecha': date.today().isoformat(),
            'duracion_min': '35', 'notas': 'Sesión de prueba'})
        self.post('/peso', {'fecha': date.today().isoformat(), 'peso_kg': '70', 'nota': ''})
        self.post('/alimentacion/guardar', {'fecha': date.today().isoformat(), 'momento': 'Almuerzo',
            'tipo': 'manual', 'nombre': 'Plato de prueba', 'gramos': '200', 'kcal': '300', 'proteinas_g': '20'})
        data = self.datos()
        self.assertEqual(len([s for s in data['Sesiones'] if s['perfil_id'] == identifier]), 1)
        self.assertEqual(len([s for s in data['Peso'] if s['perfil_id'] == identifier]), 1)
        self.assertEqual(len([s for s in data['Comidas'] if s['perfil_id'] == identifier]), 1)
        self.assertIn('Pecho y bíceps', self.cliente.get('/rutinas?dia=0').get_data(as_text=True))
        otro = self.navegador()
        self.cuenta('Otro alumno', otro)
        self.assertNotIn('Press banca', otro.get('/rutinas?dia=0').get_data(as_text=True))

    def test_22_login_y_logout_limpian_borradores(self):
        self.cuenta()
        with self.cliente.session_transaction() as ses:
            ses['entreno'] = {'dia': 0, 'valores': {'privado': 'dato'}, 'hechas': []}
            ses['reloj'] = {'duracion': 180}
        self.post('/salir')
        with self.cliente.session_transaction() as ses:
            self.assertNotIn('entreno', ses)
            self.assertNotIn('reloj', ses)

    def test_23_hilo_usa_lectura_id_y_limpia_gpio(self):
        self.cliente.get('/ingresar?modo=rfid')
        gpio = types.ModuleType('RPi.GPIO')
        gpio.setwarnings = lambda value: None
        gpio.cleanup = unittest.mock.Mock()
        rpi = types.ModuleType('RPi')
        rpi.GPIO = gpio
        lector = unittest.mock.Mock()
        lector.read_id_no_block.return_value = 1234567890
        libreria = types.ModuleType('mfrc522')
        libreria.SimpleMFRC522 = lambda: lector
        def detener(_):
            raise KeyboardInterrupt()
        with patch.dict('sys.modules', {'RPi': rpi, 'RPi.GPIO': gpio, 'mfrc522': libreria}), \
                patch.object(modulo, 'ROOT', self.raiz), patch.object(modulo.time, 'sleep', detener), \
                patch('builtins.print'):
            with self.assertRaises(KeyboardInterrupt):
                modulo._hilo_lector_rfid()
        self.assertEqual(modulo.RFID_STATE['uid'], '1234567890')
        lector.read_id_no_block.assert_called_once()
        gpio.cleanup.assert_called_once()

    def test_24_asociacion_exige_pin_configurado(self):
        identifier = self.cuenta()
        modulo.actualizar_perfil(identifier, pin_hash='')
        response = self.post('/ajustes/rfid/iniciar')
        self.assertIn('Guardá primero un PIN', response.get_data(as_text=True))
        self.assertIsNone(modulo.RFID_STATE['solicitud'])

    def test_25_pin_y_numero_invalidos_no_se_guardan(self):
        self.cuenta()
        antes = self.datos()['Perfiles']
        self.post('/ajustes/pin', {'pin': '12ab'})
        self.assertEqual(self.datos()['Perfiles'], antes)
        self.post('/peso', {'fecha': date.today().isoformat(), 'peso_kg': 'NaN'})
        self.assertFalse(any(p['perfil_id'] == antes[-1]['id'] for p in self.datos()['Peso']))

    def test_26_descarga_solo_cuenta_y_sin_credenciales(self):
        identifier = self.cuenta()
        otro = self.navegador()
        self.cuenta('Otra cuenta privada', otro)
        antes = self.excel.read_bytes()
        response = self.cliente.get('/descargar-excel')
        self.assertEqual(response.status_code, 200)
        book = modulo.load_workbook(io.BytesIO(response.data))
        try:
            perfiles = modulo.filas_excel(book, 'Perfiles')
            self.assertEqual(len(perfiles), 1)
            self.assertEqual(perfiles[0]['id'], identifier)
            self.assertFalse(perfiles[0]['clave_hash'])
            self.assertFalse(perfiles[0]['pin_hash'])
            self.assertFalse(perfiles[0]['rfid_uid'])
            self.assertTrue(all(d['perfil_id'] == identifier for d in modulo.filas_excel(book, 'Dias')))
        finally:
            book.close()
            response.close()
        self.assertEqual(self.excel.read_bytes(), antes)

    def ejercicio(self, nombre='Press banca', grupo='Pecho'):
        self.post('/rutinas/ejercicio', {'dia': '0', 'nombre': nombre, 'grupo': grupo,
            'series': '3', 'reps': '10'})
        return next(e for e in self.datos()['Ejercicios'] if e['nombre'] == nombre)

    def test_27_rutina_sin_campo_descanso(self):
        self.cuenta()
        self.ejercicio()
        html = self.cliente.get('/rutinas?dia=0').get_data(as_text=True)
        self.assertNotIn('name="descanso_s"', html)
        self.assertNotIn('s de descanso', html)
        self.assertIn('Press banca', html)

    def test_28_preferencia_descanso_persistente(self):
        identifier = self.cuenta()
        self.post('/descanso', {'accion': 'configurar', 'minutos': '2', 'segundos_extra': '15', 'dia': '0'})
        perfil = next(p for p in self.datos()['Perfiles'] if p['id'] == identifier)
        self.assertEqual(perfil['descanso_s'], 135)
        otro = self.navegador()
        self.post('/entrar', {'nombre': self.nombre_base, 'clave': 'clave-prueba'}, otro)
        html = otro.get('/entrenar?dia=0').get_data(as_text=True)
        self.assertIn('2 min 15 s', html)
        self.assertIn('segundos=135', html)

    def test_29_reloj_sigue_en_otras_paginas(self):
        self.cuenta()
        self.post('/descanso', {'accion': 'iniciar', 'segundos': '120', 'volver': 'reloj'})
        with self.cliente.session_transaction() as ses:
            fin = ses['reloj']['fin']
        for url in ['/rutinas', '/progreso', '/alimentacion', '/asistente', '/ajustes']:
            r = self.cliente.get(url)
            self.assertEqual(r.status_code, 200)
            self.assertIn('reloj-global-frame', r.get_data(as_text=True))
            with self.cliente.session_transaction() as ses:
                self.assertEqual(ses['reloj']['fin'], fin)
        reloj = self.cliente.get('/descanso/reloj').get_data(as_text=True)
        self.assertIn('http-equiv="refresh"', reloj)

    def test_30_configurar_no_reinicia_descanso_activo(self):
        self.cuenta()
        self.post('/descanso', {'accion': 'iniciar', 'segundos': '90'})
        with self.cliente.session_transaction() as ses:
            fin = ses['reloj']['fin']
        self.post('/descanso', {'accion': 'configurar', 'minutos': '4', 'segundos_extra': '0'})
        with self.cliente.session_transaction() as ses:
            self.assertEqual(ses['reloj']['fin'], fin)
            self.assertEqual(ses['reloj']['duracion'], 90)
        self.post('/descanso', {'accion': 'pausar', 'segundos': '240'})
        with self.cliente.session_transaction() as ses:
            restante = ses['reloj']['restante']
        self.post('/descanso', {'accion': 'iniciar', 'segundos': '240'})
        with self.cliente.session_transaction() as ses:
            self.assertEqual(ses['reloj']['duracion'], 90)
            self.assertLessEqual(ses['reloj']['fin'] - time.time(), restante)
        self.post('/descanso', {'accion': 'reiniciar', 'segundos': '240'})
        self.post('/descanso', {'accion': 'iniciar', 'segundos': '240'})
        with self.cliente.session_transaction() as ses:
            self.assertEqual(ses['reloj']['duracion'], 240)

    def test_31_pausa_reanuda_y_fin_reloj(self):
        self.cuenta()
        self.post('/descanso', {'accion': 'iniciar', 'segundos': '90'})
        with self.cliente.session_transaction() as ses:
            reloj = ses['reloj']
            reloj['fin'] = time.time() + 32
            ses['reloj'] = reloj
        self.post('/descanso', {'accion': 'pausar', 'segundos': '90'})
        with self.cliente.session_transaction() as ses:
            restante = ses['reloj']['restante']
            self.assertIsNone(ses['reloj']['fin'])
            self.assertTrue(25 <= restante <= 33)
        self.post('/descanso', {'accion': 'iniciar', 'segundos': '90'})
        with self.cliente.session_transaction() as ses:
            self.assertLessEqual(ses['reloj']['fin'] - time.time(), restante)
            reloj = ses['reloj']
            reloj['fin'] = time.time() - 3
            ses['reloj'] = reloj
        html = self.cliente.get('/descanso/reloj').get_data(as_text=True)
        self.assertIn('¡Descanso terminado!', html)
        self.assertNotIn('http-equiv="refresh"', html)

    def test_32_historial_conserva_rutina_realizada(self):
        identifier = self.cuenta()
        e = self.ejercicio()
        prefijo = e['id'] + '_1'
        campos = {'dia': '0', 'serie': prefijo, 'peso_' + prefijo: '20',
            'reps_' + prefijo: '8', 'tecnica_' + prefijo: 'bien',
            'rutina': 'Pecho y bíceps', 'fecha': date.today().isoformat(),
            'duracion_min': '37', 'notas': 'Usé menos carga en la última serie.'}
        self.post('/entrenar/serie', campos)
        self.post('/entrenar/guardar', campos)
        self.post('/rutinas/dia', {'dia': '0', 'titulo': 'Espalda', 'activo': 'si'})
        sesion = next(s for s in self.datos()['Sesiones'] if s['perfil_id'] == identifier)
        self.assertEqual(sesion['rutina'], 'Pecho y bíceps')
        self.assertEqual(sesion['duracion_min'], 37)
        html = self.cliente.get('/progreso').get_data(as_text=True)
        self.assertIn('Pecho y bíceps', html)
        self.assertIn('37 min', html)
        self.assertIn('Usé menos carga en la última serie.', html)

    def test_33_boton_volver_y_no_registra_iframe(self):
        self.cuenta()
        with self.cliente.session_transaction() as ses:
            ses.pop('navegacion', None)
        self.cliente.get('/rutinas?dia=0')
        self.cliente.get('/entrenar?dia=0')
        self.cliente.get('/progreso')
        self.cliente.get('/descanso/reloj')
        r = self.cliente.get('/volver', follow_redirects=True)
        self.assertEqual(r.request.path, '/entrenar')
        self.assertIn('aria-label="Volver atrás"', r.get_data(as_text=True))
        r = self.cliente.get('/volver', follow_redirects=True)
        self.assertEqual(r.request.path, '/rutinas')

    def test_34_ajustes_desplegables_y_rfid_abierto(self):
        self.cuenta()
        parser = AnalizarHTML()
        parser.feed(self.cliente.get('/ajustes').get_data(as_text=True))
        paneles = [d for d in parser.detalles if 'settings-panel' in d.get('class', '')]
        self.assertEqual(len(paneles), 7)
        self.assertTrue(all('open' not in d for d in paneles))
        r = self.post('/ajustes/rfid/iniciar')
        parser = AnalizarHTML()
        parser.feed(r.get_data(as_text=True))
        seguridad = next(d for d in parser.detalles if d.get('id') == 'seguridad')
        self.assertIn('open', seguridad)
        self.assertIn('Acercá el llavero azul', r.get_data(as_text=True))

    def test_35_migracion_amplia_columnas_sin_perder_filas(self):
        # La entrega ya conserva un Excel actualizado. Construir una copia
        # con el esquema anterior para comprobar una migración real.
        libro = modulo.load_workbook(self.excel)
        for nombre in ('Perfiles', 'Sesiones'):
            headers = modulo.SCHEMA[nombre][:-1]
            libro[nombre].delete_cols(len(headers) + 1)
            for tabla in libro[nombre].tables.values():
                tabla.ref = f'A1:{modulo.get_column_letter(len(headers))}{max(2, libro[nombre].max_row)}'
                tabla.tableColumns = [modulo.TableColumn(id=i, name=h) for i, h in enumerate(headers, 1)]
                if tabla.autoFilter is not None:
                    tabla.autoFilter.ref = tabla.ref
        libro.save(self.excel)
        libro.close()
        original = self.excel.read_bytes()
        libro = modulo.load_workbook(io.BytesIO(original))
        anteriores = {ws.title: list(ws.values) for ws in libro}
        libro.close()
        self.datos()
        libro = modulo.load_workbook(self.excel)
        try:
            for nombre, filas in anteriores.items():
                for i, fila in enumerate(filas, 1):
                    for j, valor in enumerate(fila, 1):
                        self.assertEqual(libro[nombre].cell(i, j).value, valor)
            for nombre in ('Perfiles', 'Sesiones'):
                self.assertEqual([c.value for c in libro[nombre][1]], modulo.SCHEMA[nombre])
                for tabla in libro[nombre].tables.values():
                    self.assertEqual([c.name for c in tabla.tableColumns], modulo.SCHEMA[nombre])
        finally:
            libro.close()
        copias = list((self.raiz / 'copias_excel').glob('*.xlsx'))
        self.assertTrue(any(p.read_bytes() == original for p in copias))

    def test_36_desplegable_y_borrador_conservan_datos(self):
        self.cuenta()
        e = self.ejercicio()
        self.ejercicio('Curl bíceps', 'Bíceps')
        self.post('/descanso', {'accion': 'configurar', 'minutos': '1', 'segundos_extra': '15'})
        prefijo = e['id'] + '_1'
        campos = {'dia': '0', 'serie': prefijo, 'peso_' + prefijo: '22',
            'reps_' + prefijo: '9', 'rutina': 'Pecho y bíceps',
            'fecha': date.today().isoformat(), 'duracion_min': '41', 'notas': 'Cambios de carga.'}
        r = self.post('/entrenar/serie', campos)
        parser = AnalizarHTML()
        parser.feed(r.get_data(as_text=True))
        ejercicios = [d for d in parser.detalles if 'exercise-accordion' in d.get('class', '')]
        self.assertEqual(len(ejercicios), 2)
        self.assertIn('open', next(d for d in ejercicios if d['id'] == 'ejercicio-' + e['id']))
        self.assertEqual(parser.campos['duracion_min'], '41')
        self.assertEqual(parser.campos['rutina'], 'Pecho y bíceps')
        self.assertIn('Cambios de carga.', r.get_data(as_text=True))
        self.assertNotIn('series planificadas', r.get_data(as_text=True))
        with self.cliente.session_transaction() as ses:
            self.assertEqual(ses['reloj']['duracion'], 75)

    def test_37_ocultar_y_mostrar_sin_reiniciar_reloj(self):
        self.cuenta()
        self.post('/descanso', {'accion': 'iniciar', 'segundos': '120'})
        with self.cliente.session_transaction() as ses:
            fin = ses['reloj']['fin']
        self.cliente.get('/progreso')
        html_reloj = self.cliente.get('/descanso/reloj').get_data(as_text=True)
        self.assertIn('aria-label="Ocultar contador"', html_reloj)
        self.assertIn('target="_top"', html_reloj)
        response = self.post('/descanso/ocultar')
        self.assertEqual(response.request.path, '/progreso')
        self.assertNotIn('reloj-global-frame', response.get_data(as_text=True))
        self.assertNotIn('class="con-reloj"', response.get_data(as_text=True))
        html = self.cliente.get('/entrenar?dia=0').get_data(as_text=True)
        self.assertIn('Mostrar contador de descanso', html)
        self.assertNotIn('reloj-global-frame', html)
        response = self.post('/descanso/mostrar')
        self.assertEqual(response.request.path, '/entrenar')
        self.assertIn('reloj-global-frame', response.get_data(as_text=True))
        with self.cliente.session_transaction() as ses:
            self.assertEqual(ses['reloj']['fin'], fin)

    def test_38_nueva_serie_recupera_contador_y_logout_limpia_oculto(self):
        self.cuenta()
        ejercicio = self.ejercicio()
        self.post('/descanso/ocultar')
        prefijo = ejercicio['id'] + '_1'
        response = self.post('/entrenar/serie', {'dia': '0', 'serie': prefijo,
            'peso_' + prefijo: '20', 'reps_' + prefijo: '10'})
        self.assertIn('reloj-global-frame', response.get_data(as_text=True))
        with self.cliente.session_transaction() as ses:
            self.assertNotIn('reloj_oculto', ses)

        self.post('/descanso/ocultar')
        self.post('/salir')
        with self.cliente.session_transaction() as ses:
            self.assertNotIn('reloj_oculto', ses)

    def test_39_boton_activa_lectura_y_vincula_cuenta_con_controlador(self):
        identifier = self.cuenta()
        with modulo.RFID_LOCK:
            modulo.RFID_STATE['disponible'] = False
        gpio = types.ModuleType('RPi.GPIO')
        gpio.setwarnings = lambda value: None
        gpio.cleanup = unittest.mock.Mock()
        rpi = types.ModuleType('RPi')
        rpi.GPIO = gpio
        lector = unittest.mock.Mock()
        lector.read_id_no_block.return_value = 1234567890
        libreria = types.ModuleType('mfrc522')
        libreria.SimpleMFRC522 = lambda: lector
        etapa = [0]
        def avanzar(_):
            if etapa[0] == 0:
                # Está preparado, pero aún NO consulta tarjetas hasta pulsar el botón.
                lector.read_id_no_block.assert_not_called()
                etapa[0] = 1
                response = self.post('/ajustes/rfid/iniciar')
                self.assertIn('Esperando lectura', response.get_data(as_text=True))
                self.assertEqual(modulo.RFID_STATE['motivo'], 'asociar')
                return
            raise KeyboardInterrupt()
        with patch.dict('sys.modules', {'RPi': rpi, 'RPi.GPIO': gpio, 'mfrc522': libreria}), \
                patch.object(modulo, 'ROOT', self.raiz), patch.object(modulo.time, 'sleep', avanzar), \
                patch('builtins.print'):
            with self.assertRaises(KeyboardInterrupt):
                modulo._hilo_lector_rfid()
            lectura = self.campos(self.cliente.get('/ajustes?escanear=1'))['lectura']
            self.post('/ajustes/rfid/asociar', {'lectura': lectura})
            perfil = next(p for p in self.datos()['Perfiles'] if p['id'] == identifier)
            self.assertEqual(perfil['rfid_uid'], '1234567890')
            self.assertTrue(modulo.check_password_hash(perfil['pin_hash'], '2468'))
            self.post('/salir')
            self.cliente.get('/ingresar?modo=rfid')
            with self.assertRaises(KeyboardInterrupt):
                modulo._hilo_lector_rfid()
            lectura = self.campos(self.cliente.get('/ingresar?modo=rfid'))['lectura']
            self.post('/entrar/rfid', {'lectura': lectura, 'pin': '2468'})
        with self.cliente.session_transaction() as ses:
            self.assertEqual(ses['perfil'], identifier)
        self.assertEqual(lector.read_id_no_block.call_count, 2)
        self.assertEqual(gpio.cleanup.call_count, 2)

    def test_40_boton_sin_lector_informa_error(self):
        self.cuenta()
        with modulo.RFID_LOCK:
            modulo.RFID_STATE['disponible'] = False
        response = self.post('/ajustes/rfid/iniciar')
        self.assertIn('El lector no está disponible', response.get_data(as_text=True))
        self.assertIsNone(modulo.RFID_STATE['solicitud'])


if __name__ == '__main__':
    unittest.main(verbosity=2)
