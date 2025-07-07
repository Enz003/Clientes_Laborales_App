from docxtpl import DocxTemplate, InlineImage
import uuid
from PIL import Image
from docx.shared import Mm,Cm
import requests
from io import BytesIO
from datetime import datetime
import qrcode
import os
import sys
import shutil
import time
import requests



def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        #PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def dolar_hoy():
    url = "https://v6.exchangerate-api.com/v6/cde2d5b8b6f71d1f8dd6790d/latest/USD"
    try:
        respuesta = requests.get(url)
        datos = respuesta.json()
        if datos["result"] == "success":
            valor = datos["conversion_rates"]["PYG"]
            return str(round(valor))
        else:
            return "7300"
    except Exception as e:
        print(f"Error: {e}")
        return "7300"

#----------------------------------------------------------------------------------------------------------------------
import re

UNIDADES = ['', 'UNO', 'DOS', 'TRES', 'CUATRO', 'CINCO', 'SEIS', 'SIETE', 'OCHO', 'NUEVE']
DECENAS = ['', 'DIEZ', 'VEINTE', 'TREINTA', 'CUARENTA', 'CINCUENTA', 'SESENTA', 'SETENTA', 'OCHENTA', 'NOVENTA']
DIECIS = ['DIEZ', 'ONCE', 'DOCE', 'TRECE', 'CATORCE', 'QUINCE', 'DIECISEIS', 'DIECISIETE', 'DIECIOCHO', 'DIECINUEVE']
VEINTIS = ['VEINTE', 'VEINTIUNO', 'VEINTIDOS', 'VEINTITRES', 'VEINTICUATRO', 'VEINTICINCO', 
           'VEINTISEIS', 'VEINTISIETE', 'VEINTIOCHO', 'VEINTINUEVE']
CIENTOS = ['', 'CIENTO', 'DOSCIENTOS', 'TRESCIENTOS', 'CUATROCIENTOS', 'QUINIENTOS', 
           'SEISCIENTOS', 'SETECIENTOS', 'OCHOCIENTOS', 'NOVECIENTOS']

def numero_a_letras(numero_str):
    # Limpiar el número (quitar puntos y convertir a entero)
    numero = int(re.sub(r'[^\d]', '', numero_str))
    
    if numero == 0:
        return 'CERO'
    
    letras = []
    if numero >= 1_000_000_000:
        miles_millones = numero // 1_000_000_000
        numero %= 1_000_000_000
        if miles_millones == 1:
            letras.append('MIL MILLONES')
        else:
            letras.append(convertir_grupo(miles_millones) + ' MIL MILLONES')
    
    if numero >= 1_000_000:
        millones = numero // 1_000_000
        numero %= 1_000_000
        if millones == 1:
            letras.append('UN MILLON')
        else:
            letras.append(convertir_grupo(millones) + ' MILLONES')
    
    if numero >= 1000:
        miles = numero // 1000
        numero %= 1000
        if miles == 1:
            letras.append('MIL')
        else:
            letras.append(convertir_grupo(miles) + ' MIL')
    
    if numero > 0:
        letras.append(convertir_grupo(numero))
    
    return ' '.join(letras)

def convertir_grupo(n):
    if n == 100:
        return 'CIEN'
    
    c = n // 100
    d = (n % 100) // 10
    u = n % 10
    grupo = []
    
    if c > 0:
        grupo.append(CIENTOS[c])
    
    if d == 1:
        grupo.append(DIECIS[u])
    elif d == 2 and u > 0:
        grupo.append(VEINTIS[u])
    else:
        if d > 0:
            grupo.append(DECENAS[d])
        if u > 0:
            if d > 0:
                grupo.append('Y')
            grupo.append(UNIDADES[u])
    
    return ' '.join(grupo)

#-----------------------------------------------------------------------------------------------------------------------

def formatear_fecha_conInput(dia,mes,anho):
    
    meses = [
        "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
    ]
    
    return f"{dia} de {meses[int(mes)-1]} de {anho}"

def obtener_fecha_formateada():
    # Diccionario de meses en español
    meses = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
        5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
        9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
    }
    
    # Obtener fecha actual
    hoy = datetime.now()
    dia = hoy.day
    mes_num = hoy.month
    año = hoy.year
    
    # Formatear la fecha
    fecha_formateada = f"{dia} días del mes de {meses[mes_num]} del año {año}"
    
    return fecha_formateada

def calcular_antiguedad(inicio_str, fin_str):
    # Convertimos las fechas en objetos datetime
    inicio = datetime.strptime(str(inicio_str), "%d/%m/%Y")
    fin = datetime.strptime(str(fin_str), "%d/%m/%Y")

    # Diferencia total en años y meses
    años = fin.year - inicio.year
    meses = fin.month - inicio.month
    dias = fin.day - inicio.day

    # Ajuste si los días del mes de fin son menores que los del mes de inicio
    if dias < 0:
        meses -= 1  # un mes menos si no completó el mes

    # Ajuste si los meses dan negativos
    if meses < 0:
        años -= 1
        meses += 12

    # Aplica la regla: si tiene más de 6 meses en el último año → se redondea a un año más
    if meses > 6 or (meses == 6 and dias > 0):
        años += 1

    return años


def limpiar_carpeta(ruta_carpeta, eliminar_subcarpetas=False):
    """
    Elimina todos los archivos de una carpeta local, con opción para eliminar subcarpetas.
    
    Args:
        ruta_carpeta (str): Ruta absoluta o relativa de la carpeta a limpiar
        eliminar_subcarpetas (bool): Si es True, también elimina subcarpetas y su contenido
    
    Returns:
        tuple: (archivos_eliminados, errores)
    """
    archivos_eliminados = 0
    errores = 0
    
    try:
        # Verificar si la carpeta existe
        ruta_carpeta = resource_path(ruta_carpeta)
        if not os.path.exists(ruta_carpeta):
            raise FileNotFoundError(f"La carpeta no existe: {ruta_carpeta}")
        
        # Recorrer todos los elementos en la carpeta
        for nombre in os.listdir(ruta_carpeta):
            ruta_completa = os.path.join(ruta_carpeta, nombre)
            
            try:
                if os.path.isfile(ruta_completa):
                    os.remove(ruta_completa)
                    archivos_eliminados += 1
                elif os.path.isdir(ruta_completa) and eliminar_subcarpetas:
                    shutil.rmtree(ruta_completa)
                    archivos_eliminados += 1  # Contamos la carpeta como un elemento eliminado
            except Exception as e:
                print(f"Error al eliminar {ruta_completa}: {str(e)}")
                errores += 1
        
        return (archivos_eliminados, errores)
    
    except Exception as e:
        print(f"Error general: {str(e)}")
        return (0, 1)

def generar_qr_inline(doc, enlace, ancho_mm=30):
    """
    Genera un código QR desde un enlace y lo devuelve como InlineImage para docxtpl.
    
    Parámetros:
        doc: instancia de DocxTemplate
        enlace: texto o URL a codificar en el QR
        ancho_mm: ancho del QR en milímetros (default: 30)
    
    Retorna:
        InlineImage listo para usarse en el contexto del template
    """
    # Crear imagen QR
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H)
    qr.add_data(enlace)
    qr.make(fit=True)
    
    img = qr.make_image(fill_color="black", back_color="white")

    # Guardar QR como imagen temporal
    nombre_archivo = f"Img/qr_{uuid.uuid4().hex}.png"
    nombre_archivo = resource_path(nombre_archivo)
    img.save(nombre_archivo)

    # Crear InlineImage
    return InlineImage(doc, nombre_archivo, width=Mm(ancho_mm))

# Extrae el ID de la URL de Google Drive (si está en formato "open?id=")
def convertir_url_google_drive(url):
    if "open?id=" in url:
        id_imagen = url.split("open?id=")[-1]
        return f"https://drive.google.com/uc?export=download&id={id_imagen}"
    return url  # Ya es directa


def generar_imagen_inline(doc, url, ancho_cm=14):


    nombre_archivo = f"Img/imagen_{uuid.uuid4().hex}.jpg"
    nombre_archivo = resource_path(nombre_archivo)

    url_valida = convertir_url_google_drive(url)
    response = requests.get(url_valida)

    content_type = response.headers.get('Content-Type', '')
    if 'image' not in content_type:
        raise Exception(f"La URL no contiene una imagen válida. Content-Type: {content_type}")

    try:
        img = Image.open(BytesIO(response.content))
        img.verify()
        img = Image.open(BytesIO(response.content))
    except Exception as e:
        raise Exception("El contenido descargado no es una imagen válida") from e

    os.makedirs(os.path.dirname(nombre_archivo), exist_ok=True)

    # 🔧 Solución al error RGBA
    if img.mode == 'RGBA':
        img = img.convert('RGB')

    img.save(nombre_archivo)

    width_px, height_px = img.size
    dpi = img.info.get("dpi", (96, 96))[0]
    ancho_in = ancho_cm / 2.54
    escala = (ancho_in * dpi) / width_px
    alto_cm = (height_px * escala) / dpi * 2.54

    return InlineImage(doc, nombre_archivo, width=Cm(ancho_cm), height=Cm(alto_cm))




def  FORM_DATOS_NUEVOS_PARA_TRABAJADOR(cliente):
    contexto = {
        'nombre_completo': cliente['Nombres y Apellidos completos como esta en tu Cedula.'],
        'estado_civil': cliente['Estado Civil como esta en tu cedula'],
        'nacionalidad': cliente['Nacionalidad'],
        'ci': cliente['Numero de Cedula'],
        'ciudad': cliente['Ciudad'],
        'barrio': cliente['Barrio'],
        'direccion_calle': cliente['Direccion Particular, Calles, Numero de casa'],
        'telefono': cliente['Telefono de contacto personal'],
        'empresa_que_trabajo': cliente['Empresa en la que trabajo <Razon Social>'],
        'direccion_empresa': cliente['Direccion de la Empresa'],
        'ruc_empresa':cliente['Ruc de la empresa'],
        'fecha_ingreso': cliente['Fecha de ingreso'],
        'fecha_despido': cliente['Fecha de Despido'],
        'jornada_laboral': cliente['JORNADA LABORAL. Como es o era tu Jornada Laboral? Lunes a Viernes, Lunes a Lunes?'],
        'horario_laboral': cliente['HORARIO DE TRABAJO. Como era tu Horario que Cumplias? Ej 8.00 a 18.00'],
        'salario': cliente['CUANTO ERA SALARIO. Mensual, semanal, diario?'],
        'ips': cliente['IPS'],
        'bonificacion_familiar':cliente['Bonficacion familiar por Cuantos hijos si la respuesta fue SI'],
        'tarea_realizada': cliente['Describe las tareas o funciones que desempenabas en el lugar de trabajo.'],
        'motivo_despido': cliente['MOTIVO DE DESPIDO. Cuentanos como se dio la situacion.'],
        'quien_lo_despidio': cliente['Quien te Comunico de tu despido?. Nombre Apellido, cargo en la empresa.'],
        'medio_del_despido': cliente['DESPIDO COMUNICACION. Como te comunicaron tu despido. Verbal, por escrito con nota, por llamada telefonica, por mensaje de texto?'],
        'salarios_pendientes': cliente['SALARIOS PENDIENTES. Te deben Salarios, Cuanto de cuantos dias o meses?'],
        'medio_pago_salario': cliente['PAGOS DE SALARIOS. Como recibias los pagos de salario o jornales?. Efectivo, Trasferencia, giros? via que banco?'],
        'vacaciones_pendientes': cliente['VACACIONES. Salias o tenias vacaciones? Te deben vacaciones?'],
        'aguinaldo_pendiente': cliente['AGUINALDO. Recibias Aguinaldo?. Te pagaban o te debe?'],
        'contaba_contrato': cliente['CONTRATO DE TRABAJO. Tenias contrato de Trabajo Firmado'],
        'ofrecio_liquidacion': cliente['LIQUIDACION.Te presentaron tu liquidacion de salarios y haberes al momento del despido?. Adjuntar Foto.'],
        'firmo_documento_en_blanco': cliente['Firmaste en algun momento algun Documento en blanco o pagare?'],
        'observaciones': cliente['Alguna informacion adicional que deseas agregar?'],
        'entrevistador':cliente['Entrevista realizada por']
    }

    fecha_actual = datetime.now().strftime('%d/%m/%Y')
    ruta_abrir = resource_path('Plantilla/Plantillla_Formulario_Datos.docx')
    doc = DocxTemplate(ruta_abrir)
    time.sleep(5)

    imagen_actor = generar_imagen_inline(doc, cliente['Adjunta Imagen de la Ubicacion de Google Maps de casa Trabajador.'])
    link_ubicacion_actor = cliente['Ubicacion de tu casa. Copia el link de la ubicacion de google maps']
    qr_actor = generar_qr_inline(doc,link_ubicacion_actor)
    imagen_demandado = generar_imagen_inline(doc,cliente['Adjunta Imagen de la Ubicacion de Google Maps de la empresa'])
    link_ubicacion_demandado = cliente['Ubicacion de la empresa']
    qr_demandado = generar_qr_inline(doc,link_ubicacion_demandado)
    imagenesYfecha = {'fecha_hoy': fecha_actual,
                    'imagen_actor':imagen_actor,
                    'link_ubicacion_actor': link_ubicacion_actor,
                    'imagen_demandado':imagen_demandado,
                    'link_ubicacion_demandado':link_ubicacion_demandado,
                    'qr_actor':qr_actor,
                    'qr_demandado':qr_demandado}
    contexto.update(imagenesYfecha)


    doc.render(contexto)

    carpeta_limpiar = resource_path('Img')
    limpiar_carpeta(carpeta_limpiar)

    ruta_guardar = resource_path(f'Generado/Formulario_Datos_{contexto['ci']}.docx') 
    doc.save(ruta_guardar)

def  Carta_Poder(cliente):
    contexto = {
        'fecha': obtener_fecha_formateada(),
        'nombre_completo': cliente['Nombres y Apellidos completos como esta en tu Cedula.'].upper(),
        'estado_civil': cliente['Estado Civil como esta en tu cedula'].lower(),
        'nacionalidad': cliente['Nacionalidad'].lower(),
        'ci': cliente['Numero de Cedula'],
        'ciudad': cliente['Ciudad'].capitalize(),
        'direccion_calle': cliente['Direccion Particular, Calles, Numero de casa'],
        'empresa_que_trabajo': cliente['Empresa en la que trabajo <Razon Social>'].upper(),
        'direccion_empresa': cliente['Direccion de la Empresa'],
        'ciudad_empresa': cliente['Ciudad de la empresa'].capitalize(),
        'ruc_empresa':cliente['Ruc de la empresa'] 
    }

    if cliente['Sexo']=='Femenino':
        contexto['estado_civil']=contexto['estado_civil'][:-1]+'a'

    ruta_abrir = resource_path('Plantilla/Carta_de_Poder.docx')
    doc = DocxTemplate(ruta_abrir)



    doc.render(contexto)

    ruta_guardar = resource_path(f'Generado/Carta_de_Poder_{contexto['ci']}.docx')
    doc.save(ruta_guardar)

def Carta_Compromiso(cliente):
   
    contexto = {
        'nombre_completo': cliente['Nombres y Apellidos completos como esta en tu Cedula.'].upper(),
        'fecha': obtener_fecha_formateada(),
        'ci': cliente['Numero de Cedula'],
        'ciudad': cliente['Ciudad'].capitalize(),
        'direccion_calle': cliente['Direccion Particular, Calles, Numero de casa'],
        'empresa_que_trabajo': cliente['Empresa en la que trabajo <Razon Social>'].upper(),
        'ruc_empresa':cliente['Ruc de la empresa']
    }

    ruta_abrir = resource_path('Plantilla/Carta_Compromiso.docx')
    doc = DocxTemplate(ruta_abrir)

    doc.render(contexto)

    ruta_guardar = resource_path(f'Generado/Carta_Compromiso_{contexto['ci']}.docx') 
    doc.save(ruta_guardar)
    
def Desistimiento_de_renuncia(cliente):
    contexto = {
        'nombre_completo': cliente['Nombres y Apellidos completos como esta en tu Cedula.'].upper(),
        'ci': cliente['Numero de Cedula']
    }

    ruta_abrir = resource_path('Plantilla/Desistimiento_de_Renuncia.docx')
    doc = DocxTemplate(ruta_abrir)

    doc.render(contexto)

    ruta_cerrar = resource_path(f'Generado/Desistimiento_de_renuncia_{contexto['ci']}.docx')
    doc.save(ruta_cerrar)

def Nota_de_Renuncia(cliente):
    contexto = {
        'nombre_completo': cliente['Nombres y Apellidos completos como esta en tu Cedula.'].upper(),
        'ci': cliente['Numero de Cedula'],
        'empresa_que_trabajo': cliente['Empresa en la que trabajo <Razon Social>'].upper(),
        'ruc_empresa':cliente['Ruc de la empresa']
    }

    ruta_abrir = resource_path('Plantilla/Nota_de_Renuncia.docx')
    doc = DocxTemplate(ruta_abrir)

    doc.render(contexto)

    ruta_cerrar = resource_path(f'Generado/Nota_de_Renuncia_{contexto['ci']}.docx')
    doc.save(ruta_cerrar)
    
def documento_demanda(cliente, datos_indemizacion = None):
    ruta_abrir = resource_path('Plantilla/documento_demanda.docx')
    doc = DocxTemplate(ruta_abrir)

    contexto = {
        'nombre_completo': cliente['Nombres y Apellidos completos como esta en tu Cedula.'].upper(),
        'estado_civil': cliente['Estado Civil como esta en tu cedula'].lower(),
        'nacionalidad': cliente['Nacionalidad'].lower(),
        'ci': cliente['Numero de Cedula'],
        'ciudad': cliente['Ciudad'].lower(),
        'barrio': cliente['Barrio'].lower(),
        'direccion_calle': cliente['Direccion Particular, Calles, Numero de casa'],
        'telefono': cliente['Telefono de contacto personal'],
        'empresa_que_trabajo': cliente['Empresa en la que trabajo <Razon Social>'].upper(),
        'direccion_empresa': cliente['Direccion de la Empresa'],
        'ciudad_empresa': cliente['Ciudad de la empresa'],
        'ruc_empresa':cliente['Ruc de la empresa'],
        'fecha_ingreso': cliente['Fecha de ingreso'],
        'jornada_laboral': cliente['JORNADA LABORAL. Como es o era tu Jornada Laboral? Lunes a Viernes, Lunes a Lunes?'],
        'horario_laboral': cliente['HORARIO DE TRABAJO. Como era tu Horario que Cumplias? Ej 8.00 a 18.00'],
        'salario': cliente['CUANTO ERA SALARIO. Mensual, semanal, diario?'],
        'ips': cliente['IPS'],
        'bonificacion_familiar':cliente['Bonficacion familiar por Cuantos hijos si la respuesta fue SI'],
        'tarea_realizada': cliente['Describe las tareas o funciones que desempenabas en el lugar de trabajo.'].lower(),
        'motivo_despido': cliente['MOTIVO DE DESPIDO. Cuentanos como se dio la situacion.'],
        'quien_lo_despidio': cliente['Quien te Comunico de tu despido?. Nombre Apellido, cargo en la empresa.'],
        'medio_del_despido': cliente['DESPIDO COMUNICACION. Como te comunicaron tu despido. Verbal, por escrito con nota, por llamada telefonica, por mensaje de texto?'],
        'salarios_pendientes': cliente['SALARIOS PENDIENTES. Te deben Salarios, Cuanto de cuantos dias o meses?'],
        'medio_pago_salario': cliente['PAGOS DE SALARIOS. Como recibias los pagos de salario o jornales?. Efectivo, Trasferencia, giros? via que banco?'],
        'vacaciones_pendientes': cliente['VACACIONES. Salias o tenias vacaciones? Te deben vacaciones?'],
        'aguinaldo_pendiente': cliente['AGUINALDO. Recibias Aguinaldo?. Te pagaban o te debe?'],
        'contaba_contrato': cliente['CONTRATO DE TRABAJO. Tenias contrato de Trabajo Firmado'],
        'ofrecio_liquidacion': cliente['LIQUIDACION.Te presentaron tu liquidacion de salarios y haberes al momento del despido?. Adjuntar Foto.'],
        'firmo_documento_en_blanco': cliente['Firmaste en algun momento algun Documento en blanco o pagare?'],
        'observaciones': cliente['Alguna informacion adicional que deseas agregar?'],
        'entrevistador':cliente['Entrevista realizada por']
    }

    

    fecha_ingreso = str(cliente['Fecha de ingreso']).split('/')
    dia_ingreso = fecha_ingreso[0]
    mes_ingreso = fecha_ingreso[1]
    anho_ingreso = fecha_ingreso[2]
    fecha_despido = str(cliente['Fecha de Despido']).split('/')
    dia_despido = fecha_despido[0]
    mes_despido = fecha_despido[1]
    anho_despido = fecha_despido[2]
    antiguedad = calcular_antiguedad(cliente['Fecha de ingreso'],cliente['Fecha de Despido'])

    if cliente['Sexo']=='Femenino':
        contexto['estado_civil']=contexto['estado_civil'][:-1]+'a'

    contexto.update({   
                     'antiguedad': antiguedad,
                     'fecha_ingreso':formatear_fecha_conInput(dia_ingreso,mes_ingreso,anho_ingreso),
                     'anho_ingreso':anho_ingreso,
                     'fecha_despido': formatear_fecha_conInput(dia_despido,mes_despido,anho_despido),
                     'anho_despido': anho_despido,
                     'dolar_hoy': dolar_hoy()})
    if datos_indemizacion:
        datos_cliente_indemnizacion = calcular_agregar_indemnizacion(datos_indemizacion)
        contexto.update(datos_cliente_indemnizacion)

    doc.render(contexto)

    ruta_cerrar = resource_path(f'Generado/Demanda {cliente['Nombres y Apellidos completos como esta en tu Cedula.'].split()[0]} contra {cliente['Empresa en la que trabajo <Razon Social>']}.docx')
    doc.save(ruta_cerrar)

from datetime import datetime, date
from decimal import Decimal, ROUND_HALF_UP
import calendar

class LiquidacionDespido:
    def __init__(self):
        self.salario_minimo = 2681084  # Salario mínimo 2024 en guaraníes
    
    def calcular_antiguedad_años(self, fecha_ingreso, fecha_despido):
        """Calcula la antigüedad en años completos"""
        if isinstance(fecha_ingreso, str):
            fecha_ingreso = datetime.strptime(fecha_ingreso, "%d/%m/%Y").date()
        if isinstance(fecha_despido, str):
            fecha_despido = datetime.strptime(fecha_despido, "%d/%m/%Y").date()
        
        años = fecha_despido.year - fecha_ingreso.year
        if fecha_despido.month < fecha_ingreso.month or \
           (fecha_despido.month == fecha_ingreso.month and fecha_despido.day < fecha_ingreso.day):
            años -= 1
        
        return años
    
    def calcular_dias_trabajados_ultimo_año(self, fecha_ingreso, fecha_despido):
        """Calcula los días trabajados en el último año incompleto"""
        if isinstance(fecha_ingreso, str):
            fecha_ingreso = datetime.strptime(fecha_ingreso, "%d/%m/%Y").date()
        if isinstance(fecha_despido, str):
            fecha_despido = datetime.strptime(fecha_despido, "%d/%m/%Y").date()
        
        años_completos = self.calcular_antiguedad_años(fecha_ingreso, fecha_despido)
        
        # Fecha de inicio del último período
        fecha_ultimo_año = date(fecha_ingreso.year + años_completos, 
                               fecha_ingreso.month, 
                               fecha_ingreso.day)
        
        # Si la fecha calculada es posterior al despido, usar la fecha de ingreso
        if fecha_ultimo_año > fecha_despido:
            fecha_ultimo_año = fecha_ingreso
        
        # Calcular días del último período
        dias = (fecha_despido - fecha_ultimo_año).days
        return dias
    
    def calcular_preaviso(self,dias_preaviso_cumplido, salario_mensual, años_antiguedad, tipo_despido="Injustificado"):
        """
        Calcula indemnización por preaviso
        Art. 91 - Código del Trabajo
        """
        if tipo_despido == "Justificado":
            return 0
        
        if años_antiguedad < 1:
            dias_preaviso = 30
        elif años_antiguedad < 5:
            dias_preaviso = 45
        elif años_antiguedad < 10:
            dias_preaviso = 60
        else:
            dias_preaviso = 90
        
        monto = (salario_mensual * (dias_preaviso - dias_preaviso_cumplido)) / 30
        if monto<0:
          return 0
        return monto
    
    def calcular_indemnizacion_antiguedad(self, salario_mensual, años_antiguedad, dias_ultimo_año):
        """
        Calcula indemnización por antigüedad
        Art. 92 - Código del Trabajo
        15 días de salario por cada año completo
        Proporcional por fracción de año
        """
        # 15 días por año completo
        indemnizacion_años = (salario_mensual * 15 * años_antiguedad) / 30
        
        # Proporcional por días del último año
        indemnizacion_dias = (salario_mensual * 15 * dias_ultimo_año) / (30 * 365)
        
        return indemnizacion_años + indemnizacion_dias
    
    def calcular_aguinaldo_proporcional(self, salario_mensual, fecha_ingreso, fecha_despido):
        """
        Calcula aguinaldo proporcional (1/12 por mes trabajado en el año)
        """
        if isinstance(fecha_despido, str):
            fecha_despido = datetime.strptime(fecha_despido, "%d/%m/%Y").date()
        
        año_despido = fecha_despido.year
        inicio_año = date(año_despido, 1, 1)
        
        # Si ingresó en el mismo año, usar fecha de ingreso
        if isinstance(fecha_ingreso, str):
            fecha_ingreso = datetime.strptime(fecha_ingreso, "%d/%m/%Y").date()
        
        if fecha_ingreso.year == año_despido:
            inicio_calculo = fecha_ingreso
        else:
            inicio_calculo = inicio_año
        
        # Calcular meses completos trabajados en el año
        meses_trabajados = 0
        fecha_actual = inicio_calculo
        
        while fecha_actual.replace(day=1) <= fecha_despido.replace(day=1):
            if fecha_actual.month == fecha_despido.month:
                # Último mes - verificar si trabajó más de 15 días
                dias_mes = fecha_despido.day
                if dias_mes >= 15:
                    meses_trabajados += 1
            else:
                meses_trabajados += 1
            
            # Pasar al siguiente mes
            if fecha_actual.month == 12:
                fecha_actual = fecha_actual.replace(year=fecha_actual.year + 1, month=1)
            else:
                fecha_actual = fecha_actual.replace(month=fecha_actual.month + 1)
        
        return (salario_mensual * meses_trabajados) / 12
    
    def calcular_vacaciones_causadas(self, salario_mensual, dias_vacaciones_causadas):
        """
        Calcula vacaciones causadas (pendientes del año anterior)
        """
        if dias_vacaciones_causadas <= 0:
            return 0
        
        return (salario_mensual * dias_vacaciones_causadas) / 30
    
    def calcular_vacaciones_proporcionales(self, salario_mensual, dias_vacaciones_proporcionales):
        """
        Calcula vacaciones proporcionales (otorgadas este año aún no gozadas)
        """
        if dias_vacaciones_proporcionales <= 0:
            return 0
        
        return (salario_mensual * dias_vacaciones_proporcionales) / 30
    
    def calcular_vacaciones_acumuladas(self, salario_mensual, dias_vacaciones_acumuladas):
        """
        Calcula vacaciones acumuladas (por tomar acumuladas anteriormente)
        """
        if dias_vacaciones_acumuladas <= 0:
            return 0
        
        return (salario_mensual * dias_vacaciones_acumuladas) / 30
    
    def calcular_dias_sueldo_pendientes(self, salario_mensual, dias_sueldo_pendientes):
        """
        Calcula el monto por días de sueldo aún no gozados/pagados
        """
        if dias_sueldo_pendientes <= 0:
            return 0
        
        return (salario_mensual * dias_sueldo_pendientes) / 30
    
    def calcular_descuento_ips(self, monto_total, tiene_seguro_ips):
        """
        Calcula el descuento del 9% de IPS si corresponde
        """
        if tiene_seguro_ips.lower() in ['si', 'sí', 's', 'yes', 'y']:
            return monto_total * 0.09
        return 0
    
    def calcular_liquidacion_completa(self,dias_preaviso_cumplido, salario_mensual, fecha_ingreso, fecha_despido, 
                                    tipo_despido="sin_causa", 
                                    dias_vacaciones_causadas=0,
                                    dias_vacaciones_proporcionales=0,
                                    dias_vacaciones_acumuladas=0,
                                    dias_sueldo_pendientes=0,
                                    tiene_seguro_ips="no"):
        """
        Calcula la liquidación completa por despido
        
        Parámetros:
        - dias_preaviso_cumplido: Días de preaviso cumplidos por el empleador
        - salario_mensual: Salario mensual del trabajador
        - fecha_ingreso: Fecha de ingreso (formato "YYYY/MM/DD")
        - fecha_despido: Fecha de despido (formato "YYYY/MM/DD")
        - tipo_despido: "sin_causa" o "con_causa_justa"
        - dias_vacaciones_causadas: Días de vacaciones pendientes del año anterior
        - dias_vacaciones_proporcionales: Días de vacaciones otorgadas este año aún no gozadas
        - dias_vacaciones_acumuladas: Días de vacaciones acumuladas de años anteriores
        - dias_sueldo_pendientes: Días de sueldo aún no gozados/pagados
        - tiene_seguro_ips: "Si" o "No" para descuento del 9% de IPS
        """
        años_antiguedad = self.calcular_antiguedad_años(fecha_ingreso, fecha_despido)
        dias_ultimo_año = self.calcular_dias_trabajados_ultimo_año(fecha_ingreso, fecha_despido)
        
        # Cálculos individuales
        preaviso = self.calcular_preaviso(dias_preaviso_cumplido,salario_mensual, años_antiguedad, tipo_despido)
        indemnizacion = self.calcular_indemnizacion_antiguedad(salario_mensual, años_antiguedad, dias_ultimo_año)
        aguinaldo = self.calcular_aguinaldo_proporcional(salario_mensual, fecha_ingreso, fecha_despido)
        
        # Cálculos de vacaciones
        vacaciones_causadas = self.calcular_vacaciones_causadas(salario_mensual, dias_vacaciones_causadas)
        vacaciones_proporcionales = self.calcular_vacaciones_proporcionales(salario_mensual, dias_vacaciones_proporcionales)
        vacaciones_acumuladas = self.calcular_vacaciones_acumuladas(salario_mensual, dias_vacaciones_acumuladas)
        
        # Días de sueldo pendientes
        sueldo_pendiente = self.calcular_dias_sueldo_pendientes(salario_mensual, dias_sueldo_pendientes)
        
        # Subtotal antes de descuentos
        total_vacaciones = vacaciones_causadas + vacaciones_proporcionales + vacaciones_acumuladas
        subtotal = preaviso + indemnizacion + aguinaldo + total_vacaciones + sueldo_pendiente
        
        # Descuento IPS
        descuento_ips = self.calcular_descuento_ips(subtotal, tiene_seguro_ips)
        total_final = subtotal - descuento_ips
        
        return {
            'anhos_antiguedad': años_antiguedad,
            'dias_ultimo_anho': dias_ultimo_año,
            'conceptos': {
                'preaviso': round(preaviso),
                'indemnizacion_antiguedad': round(indemnizacion),
                'aguinaldo_proporcional': round(aguinaldo),
                'vacaciones_causadas': round(vacaciones_causadas),
                'vacaciones_proporcionales': round(vacaciones_proporcionales),
                'vacaciones_acumuladas': round(vacaciones_acumuladas),
                'total_vacaciones': round(total_vacaciones),
                'sueldo_pendiente': round(sueldo_pendiente),
                'subtotal': round(subtotal),
                'descuento_ips': round(descuento_ips),
                'total_final': round(total_final)
            },
            'total_liquidacion': round(total_final),
            'detalles': {
                'salario_mensual': salario_mensual,
                'fecha_ingreso': fecha_ingreso,
                'fecha_despido': fecha_despido,
                'tipo_despido': tipo_despido,
                'dias_vacaciones_causadas': dias_vacaciones_causadas,
                'dias_vacaciones_proporcionales': dias_vacaciones_proporcionales,
                'dias_vacaciones_acumuladas': dias_vacaciones_acumuladas,
                'total_dias_vacaciones': dias_vacaciones_causadas + dias_vacaciones_proporcionales + dias_vacaciones_acumuladas,
                'dias_sueldo_pendientes': dias_sueldo_pendientes,
                'tiene_seguro_ips': tiene_seguro_ips.lower() in ['si', 'sí', 's', 'yes', 'y']
            }
        }
def formatear_con_separadorDeMiles(numero):
    return "{:,}".format(numero).replace(",", ".")

def calcular_agregar_indemnizacion(cliente):
    liquidacion = LiquidacionDespido()
    # Datos del trabajador
    salario_mensual = int(cliente['Salario promedio mensual (Coloque solo el valor numérico del salario en guaraníes)'])  
    fecha_ingreso = cliente['Fecha de ingreso']
    fecha_despido = cliente['Fecha de Despido']
    tipo_despido = cliente['El despido fue:']  
    dias_preaviso_cumplido = int(cliente['¿Cuántos días de preaviso recibió? (Inserte solo el numero de días de preaviso)'])
    
    # Días de vacaciones por concepto
    dias_vacaciones_causadas = int(cliente['¿Le quedan vacaciones pendientes por tomar del año anterior? (Inserte solo el numero de días de vacaciones por tomar)'])
    dias_vacaciones_proporcionales = int(cliente['¿Fueron otorgados días de vacaciones aun no gozados correspondientes al año en curso? (Inserte solo el numero de días de vacaciones otorgadas)'])
    dias_vacaciones_acumuladas = int(cliente['¿Le quedan vacaciones por tomar que fueron acumuladas conforme al Art. 224 del Código Laboral? (Inserte solo el numero de días de vacaciones acumulados)'])
    
    # Días de sueldo pendientes y seguro IPS
    dias_sueldo_pendientes = int(cliente['Días trabajados del mes que aun no fueron abonados (Inserte un valor numérico)'])
    tiene_seguro_ips = cliente['IPS']

    # Calcular liquidación
    resultado = liquidacion.calcular_liquidacion_completa(
        dias_preaviso_cumplido,
        salario_mensual, 
        fecha_ingreso, 
        fecha_despido, 
        tipo_despido,
        dias_vacaciones_causadas,
        dias_vacaciones_proporcionales,
        dias_vacaciones_acumuladas,
        dias_sueldo_pendientes,
        tiene_seguro_ips
    )
    montos_liquidacion = resultado['conceptos']
    contexto = {
        'indemnizacion_antiguedad': str(formatear_con_separadorDeMiles(montos_liquidacion['indemnizacion_antiguedad'])),
        'preaviso_calculado': str(formatear_con_separadorDeMiles(montos_liquidacion['preaviso'])),
        'aguinaldo_proporcional': str(formatear_con_separadorDeMiles(montos_liquidacion['aguinaldo_proporcional'])),
        'vacaciones_causadas': str(formatear_con_separadorDeMiles(montos_liquidacion['vacaciones_causadas'])),
        'vacaciones_proporcionales': str(formatear_con_separadorDeMiles(montos_liquidacion['vacaciones_proporcionales'])),
        'vacaciones_acumuladas': str(formatear_con_separadorDeMiles(montos_liquidacion['vacaciones_acumuladas'])),
        'sueldo_pendiente': str(formatear_con_separadorDeMiles(montos_liquidacion['sueldo_pendiente'])),
        'descuento_ips': str(formatear_con_separadorDeMiles(montos_liquidacion['descuento_ips'])),
        'total_liquidacion': str(formatear_con_separadorDeMiles(montos_liquidacion['total_final'])),
        'total_liquidacion_str': numero_a_letras(str(formatear_con_separadorDeMiles(montos_liquidacion['total_final']))),
        'salario_formateado':str(formatear_con_separadorDeMiles(resultado['detalles']['salario_mensual']))
    }
    return contexto
    