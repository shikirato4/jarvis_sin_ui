import pyttsx3
import speech_recognition as sr
import os
import time
import pyautogui
import webbrowser
import urllib.parse
import datetime
import google.generativeai as genai
from openai import OpenAI
import cv2
import psutil
import ctypes
import re
import requests
import random
import threading
import subprocess
import sys
import tempfile
import json
import socket
import shutil
from pathlib import Path
from difflib import SequenceMatcher

# Intentamos importar librerías adicionales. Si no están, avisamos, pero no crasheamos el programa.
try:
    import tinytuya
except ImportError:
    tinytuya = None

try:
    import wikipedia
    wikipedia.set_lang("es")
except ImportError:
    wikipedia = None

try:
    import pyperclip
except ImportError:
    pyperclip = None

# ==========================================
# 0. Configuración del "Cerebro" (APIs de IA)
# ==========================================
# NO pego claves privadas aquí. Deja las tuyas en estas variables.
API_KEY_GEMINI = "PEGA_AQUI_TU_API_KEY_GEMINI"
genai.configure(api_key=API_KEY_GEMINI)

modelo_ia = genai.GenerativeModel(
    model_name="models/gemini-2.5-flash",
    system_instruction=(
        "Eres Jarvis, el asistente virtual hiper avanzado del señor Styx. "
        "Tienes permitido dar respuestas detalladas, largas y profundas. "
        "Eres un programador experto. Si se te pide escribir código, escríbelo SIEMPRE "
        "dentro de bloques de código markdown (```). "
        "También ayudas a automatizar tareas del sistema operativo Windows, explicas con claridad "
        "y respondes de forma natural, elegante y útil. "
        "Cuando el usuario pida una acción del sistema, puedes responder en formato JSON si eso permite ejecutar la orden."
    )
)
memoria_jarvis = modelo_ia.start_chat(history=[])

API_KEY_OPENAI = "PEGA_AQUI_TU_API_KEY_OPENAI"
cliente_openai = OpenAI(api_key=API_KEY_OPENAI)

historial_respaldo = [
    {
        "role": "system",
        "content": (
            "Eres Jarvis, el asistente virtual del señor Styx. "
            "Da respuestas detalladas y envuelve el código siempre en bloques markdown. "
            "Si es útil, puedes representar intenciones como JSON para automatización."
        )
    }
]

PROMPT_RUTEO_JSON = """
Convierte la orden del usuario a JSON puro.
No expliques nada, solo devuelve JSON válido.
Formato:
{
  "modo": "accion" | "chat",
  "accion": "nombre_accion" | null,
  "parametros": {},
  "respuesta_corta": "texto breve opcional"
}

Acciones permitidas:
abrir_app, abrir_url, buscar_google, buscar_youtube, reproducir_youtube,
escribir_texto, pegar_texto, crear_carpeta, crear_archivo_txt,
tomar_captura, tomar_foto, reconocimiento_facial,
reporte_hardware, listar_procesos, cerrar_proceso,
bloquear_computadora, apagar_pc, reiniciar_pc, suspender_pc,
vaciar_papelera, minimizar_todo, cerrar_ventana, maximizar_ventana,
restaurar_ventana, cambiar_ventana,
subir_volumen, bajar_volumen, silenciar_volumen,
pausar_media, siguiente_media, anterior_media,
abrir_descargas, abrir_documentos, abrir_escritorio,
buscar_archivo, buscar_archivos_fragmento,
leer_portapapeles, resumir_portapapeles, traducir_portapapeles_ingles,
activar_dictado, desactivar_dictado,
foco_encender, foco_apagar, foco_color, foco_brillo,
mover_cursor, arrastrar_cursor, click, doble_click, click_derecho,
scroll_arriba, scroll_abajo, presionar_tecla,
rutina_estudio, rutina_gaming, rutina_noche, rutina_programacion,
abrir_entorno_programacion,
chat

Usa parametros como:
- nombre
- texto
- consulta
- color
- brillo
- x
- y
- tecla
"""

# ==========================================
# 1. Configuración general
# ==========================================
APP_NAME = "J.A.R.V.I.S."
NOMBRE_ACTIVACION = "jarvis"
USUARIO_PC = r"C:\Users\GAMER"
ESCRITORIO = r"C:\Users\GAMER\Desktop"
DOCUMENTOS = r"C:\Users\GAMER\Documents"
DESCARGAS = r"C:\Users\GAMER\Downloads"
CARPETA_CAPTURAS = os.path.join(os.path.expanduser("~"), "Desktop", "Capturas_Jarvis")
CARPETA_FOTOS = os.path.join(os.path.expanduser("~"), "Desktop", "Fotos_Jarvis")
CARPETA_CODIGO = os.path.join(os.path.expanduser("~"), "Desktop", "Proyectos_Jarvis")
CARPETA_NOTAS = os.path.join(os.path.expanduser("~"), "Desktop", "Notas_Jarvis")
ARCHIVO_MEMORIA = os.path.join(USUARIO_PC, "jarvis_memoria_persistente.json")
ARCHIVO_MACROS = os.path.join(USUARIO_PC, "jarvis_macros.json")
ARCHIVO_RUTINAS = os.path.join(USUARIO_PC, "jarvis_rutinas.json")

for carpeta in [CARPETA_CAPTURAS, CARPETA_FOTOS, CARPETA_CODIGO, CARPETA_NOTAS]:
    os.makedirs(carpeta, exist_ok=True)

# Seguridad para PyAutoGUI
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.2

# ==========================================
# 2. Configuración de la Voz
# ==========================================
engine = pyttsx3.init()
engine.setProperty("rate", 165)

voices = engine.getProperty("voices")
if voices:
    engine.setProperty("voice", voices[0].id)

def hablar(texto):
    texto = str(texto).strip()

    if not texto:
        print(f"{APP_NAME}: [respuesta vacía]")
        return

    print(f"{APP_NAME}: {texto}")
    try:
        engine.stop()
        engine.say(texto)
        engine.runAndWait()
    except Exception as e:
        print(f"Error al hablar: {e}")

# ==========================================
# 3. Memoria persistente y utilidades
# ==========================================
MODO_DICTADO_CONTINUO = False
MODO_SUSPENDIDO = False
TEMPORIZADORES = []

ALIASES_TEXTO = {
    "habré": "abre",
    "hábre": "abre",
    "abré": "abre",
    "has clic": "haz clic",
    "da click": "da clic",
    "opera ge equis": "opera gx",
    "opera equis": "opera gx",
    "brabe": "brave",
    "fayarfocs": "firefox",
    "esteren": "steren",
    "jarbis": "jarvis",
    "yarvis": "jarvis",
    "show me": "shome"
}

ALIASES_APPS = {
    "opera": "opera gx",
    "opera gx": "opera gx",
    "brave": "brave",
    "firefox": "firefox",
    "calc": "calculadora",
    "notepad": "bloc de notas",
    "explorador de archivos": "explorador",
    "consola": "cmd"
}

COLORES_FOCO = {
    "rojo": (255, 0, 0),
    "verde": (0, 255, 0),
    "azul": (0, 0, 255),
    "blanco": (255, 255, 255),
    "amarillo": (255, 255, 0),
    "morado": (128, 0, 128),
    "naranja": (255, 165, 0),
    "rosa": (255, 192, 203),
    "cian": (0, 255, 255),
    "turquesa": (64, 224, 208),
    "lila": (200, 162, 200)
}


def normalizar_texto(texto):
    texto = (texto or "").lower().strip()
    for malo, bueno in ALIASES_TEXTO.items():
        texto = texto.replace(malo, bueno)
    texto = re.sub(r"\s+", " ", texto)
    return texto


def similitud(a, b):
    return SequenceMatcher(None, a, b).ratio()


def cargar_json(ruta, por_defecto):
    try:
        if os.path.exists(ruta):
            with open(ruta, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return por_defecto


def guardar_json(ruta, data):
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def registrar_memoria(tipo, contenido):
    memoria = cargar_json(ARCHIVO_MEMORIA, {"notas": [], "log": []})
    if tipo == "nota":
        memoria["notas"].append({
            "fecha": datetime.datetime.now().isoformat(),
            "contenido": contenido
        })
    memoria["log"].append({
        "fecha": datetime.datetime.now().isoformat(),
        "tipo": tipo,
        "contenido": contenido
    })
    memoria["log"] = memoria["log"][-120:]
    guardar_json(ARCHIVO_MEMORIA, memoria)


def contexto_memoria():
    memoria = cargar_json(ARCHIVO_MEMORIA, {"notas": [], "log": []})
    notas = memoria.get("notas", [])[-10:]
    return "\n".join([f"- {n['contenido']}" for n in notas])


def obtener_ip_local():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "desconocida"


def confirmar_accion(pregunta="¿Confirmo la acción, señor?"):
    hablar(pregunta)
    respuesta = escuchar()
    return any(x in respuesta for x in ["sí", "si", "confirmo", "adelante", "hazlo", "correcto"])

# ==========================================
# 4. Configuración del Micrófono
# ==========================================
recognizer = sr.Recognizer()
recognizer.dynamic_energy_threshold = True
recognizer.energy_threshold = 300
recognizer.pause_threshold = 0.8
recognizer.phrase_threshold = 0.3
recognizer.non_speaking_duration = 0.5

def escuchar():
    with sr.Microphone() as source:
        print("\n[Escuchando...]")
        recognizer.adjust_for_ambient_noise(source, duration=1.0)
        try:
            audio = recognizer.listen(source, timeout=5, phrase_time_limit=20)
            comando = recognizer.recognize_google(audio, language="es-MX")
            comando = normalizar_texto(comando)
            print(f"Tú: {comando}")
            return comando
        except sr.WaitTimeoutError:
            return ""
        except sr.UnknownValueError:
            return ""
        except sr.RequestError:
            hablar("No pude conectar con el reconocimiento de voz, señor.")
            return ""
        except Exception as e:
            print(f"Error al escuchar: {e}")
            return ""


def escuchar_activacion():
    with sr.Microphone() as source:
        print("\n[Esperando palabra de activación: Jarvis]")
        recognizer.adjust_for_ambient_noise(source, duration=1.0)
        try:
            audio = recognizer.listen(source, timeout=None, phrase_time_limit=4)
            activacion = recognizer.recognize_google(audio, language="es-MX")
            activacion = normalizar_texto(activacion)
            print(f"Detectado: {activacion}")
            return activacion
        except Exception:
            return ""


def dictado_continuo_loop():
    global MODO_DICTADO_CONTINUO
    hablar("Modo dictado continuo activado. Diga detener dictado para salir.")

    while MODO_DICTADO_CONTINUO:
        texto = escuchar()
        if not texto:
            continue

        if "detener dictado" in texto or "sal del dictado" in texto or "termina dictado" in texto:
            MODO_DICTADO_CONTINUO = False
            hablar("Modo dictado continuo desactivado, señor.")
            break

        if "nueva línea" in texto or "nuevo párrafo" in texto:
            pyautogui.press("enter")
            continue

        if "borra eso" in texto:
            pyautogui.press("backspace")
            continue

        texto = texto.replace(" punto ", ". ")
        texto = texto.replace(" coma ", ", ")
        texto = texto.replace(" dos puntos ", ": ")
        texto = texto.replace(" punto y coma ", "; ")
        escribir_en_ventana_seguro(texto + " ")

# ==========================================
# 5. FUNCIONES AVANZADAS DE TONY STARK
# ==========================================
def reporte_hardware():
    cpu = psutil.cpu_percent(interval=1)
    ram = psutil.virtual_memory().percent

    try:
        disco = psutil.disk_usage("C:\\").percent
    except Exception:
        disco = psutil.disk_usage(os.path.abspath(os.sep)).percent

    ip_local = obtener_ip_local()
    hablar(
        f"Señor, el procesador está al {cpu} por ciento de su capacidad. "
        f"La memoria RAM está al {ram} por ciento, el almacenamiento principal está al {disco} por ciento, "
        f"y la IP local actual es {ip_local}."
    )


def listar_procesos_importantes(limite=15):
    procesos = []
    for proc in psutil.process_iter(["pid", "name", "memory_percent"]):
        try:
            procesos.append(proc.info)
        except Exception:
            pass

    procesos = sorted(procesos, key=lambda x: x.get("memory_percent", 0), reverse=True)[:limite]
    texto = "Procesos más pesados detectados:\n"
    for p in procesos:
        texto += f"- {p.get('name')} | PID {p.get('pid')} | RAM {round(p.get('memory_percent', 0), 2)}%\n"

    print(texto)
    hablar("He listado en consola los procesos más pesados, señor.")


def cerrar_proceso_por_nombre(nombre):
    cerrados = 0
    for proc in psutil.process_iter(["name"]):
        try:
            if proc.info["name"] and nombre.lower() in proc.info["name"].lower():
                proc.terminate()
                cerrados += 1
        except Exception:
            pass
    return cerrados


def tomar_captura_pantalla():
    nombre_archivo = f"Captura_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
    ruta_completa = os.path.join(CARPETA_CAPTURAS, nombre_archivo)
    pyautogui.screenshot(ruta_completa)
    hablar("Captura de pantalla realizada y asegurada en su escritorio, señor.")
    return ruta_completa


def tomar_foto_camara():
    hablar("Activando la cámara principal para tomar la fotografía. Sonría, señor.")
    cap = cv2.VideoCapture(0)
    time.sleep(2)
    ret, frame = cap.read()
    if ret:
        nombre_archivo = f"Fotografia_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
        ruta_completa = os.path.join(CARPETA_FOTOS, nombre_archivo)
        cv2.imwrite(ruta_completa, frame)
        hablar("Fotografía capturada y guardada con éxito en su escritorio.")
        cap.release()
        return ruta_completa
    else:
        hablar("Hubo un error al intentar capturar la imagen desde la cámara.")
        cap.release()
        return None


def reconocimiento_facial():
    hablar("Iniciando protocolos de visión biométrica, señor. Buscando rostro en la cámara.")
    cap = cv2.VideoCapture(0)
    face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
    rostro_detectado = False
    tiempo_inicio = time.time()

    while time.time() - tiempo_inicio < 10:
        ret, frame = cap.read()
        if not ret:
            break
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        rostros = face_cascade.detectMultiScale(gray, 1.3, 5)

        for (x, y, w, h) in rostros:
            cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)

        if len(rostros) > 0:
            rostro_detectado = True
            cv2.imshow("Escáner de Jarvis", frame)
            cv2.waitKey(1000)
            break

        cv2.imshow("Buscando al señor Styx...", frame)
        if cv2.waitKey(1) & 0xFF == ord("q"):
            break

    cap.release()
    cv2.destroyAllWindows()

    if rostro_detectado:
        hablar("Escaneo biométrico completado. Identidad confirmada. Bienvenido de vuelta a la terminal, señor Styx.")
    else:
        hablar("No logré identificar su rostro en el tiempo establecido. Manteniendo alerta.")


def detectar_extension_codigo(codigo_puro):
    c = codigo_puro.lower()
    if "<html" in c or "<div" in c or "<body" in c:
        return ".html"
    if "function " in codigo_puro or "const " in codigo_puro or "let " in codigo_puro or "var " in codigo_puro:
        return ".js"
    if "body {" in c or ".class" in c or "font-family" in c:
        return ".css"
    if "import " in codigo_puro or "def " in codigo_puro or "print(" in codigo_puro or "class " in codigo_puro:
        return ".py"
    if "void setup()" in codigo_puro or "void loop()" in codigo_puro:
        return ".ino"
    if "select " in c or "create table" in c or "insert into" in c:
        return ".sql"
    if "@echo off" in c:
        return ".bat"
    return ".txt"


def extraer_y_guardar_codigo(texto_respuesta):
    if not texto_respuesta:
        return False

    bloques_codigo = re.findall(r"```(?:[\w.+-]+)?\n(.*?)```", texto_respuesta, flags=re.DOTALL)

    if bloques_codigo:
        for i, codigo_puro in enumerate(bloques_codigo):
            codigo_puro = codigo_puro.strip()
            extension = detectar_extension_codigo(codigo_puro)
            nombre_archivo = f"codigo_generado_{datetime.datetime.now().strftime('%H%M%S')}_{i}{extension}"
            ruta_archivo = os.path.join(CARPETA_CODIGO, nombre_archivo)
            with open(ruta_archivo, "w", encoding="utf-8") as f:
                f.write(codigo_puro)
            print(f"[!] CÓDIGO AISLADO Y GUARDADO EN: {ruta_archivo}")
        return True
    return False


def controlar_multimedia(accion):
    if accion == "pausar" or accion == "reproducir":
        pyautogui.press("playpause")
    elif accion == "siguiente":
        pyautogui.press("nexttrack")
    elif accion == "anterior":
        pyautogui.press("prevtrack")


def minimizar_todo():
    pyautogui.hotkey("win", "d")
    hablar("Minimizando ventanas y limpiando el área visual, señor.")


def cerrar_ventana_actual():
    pyautogui.hotkey("alt", "f4")
    hablar("Ventana cerrada.")


def maximizar_ventana():
    pyautogui.hotkey("win", "up")
    hablar("Ventana maximizada, señor.")


def restaurar_ventana():
    pyautogui.hotkey("win", "down")
    hablar("Ventana restaurada o minimizada según el estado actual, señor.")


def cambiar_ventana():
    pyautogui.hotkey("alt", "tab")
    hablar("Cambiando de ventana, señor.")


def bloquear_computadora():
    hablar("Bloqueando la terminal. Sistemas en modo de suspensión segura, señor.")
    ctypes.windll.user32.LockWorkStation()


def vaciar_papelera():
    try:
        resultado = ctypes.windll.shell32.SHEmptyRecycleBinW(None, None, 7)
        if resultado == 0:
            hablar("Protocolo de limpieza completado. Papelera de reciclaje vaciada exitosamente, señor.")
        else:
            hablar("La papelera de reciclaje ya se encuentra vacía, señor.")
    except Exception:
        hablar("No pude acceder a la papelera en este momento.")


def apagar_pc():
    hablar("Iniciando apagado del sistema, señor.")
    os.system("shutdown /s /t 3")


def reiniciar_pc():
    hablar("Iniciando reinicio del sistema, señor.")
    os.system("shutdown /r /t 3")


def suspender_pc():
    hablar("Enviando el sistema a suspensión, señor.")
    os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")


def obtener_clima(ciudad="Villahermosa"):
    hablar(f"Consultando los satélites meteorológicos para {ciudad}, señor...")
    try:
        url = f"https://wttr.in/{urllib.parse.quote(ciudad)}?format=%t+con+%C"
        res = requests.get(url, timeout=5)
        if res.status_code == 200:
            clima = res.text
            clima_hablado = clima.replace("+", " ").replace("C", "grados celsius")
            hablar(f"El reporte indica que en {ciudad} estamos a {clima_hablado}.")
        else:
            hablar("El servidor meteorológico no responde en este momento.")
    except Exception:
        hablar("Error de conexión al obtener el clima. Revisaré los receptores más tarde.")


def consultar_wikipedia(consulta):
    if wikipedia is None:
        hablar("Señor, el módulo de Wikipedia no está instalado. Ejecute pip install wikipedia.")
        return

    hablar(f"Investigando sobre {consulta} en la base de datos de Wikipedia...")
    try:
        resumen = wikipedia.summary(consulta, sentences=2)
        hablar("Esto es lo que encontré, señor.")
        hablar(resumen)
    except wikipedia.exceptions.DisambiguationError:
        hablar("La consulta es demasiado amplia y tiene múltiples significados. Por favor, sea más específico.")
    except wikipedia.exceptions.PageError:
        hablar("No encontré ningún artículo en Wikipedia que coincida con esa búsqueda exacta.")
    except Exception:
        hablar("Hubo un fallo en la conexión con la base de datos de Wikipedia.")


def contar_chiste():
    chistes = [
        "¿Por qué los programadores prefieren la oscuridad? Porque la luz atrae a los bugs.",
        "Hay 10 tipos de personas en el mundo: los que entienden binario y los que no.",
        "Un buen programador es aquel que mira a ambos lados antes de cruzar una calle de un solo sentido.",
        "¿Cuál es el animal favorito de un programador? El ratón.",
        "El hardware es lo que puedes golpear. El software es lo que solo puedes maldecir."
    ]
    hablar(random.choice(chistes))

# -------------------------------------------------------------------
# DOMÓTICA: CONTROL DEL FOCO STEREN SHOME-120 (Tuya)
# -------------------------------------------------------------------
def obtener_foco_steren():
    if tinytuya is None:
        hablar("Señor, no detecto los controladores del foco. Instale tinytuya en el entorno virtual.")
        return None

    DEVICE_ID = "PON_TU_DEVICE_ID_AQUI"
    IP_ADDRESS = "PON_TU_IP_AQUI"
    LOCAL_KEY = "PON_TU_LOCAL_KEY_AQUI"

    if DEVICE_ID == "PON_TU_DEVICE_ID_AQUI":
        return "SIMULACION"

    try:
        foco = tinytuya.BulbDevice(DEVICE_ID, IP_ADDRESS, LOCAL_KEY)
        foco.set_version(3.3)
        return foco
    except Exception as e:
        print(f"Error creando conexión con foco: {e}")
        return None


def control_foco_steren(accion, color=None, brillo=None):
    foco = obtener_foco_steren()
    if foco is None:
        hablar("No pude inicializar el foco Steren, señor.")
        return

    try:
        if foco == "SIMULACION":
            time.sleep(0.3)
            if accion == "encender":
                hablar("Conectando con la red inteligente... Foco Steren encendido, señor.")
            elif accion == "apagar":
                hablar("Cortando energía... Foco Steren apagado.")
            elif accion == "color":
                hablar(f"Modulando matriz LED. Cambio de iluminación a tono {color} completado.")
            elif accion == "brillo":
                hablar(f"Ajustando brillo del foco al {brillo} por ciento, señor.")
            return

        if accion == "encender":
            foco.turn_on()
            hablar("Foco de la habitación encendido, señor.")

        elif accion == "apagar":
            foco.turn_off()
            hablar("Foco de la habitación apagado, señor.")

        elif accion == "color" and color:
            if color in COLORES_FOCO:
                r, g, b = COLORES_FOCO[color]
                foco.set_colour(r, g, b)
                hablar(f"Cambiando la ambientación de la luz a {color}, señor.")
            else:
                hablar("No tengo ese color registrado en mi matriz visual.")

        elif accion == "brillo" and brillo is not None:
            brillo = max(10, min(100, int(brillo)))
            brillo_tuya = int((brillo / 100) * 1000)
            try:
                foco.set_brightness_percentage(brillo)
            except Exception:
                try:
                    foco.set_brightness(brillo_tuya)
                except Exception:
                    pass
            hablar(f"Brillo del foco ajustado al {brillo} por ciento, señor.")

    except Exception as e:
        hablar("Hubo una anomalía térmica o de red al intentar conectar con el foco Steren.")
        print(f"Error técnico foco Steren: {e}")

# ==========================================
# 6. Herramientas Offline y Utilidades Básicas
# ==========================================
def buscar_archivo(nombre_archivo, directorio_base):
    for raiz, carpetas, archivos in os.walk(directorio_base):
        if nombre_archivo in archivos:
            return os.path.join(raiz, nombre_archivo)
    return None


def buscar_archivos_por_fragmento(fragmento, directorio_base, limite=20):
    encontrados = []
    for raiz, carpetas, archivos in os.walk(directorio_base):
        for archivo in archivos:
            if fragmento.lower() in archivo.lower():
                encontrados.append(os.path.join(raiz, archivo))
                if len(encontrados) >= limite:
                    return encontrados
    return encontrados


def crear_carpeta(nombre_carpeta, base=None):
    if base is None:
        base = ESCRITORIO
    ruta = os.path.join(base, nombre_carpeta)
    os.makedirs(ruta, exist_ok=True)
    return ruta


def crear_archivo_txt(nombre_archivo, contenido="", base=None):
    if base is None:
        base = ESCRITORIO
    if not nombre_archivo.lower().endswith(".txt"):
        nombre_archivo += ".txt"
    ruta = os.path.join(base, nombre_archivo)
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(contenido)
    return ruta


def escribir_en_ventana(texto):
    pyautogui.write(texto, interval=0.03)


def escribir_en_ventana_seguro(texto):
    if pyperclip is not None:
        try:
            pyperclip.copy(texto)
            pyautogui.hotkey("ctrl", "v")
            return
        except Exception:
            pass
    escribir_en_ventana(texto)


def leer_portapapeles():
    if pyperclip is None:
        return None
    try:
        return pyperclip.paste()
    except Exception:
        return None


def copiar_al_portapapeles(texto):
    if pyperclip is None:
        return False
    try:
        pyperclip.copy(texto)
        return True
    except Exception:
        return False


def mover_cursor_centro():
    ancho, alto = pyautogui.size()
    pyautogui.moveTo(ancho // 2, alto // 2, duration=0.5)


def mover_cursor_arriba():
    ancho, alto = pyautogui.size()
    pyautogui.moveTo(ancho // 2, 0, duration=0.5)


def mover_cursor_abajo():
    ancho, alto = pyautogui.size()
    pyautogui.moveTo(ancho // 2, alto - 1, duration=0.5)


def mover_cursor_izquierda():
    ancho, alto = pyautogui.size()
    pyautogui.moveTo(0, alto // 2, duration=0.5)


def mover_cursor_derecha():
    ancho, alto = pyautogui.size()
    pyautogui.moveTo(ancho - 1, alto // 2, duration=0.5)


def mover_cursor_a(x, y):
    pyautogui.moveTo(x, y, duration=0.5)


def arrastrar_cursor_a(x, y):
    pyautogui.dragTo(x, y, duration=0.5, button="left")


def hacer_scroll_arriba(cantidad=500):
    pyautogui.scroll(cantidad)


def hacer_scroll_abajo(cantidad=500):
    pyautogui.scroll(-cantidad)


def abrir_ruta_si_existe(ruta):
    ruta_expandida = os.path.expandvars(ruta)
    if os.path.exists(ruta_expandida):
        os.startfile(ruta_expandida)
        return True
    return False


def abrir_app(nombre):
    nombre = normalizar_texto(nombre)
    nombre = ALIASES_APPS.get(nombre, nombre)

    apps_comando = {
        "calculadora": "start calc",
        "bloc de notas": "start notepad",
        "paint": "start mspaint",
        "explorador": "start explorer",
        "word": "start winword",
        "excel": "start excel",
        "powerpoint": "start powerpnt",
        "cmd": "start cmd",
        "terminal": "start wt",
    }

    apps_ruta = {
        "opera gx": [
            r"C:\Users\GAMER\AppData\Local\Programs\Opera GX\opera.exe",
            r"C:\Program Files\Opera GX\opera.exe"
        ],
        "brave": [
            r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
            r"C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe"
        ],
        "firefox": [
            r"C:\Program Files\Mozilla Firefox\firefox.exe",
            r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
        ],
        "spotify": [
            r"C:\Users\GAMER\AppData\Roaming\Spotify\Spotify.exe"
        ],
        "telegram": [
            r"C:\Users\GAMER\AppData\Roaming\Telegram Desktop\Telegram.exe",
            r"C:\Program Files\Telegram Desktop\Telegram.exe"
        ]
    }

    if nombre in apps_ruta:
        for ruta in apps_ruta[nombre]:
            if abrir_ruta_si_existe(ruta):
                return True
        try:
            os.system(f'start "" "{nombre}"')
            return True
        except Exception:
            return False

    if nombre in apps_comando:
        try:
            os.system(apps_comando[nombre])
            return True
        except Exception:
            return False

    return False


def abrir_url(url):
    try:
        webbrowser.open(url)
        return True
    except Exception:
        return False


def buscar_google(consulta):
    consulta_codificada = urllib.parse.quote(consulta)
    url = f"https://www.google.com/search?q={consulta_codificada}"
    return abrir_url(url)


def buscar_youtube(consulta):
    consulta_codificada = urllib.parse.quote(consulta)
    url = f"https://www.youtube.com/results?search_query={consulta_codificada}"
    return abrir_url(url)


def reproducir_en_youtube(consulta):
    consulta_codificada = urllib.parse.quote(consulta)
    url = f"https://www.youtube.com/results?search_query={consulta_codificada}"
    return abrir_url(url)


def decir_hora():
    ahora = datetime.datetime.now()
    return ahora.strftime("Son las %I:%M %p")


def decir_fecha():
    meses = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
        5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
        9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
    }
    hoy = datetime.datetime.now()
    return f"Hoy es {hoy.day} de {meses[hoy.month]} de {hoy.year}"


def subir_volumen(veces=5):
    for _ in range(veces):
        pyautogui.press("volumeup")


def bajar_volumen(veces=5):
    for _ in range(veces):
        pyautogui.press("volumedown")


def silenciar_volumen():
    pyautogui.press("volumemute")


def abrir_descargas():
    return abrir_ruta_si_existe(DESCARGAS)


def abrir_documentos():
    return abrir_ruta_si_existe(DOCUMENTOS)


def abrir_escritorio():
    return abrir_ruta_si_existe(ESCRITORIO)


def bloquear_input_por_segundos(segundos=5):
    hablar(f"Bloqueando entrada durante {segundos} segundos, señor.")
    ctypes.windll.user32.BlockInput(True)
    time.sleep(segundos)
    ctypes.windll.user32.BlockInput(False)
    hablar("Entrada restaurada.")


def ejecutar_comando_shell(comando):
    try:
        subprocess.Popen(comando, shell=True)
        return True
    except Exception:
        return False


def crear_nota(nombre, contenido):
    ruta = os.path.join(CARPETA_NOTAS, nombre if nombre.endswith(".txt") else nombre + ".txt")
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(contenido)
    return ruta


def temporizador_segundos(segundos, mensaje="Tiempo cumplido, señor."):
    def _run():
        time.sleep(segundos)
        hablar(mensaje)
    hilo = threading.Thread(target=_run, daemon=True)
    hilo.start()
    TEMPORIZADORES.append(hilo)


def rutina_estudio():
    hablar("Activando rutina de estudio, señor.")
    control_foco_steren("encender")
    control_foco_steren("color", "blanco")
    control_foco_steren("brillo", brillo=100)
    abrir_app("word")
    time.sleep(1)
    abrir_url("https://chatgpt.com")
    abrir_url("https://www.google.com")
    hablar("Entorno de estudio preparado.")


def rutina_gaming():
    hablar("Activando rutina gaming, señor.")
    control_foco_steren("encender")
    control_foco_steren("color", "azul")
    control_foco_steren("brillo", brillo=45)
    abrir_app("opera gx")
    abrir_url("https://discord.com/app")
    hablar("Escenario gaming preparado.")


def rutina_noche():
    hablar("Activando rutina nocturna, señor.")
    control_foco_steren("encender")
    control_foco_steren("color", "morado")
    control_foco_steren("brillo", brillo=20)
    hablar("Modo noche activado.")


def rutina_programacion():
    hablar("Activando rutina de programación, señor.")
    control_foco_steren("encender")
    control_foco_steren("color", "cian")
    control_foco_steren("brillo", brillo=75)
    abrir_app("terminal")
    abrir_url("https://github.com")
    abrir_url("https://chatgpt.com")
    hablar("Entorno de programación preparado, señor.")


def abrir_entorno_programacion():
    rutina_programacion()

# ==========================================
# 7. IA híbrida
# ==========================================
def intentar_parsear_json(texto):
    try:
        return json.loads(texto.strip())
    except Exception:
        m = re.search(r"\{.*\}", texto or "", flags=re.DOTALL)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                return None
    return None


def obtener_plan_accion_ia(comando):
    try:
        respuesta = modelo_ia.generate_content(f"{PROMPT_RUTEO_JSON}\n\nORDEN DEL USUARIO:\n{comando}")
        texto = getattr(respuesta, "text", "") or ""
        return intentar_parsear_json(texto)
    except Exception:
        return None


def responder_ia_general(comando):
    historial_respaldo.append({"role": "user", "content": comando})
    try:
        respuesta = memoria_jarvis.send_message(
            f"Contexto persistente del usuario:\n{contexto_memoria()}\n\nConsulta actual:\n{comando}"
        )
        texto_respuesta = ""
        if hasattr(respuesta, "text") and respuesta.text:
            texto_respuesta = respuesta.text.strip()
        if texto_respuesta:
            historial_respaldo.append({"role": "assistant", "content": texto_respuesta})
            extraer_y_guardar_codigo(texto_respuesta)
            texto_hablado = re.sub(
                r"```.*?```",
                " He escrito la lógica de programación solicitada. El código fuente ha sido empaquetado y exportado exitosamente a su escritorio, señor. ",
                texto_respuesta,
                flags=re.DOTALL
            )
            texto_hablado = texto_hablado.replace("*", "").replace("#", "")
            hablar(texto_hablado)
            return
        hablar("Recibí un paquete de datos vacío de la matriz principal, señor.")
    except Exception as e_gemini:
        print(f"\n[Fallo crítico en red principal: {e_gemini}]")
        print("Ejecutando protocolo de contingencia. Cambiando a red de respaldo (OpenAI)...")
        try:
            respuesta_openai = cliente_openai.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=historial_respaldo
            )
            texto_respuesta = respuesta_openai.choices[0].message.content.strip()
            if texto_respuesta:
                historial_respaldo.append({"role": "assistant", "content": texto_respuesta})
                extraer_y_guardar_codigo(texto_respuesta)
                texto_hablado = re.sub(
                    r"```.*?```",
                    " El archivo de código ha sido generado por la red secundaria y exportado a su escritorio de forma segura, señor. ",
                    texto_respuesta,
                    flags=re.DOTALL
                )
                texto_hablado = texto_hablado.replace("*", "").replace("#", "")
                hablar(texto_hablado)
            else:
                hablar("La red de respaldo también arrojó una latencia crítica. No hay respuesta, señor.")
        except Exception as e_openai:
            print(f"\n--- ERROR TÉCNICO TOTAL ---\n{e_openai}\n-------------------\n")
            hablar(
                "Señor, he perdido conexión absoluta con ambas redes neuronales principales. "
                "Los módulos de inteligencia están desconectados. Revise su conexión a internet "
                "o los fondos en sus claves de API."
            )

# ==========================================
# 8. Ejecutores por acción
# ==========================================
def ejecutar_accion_json(plan):
    if not isinstance(plan, dict):
        return False
    if plan.get("modo") != "accion":
        return False

    accion = plan.get("accion")
    params = plan.get("parametros", {}) or {}
    respuesta_corta = plan.get("respuesta_corta", "")
    if respuesta_corta:
        hablar(respuesta_corta)

    if accion == "abrir_app":
        return abrir_app(params.get("nombre", ""))
    if accion == "abrir_url":
        return abrir_url(params.get("texto", "") or params.get("url", ""))
    if accion == "buscar_google":
        return buscar_google(params.get("consulta", "") or params.get("texto", ""))
    if accion == "buscar_youtube":
        return buscar_youtube(params.get("consulta", "") or params.get("texto", ""))
    if accion == "reproducir_youtube":
        return reproducir_en_youtube(params.get("consulta", "") or params.get("texto", ""))
    if accion == "escribir_texto":
        escribir_en_ventana_seguro(params.get("texto", ""))
        return True
    if accion == "pegar_texto":
        texto = params.get("texto", "")
        if copiar_al_portapapeles(texto):
            pyautogui.hotkey("ctrl", "v")
        else:
            escribir_en_ventana_seguro(texto)
        return True
    if accion == "crear_carpeta":
        crear_carpeta(params.get("nombre", "NuevaCarpeta"))
        return True
    if accion == "crear_archivo_txt":
        crear_archivo_txt(params.get("nombre", "nota"), params.get("texto", ""))
        return True
    if accion == "tomar_captura":
        tomar_captura_pantalla()
        return True
    if accion == "tomar_foto":
        tomar_foto_camara()
        return True
    if accion == "reconocimiento_facial":
        reconocimiento_facial()
        return True
    if accion == "reporte_hardware":
        reporte_hardware()
        return True
    if accion == "listar_procesos":
        listar_procesos_importantes()
        return True
    if accion == "cerrar_proceso":
        nombre = params.get("nombre", "")
        if nombre:
            cantidad = cerrar_proceso_por_nombre(nombre)
            hablar(f"He intentado cerrar {cantidad} proceso o procesos con el nombre {nombre}, señor.")
        return True
    if accion == "bloquear_computadora":
        bloquear_computadora()
        return True
    if accion == "apagar_pc":
        apagar_pc()
        return True
    if accion == "reiniciar_pc":
        reiniciar_pc()
        return True
    if accion == "suspender_pc":
        suspender_pc()
        return True
    if accion == "vaciar_papelera":
        vaciar_papelera()
        return True
    if accion == "minimizar_todo":
        minimizar_todo()
        return True
    if accion == "cerrar_ventana":
        cerrar_ventana_actual()
        return True
    if accion == "maximizar_ventana":
        maximizar_ventana()
        return True
    if accion == "restaurar_ventana":
        restaurar_ventana()
        return True
    if accion == "cambiar_ventana":
        cambiar_ventana()
        return True
    if accion == "subir_volumen":
        subir_volumen(6)
        return True
    if accion == "bajar_volumen":
        bajar_volumen(6)
        return True
    if accion == "silenciar_volumen":
        silenciar_volumen()
        return True
    if accion == "pausar_media":
        controlar_multimedia("pausar")
        return True
    if accion == "siguiente_media":
        controlar_multimedia("siguiente")
        return True
    if accion == "anterior_media":
        controlar_multimedia("anterior")
        return True
    if accion == "abrir_descargas":
        abrir_descargas()
        return True
    if accion == "abrir_documentos":
        abrir_documentos()
        return True
    if accion == "abrir_escritorio":
        abrir_escritorio()
        return True
    if accion == "buscar_archivo":
        ruta = buscar_archivo(params.get("nombre", ""), USUARIO_PC)
        if ruta:
            print(ruta)
            hablar("Archivo localizado, señor.")
        else:
            hablar("No encontré ese archivo exacto, señor.")
        return True
    if accion == "buscar_archivos_fragmento":
        encontrados = buscar_archivos_por_fragmento(params.get("texto", ""), USUARIO_PC)
        for e in encontrados:
            print(e)
        hablar(f"He encontrado {len(encontrados)} coincidencias, señor.")
        return True
    if accion == "leer_portapapeles":
        texto = leer_portapapeles()
        hablar(texto if texto else "No pude leer el portapapeles, señor.")
        return True
    if accion == "resumir_portapapeles":
        texto = leer_portapapeles()
        if texto:
            responder_ia_general(f"Resume este texto en español de forma clara:\n\n{texto}")
        return True
    if accion == "traducir_portapapeles_ingles":
        texto = leer_portapapeles()
        if texto:
            responder_ia_general(f"Traduce al inglés este texto:\n\n{texto}")
        return True
    if accion == "activar_dictado":
        global MODO_DICTADO_CONTINUO
        if not MODO_DICTADO_CONTINUO:
            MODO_DICTADO_CONTINUO = True
            threading.Thread(target=dictado_continuo_loop, daemon=True).start()
        return True
    if accion == "desactivar_dictado":
        MODO_DICTADO_CONTINUO = False
        return True
    if accion == "foco_encender":
        control_foco_steren("encender")
        return True
    if accion == "foco_apagar":
        control_foco_steren("apagar")
        return True
    if accion == "foco_color":
        control_foco_steren("color", params.get("color", "blanco"))
        return True
    if accion == "foco_brillo":
        control_foco_steren("brillo", brillo=params.get("brillo", 50))
        return True
    if accion == "mover_cursor":
        mover_cursor_a(int(params.get("x", 0)), int(params.get("y", 0)))
        return True
    if accion == "arrastrar_cursor":
        arrastrar_cursor_a(int(params.get("x", 0)), int(params.get("y", 0)))
        return True
    if accion == "click":
        pyautogui.click()
        return True
    if accion == "doble_click":
        pyautogui.doubleClick()
        return True
    if accion == "click_derecho":
        pyautogui.rightClick()
        return True
    if accion == "scroll_arriba":
        hacer_scroll_arriba()
        return True
    if accion == "scroll_abajo":
        hacer_scroll_abajo()
        return True
    if accion == "presionar_tecla":
        tecla = params.get("tecla", "enter")
        pyautogui.press(tecla)
        return True
    if accion == "rutina_estudio":
        rutina_estudio()
        return True
    if accion == "rutina_gaming":
        rutina_gaming()
        return True
    if accion == "rutina_noche":
        rutina_noche()
        return True
    if accion == "rutina_programacion":
        rutina_programacion()
        return True
    if accion == "abrir_entorno_programacion":
        abrir_entorno_programacion()
        return True

    return False

# ==========================================
# 9. Bucle Principal
# ==========================================
if __name__ == "__main__":
    hablar(
        "Todos los sistemas al mil por ciento. Doble red neuronal activa. "
        "Módulos de programación, domótica, cámara, automatización, memoria persistente, "
        "ruteo inteligente y sensores de hardware en línea. "
        "A sus órdenes, señor Styx."
    )

    carpeta_usuario = USUARIO_PC

    while True:
        activacion = escuchar_activacion()

        if NOMBRE_ACTIVACION in activacion:
            if MODO_SUSPENDIDO:
                MODO_SUSPENDIDO = False
                hablar("Modo suspendido desactivado. Sistema restaurado, señor.")
                continue

            hablar("Lo escucho, señor.")
            comando = escuchar()

            if comando:
                registrar_memoria("usuario", comando)

                # ------------------------------------------
                # DOMÓTICA / FOCO STEREN
                # ------------------------------------------
                if "enciende el foco" in comando or "prende el foco" in comando or "prende la luz" in comando or "enciende la luz" in comando:
                    control_foco_steren("encender")

                elif "apaga el foco" in comando or "apaga la luz" in comando:
                    control_foco_steren("apagar")

                elif "pon el foco en rojo" in comando or "luz roja" in comando:
                    control_foco_steren("color", "rojo")

                elif "pon el foco en azul" in comando or "luz azul" in comando:
                    control_foco_steren("color", "azul")

                elif "pon el foco en verde" in comando or "luz verde" in comando:
                    control_foco_steren("color", "verde")

                elif "pon el foco en blanco" in comando or "luz blanca" in comando:
                    control_foco_steren("color", "blanco")

                elif "pon el foco en amarillo" in comando or "luz amarilla" in comando:
                    control_foco_steren("color", "amarillo")

                elif "pon el foco en morado" in comando or "luz morada" in comando:
                    control_foco_steren("color", "morado")

                elif "pon el foco en naranja" in comando or "luz naranja" in comando:
                    control_foco_steren("color", "naranja")

                elif "pon el foco en rosa" in comando or "luz rosa" in comando:
                    control_foco_steren("color", "rosa")

                elif "brillo al 100" in comando:
                    control_foco_steren("brillo", brillo=100)

                elif "brillo al 75" in comando:
                    control_foco_steren("brillo", brillo=75)

                elif "brillo al 50" in comando:
                    control_foco_steren("brillo", brillo=50)

                elif "brillo al 25" in comando:
                    control_foco_steren("brillo", brillo=25)

                # ------------------------------------------
                # SISTEMA Y HARDWARE
                # ------------------------------------------
                elif "estado del sistema" in comando or "cómo está el hardware" in comando or "cómo está la pc" in comando:
                    reporte_hardware()

                elif "lista procesos" in comando or "procesos abiertos" in comando:
                    listar_procesos_importantes()

                elif "cierra proceso" in comando:
                    nombre = comando.replace("cierra proceso", "").strip()
                    if nombre:
                        if confirmar_accion(f"¿Confirma cerrar procesos relacionados con {nombre}, señor?"):
                            cantidad = cerrar_proceso_por_nombre(nombre)
                            hablar(f"He intentado cerrar {cantidad} proceso o procesos con el nombre {nombre}, señor.")
                        else:
                            hablar("Acción cancelada.")
                    else:
                        hablar("No entendí qué proceso debo cerrar, señor.")

                elif "bloquea la computadora" in comando or "bloquea el sistema" in comando or "modo seguridad" in comando:
                    bloquear_computadora()

                elif "apaga la computadora" in comando or "apaga la pc" in comando:
                    if confirmar_accion("¿Desea apagar la computadora, señor?"):
                        apagar_pc()
                    else:
                        hablar("Acción cancelada.")

                elif "reinicia la computadora" in comando or "reinicia la pc" in comando:
                    if confirmar_accion("¿Desea reiniciar la computadora, señor?"):
                        reiniciar_pc()
                    else:
                        hablar("Acción cancelada.")

                elif "suspende la computadora" in comando or "duerme la computadora" in comando:
                    suspender_pc()

                elif "modo suspendido" in comando or "suspéndete" in comando or "suspendete" in comando:
                    MODO_SUSPENDIDO = True
                    hablar("Modo suspendido activado. Esperaré una nueva activación para volver, señor.")

                elif "vacía la papelera" in comando or "limpia la papelera" in comando or "elimina la basura" in comando:
                    if confirmar_accion("¿Desea vaciar la papelera, señor?"):
                        vaciar_papelera()
                    else:
                        hablar("Acción cancelada.")

                elif "minimiza todo" in comando or "limpia la pantalla" in comando or "esconde todo" in comando:
                    minimizar_todo()

                elif "cierra la ventana" in comando or "cierra esto" in comando:
                    cerrar_ventana_actual()

                elif "maximiza la ventana" in comando:
                    maximizar_ventana()

                elif "restaura la ventana" in comando:
                    restaurar_ventana()

                elif "cambia de ventana" in comando or "siguiente ventana" in comando:
                    cambiar_ventana()

                elif "abre el administrador de tareas" in comando or "estado de procesos" in comando:
                    pyautogui.hotkey("ctrl", "shift", "esc")
                    hablar("Desplegando el administrador de tareas de Windows.")

                elif "abre configuraciones" in comando or "abre ajustes" in comando:
                    os.system("start ms-settings:")
                    hablar("Abriendo panel de ajustes del sistema operativo.")

                elif "bloquea entrada" in comando:
                    bloquear_input_por_segundos(5)

                elif "crea una nota llamada" in comando:
                    nombre = comando.replace("crea una nota llamada", "").strip()
                    if nombre:
                        ruta = crear_nota(nombre, "")
                        hablar("Nota creada, señor.")
                        print(ruta)
                    else:
                        hablar("No entendí el nombre de la nota, señor.")

                elif "pon un temporizador de" in comando:
                    numeros = re.findall(r"\d+", comando)
                    if numeros:
                        segundos = int(numeros[0])
                        if "minuto" in comando:
                            segundos *= 60
                        temporizador_segundos(segundos, "Señor, el temporizador ha terminado.")
                        hablar("Temporizador activado, señor.")
                    else:
                        hablar("No pude determinar el tiempo del temporizador.")

                # ------------------------------------------
                # MULTIMEDIA
                # ------------------------------------------
                elif "pausa la música" in comando or "reproduce la música" in comando or "ponle pausa" in comando:
                    controlar_multimedia("pausar")
                    hablar("Control multimedia ejecutado, señor.")

                elif "siguiente canción" in comando or "pasa la canción" in comando or "pon la que sigue" in comando:
                    controlar_multimedia("siguiente")
                    hablar("Cambiando a la siguiente pista de audio.")

                elif "canción anterior" in comando or "regresa la canción" in comando:
                    controlar_multimedia("anterior")
                    hablar("Regresando a la pista anterior.")

                elif "sube el volumen" in comando:
                    subir_volumen(6)
                    hablar("Volumen de sistema aumentado, señor.")

                elif "baja el volumen" in comando:
                    bajar_volumen(6)
                    hablar("Volumen de sistema reducido, señor.")

                elif "silencia" in comando or "mute" in comando:
                    silenciar_volumen()
                    hablar("Audio completamente silenciado, señor.")

                # ------------------------------------------
                # CÁMARA
                # ------------------------------------------
                elif "toma una captura" in comando or "captura de pantalla" in comando or "toma foto a la pantalla" in comando:
                    tomar_captura_pantalla()

                elif "tómame una foto" in comando or "foto con la cámara" in comando or "usa la cámara" in comando:
                    tomar_foto_camara()

                elif "reconocimiento facial" in comando or "activa la cámara" in comando or "quién soy" in comando or "escanéame" in comando:
                    reconocimiento_facial()

                # ------------------------------------------
                # ARCHIVOS Y CARPETAS
                # ------------------------------------------
                elif "crea una carpeta llamada" in comando:
                    nombre = comando.replace("crea una carpeta llamada", "").strip()
                    if nombre:
                        ruta = crear_carpeta(nombre)
                        hablar(f"Carpeta creada, señor. Nombre: {nombre}")
                        print("Ruta creada:", ruta)
                    else:
                        hablar("No entendí el nombre de la carpeta, señor.")

                elif "crea carpeta llamada" in comando:
                    nombre = comando.replace("crea carpeta llamada", "").strip()
                    if nombre:
                        ruta = crear_carpeta(nombre)
                        hablar(f"Carpeta creada, señor. Nombre: {nombre}")
                        print("Ruta creada:", ruta)
                    else:
                        hablar("No entendí el nombre de la carpeta, señor.")

                elif "crea archivo de texto llamado" in comando:
                    nombre = comando.replace("crea archivo de texto llamado", "").strip()
                    if nombre:
                        ruta = crear_archivo_txt(nombre, "")
                        hablar("Archivo de texto creado en el escritorio, señor.")
                        print("Ruta creada:", ruta)
                    else:
                        hablar("No entendí el nombre del archivo, señor.")

                elif "busca archivo" in comando:
                    nombre = comando.replace("busca archivo", "").strip()
                    if nombre:
                        ruta = buscar_archivo(nombre, carpeta_usuario)
                        if ruta:
                            print("Archivo encontrado:", ruta)
                            hablar("Archivo localizado, señor.")
                        else:
                            hablar("No encontré ese archivo exacto, señor.")
                    else:
                        hablar("No entendí qué archivo busca, señor.")

                elif "busca archivos que tengan" in comando:
                    fragmento = comando.replace("busca archivos que tengan", "").strip()
                    if fragmento:
                        encontrados = buscar_archivos_por_fragmento(fragmento, carpeta_usuario)
                        if encontrados:
                            print("Coincidencias encontradas:")
                            for e in encontrados:
                                print(e)
                            hablar(f"He encontrado {len(encontrados)} coincidencias. Las he listado en consola, señor.")
                        else:
                            hablar("No encontré coincidencias con ese fragmento, señor.")
                    else:
                        hablar("No entendí qué fragmento debo buscar, señor.")

                elif "abre descargas" in comando:
                    if abrir_descargas():
                        hablar("Abriendo la carpeta de descargas, señor.")
                    else:
                        hablar("No pude abrir la carpeta de descargas, señor.")

                elif "abre documentos" in comando:
                    if abrir_documentos():
                        hablar("Abriendo la carpeta de documentos, señor.")
                    else:
                        hablar("No pude abrir documentos, señor.")

                elif "abre escritorio" in comando:
                    if abrir_escritorio():
                        hablar("Abriendo el escritorio, señor.")
                    else:
                        hablar("No pude abrir el escritorio, señor.")

                elif "abre mi libro" in comando or "teoría del eterno" in comando:
                    hablar("Buscando el manuscrito, señor.")
                    ruta = buscar_archivo("La Teoría del Eterno.docx", carpeta_usuario)
                    if ruta:
                        os.startfile(ruta)
                        hablar("Archivo abierto, señor.")
                    else:
                        hablar("No logré localizar el archivo, señor.")

                elif "minecraft" in comando:
                    hablar("Buscando la imagen, señor.")
                    ruta = buscar_archivo("minecraft.jpg", carpeta_usuario)
                    if ruta:
                        os.startfile(ruta)
                        hablar("Imagen abierta, señor.")
                    else:
                        hablar("No encontré la imagen, señor.")

                # ------------------------------------------
                # ESCRITURA / PORTAPAPELES / DICTADO
                # ------------------------------------------
                elif "escribe en word" in comando:
                    texto = comando.replace("escribe en word", "").strip()
                    if texto:
                        hablar("Escribiendo en Word, señor.")
                        time.sleep(1)
                        escribir_en_ventana_seguro(texto)
                    else:
                        hablar("No me dijiste qué escribir, señor.")

                elif "escribe en el documento" in comando:
                    texto = comando.replace("escribe en el documento", "").strip()
                    if texto:
                        hablar("Escribiendo en el documento, señor.")
                        time.sleep(1)
                        escribir_en_ventana_seguro(texto)
                    else:
                        hablar("No me dijiste qué escribir, señor.")

                elif "escribe aquí" in comando:
                    texto = comando.replace("escribe aquí", "").strip()
                    if texto:
                        hablar("Escribiendo exactamente en la posición actual, señor.")
                        time.sleep(0.5)
                        escribir_en_ventana_seguro(texto)
                    else:
                        hablar("No me dijiste qué escribir, señor.")

                elif "pega esto" in comando:
                    texto = comando.replace("pega esto", "").strip()
                    if texto:
                        if copiar_al_portapapeles(texto):
                            pyautogui.hotkey("ctrl", "v")
                            hablar("Contenido pegado, señor.")
                        else:
                            escribir_en_ventana_seguro(texto)
                            hablar("Contenido escrito, señor.")
                    else:
                        hablar("No me dijiste qué pegar, señor.")

                elif "lee el portapapeles" in comando:
                    texto = leer_portapapeles()
                    if texto:
                        hablar(f"El portapapeles contiene: {texto}")
                    else:
                        hablar("No pude leer el portapapeles, señor.")

                elif "resume lo copiado" in comando or "resume el portapapeles" in comando:
                    texto = leer_portapapeles()
                    if texto:
                        responder_ia_general(f"Resume este texto en español de forma clara:\n\n{texto}")
                    else:
                        hablar("No hay texto en el portapapeles, señor.")

                elif "traduce lo copiado al inglés" in comando or "traduce el portapapeles al inglés" in comando:
                    texto = leer_portapapeles()
                    if texto:
                        responder_ia_general(f"Traduce al inglés este texto:\n\n{texto}")
                    else:
                        hablar("No hay texto en el portapapeles, señor.")

                elif "dictado continuo" in comando or "modo dictado" in comando:
                    if not MODO_DICTADO_CONTINUO:
                        MODO_DICTADO_CONTINUO = True
                        hilo = threading.Thread(target=dictado_continuo_loop, daemon=True)
                        hilo.start()
                    else:
                        hablar("El modo dictado ya está activo, señor.")

                elif "escribe" in comando:
                    texto = comando.replace("escribe", "", 1).strip()
                    if texto:
                        hablar("Escribiendo dictado, señor.")
                        time.sleep(0.5)
                        escribir_en_ventana_seguro(texto)
                    else:
                        hablar("No me dijiste qué escribir, señor.")

                elif "nueva línea" in comando or "nuevo párrafo" in comando:
                    pyautogui.press("enter")
                    hablar("Hecho, señor.")

                elif "borra eso" in comando:
                    pyautogui.press("backspace")
                    hablar("Texto eliminado, señor.")

                elif "guarda el documento" in comando or "guarda" in comando:
                    pyautogui.hotkey("ctrl", "s")
                    hablar("Documento y progreso guardado de manera segura, señor.")

                elif "selecciona todo" in comando:
                    pyautogui.hotkey("ctrl", "a")
                    hablar("Todo seleccionado, señor.")

                elif "copia eso" in comando:
                    pyautogui.hotkey("ctrl", "c")
                    hablar("Texto copiado al portapapeles, señor.")

                elif "pega eso" in comando:
                    pyautogui.hotkey("ctrl", "v")
                    hablar("Texto pegado con éxito, señor.")

                elif "corta eso" in comando:
                    pyautogui.hotkey("ctrl", "x")
                    hablar("Cortado y asegurado, señor.")

                elif "deshacer" in comando:
                    pyautogui.hotkey("ctrl", "z")
                    hablar("Acción deshecha, señor.")

                elif "rehacer" in comando:
                    pyautogui.hotkey("ctrl", "y")
                    hablar("Acción rehecha, señor.")

                # ------------------------------------------
                # CURSOR / MOUSE
                # ------------------------------------------
                elif "mueve el cursor al centro" in comando:
                    mover_cursor_centro()
                    hablar("Cursor movido al centro, señor.")

                elif "mueve el cursor arriba" in comando:
                    mover_cursor_arriba()
                    hablar("Cursor movido arriba, señor.")

                elif "mueve el cursor abajo" in comando:
                    mover_cursor_abajo()
                    hablar("Cursor movido abajo, señor.")

                elif "mueve el cursor a la izquierda" in comando:
                    mover_cursor_izquierda()
                    hablar("Cursor movido a la izquierda, señor.")

                elif "mueve el cursor a la derecha" in comando:
                    mover_cursor_derecha()
                    hablar("Cursor movido a la derecha, señor.")

                elif "mueve el cursor a" in comando:
                    numeros = re.findall(r"\d+", comando)
                    if len(numeros) >= 2:
                        x = int(numeros[0])
                        y = int(numeros[1])
                        mover_cursor_a(x, y)
                        hablar(f"Cursor movido a las coordenadas {x}, {y}, señor.")
                    else:
                        hablar("No pude detectar coordenadas válidas, señor.")

                elif "arrastra a" in comando:
                    numeros = re.findall(r"\d+", comando)
                    if len(numeros) >= 2:
                        x = int(numeros[0])
                        y = int(numeros[1])
                        arrastrar_cursor_a(x, y)
                        hablar(f"Arrastre completado hacia {x}, {y}, señor.")
                    else:
                        hablar("No pude detectar coordenadas válidas para el arrastre, señor.")

                elif "haz clic" in comando or "da clic" in comando:
                    pyautogui.click()
                    hablar("Hecho, señor.")

                elif "doble clic" in comando:
                    pyautogui.doubleClick()
                    hablar("Hecho, señor.")

                elif "clic derecho" in comando:
                    pyautogui.rightClick()
                    hablar("Hecho, señor.")

                elif "presiona enter" in comando:
                    pyautogui.press("enter")
                    hablar("Enter presionado, señor.")

                elif "scroll arriba" in comando:
                    hacer_scroll_arriba()
                    hablar("Desplazamiento hacia arriba completado, señor.")

                elif "scroll abajo" in comando:
                    hacer_scroll_abajo()
                    hablar("Desplazamiento hacia abajo completado, señor.")

                # ------------------------------------------
                # HORA Y FECHA
                # ------------------------------------------
                elif "qué hora" in comando or "dime la hora" in comando:
                    hablar(decir_hora())

                elif "qué fecha es" in comando or "qué día es hoy" in comando or "dime la fecha" in comando:
                    hablar(decir_fecha())

                # ------------------------------------------
                # APPS
                # ------------------------------------------
                elif "abre calculadora" in comando:
                    if abrir_app("calculadora"):
                        hablar("Abriendo aplicación de calculadora, señor.")
                    else:
                        hablar("No pude abrir la calculadora, señor.")

                elif "abre bloc de notas" in comando or "abre notepad" in comando:
                    if abrir_app("bloc de notas"):
                        hablar("Abriendo el bloc de notas tradicional, señor.")
                    else:
                        hablar("No pude abrir bloc de notas, señor.")

                elif "abre paint" in comando:
                    if abrir_app("paint"):
                        hablar("Abriendo Microsoft Paint, señor.")
                    else:
                        hablar("No pude abrir Paint, señor.")

                elif "abre explorador" in comando or "abre explorador de archivos" in comando:
                    if abrir_app("explorador"):
                        hablar("Abriendo explorador de archivos maestro, señor.")
                    else:
                        hablar("No pude abrir el explorador, señor.")

                elif "abre word" in comando:
                    if abrir_app("word"):
                        hablar("Lanzando procesador de texto Microsoft Word, señor.")
                    else:
                        hablar("No pude abrir Word, señor.")

                elif "abre excel" in comando:
                    if abrir_app("excel"):
                        hablar("Lanzando hojas de cálculo de Microsoft Excel, señor.")
                    else:
                        hablar("No pude abrir Excel, señor.")

                elif "abre powerpoint" in comando:
                    if abrir_app("powerpoint"):
                        hablar("Lanzando Microsoft PowerPoint, señor.")
                    else:
                        hablar("No pude abrir PowerPoint, señor.")

                elif "abre cmd" in comando or "abre consola" in comando:
                    if abrir_app("cmd"):
                        hablar("Abriendo consola de comandos principal, señor.")
                    else:
                        hablar("No pude abrir la consola, señor.")

                elif "abre terminal" in comando:
                    if abrir_app("terminal"):
                        hablar("Abriendo Windows Terminal, señor.")
                    else:
                        hablar("No pude abrir la terminal, señor.")

                elif "abre opera gx" in comando or "abre opera" in comando:
                    if abrir_app("opera gx"):
                        hablar("Navegador Opera GX en pantalla, señor.")
                    else:
                        hablar("No encontré Opera GX en la base de datos de rutas, señor.")

                elif "abre brave" in comando:
                    if abrir_app("brave"):
                        hablar("Navegador Brave ejecutándose, señor.")
                    else:
                        hablar("No encontré Brave, señor.")

                elif "abre firefox" in comando:
                    if abrir_app("firefox"):
                        hablar("Navegador Mozilla Firefox activo, señor.")
                    else:
                        hablar("No encontré Firefox, señor.")

                elif "abre spotify" in comando:
                    if abrir_app("spotify"):
                        hablar("Spotify en ejecución, señor.")
                    else:
                        hablar("No encontré Spotify, señor.")

                elif "abre telegram" in comando:
                    if abrir_app("telegram"):
                        hablar("Telegram Desktop ejecutándose, señor.")
                    else:
                        hablar("No encontré Telegram, señor.")

                elif "abre entorno de programación" in comando or "abre mi entorno de programación" in comando:
                    abrir_entorno_programacion()

                elif "rutina estudio" in comando or "modo estudio" in comando:
                    rutina_estudio()

                elif "rutina gaming" in comando or "modo gaming" in comando:
                    rutina_gaming()

                elif "rutina noche" in comando or "modo noche" in comando:
                    rutina_noche()

                elif "rutina programación" in comando or "rutina programacion" in comando:
                    rutina_programacion()

                # ------------------------------------------
                # WEB
                # ------------------------------------------
                elif "abre youtube" in comando:
                    if abrir_url("https://www.youtube.com"):
                        hablar("Abriendo plataforma de videos YouTube, señor.")
                    else:
                        hablar("No pude abrir YouTube, señor.")

                elif "abre gmail" in comando:
                    if abrir_url("https://mail.google.com"):
                        hablar("Desplegando su bandeja de correo electrónico en Gmail, señor.")
                    else:
                        hablar("No pude abrir Gmail, señor.")

                elif "abre discord" in comando:
                    if abrir_url("https://discord.com/app"):
                        hablar("Estableciendo conexión con los servidores de Discord, señor.")
                    else:
                        hablar("No pude abrir Discord, señor.")

                elif "abre chatgpt" in comando:
                    if abrir_url("https://chatgpt.com"):
                        hablar("Abriendo interfaz web de OpenAI ChatGPT, señor.")
                    else:
                        hablar("No pude abrir ChatGPT, señor.")

                elif "abre github" in comando:
                    if abrir_url("https://github.com"):
                        hablar("Accediendo a los repositorios en GitHub, señor.")
                    else:
                        hablar("No pude abrir GitHub, señor.")

                elif "abre google" in comando:
                    if abrir_url("https://www.google.com"):
                        hablar("Motor de búsqueda Google principal en pantalla, señor.")
                    else:
                        hablar("No pude abrir Google, señor.")

                elif "abre whatsapp web" in comando:
                    if abrir_url("https://web.whatsapp.com"):
                        hablar("Abriendo WhatsApp Web, señor.")
                    else:
                        hablar("No pude abrir WhatsApp Web, señor.")

                elif "busca en google" in comando:
                    consulta = comando.replace("busca en google", "").strip()
                    if consulta:
                        buscar_google(consulta)
                        hablar(f"Buscando información sobre {consulta} en la red principal, señor.")
                    else:
                        hablar("No escuché la directriz de búsqueda, señor.")

                elif "busca en youtube" in comando:
                    consulta = comando.replace("busca en youtube", "").strip()
                    if consulta:
                        buscar_youtube(consulta)
                        hablar(f"Buscando el contenido audiovisual de {consulta} en YouTube, señor.")
                    else:
                        hablar("No escuché qué desea buscar, señor.")

                elif "reproduce" in comando:
                    consulta = comando.replace("reproduce", "").strip()
                    if consulta:
                        reproducir_en_youtube(consulta)
                        hablar(f"Iniciando reproducción automática de {consulta}, señor.")
                    else:
                        hablar("No me dijo qué reproducir, señor.")

                elif "cómo está el clima" in comando or "dime el clima" in comando or "reporte del clima" in comando:
                    obtener_clima("Villahermosa")

                elif "busca en wikipedia" in comando or "investiga sobre" in comando:
                    consulta = comando.replace("busca en wikipedia", "").replace("investiga sobre", "").strip()
                    if consulta:
                        consultar_wikipedia(consulta)
                    else:
                        hablar("No capté el tema a investigar en los servidores, señor.")

                elif "cuéntame un chiste" in comando or "dime algo gracioso" in comando or "modo humor" in comando:
                    contar_chiste()

                # ------------------------------------------
                # APAGADO DEL ASISTENTE
                # ------------------------------------------
                elif "apagar" in comando or "salir" in comando or "apágate" in comando or "apagate" in comando:
                    hablar("Desconectando matriz principal. Nos vemos luego, señor Styx.")
                    break

                # ------------------------------------------
                # RUTEO IA A ACCIÓN JSON O CHAT
                # ------------------------------------------
                else:
                    plan = obtener_plan_accion_ia(comando)
                    ejecutado = ejecutar_accion_json(plan)
                    if not ejecutado:
                        print("[Procesando consulta en red neuronal...]")
                        responder_ia_general(comando)

                registrar_memoria("asistente", "Comando procesado: " + comando)


