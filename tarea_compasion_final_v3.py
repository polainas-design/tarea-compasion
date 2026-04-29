from psychopy import prefs
prefs.hardware['audioLib'] = ['PTB']
prefs.hardware['audioLatencyMode'] = 3  # Modo de baja latencia

import pandas as pd
from psychopy import visual, core, event, gui, logging, monitors
import os
import gc
import serial
from datetime import datetime

# pylsl se importa más abajo, después de la ventana de diálogo

# CONFIGURACIÓN
RUTA_FIJA = r'C:\Users\p_ull\OneDrive\Documentos\all_data_tarea_experimento\experimento_UMCE'
NOMBRE_CARPETA_VIDEOS = 'stimuli_video_mp4'
CARPETA_SALIDA = os.path.join(RUTA_FIJA, 'datos_salida')

# TIEMPOS EN SEG
TIEMPO_FIJACION = 5.0
TIEMPO_CONTEXTO = 10.0
TIEMPO_VIDEO = 10.0
TIEMPO_RATING = 6.0

# SIZE VIDEO
VIDEO_SIZE = (960, 720)

# --- MARCADORES LSL (ANT Neuro EEG) ---
LSL_FIJACION = 1
LSL_CONTEXTO = 2
LSL_VIDEO = 3
LSL_RATING = 4

# ARDUINO
ARDUINO_BAUDRATE = 115200

# PSYCHOPY
logging.console.setLevel(logging.WARNING)

try:
    os.chdir(RUTA_FIJA)
except FileNotFoundError:
    print(f"ERROR: La ruta no existe: {RUTA_FIJA}")
    core.quit()

os.makedirs(CARPETA_SALIDA, exist_ok=True)

# VENTANA DE DIALOGO INICIAL
info = {
    'Participante': '001',
    'Sesion': '1',
    'Grupo': ['A', 'B'],
    'Puerto_Arduino': 'COM5',
    'Modo_Simulacion': False,
    'Usar_LSL': True
}
if not gui.DlgFromDict(dictionary=info, title='Tarea Compasión').OK:
    core.quit()

grupo = info['Grupo'].upper().strip()
if grupo not in ('A', 'B'):
    print(f"ERROR: Grupo debe ser A o B, se recibió '{grupo}'")
    core.quit()

MODO_SIMULACION = info['Modo_Simulacion']
USAR_LSL = info['Usar_LSL']
ARDUINO_PORT = info['Puerto_Arduino'].strip()

# Importar pylsl si se activó LSL
if USAR_LSL:
    from pylsl import StreamInfo, StreamOutlet
    print("pylsl importado correctamente.")

# SET DE NOMBRE:ID_Grupo_diamesañohora.csv
timestamp_inicio = datetime.now().strftime('%d%m%Y%H%M%S')
nombre_archivo_salida = os.path.join(
    CARPETA_SALIDA,
    f"{info['Participante']}_{grupo}_{timestamp_inicio}.csv"
)

# Seleccionar archivo de condiciones según grupo
archivo_condiciones = f'condiciones_{grupo.lower()}.xlsx'
print(f"Grupo: {grupo} -> Archivo: {archivo_condiciones}")

# 2. TRIGGERS
# BNC: Pin 2 --> BNC 1 (Start), Pin 3 --> BNC 2 (Stop)
# LSL: Marcadores por fase enviados al ANT Neuro vía red

TRIG_START = ord('H')  # 72
TRIG_STOP = ord('L')   # 76


# 3. CONEXIÓN ARDUINO + LSL
def conectar_arduino():
    """Conecta al Arduino en el puerto especificado."""
    try:
        ard = serial.Serial(
            port=ARDUINO_PORT,
            baudrate=ARDUINO_BAUDRATE,
            timeout=2,
            write_timeout=1
        )
        ard.setDTR(False)
        core.wait(0.1)
        ard.setDTR(True)
        core.wait(2.0)
        ard.flushInput()
        ard.flushOutput()

        intentos = 0
        while intentos < 15:
            if ard.in_waiting > 0:
                linea = ard.readline().decode('utf-8', errors='ignore').strip()
                if 'READY' in linea:
                    print(f"Arduino conectado en {ARDUINO_PORT}")
                    return ard
            intentos += 1
            core.wait(0.1)

        print(f"Arduino conectado en {ARDUINO_PORT} (sin READY, puerto activo)")
        return ard

    except Exception as e:
        print(f"ERROR Arduino en {ARDUINO_PORT}: {e}")
        print("Verifica el puerto en la ventana de configuración.")
        core.quit()


if not MODO_SIMULACION:
    arduino = conectar_arduino()
else:
    arduino = None
    print("MODO SIMULACION activo: Arduino desactivado.")


def enviar_trigger(codigo):
    """Envía comando al Arduino ('H' para START, 'L' para STOP)."""
    if MODO_SIMULACION:
        return core.getTime()
    try:
        t_envio = core.getTime()
        arduino.write(bytes([int(codigo)]))
        return t_envio
    except Exception as e:
        print(f"Error trigger {codigo}: {e}")
        return core.getTime()


# INICIO LSL
lsl_outlet = None
if USAR_LSL:
    lsl_info = StreamInfo('CompasionTask', 'Markers', 1, 0, 'int32', 'compasion_markers')
    lsl_outlet = StreamOutlet(lsl_info)
    print("LSL stream 'CompasionTask' creado.")
else:
    print("LSL desactivado.")


def enviar_lsl(marker):
    """Envía marcador LSL al ANT Neuro."""
    if lsl_outlet is not None:
        lsl_outlet.push_sample([marker])


# 4. VENTANA Y ESTÍMULOS
mon = monitors.Monitor('expMonitor')
mon.setWidth(52)
mon.setDistance(60)
mon.setSizePix([1920, 1200])
mon.save()

win = visual.Window(
    size=[1920, 1200],
    monitor='expMonitor',
    fullscr=True,
    color=[-1, -1, -1],
    units='norm',
    waitBlanking=True,
    allowGUI=False
)
win.mouseVisible = False

fps_real = win.getActualFrameRate(nIdentical=60, nMaxFrames=120, nWarmUpFrames=20)
if fps_real is not None:
    FRAME_DUR = 1.0 / fps_real
    print(f"Monitor: {fps_real:.2f} Hz (duracion frame: {FRAME_DUR*1000:.2f} ms)")
else:
    FRAME_DUR = 1.0 / 60.0
    print(f"No se pudo medir Hz. Asumiendo 60 Hz ({FRAME_DUR*1000:.2f} ms/frame)")

FRAMES_FIJACION = int(round(TIEMPO_FIJACION / FRAME_DUR))
FRAMES_CONTEXTO = int(round(TIEMPO_CONTEXTO / FRAME_DUR))
FRAMES_VIDEO = int(round(TIEMPO_VIDEO / FRAME_DUR))
FRAMES_RATING = int(round(TIEMPO_RATING / FRAME_DUR))

print(f"Frames fijacion: {FRAMES_FIJACION} ({FRAMES_FIJACION * FRAME_DUR:.3f}s)")
print(f"Frames contexto: {FRAMES_CONTEXTO} ({FRAMES_CONTEXTO * FRAME_DUR:.3f}s)")
print(f"Frames video:    {FRAMES_VIDEO} ({FRAMES_VIDEO * FRAME_DUR:.3f}s)")
print(f"Frames rating:   {FRAMES_RATING} ({FRAMES_RATING * FRAME_DUR:.3f}s)")

# 4b. CARGAR CONDICIONES Y SEPARAR ENTRENAMIENTO / EXPERIMENTAL

try:
    df = pd.read_excel(archivo_condiciones)

    # --- Normalizar nombres de columnas ---
    col_map = {}
    for col in df.columns:
        col_limpio = col.strip().lower()
        if col_limpio in ('condición', 'condicion'):
            col_map[col] = 'condicion'
        elif col_limpio in ('nombre_video', 'video'):
            col_map[col] = 'nombre_video'
        elif col_limpio == 'contexto':
            col_map[col] = 'contexto'
        elif col_limpio == 'tipo':
            col_map[col] = 'tipo'
    df = df.rename(columns=col_map)

    columnas_requeridas = ['condicion', 'nombre_video', 'contexto']
    faltantes = [c for c in columnas_requeridas if c not in df.columns]
    if faltantes:
        raise ValueError(f"Columnas faltantes en {archivo_condiciones}: {faltantes}. "
                         f"Columnas encontradas: {list(df.columns)}")

    # --- Separar entrenamiento de experimental ---
    if 'tipo' not in df.columns:
        print("AVISO: Columna 'tipo' no encontrada. Todos los trials se tratan como experimentales.")
        df['tipo'] = 'experimental'

    df['tipo'] = df['tipo'].str.strip().str.lower()
    df_entrenamiento = df[df['tipo'] == 'entrenamiento'].copy()
    df_experimental = df[df['tipo'] != 'entrenamiento'].copy()

    # Entrenamiento: aleatorizar orden
    trials_entrenamiento = df_entrenamiento.sample(frac=1).reset_index(drop=True).to_dict('records')
    print(f"Trials de entrenamiento: {len(trials_entrenamiento)}")

    # Experimental: aleatorizar orden
    trials_experimentales = df_experimental.sample(frac=1).reset_index(drop=True).to_dict('records')
    print(f"Trials experimentales: {len(trials_experimentales)}")

except Exception as e:
    print(f"Error leyendo Excel: {e}")
    if arduino:
        arduino.close()
    win.close()
    core.quit()


# 4c. ESTÍMULOS VISUALES REUTILIZABLES
cruz = visual.TextStim(win, text='+', color='white', height=0.1)
texto_contexto = visual.TextStim(win, text='', color='white', height=0.07, wrapWidth=1.5)
texto_error = visual.TextStim(win, text='Error: Video no encontrado', color='red', height=0.06)

# --- Rating: escala 1-9 con barra de desplazamiento (flechas) ---
texto_pregunta = visual.TextStim(
    win,
    text='Valora el contenido emocional del vídeo que acabas de ver:',
    pos=(0, 0.3), color='white', height=0.06, wrapWidth=1.5
)
slider_visual = visual.Slider(
    win, ticks=(1, 2, 3, 4, 5, 6, 7, 8, 9),
    labels=['1\nMuy negativo', '2', '3', '4', '5\nNeutral', '6', '7', '8', '9\nMuy positivo'],
    granularity=1, style='slider', size=(1.0, 0.05),
    pos=(0, -0.1), color='white', labelHeight=0.045,
    font='Arial'
)

# 5. FUNCIONES DE VIDEO Y CIERRE
def precargar_video(video_path):
    """Pre-carga el video ANTES de necesitarlo."""
    if not os.path.exists(video_path):
        print(f"  VIDEO NO ENCONTRADO: {video_path}")
        return None
    print(f"  Cargando: {os.path.basename(video_path)}")
    try:
        mov = visual.MovieStim(win, filename=video_path, size=VIDEO_SIZE, noAudio=False)
        print(f"  OK (con audio)")
        return mov
    except Exception as e:
        print(f"  Error con audio: {type(e).__name__}: {e}")
    try:
        mov = visual.MovieStim(win, filename=video_path, size=VIDEO_SIZE, noAudio=True)
        print(f"  OK (sin audio, fallback)")
        return mov
    except Exception as e2:
        print(f"  Error sin audio: {type(e2).__name__}: {e2}")
        return None


def limpiar_video(mov):
    """Libera recursos de forma segura."""
    if mov is not None:
        try:
            mov.stop()
        except:
            pass
        try:
            del mov
        except:
            pass
    gc.collect()


def cerrar_todo(datos=None, archivo=None):
    """Cierre limpio."""
    if datos and archivo:
        carpeta = os.path.dirname(archivo)
        nombre = os.path.basename(archivo)
        ruta_parcial = os.path.join(carpeta, f"PARCIAL_{nombre}")
        pd.DataFrame(datos).to_csv(ruta_parcial, index=False)
        print(f"Datos parciales guardados: {ruta_parcial}")
    if arduino:
        try:
            arduino.close()
        except:
            pass
    win.close()
    core.quit()


def ejecutar_trial(trial, num_trial, fase):
    """
    Ejecuta un trial completo (fijación → contexto → video → rating).
    Envía marcadores LSL en cada fase y registra timestamps en CSV.
    """
    nombre_video_raw = str(trial['nombre_video']).strip()
    video_path = os.path.join(RUTA_FIJA, NOMBRE_CARPETA_VIDEOS, nombre_video_raw)
    condicion_str = str(trial['condicion']).strip()

    # FASE 1: FIJACIÓN + PRECARGA VIDEO 
    enviar_lsl(LSL_FIJACION)
    cruz.draw()
    t_fijacion = win.flip()

    mov = precargar_video(video_path)
    video_precargado = (mov is not None)

    for frame in range(1, FRAMES_FIJACION):
        cruz.draw()
        win.flip()
        if event.getKeys(keyList=['escape']):
            limpiar_video(mov)
            cerrar_todo(datos_guardados, nombre_archivo_salida)

    # FASE 2: CONTEXTO 
    enviar_lsl(LSL_CONTEXTO)
    texto_contexto.text = trial['contexto']
    texto_contexto.draw()
    t_contexto = win.flip()

    for frame in range(1, FRAMES_CONTEXTO):
        texto_contexto.draw()
        win.flip()
        if event.getKeys(keyList=['escape']):
            limpiar_video(mov)
            cerrar_todo(datos_guardados, nombre_archivo_salida)

    # FASE 3: VIDEO 
    enviar_lsl(LSL_VIDEO)
    t_video = None
    audio_ok = False

    if video_precargado:
        mov.draw()
        t_video = win.flip()
        audio_ok = not getattr(mov, 'noAudio', False)
        video_terminado = False

        for frame in range(1, FRAMES_VIDEO):
            if not video_terminado:
                terminado = False
                if hasattr(mov, 'isFinished') and mov.isFinished:
                    terminado = True
                if hasattr(mov, 'status') and mov.status == visual.FINISHED:
                    terminado = True
                if terminado:
                    video_terminado = True
                    limpiar_video(mov)
                    mov = None
                else:
                    mov.draw()
            win.flip()
            if event.getKeys(keyList=['escape']):
                if mov is not None:
                    limpiar_video(mov)
                cerrar_todo(datos_guardados, nombre_archivo_salida)

        if mov is not None:
            limpiar_video(mov)
            mov = None

        t_video_fin = core.getTime()
        win.flip()
    else:
        texto_error.draw()
        win.flip()
        core.wait(2)
        t_video_fin = core.getTime()

    # FASE 4: RATING 
    enviar_lsl(LSL_RATING)
    valor = 5
    slider_visual.reset()
    slider_visual.markerPos = valor
    event.clearEvents()

    texto_pregunta.draw()
    slider_visual.draw()
    t_rating = win.flip()

    for frame in range(1, FRAMES_RATING):
        keys = event.getKeys(keyList=['left', 'right', 'escape'])
        if 'escape' in keys:
            cerrar_todo(datos_guardados, nombre_archivo_salida)
        if 'left' in keys:
            valor = max(1, valor - 1)
        if 'right' in keys:
            valor = min(9, valor + 1)

        slider_visual.markerPos = valor
        texto_pregunta.draw()
        slider_visual.draw()
        win.flip()

    # REGISTRO 
    return {
        'Participante': info['Participante'],
        'Sesion': info['Sesion'],
        'Grupo': grupo,
        'Trial': num_trial,
        'Fase': fase,
        'Video': nombre_video_raw,
        'Condicion': condicion_str,
        'Rating': valor,
        'Audio_Video_OK': audio_ok,
        'T_Fijacion': round(t_fijacion, 6),
        'T_Contexto': round(t_contexto, 6),
        'T_Video_Inicio': round(t_video, 6) if t_video else 'NA',
        'T_Video_Fin': round(t_video_fin, 6),
        'T_Rating': round(t_rating, 6),
        'T_Start_Grabacion': round(t_start_grabacion, 6),
        'Frame_Hz': round(fps_real, 2) if fps_real else 'NA'
    }

# 6. INSTRUCCIONES, GRABACIÓN Y BUCLE PRINCIPAL

if not MODO_SIMULACION:
    core.wait(3.0)
    if arduino:
        arduino.flushInput()
    print("Arduino estabilizado.")

#  Pantalla de configuración BNC 
visual.TextStim(
    win,
    text='PASO 1: Conecta los cables BNC al Trigno.\n\n'
         'PASO 2: Configura Start/Stop trigger en Trigno Discover\n'
         'y presiona GRABAR.'
         '\n\n\n\n\n'
         'Presiona la barra espaciadora cuando esté listo.',
    color='white', height=0.06, wrapWidth=1.5
).draw()
win.flip()
event.waitKeys(keyList=['space'])

core.wait(1.0)
if arduino:
    arduino.flushInput()

# INSTRUCCIÓN 1 
visual.TextStim(
    win,
    text='En el siguiente experimento, verás vídeos que muestran situaciones '
         'de la vida real. Cada vídeo irá precedido de un mensaje (que podrás '
         'leer en la pantalla) describiendo el contexto en el que se grabó. '
         'Después de cada vídeo, se te pedirá que respondas algunas preguntas.'
         '\n\n\n\n\n'
         'Pulsa la barra espaciadora para continuar.',
    color='white', height=0.06, wrapWidth=1.5
).draw()
win.flip()
event.waitKeys(keyList=['space'])

# INSTRUCCIÓN 2 
visual.TextStim(
    win,
    text='Los dos primeros ensayos servirán de entrenamiento para que te '
         'familiarices con la estructura de cada ensayo y las preguntas. '
         'Para responder las preguntas, puedes seleccionar un número en una '
         'escala visual usando las flechas del teclado. '
         'Recuerda: ¡no hay respuestas correctas ni incorrectas!'
         '\n\n\n\n\n'
         'Pulsa la barra espaciadora cuando estés lista/o.',
    color='white', height=0.06, wrapWidth=1.5
).draw()
win.flip()
event.waitKeys(keyList=['space'])

# INICIO GRABACIÓN CONTINUA 
if arduino:
    arduino.flushInput()
t_start_grabacion = enviar_trigger(TRIG_START)
core.wait(0.5)
if arduino:
    arduino.flushInput()
print(f"GRABACION INICIADA: {t_start_grabacion:.6f}")

datos_guardados = []
num_trial_global = 0

#  FASE ENTRENAMIENTO 
for i, trial in enumerate(trials_entrenamiento):
    num_trial_global += 1
    datos_trial = ejecutar_trial(trial, num_trial_global, 'entrenamiento')
    datos_guardados.append(datos_trial)

    if arduino:
        try:
            arduino.flushInput()
        except:
            pass
    print(f"Entrenamiento {i+1}/{len(trials_entrenamiento)} OK: {trial['nombre_video']}")

# INSTRUCCIÓN 3 
visual.TextStim(
    win,
    text='¡El entrenamiento ha terminado!'
         '\n\n\n\n\n'
         'Pulsa la barra espaciadora para iniciar el experimento.',
    color='white', height=0.06, wrapWidth=1.5
).draw()
win.flip()
event.waitKeys(keyList=['space'])

# FASE EXPERIMENTAL
for i, trial in enumerate(trials_experimentales):
    num_trial_global += 1
    datos_trial = ejecutar_trial(trial, num_trial_global, 'experimental')
    datos_guardados.append(datos_trial)

    if arduino:
        try:
            arduino.flushInput()
        except:
            pass
    print(f"Trial {i+1}/{len(trials_experimentales)} OK: {trial['nombre_video']}")

# 7. STOP GRABACIÓN Y GUARDAR
t_stop_grabacion = enviar_trigger(TRIG_STOP)
core.wait(0.1)
if arduino:
    arduino.flushInput()
print(f"GRABACION DETENIDA: {t_stop_grabacion:.6f}")

df_salida = pd.DataFrame(datos_guardados)
df_salida['T_Stop_Grabacion'] = round(t_stop_grabacion, 6)
df_salida.to_csv(nombre_archivo_salida, index=False)

print(f"\n{'='*50}")
print("EXPERIMENTO FINALIZADO")
print(f"Grupo: {grupo}")
print(f"Trials totales: {len(datos_guardados)} "
      f"({len(trials_entrenamiento)} entrenamiento + {len(trials_experimentales)} experimentales)")
errores_audio = df_salida[df_salida['Audio_Video_OK'] == False]
if len(errores_audio) > 0:
    print(f"Videos sin audio: {len(errores_audio)}")
    for _, row in errores_audio.iterrows():
        print(f"  - {row['Video']}")
else:
    print("Todos los videos con audio OK.")
print(f"Archivo: {nombre_archivo_salida}")
print(f"{'='*50}\n")

# INSTRUCCIÓN 4 
visual.TextStim(
    win,
    text='Gracias por participar y colaborar con la ciencia.',
    color='white', height=0.07
).draw()
win.flip()
core.wait(3)

if arduino:
    try:
        arduino.close()
    except:
        pass

win.close()
core.quit()
