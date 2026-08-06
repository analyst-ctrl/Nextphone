============================================
 TRANSCRIPCIÓN DE AUDIOS A TEXTO EN ESPAÑOL
 (100% local, gratis y privado)
============================================

¿QUÉ ES?
--------
Convierte grabaciones (MP3, WAV, M4A, OGG, etc.) en texto transcrito en
español usando Whisper (faster-whisper). Todo se procesa en TU PC:
los audios de llamadas con datos de clientes NUNCA salen de tu equipo.

REQUISITOS (ya instalados en este equipo)
----------------------------------------
- Python 3.13
- faster-whisper + ctranslate2 + PyAV  (instalados)
- edge-tts (solo para generar audios de prueba)

CÓMO USARLO
-----------
Transcribir UN audio:
    python transcribir.py llamada.mp3

Transcribir TODOS los audios de una carpeta:
    python transcribir.py carpeta_de_audios/

Opciones útiles:
    --modelo base | small        Cambia el modelo (ver tabla abajo)
    --con_tiempos                Marca [mm:ss] al inicio de cada párrafo
    --formato srt                Genera subtítulos en vez de texto
    --idioma en | auto           Otro idioma (default: es)
    --beam 1 | 3 | 5             Menos/más precisión (default: 3; usa 5 para
                                 máxima precisión si puedes esperar más)

Ejemplo para llamadas de cobro (recomendado):
    python transcribir.py "Llamadas Mayo/" --modelo base --con_tiempos

El resultado se guarda junto al audio: llamada.mp3 -> llamada.txt

TIEMPOS DE ESPERA (según tu PC: i7-4650U, 8 GB RAM)
---------------------------------------------------
  Modelo   Calidad      Velocidad                 Llamada de 30 min
  -------  ----------   ------------------------  -----------------
  base     Buena        5x más rápido que el audio   ~6 min
  small    Muy buena    1.5x más rápido              ~20 min
  tiny     Regular      muy rápido                   ~3 min
  medium   Excelente    MUY lento en tu CPU          ~1-2 horas (no recomendado)

RECOMENDACIÓN
-------------
- Para el día a día (muchas llamadas):  --modelo base
- Para casos difíciles (audio malo, tecnicismos):  --modelo small

PRIMERA EJECUCIÓN
-----------------
La primera vez con cada modelo descarga el archivo de IA (~140 MB base,
~460 MB small) desde internet. Después funciona sin conexión.

NOTAS / SOLUCIÓN DE PROBLEMAS
-----------------------------
- Nombres propios/marcas ("Nextphone") pueden salir mal; es normal en
  cualquier transcritor. Se pueden corregir a mano o con un reemplazo.
- Si la consola muestra caracteres raros, ejecuta:
      python -X utf8 transcribir.py audio.mp3
- El aviso de "symlinks" de HuggingFace es inofensivo, ignorarlo.
- Se puede mejorar la velocidad con el modelo tiny si es urgente,
  pero la calidad baja.

PROBAR CON UN AUDIO DE EJEMPLO
------------------------------
    python generar_prueba.py          (crea prueba_llamada.mp3)
    python transcribir.py prueba_llamada.mp3 --modelo base --con_tiempos
