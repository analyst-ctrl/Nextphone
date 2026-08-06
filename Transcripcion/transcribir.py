#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Transcripción de audios (MP3, WAV, M4A, OGG...) a texto en español.
100% LOCAL: tus grabaciones nunca salen de este PC (privacidad).

Primera vez: descarga el modelo (~460 MB en el modelo "small").
Siguientes veces: usa el modelo ya descargado (sin internet).

Uso:
    python transcribir.py audio.mp3
    python transcribir.py carpeta_de_audios/            # transcribe todos los audios
    python transcribir.py audio.mp3 --modelo base       # más rápido, menos preciso
    python transcribir.py audio.mp3 --con_tiempos       # marca [mm:ss] por párrafo
    python transcribir.py audio.mp3 --formato srt       # subtítulos

Opciones:
    --modelo   small (default) | base | medium | large-v3
    --idioma   es (default) | en | auto
    --con_tiempos   agrega [mm:ss] al inicio de cada párrafo
    --formato  txt (default) | srt
    --salida   carpeta donde guardar (default: misma carpeta del audio)
"""

import argparse
import os
import sys
import time
from pathlib import Path

# La consola de Windows (cp1252) no imprime emojis/acentos: forzamos UTF-8
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Silencia el aviso de symlinks de HuggingFace en Windows (inofensivo)
os.environ.setdefault("HF_HUB_DISABLE_SYMLINKS_WARNING", "1")

# Extensiones de audio soportadas (faster-whisper decodifica sin ffmpeg)
EXT_AUDIO = {".mp3", ".wav", ".m4a", ".ogg", ".flac", ".aac", ".wma", ".opus", ".mp4", ".webm"}


def formatear_tiempo(seg: float) -> str:
    """Convierte segundos a mm:ss."""
    seg = max(0, int(seg))
    m, s = divmod(seg, 60)
    return f"{m:02d}:{s:02d}"


def juntar_parrafos(segments, pausa_max: float = 3.0, palabras_max: int = 90):
    """
    Une segmentos en párrafos legibles. Corta por silencio largo o cuando el
    párrafo alcanza palabras_max (el filtro VAD elimina los silencios, así que
    este tope evita párrafos gigantes en llamadas largas).
    Devuelve lista de (inicio_seg, texto).
    """
    parrafos = []
    inicio_actual = None
    texto_actual = []
    palabras_actual = 0
    fin_anterior = None

    for seg in segments:
        if inicio_actual is None:
            inicio_actual = seg.start
        pausa_larga = fin_anterior is not None and (seg.start - fin_anterior) > pausa_max
        if (pausa_larga and palabras_actual >= 10) or palabras_actual >= palabras_max:
            parrafos.append((inicio_actual, " ".join(texto_actual).strip()))
            inicio_actual = seg.start
            texto_actual = []
            palabras_actual = 0
        texto_actual.append(seg.text.strip())
        palabras_actual += len(seg.text.split())
        fin_anterior = seg.end

    if texto_actual:
        parrafos.append((inicio_actual, " ".join(texto_actual).strip()))
    return parrafos


def transcribir_archivo(ruta: Path, modelo, idioma: str, con_tiempos: bool, formato: str, salida: Path, beam: int):
    print(f"\n🎧  {ruta.name}")
    print(f"    ⏳ Transcribiendo... (esto puede tardar según la duración del audio)")

    t0 = time.time()
    segments, info = modelo.transcribe(
        str(ruta),
        language=None if idioma == "auto" else idioma,
        beam_size=beam,
        vad_filter=True,          # ignora silencios/ruido de fondo
        vad_parameters=dict(min_silence_duration_ms=500),
    )

    info_idioma = info.language if idioma == "auto" else idioma
    print(f"    🌐 Idioma detectado: {info_idioma} | Duración: {formatear_tiempo(info.duration)}")

    # Primera pasada: recolectar segmentos (la transcripción ocurre aquí)
    segmentos = list(segments)

    if formato == "srt":
        lineas = []
        for i, seg in enumerate(segmentos, 1):
            inicio = seg.start
            fin = seg.end
            t_ini = f"{int(inicio // 3600):02d}:{int(inicio % 3600 // 60):02d}:{int(inicio % 60):02d},{int((inicio % 1) * 1000):03d}"
            t_fin = f"{int(fin // 3600):02d}:{int(fin % 3600 // 60):02d}:{int(fin % 60):02d},{int((fin % 1) * 1000):03d}"
            lineas.append(f"{i}\n{t_ini} --> {t_fin}\n{seg.text.strip()}\n")
        contenido = "\n".join(lineas)
    else:
        parrafos = juntar_parrafos(segmentos)
        if con_tiempos:
            contenido = "\n\n".join(f"[{formatear_tiempo(ini)}]\n{texto}" for ini, texto in parrafos)
        else:
            contenido = "\n\n".join(texto for _, texto in parrafos)

    # Guardar
    salida.mkdir(parents=True, exist_ok=True)
    archivo_salida = salida / f"{ruta.stem}.{formato}"
    archivo_salida.write_text(contenido, encoding="utf-8")

    t_total = time.time() - t0
    print(f"    ✅ Guardado en: {archivo_salida}")
    print(f"    ⏱️  Tiempo de transcripción: {t_total:.0f}s "
          f"({(info.duration / t_total) if t_total else 0:.1f}x tiempo real)")
    return archivo_salida


def main():
    parser = argparse.ArgumentParser(
        description="Transcribe audios a texto en español (100% local).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__.split("Opciones:")[0],
    )
    parser.add_argument("entrada", help="Archivo de audio o carpeta con audios")
    parser.add_argument("--modelo", default="small",
                        choices=["tiny", "base", "small", "medium", "large-v3"],
                        help="Tamaño del modelo (default: small)")
    parser.add_argument("--beam", type=int, default=3,
                        help="Precisión vs velocidad (default: 3; usa 5 si quieres más precisión y puedes esperar)")
    parser.add_argument("--idioma", default="es", help="Idioma: es | en | auto (default: es)")
    parser.add_argument("--con_tiempos", action="store_true", help="Agrega [mm:ss] a cada párrafo")
    parser.add_argument("--formato", default="txt", choices=["txt", "srt"], help="Formato de salida")
    parser.add_argument("--salida", default=None, help="Carpeta de salida (default: junto al audio)")
    args = parser.parse_args()

    entrada = Path(args.entrada)
    if not entrada.exists():
        print(f"❌ No existe: {entrada}")
        sys.exit(1)

    if entrada.is_file():
        archivos = [entrada]
    else:
        archivos = sorted(p for p in entrada.rglob("*") if p.suffix.lower() in EXT_AUDIO)
        if not archivos:
            print(f"❌ No hay audios (.mp3, .wav, .m4a, ...) en {entrada}")
            sys.exit(1)
        print(f"📁 {len(archivos)} audios encontrados en {entrada}")

    if args.idioma != "auto" and (len(args.idioma) != 2 or not args.idioma.isalpha()):
        print("❌ --idioma debe ser un código de 2 letras (es, en, pt...) o 'auto'")
        sys.exit(1)

    print(f"🧠 Cargando modelo '{args.modelo}' (la primera vez lo descarga, puede tardar)...")
    t_carga = time.time()
    try:
        # Importamos aquí para no frenar el --help
        from faster_whisper import WhisperModel
        modelo = WhisperModel(args.modelo, device="cpu", compute_type="int8")
    except ModuleNotFoundError:
        print("❌ Falta faster-whisper. Instálalo con:  python -m pip install faster-whisper")
        sys.exit(1)
    except Exception as e:
        print(f"❌ No se pudo cargar el modelo '{args.modelo}': {e}")
        print("   La primera vez se descarga de internet; revisa tu conexión e inténtalo de nuevo.")
        sys.exit(1)
    print(f"   Modelo listo en {time.time() - t_carga:.0f}s")

    for archivo in archivos:
        try:
            salida = Path(args.salida) if args.salida else archivo.parent
            transcribir_archivo(archivo, modelo, args.idioma, args.con_tiempos, args.formato, salida, args.beam)
        except Exception as e:
            print(f"❌ Error en {archivo.name}: {e}")

    print("\n✨ ¡Listo! Revisa los archivos .txt generados.")


if __name__ == "__main__":
    main()
