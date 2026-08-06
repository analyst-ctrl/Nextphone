#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Genera un MP3 de prueba en español (voz de Microsoft Edge) para probar transcribir.py."""
import asyncio
import sys
import edge_tts
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TEXTO = (
    "Buenas tardes, habla con el departamento de cobros de Nextphone. "
    "Mi nombre es María, ¿con quién tengo el gusto? "
    "Hola María, soy el señor Pérez. Lo llamo porque tengo una consulta sobre mi factura de este mes. "
    "Claro, con gusto le ayudo. Veo en el sistema que su factura tiene un saldo pendiente de cuarenta y cinco dólares con veinticinco centavos. "
    "¿Le gustaría realizar el pago ahora con tarjeta de crédito o débito? "
    "Prefiero pagar en efectivo en la sucursal, ¿me puede dar la dirección? "
    "Por supuesto, estamos en vía España, frente al parque, de lunes a viernes de ocho de la mañana a cinco de la tarde. "
    "Perfecto, muchas gracias por la información. "
    "A usted señor Pérez, que tenga un excelente día."
)


async def main():
    salida = Path(__file__).parent / "prueba_llamada.mp3"
    voz = "es-MX-DaliaNeural"
    comunicador = edge_tts.Communicate(TEXTO, voz, rate="-5%")
    await comunicador.save(str(salida))
    print(f"✅ MP3 de prueba generado: {salida} ({salida.stat().st_size / 1024:.0f} KB)")


if __name__ == "__main__":
    asyncio.run(main())
