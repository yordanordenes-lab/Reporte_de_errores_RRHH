#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
AUDITORÍA MENSUAL DE LIQUIDACIONES VS. ASISTENCIAS
================================================================================
Herramienta experimental para automatizar la revisión cruzada entre
registros de asistencia y liquidaciones de sueldo (PDF).

⚠️  USO Y ÁMBITO:
    - Proyecto de código abierto con fines educativos y de mejora de procesos
      administrativos de RRHH.
    - Diseñado para ejecución LOCAL en entornos controlados.
    - Los datos de prueba incluidos son 100% ficticios.
    - El autor no se hace responsable del uso indebido o del procesamiento de
      datos reales sin las salvaguardas legales correspondientes (Ley N° 19.628
      sobre Protección de la Vida Privada en Chile, GDPR, etc.).

REQUISITOS:
    pip install pdfplumber pandas openpyxl

================================================================================
"""

import os
import sys
import re
import glob
import math
import argparse
import logging
import hashlib
from pathlib import Path
from typing import List, Dict, Optional, Tuple, Iterator, Union
from dataclasses import dataclass, field, asdict
from abc import ABC, abstractmethod
from collections import defaultdict

import pdfplumber
import pandas as pd

# ==============================================================================
# CONFIGURACIÓN GLOBAL
# ==============================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger("auditoria_rrhh")

# Límites de seguridad para ejecución local
MAX_PDF_SIZE_MB = 50
MAX_PDF_PAGES_ESTIMATE = 2000  # umbral para advertencia, no bloqueo


# ==============================================================================
# EXCEPCIONES PERSONALIZADAS
# ==============================================================================

class AuditoriaError(Exception):
    """Error base del sistema de auditoría."""
    pass


class PDFInvalidoError(AuditoriaError):
    """El archivo no parece un PDF válido o está corrupto."""
    pass


class RutaNoPermitidaError(AuditoriaError):
    """Intento de acceso a rutas fuera del directorio permitido."""
    pass


# ==============================================================================
# MODELOS DE DATOS (Dataclasses)
# ==============================================================================

@dataclass
class RegistroTrabajador:
    """Representa un trabajador extraído de un PDF (asistencia o liquidación)."""
    rut: str
    nombre: str = "N/A"
    dias_trabajados: Optional[float] = None
    atrasos: Optional[float] = 0.0
    cargo: Optional[str] = None
    # Campos exclusivos de liquidación
    sueldo_base_30_dias: Optional[float] = None
    sueldo_base_a_pago: Optional[float] = None
    tipo_contrato: Optional[str] = None
    afp: Optional[str] = None
    salud: Optional[str] = None
    # Metadatos
    tipo_documento: str = ""          # "asistencia" | "liquidacion"
    pagina: int = 0
    archivo_origen: str = ""
    hash_archivo: str = ""              # para trazabilidad

    def to_dict(self) -> Dict:
        return asdict(self)


@dataclass
class Hallazgo:
    """Representa una discrepancia detectada entre asistencia y liquidación."""
    rut: str
    nombre: str
    severidad: str                      # CRÍTICO | ALTO | MEDIO | BAJO
    mensaje: str
    dias_asistencia: Optional[float] = None
    dias_liquidacion: Optional[float] = None
    diferencia_dias: Optional[float] = None
    atrasos_asistencia: Optional[float] = None
    atrasos_liquidacion: Optional[float] = None
    diferencia_atrasos: Optional[float] = None
    sueldo_pagado: Optional[float] = None
    sueldo_base_30: Optional[float] = None
    sueldo_base_real: Optional[float] = None
    diferencia_sueldo: Optional[float] = None


# ==============================================================================
# UTILIDADES
# ==============================================================================

def normalizar_texto(texto: str) -> str:
    """Limpia y normaliza texto para comparaciones robustas."""
    if not texto:
        return ""
    texto = str(texto).strip().lower()
    tildes = str.maketrans("áéíóúÁÉÍÓÚñÑ", "aeiouAEIOUnn")
    texto = texto.translate(tildes)
    texto = re.sub(r"\s+", " ", texto)
    return texto


def limpiar_rut(rut: str) -> Optional[str]:
    """
    Normaliza un RUT chileno al formato 12345678-9.
    Retorna None si el formato es inválido.
    """
    if not rut:
        return None
    rut = str(rut).strip().upper().replace(".", "").replace(" ", "")
    # Validar formato básico: 7-8 dígitos, guion, dígito verificador (0-9 o K)
    if not re.match(r"^\d{7,8}-[\dK]$", rut):
        # Intentar auto-corregir si falta guion
        if "-" not in rut and len(rut) >= 2:
            rut = rut[:-1] + "-" + rut[-1]
        if not re.match(r"^\d{7,8}-[\dK]$", rut):
            logger.warning(f"RUT con formato inválido descartado: {rut}")
            return None
    return rut


def limpiar_monto(valor) -> Optional[float]:
    """Convierte un monto en string a float, soportando formatos chileno y USA."""
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip()
    s = s.replace("$", "").replace(" ", "").replace(".-", "").replace("-", "")
    if not s:
        return None

    # Detectar formato por posición de separadores
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            # 1.234.567,89 → chileno/europeo
            s = s.replace(".", "").replace(",", ".")
        else:
            # 1,234,567.89 → USA
            s = s.replace(",", "")
    elif "," in s:
        partes = s.split(",")
        if len(partes[-1]) == 2 and len(partes) == 2:
            # 1234,56 → decimal con coma
            s = s.replace(",", ".")
        else:
            # 1.234.567 → separador de miles (sin decimales)
            s = s.replace(",", "")
    else:
        # Sin comas, puede tener puntos como miles o ser entero
        if s.count(".") > 1:
            s = s.replace(".", "")
        # si tiene un solo punto y 3 decimales no es probable, asumimos decimal

    try:
        return float(s)
    except ValueError:
        logger.warning(f"No se pudo convertir monto: '{valor}'")
        return None


def limpiar_dias(valor) -> Optional[float]:
    """Convierte días/atrasos a float."""
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip()
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return None


def formatear_moneda(valor: Optional[float]) -> str:
    """Formatea un valor como moneda chilena ($1.234.567)."""
    if valor is None:
        return ""
    entero = int(valor) if valor == int(valor) else math.ceil(valor)
    return f"${entero:,.0f}".replace(",", ".")


def redondear_sueldo(valor: float, metodo: str = "up") -> int:
    """
    Redondea según método:
      - 'up'     : techo (default, común en normativa chilena)
      - 'down'   : piso
      - 'nearest': al más cercano, .5 hacia arriba (NO banker's rounding)
    """
    if valor is None:
        return 0
    if metodo == "up":
        return math.ceil(valor)
    elif metodo == "down":
        return math.floor(valor)
    elif metodo == "nearest":
        # Evitar round() de Python que usa banker's rounding
        return int(math.floor(valor + 0.5))
    else:
        return math.ceil(valor)


def validar_ruta_segura(ruta: str, base: Optional[str] = None) -> Path:
    """
    Resuelve la ruta y verifica que no escape del directorio base.
    Previene path traversal en entornos locales compartidos.
    """
    ruta_resuelta = Path(ruta).resolve()
    if base:
        base_resuelta = Path(base).resolve()
        if not str(ruta_resuelta).startswith(str(base_resuelta)):
            raise RutaNoPermitidaError(
                f"Acceso denegado: {ruta_resuelta} está fuera de {base_resuelta}"
            )
    return ruta_resuelta


def verificar_pdf_sano(ruta: Path) -> None:
    """Valida que el archivo sea un PDF real y no exceda tamaño límite."""
    if not ruta.exists():
        raise PDFInvalidoError(f"Archivo no encontrado: {ruta}")
    tamano_mb = ruta.stat().st_size / (1024 * 1024)
    if tamano_mb > MAX_PDF_SIZE_MB:
        raise PDFInvalidoError(
            f"PDF demasiado grande ({tamano_mb:.1f} MB). Límite: {MAX_PDF_SIZE_MB} MB"
        )
    with open(ruta, "rb") as f:
        header = f.read(5)
        if header != b"%PDF-":
            raise PDFInvalidoError(f"El archivo no parece un PDF válido: {ruta}")


def calcular_hash_archivo(ruta: Path) -> str:
    """SHA-256 del archivo para trazabilidad."""
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()[:16]


# ==============================================================================
# ESTRATEGIA DE EXTRACCIÓN (Strategy Pattern)
# ==============================================================================

class ExtractorStrategy(ABC):
    """Interfaz base para extractores de PDF."""

    @abstractmethod
    def extraer(self, pagina, num_pagina: int, archivo: str, hash_archivo: str) -> Optional[RegistroTrabajador]:
        """Extrae un RegistroTrabajador de una página de PDF."""
        pass

    @property
    @abstractmethod
    def tipo_documento(self) -> str:
        pass


class AsistenciaExtractor(ExtractorStrategy):
    """Extractor especializado para registros de asistencia."""

    CAMPOS_TEXTO = {
        "nombre": ["nombre trabajador", "trabajador", "nombre"],
        "rut": ["rut"],
        "dias_trabajados": ["dias trabajados", "días trabajados", "dias trabajado"],
        "cargo": ["cargo"],
        "atrasos": ["atrasos", "tardanzas"],
    }

    @property
    def tipo_documento(self) -> str:
        return "asistencia"

    def extraer(self, pagina, num_pagina: int, archivo: str, hash_archivo: str) -> Optional[RegistroTrabajador]:
        texto = pagina.extract_text() or ""
        tablas = pagina.extract_tables() or []
        datos: Dict[str, any] = {}

        # 1. Extraer desde tablas
        for tabla in tablas:
            for fila in tabla:
                if len(fila) >= 2:
                    etiqueta = normalizar_texto(str(fila[0]))
                    valor = str(fila[1]).strip()
                    for campo, variants in self.CAMPOS_TEXTO.items():
                        if any(v in etiqueta for v in variants) and campo not in datos:
                            datos[campo] = self._parsear(campo, valor)

        # 2. Fallback por líneas (corregido: sin doble incremento)
        if not datos.get("rut") or datos.get("dias_trabajados") is None:
            self._extraer_lineas(texto, datos)

        # 3. Validar mínimo indispensable
        rut = datos.get("rut")
        dias = datos.get("dias_trabajados")
        if not rut or dias is None:
            return None

        return RegistroTrabajador(
            rut=rut,
            nombre=datos.get("nombre", "N/A"),
            dias_trabajados=dias,
            atrasos=datos.get("atrasos", 0.0) or 0.0,
            cargo=datos.get("cargo"),
            tipo_documento=self.tipo_documento,
            pagina=num_pagina,
            archivo_origen=archivo,
            hash_archivo=hash_archivo,
        )

    def _extraer_lineas(self, texto: str, datos: Dict):
        """Recorre línea a línea buscando etiqueta + valor en la siguiente línea."""
        lineas = [l.strip() for l in texto.split("\n") if l.strip()]
        i = 0
        while i < len(lineas):
            linea_norm = normalizar_texto(lineas[i])
            matched = False
            for campo, variants in self.CAMPOS_TEXTO.items():
                if campo in datos:
                    continue
                if any(v in linea_norm for v in variants):
                    if i + 1 < len(lineas):
                        datos[campo] = self._parsear(campo, lineas[i + 1])
                        i += 2  # saltamos etiqueta + valor
                        matched = True
                        break
            if not matched:
                i += 1

    def _parsear(self, campo: str, valor: str):
        if campo in ("dias_trabajados", "atrasos"):
            return limpiar_dias(valor)
        elif campo == "rut":
            return limpiar_rut(valor)
        return valor.strip()


class LiquidacionExtractor(ExtractorStrategy):
    """Extractor especializado para liquidaciones de sueldo."""

    CAMPOS_TEXTO = {
        "nombre": ["nombre trabajador", "trabajador", "nombre"],
        "rut": ["rut"],
        "dias_trabajados": ["dias trabajados", "días trabajados"],
        "cargo": ["cargo"],
        "atrasos": ["atrasos"],
        "tipo_contrato": ["tipo contrato"],
        "afp": ["afp"],
        "salud": ["salud"],
    }

    REGEX = {
        "sueldo_base_30_dias": re.compile(
            r"SUELDO\s+BASE\s+A\s+30\s+DIAS\s*\n?\s*\$?([\d\.\,]+)", re.IGNORECASE
        ),
        "sueldo_base_a_pago": re.compile(
            r"SUELDO\s+BASE\s+A\s+PAGO\s*\n?\s*\$?([\d\.\,]+)", re.IGNORECASE
        ),
        "atrasos": re.compile(r"ATRASOS\s*\(\s*(\d+)\s*\)", re.IGNORECASE),
    }

    @property
    def tipo_documento(self) -> str:
        return "liquidacion"

    def extraer(self, pagina, num_pagina: int, archivo: str, hash_archivo: str) -> Optional[RegistroTrabajador]:
        texto = pagina.extract_text() or ""
        tablas = pagina.extract_tables() or []
        datos: Dict[str, any] = {}

        # 1. Tablas
        for tabla in tablas:
            for fila in tabla:
                if len(fila) >= 2:
                    etiqueta = normalizar_texto(str(fila[0]))
                    valor = str(fila[1]).strip()
                    for campo, variants in self.CAMPOS_TEXTO.items():
                        if any(v in etiqueta for v in variants) and campo not in datos:
                            datos[campo] = self._parsear(campo, valor)

        # 2. Fallback líneas
        if not datos.get("rut") or datos.get("dias_trabajados") is None:
            self._extraer_lineas(texto, datos)

        # 3. Regex específicos
        for campo, patron in self.REGEX.items():
            if campo not in datos:
                match = patron.search(texto)
                if match:
                    datos[campo] = self._parsear(campo, match.group(1))

        rut = datos.get("rut")
        dias = datos.get("dias_trabajados")
        if not rut or dias is None:
            return None

        return RegistroTrabajador(
            rut=rut,
            nombre=datos.get("nombre", "N/A"),
            dias_trabajados=dias,
            atrasos=datos.get("atrasos", 0.0) or 0.0,
            cargo=datos.get("cargo"),
            sueldo_base_30_dias=datos.get("sueldo_base_30_dias"),
            sueldo_base_a_pago=datos.get("sueldo_base_a_pago"),
            tipo_contrato=datos.get("tipo_contrato"),
            afp=datos.get("afp"),
            salud=datos.get("salud"),
            tipo_documento=self.tipo_documento,
            pagina=num_pagina,
            archivo_origen=archivo,
            hash_archivo=hash_archivo,
        )

    def _extraer_lineas(self, texto: str, datos: Dict):
        lineas = [l.strip() for l in texto.split("\n") if l.strip()]
        i = 0
        while i < len(lineas):
            linea_norm = normalizar_texto(lineas[i])
            matched = False
            for campo, variants in self.CAMPOS_TEXTO.items():
                if campo in datos:
                    continue
                if any(v in linea_norm for v in variants):
                    if i + 1 < len(lineas):
                        datos[campo] = self._parsear(campo, lineas[i + 1])
                        i += 2
                        matched = True
                        break
            if not matched:
                i += 1

    def _parsear(self, campo: str, valor: str):
        if campo in ("dias_trabajados", "atrasos"):
            return limpiar_dias(valor)
        elif campo in ("sueldo_base_30_dias", "sueldo_base_a_pago"):
            return limpiar_monto(valor)
        elif campo == "rut":
            return limpiar_rut(valor)
        return valor.strip()


# ==============================================================================
# MOTOR DE EXTRACCIÓN POR LOTES (Streaming)
# ==============================================================================

class MotorExtraccion:
    """
    Orquesta la extracción de múltiples PDFs usando estrategias.
    Soporta streaming para no cargar todo en memoria de golpe.
    """

    def __init__(self, estrategia: ExtractorStrategy):
        self.estrategia = estrategia

    def extraer_de_pdf(self, ruta_pdf: Path) -> Iterator[RegistroTrabajador]:
        """Generador: produce registros uno a uno, liberando memoria de páginas."""
        verificar_pdf_sano(ruta_pdf)
        hash_archivo = calcular_hash_archivo(ruta_pdf)
        nombre_archivo = ruta_pdf.name
        logger.info(f"📄 Procesando {self.estrategia.tipo_documento}: {nombre_archivo}")

        try:
            with pdfplumber.open(str(ruta_pdf)) as pdf:
                total_paginas = len(pdf.pages)
                if total_paginas > MAX_PDF_PAGES_ESTIMATE:
                    logger.warning(
                        f"   ⚠️  PDF grande detectado ({total_paginas} páginas). "
                        f"Considere dividir el archivo para mejor rendimiento."
                    )
                for num_pagina, pagina in enumerate(pdf.pages, 1):
                    try:
                        registro = self.estrategia.extraer(
                            pagina, num_pagina, nombre_archivo, hash_archivo
                        )
                        if registro:
                            yield registro
                    except Exception as e:
                        logger.warning(f"   ⚠️  Error página {num_pagina}: {e}")
                        continue
        except Exception as e:
            logger.error(f"No se pudo abrir {ruta_pdf}: {e}")

    def extraer_de_pdfs(self, rutas: List[Path]) -> Iterator[RegistroTrabajador]:
        """Extrae de múltiples PDFs en secuencia."""
        for ruta in rutas:
            yield from self.extraer_de_pdf(ruta)


# ==============================================================================
# AUDITORÍA MENSUAL
# ==============================================================================

class AuditoriaMensual:
    """
    Compara asistencias vs liquidaciones por RUT.
    Diseñado para manejar volúmenes medianos (hasta ~5.000 trabajadores en RAM).
    Para volúmenes mayores, usar la variante de streaming con base de datos.
    """

    def __init__(self, metodo_redondeo: str = "up"):
        self.metodo_redondeo = metodo_redondeo
        self.hallazgos: List[Hallazgo] = []
        self.duplicados: List[Dict] = []

    def auditar(
        self,
        asistencias: List[RegistroTrabajador],
        liquidaciones: List[RegistroTrabajador],
    ) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        Ejecuta la auditoría completa.
        Retorna: (DataFrame hallazgos, DataFrame duplicados)
        """
        asis_por_rut, dup_asis = self._agrupar_con_duplicados(asistencias)
        liq_por_rut, dup_liq = self._agrupar_con_duplicados(liquidaciones)
        self.duplicados = dup_asis + dup_liq

        logger.info("\n" + "=" * 60)
        logger.info("🔍 AUDITORÍA MENSUAL - Comparación Asistencia vs Liquidación")
        logger.info("=" * 60)
        logger.info(f"   • Asistencias únicas: {len(asis_por_rut)}")
        logger.info(f"   • Liquidaciones únicas: {len(liq_por_rut)}")
        logger.info(f"   • Duplicados detectados: {len(self.duplicados)}")

        ruts_asis = set(asis_por_rut.keys())
        ruts_liq = set(liq_por_rut.keys())
        todos = sorted(ruts_asis | ruts_liq)
        logger.info(f"   • Trabajadores a auditar: {len(todos)}\n")

        for rut in todos:
            asis = asis_por_rut.get(rut)
            liq = liq_por_rut.get(rut)
            nombre = asis.nombre if asis else (liq.nombre if liq else "N/A")

            if not asis:
                self._registrar_critico(rut, nombre, "CRÍTICO",
                    f"Liquidación sin registro de asistencia.",
                    liq=liq)
                continue
            if not liq:
                self._registrar_critico(rut, nombre, "CRÍTICO",
                    f"Asistencia sin liquidación.",
                    asis=asis)
                continue

            self._comparar_registros(rut, nombre, asis, liq)

        df_hallazgos = pd.DataFrame([h.__dict__ for h in self.hallazgos])
        df_duplicados = pd.DataFrame(self.duplicados) if self.duplicados else pd.DataFrame()

        logger.info("\n📊 RESULTADOS:")
        if df_hallazgos.empty:
            logger.info("   ✅ Todos los trabajadores coinciden. ¡Todo OK!")
        else:
            logger.info(f"   ⚠️  Hallazgos: {len(df_hallazgos)}")
            if "severidad" in df_hallazgos.columns:
                logger.info("\n" + str(df_hallazgos["severidad"].value_counts()))

        return df_hallazgos, df_duplicados

    def _agrupar_con_duplicados(
        self, registros: List[RegistroTrabajador]
    ) -> Tuple[Dict[str, RegistroTrabajador], List[Dict]]:
        """Agrupa por RUT y reporta duplicados en una lista separada."""
        por_rut: Dict[str, RegistroTrabajador] = {}
        duplicados: List[Dict] = []
        for reg in registros:
            rut = reg.rut
            if rut in por_rut:
                duplicados.append({
                    "rut": rut,
                    "nombre": reg.nombre,
                    "tipo": reg.tipo_documento,
                    "archivo": reg.archivo_origen,
                    "pagina": reg.pagina,
                    "motivo": "RUT duplicado dentro del mismo tipo de documento"
                })
                logger.warning(
                    f"   ⚠️  RUT duplicado en {reg.tipo_documento}: {rut} "
                    f"(pág. {reg.pagina} de {reg.archivo_origen})"
                )
            else:
                por_rut[rut] = reg
        return por_rut, duplicados

    def _comparar_registros(self, rut: str, nombre: str, asis: RegistroTrabajador, liq: RegistroTrabajador):
        dias_asis = asis.dias_trabajados
        dias_liq = liq.dias_trabajados
        atrasos_asis = asis.atrasos or 0.0
        atrasos_liq = liq.atrasos or 0.0
        sueldo_pagado = liq.sueldo_base_a_pago
        sueldo_base_30 = liq.sueldo_base_30_dias

        # Cálculo sueldo real
        sueldo_base_real = None
        diferencia_sueldo = None
        if sueldo_base_30 is not None and dias_asis is not None:
            sueldo_base_real = redondear_sueldo(
                (sueldo_base_30 / 30.0) * dias_asis, self.metodo_redondeo
            )
        if sueldo_pagado is not None and sueldo_base_real is not None:
            diferencia_sueldo = sueldo_base_real - sueldo_pagado

        hay_error_dias = False
        msg_dias = ""
        if dias_asis is None or dias_liq is None:
            hay_error_dias = True
            msg_dias = f"Días incompletos: Asistencia={dias_asis}, Liquidación={dias_liq}. "
        elif abs(dias_liq - dias_asis) > 0.1:
            hay_error_dias = True
            msg_dias = (
                f"DÍAS DIFIEREN: Liq={dias_liq}, Asis={dias_asis} "
                f"(Δ{dias_liq - dias_asis:+.1f}). "
            )

        hay_error_atrasos = False
        msg_atrasos = ""
        if atrasos_asis != atrasos_liq:
            hay_error_atrasos = True
            msg_atrasos = (
                f"ATRASOS DIFIEREN: Asis={atrasos_asis}, Liq={atrasos_liq} "
                f"(Δ{atrasos_liq - atrasos_asis:+.0f}). "
            )

        if hay_error_dias or hay_error_atrasos:
            severidad = "ALTO" if hay_error_dias else "MEDIO"
            mensaje = msg_dias + msg_atrasos + "Revisar cálculo de sueldo."
            self.hallazgos.append(Hallazgo(
                rut=rut,
                nombre=nombre,
                severidad=severidad,
                mensaje=mensaje,
                dias_asistencia=dias_asis,
                dias_liquidacion=dias_liq,
                diferencia_dias=(dias_liq - dias_asis) if (dias_asis is not None and dias_liq is not None) else None,
                atrasos_asistencia=atrasos_asis,
                atrasos_liquidacion=atrasos_liq,
                diferencia_atrasos=atrasos_liq - atrasos_asis,
                sueldo_pagado=sueldo_pagado,
                sueldo_base_30=sueldo_base_30,
                sueldo_base_real=sueldo_base_real,
                diferencia_sueldo=diferencia_sueldo,
            ))
        else:
            logger.info(f"   ✓ {nombre:30s} | RUT {rut} | {dias_asis} días | {atrasos_asis} atrasos | OK")

    def _registrar_critico(self, rut: str, nombre: str, severidad: str, mensaje: str,
                           asis: Optional[RegistroTrabajador] = None,
                           liq: Optional[RegistroTrabajador] = None):
        self.hallazgos.append(Hallazgo(
            rut=rut,
            nombre=nombre,
            severidad=severidad,
            mensaje=mensaje,
            dias_asistencia=asis.dias_trabajados if asis else None,
            dias_liquidacion=liq.dias_trabajados if liq else None,
            atrasos_asistencia=asis.atrasos if asis else None,
            atrasos_liquidacion=liq.atrasos if liq else None,
            sueldo_pagado=liq.sueldo_base_a_pago if liq else None,
            sueldo_base_30=liq.sueldo_base_30_dias if liq else None,
        ))


# ==============================================================================
# EXPORTACIÓN DE RESULTADOS
# ==============================================================================

class ExportadorExcel:
    """Maneja la escritura del reporte Excel con múltiples hojas."""

    def __init__(self, ruta_salida: Path):
        self.ruta_salida = ruta_salida

    def exportar(self, hallazgos: pd.DataFrame, duplicados: pd.DataFrame, resumen: Dict):
        """Escribe el archivo Excel con formato básico."""
        dir_salida = self.ruta_salida.parent
        if not os.access(dir_salida or ".", os.W_OK):
            raise AuditoriaError(f"Sin permisos de escritura en: {dir_salida}")

        try:
            with pd.ExcelWriter(self.ruta_salida, engine="openpyxl") as writer:
                # Hoja 1: Hallazgos
                if not hallazgos.empty:
                    # Formatear columnas de moneda para legibilidad
                    df_out = hallazgos.copy()
                    for col in ["sueldo_pagado", "sueldo_base_30", "sueldo_base_real", "diferencia_sueldo"]:
                        if col in df_out.columns:
                            df_out[col] = df_out[col].apply(lambda x: formatear_moneda(x) if pd.notna(x) else "")
                    df_out.to_excel(writer, sheet_name="Hallazgos", index=False)
                else:
                    pd.DataFrame({"Mensaje": ["Sin hallazgos. Todo coincide correctamente."]}).to_excel(
                        writer, sheet_name="Hallazgos", index=False
                    )

                # Hoja 2: Duplicados
                if not duplicados.empty:
                    duplicados.to_excel(writer, sheet_name="Duplicados", index=False)

                # Hoja 3: Resumen Ejecutivo
                df_resumen = pd.DataFrame([resumen])
                df_resumen.to_excel(writer, sheet_name="Resumen", index=False)

            logger.info(f"\n✅ Reporte guardado: {self.ruta_salida}")
        except Exception as e:
            raise AuditoriaError(f"Error escribiendo Excel: {e}")


# ==============================================================================
# MAIN / CLI
# ==============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Auditoría de liquidaciones vs asistencias desde PDFs (experimental)",
        epilog="Ejemplo: python auditoria_mensual.py --carpeta ./datos/ --redondeo up"
    )
    parser.add_argument("--carpeta", default=".", help="Carpeta donde buscar PDFs")
    parser.add_argument("--asistencia", help="Ruta específica al PDF de asistencia")
    parser.add_argument("--liquidacion", help="Ruta específica al PDF de liquidación")
    parser.add_argument("--salida", default="reporte_auditoria.xlsx", help="Nombre del archivo Excel de salida")
    parser.add_argument("--redondeo", choices=["up", "down", "nearest"], default="up",
                        help="Método de redondeo para sueldo real (default: up)")
    parser.add_argument("--debug", action="store_true", help="Activa logging DEBUG detallado")
    args = parser.parse_args()

    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)

    logger.info("=" * 60)
    logger.info("AUDITORÍA MENSUAL - Liquidaciones vs Asistencias")
    logger.info("Versión refactorizada | Uso experimental y educativo")
    logger.info("=" * 60)

    # --------------------------------------------------------------------------
    # 1. Resolver y validar rutas
    # --------------------------------------------------------------------------
    base_dir = Path(args.carpeta).resolve()
    if not base_dir.is_dir():
        logger.error(f"La carpeta no existe: {base_dir}")
        sys.exit(1)

    # --------------------------------------------------------------------------
    # 2. Descubrir PDFs
    # --------------------------------------------------------------------------
    if args.asistencia and args.liquidacion:
        pdf_asistencia = [validar_ruta_segura(args.asistencia, base_dir)]
        pdf_liquidacion = [validar_ruta_segura(args.liquidacion, base_dir)]
    else:
        pdfs = list(base_dir.glob("*.pdf"))
        if not pdfs:
            logger.error("No se encontraron archivos PDF en la carpeta.")
            sys.exit(1)
        pdf_asistencia = [
            p for p in pdfs
            if any(k in p.name.lower() for k in ("asist", "registro", "attendance"))
        ]
        pdf_liquidacion = [
            p for p in pdfs
            if any(k in p.name.lower() for k in ("liquid", "sueldo", "payroll", "remun"))
        ]
        if not pdf_asistencia or not pdf_liquidacion:
            logger.error("No se detectaron ambos tipos de PDF. Use --asistencia y --liquidacion.")
            sys.exit(1)

    logger.info(f"📁 PDFs de asistencia: {len(pdf_asistencia)}")
    for p in pdf_asistencia:
        logger.info(f"   • {p.name}")
    logger.info(f"📁 PDFs de liquidación: {len(pdf_liquidacion)}")
    for p in pdf_liquidacion:
        logger.info(f"   • {p.name}")

    # --------------------------------------------------------------------------
    # 3. Extraer datos (streaming → listas para auditoría en memoria)
    # --------------------------------------------------------------------------
    # Nota: Para volúmenes > 5.000 trabajadores, considerar base de datos SQLite
    # en lugar de listas en memoria.
    motor_asis = MotorExtraccion(AsistenciaExtractor())
    motor_liq = MotorExtraccion(LiquidacionExtractor())

    asistencias = list(motor_asis.extraer_de_pdfs(pdf_asistencia))
    liquidaciones = list(motor_liq.extraer_de_pdfs(pdf_liquidacion))

    if not asistencias:
        logger.error("No se extrajo ninguna asistencia válida.")
        sys.exit(1)
    if not liquidaciones:
        logger.error("No se extrajo ninguna liquidación válida.")
        sys.exit(1)

    logger.info(f"\n📊 Registros extraídos:")
    logger.info(f"   • Asistencias: {len(asistencias)}")
    logger.info(f"   • Liquidaciones: {len(liquidaciones)}")

    # --------------------------------------------------------------------------
    # 4. Ejecutar auditoría
    # --------------------------------------------------------------------------
    auditor = AuditoriaMensual(metodo_redondeo=args.redondeo)
    df_hallazgos, df_duplicados = auditor.auditar(asistencias, liquidaciones)

    # --------------------------------------------------------------------------
    # 5. Exportar
    # --------------------------------------------------------------------------
    ruta_salida = base_dir / args.salida
    exportador = ExportadorExcel(ruta_salida)

    resumen = {
        "total_trabajadores_asistencia": len({a.rut for a in asistencias}),
        "total_trabajadores_liquidacion": len({l.rut for l in liquidaciones}),
        "hallazgos_criticos": len(df_hallazgos[df_hallazgos["severidad"] == "CRÍTICO"]) if not df_hallazgos.empty else 0,
        "hallazgos_altos": len(df_hallazgos[df_hallazgos["severidad"] == "ALTO"]) if not df_hallazgos.empty else 0,
        "hallazgos_medios": len(df_hallazgos[df_hallazgos["severidad"] == "MEDIO"]) if not df_hallazgos.empty else 0,
        "duplicados_detectados": len(df_duplicados),
        "metodo_redondeo": args.redondeo,
    }

    exportador.exportar(df_hallazgos, df_duplicados, resumen)

    logger.info("\n📋 Estructura del reporte Excel:")
    logger.info("   1. Hallazgos  → Discrepancias detectadas")
    logger.info("   2. Duplicados → RUTs repetidos dentro del mismo documento")
    logger.info("   3. Resumen    → Métricas generales de la auditoría")
    logger.info("\n💡 Tip: Revise la hoja 'Duplicados' antes de confiar en los hallazgos.")


if __name__ == "__main__":
    main()
