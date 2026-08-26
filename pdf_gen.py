'''
┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
┃  🌟 Empresa: Ionclinics                                                                                                                                    ┃
┃  🚀 Proyecto: Auto Bach Record                                                                                                                             ┃
┃  🌈 Versión: v0.0                                                                                                                                          ┃
┃  👨‍💻 Desarrollador: Miguel Sánchez Faubel                                                                                                                   ┃        
┣━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┫
┃  📝 Descripción: Funciones realcionadas con al generacion y formato del pdf                                                                                ┃      
┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
'''
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ IMPORTACION Y CONFIGURACION DE LIBRERIAS ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase.pdfmetrics import stringWidth
from pypdf import PdfReader, PdfWriter, Transformation, PageObject
import math
import time

import os
from pathlib import Path
import pickle

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ RUTAS ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

LOGO = './data/RESOURCES/logo_ionclinics.png'
RUTA_RAIZ = './data/'
RUTA_ERP =  RUTA_RAIZ + 'PNT MDR/ERP/'
RUTA_IMD = RUTA_RAIZ + 'PNT MDR/ERP/I+D/'
RUTA_PNT =  RUTA_RAIZ +'PNT MDR/PNT.7.5-02_CONTROL DE CALIDAD/Registros MDR'
RUTA_PNT_EXCEL = RUTA_RAIZ +'PNT MDR/PNT.7.5-02_CONTROL DE CALIDAD/Registros MDR/REG.7.5-02-02_CONTROL PNT.xlsx'
RUTA_CONFIG ='./config/config.xlsx'
RUTA_ETIQUETAS = './data/IMPRESION_ETIQUETAS/' # PARA LOS REQUISITOS ESPECIALES DEL BACH RECORD DISPOSITIVOS FUERA DE ESPAÑA

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ COLORES ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

AZUL = colors.HexColor("#0070C0")
ROJO = colors.red
NEGRO = colors.black
BLANCO = colors.white

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ ESTILOS ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

estilo_celda = ParagraphStyle(
    name="Celda",
    fontName="Helvetica",
    fontSize=8,
    leading=9,
    textColor=NEGRO,
)

estilo_cabecera = ParagraphStyle(
    name="Cabecera",
    fontName="Helvetica-Bold",
    fontSize=8,
    leading=9,
    textColor=NEGRO,
)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ COMPROBAMOS LOS DIRECTORIOS ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ CLASES ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class GeneradorBatchRecord:

    def __init__(self, nombre_archivo):

        self.pdf = canvas.Canvas(nombre_archivo, pagesize=A4)
        self.nombre_archivo = nombre_archivo    # El nombre del archivo es la ruta de salida tambien
        self.secciones = []                     # Secciones del documento se usa para generar el indice
        self.numero_pagina = 1                  # Contador de páginas
        self.pagina_indice = 1                  # Indica la pagina donde debe insertarse el indice por defecto despues de la primera (la portada)
        self.numero_paginas_indice = 0
        self.paginas_externas = []
        self.contador_bookmarks = 0

    def nueva_pagina(self):
        """
        Finaliza la página actual y comienza una nueva,
        actualizando automáticamente el contador de páginas.
        """

        self.pdf.showPage()
        self.numero_pagina += 1

    def marcar_seccion(self, titulo, nivel=0):
        self.contador_bookmarks += 1
        clave = f"bookmark_{self.contador_bookmarks}"

        self.pdf.bookmarkPage(clave)
        self.pdf.addOutlineEntry(title=titulo, key=clave, level=nivel, closed=False)

        self.secciones.append({
            "clave": clave,
            "titulo": titulo,
            "nivel": nivel,
            "pagina": self.numero_pagina
        })

        return clave
    
    def crear_cabecera(self, ruta_logo= LOGO, codigo="TD01-03-A6",version="01",fecha="2024-04-22"):
        """
        Dibuja la cabecera del Batch Record en la página actual.

        Devuelve:
            Coordenada Y recomendada para empezar el contenido
            inmediatamente después de la cabecera.
        """

        # ========================================================
        # DIMENSIONES DE LA PÁGINA
        # ========================================================

        ancho_pagina, alto_pagina = self.pdf._pagesize

        margen_izq = 10 * mm
        margen_der = 10 * mm
        margen_superior = 7 * mm

        ancho_util = (
            ancho_pagina
            - margen_izq
            - margen_der
        )

        # ========================================================
        # DIMENSIONES CABECERA
        # ========================================================

        alto_cabecera = 32 * mm

        # Reparto aproximado como en tu imagen
        ancho_logo = 45 * mm
        ancho_datos = 53 * mm

        ancho_titulo = (
            ancho_util
            - ancho_logo
            - ancho_datos
        )

        # Coordenada inferior
        y_cabecera = (
            alto_pagina
            - margen_superior
            - alto_cabecera
        )

        # Coordenadas X
        x1 = margen_izq
        x2 = x1 + ancho_logo
        x3 = x2 + ancho_titulo
        x4 = margen_izq + ancho_util

        # ========================================================
        # MARCO EXTERIOR
        # ========================================================

        self.pdf.setStrokeColor(colors.black)
        self.pdf.setLineWidth(0.6)

        self.pdf.rect(
            x1,
            y_cabecera,
            ancho_util,
            alto_cabecera,
            stroke=1,
            fill=0
        )

        # Separadores verticales
        self.pdf.line(
            x2,
            y_cabecera,
            x2,
            y_cabecera + alto_cabecera
        )

        self.pdf.line(
            x3,
            y_cabecera,
            x3,
            y_cabecera + alto_cabecera
        )

        # ========================================================
        # LOGO
        # ========================================================

        margen_logo = 4 * mm

        self.pdf.drawImage(
            ruta_logo,

            x1 + margen_logo,
            y_cabecera + margen_logo,

            width=ancho_logo - 2 * margen_logo,
            height=alto_cabecera - 2 * margen_logo,

            preserveAspectRatio=True,
            anchor="c",
            mask="auto"
        )

        # ========================================================
        # TÍTULO CENTRAL
        # ========================================================

        centro_x = x2 + ancho_titulo / 2

        # LIBERACIÓN LOTES
        texto = "LIBERACIÓN LOTES"

        self.pdf.setFillColor(colors.black)
        self.pdf.setFont(
            "Helvetica-Bold",
            13
        )

        ancho_texto = stringWidth(
            texto,
            "Helvetica-Bold",
            13
        )

        self.pdf.drawString(
            centro_x - ancho_texto / 2,
            y_cabecera + 17 * mm,
            texto
        )

        # BATCH RECORDS
        texto = "BATCH RECORDS"

        self.pdf.setFillColor(
            colors.HexColor("#0070C0")
        )

        self.pdf.setFont(
            "Helvetica-BoldOblique",
            13
        )

        ancho_texto = stringWidth(
            texto,
            "Helvetica-BoldOblique",
            13
        )

        self.pdf.drawString(
            centro_x - ancho_texto / 2,
            y_cabecera + 10 * mm,
            texto
        )

        # ========================================================
        # DATOS DERECHA
        # ========================================================

        centro_datos = (
            x3
            + ancho_datos / 2
        )

        self.pdf.setFillColor(
            colors.black
        )

        self.pdf.setFont(
            "Helvetica-Bold",
            9
        )

        textos_derecha = [
            codigo,
            f"Ver: {version}",
            fecha
        ]

        posiciones_y = [
            y_cabecera + 20 * mm,
            y_cabecera + 14 * mm,
            y_cabecera + 8 * mm
        ]

        for texto, posicion_y in zip(
            textos_derecha,
            posiciones_y
        ):

            ancho_texto = stringWidth(
                texto,
                "Helvetica-Bold",
                9
            )

            self.pdf.drawString(
                centro_datos - ancho_texto / 2,
                posicion_y,
                texto
            )

        return y_cabecera

    # def page_crear_indice(self):
    #     """
    #     Crea una página de índice navegable con enlaces internos
    #     a las secciones registradas en self.secciones.

    #     Características:
    #     - "ÍNDICE /" se muestra en negro.
    #     - "INDEX" se muestra en azul.
    #     - La primera letra de cada sección se muestra en mayúscula.
    #     - Se añaden puntos guía entre el título y el número de página.
    #     - El número de página se alinea a la derecha.
    #     - Respeta el nivel jerárquico de cada sección.
    #     - Cada entrada es clicable y dirige a la sección correspondiente.

    #     IMPORTANTE:
    #     Esta función debe llamarse cuando ya se hayan registrado todas
    #     las secciones del documento.
    #     """
    #     y = self.crear_cabecera(ruta_logo= LOGO, codigo="TD01-03-A6",version="01",fecha="2024-04-22")

    #     ancho, alto = self.pdf._pagesize

    #     margen_izq = 20 * mm
    #     margen_der = 20 * mm

    #     y -= 30 * mm

    #     # ========================================================
    #     # TÍTULO: ÍNDICE / INDEX
    #     # ========================================================

    #     tamano_titulo = 18

    #     texto_es = "ÍNDICE / "
    #     texto_en = "INDEX"

    #     # Parte española en negro
    #     self.pdf.setFillColor(NEGRO)
    #     self.pdf.setFont(
    #         "Helvetica-Bold",
    #         tamano_titulo
    #     )

    #     self.pdf.drawString(
    #         margen_izq,
    #         y,
    #         texto_es
    #     )

    #     # Calcular dónde termina "ÍNDICE / "
    #     ancho_texto_es = self.pdf.stringWidth(
    #         texto_es,
    #         "Helvetica-Bold",
    #         tamano_titulo
    #     )

    #     # INDEX en azul
    #     self.pdf.setFillColor(AZUL)
    #     self.pdf.setFont(
    #         "Helvetica-BoldOblique",
    #         tamano_titulo
    #     )

    #     self.pdf.drawString(
    #         margen_izq + ancho_texto_es,
    #         y,
    #         texto_en
    #     )

    #     # Volver a negro
    #     self.pdf.setFillColor(NEGRO)

    #     y -= 15 * mm

    #     # ========================================================
    #     # SECCIONES
    #     # ========================================================

    #     for seccion in self.secciones:

    #         clave = str(seccion["clave"])

    #         titulo = str(
    #             seccion["titulo"]
    #         ).strip()

    #         nivel = int(
    #             seccion["nivel"]
    #         )

    #         pagina = str(
    #             seccion["pagina"]
    #         )

    #         # ----------------------------------------------------
    #         # Primera letra en mayúscula
    #         #
    #         # No usamos .capitalize(), porque convertiría
    #         # el resto del texto a minúsculas.
    #         # ----------------------------------------------------

    #         if titulo:
    #             titulo = (
    #                 titulo[0].upper()
    #                 + titulo[1:]
    #             )

    #         # ----------------------------------------------------
    #         # Sangría según nivel
    #         # ----------------------------------------------------

    #         sangria = nivel * 8 * mm

    #         x_titulo = (
    #             margen_izq
    #             + sangria
    #         )

    #         # ----------------------------------------------------
    #         # Fuente según nivel
    #         # ----------------------------------------------------

    #         if nivel == 0:

    #             fuente = "Helvetica-Bold"
    #             tamano = 11

    #         else:

    #             fuente = "Helvetica"
    #             tamano = 10

    #         self.pdf.setFont(
    #             fuente,
    #             tamano
    #         )

    #         self.pdf.setFillColor(
    #             NEGRO
    #         )

    #         # ====================================================
    #         # POSICIÓN DEL NÚMERO DE PÁGINA
    #         # ====================================================

    #         ancho_numero_pagina = (
    #             self.pdf.stringWidth(
    #                 pagina,
    #                 fuente,
    #                 tamano
    #             )
    #         )

    #         x_pagina = (
    #             ancho
    #             - margen_der
    #             - ancho_numero_pagina
    #         )

    #         # ====================================================
    #         # DIBUJAR TÍTULO
    #         # ====================================================

    #         self.pdf.drawString(
    #             x_titulo,
    #             y,
    #             titulo
    #         )

    #         # ====================================================
    #         # CALCULAR PUNTOS GUÍA
    #         # ====================================================

    #         ancho_titulo = (
    #             self.pdf.stringWidth(
    #                 titulo,
    #                 fuente,
    #                 tamano
    #             )
    #         )

    #         # Comienzo de los puntos
    #         x_inicio_puntos = (
    #             x_titulo
    #             + ancho_titulo
    #             + 2 * mm
    #         )

    #         # Fin de los puntos
    #         x_fin_puntos = (
    #             x_pagina
    #             - 2 * mm
    #         )

    #         ancho_disponible_puntos = (
    #             x_fin_puntos
    #             - x_inicio_puntos
    #         )

    #         if ancho_disponible_puntos > 0:

    #             ancho_punto = (
    #                 self.pdf.stringWidth(
    #                     ".",
    #                     fuente,
    #                     tamano
    #                 )
    #             )

    #             numero_puntos = int(
    #                 ancho_disponible_puntos
    #                 / ancho_punto
    #             )

    #             puntos = "." * numero_puntos

    #             self.pdf.drawString(
    #                 x_inicio_puntos,
    #                 y,
    #                 puntos
    #             )

    #         # ====================================================
    #         # NÚMERO DE PÁGINA
    #         # ====================================================

    #         self.pdf.drawString(
    #             x_pagina,
    #             y,
    #             pagina
    #         )

    #         # ====================================================
    #         # ENLACE NAVEGABLE
    #         # ====================================================

    #         self.pdf.linkAbsolute(
    #             contents=titulo,
    #             destinationname=clave,
    #             Rect=(
    #                 x_titulo,
    #                 y - 2 * mm,
    #                 ancho - margen_der,
    #                 y + 4 * mm
    #             ),
    #             thickness=0
    #         )

    #         # ----------------------------------------------------
    #         # Siguiente entrada
    #         # ----------------------------------------------------

    #         y -= 8 * mm

    #     # ========================================================
    #     # FINALIZAR PÁGINA
    #     # ========================================================

    #     self.nueva_pagina()

    def page_crear_indice(self):
        ancho, alto = A4

        margen_izquierdo = 25 * mm
        margen_derecho = 20 * mm
        y_inicio = alto - 70 * mm
        y_minimo = 22 * mm
        interlineado = 6.5 * mm

        # ========================================================
        # CALCULAR CUÁNTAS PÁGINAS NECESITARÁ EL ÍNDICE
        # ========================================================

        lineas_por_pagina = int((y_inicio - y_minimo) // interlineado)

        if lineas_por_pagina <= 0:
            raise ValueError("No hay espacio suficiente para generar el índice.")

        self.numero_paginas_indice = max(1, math.ceil(len(self.secciones) / lineas_por_pagina))

        # Como ya habíamos reservado 1 página para el índice,
        # cada página adicional desplazará el contenido posterior.
        desplazamiento_paginas = self.numero_paginas_indice - 1

        # ========================================================
        # FUNCIÓN PARA DIBUJAR EL ENCABEZADO DE CADA PÁGINA
        # ========================================================

        def dibujar_inicio_indice(numero_pagina_indice):
            self.crear_cabecera()

            self.pdf.setFillColor(NEGRO)
            self.pdf.setFont("Helvetica-Bold", 16)
            self.pdf.drawString(margen_izquierdo, alto - 55 * mm, "ÍNDICE /")

            ancho_indice = stringWidth("ÍNDICE /", "Helvetica-Bold", 16)

            self.pdf.setFillColor(AZUL)
            self.pdf.setFont("Helvetica-BoldOblique", 16)
            self.pdf.drawString(margen_izquierdo + ancho_indice + 2 * mm, alto - 55 * mm, "INDEX")

            if self.numero_paginas_indice > 1:
                self.pdf.setFillColor(colors.HexColor("#777777"))
                self.pdf.setFont("Helvetica", 8)
                self.pdf.drawRightString(ancho - margen_derecho, alto - 55 * mm, f"{numero_pagina_indice}/{self.numero_paginas_indice}")

            return y_inicio

        # ========================================================
        # DIBUJAR ÍNDICE
        # ========================================================

        pagina_indice_actual = 1
        y = dibujar_inicio_indice(pagina_indice_actual)

        for posicion, seccion in enumerate(self.secciones):

            # ----------------------------------------------------
            # NUEVA PÁGINA CUANDO NO QUEPA LA SIGUIENTE LÍNEA
            # ----------------------------------------------------

            if y < y_minimo:
                self.nueva_pagina()
                pagina_indice_actual += 1
                y = dibujar_inicio_indice(pagina_indice_actual)

            nivel = seccion.get("nivel", 0)
            titulo = str(seccion["titulo"])
            clave = str(seccion["clave"])

            # Página real después de insertar las páginas
            # adicionales del índice.
            pagina_original = seccion["pagina"]

            if pagina_original > self.pagina_indice:
                pagina_mostrada = pagina_original + desplazamiento_paginas
            else:
                pagina_mostrada = pagina_original

            # ----------------------------------------------------
            # FORMATO SEGÚN NIVEL
            # ----------------------------------------------------

            if nivel == 0:
                fuente = "Helvetica-Bold"
                tamano = 10
            elif nivel == 1:
                fuente = "Helvetica-Bold"
                tamano = 9
            else:
                fuente = "Helvetica"
                tamano = 8.5

            sangria = nivel * 7 * mm
            x_titulo = margen_izquierdo + sangria

            # ----------------------------------------------------
            # TÍTULO
            # ----------------------------------------------------

            self.pdf.setFillColor(NEGRO)
            self.pdf.setFont(fuente, tamano)
            self.pdf.drawString(x_titulo, y, titulo)

            # ----------------------------------------------------
            # NÚMERO DE PÁGINA
            # ----------------------------------------------------

            texto_pagina = str(pagina_mostrada)
            ancho_pagina = stringWidth(texto_pagina, fuente, tamano)
            x_pagina = ancho - margen_derecho - ancho_pagina

            self.pdf.drawString(x_pagina, y, texto_pagina)

            # ----------------------------------------------------
            # PUNTOS
            # ----------------------------------------------------

            ancho_titulo = stringWidth(titulo, fuente, tamano)
            x_inicio_puntos = x_titulo + ancho_titulo + 2 * mm
            x_fin_puntos = x_pagina - 3 * mm

            self.pdf.setFont("Helvetica", 7)

            if x_fin_puntos > x_inicio_puntos:
                ancho_punto = stringWidth(".", "Helvetica", 7)
                numero_puntos = int((x_fin_puntos - x_inicio_puntos) / ancho_punto)

                self.pdf.drawString(x_inicio_puntos, y, "." * numero_puntos)

            # ----------------------------------------------------
            # ENLACE INTERNO
            # ----------------------------------------------------

            self.pdf.linkRect("", clave, Rect=(x_titulo, y - 2 * mm, ancho - margen_derecho, y + 4 * mm), relative=0, thickness=0)

            y -= interlineado

    def page_crear_portada(
        self,
        lote,
        dispositivo,
        ns_inicio,
        ns_final,
        software_version,
        requisitos_especiales=None,
        preparado_por="Victoria E. González Gutiérrez",
        cargo_preparado="Regulatory Responsible",
        revisado_por="Victoria E. González Gutiérrez",
        cargo_revisado="Quality Manager",
        aprobado_por="Josep Oliver Garcia",
        cargo_aprobado="Manager",
        fecha_preparado="2024/02/15",
        fecha_revisado="2024/02/15",
        fecha_aprobado="2024/02/15",
    ):
        """
        Crea la portada del Batch Record.

        Parámetros
        ----------
        self : GeneradorBatchRecord
            Instancia del generador de Batch Records.

        lote : str
            Número o código del lote.

        ns_inicio : str
            Número de serie inicial.

        ns_final : str
            Número de serie final.

        software_version : str
            Versión de software.

        requisitos_especiales : list[dict], opcional
            Ejemplo:
            [
                {
                    "inicio": "EPB1230001",
                    "final": "EPB1230011",
                    "requisito": "V03 INGLES"
                },
                ...
            ]
        """
        self.marcar_seccion("1. HOJA RESUMEN / SUMMARY SHEET.", nivel=0)
        
        if requisitos_especiales is None:
            requisitos_especiales = []

        # --------------------------------------------------------
        # DIMENSIONES DE PÁGINA
        # --------------------------------------------------------

        ancho, alto = A4

        margen_izq = 15 * mm
        margen_der = 15 * mm

        ancho_util = ancho - margen_izq - margen_der

        # ========================================================
        # CABECERA
        # ========================================================

        y = self.crear_cabecera()

        # ========================================================
        # 1. TÍTULO
        # ========================================================

        y -= 17 * mm # MARGEN ENTRE CABECERA Y TÍTULO

        self.pdf.setFillColor(NEGRO)
        self.pdf.setFont("Helvetica-Bold", 23)

        texto_1 = "LIBERACIÓN LOTES /"
        self.pdf.drawString(margen_izq, y, texto_1)

        ancho_texto_1 = stringWidth(
            texto_1,
            "Helvetica-Bold",
            23
        )

        self.pdf.setFillColor(AZUL)
        self.pdf.setFont("Helvetica-BoldOblique", 23)

        self.pdf.drawString(
            margen_izq + ancho_texto_1 + 3 * mm,
            y,
            "BATCH RECORDS"
        )
        # ========================================================
        # 1.1 DISPOSITIVO
        # ========================================================

        y -= 16 * mm
        texto_lote = "DISPOSITIVO /"
        self.pdf.setFillColor(NEGRO)
        self.pdf.setFont("Helvetica-Bold", 18)
        self.pdf.drawString(margen_izq,y,texto_lote)
        ancho_lote1 = stringWidth(texto_lote,"Helvetica-Bold",18)

        texto_lote = " DEVICE:"
        self.pdf.setFillColor(AZUL)
        self.pdf.drawString(margen_izq + ancho_lote1 ,y,texto_lote)
        ancho_lote2 = stringWidth(texto_lote,"Helvetica-Bold",18)
    

        texto_lote = str(dispositivo).upper()
        self.pdf.setFillColor(NEGRO)
        self.pdf.setFont("Helvetica-Bold", 18)

        self.pdf.drawString(margen_izq + ancho_lote1 + ancho_lote2 + 3 * mm, y, texto_lote)

        # ========================================================
        # 2. LOTE
        # ========================================================

        y -= 11 * mm

        texto_lote = "LOTE /"
        self.pdf.setFillColor(NEGRO)
        self.pdf.setFont("Helvetica-Bold", 18)
        self.pdf.drawString(margen_izq,y,texto_lote)
        ancho_lote1 = stringWidth(texto_lote,"Helvetica-Bold",18)

        texto_lote = " BATCH:"
        self.pdf.setFillColor(AZUL)
        self.pdf.drawString(margen_izq + ancho_lote1 ,y,texto_lote)
        ancho_lote2 = stringWidth(texto_lote,"Helvetica-Bold",18)

        self.pdf.setFillColor(NEGRO)
        self.pdf.setFont("Helvetica-Bold", 18)

        self.pdf.drawString(margen_izq + ancho_lote1 + ancho_lote2 + 3 * mm,y,str(lote))

        # ========================================================
        # 3. TABLA Nº SERIE / SOFTWARE
        # ========================================================

        y -= 22 * mm

        datos_software = [
            [
                Paragraph("NS INICIO", estilo_cabecera),
                Paragraph("NS FINAL", estilo_cabecera),
                Paragraph("SOFTWARE VERSION", estilo_cabecera),
            ],
            [
                Paragraph(str(ns_inicio), estilo_celda),
                Paragraph(str(ns_final), estilo_celda),
                Paragraph(str(software_version), estilo_celda),
            ],
        ]

        tabla_software = Table(
            datos_software,
            colWidths=[
                42 * mm,
                42 * mm,
                ancho_util - 84 * mm
            ],
            rowHeights=[
                6 * mm,
                6 * mm
            ]
        )

        tabla_software.setStyle(
            TableStyle([
                ("GRID", (0, 0), (-1, -1), 0.5, NEGRO),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3 * mm),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2 * mm),
                ("TOPPADDING", (0, 0), (-1, -1), 1 * mm),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1 * mm),
            ])
        )

        w, h = tabla_software.wrap(ancho_util, 30 * mm)

        tabla_software.drawOn(
            self.pdf,
            margen_izq,
            y - h
        )

        y -= h + 16 * mm

        # ========================================================
        # 4. TABLA REQUISITOS ESPECIALES
        # ========================================================

        titulo_requisitos = Paragraph(
            'REQUISITOS ESPECIALES / '
            '<font color="#0070C0">'
            '<i>SPECIAL REQUIREMENTS</i>'
            '</font>',
            estilo_cabecera
        )

        datos_requisitos = [
            [
                Paragraph("NS INICIO", estilo_cabecera),
                Paragraph("NS FINAL", estilo_cabecera),
                titulo_requisitos
            ]
        ]

        for requisito in requisitos_especiales:

            datos_requisitos.append([
                Paragraph(
                    str(requisito.get("inicio", "")),
                    estilo_celda
                ),
                Paragraph(
                    str(requisito.get("final", "")),
                    estilo_celda
                ),
                Paragraph(
                    str(requisito.get("requisito", "")),
                    estilo_celda
                ),
            ])

        # Si no existen requisitos, mostramos una fila vacía
        if not requisitos_especiales:
            datos_requisitos.append([
                "",
                "",
                Paragraph(
                    "SIN REQUISITOS ESPECIALES / "
                    '<font color="#0070C0">'
                    "<i>NO SPECIAL REQUIREMENTS</i>"
                    "</font>",
                    estilo_celda
                )
            ])

        tabla_requisitos = Table(
            datos_requisitos,
            colWidths=[
                42 * mm,
                42 * mm,
                ancho_util - 84 * mm
            ],
            rowHeights=6 * mm,

            # Repetir cabecera cuando se divide la tabla
            repeatRows=1
        )

        tabla_requisitos.setStyle(
            TableStyle([
                ("GRID", (0, 0), (-1, -1), 0.5, NEGRO),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3 * mm),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2 * mm),
                ("TOPPADDING", (0, 0), (-1, -1), 1 * mm),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1 * mm),
            ])
        )

        # DIBUJAR TABLA REQUISITOS CON PAGINACIÓN

        margen_inferior = 20 * mm
        margen_superior = 20 * mm

        # Espacio disponible en la primera página
        alto_disponible = y - margen_inferior

        tabla_actual = tabla_requisitos


        while tabla_actual:

            # Dividir la tabla según el espacio disponible
            partes = tabla_actual.split(
                ancho_util,
                alto_disponible
            )

            if not partes:
                raise ValueError(
                    "La tabla no puede dividirse en el espacio disponible."
                )

            # Primera parte que cabe en esta página
            parte_actual = partes[0]

            w, h = parte_actual.wrap(
                ancho_util,
                alto_disponible
            )

            parte_actual.drawOn(
                self.pdf,
                margen_izq,
                y - h
            )

            # Guardamos dónde termina esta parte
            y_fin_tabla_requisitos = y - h

            # ¿QUEDA TABLA POR DIBUJAR?

            if len(partes) > 1:

                # La segunda parte contiene todo lo restante
                tabla_actual = partes[1]

                # Nueva página
                self.nueva_pagina()

                # Recuperar tamaño de página
                ancho, alto = self.pdf._pagesize

                # Reiniciar posición vertical
                y = alto - margen_superior

                # Nuevo espacio disponible
                alto_disponible = (
                    y
                    - margen_inferior
                )

            else:

                # Ya hemos terminado toda la tabla
                tabla_actual = None

        y_fin_tabla_requisitos = y - h

        # ========================================================
        # 5. BLOQUE DE FIRMAS
        # ========================================================

        y_firmas_fija = 18 * mm # Determina la posicion de la esquina inferior de las firmas a partir de la cual se crea el bloque


        altura_firma = 51 * mm
        altura_fecha = 7 * mm

        sep_requisitos_firmas = y_fin_tabla_requisitos - 5 * mm - altura_firma - altura_fecha

        if sep_requisitos_firmas < y_firmas_fija: # Si la tabla de requisitos es demasiado grande, se crea una nueva página para las firmas
            self.nueva_pagina()
            # y_firmas_fija = sep_requisitos_firmas
        
        ancho_columna = ancho_util / 3

        # Coordenadas
        x1 = margen_izq
        x2 = margen_izq + ancho_columna
        x3 = margen_izq + ancho_columna * 2
        x4 = margen_izq + ancho_util

        y_inferior = y_firmas_fija
        y_fecha = y_firmas_fija + altura_fecha
        y_superior = y_fecha + altura_firma

        # --------------------------------------------------------
        # Rectángulo exterior
        # --------------------------------------------------------

        self.pdf.setStrokeColor(NEGRO)
        self.pdf.setLineWidth(0.5)

        self.pdf.rect(
            x1,
            y_inferior,
            ancho_util,
            altura_firma + altura_fecha,
            stroke=1,
            fill=0
        )

        # Divisiones verticales
        self.pdf.line(
            x2,
            y_inferior,
            x2,
            y_superior
        )

        self.pdf.line(
            x3,
            y_inferior,
            x3,
            y_superior
        )

        # Línea horizontal de fecha
        self.pdf.line(
            x1,
            y_fecha,
            x4,
            y_fecha
        )

        # ========================================================
        # TEXTO FIRMAS
        # ========================================================

        padding = 3 * mm

        self.pdf.setFillColor(NEGRO)

        # --------------------------------------------------------
        # PREPARADO
        # --------------------------------------------------------

        self.pdf.setFont("Helvetica", 9)

        self.pdf.drawString(
            x1 + padding,
            y_superior - 6 * mm,
            "Prepared by:"
        )

        self.pdf.drawString(
            x1 + padding,
            y_superior - 12 * mm,
            preparado_por
        )

        self.pdf.drawString(
            x1 + padding,
            y_superior - 18 * mm,
            cargo_preparado
        )

        # --------------------------------------------------------
        # REVISADO
        # --------------------------------------------------------

        self.pdf.drawString(
            x2 + padding,
            y_superior - 6 * mm,
            "Revised by:"
        )

        self.pdf.drawString(
            x2 + padding,
            y_superior - 12 * mm,
            revisado_por
        )

        self.pdf.drawString(
            x2 + padding,
            y_superior - 18 * mm,
            cargo_revisado
        )

        # --------------------------------------------------------
        # APROBADO
        # --------------------------------------------------------

        self.pdf.drawString(
            x3 + padding,
            y_superior - 6 * mm,
            "Approved by:"
        )

        self.pdf.drawString(
            x3 + padding,
            y_superior - 12 * mm,
            aprobado_por
        )

        self.pdf.drawString(
            x3 + padding,
            y_superior - 18 * mm,
            cargo_aprobado
        )

        # ========================================================
        # FECHAS
        # ========================================================

        self.pdf.setFont("Helvetica", 9)

        self.pdf.drawString(
            x1 + padding,
            y_inferior + 2 * mm,
            fecha_preparado
        )

        self.pdf.drawString(
            x2 + padding,
            y_inferior + 2 * mm,
            fecha_revisado
        )

        self.pdf.drawString(
            x3 + padding,
            y_inferior + 2 * mm,
            fecha_aprobado
        )

        # ========================================================
        # FINALIZAR PÁGINA
        # ========================================================

        self.nueva_pagina() #
        self.pagina_indice = self.numero_pagina # Indicamos que la pagina donde se insertara el indice es la siguiente a la portada
        self.nueva_pagina() # Dejamos una pagina en blanco para el indice siempre despues de crear la portada

    def page_crear_separador_seccion(self, titulo, indice=1, numero=None, contexto=None, subtitulo=None, descripcion=None):
        """
        Crea una página completa utilizada exclusivamente como separador
        o portada de una sección del Batch Record.

        indice:
            0 -> Sección principal
            1 -> Subsección
            2 -> SubSUBsección

        contexto:
            Texto opcional que muestra la jerarquía superior.
            Ejemplo:
            "TRAZABILIDAD IONCLINICS / ÓRDENES DE PRODUCCIÓN"
        """

        # ========================================================
        # REGISTRAR EN ÍNDICE Y MARCADORES PDF
        # ========================================================

        nivel = indice - 1
        titulo_indice = f"{numero} {titulo}" if numero else str(titulo)
        self.marcar_seccion(titulo_indice, nivel=nivel)

        # ========================================================
        # DIMENSIONES
        # ========================================================

        ancho, alto = self.pdf._pagesize
        margen = 20 * mm

        # ========================================================
        # CREAMOS LA CABECERA
        # ========================================================

        y_titulo = self.crear_cabecera()

        # ========================================================
        # NÚMERO DE SECCIÓN
        # ========================================================

        if numero:
            self.pdf.setFillColor(colors.HexColor("#D9EAF7"))
            self.pdf.setFont("Helvetica-Bold", 55)
            self.pdf.drawRightString(ancho - margen, alto - 80 * mm, str(numero))

        # ========================================================
        # LÍNEA SUPERIOR
        # ========================================================

        y_titulo -= ( alto * 0.30)

        self.pdf.setStrokeColor(AZUL)
        self.pdf.setLineWidth(2)
        self.pdf.line(margen, y_titulo + 22 * mm, ancho - margen, y_titulo + 22 * mm)

        # ========================================================
        # SECCIÓN / SECTION
        # ========================================================

        # self.pdf.setFillColor(NEGRO)
        # self.pdf.setFont("Helvetica-Bold", 13)
        # self.pdf.drawString(margen, y + 12 * mm, texto_es)

        # ancho_texto = stringWidth(texto_es, "Helvetica-Bold", 13)

        # self.pdf.setFillColor(AZUL)
        # self.pdf.setFont("Helvetica-BoldOblique", 13)
        # self.pdf.drawString(margen + ancho_texto + 2 * mm, y + 12 * mm, f"/ {texto_en}")

        # ========================================================
        # TÍTULO PRINCIPAL
        # ========================================================

        if indice == 1:
            tamano_titulo = 27
        elif indice == 2:
            tamano_titulo = 24
        else:
            tamano_titulo = 22

        titulo = str(titulo).strip()

        if "/" in titulo:
            titulo_es, titulo_en = titulo.split("/", 1)
            titulo_es = titulo_es.strip().upper()
            titulo_en = titulo_en.strip().upper()

            texto_es = f"{titulo_es} /"
            texto_en = f" {titulo_en}"

            fuente_es = "Helvetica-Bold"
            fuente_en = "Helvetica-BoldOblique"

            ancho_es = stringWidth(texto_es, fuente_es, tamano_titulo)
            ancho_en = stringWidth(texto_en, fuente_en, tamano_titulo)
            ancho_total = ancho_es + ancho_en
            ancho_disponible = ancho - 2 * margen

            # ----------------------------------------------------
            # SI CABE TODO EN UNA LÍNEA
            # ----------------------------------------------------

            if ancho_total <= ancho_disponible:
                x = (ancho - ancho_total) / 2

                self.pdf.setFillColor(NEGRO)
                self.pdf.setFont(fuente_es, tamano_titulo)
                self.pdf.drawString(x, y_titulo, texto_es)

                self.pdf.setFillColor(AZUL)
                self.pdf.setFont(fuente_en, tamano_titulo)
                self.pdf.drawString(x + ancho_es, y_titulo, texto_en)

                y_titulo -= 12 * mm

            # ----------------------------------------------------
            # SI NO CABE -> DOS LÍNEAS
            # ----------------------------------------------------

            else:
                self.pdf.setFillColor(NEGRO)
                self.pdf.setFont(fuente_es, tamano_titulo)
                self.pdf.drawCentredString(ancho / 2, y_titulo, texto_es)

                y_titulo -= 11 * mm

                self.pdf.setFillColor(AZUL)
                self.pdf.setFont(fuente_en, tamano_titulo)
                self.pdf.drawCentredString(ancho / 2, y_titulo, titulo_en)

                y_titulo -= 12 * mm

        else:

            titulo = titulo.upper()

            self.pdf.setFillColor(NEGRO)
            self.pdf.setFont("Helvetica-Bold", tamano_titulo)
            self.pdf.drawCentredString(ancho / 2, y_titulo, titulo)

            y_titulo -= 12 * mm


        # ========================================================
        # DESCRIPCIÓN OPCIONAL
        # Antes de "/"  -> NEGRO
        # "/"           -> NEGRO
        # Después de "/" -> AZUL + CURSIVA
        # ========================================================

        if descripcion:
            y_titulo -= 4 * mm

            x_descripcion = margen
            x_actual = x_descripcion
            ancho_maximo = ancho - 2 * margen
            tamano_descripcion = 10
            interlineado = 5 * mm

            descripcion = str(descripcion).strip()

            # ----------------------------------------------------
            # CREAR TOKENS CONSERVANDO LA "/"
            # ----------------------------------------------------

            if "/" in descripcion:
                descripcion_es, descripcion_en = descripcion.split("/", 1)

                tokens = []

                for palabra in descripcion_es.strip().split():
                    tokens.append((palabra, "Helvetica", NEGRO))

                tokens.append(("/", "Helvetica", NEGRO))

                for palabra in descripcion_en.strip().split():
                    tokens.append((palabra, "Helvetica-Oblique", AZUL))

            else:
                tokens = [(palabra, "Helvetica", NEGRO) for palabra in descripcion.split()]

            # ----------------------------------------------------
            # DIBUJAR COMO UN ÚNICO PÁRRAFO
            # ----------------------------------------------------

            primera_palabra_linea = True

            for palabra, fuente, color in tokens:

                texto = palabra if primera_palabra_linea else f" {palabra}"
                ancho_texto = stringWidth(texto, fuente, tamano_descripcion)

                # ------------------------------------------------
                # SALTO SOLO SI SE SUPERA EL ANCHO DISPONIBLE
                # ------------------------------------------------

                if not primera_palabra_linea and x_actual + ancho_texto > x_descripcion + ancho_maximo:
                    y_titulo -= interlineado
                    x_actual = x_descripcion
                    texto = palabra
                    ancho_texto = stringWidth(texto, fuente, tamano_descripcion)
                    primera_palabra_linea = True

                self.pdf.setFillColor(color)
                self.pdf.setFont(fuente, tamano_descripcion)
                self.pdf.drawString(x_actual, y_titulo, texto)

                x_actual += ancho_texto
                primera_palabra_linea = False

        # ========================================================
        # SUBTÍTULO OPCIONAL
        # ========================================================

        if subtitulo:
            y_titulo -= 8 * mm
            self.pdf.setFillColor(AZUL)
            self.pdf.setFont("Helvetica-Oblique", 13)
            ancho_subtitulo = stringWidth(str(subtitulo), "Helvetica-Oblique", 13)
            self.pdf.drawString((ancho - ancho_subtitulo) / 2, y_titulo, str(subtitulo))

        # ========================================================
        # CONTEXTO / JERARQUÍA
        # ========================================================

        if contexto:
            self.pdf.setFillColor(colors.HexColor("#666666"))
            self.pdf.setFont("Helvetica", 9)
            self.pdf.drawCentredString(ancho / 2, 28 * mm, str(contexto).upper())

        # ========================================================
        # LÍNEA INFERIOR
        # ========================================================

        self.pdf.setStrokeColor(AZUL)
        self.pdf.setLineWidth(0.8)
        self.pdf.line(margen, 20 * mm, ancho - margen, 20 * mm)

        # ========================================================
        # LA PÁGINA ES EXCLUSIVAMENTE EL SEPARADOR
        # ========================================================

        self.nueva_pagina()


    def page_añadir_pdf(self, ruta_pdf, paginas=None):
        """
        Reserva páginas dentro del Batch Record para insertar posteriormente
        páginas procedentes de otro PDF.

        parametros:
            ruta_pdf:
                Ruta del PDF que se desea incorporar.

            paginas:
                None -> incorpora todas las páginas.
                int -> incorpora una única página.
                list/tuple -> incorpora las páginas indicadas.

                La numeración comienza en 1.

        Ejemplos:
            self.page_añadir_pdf("documento.pdf")
            self.page_añadir_pdf("documento.pdf", paginas=1)
            self.page_añadir_pdf("documento.pdf", paginas=[1, 2, 5])
        """

        ruta_pdf = Path(ruta_pdf)

        if not ruta_pdf.exists():
            raise FileNotFoundError(f"No existe el PDF: {ruta_pdf}")

        reader = PdfReader(str(ruta_pdf))
        numero_paginas = len(reader.pages)

        if paginas is None:
            paginas = list(range(1, numero_paginas + 1))
        elif isinstance(paginas, int):
            paginas = [paginas]
        else:
            paginas = list(paginas)

        for pagina_origen in paginas:

            if pagina_origen < 1 or pagina_origen > numero_paginas:
                raise ValueError(f"La página {pagina_origen} no existe en {ruta_pdf.name}. El PDF tiene {numero_paginas} páginas.")

            self.paginas_externas.append({
                "pagina_destino": self.numero_pagina,
                "ruta_pdf": str(ruta_pdf),
                "pagina_origen": pagina_origen
            })

            self.nueva_pagina()
            
    def guardar(self):
        """
        Guarda el PDF generado y mueve la última página a la posición indicada, 
        reemplazando la página que actualmente ocupa esa posición (Ya reservada en blanco).
        """
        def reemplazar_paginas_externas(ruta_pdf, paginas_externas, pagina_indice, numero_paginas_indice):

            if not paginas_externas:
                return

            # ========================================================
            # INSERTAR PDF EXTERNO SOBRE LA PÁGINA PLACEHOLDER
            # Ajusta automáticamente a A4 vertical u horizontal
            # ========================================================

            def insertar_pagina_externa_en_pagina_base(pagina_base, pagina_externa):

                ancho_a4, alto_a4 = A4
                margen = 5 * mm

                # Aplicar rotación real al contenido
                if pagina_externa.rotation:
                    pagina_externa.transfer_rotation_to_content()

                x0 = float(pagina_externa.cropbox.left)
                y0 = float(pagina_externa.cropbox.bottom)
                x1 = float(pagina_externa.cropbox.right)
                y1 = float(pagina_externa.cropbox.top)

                ancho_original = x1 - x0
                alto_original = y1 - y0

                if ancho_original <= 0 or alto_original <= 0:
                    raise ValueError("La página externa tiene unas dimensiones no válidas.")

                # ====================================================
                # CALCULAR SI APROVECHA MEJOR A4 VERTICAL U HORIZONTAL
                # ====================================================

                escala_vertical = min((ancho_a4 - 2 * margen) / ancho_original, (alto_a4 - 2 * margen) / alto_original)
                escala_horizontal = min((alto_a4 - 2 * margen) / ancho_original, (ancho_a4 - 2 * margen) / alto_original)

                if escala_horizontal > escala_vertical:
                    ancho_pagina = alto_a4
                    alto_pagina = ancho_a4
                    escala = escala_horizontal
                    orientacion = "HORIZONTAL"
                else:
                    ancho_pagina = ancho_a4
                    alto_pagina = alto_a4
                    escala = escala_vertical
                    orientacion = "VERTICAL"

                # ====================================================
                # CAMBIAR EL TAMAÑO DE LA MISMA PÁGINA PLACEHOLDER
                # ====================================================

                pagina_base.mediabox.lower_left = (0, 0)
                pagina_base.mediabox.upper_right = (ancho_pagina, alto_pagina)

                pagina_base.cropbox.lower_left = (0, 0)
                pagina_base.cropbox.upper_right = (ancho_pagina, alto_pagina)

                # ====================================================
                # CENTRAR EL CONTENIDO
                # ====================================================

                ancho_escalado = ancho_original * escala
                alto_escalado = alto_original * escala

                desplazamiento_x = (ancho_pagina - ancho_escalado) / 2
                desplazamiento_y = (alto_pagina - alto_escalado) / 2

                transformacion = Transformation().translate(-x0, -y0).scale(escala).translate(desplazamiento_x, desplazamiento_y)

                pagina_base.merge_transformed_page(pagina_externa, transformacion, over=True, expand=False)

                return orientacion

            # ========================================================
            # ABRIR PDF PRINCIPAL
            # ========================================================

            ruta_pdf = Path(ruta_pdf)

            reader = PdfReader(str(ruta_pdf))

            # Clonamos el PDF completo para mantener bookmarks,
            # enlaces internos, anotaciones, etc.
            writer = PdfWriter()
            writer.clone_document_from_reader(reader)

            # ========================================================
            # CORRECCIÓN POR ÍNDICE MULTIPÁGINA
            # ========================================================

            desplazamiento_indice = max(0, numero_paginas_indice - 1)

            # ========================================================
            # INSERTAR CADA PDF EXTERNO
            # ========================================================

            for documento in paginas_externas:

                pagina_destino_original = documento["pagina_destino"]
                ruta_externa = documento["ruta_pdf"]
                pagina_origen = documento["pagina_origen"]

                if pagina_destino_original > pagina_indice:
                    pagina_destino_final = pagina_destino_original + desplazamiento_indice
                else:
                    pagina_destino_final = pagina_destino_original

                if pagina_destino_final < 1 or pagina_destino_final > len(writer.pages):
                    raise ValueError(f"Página destino fuera de rango: {pagina_destino_final}. El PDF tiene {len(writer.pages)} páginas.")

                reader_externo = PdfReader(str(ruta_externa))

                if pagina_origen < 1 or pagina_origen > len(reader_externo.pages):
                    raise ValueError(f"La página {pagina_origen} no existe en {ruta_externa}.")

                pagina_externa = reader_externo.pages[pagina_origen - 1]
                pagina_base = writer.pages[pagina_destino_final - 1]

                orientacion = insertar_pagina_externa_en_pagina_base(pagina_base, pagina_externa)

            # ========================================================
            # GUARDAR
            # ========================================================

            ruta_temporal = ruta_pdf.with_name(f"{ruta_pdf.stem}_externos_temp{ruta_pdf.suffix}")

            with open(ruta_temporal, "wb") as archivo:
                writer.write(archivo)

            os.replace(ruta_temporal, ruta_pdf)

        def mover_paginas_indice(ruta_pdf, posicion, numero_paginas_indice):
            ruta_pdf = Path(ruta_pdf)

            reader = PdfReader(str(ruta_pdf))
            paginas = list(reader.pages)

            if numero_paginas_indice <= 0:
                return

            if numero_paginas_indice > len(paginas):
                raise ValueError("El número de páginas del índice es superior al número de páginas del PDF.")

            # Las N últimas páginas son el índice generado
            paginas_indice = paginas[-numero_paginas_indice:]
            paginas_documento = paginas[:-numero_paginas_indice]

            # posicion es 1-based.
            indice_insercion = posicion - 1

            # Eliminar la página en blanco que habíamos reservado originalmente
            if 0 <= indice_insercion < len(paginas_documento):
                paginas_documento.pop(indice_insercion)

            paginas_finales = paginas_documento[:indice_insercion] + paginas_indice + paginas_documento[indice_insercion:]

            writer = PdfWriter()

            for pagina in paginas_finales:
                writer.add_page(pagina)

            ruta_temporal = ruta_pdf.with_name(f"{ruta_pdf.stem}_indice_temp{ruta_pdf.suffix}")

            with open(ruta_temporal, "wb") as archivo:
                writer.write(archivo)

            os.replace(ruta_temporal, ruta_pdf)

        def reconstruir_marcadores_pdf(ruta_pdf, secciones, pagina_indice, numero_paginas_indice):

            ruta_pdf = Path(ruta_pdf)

            reader = PdfReader(str(ruta_pdf))

            writer = PdfWriter()
            writer.clone_document_from_reader(reader)

            desplazamiento_indice = max(0, numero_paginas_indice - 1)

            # ========================================================
            # ELIMINAR OUTLINE ANTERIOR SI EXISTE
            # ========================================================

            if "/Outlines" in writer._root_object:
                del writer._root_object["/Outlines"]

            # ========================================================
            # RECREAR JERARQUÍA DE MARCADORES
            # ========================================================

            padres = {}

            for seccion in secciones:

                titulo = str(seccion["titulo"])
                nivel = int(seccion.get("nivel", 0))
                pagina_original = int(seccion["pagina"])

                # Corregir número debido a las páginas adicionales del índice
                if pagina_original > pagina_indice:
                    pagina_final = pagina_original + desplazamiento_indice
                else:
                    pagina_final = pagina_original

                # pypdf trabaja con índices empezando por 0
                indice_pagina = pagina_final - 1

                if indice_pagina < 0 or indice_pagina >= len(writer.pages):
                    print(f"AVISO: marcador fuera de rango: {titulo} -> página {pagina_final}")
                    continue

                # ====================================================
                # DETERMINAR EL PADRE SEGÚN EL NIVEL
                # ====================================================

                if nivel == 0:
                    padre = None
                else:
                    padre = padres.get(nivel - 1)

                marcador = writer.add_outline_item(titulo, indice_pagina, parent=padre)

                padres[nivel] = marcador

                # Borrar niveles inferiores antiguos
                for nivel_guardado in list(padres.keys()):
                    if nivel_guardado > nivel:
                        del padres[nivel_guardado]

            # Intentar que el visor abra directamente el panel de marcadores
            writer.page_mode = "/UseOutlines"

            ruta_temporal = ruta_pdf.with_name(f"{ruta_pdf.stem}_marcadores_temp{ruta_pdf.suffix}")

            with open(ruta_temporal, "wb") as archivo:
                writer.write(archivo)

            os.replace(ruta_temporal, ruta_pdf)

        self.pdf.save()
        mover_paginas_indice(self.nombre_archivo, self.pagina_indice, self.numero_paginas_indice)
        reemplazar_paginas_externas(self.nombre_archivo, self.paginas_externas, self.pagina_indice, self.numero_paginas_indice)
        reconstruir_marcadores_pdf(self.nombre_archivo, self.secciones, self.pagina_indice, self.numero_paginas_indice)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ FUNCIONES ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

def cargar_diccionario(ruta="./temp/estructura_documental.pkl"):
    with open(ruta, "rb") as archivo:
        return pickle.load(archivo)


if __name__ == "__main__":

    estructura_documental = cargar_diccionario()
    dispositivo = "EPTEV02DEV01"
    bachredord = GeneradorBatchRecord("BATCH_RECORD_EPTE.pdf")
    requisitos = [
        {
            "inicio": "EPB1230001",
            "final": "EPB1230011",
            "requisito": "V03 INGLES",
        },
        {
            "inicio": "EPB1230023",
            "final": "EPB1230025",
            "requisito": "V03 INGLES",
        },
        {
            "inicio": "EPB1230035",
            "final": "EPB1230038",
            "requisito": "V03 INGLES",
        },
        {
            "inicio": "EPB1230061",
            "final": "EPB1230062",
            "requisito": "V03 INGLES",
        },
        {
            "inicio": "EPB1230084",
            "final": "EPB1230087",
            "requisito": "V03 INGLES",
        },
    ]

    bachredord.page_crear_portada(

        lote="XXXX",

        dispositivo="EPTEV02DEV01",

        ns_inicio="EPB1230001",
        ns_final="EPB1230090",

        software_version="4643_3385",

        requisitos_especiales=requisitos,

        preparado_por="Victoria E. González Gutiérrez",
        cargo_preparado="Regulatory Responsible",

        revisado_por="Victoria E. González Gutiérrez",
        cargo_revisado="Quality Manager",

        aprobado_por="Josep Oliver Garcia",
        cargo_aprobado="Manager",

        fecha_preparado="2024/02/15",
        fecha_revisado="2024/02/15",
        fecha_aprobado="2024/02/15",
    )

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ TRAZABILIDAD DE DISMUNTEL / DISMUNTEL TRACEABILITY ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━═════════════════════════════════════════════════════════════════════════════┛
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    
    bachredord.page_crear_separador_seccion("TRAZABILIDAD DE DISMUNTEL / DISMUNTEL TRACEABILITY", indice=1, numero="2")

    bachredord.page_crear_separador_seccion(
        "Guía fabricación / Manufacturing guide",
        descripcion="Guía de fabricación del lote liberado generada por el fabricante. / Manufacturing guide of the released batch generated by the manufacturer.",
        indice=2,
        numero="2.1",
        contexto="TRAZABILIDAD DE DISMUNTEL / TRAZABILIDAD DE DISMUNTEL"
    )
    bachredord.page_crear_separador_seccion(
        "Trazabilidad MP-Lote Fabricación / Traceability MP-Batch Production",
        descripcion="Trazabilidad de fabricación del lote liberado generada por el fabricante. / Manufacturing traceability of the released batch generated by the manufacturer.",
        indice=2,
        numero="2.2",
        contexto="TRAZABILIDAD DE DISMUNTEL / TRAZABILIDAD DE DISMUNTEL"
    )
    bachredord.page_crear_separador_seccion(
        "Lotes / Batches",
        descripcion="Albaranes de Dismuntel que conforman el lote liberado. / Dismuntel delivery notes that make up the released batch.",
        indice=2,
        numero="2.3",
        contexto="TRAZABILIDAD DE DISMUNTEL / TRAZABILIDAD DE DISMUNTEL"
    )

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ TRAZABILIDAD DE IONCLINICS / IONCLINICS TRACEABILITY ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━═════════════════════════════════════════════════════════════════════════════┛
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    
    bachredord.page_crear_separador_seccion("TRAZABILIDAD DE IONCLINICS / IONCLINICS TRACEABILITY", indice=1, numero="3")

    bachredord.page_crear_separador_seccion(
        "Trazabilidad / Traceability",
        descripcion='''Trazabilidad obtenida del software ERP de los números de serie de los dispositivos que conforman el lote
                    liberado. / Traceability obtained from the ERP software of the serial numbers of the devices that make up
                    the released batch.''',
        indice=2,
        numero="3.1",
        contexto="TRAZABILIDAD DE IONCLINICS / IONCLINICS TRACEABILITY"
    )

    bachredord.page_crear_separador_seccion(
        "Ordenes de producción / Production orders",
        indice=2,
        descripcion="Ordenes de producción obtenidas del programa ERP. / Production orders obtained from ERP software.",
        numero="3.2",
        contexto="TRAZABILIDAD DE IONCLINICS / IONCLINICS TRACEABILITY"
    )

    bachredord.page_crear_separador_seccion(
        "Dispositivo / Device",
        indice=3,
        descripcion='''Ordenes de producción obtenidas del programa ERP del dispositivo ensamblado. Une la trazabilidad de la
                    consola y los accesorios del dispositivo. / Production orders obtained from the ERP software of the
                    assembled device. Links the traceability of the console and device accessories.''',
        numero="3.2.1",
        contexto="TRAZABILIDAD DE IONCLINICS / IONCLINICS TRACEABILITY"
    )

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ Accesorios / Accessories

    bachredord.page_crear_separador_seccion(
        "Accesorios / Accessories",
        indice=4,
        descripcion='''Ordenes de producción obtenidas del programa ERP del dispositivo ensamblado. Une la trazabilidad de la
                    consola y los accesorios del dispositivo. / Production orders obtained from the ERP software of the
                    assembled device. Links the traceability of the console and device accessories.''',
        numero="3.2.2",
        contexto="TRAZABILIDAD DE IONCLINICS / IONCLINICS TRACEABILITY"
    )


    articulos_bachrecord = [articulo for articulo in estructura_documental["componentes"]]
    articulos_imd = [carpeta.name for carpeta in Path(RUTA_IMD).iterdir() if carpeta.is_dir()]
    accesorios_IMD = [articulo for articulo in estructura_documental["componentes"] if articulo in articulos_imd]
    accesorios_IMD.remove(dispositivo)
    i = 0
    for accesorio in accesorios_IMD:

        bachredord.page_crear_separador_seccion(
            f"{accesorio} / {accesorio}",
            indice=4,
            numero=f"3.2.2.{str(i)}",
            contexto=f"TRAZABILIDAD DE IONCLINICS / IONCLINICS TRACEABILITY"
        )
        bachredord.page_añadir_pdf("data\PNT MDR\ERP\I+D\CABLEPLANO\CABLEPLANO_RESUMEN_formateado.pdf") #TODO PONER LO QUE TOQUE
        i +=1


    # bachredord.page_crear_separador_seccion(
    #     "Lotes / Batches",
    #     indice=2,
    #     descripcion="Guía de fabricación del lote liberado generada por el fabricante. / Manufacturing guide of the releasedbatch generated by the manufacturer.",
    #     numero="3.3",
    #     contexto="TRAZABILIDAD DE IONCLINICS / IONCLINICS TRACEABILITY"
    # )

    bachredord.page_crear_indice()
    bachredord.guardar()