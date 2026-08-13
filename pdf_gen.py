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

from pypdf import PdfReader, PdfWriter
import os
from pathlib import Path

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ RUTAS ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

LOGO = './data/RESOURCES/logo_ionclinics.png'

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

    def nueva_pagina(self):
        """
        Finaliza la página actual y comienza una nueva,
        actualizando automáticamente el contador de páginas.
        """

        self.pdf.showPage()
        self.numero_pagina += 1

    def marcar_seccion(self, clave, titulo, nivel=0):
        """
        Registra la página actual como una sección navegable del PDF.

        Parámetros
        ----------
        clave : str
            Identificador interno único de la sección.
            No debe repetirse dentro del mismo documento.
            Ejemplos: "portada", "requisitos", "produccion".

        titulo : str
            Texto que se mostrará en el panel de marcadores del lector PDF.
            Ejemplo: "Requisitos especiales".

        nivel : int, opcional
            Nivel jerárquico del marcador dentro del índice lateral.
            El valor 0 corresponde a una sección principal, 1 a una
            subsección, 2 a una subsección de segundo nivel, etc.
            Por defecto es 0.

        Ejemplo
        -------
        self.marcar_seccion(
            clave="control_calidad",
            titulo="Control de calidad",
            nivel=0
        )
        """

        self.pdf.bookmarkPage(clave)

        self.pdf.addOutlineEntry(
            title=titulo,
            key=clave,
            level=nivel,
            closed=False
        )

        # --------------------------------------------------------
        # Guardar automáticamente la sección
        # --------------------------------------------------------
        
        self.secciones.append({
            "clave": clave,
            "titulo": titulo,
            "nivel": nivel,
            "pagina": self.numero_pagina
        })

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

    def page_crear_indice(self):
        """
        Crea una página de índice navegable con enlaces internos
        a las secciones registradas en self.secciones.

        Características:
        - "ÍNDICE /" se muestra en negro.
        - "INDEX" se muestra en azul.
        - La primera letra de cada sección se muestra en mayúscula.
        - Se añaden puntos guía entre el título y el número de página.
        - El número de página se alinea a la derecha.
        - Respeta el nivel jerárquico de cada sección.
        - Cada entrada es clicable y dirige a la sección correspondiente.

        IMPORTANTE:
        Esta función debe llamarse cuando ya se hayan registrado todas
        las secciones del documento.
        """
        y = self.crear_cabecera(ruta_logo= LOGO, codigo="TD01-03-A6",version="01",fecha="2024-04-22")

        ancho, alto = self.pdf._pagesize

        margen_izq = 20 * mm
        margen_der = 20 * mm

        y -= 30 * mm

        # ========================================================
        # TÍTULO: ÍNDICE / INDEX
        # ========================================================

        tamano_titulo = 18

        texto_es = "ÍNDICE / "
        texto_en = "INDEX"

        # Parte española en negro
        self.pdf.setFillColor(NEGRO)
        self.pdf.setFont(
            "Helvetica-Bold",
            tamano_titulo
        )

        self.pdf.drawString(
            margen_izq,
            y,
            texto_es
        )

        # Calcular dónde termina "ÍNDICE / "
        ancho_texto_es = self.pdf.stringWidth(
            texto_es,
            "Helvetica-Bold",
            tamano_titulo
        )

        # INDEX en azul
        self.pdf.setFillColor(AZUL)
        self.pdf.setFont(
            "Helvetica-BoldOblique",
            tamano_titulo
        )

        self.pdf.drawString(
            margen_izq + ancho_texto_es,
            y,
            texto_en
        )

        # Volver a negro
        self.pdf.setFillColor(NEGRO)

        y -= 15 * mm

        # ========================================================
        # SECCIONES
        # ========================================================

        for seccion in self.secciones:

            clave = str(seccion["clave"])

            titulo = str(
                seccion["titulo"]
            ).strip()

            nivel = int(
                seccion["nivel"]
            )

            pagina = str(
                seccion["pagina"]
            )

            # ----------------------------------------------------
            # Primera letra en mayúscula
            #
            # No usamos .capitalize(), porque convertiría
            # el resto del texto a minúsculas.
            # ----------------------------------------------------

            if titulo:
                titulo = (
                    titulo[0].upper()
                    + titulo[1:]
                )

            # ----------------------------------------------------
            # Sangría según nivel
            # ----------------------------------------------------

            sangria = nivel * 8 * mm

            x_titulo = (
                margen_izq
                + sangria
            )

            # ----------------------------------------------------
            # Fuente según nivel
            # ----------------------------------------------------

            if nivel == 0:

                fuente = "Helvetica-Bold"
                tamano = 11

            else:

                fuente = "Helvetica"
                tamano = 10

            self.pdf.setFont(
                fuente,
                tamano
            )

            self.pdf.setFillColor(
                NEGRO
            )

            # ====================================================
            # POSICIÓN DEL NÚMERO DE PÁGINA
            # ====================================================

            ancho_numero_pagina = (
                self.pdf.stringWidth(
                    pagina,
                    fuente,
                    tamano
                )
            )

            x_pagina = (
                ancho
                - margen_der
                - ancho_numero_pagina
            )

            # ====================================================
            # DIBUJAR TÍTULO
            # ====================================================

            self.pdf.drawString(
                x_titulo,
                y,
                titulo
            )

            # ====================================================
            # CALCULAR PUNTOS GUÍA
            # ====================================================

            ancho_titulo = (
                self.pdf.stringWidth(
                    titulo,
                    fuente,
                    tamano
                )
            )

            # Comienzo de los puntos
            x_inicio_puntos = (
                x_titulo
                + ancho_titulo
                + 2 * mm
            )

            # Fin de los puntos
            x_fin_puntos = (
                x_pagina
                - 2 * mm
            )

            ancho_disponible_puntos = (
                x_fin_puntos
                - x_inicio_puntos
            )

            if ancho_disponible_puntos > 0:

                ancho_punto = (
                    self.pdf.stringWidth(
                        ".",
                        fuente,
                        tamano
                    )
                )

                numero_puntos = int(
                    ancho_disponible_puntos
                    / ancho_punto
                )

                puntos = "." * numero_puntos

                self.pdf.drawString(
                    x_inicio_puntos,
                    y,
                    puntos
                )

            # ====================================================
            # NÚMERO DE PÁGINA
            # ====================================================

            self.pdf.drawString(
                x_pagina,
                y,
                pagina
            )

            # ====================================================
            # ENLACE NAVEGABLE
            # ====================================================

            self.pdf.linkAbsolute(
                contents=titulo,
                destinationname=clave,
                Rect=(
                    x_titulo,
                    y - 2 * mm,
                    ancho - margen_der,
                    y + 4 * mm
                ),
                thickness=0
            )

            # ----------------------------------------------------
            # Siguiente entrada
            # ----------------------------------------------------

            y -= 8 * mm

        # ========================================================
        # FINALIZAR PÁGINA
        # ========================================================

        self.nueva_pagina()

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
        self.marcar_seccion("PORTADA", "portada", nivel=0)
        
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
    
    def guardar(self):
        """
        Guarda el PDF generado y mueve la última página a la posición indicada, 
        reemplazando la página que actualmente ocupa esa posición (Ya reservada en blanco).
        """
        def mover_ultima_pagina_y_reemplazar(ruta_pdf,posicion):
            """
            Mueve la última página del PDF a la posición indicada,
            sustituyendo la página que actualmente ocupa esa posición.

            La página sustituida se elimina del documento.

            Parámetros
            ----------
            ruta_pdf : str
                Ruta del PDF que se desea modificar.

            posicion : int
                Número de página que se desea reemplazar.
                La numeración comienza en 1.

                Ejemplo:
                    posicion=2

                sustituye la página 2 por la última página del PDF.

            Ejemplo
            -------
            PDF inicial:

                Página 1 -> Portada
                Página 2 -> Página reservada
                Página 3 -> Producción
                Página 4 -> Ensayos
                Página 5 -> Índice

            Después de:

                mover_ultima_pagina_y_reemplazar(
                    "Batch_Record.pdf",
                    posicion=2
                )

            Resultado:

                Página 1 -> Portada
                Página 2 -> Índice
                Página 3 -> Producción
                Página 4 -> Ensayos
            """

            ruta_pdf = Path(ruta_pdf)

            if not ruta_pdf.exists():
                raise FileNotFoundError(
                    f"No existe el PDF: {ruta_pdf}"
                )

            reader = PdfReader(str(ruta_pdf))

            numero_paginas = len(reader.pages)

            if numero_paginas < 2:
                raise ValueError(
                    "El PDF debe contener al menos 2 páginas."
                )

            if posicion < 1 or posicion > numero_paginas:
                raise ValueError(
                    f"La posición debe estar entre 1 y {numero_paginas}."
                )

            # No tendría sentido reemplazar la última página
            # por ella misma
            if posicion == numero_paginas:
                return

            # --------------------------------------------------------
            # La última página es la que queremos mover
            # --------------------------------------------------------

            ultima_pagina = reader.pages[-1]

            writer = PdfWriter()

            # --------------------------------------------------------
            # Recorrer todas las páginas EXCEPTO la última
            # --------------------------------------------------------

            for indice in range(numero_paginas - 1):

                numero_pagina = indice + 1

                # Si llegamos a la página que queremos sustituir,
                # añadimos la última página en su lugar.
                if numero_pagina == posicion:

                    writer.add_page(
                        ultima_pagina
                    )

                else:

                    writer.add_page(
                        reader.pages[indice]
                    )

            # --------------------------------------------------------
            # Guardar en archivo temporal
            # --------------------------------------------------------

            ruta_temporal = ruta_pdf.with_name(
                ruta_pdf.stem
                + "_temp"
                + ruta_pdf.suffix
            )

            with open(
                ruta_temporal,
                "wb"
            ) as archivo:

                writer.write(
                    archivo
                )

            # --------------------------------------------------------
            # Sustituir PDF original
            # --------------------------------------------------------

            os.replace(
                ruta_temporal,
                ruta_pdf
            )

        self.pdf.save()
        mover_ultima_pagina_y_reemplazar(self.nombre_archivo, self.pagina_indice)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ FUNCIONES ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛



if __name__ == "__main__":

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

    bachredord.guardar()