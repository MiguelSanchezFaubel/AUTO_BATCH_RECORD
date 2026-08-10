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

        self.pdf = canvas.Canvas(nombre_archivo,pagesize=A4)
        self.nombre_archivo = nombre_archivo

    def crear_hoja_control_calidad(self, lote):

        c = self.pdf

        c.setFont("Helvetica-Bold", 16)

        c.drawString(
            50,
            800,
            "CONTROL DE CALIDAD"
        )

        c.drawString(
            50,
            770,
            f"Lote: {lote}"
        )

        c.showPage()

    def crear_portada(
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
        # 1. TÍTULO
        # ========================================================

        y = alto - 25 * mm

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
            rowHeights=6 * mm
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

        w, h = tabla_requisitos.wrap(
            ancho_util,
            80 * mm
        )

        tabla_requisitos.drawOn(
            self.pdf,
            margen_izq,
            y - h
        )

        # ========================================================
        # 5. BLOQUE DE FIRMAS
        # ========================================================

        y_firmas = 35 * mm

        altura_firma = 51 * mm
        altura_fecha = 7 * mm

        ancho_columna = ancho_util / 3

        # Coordenadas
        x1 = margen_izq
        x2 = margen_izq + ancho_columna
        x3 = margen_izq + ancho_columna * 2
        x4 = margen_izq + ancho_util

        y_inferior = y_firmas
        y_fecha = y_firmas + altura_fecha
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

        self.pdf.showPage()

    def guardar(self):

        self.pdf.save()



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

    bachredord.crear_portada(

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