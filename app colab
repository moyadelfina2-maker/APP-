import os
import re
import ipywidgets as widgets
from IPython.display import display, clear_output, HTML
from google.colab import files
import time # Import time for sleep

# Instalamos python-docx si no está presente
try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.text import WD_COLOR_INDEX # Import WD_COLOR_INDEX for highlighting
except ImportError:
    !pip install python-docx
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.text import WD_COLOR_INDEX # Import WD_COLOR_INDEX for highlighting

# ==========================================
# 2. VARIABLES GLOBALES Y CONFIGURACIÓN
# ==========================================
uploaded_filename = None
hallazgos_widgets_list = []
hallazgos_container = widgets.VBox([], layout=widgets.Layout(border='1px solid lightgray', padding='10px'))

# --- CONSTANTES DE ESTADOS PARA DROPDOWNS ---
CAAP_LOGICA_ESTADOS = [
    "No aplica", # Changed from "Seleccione..."
    "No iniciado, tiene CNCA vigente",
    "No iniciado, tiene CNCA en curso",
    "No iniciado, tiene CNCA vencida / No solicitada)",
    "En curso",
    "Vigente",
    "Vencido"
]
ADA_DETALLE_ESTADOS = ["Seleccione...", "Prefactibilidad con Chi 0 para todos los permisos", "Prefactibilidad vigente", "Prefactibilidad vencida", "No solicitada"]
RENPRE_DETALLE_ESTADOS = ["Seleccione...", "No aplica", "Esta inscripto y renueva", "Esta inscripto y no renovó", "Aplica pero no esta inscripto"]
FRECUENCIA_MUESTREO_OPTIONS = ["Seleccione...", "Anual", "Semestral", "Trimestral", "Mensual", "Continua", "N/A"]

# NUEVAS LISTAS PARA RESIDUOS Y CHE
RESIDUOS_ESPECIALES_STATUS_GENERAL = [
    "Seleccione...", "Empresa exenta porque no genera", "No inscripta", "No Cumple con las DDJJ",
    "Cumple con las DDJJ"
]
CHE_DETALLE_ESTADOS = [
    "Seleccione...", "Obtuvo CHE", "No obtuvo CHE", "CHE en curso"
]

# ACUMAR and SE States
ACUMAR_DETALLE_ESTADOS = ["Seleccione...", "No aplica", "Vigente", "En curso", "Vencida", "No solicitada"]
SE_DETALLE_ESTADOS = ["Seleccione...", "No aplica", "Vigente", "En curso", "Vencida", "No solicitada"]

# --- HALLAZGOS PREDEFINIDOS (ACTUALIZADOS CON EL CONTENIDO DEL ARCHIVO) ---
HALLAZGOS_PREDEFINIDOS = {
    'Ambiental': [
        {
            'observacion': 'Purgas de Aparatos a Presión (ASP) sin destino final identificado.',
            'situacion': 'Durante el relevamiento se observó que, respecto de los ASP, no se encontraba identificado el destino final de las purgas asociadas.',
            'autoridad': 'Ministerio de Ambiente de la Provincia de Buenos Aires (MAPBA).',
            'riesgo': 'Infracción durante una inspección del organismo de control.',
            'recomendacion': 'Se sugiere la verificación de la salida de las mismas, considerando que podrían contener mezclas de aceite y agua, residuo clasificado como especial según la normativa vigente.'
        },
        {
            'observacion': 'Derrame de líquido con presencia de hidrocarburos sobre suelo absorbente.',
            'situacion': 'Durante el recorrido se evidenció que hubo un derrame de líquido con presencia de hidrocarburos sobre suelo absorbente.',
            'autoridad': 'Autoridad del Agua (ADA) o Ministerio de Ambiente de la Provincia de Buenos Aires (MAPBA).',
            'riesgo': 'Infracción ante una inspección de ADA o bien, MAPBA por el derrame, en relación con el Art. 103 de la Ley 12.257 - Código de Aguas de la Provincia de Buenos Aires, o el incumplimiento de la Resolución 3722/16 que establece que se debe informar cualquier eventualidad en sus operaciones que pueda impactar en el ambiente o generar preocupación en la comunidad.',
            'recomendacion': 'Identificar la posible causa para tomar acciones con el fin de evitar este tipo de derrames o salpicaduras sobre suelo absorbente. En caso de ocurrir, se debe dar aviso ante las autoridades pertinentes.'
        },
        {
            'observacion': 'Almacenamiento Aéreo de Combustible (Sala Calderas/Red Incendio) no inscripto.',
            'situacion': 'Se constató que la empresa realiza almacenamiento de combustible en planta mediante sistemas aéreos, correspondientes al tanque de la sala de calderas y al tanque de bombas de la red de incendio. Sin embargo, dichos sistemas no se encuentran inscriptos.',
            'autoridad': 'Secretaría de Energía (SE).',
            'riesgo': 'Retrasar la emisión de permisos/habilitaciones si la autoridad detecta el tanque sin adecuar.',
            'recomendacion': 'Declarar los sistemas de almacenamiento ante la SE, inscribiéndolos en el Registro de Bocas de Expendio de Combustibles Líquidos (Res. 1102/04), para incorporarlos en las auditorías (Res. 404/94).'
        },
        {
            'observacion': 'Contenedor (bin) de 1000 Lts con combustible almacenado inadecuadamente.',
            'situacion': 'Se observó un contenedor (bin) de 1000 Lts con combustible almacenado transitoriamente, dispuesto inadecuadamente a la intemperie, sobre suelo absorbente y sin identificación.',
            'autoridad': 'Secretaría de Energía (SE).',
            'riesgo': 'Retrasar la emisión de permisos/habilitaciones si la autoridad detecta el tanque sin adecuar.',
            'recomendacion': 'Retirar de planta o, si se planea mantener, adecuar el sistema para luego proceder con su habilitación ante la SE. Si se usa un batán, revisar la normativa particular de transporte.'
        },
        {
            'observacion': 'Depósito de Residuos Especiales sin pasillos de separación de 1 metro.',
            'situacion': 'Algunos residuos almacenados en el depósito de residuos especiales no se encuentran separados por pasillos de 1 metro, impidiendo la visualización de los residuos posteriores.',
            'autoridad': 'Ministerio de Ambiente de la Provincia de Buenos Aires (MAPBA).',
            'riesgo': 'Infracción durante una inspección de MAPBA por incumplimiento de la Resolución 592/00.',
            'recomendacion': 'Ordenar los residuos para facilitar la verificación y contabilización ante una inspección. Adicionalmente, revisar que se cumpla con el etiquetado de la totalidad de residuos.'
        },
        {
            'observacion': 'Cuarto de lavado de piezas sucio, colapsado y con rejilla desbordada.',
            'situacion': 'El cuarto de lavado de piezas se observó sucio y colapsado, con la rejilla de la cámara de contención desbordada, lo que provocó el estancamiento de líquido contaminado y charcos.',
            'autoridad': 'Ministerio de Ambiente de la Provincia de Buenos Aires (MAPBA).',
            'riesgo': 'Infracción durante una inspección de MAPBA.',
            'recomendacion': 'Realizar acciones para evitar este tipo de desbordes y verificar que la capacidad de almacenamiento de la cámara sea la adecuada para los volúmenes generados.'
        },
        {
            'observacion': 'Análisis de transformador de vía húmeda realizado de manera no oficial.',
            'situacion': 'Respecto al transformador eléctrico de vía húmeda, el análisis realizado fue de manera no oficial.',
            'autoridad': 'Ministerio de Ambiente de la Provincia de Buenos Aires (MAPBA).',
            'riesgo': 'Infracción ante una inspección de la MAPBA.',
            'recomendacion': 'Se recomienda realizar un análisis de manera oficial con protocolo de informe y certificado de cadena de custodia oficial.'
        },
        {
            'observacion': 'Contenedores para residuos no especiales sin protección climática.',
            'situacion': 'El sector donde se almacenan los contenedores para residuos no especiales (para transporte y reciclaje) no cuenta con protección contra las inclemencias climáticas.',
            'autoridad': 'Ministerio de Ambiente (MDA).',
            'riesgo': 'Infracción ante una inspección de la MDA.',
            'recomendacion': 'Se recomienda colocar los volquetes en un sector con protección contra las lluvias o colocar volquetes con tapas.'
        },
        {
            'observacion': 'Bines y tambores sin identificar (falta de etiqueta de residuo/materia prima).',
            'situacion': 'Se tomó vista de una serie de bines y tambores sin identificar.',
            'autoridad': 'Ministerio de Ambiente (MDA).',
            'riesgo': 'Infracción ante una inspección de la MDA.',
            'recomendacion': 'Se recomienda identificar los bines y tambores observados definiendo si son residuos especiales (almacenar en depósito transitorio con etiquetas), materia prima (almacenar en depósito destinado para tal fin) o para devolución (definir y sectorizar un lugar).'
        },
        {
            'observacion': 'Residuos almacenados dentro del depósito de residuos especiales sin etiquetas identificatorias.',
            'situacion': 'Los residuos almacenados dentro del depósito de residuos especiales no contaban con etiquetas identificatorias.',
            'autoridad': 'Ministerio de Ambiente (MDA).',
            'riesgo': 'Infracción ante una inspección de la MDA.',
            'recomendacion': 'Se recomienda incorporar etiquetas que contengan fecha de ingreso, categoría (Y) y peligrosidad (H) en todos los residuos almacenados dentro del depósito.'
        },
        {
            'observacion': 'Efluentes industriales y de refrigeración sin separación/CAyTM incompleta (ACUMAR).',
            'situacion': 'El establecimiento no posee separación de los efluentes líquidos industriales y del proceso de refrigeración que permita evaluar la calidad previa a la CAyTM final, tal como lo solicita ACUMAR. Además, la CAyTM no cuenta con la placa para la clausura de vuelco.',
            'autoridad': 'Autoridad de Cuenca Matanza Riachuelo (ACUMAR) y Autoridad del Agua (ADA).',
            'riesgo': 'Infracción ante una inspección de la ACUMAR y ADA.',
            'recomendacion': 'Se recomienda evaluar la posibilidad de realizar 2 CAyTM (una para efluentes industriales y otra para efluentes de refrigeración) o enviar los efluentes de refrigeración a la PTEL e incorporar la placa para la clausura de vuelco.'
        },
        {
            'observacion': 'Uso de manguera para dilución en Planta de Tratamiento de Efluentes Líquidos (PTEL).',
            'situacion': 'En la PTEL se observó una manguera utilizada para incorporar agua, práctica considerada dilución del efluente y que está prohibida.',
            'autoridad': 'Autoridad del Agua (ADA).',
            'riesgo': 'Infracción por parte de la Autoridad del Agua.',
            'recomendacion': 'Quitar las mangueras que se utilicen para verter agua dentro de la PTEL.'
        },
        {
            'observacion': 'Sala de calderas sin detector de gas y monóxido de carbono.',
            'situacion': 'La sala de calderas no contaba con detector de gas y monóxido de carbono.',
            'autoridad': 'Ministerio de Ambiente.',
            'riesgo': 'Infracción ante una inspección del Ministerio de Ambiente.',
            'recomendacion': 'Se recomienda avanzar en la colocación del detector.'
        },
        {
            'observacion': 'Depósito de químicos de caldera (con contención y techo parciales) con envases fuera de la zona cubierta.',
            'situacion': 'El depósito de químicos de caldera (con contención y techo parciales) tenía envases almacenados sobre sectores donde el techo y la contención de derrames no cubrían.',
            'autoridad': 'Ministerio de Ambiente.',
            'riesgo': 'Infracción ante una inspección del Ministerio de Ambiente.',
            'recomendacion': 'Asegurar el almacenamiento en el sector adecuado del depósito o extender el techo y la contención para cubrir toda la superficie de la planta.'
        },
        {
            'observacion': 'Diámetro del Orificio Toma Muestra (OTM) de LEGA No Conforme (Res. 559/19 y Dec. 1074/18).',
            'situacion': 'Las adecuaciones implementadas en el Orificio Toma Muestra (OTM) para la LEGA no cumplen con los requisitos técnicos de las Res. 559/19 y Dec. 1074/18, ya que el diámetro de la instalación está por debajo del mínimo exigido.',
            'autoridad': 'Ministerio de Ambiente de la Provincia de Buenos Aires (MAPBA).',
            'riesgo': 'Infracción durante una inspección del organismo de control.',
            'recomendacion': 'Se recomienda proceder con la modificación del OTM a fin de garantizar el cumplimiento de las dimensiones mínimas estipuladas por la normativa.'
        },
        {
            'observacion': 'Presencia de nuevo pozo de explotación hídrica no declarado ante ADA.',
            'situacion': 'Se constató la presencia de un nuevo pozo de explotación hídrica no declarado formalmente ante la Autoridad del Agua (ADA). Esta captación no figura en los permisos de uso del recurso hídrico.',
            'autoridad': 'Autoridad del Agua (ADA).',
            'riesgo': 'Infracción por existencia de instalaciones no declaradas o por falta de condiciones del pozo.',
            'recomendacion': 'Declarar el nuevo pozo de explotación ante la ADA e incorporar un caudalímetro homologado para cumplir con los requerimientos de medición.'
        },
        {
            'observacion': 'Transformadores "Libres de PCBs" sin el análisis obligatorio actualizado.',
            'situacion': 'Se constató la correcta señalización de equipos identificados como "libres de PCBs".',
            'autoridad': 'Ministerio de Ambiente de la Provincia de Buenos Aires (MAPBA).',
            'riesgo': 'Ante auditorías, es obligatorio que haya al menos un análisis de PCBs de los transformadores.',
            'recomendacion': 'Si no se tiene un monitoreo, realizarlo. Si el monitoreo tiene una fecha mayor a 3 años, realizarlo nuevamente para conocer el estatus actual.'
        },
        {
            'observacion': 'Depósito de Residuos Especiales inaccesible por reubicación o fuera de norma.',
            'situacion': 'No fue posible acceder al depósito de residuos especiales debido a que se encontraba en proceso de reubicación, impidiendo verificar el cumplimiento de las disposiciones técnicas.',
            'autoridad': 'Ministerio de Ambiente de la Provincia de Buenos Aires (MAPBA).',
            'riesgo': 'Infracción por no poseer depósito de residuos especiales conforme a la Resolución 592/00.',
            'recomendacion': 'Debe realizarse de manera urgente la adecuación de un sector para el almacenamiento de residuos especiales conforme a la Resolución 592/00.'
        },
        {
            'observacion': 'Baldes con residuos con materia orgánica sin contención secundaria ni identificación.',
            'situacion': 'Presencia de baldes conteniendo residuos con materia orgánica, sin sistema de contención secundaria y sin identificación alguna.',
            'autoridad': 'Ministerio de Ambiente de la Provincia de Buenos Aires (MAPBA).',
            'riesgo': 'Llamado de atención de las autoridades, o derrames descontrolados que terminen en el exterior de la planta.',
            'recomendacion': 'Se recomienda realizar la adecuación para contención de derrames, protección contra inclemencias climáticas y piso impermeable sin conexión con el sistema de pluviales.'
        },
        {
            'observacion': 'Tanque contenedor de ácido con vertido directo al suelo sin contención.',
            'situacion': 'Tanque contenedor de ácido conectado a una manguera sin medidas de seguridad. El líquido era liberado directamente al suelo, sin contención secundaria ni sistemas de control de derrames.',
            'autoridad': 'Ministerio de Ambiente de la Provincia de Buenos Aires (MAPBA) y la Autoridad del Agua (ADA).',
            'riesgo': 'Infracción por eventualidades no declaradas (MAPBA) o por vertidos no declarados si se derivan al pluvial (ADA).',
            'recomendacion': 'Implementación de sistemas de contención de derrames en los puntos de carga y descarga, y un mejor guardado de la manguera.'
        },
        {
            'observacion': 'Falta de la cantidad mínima de foguistas habilitados.',
            'situacion': 'La planta no posee la cantidad mínima de foguistas habilitados.',
            'autoridad': 'Ministerio de Ambiente de la Provincia de Buenos Aires (MAPBA).',
            'riesgo': 'Infracción del MAPBA por incumplimiento al Art. 18 de la Res. 231/96 modificado por Art. 5 de la Res. 1126/07.',
            'recomendacion': 'Realizar la capacitación y habilitación correspondiente a operarios para cumplir con la cantidad mínima de foguistas de acuerdo a la cantidad de turnos.'
        },
        {
            'observacion': 'Sala de calderas sin protecciones y alarmas de detección automática de fuga de combustibles gaseosos y detectores de monóxido.',
            'situacion': 'En la sala de calderas no se evidenció la presencia de protecciones y alarmas de detección automática de fuga de combustibles gaseosos y detectores de monóxido.',
            'autoridad': 'Ministerio de Ambiente de la Provincia de Buenos Aires (MAPBA).',
            'riesgo': 'Infracción del MAPBA por incumplimiento al Art. 18 de la Res. 231/96 modificado por Art. 5 de la Res. 1126/07.',
            'recomendacion': 'Se recomienda realizar la instalación de los elementos de seguridad previamente mencionados.'
        },
        {
            'observacion': 'Sala de calderas sin libro de seguimiento foliado de generadores de vapor (Res. 1126/07).',
            'situacion': 'En la sala de calderas no se evidenció la presencia del libro de seguimiento foliado de generadores de vapor, acorde al Apéndice 3 de la Resolución 1126/07, en el que se asienten todos los controles realizados, reparaciones solicitadas y/o realizadas, y todas las anormalidades detectadas con indicación de la fecha respectiva.',
            'autoridad': 'Ministerio de Ambiente de la Provincia de Buenos Aires (MAPBA).',
            'riesgo': 'Infracción del MAPBA por incumplimiento de la Resolución 1126/07.',
            'recomendacion': 'Se recomienda confeccionar el libro rubricado y colocarlo en la sala de calderas.'
        },
        {
            'observacion': 'Residuos especiales acopiados fuera del depósito sin cobertura climática.',
            'situacion': 'Se observan residuos especiales acopiados fuera del depósito sin cobertura ante inclemencias climáticas.',
            'autoridad': 'Ministerio de Ambiente de la Provincia de Buenos Aires (MAPBA).',
            'riesgo': 'Infracción al Art. 3 inciso A de la Resolución 592/00.',
            'recomendacion': 'Almacenar los residuos especiales en el depósito transitorio con la cobertura y contención adecuadas, conforme a la Resolución 592/00.'
        }
    ],
    'Seguridad': [
        {'observacion': 'Falta de señalización en áreas de riesgo.', 'situacion': 'Áreas operativas sin cartelería.', 'autoridad': 'SRT / Ministerio Trabajo', 'riesgo': 'Seguridad Laboral', 'recomendacion': 'Instalar señalización IRAM.'}
    ]
}

# --- MARCADORES_CONDICIONALES (Esenciales para la lógica del documento) ---
MARCADORES_CONDICIONALES = {
    "HABILITACION_MUNICIPAL": {
        "cumple": {"start": "{INICIO_HAB_CUMPLE}", "end": "{FIN_HAB_CUMPLE}"}, "no cumple": {"start": "{INICIO_HAB_NO_CUMPLE}", "end": "{FIN_HAB_NO_CUMPLE}"},
        "parcial": {"start": "{INICIO_HAB_PARCIAL}", "end": "{FIN_HAB_PARCIAL}"}
    },
    "CNCA_STATUS": {
        "vigente": {"start": "{INICIO_CNCA_VIGENTE}", "end": "{FIN_CNCA_VIGENTE}"}, "superada (esta en curso el CAA)": {"start": "{INICIO_CNCA_SUPERADA}", "end": "{FIN_CNCA_SUPERADA}"},
        "vencida": {"start": "{INICIO_CNCA_VENCIDA}", "end": "{FIN_CNCA_VENCIDA}"}, "no solicitada": {"start": "{INICIO_CNCA_NO_SOLICITADA}", "end": "{FIN_CNCA_NO_SOLICITADA}"},
        "en curso": {"start": "{INICIO_CNCA_EN_CURSO}", "end": "{FIN_CNCA_EN_CURSO}"}
    },
    "CAAP_STATUS": {
        "No iniciado, tiene CNCA vigente": {"start": "{INICIO_CAAP_NO_INICIADO_CNCA_VIGENTE}", "end": "{FIN_CAAP_NO_INICIADO_CNCA_VIGENTE}"},
        "No iniciado, tiene CNCA en curso": {"start": "{INICIO_CAAP_NO_INICIADO_CNCA_ENCURSO}", "end": "{FIN_CAAP_NO_INICIADO_CNCA_ENCURSO}"},
        "No iniciado, tiene CNCA vencida / No solicitada)": {"start": "{INICIO_CAAP_NO_INICIADO_CNCA_VENCIDA_NO_SOLICITADA}", "end": "{FIN_CAAP_NO_INICIADO_CNCA_VENCIDA_NO_SOLICITADA}"},
        "En curso": {"start": "{INICIO_CAAP_EN_CURSO}", "end": "{FIN_CAAP_EN_CURSO}"},
        "Vigente": {"start": "{INICIO_CAAP_VIGENTE}", "end": "{FIN_CAAP_VIGENTE}"},
        "Vencido": {"start": "{INICIO_CAAP_VENCIDO}", "end": "{FIN_CAAP_VENCIDO}"}
    },
    "CAAF_STATUS": {
        "en_curso": {"start": "{INICIO_CAAF_EN CURSO}", "end": "{FIN_CAAF_EN_CURSO}"}, "vigente": {"start": "{INICIO_CAAF_VIGENTE}", "end": "{FIN_CAAF_VIGENTE}"},
        "no_iniciado_caap_en_curso": {"start": "{INICIO_CAAF_NO_INICIADO_CAAP_EN CURSO}", "end": "{FIN_CAAF_NO_INICIADO_CAAP_EN_CURSO}"},
        "no_iniciado_caap_vencido": {"start": "{INICIO_CAAF_NO_INICIADO_CAAP_VENCIDO}", "end": "{FIN_CAAF_NO_INICIADO_CAAP_VENCIDO}"},
        "vencido": {"start": "{INICIO_CAAF_VENCIDO}", "end": "{FIN_CAAF_VENCIDO}"},
        "eliminar_todo": {"start": "{ELIMINAR_TODO_CAAF}", "end": "{FIN_ELIMINAR_TODO_CAAF}"}
    },
    "RENOVACION_CAA_STATUS": {
        "En curso": {"start": "{INICIO_RENOVACION_CAA_ENCURSO}", "end": "{FIN_RENOVACION_CAA_ENCURSO}"}
    },
    "LEGA_STATUS": {
        "vigente": {"start": "{INICIO_LEGA_VIGENTE}", "end": "{FIN_LEGA_VIGENTE}"}, "en_curso": {"start": "{INICIO_LEGA_EN_CURSO}", "end": "{FIN_LEGA_EN_CURSO}"},
        "vencida": {"start": "{INICIO_LEGA_VENCIDA}", "end": "{FIN_LEGA_VENCIDA}"}
    },
    "RESIDUOS_ESPECIALES_STATUS": {
        "Empresa exenta porque no genera": {"start": "{INICIO_EMPRESA_EXCENTA}", "end": "{FIN_EMPRESA_EXCENTA}"},
        "No inscripta": {"start": "{INICIO_RREE_NO_INSCRIPTA}", "end": "{FIN_RREE_NO_INSCRIPTA}"},
        "No Cumple con las DDJJ": {"start": "{INICIO_RREE_NOCUMPLE_DDJJ}", "end": "{FIN_RREE_NOCUMPLE_DDJJ}"},
        "Cumple con las DDJJ": {"start": "{INICIO_RREE_CUMPLE_DDJJ}", "end": "{FIN_RREE_CUMPLE_DDJJ}"}
    },
    "CHE_STATUS": {
        "Obtuvo CHE": {"start": "{INICIO_RREE_OBTUVO_CHE}", "end": "{FIN_RREE_OBTUVO_CHE}"},
        "No obtuvo CHE": {"start": "{INICIO_RREE_NO_OBTUVO_NO_SOLICITO_CHE}", "end": "{FIN_RREE_NO_OBTUVO_NO_SOLICITO_CHE}"},
        "CHE en curso": {"start": "{INICIO_RREE_CHE_EN_CURSO}", "end": "{FIN_RREE_CHE_EN_CURSO}"}
    },
    "RESIDUOS_GESTION": {
        "Correcta": {"start": "{INICIO_CORRECTA_GESTION_RESIDUOS}", "end": "{FIN_CORRECTA_GESTION_RESIDUOS}"}, "Mala": {"start": "{INICIO_MAL_GESTION_RESIDUOS}", "end": "{FIN_MAL_GESTION_RESIDUOS}"}
    },
    "ASP_STATUS": {
        "Finalizada": {"start": "{INICIO_PRESENTACION_ASP_FINALIZADA}", "end": "{FIN_PRESENTACION_ASP_FINALIZADA}"},
        "Caratulada": {"start": "{INICIO_PRESENTACION_ASP_CARATULADA}", "end": "{FIN_PRESENTACION_ASP_CARATULADA}"},
        "No Presentado": {"start": "{INICIO_ASP_NO_PRESENTADO}", "end": "{FIN_ASP_NO_PRESENTADO}"}
    },
    "VALVULAS_CALIBRACION_STATUS": {
        "Cumple": {"start": "{INICIO_CALIBRACION_ASP_CUMPLE}", "end": "{FIN_CALIBRACION_ASP_CUMPLE}"},
        "No Cumple": {"start": "{INICIO_CALIBRACION_ASP_NOCUMPLE}", "end": "{FIN_CALIBRACION_ASP_NOCUMPLE}"}
    },
    "ADA_STATUS": {
         "Prefactibilidad con Chi 0 para todos los permisos": {"start": "{PREFA_TODOS_CHI0}", "end": "{FIN_PREFA_TODOS_CHI0}"},
         "Prefactibilidad vigente": {"start": "{PREFA_OBTENIDA}", "end": "{FIN_PREFA_OBTENIDA}"},
         "Prefactibilidad vencida": {"start": "{PREFA_VENCIDA}", "end": "{FIN_PREFA_VENCIDA}"},
         "No solicitada": {"start": "{PREFA_NO_SOLICITADA}", "end": "{FIN_PREFA_NO_SOLICITADA}"}
    },
    "RENPRE_STATUS": {
        "No aplica": {"start": "{INICIO_NO_APLICA}", "end": "{FIN_NO_APLICA}"},
        "Esta inscripto y renueva": {"start": "{INICIO_APLICA_INSCRIPTO_RENUEVA}", "end": "{FIN_APLICA_INSCRIPTO_RENUEVA}"},
        "Esta inscripto y no renovó": {"start": "{INICIO_APLICA_INSCRIPTO_NO_RENOVO}", "end": "{FIN_APLICA_NO_RENOVO}"},
        "Aplica pero no esta inscripto": {"start": "{INICIO_APLICA_NO_INSCRIPTO}", "end": "{FIN_APLICA_NO_INSCRIPTO}"}
    },
    "SEGURO_STATUS": {
        "Vigente": {"start": "{INICIO_POLIZA_VIGENTE}", "end": "{FIN_POLIZA_VIGENTE}"}, "Vencida": {"start": "{INICIO_POLIZA_VENCIDA}", "end": "{FIN_POLIZA_VENCIDA}"},
        "Nunca Tuvo": {"start": "{INICIO_NUNCA_TUVO_POLIZA}", "end": "{FIN_NUNCA_TUVO_POLIZA}"}
    },
    "ACUMAR_STATUS": {
        "Vigente": {"start": "{INICIO_ACUMAR_VIGENTE}", "end": "{FIN_ACUMAR_VIGENTE}"},
        "En curso": {"start": "{INICIO_ACUMAR_EN_CURSO}", "end": "{FIN_ACUMAR_EN_CURSO}"},
        "Vencida": {"start": "{INICIO_ACUMAR_VENCIDA}", "end": "{FIN_ACUMAR_VENCIDA}"},
        "No solicitada": {"start": "{INICIO_ACUMAR_NO_SOLICITADA}", "end": "{FIN_ACUMAR_NO_SOLICITADA}"},
        "No aplica": {"start": "{INICIO_ACUMAR_NO_APLICA}", "end": "{FIN_ACUMAR_NO_APLICA}"}
    },
    "INSCRIPCION_1102": {
        "Inscripto": {"start": "{INICIO_INSCRIPTA_1102}", "end": "{FIN_INSCRIPTA_1102}"},
        "No inscripto": {"start": "{INICIO_NOINSCRIPTA_1102}", "end": "{FIN_NOINSCRIPTA_1102}"},
        "No aplica": {"start": "{INICIO_NOAPLICA_1102}", "end": "{FIN_NOAPLICA_1102}"}

    },
    "AUDITORIA_404": {
        "Realizo": {"start": "{INICIO_REALIZO_AUDITORIA}", "end": "{FIN_REALIZO_AUDITORIA}"},
        "No realizo": {"start": "{INICIO_NO_REALIZO_AUDITORIA}", "end": "{FIN_NO_REALIZO_AUDITORIA}"},
        "No aplica": {"start": "{INICIO_NOAPLICA_AUDITORIA}", "end": "{FIN_NOAPICA_AUDITORIA}"},
      "No inscripto, no realiza": {"start": "{INICIO_NO_INSCRIPTA_NO_AUDITORIA}", "end": "{FIN_NO_INSCRIPTA_NO_AUDITORIA}"}
    },
    "INSCRIPCION_277": {
        "Inscripto": {"start": "{INICIO_INSCRIPTA_277}", "end": "{FIN_INSCRIPTA_277}"},
        "No inscripto": {"start": "{INICIO_NOINSCRIPTA_277}", "end": "{FIN_NOINSCRIPTA_277}"},
        "No aplica": {"start": "{INICIO_NOAPLICA_277}", "end": "{FIN_NOAPLICA_277}"}

    }
}

# ==========================================
# 3. FUNCIONES DE CARGA Y PROCESAMIENTO DOCX
# ==========================================

def upload_template(b):
    global uploaded_filename
    with output_text:
        clear_output()
        print("Sube tu plantilla DOCX.")
        try:
            upload = files.upload()
            if upload and len(upload) > 0: # Check if upload is a non-empty dictionary
                uploaded_filename = list(upload.keys())[0]
                print(f"✅ Plantilla '{uploaded_filename}' cargada.")
            else: # This block will be executed if upload is None, {}, or any falsy value
                print("❌ No se seleccionó ninguna plantilla o la carga fue cancelada/fallida.")
                uploaded_filename = None # Ensure uploaded_filename is cleared if upload fails
        except Exception as e:
            print(f"❌ Error al cargar el archivo: {e}")
            print("Por favor, asegúrate de seleccionar un archivo DOCX y de que la interfaz de carga no se haya cerrado inesperadamente.")
            uploaded_filename = None

def agregar_hallazgo_formateado_al_doc(doc, index, observacion, situacion, autoridad, riesgo, recomendacion):
    if not (observacion.strip() or situacion.strip()): return

    # Add the formatted title
    p_title = doc.add_paragraph()
    run_title = p_title.add_run(f"Observación de campo # {index}")
    run_title.bold = True
    run_title.underline = True
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT

    fields = [
        ("Observación:", observacion), ("Situación:", situacion), ("Autoridad:", autoridad),
        ("Riesgo:", riesgo), ("Recomendación:", recomendacion)
    ]

    # Add a separator
    doc.add_paragraph()
    # p_sep = doc.add_paragraph(f"--- Hallazgo ---") # Commented out this line to remove the subtitle
    # p_sep.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for label, value in fields:
        if value.strip():
            p = doc.add_paragraph() # Removed style='List Bullet'
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run_label = p.add_run(label + " ")
            run_label.bold = True
            p.add_run(value)

def find_paragraphs_to_remove(doc, selected_state, situation_type):
    paragraphs_to_remove = []
    markers_config = MARCADORES_CONDICIONALES.get(situation_type, {})

    # Obtenemos los textos de los marcadores tal cual están en tu config
    selected_start_marker_text = markers_config.get(selected_state, {}).get('start')

    all_start_markers = {cfg.get('start') for cfg in markers_config.values() if cfg.get('start')}
    all_end_markers = {cfg.get('end') for cfg in markers_config.values() if cfg.get('end')}

    # Este es el único interruptor que necesitamos
    in_unselected_section = False

    for p in doc.paragraphs:
        # Limpiamos espacios en blanco al inicio/final del párrafo
        text = p.text.strip()

        # 1. SEGURIDAD: Si el párrafo está vacío, solo borrarlo si estamos dentro de una zona no deseada
        if not text:
            if in_unselected_section:
                paragraphs_to_remove.append(p)
            continue

        # 2. DETECCIÓN DE INICIOS
        if text in all_start_markers:
            paragraphs_to_remove.append(p) # El marcador siempre se elimina

            # Si el marcador que encontramos NO es el que el usuario eligió:
            if text != selected_start_marker_text:
                in_unselected_section = True
            else:
                # Si es el elegido, nos aseguramos de NO borrar lo que viene
                in_unselected_section = False
            continue

        # 3. DETECCIÓN DE FINALES
        if text in all_end_markers:
            paragraphs_to_remove.append(p) # El marcador siempre se elimina
            in_unselected_section = False # DETENER EL BORRADO. Esto protege las secciones comunes.
            continue

        # 4. LÓGICA DE CONTENIDO
        if in_unselected_section:
            # Solo si el interruptor está activo, agregamos el párrafo para borrar
            paragraphs_to_remove.append(p)

        # Si in_unselected_section es False, el código no hace nada y el párrafo se mantiene.

    return paragraphs_to_remove
def reemplazar_marcadores(doc, user_data):
    def process_container_for_replacements_and_highlights(container):
        paragraphs_to_iterate = []
        if hasattr(container, 'paragraphs'): # It's a table cell
            paragraphs_to_iterate = container.paragraphs
        else: # It's a paragraph
            paragraphs_to_iterate = [container]

        for p in paragraphs_to_iterate:
            if not p.runs: # Skip empty paragraphs
                continue

            # Store original run formats of the first run as a fallback
            first_run_format = {}
            if p.runs:
                run = p.runs[0]
                first_run_format = {
                    'bold': run.bold,
                    'italic': run.italic,
                    'underline': run.underline,
                    'font_name': run.font.name,
                    'font_size': run.font.size,
                    'font_color_rgb': run.font.color.rgb if run.font.color else None
                }

            combined_text = "".join([run.text for run in p.runs])
            modified_text = combined_text
            replaced_in_paragraph = False

            for marker, value in user_data.items():
                placeholder_pattern = re.compile(r'{\s*' + re.escape(marker) + r'\s*}')

                if placeholder_pattern.search(modified_text):
                    if str(value).strip() in ['N/A', '0', '']:
                        modified_text = placeholder_pattern.sub('', modified_text)
                        replaced_in_paragraph = True
                    elif str(value).strip():
                        modified_text = placeholder_pattern.sub(str(value), modified_text)
                        replaced_in_paragraph = True

            if replaced_in_paragraph:
                # Clear all existing runs
                for i in range(len(p.runs) -1, -1, -1): # Iterate backwards to safely delete
                    p.runs[i]._element.getparent().remove(p.runs[i]._element)
                # Add a single new run with the modified text
                new_run = p.add_run(modified_text)
                # Apply format of the first original run to the new run
                if first_run_format:
                    new_run.bold = first_run_format['bold']
                    new_run.italic = first_run_format['italic']
                    new_run.underline = first_run_format['underline']
                    if first_run_format['font_name']: new_run.font.name = first_run_format['font_name']
                    if first_run_format['font_size']: new_run.font.size = first_run_format['font_size']
                    if first_run_format['font_color_rgb']: new_run.font.color.rgb = first_run_format['font_color_rgb']

            # After all replacements (or if no replacement happened), check for remaining placeholders and highlight
            # This loop now operates on the potentially new single run, or the original runs.
            for run in p.runs:
                if re.search(r'{\s*[A-Z_]+\s*}', run.text):
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW

    for p in doc.paragraphs:
        process_container_for_replacements_and_highlights(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                process_container_for_replacements_and_highlights(cell)

# ==========================================
# 4. CREACIÓN DE WIDGETS
# ==========================================

def create_input_widget(key, label_text, default_value="", widget_type='Text', dropdown_options=None, disabled=False):
    style = widgets.Layout(width='65%')
    label_layout = widgets.Layout(width='30%')

    if widget_type == 'Text':
        input_widget = widgets.Text(value=default_value, disabled=disabled, layout=style)
    elif widget_type == 'Dropdown':
        input_widget = widgets.Dropdown(options=dropdown_options, value=dropdown_options[0], disabled=disabled, layout=style)
    elif widget_type == 'Textarea':
         input_widget = widgets.Textarea(value=default_value, disabled=disabled, layout=style)

    label = widgets.Label(label_text, layout=label_layout)
    # Retorna el HBox (para las secciones normales) y el widget interno (para la tabla de muestreo)
    hbox = widgets.HBox([label, input_widget], layout=widgets.Layout(width='100%', margin='2px'))
    hbox.add_class('custom-font-widget') # Add a custom class here
    return hbox, input_widget

# Función auxiliar para obtener el widget interno (Input/Dropdown)
def get_input_widget(w_container): return w_container.children[1]


# --- DEFINICIÓN DE WIDGETS ---
main_widgets = {
    'NOMBRE_EMPRESA': create_input_widget('NOMBRE_EMPRESA', 'Razón Social:', 'EMPRESA S.A.')[0],
    'NOMBRE_PLANTA': create_input_widget('NOMBRE_PLANTA', 'Planta Industrial:', 'Planta Industrial')[0],
    'MES_AUDITORIA': create_input_widget(key='MES_AUDITORIA', label_text='Mes de relevamiento (Mes/Año):', default_value='Octubre 2025')[0],
    'DIRECCION': create_input_widget('DIRECCION_EMPRESA', 'Dirección:', '')[0],
    'MUNICIPIO': create_input_widget('MUNICIPIO_EMPRESA', 'Municipio:', '')[0],
    'RUBRO': create_input_widget('RUBRO_EMPRESA', 'Rubro:', '')[0]
}
hab_widgets = {
    'STATUS': create_input_widget('HABILITACION_MUNICIPAL_STATUS', 'Estado Habilitación:', widget_type='Dropdown', dropdown_options=['Seleccione...', 'cumple', 'no cumple', 'parcial'])[0],
    'FECHA': create_input_widget('FECHA_HABILITACIÓN', 'Fecha de obtencion de habilitación:', 'dd/mm/aaaa')[0],
    'EXPEDIENTE': create_input_widget('EXPEDIENTE_HABILITACION', 'Nº de expediente:', 'N/A')[0],
    'OBSERVACION_HAB_MUNICIPAL': create_input_widget('OBSERVACION_HAB_MUNICIPAL', 'Observaciones extra habilitacion municipal:', '', widget_type='Textarea')[0]
}
cnca_widgets = {
    'STATUS': create_input_widget('CNCA_DETALLE_STATUS', 'Estado CNCA:', widget_type='Dropdown', dropdown_options=["No aplica", "vigente", "superada (esta en curso el CAA)", "vencida", "no solicitada", "en curso"])[0],
    'FECHA': create_input_widget('FECHA_CNCA', 'Fecha de obtencion de CNCA:', 'dd/mm/aaaa')[0],
    'VENCIMIENTO': create_input_widget('FECHA_CNCA_VENCIMIENTO', 'Vencimiento de CNCA:', 'dd/mm/aaaa')[0],
    'EXPEDIENTE': create_input_widget('EXPEDIENTE_CNCA', 'Expediente de la CNCA:', 'N/A')[0],
    'CATEGORIA': create_input_widget('CATEGORIA_CNCA', 'Categoría:', 'primera/segunda/tercera')[0],
    'PUNTOS': create_input_widget('PUNTOS_CNCA', 'Puntos (solo numero):', 'ej:25')[0],
    'DISPO_CNCA': create_input_widget('DISPO_CNCA', 'Disposición de CNCA:', 'N/A')[0],
    'OBSERVACIONES_CNCA': create_input_widget('OBSERVACIONES_CNCA', 'Observaciones extra CNCA:', '', widget_type='Textarea')[0]

}
caap_caaf_widgets = {
    'CAAP_LOGICA': create_input_widget('CAAP_STATUS', 'Situación CAAP (Fase II):', widget_type='Dropdown', dropdown_options=CAAP_LOGICA_ESTADOS)[0],
    'CAAF_LOGICA': create_input_widget('CAAF_STATUS', 'Situación CAAF (Fase III):', widget_type='Dropdown', dropdown_options=["No aplica", "no_iniciado_caap_en_curso", "no_iniciado_caap_vencido", "en_curso", "vigente", "vencido"])[0],
    'RENOVACION_CAA_STATUS': create_input_widget('RENOVACION_CAA_STATUS', 'Estado Renovación CAA:', widget_type='Dropdown', dropdown_options=["No aplica", "En curso", "Finalizada", "No iniciada"])[0],
    'FECHA_CAAP': create_input_widget('FECHA_Obtención_CAAP', 'Fecha de obtencion del CAAP:', 'N/A')[0],
    'EXP_CAAP': create_input_widget('EXPEDIENTE_CAAP', 'Expediente del CAAP:', 'N/A')[0],
    'DISPO_CAAP': create_input_widget('DISPO_CAAP', 'Disposicion del CAAP:', 'N/A')[0],
    'VIGENCIA_CAAP': create_input_widget('CAAP_plazo de vigencia', 'Plazo de vigencia (años):', 'N/A')[0],
    'VTO_CAAP': create_input_widget('FECHA_Vencimiento_CAAP', 'Vencimiento del CAAP:', 'N/A')[0],
    'ESTADO_PORTAL_CAAP': create_input_widget('Estado_CAAP', 'Estado del CAAP en el portal:', 'N/A')[0],
    'OBSERVACIONES_4': create_input_widget('OBSERVACIONES_4', 'Observaciones extra CAAP:', '', widget_type='Textarea')[0],
    'FECHA_CAAF': create_input_widget('FECHA_Obtención_CAAF', 'Fecha de obtencion del CAAF:', 'N/A')[0],
    'EXPEDIENTE_CAAF': create_input_widget('EXPEDIENTE_CAAF', 'Expediente del CAAF:', 'N/A')[0],
    'DISPO_CAAF': create_input_widget('DISPO_CAAF', 'Disposicion de CAAF:', 'N/A')[0],
    'ESTADO_PORTAL_CAAF': create_input_widget('Estado_CAAF', 'Estado del CAAF en el portal:', 'N/A')[0],
    'EXPEDIENTE_RENOVACION_CAA': create_input_widget('EXPEDIENTE_RENOVACION_CAA', 'Expediente Renovación CAA:', 'N/A')[0],
    'ESTADO_PORTAL_RENOVACION_CAA': create_input_widget('ESTADO_PORTAL_RENOVACION_CAA', 'Estado de renovacion del CAA en el portal:', 'N/A')[0],
    'DISPO_RENOVACION_CAA': create_input_widget('DISPO_RENOVACION_CAA', 'Disposición Renovación CAA:', 'N/A')[0],
    'OBSERVACIONES_CAA': create_input_widget('OBSERVACIONES_CAA', 'Observaciones extra CAAP/CAAF:', '', widget_type='Textarea')[0]
}
lega_widgets = {
    'STATUS': create_input_widget('LEGA_STATUS', 'Estado LEGA:', widget_type='Dropdown', dropdown_options=["Seleccione...", "vigente", "en_curso", "vencida"])[0],
    'FECHA': create_input_widget('{FECHA_OBTENCIÓN_LEGA', 'Fecha de obtencion de la LEGA:', 'N/A')[0],
    'EXPEDIENTE': create_input_widget('EXPEDIENTE_LEGA', 'Expediente de la LEGA:', 'N/A')[0],
    'ESTADO_PORTAL': create_input_widget('Estado_LEGA', 'Estado de la LEGA en el portal:', 'N/A')[0],
    'VTO_LEGA': create_input_widget('VENCIMIENTO_LEGA', 'Vencimiento de la LEGA:', 'N/A')[0],
     'OBSERVACIONES_LEGA': create_input_widget('OBSERVACIONES_LEGA', 'Observaciones LEGA:', 'N/A')[0], # New field
}

rree_widgets = {
    'STATUS': create_input_widget('RESIDUOS_ESPECIALES_STATUS', 'Estado Generador (General):', widget_type='Dropdown', dropdown_options=RESIDUOS_ESPECIALES_STATUS_GENERAL)[0],
    'CHE_STATUS': create_input_widget('CHE_STATUS', 'Estado CHE:', widget_type='Dropdown', dropdown_options=CHE_DETALLE_ESTADOS)[0], # Nuevo dropdown para CHE
    'ANIO_CHE': create_input_widget('AÑO_CHE', 'Año de obtencion del CHE:', '2025')[0],
    'OBS_CHE': create_input_widget('OBSERVACIONES_TICKETS_CONSULTA_CHE', 'Obs. CHE:', '', widget_type='Textarea')[0],
    # Separación lógica del estado legal (RREE) de la gestión operativa (RGNL)
    'GESTION_RES': create_input_widget('GESTION_RESIDUOS_STATUS', 'Gestión operativa de residuos (general):', widget_type='Dropdown', dropdown_options=["Correcta", "Mala"])[0],
    'TIPO_RES': create_input_widget('TIPO_RESIDUO', 'Tipo Residuos (Ej: solidos/humedos/urbanos):', '')[0],
    'OBS_EXTRA': create_input_widget('OBSERVACION_EXTRA_RESIDUOS', 'Observaciones extra residuos:', '', widget_type='Textarea')[0]
}
asp_widgets = {
    'STATUS': create_input_widget('ASP_STATUS', 'Estado ASP:', widget_type='Dropdown', dropdown_options=["Seleccione...", "Finalizada", "Caratulada", "No Presentado"])[0],
    'VTO_ASP': create_input_widget('VENCIMIENTO_ASP', 'Vencimiento de presentacion de ASP:', 'N/A')[0],
    'EXPEDIENTE_ASP': create_input_widget('EXPEDIENTE_ASP', 'Expediente del ASP:', 'N/A')[0], # New field
    'OBS_ASP': create_input_widget('OBSERVACIONES_EXTRA_ASP', 'Obs. ASP:', '', widget_type='Textarea')[0],
    'VALVULAS_STATUS': create_input_widget('VALVULAS_STATUS', 'Calibracion de válvulas de seguridad:', widget_type='Dropdown', dropdown_options=["Cumple", "No Cumple"])[0],
    'VTO_VALVULAS': create_input_widget('VENCIMIENTO_CALIBRACION_ASP', 'Vencimiento de válvulas:', 'N/A')[0]
}
otros_widgets = {
    'ADA_STATUS': create_input_widget('ADA_STATUS', 'Estado ADA:', widget_type='Dropdown', dropdown_options=ADA_DETALLE_ESTADOS)[0],
    'ADA_FECHA': create_input_widget('FECHA_PREFA', 'Fecha de obtencion de Prefactibilidad:', 'N/A')[0],
    'ADA_EXP': create_input_widget('EXPEDIENTE_PREFA', 'Expediente de Prefactibilidad:', 'N/A')[0],
    'CHI_HID': create_input_widget('NCHI_HIDRAULICA', 'CHi Hidráulica:', '0/1/2/3')[0],
    'CHI_EXP': create_input_widget('NCHI_EXPLOTACION', 'CHi Explotación:', '0/1/2/3')[0],
    'CHI_VUE': create_input_widget('NCHI_VUELCO', 'CHi Vuelco:', '0/1/2/3')[0],
    # Campos de texto simples para ADA
    'ESTADO_PERMISO_HIDRAULICA': create_input_widget('ESTADO_PERMISO_HIDRAULICA', 'Estado Permiso Hidráulica:', 'N/A')[0],
    'ESTADO_PERMISO_EXPLOTACION': create_input_widget('ESTADO_PERMISO_EXPLOTACION', 'Estado Permiso Explotación:', 'N/A')[0],
    'ESTADO_PERMISO_VUELCO': create_input_widget('ESTADO_PERMISO_VUELCO', 'Estado Permiso Vuelco:', 'N/A')[0],
    'RED_MONITOREOS': create_input_widget('RED_MONITOREOS', 'Red de Monitoreos:', 'N/A')[0],
    'RENPRE_STATUS': create_input_widget('RENPRE_STATUS', 'Estado RENPRE:', widget_type='Dropdown', dropdown_options=RENPRE_DETALLE_ESTADOS)[0],
    'RENPRE_NUM': create_input_widget('NUMERO_RENPRE', 'Numero de operador:', 'N/A')[0],
    'SEGURO_STATUS': create_input_widget('SEGURO_STATUS', 'Seguro ambiental:', widget_type='Dropdown', dropdown_options=["Vigente", "Vencida", "Nunca Tuvo"])[0],
    'POLIZA_NUM': create_input_widget('NUMERO_POLIZA', 'Nº Póliza:', 'N/A')[0],
    'POLIZA_VTO': create_input_widget('VTO_POLIZA', 'Vencimeinto de póliza:', 'N/A')[0],
    'ACUMAR_STATUS': create_input_widget('ACUMAR_STATUS', 'Estado ACUMAR:', widget_type='Dropdown', dropdown_options=ACUMAR_DETALLE_ESTADOS)[0],
    'ACUMAR_EXP': create_input_widget('ACUMAR_EXPEDIENTE', 'Expediente ACUMAR:', 'N/A')[0],
    'ACUMAR_OBS': create_input_widget('ACUMAR_OBSERVACIONES', 'Observaciones ACUMAR:', '', widget_type='Textarea')[0],
    'SE_STATUS': create_input_widget('SE_STATUS', 'Estado Sec. Energía:', widget_type='Dropdown', dropdown_options=SE_DETALLE_ESTADOS)[0],
    'SE_EXP': create_input_widget('SE_EXPEDIENTE', 'Expediente Sec. Energía:', 'N/A')[0],
    'INSCRIPCION_1102': create_input_widget('INSCRIPCION_1102', 'Inscripcion 1102/04:', widget_type='Dropdown', dropdown_options=["Inscripto", "No inscripto", "No aplica"])[0],
    'NUMERO_SE': create_input_widget('NUMERO_SE', 'Numero de operador:', 'N/A')[0],
    'NUMERO_EXP': create_input_widget('NUMERO_EXP', 'Numero de expediente:', 'N/A')[0],
    'CANTIDAD_DE_TANQUES': create_input_widget('CANTIDAD_DE_TANQUES', 'Numero de tanques:', 'N/A')[0],
    'AUDITORIA_404': create_input_widget('AUDITORIA_404', 'Auditoria de seguridad 404/94:', widget_type='Dropdown', dropdown_options=["Realizo", "No realizo", "No aplica","No inscripto, no realiza"])[0],
    'VENCIMIENTO_AUDITORIA404': create_input_widget('VENCIMIENTO_AUDITORIA404', 'Vencimiento de la auditoria 404/94:', 'N/A')[0],
    'INSCRIPCION_277': create_input_widget('INSCRIPCION_277', 'Inscripcion 277/25:', widget_type='Dropdown', dropdown_options=["Inscripto", "No inscripto", "No aplica"])[0],
    'OBSERVACIONES_277': create_input_widget('OBSERVACIONES_277', 'Observaciones 277/25:', '', widget_type='Textarea')[0], # New widget
    'SE_OBS': create_input_widget('SE_OBSERVACIONES', 'Observaciones Sec. Energía:', '', widget_type='Textarea')[0]
}

from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def generar_tabla_desde_interfaz_dinamica(doc, lista_filas):
    target_text = "Plan de monitoreos"
    encontrado = False # Bandera para verificar si se encontró el texto

    for p in doc.paragraphs:
        if target_text in p.text:
            encontrado = True
            # Creamos la tabla de 5 columnas
            headers = ['Recurso', 'Organismo', 'Puntos', 'Parámetros', 'Frecuencia']
            table = doc.add_table(rows=1, cols=5)
            table.style = 'Table Grid'

            # Encabezado
            hdr_cells = table.rows[0].cells
            for i, name in enumerate(headers):
                hdr_cells[i].text = name
                run = hdr_cells[i].paragraphs[0].runs[0]
                run.bold = True
                hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Agregamos las filas que creaste en la interfaz
            for fila in lista_filas:
                # Solo agregamos si el campo Recurso no está vacío
                if fila['recurso'].value.strip():
                    row = table.add_row().cells
                    row[0].text = fila['recurso'].value
                    row[1].text = fila['organismo'].value
                    row[2].text = fila['puntos'].value
                    row[3].text = fila['parametros'].value

                    frec = fila['frecuencia'].value
                    row[4].text = frec if frec != 'Seleccione...' else ""

                    # Formato de fuente
                    for cell in row:
                        for para in cell.paragraphs:
                            for run in para.runs:
                                run.font.size = Pt(10)
            break

    if not encontrado:
        print(f"No se encontró el texto '{target_text}' en el documento.")

# --- MODO DE USO ---
# insertar_tabla_dinamica_en_doc(mi_documento, filas_widgets)

# ==========================================
# 5. LÓGICA CONDICIONAL (Habilitar/Deshabilitar Campos)
# ==========================================

def setup_conditional_fields(all_widgets):
    # Helper to get the actual input widget from the HBox structure: HBox([Label, InputWidget])
    def get_input(widgets, key):
        return widgets[key].children[1]

    # --- HABILITACIÓN MUNICIPAL (REGLA 1: Solo 'cumple' habilita detalles) ---
    status_hab = get_input(all_widgets['hab'], 'STATUS')
    details_hab = [get_input(all_widgets['hab'], k) for k in ['FECHA', 'EXPEDIENTE', 'OBSERVACION_HAB_MUNICIPAL']]

    def update_hab_fields(change):
        is_active = True # Always active
        for w in details_hab:
            w.disabled = not is_active
    status_hab.observe(update_hab_fields, names='value')
    update_hab_fields({'new': status_hab.value})

    # --- CNCA (REGLA 2: Lógica condicional en Observaciones) ---
    status_cnca = get_input(all_widgets['cnca'], 'STATUS')
    details_cnca = [get_input(all_widgets['cnca'], k) for k in ['FECHA', 'VENCIMIENTO', 'EXPEDIENTE', 'CATEGORIA', 'PUNTOS', 'DISPO_CNCA']]
    obs2 = get_input(all_widgets['cnca'], 'OBSERVACIONES_CNCA')

    def update_cnca_fields(change):
        is_details_active = True # Always active
        for w in details_cnca:
            w.disabled = not is_details_active

        obs2.disabled = False # Always active

    status_cnca.observe(update_cnca_fields, names='value')
    update_cnca_fields({'new': status_cnca.value})

    # --- CAAP/CAAF (REGLA 3: Solo 'Vigente' habilita detalles de CAAP) ---
    status_caap = get_input(all_widgets['caap_caaf'], 'CAAP_LOGICA')
    details_caap = [get_input(all_widgets['caap_caaf'], k) for k in ['FECHA_CAAP', 'EXP_CAAP', 'DISPO_CAAP', 'VIGENCIA_CAAP', 'VTO_CAAP', 'ESTADO_PORTAL_CAAP']]
    status_caaf = get_input(all_widgets['caap_caaf'], 'CAAF_LOGICA')
    details_caaf = [get_input(all_widgets['caap_caaf'], k) for k in ['FECHA_CAAF', 'EXPEDIENTE_CAAF', 'DISPO_CAAF', 'ESTADO_PORTAL_CAAF']]

    # New Renovacion CAA fields
    status_renovacion_caa = get_input(all_widgets['caap_caaf'], 'RENOVACION_CAA_STATUS')
    details_renovacion_caa = [get_input(all_widgets['caap_caaf'], k) for k in ['EXPEDIENTE_RENOVACION_CAA', 'ESTADO_PORTAL_RENOVACION_CAA', 'DISPO_RENOVACION_CAA']]

    def update_caap_fields(change):
        is_active = True # Always active
        for w in details_caap:
            w.disabled = not is_active

    def update_caaf_fields(change):
        is_active = True # Always active
        for w in details_caaf: w.disabled = not is_active

    def update_renovacion_caa_fields(change):
        is_active = True # Always active
        for w in details_renovacion_caa:
            w.disabled = not is_active

    status_caap.observe(update_caap_fields, names='value')
    update_caap_fields({'new': status_caap.value})
    status_caaf.observe(update_caaf_fields, names='value')
    update_caaf_fields({'new': status_caaf.value})
    status_renovacion_caa.observe(update_renovacion_caa_fields, names='value')
    update_renovacion_caa_fields({'new': status_renovacion_caa.value})

    # --- LEGA (REGLA 4: Solo 'vigente' o 'en_curso' habilita detalles) ---
    status_lega = get_input(all_widgets['lega'], 'STATUS')
    details_lega = [get_input(all_widgets['lega'], k) for k in ['FECHA', 'EXPEDIENTE', 'ESTADO_PORTAL', 'VTO_LEGA']]

    def update_lega_fields(change):
        is_active = True # Always active
        for w in details_lega:
            w.disabled = not is_active

    status_lega.observe(update_lega_fields, names='value')
    update_lega_fields({'new': status_lega.value})

    # --- RESIDUOS ESPECIALES (Generador) y CHE---
    # General RREE status dropdown (no longer controls CHE fields)
    status_rree_general = get_input(all_widgets['rree'], 'STATUS')
    # The general RREE status doesn't have associated fields to enable/disable based on its selection alone anymore.

    # New CHE Status dropdown and its associated fields
    status_che = get_input(all_widgets['rree'], 'CHE_STATUS')
    details_che = [get_input(all_widgets['rree'], k) for k in ['ANIO_CHE', 'OBS_CHE']]

    def update_che_fields(change):
        is_active = True # Always active
        for w in details_che:
            w.disabled = not is_active

    status_che.observe(update_che_fields, names='value')
    update_che_fields({'new': status_che.value})


    # --- ASP ---
    status_asp = get_input(all_widgets['asp'], 'STATUS')
    details_asp = [get_input(all_widgets['asp'], k) for k in ['VTO_ASP', 'EXPEDIENTE_ASP']]
    status_valvulas = get_input(all_widgets['asp'], 'VALVULAS_STATUS') # New: Status for valves
    vto_valvulas = get_input(all_widgets['asp'], 'VTO_VALVULAS') # New: Due date for valves

    def update_valvulas_fields(change):
        is_active_valvulas = True # Always active
        vto_valvulas.disabled = not is_active_valvulas

    def update_asp_fields(change):
        is_active = True # Always active
        for w in details_asp: w.disabled = not is_active

    status_asp.observe(update_asp_fields, names='value')
    update_asp_fields({'new': status_asp.value})
    status_valvulas.observe(update_valvulas_fields, names='value') # Observe VALVULAS_STATUS
    update_valvulas_fields({'new': status_valvulas.value}) # Initial call for VALVULAS_STATUS

    # --- ADA (Recursos Hídricos) ---
    status_ada = get_input(all_widgets['otros'], 'ADA_STATUS')
    details_ada_text = [get_input(all_widgets['otros'], k) for k in ['ADA_FECHA', 'ADA_EXP', 'ESTADO_PERMISO_HIDRAULICA', 'ESTADO_PERMISO_EXPLOTACION', 'ESTADO_PERMISO_VUELCO', 'RED_MONITOREOS']]
    details_ada_chi = [get_input(all_widgets['otros'], k) for k in ['CHI_HID', 'CHI_EXP', 'CHI_VUE']]

    def update_ada_fields(change):
        is_active = True # Always active
        for w in details_ada_text: w.disabled = not is_active
        for w in details_ada_chi: w.disabled = not is_active

    status_ada.observe(update_ada_fields, names='value')
    update_ada_fields({'new': status_ada.value})

    # --- RENPRE ---
    status_renpre = get_input(all_widgets['otros'], 'RENPRE_STATUS')
    details_renpre = [get_input(all_widgets['otros'], 'RENPRE_NUM')]

    def update_renpre_fields(change):
        is_active = True # Always active
        for w in details_renpre: w.disabled = not is_active
    status_renpre.observe(update_renpre_fields, names='value')
    update_renpre_fields({'new': status_renpre.value})

    # --- SEGURO AMBIENTAL ---
    status_seguro = get_input(all_widgets['otros'], 'SEGURO_STATUS')
    details_seguro = [get_input(all_widgets['otros'], k) for k in ['POLIZA_NUM', 'POLIZA_VTO']]

    def update_seguro_fields(change):
        is_active = True # Always active
        for w in details_seguro: w.disabled = not is_active
    status_seguro.observe(update_seguro_fields, names='value')
    update_seguro_fields({'new': status_seguro.value})

    # --- ACUMAR ---
    status_acumar = get_input(all_widgets['otros'], 'ACUMAR_STATUS')
    details_acumar = [get_input(all_widgets['otros'], k) for k in ['ACUMAR_EXP', 'ACUMAR_OBS']]

    def update_acumar_fields(change):
        is_active = True # Always active
        for w in details_acumar: w.disabled = not is_active
    status_acumar.observe(update_acumar_fields, names='value')
    update_acumar_fields({'new': status_acumar.value})

    # --- SECRETARÍA DE ENERGÍA ---
    # General SE Status and Exp.
    status_se = get_input(all_widgets['otros'], 'SE_STATUS')
    details_se = [get_input(all_widgets['otros'], k) for k in ['SE_EXP', 'SE_OBS']]
    def update_se_general_fields(change):
        is_active = change['new'] in ["Vigente", "En curso", "Vencida"] # Active if not "No aplica" or "No solicitada"
        for w in details_se: w.disabled = not is_active
    status_se.observe(update_se_general_fields, names='value')
    update_se_general_fields({'new': status_se.value})

    # INSCRIPCION_1102
    status_1102 = get_input(all_widgets['otros'], 'INSCRIPCION_1102')
    details_1102 = [get_input(all_widgets['otros'], k) for k in ['NUMERO_SE']]
    def update_1102_fields(change):
        is_active = change['new'] == "Inscripto" # Only active if 'Inscripto'
        for w in details_1102: w.disabled = not is_active
    status_1102.observe(update_1102_fields, names='value')
    update_1102_fields({'new': status_1102.value})

    # INSCRIPCION_277
    status_277 = get_input(all_widgets['otros'], 'INSCRIPCION_277')
    details_277 = [get_input(all_widgets['otros'], k) for k in ['OBSERVACIONES_277']]
    def update_277_fields(change):
        is_active = change['new'] == "Inscripto" # Only active if 'Inscripto'
        for w in details_277: w.disabled = not is_active
    status_277.observe(update_277_fields, names='value')
    update_277_fields({'new': status_277.value})

    # AUDITORIA_404
    status_404 = get_input(all_widgets['otros'], 'AUDITORIA_404')
    details_404 = [get_input(all_widgets['otros'], k) for k in ['CANTIDAD_DE_TANQUES', 'VENCIMIENTO_AUDITORIA404']]
    def update_404_fields(change):
        is_active = change['new'] in ["Realizo", "No realizo"] # Active if either Realizo or No realizo
        for w in details_404: w.disabled = not is_active
    status_404.observe(update_404_fields, names='value')
    update_404_fields({'new': status_404.value})

# ==========================================
# 6. LÓGICA HALLAZGOS (DINÁMICA)
# ==========================================
def update_hallazgos_display():
    global hallazgos_widgets_list, hallazgos_container
    hallazgos_vbox_children = []

    for i, h_dict in enumerate(hallazgos_widgets_list):
        h_dict['titulo'].value = f"<b><U>Observación de campo # {i + 1}</U></b>"
        # Usar lambda con argumento default para asegurar el índice correcto
        # Remove existing click handlers to prevent multiple calls
        h_dict['delete_button'].on_click(lambda b, idx=i: on_delete_hallazgo(b, idx), remove=True)
        h_dict['delete_button'].on_click(lambda b, idx=i: on_delete_hallazgo(b, idx))

        hallazgo_box = widgets.VBox([
            h_dict['header'], h_dict['dropdown'], h_dict['obs'], h_dict['sit'],
            h_dict['aut'], h_dict['rie'], h_dict['rec'], widgets.HTML("<hr>")
        ], layout=widgets.Layout(border='1px solid #ccc', padding='10px', margin='5px 0'))

        hallazgos_vbox_children.append(hallazgo_box)
    hallazgos_container.children = tuple(hallazgos_vbox_children)


def on_delete_hallazgo(b, index):
    if len(hallazgos_widgets_list) > 0:
        hallazgos_widgets_list.pop(index)
        update_hallazgos_display()

def on_add_hallazgo(b):
    idx = len(hallazgos_widgets_list) + 1

    title = widgets.HTML(f"<b><U>Observación de campo # {idx}</U></b>")
    btn_del = widgets.Button(description="Eliminar", button_style='danger', icon='trash', layout=widgets.Layout(width='auto'))
    header = widgets.HBox([title, btn_del], layout=widgets.Layout(justify_content='space-between'))

    opts = ["Autocompletar (Seleccione)..."]
    for cat, items in HALLAZGOS_PREDEFINIDOS.items():
        for item in items:
            opts.append(f"[{cat}] {item['observacion'][:80]}") # Incrementado para mejor identificación

    dd = widgets.Dropdown(options=opts, description="Preset:", layout=widgets.Layout(width='98%'))

    # Se extrae solo el widget de entrada para la lógica del hallazgo
    w_obs = create_input_widget('H_OBS', 'Observación:', '', widget_type='Textarea')[1]
    w_sit = create_input_widget('H_SIT', 'Situación:', '')[1]
    w_aut = create_input_widget('H_AUT', 'Autoridad:', '')[1]
    w_rie = create_input_widget('H_RIE', 'Riesgo:', '')[1]
    w_rec = create_input_widget('H_REC', 'Recomendación:', '', widget_type='Textarea')[1]

    # Contenedores para la visualización dentro del VBox
    obs_box = widgets.HBox([widgets.Label('Observación:', layout=widgets.Layout(width='30%')), w_obs])
    sit_box = widgets.HBox([widgets.Label('Situación:', layout=widgets.Layout(width='30%')), w_sit])
    aut_box = widgets.HBox([widgets.Label('Autoridad:', layout=widgets.Layout(width='30%')), w_aut])
    rie_box = widgets.HBox([widgets.Label('Riesgo:', layout=widgets.Layout(width='30%')), w_rie])
    rec_box = widgets.HBox([widgets.Label('Recomendación:', layout=widgets.Layout(width='30%')), w_rec])

    # Callback autocompletar
    def on_change_dd(change):
        val = change['new']
        if "Autocompletar" in val: return
        # Buscar en todos los predefinidos
        for cat, items in HALLAZGOS_PREDEFINIDOS.items():
            for item in items:
                # Reconstruct the string as it appears in the dropdown options
                dropdown_representation = f"[{cat}] {item['observacion'][:80]}"
                if dropdown_representation == val:
                    w_obs.value = item['observacion']
                    w_sit.value = item['situacion']
                    w_aut.value = item['autoridad']
                    w_rie.value = item['riesgo']
                    w_rec.value = item['recomendacion']
                    return
    dd.observe(on_change_dd, names='value')

    h_dict = {
        'header': header, 'delete_button': btn_del, 'dropdown': dd, 'titulo': title,
        'obs': obs_box, 'sit': sit_box, 'aut': aut_box, 'rie': rie_box, 'rec': rec_box
    }

    hallazgos_widgets_list.append(h_dict)
    update_hallazgos_display()

btn_add_hallazgo = widgets.Button(description="Añadir Hallazgo", button_style='success', icon='plus')
btn_add_hallazgo.on_click(on_add_hallazgo)
on_add_hallazgo(None) # Añadir uno inicial

# ==========================================
# 7. GENERACIÓN DEL INFORME (LÓGICA PRINCIPAL)
# ==========================================

def on_generar_click(b):
    with output_text:
        clear_output(wait=True) # Ensure output is cleared completely before processing
        if not uploaded_filename:
            print("❌ Carga primero la plantilla.")
            return

        # Disable the button to prevent multiple clicks
        btn_gen.disabled = True

        try:
            print("⚙️ Generando informe...")
            doc = Document(uploaded_filename)

            # Función auxiliar para extraer valor de HBox (todos los widgets son HBox), excepto los de muestreo que se extraen del map directo
            def get_val(w_container): return w_container.children[1].value

            # 1. Recolección de datos
            user_data = {
                'RAZON_SOCIAL': get_val(main_widgets['NOMBRE_EMPRESA']), 'NOMBRE_EMPRESA': get_val(main_widgets['NOMBRE_EMPRESA']),
                'NOMBRE_PLANTA': get_val(main_widgets['NOMBRE_PLANTA']), 'MES_AUDITORIA': get_val(main_widgets['MES_AUDITORIA']),
                'DIRECCION_EMPRESA': get_val(main_widgets['DIRECCION']), 'MUNICIPIO_EMPRESA': get_val(main_widgets['MUNICIPIO']), 'RUBRO_EMPRESA': get_val(main_widgets['RUBRO']),
                'FECHA_HABILITACIÓN': get_val(hab_widgets['FECHA']), 'EXPEDIENTE_HABILITACION': get_val(hab_widgets['EXPEDIENTE']), 'OBSERVACION_HAB_MUNICIPAL': get_val(hab_widgets['OBSERVACION_HAB_MUNICIPAL']),
                'FECHA_CNCA': get_val(cnca_widgets['FECHA']), 'EXPEDIENTE_CNCA': get_val(cnca_widgets['EXPEDIENTE']), 'CATEGORIA_CNCA': get_val(cnca_widgets['CATEGORIA']),
                'PUNTOS_CNCA': get_val(cnca_widgets['PUNTOS']), 'FECHA_CNCA_VENCIMIENTO': get_val(cnca_widgets['VENCIMIENTO']),
                'DISPO_CNCA': get_val(cnca_widgets['DISPO_CNCA']),
                'OBSERVACIONES_CNCA': get_val(cnca_widgets['OBSERVACIONES_CNCA']),
                'FECHA_OBTENCION_CAAP': get_val(caap_caaf_widgets['FECHA_CAAP']),
                'EXPEDIENTE_CAAP': get_val(caap_caaf_widgets['EXP_CAAP']),
                'DISPO_CAAP': get_val(caap_caaf_widgets['DISPO_CAAP']),
                'VIGENCIA_CAAP': get_val(caap_caaf_widgets['VIGENCIA_CAAP']),
                'VENCIMIENTO_CAAP': get_val(caap_caaf_widgets['VTO_CAAP']),
                'ESTADO_PORTAL_CAAP': get_val(caap_caaf_widgets['ESTADO_PORTAL_CAAP']),
                'OBSERVACIONES_4': get_val(caap_caaf_widgets['OBSERVACIONES_4']),
                'FECHA_CAAF': get_val(caap_caaf_widgets['FECHA_CAAF']),
                'EXPEDIENTE_CAAF': get_val(caap_caaf_widgets['EXPEDIENTE_CAAF']),
                'DISPO_CAAF': get_val(caap_caaf_widgets['DISPO_CAAF']),
                'ESTADO_CAAF': get_val(caap_caaf_widgets['ESTADO_PORTAL_CAAF']),
                'OBSERVACIONES_CAA': get_val(caap_caaf_widgets['OBSERVACIONES_CAA']),
                'ESTADO_RENOVACION_CAA': get_val(caap_caaf_widgets['RENOVACION_CAA_STATUS']),
                'EXPEDIENTE_RENOVACION_CAA': get_val(caap_caaf_widgets['EXPEDIENTE_RENOVACION_CAA']),
                'ESTADO_PORTAL_RENOVACION_CAA': get_val(caap_caaf_widgets['ESTADO_PORTAL_RENOVACION_CAA']),
                'DISPO_RENOVACION_CAA': get_val(caap_caaf_widgets['DISPO_RENOVACION_CAA']),
                'FECHA_OBTENCION_LEGA': get_val(lega_widgets['FECHA']),
                'EXPEDIENTE_LEGA': get_val(lega_widgets['EXPEDIENTE']),
                'ESTADO_LEGA': get_val(lega_widgets['ESTADO_PORTAL']),
                'VENCIMIENTO_LEGA': get_val(lega_widgets['VTO_LEGA']), # New field

                # Actualización de la recolección de datos para RREE y CHE
                'RESIDUOS_ESPECIALES_STATUS': get_val(rree_widgets['STATUS']), # General RREE status
                'CHE_STATUS': get_val(rree_widgets['CHE_STATUS']), # New CHE status
                'AÑO_CHE': get_val(rree_widgets['ANIO_CHE']), 'OBSERVACIONES_TICKETS_CONSULTA_CHE': get_val(rree_widgets['OBS_CHE']),

                'TIPO_RESIDUO': get_val(rree_widgets['TIPO_RES']), 'OBSERVACION_EXTRA_RESIDUOS': get_val(rree_widgets['OBS_EXTRA']),
                'VENCIMIENTO_ASP': get_val(asp_widgets['VTO_ASP']),
                'EXPEDIENTE_ASP': get_val(asp_widgets['EXPEDIENTE_ASP']), # New field
                'OBSERVACIONES_EXTRA_ASP': get_val(asp_widgets['OBS_ASP']),
                'VENCIMIENTO_CALIBRACION_ASP': get_val(asp_widgets['VTO_VALVULAS']),
                'FECHA_PREFA': get_val(otros_widgets['ADA_FECHA']), 'EXPEDIENTE_PREFA': get_val(otros_widgets['ADA_EXP']),
                'NCHI_HIDRAULICA': get_val(otros_widgets['CHI_HID']), 'NCHI_EXPLOTACION': get_val(otros_widgets['CHI_EXP']),
                'NCHI_VUELCO': get_val(otros_widgets['CHI_VUE']),
                'ESTADO_PERMISO_HIDRAULICA': get_val(otros_widgets['ESTADO_PERMISO_HIDRAULICA']),
                'ESTADO_PERMISO_EXPLOTACION': get_val(otros_widgets['ESTADO_PERMISO_EXPLOTACION']),
                'ESTADO_PERMISO_VUELCO': get_val(otros_widgets['ESTADO_PERMISO_VUELCO']),
                'RED_MONITOREOS': get_val(otros_widgets['RED_MONITOREOS']),
                'NUMERO_RENPRE': get_val(otros_widgets['RENPRE_NUM']),
                'NUMERO_POLIZA': get_val(otros_widgets['POLIZA_NUM']), 'VTO_POLIZA': get_val(otros_widgets['POLIZA_VTO']),
                'ACUMAR_STATUS': get_val(otros_widgets['ACUMAR_STATUS']), # New ACUMAR Status
                'ACUMAR_EXPEDIENTE': get_val(otros_widgets['ACUMAR_EXP']), # New ACUMAR Expediente
                'ACUMAR_OBSERVACIONES': get_val(otros_widgets['ACUMAR_OBS']), # New ACUMAR Observaciones
                'SE_STATUS': get_val(otros_widgets['SE_STATUS']),
                'SE_EXPEDIENTE': get_val(otros_widgets['SE_EXP']), # New SE Expediente
                'INSCRIPCION_1102': get_val(otros_widgets['INSCRIPCION_1102']),
                'NUMERO_SE': get_val(otros_widgets['NUMERO_SE']),
                'CANTIDAD_DE_TANQUES': get_val(otros_widgets['CANTIDAD_DE_TANQUES']),
                'AUDITORIA_404': get_val(otros_widgets['AUDITORIA_404']),
                'VENCIMIENTO_AUDITORIA404': get_val(otros_widgets['VENCIMIENTO_AUDITORIA404']),
                'INSCRIPCION_277': get_val(otros_widgets['INSCRIPCION_277']),
                'OBSERVACIONES_277': get_val(otros_widgets['OBSERVACIONES_277']), # New field
                'SE_OBSERVACIONES': get_val(otros_widgets['SE_OBS'])

            }

            # Función auxiliar para la limpieza de párrafos
            def find_and_remove(doc, key, section):
                 for p in find_paragraphs_to_remove(doc, key, section): remove_paragraph(p)

            # 2. Procesar Lógica Condicional (Borrar bloques)
            s_hab = get_val(hab_widgets['STATUS']); find_and_remove(doc, s_hab, "HABILITACION_MUNICIPAL")
            s_cnca = get_val(cnca_widgets['STATUS']); find_and_remove(doc, s_cnca, "CNCA_STATUS")

            # CAAP (Lógica compleja mapeada ahora simplificada ya que las claves de MARCADORES_CONDICIONALES coinciden con los valores del dropdown)
            s_caap = get_val(caap_caaf_widgets['CAAP_LOGICA']);
            find_and_remove(doc, s_caap, "CAAP_STATUS")

            # CAAF
            s_caaf = get_val(caap_caaf_widgets['CAAF_LOGICA']);
            if s_caaf == "No aplica": s_caaf = "eliminar_todo"
            find_and_remove(doc, s_caaf, "CAAF_STATUS")

            # Renovacion CAA
            s_renovacion_caa = get_val(caap_caaf_widgets['RENOVACION_CAA_STATUS'])
            find_and_remove(doc, s_renovacion_caa, "RENOVACION_CAA_STATUS")

            # Resto de Permisos Condicionales
            find_and_remove(doc, get_val(lega_widgets['STATUS']), "LEGA_STATUS")

            # Manejo de los dos estados separados para RREE y CHE
            find_and_remove(doc, get_val(rree_widgets['STATUS']), "RESIDUOS_ESPECIALES_STATUS")
            find_and_remove(doc, get_val(rree_widgets['CHE_STATUS']), "CHE_STATUS") # Nuevo bloque condicional para CHE

            find_and_remove(doc, get_val(rree_widgets['GESTION_RES']), "RESIDUOS_GESTION")

            # Procesamiento de ASP y Calibración de Válvulas por separado
            find_and_remove(doc, get_val(asp_widgets['STATUS']), "ASP_STATUS")
            find_and_remove(doc, get_val(asp_widgets['VALVULAS_STATUS']), "VALVULAS_CALIBRACION_STATUS")

            find_and_remove(doc, get_val(otros_widgets['ADA_STATUS']), "ADA_STATUS")
            find_and_remove(doc, get_val(otros_widgets['RENPRE_STATUS']), "RENPRE_STATUS")
            find_and_remove(doc, get_val(otros_widgets['SEGURO_STATUS']), "SEGURO_STATUS")

            # New conditional blocks for ACUMAR and SE
            find_and_remove(doc, get_val(otros_widgets['ACUMAR_STATUS']), "ACUMAR_STATUS")
            find_and_remove(doc, get_val(otros_widgets['SE_STATUS']), "SE_STATUS")
            find_and_remove(doc, get_val(otros_widgets['INSCRIPCION_1102']), "INSCRIPCION_1102")
            find_and_remove(doc, get_val(otros_widgets['AUDITORIA_404']), "AUDITORIA_404")
            find_and_remove(doc, get_val(otros_widgets['INSCRIPCION_277']), "INSCRIPCION_277")

            # 3. Reemplazo de Variables (Textos)
            print("📝 Reemplazando textos...")
            reemplazar_marcadores(doc, user_data)

            # 4. Insertar Hallazgos
            print("🔍 Insertando hallazgos...")
            for i, h in enumerate(hallazgos_widgets_list):
                 obs_val = h['obs'].children[1].value
                 sit_val = h['sit'].children[1].value
                 aut_val = h['aut'].children[1].value
                 rie_val = h['rie'].children[1].value
                 rec_val = h['rec'].children[1].value
                 agregar_hallazgo_formateado_al_doc(doc, i + 1, obs_val, sit_val, aut_val, rie_val, rec_val)

            # 5. Insertar la tabla dinámica de muestreo
            print("📊 Insertando tabla de monitoreo...")
            generar_tabla_desde_interfaz_dinamica(doc, muestreo_filas_datos)

            # 6. Guardar
            out_file = f"Informe_{user_data['RAZON_SOCIAL']}.docx"
            doc.save(out_file)
            print(f"✅ Informe generado: {out_file}")

            files.download(out_file)
            # Removed the print statement after download to keep it clean

        except Exception as e:
            print(f"❌ Error crítico durante la generación del informe: {e}")
            import traceback
            traceback.print_exc()
            print("\nPor favor, revisa los mensajes de error anteriores para más detalles.")
            print("Asegúrate de que la plantilla DOCX esté correctamente formateada y que no haya valores inesperados en los campos.")
        finally:
            # Re-enable the button regardless of success or failure
            btn_gen.disabled = False

# ==========================================
# 8. INTERFAZ FINAL (LAYOUT DE PESTAÑAS)
# ==========================================

# Custom CSS for font standardization
custom_css = """
<style>
.custom-font-widget label,
.custom-font-widget input[type='text'],
.custom-font-widget textarea,
.custom-font-widget .widget-dropdown > select {
    font-family: 'Arial', sans-serif !important;
    font-size: 14px !important;
}
.custom-font-widget .widget-html h3, .custom-font-widget .widget-html h4 {
    font-family: 'Arial', sans-serif !important;
}
</style>
"""

# 1. Datos Generales y Habilitación
general_hab_box = widgets.VBox([
    widgets.HTML("<h3>📍 Información Básica</h3>"),
    widgets.VBox(list(main_widgets.values())),
    widgets.HTML("<h3>✅ Habilitación Municipal</h3>"),
    widgets.VBox(list(hab_widgets.values()))
], layout=widgets.Layout(padding='10px')).add_class('custom-font-widget')

# 2. Permisos Centrales (CNCA, CAAP, CAAF)
cnca_group = widgets.VBox([
    cnca_widgets['STATUS'],
    cnca_widgets['FECHA'],
    cnca_widgets['VENCIMIENTO'],
    cnca_widgets['EXPEDIENTE'],
    cnca_widgets['CATEGORIA'],
    cnca_widgets['PUNTOS'],
    cnca_widgets['DISPO_CNCA'],
    cnca_widgets['OBSERVACIONES_CNCA']
])

caap_caaf_group = widgets.VBox([
    caap_caaf_widgets['CAAP_LOGICA'],
    caap_caaf_widgets['FECHA_CAAP'],
    caap_caaf_widgets['EXP_CAAP'],
    caap_caaf_widgets['DISPO_CAAP'],
    caap_caaf_widgets['VIGENCIA_CAAP'],
    caap_caaf_widgets['VTO_CAAP'],
    caap_caaf_widgets['ESTADO_PORTAL_CAAP'],
    caap_caaf_widgets['OBSERVACIONES_4']
])

caaf_details_group = widgets.VBox([
    caap_caaf_widgets['CAAF_LOGICA'],
    caap_caaf_widgets['FECHA_CAAF'],
    caap_caaf_widgets['EXPEDIENTE_CAAF'],
    caap_caaf_widgets['DISPO_CAAF'],
    caap_caaf_widgets['ESTADO_PORTAL_CAAF']
])

renovacion_caa_group = widgets.VBox([
    caap_caaf_widgets['RENOVACION_CAA_STATUS'],
    caap_caaf_widgets['EXPEDIENTE_RENOVACION_CAA'],
    caap_caaf_widgets['ESTADO_PORTAL_RENOVACION_CAA'],
    caap_caaf_widgets['DISPO_RENOVACION_CAA'],
    caap_caaf_widgets['OBSERVACIONES_CAA']
])

permisos_opds_box = widgets.VBox([
    widgets.HTML("<h3>📄 CNCA (Clasificación Nivel)</h3>"),
    cnca_group,
    widgets.HTML("<h3>🌳 CAAP (Aptitud Ambiental - Fase II)</h3>"),
    caap_caaf_group,
    widgets.HTML("<h3>🌳 CAAF (Aptitud Ambiental - Fase III)</h3>"),
    caaf_details_group,
    widgets.HTML("<h3>🔄 Renovación CAA</h3>"),
    renovacion_caa_group
], layout=widgets.Layout(padding='10px')).add_class('custom-font-widget')

from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def insertar_tabla_manual_dinamica(doc, lista_de_filas_widgets):
    """
    Crea la tabla basándose UNICAMENTE en lo que el usuario
    escribió en la interfaz, sin nombres predefinidos.
    """
    target_text = "Plan de monitoreos"
    encontrado = False # Bandera para verificar si se encontró el texto

    for p in doc.paragraphs:
        if target_text in p.text:
            encontrado = True
            # Creamos la tabla de 5 columnas (Recurso, Organismo, Puntos, Parámetros, Frecuencia)
            table = doc.add_table(rows=1, cols=5)
            table.style = 'Table Grid'

            # 1. Encabezados
            headers = ['Recurso', 'Organismo', 'Puntos', 'Parámetros', 'Frecuencia']
            hdr_cells = table.rows[0].cells
            for i, name in enumerate(headers):
                hdr_cells[i].text = name
                run = hdr_cells[i].paragraphs[0].runs[0]
                run.bold = True
                hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

            # 2. Leer cada fila de la interfaz
            # Aquí 'lista_de_filas_widgets' es la lista donde guardas cada fila nueva
            for fila in lista_de_filas_widgets:
                # Extraemos los valores de los inputs de esa fila
                # (Asegúrate de que 'fila' sea un objeto con acceso a los .value)
                recurso = fila['recurso'].value.strip()

                if recurso: # Si el recurso no está vacío, agregamos la fila al Word
                    row = table.add_row().cells
                    row[0].text = recurso
                    row[1].text = fila['organismo'].value
                    row[2].text = fila['puntos'].value
                    row[3].text = fila['parametros'].value

                    frec = fila['frecuencia'].value
                    row[4].text = frec if frec != 'Seleccione...' else ""

                    # Formato de letra tamaño 10
                    for cell in row:
                        for para in cell.paragraphs:
                            for run in para.runs:
                                run.font.size = Pt(10)
            break
    if not encontrado:
        # No print if not found, to avoid clutter in output for non-existent markers
        pass

# Lista donde guardaremos las filas para leerlas después
muestreo_filas_datos = [] # List to hold the *data* widgets for each row

# New list to hold the *display HBox widgets* for each row
muestreo_rows_display_widgets = []

def crear_fila_muestreo_manual_and_add():
    rec = widgets.Text(placeholder='Recurso...')
    org = widgets.Text(placeholder='Organismo...')
    pun = widgets.Text(placeholder='Puntos...')
    par = widgets.Text(placeholder='Parámetro...')
    fre = widgets.Dropdown(options=['Seleccione...', 'Mensual', 'Trimestral', 'Semestral', 'Anual', 'N/A'], value='Seleccione...')
    # The following widgets (cum and fec) are not used in generar_tabla_desde_interfaz_dinamica, but kept for consistency if needed later.
    cum = widgets.Dropdown(options=['SI', 'NO', 'N/A'], value='SI')
    fec = widgets.Text(placeholder='dd/mm/aaaa')

    # Store the actual input widgets for data retrieval
    muestreo_filas_datos.append({
        'recurso': rec, 'organismo': org, 'puntos': pun,
        'parametros': par, 'frecuencia': fre, 'cumplimiento': cum, 'fecha_analisis': fec
    })

    # Return the HBox for display
    return widgets.HBox([rec, org, pun, par, fre]) # Only display the relevant fields for the table

# VBox that will contain the dynamically added rows
muestreo_rows_container = widgets.VBox([], layout=widgets.Layout(border='1px solid lightgray', padding='5px'))

def on_add_muestreo_row(b=None):
    new_row_widget = crear_fila_muestreo_manual_and_add()
    muestreo_rows_display_widgets.append(new_row_widget)
    muestreo_rows_container.children = tuple(muestreo_rows_display_widgets)

# Add one initial row
on_add_muestreo_row()

# Button to add more rows
btn_add_muestreo_row = widgets.Button(description="Añadir Fila de Muestreo", button_style='success', icon='plus')
btn_add_muestreo_row.on_click(on_add_muestreo_row)

# Construimos la caja visual (VBox) para la tabla de muestreo
muestreo_table_content = widgets.VBox([
    widgets.HTML("<h3> Plan de Monitoreo (Calidad Ambiental)</h3>"),
    widgets.HTML("<b>Recurso | Organismo | Puntos | Parámetros | Frecuencia</b>"), # Header for visual guide
    muestreo_rows_container,
    btn_add_muestreo_row
], layout=widgets.Layout(padding='10px')).add_class('custom-font-widget')



# 3. LEGA y Muestreo
lega_muestreo_box = widgets.VBox([
    widgets.HTML("<h3>📜 LEGA</h3>"),
    widgets.VBox(list(lega_widgets.values())),
    muestreo_table_content # This is where the sampling table is added
], layout=widgets.Layout(padding='10px')).add_class('custom-font-widget')

# NEW: ADA Tab Content
ada_tab_content = widgets.VBox([
    widgets.HTML("<h3> ADA (Autoridad del Agua)</h3>"),
    widgets.VBox([
        otros_widgets['ADA_STATUS'], otros_widgets['ADA_FECHA'], otros_widgets['ADA_EXP'],
        otros_widgets['CHI_HID'], otros_widgets['CHI_EXP'], otros_widgets['CHI_VUE'],
        otros_widgets['ESTADO_PERMISO_HIDRAULICA'], # New field
        otros_widgets['ESTADO_PERMISO_EXPLOTACION'], # New field
        otros_widgets['ESTADO_PERMISO_VUELCO'], # New field
        otros_widgets['RED_MONITOREOS'] # New field
    ])
], layout=widgets.Layout(padding='10px')).add_class('custom-font-widget')

# 4. Otros Permisos (ACUMAR, SE, RENPRE, Seguro)
hidricos_otros_box = widgets.VBox([
    widgets.HTML("<h3> ACUMAR (Autoridad de Cuenca Matanza Riachuelo)</h3>"),
    widgets.VBox([
        otros_widgets['ACUMAR_STATUS'],
        otros_widgets['ACUMAR_EXP'],
        otros_widgets['ACUMAR_OBS']
    ]),
    widgets.HTML("<h3> Secretaría de Energía</h3>"),
    widgets.VBox([
        otros_widgets['SE_STATUS'], # General SE Status
        otros_widgets['SE_EXP'], # General SE Expediente
        # Section for Inscripcion 1102/04
        otros_widgets['INSCRIPCION_1102'],
        otros_widgets['NUMERO_SE'],
        # Section for Auditoria 404/94
        otros_widgets['AUDITORIA_404'],
        otros_widgets['CANTIDAD_DE_TANQUES'],
        otros_widgets['VENCIMIENTO_AUDITORIA404'],
        # Section for Inscripcion 277/25
        otros_widgets['INSCRIPCION_277'],
        otros_widgets['OBSERVACIONES_277'], # The new widget
        otros_widgets['SE_OBS'] # General SE Observations
    ]),
    widgets.HTML("<h3> RENPRE y Seguro Ambiental</h3>"),
    widgets.VBox([
        otros_widgets['RENPRE_STATUS'], otros_widgets['RENPRE_NUM'],
        widgets.HTML("<hr>"),
        otros_widgets['SEGURO_STATUS'], otros_widgets['POLIZA_NUM'], otros_widgets['POLIZA_VTO']
    ])
], layout=widgets.Layout(padding='10px')).add_class('custom-font-widget')


# 5. Operación y Muestreo (Residuos, ASP)
operacion_residuos_asp_box = widgets.VBox([
    widgets.HTML("<h3> 1. Residuos Especiales (Inscripción)</h3>"),
    widgets.VBox([
        rree_widgets['STATUS'] # General RREE status
    ]),
    widgets.HTML("<h3> 2. Certificado de Habilitación Especial (CHE)</h3>"),
    widgets.VBox([
        rree_widgets['CHE_STATUS'], # New CHE status
        rree_widgets['ANIO_CHE'],
        rree_widgets['OBS_CHE']
    ]),
    widgets.HTML("<h3>♻️ 3. Gestión Operativa de Residuos</h3>"),
    widgets.VBox([
        rree_widgets['GESTION_RES'],
        rree_widgets['TIPO_RES'],
        rree_widgets['OBS_EXTRA']
    ]),
    widgets.HTML("<h3>🌡️ ASP (Aparatos a Presión)</h3>"),
    widgets.VBox(list(asp_widgets.values()))
], layout=widgets.Layout(padding='10px')).add_class('custom-font-widget')

# 6. Hallazgos
hallazgos_tab_content = widgets.VBox([
    widgets.HTML("<h3>🔍 Hallazgos de Campo</h3>"),
    hallazgos_container,
    btn_add_hallazgo
], layout=widgets.Layout(padding='10px')).add_class('custom-font-widget')


# --- Creación de Pestañas (widgets.Tab) ---
tab_widget = widgets.Tab()
tab_widget.children = [
    general_hab_box,
    permisos_opds_box,
    lega_muestreo_box,
    ada_tab_content, # NEW ADA TAB
    hidricos_otros_box,
    operacion_residuos_asp_box,
    hallazgos_tab_content
]

tab_widget.set_title(0, '1. Info General')
tab_widget.set_title(1, '2. CNCA, CAAP, CAAF')
tab_widget.set_title(2, '3. LEGA y Monitoreos Ambientales')
tab_widget.set_title(3, '4. Autoridad del Agua') # NEW TITLE
tab_widget.set_title(4, '5. Otros Organismos') # Shifted index and updated title
tab_widget.set_title(5, '6. Residuos y ASP') # Shifted index
tab_widget.set_title(6, '7. Hallazgos y Recomendaciones') # Shifted index

# --- Controles de Archivo y Generación ---
btn_upload = widgets.Button(description="Cargar Plantilla DOCX", button_style='info', icon='upload')
btn_upload.on_click(upload_template)

btn_gen = widgets.Button(description="GENERAR INFORME", button_style='primary', layout=widgets.Layout(width='100%', height='50px'), icon='file-text')
btn_gen.on_click(on_generar_click)

output_text = widgets.Output()

# --- Definición de todos los mapas de widgets para la lógica condicional ---
all_widget_maps = {
    'main': main_widgets, 'hab': hab_widgets, 'cnca': cnca_widgets,
    'caap_caaf': caap_caaf_widgets, 'lega': lega_widgets, 'rree': rree_widgets,
    'asp': asp_widgets, 'otros': otros_widgets
}

# --- Llamada a la función de lógica condicional (Importante: debe ir antes de display)---
setup_conditional_fields(all_widget_maps)
# --- Mostrar la Interfaz ---
display(HTML(custom_css))
display(widgets.HTML("<h2>📝 Generador de Informes Ambientales</h2>"))
display(btn_upload)
display(tab_widget)
display(widgets.HTML("<br>"))
display(btn_gen)
display(output_text)
