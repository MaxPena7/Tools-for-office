import pandas as pd
import numpy as np

#Variable que tiene los motivos de altas; Nota: el "9" corresponde a "09"
altas = ["09", "10", "20", "24", "25", "84", "95", "96", "99", "61", "62", "63"]
#Variable que tiene los motivos de bajas: 
bajas = ["31", "32", "33", "34", "35", "36", "37", "38", "41", "42", "43", "48", "51", "52", "53"]
#variable para indicar la quincena proceso
qnaproc = 202513



#Declaramos la variable para identificar los conceptos que graban para ISR
cptosisr = ["14", "15", "07", "7A", "7B", "BC", "7C", "7D", "7E", "UA", "UB", "UF", "UC", "UD", "UE", "VA", "VB", "VF", "VC", "VD", "VE", "WA", "WB", "WF", "WC", "WD", "WE", "XA", "XB", "XF", "XC", "XD", "XE", "YA", "YB", "YF", "YC", "YD", "YE", "ZA", "ZB", "ZF", "ZC", "ZD", "ZE", "K1", "K2", "K3", "K4", "K5", "K1A", "K1B", "K1F", "K1C", "K1D", "K1E", "K2A", "K2B", "K2F", "K2C", "K2D", "K2E", "K3A", "K3B", "K3F", "K3C", "K3D", "K3E", "K4A", "K4B", "K4F", "K4C", "K4D", "K4E", "K5A", "K5B", "K5F", "K5C", "K5D", "K5E", "K6A", "K6B", "K6F", "K6C", "K6D", "K6E", "O1", "O2", "O3", "O4", "O5", "Q1", "Q2", "Q3", "Q4", "Q5", "A1", "A2", "A3", "A4", "A5", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "PA", "PB", "PC", "PD", "PE", "PF", "PG", "PH", "PI", "PJ", "PK", "PL", "PM", "PN", "PO", "PP", "PQ", "PR", "PS", "PT", "PU", "PV", "PW", "PX", "PY", "PZ", "QA", "QB", "QC", "QD", "QE", "QF", "QG", "QH", "QI", "QJ", "QK", "QL", "QM", "QN", "QO", "QP", "QQ", "QR", "QS", "QT", "QU", "QV", "QW", "QX", "QY", "QZ", "E2", "E3", "T1", "T2", "T3", "L1", "L2", "L3", "LT", "I2", "CP", "MA", "DO", "SC", "MB", "5A", "5D", "OE", "FC"]
#Declaramos la variable que indica la base gravable para ISR por rfc nómina ordinaria ("ior")
ior = None
#Declaramos la variable que indica la base gravable para ISR por rfc nómina extrordinaria ("ier")
ier = None
#Declaramos la variable que indica la base gravable para ISR por plaza nómina extraordinaria ("iep")
iep = None

#Declaramos la variable para identificar los conceptos que graban para ISSSTE
cptosisste = ["07", "7A", "7B", "BC", "7C", "7D", "7E", "UA", "UB", "UF", "UC", "UD", "UE", "VA", "VB", "VF", "VC", "VD", "VE", "WA", "WB", "WF", "WC", "WD", "WE", "XA", "XB", "XF", "XC", "XD", "XE", "YA", "YB", "YF", "YC", "YD", "YE", "ZA", "ZB", "ZF", "ZC", "ZD", "ZE", "K1", "K2", "K3", "K4", "K5", "O1", "O2", "O3", "O4", "O5", "Q1", "Q2", "Q3", "Q4", "Q5", "A1", "A2", "A3", "A4", "A5", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "PA", "PB", "PC", "PD", "PE", "PF", "PG", "PH", "PI", "PJ", "PK", "PL", "PM", "PN", "PO", "PP", "PQ", "PR", "PS", "PT", "PU", "PV", "PW", "PX", "PY", "PZ", "QA", "QB", "QC", "QD", "QE", "QF", "QG", "QH", "QI", "QJ", "QK", "QL", "QM", "QN", "QO", "QP", "QQ", "QR", "QS", "QT", "QU", "QV", "QW", "QX", "QY", "QZ", "E2", "T1", "T2", "T3", "L1", "L2", "L3", "LT", "MA", "DO"]
#Declaramos la vcvariable que indica la base gravable para ISSSTE por rfc nómina ordinaria ("isor")
isor = None
#Declaramos la vcvariable que indica la base gravable para ISSSTE por rfc nómina extraordinaria ("iser")
iser = None
#Declaramos la vcvariable que indica la base gravable para ISSSTE por plaza nómina extraordinaria ("isep")
isep = None

desctosterceros = ["03", "77", "65", "FJ", "IS", "51", "AD", "FS", "GF", "SF", "SM", "21", "VT", "GN", "FE", "57", "17", "54", "ET", "TM", "TD", "FT", "JM", "SG", "CF", "CN", "56", "RF", "MN", "18", "53", "PB", "SZ"]

# region Carga de archivos
#CARGA DE ARCHIVOS
# Abrimos el reporte de FUPS de la qna
df = pd.read_excel(r'C:\Users\Maxruso7\Desktop\REPORTE DE FUPS ACUMULADO\RptFup_202513_20250703.xlsx', index_col=None)
# Elimina columnas como 'Unnamed: 0' o numéricas mal colocadas
df = df.loc[:, ~df.columns.str.contains('^Unnamed|^0$', na=False)]

# Cargamos el archivo con todos los movimientos para buscar continuidad
movimientos_df = pd.read_excel(r'C:\Users\Maxruso7\Desktop\REPORTE DE FUPS ACUMULADO\combinado.xlsx', index_col=None)
# Elimina columnas como 'Unnamed: 0' o numéricas mal colocadas
movimientos_df = movimientos_df.loc[:, ~movimientos_df.columns.str.contains('^Unnamed|^0$', na=False)]

#Abrimos el archivo de nómina de movimientos "perded"
df_premov = pd.read_excel(r'C:\Users\Maxruso7\Desktop\SEP_HERRAMIENTAS\EXCEL\Q13\PERDED_QNA13_PRENOMINA MOVIMIENTOS.xlsx', index_col=None)
# Elimina columnas como 'Unnamed: 0' o numéricas mal colocadas
df_premov = df_premov.loc[:, ~df_premov.columns.str.contains('^Unnamed|^0$', na=False)]

#Abrimos una nómina anterior
nomina_anterior = pd.read_excel(r'C:\Users\Maxruso7\Desktop\SEP_HERRAMIENTAS\EXCEL\Q13\PERDED_QNA_12_PRENOMINA GENERAL_3RA_REVISION.xlsx', index_col=None)
# Elimina columnas como 'Unnamed: 0' o numéricas mal colocadas
nomina_anterior = nomina_anterior.loc[:, ~nomina_anterior.columns.str.contains('^Unnamed|^0$', na=False)]

# endregion



# region Análisi de liquidaciones
#---------------------------------------------------------------------------------------
#Bloque 1: Proceso para obtener Rfc, Plazas y quincenas a liquidar
#---------------------------------------------------------------------------------------

# Parte 1: Filtramos movimientos de alta
# Aseguramos que la columna de tipo de movimiento tenga formato string con ceros a la izquierda
df.iloc[:, 4] = df.iloc[:, 4].astype(str).str.zfill(2)

# Filtramos las filas donde haya altas
filtradomovaltas = df[df.iloc[:, 4].isin(altas)]

# Solo extraemos columnas 4, 7, 13 y 19 (TIPOMOV, PLAZA, RFC, QINICIO)
resultado = filtradomovaltas.iloc[:, [4, 7, 13, 19]]

# Agregamos la columna de HORAS (índice 11) al final
resultado['HORAS'] = filtradomovaltas.iloc[:, 11].values

# Agregamos la columna de HORAS (índice 11) al final
resultado['CATPUESTO'] = filtradomovaltas.iloc[:, 10].values


# Convertimos a número la columna QINICIO
resultado.iloc[:, 3] = pd.to_numeric(resultado.iloc[:, 3], errors='coerce')

# Filtramos registros cuya QINICIO sea menor a la quincena de proceso
filtradoliquidaciones = resultado[resultado.iloc[:, 3] < qnaproc].copy()

# Calculamos la diferencia de quincenas a liquidar
columna_3 = filtradoliquidaciones.iloc[:, 3].fillna(0)
filtradoliquidaciones['diferencia'] = qnaproc - columna_3

# Parte 2: Eliminamos registros con motivos de baja
# Creamos llave 'rfcplaza' en df original (RFC + PLAZA sin espacios)
df['rfcplaza'] = (df.iloc[:, 13].astype(str).str.replace(" ", "") +df.iloc[:, 7].astype(str).str.replace(" ", ""))

# Creamos misma llave en filtradoliquidaciones
filtradoliquidaciones['rfcplaza'] = (filtradoliquidaciones.iloc[:, 2].astype(str).str.replace(" ", "") +filtradoliquidaciones.iloc[:, 1].astype(str).str.replace(" ", ""))

# Filtramos las filas con motivos de baja
bajas_df = df[df.iloc[:, 4].isin(bajas)]

# Extraemos llaves de rfcplaza que tienen motivo de baja
llaves_bajas = bajas_df['rfcplaza'].unique()

# Excluimos del DataFrame de liquidaciones los registros con esas llaves
filtradoliquidaciones = filtradoliquidaciones[~filtradoliquidaciones['rfcplaza'].isin(llaves_bajas)]


# Análisis de antigüedad para altas con continuidad sin interrupciones
# Nos aseguramos de que las columnas EFEINI, EFEFIN, TIPOMOV y HORAS sean numéricas
movimientos_df.iloc[:, 19] = pd.to_numeric(movimientos_df.iloc[:, 19], errors='coerce')  # EFEINI
movimientos_df.iloc[:, 20] = pd.to_numeric(movimientos_df.iloc[:, 20], errors='coerce')  # EFEFIN
movimientos_df.iloc[:, 4] = movimientos_df.iloc[:, 4].astype(str).str.zfill(2)  # TIPOMOV
movimientos_df.iloc[:, 11] = pd.to_numeric(movimientos_df.iloc[:, 11], errors='coerce')  # HORAS

# Creamos columna jornada equivalente a 20 horas si HORAS == 0
movimientos_df['horas_equivalentes'] = movimientos_df.iloc[:, 11].apply(lambda x: 20 if x == 0 else x)

# Creamos columna para llave de comparación
movimientos_df['rfcplaza'] = (movimientos_df.iloc[:, 13].astype(str).str.replace(" ", "") + movimientos_df.iloc[:, 7].astype(str).str.replace(" ", ""))

# Generamos columna final de antigüedad respetada
antiguedades = []

for _, fila in filtradoliquidaciones.iterrows():

    # Si el movimiento actual es 24 o 25, se respeta su propia fecha sin análisis
    if fila.iloc[0] in ['24', '25']:
        antiguedades.append(fila.iloc[3])  # EFEINI actual
        continue


    rfc = fila.iloc[2]
    plaza = fila.iloc[1]
    efeini_actual = fila.iloc[3]
    horas_actual = 20 if fila['HORAS'] == 0 else fila['HORAS']
    clave_rfcplaza = fila['rfcplaza']
    
    # Buscamos movimientos anteriores del mismo RFC (no importa plaza) con mismo tipo de alta (excepto 24 y 25)
    prev_movs = movimientos_df[
        (movimientos_df.iloc[:, 13] == rfc) &
        (movimientos_df.iloc[:, 4].isin(altas)) &
        (~movimientos_df.iloc[:, 4].isin(['24', '25'])) &
        (movimientos_df.iloc[:, 20] == efeini_actual - 1)
    ]
    
    # Filtrar por horas equivalentes (igual o menor)
    prev_movs = prev_movs[
        prev_movs['horas_equivalentes'] <= horas_actual
    ]

    # Si hay al menos un movimiento válido anterior, respetamos la antigüedad
    if not prev_movs.empty:
        efeini_respetado = prev_movs.iloc[0, 19]
    else:
        efeini_respetado = efeini_actual

    antiguedades.append(efeini_respetado)

# Agregamos la nueva columna al DataFrame final
filtradoliquidaciones['EFEINI_ANTIGUEDAD'] = antiguedades

#Reordenamos el dataframe para no afectar los calculos siguientes (por el orden de las columnas)
filtradoliquidaciones = filtradoliquidaciones[
    ['TIPOMOV', 'CVE_PPTAL', 'RFC', 'EFEINI', 'diferencia', 'rfcplaza', 'HORAS', 'EFEINI_ANTIGUEDAD', 'CATPUESTO']
]

# Función para calcular quincenas hacia atrás
def calcular_quincenas_atras(quincena_actual, num_quincenas):
    """
    Calcula una quincena X número de quincenas hacia atrás
    """
    año = quincena_actual // 100
    quincena = quincena_actual % 100
    
    quincenas_restantes = num_quincenas
    
    while quincenas_restantes > 0:
        quincena -= 1
        if quincena == 0:
            quincena = 24
            año -= 1
        quincenas_restantes -= 1
    
    return año * 100 + quincena

# Función para verificar si le corresponden las quincenas de vacaciones
def verificar_quincenas_vacaciones(efeini_antiguedad, quincena_proceso):
    """
    Verifica qué quincenas de vacaciones le corresponden según antigüedad
    """
    quincenas_vacaciones = {8, 14, 15, 24}
    quincenas_permitidas = set()
    
    # Para quincena 8: debe estar activo desde 9 quincenas atrás
    limite_q8 = calcular_quincenas_atras(quincena_proceso, 9)
    if efeini_antiguedad <= limite_q8:
        quincenas_permitidas.add(8)
    
    # Para quincenas 14 y 15: debe estar activo desde cuando se paga la 14
    # Calculamos cuándo se paga la 14 (normalmente en julio, quincena 14)
    año_actual = quincena_proceso // 100
    quincena_pago_14 = año_actual * 100 + 14
    limite_q14 = calcular_quincenas_atras(quincena_pago_14, 9)
    if efeini_antiguedad <= limite_q14:
        quincenas_permitidas.update({14, 15})
    
    # Para quincena 24: debe estar activo desde 9 quincenas atrás del pago
    quincena_pago_24 = año_actual * 100 + 24
    limite_q24 = calcular_quincenas_atras(quincena_pago_24, 9)
    if efeini_antiguedad <= limite_q24:
        quincenas_permitidas.add(24)
    
    return quincenas_permitidas

# Función principal para calcular quincenas a liquidar
def calcular_quincenas_liquidar(row, qnaproc):
    """
    Calcula el número de quincenas a liquidar considerando restricciones de vacaciones
    """
    efeini = row['EFEINI']
    efeini_antiguedad = row['EFEINI_ANTIGUEDAD']
    catpuesto = str(row['CATPUESTO']).upper()
    
    # Verificar si es docente (comienza con E, excepto ED02810)
    es_docente = catpuesto.startswith('E') and catpuesto != 'ED02810'
    
    # Calcular todas las quincenas entre EFEINI y qnaproc
    año_inicio = efeini // 100
    quincena_inicio = efeini % 100
    año_proceso = qnaproc // 100
    quincena_proceso = qnaproc % 100
    
    quincenas_totales = []
    año_actual = año_inicio
    quincena_actual = quincena_inicio
    
    # Generar lista de todas las quincenas a liquidar
    while año_actual * 100 + quincena_actual < qnaproc:
        quincenas_totales.append(quincena_actual)
        quincena_actual += 1
        if quincena_actual > 24:
            quincena_actual = 1
            año_actual += 1
    
    # Si no es docente, se pagan todas las quincenas
    if not es_docente:
        return len(quincenas_totales)
    
    # Si es docente, aplicar restricciones de vacaciones
    quincenas_vacaciones_permitidas = verificar_quincenas_vacaciones(efeini_antiguedad, qnaproc)
    quincenas_a_pagar = []
    
    for q in quincenas_totales:
        if q in {8, 14, 15, 24}:  # Es quincena de vacaciones
            if q in quincenas_vacaciones_permitidas:
                quincenas_a_pagar.append(q)
            # Si es quincena 15 y no tiene derecho a 14, tampoco a 15
            elif q == 15 and 14 not in quincenas_vacaciones_permitidas:
                continue
        else:  # No es quincena de vacaciones, se paga normalmente
            quincenas_a_pagar.append(q)
    
    return len(quincenas_a_pagar)

# Aplicar la función al DataFrame, sustituyendo los valores de la columna 'diferencia'
filtradoliquidaciones['diferencia'] = filtradoliquidaciones.apply(
    lambda row: calcular_quincenas_liquidar(row, qnaproc), axis=1
)

# endregion


# region Cálculo de base gravable para ISR
#--------------------------------------------------------------------------------------
#Bloque 2: Cálculo de ISR, obtención de la base gravable para ISR, de la prenómina extraodinaria y ordinaria
#--------------------------------------------------------------------------------------

#SACAMOS LA BASE GRAVABLE POR RFC DE LA PRENÓMINA ORDINARIA
#Filtramos las filas en donde sean puras percepcciones y ademas que graben para ISR
isr_ord = df_premov[(df_premov.iloc[:, 5] == "P") & (df_premov.iloc[:, 6].isin(cptosisr))].iloc[:, [0, 2, 5, 6, 7]]
#Hacemos un subtotal (sumar.si) para sacar la base gravable por RFC
ior = isr_ord.groupby(isr_ord.columns[0])[isr_ord.columns[4]].sum().reset_index()


#SACAMOS LA BASE GRAVABLE POR RFC DE LA PRENÓNIMA EXTRAORDINARIA
#Hacemos una llave RFC PLAZA en df_permov y quitamos espacios, ademas al principio de df creamos una nueva columna con la llave
rfcplaza1 = df_premov.iloc[:, 0].astype(str).str.replace(" ","") + df_premov.iloc[:, 2].astype(str).str.replace(" ", "")
df_premov_llave = df_premov.copy()
df_premov_llave.insert(0, 'rfcplaza1', rfcplaza1)

#Hacemos una llave RFC PLAZA en filtradoliquidaciones y quitamos espacios, ademas al principio de df creamos una nueva columna con la llave
rfcplaza2 = filtradoliquidaciones.iloc[:, 2].astype(str).str.replace(" ","") + filtradoliquidaciones.iloc[:, 1].astype(str).str.replace(" ","")
filtradoliquidaciones_llave = filtradoliquidaciones.copy()
filtradoliquidaciones_llave.insert(0, 'rfcplaza2', rfcplaza2)

#Hacemos el cruce con ambos df para multiplicar la diferencia por el importe de los conceptos de pago de cada plaza a liquidar
# Renombrar la llave en filtradoliquidaciones_llave para que coincida con df_premov_llave
filtradoliquidaciones_llave = filtradoliquidaciones_llave.rename(columns={'rfcplaza2': 'rfcplaza1'})

#Cruzar los DataFrames por la llave 'rfcplaza1' usando left join para mantener todas las filas de df_premov_llave
df_cruce = pd.merge(df_premov_llave, filtradoliquidaciones_llave[['rfcplaza1', filtradoliquidaciones_llave.columns[5]]], on='rfcplaza1', how='left')

#Renombrar la columna 'diferencia'
df_cruce.rename(columns={filtradoliquidaciones_llave.columns[5]: 'diferencia'}, inplace=True)

#Multiplicamos la columna de diferencia por el importe
#Nota: primero aseguramos de que las columnas coincidan en el mismo formato que es numérico
df_cruce.iloc[:, [8, 11]] = df_cruce.iloc[:, [8, 11]].apply(pd.to_numeric, errors='coerce')

#Ahora si multiplicamos
df_cruce['nuevo importe'] = df_cruce.iloc[:, 8] * df_cruce.iloc[:, 11]

#Procedemos a filtrar solo las filas que sean percepciones y que además graven para ISR y que además en la columna diferencia tenga algun valor, o sea, estamos aplicando 3 filtros
isr_ext = df_cruce[(df_cruce.iloc[:, 6] == "P") & (df_cruce.iloc[:, 7].isin(cptosisr)) & pd.to_numeric(df_cruce.iloc[:, 11], errors='coerce').notna()]

#Hacemos un subtotal (sumar.si) para sacar la base gravable por RFC de la extraordinaria
ier = isr_ext.groupby(isr_ext.columns[1])[isr_ext.columns[12]].sum().reset_index()


#SACAMOS LA BASE GRAVABLE POR PLAZA DE LA PRENÓMINA EXTRAORDINARIA 
#Nota: Con el DF isr_ext sacamos el subtotal para sacar la base gravable por plaza de la prenómina extraordinaria
iep = isr_ext.groupby(isr_ext.columns[0]).agg({isr_ext.columns[1]: "first", isr_ext.columns[3]: "first", isr_ext.columns[12]: "sum"}).reset_index()

# endregion


# region Cálculo de ISR

#PROCESO PARA EL CÁLCULO DEL ISR
#Primero: Creamos un nuevo DataFrame en donde combinaremos todas las bases gravables para concentrar toda la información
#Hacemos un nuevo dataframe con la misma estructrua de iep
todogravableisr = iep.copy()
#Cruzamos con ier para traernos la base base gravable de la nómina extraordinaria por RFC
todogravableisr = todogravableisr.merge(ier[[ier.columns[0], ier.columns[1]]], left_on=todogravableisr.columns[1], right_on=ier.columns[0], how="left", suffixes=('', '_ier'))
#Cruzamos con ior para traernos la base base gravable de la nómina ordinaria por RFC
todogravableisr = todogravableisr.merge(ior[[ior.columns[0], ior.columns[1]]], left_on=todogravableisr.columns[1], right_on=ior.columns[0], how="left", suffixes=('', '_ier'))

#Segundo: Definimos las tablas y procesos para el cálculo, como una tipo Macro o función que nos regresa los valores calculados que despues nos mostrará cuando las llamemos
#Definimos la tabla de ISR 
tabla_isr = pd.DataFrame({
    "limite_inferior": [0.01, 746.05, 6332.06, 11128.02, 12935.83, 15487.72, 31236.50, 49233.01, 93993.91, 125325.21, 375975.62],
    "limite_superior": [746.04, 6332.05, 11128.01, 12935.82, 15487.71, 31236.49, 49233.00, 93993.90, 125325.20, 375975.61, np.inf],
    "cuota_fija": [0.00, 14.32, 371.83, 893.63, 1182.88, 1640.18, 5004.12, 9236.89, 22665.17, 32691.18, 117912.32],
    "porcentaje": [0.0192, 0.0640, 0.1088, 0.1600, 0.1792, 0.2136, 0.2352, 0.3000, 0.3200, 0.3400, 0.3500]
})

#Creamos la función de cálculo de ISR
def calcular_isr_con_subsidio(importe_qnal):
    # Convertir a mensual (x2)
    importe_mensual = importe_qnal * 2
    
    # Calcular ISR según la tabla
    fila_isr = tabla_isr[(importe_mensual >= tabla_isr["limite_inferior"]) & (importe_mensual <= tabla_isr["limite_superior"])].iloc[0]
    
    impuesto = fila_isr["cuota_fija"] + ((importe_mensual - fila_isr["limite_inferior"]) * fila_isr["porcentaje"])
    
    # Aplicar subsidio con la lógica exacta
    if importe_mensual <= 10171:  # Nota: Usamos <= 10171 (sin .01)
        impuesto_subsidio = impuesto - 475
        impuesto_final = max(impuesto_subsidio, 0)  # Si es negativo, queda 0
    else:  # importe_mensual > 10171 (no aplica subsidio)
        impuesto_final = impuesto
    
    return impuesto_final

#CALCULAMOS EL ISR DE LA PRENÓMINA ORDINARIA
#Primero preparamos el DataFrame para aplicar la función de cálculo de ISR
isr_rfc_ord = todogravableisr.drop_duplicates(subset=["RFC"], keep="first").copy()

#Calculamos el ISR para cada RFC único y agregamos una columna donde pongamos el resultado del cálculo
isr_rfc_ord["ISR_ORD"] = isr_rfc_ord["IMPORTE"].apply(calcular_isr_con_subsidio)


#CALCULAMOS EL ISR DE LA PRENÓMINA EXTRAORDINARIA
#Aprovechamos el DF de irs_rfc_ord para copiar su estructura en el nuevo DF ya con los dos impuestos calculados
isr_rfc_ord_ext = isr_rfc_ord.drop_duplicates(subset=["RFC"], keep="first").copy()
isr_rfc_ord_ext["ISR_EXT"] = (isr_rfc_ord_ext["IMPORTE"] + isr_rfc_ord_ext["nuevo importe_ier"]).apply(calcular_isr_con_subsidio)


#Procedemos a sacar la diferencia de impuesto
isr_rfc_plaza = isr_rfc_ord_ext.drop_duplicates(subset=["RFC"], keep="first").copy()
isr_rfc_plaza["DIFERENCIA"] = (isr_rfc_ord_ext["ISR_EXT"] - isr_rfc_ord_ext["ISR_ORD"])/2

#Cruzamos los DF, para crear uno nuevo en donde venga toda la información para el cálculo final del ISR
#Identificar los nombres de las columnas por índice en isr_rfc_plaza
columnas_a_agregar = isr_rfc_plaza.columns[[6, 7, 8]]  # ISR_ORD, ISR_EXT, DIFERENCIA

#Hacer el merge por la columna en índice 1 (RFC) sin modificar todogravableisr
calculo_isr_pre_final = todogravableisr.merge(
    isr_rfc_plaza.iloc[:, [1, 6, 7, 8]],  # Seleccionar columnas por índice: RFC (1), ISR_ORD (6), etc.
    left_on=todogravableisr.columns[1],   # Columna en índice 1 del left (todogravableisr)
    right_on=isr_rfc_plaza.columns[1],    # Columna en índice 1 del right (isr_rfc_plaza)
    how='left'  # Conservar todos los datos de todogravableisr
)

#Calculamos el impuesto final y por plaza
#Creamos un nuevo DF, con la misma estructura del DF calculo_isr_pre_final
isr_final = calculo_isr_pre_final.copy()

#Finalmente calculamos el ISR por plaza truncado a 2 digítos
isr_final["ISR POR PLAZA"] = ((isr_final["nuevo importe"] / isr_final["nuevo importe_ier"]) * isr_final["DIFERENCIA"]).apply(lambda x: np.trunc(x * 100) / 100)

# endregion


# region Cálculo de base gravable para ISSSTE
#--------------------------------------------------------------------------------------
#Bloque 3: Cálculo de ISSSTE, obtención de la base gravable para ISSSTE, de la prenómina extraodinaria y ordinaria
#--------------------------------------------------------------------------------------

#SACAMOS LA BASE GRAVABLE POR RFC DE LA PRENÓMINA ORDINARIA
#Filtramos las filas en donde sean puras percepcciones y ademas que graben para ISSSTE
ist_ord = df_premov[(df_premov.iloc[:, 5] == "P") & (df_premov.iloc[:, 6].isin(cptosisste))].iloc[:, [0, 2, 5, 6, 7]]
#Hacemos un subtotal (sumar.si) para sacar la base gravable por RFC
isor = ist_ord.groupby(isr_ord.columns[0])[isr_ord.columns[4]].sum().reset_index()


#SACAMOS LA BASE GRAVABLE POR RFC DE LA PRENÓNIMA EXTRAORDINARIA
#Procedemos a filtrar solo las filas que sean percepciones y que además graven para ISSSTE y que además en la columna diferencia tenga algun valor, o sea, estamos aplicando 3 filtros
isst_ext = df_cruce[(df_cruce.iloc[:, 6] == "P") & (df_cruce.iloc[:, 7].isin(cptosisste)) & pd.to_numeric(df_cruce.iloc[:, 11], errors='coerce').notna()]
#Hacemos un subtotal (sumar.si) para sacar la base gravable por RFC de la extraordinaria
iser = isst_ext.groupby(isst_ext.columns[1])[isst_ext.columns[12]].sum().reset_index()


#SACAMOS LA BASE GRAVABLE POR PLAZA DE LA PRENÓMINA EXTRAORDINARIA 
#Nota: Con el DF isst_ext sacamos el subtotal para sacar la base gravable por plaza de la prenómina extraordinaria
isep = isst_ext.groupby(isst_ext.columns[0]).agg({
    isst_ext.columns[1]: "first",
    isst_ext.columns[3]: "first",
    isst_ext.columns[12]: "sum"
}).reset_index()

# endregion


# region Cálculo de ISSSTE
#PROCESO PARA EL CÁLCULO DE CUOTAS DEL ISSSTE
# Preparamos a iser e isor con nombres personalizados
iser_temp = iser.iloc[:, [0, 1]].copy()
iser_temp.columns = ['RFC', 'RFC_EXT']

isor_temp = isor.iloc[:, [0, 1]].copy()
isor_temp.columns = ['RFC', 'RFC_ORD']

# Copiamos isep y extraemos su RFC
isep_temp = isep.copy()
isep_temp['RFC'] = isep.iloc[:, 1]  # RFC está en columna índice 1

# Guardamos columnas originales ANTES de hacer el merge
columnas_originales = isep_temp.columns.tolist()

# Hacemos el merge con iser_temp e isor_temp
resultado = (
    isep_temp
    .merge(iser_temp, on='RFC', how='left')
    .merge(isor_temp, on='RFC', how='left')
)

#Reordenamos columnas al final
resultado = resultado[columnas_originales + ['RFC_EXT', 'RFC_ORD']]

#Renombramos el DataFrame de resultado para mejor entendimiento
todo_issste = resultado.copy()

#Calculamos el factor para el cálculo de las cuotas
todo_issste["FACTOR %"] = (todo_issste["nuevo importe"] / todo_issste["RFC_EXT"])

#Hacemos un filtro para quitar a los que no llevarán cuotas por estar topados
todo_issste = todo_issste[todo_issste.iloc[:, 5] <= 16971].copy()

#Calculamos la columna CV y la agregamos al final
todo_issste["CV"] = np.where(
    todo_issste.iloc[:, 4] + todo_issste.iloc[:, 5] > 16971,
    (1039.47 - (todo_issste.iloc[:, 5] * 0.06125)) * todo_issste.iloc[:, 6],
    todo_issste.iloc[:, 3] * 0.06125
).round(2)

#Calculamos la columna SI y la agregamos al final
todo_issste["SI"] = np.where(
    todo_issste.iloc[:, 4] + todo_issste.iloc[:, 5] > 16971,
    (106.06 - (todo_issste.iloc[:, 5] * 0.00625)) * todo_issste.iloc[:, 6],
    todo_issste.iloc[:, 3] * 0.00625
).round(2)

#Calculamos la columna SO y la agregamos al final
todo_issste["SO"] = np.where(
    todo_issste.iloc[:, 4] + todo_issste.iloc[:, 5] > 16971,
    (84.85 - (todo_issste.iloc[:, 5] * 0.005)) * todo_issste.iloc[:, 6],
    todo_issste.iloc[:, 3] * 0.005
).round(2)

#Calculamos la comuna SS y la agregamos al final
todo_issste["SS"] = np.where(
    todo_issste.iloc[:, 4] + todo_issste.iloc[:, 5] > 16971,
    (572.77 - (todo_issste.iloc[:, 5] * 0.03375)) * todo_issste.iloc[:, 6],
    todo_issste.iloc[:, 3] * 0.03375
).round(2)

# endregion


# region Reemplazo de ISR y cuotas ISSSTE y filas innecesarias

#Armamos el perded final, con el nuevo ISR y cuotas ISSSTE
#Abrimos el archivo de nómina de movimientos "perded" (ya estaba guardado en df_premov), y lo renombramos para evitar confusiones
prenominamovimientos = df_premov.copy()

#Reemplazamos los valores de ISR por los calculados
#Creamos la llave temporal desde prenominamovimientos
llave_prenomina = (
    prenominamovimientos.iloc[:, 0].astype(str).str.replace(" ", "", regex=False) +
    prenominamovimientos.iloc[:, 2].astype(str).str.replace(" ", "", regex=False)
)

#Creamos un diccionario con llave: ISR desde isr_final
isr_dict = dict(
    zip(
        isr_final.iloc[:, 0].astype(str).str.replace(" ", "", regex=False),
        isr_final.iloc[:, 9]  # Valor de ISR_POR_PLAZA
    )
)

#Condición para encontrar las filas que deben ser modificadas
condicion = (
    (prenominamovimientos.iloc[:, 5] == "D") &
    (prenominamovimientos.iloc[:, 6] == "01")
)

#Solo para esas filas, reemplazamos el valor en la columna 7
#Usando la llave para buscar en el diccionario
prenominamovimientos.loc[condicion, prenominamovimientos.columns[7]] = llave_prenomina[condicion].map(isr_dict)


#Reemplazamos los conceptos de cuotas ISSSTE
#Llave temporal desde prenominamovimientos
llave_prenomina = (
    prenominamovimientos.iloc[:, 0].astype(str).str.replace(" ", "", regex=False) +
    prenominamovimientos.iloc[:, 2].astype(str).str.replace(" ", "", regex=False)
)

#Crear diccionarios desde todo_issste
cv_dict = dict(zip(todo_issste.iloc[:, 0].astype(str).str.replace(" ", "", regex=False), todo_issste.iloc[:, 7]))
si_dict = dict(zip(todo_issste.iloc[:, 0].astype(str).str.replace(" ", "", regex=False), todo_issste.iloc[:, 8]))
so_dict = dict(zip(todo_issste.iloc[:, 0].astype(str).str.replace(" ", "", regex=False), todo_issste.iloc[:, 9]))
ss_dict = dict(zip(todo_issste.iloc[:, 0].astype(str).str.replace(" ", "", regex=False), todo_issste.iloc[:, 10]))

#Condición base: TIPO == 'D'
cond_base = prenominamovimientos.iloc[:, 5] == "D"

# Reemplazos por concepto
# CV
cond_cv = cond_base & (prenominamovimientos.iloc[:, 6] == "CV")
prenominamovimientos.loc[cond_cv, prenominamovimientos.columns[7]] = llave_prenomina[cond_cv].map(cv_dict)

# SI
cond_si = cond_base & (prenominamovimientos.iloc[:, 6] == "SI")
prenominamovimientos.loc[cond_si, prenominamovimientos.columns[7]] = llave_prenomina[cond_si].map(si_dict)

# SO
cond_so = cond_base & (prenominamovimientos.iloc[:, 6] == "SO")
prenominamovimientos.loc[cond_so, prenominamovimientos.columns[7]] = llave_prenomina[cond_so].map(so_dict)

# SS
cond_ss = cond_base & (prenominamovimientos.iloc[:, 6] == "SS")
prenominamovimientos.loc[cond_ss, prenominamovimientos.columns[7]] = llave_prenomina[cond_ss].map(ss_dict)


#Eliminamos las filas que no son necesarias, esto cruzando con el DF filtrado liquidaciones
#Crear llave temporal en prenominamovimientos (columna 0 + 2)
llave_prenomina = (
    prenominamovimientos.iloc[:, 0].astype(str).str.replace(" ", "", regex=False) +
    prenominamovimientos.iloc[:, 2].astype(str).str.replace(" ", "", regex=False)
)

#Obtener llaves válidas del DataFrame filtradoliquidaciones (columna índice 5)
llaves_validas = filtradoliquidaciones.iloc[:, 5].astype(str).str.replace(" ", "", regex=False)

#Filtrar prenominamovimientos conservando solo las llaves válidas
prenominamovimientos = prenominamovimientos[llave_prenomina.isin(llaves_validas)].reset_index(drop=True)

#Convertimos la columa de importe a formato de número
prenominamovimientos.iloc[:, 7] = pd.to_numeric(prenominamovimientos.iloc[:, 7], errors='coerce')

# endregion


#region Eliminación de descuentos de terceros que no corresponden
#ELIMINACIÓN DE DESCUENTOS DE TERCEROS QUE NO CORRESPONDEN
#Procedemos a hacer un análisis para eleminar descuentos de terceros que ya se descontaron en la Qna pasada

#Obtener los nombres de columnas por índice
rfc_col = nomina_anterior.columns[0]    # columna 0: RFC
tipo_col = nomina_anterior.columns[5]   # columna 5: "D" (tipo de movimiento)
concepto_col = nomina_anterior.columns[6]  # columna 6: concepto de descuento

descuentos_previos = nomina_anterior[
    (nomina_anterior[tipo_col] == "D") & 
    (nomina_anterior[concepto_col].isin(desctosterceros))
][[rfc_col, concepto_col]].drop_duplicates()

# Eliminar de prenominamovimientos solo las filas con esos RFC + concepto
# Obtener los nombres de columnas de prenominamovimientos también
rfc_col_pre = prenominamovimientos.columns[0]
concepto_col_pre = prenominamovimientos.columns[6]

filtro = prenominamovimientos.merge(
    descuentos_previos,
    how="left",
    on=[rfc_col_pre, concepto_col_pre],
    indicator=True
)

prenominamovimientos_filtrado = filtro[filtro["_merge"] == "left_only"].drop(columns=["_merge"])

# endregion


# region Modificación de periodos en el perded
#Modificación de periodos en el perded
#Crear llaves temporales sin modificar las columnas originales
filtradoliquidaciones['llave_rfc_plaza'] = (
    filtradoliquidaciones['RFC'].astype(str).str.replace(" ", "", regex=False) +
    filtradoliquidaciones['CVE_PPTAL'].astype(str).str.replace(" ", "", regex=False)
)

prenominamovimientos_filtrado['llave_rfc_plaza'] = (
    prenominamovimientos_filtrado['RFC'].astype(str).str.replace(" ", "", regex=False) +
    prenominamovimientos_filtrado['PLAZA'].astype(str).str.replace(" ", "", regex=False)
)

#Crear diccionario con EFEINI desde filtradoliquidaciones para hacer el mapeo
dict_efeini = dict(zip(filtradoliquidaciones['llave_rfc_plaza'], filtradoliquidaciones['EFEINI']))

#Asignar EFEINI a las columnas QI y QICON usando la llave temporal
prenominamovimientos_filtrado['QI'] = prenominamovimientos_filtrado['llave_rfc_plaza'].map(dict_efeini)
prenominamovimientos_filtrado['QICON'] = prenominamovimientos_filtrado['QI']

#Asignar QF y QFCON como qnaproc - 1
prenominamovimientos_filtrado['QF'] = qnaproc - 1
prenominamovimientos_filtrado['QFCON'] = qnaproc - 1

#Eliminar la columna de llave temporal
prenominamovimientos_filtrado.drop(columns=['llave_rfc_plaza'], inplace=True)

# Crear la llave en ambos DataFrames (sin modificar las columnas originales)
llave_prenomina = prenominamovimientos_filtrado['RFC'].astype(str).str.replace(" ", "") + prenominamovimientos_filtrado['PLAZA'].astype(str).str.replace(" ", "")
llave_filtrado = filtradoliquidaciones['RFC'].astype(str).str.replace(" ", "") + filtradoliquidaciones['CVE_PPTAL'].astype(str).str.replace(" ", "")

# Crear un diccionario con las diferencias por llave
diferencias_dict = dict(zip(llave_filtrado, filtradoliquidaciones['diferencia']))

# Crear columna auxiliar con la llave en el DataFrame filtrado
prenominamovimientos_filtrado['llave'] = llave_prenomina

# Definir conceptos excluidos
conceptos_excluidos = ['01', 'CV', 'SI', 'SO', 'SS']

# Aplicar la multiplicación solo si el concepto no está excluido y existe la llave
mascara = ~prenominamovimientos_filtrado['CONCEPTO'].isin(conceptos_excluidos) & prenominamovimientos_filtrado['llave'].isin(diferencias_dict.keys())

# Multiplicar el importe por la diferencia correspondiente
prenominamovimientos_filtrado.loc[mascara, 'IMPORTE'] = (
    prenominamovimientos_filtrado.loc[mascara, 'IMPORTE'] *
    prenominamovimientos_filtrado.loc[mascara, 'llave'].map(diferencias_dict)
)

# Eliminar la columna auxiliar
prenominamovimientos_filtrado.drop(columns=['llave'], inplace=True)


# endregion


#prenominamovimientos_filtrado.to_excel(r'C:\Users\Maxruso7\Desktop\SEP_HERRAMIENTAS\EXCEL\prenominamovimientos_filtrado5.xlsx', index=False)
#filtradoliquidaciones.to_excel(r'C:\Users\Maxruso7\Desktop\SEP_HERRAMIENTAS\EXCEL\filtradoliquidaciones.xlsx', index=False)
#todo_issste.to_excel(r'C:\Users\Maxruso7\Desktop\SEP_HERRAMIENTAS\EXCEL\todo_issste.xlsx')
#isr_ord.to_excel(r'C:\Users\Maxruso7\Desktop\SEP_HERRAMIENTAS\EXCEL\isr_ord.xlsx')
#df_cruce.to_excel(r'C:\Users\Maxruso7\Desktop\SEP_HERRAMIENTAS\EXCEL\df_cruce.xlsx')
#iep.to_excel(r'C:\Users\Maxruso7\Desktop\SEP_HERRAMIENTAS\EXCEL\iep.xlsx')
df_premov.to_excel(r'C:\Users\Maxruso7\Desktop\SEP_HERRAMIENTAS\EXCEL\df_premov.xlsx')
print("FINALIZADO")


