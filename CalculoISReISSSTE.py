import pandas as pd
import numpy as np

#Declaramos la variable para identificar los conceptos que graban para ISR
cptosisr = ["14", "15", "07", "7A", "7B", "BC", "7C", "7D", "7E", "UA", "UB", "UF", "UC", "UD", "UE", "VA", "VB", "VF", "VC", "VD", "VE", "WA", "WB", "WF", "WC", "WD", "WE", "XA", "XB", "XF", "XC", "XD", "XE", "YA", "YB", "YF", "YC", "YD", "YE", "ZA", "ZB", "ZF", "ZC", "ZD", "ZE", "K1", "K2", "K3", "K4", "K5", "K1A", "K1B", "K1F", "K1C", "K1D", "K1E", "K2A", "K2B", "K2F", "K2C", "K2D", "K2E", "K3A", "K3B", "K3F", "K3C", "K3D", "K3E", "K4A", "K4B", "K4F", "K4C", "K4D", "K4E", "K5A", "K5B", "K5F", "K5C", "K5D", "K5E", "K6A", "K6B", "K6F", "K6C", "K6D", "K6E", "O1", "O2", "O3", "O4", "O5", "Q1", "Q2", "Q3", "Q4", "Q5", "A1", "A2", "A3", "A4", "A5", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "PA", "PB", "PC", "PD", "PE", "PF", "PG", "PH", "PI", "PJ", "PK", "PL", "PM", "PN", "PO", "PP", "PQ", "PR", "PS", "PT", "PU", "PV", "PW", "PX", "PY", "PZ", "QA", "QB", "QC", "QD", "QE", "QF", "QG", "QH", "QI", "QJ", "QK", "QL", "QM", "QN", "QO", "QP", "QQ", "QR", "QS", "QT", "QU", "QV", "QW", "QX", "QY", "QZ", "E2", "E3", "T1", "T2", "T3", "L1", "L2", "L3", "LT", "I2", "CP", "MA", "DO", "SC", "MB", "5A", "5D", "OE", "FC"]

#Declaramos la variable para identificar los conceptos que graban para ISSSTE
cptosisste = ["07", "7A", "7B", "BC", "7C", "7D", "7E", "UA", "UB", "UF", "UC", "UD", "UE", "VA", "VB", "VF", "VC", "VD", "VE", "WA", "WB", "WF", "WC", "WD", "WE", "XA", "XB", "XF", "XC", "XD", "XE", "YA", "YB", "YF", "YC", "YD", "YE", "ZA", "ZB", "ZF", "ZC", "ZD", "ZE", "K1", "K2", "K3", "K4", "K5", "O1", "O2", "O3", "O4", "O5", "Q1", "Q2", "Q3", "Q4", "Q5", "A1", "A2", "A3", "A4", "A5", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "PA", "PB", "PC", "PD", "PE", "PF", "PG", "PH", "PI", "PJ", "PK", "PL", "PM", "PN", "PO", "PP", "PQ", "PR", "PS", "PT", "PU", "PV", "PW", "PX", "PY", "PZ", "QA", "QB", "QC", "QD", "QE", "QF", "QG", "QH", "QI", "QJ", "QK", "QL", "QM", "QN", "QO", "QP", "QQ", "QR", "QS", "QT", "QU", "QV", "QW", "QX", "QY", "QZ", "E2", "T1", "T2", "T3", "L1", "L2", "L3", "LT", "MA", "DO"]

#Variable para base gravable para iSR
baseISR = None

#Variable para el importe ISR
ISR = None

#Variable para base gravable para ISSSTE
baseISSSTE = None

#Variable para el importe del ISSSTE
ISSSTE = None

#Carga de archivos
#Abrimos el archivo de nómina de movimientos "perded"
df_perded = pd.read_excel(r'C:\Users\Maxruso7\Desktop\INCREMENTO 2025\EDGAR\LIQ_INC_202515_15QNAS_ISR Y CUOTAS ISSSTE_MAX.xlsx', index_col=None)


# region Base gravable para ISR
#Hacemos una llave RFCPLAZA y la insertamos al inicio
df_perded1 = df_perded.iloc[:, 0].astype(str).str.replace(" ","") + df_perded.iloc[:, 2].astype(str).str.replace(" ","")
df_perded_llave = df_perded.copy()
df_perded_llave.insert(0, 'rfcplaza',df_perded1)


#SACAMOS LA BASE GRAVABLE POR RFCPLAZA
#Filtramos las filas en donde sean puras percepcciones y ademas que graben para ISR
baseISR_temporal = df_perded_llave[(df_perded_llave.iloc[:, 6] == "P") & (df_perded_llave.iloc[:, 7].isin(cptosisr))]
#Hacemos un subtotal (sumar.si) para sacar la base gravable por RFCPLAZA
# Paso 1: Agrupar por columna 0 y sumar columna 8 (como SUMAR.SI)
suma = baseISR_temporal.groupby(baseISR_temporal.columns[0])[baseISR_temporal.columns[8]].sum().reset_index()

# Paso 2: Renombramos la columna de suma (opcional)
suma.columns = [baseISR_temporal.columns[0], 'SUBTOTAL']

# Paso 3: Tomamos columnas 0, 1 y 3 del DataFrame base, quitando duplicados
datos_unicos = baseISR_temporal.iloc[:, [0, 1, 3]].drop_duplicates(subset=baseISR_temporal.columns[0])

# Paso 4: Hacemos el merge con la suma
resultado_final = datos_unicos.merge(suma, on=baseISR_temporal.columns[0], how='left')

# Paso 5: Reordenamos las columnas si es necesario (en este orden: 0, 1, 3, suma)
resultado_final = resultado_final[[baseISR_temporal.columns[0], baseISR_temporal.columns[1], baseISR_temporal.columns[3], 'SUBTOTAL']]
baseISR = resultado_final.copy()
# endregion


# region Cálculo de ISR

#Cálculo de ISR
#PROCESO PARA EL CÁLCULO DEL ISR
#Definimos las tablas y procesos para el cálculo, como una tipo Macro o función que nos regresa los valores calculados que despues nos mostrará cuando las llamemos
#Definimos la tabla de ISR 
tabla_isr = pd.DataFrame({
    "limite_inferior": [0.01, 746.05, 6332.06, 11128.02, 12935.83, 15487.72, 31236.50, 49233.01, 93993.91, 125325.21, 375975.62],
    "limite_superior": [746.04, 6332.05, 11128.01, 12935.82, 15487.71, 31236.49, 49233.00, 93993.90, 125325.20, 375975.61, np.inf],
    "cuota_fija": [0.00, 14.32, 371.83, 893.63, 1182.88, 1640.18, 5004.12, 9236.89, 22665.17, 32691.18, 117912.32],
    "porcentaje": [0.0192, 0.0640, 0.1088, 0.1600, 0.1792, 0.2136, 0.2352, 0.3000, 0.3200, 0.3400, 0.3500]
})

#Creamos la función de cálculo de ISR
def calcular_isr(importe_qnal):
    # Convertir a mensual (x2)
    importe_mensual = importe_qnal * 2
    
    # Calcular ISR según la tabla
    fila_isr = tabla_isr[(importe_mensual >= tabla_isr["limite_inferior"]) & (importe_mensual <= tabla_isr["limite_superior"])].iloc[0]
    
    impuesto = fila_isr["cuota_fija"] + ((importe_mensual - fila_isr["limite_inferior"]) * fila_isr["porcentaje"])
    
    return impuesto
    
#CALCULAMOS EL ISR DE LA PRENÓMINA ORDINARIA
#Primero preparamos el DataFrame para aplicar la función de cálculo de ISR
rfcplazaisr_temporal = baseISR.drop_duplicates(subset=["rfcplaza"], keep="first").copy()

#Calculamos el ISR para cada RFC único y agregamos una columna donde pongamos el resultado del cálculo
rfcplazaisr_temporal["ISR"] = rfcplazaisr_temporal["SUBTOTAL"].apply(calcular_isr)
ISR = rfcplazaisr_temporal.copy()



# endregion


# region Base gravable para cuotas ISSSTE

#Filtramos las filas en donde sean puras percepcciones y ademas que graben para ISSSTE
baseISSSTE_temporal = df_perded_llave[(df_perded_llave.iloc[:, 6] == "P") & (df_perded_llave.iloc[:, 7].isin(cptosisste))]



#Hacemos el subtotal por RFCPLAZA
# Paso 1: Agrupar por columna 0 y sumar columna 8 (como SUMAR.SI)
suma1 = baseISSSTE_temporal.groupby(baseISSSTE_temporal.columns[0])[baseISSSTE_temporal.columns[8]].sum().reset_index()

# Paso 2: Renombramos la columna de suma (opcional)
suma1.columns = [baseISSSTE_temporal.columns[0], 'SUBTOTAL']

# Paso 3: Tomamos columnas 0, 1 y 3 del DataFrame base, quitando duplicados
datos_unicos1 = baseISSSTE_temporal.iloc[:, [0, 1, 3]].drop_duplicates(subset=baseISSSTE_temporal.columns[0])

# Paso 4: Hacemos el merge con la suma
resultado_final1 = datos_unicos1.merge(suma1, on=baseISSSTE_temporal.columns[0], how='left')

# Paso 5: Reordenamos las columnas si es necesario (en este orden: 0, 1, 3, suma)
resultado_final1 = resultado_final1[[baseISSSTE_temporal.columns[0], baseISSSTE_temporal.columns[1], baseISSSTE_temporal.columns[3], 'SUBTOTAL']]
baseISSSTE_rfcplaza = resultado_final1.copy()


#Hacemos el subtotal por RFC
# Paso 1: Agrupar por columna 1 y sumar columna 8 (como SUMAR.SI)
suma2 = baseISSSTE_temporal.groupby(baseISSSTE_temporal.columns[1])[baseISSSTE_temporal.columns[8]].sum().reset_index()

# Paso 2: Renombramos la columna de suma (opcional)
suma2.columns = [baseISSSTE_temporal.columns[1], 'SUBTOTAL']  # <- ahora coincide con el agrupamiento

# Paso 3: Tomamos columnas 0, 1 y 3 del DataFrame base, quitando duplicados por RFC (columna 1)
datos_unicos2 = baseISSSTE_temporal.iloc[:, [0, 1, 3]].drop_duplicates(subset=baseISSSTE_temporal.columns[1])

# Paso 4: Hacemos el merge con la suma usando columna 1 (RFC)
resultado_final2 = datos_unicos2.merge(suma2, on=baseISSSTE_temporal.columns[1], how='left')

# Paso 5: Reordenamos las columnas si es necesario (en este orden: 0, 1, 3, suma)
resultado_final2 = resultado_final2[[baseISSSTE_temporal.columns[0], baseISSSTE_temporal.columns[1], baseISSSTE_temporal.columns[3], 'SUBTOTAL']]

# Guardamos resultado
baseISSSTE_rfc = resultado_final2.copy()

# endregion


# region Cálculo de Cuotas ISSSTE
#Hacemos un nuevo DF en donde hagamos un cruce para tener ambas bases gravables, por plaza y rfc
#Obtener nombre de columna RFC
col_rfc = baseISSSTE_rfcplaza.columns[1]

#Obtener columna de valor a traer desde baseISSSTE_rfc (columna índice 3)
col_valor = baseISSSTE_rfc.columns[3]

#Renombrar temporalmente columna para no generar conflictos
baseISSSTE_rfc_renamed = baseISSSTE_rfc[[col_rfc, col_valor]].copy()
baseISSSTE_rfc_renamed = baseISSSTE_rfc_renamed.rename(columns={col_valor: 'SUBTOTAL_RFC'})

#Hacemos el merge para traer el valor
baseISSSTE_rfcplaza = baseISSSTE_rfcplaza.merge(baseISSSTE_rfc_renamed, on=col_rfc, how='left')

# Aseguramos que las columnas 3 y 4 (índices) sean numéricas
baseISSSTE_rfcplaza.iloc[:, 3] = pd.to_numeric(baseISSSTE_rfcplaza.iloc[:, 3], errors='coerce')
baseISSSTE_rfcplaza.iloc[:, 4] = pd.to_numeric(baseISSSTE_rfcplaza.iloc[:, 4], errors='coerce')

# Calculamos la columna FACTOR = columna 3 / columna 4
baseISSSTE_rfcplaza['FACTOR'] = baseISSSTE_rfcplaza.apply(lambda row: row.iloc[3] / row.iloc[4] if pd.notna(row.iloc[4]) and row.iloc[4] != 0 else np.nan, axis=1)
basesgravables_ISSSTE = baseISSSTE_rfcplaza.copy()

#Calculamos la cuota CV y la agregamos en una columna al final
basesgravables_ISSSTE["CV"] = np.where(
    basesgravables_ISSSTE.iloc[:, 4] > 16971,
    (1039.47 - (basesgravables_ISSSTE.iloc[:, 4] * 0.06125)) * basesgravables_ISSSTE.iloc[:, 5],
    basesgravables_ISSSTE.iloc[:, 3] * 0.06125
).round(2)

#Calculamos la cuota SI y la agregamos en una columna al final
basesgravables_ISSSTE["SI"] = np.where(
    basesgravables_ISSSTE.iloc[:, 4] > 16971,
    (106.06 - (basesgravables_ISSSTE.iloc[:, 4] * 0.00625)) * basesgravables_ISSSTE.iloc[:, 5],
    basesgravables_ISSSTE.iloc[:, 3] * 0.00625
).round(2)

#Calculamos la cuota SO y la agregamos en una columna al final
basesgravables_ISSSTE["SO"] = np.where(
    basesgravables_ISSSTE.iloc[:, 4] > 16971,
    (84.85 - (basesgravables_ISSSTE.iloc[:, 4] * 0.005)) * basesgravables_ISSSTE.iloc[:, 5],
    basesgravables_ISSSTE.iloc[:, 3] * 0.005
).round(2)

#Calculamos la cuota SS y la agregamos en una columna al final
basesgravables_ISSSTE["SS"] = np.where(
    basesgravables_ISSSTE.iloc[:, 4] > 16971,
    (572.77 - (basesgravables_ISSSTE.iloc[:, 4] * 0.03375)) * basesgravables_ISSSTE.iloc[:, 5],
    basesgravables_ISSSTE.iloc[:, 3] * 0.03375
).round(2)


# endregion



basesgravables_ISSSTE.to_excel(r'C:\Users\Maxruso7\Desktop\INCREMENTO 2025\basesgrvables_ISSSTE_15QNAS.xlsx')
print("Finalizado")