#!/usr/bin/env python
# coding: utf-8

# ### Importación de las librerias necesarias

# In[1]:


import pandas as pd
import numpy as np


# ### Carga de los archivos pre-requisites

# In[2]:


df_espacios = pd.read_excel('1_ESPACIOS_FISICOS.xlsx', 'MAESTRA V2_Final')
df_profesionales = pd.read_excel('2_PROFESIONALES.xlsx', 'MAESTRA')
df_disp_espacios = pd.read_excel('3_DISPONIBILIDAD_ESPACIOS_FISICOS.xlsx', 'MAESTRA')
df_disp_profesionales = pd.read_excel('4_DISPONIBILIDAD_DEL_PERSONAL.xlsx', 'MAESTRA')


# ### Renombramiento y formateo de los atributos de las tablas
# 
# Renombrar columnas en df_disp_profesionales:
# 
# Se utiliza el método rename() de pandas para cambiar los nombres de las columnas en el DataFrame df_disp_profesionales. Los nombres de las columnas se modifican según un diccionario proporcionado. Por ejemplo, la columna 'Id Sede' se renombra como 'id_sede'.
# Convertir columnas a tipo de datos específico en df_disp_profesionales:
# 
# Se utilizan los métodos astype() y str.upper().str.strip() para convertir las columnas 'id_persona', 'id_sede' y 'dia' del DataFrame df_disp_profesionales a tipos de datos específicos y para realizar operaciones de limpieza en los valores de las columnas.
# Filtrar columnas y combinar DataFrames en df_disp_profesionales:
# 
# Se utiliza el método filter() con una expresión regular para seleccionar columnas que comiencen con 'h' seguidas de un número entre 0 y 23. Luego se utiliza fillna(0) para rellenar los valores faltantes en esas columnas con ceros. El DataFrame resultante se guarda en la variable tempi. Después, se utiliza pd.concat() para combinar las columnas seleccionadas con las columnas 'id_sede', 'id_persona' y 'dia' en df_disp_profesionales.
# Los pasos 1-3 se repiten para los DataFrames df_disp_espacios, df_profesionales y df_espacios con diferentes columnas y manipulaciones de datos.

# In[3]:


df_disp_profesionales = df_disp_profesionales.rename(
    columns = {
        'Id Sede' : 'id_sede',
        'Codigo Succes Factors' : 'id_persona',
        'Día' : 'dia'
    }
)
df_disp_profesionales['id_persona'] = df_disp_profesionales['id_persona'].astype(str)
df_disp_profesionales['id_sede'] = df_disp_profesionales['id_sede'].astype(str)
df_disp_profesionales['dia'] = df_disp_profesionales['dia'].str.upper().str.strip()

tempi = df_disp_profesionales.filter(regex=r'^h([0-9]|(1[0-9])|(2[0-3]))$', axis=1).fillna(0) #.astype(bool)
df_disp_profesionales = pd.concat([df_disp_profesionales[['id_sede', 'id_persona', 'dia']], tempi], axis=1)

df_disp_espacios = df_disp_espacios.rename(
    columns = {
        'Id Sede' : 'id_sede',
        'Id_Espacio_Físico' : 'id_espacio',
        'Día' : 'dia'
    }
)
df_disp_espacios['id_sede'] = df_disp_espacios['id_sede'].astype(str)
df_disp_espacios['id_espacio'] = df_disp_espacios['id_espacio'].str.upper().str.strip()
df_disp_espacios['dia'] = df_disp_espacios['dia'].str.upper().str.strip()

tempi = df_disp_espacios.filter(regex=r'^h([0-9]|(1[0-9])|(2[0-3]))$', axis=1).fillna(0) #.astype(bool)
df_disp_espacios = pd.concat([df_disp_espacios[['id_sede', 'id_espacio', 'dia']], tempi], axis=1)
df_profesionales = df_profesionales.rename(
    columns = {
        'Id_Sede' : 'id_sede',
        'Codigo Sf (Id Profesional)' : 'id_persona',
        'Jornada Semanal Sede' : 'total_horas_semana',
        'Horas Virtuales (No Consultorio) Semanal' : 'horas_virtual_semana',
        'Tipo De Contrato' : 'tipo_contrato',
        'Qualification Cargo (Id Qualification)' : 'cualificacion'
    }
)
df_profesionales['id_sede'] = df_profesionales['id_sede'].astype(str)
df_profesionales['id_persona'] = df_profesionales['id_persona'].astype(str)
df_profesionales['cualificacion'] = df_profesionales['cualificacion'].str.upper().str.strip()
df_profesionales['tipo_contrato'] = df_profesionales['tipo_contrato'].str.upper().str.strip()
df_profesionales['horas_presencial_semana'] = (df_profesionales['total_horas_semana'] - df_profesionales['horas_virtual_semana']).round(0)

df_profesionales = df_profesionales[['id_sede', 'id_persona', 'cualificacion', 'tipo_contrato', 'horas_presencial_semana']]

# solo dejar profesionales con horas presenciales
df_profesionales = df_profesionales[df_profesionales['horas_presencial_semana'] > 0]
df_profesionales['horas_presencial_quincena'] = df_profesionales['horas_presencial_semana'].apply(lambda x: x * 2)


df_espacios = df_espacios.rename(
    columns = {
        'Id_Sede' : 'id_sede',
        'Id_espacio_físico' : 'id_espacio',
        'Id_Qualification2' : 'cualificacion'
    }
)

df_espacios['id_sede'] = df_espacios['id_sede'].astype(str)
df_espacios['id_espacio'] = df_espacios['id_espacio'].str.upper().str.strip()
df_espacios['cualificacion'] = df_espacios['cualificacion'].str.upper().str.strip()
df_espacios = df_espacios[['id_sede', 'id_espacio', 'cualificacion']]
df_espacios


# ### Formato de tildes, dejar los nombres de los días sin ellas.

# In[4]:


diccionario_tildes = {
    'Á': 'A',
    'É': 'E',
    'Í': 'I',
    'Ó': 'O',
    'Ú': 'U'
}
df_disp_espacios['dia'] = df_disp_espacios['dia'].replace(diccionario_tildes, regex = True)
df_disp_profesionales['dia'] = df_disp_profesionales['dia'].replace(diccionario_tildes, regex = True)


# ### Validacion de llaves
# 
# Se crea una lista llamada errores_llaves para almacenar los errores en las claves.
# 
# Se verifica la unicidad de la clave en el DataFrame df_espacios y se agrega un diccionario a errores_llaves si hay duplicados.
# 
# Se repite el paso 2 para los DataFrames df_profesionales, df_disp_espacios y df_disp_profesionales, agregando diccionarios a errores_llaves si se encuentran duplicados en las claves correspondientes.
# 
# Se utiliza un bloque try-except para imprimir un mensaje dependiendo de si se encontraron errores en las claves.
# 
# Si se captura una excepción, se imprime la lista de errores.

# In[5]:


errores_llaves = []

# 1
dif_esp = df_espacios['id_sede'].str.cat([df_espacios['id_espacio'], df_espacios['cualificacion']], sep='-').drop_duplicates().size
if len(df_espacios) != dif_esp :
  errores_llaves.append({'Maestro': 'Espacios', 'Total': len(df_espacios), 'Unicos': dif_esp})

# 2
dif_prof = df_profesionales['id_sede'].str.cat(df_profesionales['id_persona'], sep='-').drop_duplicates().size
if len(df_profesionales) != dif_prof :
  errores_llaves.append({'Maestro': 'Profesionales', 'Total': len(df_profesionales), 'Unicos': dif_prof})

# 3
dif_disp_esp = df_disp_espacios['id_sede'].str.cat([df_disp_espacios['id_espacio'], df_disp_espacios['dia']], sep='-').drop_duplicates().size
if len(df_disp_espacios) != dif_disp_esp :
  errores_llaves.append({'Maestro': 'Disp. espacios', 'Total': len(df_disp_espacios), 'Unicos': dif_disp_esp})

# 4
dif_disp_prof = df_disp_profesionales['id_sede'].str.cat([df_disp_profesionales['id_persona'], df_disp_profesionales['dia']], sep='-').drop_duplicates().size
if len(df_disp_profesionales) != dif_disp_prof :
  errores_llaves.append({'Maestro': 'Disp. profesionales', 'Total': len(df_disp_profesionales), 'Unicos': dif_disp_prof})


try :
  if len(errores_llaves) > 0 :
    raise Exception
  print('Llaves OK')
except Exception :
  print('Estos son los maestros en donde la llave no es unica: \n', errores_llaves)


# ### Comienza la creación de las hojas

# In[76]:


# [P_PQ(p,q)] Personas :: Persona, Qualificacion, (1)

df_h_profesionales = df_profesionales[['id_sede', 'id_persona', 'cualificacion']]
df_h_profesionales['valor'] = 1
df_h_profesionales


# In[77]:


# [P_CQ(c,q)]  Espacios :: consultorio, Qualificacion, (1)

df_h_espacios = df_espacios[['id_sede', 'id_espacio', 'cualificacion']]
df_h_espacios['valor'] = 1

#Que solo estén los de la anterior hoja

# Obtener las cualificaciones únicas del primer DataFrame
cualificaciones_unicas = df_h_profesionales['cualificacion'].unique()

# Filtrar el segundo DataFrame para solo las filas donde la 'cualificacion' también esté en el primer DataFrame
df_h_espacios = df_h_espacios[df_h_espacios['cualificacion'].isin(cualificaciones_unicas)]


# In[72]:


# [P_Di_PC(p,c)] personas x consultorios :: persona, espacio, (1)

df_full_esp_prof = pd.merge(df_espacios, df_profesionales, how = 'outer', on = ['id_sede', 'cualificacion'])

# profesionales_sin_consultorios
profesionales_sin_consultorios = df_full_esp_prof[df_full_esp_prof['id_espacio'].isna()] # problemas!


# consultorios_sin_profesionales
esp_sin_prof_cualif = df_full_esp_prof[df_full_esp_prof['id_persona'].isna()]
h_personas_espacios = df_full_esp_prof[~df_full_esp_prof['id_persona'].isna() & ~df_full_esp_prof['id_espacio'].isna()]
h_personas_espacios['combinacion factible persona-consultorio (1:si)'] = 1


consultorios_sin_profesionales = pd.merge(esp_sin_prof_cualif, h_personas_espacios, how = 'left', on = ['id_sede', 'id_espacio'])
consultorios_sin_profesionales = consultorios_sin_profesionales[consultorios_sin_profesionales['id_persona_y'].isna()] # problemas tambien!


hayErrores = False

if profesionales_sin_consultorios.size > 0 :
  profesionales_sin_consultorios.to_excel('profesionales_sin_consultorios.xlsx', index = False)
  hayErrores = True

if consultorios_sin_profesionales.size > 0 :
  consultorios_sin_profesionales.to_excel('consultorios_sin_profesionales.xlsx', index = False)
  hayErrores = True


try :
  if hayErrores :
    raise Exception
  print('Todos los profesionales pueden usar al menos un espacio fisico')
  print('Todos los consultorios pueden ser usados por al menos un profesional')
except Exception :
  # si falloo por los profesionales:
  print('En el maestro de profesionales hay personas con horas presenciales que no se pueden asignar porque la sede a la que pertenecen no tiene algun espacio fisico que tenga la cualificación. Por favor revisar el archivo "profesionales_sin_consultorios.xlsx"')


# ### Eliminando registros de aquellos profesionales que no tengan consultorios y consultorios que no tengan profesionales
# 
# Esta parte del código se está utilizando para limpiar los dataframes y asegurarse de que solo contengan registros de profesionales que están asignados a consultorios y consultorios que tienen profesionales asignados. Esto es útil para asegurarse de que los datos son coherentes y están completos antes de realizar análisis adicionales o modelado.

# In[9]:


# Definiendo las listas de personas sin consultorio y consultorios sin persona
unique_id_persona_sin_consultorio = profesionales_sin_consultorios['id_persona'].unique().tolist()
unique_id_consultorio_sin_persona=consultorios_sin_profesionales['id_espacio'].unique().tolist()
#Eliminando dichos registros de los dataframes relacionados con profesionales creados hasta ahora
df_disp_profesionales.drop(df_disp_profesionales[df_disp_profesionales['id_persona'].isin(unique_id_persona_sin_consultorio)].index,inplace=True)
df_profesionales.drop(df_profesionales[df_profesionales['id_persona'].isin(unique_id_persona_sin_consultorio)].index, inplace=True)
df_h_profesionales.drop(df_h_profesionales[df_h_profesionales['id_persona'].isin(unique_id_persona_sin_consultorio)].index, inplace=True)
#Mismo proceso pero para consultorios
df_h_espacios.drop(df_h_espacios[df_h_espacios['id_espacio'].isin(unique_id_consultorio_sin_persona)].index, inplace=True)
df_disp_espacios.drop(df_disp_espacios[df_disp_espacios['id_espacio'].isin(unique_id_consultorio_sin_persona)].index,inplace=True)
df_espacios.drop(df_espacios[df_espacios['id_espacio'].isin(unique_id_consultorio_sin_persona)].index, inplace=True)


# In[10]:


# se transponen las tablas de disponibilidad
df_disp_profesionales_transp = df_disp_profesionales.melt(id_vars=['id_sede', 'id_persona', 'dia'], var_name = 'hora', value_name = 'disponible')
df_disp_espacios_transp = df_disp_espacios.melt(id_vars=['id_sede', 'id_espacio', 'dia'], var_name = 'hora', value_name = 'disponible')
df_disp_espacios_transp


# In[11]:


# quitar "h"s de las horas
df_disp_profesionales_transp['hora'] = df_disp_profesionales_transp['hora'].str.extract(r'(\d+)', expand = False).astype(int)
df_disp_espacios_transp['hora'] = df_disp_espacios_transp['hora'].str.extract(r'(\d+)', expand = False).astype(int)
df_disp_espacios_transp


# In[12]:


# [P_Di_PDH(p,d,h)] disp. personas :: persona, dia, hora, (1)
h_disp_personas = df_disp_profesionales_transp[df_disp_profesionales_transp['disponible'] == True]
h_disp_personas['esta la persona disponible (1:si)'] = 1


# In[13]:


# [P_Di_CDH(c,d,h)] disp. espacios :: espacios, dia, hora, (0)
h_disp_espacios = df_disp_espacios_transp[df_disp_espacios_transp['disponible'] == True]


# In[14]:


# cada profesional y espacio debe tener al menos una hora disponible

# profesional
columnas_suma_prof = df_disp_profesionales.columns[3:27]
df_disp_profesionales[columnas_suma_prof] = df_disp_profesionales[columnas_suma_prof].replace(' ', 0).astype(float)
df_disp_profesionales['suma'] = df_disp_profesionales[columnas_suma_prof].sum(axis = 1)

df_prof_sin_disp = pd.merge(df_profesionales, df_disp_profesionales[df_disp_profesionales['suma'] > 0], how = 'left', on = ['id_sede', 'id_persona'])
df_prof_sin_disp = df_prof_sin_disp[df_prof_sin_disp['suma'].isna()]
#print(df_esp_sin_disp)


# espacio
columnas_suma_esp = df_disp_espacios.columns[3:27]
df_disp_espacios[columnas_suma_esp] = df_disp_espacios[columnas_suma_esp].replace(' ', 0).astype(float)
df_disp_espacios['suma'] = df_disp_espacios[columnas_suma_esp].sum(axis = 1)

df_esp_sin_disp = pd.merge(df_espacios, df_disp_espacios[df_disp_espacios['suma'] > 0], how = 'left', on = ['id_sede', 'id_espacio'])
df_esp_sin_disp = df_esp_sin_disp[df_esp_sin_disp['suma'].isna()]
#print(df_esp_sin_disp)



# y si comparamos cuantas horas tienen los profesionales vs las que deben ser asignadas?
# Deben tener como minimo disponible el mismo numero de horas que deben trabajar


try :
  if len(df_prof_sin_disp) > 0 | len(df_esp_sin_disp) > 0 :
    raise Exception
  print('Todos los profesionales cuentan con al menos una hora de disponibilidad')
  print('Todos los consultorios cuentan con al menos una hora de disponibilidad')
except Exception :
  print('Hay problemas, ya sea con profesionales que no pueden ser asignados o a espacios fisicos que no pueden ser asignados. Revisar archivo(s) de salida.')
  df_prof_sin_disp.to_excel('Profesionales_sin_disponibilidad.xlsx', index = False)
  df_esp_sin_disp.to_excel('Espacios_sin_disponibilidad.xlsx', index = False)




# In[15]:


# [P_Hmax_PD(p,d)] horas maximas

df_h_horas_max = pd.DataFrame({ 'Persona' : [], 'Horas presenciales contratadas semana' : [] })
df_h_horas_max


# In[64]:


# [P_Hmin_PD(p,d)] horas minimas :: persona, dia, horas_min

h_horas_min_t = df_profesionales[~df_profesionales['tipo_contrato'].str.contains('PRESTA')][['id_sede', 'id_persona']]
h_horas_min_t['horas_min'] = 4

opciones_dia = df_disp_profesionales['dia'].unique().tolist()

h_horas_min = pd.DataFrame()

for dia in opciones_dia:
  tempi = h_horas_min_t
  tempi['dia'] = dia
  h_horas_min = pd.concat([h_horas_min, tempi])


# In[67]:


#Quitando los festivos para P_HMIN_PD

h_horas_min = h_horas_min[h_horas_min['dia'] != 'FESTIVO']
print(h_horas_min.shape)
h_horas_min = h_horas_min[h_horas_min['dia'] != 'FESTIVO 2']
print(h_horas_min.shape)


# In[66]:


h_horas_min


# In[17]:


# [P_Di_TH(t,h)] tipos de turnos :: "T"tiempo_horaini, horaini, (1)

max_horas = 12

tempi = df_disp_espacios_transp[df_disp_espacios_transp['disponible'] == True]
tempi = tempi.groupby('id_sede')['hora'].agg(['min', 'max'])

l_sedes = []
l_turnos = []
l_horas = []

for index, row in tempi.iterrows():
    for i in range(1, max_horas + 1):
        for j in range(row['min'], row['max'] + 2 - i):
            for k in range(j, j + i):
                l_sedes.append(index)
                l_turnos.append(f"T{i}_{j}")
                l_horas.append(k)

h_turnos = pd.DataFrame({'id_sede': l_sedes, 'turno': l_turnos, 'hora': l_horas})
h_turnos['el turno esta activo (1:si)'] = 1
h_turnos


# In[18]:


# [P_Hcon_P(p)] horas presenciales :: persona, horas quincena

h_horas = df_profesionales[['id_sede', 'id_persona', 'horas_presencial_quincena']]


# In[68]:


# [P_Tmax_PD(p,d)] combinaciones turnos-consultorios maximo permitido :: persona, dia, num combinaciones turnos-consultorios maximo permitido

# Función para contar los conjuntos de unos en una fila
def contar_conjuntos_unos(row):
    contador = 0
    conjuntos = 0

    for valor in row[3:]:
        if valor == 1:
            contador += 1
        else:
            if contador > 0:
                conjuntos += 1
            contador = 0

    if contador > 0:
        conjuntos += 1

    return conjuntos

# Crear nueva columna con la cantidad de conjuntos de unos por registro
df_disp_profesionales['max_turnos'] = df_disp_profesionales.apply(contar_conjuntos_unos, axis = 1)


h_combinaciones_turnos = df_disp_profesionales[['id_sede', 'id_persona', 'dia', 'max_turnos']]

h_combinaciones_turnos = h_combinaciones_turnos[h_combinaciones_turnos['dia'] != 'FESTIVO']
h_combinaciones_turnos = h_combinaciones_turnos[h_combinaciones_turnos['dia'] != 'FESTIVO 2']
h_combinaciones_turnos


# inicialmente las personas que tiene las disponibilidad continua, van a tener un unico turno (eso es lo que está plastamado en la tabla)
# es posible que al momento de ejecutar el modelo de optimización queden muchas horas sin asignar, se puede ejecutar una segunda vez, reemplazando los 1s por 0s
# si es mucho mas optimo, se modificaria este trozo del codigo



# In[20]:


# [3*] Dias semana

dias_semana = {
    'semana' : [
        'semana_1', 'semana_1', 'semana_1', 'semana_1', 'semana_1', 'semana_1', 'semana_1',
        'semana_2', 'semana_2', 'semana_2', 'semana_2', 'semana_2', 'semana_2', 'semana_2',
    ],
    'dia' : [
        'LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO', 'DOMINGO',
        'LUNES 2', 'MARTES 2', 'MIERCOLES 2', 'JUEVES 2', 'VIERNES 2','SABADO 2', 'DOMINGO 2'
    ]
}

df_h_dias_semana = pd.DataFrame(dias_semana)


# ### Creacion de los archivos

# #### Preparando los archivos, renombrando columnas y demás

# In[21]:


# Preparando las hojas para el export:
sedes=df_profesionales['id_sede'].unique()
#P_PQ(p,q)
#df_h_profesionales_test=df_h_profesionales
df_h_profesionales.rename(columns={'id_persona': 'Persona', 'cualificacion': 'Qualificacion'}, inplace=True)
df_h_profesionales = df_h_profesionales[['Persona', 'Qualificacion', 'valor','id_sede']]
df_h_profesionales

#P_CQ
df_h_espacios.rename(columns={'id_espacio':'consultorio','cualificacion':'Qualificacion'},inplace=True)
df_h_espacios=df_h_espacios[['consultorio','Qualificacion','valor','id_sede']]
df_h_espacios

#P_Di_PC

h_personas_espacios.rename(columns={'id_persona':'persona','id_espacio':'consultorio'},inplace=True)
h_personas_espacios=h_personas_espacios[['persona','consultorio','combinacion factible persona-consultorio (1:si)','id_sede']]
h_personas_espacios

#P_Di_PDH
h_disp_personas.rename(columns={'id_persona':'persona'},inplace=True)
h_disp_personas=h_disp_personas[['persona','dia','hora','esta la persona disponible (1:si)','id_sede']]
h_disp_personas

#P_Di_PCH

h_disp_espacios.rename(columns={'id_espacio':'consultorio','disponible':'esta disponible el consultorio (0:no)'},inplace=True)


#P_Hmax_PD -> df_h_horas_max --> hoja en blanco

#P_Hmin_PD

h_horas_min.rename(columns={'id_persona':'persona', 'horas_min':'horas minimas de trabajo'},inplace=True)
h_horas_min=h_horas_min[['persona','dia','horas minimas de trabajo','id_sede']]


#P_Di_TH

h_turnos=h_turnos[['turno','hora','el turno esta activo (1:si)','id_sede']]

#P_Hcon_P

h_horas.rename(columns={'id_persona':'Persona', 'horas_presencial_quincena':'Horas presenciales contratadas periodo'},inplace=True)
h_horas=h_horas[['Persona','Horas presenciales contratadas periodo','id_sede']]

#P_Tmax_PD
h_combinaciones_turnos.rename(columns={'id_persona':'persona','max_turnos':'combinaciones turnos-consultorios maximo permitido'},inplace=True)

#uno DFs en blanco --------------- PENDING

h_sim_PDD=pd.DataFrame()
h_sim_PCDD=pd.DataFrame()
h_p_fat=pd.DataFrame()

#dias semana
df_dias_semana = pd.DataFrame(dias_semana)


# In[22]:


#P_FAT10_PDHDH

# Mapeamos los días de la semana a números
dia_a_numero = {
    'LUNES': 1,
    'MARTES': 2,
    'MIERCOLES': 3,
    'JUEVES': 4,
    'VIERNES': 5,
    'SABADO': 6,
    'DOMINGO': 7,
    'LUNES 2': 8,
    'MARTES 2': 9,
    'MIERCOLES 2': 10,
    'JUEVES 2': 11,
    'VIERNES 2': 12,
    'SABADO 2': 13,
    'DOMINGO 2': 14
}

# Aplicamos la transformación al dataframe
h_disp_personas['dia_num'] = h_disp_personas['dia'].map(dia_a_numero)

# Creamos una copia ordenada del dataframe
h_disp_personas_ordenado = h_disp_personas.sort_values(['persona', 'dia_num', 'hora'])

# Creamos un nuevo dataframe con la información desplazada, solo para 'dia', 'hora', y 'dia_num'
h_disp_personas_ordenado[['dia2', 'hora2', 'dia_num2']] = h_disp_personas_ordenado.groupby('persona')[['dia', 'hora', 'dia_num']].shift(-1)

# Ahora, tenemos que tener en cuenta que cuando cambiamos de día, la hora se reinicia a 0.
# Por lo tanto, cuando calculamos la diferencia de horas, si es negativa, significa que hemos cambiado de día y debemos agregar 24 a la hora del segundo día.
h_disp_personas_ordenado['diff_hora'] = (h_disp_personas_ordenado['hora2'] + 24) - h_disp_personas_ordenado['hora']

# Creamos una máscara para filtrar solo aquellos registros donde la diferencia de tiempo es menor a 10 horas y el día es consecutivo.
mask = ((h_disp_personas_ordenado['diff_hora'] <= 10) & (h_disp_personas_ordenado['dia_num2'] - h_disp_personas_ordenado['dia_num'] == 1))

# Filtramos el dataframe original con la máscara para obtener las filas conflictivas
conflictos = h_disp_personas_ordenado.loc[mask]

# Seleccionamos solo las columnas que nos interesan, incluyendo 'id_sede'
conflictos = conflictos[['persona', 'dia', 'hora', 'id_sede', 'dia2', 'hora2']]

# Eliminamos las filas donde dia2 o hora2 son NaN (esto sucederá en la última fila de cada grupo)
conflictos = conflictos.dropna(subset=['dia2', 'hora2'])

# Convertimos hora2 a int ya que después de shift se convierte en float
conflictos['hora2'] = conflictos['hora2'].astype(int)

h_p_fat=conflictos


# In[23]:


#Escitura de los archivos:
'''
for sede in sedes:
    ruta='C:\\Users\\juanp\\Desktop\\Work\\Synaptica\\SURA\\Inputs\\Results\\'
    nombre_archivo=f'archivo_excel_{sede}.xlsx'
    writer=pd.ExcelWriter(ruta+nombre_archivo,engine='xlsxwriter')
    df_h_profesionales[df_h_profesionales['id_sede']==sede].to_excel(writer,sheet_name='P_PQ(p,q)')
    df_h_espacios[df_h_espacios['id_sede']==sede].to_excel(writer,sheet_name='P_CQ(c,q)')
    h_personas_espacios[h_personas_espacios['id_sede']==sede].to_excel(writer,sheet_name='P_Di_PC(p,c)')
    h_disp_personas[h_disp_personas['id_sede']==sede].to_excel(writer,sheet_name='P_Di_PDH(p,d,h)')
    h_disp_espacios[h_disp_espacios['id_sede']==sede].to_excel(writer,sheet_name='P_Di_CDH(c,d,h)')
    df_h_horas_max.to_excel(writer,sheet_name='P_Hmax_PD(p,d)')
    h_horas_min[h_horas_min['id_sede']==sede].to_excel(writer,sheet_name='P_Hmin_PD')
    h_turnos[h_turnos['id_sede']==sede].to_excel(writer,sheet_name='P_Di_TH(t,h)')
    h_horas[h_horas['id_sede']==sede].to_excel(writer,sheet_name='P_Hcon_P(p)')
    h_combinaciones_turnos[h_combinaciones_turnos['id_sede']==sede].to_excel(writer,sheet_name='P_Tmax_PD(p,d)')
    h_sim_PDD.to_excel(writer,sheet_name='P_SIM_PDD_(p,d,d)')
    h_sim_PCDD.to_excel(writer,sheet_name='P_SIM_PCDD_(p,c,d,d)')
    h_p_fat[h_p_fat['id_sede']==sede].to_excel(writer,sheet_name='P_Fat10_PDHDH_(p,d,h,d,h)')
    df_dias_semana.to_excel(writer,sheet_name='Dias_semana')
    writer.save()
'''
for sede in sedes:
    ruta = 'C:\\Users\\juanp\\Desktop\\Work\\Synaptica\\SURA\\Inputs\\Results\\'
    nombre_archivo = f'archivo_excel_{sede}.xlsx'
    writer = pd.ExcelWriter(ruta + nombre_archivo, engine='xlsxwriter')

    # Hoja P_PQ(p,q)
    df_h_profesionales[df_h_profesionales['id_sede'] == sede].drop('id_sede', axis=1).to_excel(writer, sheet_name='P_PQ(p,q)', index=False)

    # Hoja P_CQ(c,q)
    df_h_espacios[df_h_espacios['id_sede'] == sede].drop('id_sede', axis=1).to_excel(writer, sheet_name='P_CQ(c,q)', index=False)

    # Hoja P_Di_PC(p,c)
    h_personas_espacios[h_personas_espacios['id_sede'] == sede].drop('id_sede', axis=1).to_excel(writer, sheet_name='P_Di_PC(p,c)', index=False)

    # Hoja P_Di_PDH(p,d,h)
    h_disp_personas[h_disp_personas['id_sede'] == sede].drop(['id_sede','dia_num'], axis=1).to_excel(writer, sheet_name='P_Di_PDH(p,d,h)', index=False)

    # Hoja P_Di_CDH(c,d,h)
    h_disp_espacios[h_disp_espacios['id_sede'] == sede].drop('id_sede', axis=1).to_excel(writer, sheet_name='P_Di_CDH(c,d,h)', index=False)

    # Hoja P_Hmax_PD(p,d)
    df_h_horas_max.to_excel(writer, sheet_name='P_Hmax_PD(p,d)', index=False)

    # Hoja P_Hmin_PD
    h_horas_min[h_horas_min['id_sede'] == sede].drop('id_sede', axis=1).to_excel(writer, sheet_name='P_Hmin_PD(p,d)', index=False)

    # Hoja P_Di_TH(t,h)
    h_turnos[h_turnos['id_sede'] == sede].drop('id_sede', axis=1).to_excel(writer, sheet_name='P_Di_TH(t,h)', index=False)

    # Hoja P_Hcon_P(p)
    h_horas[h_horas['id_sede'] == sede].drop('id_sede', axis=1).to_excel(writer, sheet_name='P_Hcon_P(p)', index=False)

    # Hoja P_Tmax_PD(p,d)
    h_combinaciones_turnos[h_combinaciones_turnos['id_sede'] == sede].drop('id_sede', axis=1).to_excel(writer, sheet_name='P_Tmax_PD(p,d)', index=False)

    # Hoja P_SIM_PDD_(p,d,d)
    h_sim_PDD.to_excel(writer, sheet_name='P_Sim_PDD_(p,d,d)', index=False)

    # Hoja P_SIM_PCDD_(p,c,d,d)
    h_sim_PCDD.to_excel(writer, sheet_name='P_Sim_PCDD_(p,c,d,d)', index=False)

    # Hoja P_Fat10_PDHDH_(p,d,h,d,h)
    h_p_fat[h_p_fat['id_sede'] == sede].drop('id_sede', axis=1).to_excel(writer, sheet_name='P_Fat10_PDHDH_(p,d,h,d,h)', index=False)

    # Hoja Dias_semana
    df_dias_semana.to_excel(writer, sheet_name='Dias_semana', index=False)

    writer.save()
    


# ### Creacion del archivo .py

# In[24]:


get_ipython().system('jupyter nbconvert --to script inputs.ipynb')

