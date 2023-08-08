import pandas as pd
from mlxtend.preprocessing import TransactionEncoder
from mlxtend.frequent_patterns import apriori, association_rules
# Carga los datos desde un archivo de Excel
data = pd.read_excel('output.xlsx')
# Convierte los datos a una lista de transacciones
transactions = []
for trans_id, group in data.groupby('Nro_transaccion'):
    transactions.append(list(group['Cod_seccion'].values))
# Convierte las transacciones en una matriz binaria
te = TransactionEncoder()
te_ary = te.fit_transform(transactions)
data_bin = pd.DataFrame(te_ary, columns=te.columns_)
# Ejecuta el algoritmo Apriori
frequent_itemsets = apriori(data_bin, min_support=0.01, use_colnames=True, max_len=None)
# Crea las reglas de asociación entre secciones
rules = association_rules(frequent_itemsets, metric="lift", min_threshold=1)
# Crea una lista para almacenar los datos de la salida
output_data = []
# Agrega los datos de cada regla de asociación a la lista de salida
for index, row in rules.iterrows():
    antecedent = set(row['antecedents'])
    consequent = set(row['consequents'])
    mask = (data_bin[list(antecedent.union(consequent))] == 1).all(axis=1)
    transactions_count = mask.sum()
    output_data.append([antecedent, consequent, transactions_count])

# Agrupa las transacciones por sección y calcula el total de transacciones
transactions_per_section = data.groupby('Cod_seccion')['Nro_transaccion'].nunique().reset_index()
transactions_total = data['Nro_transaccion'].nunique()
####################
# Calcular el monto total por Cod_seccion
total_per_section = data.groupby('Cod_seccion')['Total'].sum().reset_index()
# Agregar la columna "Total por sección" al DataFrame "transactions_per_section"
transactions_per_section['TOTAL VENTA (Gs)'] = total_per_section['Total']

# Guarda los resultados en un archivo Excel
with pd.ExcelWriter('resultado_secciones.xlsx') as writer:
    frequent_itemsets.to_excel(writer, sheet_name='Frecuentes', index=False)
    rules.to_excel(writer, sheet_name='Reglas', index=False)
    pd.DataFrame(output_data, columns=['Antecedente', 'Consecuente', 'Número de transacciones']).to_excel(writer, sheet_name='Transacciones',                                                                                                         index=False)
    transactions_per_section.to_excel(writer, sheet_name='Transacciones por sección', index=False)
    pd.DataFrame({'Total de transacciones': [transactions_total]}).to_excel(writer, sheet_name='Total', index=False)
# Obtiene las transacciones únicas por sección
unique_transactions = []
for col in data_bin.columns:
    section_transactions = data_bin.loc[data_bin[col] == 1]
    other_sections = set(data_bin.columns) - {col}
    unique_mask = section_transactions[list(other_sections)].sum(axis=1) == 0
    unique_transactions.append(section_transactions[unique_mask])
# Concatena todas las transacciones únicas en un solo DataFrame
unique_transactions = pd.concat(unique_transactions)
# Crea una lista para almacenar los datos de la salida
output_data_unique = []
# Agrega los datos de cada sección a la lista de salida
for col in data_bin.columns:
    transactions_count = len(unique_transactions.loc[unique_transactions[col] == 1])
    output_data_unique.append([col, transactions_count])
# Crea un DataFrame con los datos de la salida
output_unique_df = pd.DataFrame(output_data_unique, columns=['Sección', 'Transacciones únicas'])
# Guarda los resultados en un archivo Excel
with pd.ExcelWriter('resultado_Secciones.xlsx') as writer:
    frequent_itemsets.to_excel(writer, sheet_name='Frecuentes', index=False)
    rules.to_excel(writer, sheet_name='Reglas', index=False)
    pd.DataFrame(output_data, columns=['Antecedente', 'Consecuente', 'Número de transacciones']).to_excel(writer,sheet_name='Transacciones',                                                                                                          index=False)
    output_unique_df.to_excel(writer, sheet_name='Transacciones únicas', index=False)
    transactions_per_section.to_excel(writer, sheet_name='Transacciones por sección', index=False)
    # Obtiene las transacciones únicas por sección
    unique_transactions = []
    for col in data_bin.columns:
        section_transactions = data_bin.loc[data_bin[col] == 1]
        other_sections = set(data_bin.columns) - {col}
        unique_mask = section_transactions[list(other_sections)].sum(axis=1) == 0
        unique_transactions.append(section_transactions[~unique_mask])
    # Concatena todas las transacciones únicas en un solo DataFrame
    unique_transactions = pd.concat(unique_transactions)
    # Crea una lista para almacenar los datos de la salida
    output_data_unique = []
    # Agrega los datos de cada sección a la lista de salida
    for col in data_bin.columns:
        other_sections = set(data_bin.columns) - {col}
        section_transactions = unique_transactions.loc[unique_transactions[col] == 1]
        other_sections_transactions = section_transactions[list(other_sections)].sum(axis=1) > 0
        transactions_count = other_sections_transactions.sum()
        output_data_unique.append([col, transactions_count])
    # Crea un DataFrame con los datos de la salida
        output_unique_df = pd.DataFrame(output_data_unique, columns=['Sección', 'Transacciones con otras secciones'])
    # Guarda los resultados en un archivo Excel
    with pd.ExcelWriter('resultado_Secciones.xlsx') as writer:
        frequent_itemsets.to_excel(writer, sheet_name='Frecuentes', index=False)
        rules.to_excel(writer, sheet_name='Reglas', index=False)
        pd.DataFrame(output_data, columns=['Antecedente', 'Consecuente', 'Número de transacciones']).to_excel(writer, sheet_name='Transacciones',index=False)
        # Crea una lista para almacenar los datos de las transacciones únicas
        unique_transactions = []
        # Crea una lista para almacenar los datos de las transacciones con otras secciones
        other_transactions = []
        # Separa las transacciones únicas y las transacciones con otras secciones
        for trans_id, group in data.groupby('Nro_transaccion'):
            if len(set(group['Cod_seccion'])) == 1:
                unique_transactions.append(trans_id)
            else:
                other_transactions.append(trans_id)

        # Calcula la cantidad de transacciones únicas y de transacciones con otras secciones
        total_transactions = len(data['Nro_transaccion'].unique())
        unique_transactions_count = len(unique_transactions)
        other_transactions_count = len(other_transactions)
        # Resta la cantidad de transacciones con otras secciones a la cantidad total de transacciones
        result = total_transactions - other_transactions_count
        # Imprime los resultados
        print(f'Total de transacciones: {total_transactions}')
        print(f'Transacciones únicas: {unique_transactions_count}')
        print(f'Transacciones con otras secciones: {other_transactions_count}')
# Crear una copia del DataFrame data sin modificarlo
data_sin_duplicados = data.copy()
# Agrupar los datos por proveedor y calcular el total de nro transacción por proveedores
proveedor_gastos = data.groupby(['NOMBRE_PROVEEDOR'])['Total'].sum().reset_index()
clientes_proveedor = proveedor_gastos.drop_duplicates()
# Eliminar las filas duplicadas para cada Nro Transaccion en la copia
data_sin_duplicados.drop_duplicates(subset=['Nro_transaccion', 'NOMBRE_PROVEEDOR'], inplace=True)
# Ordenar los clientes por monto gastado de mayor a menor
proveedor_gastos = proveedor_gastos.sort_values('Total', ascending=False)
# Calcular el número de transacciones por proveedor en la copia sin duplicados
transacciones_por_proveedor = data_sin_duplicados['NOMBRE_PROVEEDOR'].value_counts().reset_index()
transacciones_por_proveedor = transacciones_por_proveedor.drop_duplicates()
transacciones_por_proveedor.columns = ['NOMBRE_PROVEEDOR', 'Nro_transaccion']
transacciones_por_proveedor = pd.merge(transacciones_por_proveedor, proveedor_gastos[['NOMBRE_PROVEEDOR', 'Total']],on='NOMBRE_PROVEEDOR')
transacciones_por_proveedor = transacciones_por_proveedor.sort_values('Nro_transaccion', ascending=False)
# Obtener el número de transacciones únicas por proveedor
transacciones_unicas = data_sin_duplicados.groupby('NOMBRE_PROVEEDOR')['Nro_transaccion'].nunique().reset_index()
transacciones_unicas.columns = ['NOMBRE_PROVEEDOR', 'Nro_transacciones únicas']

# Obtener el número total de transacciones sin duplicados por proveedor
transacciones_total = data_sin_duplicados.groupby('NOMBRE_PROVEEDOR')['Nro_transaccion'].count().reset_index()

# Obtener el número de transacciones únicas por proveedor
transacciones_unicas = \
data_sin_duplicados[~data_sin_duplicados['Nro_transaccion'].duplicated(keep=False)].groupby('NOMBRE_PROVEEDOR')['Nro_transaccion'].count().reset_index()
transacciones_unicas.columns = ['NOMBRE_PROVEEDOR', 'Transacciones únicas']
# Fusionar los resultados en un solo DataFrame
resultado = transacciones_unicas.merge(transacciones_total, on='NOMBRE_PROVEEDOR', how='left')
# Calcular el número de transacciones no únicas (conjuntos)
resultado['conjuntos'] = resultado['Nro_transaccion'] - resultado['Transacciones únicas']
# Calcular el porcentaje de transacciones únicas
resultado['porcentaje'] = resultado['Transacciones únicas'] / resultado['Nro_transaccion']
# Agregar la columna "Cod_sub_seccion" al DataFrame data si no existe previamente
# Agrupar los datos por proveedor y calcular el total de nro transacción por proveedores
proveedor_gastos = data.groupby(['NOMBRE_PROVEEDOR'])['Total'].sum().reset_index()
clientes_proveedor = proveedor_gastos.drop_duplicates()
# Crear una copia del DataFrame data sin modificarlo
data_sin_duplicados = data.copy()
# Eliminar las filas duplicadas para cada Nro Transaccion en la copia
data_sin_duplicados.drop_duplicates(subset=['Nro_transaccion', 'NOMBRE_PROVEEDOR'], inplace=True)
# Ordenar los clientes por monto gastado de mayor a menor
proveedor_gastos = proveedor_gastos.sort_values('Total', ascending=False)
# Calcular el número de transacciones por proveedor en la copia sin duplicados
##
proveedores_unicos = \
data_sin_duplicados[~data_sin_duplicados['Nro_transaccion'].duplicated(keep=False)].groupby('NOMBRE_PROVEEDOR')[
    'Nro_transaccion'].count().reset_index()
proveedores_unicos.columns = ['NOMBRE_PROVEEDOR', 'Transacciones únicas']
proveedores_unicos = proveedores_unicos.merge(
    data_sin_duplicados.groupby('NOMBRE_PROVEEDOR')['Total'].sum().reset_index(), on='NOMBRE_PROVEEDOR')
proveedores_unicos.columns = ['NOMBRE_PROVEEDOR', 'Transacciones únicas', 'Total']
proveedores_cod_subseccion = \
data_sin_duplicados[~data_sin_duplicados['Nro_transaccion'].duplicated(keep=False)].groupby(['NOMBRE_PROVEEDOR', 'Cod_sub_seccion'])['Nro_transaccion'].count().reset_index()
proveedores_cod_subseccion.columns = ['NOMBRE_PROVEEDOR', 'Cod_sub_seccion', 'Transacciones únicas']
proveedores_cod_subseccion = proveedores_cod_subseccion.merge(data_sin_duplicados.groupby(['NOMBRE_PROVEEDOR', 'Cod_sub_seccion'])['Total'].sum().reset_index(),on=['NOMBRE_PROVEEDOR', 'Cod_sub_seccion'])
proveedores_cod_subseccion.columns = ['NOMBRE_PROVEEDOR', 'Cod_sub_seccion', 'Transacciones únicas', 'Total']
# proveedores_subseccion = data.groupby(['NOMBRE_PROVEEDOR', 'Cod_sub_seccion']).agg({'Nro_transaccion': 'count', 'Total': 'sum'}).reset_index()
# proveedores_subseccion.columns = ['NOMBRE_PROVEEDOR', 'Cod_sub_seccion', 'Transacciones', 'Total']
proveedores_subseccion = data.groupby(['NOMBRE_PROVEEDOR', 'Cod_sub_seccion']).agg({'Nro_transaccion': 'nunique', 'Total': 'sum'}).reset_index()
proveedores_subseccion.columns = ['NOMBRE_PROVEEDOR', 'Cod_sub_seccion', 'Transacciones', 'Total']
clientes_cod_subseccion = data_sin_duplicados[~data_sin_duplicados['Nro_transaccion'].duplicated(keep=False)].groupby(['nombre_cliente', 'Cod_sub_seccion'])['Nro_transaccion'].count().reset_index()
clientes_cod_subseccion.columns = ['nombre_cliente', 'Cod_sub_seccion', 'Transacciones únicas']
clientes_cod_subseccion = clientes_cod_subseccion.merge(data_sin_duplicados.groupby(['nombre_cliente', 'Cod_sub_seccion'])['Total'].sum().reset_index(),on=['nombre_cliente', 'Cod_sub_seccion'])
clientes_cod_subseccion.columns = ['nombre_cliente', 'Cod_sub_seccion', 'Transacciones únicas', 'Total']
clientes_subseccion = data.groupby(['nombre_cliente', 'Cod_sub_seccion']).agg({'Nro_transaccion': 'nunique', 'Total': 'sum'}).reset_index()
clientes_subseccion.columns = ['nombre_cliente', 'Cod_sub_seccion', 'Transacciones', 'Total']
# Columnas que apareceran
secciones_sub_secc = data.groupby(['Cod_seccion', 'Cod_sub_seccion']).agg({'Total': 'sum', 'Nro_transaccion': 'nunique'}).reset_index()
secciones_sub_secc = secciones_sub_secc.drop_duplicates()
secciones_sub_secc = secciones_sub_secc.sort_values('Total', ascending=False)
# Agrupar los datos por cliente y calcular el total gastado por cada cliente
clientes_gastos = data.groupby(['nombre_cliente'])['Total'].sum().reset_index()
clientes_gastos = clientes_gastos.drop_duplicates()
# Eliminar las filas duplicadas para cada Nro Transaccion
data = data.drop_duplicates(subset=['Nro_transaccion', 'nombre_cliente'])
# Ordenar los clientes por monto gastado de mayor a menor
# sort_values() metodo que ordena de mayor a menor
clientes_gastos = clientes_gastos.sort_values('Total', ascending=False)
transacciones_por_cliente = data['nombre_cliente'].value_counts().reset_index()
transacciones_por_cliente = transacciones_por_cliente.drop_duplicates()
transacciones_por_cliente.columns = ['nombre_cliente', 'Nro_transaccion']
transacciones_por_cliente = pd.merge(transacciones_por_cliente, clientes_gastos[['nombre_cliente', 'Total']],on='nombre_cliente')
transacciones_por_cliente = transacciones_por_cliente.sort_values('Nro_transaccion', ascending=False)
# Calcular el monto promedio por transacción
monto_promedio = data.groupby('nombre_cliente')['Total'].mean().reset_index()
# Obtiene las transacciones únicas por sección
unique_transactions = []
for col in data_bin.columns:
    section_transactions = data_bin.loc[data_bin[col] == 1]
    other_sections = set(data_bin.columns) - {col}
    unique_mask = section_transactions[list(other_sections)].sum(axis=1) == 0
    unique_transactions.append(section_transactions[unique_mask])

# Concatena todas las transacciones únicas en un solo DataFrame
unique_transactions = pd.concat(unique_transactions)
# Crea una lista para almacenar los datos de la salida
output_data_unique = []
# Agrega los datos de cada sección a la lista de salida
for col in data_bin.columns:
    transactions_count = len(unique_transactions.loc[unique_transactions[col] == 1])
    output_data_unique.append([col, transactions_count])
# Crea un DataFrame con los datos de la salida
output_unique_df = pd.DataFrame(output_data_unique, columns=['SECCION', 'TRANSACCIONES ÚNICAS'])
#####################
# Guarda los resultados en un archivo Excel
with pd.ExcelWriter('resultado_Secciones.xlsx') as writer:
    frequent_itemsets.to_excel(writer, sheet_name='Frecuentes', index=False)
    rules.to_excel(writer, sheet_name='Reglas', index=False)
    pd.DataFrame(output_data, columns=['Antecedente', 'Consecuente', 'Número de transacciones']).to_excel(writer, sheet_name='Transacciones', index=False)
    # Concatenar los DataFrames en uno solo
    merged_df = pd.concat([output_unique_df, transactions_per_section], axis=1)
    # Restar las columnas y guardar el resultado en una nueva columna
    merged_df['NO UNICOS'] = merged_df['Nro_transaccion'] - merged_df['TRANSACCIONES ÚNICAS']
    # Calcular el porcentaje de las transacciones únicas
    merged_df['% UNICOS'] = merged_df['TRANSACCIONES ÚNICAS'] / merged_df['Nro_transaccion']
    # Formatear los valores como porcentaje
    merged_df['% UNICOS'] = merged_df['% UNICOS'].apply(lambda x: f"{x:.2%}")
    merged_df['% NO UNICOS'] = merged_df['NO UNICOS'] / merged_df['Nro_transaccion']
    # Formatear los valores como porcentaje
    merged_df['% NO UNICOS'] = merged_df['% NO UNICOS'].apply(lambda x: f"{x:.2%}")


    # Calcular el monto de la división solitos
    # merged_df['Monto Division Solitos'] = (merged_df['division solitos'].str.rstrip('%').astype(float) / 100) * \
    # merged_df['Total por sección']
    # merged_df['Monto Division Parejas'] = (merged_df['division pareja'].str.rstrip('%').astype(float) / 100) * \
    # merged_df['Total por sección']
    transactions_per_section.to_excel(writer, sheet_name='Unicas y Juntas', index=False)
    pd.DataFrame({'Total de transacciones': [transactions_total]}).to_excel(writer, sheet_name='Total', index=False)
    # Guardar el DataFrame combinado en una hoja de cálculo
    merged_df.to_excel(writer, sheet_name='Unicas y Juntas', index=False)
    # output_unique_df.to_excel(writer, sheet_name='Transacciones únicas', index=False)
    # transactions_per_section.to_excel(writer, sheet_name='Transacciones por sección', index=False)
    # pd.DataFrame({'Total de transacciones': [transactions_total]}).to_excel(writer, sheet_name='Total', index=False)
    # clientes_gastos.to_excel(writer,sheet_name='carritos',index=False)
    secciones_sub_secc.to_excel(writer, sheet_name='Seccion y Sub', index=False)
    transacciones_por_proveedor = pd.merge(transacciones_por_proveedor, resultado, on='NOMBRE_PROVEEDOR', how='left')
    transacciones_por_proveedor.to_excel(writer, sheet_name='Transacciones total proveedor', index=False)
    # Guardar en una hoja aparte llamada "Proveedores únicos"
    # proveedores_unicos.to_excel(writer, sheet_name='Proveedores trans solitos', index=False)
    # proveedores_cod_subseccion.to_excel(writer, index=False, sheet_name='Unicos con sub')
    # proveedores_subseccion.to_excel(writer, index=False, sheet_name='prov conjuntos ')
    transacciones_por_cliente.to_excel(writer, sheet_name='Transacciones_clientes', index=False)
    clientes_subseccion.to_excel(writer, sheet_name='Total_seccion', index=False)
    clientes_cod_subseccion.to_excel(writer, sheet_name='Unicos_seccion', index=False)
    # monto_promedio.to_excel(writer, sheet_name='Monto_Promedio', index=False)
    # Crear un DataFrame con los resultados
    # results_df = pd.DataFrame({'Total de transacciones': [total_transactions], 'Transacciones únicas': [unique_transactions_count],'Transacciones con otras secciones': [other_transactions_count],'Resultado': [result]})
    # # Guardar el DataFrame en la hoja "Total"
    # results_df.to_excel(writer, sheet_name='Total', index=False)