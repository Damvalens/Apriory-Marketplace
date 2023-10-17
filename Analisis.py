# VPR
# Interfaz para vpr y consulta de datos traveler de PYTHON
import pandas as pd
from mlxtend.preprocessing import TransactionEncoder
from mlxtend.frequent_patterns import apriori, association_rules
from sqlalchemy import create_engine
import tkinter as tk
from tkinter import filedialog, messagebox

global registro_button, vpr_button, proveedor_entry, proveedor_entry_2, start_date_entry,start_date_entry_2,end_date_entry_2,end_date_entry, query_button, query_button2,query_button3, back_button, labels

# Función para consultar los datos y guardarlos en un archivo Excel
def query_and_save_data():
    # Obtener el número de proveedor ingresado por el usuario
    proveedor = proveedor_entry.get()

    # Establecer la cadena de conexión a la base de datos
    server = '10.51.80.111,1433'
    database = 'pegasus'
    username = 'sa'
    password = 'Market2023'
    connection_string = f"mssql+pyodbc://{username}:{password}@{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server"

    # Crear el motor de SQLAlchemy
    engine = create_engine(connection_string)

    # Establecer la consulta SQL que deseas ejecutar
    query = f"""
    -- La consulta SQL aquí
    SELECT
    CD.NRO_REG,
    p.codigo,
    pc.codigo_alternativo AS BARRAS,
    p.DESCRIPCION_PRODUCTO,
    CONVERT(VARCHAR(10), GETDATE(), 103) AS FECHA_ACTUAL,
    CONVERT(VARCHAR(10), ct.FECHA, 103) AS FECHA_COMPRA,
    DATEDIFF(DAY, ct.FECHA, GETDATE()) AS DIAS,
    CAST(ROUND(CD.CANTIDAD, 0) AS float) AS Q_COMPRADA,
    CAST(ROUND(SUM(MD.CANTIDAD), 0) AS float) AS STOCK_ACTUAL,
    PX.NOMBRE_PROVEEDOR,
    s.Secciones,
    SB.Sub_secciones,
    convert(int, p.PRECIO_COSTO) * cast(ROUND(SUM(MD.CANTIDAD), 0) AS float) AS STOCK_TOTAL_HOY,
	convert(int, p.PRECIO_COSTO) as COSTO_UNIT
	,CONVERT(bigint, (Pv.PRECIO)/1.1) as PRECIO_LISTA
	,convert(bigint,(((Pv.PRECIO)/1.1) - (p.PRECIO_COSTO))) as UTILIDAD
	,CAST((CONVERT(decimal(20, 2), (((Pv.PRECIO)/1.1) - (p.PRECIO_COSTO))*1.1) / NULLIF(Pv.PRECIO, 0)) AS decimal(10, 2)) as MDR -- CALCULO MDR
FROM
    productos p
    INNER JOIN productos_codigos pc ON pc.codigo = p.codigo
    INNER JOIN Seccion s ON s.Cod_seccion = p.Cod_seccion
    INNER JOIN Sub_seccion SB ON SB.Cod_sub_seccion = p.Cod_sub_seccion
    INNER JOIN MARCAS M ON M.COD_MARCA = p.COD_MARCA
    INNER JOIN PRECIOS_DET PV ON pv.CODIGO = p.codigo
    INNER JOIN PROVEEDORES PX ON PX.COD_PROVEEDOR = p.cod_proveedor
    INNER JOIN GRUPO G ON G.COD_GRUPO = p.COD_GRUPO
    INNER JOIN CATEGORIAS C ON C.COD_CATEGORIA = p.COD_CATEGORIA
    INNER JOIN MOVIMIENTOS_DEPOSITOS MD ON MD.CODIGO = p.CODIGO
    INNER JOIN DEPOSITOS D ON D.COD_DEPOSITO = MD.COD_DEPOSITO
    INNER JOIN COMPRAS_DET CD ON CD.CODIGO = p.CODIGO
    INNER JOIN compras ct ON CT.NRO_REG = CD.NRO_REG
WHERE
    pv.NRO_LISTA = 1 and CD.CANTIDAD>0 
    {'AND px.COD_PROVEEDOR IN (' + proveedor + ')' if proveedor else ''}
GROUP BY
    CD.NRO_REG,
    p.codigo,
    pc.codigo_alternativo,
    p.DESCRIPCION_PRODUCTO,
    s.Secciones,
    SB.Sub_secciones,
    G.DESCRIPCION_GRUPO,
    C.DESCRIPCION_CATE,
    M.MARCA,
    PX.NOMBRE_PROVEEDOR,
    p.PRECIO_COSTO,
    pv.PRECIO,
    CD.CANTIDAD,
    ct.FECHA,
	p.PRECIO_COSTO,
	pv.PRECIO,
    DATEDIFF(DAY, ct.FECHA, GETDATE());
"""

    # Ejecutar la consulta y obtener los resultados en un DataFrame
    results = pd.read_sql(query, engine)

    # Cerrar la conexión después de obtener los resultados
    engine.dispose()

    # Verificar si se obtuvieron datos antes de guardar en el archivo Excel o csv
    if not results.empty:
        # Abrir el cuadro de diálogo para seleccionar el tipo de archivo
        filetypes = [("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=filetypes)
        if file_path:
            # Obtener la extensión del archivo seleccionado para determinar el formato de salida
            file_extension = file_path.split(".")[-1].lower()
            # Guardar los resultados en el archivo Excel o CSV según la extensión seleccionada
            if file_extension == "xlsx":
                results.to_excel(file_path, index=False)
            elif file_extension == "csv":
                results.to_csv(file_path, index=False)
            messagebox.showinfo("Guardar", f"Los resultados se han guardado exitosamente en el archivo.")

def create_and_run_interface():
    global proveedor_entry, registro_button, query_button, back_button, labels

    # Desactivar el botón "VPR" para evitar abrir múltiples interfaces
    registro_button.config(state=tk.DISABLED)
    # ocultar el botón "VPR"
    vpr_button.grid_remove()
    #ocultar el boton carrito
    carrito_button.grid_remove()
    # Crear cuadros de texto para el rango el número de proveedor
    proveedor_entry = tk.Entry(root)

    # Crear un botón para consultar los datos y guardarlos en un archivo Excel
    query_button = tk.Button(root, text="Consultar y Guardar", command=query_and_save_data)

    # Crear un botón "Atrás" para volver a la vista del botón "VPR"
    back_button = tk.Button(root, text="Atrás", command=show_vpr_button)

    # Almacenar la etiqueta en una lista para poder ocultarla más tarde
    labels = [tk.Label(root, text="Número de proveedor:")]

    # Posicionar los widgets en la ventana usando el gestor de geometría "grid"
    labels[0].grid(row=0, column=0)
    proveedor_entry.grid(row=0, column=1)
    query_button.grid(row=3, column=0, columnspan=2)
    back_button.grid(row=4, column=0, columnspan=2)
    # reactivar el botón "VPR"
    vpr_button.config(state=tk.NORMAL)
    # Ocultar el botón "registro" después de presionarlo
    registro_button.grid_remove()


# Función para mostrar nuevamente el botón "VPR" y ocultar los demás widgets
def show_vpr_button():
    global registro_button, query_button, back_button, labels
    #mostrar el boton carrito
    carrito_button.grid(row=0, column=2, padx=10, pady=10)
    # Mostrar el botón "registro"
    registro_button.grid(row=0, column=0, padx=10, pady=10)
    # mostrar el boton "VPR"
    vpr_button.grid(row=0, column=1, padx=10, pady=10)
    # Ocultar los demás widgets
    for label in labels:
        label.grid_remove()
    proveedor_entry.grid_remove()
    query_button.grid_remove()
    back_button.grid_remove()
    # Reactivar el botón "registro"
    registro_button.config(state=tk.NORMAL)


def query_and_save_data_2():
    global start_date_entry, end_date_entry, proveedor_entry_2
    # Obtener el número de proveedor ingresado por el usuario
    proveedor = proveedor_entry_2.get()
    # Obtener las fechas de inicio y finalización
    start_date = start_date_entry.get()
    end_date = end_date_entry.get()

    # Establecer la cadena de conexión a la base de datos
    server = '10.51.80.111,1433'
    database = 'pegasus'
    username = 'sa'
    password = 'Market2023'
    connection_string = f"mssql+pyodbc://{username}:{password}@{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server"

    # Crear el motor de SQLAlchemy
    engine = create_engine(connection_string)

    # Establecer la consulta SQL que deseas ejecutar
    query = f"""
    -- La consulta SQL aquíuse pegasus;
select a.nro_reg,
convert(date,a.fecha) as fecha,
a.tipo_documen,
case 
when a.TIPO_DOCUMEN = 223 then 'Cajas Asuncion'
when a.TIPO_DOCUMEN = 225 then 'Cajas San Bernardino'
when a.TIPO_DOCUMEN = 226 then 'Fac San Bernardino'
when a.TIPO_DOCUMEN in (449,451) then 'Ashley Credito'
when a.TIPO_DOCUMEN in (450,452) then 'Ashley Contado'
else 'Atencion al Cliente'
end as punto_venta,
b.codigo CODIGO_PROD,
(select top 1 codigo_alternativo from productos_codigos pc where pc.codigo=b.codigo) as cod_barra,
c.DESCRIPCION_PRODUCTO,
G.Secciones SECCION, 
H.Sub_secciones SUB_SECCIONES,
d.DESCRIPCION_CATE CATEGORIA,
f.DESCRIPCION_SUB_CATE,
E.DESCRIPCION_GRUPO,
I.MARCA,
k.COD_PROVEEDOR,
K.NOMBRE_PROVEEDOR,

  CASE
            WHEN pc.cod_condicion = 1 AND pc.tipo_doc = 20 THEN 'IMPORTACION'
            WHEN pc.cod_condicion = 2 AND pc.tipo_doc = 20 THEN 'IMPORTACION'
            WHEN pc.cod_condicion = 1 THEN 'CONTADO'
            WHEN pc.cod_condicion = 2 THEN 'CREDITO'
            WHEN pc.cod_condicion = 3 THEN 'CONSIGNACION'
            ELSE 'OTROS'
       END AS condicion,cl.NOMBRE_CLIENTE,a.observacion,cc.categoria_cliente,
/*
precio de venta
precio final trae cuando es producto con promo en linea de cajas
si no es precio con promo sumal el total y el importe
*/
convert (bigint,sum (b.unidades)) CANTIDAD,
convert(bigint,b.PRECIO_NETO) as pventa_und,
(
select replace(convert(decimal(10,2),coti.cotizacion),'.',',') as cotizacion from cotizaciones coti where convert(date, coti.fecha) = convert(date, a.fecha) 
and coti.cod_moneda = 2 and coti.cod_moneda_destino = 1 and mul_div = 0

) as cotizacion_del_dia,

case
when a.TIPO_DOCUMEN in (223,225,13,19,18,225,226,217,218,449,450,451,452,453,454,455,456,457,458,462,463,464,465,466,467,468,469,470,471,472,473,474,475,476) then
        case 
			when a.COD_MONEDA = 2  then
			CONVERT (
			BIGINT,SUM (b.TOT_PRECIO+b.TOT_IMP) * 
			(select coti.cotizacion from cotizaciones coti where convert(date, coti.fecha) = convert(date, a.fecha) 
			and coti.cod_moneda = 2 and coti.cod_moneda_destino = 1 and mul_div = 0))
		else
			CONVERT (BIGINT,SUM (b.TOT_PRECIO+b.TOT_IMP)/1.1,0) -- VENTA SIN IVA
		end
END AS TOTAL_VENTA,
/* CALCULO PARA COSTO DE PRODUCTOS SI ES DOLARES O GS*/
case
	when a.TIPO_DOCUMEN in (223,225,13,19,18,225,226,217,218,449,450,451,452,453,454,455,456,457,458,462,463,464,465,466,467,468,469,470,471,472,473,474,475,476) then
		case 
			when a.COD_MONEDA = 2  then
				/*costo calculado de la ficha de productos*/
				CONVERT(BIGINT,(b.TOT_COSTO)*
				(select coti.cotizacion from cotizaciones coti where convert(date, coti.fecha) = convert(date, a.fecha) 
				and coti.cod_moneda = 2 and coti.cod_moneda_destino = 1 and mul_div = 0))
			else
				CONVERT(BIGINT,(b.TOT_COSTO))
			end				
END AS COSTO_VENTA, -- COSTO SIN IVA
/*CALCULO PARA UTILIDAD SI ES EN DOLARES O GS*/
case
	when a.TIPO_DOCUMEN in (223,225,13,19,18,225,226,217,218,449,450,451,452,453,454,455,456,457,458,462,463,464,465,466,467,468,469,470,471,472,473,474,475,476) then
		case 
			when a.COD_MONEDA = 2  then
				convert(bigint, SUM ((b.TOT_PRECIO+b.TOT_IMP) - (b.TOT_COSTO*1.1)) *
				(select coti.cotizacion from cotizaciones coti where convert(date, coti.fecha) = convert(date, a.fecha) 
				and coti.cod_moneda = 2 and coti.cod_moneda_destino = 1 and mul_div = 0))
			ELSE
				convert(bigint, SUM ((b.TOT_PRECIO+b.TOT_IMP) - (b.TOT_COSTO*1.1))/1.1,0)--SIN IVA
			END
END as UTILIDAD,
/* calculo de MD */
replace(convert(decimal(20,2),(SUM(b.TOT_PRECIO+b.TOT_IMP) - (b.TOT_COSTO*1.1)) / nullif(SUM (b.TOT_PRECIO+b.TOT_IMP-b.DESC_VEN),0)),'.',',')
as MDR,
/* calculo de MU */
replace(convert(decimal(20,2),(SUM(b.TOT_PRECIO+b.TOT_IMP-b.DESC_VEN) - (sum(b.unidades) * c.PRECIO_COSTO*1.1)) / nullif(sum(b.unidades) * (c.PRECIO_COSTO*1.1),0)),'.',',')
as MU
from dbo.ventas a 
inner join DBO.ventas_det b on (a.NRO_REG = b.NRO_REG)
LEFT JOIN CLIENTES cl on (a.COD_CLIENTE=cl.COD_CLIENTE)
INNER JOIN CLIENTES_CATE cc ON (cl.COD_CATEGORIA = cc.COD_CATEGORIA)
LEFT join DBO.productos c on (b.CODIGO = c.codigo)
LEFT join DBO.CATEGORIAS d on (c.COD_CATEGORIA = d.COD_CATEGORIA)
LEFT JOIN DBO.GRUPO E ON (C.COD_GRUPO = E.COD_GRUPO) 
LEFT JOIN DBO.SUB_CATEGORIAS F ON (c.COD_SUB_CATEGORIA = f.COD_SUB_CATEGORIA)
LEFT join DBO.Seccion g on (c.Cod_seccion = G.Cod_seccion)
LEFT JOIN DBO.Sub_seccion H ON (C.Cod_sub_seccion = H.Cod_sub_seccion)
LEFT JOIN DBO.MARCAS I ON (C.COD_MARCA = I.COD_MARCA)
LEFT JOIN DBO.PROVEEDORES K ON (B.cod_proveedor_ven = K.COD_PROVEEDOR)
LEFT JOIN proveedor_condicion pc ON pc.cod_proveedor=K.COD_PROVEEDOR

--cambiar el rango de fecha
WHERE convert(date,a.fecha) BETWEEN '{start_date}' AND '{end_date}'
-- Si se ingresaron números de proveedores, filtrar por ellos
{'AND k.COD_PROVEEDOR IN (' + proveedor + ')' if proveedor else ''}
--and a.TIPO_DOCUMEN in (449,450,451,452) 



group by pc.cod_condicion,pc.tipo_doc,a.NRO_REG,CONVERT(date,a.fecha),SUBSTRING (convert (VARCHAR (50), A.FECHA,112),1,6),b.codigo,c.codigo_barras,c.DESCRIPCION_PRODUCTO,d.DESCRIPCION_CATE,
E.DESCRIPCION_GRUPO,f.DESCRIPCION_SUB_CATE,G.Secciones,H.Sub_secciones,I.MARCA, b.PRECIO_NETO, b.TOT_COSTO, a.COD_MONEDA
,K.NOMBRE_PROVEEDOR,a.observacion,cc.CATEGORIA_CLIENTE, c.PRECIO_COSTO,cl.NOMBRE_CLIENTE, k.COD_PROVEEDOR, B.precio_final,a.TIPO_DOCUMEN,a.nro_reg, a.cotizacion
order by CONVERT(date,a.fecha);

    """

    # Ejecutar la consulta y obtener los resultados en un DataFrame
    results = pd.read_sql(query, engine)

    # Cerrar la conexión después de obtener los resultados
    engine.dispose()

    # Abrir el cuadro de diálogo para seleccionar el tipo de archivo
    filetypes = [("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=filetypes)
    if file_path:
        # Obtener la extensión del archivo seleccionado para determinar el formato de salida
        file_extension = file_path.split(".")[-1].lower()
        # Guardar los resultados en el archivo Excel o CSV según la extensión seleccionada
        if file_extension == "xlsx":
            results.to_excel(file_path, index=False)
        elif file_extension == "csv":
            results.to_csv(file_path, index=False)
        messagebox.showinfo("Guardar", f"Los resultados se han guardado exitosamente en el archivo.")
    # else:
    #     messagebox.showwarning("Advertencia", "No se encontraron resultados para el rango de fechas especificado.")


def create_and_run_interface_2():
    global start_date_entry, end_date_entry, proveedor_entry_2, vpr_button, query_button2, back_button, labels

    # Desactivar el botón "VPR" para evitar abrir múltiples interfaces
    vpr_button.config(state=tk.DISABLED)
    # ocultar el boton "registro"
    registro_button.grid_remove()
    # ocultar el boton "carrito"
    carrito_button.grid_remove()

    # Crear cuadros de texto para el rango de fechas y el número de proveedor
    start_date_entry = tk.Entry(root)
    end_date_entry = tk.Entry(root)
    proveedor_entry_2 = tk.Entry(root)

    # Crear un botón para consultar los datos y guardarlos en un archivo Excel
    query_button2 = tk.Button(root, text="Consultar y Guardar", command=query_and_save_data_2)

    # Crear un botón "Atrás" para volver a la vista del botón "VPR"
    back_button = tk.Button(root, text="Atrás", command=show_vpr_button_2)

    # Almacenar las etiquetas en una lista para poder ocultarlas más tarde
    labels = [tk.Label(root, text="Número de proveedor:"), tk.Label(root, text="Fecha de inicio (YYYY-MM-DD):"),
              tk.Label(root, text="Fecha de finalización (YYYY-MM-DD):")]

    # Posicionar los widgets en la ventana usando el gestor de geometría "grid"
    labels[0].grid(row=0, column=0)
    proveedor_entry_2.grid(row=0, column=1)
    labels[1].grid(row=1, column=0)
    start_date_entry.grid(row=1, column=1)
    labels[2].grid(row=2, column=0)
    end_date_entry.grid(row=2, column=1)
    query_button2.grid(row=3, column=0, columnspan=2)
    back_button.grid(row=4, column=0, columnspan=2)

    # Ocultar el botón "VPR" después de presionarlo
    vpr_button.grid_remove()
    # reactivar el botón REGISTRO
    registro_button.config(state=tk.NORMAL)
    # REACTIVAR EL BOTON CARRITO
    carrito_button.config(state=tk.NORMAL)


# Función para mostrar nuevamente el botón "VPR" y ocultar los demás widgets
def show_vpr_button_2():
    global vpr_button, query_button2, back_button, labels
    #mostar el boton "carrito"
    carrito_button.grid(row=0, column=2, padx=10, pady=10)
    # Mostrar el botón "VPR"
    vpr_button.grid(row=0, column=1, padx=10, pady=10)
    # mostrar el botón "registro"
    registro_button.grid(row=0, column=0, padx=10, pady=10)
    # Ocultar los demás widgets
    for label in labels:
        label.grid_remove()
    start_date_entry.grid_remove()
    end_date_entry.grid_remove()
    proveedor_entry_2.grid_remove()
    query_button2.grid_remove()
    back_button.grid_remove()

    # Reactivar el botón "VPR"
    vpr_button.config(state=tk.NORMAL)
    # REACTIVAR EL BOTON REGISTRO
    registro_button.config(state=tk.NORMAL)

def query_and_save_data_3():
    # Obtener el número fecha de inicio y final
    start_date_2 = start_date_entry_2.get()
    end_date_2 = end_date_entry_2.get()
    # Establecer la cadena de conexión a la base de datos
    server = '10.51.80.111,1433'
    database = 'pos_central'
    username = 'sa'
    password = 'Market2023'
    connection_string = f"mssql+pyodbc://{username}:{password}@{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server"
    # Crear el motor de base de datos
    engine = create_engine(connection_string)
    # Crear la consulta SQL
    query = f"""select  CONCAT(
    RIGHT(CAST(vd.zeta AS VARCHAR(50)), 50), '00',
    RIGHT(CAST(vp.nro_caja AS VARCHAR(50)), 50), '00',
    RIGHT(CAST(vp.nro_ticket AS VARCHAR(50)), 50),'00') AS Nro_transaccion,
vd.zeta, vp.nro_caja, vp.nro_ticket, vd.codigo, convert(bigint, vd.unidades) as cantidad, p.descripcion_producto,
convert(bigint,vd.precio_final) as precio, CAST(vd.precio_final * CONVERT(BIGINT, vd.unidades)AS INT) AS Total, vp.documento, vp.nombre_cliente, pr.NOMBRE_PROVEEDOR, s.Cod_seccion, ss.Cod_sub_seccion,s.Secciones,ss.Sub_secciones,
g.DESCRIPCION_GRUPO,  d.DESCRIPCION_CATE, f.DESCRIPCION_SUB_CATE
from ventas_pos vp
join ventas_det_pos vd on vp.zeta = vd.zeta and vp.nro_caja = vd.nro_caja and vp.nro_ticket = vd.nro_ticket
join pegasus.DBO.productos p on vd.codigo = p.codigo
LEFT join pegasus.DBO.Seccion S on (p.Cod_seccion = s.Cod_seccion)
LEFT JOIN pegasus.DBO.Sub_seccion ss ON (p.Cod_sub_seccion = ss.Cod_sub_seccion)
LEFT JOIN pegasus.DBO.GRUPO g ON (p.COD_GRUPO = g.COD_GRUPO)
LEFT join pegasus.DBO.CATEGORIAS d on (p.COD_CATEGORIA = d.COD_CATEGORIA)
LEFT JOIN pegasus.DBO.SUB_CATEGORIAS F ON (p.COD_SUB_CATEGORIA = f.COD_SUB_CATEGORIA)
LEFT JOIN pegasus.DBO.proveedores pr ON (p.cod_proveedor = pr.cod_proveedor)
WHERE convert(date,vp.fecha) BETWEEN '{start_date_2}' AND '{end_date_2}'

group by vp.nro_caja, vp.nro_ticket, vd.codigo, vd.unidades, p.descripcion_producto,
vd.precio_final,  vp.documento, vp.nombre_cliente, pr.NOMBRE_PROVEEDOR,s.Cod_seccion, ss.Cod_sub_seccion,
g.DESCRIPCION_GRUPO,  d.DESCRIPCION_CATE, f.DESCRIPCION_SUB_CATE,vd.zeta,s.Secciones,ss.Sub_secciones"""

    # Ejecutar la consulta y obtener los resultados en un DataFrame
    results = pd.read_sql(query, engine)

    # Cerrar la conexión después de obtener los resultados
    engine.dispose()
    # Abrir el cuadro de diálogo para seleccionar el tipo de archivo
    filetypes = [("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=filetypes)
    if file_path:
        # Obtener la extensión del archivo seleccionado para determinar el formato de salida
        file_extension = file_path.split(".")[-1].lower()
        # Guardar los resultados en el archivo Excel o CSV según la extensión seleccionada
        if file_extension == "xlsx":
            results.to_excel(file_path, index=False)
        elif file_extension == "csv":
            results.to_csv(file_path, index=False)
        messagebox.showinfo("Guardar", f"Los resultados se han guardado exitosamente en el archivo.")
    # else:
    #     messagebox.showwarning("Advertencia", "No se encontraron resultados para el rango de fechas especificado.")
# algoritmo apriori
    # Ejecutar la consulta y cargar los resultados en un DataFrame de Pandas
    df = pd.read_sql(query, engine)
    # Carga el archivo guardado de la consulta SQL ejecutada anteriormente

    data = pd.read_excel(file_path)
    # Convierte los datos a una lista de transacciones
    transactions = []
    for trans_id, group in data.groupby('Nro_transaccion'):
        transactions.append(list(group['Secciones'].values))
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
    transactions_per_section = data.groupby('Secciones')['Nro_transaccion'].nunique().reset_index()
    transactions_total = data['Nro_transaccion'].nunique()
    ####################
    # Calcular el monto total por Cod_seccion
    total_per_section = data.groupby('Secciones')['Total'].sum().reset_index()
    # Agregar la columna "Total por sección" al DataFrame "transactions_per_section"
    transactions_per_section['TOTAL VENTA (Gs)'] = total_per_section['Total']
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
        pd.DataFrame(output_data, columns=['Antecedente', 'Consecuente', 'Número de transacciones']).to_excel(writer,sheet_name='Transacciones',index=False)
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
            output_unique_df = pd.DataFrame(output_data_unique,columns=['Sección', 'Transacciones con otras secciones'])
        # Guarda los resultados en un archivo Excel
        with pd.ExcelWriter('resultado_Secciones.xlsx') as writer:
            frequent_itemsets.to_excel(writer, sheet_name='Frecuentes', index=False)
            rules.to_excel(writer, sheet_name='Reglas', index=False)
            pd.DataFrame(output_data, columns=['Antecedente', 'Consecuente', 'Número de transacciones']).to_excel(
                writer, sheet_name='Transacciones', index=False)
            # Crea una lista para almacenar los datos de las transacciones únicas
            unique_transactions = []
            # Crea una lista para almacenar los datos de las transacciones con otras secciones
            other_transactions = []
            # Separa las transacciones únicas y las transacciones con otras secciones
            for trans_id, group in data.groupby('Nro_transaccion'):
                if len(set(group['Secciones'])) == 1:
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
        data_sin_duplicados[~data_sin_duplicados['Nro_transaccion'].duplicated(keep=False)].groupby('NOMBRE_PROVEEDOR')[
            'Nro_transaccion'].count().reset_index()
    transacciones_unicas.columns = ['NOMBRE_PROVEEDOR', 'Transacciones únicas']
    # Fusionar los resultados en un solo DataFrame
    resultado = transacciones_unicas.merge(transacciones_total, on='NOMBRE_PROVEEDOR', how='left')
    # Calcular el número de transacciones no únicas (conjuntos)
    resultado['conjuntos'] = resultado['Nro_transaccion'] - resultado['Transacciones únicas']
    # Calcular el porcentaje de transacciones únicas
    resultado['% UNICOS'] = resultado['Transacciones únicas'] / resultado['Nro_transaccion']
    resultado['% NO UNICOS'] = resultado['conjuntos'] / resultado['Nro_transaccion']
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
        data_sin_duplicados[~data_sin_duplicados['Nro_transaccion'].duplicated(keep=False)].groupby(
            ['NOMBRE_PROVEEDOR', 'Sub_secciones'])['Nro_transaccion'].count().reset_index()
    proveedores_cod_subseccion.columns = ['NOMBRE_PROVEEDOR', 'Sub_secciones', 'Transacciones únicas']
    proveedores_cod_subseccion = proveedores_cod_subseccion.merge(
        data_sin_duplicados.groupby(['NOMBRE_PROVEEDOR', 'Sub_secciones'])['Total'].sum().reset_index(),on=['NOMBRE_PROVEEDOR', 'Sub_secciones'])
    proveedores_cod_subseccion.columns = ['NOMBRE_PROVEEDOR', 'Sub_secciones', 'Transacciones únicas', 'Total']
    # proveedores_subseccion = data.groupby(['NOMBRE_PROVEEDOR', 'Cod_sub_seccion']).agg({'Nro_transaccion': 'count', 'Total': 'sum'}).reset_index()
    # proveedores_subseccion.columns = ['NOMBRE_PROVEEDOR', 'Cod_sub_seccion', 'Transacciones', 'Total']
    proveedores_subseccion = data.groupby(['NOMBRE_PROVEEDOR', 'Sub_secciones']).agg({'Nro_transaccion': 'nunique', 'Total': 'sum'}).reset_index()
    proveedores_subseccion.columns = ['NOMBRE_PROVEEDOR', 'Sub_secciones', 'Transacciones', 'Total']
    clientes_cod_subseccion = \
    data_sin_duplicados[~data_sin_duplicados['Nro_transaccion'].duplicated(keep=False)].groupby(['nombre_cliente', 'Sub_secciones'])['Nro_transaccion'].count().reset_index()
    clientes_cod_subseccion.columns = ['nombre_cliente', 'Sub_secciones', 'Transacciones únicas']
    clientes_cod_subseccion = clientes_cod_subseccion.merge(data_sin_duplicados.groupby(['nombre_cliente', 'Sub_secciones'])['Total'].sum().reset_index(),on=['nombre_cliente', 'Sub_secciones'])
    clientes_cod_subseccion.columns = ['nombre_cliente', 'Sub_secciones', 'Transacciones únicas', 'Total']
    clientes_subseccion = data.groupby(['nombre_cliente', 'Sub_secciones']).agg({'Nro_transaccion': 'nunique', 'Total': 'sum'}).reset_index()
    clientes_subseccion.columns = ['nombre_cliente', 'Sub_secciones', 'Transacciones', 'Total']
    # Columnas que apareceran
    secciones_sub_secc = data.groupby(['Secciones', 'Sub_secciones']).agg({'Total': 'sum', 'Nro_transaccion': 'nunique'}).reset_index()
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
    with pd.ExcelWriter('carrito.xlsx') as writer:
        frequent_itemsets.to_excel(writer, sheet_name='Frecuentes', index=False)
        rules.to_excel(writer, sheet_name='Reglas', index=False)
        pd.DataFrame(output_data, columns=['Antecedente', 'Consecuente', 'Número de transacciones']).to_excel(writer, sheet_name='Transacciones',index=False)
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
        transactions_per_section.to_excel(writer, sheet_name='Unicas y Juntas', index=False)
        pd.DataFrame({'Total de transacciones': [transactions_total]}).to_excel(writer, sheet_name='Total', index=False)
        # Guardar el DataFrame combinado en una hoja de cálculo
        merged_df.to_excel(writer, sheet_name='Unicas y Juntas', index=False)
        secciones_sub_secc.to_excel(writer, sheet_name='Seccion y Sub', index=False)
        transacciones_por_proveedor = pd.merge(transacciones_por_proveedor, resultado, on='NOMBRE_PROVEEDOR',how='left')
        transacciones_por_proveedor.to_excel(writer, sheet_name='Transacciones total proveedor', index=False)
        transacciones_por_cliente.to_excel(writer, sheet_name='Transacciones_clientes', index=False)
        clientes_subseccion.to_excel(writer, sheet_name='Total_seccion', index=False)
        clientes_cod_subseccion.to_excel(writer, sheet_name='Unicos_seccion', index=False)

#########################



def create_and_run_interface_3():
    global start_date_entry_2, end_date_entry_2, query_button3, back_button, labels, carrito_button

    # Desactivar el botón "VPR"
    vpr_button.config(state=tk.DISABLED)
    # Desactivar el botón "registro"
    registro_button.config(state=tk.DISABLED)
    # Desactivar el botón "carrito"
    carrito_button.config(state=tk.DISABLED)

    # Crear las entradas de texto para la fecha de inicio y final
    start_date_entry_2 = tk.Entry(root)
    end_date_entry_2 = tk.Entry(root)

    # Crear el botón para consultar los datos
    query_button3 = tk.Button(root, text="Generar", command=query_and_save_data_3)

    # Crear el botón para volver a la pantalla anterior
    back_button = tk.Button(root, text="Volver", command=show_vpr_button_3)

    # Almacenar las etiquetas en una lista para poder ocultarlas más tarde
    labels = [tk.Label(root, text="Fecha de inicio (YYYY-MM-DD):"),
              tk.Label(root, text="Fecha de finalización (YYYY-MM-DD):")]

    # Posicionar los widgets en la ventana usando el gestor de geometría grid
    labels[0].grid(row=0, column=0)
    start_date_entry_2.grid(row=0, column=1)
    labels[1].grid(row=1, column=0)
    end_date_entry_2.grid(row=1, column=1)
    query_button3.grid(row=3, column=1,columnspan=2)
    back_button.grid(row=4, column=1,columnspan=2)

    # Ocultar el botón "VPR" después de presionarlo
    vpr_button.grid_remove()
    # Ocultar el botón "registro" después de presionarlo
    registro_button.grid_remove()
    # Ocultar el botón "carrito" después de presionarlo
    carrito_button.grid_remove()


def show_vpr_button_3():
    global vpr_button, registro_button, labels, carrito_button

    # Mostrar el botón "VPR"
    vpr_button.grid(row=0, column=1, padx=10, pady=10)

    # Mostrar el botón "carrito"
    carrito_button.grid(row=0, column=2, padx=10, pady=10)

    # Mostrar el botón "registro"
    registro_button.grid(row=0, column=0, padx=10, pady=10)

    # Ocultar el botón "query"
    query_button3.grid_remove()

    # Ocultar el botón "back"
    back_button.grid_remove()

    # Ocultar los labels
    for label in labels:
        label.grid_remove()

    # Ocultar las entradas de texto
    start_date_entry_2.grid_remove()
    end_date_entry_2.grid_remove()

    # Reactivar el botón "REGISTRO"
    registro_button.config(state=tk.NORMAL)
    # Reactivar el botón "carrito"
    carrito_button.config(state=tk.NORMAL)
    # Reactivar el botón "VPR"
    vpr_button.config(state=tk.NORMAL)

# Ejecutar el programa principal
if __name__ == "__main__":
    # Crear una ventana principal
    root = tk.Tk()
    root.title("MKP Gestion de Datos - Marca Registrada- VERSION: ALPHA")
    root.geometry("500x150")
    root.resizable(False, False)
    # CREAR EL ICONO DE LA APLICACION
    root.iconbitmap("PO.jpeg")

    # Crear el botón "registro" para abrir la interfaz gráfica
    registro_button = tk.Button(root, text="Registros", command=create_and_run_interface)
    registro_button.grid(row=0, column=0, padx=10, pady=10)  # Posicionar en la esquina superior izquierda

    # Crear el botón "VPR"
    vpr_button = tk.Button(root, text="VPR", command=create_and_run_interface_2)
    vpr_button.grid(row=0, column=1, padx=10, pady=10)  # Posicionar a lado del botón "Registro"

    # Crear un botón para el carrito de compras
    carrito_button = tk.Button(root, text="CARRITO", command=create_and_run_interface_3)
    carrito_button.grid(row=0, column=2, padx=10, pady=10)

    # Iniciar el bucle de eventos principal de la interfaz gráfica
    root.mainloop()