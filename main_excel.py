import pandas as pd
import sys
import matplotlib.pyplot as plt
import seaborn as sns
import xlsxwriter 
from openpyxl.chart import BarChart, Reference



def salida_controlada() :

    print('\nFinalizando programa\nmaven_pizzas ')
    print('\nPrograma finalizado')
    sys.exit()


def extract():

    ordersdetails = pd.read_csv('order_details_2016.csv', sep=';')
    ingredients = pd.read_csv('pizza_types.csv', encoding='latin1')
    pizza_price = pd.read_csv('pizzas.csv')
    orders = pd.read_csv('orders_2016.csv', sep=';')
    return [ingredients, orders, pizza_price, ordersdetails]


def tablas(fichero, fichero_2, details):
    # tabla de los ingredientes
    fig = plt.figure(figsize=(25,7))
    sns.barplot(x='Ingredientes_necesarios', y='Cantidad a comprar por semana', data=fichero,
                palette=sns.color_palette("Blues_d", len(fichero)))
    plt.xlabel('Ingredientes')
    plt.ylabel('Cantidad')
    plt.xticks(rotation=45)
    plt.title("Ingredientes mas usados")
    plt.savefig("Ingredientes_semanales.jpg")

    # tabla de las pizzas y sus precios respectivos.
    fig = plt.figure(figsize=(30, 30))
    sns.barplot(x='pizza_id', y='price', data=fichero_2, palette=sns.color_palette("Blues_d", len(fichero_2)))
    plt.xlabel("Identificador pizza")
    plt.title("Precios de las pizzas")
    plt.xticks(rotation=45)
    plt.savefig("Preciosdelaspizzas.jpg")

    # por ultimo , haremos una tabla de tipos de pizza, veggie versus chicken
    fig = plt.figure(figsize=(30, 30))
    sns.barplot(x='pizza_id', y ='order_id' , data=details, palette=sns.color_palette("Oranges_d", len(details)))
    plt.xlabel('Identificador_PIZZA')
    plt.ylabel(' identificador pedido')
    plt.title(" Media de las cantidad de pedidos de cada pizza")
    plt.xticks(rotation=45)
    plt.savefig("categorias_del pedido.jpg")
    return


def crear_excel(fichero,fichero2,fichero3):

    # primera hoja del excel

    workbook= xlsxwriter.Workbook('Reporte_ejecutivo_maven.xlsx')
    worksheet_1=workbook.add_worksheet('Hoja de ingredientes')
    worksheet_1.write(0,0,'Ingrediente necesario')
    worksheet_1.write(0,1,' Cantidad necesaria para la proxima semana')
    for i in range(1,len(fichero['Ingredientes_necesarios'])) :
        worksheet_1.write(i,0,fichero['Ingredientes_necesarios'][i])
    for j in range(1,len(fichero['Cantidad a comprar por semana'])):
        worksheet_1.write(j, 1, fichero['Cantidad a comprar por semana'][j])
    # insertamos la imagen de la grafica de ingredientes ( en las siguientes hojas crearemos las graficas directamente en excel)
    img = 'Ingredientes_semanales.jpg'
    worksheet_1.insert_image('E7', img)


    # segunda hoja en el excel

    worksheet_2= workbook.add_worksheet('Hoja de pedidos')
    #vamos a añadir una grafica con las 5 pizzas más pedidas, las tenemos ordenadas de mayor a menor: 
    qantity_pizzas = fichero3.groupby(by='pizza_id')['quantity'].sum().sort_values(ascending=False)
    x1 =['big meat s','five cheese large','thai chicken l ','four cheese l',
        'classic dxl medium', 'the greek xxl','chicken alfredo s','green garden l ','calabrese l','mexicana s']
    y1= [1544,1142,1104,1036,934,22,78,79,81,133]
    fig = plt.figure(figsize=(10, 10))
    sns.barplot(x=x1, y=y1 , data= qantity_pizzas, palette=sns.color_palette("Oranges_d", 10))
    plt.xlabel('Identificador_PIZZA')
    plt.ylabel('Numero de pedidos ')
    plt.title(" Pizzas con mas y menos pedidos")
    plt.xticks(rotation=40)
    plt.savefig("masymenospedidas.jpg")
    worksheet_2.insert_image('B3', 'masymenospedidas.jpg')


    # tercera hoja del excel, con el precio de las pizzas mas y menos caras  , con fichero 2 
    worksheet_3= workbook.add_worksheet('Hojaejecutiva') 
    fichero33= fichero2.groupby(by='price')['price'].sum().sort_values(ascending=False)

    worksheet_3.write(0,0,'Pizza')
    worksheet_3.write(0,1,' Precio')
    for i in range(1,len(fichero2['pizza_id'])) :
        worksheet_3.write(i,0,fichero2['pizza_id'][i])
    for j in range(1,len(fichero2['pizza_id'])):
        worksheet_3.write(j, 1, fichero2['price'][j])

    # Ahora haremos una grafica para ver los 5 top y los 5 menores 
    chart = workbook.add_chart({'type': 'line'})
    chart.add_series({
    'values': 'Hojaejecutiva!$B$2:$B$93',
    'line':   {'color': '#FF9900'},
})
    chart.set_x_axis({
    'name': 'Pizzas',
    'name_font': {
        'name': 'Courier New',
        'color': '#92D050'
    },
    'num_font': {
        'name': 'Arial',
        'color': '#00B0F0',
    },
})

    chart.set_y_axis({
    'name': 'Prices',
    'name_font': {
        'name': 'Century',
        'color': 'red'
    },
    'num_font': {
        'bold': True,
        'italic': True,
        'underline': True,
        'color': '#7030A0',
    },
})
    worksheet_3.insert_chart('D2', chart)

    # cerrar el excel (más correcto que dejarlo abierto)

    workbook.close()

    return 
    

if __name__ == '__main__':
    dfs = extract()

    # fichero que hemos obtenido en la practica 2, fichero py adjuntado en el repo.
    fichero = pd.read_csv('resultado_pizzas_2.csv', sep=',')

    orderdetail = dfs[3]

    for i in range(len(orderdetail['pizza_id'])):
        if orderdetail['quantity'][i] == None:
            orderdetail.iloc[i]
    orderdetail=orderdetail[orderdetail['pizza_id'].notna()] 
    orderdetail=orderdetail[orderdetail['quantity'].notna()] #quita filas de Nan ya que no podemos dar un valor, no hay pistas 
    identifier = []
    identifier_2 = []
    # diccionario con claves erroneas que vamos a sustituir por aquellos valores correctos
    # a los que nos queremos referir.
    diccionario = { '@':'a','3':'e',
                 '0':'o',' ':'_',
                 '-':'_' }

    diccionario1 = { '@':'a','3':'e',
                 'One':'1','one':'1',
                 'two':'2','Two':'2',
                 '0':'o',' ':'_',
                 '-':'_', 'O': 'o','_1':'1', '_2':'2',
                 'e':'3'}

    for id in orderdetail['pizza_id']:
        id = str(id)
        for key in diccionario:
            pizza = id.replace(key,diccionario[key])
            id=pizza
        identifier.append(id)

    for cuanta_final in orderdetail['quantity']:
        cuanta_final=str(cuanta_final)
        for key in diccionario1:
            cuentas = cuanta_final.replace(key,diccionario1[key])
            cuanta_final = cuentas
        identifier_2.append(cuanta_final)
    lista_1=identifier
    lista_2=identifier_2

    orderdetail['pizza_id'] = lista_1
    # Reescribimos las dos columnas del dataframe que acabamos de procesar
    orderdetail['quantity'] = lista_2
    detalles_final = orderdetail
    detalles_final.reset_index(drop=True, inplace=True)

    tablas(fichero, dfs[2], detalles_final)
    crear_excel(fichero,dfs[2],detalles_final)
    # fichero con la informacion de los ingredientes necesarios por semana
    # guardarlo en el csv designado en la fucnion load
    salida_controlada()