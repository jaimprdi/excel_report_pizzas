# excel_report_pizzas
Generar un fichero excel con un reporte ejecutivo, un reporte de ingredientes, un reporte de pedidos (uno por cada hoja en el fichero de excel) para el dataset de Maven Pizzas trabajado en el bloque 3. El obtencion_resultados_2.py se ejecuta y se obtiene el fichero resultados (csv) que guarda el numero de ingredientes por semana. Luego usaremos este fichero en el main. 
IMportar : 
(necesario tenerlas instaladas) 

import pandas as pd
import sys
import matplotlib.pyplot as plt
import seaborn as sns
import xlsxwriter 
from openpyxl.chart import BarChart, Reference
