import smtplib, ssl
import getpass
import openpyxl

book = openpyxl.load_workbook('excel/prueba.xlsx', data_only=True)
hoja = book.active

celdas = hoja['A2' : 'H6']

lista_empleados = []

for fila in celdas:
    empleado = [celda.value for celda in fila]
    lista_empleados.append(empleado)

username = input('ingreses su nombre de usuario: ')
password = getpass.getpass('ingrese su password: ')

context = ssl.create_default_context()

with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as server:
    server.login(username, password)
    print('inició sesión!')

    for empleado in lista_empleados:
        destinatario = empleado[4]
        mensaje = f'hola {empleado[1]}, este mes ganaste {empleado[7]}'
        server.sendmail(username, destinatario, mensaje)
        print('Mensaje enviado')
