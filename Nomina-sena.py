#DEBE INSTALAR CON PIP INSTALL LA LIBRERÍA DE PANDAS Y LA DE OPENPYXL
import pandas as pd
import os as os
from openpyxl import load_workbook
# Configurar pandas para mostrar todas las columnas
pd.set_option('display.max_columns', None)
# Configurar pandas para mostrar números en formato decimal completo (evitar que transforme documento de identidad en 1090+e15)
pd.set_option('display.float_format', '{:.0f}'.format)
nomina = {"DOC. DE IDENTIDAD" : [11807222,14449722,21466090,11789567], "Nombres":["Camilo","Lina Maria","Ana","Luis"],"Apellidos":["Perdomo","Calle","García","Maturana"],"Salario":[1000000,1562484,3000000,2750000],"HED":[3,5,7,10],"HEN":[2,3,0,1],"Días trab.":[30,30,22,30],"Valor día":[33333,52083,100000,91667],"Básico":[1000000,1562484,2200000,2750000],"VHED":[15625,40690,109375,143229],"VHEN":[14583,34179,0,20052],"Subtotal":[1030208,1637353,2309375,2913281],"Salud":[128776,204669,288672,364160],"Pensión":[164833,261976,369500,466125],"Aux. transporte":[106454,106454,0,0],"Total":[843053,1277161,1651203,2082996]}
def AgregarE():
    nomina["DOC. DE IDENTIDAD"].append(Identificacion())
    nomina["Nombres"].append(input("Ingrese el/los nombre/s del empleado: "))
    nomina["Apellidos"].append(input("Ingrese los apellidos del empleado: "))
    while True:
        salario = int(input("Ingrese el monto correspondiente al salario del empleado: "))
        if salario < 1000000:
            print("El salario no puede ser menor al salario mínimo vigente (1.000.000 COP)")
        else:
            nomina["Salario"].append(salario)
            break
    while True:
        hed = int(input("Ingrese el número de horas extras diurnas\ntrabajadas del empleado en el último mes: "))
        if hed < 0 or hed >30:
            print("Las horas extras diurnas trabajadas no corresponden al rango preestablecido (0 a 30 horas máximo)")
        else:
            nomina["HED"].append(hed)
            break
    while True:
        hen = int(input("Ingrese el número de horas extras nocturnas\ntrabajadas del empleado en el último mes: "))
        if hen < 0 or hen >30:
            print("Las horas extras nocturnas trabajadas no corresponden al rango preestablecido (0 a 30 horas máximo)")
        else:
            nomina["HEN"].append(hen)
            break
    while True:
        dias = int(input("Ingrese el número de días trabajados en el mes: "))
        if dias < 1 or dias >30:
            print("El número de días trabajadas no corresponden al rango preestablecido (1 a 30 días )")
        else:
            nomina["Días trab."].append(dias)
            break
    #Procesos secuenciales
    valorDia = salario/30
    nomina["Valor día"].append(valorDia)
    basico = dias*valorDia
    nomina["Básico"].append(basico)
    vhed = ((valorDia/8)*1.25)*hed
    nomina["VHED"].append(vhed)
    vhen = ((valorDia/8)*1.75)*hen
    nomina["VHEN"].append(vhen)
    subT = basico+vhed+vhen
    nomina["Subtotal"].append(subT)
    salud = subT*0.125
    nomina["Salud"].append(salud)
    pension = subT*0.16
    nomina["Pensión"].append(pension)
    if salario > 2000000:
        aux = 0
    else:
        aux = 106454
    nomina["Aux. transporte"].append(aux)
    neto = ((subT-(salud+pension))+aux)
    nomina["Total"].append(neto)
    
    print("\n\n\nempleado agregado correctamente! \n")
def MenuP():
    while True:
        opcion = int(input("::::::Menú Nómina SENA::::::\n1.Para ver el estado actual de la Nómina \n2.Para abrir la Nómina en Excel\n3.Para agregar empleados a la nómina actual\n4.Para actualizar la información de un empleado de la nómina\n5.Para eliminar un empleado de la nómina\n6.Para salir de la aplicación\nSeleccione una opción: "))
        if opcion == 1:
            df = pd.DataFrame(nomina)
            print(f"\n\n{df}\n TOTAL NÓMINA : ",sum(nomina["Total"]),"\n")
        elif opcion ==3:
            rango = int(input("\n\nIngrese el número de empleados que vas a agregar a la nómina: "))
            for i in range(rango):
                AgregarE()
            print("\n\nempleados agregados con éxito, volviendo al menú...\n\n")
        elif opcion ==2:
            df = pd.DataFrame(nomina)
            df["total"] = sum(nomina["Total"])
            df.to_excel("NominaSENA.xlsx")
            libro = load_workbook("NominaSENA.xlsx")
            hoja = libro.active
            columnasAjustar = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R"]
            for columna in columnasAjustar:
                hoja.column_dimensions[columna].width = 20
            libro.save("NominaSENA.xlsx")
            os.startfile("NominaSENA.xlsx")
            print("\n\n\nAbriendo documento de excel, recuerde cerrarlo al finalizar y/o modificar cambios para que sean aplicados\n\n \n")
        elif opcion ==4:
            documento = int(input("\n\n\nSolo podrás modificar un usuario mediante su doc. de identidad, sí el doc. de identidad es el erróneo, se recomienda borrar el empleado\nIngrese el número de documento del empleado que desea actualizar: "))
            Modificarempleado(documento)
        elif opcion ==5:
            documento = int(input("\n\n\nIngrese el número de documento del empleado que desea eliminar: "))
            Eliminarempleado(documento)
        elif opcion ==6:
            print("\n\n\n...Saliendo de la nómina...")
            break
def Identificacion(numero = None):
    while True:
        opcion = int(input("::::::Menú identificación::::::\n1.Para Cédula de Ciudadanía\n2.Para Cédula de extranjería\n3.Para Permiso especial de Permanencia\n4.Tarjeta de identidad\nSeleccione el número correspondiente a su documento de identidad: "))
        if opcion == 1:
            numero = int(input("Ingrese el número correspondiente a la Cédula de Ciudadanía: "))
            if len(str(numero))== 10:
                print("\n\n\nInformación agregada correctamente.\n\n")
                break
            else:
                print("\n\n\nel número ingresado es incorrecto, la cédula de ciudadanía debe tener 10 números\n\n")
        elif opcion == 2:
            numero = int(input("Ingrese el número correspondiente a la Cédula de extranjería: "))
            if len(str(numero))== 6:
                print("\n\n\nInformación agregada correctamente.\n\n")
                break
            else:
                print("\n\n\nel número ingresado es incorrecto, la cédula de extranjería debe tener 10 números\n\n")
        elif opcion == 3:
            numero = int(input("Ingrese el número correspondiente al PEP: "))
            if len(str(numero))== 15:
                print("\n\n\nInformación agregada correctamente.\n\n")
                break
            else:
                print("\n\n\nel número ingresado es incorrecto, el PEP debe tener 15 números\n\n")
        elif opcion == 4:
            numero = int(input("Ingrese el número correspondiente a la T.I: "))
            if len(str(numero))== 10:
                print("\n\n\nInformación agregada correctamente.\n\n")
                break
            else:
                print("\n\n\nel número ingresado es incorrecto, la Tarjeta de identidad debe tener 10 números\n\n")
           
        else:
            print("Opción incorrecta, intentálo de nuevo")
    return numero
def Eliminarempleado(documento):
    existe = False
    docI = 0
    for i in nomina["DOC. DE IDENTIDAD"]:
        if i == documento:
            existe = True
            break
        docI+= 1
    if existe:
        for claves in nomina:
            nomina[claves].pop(docI)
    else:
        print(f"\n\n\nEl empleado con documento {documento} no figura en la nómina\n\n\n")
def Modificarempleado(documento):
    existe = False
    docI = 0
    for i in nomina["DOC. DE IDENTIDAD"]:
        if i == documento:
            existe = True
            break
        docI+= 1
    if existe:
        nomina["Nombres"][docI] =input("Ingrese el/los nuevo/s nombre/s del empleado: ")
        nomina["Apellidos"][docI] =input("Ingrese los nuevos apellidos del empleado: ")
        while True:
            salario = int(input("Actualice el monto correspondiente al salario del empleado: "))
            if salario < 1000000:
                print("El salario no puede ser menor al salario mínimo vigente (1.000.000 COP)")
            else:
                nomina["Salario"][docI] = salario
                break
        while True:
            hed = int(input("Ingrese el número de horas extras diurnas\ntrabajadas del empleado en el último mes: "))
            if hed < 0 or hed >30:
                print("Las horas extras diurnas trabajadas no corresponden al rango preestablecido (0 a 30 horas máximo)")
            else:
                nomina["HED"][docI] = hed
                break
        while True:
            hen = int(input("Ingrese el número de horas extras nocturnas\ntrabajadas del empleado en el último mes: "))
            if hen < 0 or hen >30:
                print("Las horas extras nocturnas trabajadas no corresponden al rango preestablecido (0 a 30 horas máximo)")
            else:
                nomina["HEN"][docI] = hen
                break
        while True:
            dias = int(input("Ingrese el número de días trabajados en el mes: "))
            if dias < 1 or dias >30:
                print("El número de días trabajadas no corresponden al rango preestablecido (1 a 30 días )")
            else:
                nomina["Días trab."][docI] = dias
                break
        #Procesos secuenciales
        valorDia = salario/30
        nomina["Valor día"][docI] =valorDia
        basico = dias*valorDia
        nomina["Básico"][docI] = basico
        vhed = ((valorDia/8)*1.25)*hed
        nomina["VHED"][docI] = vhed
        vhen = ((valorDia/8)*1.75)*hen
        nomina["VHEN"][docI] = vhen
        subT = basico+vhed+vhen
        nomina["Subtotal"][docI] = subT
        salud = subT*0.125
        nomina["Salud"][docI] = salud
        pension = subT*0.16
        nomina["Pensión"][docI] = pension
        if salario > 2000000:
            aux = 0
        else:
            aux = 106454
        nomina["Aux. transporte"][docI] = aux
        neto = ((subT-(salud+pension))+aux)
        nomina["Total"][docI] = neto
        
        print("\n\n\nempleado modificado correctamente! \n")
        
    else:
        print(f"\n\n\nEl empleado con documento {documento} no figura en la nómina\n\n\n")
        
            
            
MenuP()