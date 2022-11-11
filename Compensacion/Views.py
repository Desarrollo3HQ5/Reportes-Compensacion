from django.http import HttpResponse
# Importar para manejar el contexto y asi pasar atributos o informacion
from django.template import Template, Context
# Importar cargador de plantillas
from django.template import loader
#improtar shortcuts, metodo para renderisar plantillas y optimizar codigo al cargar plantillas
from django.shortcuts import render
##Librerias para el proceso del reporte horizontal
import pandas as pd
from xlsxwriter import Workbook
from io import BytesIO
import io,csv
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from openpyxl import load_workbook
def Base(request):
    
    # nombre = "JUAN ROJAS"
    # ctx = {"Nombre_persona":nombre}
    # return render(request, "index.html", ctx)
    return render(request, "Base.html")
##HTML de vistas
def Home(request):
    
    # nombre = "JUAN ROJAS"
    # ctx = {"Nombre_persona":nombre}
    # return render(request, "index.html", ctx)
    return render(request, "Home.html")
def ReporteHorizontal(request):
    
    # nombre = "JUAN ROJAS"
    # ctx = {"Nombre_persona":nombre}
    # return render(request, "index.html", ctx)
    return render(request, "ReporteHorizontal.html")
def txtSS(request):
    
    # nombre = "JUAN ROJAS"
    # ctx = {"Nombre_persona":nombre}
    # return render(request, "index.html", ctx)
    return render(request, "TXTSS.html")

##Procesos
def procesar(request):
    
    IdPeriodo = request.GET.get("id")
    # IdPeriodo = "35346"
    # Dataframe final
    Horizontal = pd.DataFrame()
    #Traer información
    URL= "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Conceptos_Nomina_Desarrollo/jwhRFUOR47TqCS9AAT82eCybwgdmgeArEtKG7U8H9s3hSjTzBd3G8bPdg37PHVygvxurxwCQvMCgHRG68dOCWKTmMWaQJU2TMwnr?ID_Periodo="+IdPeriodo
    # url_ = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Prenomina/WWjRAOJ2MGyyNGd5BxdvwApYGzgq5A9AQ5Q6bUmpsTQvWTMJE4qE5MyKnY4KKPXneurq8RnTZ2O698AO8N2KQ7Fa7qt4hpwSet0K?Periodo=" + _idPeriodo + "&zc_FileName=PreNomina_" + _idPeriodo;
    df = pd.read_excel(URL)
    df1 = pd.DataFrame(df)

    
    Contrato  = df1['Numero de Contrato'].unique().tolist()
    # print(Contrato)
    #Filtrar cada concepto unico que existe en ese reporte
    Conceptos = df1['Concepto'].unique().tolist()
    Conceptos.sort()
    ConceptosDev = []
    ConceptosDed = []
    for conceptosx in Conceptos:
        Valores = df1['Concepto'] == str(conceptosx)
        ContratoPos = df1[Valores]
        #Sumatoria
        Total = ContratoPos['Neto'].sum()
        if(Total >= 0 ):
            ConceptosDev.append(conceptosx)
        else:
            ConceptosDed.append(conceptosx)
    Conceptos.clear()
    Conceptos= ConceptosDev + ConceptosDed
    # print(Conceptos)
    #Se obtienen los datos dependiendo del empleado

    #Obtener información para agregar al nuevo data frame
    for i in Contrato:
        Valores = df1['Numero de Contrato'] == str(i)
        ContratoPos = df1[Valores]
        FilaAgregar = {}
        ##Informacion general inicial
        FilaAgregar["Temporal"] = ContratoPos.iloc[0]['Temporal']
        FilaAgregar["Empresa"] = ContratoPos.iloc[0]['Empresa']
        FilaAgregar["ID Periodo"] = ContratoPos.iloc[0]['ID Periodo']
        FilaAgregar["Tipo de Perido"] = ContratoPos.iloc[0]['Tipo de Perido']
        FilaAgregar["Mes"] = ContratoPos.iloc[0]['Mes']
        FilaAgregar["Numero de Contrato"] = ContratoPos.iloc[0]['Numero de Contrato']
        FilaAgregar["Nombres y Apellidos"] = ContratoPos.iloc[0]['Nombres y Apellidos']
        FilaAgregar["Numero de Identificación"] = ContratoPos.iloc[0]['Numero de Identificación']
        FilaAgregar["Centro de Costo"] = ContratoPos.iloc[0]['Centro de Costo']
        
        if(ContratoPos.iloc[0]['Dependencia']):
            FilaAgregar["Dependencia"] = ContratoPos.iloc[0]['Dependencia']
        if(ContratoPos.iloc[0]['Proceso']):
            FilaAgregar["Proceso"] = ContratoPos.iloc[0]['Proceso']
        # if(ContratoPos.iloc[0]['Area']):
        #     FilaAgregar["Area"] = ContratoPos.iloc[0]['Area']
        # if(ContratoPos.iloc[0]['Nivel']):
        #     FilaAgregar["Nivel"] = ContratoPos.iloc[0]['Nivel']
        FilaAgregar["Fecha Ingreso"] = pd.to_datetime(ContratoPos.iloc[0]['Fecha Ingreso']).date()
        FilaAgregar["Fecha Retiro"] = pd.to_datetime(ContratoPos.iloc[0]['Fecha Retiro']).date()
        FilaAgregar["Cargo"] = ContratoPos.iloc[0]['Cargo']
        FilaAgregar["Salario Base"] = ContratoPos.iloc[0]['Salario Base']
        SumatoriaNetoDev = 0
        SumatoriaNetoDed = 0
        #Ciclo para tomar informacion de los conceptos
        for elemento in ConceptosDev:
            de = ContratoPos["Concepto"] == str(elemento)
            Conce= ContratoPos[de]
            Unidades = 0
            Neto = 0
            if (Conce.empty == False):
                # print(Conce)
                # print(elemento + " / Unidades")
                Unidades = Conce["Horas"].sum()
                Neto = Conce["Neto"].sum()
                SumatoriaNetoDev += Neto
            if (elemento + " / Neto" in FilaAgregar):
                FilaAgregar[elemento + " / Unidades"] += Unidades
                FilaAgregar[elemento + " / Neto"] += Neto 
            else:
                FilaAgregar[elemento + " / Unidades"] = Unidades
                FilaAgregar[elemento + " / Neto"] = Neto 
        FilaAgregar["Total Devengo"] = SumatoriaNetoDev
        for elemento in ConceptosDed:
            de = ContratoPos["Concepto"] == str(elemento)
            Conce= ContratoPos[de]
            Unidades = 0
            Neto = 0
            if (Conce.empty == False):
                # print(Conce)
                # print(elemento + " / Unidades")
                Unidades = Conce["Horas"].sum()
                Neto = Conce["Neto"].sum()
                SumatoriaNetoDed += Neto
            if (elemento + " / Neto" in FilaAgregar):
                FilaAgregar[elemento + " / Unidades"] += Unidades
                FilaAgregar[elemento + " / Neto"] += Neto 
            else:
                FilaAgregar[elemento + " / Unidades"] = Unidades
                FilaAgregar[elemento + " / Neto"] = Neto 
        FilaAgregar["Total Deduccion"] = SumatoriaNetoDed
        FilaAgregar["Neto A Pagar"] = SumatoriaNetoDev - abs(SumatoriaNetoDed)
        # Informacion general de provisiones y SS 
        FilaAgregar["EPS"] = ContratoPos.iloc[0]['EPS']
        FilaAgregar["AFP"] = ContratoPos.iloc[0]['AFP']
        FilaAgregar["ARL"] = ContratoPos.iloc[0]['ARL']
        FilaAgregar["Riesgo ARL"] = ContratoPos.iloc[0]['Riesgo ARL']
        FilaAgregar["CCF"] = ContratoPos.iloc[0]['CCF']
        FilaAgregar["SENA"] = ContratoPos.iloc[0]['SENA']
        FilaAgregar["ICBF"] = ContratoPos.iloc[0]['ICBF']
        FilaAgregar["Total Seguridad Social"] = ContratoPos.iloc[0]['EPS'] + ContratoPos.iloc[0]['AFP'] + ContratoPos.iloc[0]['ARL'] + ContratoPos.iloc[0]['CCF'] + ContratoPos.iloc[0]['SENA'] + ContratoPos.iloc[0]['ICBF']
        FilaAgregar["Vacaciones tiempo"] = ContratoPos.iloc[0]['Vacaciones tiempo']
        FilaAgregar["Prima"] = ContratoPos.iloc[0]['Prima']
        FilaAgregar["Cesantías"] = ContratoPos.iloc[0]['Cesantías']
        FilaAgregar["Interés cesantías"] = ContratoPos.iloc[0]['Interés cesantías']
        FilaAgregar["Total provisiones"] = ContratoPos.iloc[0]['Vacaciones tiempo'] + ContratoPos.iloc[0]['Prima'] + ContratoPos.iloc[0]['Cesantías'] + ContratoPos.iloc[0]['Interés cesantías']
        
        Horizontal = Horizontal.append(FilaAgregar, ignore_index=True)
    NombreDocumento = "Horizontal " + Horizontal.iloc[0]['Empresa'] +"-"+ str(Horizontal.iloc[0]['Mes'])+ "-" + str(Horizontal.iloc[0]['Tipo de Perido'])
    
    heads = Horizontal.columns.values
    FilaAgregar = {}
    Validador = False
    # Horizontal.count(1)
    for k in heads:
        if(str(k).__contains__("Salario Base")):
            Validador = True
        if(Validador):
            Horizontal[k] = Horizontal[k].astype('float')
            FilaAgregar[k] = sum(Horizontal[k])
        
    Horizontal = Horizontal.append(FilaAgregar, ignore_index=True)

    wb = Workbook()
    ws = wb.active
    for r in dataframe_to_rows(Horizontal, index=False, header=True):
        ws.append(r)

    ws.insert_rows(1)
    ws.insert_rows(1)
    ws.insert_rows(1)
    
    Horizontal = pd.DataFrame(ws.values)
    with BytesIO() as b:
        # Use the StringIO object as the filehandle.
        writer = pd.ExcelWriter(b, engine='xlsxwriter')
        Horizontal.to_excel(writer, sheet_name='Sheet1',index = False, header = False)
        # Horizontal.to_excel(writer, sheet_name='Sheet1',index = False, header = False)
        # Edicion del estilo del excel
        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]
        format = workbook.add_format()
        format.set_pattern(1)
        format.set_bg_color('#AFAFAF')
        format.set_bold(True) 
        
        worksheet.write_string(1, 1,str(ContratoPos.iloc[0]['Temporal']),format)
        worksheet.write_string(1, 2,str(ContratoPos.iloc[0]['Empresa']),format)
        worksheet.write_string(1, 3,str(ContratoPos.iloc[0]['ID Periodo']),format)
        worksheet.write_string(1, 4,str(ContratoPos.iloc[0]['Mes']),format)
        worksheet.write_string(1, 5,str(ContratoPos.iloc[0]['Tipo de Perido']),format)
        worksheet.write_string(1, 1,str(ContratoPos.iloc[0]['Temporal']),format)
        worksheet.write_string(1, 2,str(ContratoPos.iloc[0]['Empresa']),format)
        worksheet.write_string(1, 3,str(ContratoPos.iloc[0]['ID Periodo']),format)
        worksheet.write_string(1, 4,str(ContratoPos.iloc[0]['Mes']),format)
        worksheet.write_string(1, 5,str(ContratoPos.iloc[0]['Tipo de Perido']),format)
        
        contador = 0
        MaxFilas = len(Horizontal.axes[0])
        Totales = Horizontal.loc[MaxFilas -1]
        for k in heads:
            
            worksheet.write_string(3, contador,str(k),format)
            contador += 1
            
        contador = 0
        for k in Totales:
            Dato = ""
            if(str(k) != "nan"):
                Dato = str(k)
            worksheet.write_string(MaxFilas-1, contador,Dato ,format)
            contador += 1
        ##Modificar el excel
        writer.save()
        # Set up the Http response.
        filename = NombreDocumento+'.xlsx'
        response = HttpResponse(
            b.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = 'attachment; filename=%s' % filename
        return response
    return render(request, "Resultado.html")
def procesarTXTSS(request):
    # request.GET.get("id")
    # Dataframe final
    # return render(request, "Resultado.html")
    NombreTemporal = request.GET.get("Empresa")
    NombreTemporal_ = NombreTemporal.replace(" ","%20")
    Anio = request.GET.get("Anio")
    Mes = request.GET.get("Mes")
    # NombreTemporal = "HQ5 S.A.S"
    # Anio = "2022"
    # Mes = "10"
    TXT_Final = pd.DataFrame()
    ##URL DEL XLS DE SS
    URL = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/TXT_SS_DESARROLLO/3HO1RZORhePyRgar44EefyEhhD27umsJE7GeJJhCDwwx2ngQ2KEHHGTCB1mYQtFktzmgSyHG2qRWsnu3ZGbW8N97TZtX709N3DAC?NOMBRE_EMPRESA=" + NombreTemporal_ + "&PENSION_ANO=" + Anio + "&PENSION_MES=" + Mes
    # URL = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/TXT_SS_DESARROLLO/3HO1RZORhePyRgar44EefyEhhD27umsJE7GeJJhCDwwx2ngQ2KEHHGTCB1mYQtFktzmgSyHG2qRWsnu3ZGbW8N97TZtX709N3DAC?NOMBRE_EMPRESA=HQ5%20S.A.S&PENSION_ANO=2022&PENSION_MES=10"
    #CNVERTIR XLS EN DATAFRAME
    df = pd.read_excel(URL)
    df1 = pd.DataFrame(df)
    if(df1.empty):
        ctx = {"Mensaje":"El reporte esta vacio para este periodo y temporal, validar informacion"}
        return render(request, "Resultado.html", ctx)
    else:
        #Armar la primera linea dependiendo temporal con nit y demas
        #Tomar sumatoria de ibc
        IBCTotal_ = sum(df1["IBC EPS"])
        ##Contador total de lineas del txt
        MaxFilas = len(df1.axes[0])
        ##Datos fijos del txt
        E_TRegristro_ = "01"
        E_Modalidad_ = "2"
        E_Secuencia_ = "0001"
        #Funcion para reemplazar acentos del txt
        def normalize(s):
            replacements = (
                ("á", "a"),
                ("é", "e"),
                ("í", "i"),
                ("ó", "o"),
                ("ú", "u"),
            )
            for a, b in replacements:
                s = s.replace(a, b).replace(a.upper(), b.upper())
            return s
        #Normalizar nombre de la temporal
        NombreTemporal = normalize(NombreTemporal)
        NombreTemporal = NombreTemporal.replace(".", "")
        NombreTemporal = NombreTemporal.replace("-", "_")
        #Completar 200 espacios y se rellenan con espacios en blanco
        Espacios_ = 200 - len(NombreTemporal)
        NombreTemporal = NombreTemporal + (Espacios_ * " ")
        E_RazonSocial_ = NombreTemporal
        #Nit tempora - se obtiene del dataframe y ademas se completan 16 espacios en blanco
        E_TDocumento_ = "NI"
        E_NDocumento_ = str(df.iloc[0]['NIT'])
        Espacios_ = 16 - len(E_NDocumento_)
        E_NDocumento_ = E_NDocumento_ + (Espacios_ * " ")
        #Numero de verificacion - se obtiene del dataframe
        E_DVerificacion_ = str(df.iloc[0]['Número de verificacación'])
        ##Informacion fija, planilla E 
        E_TPlantilla_ = "E"
        E_Blanco1_ = 20 * " "
        E_FPresentacion_ = "U"
        #Datos fijos y completar con espacios en blanco
        E_CSucursal_ = "01" + (8 * " ")
        E_NSucursal_ = "01" + (38 * " ")
        #Codigo de la ARL, se obtiene del data frame y se completan 6 espacios
        E_CodARL_ = str(df.iloc[0]['Código ARL'])
        Espacios_ = 6 - len(E_CodARL_)
        E_CodARL_ = E_CodARL_ + (Espacios_ * " ")
        ##Fechas de pension y salud
        E_FechaPension_ = str(df.iloc[0]['Año pensión']) + "-" + str(df.iloc[0]['Mes pensión'])
        E_FechaSalud_ = str(df.iloc[0]['Año pensión']) + "-" + str(int(df.iloc[0]['Mes pensión'] + 1))
        #Informaicon fija, con 10 espacios en 0
        E_NRadicacion_ = "0000000000"
        E_FechaPago = 10 * "0"

        ##Total de empleados, completadno con 0 a la izquierda
        E_TotalEmpleados_ = str(MaxFilas)
        Ceros_ = 5 - len(E_TotalEmpleados_)
        E_TotalEmpleados_ = (Ceros_ * "0") + E_TotalEmpleados_ 
        ##Total de ibc a pagar completando con 0 espacios hasta 12 
        E_TotalNomina_ = str(IBCTotal_).replace(".", "")
        Ceros_ = 12 - len(E_TotalNomina_)
        E_TotalNomina_ = (Ceros_ * "0") + E_TotalNomina_
        #Informacion Fija
        E_TAportente_ = "01"
        E_CodOperador_ = "00"
        #Se concatena la informacio anterior en una variable y luego se agrega al diccionario
        TXT_ = E_TRegristro_ + E_Modalidad_ + E_Secuencia_ + E_RazonSocial_ + E_TDocumento_ + E_NDocumento_ + E_DVerificacion_ + E_TPlantilla_ + E_Blanco1_ + E_FPresentacion_ + E_CSucursal_ + E_NSucursal_ + E_CodARL_ + E_FechaPension_ + E_FechaSalud_ + E_NRadicacion_ + E_FechaPago + E_TotalEmpleados_ + E_TotalNomina_ + E_TAportente_ + E_CodOperador_
        FilaAgregar = {}
        FilaAgregar["TXT"] = TXT_
        ##Se agrega el diccionario a un nuevo dataframe
        TXT_Final = TXT_Final.append(FilaAgregar, ignore_index=True)
        
        ## Dar consecutivo y completar linea de empleados
        ## Se añade al dataframe final
        for i in range(len(df)):
            FilaAgregar = {}
            Ceros_ = 5 - len(str(i +1))
            Contador = (Ceros_* "0") + str(i+1)
            FilaAgregar["TXT"] =  "02"+ Contador + str(df.iloc[i]['TXT'])
            TXT_Final = TXT_Final.append(FilaAgregar, ignore_index=True)

        ##Para exportar en txt
        NombreTXT_ = "TXT-" + NombreTemporal_ + "-" + Anio + "-" + Mes 
        file_name = open(NombreTXT_ + ".txt", "w+")
        Texto_ = ""
        for fila in TXT_Final["TXT"]:
            Texto_ += (str(fila)+"\n")
        file_name.write(Texto_)
        file_name.close()

        # to read the content of it
        read_file = open(NombreTXT_ + ".txt", "r")
        response = HttpResponse(read_file.read(), content_type="text/plain,charset=utf8")
        read_file.close()

        response['Content-Disposition'] = 'attachment; filename="{}.txt"'.format(NombreTXT_)
        return response
        