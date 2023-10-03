from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import time
from datetime import datetime
import sys
import ntpath
import xlsxwriter
import os



def config():
    #driver_path = "C:/Users/zuriel/AppData/Local/Programs/Python/Python310/Lib/site-packages/selenium/webdriver/chrome/chromedriver.exe"
    #brave_path = "C:/Program Files/BraveSoftware/Brave-Browser/Application/brave.exe"
    #chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe"

    option = webdriver.ChromeOptions()
    #option.binary_location = brave_path
    #Para descargar PDF's automáticamente
    profile = {"download.default_directory": 'C:/Users/zuriel/Desktop/OrdersBot/Prueba',
            "download.prompt_for_download": False, #Autodescarga
            "download.directory_upgrade": True, #Utiliza el directorio nuevo
            "safebrowsing.enabled": True, #Navegación segura :)
            "plugins.always_open_pdf_externally": True} #No muestra el PDF
    option.add_experimental_option("prefs", profile)
    option.add_argument("--disable-extensions")
    option.add_argument("--disable-print-preview")
    # option.add_argument("--incognito") OPCIONAL
    # option.add_argument("--headless") OPCIONAL

    # Nueva instancia de Chrome
    #browser = webdriver.Chrome(executable_path=driver, options=option)
    browser = webdriver.Chrome(ChromeDriverManager().install(),options=option)
    return browser




def login(browser,username,password):
    #Manejamos la pagina desde el navegador
    print("\n Abriendo el navegador...")
    browser.get("http://edifast1.com.mx/login.aspx")

    #Abrimos el navegador y buscamos id de los campos y posteamos la información
    print("\n Haciendo login como "+username+"...")
    campo = browser.find_element("id",'txtUsuario')
    campo.send_keys(username)
    campo = browser.find_element("id",'txtPass')
    campo.send_keys(password)

    browser.find_element("id",'btnEnviar').click() #Enviamos formulario

    browser.get('http://edifast1.com.mx/inicio.aspx')

    quitaAvisos(browser)

    return browser


def quitaAvisos(browser):
    time.sleep(2)
    browser.find_element("xpath",'//*[@id="modalNotification"]/div/div/div[3]/button').click()

def archivoXLSX():
    #Creamos el Excel
    dt = datetime.now()
    anio = str(dt.year)
    mes = str(dt.month)
    dia = str(dt.day)
    hora = str(dt.hour)
    min = str(dt.minute)
    seg = str(dt.second)

    orderbot_path = os.path.abspath("")
    nombre = "\\"+"REPORTE "+dia+"-"+mes+"-"+anio+" "+hora+"_"+min+"_"+seg+".xlsx"
    xlsx_path = orderbot_path+nombre
    print(xlsx_path)
    wb = xlsxwriter.Workbook(xlsx_path)
    hoja = wb.add_worksheet()
    negrita = wb.add_format({'bold': True}) #Formato de celda

    # Insertamos primeros valores
    hoja.write('A1','Orden de compra',negrita)
    hoja.write('B1','Tienda',negrita)
    hoja.write('C1','No. de Linea',negrita)
    hoja.write('D1','Descripcion',negrita)
    hoja.write('E1','Cantidad',negrita)
    hoja.write('F1','Precio',negrita)
    hoja.write('G1','Monto total de linea',negrita)
    hoja.write('H1','Monto total',negrita) 

    wb.close()


def analizarPagina(browser,inicio,fin):
    #Vamos a la bandeja de entrada
    try:
        browser.find_element("id",'lblBandejaEntrada').click()
    except:
        print("Credenciales incorrectas!!!")
        print("\n Saliendo del navegador...")
        browser.quit()
        print("\n Fin de la ejecución.")
        exit()
    print("\nLogin correcto!!!")
    print("\n Direccionando a la bandeja de entrada...")

    #Creamos el Excel
    dt = datetime.now()
    anio = str(dt.year)
    mes = str(dt.month)
    dia = str(dt.day)
    hora = str(dt.hour)
    min = str(dt.minute)
    seg = str(dt.second)

    orderbot_path = os.path.abspath("")
    nombre = "\\"+"REPORTE "+dia+"-"+mes+"-"+anio+" "+hora+"_"+min+"_"+seg+".xlsx"
    xlsx_path = orderbot_path+nombre
    print(xlsx_path)
    wb = xlsxwriter.Workbook(xlsx_path)
    hoja = wb.add_worksheet()
    negrita = wb.add_format({'bold': True}) #Formato de celda

    # Insertamos primeros valores
    hoja.write('A1','Orden de compra',negrita)
    hoja.write('B1','Tienda',negrita)
    hoja.write('C1','No. de Linea',negrita)
    hoja.write('D1','Descripcion',negrita)
    hoja.write('E1','Cantidad',negrita)
    hoja.write('F1','Precio',negrita)
    hoja.write('G1','Monto total de linea',negrita)
    hoja.write('H1','Monto total',negrita) 

    #Descargamos archivos...
    print("\n Analizando archivos...")

    #Ya estamos en la bandeja de entrada.
    html_id = "ContentPlaceHolder1_jGridConsulta_btnVerGrid_"
    pdf_id = "ContentPlaceHolder1_jGridConsulta_btnPDFGrid_"
    flag_1 = 0 #Banderas para controlar el flujo del análisis
    fila = 2
    pestania_id = "ContentPlaceHolder1_btn"
    
    #Analizamos el rango
    ind = str(int(inicio/100))
    lim = int(ind+"00")
    
    if inicio < lim:
        pest_id = int(ind)
        archivo_id = int(inicio%100)-2
    elif inicio == lim:
        pest_id = int(ind)-1
        archivo_id = 98
    elif inicio > lim:
        pest_id = int(inicio/100)
        archivo_id = int(inicio%100)-2

    
    orden = inicio #Contador de ordenes de compra

    while(flag_1 != 1 or orden <= fin):
        #Hacemos click en la siguiente pestaña.
        flag_2 = 0
        pest_id += 1
        aux = pestania_id+str(pest_id)
        print(aux)
        try:
            #Hace click en la siguiente pestaña. Si hay un error, termina.
            browser.find_element("id",aux).click()
            time.sleep(1)
            while(flag_2 != 1): 
                #Hacemos click en el siguiente archivo. Si hay un error, va a la siguiente pestaña.
                archivo_id+=1
                print("\n")
                aux = html_id+str(archivo_id) #ID del archivo HTML
                print(aux)
                #aux2 = pdf_id+str(archivo_id) #ID del PDF
                
                try:
                    #Ingresa al archivo HTML para extraer la información
                    boton = browser.find_element("id",aux)
                    browser.execute_script("arguments[0].click();", boton)
                    # Verifica si es una orden de compra
                    try:
                        prueba = browser.find_element("xpath",'/html/body/table/tbody/tr[1]/td[2]/strong/span').text
                        if prueba == "ORDEN DE COMPRA":
                            print("\n\n ========================================> ORDEN: "+str(orden))
                            print("========================================> FIN:   "+str(fin)+"\n")

                            #ORDEN DE COMPRA                    
                            oc = browser.find_element("xpath",'/html/body/table/tbody/tr[3]/td/center/table[2]/tbody/tr[2]/td[2]/b').text
                            print("Orden de compra: "+oc, end="")
                            # EMBARCAR A
                            embarca = browser.find_element("xpath",'/html/body/table/tbody/tr[3]/td/center/table[2]/tbody/tr[8]/td[2]/b').text
                            i=15
                            tienda = ""
                            while i < len(embarca):
                                tienda = tienda + embarca[i]
                                i+=1
                            print("Tienda: "+tienda, end="")
                            # NOTA
                            nota = browser.find_element("xpath",'/html/body/table/tbody/tr[3]/td/center/table[2]/tbody/tr[9]/td[2]/b').text
                            print("\n NOTA: "+nota, end="")
                            # Monto total
                            total = browser.find_element("xpath",'/html/body/table/tbody/tr[3]/td/table[3]/tbody/tr[2]/td[2]').text
                            print("Monto total: "+total, end="")
                            
                            #Iteramos sobre las filas de la orden de compra
                            i = 2
                            flag_3 = 0
                            while(flag_3 != 1):
                                try: 
                                    numStr = str(i)
                                    # No. DE LINEA
                                    path = "/html/body/table/tbody/tr[3]/td/table[2]/tbody/tr["+numStr+"]/td[1]"
                                    linea = browser.find_element("xpath",path).text

                                    print("\n No de Linea: "+linea, end="")
                                    # CANTIDAD
                                    path = "/html/body/table/tbody/tr[3]/td/table[2]/tbody/tr["+numStr+"]/td[2]"
                                    cant = browser.find_element("xpath",path).text
                                    print("\n Cantidad: "+cant, end="")
                                    # DESCRIPCION
                                    path = "/html/body/table/tbody/tr[3]/td/table[2]/tbody/tr["+numStr+"]/td[9]"
                                    desc = browser.find_element("xpath",path).text
                                    print("\n Descripcion: "+desc, end="")
                                    # PRECIO
                                    path = "/html/body/table/tbody/tr[3]/td/table[2]/tbody/tr["+numStr+"]/td[12]"
                                    precio = browser.find_element("xpath",path).text
                                    print("\n Precio: "+precio, end="")
                                    #MONTO TOTAL DE LINEA
                                    path = "/html/body/table/tbody/tr[3]/td/table[2]/tbody/tr["+numStr+"]/td[13]"
                                    monto = browser.find_element("xpath",path).text
                                    print("\n Monto total de linea: "+monto, end="")
                                    #Escribimos en la hoja...
                                    row = str(fila)
                                    a = "A"+row
                                    b = "B"+row
                                    c = "C"+row
                                    d = "D"+row
                                    e = "E"+row
                                    f = "F"+row
                                    g = "G"+row
                                    h = "H"+row
                                    hoja.write(a,oc)
                                    hoja.write(b,tienda)
                                    hoja.write(c,linea)
                                    hoja.write(d,desc)
                                    hoja.write(e,cant)
                                    hoja.write(f,precio)
                                    hoja.write(g,monto)
                                    hoja.write(h,total) 
                                    i+=1
                                    fila+=1
                                except:
                                    print("\nFin de Orden de compra "+oc)
                                    flag_3 = 1
                            #Regresamos a bandeja de entrada...
                            browser.execute_script("window.history.go(-1)")
                            #Abrimos PDF
                            #browser.find_element("id",aux2).click() #Se descarga automáticamente
                            #print(aux2+" descargado.")
                            orden+=1 #Siguiente orden de compra...
                            if orden > fin:
                                flag_2 = 1
                                flag_1 = 1
                                print("\nHemos terminado ;)")
                                time.sleep(1)
                                print("\n Saliendo del navegador...")
                                browser.quit()
                                print("\n Fin de la ejecución.")
                                wb.close()

                        else:
                            print("Esto es una: "+prueba)
                            #Regresamos a bandeja de entrada...
                            browser.execute_script("window.history.go(-1)")
                    except:
                        print(sys.exc_info()[0])
                        print("Algo salio mal!!!!!!!")
                except:
                    print(sys.exc_info()[0])
                    print("OOOPS! Llegamos al final de esta pestaña... Vamos a la que sigue! :)")
                    flag_2 = 1
                    archivo_id = -1
        except:
            flag_2 = 1
            flag_1 = 1
            print(sys.exc_info()[0])
            print("No existen más ordenes de compra.")
            time.sleep(1)
            print("\n Saliendo del navegador...")
            browser.quit()
            print("\n Fin de la ejecución.")
            wb.close()
            exit()


def main():
    os.system("cls")
    print(
        "                                                 .     .,,,                                                    \n"
        "                                           ,.     ,***        ,,,,                                             \n"                            
        "                                       .     .///,       ,///*                                                 \n"                               
        "                                    .,,**////(        ((((,        **,.                                        \n"                      
        "                                   ,,**///(.       ((((.        /(///**,.                                      \n"                    
        "                                 .,***//*       (###,        /((((((//***,                                     \n"                   
        "                                 ,,          (###*        *###(           ,                                    \n"                   
        "                                .,**//((((((##/        ,####,       ,(//**,.                                   \n"                  
        "                                ,***//((((((        .####/        ((((//**,.                                   \n"                 
        "                                ,,***//(((*        (###/        *((((((//**,                                   \n"                  
        "                                ,   .,**.       /###(        .##((((((/*,,.                                    \n"                   
        "                                 ,**///((((//(##(         (###*      ./**,.                                    \n"                  
        "                                  ,***//((((((.        (###*       ///**,.                                     \n"                   
        "                                   ,,**///(.        ((((*       *///***,                                       \n"                  
        "                                       ,         ((((.       (////***,                                         \n"                 
        "                                       .,,****///*       ,///,     .                                           \n"                 
        "                                           ,,.       ****.    ,,                                               \n"                    
        "                                                     ..                                                        \n"                 
        "                                                                                                               \n"                                                            
        "                                                                                                               \n"                                                            
        "                                &&%  &&&   ,&&&  &&%    &&%   #&&&&&%.  %&&                                    \n"                  
        "                                &&%  &&& .&&&    &&%    &&%  &&&    /   %&&                                    \n"                  
        "                                &&%  &&&&&&%     &&%    &&%  #&&&&&&#   %&&                                    \n"                  
        "                                &&%  &&&& &&&    &&%    &&%       .%&&, %&&                                    \n"                  
        "                                &&%  &&&   (&&#  ,&&&&&&&&. /&&&&%&&&&  %&&                                    \n"
        "                                                                                                               \n"                                                            
        "                                                                                                               \n"
        "                                                                                                               \n"                                                            
        "    ▄▄▄▄▄▄▄▄▄▄▄  ▄▄▄▄▄▄▄▄▄▄▄  ▄▄▄▄▄▄▄▄▄▄   ▄▄▄▄▄▄▄▄▄▄▄  ▄▄▄▄▄▄▄▄▄▄▄  ▄▄▄▄▄▄▄▄▄▄   ▄▄▄▄▄▄▄▄▄▄▄  ▄▄▄▄▄▄▄▄▄▄▄     \n"
        "   ▐░░░░░░░░░░░▌▐░░░░░░░░░░░▌▐░░░░░░░░░░▌ ▐░░░░░░░░░░░▌▐░░░░░░░░░░░▌▐░░░░░░░░░░▌ ▐░░░░░░░░░░░▌▐░░░░░░░░░░░▌    \n"
        "   ▐░█▀▀▀▀▀▀▀█░▌▐░█▀▀▀▀▀▀▀█░▌▐░█▀▀▀▀▀▀▀█░▌▐░█▀▀▀▀▀▀▀▀▀ ▐░█▀▀▀▀▀▀▀█░▌▐░█▀▀▀▀▀▀▀█░▌▐░█▀▀▀▀▀▀▀█░▌ ▀▀▀▀█░█▀▀▀▀     \n"
        "   ▐░▌       ▐░▌▐░▌       ▐░▌▐░▌       ▐░▌▐░▌          ▐░▌       ▐░▌▐░▌       ▐░▌▐░▌       ▐░▌     ▐░▌         \n"
        "   ▐░▌       ▐░▌▐░█▄▄▄▄▄▄▄█░▌▐░▌       ▐░▌▐░█▄▄▄▄▄▄▄▄▄ ▐░█▄▄▄▄▄▄▄█░▌▐░█▄▄▄▄▄▄▄█░▌▐░▌       ▐░▌     ▐░▌         \n"
        "   ▐░▌       ▐░▌▐░░░░░░░░░░░▌▐░▌       ▐░▌▐░░░░░░░░░░░▌▐░░░░░░░░░░░▌▐░░░░░░░░░░▌ ▐░▌       ▐░▌     ▐░▌         \n"
        "   ▐░▌       ▐░▌▐░█▀▀▀▀█░█▀▀ ▐░▌       ▐░▌▐░█▀▀▀▀▀▀▀▀▀ ▐░█▀▀▀▀█░█▀▀ ▐░█▀▀▀▀▀▀▀█░▌▐░▌       ▐░▌     ▐░▌         \n"
        "   ▐░▌       ▐░▌▐░▌     ▐░▌  ▐░▌       ▐░▌▐░▌          ▐░▌     ▐░▌  ▐░▌       ▐░▌▐░▌       ▐░▌     ▐░▌         \n"
        "   ▐░█▄▄▄▄▄▄▄█░▌▐░▌      ▐░▌ ▐░█▄▄▄▄▄▄▄█░▌▐░█▄▄▄▄▄▄▄▄▄ ▐░▌      ▐░▌ ▐░█▄▄▄▄▄▄▄█░▌▐░█▄▄▄▄▄▄▄█░▌     ▐░▌         \n"
        "   ▐░░░░░░░░░░░▌▐░▌       ▐░▌▐░░░░░░░░░░▌ ▐░░░░░░░░░░░▌▐░▌       ▐░▌▐░░░░░░░░░░▌ ▐░░░░░░░░░░░▌     ▐░▌         \n"
        "    ▀▀▀▀▀▀▀▀▀▀▀  ▀         ▀  ▀▀▀▀▀▀▀▀▀▀   ▀▀▀▀▀▀▀▀▀▀▀  ▀         ▀  ▀▀▀▀▀▀▀▀▀▀   ▀▀▀▀▀▀▀▀▀▀▀       ▀          \n"
        "                                                                                                               \n"
        "                                                                                                               \n"
        " ======================================== B  I  E  N  V  E  N  I  D  O ===================================     \n"
        "                                                                                                               \n"
        " ************************************************** I N F O **********************************************     \n"
        " [+] Este bot ingresa a http://http://edifast1.com.mx/ y obtiene información relevante de las ordenes de compra\n\n"
        " [+] Para poder ingresar, ten a la mano el nombre de usuario y la contraseña que sueles utilizar                \n\n"
        " [+] Necesitas indicar en qué orden de compra quieres iniciar. Esto no es el identificador de la orden de compra,\n"
        "     es el número de orden de compra que aparece en la página. Se marca con un '#'                           \n\n"
        " [+] El programa recorre las ordenes de compra de una en una. Por ello, debes indicar también el número de la   \n"
        "     última orden de compra. Esto establece un rango de descarga de datos.                                   \n\n"
        " [+] Los archivos XLSX generados se guardan en la ruta donde estás ahora mismo  :)                 \n\n"
        " [+] Ojalá te sirva! :D                                                                                      \n\n"
    )
    username = input("Ingrese el usuario: ")
    password = input("Ingrese la contraseña: ")

    flag = 0
    while(flag != 1):
        print("\n")
        inicio = int(input("Quiero iniciar en la orden de compra numero: "))
        
        if inicio < 1:
            print(">>>>>> ERROR: El inicio tiene que ser mínimamente 1...")
        else:
            flag = 1
    flag = 0
    while(flag != 1):
        fin = int(input("Quiero que termine en la orden de compra numero (Hasta donde quieras): "))

        if fin < 1:
            print(">>>>>> ERROR: El fin tiene que ser mínimamente 1 o igual que el inicio")
        else:
            if fin < inicio:
                print(">>>>>> ERROR: El fin no puede ser menor que el inicio.")
            else:
                flag = 1

    browser = config()
    browser = login(browser,username,password)
    analizarPagina(browser,inicio,fin)



main()