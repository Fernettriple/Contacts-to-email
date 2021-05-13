import pandas as pd
import pyperclip
import os


roles = {
    1: "Principal Investigator",
    2: "Sub-Investigator",
    3: "Study Coordinator",
    4: "Ethics/Regulatory Contact",
    5: "IP/Device Shipping Contact"    
}

roles_lista = '\n'.join([f"{str(key)}: {values}" for key,values in roles.items()])

mensaje_bienvenida = ('''
Bienvenid@!

Este programita es MUY sencillo. Vos le decis que sitio necesitas copiar la direcion de mails y te lo busca en el reporte de CTMS 
y te lo copia al portapapeles, cosa que vos directamente apretas control + v y ya podes pegarlo en un mail.
Para que funcione correctamente, necesitas tener presente un archivo llamado "contactos.xlsx" que obtenes al bajar el reporte de CTMS
para un determinado protocolo. 

Para generar correctamente este reporte, tenes que ir a CTMS > Protocols > Entras al protocolo que te 
interesa, > Pestaña de Contacts > Icono de Engranaje(o ruedita, o gear) > Export > Y ahi seleccionas "All rows in current Query" y "Tab Delimited Text File".
Ahora, PARA COLMO, cuando te baje el reporte tenes que entrar, y guardarlo con el nombre "contactos", en formato xlsx (Excel Workbook).
Lo bueno es que una vez que lo tenes seteado, no tenes que hacer esto siempre (si no seria un bodrio) a menos que el reporte necesite una actualizacion

Ahora, lo primero es checkear que lo hagas hecho bien! 
Presiona cualquier tecla para continuar.
''')
mensaje_sitio='''
Decime el numero de sitio
'''
mensaje_roles = '''
Ahora te voy a preguntar que roles te interesa que te copie. Podes seleccionar varios o uno solo. \n'''+roles_lista+'''
Lo UNICO importante es que los separes con una coma(,) o punto (.). Si no el programa explota y Nico llora.
Por ejemplo:
1,2,3 Si quiero PI, SubI y SC (no agreges coma o punto al final que si no tira error)
'''
mensaje_final = '''
Listo! Ahora podes apretar Ctrl + v y pegar los mails. Si queres que el programa se ejecute de nuevo apretar enter, y si no o cerrado lo apreta cualqueir otra cosa.
'''
input(mensaje_bienvenida)
while 1:
    if "contactos.xlsx" in os.listdir("."):    
        try:        
            contacts = pd.read_excel ("./contactos.xlsx")
        except:
            print("No puedo acceder al archivo. Verifica que no lo tengas abierto.")
        contacts.rename(columns = {"Site #":"Site"}, inplace = True)
        contacts = contacts[["Country","Site","Role","Last Name", "First Name","Start Date", "End Date", "E-Mail" ]]
        contacts.dropna(subset = ["Site"],inplace=True)
        contacts["Site"] = contacts["Site"].astype(int)
        contacts = contacts.loc[contacts["End Date"].isna()].sort_values(by = "Site")
        print(mensaje_sitio)
        while 1:
            sitio = input()
            try:
                sitio = int(sitio)
                break
            except:
                print(f'Error: No se encontro el sitio escrito {sitio}. Puede ser que haya sido mal tipeado.')  
        print(mensaje_roles)
        while 1:
            a_mandar = input()
            try:
                if "," in a_mandar:
                    if a_mandar.endswith(","):
                        a_mandar = a_mandar.strip(",")
                    numero_roles = a_mandar.split(',')
                elif "." in a_mandar:
                    if "." in a_mandar:
                        if a_mandar.endswith("."):
                            a_mandar = a_mandar.strip(".") 
                            numero_roles = a_mandar.split('.')
                else:
                    raise Exception
                lista_roles = [roles[int(x)] for x in numero_roles]
                site_contact = contacts.loc[(contacts["Site"] == sitio) & (contacts["Role"].isin(lista_roles))]
                email_list = '; '.join(site_contact["E-Mail"].to_list())
                break
            except:
                print('Error! (*cries*). Este mensaje es un error general. Revisa bien que escribiste bien los roles.\nIntentalo de nuevo.')
        pyperclip.copy(email_list)    
        if input(mensaje_final) != '':
            exit()

    else:
        print('No encontre el archivo "contactos.xlsx". Por favor revisa que esta presente en la carpeta donde esta este archivo.')
