import csv
import sys
import os
import subprocess

# Función para instalar un paquete si no está instalado
def install_package(package):
    try:
        __import__(package)
        print(f"'{package}' ya está instalado.")
    except ImportError:
        print(f"'{package}' no encontrado. Instalando...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        print(f"'{package}' instalado correctamente.")

# Instalar pywin32 si es necesario
install_package("pywin32")

# Verificar la instalación de win32com
try:
    import win32com.client
    print("win32com.client está instalado correctamente.")
except ImportError:
    print("Error: win32com.client no se pudo importar después de la instalación.")
    sys.exit(1)

# Ejecutar pywin32_postinstall (Opcional, si hay problemas con la instalación)
try:
    subprocess.run([sys.executable, "-m", "pywin32_postinstall"], check=True)
    print("pywin32_postinstall ejecutado correctamente.")
except Exception as e:
    print(f"Error ejecutando pywin32_postinstall: {e}")

# Verificar si se ha proporcionado el archivo de datos como argumento
if len(sys.argv) < 2:
    print("Error: No se proporcionó el archivo de datos como argumento.")
    sys.exit(1)

archivo_txt = sys.argv[1]

# Verificar si el archivo existe
if not os.path.exists(archivo_txt):
    print(f"Error: No se encontró el archivo {archivo_txt}.")
    sys.exit(1)

# Conectar con Outlook
try:
    outlook = win32com.client.Dispatch("Outlook.Application")
except Exception as e:
    print(f"Error al conectar con Outlook: {e}")
    sys.exit(1)

# Leer el archivo de texto usando csv.reader para manejar comas dentro de comillas
with open(archivo_txt, "r", encoding="utf-8") as file:
    reader = csv.reader(file, quotechar='"', delimiter=',')
    
    for index, row in enumerate(reader, start=2):
        if len(row) < 5:
            print(f"Fila {index}: Datos incompletos, se omite el envío.")
            continue

        destinatario, cc, asunto, mensaje, adjuntos = row

        # Limpiar espacios en blanco
        destinatario = destinatario.strip()
        cc = cc.strip()
        asunto = asunto.strip()
        mensaje = mensaje.strip()
        adjuntos = adjuntos.strip()

        # Omitir si el destinatario contiene "NO EXISTE CORREO PRINCIPAL"
        if "NO EXISTE CORREO PRINCIPAL" in destinatario.upper():
            print(f"Fila {index}: No se enviará el correo. Asunto: '{asunto}'")
            continue

        try:
            # Crear el mensaje en Outlook
            mail = outlook.CreateItem(0)  # 0 = Email
            mail.To = destinatario
            print(f'destinatario: {destinatario}')
            mail.Subject = asunto
            mail.CC = cc
            mail.Body = mensaje

            # Agregar adjuntos si existen
            if adjuntos and adjuntos != '""':  # Verifica si el campo adjuntos está vacío
                adjuntos_lista = adjuntos.split(";")  # Separar múltiples archivos por punto y coma (;)
                for adjunto in adjuntos_lista:
                    adjunto = adjunto.strip()
                    if os.path.exists(adjunto):  # Verificar si el archivo existe antes de adjuntar
                        mail.Attachments.Add(adjunto)
                        print(f"Adjunto agregado: {adjunto}")
                    else:
                        print(f"Advertencia: No se encontró el archivo adjunto '{adjunto}', se omite.")

            # Enviar el correo
            mail.Send()
            print(f"Correo enviado a {destinatario}")

        except Exception as e:
            print(f"Error al enviar correo a {destinatario}: {e}")

print("Proceso de envío de correos completado.")
