import pandas as pd
import os
import smtplib
from email.mime.text import MIMEText
from datetime import datetime, timedelta

#Nombre del archivo Excel
nombre_archivo = "inventario_radios.xlsx"

# Verificar si el archivo existe para cargar datos previos
if os.path.exists(nombre_archivo):
    df = pd.read_excel(nombre_archivo)
    inventario_radios = df.to_dict(orient="records")
else:
    inventario_radios = []

# Configuración del correo
EMAIL_ORIGEN = "tu_correo@gmail.com"  # Cambia esto
CONTRASENA = "tu_contraseña"  # Cambia esto (preferiblemente usa una App Password si tienes 2FA)
EMAIL_DESTINO = "destinatario@gmail.com"  # Cambia esto

# Función para enviar reportes por correo
def enviar_reporte_por_correo(asunto, contenido):
    try:
        msg = MIMEText(contenido, "plain", "utf-8")
        msg["Subject"] = asunto
        msg["From"] = EMAIL_ORIGEN
        msg["To"] = EMAIL_DESTINO

        # Establecer conexión SMTP
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as servidor:
            servidor.login(EMAIL_ORIGEN, CONTRASENA)
            servidor.sendmail(EMAIL_ORIGEN, EMAIL_DESTINO, msg.as_string())
        
        print(f"📧 Reporte enviado a {EMAIL_DESTINO} exitosamente.\n")
    except Exception as e:
        print(f"❌ Error al enviar el correo: {e}\n")

# Función para alertar sobre mantenimiento
def alertar_mantenimiento():
    print("\n--- Verificando Alertas de Mantenimiento ---")
    radios_alerta = []
    hoy = datetime.now()

    for radio in inventario_radios:
        if radio["Fecha_Ultima_Revision"]:
            fecha_revision = datetime.strptime(radio["Fecha_Ultima_Revision"], "%Y-%m-%d")
            if hoy - fecha_revision > timedelta(days=180):
                radios_alerta.append(f"{radio['ID_Serie']} | {radio['Marca_Modelo']} | Última Revisión: {radio['Fecha_Ultima_Revision']}")

    if radios_alerta:
        contenido_alerta = "\n".join(radios_alerta)
        print("⚠️ Radios que requieren mantenimiento:")
        print(contenido_alerta)
        enviar_reporte_por_correo("⚠️ Alertas de Mantenimiento de Radios", contenido_alerta)
    else:
        print("✅ No hay radios que requieran mantenimiento.\n")

# Función para agregar un nuevo radio
def agregar_radio():
    print("\n--- Agregar Nuevo Radio ---")
    
    id_serie = input("ID/No. de Serie: ")
    marca_modelo = input("Marca/Modelo: ")
    frecuencia_canales = input("Frecuencia/Canales: ")
    estado = input("Estado (Disponible, Rentado, En mantenimiento): ")
    fecha_ingreso = input("Fecha de Ingreso (YYYY-MM-DD): ")
    cliente_ubicacion = input("Cliente/Ubicación Actual: ")
    fecha_revision = input("Fecha de Última Revisión (YYYY-MM-DD): ")
    accesorios = input("Accesorios Incluidos: ")

    nuevo_radio = {
        "ID_Serie": id_serie,
        "Marca_Modelo": marca_modelo,
        "Frecuencia_Canales": frecuencia_canales,
        "Estado": estado,
        "Fecha_Ingreso": fecha_ingreso,
        "Cliente_Ubicacion": cliente_ubicacion,
        "Fecha_Ultima_Revision": fecha_revision,
        "Accesorios_Incluidos": accesorios
    }

    inventario_radios.append(nuevo_radio)
    print(f"\n✅ Radio {id_serie} agregado exitosamente.\n")

# Función para modificar un radio existente
def modificar_radio():
    print("\n--- Modificar Radio Existente ---")
    id_serie = input("Ingrese el ID/No. de Serie del radio a modificar: ")
    
    encontrado = False
    for radio in inventario_radios:
        if radio["ID_Serie"] == id_serie:
            print(f"🎯 Modificando {radio['Marca_Modelo']} ({id_serie})")
            radio["Marca_Modelo"] = input(f"Marca/Modelo [{radio['Marca_Modelo']}]: ") or radio["Marca_Modelo"]
            radio["Frecuencia_Canales"] = input(f"Frecuencia/Canales [{radio['Frecuencia_Canales']}]: ") or radio["Frecuencia_Canales"]
            radio["Estado"] = input(f"Estado [{radio['Estado']}]: ") or radio["Estado"]
            radio["Fecha_Ingreso"] = input(f"Fecha de Ingreso [{radio['Fecha_Ingreso']}]: ") or radio["Fecha_Ingreso"]
            radio["Cliente_Ubicacion"] = input(f"Cliente/Ubicación [{radio['Cliente_Ubicacion']}]: ") or radio["Cliente_Ubicacion"]
            radio["Fecha_Ultima_Revision"] = input(f"Fecha de Última Revisión [{radio['Fecha_Ultima_Revision']}]: ") or radio["Fecha_Ultima_Revision"]
            radio["Accesorios_Incluidos"] = input(f"Accesorios Incluidos [{radio['Accesorios_Incluidos']}]: ") or radio["Accesorios_Incluidos"]
            
            print(f"✅ Radio {id_serie} modificado exitosamente.\n")
            encontrado = True
            break
    
    if not encontrado:
        print("❌ Radio no encontrado.")

# Función para exportar el inventario a Excel
def exportar_excel():
    df = pd.DataFrame(inventario_radios)
    df.to_excel(nombre_archivo, index=False)
    print(f"📊 Inventario exportado a '{nombre_archivo}' exitosamente.\n")

# Función para generar reportes personalizados
def generar_reporte():
    print("\n--- Generar Reporte ---")
    print("1. Mostrar Radios Disponibles")
    print("2. Mostrar Radios Rentados")
    print("3. Mostrar Radios en Mantenimiento")
    print("4. Mostrar Todos los Radios")
    
    opcion_reporte = input("Selecciona una opción (1-4): ")
    
    if opcion_reporte == "1":
        reporte = [radio for radio in inventario_radios if radio["Estado"].lower() == "disponible"]
    elif opcion_reporte == "2":
        reporte = [radio for radio in inventario_radios if radio["Estado"].lower() == "rentado"]
    elif opcion_reporte == "3":
        reporte = [radio for radio in inventario_radios if radio["Estado"].lower() == "en mantenimiento"]
    elif opcion_reporte == "4":
        reporte = inventario_radios
    else:
        print("❌ Opción no válida. Volviendo al menú principal.\n")
        return

    if reporte:
        contenido_reporte = "\n".join([f"{radio['ID_Serie']} | {radio['Marca_Modelo']} | {radio['Estado']}" for radio in reporte])
        print("\n--- Reporte Generado ---")
        print(contenido_reporte)
        enviar_reporte_por_correo("📡 Reporte de Inventario de Radios", contenido_reporte)
    else:
        print("⚠️ No hay datos para mostrar.")

# Menú principal
while True:
    print("\n📡 GESTIÓN DE INVENTARIO DE RADIOS 📡")
    print("1. Agregar Nuevo Radio")
    print("2. Modificar Radio")
    print("3. Ver Alertas de Mantenimiento")
    print("4. Exportar a Excel")
    print("5. Generar Reporte")
    print("6. Salir")
    
    opcion = input("Selecciona una opción (1-6): ")

    if opcion == "1":
        agregar_radio()
    elif opcion == "2":
        modificar_radio()
    elif opcion == "3":
        alertar_mantenimiento()
    elif opcion == "4":
        exportar_excel()
    elif opcion == "5":
        generar_reporte()
    elif opcion == "6":
        print("👋 Saliendo del programa. ¡Hasta luego!")
        break
    else:
        print("❌ Opción no válida. Inténtalo de nuevo.\n")
