import re
import pandas as pd
import os
from datetime import datetime

# Print the current working directory
print("Current working directory:", os.getcwd())

# Definir las fechas de inicio y fin
start_date = "01/01/2025"
end_date = "31/01/2025"

# Convertir las fechas a objetos datetime
start_date = datetime.strptime(start_date, "%d/%m/%Y")
end_date = datetime.strptime(end_date, "%d/%m/%Y")

# Leer el archivo de texto
with open('expenses/chat.txt', 'r', encoding='utf-8') as file:
    text = file.read()

# Patrón regex para extraer fecha, hora, nombre y mensaje
pattern = r"(\d{1,2}/\d{1,2}/\d{2,4}), (\d{1,2}:\d{2}\s?[APMapm]*) - ([^:]+): (.*)"

# Extraer datos usando regex
matches = re.findall(pattern, text)

# Patrón regex para extraer el número del mensaje
number_pattern = r"(\d+\.?\d*)"

# Crear DataFrame
data = []
for m in matches:
    date_str = m[0]
    try:
        message_date = datetime.strptime(date_str, "%m/%d/%Y")
    except ValueError:
        try:
            message_date = datetime.strptime(date_str, "%m/%d/%y")
        except ValueError:
            print(f"Error parsing date: {date_str}")
            continue
    
    # Filtrar mensajes por rango de fechas
    if start_date <= message_date <= end_date:
        message = m[3]
        
        # Verificar si el mensaje contiene una imagen
        if "IMG-" in message or "STK-" in message or "PTT-" in message or "Comprobante_" in message:
            number = 0
        else:
            number_match = re.search(number_pattern, message)
            number = number_match.group(0) if number_match else None
        
        data.append({"Fecha": m[0], "Hora": m[1], "Nombre": m[2], "Mensaje": message, "Número": number})

df = pd.DataFrame(data)

# Guardar en Excel
df.to_excel('expenses/chat_whatsapp.xlsx', index=False)
