import pandas as pd
import pyautogui as pg
import webbrowser as web
import time 

data = pd.read_excel("Book1.xlsx")

for i in range(len(data)):
    celular = str(data.loc[i, 'CELULAR'])
    mensaje = "Certificación en varias áreas de competencias técnicas laborales y calificaciones técnicas. Algunas características clave son: - Certificación y calificaciones directas en un plazo de 24 horas, con documentos físicos registrados en SENESCYT (Secretaría de Educación Superior, Ciencia, Tecnología e Innovación) y el Ministerio de Trabajo. - Clases pregrabadas con evaluación teórica y práctica al final, tras la cual se emiten los documentos. - Capacitación en línea disponible. F. Cárdenas es el asesor, quien está disponible para resolver cualquier duda o inquietud."

    # Abriendo la URL en el navegador predeterminado
    web.open("https://web.whatsapp.com/send?phone=" + celular + "&text=" + mensaje)

    time.sleep(15)  # Espera 15 segundos para cargar la página de WhatsApp

    # Moviendo el cursor a la posición del botón de enviar mensaje y clic
    pg.click(1048, 1000)  
    time.sleep(1)
    pg.press('enter')  

    time.sleep(1.5)
    # Cerrando la pestaña del navegador
    pg.hotkey('ctrl', 'w')  
    time.sleep(1)
