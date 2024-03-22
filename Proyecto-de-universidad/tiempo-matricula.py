import win32com.client as client
outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)


message.Subject = "¿Hasta cuándo me puedo matricular?"
message.Body = '''Saludos,\n

Me complace informarle sobre las fechas de inscripción para los programas de maestría para el
año 2024. Estas fechas son cruciales para aquellos interesados en matricularse y avanzar en
su desarrollo académico. A continuación, detallo las fechas correspondientes a cada semestre:\n

Primer Semestre:
Fecha de Inscripción: Lunes 15 de enero al sábado 16 de marzo.\n

Segundo Semestre:
Fecha de Inscripción: Lunes 10 de junio al sábado 27 de julio.\n

Es fundamental tener en cuenta estas fechas para garantizar una matrícula oportuna y evitar
contratiempos en el proceso de admisión. Le insto a que marque estas fechas en su calendario
y tome las medidas necesarias para cumplir con los plazos establecidos.\n

Si tiene alguna pregunta o necesita más información sobre los programas de maestría
disponibles, los requisitos de admisión o cualquier otro detalle relacionado, no dude en ponerse
en contacto con nosotros.\n

Estamos aquí para ayudarlo/a en cada paso del proceso de admisión.

'''

message.Display()