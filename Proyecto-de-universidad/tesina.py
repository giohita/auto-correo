import win32com.client as client
outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)

message.Subject = "¿Qué es una tesina?"
message.Body = '''Saludos,\n

La tesina se define como un subtipo de tesis que requiere a los estudiantes un primer
acercamiento a un proyecto de investigación. Este tipo de documento puede adoptar la forma
de un trabajo documental similar a las monografías o ensayos, centrándose en un tema
específico de estudio.\n

Aunque la tesina suele ser menos exigente que una tesis en términos de profundidad y
extensión, sigue siendo importante destacar que debe cumplir con ciertas normativas y
requerimientos académicos. Esto incluye la adecuada utilización de citas, referencias
bibliográficas y metodologías apropiadas para garantizar la calidad y validez del trabajo
realizado.\n

Espero que esta información le sea útil a medida que considere las opciones disponibles para
su proyecto de grado. Si tiene alguna pregunta adicional o necesita más orientación sobre este
tema, no dude en ponerse en contacto conmigo.\n

Quedo a su disposición para cualquier consulta que pueda surgir.\n

Saludos cordiales,
'''

message.Display()