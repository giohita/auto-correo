import win32com.client as client
outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)

message.SentOnBehalfOfName = 'prueba269@oulook.es'
message.Subject = "¿Qué es un trabajo de campo?"
message.Body = '''Saludos,\n

El trabajo de campo en investigación es una actividad esencial en el proceso de recopilación de
datos para el desarrollo de un texto académico. En este método, el autor se sitúa físicamente
en el lugar o contexto donde se encuentran los datos necesarios para su investigación. Durante
este proceso, se recopila información, se realizan observaciones y se estudia el área con el
objetivo de generar una hipótesis.\n

Podemos definir el trabajo de campo como un método de observación en el cual el investigador
no opera en entornos semicontrolados o controlados, como podría ser un laboratorio o un aula.
Por el contrario, el trabajo de campo se lleva a cabo en la naturaleza, en el ambiente real
donde ocurren los fenómenos que se están investigando.\n

Este enfoque permite una inmersión más profunda en el contexto de estudio, lo que facilita una
comprensión más completa de los procesos y fenómenos investigados.\n

Espero que esta descripción le sea útil para comprender la importancia y el alcance del trabajo
de campo en investigación. Si tiene alguna pregunta adicional o necesita más información
sobre este tema, no dude en ponerse en contacto conmigo.\n

Quedo a su disposición para cualquier consulta que pueda surgir.\n

Saludos cordiales,
'''

message.Display()