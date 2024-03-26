import win32com.client as client
outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)

message.SentOnBehalfOfName = 'prueba269@oulook.es'
message.Subject = "¿Qué es un examen escrito y sustentado?"
message.Body = '''Saludos,\n

Me complace proporcionarle información sobre el examen escrito y sustentado, también
conocido como examen de grado, que es una evaluación final fundamental para obtener su tan
deseado título de grado. Este examen representa una etapa crucial en su camino académico y
profesional, y es importante entender su naturaleza y propósito.\n

El examen escrito y sustentado es una prueba rigurosa y detallada diseñada para evaluar sus
conocimientos, habilidades y comprensión en su campo de estudio. Esta evaluación final es
una oportunidad para demostrar todo lo que ha aprendido durante su tiempo en la universidad y
para mostrar su capacidad para aplicar estos conocimientos de manera efectiva.\n

Durante el examen, se espera que demuestre una comprensión sólida de los conceptos clave,
así como la capacidad de analizar y sintetizar información de manera crítica. Además, puede
incluir una presentación oral donde pueda defender y sustentar sus respuestas, lo que añade
un componente de comunicación y argumentación a la evaluación.\n

Es fundamental que se prepare adecuadamente para este examen, revisando exhaustivamente
los materiales de estudio relevantes y practicando la aplicación de sus conocimientos en
diferentes contextos. Además, no dude en consultar con sus profesores o asesores
académicos para recibir orientación adicional sobre cómo prepararse de manera efectiva para
el examen.\n

Recuerde que este examen es una oportunidad para demostrar su capacidad y dedicación, y
estoy seguro/a de que lo abordará con el mismo compromiso y determinación que ha mostrado
durante su trayectoria universitaria.\n

Si tiene alguna pregunta o necesita más información sobre el examen de grado, no dude en
ponerse en contacto conmigo. Estoy aquí para ayudarle en cualquier aspecto relacionado con
su proceso de graduación.\n

Le deseamos mucho éxito en su preparación para el examen y en todos sus futuros proyectos
académicos y profesionales.
'''

message.Display()