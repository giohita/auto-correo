import win32com.client as client
outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)

message.SentOnBehalfOfName = 'prueba269@oulook.es'
message.Subject = "¿Puedo inscribir cursos de idiomas mientras hago mi trabajo de graduación?"
message.Body = '''Saludos,\n

Es importante destacar que, si tiene la intención de inscribirse en la materia de opción de
grado, es fundamental cumplir con los requisitos del idioma con anticipación. Para ello, deberá
obtener la certificación requerida un mes antes de la fecha programada para el examen,
teniendo en cuenta que deberá presentar dicha certificación en el momento de la inscripción.\n

La certificación del idioma es un requisito indispensable para garantizar que usted esté
preparado/a para abordar los desafíos académicos y profesionales que implica su opción de
grado. Por lo tanto, le insto a que planifique con suficiente antelación y tome las medidas
necesarias para obtener esta certificación dentro del plazo establecido.\n

Recuerde que estamos aquí para brindarle apoyo y orientación en este proceso. Si tiene alguna
pregunta o necesita más información sobre los requisitos específicos del idioma o cualquier
otro aspecto relacionado con su opción de grado, no dude en ponerse en contacto con
nosotros.\n

Esperamos con interés recibir su inscripción y acompañarle en este importante paso hacia la
culminación de su carrera académica.\n

Quedo a su disposición para cualquier consulta adicional que pueda surgir.\n

Saludos cordiales,

'''

message.Display()