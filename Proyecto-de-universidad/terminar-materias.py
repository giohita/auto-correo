import win32com.client as client
outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)


message.Subject = "¿Qué sigue cuando termine las materias de mi maestría?"
message.Body = '''Saludos,\n

Primero es importante que procedas con la matrícula de la materia de trabajo de graduación
según lo establecido por las políticas académicas. Una vez completada esta etapa, deberás
llenar las tarjetas del trabajo de graduación de acuerdo con las pautas y formatos
proporcionados por la institución.\n

Posteriormente, te solicitamos que procedas a solicitar la revisión de tu expediente. Este paso
es crucial para asegurar que todos los requisitos académicos han sido cumplidos de manera
satisfactoria. Una vez tu expediente haya sido revisado, recibirás una notificación por correo
electrónico.\n

Es importante destacar que, una vez se haya completado la revisión de tu expediente y se haya
confirmado tu elegibilidad para la graduación, recibirás un correo electrónico adicional. En este
correo se incluirá información sobre los siguientes pasos a seguir, que incluyen la solicitud del
escribano, la solicitud de paz y salvo de egresado, y el pago de una tarifa administrativa de 396
dólares.\n

Este pago debe realizarse antes de la entrega del proyecto de opción de grado. Te
recomendamos estar atento/a a tu bandeja de entrada y revisar regularmente tu correo
electrónico para cualquier actualización o instrucción adicional.\n

Por favor, no dudes en comunicarte conmigo si tienes alguna pregunta o necesitas asistencia
adicional en este proceso. Estoy aquí para ayudarte en todo lo que necesites.\n

¡Felicitaciones por llegar a esta etapa crucial en tu camino académico!
'''

message.Display()