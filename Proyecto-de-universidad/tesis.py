import win32com.client as client
outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)

message.Subject = "¿Qué es una tesis?"
message.Body = '''Saludos,\n

Un trabajo de investigación, por lo general, es un estudio monográfico o investigativo que
involucra una disertación y la verificación de hipótesis previamente establecidas. Su objetivo
principal es demostrar la capacidad analítica del investigador y su habilidad para aplicar
procedimientos de investigación de manera efectiva.\n

Este tipo de trabajo académico implica un proceso riguroso que puede incluir la revisión de
literatura existente, la recolección y análisis de datos pertinentes, así como la formulación y
prueba de hipótesis específicas. A través de este proceso, el investigador busca aportar nuevos
conocimientos o profundizar en la comprensión de un tema particular dentro de su campo de
estudio.\n

Además, un trabajo de investigación también puede requerir la presentación de conclusiones
fundamentadas en evidencia sólida, así como la discusión de las implicaciones de los hallazgos
para el campo en cuestión.\n

En resumen, un trabajo de investigación es una empresa intelectualmente desafiante que
ofrece la oportunidad de desarrollar habilidades analíticas, de investigación y de comunicación,
mientras se contribuye al avance del conocimiento en un área específica.
Si tiene alguna pregunta adicional sobre este tema o necesita más información, no dude en
comunicarse conmigo.\n

Quedo a su disposición para cualquier consulta que pueda surgir.
'''

message.Display()