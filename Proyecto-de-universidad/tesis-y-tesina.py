import win32com.client as client
outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)

message.Subject = "¿Cuál es la diferencia entre tesis y tesina?"
message.Body = '''Saludos,\n

Me complace proporcionarle información sobre las principales diferencias entre la tesis y la
tesina, con el fin de ayudarle a entender mejor las características distintivas de cada una:\n

1. Enfoque y Metodología:\n

- En la tesis, el enfoque es principalmente experimental y teórico, mientras que en la tesina es
del tipo monográfico y recopilatorio.\n

- En la tesis, se requiere una hipótesis, mientras que en la tesina no es necesaria. Además,
los objetivos en la tesina deben ser limitados, mientras que en la tesis pueden ser tan amplios
como lo requiera la investigación.\n

2. Investigación:\n

- La tesis busca generar un aporte significativo a su campo a través de una investigación
profunda y exhaustiva.\n

- Por otro lado, la tesina se enfoca más en recopilar información existente y no busca
necesariamente aportar nuevos conocimientos. Por lo general, las tesinas se centran en una
revisión bibliográfica, mientras que las tesis son de carácter investigativo.\n

3. Extensión:\n

- Una tesis típicamente tiene una extensión mínima de 120 hojas y un máximo de 500.\n

- En cambio, una tesina suele tener una extensión más reducida, oscilando entre 20 y 40 o 50
páginas como máximo, según nuestra experiencia.\n

Espero que esta información sea útil para usted a medida que considere cuál de estas
opciones se adapta mejor a sus necesidades académicas y profesionales. Si tiene alguna
pregunta adicional o necesita más orientación sobre este tema, no dude en ponerse en
contacto conmigo.\n

Quedo a su disposición para cualquier consulta que pueda surgir.\n

Saludos cordiales,
'''

message.Display()