import pathlib
from tkinter import messagebox
import win32com.client as client
from tkinter import *
import customtkinter

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme('dark-blue')

root = customtkinter.CTk()
root.title('Tkinter.comm - Custom Tkinter Buttons')
root.geometry('600x350')


def cantidad_de_opciones():
    try:
        outlook = client.Dispatch("Outlook.Application")
        opciones_path = pathlib.Path("opciones.pdf")
        opciones_absolute = str(opciones_path.absolute())
        message = outlook.CreateItem(0) 
        message.SentOnBehalfOfName = 'prueba269@outlook.es'
        message.Attachments.Add(opciones_absolute)
        message.Subject = "¿Qué sigue cuando termine las materias de mi maestría?"
        message.Body = '''Saludos,\n

Me complace informarle sobre las opciones de grado disponibles según cada programa de
maestría en nuestra institución. Es fundamental que esté al tanto de estas alternativas, ya que
su elección determinará el enfoque final de sus estudios y el tipo de proyecto que llevará a
cabo para culminar su programa de maestría.\n

Le insto a revisar detenidamente estas opciones y considerar cuál se ajusta mejor a sus
intereses académicos y profesionales. Recuerde que su elección también debe estar en
consonancia con los requisitos específicos del programa y las expectativas de su departamento
académico.\n

Si tiene alguna pregunta o necesita orientación adicional para tomar una decisión informada, no
dude en ponerse en contacto con su asesor académico. Estamos aquí para ayudarlo/a en cada
paso del camino hacia la culminación exitosa de su programa de maestría.\n
Quedo a su disposición para cualquier consulta adicional que pueda surgir.
        '''
        message.Display()
        message.Save()
    except Exception as e:
        messagebox.showerror("Error", f"Un error a ocurrido!: {e}")

def examen_escrito():
    try:
        outlook = client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)

        message.SentOnBehalfOfName = 'prueba269@outlook.es'
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
        message.Send()
    except Exception as e:
        messagebox.showerror("Error", f"Un error a ocurrido!: {e}")

def idiomas():
    try:
        outlook = client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)

        message.SentOnBehalfOfName = 'prueba269@outlook.es'
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
    except Exception as e:
        messagebox.showerror("Error", f"Un error a ocurrido!: {e}")

def opciones_grado():
    try:
        outlook = client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)

        message.SentOnBehalfOfName = 'prueba269@outlook.es'
        message.Subject = "¿Cuáles opciones de grado tengo?"
        message.Body = '''Saludos,\n

Las opciones para su grado incluyen:\n
    
    - Tesis: Esta opción implica la investigación y redacción de un documento original sobre
    un tema específico dentro de su campo de estudio. La tesis le brinda la oportunidad de
    explorar un tema en profundidad y contribuir al conocimiento existente en su área.
    
    - Trabajo de Campo: Si prefiere una experiencia más práctica, el trabajo de campo
    puede ser la opción adecuada. Esto implica realizar investigaciones en el terreno,
    recolectar datos y analizarlos para abordar un problema o una pregunta de
    investigación.\n
    
    - Tesina: Similar a una tesis, pero generalmente más corta en longitud y alcance. Una
    tesina le permite realizar una investigación más enfocada en un área particular de
    interés, pero con menos extensión que una tesis tradicional.\n

    - Examen Escrito y Sustentado: Esta opción implica la realización de un examen que
    evaluará su comprensión y dominio de los conceptos clave en su área de estudio. Este
    examen puede ser acompañado por una presentación oral para defender su perspectiva
    y respuestas.\n

Cada una de estas opciones tiene sus propias ventajas y requisitos, por lo que le animo a
considerar cuidadosamente cuál se ajusta mejor a sus intereses, habilidades y objetivos
profesionales.\n

Además, recuerde que la elección de su opción de grado también dependerá de la resolución
de su plan de estudios. Por lo tanto, es importante que se comunique con su asesor académico
para discutir cómo su elección se alinea con los requisitos de su programa y cualquier otro
detalle específico que pueda ser relevante.\n

Si tiene alguna pregunta o necesita más orientación sobre estas opciones, no dude en ponerse
en contacto conmigo o con su asesor académico. Estamos aquí para apoyarlo/a en este
proceso crucial para su desarrollo académico y profesional.\n

Quedo a su disposición para cualquier consulta adicional.\n

Saludos cordiales,
'''

        message.Display()
    except Exception as e:
        messagebox.showerror("Error", f"Un error a ocurrido!: {e}")

def terminar_materias():
    try:
        outlook = client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)

        message.SentOnBehalfOfName = 'prueba269@outlook.es'
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
    except Exception as e:
        messagebox.showerror("Error", f"Un error a ocurrido!: {e}")

def tesina():
    try:
        outlook = client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)

        message.SentOnBehalfOfName = 'prueba269@outlook.es'
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
    except Exception as e:
        messagebox.showerror("Error", f"Un error a ocurrido!: {e}")

def tesis_tesina():
    try:
        outlook = client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)

        message.SentOnBehalfOfName = 'prueba269@outlook.es'
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
    except Exception as e:
        messagebox.showerror("Error", f"Un error a ocurrido!: {e}")

def tesis():
    try:
        outlook = client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)

        message.SentOnBehalfOfName = 'prueba269@outlook.es'
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
    except Exception as e:
        messagebox.showerror("Error", f"Un error a ocurrido!: {e}")

def tiempo_matricula():
    try:
        outlook = client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)

        message.SentOnBehalfOfName = 'prueba269@outlook.es'
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
    except Exception as e:
        messagebox.showerror("Error", f"Un error a ocurrido!: {e}")

def trabajo_campo():
    try:
        outlook = client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)

        message.SentOnBehalfOfName = 'prueba269@outlook.es'
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
    except Exception as e:
        messagebox.showerror("Error", f"Un error a ocurrido!: {e}")



my_button = customtkinter.CTkButton(root, text='Cantidad de opciones', command=cantidad_de_opciones)
my_button.pack(pady=5)

my_button = customtkinter.CTkButton(root, text='Examen escrito', command=examen_escrito)
my_button.pack(pady=5)

my_button = customtkinter.CTkButton(root, text='Idiomas', command=idiomas)
my_button.pack(pady=5)

my_button = customtkinter.CTkButton(root, text='Opciones de grado', command=opciones_grado)
my_button.pack(pady=5)

my_button = customtkinter.CTkButton(root, text='Tesina', command=tesina)
my_button.pack(pady=5)

my_button = customtkinter.CTkButton(root, text='Tesis y tesina', command=tesis_tesina)
my_button.pack(pady=5)

my_button = customtkinter.CTkButton(root, text='Tesis', command=tesis)
my_button.pack(pady=5)

my_button = customtkinter.CTkButton(root, text='Tiempo de matricula', command=tiempo_matricula)
my_button.pack(pady=5)

my_button = customtkinter.CTkButton(root, text='Trabajo de campo', command=trabajo_campo)
my_button.pack(pady=5)


root.mainloop()
