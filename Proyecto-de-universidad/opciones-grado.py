import win32com.client as client
outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)


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