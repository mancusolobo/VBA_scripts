# Administrador de sesiones de SAP

Teníamos un problema al momento de ejecutar scripts en SAP y era que los mismos cerraban SAP para luego iniciarlo y ejecutar el script correctamente o lo ejecutaban sobre la sesion sin posibilidad de abrir una nueva sesion si disponías de limite o tomar alguna sesión libre de transacciones activas.
Con ello elaboré este módulo el cual toma cierta lógica al momento de iniciar o tomar SAP.
La logica que sigue la detallo:
- Verifica que el proceso esté iniciado o sino lo iniciará.
- Verifica si hay conexiones activas y busca la que detallamos o sino la creará.
    - En caso de encontrarla buscará una sesion libre de transacciones.
    - Si no encuentra una sesión libre verificará el limité de sesiones.
        - Si tenemos limite creará una sesion nueva.
        - Si no tenemos límite, tomará la ultima sesión que no esté ocupada.

Con ello devuelve un objeto de tipo *SAPFEWSELib.GuiSession* para ejecutar las transaccionesque se necesiten.
Este modulo utiliza la librería *sapfewse.ocx* de SAP.

En caso de algún error respecto a las referencias en el libro.
Tienen que ir al menú de programador.

Pestaña "Herramientas" > "Referencias" > "Examinar"

Pegar las siguientes direcciones en donde dice nombre y aceptar. (Puede variar de una máquina a otra depende donde tengan SAP).

_"C:\Program Files (x86)\SAP\FrontEnd\SapGui\sapfewse.ocx"
"C:\Program Files\SAP\FrontEnd\SAPgui\sapfewse.ocx"_

Espero les sea de utilidad tanto como a mí. Saludos!