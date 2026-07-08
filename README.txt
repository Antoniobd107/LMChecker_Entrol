================================================================
 COMPROBADOR DE LM - ENTROL (Oficina Tecnica)
================================================================

CONTENIDO DE LA CARPETA
------------------------------------------------------------
- lm_reglas.py        Modulo con la logica de negocio (no se ejecuta
                       directamente).
- lm_checker_qt.py     Aplicacion de escritorio (PyQt6) que usa el
                       archivo anterior. Es el ejecutable final
                       para el usuario.
- logo_entrol.png       Logo de la app (opcional).


DESCRIPCION DE LOS ARCHIVOS PRINCIPALES
------------------------------------------------------------

lm_reglas.py
    Modulo central que concentra toda la logica del comprobador:
    lectura del catalogo maestro (.ods), lectura del LM (.odt),
    deteccion de piezas obsoletas/no encontradas/huecos en el LM,
    aplicacion de las reglas especiales del departamento (sustitu-
    ciones, correcciones de grupo, condiciones por pais/debriefing/
    FLIR/1000x, etc.) y generacion del informe final (.xlsx/.csv).
    No tiene interfaz propia; es importado tanto por lm_checker_qt.py.
    Cualquier cambio en una norma del departamento se realiza editando 
    el bloque de constantes señalado al inicio de este archivo, sin tocar
    el resto de la logica.

lm_checker_qt.py
    Interfaz grafica de escritorio que es la aplicacion final que
    utiliza el tecnico. Permite seleccionar el catalogo, el LM y 
    la carpeta de salida, marcar la configuracion del simulador (pais, debriefing, FLIR, 1000x,
    verificacion de ultimas referencias) y lanzar el analisis. Las
    tareas largas (carga de catalogo, analisis) se ejecutan en
    segundo plano mediante hilos, sin congelar la ventana. Toda la
    logica de negocio se delega en lm_reglas.py; este archivo solo
    gestiona la presentacion y la interaccion con el usuario.


COMPILACION A .EXE (PyInstaller)
------------------------------------------------------------
El ejecutable a distribuir es el de lm_checker_qt.py (incluye a
lm_reglas.py como dependencia, no se compila por separado).

1. Requisitos previos (entorno con Python 3.11+ recomendado):

   pip install pyinstaller pandas lxml odfpy openpyxl PyQt6 PyQt6-Fluent-Widgets

2. IMPORTANTE - antes de compilar:
   Revisar en lm_checker_qt.py la constante RUTA_CATALOGO_DEFECTO
   y confirmar que apunta a la ruta de red vigente del catalogo
   maestro compartido.

3. Comando de compilacion, ejecutado desde la carpeta del proyecto
   (con lm_checker_qt.py, lm_reglas.py y logo_entrol.png juntos):

python -m PyInstaller --onedir --windowed --name "ComprobadorLM" --icon=logo_entrol.ico --add-data "logo_entrol.png;." lm_checker_qt.py

4. El ejecutable resultante queda en la carpeta dist\ComprobadorLM.exe.
   Es el unico archivo que hay que distribuir a los tecnicos.

5. La configuracion (ultima carpeta/catalogo usados) y el registro
   de errores se generan automaticamente en tiempo de ejecucion en:
   %LOCALAPPDATA%\ComprobadorLM\ (no requieren empaquetarse).
