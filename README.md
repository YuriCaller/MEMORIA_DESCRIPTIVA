Plugin: Memoria Descriptiva
===================================

Este plugin genera una memoria descriptiva en formato Word a partir de tres capas de QGIS:
- Una capa de polígonos para obtener datos generales y área
- Una capa de puntos para obtener coordenadas
- Una capa de líneas para obtener datos perimétricos

Nuevas funcionalidades:
----------------------
- Detección automática de colindantes
- Procesamiento mejorado de coordenadas
- Interfaz completa con todas las secciones de la memoria descriptiva
- Generación de documento Word con formato profesional
- Opciones avanzadas de personalización

Requisitos:
-----------
- QGIS 3.0 o superior
- Python 3.x
- Biblioteca python-docx (puede instalarse con pip: pip install python-docx)

Instalación:
-----------
1. Descomprima el archivo zip en la carpeta de plugins de QGIS:
   - Windows: C:\Users\{username}\AppData\Roaming\QGIS\QGIS3\profiles\default\python\plugins
   - Linux: ~/.local/share/QGIS/QGIS3/profiles/default/python/plugins
   - macOS: ~/Library/Application Support/QGIS/QGIS3/profiles/default/python/plugins

2. Abra QGIS y active el plugin en el Administrador de Complementos:
   - Menú Complementos > Administrar e instalar complementos
   - Busque "Memoria Descriptiva" y active la casilla correspondiente

Uso:
----
1. Haga clic en el icono del plugin en la barra de herramientas o acceda desde el menú Vector > Memoria Descriptiva
2. Complete los datos del solicitante (nombre y DNI)
3. Ingrese la información de ubicación (sector, zona, distrito, provincia, departamento)
4. Seleccione las tres capas requeridas (polígonos, puntos y líneas)
5. Especifique la ubicación del archivo de salida (.docx)
6. Configure las opciones avanzadas según sus necesidades:
   - Detección automática de colindantes
   - Personalización del texto de generalidades
   - Configuración de la información técnica del mapa
7. Haga clic en "Generar Memoria" para crear el documento

Contacto:
---------
Para soporte o consultas, contacte a: usuario@example.com
