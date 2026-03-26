# -*- coding: utf-8 -*-
"""
Funciones para la identificación de colindantes y procesamiento de sus datos
"""

from qgis.core import QgsProject, QgsVectorLayer, QgsFeature, QgsGeometry, QgsPointXY
import math

def identificar_colindantes_completo(poligono_layer, capas_adyacentes=None):
    """
    Identifica los colindantes del polígono principal con información completa
    
    Args:
        poligono_layer: Capa de polígono principal
        capas_adyacentes: Lista de capas a revisar (opcional)
        
    Returns:
        Diccionario con información completa de colindantes por posición
    """
    # Obtener el polígono principal
    features = list(poligono_layer.getFeatures())
    if not features:
        return {
            'NORTE': {'nombre': 'Terrenos del estado', 'observacion': ''},
            'SUR': {'nombre': 'Terrenos del estado', 'observacion': ''},
            'ESTE': {'nombre': 'Terrenos del estado', 'observacion': ''},
            'OESTE': {'nombre': 'Terrenos del estado', 'observacion': ''}
        }
    
    poligono_principal = features[0]
    geometria_principal = poligono_principal.geometry()
    
    # Si no se proporcionan capas adyacentes, buscar en todas las capas del proyecto
    if capas_adyacentes is None:
        capas_adyacentes = []
        for layer in QgsProject.instance().mapLayers().values():
            if isinstance(layer, QgsVectorLayer) and layer.geometryType() == 2:  # Polígonos
                if layer.id() != poligono_layer.id():  # Evitar la capa principal
                    capas_adyacentes.append(layer)
    
    # Inicializar diccionario de colindantes con valores por defecto
    colindantes = {
        'NORTE': {'nombre': 'Terrenos del estado', 'observacion': ''},
        'SUR': {'nombre': 'Terrenos del estado', 'observacion': ''},
        'ESTE': {'nombre': 'Terrenos del estado', 'observacion': ''},
        'OESTE': {'nombre': 'Terrenos del estado', 'observacion': ''}
    }
    
    # Obtener límites del polígono principal para ayudar a determinar direcciones
    limites_principal = obtener_limites_poligono(geometria_principal)
    
    # Buscar colindantes en cada capa
    for layer in capas_adyacentes:
        for feature in layer.getFeatures():
            if feature.geometry().touches(geometria_principal):
                # Determinar la posición relativa con mayor precisión
                posicion = determinar_posicion_relativa_mejorada(
                    geometria_principal, 
                    feature.geometry(),
                    limites_principal
                )
                
                # Extraer nombre y observación del colindante
                nombre = extraer_nombre_colindante_mejorado(feature, layer)
                observacion = extraer_observacion_colindante(feature, layer)
                
                if posicion:
                    colindantes[posicion] = {
                        'nombre': nombre,
                        'observacion': observacion
                    }
    
    # Verificar si hay información del propietario en el polígono principal
    # para usarla como colindante en alguna dirección si es necesario
    nombre_propietario = extraer_nombre_colindante_mejorado(poligono_principal, poligono_layer)
    if nombre_propietario:
        # Buscar en atributos del polígono principal si hay información sobre colindantes
        for direccion in ['NORTE', 'SUR', 'ESTE', 'OESTE']:
            campo_colindante = f"colindante_{direccion.lower()}"
            idx = poligono_principal.fieldNameIndex(campo_colindante)
            if idx != -1 and poligono_principal[campo_colindante]:
                colindantes[direccion] = {
                    'nombre': poligono_principal[campo_colindante],
                    'observacion': ''
                }
    
    return colindantes

def determinar_posicion_relativa_mejorada(geom_principal, geom_colindante, limites_principal=None):
    """
    Determina con mayor precisión si un polígono colindante está al N, S, E u O del principal
    
    Args:
        geom_principal: Geometría del polígono principal
        geom_colindante: Geometría del polígono colindante
        limites_principal: Límites del polígono principal (opcional)
        
    Returns:
        String con la posición relativa ('NORTE', 'SUR', 'ESTE', 'OESTE')
    """
    # Si no se proporcionan límites, calcularlos
    if limites_principal is None:
        limites_principal = obtener_limites_poligono(geom_principal)
    
    # Obtener límites del colindante
    limites_colindante = obtener_limites_poligono(geom_colindante)
    
    # Calcular la intersección entre los polígonos
    # Usamos buffer para asegurar que toquen incluso con pequeñas imprecisiones
    buffer_principal = geom_principal.buffer(0.1, 5)
    buffer_colindante = geom_colindante.buffer(0.1, 5)
    
    interseccion = buffer_principal.intersection(buffer_colindante)
    
    if interseccion.isEmpty():
        # Si no hay intersección clara, usar centroides
        centroide_principal = geom_principal.centroid().asPoint()
        centroide_colindante = geom_colindante.centroid().asPoint()
        
        # Calcular diferencias
        dx = centroide_colindante.x() - centroide_principal.x()
        dy = centroide_colindante.y() - centroide_principal.y()
        
        # Determinar dirección predominante
        if abs(dx) > abs(dy):
            return 'ESTE' if dx > 0 else 'OESTE'
        else:
            return 'NORTE' if dy > 0 else 'SUR'
    else:
        # Hay intersección, determinar en qué borde está principalmente
        centroide_interseccion = interseccion.centroid().asPoint()
        
        # Calcular distancias a los bordes del polígono principal
        dist_norte = abs(centroide_interseccion.y() - limites_principal['norte'].y())
        dist_sur = abs(centroide_interseccion.y() - limites_principal['sur'].y())
        dist_este = abs(centroide_interseccion.x() - limites_principal['este'].x())
        dist_oeste = abs(centroide_interseccion.x() - limites_principal['oeste'].x())
        
        # Determinar el borde más cercano
        min_dist = min(dist_norte, dist_sur, dist_este, dist_oeste)
        
        if min_dist == dist_norte:
            return 'NORTE'
        elif min_dist == dist_sur:
            return 'SUR'
        elif min_dist == dist_este:
            return 'ESTE'
        else:
            return 'OESTE'

def extraer_nombre_colindante_mejorado(feature, layer):
    """
    Extrae el nombre del colindante con búsqueda mejorada en atributos
    
    Args:
        feature: Feature del colindante
        layer: Capa del colindante
        
    Returns:
        String con el nombre del colindante formateado
    """
    # Campos posibles que pueden contener nombres o partes del nombre
    campos_nombre = ['nombre', 'nom_tit', 'propietario', 'titular', 'name', 'owner', 'dueño']
    campos_apellido = ['apellido', 'apellidos', 'last_name', 'surname']
    campos_dni = ['dni', 'doc', 'documento', 'id', 'identificacion']
    
    nombre_completo = ""
    dni = ""
    
    # Buscar nombre
    for campo in campos_nombre:
        idx = feature.fieldNameIndex(campo)
        if idx != -1 and feature[campo]:
            nombre_completo = feature[campo]
            break
    
    # Si no encontró nombre, buscar apellido
    if not nombre_completo:
        for campo in campos_apellido:
            idx = feature.fieldNameIndex(campo)
            if idx != -1 and feature[campo]:
                nombre_completo = feature[campo]
                break
    
    # Buscar DNI
    for campo in campos_dni:
        idx = feature.fieldNameIndex(campo)
        if idx != -1 and feature[campo]:
            dni = feature[campo]
            break
    
    # Formatear resultado
    if nombre_completo:
        if dni:
            return f"la parcela de {nombre_completo} con DNI {dni}"
        else:
            return f"la parcela de {nombre_completo}"
    else:
        # Si no encuentra, buscar cualquier campo que pueda identificar la parcela
        for field in feature.fields():
            field_name = field.name().lower()
            if field_name not in ['id', 'fid', 'objectid', 'shape_area', 'shape_length', 'area', 'perimeter']:
                value = feature[field.name()]
                if value:
                    return f"la parcela con {field.name()}={value}"
        
        # Si no encuentra nada útil
        return "Parcela sin nombre identificado"

def extraer_observacion_colindante(feature, layer):
    """
    Extrae observaciones sobre el colindante de los atributos
    
    Args:
        feature: Feature del colindante
        layer: Capa del colindante
        
    Returns:
        String con observaciones o cadena vacía
    """
    # Campos posibles que pueden contener observaciones
    campos_obs = ['observacion', 'obs', 'descripcion', 'desc', 'notas', 'comentario', 'tipo', 'type']
    
    for campo in campos_obs:
        idx = feature.fieldNameIndex(campo)
        if idx != -1 and feature[campo]:
            return feature[campo]
    
    # Buscar campos que contengan palabras clave
    palabras_clave = ['quebrada', 'rio', 'camino', 'carretera', 'via', 'calle', 'avenida']
    
    for field in feature.fields():
        value = str(feature[field.name()]).lower()
        for palabra in palabras_clave:
            if palabra in value:
                return value.capitalize()
    
    # Si no encuentra, devolver cadena vacía
    return ""

def obtener_limites_poligono(poligono_geometry):
    """
    Obtiene los límites extremos del polígono
    
    Args:
        poligono_geometry: Geometría del polígono
        
    Returns:
        Diccionario con puntos extremos (norte, sur, este, oeste)
    """
    bbox = poligono_geometry.boundingBox()
    
    return {
        'norte': QgsPointXY(bbox.center().x(), bbox.yMaximum()),
        'sur': QgsPointXY(bbox.center().x(), bbox.yMinimum()),
        'este': QgsPointXY(bbox.xMaximum(), bbox.center().y()),
        'oeste': QgsPointXY(bbox.xMinimum(), bbox.center().y())
    }

def buscar_colindantes_por_nombre(nombre_buscar, capas_disponibles):
    """
    Busca en todas las capas disponibles features que contengan el nombre buscado
    
    Args:
        nombre_buscar: Nombre o parte del nombre a buscar
        capas_disponibles: Lista de capas donde buscar
        
    Returns:
        Lista de features que coinciden con el nombre buscado
    """
    resultados = []
    
    for layer in capas_disponibles:
        if not isinstance(layer, QgsVectorLayer):
            continue
            
        # Buscar en todos los campos de texto
        for field in layer.fields():
            if field.type() in [10, 4]:  # Tipos de campo de texto en QGIS
                # Crear expresión de filtro
                expresion = f"lower(\"{field.name()}\") LIKE '%{nombre_buscar.lower()}%'"
                
                # Buscar features que coincidan
                for feature in layer.getFeatures(expresion):
                    resultados.append({
                        'feature': feature,
                        'layer': layer,
                        'field': field.name(),
                        'value': feature[field.name()]
                    })
    
    return resultados
