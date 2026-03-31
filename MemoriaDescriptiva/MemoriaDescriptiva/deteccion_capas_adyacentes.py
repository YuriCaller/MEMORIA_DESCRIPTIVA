# -*- coding: utf-8 -*-
"""
Funciones para la detección de capas adyacentes y colindantes
"""

from qgis.core import QgsProject, QgsVectorLayer, QgsFeature, QgsGeometry, QgsPointXY
import math

def detectar_capas_adyacentes(poligono_layer):
    """
    Detecta todas las capas de polígonos que podrían contener colindantes
    
    Args:
        poligono_layer: Capa de polígono principal
        
    Returns:
        Lista de capas que contienen polígonos adyacentes al polígono principal
    """
    # Obtener el polígono principal
    features = list(poligono_layer.getFeatures())
    if not features:
        return []
    
    poligono_principal = features[0]
    geometria_principal = poligono_principal.geometry()
    
    # Buscar todas las capas de polígonos en el proyecto
    capas_adyacentes = []
    
    for layer in QgsProject.instance().mapLayers().values():
        # Verificar que sea una capa vectorial de polígonos
        if not isinstance(layer, QgsVectorLayer) or layer.geometryType() != 2:
            continue
            
        # Evitar la capa principal
        if layer.id() == poligono_layer.id():
            continue
            
        # Verificar si algún polígono de esta capa toca el polígono principal
        for feature in layer.getFeatures():
            if feature.geometry().touches(geometria_principal):
                capas_adyacentes.append(layer)
                break  # Si encontramos al menos un polígono adyacente, añadimos la capa y pasamos a la siguiente
    
    return capas_adyacentes

def identificar_colindantes(poligono_layer, capas_adyacentes=None):
    """
    Identifica los colindantes del polígono principal y su posición relativa
    
    Args:
        poligono_layer: Capa de polígono principal
        capas_adyacentes: Lista de capas a revisar (opcional, si no se proporciona se detectan automáticamente)
        
    Returns:
        Diccionario con colindantes por posición (NORTE, SUR, ESTE, OESTE)
    """
    # Obtener el polígono principal
    features = list(poligono_layer.getFeatures())
    if not features:
        return {'NORTE': [], 'SUR': [], 'ESTE': [], 'OESTE': []}
    
    poligono_principal = features[0]
    geometria_principal = poligono_principal.geometry()
    
    # Si no se proporcionan capas adyacentes, detectarlas
    if capas_adyacentes is None:
        capas_adyacentes = detectar_capas_adyacentes(poligono_layer)
    
    # Inicializar diccionario de colindantes
    colindantes = {'NORTE': [], 'SUR': [], 'ESTE': [], 'OESTE': []}
    
    # Buscar colindantes en cada capa
    for layer in capas_adyacentes:
        for feature in layer.getFeatures():
            if feature.geometry().touches(geometria_principal):
                # Determinar la posición relativa
                posicion = determinar_posicion_relativa(geometria_principal, feature.geometry())
                
                # Extraer nombre del colindante
                nombre = extraer_nombre_colindante(feature, layer)
                observacion = extraer_observacion_colindante(feature, layer)
                
                if nombre and posicion:
                    colindantes[posicion].append({
                        'nombre': nombre,
                        'observacion': observacion,
                        'feature': feature,
                        'layer': layer.name()
                    })
    
    # Si no hay colindantes en alguna dirección, añadir "Terrenos del estado" como valor por defecto
    for direccion in colindantes:
        if not colindantes[direccion]:
            colindantes[direccion].append({
                'nombre': "Terrenos del estado",
                'observacion': "",
                'feature': None,
                'layer': None
            })
    
    return colindantes

def determinar_posicion_relativa(geom_principal, geom_colindante):
    """
    Determina si un polígono colindante está al N, S, E u O del principal
    
    Args:
        geom_principal: Geometría del polígono principal
        geom_colindante: Geometría del polígono colindante
        
    Returns:
        String con la posición relativa ('NORTE', 'SUR', 'ESTE', 'OESTE')
    """
    # Obtener centroides
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

def extraer_nombre_colindante(feature, layer):
    """
    Extrae el nombre del colindante de los atributos
    
    Args:
        feature: Feature del colindante
        layer: Capa del colindante
        
    Returns:
        String con el nombre del colindante o valor por defecto
    """
    # Campos posibles que pueden contener nombres
    campos_nombre = ['nombre', 'nom_tit', 'propietario', 'titular', 'name', 'owner', 'dueño']
    
    for campo in campos_nombre:
        idx = feature.fieldNameIndex(campo)
        if idx != -1 and feature[campo]:
            return feature[campo]
    
    # Si no encuentra, devolver nombre genérico
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
    campos_obs = ['observacion', 'obs', 'descripcion', 'desc', 'notas', 'comentario']
    
    for campo in campos_obs:
        idx = feature.fieldNameIndex(campo)
        if idx != -1 and feature[campo]:
            return feature[campo]
    
    # Si no encuentra, devolver cadena vacía
    return ""

def obtener_limites_poligono(poligono_geometry):
    """
    Obtiene los límites extremos del polígono para ayudar a determinar colindantes
    
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

def verificar_colindancia_por_limite(poligono_geometry, otro_poligono_geometry, direccion):
    """
    Verifica si dos polígonos colindan en una dirección específica
    
    Args:
        poligono_geometry: Geometría del polígono principal
        otro_poligono_geometry: Geometría del posible colindante
        direccion: Dirección a verificar ('NORTE', 'SUR', 'ESTE', 'OESTE')
        
    Returns:
        Boolean indicando si colindan en esa dirección
    """
    # Obtener límites de ambos polígonos
    limites_principal = obtener_limites_poligono(poligono_geometry)
    limites_otro = obtener_limites_poligono(otro_poligono_geometry)
    
    # Verificar colindancia según dirección
    if direccion == 'NORTE':
        return (poligono_geometry.touches(otro_poligono_geometry) and 
                limites_otro['sur'].y() >= limites_principal['norte'].y())
    elif direccion == 'SUR':
        return (poligono_geometry.touches(otro_poligono_geometry) and 
                limites_otro['norte'].y() <= limites_principal['sur'].y())
    elif direccion == 'ESTE':
        return (poligono_geometry.touches(otro_poligono_geometry) and 
                limites_otro['oeste'].x() >= limites_principal['este'].x())
    elif direccion == 'OESTE':
        return (poligono_geometry.touches(otro_poligono_geometry) and 
                limites_otro['este'].x() <= limites_principal['oeste'].x())
    
    return False
