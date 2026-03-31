# -*- coding: utf-8 -*-
"""
Procesamiento de coordenadas para Memoria Descriptiva.
VERSIÓN 3.0 — Lógica basada en la estructura real:
  - AREA_TOTAL (polígonos): fid, NombresApellidos, Area_ha, Perímetro
  - Puntos: ID_Poligono, ID_Vertice, LADO, Este, Norte, Distancia, Azimut
  - perimetros: ID_Poligono, ID_Segmento, longitud, azimut (opcional)
"""

import math
from qgis.core import QgsDistanceArea, QgsFeatureRequest


def obtener_vertices_de_poligono(punto_layer, id_poligono, campos_config=None):
    """
    Obtiene los vértices de UN polígono específico filtrando por ID_Poligono.

    Estrategia de filtrado (en orden de preferencia):
      1. Campo configurado en campos_config['campo_id_poligono']
      2. Auto-detección por nombres conocidos: ID_Poligono, id_poligono, fid_area, poligono_id
      3. Si no existe campo de relación, usa TODOS los puntos ordenados

    Args:
        punto_layer   : QgsVectorLayer de puntos
        id_poligono   : valor del ID a filtrar (int o str)
        campos_config : dict con configuración de campos

    Returns:
        Lista de dicts con: vertice, lado, este, norte, distancia, azimut
    """
    if campos_config is None:
        campos_config = {}

    # ── Detectar campo de relación ────────────────────────────────────────────
    campo_rel = campos_config.get('campo_id_poligono')
    if not campo_rel:
        campo_rel = _detectar_campo(punto_layer,
            ['ID_Poligono', 'id_poligono', 'fid_area', 'poligono_id',
             'id_pol', 'pol_id', 'ID_POL'])

    # ── Filtrar features ──────────────────────────────────────────────────────
    if campo_rel:
        try:
            id_val = int(id_poligono)
        except (ValueError, TypeError):
            id_val = id_poligono

        # Intentar expresión SQL primero (rápido)
        puntos = []
        try:
            campo_field = next((f for f in punto_layer.fields()
                                if f.name() == campo_rel), None)
            # QVariant tipos numéricos: Int=2, LongLong=4, Double=6
            es_num = (campo_field is not None and campo_field.type() in (1, 2, 4, 6, 10, 35))
            if es_num:
                expr = '"{}" = {}'.format(campo_rel, id_val)
            else:
                expr = '"{}" = \'{}\''.format(campo_rel, id_val)
            req = QgsFeatureRequest()
            req.setFilterExpression(expr)
            puntos = list(punto_layer.getFeatures(req))
            print("  Filtro SQL '{}': {} puntos".format(expr, len(puntos)))
        except Exception as e:
            print("  Error filtro SQL: {}".format(e))
            puntos = []

        # Fallback manual si SQL no devolvió nada
        if not puntos:
            print("  Fallback manual: buscando {} = {}...".format(campo_rel, id_val))
            for f in punto_layer.getFeatures():
                val = f[campo_rel]
                if val is None:
                    continue
                try:
                    if int(val) == int(id_val):
                        puntos.append(f)
                except (ValueError, TypeError):
                    if str(val).strip() == str(id_val).strip():
                        puntos.append(f)
            print("  Fallback encontró: {} puntos".format(len(puntos)))
    else:
        puntos = list(punto_layer.getFeatures())
        print("  Sin campo relación: usando todos los {} puntos".format(len(puntos)))

    if not puntos:
        return []

    # ── Ordenar por ID_Vertice ────────────────────────────────────────────────
    campo_orden = campos_config.get('campo_vertice') or _detectar_campo(
        punto_layer, ['ID_Vertice', 'id_vertice', 'orden', 'id', 'fid', 'secuencia'])

    if campo_orden:
        try:
            puntos.sort(key=lambda f: _to_float(f[campo_orden]))
        except Exception as e:
            print("  Advertencia ordenando por {}: {}".format(campo_orden, e))
    
    # ── Extraer datos de cada punto ───────────────────────────────────────────
    # Detectar campos de datos
    c_lado  = campos_config.get('campo_lado')      or _detectar_campo(punto_layer, ['LADO', 'lado', 'side', 'segment', 'tramo'])
    c_este  = campos_config.get('campo_este')      or _detectar_campo(punto_layer, ['Este', 'este', 'ESTE', 'X', 'x', 'coord_x', 'easting'])
    c_norte = campos_config.get('campo_norte')     or _detectar_campo(punto_layer, ['Norte', 'norte', 'NORTE', 'Y', 'y', 'coord_y', 'northing'])
    c_dist  = campos_config.get('campo_distancia') or _detectar_campo(punto_layer, ['Distancia', 'distancia', 'DISTANCIA', 'distance', 'dist', 'longitud'])
    c_az    = campos_config.get('campo_azimut')    or _detectar_campo(punto_layer, ['Azimut', 'azimut', 'AZIMUT', 'azimuth', 'rumbo', 'bearing'])
    c_vid   = campos_config.get('campo_vertice')   or _detectar_campo(punto_layer, ['ID_Vertice', 'id_vertice', 'vertice', 'ID', 'id'])

    vertices = []
    n = len(puntos)

    for i, punto in enumerate(puntos):
        geom = punto.geometry()
        if not geom or geom.isEmpty():
            continue
        coord = geom.asPoint()

        # ID del vértice
        vid = _get_val_str(punto, c_vid)
        if not vid:
            vid = 'V{:02d}'.format(i + 1)
        else:
            vid = 'V{:02d}'.format(int(_to_float(vid))) if str(vid).isdigit() else str(vid)

        # Coordenadas (campo BD o geometría)
        este  = _get_val_num(punto, c_este)
        norte = _get_val_num(punto, c_norte)
        if este  is None: este  = coord.x()
        if norte is None: norte = coord.y()

        # Distancia y azimut (campo BD o calcular después)
        distancia = _get_val_num(punto, c_dist)
        azimut    = _get_val_num(punto, c_az)

        # Lado
        lado = _get_val_str(punto, c_lado)

        vertices.append({
            'vertice': vid, 'lado': lado,
            'este': este,   'norte': norte,
            'distancia': distancia, 'azimut': azimut,
            '_x': coord.x(), '_y': coord.y()
        })

    # ── Completar distancias/azimuts/lados faltantes ─────────────────────────
    n = len(vertices)
    for i, v in enumerate(vertices):
        sig = vertices[(i + 1) % n]

        if not v['lado']:
            v['lado'] = '{} a {}'.format(v['vertice'], sig['vertice'])

        if v['distancia'] is None or v['distancia'] == 0.0:
            dx = sig['_x'] - v['_x']; dy = sig['_y'] - v['_y']
            v['distancia'] = round(math.sqrt(dx*dx + dy*dy), 4)

        if v['azimut'] is None or v['azimut'] == 0.0:
            dx = sig['_x'] - v['_x']; dy = sig['_y'] - v['_y']
            az = math.degrees(math.atan2(dx, dy))
            v['azimut'] = round(az + 360 if az < 0 else az, 4)

        # Limpiar claves internas
        v.pop('_x', None); v.pop('_y', None)

    return vertices


def calcular_area_perimetro_feature(feature, pol_layer, campos_config=None):
    """
    Obtiene área y perímetro de un feature de polígono.
    Prioridad: campo BD → geometría WGS84.

    Returns:
        dict: area (ha), perimetro (m), fuente_area, fuente_perimetro
    """
    if campos_config is None:
        campos_config = {}

    geom = feature.geometry()

    # ── Campo BD para área ────────────────────────────────────────────────────
    c_area = campos_config.get('campo_area') or _detectar_campo(
        pol_layer, ['Area_ha', 'area_ha', 'AREA_HA', 'area', 'AREA',
                    'hectareas', 'Hectareas', 'HECTAREAS', 'superficie'])
    area_ha = None; fuente_area = 'geometría'

    if c_area:
        val = feature[c_area] if c_area in [f.name() for f in feature.fields()] else None
        if val is not None:
            try:
                area_ha = float(val)
                if area_ha > 5000:  # probablemente en m²
                    area_ha = round(area_ha / 10000, 6)
                    fuente_area = 'campo BD "{}" (convertido de m²)'.format(c_area)
                else:
                    fuente_area = 'campo BD "{}"'.format(c_area)
            except (ValueError, TypeError):
                pass

    if area_ha is None:
        da = QgsDistanceArea(); da.setEllipsoid('WGS84')
        try:   area_ha = round(da.measureArea(geom) / 10000, 6)
        except: area_ha = round(geom.area() / 10000, 6)
        fuente_area = 'geometría (WGS84)'

    # ── Campo BD para perímetro ───────────────────────────────────────────────
    c_perim = campos_config.get('campo_perimetro') or _detectar_campo(
        pol_layer, ['Perímetro', 'Perimetro', 'PERIMETRO', 'perimetro',
                    'perimeter', 'shape_length', 'Shape_Length'])
    perim_m = None; fuente_perim = 'geometría'

    if c_perim:
        val = feature[c_perim] if c_perim in [f.name() for f in feature.fields()] else None
        if val is not None:
            try:
                perim_m = float(val)
                fuente_perim = 'campo BD "{}"'.format(c_perim)
            except (ValueError, TypeError):
                pass

    if perim_m is None:
        da = QgsDistanceArea(); da.setEllipsoid('WGS84')
        try:   perim_m = round(da.measurePerimeter(geom), 4)
        except: perim_m = round(geom.length(), 4)
        fuente_perim = 'geometría (WGS84)'

    print("  Área: {:.4f} ha [{}]  |  Perímetro: {:.2f} m [{}]".format(
        area_ha, fuente_area, perim_m, fuente_perim))

    return {'area': area_ha, 'perimetro': perim_m,
            'fuente_area': fuente_area, 'fuente_perimetro': fuente_perim}


def generar_descripcion_linderos(vertices):
    """Descripción narrativa de linderos con rumbos."""
    if not vertices:
        return "No se encontraron vértices para este polígono."
    n = len(vertices)
    partes = ["Comienza en el vértice {}".format(vertices[0]['vertice'])]
    for i, v in enumerate(vertices):
        sig = vertices[(i + 1) % n]
        partes.append(
            "con rumbo {} y una distancia de {:.2f} m llega al vértice {}".format(
                _az_rumbo(v['azimut']), v['distancia'], sig['vertice']))
    return "; ".join(partes) + "; cerrando así el perímetro del predio."


def obtener_info_sistema_coordenadas(layer):
    """Info CRS de la capa."""
    import re
    crs = layer.crs()
    desc = crs.description()
    info = {'Sistema de coordenadas': desc, 'Unidades': 'Metros',
            'Elipsoide': crs.ellipsoidAcronym(), 'Grillado': 'Cada 1 000 metros'}
    zm = re.search(r'zone\s*(\d+\s*[ns]?)', desc.lower())
    if zm:
        info['Sistema de coordenadas'] = 'Datum WGS 84 / UTM zona {}'.format(
            zm.group(1).upper().strip())
    return info


# ── Helpers ──────────────────────────────────────────────────────────────────

def _detectar_campo(layer, nombres_posibles):
    """Busca el primer campo que coincida con la lista (case-insensitive)."""
    campos = {f.name().lower(): f.name() for f in layer.fields()}
    for n in nombres_posibles:
        if n.lower() in campos:
            return campos[n.lower()]
    return None


def _get_val_num(feature, campo):
    if not campo: return None
    try:
        v = feature[campo]
        return float(v) if v is not None else None
    except: return None


def _get_val_str(feature, campo):
    if not campo: return None
    try:
        v = feature[campo]
        return str(v).strip() if v is not None else None
    except: return None


def _to_float(v):
    try: return float(v)
    except: return 0.0


def _az_rumbo(az_deg):
    try:
        az = float(az_deg) % 360.0
        g = int(az); md = (az - g)*60; m = int(md); s = int((md-m)*60)
        if az <= 90:   return "N {}°{:02d}'{:02d}\" E".format(g,m,s)
        elif az <= 180:
            a=180-az; g2=int(a); md2=(a-g2)*60; m2=int(md2); s2=int((md2-m2)*60)
            return "S {}°{:02d}'{:02d}\" E".format(g2,m2,s2)
        elif az <= 270:
            a=az-180; g2=int(a); md2=(a-g2)*60; m2=int(md2); s2=int((md2-m2)*60)
            return "S {}°{:02d}'{:02d}\" O".format(g2,m2,s2)
        else:
            a=360-az; g2=int(a); md2=(a-g2)*60; m2=int(md2); s2=int((md2-m2)*60)
            return "N {}°{:02d}'{:02d}\" O".format(g2,m2,s2)
    except: return "{:.4f}°".format(az_deg)
