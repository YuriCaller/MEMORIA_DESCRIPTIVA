# -*- coding: utf-8 -*-
"""
Procesamiento de coordenadas y cálculo de área/perímetro.
VERSIÓN 2.1 - Extrae área desde BD, calcula geométricamente como fallback,
              soporte para filtrado por feature (modo atlas).
"""

from qgis.core import QgsProject, QgsVectorLayer, QgsFeature, QgsDistanceArea
import math


# ===========================================================================
# PROCESAMIENTO DE COORDENADAS
# ===========================================================================

def procesar_coordenadas(punto_layer, linea_layer=None, campos_config=None, filtro_expresion=None):
    """
    Procesa la capa de puntos para obtener vértices con coordenadas, distancias y azimuts.

    Args:
        punto_layer     : Capa de puntos
        linea_layer     : No usada (compatibilidad)
        campos_config   : Dict con configuración de campos
        filtro_expresion: Expresión QGIS para filtrar puntos (modo atlas, ej. "\"nombre\"='NOEMIA'")

    Returns:
        Lista de dicts con info de cada vértice
    """
    if campos_config is None:
        campos_config = {}

    campo_orden = campos_config.get('orden_punto')

    # Aplicar filtro si se proporciona (modo atlas)
    if filtro_expresion:
        request = punto_layer.getFeatures(filtro_expresion)
        puntos = list(request)
    else:
        puntos = list(punto_layer.getFeatures())

    if not puntos:
        print("AVISO: No se encontraron puntos{}".format(
            " con filtro: " + filtro_expresion if filtro_expresion else ""))
        return []

    # Ordenar
    puntos = _ordenar_puntos(punto_layer, puntos, campo_orden)
    print("Procesando {} puntos...".format(len(puntos)))

    datos_vertices = []
    for i, punto in enumerate(puntos):
        geom = punto.geometry()
        if not geom or geom.isEmpty():
            continue
        coord = geom.asPoint()

        este  = _num(punto, campos_config.get('este'),  ['este', 'east', 'x', 'coord_x', 'easting'])
        norte = _num(punto, campos_config.get('norte'), ['norte', 'north', 'y', 'coord_y', 'northing'])
        if este  is None: este  = coord.x()
        if norte is None: norte = coord.y()

        vid  = _txt(punto, campos_config.get('vertice_id'), ['vertice', 'vertice_id', 'id', 'punto_id', 'fid'])
        if not vid: vid = 'V-{:02d}'.format(i + 1)

        lado      = _txt(punto, campos_config.get('lado'),     ['lado', 'side', 'segment', 'tramo'])
        distancia = _num(punto, campos_config.get('distancia'), ['distancia', 'distance', 'dist', 'longitud'])
        azimut    = _num(punto, campos_config.get('azimut'),    ['azimut', 'azimuth', 'rumbo', 'bearing'])

        datos_vertices.append({
            'vertice': vid, 'lado': lado,
            'este': este,   'norte': norte,
            'distancia': distancia, 'azimut': azimut,
            '_x': coord.x(), '_y': coord.y()
        })

    if not datos_vertices:
        return []

    # Completar distancias / azimuts / lados geométricamente
    n = len(datos_vertices)
    for i, v in enumerate(datos_vertices):
        sig = datos_vertices[(i + 1) % n]
        if not v['lado']:
            v['lado'] = '{}-{}'.format(v['vertice'], sig['vertice'])
        if v['distancia'] is None or v['distancia'] == 0.0:
            dx = sig['_x'] - v['_x'];  dy = sig['_y'] - v['_y']
            v['distancia'] = math.sqrt(dx*dx + dy*dy)
        if v['azimut'] is None or v['azimut'] == 0.0:
            dx = sig['_x'] - v['_x'];  dy = sig['_y'] - v['_y']
            az = math.degrees(math.atan2(dx, dy))
            if az < 0: az += 360.0
            v['azimut'] = round(az, 4)
        v.pop('_x', None); v.pop('_y', None)

    return datos_vertices


# ===========================================================================
# ÁREA Y PERÍMETRO — con prioridad a campo BD
# ===========================================================================

def calcular_area_perimetro(poligono_layer, campos_config=None, feature=None):
    """
    Obtiene área (ha) y perímetro (m) del polígono.

    Estrategia:
      1. Si campos_config especifica campos, extrae de la BD.
      2. Si no hay campo o el valor es nulo, calcula desde la geometría.
      3. Devuelve también 'fuente_area' y 'fuente_perimetro' para informar al usuario.

    Args:
        poligono_layer: Capa de polígonos
        campos_config : Dict con 'area' y 'perimetro' (nombres de campo)
        feature       : Feature específico (modo atlas); si None usa el primero
    """
    if campos_config is None:
        campos_config = {}

    if feature is None:
        feats = list(poligono_layer.getFeatures())
        if not feats:
            return {'area': 0, 'perimetro': 0, 'fuente_area': 'sin datos', 'fuente_perimetro': 'sin datos'}
        feature = feats[0]

    geom = feature.geometry()

    # ── Área ──────────────────────────────────────────────────────────────────
    area_ha = None
    fuente_area = 'geometría'
    campo_area = campos_config.get('area')
    if campo_area:
        field_names = [f.name() for f in feature.fields()]
        if campo_area in field_names:
            val = feature[campo_area]
            if val is not None:
                try:
                    area_ha = float(val)
                    # Detectar si está en m² (>5000 implica probablemente m²)
                    if area_ha > 5000:
                        area_ha = round(area_ha / 10000, 6)
                        fuente_area = 'campo BD "{}" (convertido de m²)'.format(campo_area)
                    else:
                        fuente_area = 'campo BD "{}"'.format(campo_area)
                except (ValueError, TypeError):
                    pass

    if area_ha is None:
        # Calcular con QgsDistanceArea para mayor precisión
        da = QgsDistanceArea()
        da.setEllipsoid('WGS84')
        try:
            area_m2 = da.measureArea(geom)
        except Exception:
            area_m2 = geom.area()
        area_ha = round(area_m2 / 10000, 6)
        fuente_area = 'geometría (calculada)'

    # ── Perímetro ─────────────────────────────────────────────────────────────
    perim_m = None
    fuente_perim = 'geometría'
    campo_perim = campos_config.get('perimetro')
    if campo_perim:
        field_names = [f.name() for f in feature.fields()]
        if campo_perim in field_names:
            val = feature[campo_perim]
            if val is not None:
                try:
                    perim_m = float(val)
                    fuente_perim = 'campo BD "{}"'.format(campo_perim)
                except (ValueError, TypeError):
                    pass

    if perim_m is None:
        da = QgsDistanceArea()
        da.setEllipsoid('WGS84')
        try:
            perim_m = da.measurePerimeter(geom)
        except Exception:
            perim_m = geom.length()
        perim_m = round(perim_m, 4)
        fuente_perim = 'geometría (calculada)'

    print("Área: {:.6f} ha [{}]  |  Perímetro: {:.4f} m [{}]".format(
        area_ha, fuente_area, perim_m, fuente_perim))

    return {
        'area': area_ha, 'perimetro': perim_m,
        'fuente_area': fuente_area, 'fuente_perimetro': fuente_perim
    }


# ===========================================================================
# DESCRIPCIÓN DE LINDEROS
# ===========================================================================

def generar_descripcion_linderos(datos_vertices):
    """Descripción textual profesional de linderos con rumbos."""
    if not datos_vertices:
        return "No hay datos de vértices disponibles."

    n = len(datos_vertices)
    partes = ["Comienza en el vértice {}".format(datos_vertices[0]['vertice'])]
    for i, v in enumerate(datos_vertices):
        sig = datos_vertices[(i + 1) % n]
        partes.append(
            "con rumbo {} y una distancia de {:.2f} m llega al vértice {}".format(
                _az_a_rumbo(v['azimut']), v['distancia'], sig['vertice']))
    return "; ".join(partes) + "; cerrando así el perímetro del predio."


def _az_a_rumbo(az_deg):
    try:
        az = float(az_deg) % 360.0
        g = int(az); md = (az - g) * 60; m = int(md); s = int((md - m) * 60)
        if az <= 90:   return "N {}°{:02d}'{:02d}\" E".format(g, m, s)
        elif az <= 180:
            a = 180 - az; g2 = int(a); md2=(a-g2)*60; m2=int(md2); s2=int((md2-m2)*60)
            return "S {}°{:02d}'{:02d}\" E".format(g2, m2, s2)
        elif az <= 270:
            a = az-180;   g2=int(a); md2=(a-g2)*60; m2=int(md2); s2=int((md2-m2)*60)
            return "S {}°{:02d}'{:02d}\" O".format(g2, m2, s2)
        else:
            a = 360-az;   g2=int(a); md2=(a-g2)*60; m2=int(md2); s2=int((md2-m2)*60)
            return "N {}°{:02d}'{:02d}\" O".format(g2, m2, s2)
    except Exception:
        return "{:.4f}°".format(az_deg)


# ===========================================================================
# SISTEMA DE COORDENADAS
# ===========================================================================

def obtener_info_sistema_coordenadas(poligono_layer):
    crs = poligono_layer.crs()
    desc = crs.description()
    info = {
        'Sistema de coordenadas': desc,
        'Unidades': 'Metros',
        'Elipsoide': crs.ellipsoidAcronym(),
        'Grillado': 'Cada 1 000 metros'
    }
    import re
    zm = re.search(r'zone\s*(\d+\s*[ns]?)', desc.lower())
    if zm:
        info['Sistema de coordenadas'] = 'Datum WGS 84 / UTM zona {}'.format(zm.group(1).upper())
    return info


# ===========================================================================
# HELPERS PRIVADOS
# ===========================================================================

def _num(feature, campo_cfg, nombres):
    if campo_cfg:
        fnames = [f.name() for f in feature.fields()]
        if campo_cfg in fnames:
            v = feature[campo_cfg]
            if v is not None:
                try: return float(v)
                except: pass
    for n in nombres:
        for f in feature.fields():
            if f.name().lower() == n.lower():
                v = feature[f.name()]
                if v is not None:
                    try: return float(v)
                    except: pass
    return None


def _txt(feature, campo_cfg, nombres):
    if campo_cfg:
        fnames = [f.name() for f in feature.fields()]
        if campo_cfg in fnames:
            v = feature[campo_cfg]
            if v is not None and str(v).strip():
                return str(v).strip()
    for n in nombres:
        for f in feature.fields():
            if f.name().lower() == n.lower():
                v = feature[f.name()]
                if v is not None and str(v).strip():
                    return str(v).strip()
    return None


def _ordenar_puntos(layer, puntos, campo_orden=None):
    if campo_orden and campo_orden in [f.name() for f in layer.fields()]:
        try:
            puntos.sort(key=lambda f: float(f[campo_orden]) if f[campo_orden] is not None else 999999)
            return puntos
        except: pass
    for campo in ['id','ID','fid','FID','orden','order','secuencia','num','vertice']:
        if layer.fields().indexFromName(campo) != -1:
            try:
                puntos.sort(key=lambda f: float(f[campo]) if f[campo] is not None else 999999)
                return puntos
            except:
                try:
                    puntos.sort(key=lambda f: str(f[campo]) if f[campo] is not None else 'ZZZ')
                    return puntos
                except: pass
    return _orden_espacial(puntos)


def _orden_espacial(puntos):
    if not puntos: return puntos
    xs = [p.geometry().asPoint().x() for p in puntos if p.geometry() and not p.geometry().isEmpty()]
    ys = [p.geometry().asPoint().y() for p in puntos if p.geometry() and not p.geometry().isEmpty()]
    if not xs: return puntos
    cx, cy = sum(xs)/len(xs), sum(ys)/len(ys)
    def ang(p):
        x = p.geometry().asPoint().x()-cx; y = p.geometry().asPoint().y()-cy
        a = math.degrees(math.atan2(x, y))
        return a if a >= 0 else a+360
    puntos.sort(key=ang)
    return puntos


# Alias para compatibilidad
def extraer_campo_numerico(feature, nombres_posibles):
    return _num(feature, None, nombres_posibles)

def extraer_campo_texto(feature, nombres_posibles):
    return _txt(feature, None, nombres_posibles)
