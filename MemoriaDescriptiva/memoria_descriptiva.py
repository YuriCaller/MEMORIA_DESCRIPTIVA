# -*- coding: utf-8 -*-
"""
Plugin principal Memoria Descriptiva.
VERSIÓN 2.1:
  - Modo Atlas: genera un documento por cada polígono/feature
  - Modo Único: un documento para el primer polígono
  - Verifica y reporta fuente del área (BD vs geometría)
  - Barra de progreso con cancelación
  - Nuevo ícono
"""

from qgis.PyQt.QtCore import QSettings, QTranslator, QCoreApplication, Qt
from qgis.PyQt.QtGui import QIcon
from qgis.PyQt.QtWidgets import (QAction, QFileDialog, QMessageBox,
                                  QProgressDialog, QApplication)
from qgis.core import QgsProject, QgsVectorLayer, QgsFeatureRequest

from .resources import *
from .memoria_descriptiva_dialog import MemoriaDescriptivaDialog
import os, sys, traceback

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

sys.path.insert(0, os.path.dirname(__file__))
try:
    from .deteccion_capas_adyacentes import detectar_capas_adyacentes
    from .identificacion_colindantes import identificar_colindantes_completo
    from .procesamiento_coordenadas import (
        procesar_coordenadas, calcular_area_perimetro,
        generar_descripcion_linderos, obtener_info_sistema_coordenadas)
    from .generacion_documento_word import generar_documento_word
    _MODS_OK = True; _MODS_ERR = ''
except ImportError as e:
    _MODS_OK = False; _MODS_ERR = str(e)


class MemoriaDescriptivaPlugin:

    def __init__(self, iface):
        self.iface = iface
        self.plugin_dir = os.path.dirname(__file__)
        locale = QSettings().value('locale/userLocale')[0:2]
        lp = os.path.join(self.plugin_dir, 'i18n', 'MemoriaDescriptiva_{}.qm'.format(locale))
        if os.path.exists(lp):
            self.translator = QTranslator()
            self.translator.load(lp)
            QCoreApplication.installTranslator(self.translator)
        self.actions = []
        self.menu = self.tr(u'&Memoria Descriptiva')
        self.first_start = True
        self.dlg = None

    def tr(self, msg):
        return QCoreApplication.translate('MemoriaDescriptiva', msg)

    def add_action(self, icon_path, text, callback, parent=None):
        icon = QIcon(icon_path)
        action = QAction(icon, text, parent)
        action.triggered.connect(callback)
        self.iface.addToolBarIcon(action)
        self.iface.addPluginToVectorMenu(self.menu, action)
        self.actions.append(action)
        return action

    def initGui(self):
        self.add_action(
            os.path.join(self.plugin_dir, 'icon.png'),
            text=self.tr(u'Generar Memoria Descriptiva'),
            callback=self.run,
            parent=self.iface.mainWindow())
        self.first_start = True

    def unload(self):
        for a in self.actions:
            self.iface.removePluginVectorMenu(self.tr(u'&Memoria Descriptiva'), a)
            self.iface.removeToolBarIcon(a)

    # =========================================================================
    # RUN
    # =========================================================================

    def run(self):
        if not HAS_DOCX:
            QMessageBox.critical(
                self.iface.mainWindow(), "Dependencia faltante",
                "Este plugin requiere <b>python-docx</b>.<br><br>"
                "Instale con:<br><code>pip install python-docx</code>"); return
        if not _MODS_OK:
            QMessageBox.critical(
                self.iface.mainWindow(), "Error de módulos",
                "Error al cargar módulos:<br>{}".format(_MODS_ERR)); return

        if self.first_start:
            self.first_start = False
            self.dlg = MemoriaDescriptivaDialog()
            self.dlg.btnBrowse.clicked.connect(self._select_output)
            self.dlg.btnGenerar.clicked.connect(self._generar)

        self._cargar_capas()
        self._autodetectar_crs()
        self.dlg.show()
        self.dlg.exec_()

    def _cargar_capas(self):
        for cbo in [self.dlg.cboPoligonos, self.dlg.cboPuntos, self.dlg.cboLineas]:
            cbo.blockSignals(True); cbo.clear()
        self.dlg.cboPoligonos.addItem('-- Seleccione --', None)
        self.dlg.cboPuntos.addItem('-- Seleccione --', None)
        self.dlg.cboLineas.addItem('-- Opcional --', None)

        for layer in QgsProject.instance().mapLayers().values():
            if isinstance(layer, QgsVectorLayer):
                gt = layer.geometryType()
                if gt == 2:   self.dlg.cboPoligonos.addItem(layer.name(), layer.id())
                elif gt == 0: self.dlg.cboPuntos.addItem(layer.name(), layer.id())
                elif gt == 1: self.dlg.cboLineas.addItem(layer.name(), layer.id())

        for cbo in [self.dlg.cboPoligonos, self.dlg.cboPuntos, self.dlg.cboLineas]:
            cbo.blockSignals(False)

        if self.dlg.cboPoligonos.count() > 1:
            self.dlg.cboPoligonos.setCurrentIndex(1)
            self.dlg.actualizar_campos_poligono()
        if self.dlg.cboPuntos.count() > 1:
            self.dlg.cboPuntos.setCurrentIndex(1)
            self.dlg.actualizar_campos_puntos()

    def _autodetectar_crs(self):
        lid = self.dlg.cboPoligonos.currentData()
        if not lid: return
        layer = QgsProject.instance().mapLayer(lid)
        if not layer: return
        try:
            info = obtener_info_sistema_coordenadas(layer)
            campos = [('txtSistema','Sistema de coordenadas'),('txtUnidades','Unidades'),
                      ('txtElipsoide','Elipsoide'),('txtGrillado','Grillado')]
            for attr, key in campos:
                w = getattr(self.dlg, attr, None)
                if w and not w.text():
                    w.setText(info.get(key,''))
        except Exception as e:
            print("CRS autodetect error: {}".format(e))

    def _select_output(self):
        fn, _ = QFileDialog.getSaveFileName(
            self.dlg, "Guardar Memoria Descriptiva",
            os.path.expanduser('~'), "Documentos Word (*.docx)")
        if fn:
            if not fn.lower().endswith('.docx'): fn += '.docx'
            self.dlg.txtOutputFile.setText(fn)

    # =========================================================================
    # GENERACIÓN PRINCIPAL
    # =========================================================================

    def _generar(self):
        if not self.dlg.validar_formulario(): return
        datos_form = self.dlg.obtener_datos_formulario()
        modo_atlas = datos_form['atlas']['activo']

        if modo_atlas:
            self._generar_atlas(datos_form)
        else:
            self._generar_unico(datos_form)

    # ── Modo único ────────────────────────────────────────────────────────────

    def _generar_unico(self, datos_form):
        pol_layer = QgsProject.instance().mapLayer(datos_form['capas']['poligono_id'])
        pnt_layer = QgsProject.instance().mapLayer(datos_form['capas']['punto_id'])
        lin_id    = datos_form['capas'].get('linea_id')
        lin_layer = QgsProject.instance().mapLayer(lin_id) if lin_id else None

        feats = list(pol_layer.getFeatures())
        if not feats:
            QMessageBox.warning(self.dlg, "Sin datos", "La capa de polígonos no tiene objetos."); return

        feature = feats[0]

        prog = self._progress("Generando Memoria Descriptiva...", 6)
        try:
            dp = self._procesar_feature(
                datos_form, feature, pol_layer, pnt_layer, lin_layer,
                filtro_puntos=None, progress=prog)
            if dp is None: return  # Cancelado

            datos_form['nombre_predio'] = ''
            out = generar_documento_word(datos_form, dp)
            prog.close()

            QMessageBox.information(
                self.dlg, "¡Éxito!",
                "<b>Memoria descriptiva generada:</b><br>{}<br><br>"
                "Vértices: {} | Área: {:.4f} ha | Perímetro: {:.2f} m".format(
                    out, len(dp['vertices']), dp['area'], dp['perimetro']))
            self.dlg.accept()
        except Exception as e:
            prog.close()
            QMessageBox.critical(self.dlg, "Error", "{}\n\n{}".format(str(e), traceback.format_exc()))

    # ── Modo atlas ────────────────────────────────────────────────────────────

    def _generar_atlas(self, datos_form):
        pol_layer  = QgsProject.instance().mapLayer(datos_form['capas']['poligono_id'])
        pnt_layer  = QgsProject.instance().mapLayer(datos_form['capas']['punto_id'])
        lin_id     = datos_form['capas'].get('linea_id')
        lin_layer  = QgsProject.instance().mapLayer(lin_id) if lin_id else None

        campo_atlas  = datos_form['atlas']['campo_atlas']
        campo_filtro = datos_form['atlas']['campo_filtro']
        solo_sel     = datos_form['atlas']['solo_seleccion']

        # Obtener features
        if solo_sel:
            features = list(pol_layer.selectedFeatures())
            if not features:
                QMessageBox.warning(self.dlg, "Atlas",
                    "No hay objetos seleccionados en la capa de polígonos."); return
        else:
            features = list(pol_layer.getFeatures())

        if not features:
            QMessageBox.warning(self.dlg, "Atlas",
                "La capa de polígonos no tiene objetos."); return

        total = len(features)
        prog = self._progress("Modo Atlas — procesando predios...", total)
        generados = []; errores = []

        for i, feature in enumerate(features):
            if prog.wasCanceled(): break

            # Nombre del predio desde campo atlas
            nombre_predio = ''
            if campo_atlas:
                fnames = [f.name() for f in feature.fields()]
                if campo_atlas in fnames:
                    v = feature[campo_atlas]
                    nombre_predio = str(v).strip() if v is not None else ''

            if not nombre_predio:
                nombre_predio = 'predio_{:03d}'.format(i + 1)

            prog.setLabelText(
                "Atlas {}/{}: {}".format(i+1, total, nombre_predio))
            prog.setValue(i)
            QApplication.processEvents()

            # Expresión de filtro para puntos
            filtro_puntos = None
            if campo_filtro and nombre_predio:
                # Usar el mismo valor del campo atlas para filtrar puntos
                campo_f_puntos = campo_filtro
                filtro_puntos = '"{}" = \'{}\''.format(
                    campo_f_puntos, nombre_predio.replace("'", "\\'"))
            elif campo_atlas and nombre_predio:
                # Intentar con el mismo campo en la capa de puntos
                pnt_fields = [f.name() for f in pnt_layer.fields()]
                if campo_atlas in pnt_fields:
                    filtro_puntos = '"{}" = \'{}\''.format(
                        campo_atlas, nombre_predio.replace("'", "\\'"))

            try:
                # Clonar datos_form para este predio
                datos_form_i = dict(datos_form)
                datos_form_i['nombre_predio'] = nombre_predio

                dp = self._procesar_feature(
                    datos_form_i, feature, pol_layer, pnt_layer, lin_layer,
                    filtro_puntos=filtro_puntos, progress=None)

                if dp is None: continue

                out = generar_documento_word(datos_form_i, dp, sufijo_archivo=nombre_predio)
                generados.append((nombre_predio, out))

            except Exception as e:
                errores.append((nombre_predio, str(e)))
                print("ERROR en predio {}: {}".format(nombre_predio, traceback.format_exc()))

        prog.setValue(total)
        prog.close()

        # Resumen final
        if generados:
            msg = "<b>Atlas generado: {} memorias.</b><br><br>".format(len(generados))
            msg += "<b>Archivos generados:</b><br>"
            for nombre, path in generados[:20]:
                msg += "✓ {} → <small>{}</small><br>".format(nombre, os.path.basename(path))
            if len(generados) > 20:
                msg += "... y {} más<br>".format(len(generados) - 20)
            if errores:
                msg += "<br><b style='color:red'>Errores ({}):</b><br>".format(len(errores))
                for nombre, err in errores[:5]:
                    msg += "✗ {}: {}<br>".format(nombre, err[:80])
            QMessageBox.information(self.dlg, "Atlas completado", msg)
            self.dlg.accept()
        else:
            msg = "No se generó ningún documento."
            if errores:
                msg += "\n\nErrores:\n" + "\n".join(
                    "• {}: {}".format(n, e) for n,e in errores[:10])
            QMessageBox.warning(self.dlg, "Atlas sin resultados", msg)

    # =========================================================================
    # PROCESAMIENTO DE UN FEATURE
    # =========================================================================

    def _procesar_feature(self, datos_form, feature, pol_layer,
                          pnt_layer, lin_layer, filtro_puntos=None, progress=None):
        """
        Procesa un feature de polígono y retorna datos_procesados.
        progress puede ser None (modo atlas con su propio progreso).
        """
        def _step(n, msg):
            if progress:
                if progress.wasCanceled(): return False
                progress.setValue(n)
                progress.setLabelText(msg)
                QApplication.processEvents()
            return True

        campos_config = {
            'vertice_id':  datos_form['campos'].get('vertice_id'),
            'orden_punto': datos_form['campos'].get('orden_punto'),
            'distancia':   datos_form['campos'].get('distancia'),
            'azimut':      datos_form['campos'].get('azimut'),
            'este':        datos_form['campos'].get('este'),
            'norte':       datos_form['campos'].get('norte'),
            'lado':        datos_form['campos'].get('lado'),
        }

        dp = {}

        # Colindantes
        if not _step(1, "Identificando colindantes..."): return None
        if datos_form['colindantes']['detectar_automatico']:
            # Crear capa temporal con solo este feature para detectar colindantes
            capas_adj = detectar_capas_adyacentes(pol_layer)
            dp['colindantes'] = identificar_colindantes_completo(pol_layer, capas_adj)
        else:
            m = datos_form['colindantes']['manual']
            dp['colindantes'] = {
                d: {'nombre': m.get(d,'Terrenos del Estado'), 'observacion': ''}
                for d in ['NORTE','SUR','ESTE','OESTE']
            }

        # Coordenadas de vértices (con filtro en modo atlas)
        if not _step(2, "Procesando coordenadas..."): return None
        dp['vertices'] = procesar_coordenadas(
            pnt_layer, lin_layer, campos_config,
            filtro_expresion=filtro_puntos)

        # Área y perímetro (con el feature específico)
        if not _step(3, "Calculando área y perímetro..."): return None
        mostrar_fuente = datos_form['atlas'].get('mostrar_fuente_area', True)
        ap = calcular_area_perimetro(
            pol_layer,
            {'area': datos_form['campos'].get('area'),
             'perimetro': datos_form['campos'].get('perimetro')},
            feature=feature)
        dp['area']             = ap['area']
        dp['perimetro']        = ap['perimetro']
        dp['fuente_area']      = ap['fuente_area']    if mostrar_fuente else ''
        dp['fuente_perimetro'] = ap['fuente_perimetro'] if mostrar_fuente else ''

        # Descripción de linderos
        if not _step(4, "Generando descripción de linderos..."): return None
        dp['descripcion_linderos'] = generar_descripcion_linderos(dp['vertices'])

        # Info del mapa desde CRS si no se llenó
        if not _step(5, "Completando info del mapa..."): return None
        if not any(v for v in datos_form.get('info_mapa',{}).values()):
            datos_form['info_mapa'] = obtener_info_sistema_coordenadas(pol_layer)

        if not _step(6, "Generando documento Word..."): return None
        return dp

    # =========================================================================
    # HELPERS
    # =========================================================================

    def _progress(self, title, maximo):
        prog = QProgressDialog(title, "Cancelar", 0, maximo, self.dlg)
        prog.setWindowTitle("Memoria Descriptiva")
        prog.setWindowModality(Qt.WindowModal)
        prog.setMinimumDuration(0)
        prog.setValue(0)
        prog.show()
        QApplication.processEvents()
        return prog
