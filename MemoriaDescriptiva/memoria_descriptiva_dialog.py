# -*- coding: utf-8 -*-
"""
Diálogo del plugin Memoria Descriptiva.
VERSIÓN 2.1:
  - Modo ATLAS: genera una memoria por cada polígono/feature
  - Modo ÚNICO: genera una sola memoria para el polígono seleccionado
  - Capa de líneas OPCIONAL
  - Detección de campo atlas (nombre, clave, etc.)
  - Verificación de fuente de área (BD vs geometría)
"""

import os
from qgis.PyQt import uic, QtWidgets
from qgis.PyQt.QtCore import Qt
from qgis.core import QgsProject, QgsVectorLayer

FORM_CLASS, _ = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'memoria_descriptiva_dialog_base.ui'))


class MemoriaDescriptivaDialog(QtWidgets.QDialog, FORM_CLASS):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)

        self._crear_pestana_atlas()
        self._crear_pestana_campos()

        # Conexiones
        self.chkDetectarColindantes.toggled.connect(self._toggle_colindantes)
        self.chkTextoDefault.toggled.connect(self._toggle_generalidades)
        self.cboPoligonos.currentIndexChanged.connect(self.actualizar_campos_poligono)
        self.cboPuntos.currentIndexChanged.connect(self.actualizar_campos_puntos)
        self.cboLineas.currentIndexChanged.connect(self.actualizar_campos_lineas)

        # Estado inicial
        self._toggle_colindantes(self.chkDetectarColindantes.isChecked())
        self._toggle_generalidades(self.chkTextoDefault.isChecked())

    # =========================================================================
    # PESTAÑA ATLAS
    # =========================================================================

    def _crear_pestana_atlas(self):
        """Crea la pestaña de configuración del modo Atlas."""
        self.tabAtlas = QtWidgets.QWidget()
        self.tabWidget.insertTab(1, self.tabAtlas, "🗺  Modo Atlas")

        scroll = QtWidgets.QScrollArea(); scroll.setWidgetResizable(True)
        container = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(container)
        scroll.setWidget(container)
        outer = QtWidgets.QVBoxLayout(self.tabAtlas)
        outer.setContentsMargins(0,0,0,0); outer.addWidget(scroll)

        # ── Banner informativo ────────────────────────────────────────────────
        banner = QtWidgets.QLabel(
            "<b>🗺 Modo Atlas — Generar una memoria por cada polígono</b><br><br>"
            "Activa esta opción cuando tu capa de polígonos representa <b>múltiples predios</b> "
            "(como en un mapa atlas). El plugin generará un documento Word <b>independiente "
            "por cada polígono</b>, usando el campo que elijas como identificador del predio."
        )
        banner.setWordWrap(True)
        banner.setStyleSheet(
            "background:#E8F4FD; padding:10px; border-radius:6px; "
            "border:1px solid #AED6F1;")
        layout.addWidget(banner)

        # ── Activar modo atlas ────────────────────────────────────────────────
        self.grpAtlas = QtWidgets.QGroupBox("Configuración del Modo Atlas")
        gl = QtWidgets.QVBoxLayout()

        self.chkModoAtlas = QtWidgets.QCheckBox(
            "Activar modo Atlas (generar una memoria por cada polígono)")
        self.chkModoAtlas.setStyleSheet("font-weight:bold; font-size:11px;")
        self.chkModoAtlas.toggled.connect(self._toggle_atlas)
        gl.addWidget(self.chkModoAtlas)

        # ── Panel de opciones atlas (se muestra al activar) ───────────────────
        self.panelAtlas = QtWidgets.QWidget()
        form_a = QtWidgets.QFormLayout(self.panelAtlas)

        # Campo identificador (para el nombre del archivo y subtítulo)
        self.cboCampoAtlas = QtWidgets.QComboBox()
        self.cboCampoAtlas.setToolTip(
            "Campo de la capa de polígonos que identifica cada predio.\n"
            "Se usará como nombre en el documento y en el archivo generado.\n"
            "Ejemplo: campo 'Layer' o 'NOMBRE' de tu tabla de atributos.")
        form_a.addRow("Campo identificador del predio:", self.cboCampoAtlas)

        # Campo para filtrar puntos (relacionar puntos con polígono)
        self.cboCampoFiltro = QtWidgets.QComboBox()
        self.cboCampoFiltro.setToolTip(
            "Campo de la capa de PUNTOS que permite relacionarlos con cada polígono.\n"
            "Debe contener los mismos valores que el campo identificador del polígono.\n"
            "Ejemplo: campo 'Layer' en puntos = campo 'Layer' en polígonos.")
        form_a.addRow("Campo de relación en capa de Puntos:", self.cboCampoFiltro)

        # Carpeta de salida (en atlas se genera un archivo por predio)
        lbl_carpeta = QtWidgets.QLabel(
            "<i>En modo Atlas, los archivos se guardan en la carpeta del archivo base "
            "con el nombre: <b>MemoriaDescriptiva_[NOMBRE_PREDIO].docx</b></i>")
        lbl_carpeta.setWordWrap(True)
        lbl_carpeta.setStyleSheet("color:#555;")
        form_a.addRow("", lbl_carpeta)

        # Opciones de filtro
        self.chkSoloSeleccionados = QtWidgets.QCheckBox(
            "Procesar solo los objetos seleccionados en la capa")
        form_a.addRow("", self.chkSoloSeleccionados)

        # Resumen
        self.lblAtlasResumen = QtWidgets.QLabel("")
        self.lblAtlasResumen.setStyleSheet(
            "background:#FFF9E3; padding:6px; border-radius:4px; border:1px solid #F0C040;")
        self.lblAtlasResumen.setWordWrap(True)
        self.lblAtlasResumen.setVisible(False)
        form_a.addRow("", self.lblAtlasResumen)

        # Botón previsualizar
        self.btnPrevisualizarAtlas = QtWidgets.QPushButton("👁  Previsualizar predios a procesar")
        self.btnPrevisualizarAtlas.clicked.connect(self._previsualizar_atlas)
        form_a.addRow("", self.btnPrevisualizarAtlas)

        gl.addWidget(self.panelAtlas)
        self.grpAtlas.setLayout(gl)
        layout.addWidget(self.grpAtlas)

        # ── Info fuente de área ───────────────────────────────────────────────
        grp_area = QtWidgets.QGroupBox("Verificación de Fuente de Área")
        gl2 = QtWidgets.QFormLayout()

        lbl_area_info = QtWidgets.QLabel(
            "El plugin puede obtener el área del predio desde:<br>"
            "1. <b>Campo de la BD</b> (campo AREA, HECTAREAS, etc.) — más preciso si fue calculado con proyección correcta<br>"
            "2. <b>Geometría</b> — calculado en tiempo real con elipsoide WGS84<br><br>"
            "Configura el campo de área en la pestaña <i>Configuración de Campos</i>. "
            "Si no se configura, se calculará desde la geometría."
        )
        lbl_area_info.setWordWrap(True)
        lbl_area_info.setStyleSheet("color:#333;")
        gl2.addRow(lbl_area_info)

        self.chkMostrarFuenteArea = QtWidgets.QCheckBox(
            "Mostrar fuente del área en el documento (campo BD / geometría calculada)")
        self.chkMostrarFuenteArea.setChecked(True)
        gl2.addRow(self.chkMostrarFuenteArea)

        grp_area.setLayout(gl2)
        layout.addWidget(grp_area)

        layout.addStretch()

        # Estado inicial del panel atlas
        self.panelAtlas.setEnabled(False)

    # =========================================================================
    # PESTAÑA CONFIGURACIÓN DE CAMPOS
    # =========================================================================

    def _crear_pestana_campos(self):
        self.tabCampos = QtWidgets.QWidget()
        self.tabWidget.addTab(self.tabCampos, "⚙  Campos")

        scroll = QtWidgets.QScrollArea(); scroll.setWidgetResizable(True)
        container = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(container)
        scroll.setWidget(container)
        outer = QtWidgets.QVBoxLayout(self.tabCampos)
        outer.setContentsMargins(0,0,0,0); outer.addWidget(scroll)

        lbl = QtWidgets.QLabel(
            "<b>Configuración de campos:</b> seleccione el campo de cada dato.<br>"
            "Dejando <i>-- Automático --</i> el plugin detecta o calcula el valor.<br>"
            "<b>Tip:</b> Distancias y azimuts se calculan geométricamente si no existen en la capa.")
        lbl.setWordWrap(True)
        lbl.setStyleSheet("background:#EAF4EA; padding:8px; border-radius:4px;")
        layout.addWidget(lbl)

        # ── Capa de Puntos ────────────────────────────────────────────────────
        grp_p = QtWidgets.QGroupBox("Campos de la Capa de Puntos")
        fp = QtWidgets.QFormLayout()
        self.cboCampoVerticeID   = QtWidgets.QComboBox()
        self.cboCampoOrdenPunto  = QtWidgets.QComboBox()
        self.cboCampoDistancia   = QtWidgets.QComboBox()
        self.cboCampoAzimut      = QtWidgets.QComboBox()
        self.cboCampoEste        = QtWidgets.QComboBox()
        self.cboCampoNorte       = QtWidgets.QComboBox()
        self.cboCampoLado        = QtWidgets.QComboBox()
        fp.addRow("ID / Etiqueta del vértice:",     self.cboCampoVerticeID)
        fp.addRow("Campo de orden / secuencia:",    self.cboCampoOrdenPunto)
        fp.addRow("Distancia al siguiente vértice:", self.cboCampoDistancia)
        fp.addRow("Azimut del lado:",               self.cboCampoAzimut)
        fp.addRow("Coordenada Este (X):",           self.cboCampoEste)
        fp.addRow("Coordenada Norte (Y):",          self.cboCampoNorte)
        fp.addRow("Nombre del lado:",               self.cboCampoLado)
        grp_p.setLayout(fp); layout.addWidget(grp_p)

        # ── Capa de Polígonos ─────────────────────────────────────────────────
        grp_pol = QtWidgets.QGroupBox("Campos de la Capa de Polígonos")
        fpol = QtWidgets.QFormLayout()
        self.cboCampoArea      = QtWidgets.QComboBox()
        self.cboCampoPerimetro = QtWidgets.QComboBox()
        fpol.addRow("Campo Área (ha o m²):",   self.cboCampoArea)
        fpol.addRow("Campo Perímetro (m):", self.cboCampoPerimetro)
        grp_pol.setLayout(fpol); layout.addWidget(grp_pol)

        # Botón auto
        btn = QtWidgets.QPushButton("🔍  Auto-detectar Campos")
        btn.setStyleSheet("padding:6px; font-weight:bold;")
        btn.clicked.connect(self._autodetectar_todos)
        layout.addWidget(btn)
        layout.addStretch()

    # =========================================================================
    # TOGGLES
    # =========================================================================

    def _toggle_colindantes(self, checked):
        self.groupColindantesManual.setEnabled(not checked)

    def _toggle_generalidades(self, checked):
        self.txtGeneralidades.setReadOnly(checked)

    def _toggle_atlas(self, checked):
        self.panelAtlas.setEnabled(checked)
        self.lblAtlasResumen.setVisible(False)
        if checked:
            self._actualizar_combos_atlas()

    # =========================================================================
    # ACTUALIZACIÓN DE CAMPOS
    # =========================================================================

    def actualizar_campos_poligono(self):
        layer_id = self.cboPoligonos.currentData()
        if not layer_id: return
        layer = QgsProject.instance().mapLayer(layer_id)
        if not layer: return
        campos = [f.name() for f in layer.fields()]

        for cbo in [self.cboCampoNombre, self.cboCampoDNI]:
            cbo.clear(); cbo.addItem('-- Automático --', None)
            for c in campos: cbo.addItem(c, c)

        for cbo in [self.cboCampoArea, self.cboCampoPerimetro]:
            cbo.clear(); cbo.addItem('-- Calcular automáticamente --', None)
            for c in campos: cbo.addItem(c, c)

        self._sel(self.cboCampoNombre,    ['nombre','nom_tit','propietario','titular','name'])
        self._sel(self.cboCampoDNI,       ['dni','doc','documento','ruc'])
        self._sel(self.cboCampoArea,      ['area','area_ha','hectareas','hectarea','superficie','shape_area'])
        self._sel(self.cboCampoPerimetro, ['perimetro','perimeter','longitud','shape_length'])

        self._actualizar_combos_atlas()

    def actualizar_campos_puntos(self):
        layer_id = self.cboPuntos.currentData()
        if not layer_id: return
        layer = QgsProject.instance().mapLayer(layer_id)
        if not layer: return
        campos = [f.name() for f in layer.fields()]

        for cbo in [self.cboCampoOrden, self.cboCampoID]:
            cbo.clear(); cbo.addItem('-- Automático --', None)
            for c in campos: cbo.addItem(c, c)

        auto_map = {
            self.cboCampoVerticeID:  '-- Generar automáticamente --',
            self.cboCampoOrdenPunto: '-- Detectar automáticamente --',
            self.cboCampoDistancia:  '-- Calcular geométricamente --',
            self.cboCampoAzimut:     '-- Calcular geométricamente --',
            self.cboCampoEste:       '-- Usar coordenada X del punto --',
            self.cboCampoNorte:      '-- Usar coordenada Y del punto --',
            self.cboCampoLado:       '-- Generar automáticamente --',
        }
        for cbo, lbl in auto_map.items():
            cbo.clear(); cbo.addItem(lbl, None)
            for c in campos: cbo.addItem(c, c)

        self._sel(self.cboCampoOrden,      ['id','orden','order','secuencia','fid'])
        self._sel(self.cboCampoID,         ['id','fid','objectid','punto_id','vertice'])
        self._sel(self.cboCampoVerticeID,  ['id','vertice','vertice_id','punto_id','fid'])
        self._sel(self.cboCampoOrdenPunto, ['id','orden','order','secuencia','fid'])
        self._sel(self.cboCampoDistancia,  ['distancia','distance','dist','longitud'])
        self._sel(self.cboCampoAzimut,     ['azimut','azimuth','rumbo','bearing'])
        self._sel(self.cboCampoEste,       ['este','east','x','coord_x'])
        self._sel(self.cboCampoNorte,      ['norte','north','y','coord_y'])
        self._sel(self.cboCampoLado,       ['lado','side','segment','tramo'])

        # Actualizar combo de filtro atlas
        if hasattr(self, 'cboCampoFiltro'):
            self.cboCampoFiltro.clear()
            self.cboCampoFiltro.addItem('-- Mismo campo que identificador --', None)
            for c in campos: self.cboCampoFiltro.addItem(c, c)
            self._sel(self.cboCampoFiltro, ['layer','nombre','name','predio','parcela','id'])

    def actualizar_campos_lineas(self):
        pass  # Líneas es opcional, no se usan para datos

    def _actualizar_combos_atlas(self):
        """Actualiza el combo del campo atlas con los campos del polígono."""
        layer_id = self.cboPoligonos.currentData()
        if not layer_id or not hasattr(self, 'cboCampoAtlas'): return
        layer = QgsProject.instance().mapLayer(layer_id)
        if not layer: return
        campos = [f.name() for f in layer.fields()]
        self.cboCampoAtlas.clear()
        self.cboCampoAtlas.addItem('-- Seleccione campo identificador --', None)
        for c in campos: self.cboCampoAtlas.addItem(c, c)
        # Auto-selección para campos típicos de atlas
        self._sel(self.cboCampoAtlas, ['layer','nombre','name','predio','parcela','id_predio','codigo'])

    def _autodetectar_todos(self):
        self.actualizar_campos_poligono()
        self.actualizar_campos_puntos()
        QtWidgets.QMessageBox.information(
            self, "Auto-detección completada",
            "Los campos han sido detectados automáticamente.\n"
            "Puede modificarlos si es necesario.")

    def _sel(self, combo, nombres):
        for i in range(1, combo.count()):
            d = combo.itemData(i)
            if d and d.lower() in [n.lower() for n in nombres]:
                combo.setCurrentIndex(i); return

    # =========================================================================
    # PREVISUALIZACIÓN ATLAS
    # =========================================================================

    def _previsualizar_atlas(self):
        """Muestra un resumen de los predios que se procesarán en modo atlas."""
        layer_id = self.cboPoligonos.currentData()
        if not layer_id:
            QtWidgets.QMessageBox.warning(self, "Atlas", "Seleccione primero una capa de polígonos.")
            return
        layer = QgsProject.instance().mapLayer(layer_id)
        if not layer:
            QtWidgets.QMessageBox.warning(self, "Atlas", "No se pudo cargar la capa.")
            return

        campo_atlas = self.cboCampoAtlas.currentData()
        solo_sel    = self.chkSoloSeleccionados.isChecked()

        if solo_sel:
            features = list(layer.selectedFeatures())
        else:
            features = list(layer.getFeatures())

        if not features:
            self.lblAtlasResumen.setText("⚠ No hay objetos para procesar.")
            self.lblAtlasResumen.setVisible(True)
            return

        nombres_predios = []
        for feat in features:
            if campo_atlas and campo_atlas in [f.name() for f in feat.fields()]:
                val = feat[campo_atlas]
                nombres_predios.append(str(val) if val else "(sin nombre)")
            else:
                nombres_predios.append("FID={}".format(feat.id()))

        resumen = ("<b>Se generarán {} memorias descriptivas:</b><br>".format(len(nombres_predios)) +
                   "<br>".join("• {}".format(n) for n in nombres_predios[:15]) +
                   ("<br><i>... y {} más</i>".format(len(nombres_predios)-15) if len(nombres_predios)>15 else ""))

        self.lblAtlasResumen.setText(resumen)
        self.lblAtlasResumen.setVisible(True)

    # =========================================================================
    # OBTENER DATOS DEL FORMULARIO
    # =========================================================================

    def obtener_datos_formulario(self):
        return {
            'solicitante': {
                'nombre': self.txtNombre.text().strip(),
                'dni':    self.txtDNI.text().strip()
            },
            'ubicacion': {
                'sector':       self.txtSector.text().strip(),
                'zona':         self.txtZona.text().strip(),
                'distrito':     self.txtDistrito.text().strip(),
                'provincia':    self.txtProvincia.text().strip(),
                'departamento': self.txtDepartamento.text().strip()
            },
            'capas': {
                'poligono_id': self.cboPoligonos.currentData(),
                'punto_id':    self.cboPuntos.currentData(),
                'linea_id':    self.cboLineas.currentData()
            },
            'colindantes': {
                'detectar_automatico': self.chkDetectarColindantes.isChecked(),
                'manual': {
                    'NORTE': self.txtNorte.text().strip(),
                    'SUR':   self.txtSur.text().strip(),
                    'ESTE':  self.txtEste.text().strip(),
                    'OESTE': self.txtOeste.text().strip()
                }
            },
            'opciones': {
                'calcular_automatico': self.chkCalcularAutomatico.isChecked(),
                'incluir_mapa':  self.chkIncluirMapa.isChecked(),
                'incluir_tabla': self.chkIncluirTabla.isChecked(),
                'texto_default': self.chkTextoDefault.isChecked()
            },
            'generalidades':    self.txtGeneralidades.toPlainText().strip(),
            'info_mapa': {
                'Sistema de coordenadas': self.txtSistema.text().strip(),
                'Unidades':   self.txtUnidades.text().strip(),
                'Elipsoide':  self.txtElipsoide.text().strip(),
                'Grillado':   self.txtGrillado.text().strip()
            },
            'campos': {
                'orden':       self.cboCampoOrden.currentData(),
                'id':          self.cboCampoID.currentData(),
                'nombre':      self.cboCampoNombre.currentData(),
                'dni':         self.cboCampoDNI.currentData(),
                'vertice_id':  self.cboCampoVerticeID.currentData(),
                'orden_punto': self.cboCampoOrdenPunto.currentData(),
                'distancia':   self.cboCampoDistancia.currentData(),
                'azimut':      self.cboCampoAzimut.currentData(),
                'este':        self.cboCampoEste.currentData(),
                'norte':       self.cboCampoNorte.currentData(),
                'lado':        self.cboCampoLado.currentData(),
                'area':        self.cboCampoArea.currentData(),
                'perimetro':   self.cboCampoPerimetro.currentData()
            },
            # Atlas
            'atlas': {
                'activo':          self.chkModoAtlas.isChecked(),
                'campo_atlas':     self.cboCampoAtlas.currentData() if hasattr(self,'cboCampoAtlas') else None,
                'campo_filtro':    self.cboCampoFiltro.currentData() if hasattr(self,'cboCampoFiltro') else None,
                'solo_seleccion':  self.chkSoloSeleccionados.isChecked() if hasattr(self,'chkSoloSeleccionados') else False,
                'mostrar_fuente_area': self.chkMostrarFuenteArea.isChecked() if hasattr(self,'chkMostrarFuenteArea') else True,
            },
            'output_file': self.txtOutputFile.text().strip()
        }

    # =========================================================================
    # VALIDACIÓN
    # =========================================================================

    def validar_formulario(self):
        def _warn(msg, tab=0, w=None):
            self.tabWidget.setCurrentIndex(tab)
            if w: w.setFocus()
            QtWidgets.QMessageBox.warning(self, "Campo requerido", msg)
            return False

        if not self.txtNombre.text().strip():
            return _warn("Ingrese el nombre del solicitante.", 0, self.txtNombre)
        if not self.txtDNI.text().strip():
            return _warn("Ingrese el DNI del solicitante.", 0, self.txtDNI)
        if not self.txtSector.text().strip():
            return _warn("Ingrese el sector o localidad de ubicación.", 0, self.txtSector)
        if not self.cboPoligonos.currentData():
            return _warn("Seleccione una capa de polígonos.", 0)
        if not self.cboPuntos.currentData():
            return _warn("Seleccione una capa de puntos.", 0)
        if not self.txtOutputFile.text().strip():
            return _warn("Especifique la ruta del archivo de salida (.docx).", 0, self.txtOutputFile)

        # Modo atlas: campo identificador obligatorio
        if self.chkModoAtlas.isChecked():
            if not self.cboCampoAtlas.currentData():
                return _warn(
                    "En modo Atlas debe seleccionar el campo identificador del predio.",
                    1, self.cboCampoAtlas)

        # Colindantes manuales
        if not self.chkDetectarColindantes.isChecked():
            for nombre, widget in [('Norte', self.txtNorte), ('Sur', self.txtSur),
                                   ('Este', self.txtEste), ('Oeste', self.txtOeste)]:
                if not widget.text().strip():
                    return _warn("Ingrese el colindante {}.".format(nombre), 2, widget)

        return True
