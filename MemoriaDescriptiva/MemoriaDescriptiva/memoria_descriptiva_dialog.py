# -*- coding: utf-8 -*-
"""
Diálogo Memoria Descriptiva v3.1

LÓGICA DE DATOS DEL SOLICITANTE:
─────────────────────────────────
  MODO ÚNICO:
    • groupSolicitante ACTIVO — el usuario escribe nombre y DNI manualmente
    • Esos datos van a TODAS las secciones del documento

  MODO ATLAS (completo o selección):
    • groupSolicitante DESACTIVADO (se atenúa)
    • Aparecen dos combos para elegir el CAMPO de la capa de polígonos
      que contiene el nombre y DNI de cada propietario
    • En cada iteración se lee el valor del campo para ESE polígono
    • Los datos comunes (ubicación, datum, generalidades, colindantes)
      se ingresan UNA sola vez y se repiten en todas las memorias

DATOS QUE ITERAN con cada polígono (atlas):
  - Nombre del propietario    ← campo elegido en la BD
  - DNI del propietario       ← campo elegido en la BD
  - Nombre del predio/archivo ← campo elegido en la BD

DATOS QUE SE REPITEN en todas las memorias (atlas):
  - Ubicación (sector, distrito, provincia, departamento, zona UTM)
  - Generalidades
  - Info del mapa (datum, elipsoide, grillado)
  - Colindantes (manual o auto-detectados)
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

        self._crear_panel_atlas_solicitante()  # Panel que reemplaza al solicitante en atlas
        self._crear_tab_modo()                 # Pestaña "Modo de Trabajo"
        self._crear_tab_campos()               # Pestaña "Campos"

        # ── Conexiones ──────────────────────────────────────────────────────
        self.chkDetectarColindantes.toggled.connect(
            lambda c: self.groupColindantesManual.setEnabled(not c))
        self.chkTextoDefault.toggled.connect(
            lambda c: self.txtGeneralidades.setReadOnly(c))

        self.cboPoligonos.currentIndexChanged.connect(self.actualizar_campos_poligono)
        self.cboPuntos.currentIndexChanged.connect(self.actualizar_campos_puntos)

        # Cambio de modo → actualizar UI
        self.rbUnico.toggled.connect(self._on_modo_changed)
        self.rbAtlasCompleto.toggled.connect(self._on_modo_changed)
        self.rbAtlasSeleccion.toggled.connect(self._on_modo_changed)

        # Estado inicial
        self.groupColindantesManual.setEnabled(not self.chkDetectarColindantes.isChecked())
        self.txtGeneralidades.setReadOnly(self.chkTextoDefault.isChecked())
        self._on_modo_changed()  # Aplica estado inicial según modo seleccionado

    # =========================================================================
    # PANEL ATLAS-SOLICITANTE (se inserta debajo del groupSolicitante)
    # =========================================================================

    def _crear_panel_atlas_solicitante(self):
        """
        Crea un panel que aparece SOLO en modo atlas, debajo del groupSolicitante,
        con combos para elegir el campo de nombre y DNI de la capa de polígonos.
        """
        # Encontrar el layout de la pestaña Datos Básicos
        tab_layout = self.tabDatosBasicos.layout()

        # Panel contenedor
        self.panelAtlasSolicitante = QtWidgets.QGroupBox(
            "📋  Datos del Solicitante — Modo Atlas")
        self.panelAtlasSolicitante.setStyleSheet(
            "QGroupBox { font-weight: bold; color: #17375E; "
            "background: #EAF4EA; border: 2px solid #2E6E3E; "
            "border-radius: 6px; margin-top: 8px; padding-top: 8px; } "
            "QGroupBox::title { subcontrol-origin: margin; left: 10px; }")

        form = QtWidgets.QFormLayout()
        form.setSpacing(8)

        # Nota explicativa
        lbl_nota = QtWidgets.QLabel(
            "<i>En modo Atlas, el nombre y DNI se leen automáticamente<br>"
            "del campo de la capa de polígonos para cada predio.</i>")
        lbl_nota.setWordWrap(True)
        lbl_nota.setStyleSheet("color: #2E6E3E; font-size: 10px; padding: 4px;")
        form.addRow(lbl_nota)

        # Combo: campo nombre/propietario
        self.cboAtlasCampoNombre = QtWidgets.QComboBox()
        self.cboAtlasCampoNombre.setToolTip(
            "Campo de la capa de polígonos con el nombre del propietario.\n"
            "Ejemplo: NombresApellidos, nombre, titular")
        form.addRow("Campo Nombre / Propietario:", self.cboAtlasCampoNombre)

        # Combo: campo DNI
        self.cboAtlasCampoDNI = QtWidgets.QComboBox()
        self.cboAtlasCampoDNI.setToolTip(
            "Campo con el DNI o documento de identidad.\n"
            "Déjalo en '-- Sin DNI --' si no existe en la capa.")
        form.addRow("Campo DNI:", self.cboAtlasCampoDNI)

        # Vista previa del primer predio
        self.lblAtlasPreviewPredio = QtWidgets.QLabel("")
        self.lblAtlasPreviewPredio.setStyleSheet(
            "color: #555; font-size: 10px; padding: 2px 4px;")
        self.cboAtlasCampoNombre.currentIndexChanged.connect(self._actualizar_preview_predio)
        form.addRow("Vista previa:", self.lblAtlasPreviewPredio)

        self.panelAtlasSolicitante.setLayout(form)

        # Insertar en la posición 0 (encima del groupSolicitante) en el tab
        tab_layout.insertWidget(0, self.panelAtlasSolicitante)
        self.panelAtlasSolicitante.setVisible(False)  # Oculto hasta activar atlas

    def _actualizar_preview_predio(self):
        """Muestra un ejemplo con el valor del primer polígono."""
        lid = self.cboPoligonos.currentData()
        if not lid: return
        layer = QgsProject.instance().mapLayer(lid)
        if not layer: return
        campo = self.cboAtlasCampoNombre.currentData()
        if not campo: return
        feats = list(layer.getFeatures())
        if not feats: return
        try:
            val = feats[0][campo]
            self.lblAtlasPreviewPredio.setText(
                "1er predio: <b>{}</b>".format(str(val) if val else '(vacío)'))
        except Exception:
            self.lblAtlasPreviewPredio.setText("")

    # =========================================================================
    # PESTAÑA MODO DE TRABAJO
    # =========================================================================

    def _crear_tab_modo(self):
        self.tabModo = QtWidgets.QWidget()
        self.tabWidget.insertTab(1, self.tabModo, "📋 Modo de Trabajo")

        scroll = QtWidgets.QScrollArea(); scroll.setWidgetResizable(True)
        cont = QtWidgets.QWidget(); layout = QtWidgets.QVBoxLayout(cont)
        scroll.setWidget(cont)
        outer = QtWidgets.QVBoxLayout(self.tabModo)
        outer.setContentsMargins(0, 0, 0, 0); outer.addWidget(scroll)

        # ── Selector de modo ──────────────────────────────────────────────────
        grp_modo = QtWidgets.QGroupBox("¿Cuántas memorias quieres generar?")
        grp_modo.setStyleSheet("QGroupBox { font-weight: bold; }")
        vl = QtWidgets.QVBoxLayout()

        self.rbUnico = QtWidgets.QRadioButton(
            "🔵  Polígono único  —  1 memoria para el polígono seleccionado")
        self.rbAtlasCompleto = QtWidgets.QRadioButton(
            "🟢  Atlas completo  —  1 memoria por CADA polígono de la capa")
        self.rbAtlasSeleccion = QtWidgets.QRadioButton(
            "🟡  Atlas selección  —  1 memoria por cada polígono SELECCIONADO en QGIS")

        self.rbUnico.setChecked(True)
        for rb in [self.rbUnico, self.rbAtlasCompleto, self.rbAtlasSeleccion]:
            rb.setStyleSheet("font-size: 11px; padding: 5px;")
            vl.addWidget(rb)

        grp_modo.setLayout(vl); layout.addWidget(grp_modo)

        # ── Relación polígono ↔ puntos ────────────────────────────────────────
        grp_rel = QtWidgets.QGroupBox(
            "Relación entre capas  (¿qué campo vincula polígonos con puntos?)")
        fl = QtWidgets.QFormLayout()

        lbl = QtWidgets.QLabel(
            "<b>El plugin filtra los puntos de cada polígono usando un campo compartido.</b><br>"
            "<i>Ejemplo con tus capas:</i><br>"
            "• <b>AREA_TOTAL</b>: campo <b>fid</b> = 1, 2, 3, 4, 5, 6<br>"
            "• <b>Puntos</b>: campo <b>ID_Poligono</b> = 1,1,1... 2,2,2... 3,3,3...")
        lbl.setWordWrap(True)
        lbl.setStyleSheet(
            "background: #E8F4FD; padding: 8px; border-radius: 4px; "
            "border: 1px solid #AED6F1; font-size: 10px;")
        fl.addRow(lbl)

        self.cboCampoIdPoligono = QtWidgets.QComboBox()
        self.cboCampoIdPoligono.setToolTip("Campo ID único de cada polígono (ej: fid)")
        fl.addRow("Campo ID del polígono:", self.cboCampoIdPoligono)

        self.cboCampoRelPuntos = QtWidgets.QComboBox()
        self.cboCampoRelPuntos.setToolTip("Campo en la capa de Puntos que contiene el ID del polígono (ej: ID_Poligono)")
        fl.addRow("Campo relación en Puntos:", self.cboCampoRelPuntos)

        grp_rel.setLayout(fl); layout.addWidget(grp_rel)

        # ── Previsualización ──────────────────────────────────────────────────
        self.grpPreview = QtWidgets.QGroupBox("Previsualización del Atlas")
        vl2 = QtWidgets.QVBoxLayout()
        self.lblPreview = QtWidgets.QLabel(
            "<i>Activa modo Atlas para ver los predios que se procesarán.</i>")
        self.lblPreview.setWordWrap(True)
        self.lblPreview.setStyleSheet("color: #666; padding: 4px; font-size: 10px;")
        self.btnPreview = QtWidgets.QPushButton("👁  Ver predios a procesar")
        self.btnPreview.clicked.connect(self._previsualizar)
        self.btnPreview.setEnabled(False)
        vl2.addWidget(self.lblPreview); vl2.addWidget(self.btnPreview)
        self.grpPreview.setLayout(vl2); layout.addWidget(self.grpPreview)

        layout.addStretch()

    # =========================================================================
    # PESTAÑA CAMPOS
    # =========================================================================

    def _crear_tab_campos(self):
        self.tabCampos = QtWidgets.QWidget()
        self.tabWidget.addTab(self.tabCampos, "⚙ Campos")

        scroll = QtWidgets.QScrollArea(); scroll.setWidgetResizable(True)
        cont = QtWidgets.QWidget(); layout = QtWidgets.QVBoxLayout(cont)
        scroll.setWidget(cont)
        outer = QtWidgets.QVBoxLayout(self.tabCampos)
        outer.setContentsMargins(0, 0, 0, 0); outer.addWidget(scroll)

        lbl = QtWidgets.QLabel(
            "<b>Campos de datos:</b> deja <i>-- Automático --</i> para detección automática.<br>"
            "Distancias y azimuts se calculan geométricamente si no existen como atributos.")
        lbl.setWordWrap(True)
        lbl.setStyleSheet("background: #EAF4EA; padding: 8px; border-radius: 4px;")
        layout.addWidget(lbl)

        # Campos en Puntos
        grp_p = QtWidgets.QGroupBox("Campos en la Capa de Puntos")
        fp = QtWidgets.QFormLayout()
        self.cboCampoVerticeID  = QtWidgets.QComboBox()
        self.cboCampoOrdenPunto = QtWidgets.QComboBox()
        self.cboCampoDistancia  = QtWidgets.QComboBox()
        self.cboCampoAzimut     = QtWidgets.QComboBox()
        self.cboCampoEste       = QtWidgets.QComboBox()
        self.cboCampoNorte      = QtWidgets.QComboBox()
        self.cboCampoLado       = QtWidgets.QComboBox()
        fp.addRow("ID / Etiqueta vértice:",   self.cboCampoVerticeID)
        fp.addRow("Orden / Secuencia:",        self.cboCampoOrdenPunto)
        fp.addRow("Distancia (m):",            self.cboCampoDistancia)
        fp.addRow("Azimut (°):",               self.cboCampoAzimut)
        fp.addRow("Coordenada Este / X:",      self.cboCampoEste)
        fp.addRow("Coordenada Norte / Y:",     self.cboCampoNorte)
        fp.addRow("Nombre del Lado:",          self.cboCampoLado)
        grp_p.setLayout(fp); layout.addWidget(grp_p)

        # Campos en Polígonos
        grp_pol = QtWidgets.QGroupBox("Campos en la Capa de Polígonos")
        fpol = QtWidgets.QFormLayout()
        self.cboCampoArea      = QtWidgets.QComboBox()
        self.cboCampoPerimetro = QtWidgets.QComboBox()
        fpol.addRow("Área (ha):",        self.cboCampoArea)
        fpol.addRow("Perímetro (m):",    self.cboCampoPerimetro)
        grp_pol.setLayout(fpol); layout.addWidget(grp_pol)

        btn = QtWidgets.QPushButton("🔍  Auto-detectar todos los campos")
        btn.setStyleSheet("padding: 6px; font-weight: bold;")
        btn.clicked.connect(self._autodetectar)
        layout.addWidget(btn)
        layout.addStretch()

    # =========================================================================
    # CONTROL DE MODO — lógica central de la UI
    # =========================================================================

    def _on_modo_changed(self):
        """
        Cuando cambia el modo:
        - ÚNICO:  groupSolicitante activo, panelAtlasSolicitante oculto
        - ATLAS:  groupSolicitante desactivado+atenuado, panelAtlasSolicitante visible
        """
        es_atlas = self.rbAtlasCompleto.isChecked() or self.rbAtlasSeleccion.isChecked()

        # Solicitante manual: activo solo en modo único
        self.groupSolicitante.setEnabled(not es_atlas)
        self.groupSolicitante.setStyleSheet(
            "" if not es_atlas else
            "QGroupBox { color: #999; } QGroupBox * { color: #999; }"
        )

        # Panel atlas solicitante: visible solo en atlas
        self.panelAtlasSolicitante.setVisible(es_atlas)

        # Botón previsualizar
        self.btnPreview.setEnabled(es_atlas)
        if not es_atlas:
            self.lblPreview.setText(
                "<i>Activa modo Atlas para ver los predios que se procesarán.</i>")

        # Actualizar título del groupSolicitante para que sea claro
        if es_atlas:
            self.groupSolicitante.setTitle(
                "1. Datos del Solicitante  ⚠ (desactivado en modo Atlas — se lee de la BD)")
        else:
            self.groupSolicitante.setTitle("1. Datos del Solicitante")

    # =========================================================================
    # ACTUALIZACIÓN DE CAPAS Y CAMPOS
    # =========================================================================

    def actualizar_campos_poligono(self):
        lid = self.cboPoligonos.currentData()
        if not lid: return
        layer = QgsProject.instance().mapLayer(lid)
        if not layer: return
        campos = [f.name() for f in layer.fields()]

        # Combos del tab principal (modo único)
        for cbo in [self.cboCampoNombre, self.cboCampoDNI]:
            cbo.clear(); cbo.addItem('-- Automático --', None)
            for c in campos: cbo.addItem(c, c)

        # Combos del panel atlas solicitante
        self.cboAtlasCampoNombre.clear()
        self.cboAtlasCampoNombre.addItem('-- Seleccione campo --', None)
        for c in campos: self.cboAtlasCampoNombre.addItem(c, c)

        self.cboAtlasCampoDNI.clear()
        self.cboAtlasCampoDNI.addItem('-- Sin DNI --', None)
        for c in campos: self.cboAtlasCampoDNI.addItem(c, c)

        # Combos de campos de polígono (pestaña Campos)
        for cbo in [self.cboCampoArea, self.cboCampoPerimetro]:
            cbo.clear(); cbo.addItem('-- Calcular automáticamente --', None)
            for c in campos: cbo.addItem(c, c)

        # Campo ID del polígono (pestaña Modo)
        if hasattr(self, 'cboCampoIdPoligono'):
            self.cboCampoIdPoligono.clear()
            self.cboCampoIdPoligono.addItem('-- Seleccione --', None)
            for c in campos: self.cboCampoIdPoligono.addItem(c, c)

        # Auto-selección
        self._sel(self.cboCampoNombre,        ['NombresApellidos','nombre','nom_tit','propietario','titular'])
        self._sel(self.cboCampoDNI,           ['dni','DNI','doc','documento'])
        self._sel(self.cboAtlasCampoNombre,   ['NombresApellidos','nombre','nom_tit','propietario','titular'])
        self._sel(self.cboAtlasCampoDNI,      ['dni','DNI','doc','documento'])
        self._sel(self.cboCampoArea,          ['Area_ha','area_ha','AREA_HA','area','AREA','hectareas'])
        self._sel(self.cboCampoPerimetro,     ['Perímetro','Perimetro','PERIMETRO','perimetro','perimeter'])
        self._sel(self.cboCampoIdPoligono,    ['fid','FID','id','ID','objectid'])

        self._actualizar_preview_predio()

    def actualizar_campos_puntos(self):
        lid = self.cboPuntos.currentData()
        if not lid: return
        layer = QgsProject.instance().mapLayer(lid)
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

        # Campo de relación en puntos
        if hasattr(self, 'cboCampoRelPuntos'):
            self.cboCampoRelPuntos.clear()
            self.cboCampoRelPuntos.addItem('-- Seleccione --', None)
            for c in campos: self.cboCampoRelPuntos.addItem(c, c)
            self._sel(self.cboCampoRelPuntos,
                      ['ID_Poligono','id_poligono','fid_area','poligono_id','pol_id'])

        self._sel(self.cboCampoOrden,      ['ID_Vertice','id_vertice','orden','order','id','fid'])
        self._sel(self.cboCampoID,         ['ID_Vertice','id_vertice','fid','id'])
        self._sel(self.cboCampoVerticeID,  ['ID_Vertice','id_vertice','vertice','id'])
        self._sel(self.cboCampoOrdenPunto, ['ID_Vertice','id_vertice','orden','order','fid'])
        self._sel(self.cboCampoDistancia,  ['Distancia','distancia','DISTANCIA','distance','dist'])
        self._sel(self.cboCampoAzimut,     ['Azimut','azimut','AZIMUT','azimuth','rumbo'])
        self._sel(self.cboCampoEste,       ['Este','este','ESTE','X','x','coord_x'])
        self._sel(self.cboCampoNorte,      ['Norte','norte','NORTE','Y','y','coord_y'])
        self._sel(self.cboCampoLado,       ['LADO','Lado','lado','side','segment'])

    def actualizar_campos_lineas(self): pass

    def _autodetectar(self):
        self.actualizar_campos_poligono()
        self.actualizar_campos_puntos()
        QtWidgets.QMessageBox.information(
            self, "Auto-detección completada",
            "Los campos han sido detectados automáticamente.\n"
            "Revisa los valores y ajusta si es necesario.")

    def _sel(self, combo, nombres):
        """Selecciona el primer campo que coincida con la lista de nombres."""
        for i in range(1, combo.count()):
            d = combo.itemData(i)
            if d and d.lower() in [n.lower() for n in nombres]:
                combo.setCurrentIndex(i); return

    # =========================================================================
    # PREVISUALIZACIÓN ATLAS
    # =========================================================================

    def _previsualizar(self):
        lid = self.cboPoligonos.currentData()
        if not lid:
            QtWidgets.QMessageBox.warning(self,"Atlas","Selecciona una capa de polígonos."); return
        layer = QgsProject.instance().mapLayer(lid)
        if not layer: return

        feats = (list(layer.selectedFeatures())
                 if self.rbAtlasSeleccion.isChecked()
                 else list(layer.getFeatures()))

        if not feats:
            self.lblPreview.setText("<b style='color:red'>No hay polígonos para procesar.</b>"); return

        campo_nombre = self.cboAtlasCampoNombre.currentData()
        campo_id     = self.cboCampoIdPoligono.currentData()
        fnames       = [f.name() for f in feats[0].fields()]

        filas = []
        for feat in feats:
            nombre = ''
            if campo_nombre and campo_nombre in fnames:
                nombre = str(feat[campo_nombre] or '')
            if not nombre and campo_id and campo_id in fnames:
                nombre = 'ID={}'.format(feat[campo_id])
            if not nombre:
                nombre = 'FID={}'.format(feat.id())
            filas.append(nombre)

        modo = "seleccionados" if self.rbAtlasSeleccion.isChecked() else "totales"
        resumen = ("<b>{} predios {} → {} memorias a generar:</b><br>".format(
                   len(filas), modo, len(filas)) +
                   "<br>".join("✓ {}".format(n) for n in filas[:15]) +
                   ("<br><i>... y {} más</i>".format(len(filas)-15) if len(filas)>15 else ""))
        self.lblPreview.setText(resumen)

    # =========================================================================
    # OBTENER DATOS DEL FORMULARIO
    # =========================================================================

    def obtener_datos_formulario(self):
        es_atlas = self.rbAtlasCompleto.isChecked() or self.rbAtlasSeleccion.isChecked()

        return {
            # ── Datos del solicitante ──────────────────────────────────────────
            # En modo único: valores manuales del formulario
            # En modo atlas: None (se leerán de la BD en cada iteración)
            'solicitante': {
                'nombre': self.txtNombre.text().strip() if not es_atlas else None,
                'dni':    self.txtDNI.text().strip()    if not es_atlas else None,
            },

            # ── Campos BD para solicitante en modo atlas ───────────────────────
            'atlas_solicitante': {
                'campo_nombre': self.cboAtlasCampoNombre.currentData(),
                'campo_dni':    self.cboAtlasCampoDNI.currentData(),
            },

            # ── Datos comunes (se repiten en todas las memorias) ───────────────
            'ubicacion': {
                'sector':       self.txtSector.text().strip(),
                'zona':         self.txtZona.text().strip(),
                'distrito':     self.txtDistrito.text().strip(),
                'provincia':    self.txtProvincia.text().strip(),
                'departamento': self.txtDepartamento.text().strip()
            },
            'generalidades':    self.txtGeneralidades.toPlainText().strip(),
            'info_mapa': {
                'Sistema de coordenadas': self.txtSistema.text().strip(),
                'Unidades':   self.txtUnidades.text().strip(),
                'Elipsoide':  self.txtElipsoide.text().strip(),
                'Grillado':   self.txtGrillado.text().strip()
            },
            'colindantes': {
                'detectar_automatico': self.chkDetectarColindantes.isChecked(),
                'manual': {
                    'NORTE': self.txtNorte.text().strip(), 'SUR':   self.txtSur.text().strip(),
                    'ESTE':  self.txtEste.text().strip(),  'OESTE': self.txtOeste.text().strip()
                }
            },

            # ── Capas ──────────────────────────────────────────────────────────
            'capas': {
                'poligono_id': self.cboPoligonos.currentData(),
                'punto_id':    self.cboPuntos.currentData(),
                'linea_id':    self.cboLineas.currentData()
            },

            # ── Modo ───────────────────────────────────────────────────────────
            'modo': ('unico'           if self.rbUnico.isChecked()
                     else 'atlas_seleccion' if self.rbAtlasSeleccion.isChecked()
                     else 'atlas_completo'),

            # ── Relación entre capas ───────────────────────────────────────────
            'relacion': {
                'campo_id_poligono': self.cboCampoIdPoligono.currentData(),
                'campo_rel_puntos':  self.cboCampoRelPuntos.currentData(),
            },

            # ── Configuración de campos ────────────────────────────────────────
            'campos': {
                'campo_id_poligono': self.cboCampoIdPoligono.currentData(),
                'campo_id_rel_pts':  self.cboCampoRelPuntos.currentData(),
                'campo_vertice':     self.cboCampoVerticeID.currentData(),
                'campo_orden':       self.cboCampoOrdenPunto.currentData(),
                'campo_distancia':   self.cboCampoDistancia.currentData(),
                'campo_azimut':      self.cboCampoAzimut.currentData(),
                'campo_este':        self.cboCampoEste.currentData(),
                'campo_norte':       self.cboCampoNorte.currentData(),
                'campo_lado':        self.cboCampoLado.currentData(),
                'campo_area':        self.cboCampoArea.currentData(),
                'campo_perimetro':   self.cboCampoPerimetro.currentData(),
                # Solo usados en modo único
                'campo_nombre':      self.cboCampoNombre.currentData(),
                'campo_dni':         self.cboCampoDNI.currentData(),
            },
            'output_file': self.txtOutputFile.text().strip()
        }

    # =========================================================================
    # VALIDACIÓN
    # =========================================================================

    def validar_formulario(self):
        def _w(msg, tab=0, w=None):
            self.tabWidget.setCurrentIndex(tab)
            if w: w.setFocus()
            QtWidgets.QMessageBox.warning(self, "Campo requerido", msg)
            return False

        es_atlas = self.rbAtlasCompleto.isChecked() or self.rbAtlasSeleccion.isChecked()

        # En modo único: nombre y DNI son obligatorios
        if not es_atlas:
            if not self.txtNombre.text().strip():
                return _w("Ingrese el nombre del solicitante.", 0, self.txtNombre)
            if not self.txtDNI.text().strip():
                return _w("Ingrese el DNI del solicitante.", 0, self.txtDNI)

        # En modo atlas: campo nombre obligatorio
        if es_atlas:
            if not self.cboAtlasCampoNombre.currentData():
                return _w(
                    "En modo Atlas debes seleccionar el campo que contiene\n"
                    "el nombre del propietario en la capa de polígonos.\n\n"
                    "Revisa la pestaña 'Datos Básicos'.", 0, self.cboAtlasCampoNombre)

        # Ubicación mínima
        if not self.txtSector.text().strip():
            return _w("Ingresa el sector o localidad de ubicación.", 0, self.txtSector)

        # Capas
        if not self.cboPoligonos.currentData():
            return _w("Selecciona la capa de polígonos (ej: AREA_TOTAL).", 0)
        if not self.cboPuntos.currentData():
            return _w("Selecciona la capa de puntos (ej: Puntos).", 0)

        # Archivo de salida
        if not self.txtOutputFile.text().strip():
            return _w("Especifica la ruta del archivo de salida.", 0, self.txtOutputFile)

        # Atlas selección: verificar que haya objetos seleccionados
        if self.rbAtlasSeleccion.isChecked():
            lid = self.cboPoligonos.currentData()
            if lid:
                layer = QgsProject.instance().mapLayer(lid)
                if layer and layer.selectedFeatureCount() == 0:
                    return _w(
                        "No hay polígonos seleccionados en la capa.\n"
                        "Selecciona al menos un polígono en QGIS antes de continuar.", 1)

        # Colindantes manuales
        if not self.chkDetectarColindantes.isChecked():
            for nm, wid in [('Norte', self.txtNorte), ('Sur', self.txtSur),
                             ('Este', self.txtEste), ('Oeste', self.txtOeste)]:
                if not wid.text().strip():
                    return _w("Ingresa el colindante {}.".format(nm), 2, wid)

        return True
