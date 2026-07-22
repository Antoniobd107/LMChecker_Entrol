#!/usr/bin/env python3
"""
Comprobador de LM - version PyQt6 + Fluent Widgets
=====================================================
Misma logica que las versiones anteriores (lm_reglas.py, sin cambios), pero
con una interfaz nativa hecha con PyQt6 y la libreria PyQt6-Fluent-Widgets
(aspecto "Fluent Design" de Windows 11: tarjetas, colores de marca, controles
redondeados) - SIN depender de WebView2 ni de ningun navegador embebido, por
los problemas de estabilidad que tuvimos con esa version.

Las actualizaciones desde hilos en segundo plano (cargar catalogo, analizar)
se hacen con señales de Qt (pyqtSignal/pyqtSlot), el mecanismo nativo y
robusto de Qt para esto - mucho mas fiable que un puente JavaScript.

IMPORTANTE - ANTES DE EMPAQUETAR:
Edita la constante RUTA_CATALOGO_DEFECTO mas abajo con la ruta de red real
del catalogo master compartido del departamento.

Archivos necesarios en la misma carpeta:
    - lm_checker_qt.py   (este archivo)
    - lm_reglas.py       (logica de negocio, sin cambios)
    - logo_entrol.png    (logo de la empresa, opcional - si no esta, se omite)

Diseñado por ACC
"""

import os
import sys
import json
import threading
import traceback
from pathlib import Path
from datetime import datetime

import pandas as pd

from PyQt6.QtCore import Qt, QObject, pyqtSignal, QUrl
from PyQt6.QtGui import QPixmap, QDesktopServices, QIcon
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QFileDialog, QMessageBox,
    QScrollArea, QFrame,
)

from qfluentwidgets import (
    setTheme, Theme, setThemeColor,
    CardWidget, TitleLabel, BodyLabel, CaptionLabel, StrongBodyLabel,
    PrimaryPushButton, PushButton, ProgressBar, RadioButton, CheckBox,
    TextEdit, InfoBar, InfoBarPosition,
)

from lm_reglas import (
    cargar_catalogo, extraer_tabla_odt, analizar_lm,
    aplicar_reglas_especiales, generar_excel,
    extraer_cabecera_modelo, analizar_referencias_modelo, analizar_plates,
    VISUAL_TIPO_A, VISUAL_TIPO_B, VISUAL_TIPO_OTRO,
)

# ============================================================
# CONFIGURACION - EDITAR ANTES DE DISTRIBUIR
# ============================================================
RUTA_CATALOGO_DEFECTO = r"Z:\Materiales\BUSCADOR DE PIEZAS - MASTER.ods"
NOMBRE_APP = "Comprobador de LM"
COLOR_MARCA = "#00a8e0"  # cian de Entrol

_CARPETA_APP_LOCAL = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "ComprobadorLM"
CARPETA_SALIDA_DEFECTO = _CARPETA_APP_LOCAL / "Informes"
RUTA_CONFIG = _CARPETA_APP_LOCAL / "config.json"
RUTA_LOG_ERRORES = _CARPETA_APP_LOCAL / "error.log"
# ============================================================


def registrar_error(contexto: str, excepcion: Exception):
    try:
        _CARPETA_APP_LOCAL.mkdir(parents=True, exist_ok=True)
        with open(RUTA_LOG_ERRORES, "a", encoding="utf-8") as f:
            f.write(f"\n[{datetime.now()}] {contexto}\n")
            f.write(traceback.format_exc())
            f.write("\n" + "-" * 60 + "\n")
    except Exception:
        pass


def cargar_config() -> dict:
    try:
        with open(RUTA_CONFIG, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def guardar_config(actualizaciones: dict):
    cfg = cargar_config()
    cfg.update(actualizaciones)
    try:
        _CARPETA_APP_LOCAL.mkdir(parents=True, exist_ok=True)
        with open(RUTA_CONFIG, "w", encoding="utf-8") as f:
            json.dump(cfg, f)
    except Exception:
        pass


def resource_path(rel: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


# ============================================================
# Señales para comunicar los hilos en segundo plano con la ventana
# ============================================================

class Senales(QObject):
    catalogo_listo = pyqtSignal(dict)
    progreso = pyqtSignal(int)
    analisis_listo = pyqtSignal(dict)


# ============================================================
# Ventana principal
# ============================================================

class VentanaPrincipal(QWidget):
    def __init__(self):
        super().__init__()
        self.catalogo = None
        self.ruta_lm = None
        self.ruta_xlsx = None
        self.ruta_csv = None
        cfg = cargar_config()
        self.carpeta_salida = Path(cfg.get('carpeta_salida', str(CARPETA_SALIDA_DEFECTO)))
        self.ruta_catalogo_actual = cfg.get('ruta_catalogo', RUTA_CATALOGO_DEFECTO)

        self.senales = Senales()
        self.senales.catalogo_listo.connect(self._on_catalogo_listo)
        self.senales.progreso.connect(self._on_progreso)
        self.senales.analisis_listo.connect(self._on_analisis_listo)

        self._construir_ui()
        self._cargar_catalogo_async(self.ruta_catalogo_actual)

    # ---------------- Construccion de la interfaz ----------------

    def _construir_ui(self):
        self.setWindowTitle(NOMBRE_APP)
        ruta_icono = resource_path("logo_entrol.png")
        if os.path.exists(ruta_icono):
            self.setWindowIcon(QIcon(ruta_icono))
        self.resize(780, 860)
        self.setMinimumSize(680, 500)

        # Layout raiz de la ventana: solo contiene el scroll area. Todo el
        # contenido real vive dentro de 'contenido' / 'layout'. Con esto, si
        # la ventana se maximiza en un monitor muy ancho pero no muy alto (o
        # se hace mas pequeña de lo que el contenido necesita), aparece una
        # barra de scroll en vez de que los widgets se compriman y se
        # solapen entre si (que es lo que pasaba antes).
        layout_raiz = QVBoxLayout(self)
        layout_raiz.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        layout_raiz.addWidget(scroll)

        contenido = QWidget()
        scroll.setWidget(contenido)

        layout = QVBoxLayout(contenido)
        layout.setContentsMargins(20, 20, 20, 16)
        layout.setSpacing(14)

        # Cabecera con logo
        cabecera = QHBoxLayout()
        ruta_logo = resource_path("logo_entrol.png")
        if os.path.exists(ruta_logo):
            lbl_logo = BodyLabel()
            pix = QPixmap(ruta_logo).scaledToHeight(32, Qt.TransformationMode.SmoothTransformation)
            lbl_logo.setPixmap(pix)
            cabecera.addWidget(lbl_logo)
            cabecera.addSpacing(10)
        titulo = TitleLabel(NOMBRE_APP)
        cabecera.addWidget(titulo)
        cabecera.addStretch()
        layout.addLayout(cabecera)

        # Tarjeta: catalogo
        tarjeta_cat = CardWidget()
        v = QVBoxLayout(tarjeta_cat)
        v.addWidget(StrongBodyLabel("Catálogo master"))
        self.lbl_catalogo = BodyLabel("Cargando catálogo...")
        self.lbl_catalogo.setWordWrap(True)
        v.addWidget(self.lbl_catalogo)
        btn_cat = PushButton("Cambiar catálogo...")
        btn_cat.clicked.connect(self._elegir_catalogo_manual)
        v.addWidget(btn_cat, alignment=Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(tarjeta_cat)

        # Tarjeta: LM
        tarjeta_lm = CardWidget()
        h = QHBoxLayout(tarjeta_lm)
        col = QVBoxLayout()
        col.addWidget(StrongBodyLabel("Lista de Materiales (.odt)"))
        self.lbl_lm = BodyLabel("Ningún archivo seleccionado")
        col.addWidget(self.lbl_lm)
        h.addLayout(col)
        h.addStretch()
        btn_lm = PushButton("Seleccionar LM...")
        btn_lm.clicked.connect(self._elegir_lm)
        h.addWidget(btn_lm)
        layout.addWidget(tarjeta_lm)

        # Tarjeta: carpeta de salida
        tarjeta_carpeta = CardWidget()
        h2 = QHBoxLayout(tarjeta_carpeta)
        col2 = QVBoxLayout()
        col2.addWidget(StrongBodyLabel("Carpeta de guardado de los informes"))
        self.lbl_carpeta = BodyLabel(str(self.carpeta_salida))
        self.lbl_carpeta.setWordWrap(True)
        col2.addWidget(self.lbl_carpeta)
        h2.addLayout(col2)
        h2.addStretch()
        btn_carpeta = PushButton("Cambiar carpeta...")
        btn_carpeta.clicked.connect(self._elegir_carpeta_salida)
        h2.addWidget(btn_carpeta)
        layout.addWidget(tarjeta_carpeta)

        # Tarjeta: configuracion (en dos columnas para aprovechar mejor el
        # ancho y no depender de tanto espacio vertical apilado)
        tarjeta_cfg = CardWidget()
        vc = QVBoxLayout(tarjeta_cfg)
        vc.addWidget(StrongBodyLabel("Configuración de este simulador"))

        columnas = QHBoxLayout()
        columnas.setSpacing(28)
        vc.addLayout(columnas)

        # --- Columna izquierda: Destino + checkboxes ---
        col_izq = QVBoxLayout()
        fila_pais = QHBoxLayout()
        fila_pais.addWidget(BodyLabel("Destino:"))
        self.radio_ue = RadioButton("Unión Europea")
        self.radio_usa = RadioButton("USA")
        self.radio_ue.setChecked(True)
        fila_pais.addWidget(self.radio_ue)
        fila_pais.addWidget(self.radio_usa)
        fila_pais.addStretch()
        col_izq.addLayout(fila_pais)
        self.chk_debriefing = CheckBox("Lleva Debriefing")
        self.chk_flir = CheckBox("Lleva FLIR handheld")
        self.chk_1000x = CheckBox("Es simulador 1000x")
        self.chk_referencias = CheckBox("Verificar últimas referencias")
        self.chk_plates = CheckBox("PLATES")
        col_izq.addWidget(self.chk_debriefing)
        col_izq.addWidget(self.chk_flir)
        col_izq.addWidget(self.chk_1000x)
        col_izq.addWidget(self.chk_referencias)
        col_izq.addWidget(self.chk_plates)
        col_izq.addStretch()
        columnas.addLayout(col_izq, 1)

        # --- Columna derecha: "Marca el visual" ---
        # En su propio widget contenedor (frame_visual) para que Qt no mezcle
        # estos radios con los de Destino (UE/USA) en el mismo grupo
        # autoExclusive - los radios de un mismo parent widget son
        # mutuamente excluyentes entre si en Qt.
        col_der = QVBoxLayout()
        col_der.addWidget(StrongBodyLabel("Marca el visual"))
        frame_visual = QWidget()
        visual_layout = QVBoxLayout(frame_visual)
        visual_layout.setContentsMargins(0, 0, 0, 0)
        visual_layout.setSpacing(6)
        self.radio_vis_a = RadioButton("VIS.014.X / VIS.020.X")
        self.radio_vis_b = RadioButton("VIS.030.X / VIS.031.X / VIS.032.X / VIS.033.X")
        self.radio_vis_otro = RadioButton("Otro visual (LED)")
        self.radio_vis_otro.setChecked(True)  # opcion por defecto, la mas comun
        visual_layout.addWidget(self.radio_vis_a)
        visual_layout.addWidget(self.radio_vis_b)
        visual_layout.addWidget(self.radio_vis_otro)
        col_der.addWidget(frame_visual)
        col_der.addStretch()
        columnas.addLayout(col_der, 1)

        layout.addWidget(tarjeta_cfg)

        # Boton analizar + progreso
        self.btn_analizar = PrimaryPushButton("Analizar LM")
        self.btn_analizar.setEnabled(False)
        self.btn_analizar.clicked.connect(self._analizar)
        layout.addWidget(self.btn_analizar)

        self.barra_progreso = ProgressBar()
        self.barra_progreso.setValue(0)
        layout.addWidget(self.barra_progreso)

        # Tarjeta: resultado
        tarjeta_res = CardWidget()
        vr = QVBoxLayout(tarjeta_res)
        vr.addWidget(StrongBodyLabel("Resultado"))
        self.txt_resultado = TextEdit()
        self.txt_resultado.setReadOnly(True)
        self.txt_resultado.setPlainText("—")
        self.txt_resultado.setMinimumHeight(160)
        vr.addWidget(self.txt_resultado)
        layout.addWidget(tarjeta_res, stretch=1)

        # Botones abrir
        fila_botones = QHBoxLayout()
        self.btn_abrir_excel = PushButton("Abrir Excel")
        self.btn_abrir_excel.setEnabled(False)
        self.btn_abrir_excel.clicked.connect(self._abrir_excel)
        self.btn_abrir_csv = PushButton("Abrir CSV")
        self.btn_abrir_csv.setEnabled(False)
        self.btn_abrir_csv.clicked.connect(self._abrir_csv)
        fila_botones.addWidget(self.btn_abrir_excel)
        fila_botones.addWidget(self.btn_abrir_csv)
        layout.addLayout(fila_botones)

        # Pie
        pie = QHBoxLayout()
        self.lbl_estado_guardado = CaptionLabel("")
        pie.addWidget(self.lbl_estado_guardado)
        pie.addStretch()
        pie.addWidget(CaptionLabel("Diseñado por ACC"))
        layout.addLayout(pie)

    # ---------------- Catalogo ----------------

    def _cargar_catalogo_async(self, ruta):
        self.lbl_catalogo.setText(f"Cargando catálogo desde:\n{ruta}")
        threading.Thread(target=self._trabajo_cargar_catalogo, args=(ruta,), daemon=True).start()

    def _trabajo_cargar_catalogo(self, ruta):
        try:
            catalogo = cargar_catalogo(ruta)
            self.senales.catalogo_listo.emit({'ok': True, 'ruta': ruta, 'n': len(catalogo), 'catalogo': catalogo})
        except Exception as e:
            registrar_error(f"Carga de catálogo ({ruta})", e)
            self.senales.catalogo_listo.emit({'ok': False, 'ruta': ruta, 'error': str(e)})

    def _on_catalogo_listo(self, r: dict):
        if r['ok']:
            self.catalogo = r['catalogo']
            self.lbl_catalogo.setText(f"✓ Catálogo cargado ({r['n']} referencias)\n{r['ruta']}")
        else:
            self.lbl_catalogo.setText(f"⚠ No se pudo cargar el catálogo automáticamente.\n"
                                       f"Ruta esperada: {r['ruta']}\n({r['error']})")
        self._actualizar_boton_analizar()

    def _elegir_catalogo_manual(self):
        ruta, _ = QFileDialog.getOpenFileName(
            self, "Selecciona el catálogo master",
            os.path.dirname(self.ruta_catalogo_actual) or "",
            "Catálogo ODS (*.ods)")
        if not ruta:
            return
        self.ruta_catalogo_actual = ruta
        guardar_config({'ruta_catalogo': ruta})
        self._cargar_catalogo_async(ruta)

    # ---------------- LM ----------------

    def _elegir_lm(self):
        ruta, _ = QFileDialog.getOpenFileName(
            self, "Selecciona el LM",
            os.path.dirname(self.ruta_lm) if self.ruta_lm else "",
            "Documento ODT (*.odt)")
        if not ruta:
            return
        self.ruta_lm = ruta
        self.lbl_lm.setText(os.path.basename(ruta))
        self._actualizar_boton_analizar()

    # ---------------- Carpeta de salida ----------------

    def _elegir_carpeta_salida(self):
        ruta = QFileDialog.getExistingDirectory(
            self, "Selecciona la carpeta donde guardar los informes",
            str(self.carpeta_salida) if self.carpeta_salida.exists() else str(_CARPETA_APP_LOCAL))
        if not ruta:
            return
        self.carpeta_salida = Path(ruta)
        self.lbl_carpeta.setText(str(self.carpeta_salida))
        guardar_config({'carpeta_salida': str(self.carpeta_salida)})

    def _actualizar_boton_analizar(self):
        self.btn_analizar.setEnabled(self.catalogo is not None and self.ruta_lm is not None)

    # ---------------- Analisis ----------------

    def _analizar(self):
        self.btn_analizar.setEnabled(False)
        self.btn_abrir_excel.setEnabled(False)
        self.btn_abrir_csv.setEnabled(False)
        self.txt_resultado.setPlainText("Analizando...")
        self.barra_progreso.setValue(0)

        if self.radio_vis_a.isChecked():
            visual_tipo = VISUAL_TIPO_A
        elif self.radio_vis_b.isChecked():
            visual_tipo = VISUAL_TIPO_B
        else:
            visual_tipo = VISUAL_TIPO_OTRO

        checkboxes = {
            'pais': 'USA' if self.radio_usa.isChecked() else 'UE',
            'debriefing': self.chk_debriefing.isChecked(),
            'flir': self.chk_flir.isChecked(),
            'sim_1000x': self.chk_1000x.isChecked(),
            'verificar_referencias': self.chk_referencias.isChecked(),
            'verificar_plates': self.chk_plates.isChecked(),
            'visual_tipo': visual_tipo,
        }
        threading.Thread(target=self._trabajo_analizar, args=(checkboxes,), daemon=True).start()

    def _trabajo_analizar(self, checkboxes):
        try:
            df_lm = extraer_tabla_odt(self.ruta_lm)

            def progreso_cb(actual, total):
                pct = int((actual / total) * 100) if total else 100
                self.senales.progreso.emit(pct)

            df_resultado, df_incidencias_fila = analizar_lm(df_lm, self.catalogo, progreso_cb=progreso_cb)
            avisos_generales = aplicar_reglas_especiales(df_lm, checkboxes)

            tabla_referencias = None
            if checkboxes.get('verificar_referencias'):
                cabecera = extraer_cabecera_modelo(self.ruta_lm)
                tabla_referencias = analizar_referencias_modelo(df_lm, cabecera['modelo'], cabecera['declarado'])

            tabla_plates = None
            if checkboxes.get('verificar_plates'):
                tabla_plates = analizar_plates(df_lm)

            df_avisos = pd.DataFrame(avisos_generales)
            df_todas = (pd.concat([df_incidencias_fila, df_avisos], ignore_index=True)
                        if not df_avisos.empty else df_incidencias_fila)

            self.carpeta_salida.mkdir(parents=True, exist_ok=True)
            base = Path(self.ruta_lm).stem
            ts = datetime.now().strftime("%Y%m%d_%H%M")
            ruta_xlsx = self.carpeta_salida / f"Informe_{base}_{ts}.xlsx"
            ruta_csv = self.carpeta_salida / f"Informe_{base}_{ts}.csv"

            generar_excel(df_resultado, str(ruta_xlsx), avisos_generales, tabla_referencias, tabla_plates)
            df_todas.to_csv(str(ruta_csv), index=False, encoding="utf-8-sig")

            conteo_fila = {}
            if len(df_incidencias_fila):
                conteo_fila = (df_incidencias_fila['Tipo'].str.split(', ').explode()
                               .value_counts().to_dict())

            self.senales.analisis_listo.emit({
                'ok': True,
                'conteo_fila': conteo_fila,
                'total_fila': len(df_incidencias_fila),
                'avisos': avisos_generales,
                'tabla_referencias': tabla_referencias,
                'tabla_plates': tabla_plates,
                'ruta_xlsx': str(ruta_xlsx),
                'ruta_csv': str(ruta_csv),
                'carpeta': str(self.carpeta_salida),
            })
        except Exception as e:
            registrar_error(f"Análisis del LM ({self.ruta_lm})", e)
            self.senales.analisis_listo.emit({'ok': False, 'error': str(e)})

    def _on_progreso(self, pct: int):
        self.barra_progreso.setValue(pct)

    def _on_analisis_listo(self, r: dict):
        self._actualizar_boton_analizar()

        if not r['ok']:
            self.txt_resultado.setPlainText(f"✗ Error durante el análisis:\n{r['error']}")
            InfoBar.error("Error", r['error'], parent=self, position=InfoBarPosition.TOP, duration=6000)
            return

        lineas = []
        total = r['total_fila']
        if total == 0:
            lineas.append("✓ Sin incidencias fila a fila.")
        else:
            lineas.append(f"Incidencias fila a fila: {total}")
            for tipo, n in r['conteo_fila'].items():
                lineas.append(f"  • {tipo}: {n}")

        lineas.append("")
        if r['avisos']:
            lineas.append(f"Avisos de reglas especiales: {len(r['avisos'])}")
            for av in r['avisos']:
                fila_txt = f" (fila {av['Fila']})" if av['Fila'] != 'General' else ""
                lineas.append(f"  • {av['Tipo']}{fila_txt}: {av['Detalle']}")
        else:
            lineas.append("✓ Sin avisos de reglas especiales (con la configuración marcada).")

        if r.get('tabla_referencias'):
            lineas.append("")
            lineas.append("Verificación de últimas referencias:")
            for fila_ref in r['tabla_referencias']:
                base_txt = (f"  • {fila_ref['Prefijo']}: cabecera={fila_ref['Cabecera indica']}, "
                            f"máximo real={fila_ref['Máximo en este simulador']}")
                lineas.append(base_txt)
                if fila_ref['Aviso']:
                    lineas.append(f"      ⚠ {fila_ref['Aviso']}")

        if r.get('tabla_plates') is not None:
            tabla_plt = r['tabla_plates']
            mal = [f for f in tabla_plt if f['Aviso']]
            lineas.append("")
            lineas.append(f"Plates encontradas: {len(tabla_plt)} (fuera de Gr 6: {len(mal)})")
            for f in mal:
                lineas.append(f"  ⚠ Fila {f['Fila']} | {f['Referencia PLT']} | {f['Description']}: {f['Aviso']} (Gr actual={f['Gr']})")

        self.txt_resultado.setPlainText("\n".join(lineas))
        self.ruta_xlsx = r['ruta_xlsx']
        self.ruta_csv = r['ruta_csv']
        self.btn_abrir_excel.setEnabled(True)
        self.btn_abrir_csv.setEnabled(True)
        self.lbl_estado_guardado.setText(f"Guardado en: {r['carpeta']}")
        InfoBar.success("Análisis completado", f"{total} incidencias fila a fila",
                         parent=self, position=InfoBarPosition.TOP, duration=3000)

    # ---------------- Abrir archivos ----------------

    def _abrir_excel(self):
        self._abrir_archivo(self.ruta_xlsx)

    def _abrir_csv(self):
        self._abrir_archivo(self.ruta_csv)

    def _abrir_archivo(self, ruta):
        if not ruta or not os.path.exists(ruta):
            QMessageBox.warning(self, "Archivo no encontrado", f"No se encuentra:\n{ruta}")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(ruta))


def main():
    app = QApplication(sys.argv)
    ruta_icono = resource_path("logo_entrol.png")
    if os.path.exists(ruta_icono):
        app.setWindowIcon(QIcon(ruta_icono))
    setTheme(Theme.LIGHT)
    setThemeColor(COLOR_MARCA)
    ventana = VentanaPrincipal()
    ventana.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
