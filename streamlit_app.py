import streamlit as st
import tempfile, sys
from pathlib import Path

st.set_page_config(page_title="Automatización Lavalozas", page_icon="📊", layout="centered")

BASE_DIR      = Path(__file__).parent
TEMPLATE_XLSX = BASE_DIR / "Template_tablas_y_graficas_excel.xlsx"
TEMPLATE_PPTX = BASE_DIR / "Template_presentaciones_power_point.pptx"
sys.path.insert(0, str(BASE_DIR))
from pipeline import run_pipeline

st.markdown("""
<style>
  .block-container { padding-top:1.8rem; max-width:700px; }
  .nivel-row { display:flex; align-items:center; gap:14px;
               background:#f8faff; border:1.5px solid #dce4f5;
               border-radius:12px; padding:14px 18px; margin-bottom:12px; }
  .nivel-icon { font-size:22px; }
  .nivel-info { flex:1; }
  .nivel-title { font-weight:700; font-size:14px; color:#1a1a2e; }
  .nivel-hint  { font-size:11px; color:#9ca3af; margin-top:1px; }
  .badge-ok   { background:#dcfce7; color:#166534; border-radius:6px;
                padding:2px 9px; font-size:11px; font-weight:700; }
  .badge-opt  { background:#f3f4f6; color:#9ca3af; border-radius:6px;
                padding:2px 9px; font-size:11px; }
</style>
""", unsafe_allow_html=True)

st.title("📊 Automatización Lavalozas")
st.caption("Sube 1, 2 o 3 archivos Excel → un PowerPoint con una sección por nivel")
st.divider()

NIVELES = [
    ("marca",        "Marca",        "marca.xlsx",        "🏷️"),
    ("segmento",     "Segmento",     "segmento.xlsx",     "📦"),
    ("subcategoria", "Subcategoría", "subcategoria.xlsx", "🔍"),
]

uploaded = {}
for key, label, hint, icon in NIVELES:
    left, right = st.columns([5, 3])
    with left:
        st.markdown(f"""
        <div class="nivel-row">
          <div class="nivel-icon">{icon}</div>
          <div class="nivel-info">
            <div class="nivel-title">{label}</div>
            <div class="nivel-hint">Nombre esperado: <code>{hint}</code></div>
          </div>
        </div>""", unsafe_allow_html=True)
    with right:
        f = st.file_uploader("_", type=["xlsx"],
                             label_visibility="collapsed", key=f"up_{key}")
        if f:
            uploaded[key] = f
            st.markdown('<span class="badge-ok">✅ Cargado</span>', unsafe_allow_html=True)
        else:
            st.markdown('<span class="badge-opt">Opcional</span>', unsafe_allow_html=True)

st.divider()

n = len(uploaded)
if n == 0:
    st.info("👆 Sube al menos un archivo para continuar.")
else:
    niveles_str = " · ".join(label for key, label, _, __ in NIVELES if key in uploaded)
    st.success(f"**{n} archivo(s) listo(s):** {niveles_str}")

    if st.button("🚀 Generar Presentación", type="primary", use_container_width=True):
        bar = st.progress(0, text="Iniciando…")
        try:
            with tempfile.TemporaryDirectory() as tmp:
                tmp = Path(tmp)
                input_files = {}
                for key, f in uploaded.items():
                    p = tmp / f"{key}.xlsx"
                    p.write_bytes(f.read())
                    input_files[key] = p

                bar.progress(20, text="Procesando datos…")
                pptx_bytes, summary = run_pipeline(
                    input_files, TEMPLATE_XLSX, TEMPLATE_PPTX, tmp)
                bar.progress(100, text="¡Listo!")

            st.balloons()
            st.success(f"✅ **Presentación generada** — {summary}")
            st.download_button(
                label="⬇️  Descargar PowerPoint",
                data=pptx_bytes,
                file_name="Presentacion_Lavalozas.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )
        except Exception as e:
            bar.empty()
            st.error(f"❌ {e}")
            st.exception(e)
