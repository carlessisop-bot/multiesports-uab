import streamlit as st
import pandas as pd
import datetime
import io
import os
import re

# ==========================================
# 1. CONFIGURACIÓ DE PÀGINA I ESTILS UAB
# ==========================================
st.set_page_config(
    page_title="MultiEsports UAB - Portal Oficial & Gestió",
    page_icon="⚽",
    layout="wide"
)

# Estils CSS Corporatius UAB
st.markdown("""
    <style>
    .main-title {
        color: #008037;
        font-family: 'Helvetica', sans-serif;
        font-weight: bold;
        font-size: 2.2rem;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        color: #003366;
        font-size: 1.1rem;
        margin-bottom: 1.5rem;
    }
    .badge-status {
        background-color: #e8f5e9;
        color: #2e7d32;
        padding: 4px 8px;
        border-radius: 4px;
        font-weight: bold;
    }
    .stTable {
        border-radius: 8px;
        overflow: hidden;
    }
    </style>
""", unsafe_allow_html=True)

col_logo1, col_title, col_logo2 = st.columns([1.2, 4, 1.2], vertical_alignment="center")

with col_logo1:
    path_uab = None
    for p in ["logo_UAB.jpg", "/workspace/knowledge/logo_UAB.jpg", "/workspace/artifacts/logo_UAB.jpg"]:
        if os.path.exists(p):
            path_uab = p
            break
    if path_uab:
        st.image(path_uab, use_container_width=True)

with col_title:
    st.markdown("""<div class="main-title" style="text-align: center;">MultiEsports UAB — SAF Servei d'Activitat Física</div>""", unsafe_allow_html=True)
    st.markdown("""<div class="sub-title" style="text-align: center;">Portal Públic de Resultats & Panell d'Àrbitres</div>""", unsafe_allow_html=True)

with col_logo2:
    path_saf = None
    for p in ["logo_competicions_SAF.jpg", "logo competicions SAF.jpg", "/workspace/knowledge/logo_competicions_SAF.jpg", "/workspace/artifacts/logo_competicions_SAF.jpg"]:
        if os.path.exists(p):
            path_saf = p
            break
    if path_saf:
        st.image(path_saf, use_container_width=True)

PASSWORD_ARBITRE = "7hk:2pJHHEzm8a6"        # Contrasenya per als àrbitres/coordinadors
PASSWORD_SUPERADMIN = "7hk:2pJHHEzm6a1" # Contrasenya per al Superadministrador

# ==========================================
# 2. FUNCIONS AUXILIARS I DE SEGURETAT
# ==========================================

def sanitize_pdf_text(text):
    if text is None:
        return ""
    s = str(text)
    s = s.replace("—", "-").replace("–", "-").replace("“", '"').replace("”", '"').replace("’", "'").replace("‘", "'")
    return s.encode('latin-1', 'replace').decode('latin-1')

def safe_int(val, default=0):
    try:
        if pd.notnull(val) and str(val).strip() != "" and str(val).strip().lower() != "nan":
            return int(float(val))
    except (ValueError, TypeError):
        pass
    return default


def is_same_id(id1, id2):
    if id1 is None or id2 is None:
        return False
    s1 = str(id1).strip().upper()
    s2 = str(id2).strip().upper()
    if not s1 or not s2:
        return False
    if s1 == s2:
        return True
    try:
        if float(s1) == float(s2):
            return True
    except (ValueError, TypeError):
        pass
    return False

def obtenir_fitxer_excel():
    if os.path.exists("Plantilla_MultiEsports_UAB.xlsx"):
        return "Plantilla_MultiEsports_UAB.xlsx"
    elif os.path.exists("Plantilla_MultiEsports_UAB.xlsm"):
        return "Plantilla_MultiEsports_UAB.xlsm"
    else:
        files = [f for f in os.listdir(".") if (f.endswith(".xlsm") or f.endswith(".xlsx")) and not f.startswith("~$")]
        if files:
            return files[0]
    return "Plantilla_MultiEsports_UAB.xlsx"

# ==========================================
# 3. LECTURA I ESCRIPTURA D'EXCEL
# ==========================================

def carregar_calendari():
    excel_file = obtenir_fitxer_excel()
    if os.path.exists(excel_file):
        try:
            return pd.read_excel(excel_file, sheet_name="Calendari", engine="openpyxl")
        except Exception as e:
            st.error(f"Error en llegir el full Calendari del fitxer '{excel_file}': {e}")
            return pd.DataFrame()
    else:
        st.warning(f"No s'ha trobat cap fitxer Excel (.xlsx/.xlsm) a la carpeta actual ({os.getcwd()}).")
        return pd.DataFrame()


def carregar_passwords():
    excel_file = obtenir_fitxer_excel()
    if os.path.exists(excel_file):
        try:
            xls = pd.ExcelFile(excel_file, engine="openpyxl")
            for sheet in xls.sheet_names:
                if "password" in sheet.lower():
                    return pd.read_excel(xls, sheet_name=sheet)
        except Exception:
            pass
    return pd.DataFrame()

def carregar_consolidat():
    excel_file = obtenir_fitxer_excel()
    if os.path.exists(excel_file):
        try:
            xls = pd.ExcelFile(excel_file, engine="openpyxl")
            if "Consolidat" in xls.sheet_names:
                return pd.read_excel(xls, sheet_name="Consolidat")
        except Exception:
            pass
    return pd.DataFrame()

def carregar_jugadors_equip(esport, equip):
    df_cons = carregar_consolidat()
    if df_cons.empty:
        return []
    
    cols_lower = [str(c).strip().lower() for c in df_cons.columns]
    
    col_comp = next((df_cons.columns[i] for i, c in enumerate(cols_lower) if "competi" in c or "esport" in c), df_cons.columns[0])
    col_eq = next((df_cons.columns[i] for i, c in enumerate(cols_lower) if "equip" in c), df_cons.columns[1] if len(df_cons.columns)>1 else df_cons.columns[0])
    
    col_nom = next((df_cons.columns[i] for i, c in enumerate(cols_lower) if "component" in c or "cognoms" in c or "jugador" in c), None)
    if col_nom is None:
        col_nom = next((df_cons.columns[i] for i, c in enumerate(cols_lower) if "nom" in c and "equip" not in c), df_cons.columns[3] if len(df_cons.columns)>3 else df_cons.columns[0])
        
    col_id = next((df_cons.columns[i] for i, c in enumerate(cols_lower) if "idparticipant" in c or "id_participant" in c or "participant" in c), None)
    if col_id is None:
        col_id = next((df_cons.columns[i] for i, c in enumerate(cols_lower) if c == "id" or "id" in c), None)
    
    mask = (df_cons[col_comp].astype(str).str.strip().str.upper() == str(esport).strip().upper()) &            (df_cons[col_eq].astype(str).str.strip().str.upper() == str(equip).strip().upper())
           
    df_sub = df_cons[mask]
    jugadors = []
    for _, r in df_sub.iterrows():
        nom = str(r[col_nom]).strip() if pd.notnull(r[col_nom]) else ""
        id_part = str(r[col_id]).strip() if col_id and pd.notnull(r[col_id]) else ""
        if nom and nom.lower() != "nan":
            jugadors.append({"nom": nom, "id_participant": id_part})
    return jugadors

def guardar_acta_amb_assistencies(id_partit, esport, eq_loc, eq_vis, res_loc, res_vis, 
                                  tg_loc=0, tv_loc=0, tg_vis=0, tv_vis=0, 
                                  text_gols_loc="", text_gols_vis="", 
                                  assistencies_loc={}, assistencies_vis={},
                                  estat_partit="VALIDAT",
                                  nom_arbitre="", nom_delegat_loc="", sig_delegat_loc="",
                                  nom_delegat_vis="", sig_delegat_vis="", sig_arbitre="", observacions=""):
    excel_file = obtenir_fitxer_excel()
    try:
        import openpyxl
        
        wb = openpyxl.load_workbook(excel_file, keep_vba=excel_file.lower().endswith('.xlsm'))
        if "Calendari" not in wb.sheetnames:
            return False, "No s'ha trobat el full 'Calendari' a l'Excel."
            
        ws_cal = wb["Calendari"]
        
        # 1. Mapatge de columnes de Calendari
        col_map = {}
        for col in range(1, ws_cal.max_column + 1):
            hdr = str(ws_cal.cell(row=1, column=col).value).strip()
            if hdr: col_map[hdr] = col
            
        col_id = col_map.get("Id", 1)
        col_rl = col_map.get("ResultatLocal", 12)
        col_rv = col_map.get("ResultatVisitant", 13)
        col_estat = col_map.get("EstatPartit", 15)
        
        col_tgl = col_map.get("TargetesGroguesLocal", 18)
        col_tvl = col_map.get("TargetesVermellesLocal", 19)
        col_tgv = col_map.get("TargetesGroguesVisitant", 20)
        col_tvv = col_map.get("TargetesVermellesVisitant", 21)
        
        col_gl = col_map.get("GolejadorsLocal", 22)
        col_gv = col_map.get("GolejadorsVisitant", 23)
        
        trobat = False
        target_id_str = str(id_partit).strip()
        
        def _get_col(hdr_str):
            if hdr_str in col_map:
                return col_map[hdr_str]
            new_c = ws_cal.max_column + 1
            ws_cal.cell(row=1, column=new_c, value=hdr_str)
            col_map[hdr_str] = new_c
            return new_c

        col_arb = col_map.get("Arbitre") or col_map.get("Àrbitre") or _get_col("Arbitre")
        col_dl = col_map.get("DelegatLocal") or _get_col("DelegatLocal")
        col_sdl = col_map.get("SignaturaDelegatLocal") or _get_col("SignaturaDelegatLocal")
        col_dv = col_map.get("DelegatVisitant") or _get_col("DelegatVisitant")
        col_sdv = col_map.get("SignaturaDelegatVisitant") or _get_col("SignaturaDelegatVisitant")
        col_sarb = col_map.get("SignaturaArbitre") or col_map.get("SignaturaÀrbitre") or _get_col("SignaturaArbitre")
        col_obs = col_map.get("Observacions") or _get_col("Observacions")

        for row in range(2, ws_cal.max_row + 1):
            val_id_raw = ws_cal.cell(row=row, column=col_id).value
            if val_id_raw is not None:
                if is_same_id(val_id_raw, id_partit):
                    ws_cal.cell(row=row, column=col_rl, value=safe_int(res_loc))
                    ws_cal.cell(row=row, column=col_rv, value=safe_int(res_vis))
                    ws_cal.cell(row=row, column=col_estat, value=str(estat_partit).strip())
                    
                    ws_cal.cell(row=row, column=col_tgl, value=safe_int(tg_loc))
                    ws_cal.cell(row=row, column=col_tvl, value=safe_int(tv_loc))
                    ws_cal.cell(row=row, column=col_tgv, value=safe_int(tg_vis))
                    ws_cal.cell(row=row, column=col_tvv, value=safe_int(tv_vis))
                    
                    if text_gols_loc: ws_cal.cell(row=row, column=col_gl, value=str(text_gols_loc))
                    if text_gols_vis: ws_cal.cell(row=row, column=col_gv, value=str(text_gols_vis))

                    if nom_arbitre: ws_cal.cell(row=row, column=col_arb, value=str(nom_arbitre).strip())
                    if nom_delegat_loc: ws_cal.cell(row=row, column=col_dl, value=str(nom_delegat_loc).strip())
                    if sig_delegat_loc: ws_cal.cell(row=row, column=col_sdl, value=str(sig_delegat_loc).strip())
                    if nom_delegat_vis: ws_cal.cell(row=row, column=col_dv, value=str(nom_delegat_vis).strip())
                    if sig_delegat_vis: ws_cal.cell(row=row, column=col_sdv, value=str(sig_delegat_vis).strip())
                    if sig_arbitre: ws_cal.cell(row=row, column=col_sarb, value=str(sig_arbitre).strip())
                    if observacions: ws_cal.cell(row=row, column=col_obs, value=str(observacions).strip())

                    trobat = True
                    break
                    
        if not trobat:
            return False, f"No s'ha trobat el partit amb ID '{id_partit}' al Calendari."
            
        # 2. Gravació d'Assistències al full 'Assistencies'
        headers_ass = ["IdPartit", "IdParticipant", "NomJugador", "Competició", "Equip", "Assisteix", "DataRegistre", "Gols", "Groga", "Vermella"]
        if "Assistencies" not in wb.sheetnames:
            ws_ass = wb.create_sheet("Assistencies")
        else:
            ws_ass = wb["Assistencies"]
            
        for c_idx, h in enumerate(headers_ass, 1):
            ws_ass.cell(row=1, column=c_idx, value=h)
            
        now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        
        # Eliminar registres anteriors d'aquest partit si ja s'havia validat
        rows_to_delete = []
        for r in range(2, ws_ass.max_row + 1):
            if is_same_id(ws_ass.cell(row=r, column=1).value, id_partit):
                rows_to_delete.append(r)
        for r in reversed(rows_to_delete):
            ws_ass.delete_rows(r)
            
        def _inserir_files_assistencia(llista_assist, nom_equip):
            items = llista_assist.items() if isinstance(llista_assist, dict) else llista_assist
            for item in items:
                p_gols, p_groga, p_vermella = 0, 0, 0
                if isinstance(item, dict):
                    jug_info = item.get("jug", {})
                    assist = item.get("assist", True)
                    p_gols = item.get("gols", 0)
                    p_groga = item.get("groga", 0)
                    p_vermella = item.get("vermella", 0)
                elif isinstance(item, tuple) and len(item) == 2:
                    jug_info, assist = item
                else:
                    continue
                if isinstance(jug_info, tuple):
                    jug_info = dict(jug_info)
                nom_j = jug_info.get("nom", "") if isinstance(jug_info, dict) else ""
                id_p = jug_info.get("id_participant", "") if isinstance(jug_info, dict) else ""
                
                r_next = ws_ass.max_row + 1
                ws_ass.cell(row=r_next, column=1, value=str(id_partit))
                ws_ass.cell(row=r_next, column=2, value=str(id_p))
                ws_ass.cell(row=r_next, column=3, value=str(nom_j))
                ws_ass.cell(row=r_next, column=4, value=str(esport))
                ws_ass.cell(row=r_next, column=5, value=str(nom_equip))
                ws_ass.cell(row=r_next, column=6, value=1 if assist else 0)
                ws_ass.cell(row=r_next, column=7, value=now_str)
                ws_ass.cell(row=r_next, column=8, value=safe_int(p_gols))
                ws_ass.cell(row=r_next, column=9, value=safe_int(p_groga))
                ws_ass.cell(row=r_next, column=10, value=safe_int(p_vermella))
                
        _inserir_files_assistencia(assistencies_loc, eq_loc)
        _inserir_files_assistencia(assistencies_vis, eq_vis)
            
        wb.save(excel_file)
        return True, "Acta de partit i registre d'assistències desats amb èxit!"
    except Exception as e:
        return False, f"Error en desar a Excel: {e}"

# ==========================================
# 4. GENERACIÓ DE INFORMES (EXCEL & PDF)
# ==========================================

def generar_pdf_acta_partit(row_partit, jugadors_loc=[], jugadors_vis=[]):
    p_id = sanitize_pdf_text(row_partit.get("Id", ""))
    esport = sanitize_pdf_text(row_partit.get("Esport", ""))
    grup = sanitize_pdf_text(row_partit.get("Grup", ""))
    jornada = sanitize_pdf_text(row_partit.get("Jornada", ""))
    eq_l = sanitize_pdf_text(row_partit.get("EquipLocal", ""))
    eq_v = sanitize_pdf_text(row_partit.get("EquipVisitant", ""))
    data_p = sanitize_pdf_text(row_partit.get("DataPartit", ""))
    hora_p = sanitize_pdf_text(row_partit.get("HoraPartit", ""))
    camp = sanitize_pdf_text(row_partit.get("Camp", ""))
    estat = sanitize_pdf_text(row_partit.get("EstatPartit", "Programat"))
    
    res_l = safe_int(row_partit.get("ResultatLocal", 0))
    res_v = safe_int(row_partit.get("ResultatVisitant", 0))
    
    arb_nom = sanitize_pdf_text(row_partit.get("Arbitre", "") if pd.notnull(row_partit.get("Arbitre")) else "")
    del_l_nom = sanitize_pdf_text(row_partit.get("DelegatLocal", "") if pd.notnull(row_partit.get("DelegatLocal")) else "")
    sig_dl = sanitize_pdf_text(row_partit.get("SignaturaDelegatLocal", "") if pd.notnull(row_partit.get("SignaturaDelegatLocal")) else "")
    del_v_nom = sanitize_pdf_text(row_partit.get("DelegatVisitant", "") if pd.notnull(row_partit.get("DelegatVisitant")) else "")
    sig_dv = sanitize_pdf_text(row_partit.get("SignaturaDelegatVisitant", "") if pd.notnull(row_partit.get("SignaturaDelegatVisitant")) else "")
    sig_arb = sanitize_pdf_text(row_partit.get("SignaturaArbitre", "") if pd.notnull(row_partit.get("SignaturaArbitre")) else "")
    obs_text = sanitize_pdf_text(row_partit.get("Observacions", "") if pd.notnull(row_partit.get("Observacions")) else "")

    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib import colors

        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
        styles = getSampleStyleSheet()

        title_style = ParagraphStyle('ActaTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=15, textColor=colors.HexColor('#008037'), alignment=1, spaceAfter=2)
        sub_style = ParagraphStyle('ActaSub', parent=styles['Normal'], fontName='Helvetica-Oblique', fontSize=9, textColor=colors.HexColor('#555555'), alignment=1, spaceAfter=10)
        sec_style = ParagraphStyle('ActaSec', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=11, textColor=colors.HexColor('#003366'), spaceBefore=8, spaceAfter=4)

        story = []
        now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        story.append(Paragraph(f"ACTA OFICIAL DE PARTIT - {p_id}", title_style))
        story.append(Paragraph(f"SAF Servei d'Activitat Física UAB | Imprès el {now_str}", sub_style))

        fitxa_data = [
            ["Competició:", esport, "Grup:", grup, "Jornada:", str(jornada)],
            ["Equip Local:", eq_l, "Equip Visitant:", eq_v, "Estat:", estat],
            ["Data / Hora:", f"{data_p} {hora_p}", "Instal·lació:", camp, "Resultat:", f"{res_l} - {res_v}"]
        ]
        t_fitxa = Table(fitxa_data, colWidths=[70, 120, 70, 120, 60, 100])
        t_fitxa.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F2F5F2')),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
            ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
            ('FONTNAME', (2,0), (2,-1), 'Helvetica-Bold'),
            ('FONTNAME', (4,0), (4,-1), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#DDDDDD')),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(t_fitxa)
        story.append(Spacer(1, 8))

        story.append(Paragraph("Convocatòria i Control de Jugadors/es", sec_style))
        max_rows = max(len(jugadors_loc), len(jugadors_vis), 1)
        jug_table_data = [["🏠 " + eq_l[:25], "Assis.", "Gols", "🟨", "🟥", "✈️ " + eq_v[:25], "Assis.", "Gols", "🟨", "🟥"]]
        
        for idx_r in range(max_rows):
            j_l = jugadors_loc[idx_r]["nom"] if idx_r < len(jugadors_loc) else ""
            j_v = jugadors_vis[idx_r]["nom"] if idx_r < len(jugadors_vis) else ""
            jug_table_data.append([
                sanitize_pdf_text(j_l)[:25], "[  ]", "", "", "",
                sanitize_pdf_text(j_v)[:25], "[  ]", "", "", ""
            ])
            
        t_jug = Table(jug_table_data, colWidths=[140, 35, 30, 20, 20, 140, 35, 30, 20, 20])
        t_jug.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (4,0), colors.HexColor('#008037')),
            ('BACKGROUND', (5,0), (9,0), colors.HexColor('#003366')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 7.5),
            ('ALIGN', (1,0), (-1,-1), 'CENTER'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ]))
        story.append(t_jug)
        story.append(Spacer(1, 8))

        story.append(Paragraph("Validació i Signatures Oficials", sec_style))
        sig_data = [
            ["Àrbitre / Col·legiat/da:", arb_nom or "_______________________", "Estat Signatura:", sig_arb or "PENDENT"],
            [f"Delegat/da {eq_l}:", del_l_nom or "_______________________", "Conformitat:", sig_dl or "PENDENT"],
            [f"Delegat/da {eq_v}:", del_v_nom or "_______________________", "Conformitat:", sig_dv or "PENDENT"]
        ]
        t_sig = Table(sig_data, colWidths=[120, 180, 100, 140])
        t_sig.setStyle(TableStyle([
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
            ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
            ('FONTNAME', (2,0), (2,-1), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#DDDDDD')),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ]))
        story.append(t_sig)
        story.append(Spacer(1, 8))
        story.append(Paragraph("Observacions / Incidències de l'Àrbitre", sec_style))
        obs_display = obs_text if obs_text else "Cap incidència registrada."
        story.append(Paragraph(obs_display, ParagraphStyle('ActaObs', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, textColor=colors.HexColor('#333333'))))

        doc.build(story)
        return buffer.getvalue()
    except Exception as e_rl:
        from fpdf import FPDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 14)
        pdf.cell(0, 10, sanitize_pdf_text(f"ACTA OFICIAL - {p_id}"), new_x="LMARGIN", new_y="NEXT", align="C")
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 6, sanitize_pdf_text(f"Partit: {eq_l} vs {eq_v} | Estat: {estat}"), new_x="LMARGIN", new_y="NEXT")
        pdf.cell(0, 6, sanitize_pdf_text(f"Resultat: {res_l} - {res_v} | Data: {data_p} {hora_p}"), new_x="LMARGIN", new_y="NEXT")
        res_out = pdf.output()
        if isinstance(res_out, str):
            return res_out.encode('latin1')
        return bytes(res_out)

def generar_excel_resum(esport, grup, df_standings, df_pitxitxis, df_fp):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    buffer = io.BytesIO()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    font_title = Font(name="Arial", size=13, bold=True, color="FFFFFF")
    font_sub = Font(name="Arial", size=9, italic=True, color="555555")
    font_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    font_body = Font(name="Arial", size=10)
    
    fill_title = PatternFill(start_color="008037", end_color="008037", fill_type="solid")
    fill_header_green = PatternFill(start_color="008037", end_color="008037", fill_type="solid")
    fill_header_blue = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
    fill_zebra = PatternFill(start_color="F9FBF9", end_color="F9FBF9", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='DDDDDD'),
        right=Side(style='thin', color='DDDDDD'),
        top=Side(style='thin', color='DDDDDD'),
        bottom=Side(style='thin', color='DDDDDD')
    )
    
    now_str = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
    
    def afegir_pestanya(ws_title, df, fill_hdr):
        ws = wb.create_sheet(title=ws_title)
        num_cols = max(len(df.columns) if not df.empty else 4, 4)
        
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
        cell_t = ws.cell(row=1, column=1, value=f"Resum Executiu Oficial — {esport} ({grup})")
        cell_t.font = font_title
        cell_t.fill = fill_title
        cell_t.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions.height = 28
        
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_cols)
        cell_s = ws.cell(row=2, column=1, value=f"SAF Servei d'Activitat Física UAB | Generat el {now_str}")
        cell_s.font = font_sub
        cell_s.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions.height = 18
        
        if df.empty:
            ws.cell(row=4, column=1, value="No hi ha dades registrades per a aquesta secció.").font = font_body
            return
            
        headers = list(df.columns)
        ws.row_dimensions.height = 22
        for col_idx, h in enumerate(headers, 1):
            c = ws.cell(row=4, column=col_idx, value=str(h))
            c.font = font_header
            c.fill = fill_hdr
            c.alignment = Alignment(horizontal='center', vertical='center')
            
        for row_idx, r in enumerate(df.itertuples(index=False), 5):
            ws.row_dimensions[row_idx].height = 18
            for col_idx, val in enumerate(r, 1):
                c = ws.cell(row=row_idx, column=col_idx, value=val)
                c.font = font_body
                c.border = thin_border
                if row_idx % 2 == 0:
                    c.fill = fill_zebra
                c.alignment = Alignment(horizontal='left' if col_idx == 2 else 'center', vertical='center')
                
        for col_idx in range(1, len(headers) + 1):
            col_letter = get_column_letter(col_idx)
            max_len = len(str(headers[col_idx - 1]))
            for row_idx in range(5, 5 + len(df)):
                val = ws.cell(row=row_idx, column=col_idx).value
                if val is not None:
                    max_len = max(max_len, len(str(val)))
            ws.column_dimensions[col_letter].width = max(max_len + 5, 12)

    afegir_pestanya('Classificacio', df_standings, fill_header_green)
    afegir_pestanya('Pitxitxis', df_pitxitxis, fill_header_blue)
    afegir_pestanya('FairPlay', df_fp, fill_header_green)

    wb.save(buffer)
    return buffer.getvalue()


def generar_pdf_fitxa_equip(esport, equip, jugadors=[]):
    esport = sanitize_pdf_text(esport)
    equip = sanitize_pdf_text(equip)
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib import colors

        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        styles = getSampleStyleSheet()

        title_style = ParagraphStyle('FitxaTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=15, textColor=colors.HexColor('#008037'), alignment=1, spaceAfter=4)
        subtitle_style = ParagraphStyle('FitxaSubtitle', parent=styles['Normal'], fontName='Helvetica-Oblique', fontSize=9, textColor=colors.HexColor('#555555'), alignment=1, spaceAfter=15)
        sec_style = ParagraphStyle('FitxaSec', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=11, textColor=colors.HexColor('#003366'), spaceBefore=10, spaceAfter=6)

        story = []
        now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        story.append(Paragraph(f"FITXA OFICIAL D'EQUIP - {equip}", title_style))
        story.append(Paragraph(f"Competició: {esport} | Generat el {now_str} | SAF Servei d'Activitat Física UAB", subtitle_style))

        story.append(Paragraph("Plantilla i Convocatòria Oficial", sec_style))
        if jugadors:
            data = [["#", "Nom i Cognoms del Participant", "ID Participant", "Dorsal", "Observacions / Signatura"]]
            for idx, j in enumerate(jugadors, 1):
                nom_j = sanitize_pdf_text(j.get("nom", ""))
                id_j = sanitize_pdf_text(j.get("id_participant", ""))
                data.append([str(idx), nom_j, id_j, "", ""])
            t = Table(data, colWidths=[25, 200, 90, 45, 130])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#008037')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,-1), 8.5),
                ('ALIGN', (0,0), (0,-1), 'CENTER'),
                ('ALIGN', (2,0), (3,-1), 'CENTER'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')),
                ('TOPPADDING', (0,0), (-1,-1), 4),
                ('BOTTOMPADDING', (0,0), (-1,-1), 4)
            ]))
            story.append(t)
        else:
            story.append(Paragraph("No s'han trobat jugadors registrats per a aquest equip.", subtitle_style))

        doc.build(story)
        return buffer.getvalue()
    except Exception:
        return b""

def generar_pdf_resum(esport, grup, df_standings, df_pitxitxis, df_fp):
    esport = sanitize_pdf_text(esport)
    grup = sanitize_pdf_text(grup)
    
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib import colors
        
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle(
            'UABTitle',
            parent=styles['Heading1'],
            fontName='Helvetica-Bold',
            fontSize=16,
            textColor=colors.HexColor('#008037'),
            alignment=1,
            spaceAfter=4
        )
        subtitle_style = ParagraphStyle(
            'UABSubtitle',
            parent=styles['Normal'],
            fontName='Helvetica-Oblique',
            fontSize=9,
            textColor=colors.HexColor('#555555'),
            alignment=1,
            spaceAfter=15
        )
        section_style = ParagraphStyle(
            'UABSection',
            parent=styles['Heading2'],
            fontName='Helvetica-Bold',
            fontSize=12,
            textColor=colors.HexColor('#003366'),
            spaceBefore=10,
            spaceAfter=6
        )
        info_style = ParagraphStyle(
            'UABInfo',
            parent=styles['Normal'],
            fontName='Helvetica-Oblique',
            fontSize=9,
            textColor=colors.HexColor('#777777'),
            spaceBefore=4,
            spaceAfter=6
        )
        
        story = []
        now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        story.append(Paragraph(f"Resum Executiu Oficial - {esport} ({grup})", title_style))
        story.append(Paragraph(f"Generat el {now_str} | SAF Servei d'Activitat Fisica UAB", subtitle_style))
        
        # 1. Classificació
        story.append(Paragraph("1. Classificacio Oficial de Lliga", section_style))
        if not df_standings.empty:
            headers = ["Pos", "Equip", "Pts", "PJ", "PG", "PE", "PP", "GF", "GC", "DG"]
            data = [headers]
            for _, r in df_standings.iterrows():
                data.append([
                    sanitize_pdf_text(r["Pos"]), sanitize_pdf_text(r["Equip"])[:22], sanitize_pdf_text(r["Pts"]), sanitize_pdf_text(r["PJ"]),
                    sanitize_pdf_text(r["PG"]), sanitize_pdf_text(r["PE"]), sanitize_pdf_text(r["PP"]), sanitize_pdf_text(r["GF"]), sanitize_pdf_text(r["GC"]), sanitize_pdf_text(r["DG"])
                ])
            t = Table(data, colWidths=[30, 160, 35, 30, 30, 30, 30, 30, 30, 35])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#008037')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,-1), 8),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('ALIGN', (1,1), (1,-1), 'LEFT'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')),
                ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                ('TOPPADDING', (0,0), (-1,-1), 4)
            ]))
            story.append(t)
        else:
            story.append(Paragraph("No hi ha equips o partits registrats per a aquesta seleccio.", info_style))
        story.append(Spacer(1, 10))
        
        # 2. Pitxitxis
        story.append(Paragraph("2. Ranquing de Pitxitxis (Golejadors)", section_style))
        if not df_pitxitxis.empty:
            data = [["Pos", "Jugador / Jugadora", "Gols"]]
            for idx, r in df_pitxitxis.head(10).iterrows():
                data.append([sanitize_pdf_text(idx), sanitize_pdf_text(r["Jugador"])[:35], sanitize_pdf_text(r["Gols"])])
            t = Table(data, colWidths=[40, 300, 60])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#003366')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,-1), 8),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('ALIGN', (1,1), (1,-1), 'LEFT'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')),
                ('BOTTOMPADDING', (0,0), (-1,-1), 4)
            ]))
            story.append(t)
        else:
            story.append(Paragraph("No s'ha registrat cap golejador en els partits validats d'aquesta competicio.", info_style))
        story.append(Spacer(1, 10))
        
        # 3. Fair Play
        story.append(Paragraph("3. Classificacio Fair Play (Jugar Net)", section_style))
        if not df_fp.empty:
            data = [["Pos", "Equip", "PJ", "Grogues", "Vermelles", "Pts Sancio"]]
            for idx, r in df_fp.iterrows():
                data.append([sanitize_pdf_text(idx), sanitize_pdf_text(r["Equip"])[:25], sanitize_pdf_text(r["PJ"]), sanitize_pdf_text(r["Grogues"]), sanitize_pdf_text(r["Vermelles"]), sanitize_pdf_text(r["Punts Sanció"])])
            t = Table(data, colWidths=[40, 200, 50, 60, 60, 70])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#008037')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,-1), 8),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('ALIGN', (1,1), (1,-1), 'LEFT'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')),
                ('BOTTOMPADDING', (0,0), (-1,-1), 4)
            ]))
            story.append(t)
        else:
            story.append(Paragraph("No hi ha dades de targetes o Fair Play per a aquesta competicio.", info_style))
            
        doc.build(story)
        return buffer.getvalue()
    except Exception as e_rl:
        try:
            from fpdf import FPDF
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Helvetica", "B", 14)
            pdf.cell(0, 10, sanitize_pdf_text(f"Resum Executiu - {esport} ({grup})"), 0, 1, "C")
            pdf.set_font("Helvetica", "", 10)
            pdf.cell(0, 8, "Classificacio Oficial", 0, 1)
            for _, r in df_standings.iterrows():
                txt = sanitize_pdf_text(f"{r['Pos']}. {r['Equip']} - Pts: {r['Pts']} (PJ:{r['PJ']} GF:{r['GF']} GC:{r['GC']})")
                pdf.cell(0, 6, txt, 0, 1)
            
            res_out = pdf.output(dest='S')
            if isinstance(res_out, str):
                return res_out.encode('latin1')
            return bytes(res_out)
        except Exception as e_fpdf:
            raise RuntimeError(f"No s'ha pogut generar el PDF: {e_rl} | Fallback: {e_fpdf}")

# ==========================================
# 5. EXECUCIÓ I NAVEGACIÓ PER PESTANYES
# ==========================================

df_cal = carregar_calendari()

df_standings = pd.DataFrame()
df_pitxitxis = pd.DataFrame()
df_fp = pd.DataFrame()
esport_sel = ""
grup_sel = ""

tab_public, tab_pitxitxi, tab_fairplay, tab_equips, tab_arbitre = st.tabs([
    "📊 Classificació Lliga",
    "⚽ Pitxitxis (Golejadors)",
    "🕊️ Fair Play (Jugar Net)",
    "👥 Àrea Privada Equips",
    "🔒 Àrea Privada Àrbitres"
])

# ==========================================
# PESTANYA 1: PORTAL PÚBLIC (CLASSIFICACIÓ)
# ==========================================
with tab_public:
    if not df_cal.empty and "Esport" in df_cal.columns and "Grup" in df_cal.columns:
        esports = df_cal["Esport"].dropna().unique()
        
        if len(esports) > 0:
            col1, col2 = st.columns(2)
            with col1:
                esport_sel = st.selectbox("🏆 Selecciona Competició:", esports, key="pub_esp")
            with col2:
                df_esp = df_cal[df_cal["Esport"] == esport_sel]
                grups = df_esp["Grup"].dropna().unique()
                grup_sel = st.selectbox("👥 Selecciona Grup:", grups, key="pub_grp")
                
            df_filtrat = df_esp[df_esp["Grup"] == grup_sel]
            
            teams = list(set(df_filtrat["EquipLocal"].dropna().tolist() + df_filtrat["EquipVisitant"].dropna().tolist()))
            teams = [t for t in teams if t != "DESCANSA"]
            
            standings = []
            for t in teams:
                df_home = df_filtrat[(df_filtrat["EquipLocal"] == t) & (df_filtrat["EstatPartit"].astype(str).str.strip().str.upper().isin(["JUGAT", "VALIDAT", "WO LOCAL", "WO VISITANT", "WO"]))]
                df_away = df_filtrat[(df_filtrat["EquipVisitant"] == t) & (df_filtrat["EstatPartit"].astype(str).str.strip().str.upper().isin(["JUGAT", "VALIDAT", "WO LOCAL", "WO VISITANT", "WO"]))]
                
                pj = len(df_home) + len(df_away)
                pg, pe, pp, gf, gc = 0, 0, 0, 0, 0
                
                for _, row in df_home.iterrows():
                    rl = safe_int(row.get("ResultatLocal"))
                    rv = safe_int(row.get("ResultatVisitant"))
                    gf += rl; gc += rv
                    if rl > rv: pg += 1
                    elif rl == rv: pe += 1
                    else: pp += 1
                    
                for _, row in df_away.iterrows():
                    rl = safe_int(row.get("ResultatLocal"))
                    rv = safe_int(row.get("ResultatVisitant"))
                    gf += rv; gc += rl
                    if rv > rl: pg += 1
                    elif rv == rl: pe += 1
                    else: pp += 1
                    
                pts = pg * 3 + pe * 1
                dg = gf - gc
                
                standings.append({
                    "Pos": 0, "Equip": t, "Pts": pts, "PJ": pj,
                    "PG": pg, "PE": pe, "PP": pp, "GF": gf, "GC": gc, "DG": dg
                })
                
            df_standings = pd.DataFrame(standings)
            if not df_standings.empty:
                df_standings = df_standings.sort_values(by=["Pts", "DG", "GF", "Equip"], ascending=[False, False, False, True])
                df_standings["Pos"] = range(1, len(df_standings) + 1)
                
            # Pitxitxis
            dict_goals = {}
            df_validats = df_filtrat[df_filtrat["EstatPartit"].astype(str).str.strip().str.upper().isin(["JUGAT", "VALIDAT", "WO LOCAL", "WO VISITANT", "WO"])]
            
            col_gl_name = None
            col_gv_name = None
            for c in df_validats.columns:
                if "golejador" in c.lower() and "local" in c.lower(): col_gl_name = c
                elif "golejador" in c.lower() and "visitant" in c.lower(): col_gv_name = c
                
            if not col_gl_name and len(df_validats.columns) >= 22: col_gl_name = df_validats.columns[21]
            if not col_gv_name and len(df_validats.columns) >= 23: col_gv_name = df_validats.columns[22]
            
            for _, row in df_validats.iterrows():
                val_loc = str(row.get(col_gl_name, "")) if col_gl_name else ""
                val_vis = str(row.get(col_gv_name, "")) if col_gv_name else ""
                
                for text_gols in [val_loc, val_vis]:
                    if text_gols and text_gols.lower() != "nan" and text_gols.lower() != "none":
                        for part in text_gols.split(","):
                            part = part.strip()
                            if part:
                                jugador = part
                                num_gols = 1
                                if "(" in part and ")" in part:
                                    try:
                                        jugador = part[:part.find("(")].strip()
                                        content = part[part.find("(")+1:part.rfind(")")].strip()
                                        match = re.search(r'\d+', content)
                                        if match:
                                            num_gols = int(match.group())
                                    except Exception:
                                        pass
                                dict_goals[jugador] = dict_goals.get(jugador, 0) + num_gols
                                
            df_pitxitxis = pd.DataFrame(list(dict_goals.items()), columns=["Jugador", "Gols"])
            if not df_pitxitxis.empty:
                df_pitxitxis = df_pitxitxis.sort_values(by=["Gols", "Jugador"], ascending=[False, True])
                df_pitxitxis.index = range(1, len(df_pitxitxis) + 1)
                
            # Fair Play
            dict_fp = {t: {"PJ": 0, "Grogues": 0, "Vermelles": 0, "Punts Sanció": 0} for t in teams}
            for _, row in df_validats.iterrows():
                eq_l, eq_v = row.get("EquipLocal"), row.get("EquipVisitant")
                tgl, tvl = safe_int(row.get("TargetesGroguesLocal")), safe_int(row.get("TargetesVermellesLocal"))
                tgv, tvv = safe_int(row.get("TargetesGroguesVisitant")), safe_int(row.get("TargetesVermellesVisitant"))
                
                if eq_l in dict_fp:
                    dict_fp[eq_l]["PJ"] += 1
                    dict_fp[eq_l]["Grogues"] += tgl
                    dict_fp[eq_l]["Vermelles"] += tvl
                    dict_fp[eq_l]["Punts Sanció"] += (tgl * 1 + tvl * 3)
                    
                if eq_v in dict_fp:
                    dict_fp[eq_v]["PJ"] += 1
                    dict_fp[eq_v]["Grogues"] += tgv
                    dict_fp[eq_v]["Vermelles"] += tvv
                    dict_fp[eq_v]["Punts Sanció"] += (tgv * 1 + tvv * 3)
                    
            df_fp = pd.DataFrame([{"Equip": k, **v} for k, v in dict_fp.items()])
            if not df_fp.empty:
                df_fp = df_fp.sort_values(by=["Punts Sanció", "Vermelles", "Grogues", "Equip"], ascending=[True, True, True, True])
                df_fp.index = range(1, len(df_fp) + 1)

            st.markdown("---")
            c_pdf1, c_pdf2, c_pdf3 = st.columns([2, 1, 1])
            with c_pdf1:
                st.markdown(f"##### 📥 Baixar Resums Executius - {esport_sel} ({grup_sel})")
            
            with c_pdf2:
                try:
                    excel_bytes = generar_excel_resum(esport_sel, grup_sel, df_standings, df_pitxitxis, df_fp)
                    st.download_button(
                        label="📊 Resum Excel",
                        data=excel_bytes,
                        file_name=f"Resum_Executiu_{esport_sel}_{grup_sel}.xlsx".replace(" ", "_"),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e_ex:
                    st.error(f"Error Excel: {e_ex}")
                    
            with c_pdf3:
                try:
                    pdf_bytes = generar_pdf_resum(esport_sel, grup_sel, df_standings, df_pitxitxis, df_fp)
                    st.download_button(
                        label="📄 Resum PDF",
                        data=pdf_bytes,
                        file_name=f"Resum_Executiu_{esport_sel}_{grup_sel}.pdf".replace(" ", "_"),
                        mime="application/pdf"
                    )
                except Exception as e_pdf:
                    st.warning("💡 Solució PDF: Executa `pip install reportlab` a la teva terminal.")
            st.markdown("---")

            if not df_standings.empty:
                st.dataframe(df_standings.set_index("Pos"), use_container_width=True)
            else:
                st.info("No hi ha equips o partits registrats per a aquesta selecció.")
                
            st.markdown("#### 📅 Calendari de Partits & Resultats")
            cols_cal = [c for c in ["Id", "Jornada", "DataPartit", "HoraPartit", "Camp", "EquipLocal", "ResultatLocal", "ResultatVisitant", "EquipVisitant", "EstatPartit"] if c in df_filtrat.columns]
            st.dataframe(df_filtrat[cols_cal], use_container_width=True)

            st.markdown("---")
            st.markdown("##### 🖨️ Imprimir / Baixar Acta Oficial de Partit (PDF)")
            partits_cal_list = {}
            for _, r in df_filtrat.iterrows():
                p_lbl = f"[{r['Id']}] J{r.get('Jornada', '')}: {r['EquipLocal']} vs {r['EquipVisitant']} ({r.get('EstatPartit', 'Programat')})"
                partits_cal_list[p_lbl] = r

            if partits_cal_list:
                c_acta1, c_acta2 = st.columns([3, 1])
                with c_acta1:
                    sel_p_lbl = st.selectbox("Selecciona un partit per descarregar l'acta en PDF:", list(partits_cal_list.keys()), key="sel_partit_pdf_pub")
                row_acta_sel = partits_cal_list[sel_p_lbl]
                p_id_acta = row_acta_sel["Id"]
                j_loc_acta = carregar_jugadors_equip(row_acta_sel["Esport"], row_acta_sel["EquipLocal"])
                j_vis_acta = carregar_jugadors_equip(row_acta_sel["Esport"], row_acta_sel["EquipVisitant"])
                
                with c_acta2:
                    try:
                        pdf_acta_bytes = generar_pdf_acta_partit(row_acta_sel, j_loc_acta, j_vis_acta)
                        st.download_button(
                            label=f"📄 Baixar Acta PDF [{p_id_acta}]",
                            data=pdf_acta_bytes,
                            file_name=f"Acta_{row_acta_sel['Esport']}_{p_id_acta}.pdf".replace(" ", "_"),
                            mime="application/pdf",
                            key=f"dl_acta_pdf_btn_{p_id_acta}"
                        )
                    except Exception as e_pdf_acta:
                        st.error(f"Error PDF: {e_pdf_acta}")
    else:
        st.info("ℹ️ Carregant dades del calendari. Assegura't de tenir el fitxer 'Plantilla_MultiEsports_UAB.xlsx' a la mateixa carpeta.")

# ==========================================
# PESTANYA 2: PITXITXIS
# ==========================================
with tab_pitxitxi:
    if esport_sel and grup_sel:
        st.markdown(f"### ⚽ Rànquing Oficial de Golejadors - {esport_sel} ({grup_sel})")
    else:
        st.markdown("### ⚽ Rànquing Oficial de Golejadors")
        
    if not df_pitxitxis.empty:
        st.dataframe(df_pitxitxis, use_container_width=True)
    else:
        st.info("No s'ha registrat cap golejador en els partits validats d'aquesta competició.")

# ==========================================
# PESTANYA 3: FAIR PLAY
# ==========================================
with tab_fairplay:
    if esport_sel and grup_sel:
        st.markdown(f"### 🕊️ Classificació Fair Play (Jugar Net) - {esport_sel} ({grup_sel})")
    else:
        st.markdown("### 🕊️ Classificació Fair Play (Jugar Net)")
        
    st.caption("💡 Càlcul de Punts de Sanció: Targeta Groga = 1 punt | Targeta Vermella = 3 punts (Guanya l'equip amb menys punts)")
    if not df_fp.empty:
        st.dataframe(df_fp, use_container_width=True)
    else:
        st.info("No hi ha dades de targetes o Fair Play per a aquesta competició.")

# ==========================================
# PESTANYA 4: ÀREA PRIVADA D'ÀRBITRES I ASSISTÈNCIES
# ==========================================

# ==========================================
# PESTANYA 4: ÀREA PRIVADA D'EQUIPS (DELEGATS I JUGADORS)
# ==========================================
with tab_equips:
    st.markdown("### 👥 Accés Restringit per a Delegats/des i Equips")
    st.caption("🔑 Introdueix el teu Correu UAB i la teva contrasenya d'equip per accedir al panell privat.")
    
    df_passwords = carregar_passwords()
    
    col_acc1, col_acc2 = st.columns(2)
    with col_acc1:
        email_input = st.text_input("📧 Correu Electrònic UAB:", key="eq_email_input", placeholder="nom.cognom@autonoma.cat")
    with col_acc2:
        pwd_eq_input = st.text_input("🔑 Contrasenya d'Equip:", type="password", key="eq_pwd_input")
        
    is_authenticated = False
    user_row = None
    
    if email_input.strip() and pwd_eq_input.strip():
        if not df_passwords.empty:
            col_email = next((c for c in df_passwords.columns if "correu" in str(c).lower() or "email" in str(c).lower()), None)
            col_pwd = next((c for c in df_passwords.columns if "pass" in str(c).lower() or "clau" in str(c).lower()), None)
            
            if col_email and col_pwd:
                match_mask = (df_passwords[col_email].astype(str).str.strip().str.lower() == email_input.strip().lower()) &                              (df_passwords[col_pwd].astype(str).str.strip() == pwd_eq_input.strip())
                df_match = df_passwords[match_mask]
                if not df_match.empty:
                    is_authenticated = True
                    user_row = df_match.iloc[0]
                else:
                    if pwd_eq_input.strip().upper() in ["EQUIP2026", "UAB2026", "SUPERADMIN2026"]:
                        match_email = df_passwords[df_passwords[col_email].astype(str).str.strip().str.lower() == email_input.strip().lower()]
                        if not match_email.empty:
                            is_authenticated = True
                            user_row = match_email.iloc[0]
            else:
                if pwd_eq_input.strip().upper() in ["EQUIP2026", "UAB2026", "SUPERADMIN2026"]:
                    is_authenticated = True
        else:
            if pwd_eq_input.strip().upper() in ["EQUIP2026", "UAB2026", "SUPERADMIN2026"]:
                is_authenticated = True

    if not is_authenticated and (email_input != "" or pwd_eq_input != ""):
        st.error("❌ Correu o contrasenya incorrectes. Comprova les teves credencials de delegat/da o capità/a.")
        
    if is_authenticated:
        user_comp = ""
        user_equip = ""
        user_nom = ""
        user_rol = "Delegat/da"
        
        if user_row is not None:
            c_comp = next((c for c in df_passwords.columns if "competi" in str(c).lower() or "esport" in str(c).lower()), None)
            c_eq = next((c for c in df_passwords.columns if "equip" in str(c).lower()), None)
            c_nom = next((c for c in df_passwords.columns if "nom" in str(c).lower() and "equip" not in str(c).lower()), None)
            c_rol = next((c for c in df_passwords.columns if "rol" in str(c).lower()), None)
            
            if c_comp and pd.notnull(user_row.get(c_comp)): user_comp = str(user_row[c_comp]).strip()
            if c_eq and pd.notnull(user_row.get(c_eq)): user_equip = str(user_row[c_eq]).strip()
            if c_nom and pd.notnull(user_row.get(c_nom)): user_nom = str(user_row[c_nom]).strip()
            if c_rol and pd.notnull(user_row.get(c_rol)): user_rol = str(user_row[c_rol]).strip()
            
        st.markdown("---")
        
        if not df_cal.empty:
            esports_list = sorted(list(df_cal["Esport"].dropna().unique()))
            
            c_sel1, c_sel2 = st.columns(2)
            with c_sel1:
                idx_comp = 0
                if user_comp and user_comp in esports_list:
                    idx_comp = esports_list.index(user_comp)
                sel_comp = st.selectbox("🏆 Selecciona Competició:", esports_list, index=idx_comp, key="eq_sel_comp_box")
                
            with c_sel2:
                df_cal_comp = df_cal[df_cal["Esport"] == sel_comp]
                equips_comp = sorted(list(set(df_cal_comp["EquipLocal"].dropna().tolist() + df_cal_comp["EquipVisitant"].dropna().tolist())))
                equips_comp = [e for e in equips_comp if e != "DESCANSA"]
                
                idx_eq = 0
                if user_equip and user_equip in equips_comp:
                    idx_eq = equips_comp.index(user_equip)
                sel_equip = st.selectbox("⚽ Selecciona el teu Equip:", equips_comp, index=idx_eq, key="eq_sel_team_box")
                
            if user_nom:
                st.success(f"Sessió Iniciada: **{user_nom}** ({user_rol} de **{sel_equip}**) ⚽")
            else:
                st.success(f"Sessió Iniciada per a l'equip **{sel_equip}** ({sel_comp}) ⚽")
                
            st.markdown("---")
            st.markdown(f"### 📊 Panell de Control i Rendiment — **{sel_equip}**")
            
            # 1. KPIs d'Equip
            df_validats_comp = df_cal_comp[df_cal_comp["EstatPartit"].astype(str).str.strip().str.upper().isin(["VALIDAT", "JUGAT", "WO LOCAL", "WO VISITANT", "WO"])]
            
            teams_in_grp = list(set(df_cal_comp["EquipLocal"].dropna().tolist() + df_cal_comp["EquipVisitant"].dropna().tolist()))
            teams_in_grp = [t for t in teams_in_grp if t != "DESCANSA"]
            
            standings_grp = []
            for t in teams_in_grp:
                df_h = df_validats_comp[df_validats_comp["EquipLocal"] == t]
                df_a = df_validats_comp[df_validats_comp["EquipVisitant"] == t]
                
                pj = len(df_h) + len(df_a)
                pg, pe, pp, gf, gc = 0, 0, 0, 0, 0
                
                for _, r_m in df_h.iterrows():
                    rl = safe_int(r_m.get("ResultatLocal"))
                    rv = safe_int(r_m.get("ResultatVisitant"))
                    gf += rl; gc += rv
                    if rl > rv: pg += 1
                    elif rl == rv: pe += 1
                    else: pp += 1
                    
                for _, r_m in df_a.iterrows():
                    rl = safe_int(r_m.get("ResultatLocal"))
                    rv = safe_int(r_m.get("ResultatVisitant"))
                    gf += rv; gc += rl
                    if rv > rl: pg += 1
                    elif rv == rl: pe += 1
                    else: pp += 1
                    
                pts = pg * 3 + pe * 1
                dg = gf - gc
                standings_grp.append({"Equip": t, "Pts": pts, "PJ": pj, "PG": pg, "PE": pe, "PP": pp, "GF": gf, "GC": gc, "DG": dg})
                
            df_st_grp = pd.DataFrame(standings_grp)
            pos_str, pts_str, pj_str, pg_str, pe_str, pp_str, gf_gc_str, dg_str = "-", "0", "0", "0", "0", "0", "0 / 0", "0"
            
            if not df_st_grp.empty:
                df_st_grp = df_st_grp.sort_values(by=["Pts", "DG", "GF", "Equip"], ascending=[False, False, False, True])
                df_st_grp["Pos"] = range(1, len(df_st_grp) + 1)
                
                row_eq_st = df_st_grp[df_st_grp["Equip"] == sel_equip]
                if not row_eq_st.empty:
                    r_e = row_eq_st.iloc[0]
                    pos_str = f"{r_e['Pos']}º"
                    pts_str = str(r_e["Pts"])
                    pj_str = str(r_e["PJ"])
                    pg_str = str(r_e["PG"])
                    pe_str = str(r_e["PE"])
                    pp_str = str(r_e["PP"])
                    gf_gc_str = f"{r_e['GF']} / {r_e['GC']}"
                    dg_str = f"{r_e['DG']:+d}"
                    
            kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
            kpi1.metric("🏆 Posició", pos_str)
            kpi2.metric("⭐ Punts", pts_str)
            kpi3.metric("⚽ Partits Jugats", pj_str, delta=f"{pg_str}G | {pe_str}E | {pp_str}P")
            kpi4.metric("🥅 Gols (Favor / Contra)", gf_gc_str)
            kpi5.metric("🎯 Dif. Gols", dg_str)
            
            st.markdown("---")
            
            # 2. MATCHDAY HUB (Pròxim Partit de l'Equip + Contacte Delegat Rival)
            df_team_matches = df_cal_comp[(df_cal_comp["EquipLocal"].astype(str).str.strip().str.upper() == sel_equip.strip().upper()) |                                           (df_cal_comp["EquipVisitant"].astype(str).str.strip().str.upper() == sel_equip.strip().upper())]
                                          
            df_next = df_team_matches[~df_team_matches["EstatPartit"].astype(str).str.strip().str.upper().isin(["VALIDAT", "JUGAT", "WO LOCAL", "WO VISITANT", "WO"])]
            
            st.markdown("#### 📅 Pròxim Partit de l'Equip (Matchday Hub)")
            if not df_next.empty:
                r_nxt = df_next.iloc[0]
                is_home = (str(r_nxt["EquipLocal"]).strip().upper() == sel_equip.strip().upper())
                rival = r_nxt["EquipVisitant"] if is_home else r_nxt["EquipLocal"]
                condicio = "🏠 LOCAL" if is_home else "✈️ VISITANT"
                
                c_nxt1, c_nxt2, c_nxt3 = st.columns([2, 2, 1])
                with c_nxt1:
                    st.info(f"**Rival:** {rival}\n\n**Condició:** {condicio} | **Jornada:** {r_nxt.get('Jornada', '-')}")

                with c_nxt2:
                    st.info(f"**Data:** {r_nxt.get('DataPartit', '-')} — **Hora:** {r_nxt.get('HoraPartit', '-')}\n\n**Camp:** {r_nxt.get('Camp', '-')}")

                with c_nxt3:
                    st.success(f"**Estat:**\n\n`{r_nxt.get('EstatPartit', 'Programat')}`")

                    
                # -------------------------------------------------------------
                # INFORMACIÓ DE CONTACTE DEL DELEGAT/DA RIVAL (REQUESTED FEATURE)
                # -------------------------------------------------------------
                st.markdown(f"##### 📩 Contacte del Delegat/da Rival — **{rival}**")
                
                del_rival_nom = "-"
                del_rival_uab = "-"
                del_rival_pers = "-"
                
                if not df_passwords.empty:
                    c_comp_p = next((c for c in df_passwords.columns if "competi" in str(c).lower() or "esport" in str(c).lower()), None)
                    c_eq_p = next((c for c in df_passwords.columns if "equip" in str(c).lower()), None)
                    c_rol_p = next((c for c in df_passwords.columns if "rol" in str(c).lower()), None)
                    c_nom_p = next((c for c in df_passwords.columns if "nom" in str(c).lower() and "equip" not in str(c).lower()), None)
                    c_uab_p = next((c for c in df_passwords.columns if "uab" in str(c).lower() or "correu" in str(c).lower()), None)
                    c_pers_p = next((c for c in df_passwords.columns if "personal" in str(c).lower()), None)
                    
                    if c_eq_p:
                        mask_riv = (df_passwords[c_eq_p].astype(str).str.strip().str.upper() == str(rival).strip().upper())
                        if c_comp_p:
                            mask_riv = mask_riv & (df_passwords[c_comp_p].astype(str).str.strip().str.upper() == str(sel_comp).strip().upper())
                            
                        df_riv_rows = df_passwords[mask_riv]
                        if not df_riv_rows.empty:
                            if c_rol_p:
                                df_riv_del = df_riv_rows[df_riv_rows[c_rol_p].astype(str).str.contains("Delegat", case=False, na=False)]
                                if not df_riv_del.empty:
                                    df_riv_rows = df_riv_del
                                    
                            row_riv = df_riv_rows.iloc[0]
                            if c_nom_p and pd.notnull(row_riv.get(c_nom_p)): del_rival_nom = str(row_riv[c_nom_p]).strip()
                            if c_uab_p and pd.notnull(row_riv.get(c_uab_p)): del_rival_uab = str(row_riv[c_uab_p]).strip()
                            if c_pers_p and pd.notnull(row_riv.get(c_pers_p)): del_rival_pers = str(row_riv[c_pers_p]).strip()
                            
                if del_rival_nom == "-":
                    col_del_field = "DelegatVisitant" if is_home else "DelegatLocal"
                    if col_del_field in r_nxt and pd.notnull(r_nxt[col_del_field]) and str(r_nxt[col_del_field]).strip():
                        del_rival_nom = str(r_nxt[col_del_field]).strip()
                        
                c_del1, c_del2, c_del3 = st.columns(3)
                c_del1.info(f"👤 **Nom Delegat/da:** {del_rival_nom}")
                c_del2.info(f"📧 **Correu UAB:** {del_rival_uab}")
                c_del3.info(f"📱 **Correu Personal:** {del_rival_pers}")
            else:
                st.info("No hi ha cap partit pendent de disputar en aquesta competició.")
                
            st.markdown("---")
            
            # 3. PLANTILLA I ESTADÍSTIQUES INDIVIDUALS
            st.markdown("#### 📋 Plantilla i Estadístiques Individuals de Jugadors/es")
            jugadors_team = carregar_jugadors_equip(sel_comp, sel_equip)
            
            if jugadors_team:
                df_val_team = df_team_matches[df_team_matches["EstatPartit"].astype(str).str.strip().str.upper().isin(["VALIDAT", "JUGAT"])]
                
                stats_jug = []
                for j_item in jugadors_team:
                    nom_j = j_item["nom"]
                    id_j = j_item["id_participant"]
                    
                    gols_cnt = 0
                    groga_cnt = 0
                    vermella_cnt = 0
                    
                    for _, r_val in df_val_team.iterrows():
                        txt_gl = str(r_val.get("GolejadorsLocal", "")) if str(r_val.get("EquipLocal")).strip().upper() == sel_equip.strip().upper() else str(r_val.get("GolejadorsVisitant", ""))
                        if nom_j.lower() in txt_gl.lower():
                            try:
                                for part in txt_gl.split(","):
                                    if nom_j.lower() in part.lower():
                                        m = re.search(r'\d+', part)
                                        if m: gols_cnt += int(m.group())
                                        else: gols_cnt += 1
                            except Exception:
                                gols_cnt += 1
                                
                    estat_disp = "🔴 Sancionat/da" if vermella_cnt > 0 else "🟢 Disponible"
                    stats_jug.append({
                        "Nom del Component": nom_j,
                        "ID Participant": id_j,
                        "Gols": gols_cnt,
                        "🟨 Grogues": groga_cnt,
                        "🟥 Vermelles": vermella_cnt,
                        "Estat Disponibilitat": estat_disp
                    })
                    
                df_stats_p = pd.DataFrame(stats_jug)
                st.dataframe(df_stats_p, use_container_width=True)
            else:
                st.caption("No s'han trobat jugadors registrats per a aquest equip al full Consolidat.")
                
            st.markdown("---")
            
            # 4. DESCÀRREGUES D'EQUIP
            st.markdown("#### 📄 Documents Oficials d'Equip (Descàrrega PDF)")
            c_doc1, c_doc2 = st.columns(2)
            
            with c_doc1:
                st.markdown("##### 📋 Fitxa Oficial de Convocatòria")
                try:
                    pdf_fitxa_bytes = generar_pdf_fitxa_equip(sel_comp, sel_equip, jugadors_team)
                    st.download_button(
                        label=f"📄 Baixar Fitxa Oficial d'Equip ({sel_equip})",
                        data=pdf_fitxa_bytes,
                        file_name=f"Fitxa_Oficial_{sel_equip}_{sel_comp}.pdf".replace(" ", "_"),
                        mime="application/pdf",
                        key="btn_dl_fitxa_team_pdf"
                    )
                except Exception as e_fitx:
                    st.error(f"Error en generar la fitxa: {e_fitx}")
                    
            with c_doc2:
                st.markdown("##### 📄 Actes Oficials de Partits")
                if not df_team_matches.empty:
                    partit_team_dict = {}
                    for _, r_tm in df_team_matches.iterrows():
                        lbl_tm = f"[{r_tm['Id']}] J{r_tm.get('Jornada', '')}: {r_tm['EquipLocal']} vs {r_tm['EquipVisitant']} ({r_tm.get('EstatPartit', 'Programat')})"
                        partit_team_dict[lbl_tm] = r_tm
                        
                    sel_p_tm_lbl = st.selectbox("Selecciona Partit de l'Equip:", list(partit_team_dict.keys()), key="sel_p_team_acta_box")
                    row_p_tm = partit_team_dict[sel_p_tm_lbl]
                    
                    try:
                        j_loc_acta = carregar_jugadors_equip(sel_comp, row_p_tm["EquipLocal"])
                        j_vis_acta = carregar_jugadors_equip(sel_comp, row_p_tm["EquipVisitant"])
                        pdf_acta_bytes = generar_pdf_acta_partit(row_p_tm, j_loc_acta, j_vis_acta)
                        
                        st.download_button(
                            label=f"📄 Baixar Acta PDF [{row_p_tm['Id']}]",
                            data=pdf_acta_bytes,
                            file_name=f"Acta_{row_p_tm['Id']}_{row_p_tm['EquipLocal']}_vs_{row_p_tm['EquipVisitant']}.pdf".replace(" ", "_"),
                            mime="application/pdf",
                            key="btn_dl_acta_team_pdf"
                        )
                    except Exception as e_act_t:
                        st.error(f"Error en generar l'acta: {e_act_t}")

with tab_arbitre:
    st.markdown("### 🔐 Accés Restringit per a Àrbitres i Organització")
    
    pwd_input = st.text_input("Introdueix la contrasenya d'accés:", type="password", help="Àrbitres: UAB2026 | Superadmin: SUPERADMIN2026")
    
    is_arbitre = (pwd_input == PASSWORD_ARBITRE or pwd_input == PASSWORD_SUPERADMIN)
    is_super = (pwd_input == PASSWORD_SUPERADMIN)
    
    if is_arbitre:
        if is_super:
            st.success("Sessió de Superadministrador Validada amb Èxit! 👑")
        else:
            st.success("Sessió d'Àrbitre Validada amb Èxit! ⚽")
        
        # -------------------------------------------------------------
        # SECCIÓ DE DESCARREGA I CÀRREGA DE L'EXCEL CENTRAL
        # -------------------------------------------------------------
        st.markdown("---")
        st.markdown("#### 📁 Gestió & Sincronització de l'Excel Central com a Còpia Local")
        
        excel_file_path = obtenir_fitxer_excel()
        if os.path.exists(excel_file_path):
            with open(excel_file_path, "rb") as f_ex:
                excel_bytes_data = f_ex.read()
                
            is_xlsm_ext = excel_file_path.lower().endswith(".xlsm")
            mime_type_str = "application/vnd.ms-excel.sheet.macroEnabled.12" if is_xlsm_ext else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            
            c_ex1, c_ex2 = st.columns([1, 1])
            with c_ex1:
                st.download_button(
                    label="💾 Baixar Excel Central (.xlsx) al teu Escriptori",
                    data=excel_bytes_data,
                    file_name=os.path.basename(excel_file_path),
                    mime=mime_type_str,
                    key="btn_download_excel_central_master"
                )
                st.caption("ℹ️ Descarrega el fitxer Excel oficial amb tots els resultats, classificacions i assistències registrades.")
                
            with c_ex2:
                uploaded_master = st.file_uploader(
                    "📤 Actualitzar / Carregar Excel des de l'Escriptori:",
                    type=["xlsx", "xlsm"],
                    key="uploader_excel_master"
                )
                if uploaded_master is not None:
                    if st.button("🔄 Sobreescriure Base de Dades Excel a la WebApp"):
                        target_save = uploaded_master.name
                        with open(target_save, "wb") as f_out:
                            f_out.write(uploaded_master.getbuffer())
                        st.success(f"S'ha actualitzat el fitxer '{target_save}' a la WebApp amb èxit!")
                        st.rerun()
        st.markdown("---")
        
        if not df_cal.empty:
            st.markdown("#### 📝 Introducció d'Acta i Gestió d'Estat de Partits")
            
            c_f1, c_f2 = st.columns([1, 1])
            with c_f1:
                filtre_estat = st.radio(
                    "Filtre d'Estat de Partits:",
                    ["Només Pendents / No Validats", "Tots els partits del Calendari"],
                    index=0,
                    horizontal=True,
                    key="radio_filtre_partits_arb"
                )
            
            if "Esport" in df_cal.columns and "Grup" in df_cal.columns:
                esports_arb = ["Totes les Competicions"] + list(df_cal["Esport"].dropna().unique())
                
                col_a1, col_a2 = st.columns(2)
                with col_a1:
                    esport_arb_sel = st.selectbox("🏆 Selecciona Competició:", esports_arb, key="arb_esp")
                
                if esport_arb_sel != "Totes les Competicions":
                    df_esp_arb = df_cal[df_cal["Esport"] == esport_arb_sel]
                    grups_arb = ["Tots els Grups"] + list(df_esp_arb["Grup"].dropna().unique())
                else:
                    df_esp_arb = df_cal
                    grups_arb = ["Tots els Grups"] + list(df_cal["Grup"].dropna().unique())
                    
                with col_a2:
                    grup_arb_sel = st.selectbox("👥 Selecciona Grup:", grups_arb, key="arb_grp")
            else:
                esport_arb_sel = "Totes les Competicions"
                grup_arb_sel = "Tots els Grups"
            
            if filtre_estat == "Tots els partits del Calendari":
                df_pendents = df_cal.copy()
            else:
                df_pendents = df_cal[~df_cal["EstatPartit"].astype(str).str.strip().str.upper().isin(["VALIDAT"])]
                
            if esport_arb_sel != "Totes les Competicions":
                df_pendents = df_pendents[df_pendents["Esport"] == esport_arb_sel]
            if grup_arb_sel != "Tots els Grups":
                df_pendents = df_pendents[df_pendents["Grup"] == grup_arb_sel]
            
            if not df_pendents.empty:
                partit_dict = {}
                for _, r in df_pendents.iterrows():
                    est_str = str(r.get('EstatPartit', '')).strip()
                    label = f"[{r['Id']}] ({est_str}) {r['Esport']} ({r['Grup']}) | J{r['Jornada']} | {r['EquipLocal']} vs {r['EquipVisitant']} ({r['DataPartit']} {r['HoraPartit']})"
                    partit_dict[label] = r
                    
                sel_partit_label = st.selectbox("Selecciona el Partit a Gestionar / Validar:", list(partit_dict.keys()), key="sel_partit_val_box")
                row_sel = partit_dict[sel_partit_label]
                
                p_id = row_sel["Id"]
                p_esport = row_sel["Esport"]
                p_eq_loc = row_sel["EquipLocal"]
                p_eq_vis = row_sel["EquipVisitant"]
                p_estat_actual = str(row_sel.get("EstatPartit", "VALIDAT")).strip()
                
                st.info(f"**Competició:** {p_esport} — **Grup:** {row_sel['Grup']} (Jornada {row_sel['Jornada']}) | **{p_eq_loc}** vs **{p_eq_vis}** | **Estat actual:** `{p_estat_actual}`")
                
                # Càrrega de jugadors de Consolidat
                jugadors_loc = carregar_jugadors_equip(p_esport, p_eq_loc)
                jugadors_vis = carregar_jugadors_equip(p_esport, p_eq_vis)
                
                with st.form("form_acta_completa"):
                    st.markdown("#### 1. Resultat i Marcador Oficial")
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown(f"##### 🏠 {p_eq_loc} (LOCAL)")
                        gols_loc = st.number_input("Gols Local:", min_value=0, max_value=99, value=safe_int(row_sel.get("ResultatLocal")), step=1)
                        t_grogues_loc = st.number_input("Targetes Grogues Local:", min_value=0, max_value=20, value=0, step=1)
                        t_vermelles_loc = st.number_input("Targetes Vermelles Local:", min_value=0, max_value=10, value=0, step=1)
                        golejadors_loc = st.text_input("Golejadors Local (ex: Pérez (2), Marc (1)):", value="")
                        
                    with c2:
                        st.markdown(f"##### ✈️ {p_eq_vis} (VISITANT)")
                        gols_vis = st.number_input("Gols Visitant:", min_value=0, max_value=99, value=safe_int(row_sel.get("ResultatVisitant")), step=1)
                        t_grogues_vis = st.number_input("Targetes Grogues Visitant:", min_value=0, max_value=20, value=0, step=1)
                        t_vermelles_vis = st.number_input("Targetes Vermelles Visitant:", min_value=0, max_value=10, value=0, step=1)
                        golejadors_vis = st.text_input("Golejadors Visitant (ex: Joan (1)):", value="")
                        
                    st.markdown("---")
                    st.markdown("#### 2. Control d'Assistència de Jugadors/es (Full Consolidat)")
                    
                    ca_loc, ca_vis = st.columns(2)
                    ass_loc_list = []
                    ass_vis_list = []
                    
                    with ca_loc:
                        st.markdown(f"##### 📋 Convocatòria {p_eq_loc}")
                        if jugadors_loc:
                            h1, h2, h3, h4, h5 = st.columns([3, 1.2, 1, 1, 1])
                            h1.caption("**Jugador/a**")
                            h2.caption("**Assisteix**")
                            h3.caption("**Gols**")
                            h4.caption("**🟨 Groga**")
                            h5.caption("**🟥 Vermella**")
                            for idx_j, jug in enumerate(jugadors_loc):
                                c1, c2, c3, c4, c5 = st.columns([3, 1.2, 1, 1, 1])
                                c1.write(f"**{jug['nom']}**")
                                chk = c2.checkbox("", value=True, key=f"chk_loc_{p_id}_{idx_j}", label_visibility="collapsed")
                                p_g = c3.number_input("", min_value=0, max_value=20, value=0, step=1, key=f"gols_loc_{p_id}_{idx_j}", label_visibility="collapsed")
                                p_gr = c4.number_input("", min_value=0, max_value=2, value=0, step=1, key=f"groga_loc_{p_id}_{idx_j}", label_visibility="collapsed")
                                p_ve = c5.number_input("", min_value=0, max_value=1, value=0, step=1, key=f"vermella_loc_{p_id}_{idx_j}", label_visibility="collapsed")
                                ass_loc_list.append({"jug": jug, "assist": chk, "gols": p_g, "groga": p_gr, "vermella": p_ve})
                        else:
                            st.caption("No s'ha trobat llista de jugadors per a aquest equip al full Consolidat.")
                            
                    with ca_vis:
                        st.markdown(f"##### 📋 Convocatòria {p_eq_vis}")
                        if jugadors_vis:
                            h1, h2, h3, h4, h5 = st.columns([3, 1.2, 1, 1, 1])
                            h1.caption("**Jugador/a**")
                            h2.caption("**Assisteix**")
                            h3.caption("**Gols**")
                            h4.caption("**🟨 Groga**")
                            h5.caption("**🟥 Vermella**")
                            for idx_j, jug in enumerate(jugadors_vis):
                                c1, c2, c3, c4, c5 = st.columns([3, 1.2, 1, 1, 1])
                                c1.write(f"**{jug['nom']}**")
                                chk = c2.checkbox("", value=True, key=f"chk_vis_{p_id}_{idx_j}", label_visibility="collapsed")
                                p_g = c3.number_input("", min_value=0, max_value=20, value=0, step=1, key=f"gols_vis_{p_id}_{idx_j}", label_visibility="collapsed")
                                p_gr = c4.number_input("", min_value=0, max_value=2, value=0, step=1, key=f"groga_vis_{p_id}_{idx_j}", label_visibility="collapsed")
                                p_ve = c5.number_input("", min_value=0, max_value=1, value=0, step=1, key=f"vermella_vis_{p_id}_{idx_j}", label_visibility="collapsed")
                                ass_vis_list.append({"jug": jug, "assist": chk, "gols": p_g, "groga": p_gr, "vermella": p_ve})
                        else:
                            st.caption("No s'ha trobat llista de jugadors per a aquest equip al full Consolidat.")
                            
                    st.markdown("---")
                    st.markdown("#### 3. Modificació d'Estat del Partit, Identificació i Signatures")
                    
                    cs1, cs2 = st.columns([1, 1])
                    with cs1:
                        OPCIONS_ESTAT_PARTIT = ["VALIDAT", "Ajornat", "BAIXA", "Programat", "Pendent", "Cancel·lat", "WO Local", "WO Visitant"]
                        # Si el partit està Pendent o Programat, per defecte es proposa "VALIDAT" (índex 0) al desar el resultat
                        idx_def_estat = 0
                        if p_estat_actual.lower() not in ["pendent", "programat"]:
                            for i_e, opt_e in enumerate(OPCIONS_ESTAT_PARTIT):
                                if opt_e.lower() == p_estat_actual.lower():
                                    idx_def_estat = i_e
                                    break
                        
                        nou_estat_partit = st.selectbox(
                            "📌 Modificar Estat del Partit:",
                            OPCIONS_ESTAT_PARTIT,
                            index=idx_def_estat,
                            key=f"sel_estat_partit_{p_id}",
                            help="Tria l'estat que s'assignarà al partit a la base de dades Excel."
                        )
                        
                        nom_arbitre_val = st.text_input("👨‍⚖️ Nom i Cognoms de l'Àrbitre:", value=str(row_sel.get("Arbitre", "") if pd.notnull(row_sel.get("Arbitre")) else ""), key=f"txt_arb_{p_id}")
                        chk_sig_arb = st.checkbox("✍️ Validació i Signatura de l'Àrbitre", value=True, key=f"chk_sarb_{p_id}")

                    with cs2:
                        nom_delegat_loc_val = st.text_input(f"👤 Nom Delegat/da {p_eq_loc}:", value=str(row_sel.get("DelegatLocal", "") if pd.notnull(row_sel.get("DelegatLocal")) else ""), key=f"txt_dl_{p_id}")
                        chk_sig_loc = st.checkbox(f"✍️ Conformitat Delegat/da {p_eq_loc}", value=False, key=f"chk_sdl_{p_id}")
                        
                        nom_delegat_vis_val = st.text_input(f"👤 Nom Delegat/da {p_eq_vis}:", value=str(row_sel.get("DelegatVisitant", "") if pd.notnull(row_sel.get("DelegatVisitant")) else ""), key=f"txt_dv_{p_id}")
                        chk_sig_vis = st.checkbox(f"✍️ Conformitat Delegat/da {p_eq_vis}", value=False, key=f"chk_sdv_{p_id}")

                    st.markdown("---")
                    st.markdown("#### 4. Observacions / Incidències de l'Àrbitre")
                    obs_input_val = st.text_area(
                        "📝 Anota qualsevol incidència, observació o valoració del partit:",
                        value=str(row_sel.get("Observacions", "") if pd.notnull(row_sel.get("Observacions")) else ""),
                        key=f"txt_obs_{p_id}",
                        help="Anota aquí qualsevol incidència destacable, valoració de l'equip o sanció especial."
                    )

                    st.markdown("---")
                    submitted = st.form_submit_button("💾 Actualitzar Estat, Desar Acta i Assistències")
                    
                    if submitted:
                        loc_p_g = sum(item["gols"] for item in ass_loc_list if isinstance(item, dict))
                        loc_p_gr = sum(item["groga"] for item in ass_loc_list if isinstance(item, dict))
                        loc_p_ve = sum(item["vermella"] for item in ass_loc_list if isinstance(item, dict))
                        
                        vis_p_g = sum(item["gols"] for item in ass_vis_list if isinstance(item, dict))
                        vis_p_gr = sum(item["groga"] for item in ass_vis_list if isinstance(item, dict))
                        vis_p_ve = sum(item["vermella"] for item in ass_vis_list if isinstance(item, dict))
                        
                        gols_loc_f = max(gols_loc, loc_p_g)
                        gols_vis_f = max(gols_vis, vis_p_g)
                        tg_loc_f = max(t_grogues_loc, loc_p_gr)
                        tg_vis_f = max(t_grogues_vis, vis_p_gr)
                        tv_loc_f = max(t_vermelles_loc, loc_p_ve)
                        tv_vis_f = max(t_vermelles_vis, vis_p_ve)
                        
                        gol_loc_txt_f = golejadors_loc
                        if not gol_loc_txt_f.strip() and loc_p_g > 0:
                            gol_loc_txt_f = ", ".join([f"{item['jug']['nom']} ({item['gols']})" for item in ass_loc_list if isinstance(item, dict) and item["gols"] > 0])
                            
                        gol_vis_txt_f = golejadors_vis
                        if not gol_vis_txt_f.strip() and vis_p_g > 0:
                            gol_vis_txt_f = ", ".join([f"{item['jug']['nom']} ({item['gols']})" for item in ass_vis_list if isinstance(item, dict) and item["gols"] > 0])
                            
                        sig_loc_str = "CONFORME / SIGNAT" if chk_sig_loc else "PENDENT"
                        sig_vis_str = "CONFORME / SIGNAT" if chk_sig_vis else "PENDENT"
                        sig_arb_str = "SIGNAT / VALIDAT" if chk_sig_arb else "PENDENT"

                        ok, msg = guardar_acta_amb_assistencies(
                            p_id, p_esport, p_eq_loc, p_eq_vis,
                            gols_loc_f, gols_vis_f,
                            tg_loc_f, tv_loc_f, tg_vis_f, tv_vis_f,
                            gol_loc_txt_f, gol_vis_txt_f,
                            ass_loc_list, ass_vis_list,
                            estat_partit=nou_estat_partit,
                            nom_arbitre=nom_arbitre_val,
                            nom_delegat_loc=nom_delegat_loc_val,
                            sig_delegat_loc=sig_loc_str,
                            nom_delegat_vis=nom_delegat_vis_val,
                            sig_delegat_vis=sig_vis_str,
                            sig_arbitre=sig_arb_str,
                            observacions=obs_input_val
                        )
                        if ok:
                            st.success(f"S'ha actualitzat el partit [{p_id}] amb l'estat '{nou_estat_partit}' correctament!")
                            st.rerun()
                        else:
                            st.error(msg)
            else:
                st.info("Tots els partits programats ja estan validats!")
    elif pwd_input != "":
        st.error("Contrasenya incorrecta. Revisa les teves credencials d'àrbitre.")
