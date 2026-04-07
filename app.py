import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
from datetime import datetime
import xlsxwriter
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.graphics import renderPDF
import textwrap
import hashlib
import xmlrpc.client
import base64
import os

# ── Odoo Config ──
ODOO_URL     = os.environ.get("ODOO_URL",     "https://inovues.odoo.com")
ODOO_DB      = os.environ.get("ODOO_DB",      "inovues")
ODOO_USER    = os.environ.get("ODOO_USER",    "sketterer@inovues.com")
ODOO_API_KEY = os.environ.get("ODOO_API_KEY", "")

@st.cache_data(ttl=300)
def get_odoo_projects():
    """Fetch active projects from Odoo"""
    try:
        common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
        uid = common.authenticate(ODOO_DB, ODOO_USER, ODOO_API_KEY, {})
        if not uid:
            return []
        models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")
        projects = models.execute_kw(
            ODOO_DB, uid, ODOO_API_KEY,
            "project.project", "search_read",
            [[["active", "=", True]]],
            {"fields": ["id", "name"], "order": "name asc"}
        )
        return projects
    except Exception as e:
        return []


# Page config
st.set_page_config(
    page_title="SWR Cutting Optimizer",
    page_icon="✂️",
    layout="wide"
)

# Title
st.title("✂️ SWR Cutting Optimizer")
st.markdown("Optimize metal cutting to minimize waste from stock lengths")

# Initialize session state
if 'optimization_results' not in st.session_state:
    st.session_state.optimization_results = None
if 'project_info' not in st.session_state:
    st.session_state.project_info = {}
if 'cutting_data' not in st.session_state:
    st.session_state.cutting_data = None
if 'stock_inventory' not in st.session_state:
    st.session_state.stock_inventory = []
if 'generated_files' not in st.session_state:
    st.session_state.generated_files = {}

# Cut loss constant
CUT_LOSS = 0.5  # 1/2 inch loss per cut

def parse_excel_file(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0, header=None)
        data_start_row = None
        for idx, row in df.iterrows():
            if row[0] == "Finished Length in":
                data_start_row = idx
                break
        if data_start_row is None:
            st.error("Could not find 'Finished Length in' header in the file")
            return None, None
        project_info = {}
        for i in range(data_start_row):
            if pd.notna(df.iloc[i, 0]) and pd.notna(df.iloc[i, 1]):
                key = str(df.iloc[i, 0]).replace(":", "")
                value = df.iloc[i, 1]
                project_info[key] = value
        headers = df.iloc[data_start_row].values
        s_columns = []
        s_column_indices = []
        for i, col in enumerate(headers):
            if pd.notna(col) and str(col).startswith('S') and len(str(col)) > 1:
                if str(col)[1:].replace('.0', '').isdigit():
                    s_columns.append(str(col))
                    s_column_indices.append(i)
        try:
            length_col = np.where(headers == 'Finished Length in')[0][0]
            part_col = np.where(headers == 'Part #')[0][0]
            total_col = None
            for i, col in enumerate(headers):
                if pd.notna(col) and 'total' in str(col).lower() and 'qty' in str(col).lower():
                    total_col = i
                    break
            if total_col is None:
                st.error("Could not find 'Total QTY' column")
                return None, None
        except IndexError as e:
            st.error(f"Could not find required columns: {e}")
            return None, None
        cutting_data = []
        for i in range(data_start_row + 1, len(df)):
            if pd.notna(df.iloc[i, length_col]):
                row_data = {
                    'length': float(df.iloc[i, length_col]),
                    'part_number': df.iloc[i, part_col] if pd.notna(df.iloc[i, part_col]) else '',
                    'total_qty': int(df.iloc[i, total_col]) if pd.notna(df.iloc[i, total_col]) else 0,
                    's_quantities': {}
                }
                for s_col, s_idx in zip(s_columns, s_column_indices):
                    qty = df.iloc[i, s_idx]
                    row_data['s_quantities'][s_col] = int(qty) if pd.notna(qty) and qty != 0 else 0
                cutting_data.append(row_data)
        return project_info, cutting_data
    except Exception as e:
        st.error(f"Error parsing file: {str(e)}")
        return None, None

def validate_stock_availability(cutting_data, stock_inventory):
    total_length_needed = 0
    pieces_too_long = []
    max_stock_length = max([item['length_inches'] for item in stock_inventory]) if stock_inventory else 0
    for item in cutting_data:
        actual_length = item['length'] + CUT_LOSS
        if actual_length > max_stock_length:
            pieces_too_long.append({'length': actual_length, 'original': item['length'], 'part': item['part_number'], 'qty': item['total_qty']})
        total_length_needed += actual_length * item['total_qty']
    total_stock_available = sum(item['length_inches'] * item['quantity'] for item in stock_inventory)
    return {
        'has_enough': total_stock_available >= total_length_needed,
        'total_needed': total_length_needed,
        'total_available': total_stock_available,
        'pieces_too_long': pieces_too_long,
        'shortage': max(0, total_length_needed - total_stock_available)
    }

def get_cutting_pattern_hash(pieces):
    sorted_pieces = sorted(pieces, key=lambda x: (x['length'], x['part_number'], x['s_column']))
    pattern_str = ""
    for piece in sorted_pieces:
        pattern_str += f"{piece['length']:.3f}-{piece['part_number']}-{piece['s_column']}"
    return hashlib.md5(pattern_str.encode()).hexdigest()

def optimize_cutting(cutting_data, stock_inventory):
    available_stock = []
    for stock_item in stock_inventory:
        for _ in range(stock_item['quantity']):
            available_stock.append(stock_item['length_inches'])
    available_stock.sort(reverse=True)
    pieces_to_cut = []
    for item in cutting_data:
        has_s_column_data = any(qty > 0 for qty in item['s_quantities'].values())
        if has_s_column_data:
            for s_col, qty in item['s_quantities'].items():
                if qty > 0:
                    for _ in range(qty):
                        pieces_to_cut.append({'length': item['length'], 'actual_length': item['length'] + CUT_LOSS, 'part_number': item['part_number'], 's_column': s_col})
        else:
            if item['total_qty'] > 0:
                for _ in range(item['total_qty']):
                    pieces_to_cut.append({'length': item['length'], 'actual_length': item['length'] + CUT_LOSS, 'part_number': item['part_number'], 's_column': 'ALL'})
    pieces_to_cut.sort(key=lambda x: x['actual_length'], reverse=True)
    bins = []
    used_stock_indices = []
    uncut_pieces = []
    for piece in pieces_to_cut:
        placed = False
        for bin_data in bins:
            remaining = bin_data['stock_length'] - sum(p['actual_length'] for p in bin_data['pieces'])
            if remaining >= piece['actual_length']:
                bin_data['pieces'].append(piece)
                placed = True
                break
        if not placed:
            for i, stock_length in enumerate(available_stock):
                if i not in used_stock_indices and stock_length >= piece['actual_length']:
                    bins.append({'stock_length': stock_length, 'pieces': [piece]})
                    used_stock_indices.append(i)
                    placed = True
                    break
        if not placed:
            uncut_pieces.append(piece)
    pattern_groups = {}
    for bin_data in bins:
        pattern_hash = get_cutting_pattern_hash(bin_data['pieces'])
        if pattern_hash not in pattern_groups:
            pattern_groups[pattern_hash] = {'stock_length': bin_data['stock_length'], 'pieces': bin_data['pieces'], 'count': 1}
        else:
            pattern_groups[pattern_hash]['count'] += 1
    sorted_pattern_groups = dict(sorted(pattern_groups.items(), key=lambda x: x[1]['count'], reverse=True))
    total_stock_used = sum(b['stock_length'] for b in bins)
    total_material_needed = sum(p['actual_length'] for p in pieces_to_cut)
    total_waste = total_stock_used - sum(p['actual_length'] for b in bins for p in b['pieces'])
    total_waste_feet = total_waste / 12
    efficiency = ((total_material_needed - sum(p['actual_length'] for p in uncut_pieces)) / total_stock_used * 100) if total_stock_used > 0 else 0
    stock_usage = {}
    for bin_data in bins:
        length = bin_data['stock_length']
        if length not in stock_usage:
            stock_usage[length] = 0
        stock_usage[length] += 1
    return {
        'bins': bins, 'pattern_groups': sorted_pattern_groups,
        'total_pieces': len(pieces_to_cut), 'pieces_cut': len(pieces_to_cut) - len(uncut_pieces),
        'uncut_pieces': uncut_pieces, 'total_stock_used': total_stock_used,
        'total_material_needed': total_material_needed, 'total_waste': total_waste,
        'total_waste_feet': total_waste_feet, 'efficiency': efficiency,
        'num_stock_pieces': len(bins), 'stock_usage': stock_usage
    }

def create_cutting_diagram(bins, max_bins_to_show=10):
    bins_to_display = bins[:max_bins_to_show]
    fig = make_subplots(
        rows=len(bins_to_display), cols=1,
        subplot_titles=[f"Stock #{i+1}: {bin_data['stock_length']:.3f}″ ({bin_data['stock_length']/12:.1f} ft)" for i, bin_data in enumerate(bins_to_display)],
        vertical_spacing=0.05
    )
    piece_colors = ['#FFD700','#32CD32','#1E90FF','#FF69B4','#FFA500','#9370DB','#00CED1','#FF6347','#90EE90','#FFB6C1']
    for bin_idx, bin_data in enumerate(bins_to_display):
        row = bin_idx + 1
        for piece_idx, piece in enumerate(bin_data['pieces']):
            piece_color = piece_colors[piece_idx % len(piece_colors)]
            fig.add_trace(go.Bar(x=[piece['length']], y=[f"Stock {bin_idx+1}"], orientation='h', name=f"{piece['part_number']} ({piece['s_column']})", text=f"{piece['length']:.3f}″\n{piece['s_column']}", textposition='inside', marker_color=piece_color, showlegend=False, hovertemplate=f"Part: {piece['part_number']}<br>Length: {piece['length']:.3f}″<br>Section: {piece['s_column']}<br>Cut Loss: {CUT_LOSS}″<extra></extra>"), row=row, col=1)
            fig.add_trace(go.Bar(x=[CUT_LOSS], y=[f"Stock {bin_idx+1}"], orientation='h', name="Cut Loss", text="CUT", textposition='inside', marker_color='#FFD700', showlegend=False, hovertemplate=f"Cut Loss: {CUT_LOSS}″<extra></extra>"), row=row, col=1)
        used_length = sum(p['actual_length'] for p in bin_data['pieces'])
        waste = bin_data['stock_length'] - used_length
        if waste > 0.1:
            fig.add_trace(go.Bar(x=[waste], y=[f"Stock {bin_idx+1}"], orientation='h', name='Waste', text=f"Waste: {waste:.3f}″", textposition='inside', marker_color='#D3D3D3', showlegend=False, hovertemplate=f"Waste: {waste:.3f}″<extra></extra>"), row=row, col=1)
        fig.update_xaxes(title_text="Length (inches)" if bin_idx == len(bins_to_display)-1 else "", range=[0, bin_data['stock_length']*1.05], row=row, col=1)
        fig.update_yaxes(visible=False, row=row, col=1)
    fig.update_layout(height=100*len(bins_to_display)+200, title_text="Cutting Layout Diagram (Gold = Cut Loss)", showlegend=False, barmode='stack')
    return fig

def create_pattern_diagram(pieces, stock_length, pattern_num, count):
    drawing = Drawing(500, 60)
    piece_colors = [colors.Color(1,0.8,0.8), colors.Color(0.8,0.8,1), colors.Color(0.8,1,0.8), colors.Color(1,1,0.8), colors.Color(1,0.8,1), colors.Color(0.8,1,1), colors.Color(1,0.9,0.7), colors.Color(0.9,0.8,1)]
    cut_color = colors.Color(1,0.84,0)
    waste_color = colors.Color(0.83,0.83,0.83)
    scale = 480 / stock_length
    x_pos = 10
    y_pos = 10
    height = 30
    for i, piece in enumerate(pieces):
        piece_color = piece_colors[i % len(piece_colors)]
        piece_width = piece['length'] * scale
        cut_width = CUT_LOSS * scale
        rect = Rect(x_pos, y_pos, piece_width, height)
        rect.fillColor = piece_color
        rect.strokeColor = colors.black
        rect.strokeWidth = 0.5
        drawing.add(rect)
        if piece_width > 30:
            text = String(x_pos + piece_width/2, y_pos + height/2 - 3, f"{piece['length']:.1f}″", textAnchor='middle')
            text.fontSize = 8
            text.fillColor = colors.black
            drawing.add(text)
        x_pos += piece_width
        cut_rect = Rect(x_pos, y_pos, cut_width, height)
        cut_rect.fillColor = cut_color
        cut_rect.strokeColor = colors.black
        cut_rect.strokeWidth = 0.5
        drawing.add(cut_rect)
        x_pos += cut_width
    used_length = sum(p['actual_length'] for p in pieces)
    waste = stock_length - used_length
    if waste > 0.1:
        waste_width = waste * scale
        waste_rect = Rect(x_pos, y_pos, waste_width, height)
        waste_rect.fillColor = waste_color
        waste_rect.strokeColor = colors.black
        waste_rect.strokeWidth = 0.5
        drawing.add(waste_rect)
        if waste_width > 30:
            waste_text = String(x_pos + waste_width/2, y_pos + height/2 - 3, f"Waste: {waste:.1f}″", textAnchor='middle')
            waste_text.fontSize = 7
            waste_text.fillColor = colors.black
            drawing.add(waste_text)
    return drawing

def create_cutting_patterns_pdf(pattern_groups, project_info):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, topMargin=0.5*inch, bottomMargin=0.5*inch)
    story = []
    styles = getSampleStyleSheet()
    title_style = styles['Title']
    title_style.alignment = TA_CENTER
    project_name = project_info.get('Project Name', project_info.get('Project', 'Cutting Patterns'))
    story.append(Paragraph(f"SWR Cutting Patterns - {project_name}", title_style))
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles['Normal']))
    story.append(Paragraph(f"Cut Loss: {CUT_LOSS}″ per cut", styles['Normal']))
    story.append(Spacer(1, 12))
    pattern_num = 1
    for pattern_hash, pattern_data in pattern_groups.items():
        stock_length_ft = pattern_data['stock_length'] / 12
        header_text = f"Pattern #{pattern_num} - {stock_length_ft:.1f}ft ({pattern_data['stock_length']:.3f}″) Stock - Used {pattern_data['count']} times"
        pattern_content = []
        pattern_content.append(Paragraph(header_text, styles['Heading2']))
        diagram = create_pattern_diagram(pattern_data['pieces'], pattern_data['stock_length'], pattern_num, pattern_data['count'])
        pattern_content.append(diagram)
        pattern_content.append(Spacer(1, 8))
        table_data = [['Cut #', 'Part Number', 'Length (in)', 'Section', 'Cut Loss', 'Running Total']]
        running_total = 0
        for i, piece in enumerate(pattern_data['pieces']):
            running_total += piece['actual_length']
            table_data.append([str(i+1), piece['part_number'], f"{piece['length']:.3f}″", piece['s_column'], f"{CUT_LOSS:.1f}″", f"{running_total:.3f}″"])
        waste = pattern_data['stock_length'] - running_total
        table_data.append(['', 'WASTE', f"{waste:.3f}″", '', '', f"{pattern_data['stock_length']:.3f}″"])
        table = Table(table_data, colWidths=[0.6*inch, 1.3*inch, 1*inch, 0.8*inch, 0.8*inch, 1.1*inch])
        bright_colors = [colors.Color(1,0.8,0.8), colors.Color(0.8,0.8,1), colors.Color(0.8,1,0.8), colors.Color(1,1,0.8), colors.Color(1,0.8,1), colors.Color(0.8,1,1), colors.Color(1,0.9,0.7), colors.Color(0.9,0.8,1)]
        table_style = [('BACKGROUND',(0,0),(-1,0),colors.Color(0.2,0.2,0.8)), ('TEXTCOLOR',(0,0),(-1,0),colors.white), ('ALIGN',(0,0),(-1,-1),'CENTER'), ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'), ('FONTSIZE',(0,0),(-1,0),9), ('BOTTOMPADDING',(0,0),(-1,0),8), ('BACKGROUND',(0,1),(-1,-2),colors.Color(0.98,0.98,0.95)), ('GRID',(0,0),(-1,-1),1,colors.black), ('FONTSIZE',(0,1),(-1,-1),8)]
        for i in range(1, len(table_data)-1):
            color_idx = (i-1) % len(bright_colors)
            table_style.append(('BACKGROUND',(0,i),(-1,i),bright_colors[color_idx]))
        table_style.append(('BACKGROUND',(0,-1),(-1,-1),colors.Color(0.9,0.9,0.9)))
        table_style.append(('FONTNAME',(0,-1),(2,-1),'Helvetica-Bold'))
        table.setStyle(TableStyle(table_style))
        pattern_content.append(table)
        material_used = sum(p['actual_length'] for p in pattern_data['pieces'])
        efficiency = (material_used / pattern_data['stock_length']) * 100
        efficiency_text = f"Efficiency: {efficiency:.1f}% | Waste: {waste:.3f}″ ({waste/12:.3f}ft) | Pieces: {len(pattern_data['pieces'])}"
        pattern_content.append(Spacer(1, 6))
        pattern_content.append(Paragraph(efficiency_text, styles['Normal']))
        pattern_content.append(Spacer(1, 12))
        story.append(KeepTogether(pattern_content))
        pattern_num += 1
    story.append(Paragraph("Summary", styles['Heading1']))
    total_patterns = len(pattern_groups)
    total_repetitions = sum(p['count'] for p in pattern_groups.values())
    summary_data = [['Unique Patterns', str(total_patterns)], ['Total Stock Pieces', str(total_repetitions)], ['Average Uses per Pattern', f"{total_repetitions/total_patterns:.1f}"], ['Cut Loss per Piece', f"{CUT_LOSS}″"]]
    summary_table = Table(summary_data, colWidths=[2*inch, 1*inch])
    summary_table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),colors.Color(0.7,0.9,1)), ('GRID',(0,0),(-1,-1),1,colors.black), ('FONTNAME',(0,0),(-1,-1),'Helvetica-Bold'), ('ALIGN',(0,0),(-1,-1),'CENTER'), ('FONTSIZE',(0,0),(-1,-1),10)]))
    story.append(summary_table)
    doc.build(story)
    buffer.seek(0)
    return buffer

def create_avery_5160_labels(bins, project_info):
    buffer = io.BytesIO()
    page_width, page_height = letter
    label_width = 2.625 * inch
    label_height = 1.0 * inch
    labels_per_row = 3
    labels_per_col = 10
    labels_per_page = 30
    margin_x = 0.15 * inch
    margin_y = 0.5 * inch
    c = canvas.Canvas(buffer, pagesize=letter)
    project_name = project_info.get('Project Name', project_info.get('Project', 'Test Project'))
    if isinstance(project_name, (int, float)):
        project_name = str(project_name)
    all_pieces = []
    for bin_data in bins:
        for piece in bin_data['pieces']:
            all_pieces.append(piece)
    label_count = 0
    page_count = 1
    for piece in all_pieces:
        row = (label_count % labels_per_page) // labels_per_row
        col = (label_count % labels_per_page) % labels_per_row
        x = margin_x + (col * label_width)
        y = page_height - margin_y - ((row + 1) * label_height)
        length_text = f"{piece['length']:.3f}\""
        s_column_text = piece['s_column']
        cut_loss_text = f"Cut: +{CUT_LOSS}\""
        project_font_size = 9
        length_font_size = 12
        s_column_font_size = 10
        cut_loss_font_size = 8
        label_center_x = x + (label_width / 2)
        label_center_y = y + (label_height / 2)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", project_font_size)
        text_width = c.stringWidth(str(project_name), "Helvetica-Bold", project_font_size)
        text_x = label_center_x - (text_width / 2)
        text_y = label_center_y + (label_height * 0.3)
        c.drawString(text_x, text_y, str(project_name)[:24])
        c.setFont("Helvetica-Bold", length_font_size)
        text_width = c.stringWidth(length_text, "Helvetica-Bold", length_font_size)
        text_x = label_center_x - (text_width / 2)
        text_y = label_center_y + (label_height * 0.05)
        c.drawString(text_x, text_y, length_text)
        c.setFont("Helvetica", s_column_font_size)
        text_width = c.stringWidth(s_column_text, "Helvetica", s_column_font_size)
        text_x = label_center_x - (text_width / 2)
        text_y = label_center_y + (label_height * -0.15)
        c.drawString(text_x, text_y, s_column_text)
        c.setFont("Helvetica", cut_loss_font_size)
        text_width = c.stringWidth(cut_loss_text, "Helvetica", cut_loss_font_size)
        text_x = label_center_x - (text_width / 2)
        text_y = label_center_y + (label_height * -0.35)
        c.drawString(text_x, text_y, cut_loss_text)
        label_count += 1
        if label_count % labels_per_page == 0 and label_count < len(all_pieces):
            c.showPage()
            page_count += 1
    c.save()
    buffer.seek(0)
    return buffer, len(all_pieces), page_count

def generate_excel(results, project_info, stock_inventory):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('Cutting List')
    row = 0
    for key, value in project_info.items():
        worksheet.write(row, 0, key)
        worksheet.write(row, 1, str(value))
        row += 1
    row += 1
    worksheet.write(row, 0, "Generated:")
    worksheet.write(row, 1, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    row += 1
    worksheet.write(row, 0, "Cut Loss per Piece:")
    worksheet.write(row, 1, f"{CUT_LOSS} inches")
    row += 2
    worksheet.write(row, 0, "OPTIMIZATION SUMMARY")
    row += 1
    worksheet.write(row, 0, "Total Stock Pieces Used:")
    worksheet.write(row, 1, results['num_stock_pieces'])
    row += 1
    worksheet.write(row, 0, "Pieces Cut:")
    worksheet.write(row, 1, f"{results['pieces_cut']} / {results['total_pieces']}")
    row += 1
    worksheet.write(row, 0, "Efficiency:")
    worksheet.write(row, 1, f"{results['efficiency']:.1f}%")
    row += 1
    worksheet.write(row, 0, "Total Waste:")
    worksheet.write(row, 1, f"{results['total_waste']:.3f} inches ({results['total_waste_feet']:.3f} feet)")
    row += 2
    worksheet.write(row, 0, "DETAILED CUT LIST")
    row += 1
    worksheet.write(row, 0, "Stock #")
    worksheet.write(row, 1, "Stock Length (ft)")
    worksheet.write(row, 2, "Cut #")
    worksheet.write(row, 3, "Part Number")
    worksheet.write(row, 4, "Original Length (in)")
    worksheet.write(row, 5, "Cut Loss (in)")
    worksheet.write(row, 6, "Total Length (in)")
    worksheet.write(row, 7, "Section")
    worksheet.write(row, 8, "Waste (in)")
    row += 1
    for i, bin_data in enumerate(results['bins']):
        worksheet.write(row, 0, i+1)
        worksheet.write(row, 1, f"{bin_data['stock_length']/12:.1f}")
        for j, piece in enumerate(bin_data['pieces']):
            if j > 0:
                row += 1
            worksheet.write(row, 2, j+1)
            worksheet.write(row, 3, piece['part_number'])
            worksheet.write(row, 4, f"{piece['length']:.3f}")
            worksheet.write(row, 5, f"{CUT_LOSS:.1f}")
            worksheet.write(row, 6, f"{piece['actual_length']:.3f}")
            worksheet.write(row, 7, piece['s_column'])
            waste = bin_data['stock_length'] - sum(p['actual_length'] for p in bin_data['pieces'])
            worksheet.write(row, 8, f"{waste:.3f}")
        row += 1
    workbook.close()
    output.seek(0)
    return output

# ── Main Layout ──
col1, col2 = st.columns([1, 2])

with col1:
    st.header("📁 Input Settings")

    # ── Helper: fetch AggCutOnly attachments for a project's tasks ──
    @st.cache_data(ttl=60)
    def get_aggcutonly_files(project_id):
        """Return list of AggCutOnly attachments across all tasks in the project."""
        try:
            common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
            uid = common.authenticate(ODOO_DB, ODOO_USER, ODOO_API_KEY, {})
            if not uid:
                return []
            models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")
            # Get all task IDs for this project
            task_ids = models.execute_kw(
                ODOO_DB, uid, ODOO_API_KEY,
                "project.task", "search",
                [[["project_id", "=", project_id]]]
            )
            if not task_ids:
                return []
            # Find attachments on those tasks whose name contains AggCutOnly
            attachments = models.execute_kw(
                ODOO_DB, uid, ODOO_API_KEY,
                "ir.attachment", "search_read",
                [[
                    ["res_model", "=", "project.task"],
                    ["res_id", "in", task_ids],
                    ["name", "ilike", "AggCutOnly"]
                ]],
                {"fields": ["id", "name", "create_date", "datas"], "order": "create_date desc"}
            )
            return attachments
        except Exception:
            return []

    def load_file_from_attachment(attachment):
        """Decode base64 datas from Odoo attachment into a BytesIO object."""
        raw = base64.b64decode(attachment["datas"])
        return io.BytesIO(raw)

    # ── DEFAULT: Odoo Project Picker ──
    st.subheader("Load from Odoo")
    projects = get_odoo_projects()
    project_map = {p["name"]: p["id"] for p in projects} if projects else {}

    if not project_map:
        st.warning("⚠️ Could not load Odoo projects. Use manual upload below.")
    else:
        selected_project_name = st.selectbox(
            "Select Project",
            options=[""] + list(project_map.keys()),
            index=0,
            format_func=lambda x: "— choose a project —" if x == "" else x,
            key="odoo_project_select"
        )

        if selected_project_name:
            project_id = project_map[selected_project_name]
            attachments = get_aggcutonly_files(project_id)

            if not attachments:
                st.warning("No AggCutOnly files found for this project.")
            else:
                # Build display labels — mark newest with ⭐
                def fmt_attachment(i, att):
                    label = att["name"]
                    date_str = att.get("create_date", "")[:10]
                    suffix = f" ({date_str})"
                    return ("⭐ " if i == 0 else "") + label + suffix

                options = [fmt_attachment(i, a) for i, a in enumerate(attachments)]
                selected_label = st.selectbox(
                    "Select AggCutOnly file (⭐ = newest)",
                    options=options,
                    index=0,
                    key="odoo_file_select"
                )
                selected_idx = options.index(selected_label)
                selected_attachment = attachments[selected_idx]

                if st.button("📥 Load File from Odoo", use_container_width=True):
                    try:
                        file_obj = load_file_from_attachment(selected_attachment)
                        project_info, cutting_data = parse_excel_file(file_obj)
                        if project_info and cutting_data:
                            st.session_state.project_info = project_info
                            st.session_state.cutting_data = cutting_data
                            st.success(f"✅ Loaded: {selected_attachment['name']}")
                    except Exception as e:
                        st.error(f"❌ Failed to load file: {e}")

    # ── Show parsed file info if loaded (works for both Odoo and manual) ──
    if st.session_state.cutting_data:
        project_info = st.session_state.project_info
        cutting_data = st.session_state.cutting_data
        with st.expander("📋 Project Information", expanded=True):
            for key, value in project_info.items():
                st.text(f"{key}: {value}")
        with st.expander(f"📊 Cutting List ({len(cutting_data)} unique lengths)", expanded=False):
            total_pieces = sum(item['total_qty'] for item in cutting_data)
            st.text(f"Total pieces to cut: {total_pieces}")
            st.text(f"Cut loss per piece: {CUT_LOSS}″")
            for item in cutting_data[:5]:
                st.text(f"• {item['length']:.3f}″ (+{CUT_LOSS}″ cut) × {item['total_qty']} pcs")
            if len(cutting_data) > 5:
                st.text(f"... and {len(cutting_data)-5} more lengths")

    # ── FALLBACK: Manual Upload ──
    with st.expander("📁 Manual Upload", expanded=False):
        uploaded_file = st.file_uploader("Upload Aggcutonly file (.xlsx)", type=['xlsx'], help="Upload the Aggcutonly cutting list Excel file")
        if uploaded_file:
            project_info, cutting_data = parse_excel_file(uploaded_file)
            if project_info and cutting_data:
                st.session_state.project_info = project_info
                st.session_state.cutting_data = cutting_data
                st.success("✅ File loaded successfully!")

    st.divider()
    st.header("📦 Stock Inventory")
    st.markdown("Enter available stock lengths in **feet** and quantities")
    stock_data = []
    col_length, col_qty = st.columns(2)
    with col_length:
        st.markdown("**Length (feet)**")
    with col_qty:
        st.markdown("**Quantity**")
    for i in range(3):
        col_l, col_q = st.columns(2)
        with col_l:
            length = st.number_input(f"Length {i+1}", min_value=0.0, value=20.0 if i==0 else (12.0 if i==1 else 10.0), step=1.0, key=f"length_{i}", label_visibility="collapsed")
        with col_q:
            qty = st.number_input(f"Qty {i+1}", min_value=0, value=20 if i==0 else (10 if i==1 else 0), step=1, key=f"qty_{i}", label_visibility="collapsed")
        if length > 0 and qty > 0:
            stock_data.append({'length_feet': length, 'length_inches': length*12, 'quantity': qty})
    st.session_state.stock_inventory = stock_data
    if stock_data:
        st.success(f"**Stock Summary:** {sum(item['quantity'] for item in stock_data)} pieces")
        for item in stock_data:
            st.text(f"• {item['length_feet']:.0f} ft × {item['quantity']} pcs")

    st.divider()
    st.info(f"ℹ️ **Cut Loss:** {CUT_LOSS}″ per cut is automatically added to each piece")
    st.divider()

    if st.button("🚀 Optimize Cutting", type="primary", use_container_width=True):
        if st.session_state.cutting_data and st.session_state.stock_inventory:
            validation = validate_stock_availability(st.session_state.cutting_data, st.session_state.stock_inventory)
            if validation['pieces_too_long']:
                st.error("❌ Some pieces are too long for available stock!")
                for piece in validation['pieces_too_long']:
                    st.text(f"• {piece['original']:.3f}″ (+{CUT_LOSS}″ cut = {piece['length']:.3f}″) × {piece['qty']} pcs")
            else:
                with st.spinner("Optimizing..."):
                    results = optimize_cutting(st.session_state.cutting_data, st.session_state.stock_inventory)
                    st.session_state.optimization_results = results
                    st.session_state.generated_files = {}  # reset generated files
                st.success("✅ Optimization complete!")
        elif not st.session_state.cutting_data:
            st.error("Please upload a file first")
        else:
            st.error("Please enter stock inventory")

with col2:
    st.header("📊 Optimization Results")
    if st.session_state.optimization_results:
        results = st.session_state.optimization_results
        if results['uncut_pieces']:
            st.error(f"⚠️ {len(results['uncut_pieces'])} pieces could not be cut from available stock!")
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            st.metric("Stock Used", results['num_stock_pieces'])
        with col_m2:
            st.metric("Pieces Cut", f"{results['pieces_cut']}/{results['total_pieces']}")
        with col_m3:
            st.metric("Efficiency", f"{results['efficiency']:.1f}%")
        with col_m4:
            st.metric("Total Waste", f"{results['total_waste_feet']:.3f} ft")

        st.divider()
        st.subheader("🔄 Cutting Patterns (Sorted by Frequency)")
        pattern_summary_data = []
        for i, (pattern_hash, pattern_data) in enumerate(results['pattern_groups'].items()):
            stock_length_ft = pattern_data['stock_length'] / 12
            used_length = sum(p['actual_length'] for p in pattern_data['pieces'])
            waste = pattern_data['stock_length'] - used_length
            efficiency = (used_length / pattern_data['stock_length']) * 100
            pattern_summary_data.append({'Pattern': f"#{i+1}", 'Stock Length': f"{stock_length_ft:.1f} ft", 'Pieces': len(pattern_data['pieces']), 'Used': f"{pattern_data['count']} times", 'Efficiency': f"{efficiency:.1f}%", 'Waste': f"{waste:.3f}″"})
        if pattern_summary_data:
            df_patterns = pd.DataFrame(pattern_summary_data)
            st.dataframe(df_patterns, use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("📦 Stock Requirements")
        for length_inches, count in results['stock_usage'].items():
            length_feet = length_inches / 12
            orig_qty = next((item['quantity'] for item in st.session_state.stock_inventory if abs(item['length_inches'] - length_inches) < 0.01), 0)
            col_stock, col_used, col_available = st.columns([2, 1, 1])
            with col_stock:
                st.text(f"{length_feet:.0f} ft stock:")
            with col_used:
                st.text(f"Need: {count}")
            with col_available:
                if count <= orig_qty:
                    st.success(f"Have: {orig_qty} ✅")
                else:
                    st.error(f"Have: {orig_qty} ❌")

        st.divider()
        st.subheader("✂️ Cutting Layout")
        if len(results['bins']) > 10:
            st.info(f"Showing first 10 of {len(results['bins'])} stock pieces")
        if results['bins']:
            fig = create_cutting_diagram(results['bins'])
            st.plotly_chart(fig, use_container_width=True)

        st.divider()
        st.subheader("📥 Export Results")
        export_col1, export_col2, export_col3 = st.columns(3)

        with export_col1:
            st.write("**🏷️ Labels**")
            if st.button("🖨️ Generate Labels", use_container_width=True):
                with st.spinner("Creating labels..."):
                    pdf_buffer, total_labels, page_count = create_avery_5160_labels(results['bins'], st.session_state.project_info)
                    st.session_state.generated_files['labels'] = {'data': pdf_buffer.getvalue(), 'filename': f"cutting_labels_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", 'mime': 'application/pdf'}
                st.success(f"✅ {total_labels} labels on {page_count} page(s)")
            if 'labels' in st.session_state.generated_files:
                f = st.session_state.generated_files['labels']
                st.download_button(label="📄 Download Labels", data=f['data'], file_name=f['filename'], mime=f['mime'], use_container_width=True)

        with export_col2:
            st.write("**📊 Excel Report**")
            if st.button("📊 Generate Excel", use_container_width=True):
                with st.spinner("Creating Excel..."):
                    output = generate_excel(results, st.session_state.project_info, st.session_state.stock_inventory)
                    st.session_state.generated_files['excel'] = {'data': output.getvalue(), 'filename': f"cutting_list_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", 'mime': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}
                st.success("✅ Excel ready")
            if 'excel' in st.session_state.generated_files:
                f = st.session_state.generated_files['excel']
                st.download_button(label="💾 Download Excel", data=f['data'], file_name=f['filename'], mime=f['mime'], use_container_width=True)

        with export_col3:
            st.write("**📋 Cut Patterns**")
            if st.button("📋 Generate Patterns", use_container_width=True):
                with st.spinner("Creating PDF..."):
                    pdf_buffer = create_cutting_patterns_pdf(results['pattern_groups'], st.session_state.project_info)
                    st.session_state.generated_files['patterns'] = {'data': pdf_buffer.getvalue(), 'filename': f"cutting_patterns_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", 'mime': 'application/pdf'}
                st.success(f"✅ {len(results['pattern_groups'])} patterns ready")
            if 'patterns' in st.session_state.generated_files:
                f = st.session_state.generated_files['patterns']
                st.download_button(label="📄 Download Patterns", data=f['data'], file_name=f['filename'], mime=f['mime'], use_container_width=True)

        # ── Odoo Integration ──
        st.divider()
        st.subheader("💾 Save to Odoo Project")

        projects = get_odoo_projects()
        if not projects:
            st.error("Could not connect to Odoo or no active projects found.")
        else:
            project_options = {p['name']: p['id'] for p in projects}
            selected_project_name = st.selectbox(
                "Select Odoo Project to attach files to:",
                options=list(project_options.keys()),
                index=None,
                placeholder="Choose a project..."
            )

            files_ready = list(st.session_state.generated_files.keys())
            if not files_ready:
                st.warning("Generate at least one file above before saving to Odoo.")
            else:
                st.info(f"Files ready to upload: {', '.join(files_ready)}")

                if st.button("📎 Attach all files to selected project", type="primary",
                             use_container_width=True, disabled=selected_project_name is None):
                    project_id = project_options[selected_project_name]
                    with st.spinner(f"Attaching files to '{selected_project_name}'..."):
                        try:
                            common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
                            uid    = common.authenticate(ODOO_DB, ODOO_USER, ODOO_API_KEY, {})
                            if not uid:
                                st.error("❌ Odoo authentication failed — check ODOO_API_KEY env var.")
                            else:
                                models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")

                                def odoo_call(model, method, args, kwargs={}):
                                    return models.execute_kw(ODOO_DB, uid, ODOO_API_KEY, model, method, args, kwargs)

                                # Build task name from project info
                                proj_name = st.session_state.project_info.get(
                                    'Project Name',
                                    st.session_state.project_info.get('Project', '')
                                )
                                task_date = datetime.now().strftime('%Y-%m-%d')
                                task_title = f"Cutting Optimization — {proj_name} — {task_date}" if proj_name else f"Cutting Optimization — {task_date}"

                                # Create task in Engineering stage, set to Approved
                                task_id = odoo_call("project.task", "create", [{
                                    "name":       task_title,
                                    "project_id": project_id,
                                    "stage_id":   8,            # Engineering
                                    "state":      "03_approved",
                                }])

                                # Attach all generated files to the task
                                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                                for file_key, file_data in st.session_state.generated_files.items():
                                    odoo_call("ir.attachment", "create", [{
                                        "name":      file_data['filename'],
                                        "type":      "binary",
                                        "datas":     base64.b64encode(file_data['data']).decode("utf-8"),
                                        "res_model": "project.task",
                                        "res_id":    task_id,
                                        "mimetype":  file_data['mime'],
                                    }])

                                # Build chatter message with optimization summary
                                results = st.session_state.optimization_results
                                odoo_call("project.task", "message_post", [[task_id]], {
                                    "body": (
                                        f"<b>✂️ Cutting Optimization attached</b><br/>"
                                        f"Project: {proj_name}<br/>"
                                        f"Stock Pieces Used: {results['num_stock_pieces']}<br/>"
                                        f"Pieces Cut: {results['pieces_cut']}/{results['total_pieces']}<br/>"
                                        f"Efficiency: {results['efficiency']:.1f}%<br/>"
                                        f"Total Waste: {results['total_waste_feet']:.3f} ft<br/>"
                                        f"Files: {', '.join(files_ready)}<br/>"
                                        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
                                    ),
                                    "message_type": "comment",
                                    "subtype_xmlid": "mail.mt_comment",
                                })

                                st.success(
                                    f"✅ {len(files_ready)} file(s) attached to task **{task_title}** "
                                    f"in **{selected_project_name}** (Engineering → Approved)!"
                                )

                        except xmlrpc.client.Fault as e:
                            st.error(f"❌ Odoo API error: {e.faultString}")
                        except Exception as e:
                            st.error(f"❌ Error: {str(e)}")

    else:
        st.info("👈 Upload a file and click 'Optimize Cutting' to see results")

st.divider()
st.caption(f"SWR Cutting Optimizer v5.0 - With Odoo Integration")
