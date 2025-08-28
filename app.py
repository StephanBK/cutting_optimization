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
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
import textwrap

# Page config
st.set_page_config(
    page_title="Metal Cutting Optimizer",
    page_icon="✂️",
    layout="wide"
)

# Title
st.title("✂️ Metal Cutting Optimizer")
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


def parse_excel_file(uploaded_file):
    """Parse the Excel file with dynamic column detection"""
    try:
        # Read Excel file
        df = pd.read_excel(uploaded_file, sheet_name=0, header=None)

        # Find where actual data starts (look for "Finished Length in")
        data_start_row = None
        for idx, row in df.iterrows():
            if row[0] == "Finished Length in":
                data_start_row = idx
                break

        if data_start_row is None:
            st.error("Could not find 'Finished Length in' header in the file")
            return None, None

        # Extract project information
        project_info = {}
        for i in range(data_start_row):
            if pd.notna(df.iloc[i, 0]) and pd.notna(df.iloc[i, 1]):
                key = str(df.iloc[i, 0]).replace(":", "")
                value = df.iloc[i, 1]
                project_info[key] = value

        # Get headers from the data_start_row
        headers = df.iloc[data_start_row].values

        # Find S-columns dynamically
        s_columns = []
        s_column_indices = []
        for i, col in enumerate(headers):
            if pd.notna(col) and str(col).startswith('S') and len(str(col)) > 1:
                # Check if it's like S1, S2, etc.
                if str(col)[1:].replace('.0', '').isdigit():
                    s_columns.append(str(col))
                    s_column_indices.append(i)

        # Find important column indices
        try:
            length_col = np.where(headers == 'Finished Length in')[0][0]
            part_col = np.where(headers == 'Part #')[0][0]
            miter_col = np.where(headers == 'Miter')[0][0]

            # Total QTY might be "Total QTY" or "Total Qty" - case insensitive search
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

        # Extract cutting data
        cutting_data = []
        for i in range(data_start_row + 1, len(df)):
            if pd.notna(df.iloc[i, length_col]):
                row_data = {
                    'length': float(df.iloc[i, length_col]),
                    'part_number': df.iloc[i, part_col] if pd.notna(df.iloc[i, part_col]) else '',
                    'has_miter': str(df.iloc[i, miter_col]) == '**' if pd.notna(df.iloc[i, miter_col]) else False,
                    'total_qty': int(df.iloc[i, total_col]) if pd.notna(df.iloc[i, total_col]) else 0,
                    's_quantities': {}
                }

                # Get S-column quantities
                for s_col, s_idx in zip(s_columns, s_column_indices):
                    qty = df.iloc[i, s_idx]
                    row_data['s_quantities'][s_col] = int(qty) if pd.notna(qty) and qty != 0 else 0

                cutting_data.append(row_data)

        return project_info, cutting_data

    except Exception as e:
        st.error(f"Error parsing file: {str(e)}")
        return None, None


def validate_stock_availability(cutting_data, stock_inventory, miter_allowance):
    """Check if we have enough stock to cut all pieces"""

    # Calculate total length needed
    total_length_needed = 0
    pieces_too_long = []
    max_stock_length = max([item['length_inches'] for item in stock_inventory]) if stock_inventory else 0

    for item in cutting_data:
        actual_length = item['length']
        if item['has_miter']:
            actual_length += 2 * miter_allowance

        # Check if piece fits in any stock
        if actual_length > max_stock_length:
            pieces_too_long.append({
                'length': actual_length,
                'original': item['length'],
                'part': item['part_number'],
                'qty': item['total_qty']
            })

        total_length_needed += actual_length * item['total_qty']

    # Calculate total stock available
    total_stock_available = sum(item['length_inches'] * item['quantity'] for item in stock_inventory)

    return {
        'has_enough': total_stock_available >= total_length_needed,
        'total_needed': total_length_needed,
        'total_available': total_stock_available,
        'pieces_too_long': pieces_too_long,
        'shortage': max(0, total_length_needed - total_stock_available)
    }


def optimize_cutting(cutting_data, stock_inventory, miter_allowance):
    """Optimize cutting using First Fit Decreasing algorithm with stock quantity management"""

    # Create list of available stock pieces
    available_stock = []
    for stock_item in stock_inventory:
        for _ in range(stock_item['quantity']):
            available_stock.append(stock_item['length_inches'])

    # Sort available stock by length (use longer pieces first)
    available_stock.sort(reverse=True)

    # Prepare all pieces to cut
    pieces_to_cut = []
    for item in cutting_data:
        actual_length = item['length']
        if item['has_miter']:
            actual_length += 2 * miter_allowance

        # Create individual pieces with their S-column assignments
        for s_col, qty in item['s_quantities'].items():
            if qty > 0:
                for _ in range(qty):
                    pieces_to_cut.append({
                        'original_length': item['length'],
                        'actual_length': actual_length,
                        'part_number': item['part_number'],
                        'has_miter': item['has_miter'],
                        's_column': s_col
                    })

    # Sort pieces by length (largest first) for better optimization
    pieces_to_cut.sort(key=lambda x: x['actual_length'], reverse=True)

    # Initialize bins (stock pieces)
    bins = []
    used_stock_indices = []

    # First Fit Decreasing algorithm with stock quantity tracking
    uncut_pieces = []

    for piece in pieces_to_cut:
        placed = False

        # Try to fit in existing bins
        for bin_data in bins:
            remaining = bin_data['stock_length'] - sum(p['actual_length'] for p in bin_data['pieces'])
            if remaining >= piece['actual_length']:
                bin_data['pieces'].append(piece)
                placed = True
                break

        # If doesn't fit in any existing bin, try to use new stock
        if not placed:
            # Find the best available stock piece (smallest that fits)
            for i, stock_length in enumerate(available_stock):
                if i not in used_stock_indices and stock_length >= piece['actual_length']:
                    bins.append({
                        'stock_length': stock_length,
                        'pieces': [piece]
                    })
                    used_stock_indices.append(i)
                    placed = True
                    break

            if not placed:
                uncut_pieces.append(piece)

    # Calculate statistics
    total_stock_used = sum(b['stock_length'] for b in bins)
    total_material_needed = sum(p['actual_length'] for p in pieces_to_cut)
    total_waste = total_stock_used - sum(p['actual_length'] for b in bins for p in b['pieces'])
    total_waste_feet = total_waste / 12  # Convert to feet
    efficiency = ((total_material_needed - sum(
        p['actual_length'] for p in uncut_pieces)) / total_stock_used * 100) if total_stock_used > 0 else 0

    # Count stock usage by length
    stock_usage = {}
    for bin_data in bins:
        length = bin_data['stock_length']
        if length not in stock_usage:
            stock_usage[length] = 0
        stock_usage[length] += 1

    return {
        'bins': bins,
        'total_pieces': len(pieces_to_cut),
        'pieces_cut': len(pieces_to_cut) - len(uncut_pieces),
        'uncut_pieces': uncut_pieces,
        'total_stock_used': total_stock_used,
        'total_material_needed': total_material_needed,
        'total_waste': total_waste,
        'total_waste_feet': total_waste_feet,
        'efficiency': efficiency,
        'num_stock_pieces': len(bins),
        'stock_usage': stock_usage
    }


def create_cutting_diagram(bins, max_bins_to_show=10):
    """Create visual cutting diagram using Plotly"""

    # Limit the number of bins to show for performance
    bins_to_display = bins[:max_bins_to_show]

    fig = make_subplots(
        rows=len(bins_to_display),
        cols=1,
        subplot_titles=[f"Stock #{i + 1}: {bin_data['stock_length']:.3f}″ ({bin_data['stock_length'] / 12:.1f} ft)" for
                        i, bin_data in enumerate(bins_to_display)],
        vertical_spacing=0.05
    )

    colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD', '#98D8C8', '#6C5CE7']

    for bin_idx, bin_data in enumerate(bins_to_display):
        row = bin_idx + 1
        x_position = 0

        for piece_idx, piece in enumerate(bin_data['pieces']):
            # Create a rectangle for each piece
            fig.add_trace(
                go.Bar(
                    x=[piece['actual_length']],
                    y=[f"Stock {bin_idx + 1}"],
                    orientation='h',
                    name=f"{piece['part_number']} ({piece['s_column']})",
                    text=f"{piece['original_length']:.3f}″<br>{piece['s_column']}",
                    textposition='inside',
                    marker_color=colors[piece_idx % len(colors)],
                    showlegend=False,
                    hovertemplate=f"Part: {piece['part_number']}<br>Length: {piece['original_length']:.3f}″<br>Section: {piece['s_column']}<br>Miter: {'Yes' if piece['has_miter'] else 'No'}<extra></extra>"
                ),
                row=row, col=1
            )

        # Add waste section
        used_length = sum(p['actual_length'] for p in bin_data['pieces'])
        waste = bin_data['stock_length'] - used_length
        if waste > 0.1:  # Only show if waste is significant
            fig.add_trace(
                go.Bar(
                    x=[waste],
                    y=[f"Stock {bin_idx + 1}"],
                    orientation='h',
                    name='Waste',
                    text=f"Waste: {waste:.3f}″",
                    textposition='inside',
                    marker_color='lightgray',
                    showlegend=False,
                    hovertemplate=f"Waste: {waste:.3f}″<extra></extra>"
                ),
                row=row, col=1
            )

        # Update x-axis for this subplot
        fig.update_xaxes(
            title_text="Length (inches)" if bin_idx == len(bins_to_display) - 1 else "",
            range=[0, bin_data['stock_length'] * 1.05],
            row=row, col=1
        )
        fig.update_yaxes(visible=False, row=row, col=1)

    fig.update_layout(
        height=100 * len(bins_to_display) + 200,
        title_text="Cutting Layout Diagram",
        showlegend=False,
        barmode='stack'
    )

    return fig


def create_avery_5160_labels(bins, project_info):
    """Create PDF with Avery 5160 labels for cut pieces"""

    # Create PDF buffer
    buffer = io.BytesIO()

    # Avery 5160 specifications (in points - ReportLab uses points)
    page_width, page_height = letter  # 8.5" x 11" in points (612 x 792)
    label_width = 2.625 * inch  # 2-5/8"
    label_height = 1.0 * inch  # 1"
    labels_per_row = 3
    labels_per_col = 10
    labels_per_page = 30

    # Calculate margins - standard Avery margins
    margin_x = 0.15 * inch  # Left margin
    margin_y = 0.5 * inch  # Top margin

    # Create PDF
    c = canvas.Canvas(buffer, pagesize=letter)

    # Get project name
    project_name = project_info.get('Project Name', project_info.get('Project', 'Test Project'))
    if isinstance(project_name, (int, float)):
        project_name = str(project_name)

    # Collect all pieces to create labels for
    all_pieces = []
    for bin_data in bins:
        for piece in bin_data['pieces']:
            all_pieces.append(piece)

    # Create labels
    label_count = 0
    page_count = 1

    for piece in all_pieces:
        # Calculate position on current page
        row = (label_count % labels_per_page) // labels_per_row
        col = (label_count % labels_per_page) % labels_per_row

        # Calculate x, y coordinates (from bottom-left origin)
        x = margin_x + (col * label_width)
        y = page_height - margin_y - ((row + 1) * label_height)

        # Create label content with 3 decimal places
        length_text = f"{piece['original_length']:.3f}\""
        s_column_text = piece['s_column']

        # Font sizes
        project_font_size = 9
        length_font_size = 12
        s_column_font_size = 10

        # Handle project name wrapping
        project_lines = []
        max_chars = 24  # Characters that fit in label width

        if len(str(project_name)) > max_chars:
            words = str(project_name).split()
            current_line = ""
            for word in words:
                test_line = current_line + (" " if current_line else "") + word
                if len(test_line) <= max_chars:
                    current_line = test_line
                else:
                    if current_line:
                        project_lines.append(current_line)
                    current_line = word
            if current_line:
                project_lines.append(current_line)
            project_font_size = 8
        else:
            project_lines = [str(project_name)]

        # Draw border for debugging (uncomment to see label boundaries)
        # c.setStrokeColor(colors.red)
        # c.rect(x, y, label_width, label_height)

        # Calculate label center
        label_center_x = x + (label_width / 2)
        label_center_y = y + (label_height / 2)

        # Draw project name (top part of label)
        c.setFillColor(colors.black)
        if len(project_lines) == 1:
            # Single line
            c.setFont("Helvetica-Bold", project_font_size)
            text_width = c.stringWidth(project_lines[0], "Helvetica-Bold", project_font_size)
            text_x = label_center_x - (text_width / 2)
            text_y = label_center_y + (label_height * 0.25)
            c.drawString(text_x, text_y, project_lines[0])
        else:
            # Multiple lines
            line_height = project_font_size + 1
            total_height = len(project_lines) * line_height
            start_y = label_center_y + (label_height * 0.25) + (total_height / 2)

            c.setFont("Helvetica-Bold", project_font_size)
            for i, line in enumerate(project_lines):
                text_width = c.stringWidth(line, "Helvetica-Bold", project_font_size)
                text_x = label_center_x - (text_width / 2)
                text_y = start_y - (i * line_height)
                c.drawString(text_x, text_y, line)

        # Draw length (center of label) with 3 decimal places
        c.setFont("Helvetica-Bold", length_font_size)
        text_width = c.stringWidth(length_text, "Helvetica-Bold", length_font_size)
        text_x = label_center_x - (text_width / 2)
        text_y = label_center_y - 2
        c.drawString(text_x, text_y, length_text)

        # Draw S-column (bottom of label)
        c.setFont("Helvetica", s_column_font_size)
        text_width = c.stringWidth(s_column_text, "Helvetica", s_column_font_size)
        text_x = label_center_x - (text_width / 2)
        text_y = label_center_y - (label_height * 0.3)
        c.drawString(text_x, text_y, s_column_text)

        label_count += 1

        # Start new page if current page is full
        if label_count % labels_per_page == 0 and label_count < len(all_pieces):
            c.showPage()
            page_count += 1

    # Finalize PDF
    c.save()
    buffer.seek(0)
    return buffer, len(all_pieces), page_count


# Main app layout
col1, col2 = st.columns([1, 2])

with col1:
    st.header("📁 Input Settings")

    # File upload
    uploaded_file = st.file_uploader(
        "Upload Excel file (.xlsx)",
        type=['xlsx'],
        help="Upload the cutting list Excel file"
    )

    if uploaded_file:
        project_info, cutting_data = parse_excel_file(uploaded_file)

        if project_info and cutting_data:
            st.session_state.project_info = project_info
            st.session_state.cutting_data = cutting_data

            # Display project info
            st.success("✅ File loaded successfully!")

            with st.expander("📋 Project Information", expanded=True):
                for key, value in project_info.items():
                    st.text(f"{key}: {value}")

            with st.expander(f"📊 Cutting List ({len(cutting_data)} unique lengths)", expanded=False):
                total_pieces = sum(item['total_qty'] for item in cutting_data)
                st.text(f"Total pieces to cut: {total_pieces}")
                st.text("")
                for item in cutting_data[:5]:  # Show first 5 items
                    st.text(
                        f"• {item['length']:.3f}″ × {item['total_qty']} pcs {'(Miter)' if item['has_miter'] else ''}")
                if len(cutting_data) > 5:
                    st.text(f"... and {len(cutting_data) - 5} more lengths")

    st.divider()

    # Stock Settings
    st.header("📦 Stock Inventory")
    st.markdown("Enter available stock lengths in **feet** and quantities")

    # Stock input table
    stock_data = []

    col_length, col_qty = st.columns(2)
    with col_length:
        st.markdown("**Length (feet)**")
    with col_qty:
        st.markdown("**Quantity**")

    # Three input rows for stock
    for i in range(3):
        col_l, col_q = st.columns(2)
        with col_l:
            length = st.number_input(
                f"Length {i + 1}",
                min_value=0.0,
                value=20.0 if i == 0 else (12.0 if i == 1 else 10.0),
                step=1.0,
                key=f"length_{i}",
                label_visibility="collapsed"
            )
        with col_q:
            qty = st.number_input(
                f"Qty {i + 1}",
                min_value=0,
                value=20 if i == 0 else (10 if i == 1 else 0),
                step=1,
                key=f"qty_{i}",
                label_visibility="collapsed"
            )

        if length > 0 and qty > 0:
            stock_data.append({
                'length_feet': length,
                'length_inches': length * 12,
                'quantity': qty
            })

    # Update session state
    st.session_state.stock_inventory = stock_data

    # Display stock summary
    if stock_data:
        st.success(f"**Stock Summary:** {sum(item['quantity'] for item in stock_data)} pieces")
        for item in stock_data:
            st.text(f"• {item['length_feet']:.0f} ft × {item['quantity']} pcs")

    st.divider()

    # Miter allowance
    st.header("⚙️ Settings")
    miter_allowance = st.number_input(
        "Miter allowance per side (inches)",
        min_value=0.0,
        max_value=10.0,
        value=2.0,
        step=0.25,
        help="Extra length added to each side for mitered cuts"
    )

    st.divider()

    # Optimize button
    if st.button("🚀 Optimize Cutting", type="primary", use_container_width=True):
        if st.session_state.cutting_data and st.session_state.stock_inventory:
            # First validate stock availability
            validation = validate_stock_availability(
                st.session_state.cutting_data,
                st.session_state.stock_inventory,
                miter_allowance
            )

            if validation['pieces_too_long']:
                st.error("❌ Some pieces are too long for available stock!")
                for piece in validation['pieces_too_long']:
                    st.text(
                        f"• {piece['original']:.3f}″ (needs {piece['length']:.3f}″ with miter) × {piece['qty']} pcs")
                st.text(
                    f"Max stock length available: {max(item['length_inches'] for item in st.session_state.stock_inventory):.3f}″")
            elif not validation['has_enough']:
                st.warning(f"⚠️ Warning: May not have enough total stock")
                st.text(f"Total length needed: {validation['total_needed']:.3f}″")
                st.text(f"Total stock available: {validation['total_available']:.3f}″")
                st.text(f"Shortage: {validation['shortage']:.3f}″")
                st.info("Optimization will show which pieces can't be cut")

                # Still optimize to show what can be done
                with st.spinner("Optimizing..."):
                    results = optimize_cutting(
                        st.session_state.cutting_data,
                        st.session_state.stock_inventory,
                        miter_allowance
                    )
                    st.session_state.optimization_results = results
            else:
                st.success("✅ Stock validation passed!")
                with st.spinner("Optimizing..."):
                    results = optimize_cutting(
                        st.session_state.cutting_data,
                        st.session_state.stock_inventory,
                        miter_allowance
                    )
                    st.session_state.optimization_results = results
                    st.success("✅ Optimization complete!")
        elif not st.session_state.cutting_data:
            st.error("Please upload a file first")
        else:
            st.error("Please enter stock inventory")

# Right column - Results
with col2:
    st.header("📊 Optimization Results")

    if st.session_state.optimization_results:
        results = st.session_state.optimization_results

        # Check for uncut pieces
        if results['uncut_pieces']:
            st.error(f"⚠️ {len(results['uncut_pieces'])} pieces could not be cut from available stock!")
            with st.expander("Show uncut pieces"):
                for piece in results['uncut_pieces']:
                    st.text(f"• {piece['original_length']:.3f}″ - {piece['part_number']} ({piece['s_column']})")

        # Summary metrics
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            st.metric("Stock Used", results['num_stock_pieces'])
        with col_m2:
            st.metric("Pieces Cut", f"{results['pieces_cut']}/{results['total_pieces']}")
        with col_m3:
            st.metric("Efficiency", f"{results['efficiency']:.1f}%")
        with col_m4:
            st.metric("Total Waste", f"{results['total_waste_feet']:.3f} ft")

        # Stock Requirements Display
        st.divider()
        st.subheader("📦 Stock Requirements")

        # Create a nice display showing what stock is needed vs available
        for length_inches, count in results['stock_usage'].items():
            length_feet = length_inches / 12
            # Find original quantity for this length
            orig_qty = next((item['quantity'] for item in st.session_state.stock_inventory
                             if abs(item['length_inches'] - length_inches) < 0.01), 0)

            # Create columns for better display
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

        # Cutting diagram
        st.subheader("Cutting Layout")
        if len(results['bins']) > 10:
            st.info(f"Showing first 10 of {len(results['bins'])} stock pieces")

        if results['bins']:
            fig = create_cutting_diagram(results['bins'])
            st.plotly_chart(fig, use_container_width=True)

        # Detailed cut list
        with st.expander("📋 Detailed Cut List", expanded=False):
            for i, bin_data in enumerate(results['bins'][:20]):  # Show first 20
                length_feet = bin_data['stock_length'] / 12
                st.write(f"**Stock #{i + 1} ({length_feet:.1f} ft / {bin_data['stock_length']:.3f}″):**")
                for j, piece in enumerate(bin_data['pieces']):
                    st.text(
                        f"  Cut {j + 1}: {piece['original_length']:.3f}″ - {piece['part_number']} ({piece['s_column']}) {'[M]' if piece['has_miter'] else ''}")
                waste = bin_data['stock_length'] - sum(p['actual_length'] for p in bin_data['pieces'])
                st.text(f"  Waste: {waste:.3f}″")
                st.text("")

        # Export options
        st.divider()
        st.subheader("📥 Export Results")

        # Sticker generation
        st.write("**🏷️ Generate Avery 5160 Labels**")
        if st.button("🖨️ Generate Labels PDF", use_container_width=True):
            try:
                with st.spinner("Creating labels..."):
                    pdf_buffer, total_labels, page_count = create_avery_5160_labels(results['bins'],
                                                                                    st.session_state.project_info)

                st.success(f"✅ Created {total_labels} labels on {page_count} page(s)")
                st.download_button(
                    label=f"📄 Download Labels PDF ({total_labels} labels)",
                    data=pdf_buffer.getvalue(),
                    file_name=f"cutting_labels_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                    mime="application/pdf"
                )
            except ImportError:
                st.error(
                    "❌ PDF generation requires the reportlab library. Please install it with: pip install reportlab")
            except Exception as e:
                st.error(f"❌ Error generating labels: {str(e)}")

        st.divider()

        # Excel export
        if st.button("📊 Generate Excel Report", use_container_width=True):
            # Create Excel file in memory
            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet('Cutting List')

            # Write project info
            row = 0
            for key, value in st.session_state.project_info.items():
                worksheet.write(row, 0, key)
                worksheet.write(row, 1, str(value))
                row += 1

            row += 1
            worksheet.write(row, 0, "Generated:")
            worksheet.write(row, 1, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

            # Write optimization summary
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

            # Stock requirements section
            row += 2
            worksheet.write(row, 0, "STOCK REQUIREMENTS")
            row += 1
            worksheet.write(row, 0, "Stock Length (ft)")
            worksheet.write(row, 1, "Required Quantity")
            worksheet.write(row, 2, "Available Quantity")
            worksheet.write(row, 3, "Status")
            row += 1
            for length_inches, count in results['stock_usage'].items():
                length_feet = length_inches / 12
                orig_qty = next((item['quantity'] for item in st.session_state.stock_inventory
                                 if abs(item['length_inches'] - length_inches) < 0.01), 0)
                worksheet.write(row, 0, f"{length_feet:.0f} ft")
                worksheet.write(row, 1, count)
                worksheet.write(row, 2, orig_qty)
                worksheet.write(row, 3, "✅ OK" if count <= orig_qty else "❌ SHORTAGE")
                row += 1

            # Write detailed cut list
            row += 2
            worksheet.write(row, 0, "DETAILED CUT LIST")
            row += 1
            worksheet.write(row, 0, "Stock #")
            worksheet.write(row, 1, "Stock Length (ft)")
            worksheet.write(row, 2, "Stock Length (in)")
            worksheet.write(row, 3, "Cut #")
            worksheet.write(row, 4, "Part Number")
            worksheet.write(row, 5, "Original Length (in)")
            worksheet.write(row, 6, "Section")
            worksheet.write(row, 7, "Miter")
            worksheet.write(row, 8, "Waste (in)")

            row += 1
            for i, bin_data in enumerate(results['bins']):
                worksheet.write(row, 0, i + 1)
                worksheet.write(row, 1, f"{bin_data['stock_length'] / 12:.1f}")
                worksheet.write(row, 2, f"{bin_data['stock_length']:.3f}")

                for j, piece in enumerate(bin_data['pieces']):
                    if j > 0:
                        row += 1
                    worksheet.write(row, 3, j + 1)
                    worksheet.write(row, 4, piece['part_number'])
                    worksheet.write(row, 5, f"{piece['original_length']:.3f}")
                    worksheet.write(row, 6, piece['s_column'])
                    worksheet.write(row, 7, "Yes" if piece['has_miter'] else "No")

                waste = bin_data['stock_length'] - sum(p['actual_length'] for p in bin_data['pieces'])
                worksheet.write(row, 8, f"{waste:.3f}")
                row += 1

            # Write uncut pieces if any
            if results['uncut_pieces']:
                row += 2
                worksheet.write(row, 0, "UNCUT PIECES")
                row += 1
                worksheet.write(row, 0, "Part Number")
                worksheet.write(row, 1, "Original Length (in)")
                worksheet.write(row, 2, "Section")
                worksheet.write(row, 3, "Miter")
                row += 1
                for piece in results['uncut_pieces']:
                    worksheet.write(row, 0, piece['part_number'])
                    worksheet.write(row, 1, f"{piece['original_length']:.3f}")
                    worksheet.write(row, 2, piece['s_column'])
                    worksheet.write(row, 3, "Yes" if piece['has_miter'] else "No")
                    row += 1

            workbook.close()

            st.download_button(
                label="💾 Download Excel",
                data=output.getvalue(),
                file_name=f"cutting_list_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("👈 Upload a file and click 'Optimize Cutting' to see results")

# Footer
st.divider()
st.caption("Metal Cutting Optimizer v2.2 - Now with stock requirements and 3-decimal precision!")