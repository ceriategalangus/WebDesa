import os
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_border(cell, **kwargs):
    """
    Set cell`s border
    Usage:
    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "000000", "space": "0"},
        bottom={"sz": 12, "val": "single", "color": "000000", "space": "0"},
        start={"sz": 12, "val": "single", "color": "000000", "space": "0"},
        end={"sz": 12, "val": "single", "color": "000000", "space": "0"},
    )
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            # check for existing to replace
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))

def make_boxes_in_cell(cell, num_boxes, text_list=None):
    # This creates a nested table inside a cell to represent boxes
    if text_list is None:
        text_list = [""] * num_boxes
    table = cell.add_table(rows=1, cols=num_boxes)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    for i in range(num_boxes):
        c = table.cell(0, i)
        c.text = str(text_list[i]) if i < len(text_list) else ""
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Make the cell square-ish
        set_cell_border(c, 
            top={"sz": 4, "val": "single", "color": "000000"},
            bottom={"sz": 4, "val": "single", "color": "000000"},
            start={"sz": 4, "val": "single", "color": "000000"},
            end={"sz": 4, "val": "single", "color": "000000"}
        )
        # width
        c.width = Cm(0.5)

def create_form():
    doc = Document()
    
    # Custom Page Size: 21.7 cm x 15.1 cm
    section = doc.sections[0]
    section.page_width = Cm(21.7)
    section.page_height = Cm(15.1)
    
    # Margins
    section.top_margin = Cm(0.8)
    section.bottom_margin = Cm(0.8)
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)
    
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(9)
    
    # F-1.21 Box at top right
    p_f121 = doc.add_paragraph()
    p_f121.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p_f121.add_run("  F-1.21  ")
    run.font.bold = True
    run.font.size = Pt(11)
    # fake border with a 1x1 right-aligned table instead
    p_f121.clear()
    
    table_f121 = doc.add_table(rows=1, cols=1)
    table_f121.alignment = WD_TABLE_ALIGNMENT.RIGHT
    cell = table_f121.cell(0, 0)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("F-1.21")
    r.font.bold = True
    r.font.size = Pt(10)
    set_cell_border(cell, 
        top={"sz": 4, "val": "single", "color": "000000"},
        bottom={"sz": 4, "val": "single", "color": "000000"},
        start={"sz": 4, "val": "single", "color": "000000"},
        end={"sz": 4, "val": "single", "color": "000000"}
    )
    
    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("FORMULIR PERMOHONAN KARTU TANDA PENDUDUK ( KTP ) WARGA NEGARA INDONESIA")
    run_title.font.bold = True
    run_title.font.size = Pt(11)
    
    # Perhatian box
    table_perhatian = doc.add_table(rows=1, cols=1)
    table_perhatian.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_perhatian.autofit = False
    table_perhatian.columns[0].width = Cm(19.5)
    cell_perhatian = table_perhatian.cell(0, 0)
    
    p_perh = cell_perhatian.paragraphs[0]
    p_perh.add_run("Perhatian :\n").font.size = Pt(8)
    p_perh.add_run("1. Harap diisi dengan huruf cetak dan menggunakan tinta hitam\n").font.size = Pt(8)
    p_perh.add_run("2. Untuk kolom pilihan, harap memberi tanda silang ( X ) pada kotak pilihan.\n").font.size = Pt(8)
    p_perh.add_run("3. Setelah formulir ini diisi dan ditanda tangani, harap diserahkan kembali kekantor Desa/Kelurahan").font.size = Pt(8)
    set_cell_border(cell_perhatian, 
        top={"sz": 4, "val": "single", "color": "000000"},
        bottom={"sz": 4, "val": "single", "color": "000000"},
        start={"sz": 4, "val": "single", "color": "000000"},
        end={"sz": 4, "val": "single", "color": "000000"}
    )
    
    doc.add_paragraph() # spacer
    
    # Government Region Section
    table_gov = doc.add_table(rows=4, cols=4)
    table_gov.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_gov.autofit = False
    
    # Adjust column widths
    table_gov.columns[0].width = Cm(5.5)
    table_gov.columns[1].width = Cm(0.5)
    table_gov.columns[2].width = Cm(2.0)
    table_gov.columns[3].width = Cm(11.5)
    
    data_gov = [
        ("PEMERINTAH PROPINSI", "36", "BANTEN"),
        ("PEMERINTAH KABUPATEN/KOTA", "03", "TANGERANG"),
        ("KECAMATAN", "13", "TELUKNAGA"),
        ("KELURAHAN/DESA", "2011", "TEGALANGUS")
    ]
    
    for i, (label, code, name) in enumerate(data_gov):
        row = table_gov.rows[i]
        row.cells[0].text = label
        row.cells[0].paragraphs[0].runs[0].font.bold = True
        row.cells[1].text = ":"
        
        # Boxes for code
        make_boxes_in_cell(row.cells[2], len(code), list(code))
        
        # Underlined Name
        p_name = row.cells[3].paragraphs[0]
        run_name = p_name.add_run(" " + name + " ")
        run_name.underline = True
        run_name.font.bold = True
    
    doc.add_paragraph() # spacer
    
    # Permohonan KTP Row
    table_p = doc.add_table(rows=1, cols=4)
    table_p.alignment = WD_TABLE_ALIGNMENT.LEFT
    table_p.autofit = False
    table_p.columns[0].width = Cm(5.0)
    table_p.columns[1].width = Cm(3.5)
    table_p.columns[2].width = Cm(4.0)
    table_p.columns[3].width = Cm(4.0)
    
    row_p = table_p.rows[0]
    p_label = row_p.cells[0].paragraphs[0]
    r_label = p_label.add_run("PERMOHONAN KTP")
    r_label.font.bold = True
    r_label.underline = True
    r_label.font.italic = True
    
    def add_checkbox_item(cell, text):
        p = cell.paragraphs[0]
        # Adding a box with borders inside the cell is hard, let's just use nested table for the box
        inner = cell.add_table(rows=1, cols=2)
        inner.columns[0].width = Cm(0.6)
        inner.columns[1].width = Cm(2.5)
        box_cell = inner.cell(0, 0)
        set_cell_border(box_cell, 
            top={"sz": 4, "val": "single", "color": "000000"},
            bottom={"sz": 4, "val": "single", "color": "000000"},
            start={"sz": 4, "val": "single", "color": "000000"},
            end={"sz": 4, "val": "single", "color": "000000"}
        )
        text_cell = inner.cell(0, 1)
        text_cell.text = text
        text_cell.paragraphs[0].runs[0].font.size = Pt(8)
        
    add_checkbox_item(row_p.cells[1], " A. Baru")
    add_checkbox_item(row_p.cells[2], " B. Perpanjangan")
    add_checkbox_item(row_p.cells[3], " C. Penggantian")

    doc.add_paragraph() # spacer
    
    # Main Form Fields
    table_fields = doc.add_table(rows=4, cols=2)
    table_fields.alignment = WD_TABLE_ALIGNMENT.LEFT
    table_fields.autofit = False
    table_fields.columns[0].width = Cm(4.5)
    table_fields.columns[1].width = Cm(15.0)
    
    fields = [
        ("1. Nama Lengkap", 30),
        ("2. No. KK", 16),
        ("3. NIK", 16),
        ("4. Alamat", 30)
    ]
    
    for i, (label, box_count) in enumerate(fields):
        row = table_fields.rows[i]
        # Label cell
        cell_label = row.cells[0]
        cell_label.text = label
        set_cell_border(cell_label, 
            top={"sz": 4, "val": "single", "color": "000000"},
            bottom={"sz": 4, "val": "single", "color": "000000"},
            start={"sz": 4, "val": "single", "color": "000000"},
            end={"sz": 4, "val": "single", "color": "000000"}
        )
        cell_label.paragraphs[0].runs[0].font.size = Pt(8)
        
        # Boxes cell
        cell_boxes = row.cells[1]
        make_boxes_in_cell(cell_boxes, box_count)
        
        # Alamat has a second line for RT, RW, Kode Pos
        if i == 3:
            # Add second row of boxes inside the Alamat cell or as a new paragraph?
            # Easiest is to add a small table below the main boxes
            p = cell_boxes.add_paragraph()
            p.add_run() # spacer
            inner_tbl = cell_boxes.add_table(rows=1, cols=6)
            inner_tbl.columns[0].width = Cm(0.8) # RT
            inner_tbl.columns[1].width = Cm(2.0) # RT boxes
            inner_tbl.columns[2].width = Cm(0.8) # RW
            inner_tbl.columns[3].width = Cm(2.0) # RW boxes
            inner_tbl.columns[4].width = Cm(2.0) # Kode Pos
            inner_tbl.columns[5].width = Cm(3.0) # Kode Pos boxes
            
            # RT
            inner_tbl.cell(0, 0).text = "RT"
            inner_tbl.cell(0, 0).paragraphs[0].runs[0].font.size = Pt(8)
            set_cell_border(inner_tbl.cell(0, 0), top={"sz": 4, "val": "single"}, bottom={"sz": 4, "val": "single"}, start={"sz": 4, "val": "single"}, end={"sz": 4, "val": "single"})
            make_boxes_in_cell(inner_tbl.cell(0, 1), 3)
            
            # RW
            inner_tbl.cell(0, 2).text = "RW"
            inner_tbl.cell(0, 2).paragraphs[0].runs[0].font.size = Pt(8)
            set_cell_border(inner_tbl.cell(0, 2), top={"sz": 4, "val": "single"}, bottom={"sz": 4, "val": "single"}, start={"sz": 4, "val": "single"}, end={"sz": 4, "val": "single"})
            make_boxes_in_cell(inner_tbl.cell(0, 3), 3)
            
            # Kode Pos
            inner_tbl.cell(0, 4).text = "Kode Pos :"
            inner_tbl.cell(0, 4).paragraphs[0].runs[0].font.size = Pt(8)
            set_cell_border(inner_tbl.cell(0, 4), top={"sz": 4, "val": "single"}, bottom={"sz": 4, "val": "single"}, start={"sz": 4, "val": "single"}, end={"sz": 4, "val": "single"})
            make_boxes_in_cell(inner_tbl.cell(0, 5), 5)
            
    doc.add_paragraph()
    
    # Signature Section
    table_sig = doc.add_table(rows=2, cols=3)
    table_sig.alignment = WD_TABLE_ALIGNMENT.LEFT
    table_sig.autofit = False
    
    # Col 0: TTD box (left side)
    # Col 1: Empty space
    # Col 2: Signatures (right side)
    table_sig.columns[0].width = Cm(9.0)
    table_sig.columns[1].width = Cm(2.0)
    table_sig.columns[2].width = Cm(8.5)
    
    cell_ttd_box = table_sig.cell(0, 0)
    # merge with row 1
    cell_ttd_box.merge(table_sig.cell(1, 0))
    
    # Create the internal table for photo, thumb, signature
    inner_ttd = cell_ttd_box.add_table(rows=2, cols=3)
    inner_ttd.columns[0].width = Cm(2.5)
    inner_ttd.columns[1].width = Cm(2.5)
    inner_ttd.columns[2].width = Cm(4.0)
    
    # Headers
    h0 = inner_ttd.cell(0, 0)
    h0.text = "Pas Photo (2x3)"
    h0.paragraphs[0].runs[0].font.size = Pt(7)
    h0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_cell_border(h0, top={"sz": 4, "val": "single"}, bottom={"sz": 4, "val": "single"}, start={"sz": 4, "val": "single"}, end={"sz": 4, "val": "single"})
    
    h1 = inner_ttd.cell(0, 1)
    h1.text = "Cap Jempol"
    h1.paragraphs[0].runs[0].font.size = Pt(7)
    h1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_cell_border(h1, top={"sz": 4, "val": "single"}, bottom={"sz": 4, "val": "single"}, start={"sz": 4, "val": "single"}, end={"sz": 4, "val": "single"})
    
    h2 = inner_ttd.cell(0, 2)
    h2.text = "Spesimen Tanda Tangan"
    h2.paragraphs[0].runs[0].font.size = Pt(7)
    h2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_cell_border(h2, top={"sz": 4, "val": "single"}, bottom={"sz": 4, "val": "single"}, start={"sz": 4, "val": "single"}, end={"sz": 4, "val": "single"})
    
    # Bodies
    b0 = inner_ttd.cell(1, 0)
    p0 = b0.paragraphs[0]
    p0.add_run("\n\n\n\n") # space for photo
    set_cell_border(b0, top={"sz": 4, "val": "single"}, bottom={"sz": 4, "val": "single"}, start={"sz": 4, "val": "single"}, end={"sz": 4, "val": "single"})
    
    b1 = inner_ttd.cell(1, 1)
    b1.text = "\n\n\n\n"
    set_cell_border(b1, top={"sz": 4, "val": "single"}, bottom={"sz": 4, "val": "single"}, start={"sz": 4, "val": "single"}, end={"sz": 4, "val": "single"})
    
    b2 = inner_ttd.cell(1, 2)
    b2.text = "Atau ->\n\n\n\n\nKet: Cap Jempol/Tanda tangan"
    b2.paragraphs[0].runs[0].font.size = Pt(7)
    set_cell_border(b2, top={"sz": 4, "val": "single"}, bottom={"sz": 4, "val": "single"}, start={"sz": 4, "val": "single"}, end={"sz": 4, "val": "single"})
    
    
    # Right side signatures
    cell_pemohon = table_sig.cell(0, 2)
    p_pemohon_date = cell_pemohon.paragraphs[0]
    p_pemohon_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_date = p_pemohon_date.add_run("...........................................................")
    run_date.font.size = Pt(8)
    
    p_pemohon_title = cell_pemohon.add_paragraph("Pemohon,")
    p_pemohon_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_pemohon_title.runs[0].font.size = Pt(8)
    
    cell_pemohon.add_paragraph("\n\n")
    
    p_pemohon_name = cell_pemohon.add_paragraph("( .................................................... )")
    p_pemohon_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_pemohon_name.runs[0].font.size = Pt(8)
    
    cell_camat = table_sig.cell(1, 2)
    p_camat = cell_camat.paragraphs[0]
    p_camat.add_run("\n\nMengetahui,\n")
    
    # Layout for Camat & Kades
    inner_camat = cell_camat.add_table(rows=1, cols=2)
    inner_camat.columns[0].width = Cm(4.0)
    inner_camat.columns[1].width = Cm(4.5)
    
    c1 = inner_camat.cell(0, 0)
    c1.text = "\n\nCamat ............................"
    c1.paragraphs[0].runs[0].font.size = Pt(8)
    c1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    c2 = inner_camat.cell(0, 1)
    c2.text = "\n\nKepala Desa/Lurah.................."
    c2.paragraphs[0].runs[0].font.size = Pt(8)
    c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.save("Formulir_KTP_F121.docx")
    print("Document successfully created: Formulir_KTP_F121.docx")

if __name__ == "__main__":
    create_form()
