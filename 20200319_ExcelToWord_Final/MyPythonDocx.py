#!/usr/bin/env python
# coding: utf-8

# In[1]:

from docx import Document
from docx.shared import RGBColor
from docx.shared import Pt
from docx.shared import Cm
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH # 處理字串對齊
from docx.enum.text import WD_LINE_SPACING
from docx.enum.text import WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT # 處理表格對齊
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.style import WD_STYLE_TYPE

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement

from docx.text.paragraph import Paragraph

GRAY = RGBColor(204,204,204)
BLACK = RGBColor(0, 0, 0)


# In[2]:


class Table():
    
    def set_cell_color(cell, rgbColor):
        shading_elm_1 = parse_xml(r'<w:shd {} w:fill="{color_value}"/>'.format(nsdecls('w'), color_value=rgbColor))
        cell._tc.get_or_add_tcPr().append(shading_elm_1)
    
    def set_content_font(table, size=Pt(14), bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, name=u'標楷體'):
        for row in table.rows:
            Table.set_row_height(row)
            for cell in row.cells: 
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                Parag.set_font(cell.paragraphs, size, bold, align, name)
    
    # https://stackoverflow.com/questions/43051462/python-docx-how-to-set-cell-width-in-tables/43053996
    def col_widths(table, *widths): 
        for row in table.rows:
            for idx, width in enumerate(widths):
                row.cells[idx].width = Inches(width)

                
    def row_height(table, height):
        for i, row in enumerate(table.rows):
            row.height = height
            
    # https://stackoverflow.com/questions/37532283/python-how-to-adjust-row-height-of-table-in-docx
    def set_row_height(row):
        tr = row._tr
        trPr = tr.get_or_add_trPr()
        trHeight = OxmlElement('w:trHeight')
        trHeight.set(qn('w:val'), "100")
        trHeight.set(qn('w:hRule'), "auto")
        trPr.append(trHeight)

    # https://stackoverflow.com/questions/55545494/in-python-docx-how-do-i-delete-a-table-row
    def remove_row(table, row):
        tbl = table._tbl
        tr = row._tr
        tbl.remove(tr)
        
    # https://github.com/python-openxml/python-docx/issues/156
    def move_table_after(table, paragraph):
        tbl, p = table._tbl, paragraph._p
        p.addnext(tbl)
        
    def add(table, data):
        for row in data:
            row_cell = table.add_row().cells
            for i, cell in enumerate(row):
                row_cell[i].text = str(cell)


# In[3]:


class Parag():
    
    def set_font(parags, size=Pt(14), bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, name=u'標楷體'):

        if isinstance(parags, list):
            for paragraph in parags: 
                paragraph.paragraph_format.alignment = align
                paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                for run in paragraph.runs:
                    run.bold = bold 
                    run.font.size = size                        
                    run.font.name = 'Arial' # 設置英文字體
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), name) # 設置中文字體

        else:
            parags.paragraph_format.alignment = align
            parags.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            for run in parags.runs:
                run.bold = bold 
                run.font.size = size                        
                run.font.name = 'Arial'
                run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
                
    def delete(paragraph):
        p = paragraph._element
        p.getparent().remove(p)
        p._p = p._element = None
        
    # Insert a new paragraph after the given paragraph.
    def insert_paragraph_after(paragraph, text=None, style=None):
        new_p = OxmlElement("w:p")
        paragraph._p.addnext(new_p)
        new_para = Paragraph(new_p, paragraph._parent)
        if text:
            new_para.add_run(text)
        if style is not None:
            new_para.style = style
        return new_para
    
    # https://github.com/python-openxml/python-docx/issues/217
    def create_list(paragraph, list_type):
        p = paragraph._p # access to xml paragraph element
        pPr = p.get_or_add_pPr() # access paragraph properties
        numPr = OxmlElement('w:numPr') # create number properties element
        numId = OxmlElement('w:numId') # create numId element - sets bullet type
        numId.set(qn('w:val'), list_type) # set list type/indentation
        numPr.append(numId) # add bullet type to number properties list
        pPr.append(numPr) # add number properties to paragraph

