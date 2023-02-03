# First set directory where you want to save the file

txt = '                                                                                Porcentaje de tareas ejecutadas de una Plataforma para atención al publico para el "Software de Procesamiento de señales e Imágenes empleando Técnicas de Inteligencia Artificial Versión 1"  con cumplimiento de meta física de acuerdo a lo programado en el formato de seguimiento y  control de acciones para el cumplimiento de la actividad operativa  2022                                                                          '
# x = txt.split('\r\n')
x = txt.split()
print('x: ', x)


# import os
# os.chdir(r"C:\Users\user\Desktop\Dann Uc")
#
# #Now import required packages
#
# # import docx
# from docx import Document
# from docx.oxml.ns import qn
# from docx.oxml import OxmlElement
# from functions import title_0
#
# # Initialising document to make word file using python
#
# document = Document()
#
# # Code for making Table of Contents
#
# paragraph = document.add_paragraph()
# run = paragraph.add_run()
# fldChar = OxmlElement('w:fldChar')  # creates a new element
# fldChar.set(qn('w:fldCharType'), 'begin')  # sets attribute on element
# instrText = OxmlElement('w:instrText')
# instrText.set(qn('xml:space'), 'preserve')  # sets attribute on element
# instrText.text = 'TOC \\o "1-3" \\h \\z \\u'   # change 1-3 depending on heading levels you need
#
# fldChar2 = OxmlElement('w:fldChar')
# fldChar2.set(qn('w:fldCharType'), 'separate')
# # fldChar3 = OxmlElement('w:t')
# fldChar3 = OxmlElement('w:updateFields')
# fldChar3.set(qn('w:val'), 'true')
# # fldChar3.text = "Right-click to update field."
# fldChar2.append(fldChar3)
#
# fldChar4 = OxmlElement('w:fldChar')
# fldChar4.set(qn('w:fldCharType'), 'end')
#
# r_element = run._r
# r_element.append(fldChar)
# r_element.append(instrText)
# r_element.append(fldChar2)
# r_element.append(fldChar4)
# p_element = paragraph._p
#
# #Giving headings that need to be included in Table of contents
#
# # document.add_heading("Introduccion")
# title_0(document, 'INTRODUCCIÓN', 1)
# title_0(document, 'ANTECEDENTES', 2)
# title_0(document, 'OBJETIVOS', 3)
#
# #Saving the word file by giving name to the file
#
# name = "mdh2"
# document.save(name+".docx")

# Now check word file which got created
#
# Select "Right-click to update field text"
# Now right click and then select update field option
# and then click on update entire table
#
# Now,You will find Automatic Table of Contents







# from docx.oxml.shared import OxmlElement, qn
# from docx import Document
#
# document = Document()
# paragraph = document.add_paragraph()
# run = paragraph.add_run()
# fldChar = OxmlElement('w:fldChar')  # creates a new element
# fldChar.set(qn('w:fldCharType'), 'begin')  # sets attribute on element
# fldChar.text = 'foobar'  # not needed for this element, but this is how you set the text it contains
# r_element = run._r
# r_element.append(fldChar)  # adds new element as last child
# p_element = paragraph._p
# print(p_element.xml)  # shows XML so you can track your progress



#
# from docx import Document
# import os
# import docx
# from docx.shared import Inches, Pt, Cm, Mm
#
# os.chdir(r"C:\Users\user\Desktop\Dann Uc")
#
# # Create doc
# # document = docx.Document()
# document = Document()
#
# # Add black title
# styles = document.styles
# styles['Heading 1'].font.color.rgb = docx.shared.RGBColor(0, 0, 0)
# styles['Heading 1'].font.size = Pt(14)
# styles['Heading 1'].font.italic = False
# styles['Heading 1'].font.bold = True
# styles['Heading 1'].font.name = 'Arial'
#
# # styles = document.styles['Heading 1']
# # font = styles.font
# # font.name = 'Arial'
# # font.size = Pt(20)
# # font.bold = True
# # font.italic = True
#
# # styles['Heading 2'].font.name = 'Arial'
# # styles['Heading 2'].font.color.rgb = docx.shared.RGBColor(0, 0, 0)
# # styles['Heading 2'].font.size = Pt(16)
# # styles['Heading 2'].font.italic = False
# # styles['Heading 2'].font.bold = True
# #
# # styles['Heading 3'].font.name = 'Arial'
# # styles['Heading 3'].font.color.rgb = docx.shared.RGBColor(0, 0, 0)
# # styles['Heading 3'].font.size = Pt(12)
# # styles['Heading 3'].font.italic = False
# # styles['Heading 3'].font.bold = True
#
# document.add_heading('1.  INTRODUCCION', level=1)
# document.add_heading('2.  OBJETIVOS', level=1)
# document.add_heading('3.  ALCANCES', level=1)
#
# # Add text
# paragraph = document.add_paragraph()
# paragraph.add_run('text')
#
# #Saving the word file by giving name to the file
#
# name = "mdh2"
# document.save(name+".docx")



#
# from docx import Document
# import os
# import docx
# from docx.shared import Inches, Pt, Cm, Mm
# from docx.oxml.ns import qn
#
# os.chdir(r"C:\Users\user\Desktop\Dann Uc")
#
# # Create doc
# # document = docx.Document()
# document = Document()
#
# # # Add black title
# # heading = document.add_heading("0.  TITULO GENERAL", 1)
# # title_style = heading.style
# # rFonts = title_style.element.rPr.rFonts
# # rFonts.set(qn("w:asciiTheme"), "Times New Roman")
#
#
# styles = document.styles
# styles['Heading 1'].font.color.rgb = docx.shared.RGBColor(0, 0, 0)
# styles['Heading 1'].font.size = Pt(20)
# styles['Heading 1'].font.italic = False
# styles['Heading 1'].font.bold = True
# styles['Heading 1'].font.name = 'Times New Roman'
#
# a = document.add_heading('1.  INTRODUCCION', level=1)
# # a1 = a.style
# # rFonts = a1.element.rPr.rFonts
# # rFonts.set(qn("w:asciiTheme"), "Comic Sans")
#
# document.add_heading('2.  OBJETIVOS', level=1)
# document.add_heading('3.  ALCANCES', level=1)
#
# # Add text
# paragraph = document.add_paragraph()
# paragraph.add_run('text text text text text text text text')
#
# #Saving the word file by giving name to the file
#
# name = "mdh2"
# document.save(name+".docx")




##################################################################################################################














##################################################################################################################

#
# from docx import Document
# from docx.oxml import OxmlElement, ns
# from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
# from docx.shared import Inches, Pt, Cm, Mm
# from docx.enum.section import WD_SECTION
#
#
#
# def create_element(name):
#     return OxmlElement(name)
#
#
# def create_attribute(element, name, value):
#     element.set(ns.qn(name), value)


# def add_page_number(run):
#     fldChar1 = create_element('w:fldChar')
#     create_attribute(fldChar1, 'w:fldCharType', 'begin')
#
#     instrText = create_element('w:instrText')
#     create_attribute(instrText, 'xml:space', 'preserve')
#     instrText.text = "PAGE"
#
#     fldChar2 = create_element('w:fldChar')
#     create_attribute(fldChar2, 'w:fldCharType', 'end')
#
#     run._r.append(fldChar1)
#     run._r.append(instrText)
#     run._r.append(fldChar2)

# def add_page_number(doc_sec):
#     doc_sec.footer.is_linked_to_previous = False
#     # add_page_number(document.sections[1].footer.paragraphs[0])
#     # document.sections[1].footer.footer_distance = Cm(0.8)
#     doc_sec.footer.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT                # Alineacion del pie de página
#
#     # page_run = paragraph.add_run()
#     # t1 = create_element('w:t')
#     # create_attribute(t1, 'xml:space', 'preserve')
#     # t1.text = 'Page '
#     # page_run._r.append(t1)
#
#     page_num_run = doc_sec.footer.paragraphs[0].add_run()
#     # doc_sec.footer_distance = Cm(1.0)
#     page_num_run.font.size = Pt(12)
#     page_num_run.font.name = 'Calibri'
#
#     fldChar1 = create_element('w:fldChar')
#     create_attribute(fldChar1, 'w:fldCharType', 'begin')
#
#     instrText = create_element('w:instrText')
#     create_attribute(instrText, 'xml:space', 'preserve')
#     instrText.text = "PAGE"
#
#     fldChar2 = create_element('w:fldChar')
#     create_attribute(fldChar2, 'w:fldCharType', 'end')
#
#     page_num_run._r.append(fldChar1)
#     page_num_run._r.append(instrText)
#     page_num_run._r.append(fldChar2)
#
#     # of_run = paragraph.add_run()
#     # t2 = create_element('w:t')
#     # create_attribute(t2, 'xml:space', 'preserve')
#     # t2.text = ' of '
#     # of_run._r.append(t2)
#
#     # fldChar3 = create_element('w:fldChar')
#     # create_attribute(fldChar3, 'w:fldCharType', 'begin')
#     #
#     # instrText2 = create_element('w:instrText')
#     # create_attribute(instrText2, 'xml:space', 'preserve')
#     # instrText2.text = "NUMPAGES"
#     #
#     # fldChar4 = create_element('w:fldChar')
#     # create_attribute(fldChar4, 'w:fldCharType', 'end')
#
#     # num_pages_run = paragraph.add_run()
#     # num_pages_run._r.append(fldChar3)
#     # num_pages_run._r.append(instrText2)
#     # num_pages_run._r.append(fldChar4)
#
# # document = Document()
# # document.add_section(WD_SECTION.NEW_PAGE)
# # document.sections[1].footer.is_linked_to_previous = False
# # add_page_number(document.sections[1])
# # document.save(r'C:\Users\user\Desktop\Dann Uc\Sistema de gestión de documentos\Pruebas\index.docx')







