import aspose.words as aw

doc = aw.Document()
builder = aw.DocumentBuilder(doc)
curentSection = builder.current_section
pageSetup = curentSection.page_setup

pageSetup.different_first_page_header_footer = True
pageSetup.header_distance = 20
builder.move_to_header_footer(aw.HeaderFooterType.HEADER_FIRST)

font = builder.font
font.size = 11
font.name = 'TimesNewRoman'

table = builder.start_table()

builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER

builder.insert_cell()
builder.cell_format.width = 100

builder.writeln('ROMÂNIA')
builder.writeln('AVIATION SRL')
unitate = input('Introduceti denumirea unitatii: ').upper()
nume = input('Intorduceti numele dumneavoastra: ').upper()
builder.writeln(f'{unitate}')
builder.writeln(f'{nume}')
builder.writeln('Nr.                   din         ')

builder.insert_cell()

builder.cell_format.width = 50
builder.paragraph_format.alignment = aw.ParagraphAlignment.RIGHT

builder.writeln('Exemplar unic')

builder.end_row()

builder.end_table()

table_style = doc.styles.add(aw.StyleType.TABLE, 'Mystyle').as_table_style()
table_style.borders.line_style = aw.LineStyle.NONE
table.style = table_style

builder.move_to_document_start()

for _ in range(6):
    builder.writeln()

font.size =  18
font.bold = True
builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
destinatar = input('Cui va adresati: ').upper()
builder.writeln(f'{destinatar}')


for _ in range(4):
    builder.writeln()

font.size = 12
font.name = 'TimesNewRoman'
font.bold = False
font = builder.font
paragraphformat = builder.paragraph_format

paragraphformat.alignment = aw.ParagraphAlignment.JUSTIFY

data_concediu_anterior = input('Introduceti concediul pe care doriti sa-l modificati: ')
data_concediu_corectat = input('Introduceti noua perioada de concediu: ')

motiv = input('Introduceti motivul: ')
builder.writeln('Raportez:')
builder.paragraph_format.left_indent = 8
builder.writeln(f'''Va rog sa-mi aprobati modificarea concediului din perioada {data_concediu_anterior} in perioada {data_concediu_corectat}, din cauza faptului ca {motiv}''')

for _ in range(4):
    builder.writeln()

table2 = builder.start_table()

builder.paragraph_format.alignment = aw.ParagraphAlignment.LEFT

builder.insert_cell()
builder.cell_format.width = 100

builder.writeln('Data: ')
data = input('Intorduceti data de astazi: ')
builder.writeln(f'{data}')

builder.insert_cell()

builder.cell_format.width = 50
builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER

builder.writeln('Semnatura: ')


builder.end_row()
builder.end_table()

table_style2 = doc.styles.add(aw.StyleType.TABLE, 'Mystyle2').as_table_style()
table_style2.borders.line_style = aw.LineStyle.NONE
table2.style = table_style2


builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_FIRST)
table3 = builder.start_table()


builder.insert_cell()
builder.current_paragraph.paragraph_format.alignment = aw.ParagraphAlignment.CENTER

builder.insert_field('PAGE','')
builder.write(' din ')
builder.insert_field('NUMPAGES','')
builder.end_row()
builder.end_table()

table_style3 = doc.styles.add(aw.StyleType.TABLE, 'Mystyle3').as_table_style()
table_style3.borders.line_style = aw.LineStyle.NONE
table3.style = table_style3

doc.save("E:\\2023\\2023_ITSchool\\ITSchool Lucru individual\\docum.docx")
