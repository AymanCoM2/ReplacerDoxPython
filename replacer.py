from python_docx_replace import docx_replace
from docx import Document
from docx2pdf import convert
doc = Document("Invoice.docx")

docx_replace(doc, account_number="123")
docx_replace(doc, due_date="12/12/2023")
docx_replace(doc, order_ref="45234")
docx_replace(doc, invoice_number="233453")
docx_replace(doc, invoice_date="12/12/2023")
docx_replace(doc, invoice_type="Normal Invoice")
docx_replace(doc, customer_name="Ayman Salah")

# line  = "${f_1}           ${f_2}  ${f_3}   ${f_4}   ${f_5}  ${f_6}                                   ${f_7}        ${f_8}  ${f_9}"
# lo = "2323             343        20%      567     Good Lamp to     15                   This is Going to be the Long One          4    1"
docx_replace(doc, a_1="23")
docx_replace(doc,   a_2="45", a_3="fgf")
docx_replace(doc, a_4="23", a_5="45", a_6="fgf")
docx_replace(doc, a_7="23", a_8="45", a_9="fgf")
# docx_replace(doc, line=lo)
doc.save("replaced.docx")

# convert("replaced.docx", "output.pdf")
