from docx import Document
from docx2pdf import convert
document = Document('Invoice.docx')

dic = {
    '${a_1}': '1,380.00',
    '${a_2}': '23.00',
    '${a_3}': '0.00 %',
    '${a_4}': '23.00',
    '${a_5}': 'حبة E',
    '${a_6}': '60',
    '${a_7}': 'This is Going to be the long text to',
    # TODO , Replacing this For the Tempalte 
    '${a_8}': 'abcvvv',
    '${a_9}': '1',
}
for p in document.paragraphs:
    inline = p.runs
    for i in range(len(inline)):
        text = inline[i].text
        for key in dic.keys():
            if key in text:
                text = text.replace(key, dic[key])
                inline[i].text = text

document.save('new.docx')

# https://stackoverflow.com/questions/34779724/python-docx-replace-string-in-paragraph-while-keeping-style/57365257#57365257
# convert("new.docx", "output.pdf")
