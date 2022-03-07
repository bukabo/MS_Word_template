"""
https://pbpython.com/python-word-template.html
conda install lxml
pip install docx-mailmerge
in word: Insert -> Quick Parts -> Field
From the Field dialog box, select the “MergeField” option from the Field Names list.
In the Field Name, enter the name you want for the field.
"""

from __future__ import print_function
from mailmerge import MailMerge
from datetime import date

template = "template.docx"


def write_docx(document, name, train):
    document.merge(
        text_field=name,
        num_train=train,
        date=f'{date.today()}')

    document.write('test-output.docx')


def write_docx_multipage(document, name, train):
    page_1 = {
        'text_field': name,
        'num_train': train,
        'date': '{:%d-%b-%Y}'.format(date.today())}
    page_2 = {
        'text_field': name,
        'num_train': train,
        'date': f'{date.today()}'}
    page_3 = {
        'text_field': name,
        'num_train': train,
        'date': f'{date.today()}'}
    document.merge_templates([page_1, page_2, page_3], separator='page_break')
    document.write('test-output-multi-page.docx')


if __name__ == '__main__':
    document = MailMerge(template)
    document2 = MailMerge(template)
    print(document.get_merge_fields())
    write_docx(document, 'Александров Д.В.', '001A')
    write_docx_multipage(document2, 'Александров Д.В.', '0A')
