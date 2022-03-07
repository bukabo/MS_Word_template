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


if __name__ == '__main__':
    document = MailMerge(template)
    print(document.get_merge_fields())
    write_docx(document, 'Александров Д.В.', '001A')
