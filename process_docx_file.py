import zipfile
from xml.etree.ElementTree import XML

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
NUMBER = WORD_NAMESPACE + 'numPr'
FOOTNOTE = WORD_NAMESPACE + 'footnote'


def get_file_content(in_file):
    with zipfile.ZipFile(in_file) as docx:
        tree = XML(docx.read('word/document.xml'))

    paragraphs = []
    for paragraph in tree.iter(PARA):
        texts = [node.text
                 for node in paragraph.iter(TEXT)
                 if node.text]
        if texts:
            paragraphs.append(''.join(texts))

    final_text = ''
    for p in paragraphs:
        p = p.strip()
        final_text += p + '\n'

    return final_text


def get_file_footnotes(in_file):
    with zipfile.ZipFile(in_file) as docx:
        tree = XML(docx.read('word/footnotes.xml'))

    paragraphs = []
    for paragraph in tree.iter(FOOTNOTE):
        temp = ''  # Used to read the list in footnotes ( multiple lines instead of one line contain all data )
        for par in paragraph.iter(PARA):
            texts = [node.text
                     for node in par.getiterator(TEXT)
                     if node.text]
            if texts:
                temp += ''.join(texts) + "\n"
        paragraphs.append(temp)

    final_text = ''
    for p in paragraphs:
        if p:
            final_text += p + '\n'

    return final_text


def run(in_file, content_file, footnotes_file):
    # Write the content
    paragraphs = get_file_content(in_file)
    f = open(content_file, "w+", encoding='utf-8')
    f.write(paragraphs)
    f.close()

    # Write the foornotes
    footnotes = get_file_footnotes(in_file)
    f = open(footnotes_file, "w+", encoding='utf-8')
    f.write(footnotes)
    f.close()


if __name__ == "__main__":
    run(r"sample.docx", "sample.txt", "sample_footnotes.txt")
