# python-docx

*python-docx* is a Python library for reading, creating, and updating Microsoft Word 2007+ (.docx) files.

## ✨ DOCM Support (Macro-Enabled Documents)

This fork adds full support for `.docm` files with intelligent macro preservation:

- **Open DOCM files** just like DOCX files
- **Automatically strips macros** when saving to `.docx`
- **Automatically preserves macros** when saving to `.docm`
- **Explicit control** via `preserve_macros` parameter

```python
from docx import Document

# Open and work with DOCM files
doc = Document('macro_document.docm')

# Save as DOCX (strips macros)
doc.save('output.docx')

# Save as DOCM (preserves macros)
doc.save('output.docm')
```

📖 **[Full DOCM Documentation →](DOCM_SUPPORT.md)**

## Installation

```
pip install python-docx
```

## Example

```python
>>> from docx import Document

>>> document = Document()
>>> document.add_paragraph("It was a dark and stormy night.")
<docx.text.paragraph.Paragraph object at 0x10f19e760>
>>> document.save("dark-and-stormy.docx")

>>> document = Document("dark-and-stormy.docx")
>>> document.paragraphs[0].text
'It was a dark and stormy night.'
```

More information is available in the [python-docx documentation](https://python-docx.readthedocs.org/en/latest/)
