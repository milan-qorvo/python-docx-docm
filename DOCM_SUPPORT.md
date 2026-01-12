# DOCM File Support

This fork adds comprehensive support for DOCM (macro-enabled Word documents) files with intelligent macro preservation.

## Features

### ✅ Open DOCM Files
Load and read DOCM files just like DOCX files:
```python
from docx import Document

doc = Document('my_document.docm')
# Work with the document normally
for para in doc.paragraphs:
    print(para.text)
```

### ✅ Conditional Macro Preservation

The library automatically detects your intent based on the file extension:

#### Save as DOCX → Macros are stripped
```python
doc = Document('input.docm')
doc.save('output.docx')  # Macros automatically removed
```

#### Save as DOCM → Macros are preserved
```python
doc = Document('input.docm')
doc.save('output.docm')  # Macros preserved
```

### ✅ Explicit Control

Override the automatic behavior with the `preserve_macros` parameter:

```python
doc = Document('input.docm')

# Force strip macros even with .docm extension
# (Note: Extension will be auto-corrected to .docx)
doc.save('output.docm', preserve_macros=False)  # Saves as output.docx

# Force preserve macros even with .docx extension
# (Content type will be DOCM)
doc.save('output.docx', preserve_macros=True)
```

### ✅ Stream Support

When saving to file-like objects (BytesIO, etc.):

```python
from io import BytesIO

doc = Document('input.docm')

# Default: strips macros (safer)
stream = BytesIO()
doc.save(stream)

# Explicit: preserve macros
stream = BytesIO()
doc.save(stream, preserve_macros=True)
```

## What Gets Removed When Stripping Macros

When converting DOCM → DOCX, the following are removed:

1. ✅ Content type changed from `macroEnabled.main` to `document.main`
2. ✅ VBA project relationships removed
3. ✅ WordVbaData relationships removed
4. ✅ ActiveX control relationships removed
5. ✅ `<w:control>` XML elements removed from document
6. ✅ VBA binary parts (`/word/vbaProject.bin`) removed from package
7. ✅ VBA data parts removed from package
8. ✅ All orphaned relationships cleaned up

## Extension Auto-Correction

To prevent creating invalid files that Word cannot open:

- When stripping macros (DOCM → DOCX conversion), if you specify a `.docm` extension, it will automatically be corrected to `.docx`
- This ensures the file extension matches the content type

Example:
```python
doc = Document('input.docm')
doc.save('output.docm', preserve_macros=False)
# Actually saves to: output.docx (auto-corrected)
```

## Implementation Details

### Modified Files

1. **src/docx/api.py**
   - Updated `Document()` function to accept both DOCX and DOCM content types

2. **src/docx/opc/constants.py**
   - Added `WML_DOCUMENT_MACRO_ENABLED_MAIN` content type constant

3. **src/docx/__init__.py**
   - Registered DOCM content type to use `DocumentPart`

4. **src/docx/parts/document.py**
   - Added `preserve_macros` parameter to `save()` method
   - Added `_should_preserve_macros()` helper for auto-detection
   - Added `_remove_macro_relationships()` to clean up VBA relationships and ActiveX controls
   - Added `_remove_vba_parts()` to remove VBA binary files from package
   - Added extension auto-correction for DOCM→DOCX conversions

5. **src/docx/document.py**
   - Updated `save()` method to pass through `preserve_macros` parameter

6. **tests/test_document.py**
   - Updated test to expect new `preserve_macros` parameter

7. **tests/parts/test_document.py**
   - Added 7 comprehensive unit tests for DOCM functionality

8. **tests/test_api.py**
   - Added test for opening DOCM files

## Testing

```bash
# Run all tests
pytest -W ignore::DeprecationWarning

# Run just DOCM-specific tests
pytest tests/parts/test_document.py::DescribeDocumentPart -k "macro" -v
pytest tests/test_api.py::DescribeDocument::it_opens_a_docm_file -v
```

Unit tests cover:
- ✅ File extension-based macro preservation detection
- ✅ Explicit `preserve_macros` flag override
- ✅ BytesIO/stream handling (default strip, explicit preserve)
- ✅ DOCM → DOCX conversion with macro stripping
- ✅ DOCM → DOCM preservation
- ✅ Extension auto-correction for invalid combinations
- ✅ Regular DOCX files unaffected by DOCM logic
- ✅ Opening DOCM files via Document() API

**All 1617 tests pass** (including 7 new DOCM-specific tests).

## Caveats

1. **No Macro Execution**: This library never executes macros, only preserves/removes them
2. **No Macro Reading**: You cannot read or inspect macro code
3. **Binary Preservation Only**: When preserving, macros are kept as-is in their binary format
4. **One-Way Stripping**: Once stripped, macros cannot be recovered (keep backups!)

## Safety

Default behavior prioritizes safety:
- Saving to streams defaults to stripping macros
- Explicit control available when needed
- Invalid file configurations are prevented (extension auto-correction)
- All VBA content is thoroughly removed when stripping

## Compatibility

- Works with all existing python-docx code
- Backward compatible (existing DOCX workflows unchanged)
- DOCM files can be opened and manipulated just like DOCX files
- Output files successfully open in Microsoft Word