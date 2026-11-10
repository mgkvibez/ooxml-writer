"""Generate tests/fixtures/deck.pptx, the real deck the round-trip tests edit.

Committed as a binary so the suite needs no PowerPoint toolchain; this script is the provenance record. It
deliberately includes the three things a naive editor destroys: one phrase split across runs, a DrawingML
table (``<a:tbl>``, not the
WordprocessingML ``<w:tbl>`` a DOCX uses), and a binary image part that must survive byte-for-byte.

    python3 -m pip install python-pptx && python3 tools/make-fixture.py
"""

from PIL import Image
from pptx import Presentation
from pptx.util import Inches

OUT = "tests/fixtures/deck.pptx"

# a real PNG so the package contains a binary part with its own compression and content type
Image.new("RGB", (16, 16), (30, 120, 200)).save("tests/fixtures/pixel.png")

presentation = Presentation()
slide = presentation.slides.add_slide(presentation.slide_layouts[1])  # Title and Content

title = slide.shapes.title
title.text = "Quarterly review"
body = slide.placeholders[1].text_frame
body.text = "Revenue is up."
paragraph = body.add_paragraph()
paragraph.text = "Costs are flat."
paragraph = body.add_paragraph()
paragraph.text = "Outlook is steady."

# a phrase split across runs, which is how PowerPoint stores mixed formatting
textbox = slide.shapes.add_textbox(Inches(1), Inches(5), Inches(4), Inches(1))
frame = textbox.text_frame
frame.text = "Hello "
frame.paragraphs[0].add_run().text = "wor"
frame.paragraphs[0].add_run().text = "ld"

# a DrawingML table: a:tbl / a:tr / a:tc, which is a different namespace from Word's
rows, columns = 3, 2
table = slide.shapes.add_table(rows, columns, Inches(5), Inches(4.5), Inches(3), Inches(1.2)).table
data = [["Region", "Revenue"], ["North", "120"], ["South", "98"]]
for r, row in enumerate(data):
    for c, value in enumerate(row):
        table.cell(r, c).text = value

slide.shapes.add_picture("tests/fixtures/pixel.png", Inches(9), Inches(0.2))

second = presentation.slides.add_slide(presentation.slide_layouts[1])
second.shapes.title.text = "Appendix"
second.placeholders[1].text_frame.text = "Nothing to report."

presentation.save(OUT)
print("wrote", OUT)
