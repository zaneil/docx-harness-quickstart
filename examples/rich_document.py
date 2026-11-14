"""Lists, links, a picture, a header and footer, and table borders, using the
document layer on its own (no model, no API key).

Run: uv run python examples/rich_document.py [picture.png]
"""

import sys

from docx_harness import WordDocument

doc = WordDocument.open()
doc.insert_paragraph("Launch checklist", "Title")
intro = doc.insert_paragraph("Everything the team needs before release day. Details are in the handbook.")
doc.add_hyperlink(intro, "https://example.com/handbook", match="the handbook")

doc.insert_paragraph("Steps", "Heading 1")
doc.insert_paragraph("Freeze the release branch", "List Number")           # a new numbered list: starts at 1
doc.insert_paragraph("Tag the last commit", "List Bullet", level=1)        # nested under step 1
doc.insert_paragraph("Tell the other teams", "List Bullet", level=1)
doc.insert_paragraph("Run the full test suite", "List Number")             # step 2 of the same list

doc.insert_paragraph("After release", "Heading 1")
doc.insert_paragraph("Watch the dashboards for an hour", "List Number")    # a separate list: starts at 1 again
doc.insert_paragraph("Write the release notes", "List Number")

doc.insert_paragraph("Owners", "Heading 4")                                # not in the template: added on first use
table = doc.insert_table([["Area", "Owner", "Status"], ["Backend", "Priya", "Ready"], ["Docs", "Tom", "In review"]])
doc.set_table_format(table, borders="horizontal", cell_margins_in=0.05)
doc.format_cells(f"{table}[0,*]", shading="D9E2F3", border="thick", border_sides=["bottom"])

if len(sys.argv) > 1:                                                      # PNG, JPEG or GIF
    figure = doc.insert_image(sys.argv[1], width_in=4, alt_text="Release timeline", alignment="center")
    doc.insert_paragraph("Figure 1: release timeline", "Caption", after=figure)

doc.set_header("Launch checklist\tInternal")                               # a tab moves to the centre
doc.set_footer("Page {PAGE} of {NUMPAGES}", alignment="center")

for entry in doc.outline():
    print(entry["id"], entry.get("style", entry["type"]), "|", entry["preview"])
print(doc.save("checklist.docx"))
