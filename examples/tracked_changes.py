"""Edit a document with tracked changes, so every edit is a redline to review in Word.

Run: uv run python examples/tracked_changes.py   (no API key needed for part 1)
"""

from docx_harness import WordDocument

# A small document to review.
doc = WordDocument.open()
doc.insert_paragraph("Services Agreement", "Title")
terms = doc.insert_paragraph("The Supplier shall deliver the goods within 30 days of the order date.")
payment = doc.insert_paragraph("Payment is due within 60 days of invoice.")
doc.insert_paragraph("The Customer may cancel at any time without notice.")
doc.save("agreement.docx")

# 1. Direct API: reopen with tracking on. Nothing below changes the text
#    outright; each edit becomes a revision by "Review Agent".
doc = WordDocument.open("agreement.docx", track_changes=True, author="Review Agent")
doc.replace_text("30 days", "14 days")
doc.set_paragraph_text("b3", "Payment is due within 45 days of a valid invoice.")
doc.delete_block("b4")
interest = doc.insert_paragraph("Late payments accrue interest at 2% per month.", after="b3")
doc.move_block("b2", after=interest)         # recorded as a move, not a deletion plus an insertion
doc.save("agreement-redline.docx")

print(doc.text(), "\n")                      # the document as it reads once accepted
for revision in doc.revisions():
    print(revision["type"], repr(revision["text"]) if revision["text"] else f"(whole {revision['target']})")

# The reviewer's side: reject the move, accept everything else.
moves = [revision["id"] for revision in doc.revisions() if revision["type"] == "move"]
doc.reject_revisions(ids=moves)
doc.accept_revisions()
doc.save("agreement-reviewed.docx")
print("\n" + doc.text())

# 2. The same through the agent (needs OPENROUTER_API_KEY in .env):
#
#   from dotenv import load_dotenv
#   from docx_harness import draft
#   load_dotenv()
#   draft("Make the cancellation clause require 30 days written notice",
#         "agreement.docx", "agreement-agent-redline.docx",
#         track_changes=True, author="Review Agent")
