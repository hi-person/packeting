from openpyxl import load_workbook
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from math import floor, ceil
wb = load_workbook("2026 smh invitational cup question drive.xlsx")
subjects = ["Biology", "Physics", "Math", "Chemistry", "Earth and Space", "Energy", "Computer Science"]
doc = Document()
t = 0
#i is row, t is question number, a is subject
def qwrite(i, t, a):
   p = doc.add_paragraph()
   p.alignment = WD_ALIGN_PARAGRAPH.CENTER
   run = p.add_run("TOSS-UP").bold = True
   if ws[f"E{i}"].value == "SA":
      p = doc.add_paragraph()
      p.add_run(f"{1 + round(floor(t / 2)) % 22})  {a}  ")
      p.add_run("Short Answer  ").italic = True
      p.add_run(f"{ws[f"F{i}"].value}")
      t += 1
   elif ws[f"E{i}"].value == "MC":
      p = doc.add_paragraph()
      p.add_run(f"{1 + round(floor(t / 2)) % 22})  {a}  ")
      p.add_run("Multiple Choice  ").italic = True
      p.add_run(f"{ws[f"F{i}"].value}")
      t += 1

   p = doc.add_paragraph()
   run = p.add_run(f"Answer: {ws[f"G{i}"].value} [{ws[f"B{i}"].value}]  ")
   if ws[f"H{i}"].value is not None:
      run = p.add_run(f"(Do Not Accept: {ws[f"H{i}"].value})")

   # bonus section
   p = doc.add_paragraph()
   p.alignment = WD_ALIGN_PARAGRAPH.CENTER
   run = p.add_run("BONUS").bold = True
   doc.add_paragraph()

   if ws[f"I{i}"].value == "SA":
      p = doc.add_paragraph()
      p.add_run(f"{1 + round(floor(t / 2)) % 22})  {a}  ")
      p.add_run("Short Answer  ").italic = True
      p.add_run(f"{ws[f"J{i}"].value}")
   elif ws[f"I{i}"].value == "MC":
      p = doc.add_paragraph()
      p.add_run(f"{1 + round(floor(t / 2)) % 22})  {a}  ")
      p.add_run("Multiple Choice  ").italic = True
      p.add_run(f"{ws[f"J{i}"].value}")
   elif ws[f"I{i}"].value == "VB":
      p = doc.add_paragraph()
      p.add_run(f"{1 + round(floor(t / 2)) % 22})  {a}  ")
      p.add_run("Visual Bonus  ").italic = True
      p.add_run(f"{ws[f"J{i}"].value}")

   p = doc.add_paragraph()
   run = p.add_run(f"Answer: {ws[f"K{i}"].value} [{ws[f"B{i}"].value}]  ")
   if ws[f"L{i}"].value is not None:
      run = p.add_run(f"(Do Not Accept: {ws[f"L{i}"].value})")


for i in range(2, 71):
   for a in subjects:
      ws = wb[a]
      if a not in ["Energy", "Computer Science"] and ws[f"R{i}"].value == "y":
         qwrite(i, t, a)
         t += 2
      elif i % 4 == 3 and ws[f"R{i}"].value == "y" and a == "Energy":
         qwrite(ceil(i/4)+1, t, a)
         t += 2
      elif i % 4 == 1 and ws[f"R{ceil(i/2)}"].value == "y" and a == "Computer Science":
         qwrite(ceil(i/4), t, a)
         t += 2
      if t % 4 == 0 and a != "Computer Science" and a != "Energy":
         doc.add_page_break()
      elif t % 44 == 0:
         doc.add_page_break()
      else:
         doc.add_paragraph()


doc.save(f"doc.docx")
