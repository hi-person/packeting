from openpyxl import load_workbook
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from math import floor
wb = load_workbook("2026 smh invitational cup question drive old.xlsx")
#subjects = ["bio", "physics", "math", "chem", "eSS", "csenergy"]
subjects = ["math", "physics"]
doc = Document()
t = 2
for i in range(2, 51):
   for a in subjects:
      ws = wb[a]
      if ws[f"R{i}"].value == "y":
         # toss-up section
         p = doc.add_paragraph()
         p.alignment = WD_ALIGN_PARAGRAPH.CENTER
         run = p.add_run("TOSS-UP").bold = True
         if ws[f"E{i}"].value == "SA":
            p = doc.add_paragraph()
            p.add_run(f"{round(floor(t/2))})  {a}  ")
            p.add_run("Short Answer  ").italic = True
            p.add_run(f"{ws[f"F{i}"].value}")
            t += 1
         elif ws[f"E{i}"].value == "MC":
            p = doc.add_paragraph()
            p.add_run(f"{round(floor(t/2))})  {a}  ")
            p.add_run("Multiple Choice  ").italic = True
            p.add_run(f"{ws[f"F{i}"].value}")
            t += 1

         p = doc.add_paragraph()
         run = p.add_run(f"Answer: {ws[f"G{i}"].value} [{ws[f"B{i}"].value}]  ")
         if ws[f"H{i}"].value is not None:
            run = p.add_run(f"(Do Not Accept: {ws[f"H{i}"].value})")

         #bonus section
         p = doc.add_paragraph()
         p.alignment = WD_ALIGN_PARAGRAPH.CENTER
         run = p.add_run("BONUS").bold = True
         doc.add_paragraph()

         if ws[f"I{i}"].value == "SA":
            p = doc.add_paragraph()
            p.add_run(f"{round(floor(t/2))})  {a}  ")
            p.add_run("Short Answer  ").italic = True
            p.add_run(f"{ws[f"J{i}"].value}")
            t += 1
         elif ws[f"I{i}"].value == "MC":
            p = doc.add_paragraph()
            p.add_run(f"{round(floor(t/2))})  {a}  ")
            p.add_run("Multiple Choice  ").italic = True
            p.add_run(f"{ws[f"J{i}"].value}")
            t += 1
         elif ws[f"I{i}"].value == "VB":
            p = doc.add_paragraph()
            p.add_run(f"{round(floor(t/2))})  {a}  ")
            p.add_run("Visual Bonus  ").italic = True
            p.add_run(f"{ws[f"J{i}"].value}")
            t += 1

         p = doc.add_paragraph()
         run = p.add_run(f"Answer: {ws[f"K{i}"].value} [{ws[f"B{i}"].value}]  ")
         if ws[f"H{i}"].value is not None:
            run = p.add_run(f"(Do Not Accept: {ws[f"L{i}"].value})")
      if t % 4 == 2:
         doc.add_page_break()
      else:
         doc.add_paragraph()


doc.save("doc.docx")