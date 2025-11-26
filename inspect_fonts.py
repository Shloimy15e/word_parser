from word_parser.readers.docx_reader import DocxReader
from pathlib import Path

from word_parser.readers.docx_reader import DocxReader
from pathlib import Path

from word_parser.readers.docx_reader import DocxReader
from pathlib import Path

from word_parser.readers.docx_reader import DocxReader
from pathlib import Path

file_path = Path(r"C:\Users\shloi\Desktop\word_parser\docs\זכר חלק ו די גאנצע.docx")
reader = DocxReader()
doc = reader.read(file_path)

print(f"Total paragraphs: {len(doc.paragraphs)}")

with open("inspection_output.txt", "w", encoding="utf-8") as f:
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        if (
            "יד." in text
            or "דער רבי נעמט" in text
            or "טו." in text
            or "אם ראשונים" in text
        ):
            f.write(f"Para {i}: '{text}'\n")
            for run in para.runs:
                f.write(f"  Run: '{run.text}' Size: {run.style.font_size}\n")
            f.write("-" * 20 + "\n")
