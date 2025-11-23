from docx import Document
from pathlib import Path

output_file = Path(r'output\שס\ביצה\PEREK1-formatted.docx')

doc = Document(output_file)

# Write to file instead of printing
with open('output_analysis.txt', 'w', encoding='utf-8') as f:
    f.write("First 20 paragraphs from output:\n\n")
    f.write("=" * 80 + "\n")
    
    for i, para in enumerate(doc.paragraphs[:20], 1):
        text = para.text.strip()
        if text:
            style = para.style.name if para.style else "Normal"
            f.write(f"{i}. [{style}]\n")
            # Show first 150 chars
            display = text[:150] if len(text) > 150 else text
            f.write(f"   {display}\n")
            if len(text) > 150:
                f.write(f"   ... ({len(text)} chars total)\n")
            f.write("\n")

print("Output written to output_analysis.txt")

