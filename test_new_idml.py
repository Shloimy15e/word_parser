from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))

from main import extract_text_from_idml

idml_file = Path(r'backup_input\שס 2\ביצה\PEREK1.idml')

print("Testing improved IDML extraction...\n")
print("=" * 80)

paragraphs = extract_text_from_idml(idml_file)

print(f"\nExtracted {len(paragraphs)} paragraphs\n")
print("First 10 paragraphs:\n")
print("=" * 80)

for i, para in enumerate(paragraphs[:10], 1):
    print(f"\nParagraph {i}:")
    print(f"  Length: {len(para)} chars")
    # Show first 200 chars
    display = para[:200] if len(para) > 200 else para
    print(f"  Text: {display}")
    if len(para) > 200:
        print(f"  ... (+ {len(para) - 200} more chars)")

# Write full output to file
with open('idml_extracted_text.txt', 'w', encoding='utf-8') as f:
    for i, para in enumerate(paragraphs, 1):
        f.write(f"\n{'=' * 60}\n")
        f.write(f"Paragraph {i} ({len(para)} chars):\n")
        f.write(f"{'=' * 60}\n")
        f.write(para)
        f.write("\n")

print(f"\n\nFull output written to idml_extracted_text.txt")

