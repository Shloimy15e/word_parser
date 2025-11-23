from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))

from main import extract_text_from_idml

idml_file = Path(r'backup_input\שס 2\ביצה\PEREK1.idml')

paragraphs = extract_text_from_idml(idml_file)

with open('test_output.txt', 'w', encoding='utf-8') as f:
    f.write(f"Total paragraphs extracted: {len(paragraphs)}\n\n")
    f.write("=" * 80 + "\n\n")
    
    for i, para in enumerate(paragraphs[:15], 1):
        f.write(f"Paragraph {i} (Length: {len(para)}):\n")
        f.write(para)
        f.write("\n\n")
        f.write("-" * 60 + "\n\n")

print(f"Extracted {len(paragraphs)} paragraphs - written to test_output.txt")

