import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

idml_path = Path(r'backup_input\שס 2\ביצה\PEREK1.idml')

with zipfile.ZipFile(idml_path, 'r') as z:
    # Focus on the largest story only
    story_file = 'Stories/Story_ue38.xml'
    
    with z.open(story_file) as f:
        tree = ET.parse(f)
        root = tree.getroot()
        
        paragraphs = []
        
        for para_elem in root.iter():
            tag_name = para_elem.tag.split('}')[-1] if '}' in para_elem.tag else para_elem.tag
            
            if tag_name == 'ParagraphStyleRange':
                # Collect all Content text
                texts = []
                for content in para_elem.iter():
                    c_tag = content.tag.split('}')[-1] if '}' in content.tag else content.tag
                    if c_tag == 'Content' and content.text:
                        texts.append(content.text)
                
                if texts:
                    para = ''.join(texts).strip()
                    if para and len(para) > 10:  # Only paragraphs > 10 chars
                        paragraphs.append(para)

with open('main_story_output.txt', 'w', encoding='utf-8') as f:
    f.write(f"Extracted {len(paragraphs)} paragraphs from main story\n\n")
    f.write("=" * 80 + "\n\n")
    
    for i, para in enumerate(paragraphs[:10], 1):
        f.write(f"Paragraph {i} ({len(para)} chars):\n")
        f.write(para[:500])  # First 500 chars
        if len(para) > 500:
            f.write(f"\n... (+ {len(para) - 500} more chars)")
        f.write("\n\n" + "-" * 60 + "\n\n")

print(f"Extracted {len(paragraphs)} paragraphs - written to main_story_output.txt")

