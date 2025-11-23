import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

idml_path = Path(r'backup_input\שס 2\ביצה\PEREK1.idml')

with zipfile.ZipFile(idml_path, 'r') as z:
    # Get all story files
    story_files = []
    for name in z.namelist():
        if name.startswith('Stories/') and name.endswith('.xml'):
            info = z.getinfo(name)
            story_files.append((name, info.file_size))
    
    story_files.sort(key=lambda x: x[1], reverse=True)
    
    with open('page_structure_analysis.txt', 'w', encoding='utf-8') as out:
        out.write("IDML Page Structure Analysis\n")
        out.write("=" * 80 + "\n\n")
        
        # Find page reference patterns
        page_references = []
        
        for story_file, size in story_files:
            with z.open(story_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                
                # Extract text
                texts = []
                for elem in root.iter():
                    tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
                    if tag == 'Content' and elem.text:
                        texts.append(elem.text.strip())
                
                combined = ' '.join(texts).strip()
                
                # Check if this looks like a page reference
                # Pattern: "ביצה" followed by Hebrew numerals and colons
                if combined and 'ביצה' in combined and ':' in combined:
                    page_references.append({
                        'file': story_file,
                        'size': size,
                        'text': combined
                    })
        
        out.write(f"Found {len(page_references)} page reference markers:\n\n")
        
        for ref in page_references:
            out.write(f"File: {ref['file']}\n")
            out.write(f"Size: {ref['size']} bytes\n")
            out.write(f"Text: {ref['text']}\n")
            
            # Try to parse the reference
            text = ref['text']
            
            # Remove "ביצה" prefix
            if text.startswith('ביצה '):
                page_part = text[5:].strip()
                
                out.write(f"Page part: {page_part}\n")
                
                # Check for range (contains dash or multiple colons)
                if '-' in page_part or page_part.count(':') > 1:
                    out.write(f"Type: PAGE RANGE\n")
                elif ':' in page_part:
                    out.write(f"Type: SINGLE PAGE\n")
                    parts = page_part.split(':')
                    out.write(f"  Daf (page): {parts[0].strip()}\n")
                    if len(parts) > 1 and parts[1].strip():
                        out.write(f"  Possibly Amud marker\n")
                
            out.write("\n" + "-" * 60 + "\n\n")

print("Analysis written to page_structure_analysis.txt")

