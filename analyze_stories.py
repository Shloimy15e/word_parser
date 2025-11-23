import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

idml_path = Path(r'backup_input\שס 2\ביצה\PEREK1.idml')

with zipfile.ZipFile(idml_path, 'r') as z:
    story_files = [name for name in z.namelist() if name.startswith('Stories/') and name.endswith('.xml')]
    
    # Get story sizes
    story_sizes = []
    for story_file in story_files:
        info = z.getinfo(story_file)
        story_sizes.append((story_file, info.file_size))
    
    story_sizes.sort(key=lambda x: x[1], reverse=True)
    
    print("Analyzing all story files for content type:\n")
    print("=" * 80)
    
    main_content_size = story_sizes[0][1]  # Largest is main content
    
    for story_file, size in story_sizes:
        with z.open(story_file) as f:
            content = f.read().decode('utf-8')
            tree = ET.fromstring(content.encode('utf-8'))
            
            # Get all text content
            texts = []
            for elem in tree.iter():
                tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
                if 'Content' in tag and elem.text:
                    texts.append(elem.text)
            
            all_text = ' '.join(texts).strip()
            
            # Categorize
            is_main = size == main_content_size
            has_brackets = '[' in all_text or ']' in all_text
            is_small = size < 2000
            
            status = "MAIN CONTENT" if is_main else "Small text" if is_small else "Unknown"
            
            print(f"{story_file}")
            print(f"  Size: {size} bytes | {len(texts)} text elements")
            print(f"  Type: {status}")
            if has_brackets:
                print(f"  Contains brackets: YES")
            if all_text and len(all_text) < 200:
                safe_text = all_text[:100].replace('\n', ' ')
                print(f"  Text: {safe_text}")
            elif all_text:
                safe_text = all_text[:100].replace('\n', ' ')
                print(f"  Text preview: {safe_text}...")
            print()

