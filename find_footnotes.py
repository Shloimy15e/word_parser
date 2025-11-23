import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

idml_path = Path(r'backup_input\שס 2\ביצה\PEREK1.idml')

print("Looking for footnotes and text structure...")

with zipfile.ZipFile(idml_path, 'r') as z:
    story_files = [name for name in z.namelist() if name.startswith('Stories/') and name.endswith('.xml')]
    
    # Sort by size - larger files likely have more content
    story_sizes = []
    for story_file in story_files:
        info = z.getinfo(story_file)
        story_sizes.append((story_file, info.file_size))
    
    story_sizes.sort(key=lambda x: x[1], reverse=True)
    
    print(f"\nTop 10 largest story files:")
    for story, size in story_sizes[:10]:
        print(f"  {story}: {size} bytes")
    
    # Examine the largest story files
    print("\n" + "=" * 80)
    print("Examining largest story files for structure:")
    print("=" * 80)
    
    for story_file, size in story_sizes[:3]:
        print(f"\n{'=' * 60}")
        print(f"Story: {story_file} ({size} bytes)")
        print(f"{'=' * 60}")
        
        with z.open(story_file) as f:
            content = f.read().decode('utf-8')
            
            # Save largest story
            if story_file == story_sizes[0][0]:
                with open('largest_story.xml', 'w', encoding='utf-8') as out:
                    out.write(content)
                print(f"Saved to largest_story.xml")
            
            tree = ET.fromstring(content.encode('utf-8'))
            
            # Count paragraphs and text content
            para_count = 0
            char_count = 0
            
            # Extract all text in order
            texts = []
            for para in tree.iter():
                tag = para.tag.split('}')[-1] if '}' in para.tag else para.tag
                if 'ParagraphStyleRange' in tag:
                    para_count += 1
                elif 'Content' in tag and para.text:
                    texts.append(para.text)
                    char_count += len(para.text)
            
            print(f"Paragraphs: {para_count}")
            print(f"Text elements: {len(texts)}")
            print(f"Total characters: {char_count}")
            
            if texts:
                print(f"\nFirst 5 text elements:")
                for i, text in enumerate(texts[:5], 1):
                    # Print safely
                    display = text[:80].replace('\n', ' ')
                    try:
                        print(f"  {i}. {display}")
                    except:
                        print(f"  {i}. [Hebrew text - {len(text)} chars]")

print("\nDone! Check largest_story.xml for full content structure.")

