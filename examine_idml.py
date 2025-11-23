import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

idml_path = Path(r'backup_input\שס 2\ביצה\PEREK1.idml')

print("=" * 80)
print("IDML File Structure:")
print("=" * 80)

with zipfile.ZipFile(idml_path, 'r') as z:
    print("\nFiles in IDML:")
    for name in z.namelist():
        print(f"  {name}")
    
    print("\n" + "=" * 80)
    print("Examining Story Files:")
    print("=" * 80)
    
    story_files = [name for name in z.namelist() if name.startswith('Stories/') and name.endswith('.xml')]
    
    for story_file in sorted(story_files)[:5]:  # Look at first 5 story files
        print(f"\n{'=' * 60}")
        print(f"Story: {story_file}")
        print(f"{'=' * 60}")
        
        with z.open(story_file) as f:
            content = f.read().decode('utf-8')
            
            # Save first story file for examination
            if story_file == sorted(story_files)[0]:
                with open('sample_story.xml', 'w', encoding='utf-8') as out:
                    out.write(content)
                print(f"\nSaved {story_file} to sample_story.xml for examination")
            
            # Look for footnote references
            if 'Footnote' in content or 'Note' in content or 'MarkerType' in content:
                print(f"\n*** CONTAINS FOOTNOTE/MARKER ELEMENTS in {story_file} ***")
                
            # Parse and examine structure
            tree = ET.fromstring(content.encode('utf-8'))
            
            # Look for all unique element tags
            tags = set()
            text_samples = []
            for elem in tree.iter():
                tag_name = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
                tags.add(tag_name)
                
                # Collect text samples
                if tag_name == 'Content' and elem.text and len(elem.text) > 3:
                    text_samples.append(elem.text[:50])
            
            print(f"\nElement types found: {', '.join(sorted(tags))}")
            if text_samples:
                print(f"Sample text (first 3): {text_samples[:3]}")

