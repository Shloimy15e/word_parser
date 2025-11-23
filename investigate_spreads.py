import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

idml_path = Path(r'backup_input\שס 2\ביצה\PEREK1.idml')

with zipfile.ZipFile(idml_path, 'r') as z:
    # List all files
    all_files = z.namelist()
    
    print("Files in IDML:")
    print("=" * 80)
    
    # Find spread files
    spread_files = [f for f in all_files if f.startswith('Spreads/')]
    print(f"\nSpread files: {len(spread_files)}")
    for f in spread_files:
        print(f"  {f}")
    
    # Examine first spread
    if spread_files:
        print("\n" + "=" * 80)
        print("Examining first spread file:")
        print("=" * 80)
        
        with z.open(spread_files[0]) as f:
            content = f.read().decode('utf-8')
            
            # Save to file for examination
            with open('spread_sample.xml', 'w', encoding='utf-8') as out:
                out.write(content)
            
            print(f"\nSaved {spread_files[0]} to spread_sample.xml")
            
            # Parse and look for story references
            tree = ET.fromstring(content.encode('utf-8'))
            
            # Find all elements and their types
            elements = {}
            for elem in tree.iter():
                tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
                if tag not in elements:
                    elements[tag] = 0
                elements[tag] += 1
            
            print(f"\nElement types found:")
            for tag, count in sorted(elements.items()):
                print(f"  {tag}: {count}")
            
            # Look for story references
            print(f"\nLooking for story references...")
            story_refs = []
            for elem in tree.iter():
                # Check ParentStory attributes
                if 'ParentStory' in elem.attrib:
                    story_id = elem.attrib['ParentStory']
                    story_refs.append(story_id)
            
            if story_refs:
                print(f"Found {len(story_refs)} story references:")
                for ref in sorted(set(story_refs))[:10]:
                    print(f"  Story ID: {ref}")

print("\nDone! Check spread_sample.xml for detailed structure")

