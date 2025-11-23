import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

idml_path = Path(r'backup_input\שס 2\ביצה\PEREK1.idml')

with zipfile.ZipFile(idml_path, 'r') as z:
    # Get main story
    story_files = [f for f in z.namelist() if f.startswith('Stories/') and f.endswith('.xml')]
    story_sizes = [(f, z.getinfo(f).file_size) for f in story_files]
    main_story = max(story_sizes, key=lambda x: x[1])[0]
    
    print("=" * 80)
    print("STORY ANALYSIS - Looking for frame references")
    print("=" * 80)
    
    with z.open(main_story) as f:
        tree = ET.parse(f)
        root = tree.getroot()
        
        # Check Story element attributes
        print(f"\nStory root tag: {root.tag}")
        print(f"Story root attributes: {root.attrib}")
        
        # Check first few paragraph elements
        print("\n--- First 3 ParagraphStyleRange elements ---")
        count = 0
        for elem in root.iter():
            tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
            if tag == 'ParagraphStyleRange' and count < 3:
                print(f"\nParagraph {count}:")
                print(f"  Tag: {elem.tag}")
                print(f"  Attributes: {elem.attrib}")
                
                # Check children
                for child in elem:
                    child_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    print(f"  Child: {child_tag}, Attribs: {child.attrib}")
                
                count += 1
    
    print("\n" + "=" * 80)
    print("SPREAD/TEXTFRAME ANALYSIS - Looking for content references")
    print("=" * 80)
    
    spread_files = [f for f in z.namelist() if f.startswith('Spreads/') and f.endswith('.xml')]
    
    # Check first spread
    with z.open(spread_files[0]) as f:
        tree = ET.parse(f)
        root = tree.getroot()
        
        print("\n--- First TextFrame element ---")
        for elem in root.iter():
            tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
            if tag == 'TextFrame':
                print(f"\nTextFrame attributes: {elem.attrib}")
                
                # Look for any child elements that might reference content
                print("TextFrame children:")
                for child in elem:
                    child_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    print(f"  {child_tag}: {child.attrib}")
                
                break  # Just check first one
    
    print("\n" + "=" * 80)
    print("LOOKING FOR CHARACTER RANGES OR CONTENT REFERENCES")
    print("=" * 80)
    
    # Search for any elements with "character", "index", "range", "offset" in attributes
    print("\nSearching all Spread XMLs for content reference attributes...")
    keywords = ['character', 'index', 'range', 'offset', 'start', 'end', 'length']
    
    for spread_file in spread_files[:2]:  # Check first 2 spreads
        with z.open(spread_file) as f:
            content = f.read().decode('utf-8')
            
            for keyword in keywords:
                if keyword.lower() in content.lower():
                    print(f"\n  Found '{keyword}' in {spread_file}")
                    # Show a snippet
                    idx = content.lower().find(keyword.lower())
                    snippet = content[max(0, idx-50):min(len(content), idx+100)]
                    print(f"    Context: ...{snippet}...")
                    break

print("\n" + "=" * 80)
print("Investigation complete. Check output above for mapping clues.")
print("=" * 80)

