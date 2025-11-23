import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from collections import defaultdict

idml_path = Path(r'backup_input\שס 2\ביצה\PEREK1.idml')

# Data structures to store the mapping
page_to_stories = defaultdict(list)  # page_name -> list of story IDs
story_to_pages = defaultdict(list)   # story_id -> list of page names

with zipfile.ZipFile(idml_path, 'r') as z:
    # Get all spreads
    spread_files = [f for f in z.namelist() if f.startswith('Spreads/') and f.endswith('.xml')]
    
    print(f"Analyzing {len(spread_files)} spreads...")
    print("=" * 80)
    
    for spread_file in sorted(spread_files):
        with z.open(spread_file) as f:
            tree = ET.parse(f)
            root = tree.getroot()
            
            # Find all Page elements and their names
            for page_elem in root.iter():
                tag = page_elem.tag.split('}')[-1] if '}' in page_elem.tag else page_elem.tag
                
                if tag == 'Page':
                    page_name = page_elem.attrib.get('Name', 'Unknown')
                    
                    # Find all TextFrame elements on this page (they are siblings/children)
                    # In the spread, after a Page element, the TextFrames follow
                    parent = root
                    for text_frame in parent.iter():
                        frame_tag = text_frame.tag.split('}')[-1] if '}' in text_frame.tag else text_frame.tag
                        
                        if frame_tag == 'TextFrame':
                            story_id = text_frame.attrib.get('ParentStory', '')
                            if story_id:
                                page_to_stories[page_name].append(story_id)
                                story_to_pages[story_id].append(page_name)
    
    with open('page_mapping.txt', 'w', encoding='utf-8') as out:
        out.write("\nPage to Story Mapping:\n")
        out.write("=" * 80 + "\n")
        for page_name in sorted(page_to_stories.keys()):
            stories = set(page_to_stories[page_name])
            out.write(f"\nPage: {page_name}\n")
            for story_id in stories:
                # Get story info
                story_file = f'Stories/Story_{story_id}.xml'
                try:
                    with z.open(story_file) as f:
                        content = f.read().decode('utf-8')
                        # Extract short preview
                        tree_story = ET.fromstring(content.encode('utf-8'))
                        texts = []
                        for elem in tree_story.iter():
                            t = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
                            if t == 'Content' and elem.text:
                                texts.append(elem.text.strip())
                        preview = ' '.join(texts)[:80]
                        info = z.getinfo(story_file)
                        out.write(f"  Story {story_id} ({info.file_size} bytes): {preview}\n")
                except:
                    out.write(f"  Story {story_id}: [Could not read]\n")

print("Done! Output written to page_mapping.txt")

