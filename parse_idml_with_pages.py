import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from collections import defaultdict, OrderedDict

idml_path = Path(r'backup_input\שס 2\ביצה\PEREK1.idml')

# Data structures
frame_to_page = {}  # frame_id -> page_name
page_markers = {}   # page_name -> marker text (like "ביצה ב:")
frame_chain = {}    # frame_id -> next_frame_id
main_story_id = None

log = []

with zipfile.ZipFile(idml_path, 'r') as z:
    log.append("Step 1: Finding main story (largest)...")
    story_files = [f for f in z.namelist() if f.startswith('Stories/') and f.endswith('.xml')]
    story_sizes = [(f, z.getinfo(f).file_size) for f in story_files]
    story_sizes.sort(key=lambda x: x[1], reverse=True)
    main_story_file = story_sizes[0][0]
    main_story_id = main_story_file.split('_')[1].replace('.xml', '')
    log.append(f"  Main story: {main_story_id} ({story_sizes[0][1]} bytes)")
    
    log.append("\nStep 2: Mapping pages and frames from spreads...")
    spread_files = [f for f in z.namelist() if f.startswith('Spreads/') and f.endswith('.xml')]
    
    for spread_file in sorted(spread_files):
        with z.open(spread_file) as f:
            tree = ET.parse(f)
            root = tree.getroot()
            
            current_page = None
            
            # Process all elements in order
            for elem in root.iter():
                tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
                
                if tag == 'Page':
                    current_page = elem.attrib.get('Name', 'Unknown')
                    log.append(f"  Found page: {current_page}")
                
                elif tag == 'TextFrame' and current_page:
                    frame_id = elem.attrib.get('Self', '')
                    parent_story = elem.attrib.get('ParentStory', '')
                    next_frame = elem.attrib.get('NextTextFrame', 'n')
                    prev_frame = elem.attrib.get('PreviousTextFrame', 'n')
                    
                    if frame_id:
                        frame_to_page[frame_id] = current_page
                        
                        if next_frame != 'n':
                            frame_chain[frame_id] = next_frame
                        
                        # Check if this is a page marker (small story on this page)
                        if parent_story != main_story_id and parent_story:
                            # Extract text from this story
                            try:
                                story_file = f'Stories/Story_{parent_story}.xml'
                                with z.open(story_file) as sf:
                                    story_content = sf.read().decode('utf-8')
                                    story_tree = ET.fromstring(story_content.encode('utf-8'))
                                    texts = []
                                    for se in story_tree.iter():
                                        st = se.tag.split('}')[-1] if '}' in se.tag else se.tag
                                        if st == 'Content' and se.text:
                                            texts.append(se.text.strip())
                                    marker_text = ' '.join(texts).strip()
                                    if marker_text and 'ביצה' in marker_text and ':' in marker_text:
                                        # This is a page reference marker
                                        if current_page not in page_markers:
                                            page_markers[current_page] = marker_text
                                            log.append(f"    Page marker: {marker_text}")
                            except:
                                pass
    
    log.append(f"\nStep 3: Extracting main story content...")
    with z.open(main_story_file) as f:
        tree = ET.parse(f)
        root = tree.getroot()
        
        # Find all paragraphs in document order
        paragraphs = []
        for para_elem in root.iter():
            tag_name = para_elem.tag.split('}')[-1] if '}' in para_elem.tag else para_elem.tag
            
            if tag_name == 'ParagraphStyleRange':
                # Collect all text within this paragraph
                para_texts = []
                for content_elem in para_elem.iter():
                    content_tag = content_elem.tag.split('}')[-1] if '}' in content_elem.tag else content_elem.tag
                    if content_tag == 'Content' and content_elem.text:
                        para_texts.append(content_elem.text)
                
                if para_texts:
                    para_text = ''.join(para_texts).strip()
                    para_text = para_text.replace('&apos;', "'").replace('&quot;', '"')
                    para_text = para_text.replace('\ufeff', '')
                    para_text = ' '.join(para_text.split())
                    
                    if para_text and para_text != "0" and len(para_text) > 1:
                        paragraphs.append(para_text)
        
        log.append(f"  Extracted {len(paragraphs)} paragraphs")
    
    log.append(f"\nStep 4: Mapping paragraphs to pages...")
    # Get pages that have the main story
    pages_with_main_story = sorted(set([pg for fid, pg in frame_to_page.items()]))
    
    log.append(f"  Story appears on {len(pages_with_main_story)} pages")
    
    # Distribute paragraphs evenly across pages (approximation)
    paras_per_page = len(paragraphs) // len(pages_with_main_story) if pages_with_main_story else len(paragraphs)
    
    log.append(f"\nStep 5: Creating output with page markers...")
    with open('idml_with_page_markers.txt', 'w', encoding='utf-8') as out:
        out.write("IDML Content with Page Markers\n")
        out.write("=" * 80 + "\n\n")
        
        para_idx = 0
        for page_idx, page_name in enumerate(pages_with_main_story):
            # Insert page marker
            marker = page_markers.get(page_name, '')
            if marker:
                out.write(f"\n{'=' * 60}\n")
                out.write(f"PAGE MARKER: {marker}\n")
                out.write(f"Page: {page_name}\n")
                out.write(f"{'=' * 60}\n\n")
            
            # Write paragraphs for this page
            end_idx = min(para_idx + paras_per_page, len(paragraphs))
            if page_idx == len(pages_with_main_story) - 1:
                end_idx = len(paragraphs)  # Last page gets remaining paras
            
            for i in range(para_idx, end_idx):
                out.write(f"{paragraphs[i]}\n\n")
            
            para_idx = end_idx

# Write log
with open('parse_log.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(log))
    f.write(f"\n\nSummary:\n")
    f.write(f"  - Main story: {main_story_id}\n")
    f.write(f"  - Total paragraphs: {len(paragraphs)}\n")
    f.write(f"  - Pages with content: {len(pages_with_main_story)}\n")
    f.write(f"  - Page markers found: {len(page_markers)}\n")

print("Done! Check parse_log.txt and idml_with_page_markers.txt")

