import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

idml_path = Path(r'backup_input\שס 2\ביצה\PEREK1.idml')

with zipfile.ZipFile(idml_path, 'r') as z:
    story_files = [name for name in z.namelist() if name.startswith('Stories/') and name.endswith('.xml')]
    
    # Get story sizes
    story_info = []
    for story_file in story_files:
        info = z.getinfo(story_file)
        story_info.append((story_file, info.file_size))
    
    # Sort by size
    story_info.sort(key=lambda x: x[1], reverse=True)
    
    with open('small_stories_analysis.txt', 'w', encoding='utf-8') as out:
        out.write("Analysis of all story files in IDML\n")
        out.write("=" * 80 + "\n\n")
        
        out.write(f"Total story files: {len(story_info)}\n\n")
        
        # Show largest first
        out.write("LARGEST STORY (Main Content):\n")
        out.write("-" * 80 + "\n")
        largest = story_info[0]
        out.write(f"{largest[0]} - {largest[1]} bytes\n\n")
        
        # Now examine smaller stories
        out.write("\nSMALLER STORY FILES (Non-Main Content):\n")
        out.write("=" * 80 + "\n\n")
        
        for story_file, size in story_info[1:20]:  # Look at next 20 files
            out.write(f"\n{'=' * 60}\n")
            out.write(f"Story: {story_file}\n")
            out.write(f"Size: {size} bytes\n")
            out.write(f"{'=' * 60}\n")
            
            with z.open(story_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                
                # Extract all text
                all_text = []
                for elem in root.iter():
                    tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
                    if tag == 'Content' and elem.text:
                        all_text.append(elem.text)
                
                combined_text = ' '.join(all_text).strip()
                
                out.write(f"Text content ({len(combined_text)} chars):\n")
                if combined_text:
                    # Show first 200 chars
                    display = combined_text[:200] if len(combined_text) > 200 else combined_text
                    out.write(display)
                    if len(combined_text) > 200:
                        out.write(f"\n... (+ {len(combined_text) - 200} more chars)")
                else:
                    out.write("[EMPTY OR NO TEXT CONTENT]")
                out.write("\n\n")
                
                # Check for specific patterns
                if '[' in combined_text or ']' in combined_text:
                    out.write("Contains brackets: YES\n")
                if any(x in combined_text for x in ['ע"א', 'ע"ב', "ע'א", "ע'ב"]):
                    out.write("Contains page references (ע\"א/ע\"ב): YES\n")
                if combined_text.startswith('[') or (len(combined_text) < 30 and '[' in combined_text):
                    out.write("Type: Likely PAGE REFERENCE or CITATION\n")
                    
            out.write("\n")

print("Analysis written to small_stories_analysis.txt")

