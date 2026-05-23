import os
import sys

# add path to minor_project to allow imports if needed
sys.path.append(r"c:\Users\manoj\OneDrive\Desktop\Minor\minor_project")

from app import markdown_to_pptx_bytes

md = """# Title
## Welcome
- Bullet 1
- Bullet 2

### Sub-bullet
Some interesting text!
"""

try:
    pptx_bytes = markdown_to_pptx_bytes(md, title="Test Presentation")
    print("PPTX generation successful!")
    print(f"Bytes size: {len(pptx_bytes.getvalue())}")
except Exception as e:
    print(f"PPTX Error: {e}")
    import traceback
    traceback.print_exc()
