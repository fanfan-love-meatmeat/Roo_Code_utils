"""Verify ordered list / claim structure in new MD"""
import markdown_it, re

md_path = r'D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\patent_md\完整专利申请书_20260531_1235_v2.1.md'
with open(md_path, 'r', encoding='utf-8') as f:
    text = f.read()

md = markdown_it.MarkdownIt('commonmark', {'breaks': False})
tokens = md.parse(text)

# Find ordered_list sections and show their structure
in_ordered_list = False
in_list_item = False
list_depth = 0
list_content = []

for i, t in enumerate(tokens):
    if t.type == 'ordered_list_open':
        in_ordered_list = True
        list_depth += 1
        print(f"\n=== ORDERED LIST START (depth={list_depth}) ===")
        list_content = []
    elif t.type == 'ordered_list_close':
        print(f"\n=== LIST ITEMS: {len([x for x in list_content if x == 'ITEM_START'])} ===")
        list_depth -= 1
        if list_depth == 0:
            in_ordered_list = False
            print(f"\nSample claims:")
            for item in list_content[:3]:
                print(f"  {item[:80]}")
    elif t.type == 'list_item_open':
        in_list_item = True
        list_content.append(f"ITEM_START")
    elif t.type == 'list_item_close':
        in_list_item = False
    elif in_list_item and t.type == 'inline':
        content = t.content.strip()[:100]
        list_content.append(content)

print("\n\n=== PARAGRAPH-LEVEL CLASSIFICATION (level=1 only) ===")
for i, t in enumerate(tokens):
    if t.type == 'inline' and t.level == 1:
        raw = t.content.strip()
        if re.match(r'^\d+\.\s', raw) and len(raw) < 200:
            print(f"  TOP-LEVEL CLAIM-LIKE: {raw[:80]}")
