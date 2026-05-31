import markdown_it, re, collections, os

md_path = r'D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\patent_md\完整专利申请书_20260531_1235_v2.1.md'
with open(md_path, 'r', encoding='utf-8') as f:
    text = f.read()

md = markdown_it.MarkdownIt('commonmark', {'breaks': False})
tokens = md.parse(text)

# Token types
tc = collections.Counter()
hc = collections.Counter()
for t in tokens:
    tc[t.type] += 1
    if t.type == 'heading_open':
        hc[t.tag] += 1

print("=" * 60)
print("AST TOKEN DISTRIBUTION")
print("=" * 60)
for k, v in tc.most_common(20):
    print(f"  {k:25s} {v:5d}")

print("\nHEADING LEVELS:")
for tag, cnt in sorted(hc.items()):
    print(f"  {tag}: {cnt}")

# Paragraph classification
numbered = claim = step = normal = 0
for i, t in enumerate(tokens):
    if t.type == 'inline' and t.level == 1:
        raw = t.content.strip()
        if re.match(r'^\[(\d{4})\]\s', raw):
            numbered += 1
        elif re.match(r'^\d+\.\s', raw):
            claim += 1
        elif re.match(r'^S\d+\.\s', raw):
            step += 1
        elif raw:
            normal += 1

tables = sum(1 for t in tokens if t.type == 'table_open')
hrs = sum(1 for t in tokens if t.type == 'hr')
imgs = sum(1 for t in tokens if t.type == 'image')

# Formula counts
inline_f = len(re.findall(r'(?<!\$)\$[^$]+\$(?!\$)', text))
block_f = len(re.findall(r'\$\$[^$]+\$\$', text))
standalone = sum(1 for line in text.split('\n') 
                 if re.match(r'^\$[^$]+\$$', line.strip()) and len(line.strip()) > 2)

print("\nPARAGRAPH CLASSIFICATION:")
print(f"  numbered_para [NNNN]: {numbered}")
print(f"  claim   (N. ...):     {claim}")
print(f"  step    (S1. ...):    {step}")
print(f"  normal:               {normal}")

print("\nSTRUCTURAL ELEMENTS:")
print(f"  tables:               {tables}")
print(f"  hr (---):             {hrs}")
print(f"  images:               {imgs}")

print("\nFORMULAS:")
print(f"  inline $...$:         {inline_f}")
print(f"  block $$...$$:        {block_f}")
print(f"  standalone $...$:     {standalone}")

# H1/H2/H3 titles
print("\nHEADING TITLES:")
for i, t in enumerate(tokens):
    if t.type == 'heading_open':
        title = tokens[i + 1].content.strip()[:60]
        print(f"  {t.tag}: {title}")
