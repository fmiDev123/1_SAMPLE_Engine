"""Compare two .pptx files part by part. Usage: cmp_pptx.py A B
Reports parts only in one, parts whose bytes differ (docProps dates ignored)."""
import sys, zipfile, re, hashlib
a, b = zipfile.ZipFile(sys.argv[1]), zipfile.ZipFile(sys.argv[2])
na, nb = set(a.namelist()), set(b.namelist())
def norm(name, data):
    if name == "docProps/core.xml":
        data = re.sub(rb"<dcterms:(created|modified)[^>]*>[^<]*</dcterms:\1>", b"", data)
    return data
only_a, only_b = sorted(na - nb), sorted(nb - na)
diff = [n for n in sorted(na & nb) if norm(n, a.read(n)) != norm(n, b.read(n))]
print(f"parts: A={len(na)} B={len(nb)} | only A: {len(only_a)} | only B: {len(only_b)} | differing: {len(diff)}")
for n in only_a[:10]: print("  only A:", n)
for n in only_b[:10]: print("  only B:", n)
for n in diff[:15]: print("  differs:", n, len(a.read(n)), len(b.read(n)))
sys.exit(0 if not (only_a or only_b or diff) else 1)
