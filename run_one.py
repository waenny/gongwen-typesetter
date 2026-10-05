import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gwps_engine as E

inp, outp = sys.argv[1], sys.argv[2]
proc = E.WordProcessor(config={}, log=lambda m: None)
rep = proc.format_document(inp, outp)
print("TITLE=", proc.last_title)
print("FNAME=", E.build_filename(proc.last_title))
for n, ok, d in rep:
    print(("PASS" if ok else "需核"), n)
