import os
import struct
import pefile

SRC = r"e:\刘相玮\科技信息化 (1)\公文一键排版助手（原版）.exe"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "app.ico")
os.makedirs(os.path.dirname(OUT), exist_ok=True)

RT_ICON = 3
RT_GROUP_ICON = 14

pe = pefile.PE(SRC)
icons = {}      # id -> bytes
groups = []     # list of (id, bytes)


def walk(entry, depth, chain):
    # entry is a ResourceEntry; directory id/name is in entry.name or entry.string
    if hasattr(entry, "directory"):
        for e in entry.directory.entries:
            walk(e, depth + 1, chain)
    elif hasattr(entry, "data"):
        # entry is a leaf resource; its parent chain gives type/id
        pass


def parse_type_dir(top):
    """top = DIRECTORY_ENTRY_RESOURCE entries (per type)."""
    for type_entry in top.entries:
        type_id = getattr(type_entry, "id", None)
        if type_id not in (RT_ICON, RT_GROUP_ICON):
            continue
        for id_entry in type_entry.directory.entries:
            res_id = getattr(id_entry, "id", None)
            for lang_entry in id_entry.directory.entries:
                data = lang_entry.data
                rva = data.struct.OffsetToData
                size = data.struct.Size
                blob = pe.get_memory_mapped_image()[rva:rva + size]
                if type_id == RT_ICON:
                    icons[res_id] = bytes(blob)
                else:
                    groups.append((res_id, bytes(blob)))


for entry in pe.DIRECTORY_ENTRY_RESOURCE.entries:
    if getattr(entry, "id", None) in (RT_ICON, RT_GROUP_ICON):
        for id_entry in entry.directory.entries:
            res_id = getattr(id_entry, "id", None)
            for lang_entry in id_entry.directory.entries:
                d = lang_entry.data
                blob = bytes(pe.get_memory_mapped_image()[d.struct.OffsetToData:d.struct.OffsetToData + d.struct.Size])
                if entry.id == RT_ICON:
                    icons[res_id] = blob
                else:
                    groups.append((res_id, blob))

if not groups:
    print("NO_ICON_GROUP")
    raise SystemExit(1)

groups.sort(key=lambda x: x[0])
grp = groups[0][1]
# GRPICONDIR: <HHH> reserved,type,count
_r, _t, count = struct.unpack("<HHH", grp[:6])
entries = []
off = 6
for i in range(count):
    w, h, ncol, res, planes, bitcnt, nbytes, iid = struct.unpack("<BBBBHHIH", grp[off:off + 14])
    off += 14
    entries.append([w, h, ncol, res, planes, bitcnt, nbytes, iid])

data_dir = 6 + 16 * len(entries)
ico = struct.pack("<HHH", 0, 1, len(entries))
body = b""
cursor = data_dir
for w, h, ncol, res, planes, bitcnt, nbytes, iid in entries:
    img = icons.get(iid, b"")
    ico += struct.pack("<BBBBHHII", w, h, ncol, res, planes, bitcnt, len(img), cursor)
    cursor += len(img)
    body += img

with open(OUT, "wb") as f:
    f.write(ico + body)

# 校验尺寸
from PIL import Image
probe = Image.open(OUT)
sizes = sorted(probe.ico.sizes())
print("ICO_SAVED", OUT, os.path.getsize(OUT), "bytes")
print("ICONS_IN_EXE", len(icons), "GROUPS", len(groups), "SIZES", sizes)
