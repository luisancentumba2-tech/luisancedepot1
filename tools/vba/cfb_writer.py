"""Minimal from-scratch OLE2 / MS-CFB (Compound File Binary) writer.

Only implements what's needed to build a small vbaProject.bin: a two-level
storage hierarchy (root + one sub-storage), streams routed to either the
mini-stream (<4096 bytes) or regular FAT sectors (>=4096 bytes), a single
FAT sector (fine for the small sector counts this produces), and a
balanced (not necessarily red/black-colored) binary search tree per
storage level ordered per MS-CFB 2.6.4 (name length, then case-insensitive
compare) so name-based lookups in real readers work correctly.
"""
import struct

SECTOR_SIZE = 512
MINI_SECTOR_SIZE = 64
MINI_STREAM_CUTOFF = 4096
FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF


def _name_key(name):
    return (len(name), name.upper())


def _build_bst(sids_names):
    """sids_names: list of (sid, name). Returns dict sid -> (left_sid, right_sid)."""
    ordered = sorted(sids_names, key=lambda t: _name_key(t[1]))
    tree = {}

    def build(items):
        if not items:
            return NOSTREAM
        mid = len(items) // 2
        sid, _name = items[mid]
        left = build(items[:mid])
        right = build(items[mid + 1:])
        tree[sid] = (left, right)
        return sid

    root = build(ordered)
    return tree, root


def _pad(data, size, unit):
    n = len(data)
    padded_len = ((n + unit - 1) // unit) * unit
    if padded_len == 0:
        padded_len = 0
    return data + b"\x00" * (padded_len - n)


def _dir_entry(name, entry_type, color, left, right, child, start_sector, size):
    name_utf16 = name.encode("utf-16-le") + b"\x00\x00"
    assert len(name_utf16) <= 64, name
    name_field = name_utf16 + b"\x00" * (64 - len(name_utf16))
    name_len = len(name_utf16)
    clsid = b"\x00" * 16
    state_bits = 0
    ctime = 0
    mtime = 0
    return struct.pack(
        "<64sHBBIII16sIQQIQ",
        name_field, name_len, entry_type, color,
        left, right, child,
        clsid, state_bits, ctime, mtime,
        start_sector, size,
    )


def build_compound_file(streams):
    """
    streams: list of dicts, each:
      {"name": str, "storage": "root"|"VBA", "data": bytes}
    Builds a two-level container: Root Entry directly holding "root"-storage
    streams plus one sub-storage named "VBA" holding "VBA"-storage streams.
    Returns the full file bytes.
    """
    root_streams = [s for s in streams if s["storage"] == "root"]
    vba_streams = [s for s in streams if s["storage"] == "VBA"]

    # --- classify each stream as mini (ministream) or regular (big) ---
    for s in streams:
        s["is_mini"] = len(s["data"]) < MINI_STREAM_CUTOFF

    mini_streams = [s for s in streams if s["is_mini"]]
    big_streams = [s for s in streams if not s["is_mini"]]

    # --- assign directory entry (SID) indices ---
    # 0 = Root Entry, 1 = VBA storage, then root streams, VBA streams
    entries = []  # list of dict describing each directory entry, index = sid
    entries.append({"name": "Root Entry", "type": 5})   # sid 0
    entries.append({"name": "VBA", "type": 1})           # sid 1
    sid_of = {}
    next_sid = 2
    for s in root_streams:
        sid = next_sid
        next_sid += 1
        entries.append({"name": s["name"], "type": 2, "stream": s})
        sid_of[("root", s["name"])] = sid
    for s in vba_streams:
        sid = next_sid
        next_sid += 1
        entries.append({"name": s["name"], "type": 2, "stream": s})
        sid_of[("VBA", s["name"])] = sid

    root_sid = 0
    vba_sid = 1

    # --- build BSTs ---
    top_level = [(vba_sid, "VBA")] + [(sid_of[("root", s["name"])], s["name"]) for s in root_streams]
    top_tree, top_root = _build_bst(top_level)

    vba_level = [(sid_of[("VBA", s["name"])], s["name"]) for s in vba_streams]
    vba_tree, vba_root = _build_bst(vba_level)

    # --- lay out the mini-stream (concatenation of all mini stream data, 64-byte aligned) ---
    mini_stream_data = bytearray()
    mini_start_sector_in_ministream = {}  # sid -> mini-sector index
    for s in mini_streams:
        start_minisector = len(mini_stream_data) // MINI_SECTOR_SIZE
        padded = _pad(s["data"], len(s["data"]), MINI_SECTOR_SIZE)
        mini_start_sector_in_ministream[id(s)] = start_minisector
        mini_stream_data.extend(padded)
    mini_stream_data = bytes(mini_stream_data)
    n_minisectors = len(mini_stream_data) // MINI_SECTOR_SIZE

    # --- MiniFAT: chain per mini-stream stream ---
    minifat = [FREESECT] * n_minisectors
    for s in mini_streams:
        start = mini_start_sector_in_ministream[id(s)]
        length = len(s["data"])
        count = (length + MINI_SECTOR_SIZE - 1) // MINI_SECTOR_SIZE if length > 0 else 0
        for i in range(count):
            sect = start + i
            minifat[sect] = (start + i + 1) if i < count - 1 else ENDOFCHAIN

    # --- regular-sector layout: FAT sector(s), directory sectors, minifat sectors,
    #     ministream data sectors, big-stream sectors ---
    n_dir_entries = len(entries)
    n_dir_sectors = max(1, (n_dir_entries * 4 + SECTOR_SIZE - 1) // (SECTOR_SIZE // 4)) if False else \
        max(1, -(-n_dir_entries // (SECTOR_SIZE // 128)))
    # each directory sector holds SECTOR_SIZE/128 = 4 entries
    n_dir_sectors = max(1, -(-n_dir_entries // (SECTOR_SIZE // 128)))

    n_minifat_bytes = len(minifat) * 4
    n_minifat_sectors = -(-n_minifat_bytes // SECTOR_SIZE) if n_minifat_bytes > 0 else 0

    n_ministream_sectors = -(-len(mini_stream_data) // SECTOR_SIZE) if len(mini_stream_data) > 0 else 0

    big_sector_counts = {}
    n_big_sectors_total = 0
    for s in big_streams:
        cnt = -(-len(s["data"]) // SECTOR_SIZE)
        big_sector_counts[id(s)] = cnt
        n_big_sectors_total += cnt

    # fixed point for number of FAT sectors (each covers 128 sector-entries)
    n_fat_sectors = 1
    while True:
        total_sectors = n_fat_sectors + n_dir_sectors + n_minifat_sectors + n_ministream_sectors + n_big_sectors_total
        needed = -(-total_sectors // 128)
        if needed == n_fat_sectors:
            break
        n_fat_sectors = needed
    assert n_fat_sectors <= 109, "would need DIFAT sectors, not implemented"

    # --- assign absolute sector numbers (0-based, right after the 512-byte header) ---
    cur = 0
    fat_sector_start = cur
    cur += n_fat_sectors
    dir_sector_start = cur
    cur += n_dir_sectors
    minifat_sector_start = cur
    cur += n_minifat_sectors
    ministream_sector_start = cur
    cur += n_ministream_sectors
    big_sector_start = cur
    cur += n_big_sectors_total
    total_sectors = cur

    # --- build the FAT array (one entry per regular sector) ---
    fat = [FREESECT] * total_sectors
    for i in range(n_fat_sectors):
        fat[fat_sector_start + i] = FATSECT
    for i in range(n_dir_sectors):
        sect = dir_sector_start + i
        fat[sect] = (sect + 1) if i < n_dir_sectors - 1 else ENDOFCHAIN
    for i in range(n_minifat_sectors):
        sect = minifat_sector_start + i
        fat[sect] = (sect + 1) if i < n_minifat_sectors - 1 else ENDOFCHAIN
    for i in range(n_ministream_sectors):
        sect = ministream_sector_start + i
        fat[sect] = (sect + 1) if i < n_ministream_sectors - 1 else ENDOFCHAIN

    big_start_sector_of = {}
    off = big_sector_start
    for s in big_streams:
        cnt = big_sector_counts[id(s)]
        big_start_sector_of[id(s)] = off if cnt > 0 else ENDOFCHAIN
        for i in range(cnt):
            sect = off + i
            fat[sect] = (sect + 1) if i < cnt - 1 else ENDOFCHAIN
        off += cnt

    # --- fill in per-entry start sector / size for streams ---
    for e in entries:
        if e["type"] != 2:
            continue
        s = e["stream"]
        if s["is_mini"]:
            e["start"] = mini_start_sector_in_ministream[id(s)] if len(s["data"]) > 0 else ENDOFCHAIN
            e["size"] = len(s["data"])
        else:
            e["start"] = big_start_sector_of[id(s)]
            e["size"] = len(s["data"])

    # Root Entry: start = first sector of the ministream container (in regular FAT), size = its length
    entries[root_sid]["start"] = ministream_sector_start if n_ministream_sectors > 0 else ENDOFCHAIN
    entries[root_sid]["size"] = len(mini_stream_data)
    entries[root_sid]["child"] = top_root
    entries[vba_sid]["start"] = 0
    entries[vba_sid]["size"] = 0
    entries[vba_sid]["child"] = vba_root

    for sid, (l, r) in top_tree.items():
        entries[sid]["left"] = l
        entries[sid]["right"] = r
    for sid, (l, r) in vba_tree.items():
        entries[sid]["left"] = l
        entries[sid]["right"] = r
    for e in entries:
        e.setdefault("left", NOSTREAM)
        e.setdefault("right", NOSTREAM)
        e.setdefault("child", NOSTREAM)

    # --- serialize directory sectors ---
    dir_bytes = bytearray()
    for sid, e in enumerate(entries):
        color = 1  # black; a valid (if not depth-balanced) tree
        dir_bytes.extend(_dir_entry(
            e["name"], e["type"], color, e["left"], e["right"], e["child"],
            e.get("start", 0), e.get("size", 0),
        ))
    # pad remaining slots in the last directory sector with empty (unused) entries
    n_dir_slots = n_dir_sectors * 4
    while len(entries) < n_dir_slots:
        dir_bytes.extend(_dir_entry("", 0, 0, NOSTREAM, NOSTREAM, NOSTREAM, 0, 0))
        entries.append(None)
    assert len(dir_bytes) == n_dir_sectors * SECTOR_SIZE

    # --- serialize FAT sectors ---
    fat_bytes = bytearray()
    for v in fat:
        fat_bytes.extend(struct.pack("<I", v & 0xFFFFFFFF))
    # pad FAT array to a whole number of sectors with FREESECT
    while len(fat_bytes) < n_fat_sectors * SECTOR_SIZE:
        fat_bytes.extend(struct.pack("<I", FREESECT))
    assert len(fat_bytes) == n_fat_sectors * SECTOR_SIZE

    # --- serialize MiniFAT sectors ---
    minifat_bytes = bytearray()
    for v in minifat:
        minifat_bytes.extend(struct.pack("<I", v & 0xFFFFFFFF))
    while len(minifat_bytes) < n_minifat_sectors * SECTOR_SIZE:
        minifat_bytes.extend(struct.pack("<I", FREESECT))
    assert len(minifat_bytes) == n_minifat_sectors * SECTOR_SIZE

    # --- serialize ministream data sectors ---
    ministream_bytes = _pad(mini_stream_data, len(mini_stream_data), SECTOR_SIZE)
    assert len(ministream_bytes) == n_ministream_sectors * SECTOR_SIZE

    # --- serialize big stream sectors ---
    big_bytes = bytearray()
    for s in big_streams:
        big_bytes.extend(_pad(s["data"], len(s["data"]), SECTOR_SIZE))
    assert len(big_bytes) == n_big_sectors_total * SECTOR_SIZE

    # --- header ---
    header = bytearray(512)
    header[0:8] = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
    header[8:24] = b"\x00" * 16  # CLSID
    struct.pack_into("<H", header, 24, 0x003E)  # minor version
    struct.pack_into("<H", header, 26, 0x0003)  # major version (3 = 512-byte sectors)
    struct.pack_into("<H", header, 28, 0xFFFE)  # byte order
    struct.pack_into("<H", header, 30, 0x0009)  # sector shift (2^9=512)
    struct.pack_into("<H", header, 32, 0x0006)  # mini sector shift (2^6=64)
    # bytes 34-39 reserved = 0
    struct.pack_into("<I", header, 40, 0)  # number of directory sectors (0 for v3)
    struct.pack_into("<I", header, 44, n_fat_sectors)
    struct.pack_into("<I", header, 48, dir_sector_start)  # first directory sector
    struct.pack_into("<I", header, 52, 0)  # transaction signature
    struct.pack_into("<I", header, 56, MINI_STREAM_CUTOFF)
    struct.pack_into("<I", header, 60, minifat_sector_start if n_minifat_sectors > 0 else ENDOFCHAIN)
    struct.pack_into("<I", header, 64, n_minifat_sectors)
    struct.pack_into("<I", header, 68, ENDOFCHAIN)  # first DIFAT sector (none)
    struct.pack_into("<I", header, 72, 0)  # number of DIFAT sectors
    # DIFAT array: 109 entries at offset 76, 4 bytes each
    for i in range(109):
        val = fat_sector_start + i if i < n_fat_sectors else FREESECT
        struct.pack_into("<I", header, 76 + i * 4, val)

    out = bytearray()
    out.extend(header)
    out.extend(fat_bytes)
    out.extend(dir_bytes)
    out.extend(minifat_bytes)
    out.extend(ministream_bytes)
    out.extend(big_bytes)
    return bytes(out)
