"""Minimal MS-OVBA (2.4.1) compressor: literal-only tokens (no back-references).
Produces spec-valid CompressedContainer bytes. Round-trip verified against
oletools.olevba.decompress_stream (the reference decompressor)."""


# A literal-only "compressed" chunk has overhead of 1 flag byte per 8 literal
# bytes, so its encoded chunk_data = N + ceil(N/8), which must fit in the
# format's 12-bit size field (max chunk_data = 4096 bytes). N=3640 is the
# largest chunk size that stays safely under that cap for every source chunk.
MAX_CHUNK = 3640


def compress_chunk_literal(data: bytes) -> bytes:
    assert len(data) <= MAX_CHUNK
    out = bytearray()
    i = 0
    n = len(data)
    while i < n:
        group = data[i:i + 8]
        out.append(0x00)  # flag byte: all 8 (or fewer) tokens are literal bytes
        out.extend(group)
        i += 8
    chunk_data = bytes(out)
    chunk_size_field = (2 + len(chunk_data)) - 3
    assert 0 <= chunk_size_field <= 0xFFF, chunk_size_field
    header = chunk_size_field | (0b011 << 12) | (1 << 15)
    return header.to_bytes(2, "little") + chunk_data


def compress_stream(data: bytes) -> bytes:
    out = bytearray()
    out.append(0x01)  # CompressedContainer signature
    i = 0
    n = len(data)
    while i < n:
        out.extend(compress_chunk_literal(data[i:i + MAX_CHUNK]))
        i += MAX_CHUNK
    return bytes(out)


if __name__ == "__main__":
    # Self-test: round-trip through oletools' reference decompressor.
    # Requires `pip install oletools` (not a runtime dependency of build_xlsm.py).
    from oletools.olevba import decompress_stream

    tests = [
        b"",
        b"hello",
        b"Attribute VB_Name = \"Module1\"\r\nSub Test()\r\nEnd Sub\r\n",
        ("Attribute VB_Name = \"Module1\"\r\n"
         "Sub CloturerLeMois()\r\n"
         "    Dim x As Long\r\n"
         "    ' commentaire avec accents : éèàôûçÉ\r\n"
         "End Sub\r\n").encode("cp1252"),
        (b"A" * 4096),
        (b"B" * 4097),
        (b"C" * 8192),
        (b"D" * 9000),
        bytes(range(256)) * 20,
    ]
    for t in tests:
        c = compress_stream(t)
        d = bytes(decompress_stream(bytearray(c)))
        status = "OK" if d == t else "MISMATCH"
        print(status, "len(in)=", len(t), "len(compressed)=", len(c), "len(out)=", len(d))
        if d != t:
            print("  expected:", t[:80])
            print("  got     :", d[:80])
