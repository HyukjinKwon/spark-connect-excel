# SPDX-License-Identifier: Apache-2.0
#
# gen_icons.py — Generate solid-color placeholder PNGs for the Office Add-in
# ribbon icons.  Uses only Python stdlib (zlib + struct); no Pillow required.
#
# Spark orange: #E25A1C  (R=226, G=90, B=28)
# Run from the repo root:  python assets/gen_icons.py

import struct
import zlib
import os

SPARK_ORANGE = (226, 90, 28)  # #E25A1C


def make_png(size: int, rgb: tuple[int, int, int]) -> bytes:
    """Return bytes for a minimal valid PNG: <size>x<size> solid RGB image."""
    r, g, b = rgb

    # --- PNG signature ---
    sig = b"\x89PNG\r\n\x1a\n"

    # --- IHDR chunk: width, height, bit-depth=8, color-type=2 (RGB), etc. ---
    ihdr_data = struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0)
    ihdr = _chunk(b"IHDR", ihdr_data)

    # --- IDAT chunk: raw image data (filter byte 0 + RGB triplets per row) ---
    row = b"\x00" + bytes([r, g, b] * size)   # filter=None (0), then RGB pixels
    raw = row * size                            # repeat for every row
    compressed = zlib.compress(raw, 9)
    idat = _chunk(b"IDAT", compressed)

    # --- IEND chunk ---
    iend = _chunk(b"IEND", b"")

    return sig + ihdr + idat + iend


def _chunk(tag: bytes, data: bytes) -> bytes:
    """Wrap raw data in a PNG chunk (length + tag + data + CRC)."""
    length = struct.pack(">I", len(data))
    crc = struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
    return length + tag + data + crc


def main() -> None:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    sizes = [16, 32, 64, 80]
    for size in sizes:
        path = os.path.join(script_dir, f"icon-{size}.png")
        data = make_png(size, SPARK_ORANGE)
        with open(path, "wb") as f:
            f.write(data)
        # Quick sanity: verify PNG signature
        assert data[:8] == b"\x89PNG\r\n\x1a\n", f"Bad PNG signature for {path}"
        print(f"  wrote {path}  ({len(data)} bytes, {size}x{size} px)")
    print("Done — all icons verified (PNG signature OK).")


if __name__ == "__main__":
    main()
