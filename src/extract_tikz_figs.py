"""Extract the architecture diagram (page 4) and flowchart (page 6) from
main.pdf and crop them to just the figure region, saving as PNG.

This avoids running standalone TikZ compiles -- it just reuses what
main.pdf already rendered.
"""
import os
import subprocess
import sys

from PIL import Image

REPORT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                          "report")
FIG_DIR = os.path.join(REPORT_DIR, "figs")
# the report PDF may have been renamed by the team -- pick the first match
_CANDIDATES = ["main.pdf", "Research_Paper.pdf", "AO_PSO_Paper.pdf"]
PDF = next((os.path.join(REPORT_DIR, c) for c in _CANDIDATES
             if os.path.exists(os.path.join(REPORT_DIR, c))),
           os.path.join(REPORT_DIR, "main.pdf"))

DPI = 200


def render(page: int, tag: str) -> str:
    out_prefix = os.path.join(FIG_DIR, f"_tmp_{tag}")
    subprocess.run(
        ["pdftoppm", "-r", str(DPI), "-png", "-f", str(page), "-l", str(page),
         PDF, out_prefix],
        check=True,
    )
    return f"{out_prefix}-{page}.png"


def crop_and_save(src: str, box: tuple, dest: str) -> None:
    im = Image.open(src)
    cropped = im.crop(box)
    cropped.save(dest, "PNG")
    print(f"wrote {dest} ({cropped.size[0]}x{cropped.size[1]})")


def main() -> None:
    arch_src = render(4, "arch")
    flow_src = render(6, "flow")

    arch_im = Image.open(arch_src)
    flow_im = Image.open(flow_src)
    print(f"page 4 size: {arch_im.size}")
    print(f"page 6 size: {flow_im.size}")

    # page 4: architecture is the top portion (spans both columns)
    # IEEE page is ~1700 wide at 200 dpi. Figure runs from y~150 to y~750
    aw, ah = arch_im.size
    arch_box = (int(aw * 0.04), int(ah * 0.05),
                int(aw * 0.96), int(ah * 0.36))
    crop_and_save(arch_src, arch_box,
                  os.path.join(FIG_DIR, "architecture.png"))

    # page 6: flowchart is the LEFT column (the algorithm is on the right)
    fw, fh = flow_im.size
    flow_box = (int(fw * 0.06), int(fh * 0.05),
                int(fw * 0.48), int(fh * 0.605))
    crop_and_save(flow_src, flow_box,
                  os.path.join(FIG_DIR, "flowchart.png"))

    # algorithm box from page 6, right column
    algo_box = (int(fw * 0.50), int(fh * 0.05),
                int(fw * 0.96), int(fh * 0.50))
    crop_and_save(flow_src, algo_box,
                  os.path.join(FIG_DIR, "algorithm.png"))

    for tmp in (arch_src, flow_src):
        try:
            os.remove(tmp)
        except OSError:
            pass


if __name__ == "__main__":
    main()
