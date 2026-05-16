"""Generate a pptx with text shadow / text outline / RTL paragraphs for VRT.

python-pptx doesn't expose rPr > effectLst / rPr > a:ln / pPr@rtl directly,
so we hand-assemble the run XML by reaching into the lxml element tree.

Run: python3 _gen_text_effects.py
Output: text-effects.pptx (gitignored — public/private/ is local-only).
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from lxml import etree

OUT = "text-effects.pptx"

NSMAP = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


def make_run(text, *, size_pt=40, bold=False, color="000000",
             shadow=False, outline_w=None, outline_color=None):
    """Build an <a:r> element with rPr-level effectLst / a:ln if requested."""
    r = etree.SubElement(etree.Element(qn("a:p")), qn("a:r"))
    rPr = etree.SubElement(r, qn("a:rPr"))
    rPr.set("lang", "en-US")
    rPr.set("sz", str(int(size_pt * 100)))
    if bold:
        rPr.set("b", "1")
    if outline_w is not None:
        # <a:ln w="EMU"><a:solidFill><a:srgbClr val="HEX"/></a:solidFill></a:ln>
        ln = etree.SubElement(rPr, qn("a:ln"))
        ln.set("w", str(int(outline_w)))
        sf = etree.SubElement(ln, qn("a:solidFill"))
        c = etree.SubElement(sf, qn("a:srgbClr"))
        c.set("val", outline_color or "000000")
    # text fill
    fill = etree.SubElement(rPr, qn("a:solidFill"))
    c = etree.SubElement(fill, qn("a:srgbClr"))
    c.set("val", color)
    # shadow
    if shadow:
        eff = etree.SubElement(rPr, qn("a:effectLst"))
        shdw = etree.SubElement(eff, qn("a:outerShdw"))
        shdw.set("blurRad", "50800")  # 4pt blur
        shdw.set("dist", "38100")     # 3pt offset
        shdw.set("dir", "2700000")    # 45° (degrees * 60000)
        shdw.set("algn", "tl")
        scol = etree.SubElement(shdw, qn("a:srgbClr"))
        scol.set("val", "808080")
        alpha = etree.SubElement(scol, qn("a:alpha"))
        alpha.set("val", "65000")
    t = etree.SubElement(r, qn("a:t"))
    t.text = text
    return r


def make_paragraph(*, algn=None, rtl=False):
    p = etree.Element(qn("a:p"))
    pPr = etree.SubElement(p, qn("a:pPr"))
    if algn:
        pPr.set("algn", algn)
    if rtl:
        pPr.set("rtl", "1")
    return p


def add_text(slide, *, left, top, width, height, paragraphs):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tx = tb.text_frame
    # Clear the default paragraph python-pptx inserts.
    txBody = tx._txBody
    for p in list(txBody.findall(qn("a:p"))):
        txBody.remove(p)
    for p in paragraphs:
        txBody.append(p)


prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]

# ---------------------------------------------------------------------------
# Slide 1 — text shadow + text outline samples
# ---------------------------------------------------------------------------
s1 = prs.slides.add_slide(blank)

title_p = make_paragraph()
title_p.append(make_run("Run-level text effects", size_pt=28, bold=True))
add_text(s1, left=Inches(0.4), top=Inches(0.2), width=Inches(12.5), height=Inches(0.8), paragraphs=[title_p])

# Shadow demo
shadow_p = make_paragraph()
shadow_p.append(make_run("Text with shadow", size_pt=48, bold=True, color="1F4E79", shadow=True))
add_text(s1, left=Inches(0.5), top=Inches(1.4), width=Inches(12), height=Inches(1.2), paragraphs=[shadow_p])

# Outline-only demo (white fill, blue outline)
outline_p = make_paragraph()
outline_p.append(make_run("Text with outline", size_pt=64, bold=True,
                          color="FFFFFF",
                          outline_w=19050, outline_color="C00000"))  # 1.5pt outline
add_text(s1, left=Inches(0.5), top=Inches(3.2), width=Inches(12), height=Inches(1.6), paragraphs=[outline_p])

# Combined shadow + outline
combined_p = make_paragraph()
combined_p.append(make_run("Shadow + outline together", size_pt=44, bold=True,
                           color="FFD966",
                           outline_w=12700, outline_color="8B4513",
                           shadow=True))
add_text(s1, left=Inches(0.5), top=Inches(5.3), width=Inches(12), height=Inches(1.4), paragraphs=[combined_p])

# ---------------------------------------------------------------------------
# Slide 2 — RTL paragraph (default alignment → right)
# ---------------------------------------------------------------------------
s2 = prs.slides.add_slide(blank)

t2 = make_paragraph()
t2.append(make_run("Paragraph @rtl — right alignment default", size_pt=24, bold=True))
add_text(s2, left=Inches(0.4), top=Inches(0.3), width=Inches(12.5), height=Inches(0.7), paragraphs=[t2])

# Non-RTL (control)
p_ltr = make_paragraph()
p_ltr.append(make_run("LTR control — no rtl flag, defaults to left", size_pt=28))
add_text(s2, left=Inches(0.5), top=Inches(1.6), width=Inches(12), height=Inches(1), paragraphs=[p_ltr])

# RTL paragraph — no algn → should auto-align right
p_rtl_auto = make_paragraph(rtl=True)
p_rtl_auto.append(make_run("RTL, no algn — auto right alignment", size_pt=28))
add_text(s2, left=Inches(0.5), top=Inches(2.8), width=Inches(12), height=Inches(1), paragraphs=[p_rtl_auto])

# RTL paragraph with explicit center
p_rtl_ctr = make_paragraph(algn="ctr", rtl=True)
p_rtl_ctr.append(make_run("RTL + algn=ctr — explicit centred", size_pt=28))
add_text(s2, left=Inches(0.5), top=Inches(4.0), width=Inches(12), height=Inches(1), paragraphs=[p_rtl_ctr])

# Hebrew-ish glyph sample (uses Latin chars but visible RTL effect)
p_rtl_demo = make_paragraph(rtl=True)
p_rtl_demo.append(make_run("שלום עולם — Hebrew sample (rtl=1)", size_pt=32, bold=True))
add_text(s2, left=Inches(0.5), top=Inches(5.2), width=Inches(12), height=Inches(1.2), paragraphs=[p_rtl_demo])

prs.save(OUT)
print(f"wrote {OUT}")
