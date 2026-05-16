"""Generate a pptx with pie/doughnut/scatter/bubble charts for VRT.

Run: python3 _gen_chart_demo.py
Output: chart-demo.pptx (gitignored — public/private/ is local-only).
"""
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData, BubbleChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches, Pt

OUT = "chart-demo.pptx"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]


def add_title(slide, text):
    tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.2), Inches(12.5), Inches(0.6))
    tf = tb.text_frame
    tf.text = text
    tf.paragraphs[0].runs[0].font.size = Pt(24)
    tf.paragraphs[0].runs[0].font.bold = True


def add_category_chart(slide, chart_type, title, left=Inches(0.5), top=Inches(1.0), width=Inches(6.0), height=Inches(5.5)):
    data = CategoryChartData()
    data.categories = ["Q1", "Q2", "Q3", "Q4"]
    data.add_series("North", (32, 28, 41, 36))
    data.add_series("South", (24, 30, 22, 28)) if chart_type not in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT) else None
    chart = slide.shapes.add_chart(chart_type, left, top, width, height, data).chart
    chart.has_title = True
    chart.chart_title.text_frame.text = title
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    return chart


def add_xy_chart(slide, chart_type, title, left=Inches(0.5), top=Inches(1.0), width=Inches(6.0), height=Inches(5.5)):
    data = XyChartData()
    s1 = data.add_series("Sample A")
    for x, y in [(1.0, 2.3), (2.4, 1.1), (3.7, 4.5), (4.2, 3.0), (5.0, 5.8)]:
        s1.add_data_point(x, y)
    s2 = data.add_series("Sample B")
    for x, y in [(0.8, 4.1), (2.1, 3.6), (3.0, 2.0), (4.5, 4.7), (5.4, 3.2)]:
        s2.add_data_point(x, y)
    chart = slide.shapes.add_chart(chart_type, left, top, width, height, data).chart
    chart.has_title = True
    chart.chart_title.text_frame.text = title
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    return chart


def add_bubble_chart(slide, title, left=Inches(0.5), top=Inches(1.0), width=Inches(6.0), height=Inches(5.5)):
    data = BubbleChartData()
    s1 = data.add_series("Series 1")
    for x, y, sz in [(1.0, 2.0, 10), (2.5, 3.2, 18), (4.0, 1.8, 25), (5.5, 4.5, 12)]:
        s1.add_data_point(x, y, sz)
    s2 = data.add_series("Series 2")
    for x, y, sz in [(1.5, 4.0, 15), (3.0, 2.7, 22), (4.5, 3.6, 8), (5.0, 1.2, 20)]:
        s2.add_data_point(x, y, sz)
    chart = slide.shapes.add_chart(XL_CHART_TYPE.BUBBLE, left, top, width, height, data).chart
    chart.has_title = True
    chart.chart_title.text_frame.text = title
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    return chart


# Slide 1: Pie + Doughnut
s1 = prs.slides.add_slide(blank)
add_title(s1, "Pie & Doughnut")
add_category_chart(s1, XL_CHART_TYPE.PIE, "Pie", left=Inches(0.3), width=Inches(6.4))
add_category_chart(s1, XL_CHART_TYPE.DOUGHNUT, "Doughnut", left=Inches(6.7), width=Inches(6.4))

# Slide 2: Scatter variants
s2 = prs.slides.add_slide(blank)
add_title(s2, "Scatter — markers / smooth / lines")
add_xy_chart(s2, XL_CHART_TYPE.XY_SCATTER, "Markers only", left=Inches(0.3), width=Inches(4.2))
add_xy_chart(s2, XL_CHART_TYPE.XY_SCATTER_SMOOTH, "Smooth lines + markers", left=Inches(4.6), width=Inches(4.2))
add_xy_chart(s2, XL_CHART_TYPE.XY_SCATTER_LINES, "Straight lines + markers", left=Inches(8.9), width=Inches(4.2))

# Slide 3: Bubble
s3 = prs.slides.add_slide(blank)
add_title(s3, "Bubble")
add_bubble_chart(s3, "Bubble — 2 series", left=Inches(3.0), width=Inches(7.0))

prs.save(OUT)
print(f"wrote {OUT}")
