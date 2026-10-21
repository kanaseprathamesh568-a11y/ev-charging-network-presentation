import os
import pptx
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image

# Initialize Presentation 16:9 Widescreen
prs = pptx.Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# Theme Colors (Exact match with Web Deck)
BG_COLOR = RGBColor(2, 6, 23)        # #020617 (slate-950)
CARD_BG = RGBColor(15, 23, 42)       # #0f172a (slate-900)
CARD_BORDER = RGBColor(30, 41, 59)   # #1e293b (slate-800)
CARD_ACTIVE = RGBColor(8, 51, 68)    # #083344 (cyan-950)
BORDER_ACTIVE = RGBColor(6, 182, 212)# #06b6d4 (cyan-500)

CYAN = RGBColor(34, 211, 238)        # #22d3ee (cyan-400)
CYAN_LIGHT = RGBColor(103, 232, 249) # #67e8f9 (cyan-300)
EMERALD = RGBColor(52, 211, 153)     # #34d399 (emerald-400)
AMBER = RGBColor(251, 191, 36)       # #fbbf24 (amber-400)
ROSE = RGBColor(248, 113, 113)       # #f87171 (rose-400)

TEXT_WHITE = RGBColor(255, 255, 255) # white
TEXT_LIGHT = RGBColor(203, 213, 225) # slate-300
TEXT_MUTED = RGBColor(148, 163, 184) # slate-400
TEXT_DIM = RGBColor(100, 116, 139)   # slate-500
FOOTER_TEXT = RGBColor(71, 85, 105)  # slate-600

# Cache directory for aspect-cropped images to avoid distortion
CROPPED_DIR = 'images/pptx_cropped'
os.makedirs(CROPPED_DIR, exist_ok=True)

def get_cropped_image(image_path, target_w_inch, target_h_inch):
    """Crops image to exact target aspect ratio using center crop to match CSS object-fit: cover."""
    basename = os.path.basename(image_path)
    tag = f"{target_w_inch:.2f}x{target_h_inch:.2f}"
    out_name = f"{os.path.splitext(basename)[0]}_{tag}.jpg"
    out_path = os.path.join(CROPPED_DIR, out_name)
    
    if os.path.exists(out_path):
        return out_path
        
    with Image.open(image_path) as im:
        im_w, im_h = im.size
        target_ratio = target_w_inch / target_h_inch
        current_ratio = im_w / im_h

        if current_ratio > target_ratio:
            new_w = int(im_h * target_ratio)
            left = (im_w - new_w) // 2
            im_cropped = im.crop((left, 0, left + new_w, im_h))
        else:
            new_h = int(im_w / target_ratio)
            top = (im_h - new_h) // 2
            im_cropped = im.crop((0, top, im_w, top + new_h))
        
        im_cropped.save(out_path, quality=95)
    return out_path

def set_slide_background(slide):
    bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = BG_COLOR
    bg_shape.line.fill.background()
    return bg_shape

def add_header(slide, tag, title, subtitle=None):
    tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.45), Inches(11.733), Inches(0.95))
    tf = tx_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    p_tag = tf.paragraphs[0]
    p_tag.text = tag.upper()
    p_tag.font.size = Pt(10)
    p_tag.font.bold = True
    p_tag.font.color.rgb = CYAN
    p_tag.space_after = Pt(2)

    p_title = tf.add_paragraph()
    p_title.text = title
    p_title.font.size = Pt(22)
    p_title.font.bold = True
    p_title.font.color.rgb = TEXT_WHITE

    if subtitle:
        p_title.space_after = Pt(2)
        p_sub = tf.add_paragraph()
        p_sub.text = subtitle
        p_sub.font.size = Pt(11)
        p_sub.font.color.rgb = TEXT_MUTED

def add_footer(slide, current, total=13):
    tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(11.733), Inches(0.3))
    tf = tx_box.text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = f"EV CHARGING NETWORK : Slide {current:02d} of {total:02d}"
    p.font.size = Pt(9)
    p.font.color.rgb = FOOTER_TEXT

def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=CARD_BORDER, border_width=1.0):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(border_width)
    return card

def add_photo_with_badge(slide, img_path, left, top, width, height, badge_title=None, badge_sub=None, badge_color=CYAN):
    cropped_path = get_cropped_image(img_path, width, height)
    slide.shapes.add_picture(cropped_path, Inches(left), Inches(top), width=Inches(width), height=Inches(height))
    
    # Border frame around picture
    frame = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    frame.fill.background()
    frame.line.color.rgb = CARD_BORDER
    frame.line.width = Pt(1.5)

    # Optional overlay badge card at bottom of image
    if badge_title or badge_sub:
        badge_h = 0.75 if badge_sub else 0.45
        badge_w = width - 0.3
        b_left = left + 0.15
        b_top = top + height - badge_h - 0.15
        
        b_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(b_left), Inches(b_top), Inches(badge_w), Inches(badge_h))
        b_card.fill.solid()
        b_card.fill.fore_color.rgb = CARD_BG
        b_card.line.color.rgb = CARD_BORDER
        b_card.line.width = Pt(1.0)
        
        tb = slide.shapes.add_textbox(Inches(b_left + 0.15), Inches(b_top + 0.08), Inches(badge_w - 0.3), Inches(badge_h - 0.16))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        
        if badge_title:
            p0 = tf.paragraphs[0]
            p0.text = badge_title.upper()
            p0.font.size = Pt(8.5)
            p0.font.bold = True
            p0.font.color.rgb = badge_color
        
        if badge_sub:
            p1 = tf.add_paragraph()
            p1.text = badge_sub
            p1.font.size = Pt(9.5)
            p1.font.color.rgb = TEXT_WHITE
            p1.space_before = Pt(2)


# ==============================================================================
# SLIDE 1: TITLE SLIDE (WITH REAL HERO PHOTO)
# ==============================================================================
s1 = prs.slides.add_slide(blank_layout)
set_slide_background(s1)

# Left Hero Text Column
tb1 = s1.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(6.5), Inches(2.2))
tf1 = tb1.text_frame
tf1.word_wrap = True
tf1.margin_left = tf1.margin_top = tf1.margin_right = tf1.margin_bottom = 0

p_tag1 = tf1.paragraphs[0]
p_tag1.text = "TECHNICAL ARCHITECTURE DECK"
p_tag1.font.size = Pt(10)
p_tag1.font.bold = True
p_tag1.font.color.rgb = CYAN
p_tag1.space_after = Pt(4)

p_tit1 = tf1.add_paragraph()
p_tit1.text = "EV CHARGING NETWORK"
p_tit1.font.size = Pt(36)
p_tit1.font.bold = True
p_tit1.font.color.rgb = TEXT_WHITE
p_tit1.space_after = Pt(4)

p_sub1 = tf1.add_paragraph()
p_sub1.text = "A secure, convenient and scalable charging ecosystem"
p_sub1.font.size = Pt(14)
p_sub1.font.color.rgb = TEXT_MUTED

# Quote Box
add_card(s1, 0.8, 3.1, 6.4, 0.75, bg_color=CARD_BG, border_color=CARD_BORDER)
tb_q = s1.shapes.add_textbox(Inches(1.0), Inches(3.22), Inches(6.0), Inches(0.5))
tf_q = tb_q.text_frame
tf_q.word_wrap = True
tf_q.margin_left = tf_q.margin_top = tf_q.margin_right = tf_q.margin_bottom = 0
pq = tf_q.paragraphs[0]
pq.text = '"YOU DON\'T NEED TO FIND A CHARGER. WE BRING CHARGING TO YOU."'
pq.font.size = Pt(11.5)
pq.font.bold = True
pq.font.color.rgb = CYAN_LIGHT

# 2 Model Cards
# Model 1
add_card(s1, 0.8, 4.05, 3.1, 2.0, bg_color=CARD_BG, border_color=CARD_BORDER)
tb_m1 = s1.shapes.add_textbox(Inches(0.95), Inches(4.2), Inches(2.8), Inches(1.7))
tf_m1 = tb_m1.text_frame
tf_m1.word_wrap = True
tf_m1.margin_left = tf_m1.margin_top = tf_m1.margin_right = tf_m1.margin_bottom = 0
pm1_tag = tf_m1.paragraphs[0]
pm1_tag.text = "MODEL 1 : URBAN"
pm1_tag.font.size = Pt(9.5)
pm1_tag.font.bold = True
pm1_tag.font.color.rgb = CYAN
pm1_tit = tf_m1.add_paragraph()
pm1_tit.text = "Charging Hub + Valet"
pm1_tit.font.size = Pt(13)
pm1_tit.font.bold = True
pm1_tit.font.color.rgb = TEXT_WHITE
pm1_tit.space_before = Pt(2)
pm1_tit.space_after = Pt(4)
pm1_desc = tf_m1.add_paragraph()
pm1_desc.text = "Vertical parking towers with dedicated chargers & solar. Scheduled evening pickup and 7:00 AM delivery."
pm1_desc.font.size = Pt(10)
pm1_desc.font.color.rgb = TEXT_LIGHT

# Model 2
add_card(s1, 4.1, 4.05, 3.1, 2.0, bg_color=CARD_BG, border_color=CARD_BORDER)
tb_m2 = s1.shapes.add_textbox(Inches(4.25), Inches(4.2), Inches(2.8), Inches(1.7))
tf_m2 = tb_m2.text_frame
tf_m2.word_wrap = True
tf_m2.margin_left = tf_m2.margin_top = tf_m2.margin_right = tf_m2.margin_bottom = 0
pm2_tag = tf_m2.paragraphs[0]
pm2_tag.text = "MODEL 2 : HIGHWAY"
pm2_tag.font.size = Pt(9.5)
pm2_tag.font.bold = True
pm2_tag.font.color.rgb = TEXT_MUTED
pm2_tit = tf_m2.add_paragraph()
pm2_tit.text = "Secure Distributed"
pm2_tit.font.size = Pt(13)
pm2_tit.font.bold = True
pm2_tit.font.color.rgb = TEXT_WHITE
pm2_tit.space_before = Pt(2)
pm2_tit.space_after = Pt(4)
pm2_desc = tf_m2.add_paragraph()
pm2_desc.text = "Verified host network with smart lock & virtual security. Monetizes private residential & commercial EVSEs."
pm2_desc.font.size = Pt(10)
pm2_desc.font.color.rgb = TEXT_LIGHT

# Subtext
tb_subnote = s1.shapes.add_textbox(Inches(0.8), Inches(6.2), Inches(6.4), Inches(0.4))
tf_subnote = tb_subnote.text_frame
tf_subnote.word_wrap = True
tf_subnote.margin_left = tf_subnote.margin_top = tf_subnote.margin_right = tf_subnote.margin_bottom = 0
ps = tf_subnote.paragraphs[0]
ps.text = "Dual-model architecture unifies high-density urban demand with distributed highway long-distance transit."
ps.font.size = Pt(9.5)
ps.font.color.rgb = TEXT_DIM

# Right Hero Photo with Badge
add_photo_with_badge(
    s1, 'images/ev_charging_hero.jpg',
    left=7.5, top=0.8, width=5.033, height=5.25,
    badge_title="Infrastructure Vision",
    badge_sub="Coordinated automated charging for urban & highway mobility."
)
add_footer(s1, 1)


# ==============================================================================
# SLIDE 2: 01 : PROBLEM
# ==============================================================================
s2 = prs.slides.add_slide(blank_layout)
set_slide_background(s2)
add_header(s2, "01 : PROBLEM", "The problem in one picture", "Empirical friction in urban EV ownership and long-distance travel")

# Left: 3 Node Cards
nodes = [
    ("[ Node 01 ] Apartment", "No Dedicated Charger", "Shared garages lack dedicated 240V circuits. Level 1 outlet adds only 4 miles/hr; full recharge requires 40+ hours."),
    ("[ Node 02 ] EV OWNER", "Needs Routine Charging", "Requires predictable schedule without losing 3-4 hours every week sitting at crowded public charging hubs."),
    ("[ Node 03 ] Public Charger", "Travel + Wait + Availability Uncertainty", "15%-25% dispenser failure rate, 20-45 minute peak queues, and unpredictable station availability.")
]

y_pos = 1.6
for tag, tit, desc in nodes:
    add_card(s2, 0.8, y_pos, 6.4, 1.35, bg_color=CARD_BG, border_color=CARD_BORDER)
    tb = s2.shapes.add_textbox(Inches(1.0), Inches(y_pos + 0.12), Inches(6.0), Inches(1.1))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p0 = tf.paragraphs[0]
    p0.text = tag
    p0.font.size = Pt(9.5)
    p0.font.bold = True
    p0.font.color.rgb = CYAN
    p1 = tf.add_paragraph()
    p1.text = tit
    p1.font.size = Pt(12.5)
    p1.font.bold = True
    p1.font.color.rgb = TEXT_WHITE
    p1.space_before = Pt(1)
    p1.space_after = Pt(2)
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.size = Pt(10)
    p2.font.color.rgb = TEXT_LIGHT
    y_pos += 1.45

# Right Photo with Badges
add_photo_with_badge(
    s2, 'images/public_charger_queue.jpg',
    left=7.5, top=1.6, width=5.033, height=4.25,
    badge_title="Highway Range Gap",
    badge_sub="Long-distance corridors have charging voids while private chargers sit idle."
)

# Top badge on photo
top_b = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.65), Inches(1.75), Inches(4.733), Inches(0.65))
top_b.fill.solid()
top_b.fill.fore_color.rgb = CARD_BG
top_b.line.color.rgb = CARD_BORDER
tb_tb = s2.shapes.add_textbox(Inches(7.8), Inches(1.82), Inches(4.4), Inches(0.5))
tf_tb = tb_tb.text_frame
tf_tb.word_wrap = True
tf_tb.margin_left = tf_tb.margin_top = tf_tb.margin_right = tf_tb.margin_bottom = 0
ptb0 = tf_tb.paragraphs[0]
ptb0.text = "URBAN BARRIER METRIC"
ptb0.font.size = Pt(8.5)
ptb0.font.bold = True
ptb0.font.color.rgb = ROSE
ptb1 = tf_tb.add_paragraph()
ptb1.text = "42% of urban EV drivers lack home charging"
ptb1.font.size = Pt(10.5)
ptb1.font.bold = True
ptb1.font.color.rgb = TEXT_WHITE

# Bottom Opportunity Box
add_card(s2, 0.8, 6.0, 11.733, 0.65, bg_color=CARD_BG, border_color=CARD_BORDER)
tb_opp = s2.shapes.add_textbox(Inches(1.0), Inches(6.12), Inches(11.333), Inches(0.4))
tf_opp = tb_opp.text_frame
tf_opp.word_wrap = True
tf_opp.margin_left = tf_opp.margin_top = tf_opp.margin_right = tf_opp.margin_bottom = 0
po = tf_opp.paragraphs[0]
po.text = "Opportunity: connect convenience with underutilized charging capacity."
po.font.size = Pt(11)
po.font.bold = True
po.font.color.rgb = CYAN_LIGHT
po.alignment = PP_ALIGN.CENTER
add_footer(s2, 2)


# ==============================================================================
# SLIDE 3: 02 : SOLUTION
# ==============================================================================
s3 = prs.slides.add_slide(blank_layout)
set_slide_background(s3)
add_header(s3, "02 : SOLUTION", "One platform. Two charging models.", "Unifying high-density company-owned hubs with an asset-light distributed network")

# Model 1 Card (Left)
add_card(s3, 0.8, 1.6, 5.7, 5.25, bg_color=CARD_BG, border_color=CARD_BORDER)
add_photo_with_badge(s3, 'images/vertical_parking_tower.jpg', left=1.0, top=1.8, width=5.3, height=1.9, badge_title="Company-Owned Infrastructure")
tb_sol1 = s3.shapes.add_textbox(Inches(1.0), Inches(3.85), Inches(5.3), Inches(2.8))
tf_sol1 = tb_sol1.text_frame
tf_sol1.word_wrap = True
tf_sol1.margin_left = tf_sol1.margin_top = tf_sol1.margin_right = tf_sol1.margin_bottom = 0
p = tf_sol1.paragraphs[0]
p.text = "MODEL 1 : CHARGING HUB"
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = TEXT_WHITE
p.space_after = Pt(4)
for b in [
    "- Company-owned hub: Dedicated multi-megawatt facility.",
    "- Vertical parking: 3x vehicle density per land footprint.",
    "- Dedicated chargers: Every stall has dedicated 22kW port.",
    "- Optional solar: Rooftop array with battery storage peak shaving."
]:
    pb = tf_sol1.add_paragraph()
    pb.text = b
    pb.font.size = Pt(10)
    pb.font.color.rgb = TEXT_LIGHT
    pb.space_after = Pt(2)
p_foot1 = tf_sol1.add_paragraph()
p_foot1.text = "Valet Concierge: Scheduled evening pickup and 7:00 AM delivery to owner bay."
p_foot1.font.size = Pt(9.5)
p_foot1.font.color.rgb = CYAN
p_foot1.space_before = Pt(4)

# Model 2 Card (Right)
add_card(s3, 6.833, 1.6, 5.7, 5.25, bg_color=CARD_BG, border_color=CARD_BORDER)
add_photo_with_badge(s3, 'images/residential_driveway_charger.jpg', left=7.033, top=1.8, width=5.3, height=1.9, badge_title="Asset-Light Host Hardware", badge_color=TEXT_MUTED)
tb_sol2 = s3.shapes.add_textbox(Inches(7.033), Inches(3.85), Inches(5.3), Inches(2.8))
tf_sol2 = tb_sol2.text_frame
tf_sol2.word_wrap = True
tf_sol2.margin_left = tf_sol2.margin_top = tf_sol2.margin_right = tf_sol2.margin_bottom = 0
p2 = tf_sol2.paragraphs[0]
p2.text = "MODEL 2 : DISTRIBUTED NETWORK"
p2.font.size = Pt(13)
p2.font.bold = True
p2.font.color.rgb = TEXT_WHITE
p2.space_after = Pt(4)
for b in [
    "- Verified host chargers: Monetizes private residential & commercial EVSEs.",
    "- Secure self-service: Zero host contact or physical keys required.",
    "- Highway availability: Fills long-distance route gaps between fast chargers.",
    "- Smart IoT control: Solenoid holster locks only during booked windows."
]:
    pb2 = tf_sol2.add_paragraph()
    pb2.text = b
    pb2.font.size = Pt(10)
    pb2.font.color.rgb = TEXT_LIGHT
    pb2.space_after = Pt(2)
p_foot2 = tf_sol2.add_paragraph()
p_foot2.text = "Distributed Scaling: Expands coverage across highways with zero construction CapEx."
p_foot2.font.size = Pt(9.5)
p_foot2.font.color.rgb = CYAN
p_foot2.space_before = Pt(4)
add_footer(s3, 3)


# ==============================================================================
# SLIDE 4: 03 : URBAN
# ==============================================================================
s4 = prs.slides.add_slide(blank_layout)
set_slide_background(s4)
add_header(s4, "03 : URBAN", "Model 1 - Charging Hub + Valet", "Interactive Lifecycle Demonstration: 5-step operational workflow")

# 5 Stepper Buttons Bar
steps = [
    ("01 : BOOK", "Schedule", True),
    ("02 : PICKUP", "Verified Valet", False),
    ("03 : TOWER", "Stack Parking", False),
    ("04 : CHARGE", "Solar + 22kW", False),
    ("05 : RETURN", "Morning Drop", False)
]
btn_w = 2.25
for i, (tag, lbl, is_active) in enumerate(steps):
    bx = 0.8 + i * (btn_w + 0.12)
    bg = CARD_ACTIVE if is_active else CARD_BG
    br = BORDER_ACTIVE if is_active else CARD_BORDER
    add_card(s4, bx, 1.55, btn_w, 0.7, bg_color=bg, border_color=br)
    tb_b = s4.shapes.add_textbox(Inches(bx + 0.1), Inches(1.6), Inches(btn_w - 0.2), Inches(0.55))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True
    tf_b.margin_left = tf_b.margin_top = tf_b.margin_right = tf_b.margin_bottom = 0
    p0 = tf_b.paragraphs[0]
    p0.text = tag
    p0.font.size = Pt(8.5)
    p0.font.bold = True
    p0.font.color.rgb = CYAN if is_active else TEXT_MUTED
    p1 = tf_b.add_paragraph()
    p1.text = lbl
    p1.font.size = Pt(10)
    p1.font.bold = True
    p1.font.color.rgb = TEXT_WHITE if is_active else TEXT_LIGHT

# Photo on Left
add_photo_with_badge(
    s4, 'images/vertical_parking_tower.jpg',
    left=0.8, top=2.45, width=4.8, height=2.6,
    badge_title="Tower Stacking System",
    badge_sub="Automated car lift positions vehicle in dedicated stall."
)

# Dynamic Operation Details on Right
add_card(s4, 5.8, 2.45, 6.733, 2.6, bg_color=CARD_BG, border_color=CARD_BORDER)
tb_op = s4.shapes.add_textbox(Inches(6.0), Inches(2.6), Inches(6.333), Inches(2.3))
tf_op = tb_op.text_frame
tf_op.word_wrap = True
tf_op.margin_left = tf_op.margin_top = tf_op.margin_right = tf_op.margin_bottom = 0
p = tf_op.paragraphs[0]
p.text = "STAGE 01 : BOOKING"
p.font.size = Pt(9.5)
p.font.bold = True
p.font.color.rgb = CYAN
p_tit = tf_op.add_paragraph()
p_tit.text = "Owner Schedules Pickup Window via Mobile App"
p_tit.font.size = Pt(13.5)
p_tit.font.bold = True
p_tit.font.color.rgb = TEXT_WHITE
p_tit.space_before = Pt(2)
p_tit.space_after = Pt(4)
p_desc = tf_op.add_paragraph()
p_desc.text = "Owner specifies target state of charge (80% or 100%) and selects 9:00 PM pickup with 7:00 AM delivery. Cloud platform reserves dedicated vertical hub stall and dispatches verified concierge valet."
p_desc.font.size = Pt(10.5)
p_desc.font.color.rgb = TEXT_LIGHT
p_desc.space_after = Pt(8)

p_met = tf_op.add_paragraph()
p_met.text = "Status: Reserved  |  Target SoC: 80%  |  Route: Optimal Stacking Level 3"
p_met.font.size = Pt(9.5)
p_met.font.bold = True
p_met.font.color.rgb = CYAN_LIGHT

# Bottom 3 Pillars
pillars = [
    ("VERTICAL PARKING:", "3x land density through multi-tier automated stacking."),
    ("DEDICATED CHARGERS:", "Each parking stall wired with dedicated 22kW AC port."),
    ("SOLAR POTENTIAL:", "Rooftop generation captures clean power for peak shaving.")
]
pil_w = 3.8
for i, (title, text) in enumerate(pillars):
    px = 0.8 + i * (pil_w + 0.16)
    add_card(s4, px, 5.25, pil_w, 1.45, bg_color=CARD_BG, border_color=CARD_BORDER)
    tb_p = s4.shapes.add_textbox(Inches(px + 0.2), Inches(5.4), Inches(pil_w - 0.4), Inches(1.1))
    tf_p = tb_p.text_frame
    tf_p.word_wrap = True
    tf_p.margin_left = tf_p.margin_top = tf_p.margin_right = tf_p.margin_bottom = 0
    p0 = tf_p.paragraphs[0]
    p0.text = title
    p0.font.size = Pt(11)
    p0.font.bold = True
    p0.font.color.rgb = TEXT_WHITE
    p0.space_after = Pt(2)
    p1 = tf_p.add_paragraph()
    p1.text = text
    p1.font.size = Pt(10)
    p1.font.color.rgb = TEXT_MUTED
add_footer(s4, 4)


# ==============================================================================
# SLIDE 5: 04 : HIGHWAY
# ==============================================================================
s5 = prs.slides.add_slide(blank_layout)
set_slide_background(s5)
add_header(s5, "04 : HIGHWAY", "Model 2 - Secure Distributed Charging", "Turn unused charging capacity into verified network capacity.")

# 5 Phase Stepper
phases = [
    ("01 : LIST", "Host Node", True),
    ("02 : BOOK", "Route Reserve", False),
    ("03 : AUTH", "BLE Token", False),
    ("04 : CHARGE", "Self-Plug", False),
    ("05 : LOCK", "Auto-Latch", False)
]
for i, (tag, lbl, is_active) in enumerate(phases):
    bx = 0.8 + i * (btn_w + 0.12)
    bg = CARD_ACTIVE if is_active else CARD_BG
    br = BORDER_ACTIVE if is_active else CARD_BORDER
    add_card(s5, bx, 1.55, btn_w, 0.7, bg_color=bg, border_color=br)
    tb_b = s5.shapes.add_textbox(Inches(bx + 0.1), Inches(1.6), Inches(btn_w - 0.2), Inches(0.55))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True
    tf_b.margin_left = tf_b.margin_top = tf_b.margin_right = tf_b.margin_bottom = 0
    p0 = tf_b.paragraphs[0]
    p0.text = tag
    p0.font.size = Pt(8.5)
    p0.font.bold = True
    p0.font.color.rgb = CYAN if is_active else TEXT_MUTED
    p1 = tf_b.add_paragraph()
    p1.text = lbl
    p1.font.size = Pt(10)
    p1.font.bold = True
    p1.font.color.rgb = TEXT_WHITE if is_active else TEXT_LIGHT

# Photo on Left
add_photo_with_badge(
    s5, 'images/residential_driveway_charger.jpg',
    left=0.8, top=2.45, width=4.8, height=2.6,
    badge_title="Verified Host Hardware",
    badge_sub="Residential Level 2 charger unlocked for road trip stopovers."
)

# Phase Details on Right
add_card(s5, 5.8, 2.45, 6.733, 2.6, bg_color=CARD_BG, border_color=CARD_BORDER)
tb_hp = s5.shapes.add_textbox(Inches(6.0), Inches(2.6), Inches(6.333), Inches(2.3))
tf_hp = tb_hp.text_frame
tf_hp.word_wrap = True
tf_hp.margin_left = tf_hp.margin_top = tf_hp.margin_right = tf_hp.margin_bottom = 0
p = tf_hp.paragraphs[0]
p.text = "PHASE 01 : LISTING"
p.font.size = Pt(9.5)
p.font.bold = True
p.font.color.rgb = CYAN
p_tit = tf_hp.add_paragraph()
p_tit.text = "Host Publishes Availability Window & Tariff"
p_tit.font.size = Pt(13.5)
p_tit.font.bold = True
p_tit.font.color.rgb = TEXT_WHITE
p_tit.space_before = Pt(2)
p_tit.space_after = Pt(4)
p_desc = tf_hp.add_paragraph()
p_desc.text = "Homeowner configures idle hours (e.g., 9:00 AM - 5:00 PM). Smart controller links to platform with zero manual host intervention. Platform coordinates identity, payment escrow, and charger solenoid unlocking."
p_desc.font.size = Pt(10.5)
p_desc.font.color.rgb = TEXT_LIGHT
p_desc.space_after = Pt(8)

p_met = tf_hp.add_paragraph()
p_met.text = "Hardware: 240V / 32A  |  Availability: Mon-Fri 8 hrs  |  Rate: Tariff + Margin"
p_met.font.size = Pt(9.5)
p_met.font.bold = True
p_met.font.color.rgb = CYAN_LIGHT

# Critical Design Decision Card
add_card(s5, 0.8, 5.25, 11.733, 1.45, bg_color=CARD_BG, border_color=CARD_BORDER)
tb_crit = s5.shapes.add_textbox(Inches(1.0), Inches(5.4), Inches(11.333), Inches(1.1))
tf_crit = tb_crit.text_frame
tf_crit.word_wrap = True
tf_crit.margin_left = tf_crit.margin_top = tf_crit.margin_right = tf_crit.margin_bottom = 0
pc0 = tf_crit.paragraphs[0]
pc0.text = "CRITICAL DESIGN DECISION"
pc0.font.size = Pt(9.5)
pc0.font.bold = True
pc0.font.color.rgb = CYAN
pc1 = tf_crit.add_paragraph()
pc1.text = "Autonomous Self-Service Access"
pc1.font.size = Pt(12)
pc1.font.bold = True
pc1.font.color.rgb = TEXT_WHITE
pc1.space_before = Pt(1)
pc1.space_after = Pt(2)
pc2 = tf_crit.add_paragraph()
pc2.text = "A highway traveller should not need a company driver at every residential charger. The system therefore supports secure self-service: automated smart solenoid locks, Bluetooth cryptographic handshake, and virtual perimeter surveillance."
pc2.font.size = Pt(10)
pc2.font.color.rgb = TEXT_LIGHT
add_footer(s5, 5)


# ==============================================================================
# SLIDE 6: 05 : TRUST
# ==============================================================================
s6 = prs.slides.add_slide(blank_layout)
set_slide_background(s6)
add_header(s6, "05 : TRUST", "Security: a charger is not an open door", "Interactive Virtual Perimeter: 3-Zone CCTV monitoring framework")

# 3-Zone Bar
z_w = 3.8
zones = [
    ("ZONE 1 : GREEN", "Safe Authorized Stall", EMERALD, True),
    ("ZONE 2 : YELLOW", "Warning Buffer (3m)", AMBER, False),
    ("ZONE 3 : RED", "Restricted Residence", ROSE, False)
]
for i, (tag, lbl, col, is_act) in enumerate(zones):
    zx = 0.8 + i * (z_w + 0.16)
    bg = CARD_BG
    br = col if is_act else CARD_BORDER
    add_card(s6, zx, 1.55, z_w, 0.7, bg_color=bg, border_color=br, border_width=1.5 if is_act else 1.0)
    tb_z = s6.shapes.add_textbox(Inches(zx + 0.15), Inches(1.6), Inches(z_w - 0.3), Inches(0.55))
    tf_z = tb_z.text_frame
    tf_z.word_wrap = True
    tf_z.margin_left = tf_z.margin_top = tf_z.margin_right = tf_z.margin_bottom = 0
    p0 = tf_z.paragraphs[0]
    p0.text = tag
    p0.font.size = Pt(9.5)
    p0.font.bold = True
    p0.font.color.rgb = col
    p1 = tf_z.add_paragraph()
    p1.text = lbl
    p1.font.size = Pt(10)
    p1.font.bold = True
    p1.font.color.rgb = TEXT_WHITE if is_act else TEXT_MUTED

# Left Image with Optical HUD
add_photo_with_badge(
    s6, 'images/cctv_security_driveway.jpg',
    left=0.8, top=2.45, width=5.5, height=2.6,
    badge_title="CAM 01 : GREEN SAFE ZONE  |  REC 1080P",
    badge_sub="Vehicle in Bay : Normal Charging Allowed",
    badge_color=EMERALD
)

# Right Status Details
add_card(s6, 6.5, 2.45, 6.033, 2.6, bg_color=CARD_BG, border_color=CARD_BORDER)
tb_sec = s6.shapes.add_textbox(Inches(6.7), Inches(2.6), Inches(5.633), Inches(2.3))
tf_sec = tb_sec.text_frame
tf_sec.word_wrap = True
tf_sec.margin_left = tf_sec.margin_top = tf_sec.margin_right = tf_sec.margin_bottom = 0
p = tf_sec.paragraphs[0]
p.text = "ACTIVE ZONE : GREEN (SAFE)"
p.font.size = Pt(9.5)
p.font.bold = True
p.font.color.rgb = EMERALD
p_tit = tf_sec.add_paragraph()
p_tit.text = "Driveway Staging Bay & Charger Port Area"
p_tit.font.size = Pt(13)
p_tit.font.bold = True
p_tit.font.color.rgb = TEXT_WHITE
p_tit.space_before = Pt(2)
p_tit.space_after = Pt(4)
p_desc = tf_sec.add_paragraph()
p_desc.text = "Vehicle parked in bay. Charging cable connected. No unauthorized proximity detected. Host family perimeter is untouched. Optical AI verifies human presence is restricted exclusively to vehicle charging port."
p_desc.font.size = Pt(10)
p_desc.font.color.rgb = TEXT_LIGHT
p_desc.space_after = Pt(8)

p_met = tf_sec.add_paragraph()
p_met.text = "Perimeter Status: Normal  |  Audio Alert: Inactive  |  Solenoid Lock: Engaged"
p_met.font.size = Pt(9.5)
p_met.font.bold = True
p_met.font.color.rgb = EMERALD

# Bottom Card
add_card(s6, 0.8, 5.25, 11.733, 1.45, bg_color=CARD_BG, border_color=CARD_BORDER)
tb_secbot = s6.shapes.add_textbox(Inches(1.0), Inches(5.4), Inches(11.333), Inches(1.1))
tf_secbot = tb_secbot.text_frame
tf_secbot.word_wrap = True
tf_secbot.margin_left = tf_secbot.margin_top = tf_secbot.margin_right = tf_secbot.margin_bottom = 0
p0 = tf_secbot.paragraphs[0]
p0.text = "CONCEPTUAL SAFETY FRAMEWORK"
p0.font.size = Pt(9.5)
p0.font.bold = True
p0.font.color.rgb = CYAN
p1 = tf_secbot.add_paragraph()
p1.text = "CCTV + Virtual Perimeter Monitoring"
p1.font.size = Pt(12)
p1.font.bold = True
p1.font.color.rgb = TEXT_WHITE
p1.space_before = Pt(1)
p1.space_after = Pt(2)
p2 = tf_secbot.add_paragraph()
p2.text = "CCTV + virtual perimeter monitoring keeps the charging interaction controlled. The zones are a conceptual safety framework designed to reduce unwanted movement and false alarms: green allows charging, yellow sounds polite audio guidance, red trips security alerts and terminates the session."
p2.font.size = Pt(10)
p2.font.color.rgb = TEXT_LIGHT
add_footer(s6, 6)


# ==============================================================================
# SLIDE 7: 06 : ACCESS
# ==============================================================================
s7 = prs.slides.add_slide(blank_layout)
set_slide_background(s7)
add_header(s7, "06 : ACCESS", "Smart charger control", "Hardware-level access security and dual-mode separation")

# Mode Selector
m_w = 5.75
add_card(s7, 0.8, 1.55, m_w, 0.7, bg_color=CARD_ACTIVE, border_color=BORDER_ACTIVE)
tb_m1 = s7.shapes.add_textbox(Inches(0.95), Inches(1.6), Inches(m_w - 0.3), Inches(0.55))
tf_m1 = tb_m1.text_frame
tf_m1.word_wrap = True
tf_m1.margin_left = tf_m1.margin_top = tf_m1.margin_right = tf_m1.margin_bottom = 0
p0 = tf_m1.paragraphs[0]
p0.text = "MODE 01 : OWNER"
p0.font.size = Pt(9)
p0.font.bold = True
p0.font.color.rgb = CYAN
p1 = tf_m1.add_paragraph()
p1.text = "Owner Authenticates -> Uses -> Locks"
p1.font.size = Pt(10.5)
p1.font.bold = True
p1.font.color.rgb = TEXT_WHITE

add_card(s7, 6.783, 1.55, m_w, 0.7, bg_color=CARD_BG, border_color=CARD_BORDER)
tb_m2 = s7.shapes.add_textbox(Inches(6.933), Inches(1.6), Inches(m_w - 0.3), Inches(0.55))
tf_m2 = tb_m2.text_frame
tf_m2.word_wrap = True
tf_m2.margin_left = tf_m2.margin_top = tf_m2.margin_right = tf_m2.margin_bottom = 0
p0 = tf_m2.paragraphs[0]
p0.text = "MODE 02 : TRAVELLER"
p0.font.size = Pt(9)
p0.font.bold = True
p0.font.color.rgb = TEXT_MUTED
p1 = tf_m2.add_paragraph()
p1.text = "Temporary Token Authorization -> Locks"
p1.font.size = Pt(10.5)
p1.font.bold = True
p1.font.color.rgb = TEXT_LIGHT

# Photo on Left
add_photo_with_badge(
    s7, 'images/smart_phone_nfc_charger.jpg',
    left=0.8, top=2.45, width=4.8, height=2.6,
    badge_title="NFC & BLE Tap Interface",
    badge_sub="Hardware unlocks only upon verified digital key handshake."
)

# Right Mode Details
add_card(s7, 5.8, 2.45, 6.733, 2.6, bg_color=CARD_BG, border_color=CARD_BORDER)
tb_acc = s7.shapes.add_textbox(Inches(6.0), Inches(2.6), Inches(6.333), Inches(2.3))
tf_acc = tb_acc.text_frame
tf_acc.word_wrap = True
tf_acc.margin_left = tf_acc.margin_top = tf_acc.margin_right = tf_acc.margin_bottom = 0
p = tf_acc.paragraphs[0]
p.text = "FIRMWARE MODE : OWNER MODE"
p.font.size = Pt(9.5)
p.font.bold = True
p.font.color.rgb = CYAN
p_tit = tf_acc.add_paragraph()
p_tit.text = "Owner NFC / Smartphone Bypass"
p_tit.font.size = Pt(13.5)
p_tit.font.bold = True
p_tit.font.color.rgb = TEXT_WHITE
p_tit.space_before = Pt(2)
p_tit.space_after = Pt(4)
p_desc = tf_acc.add_paragraph()
p_desc.text = "Owner registers master keycard or smartphone. Instant proximity tap releases holster latch. No billing fees or reservations required for personal electricity usage. System switches automatically to Guest Mode when reservation begins."
p_desc.font.size = Pt(10.5)
p_desc.font.color.rgb = TEXT_LIGHT
p_desc.space_after = Pt(8)

p_met = tf_acc.add_paragraph()
p_met.text = "Authorization: Master Key  |  Billing: Direct Home  |  Latch: Instant Tap"
p_met.font.size = Pt(9.5)
p_met.font.bold = True
p_met.font.color.rgb = CYAN_LIGHT

# Bottom Card
add_card(s7, 0.8, 5.25, 11.733, 1.45, bg_color=CARD_BG, border_color=CARD_BORDER)
tb_accbot = s7.shapes.add_textbox(Inches(1.0), Inches(5.4), Inches(11.333), Inches(1.1))
tf_accbot = tb_accbot.text_frame
tf_accbot.word_wrap = True
tf_accbot.margin_left = tf_accbot.margin_top = tf_accbot.margin_right = tf_accbot.margin_bottom = 0
p0 = tf_accbot.paragraphs[0]
p0.text = "SAFETY SPECIFICATION"
p0.font.size = Pt(9.5)
p0.font.bold = True
p0.font.color.rgb = CYAN
p1 = tf_accbot.add_paragraph()
p1.text = "Dual-Channel Separation"
p1.font.size = Pt(12)
p1.font.bold = True
p1.font.color.rgb = TEXT_WHITE
p1.space_before = Pt(1)
p1.space_after = Pt(2)
p2 = tf_accbot.add_paragraph()
p2.text = "Safety: the external module is a security/monitoring controller. High-current charging remains under a properly rated EVSE and supported control mechanisms. Solenoid holster latch prevents unauthorized plug removal until session completion."
p2.font.size = Pt(10)
p2.font.color.rgb = TEXT_LIGHT
add_footer(s7, 7)


# ==============================================================================
# SLIDE 8: 07 : VERIFICATION
# ==============================================================================
s8 = prs.slides.add_slide(blank_layout)
set_slide_background(s8)
add_header(s8, "07 : VERIFICATION", "Every session leaves evidence", "Tripartite photographic inspection workbench & computer vision comparison")

# Left Photo with CV Inspection Badges
add_photo_with_badge(
    s8, 'images/ev_port_inspection.jpg',
    left=0.8, top=1.6, width=4.8, height=3.45,
    badge_title="Port Latch & Contact Verification",
    badge_sub="Automated high-resolution macro visual evidence logging."
)

# Right: 3 Tripartite Inspection Cards
steps_v = [
    ("[ 01 : BEFORE ] Pre-Session Baseline Scan", "Vehicle port flap, connector pins, cable sheath, and driveway baseline recorded prior to charging."),
    ("[ 02 : CHARGING ] Continuous Session Logging", "1 Hz telemetry stream + CCTV perimeter motion clip recording throughout the active charging session."),
    ("[ 03 : AFTER ] Post-Session Diff", "Repeat matching photos + computer vision condition comparison to verify zero damage or cable strain.")
]
vy = 1.6
for title, text in steps_v:
    add_card(s8, 5.8, vy, 6.733, 1.05, bg_color=CARD_BG, border_color=CARD_BORDER)
    tb_v = s8.shapes.add_textbox(Inches(6.0), Inches(vy + 0.1), Inches(6.333), Inches(0.85))
    tf_v = tb_v.text_frame
    tf_v.word_wrap = True
    tf_v.margin_left = tf_v.margin_top = tf_v.margin_right = tf_v.margin_bottom = 0
    p0 = tf_v.paragraphs[0]
    p0.text = title
    p0.font.size = Pt(10)
    p0.font.bold = True
    p0.font.color.rgb = CYAN
    p1 = tf_v.add_paragraph()
    p1.text = text
    p1.font.size = Pt(9.5)
    p1.font.color.rgb = TEXT_LIGHT
    p1.space_before = Pt(2)
    vy += 1.2

# Bottom Workbench Card
add_card(s8, 0.8, 5.25, 11.733, 1.45, bg_color=CARD_BG, border_color=CARD_BORDER)
tb_vbot = s8.shapes.add_textbox(Inches(1.0), Inches(5.4), Inches(11.333), Inches(1.1))
tf_vbot = tb_vbot.text_frame
tf_vbot.word_wrap = True
tf_vbot.margin_left = tf_vbot.margin_top = tf_vbot.margin_right = tf_vbot.margin_bottom = 0
p0 = tf_vbot.paragraphs[0]
p0.text = "AUTOMATED AI DIFFERENCE VERIFICATION"
p0.font.size = Pt(9.5)
p0.font.bold = True
p0.font.color.rgb = CYAN
p1 = tf_vbot.add_paragraph()
p1.text = "16-Point Computer Vision Diff Analysis"
p1.font.size = Pt(12)
p1.font.bold = True
p1.font.color.rgb = TEXT_WHITE
p1.space_before = Pt(1)
p1.space_after = Pt(2)
p2 = tf_vbot.add_paragraph()
p2.text = "Algorithms compare pre-charge and departure imagery to detect cosmetic scratches, cable strain, or mechanical port wear before liability escrow is released. Status: Verified Baseline Passed (Zero Anomaly Detected)."
p2.font.size = Pt(10)
p2.font.color.rgb = TEXT_LIGHT
add_footer(s8, 8)


# ==============================================================================
# SLIDE 9: 08 : MONITORING
# ==============================================================================
s9 = prs.slides.add_slide(blank_layout)
set_slide_background(s9)
add_header(s9, "08 : MONITORING", "Protect the charger and the vehicle", "Real-time sensor telemetry and automated fault isolation controls")

# Left Photo
add_photo_with_badge(
    s9, 'images/ev_telemetry_dashboard.jpg',
    left=0.8, top=1.6, width=4.8, height=2.6,
    badge_title="Live Telemetry Supervision",
    badge_sub="Continuous monitoring protects charger hardware and vehicle battery."
)

# Right: 4 Gauges
gauges = [
    ("TEMPERATURE", "32.4 deg C", "Thermocouple (Safe)", CYAN_LIGHT),
    ("VOLT / CURRENT", "230V / 32A", "Phase Health: Nominal", TEXT_WHITE),
    ("POWER DELIVERY", "7.36 kW", "Class 1 Meter Output", CYAN_LIGHT),
    ("FAULT CODE", "0x00 NORMAL", "GFCI State: Active Protected", EMERALD)
]
gw = 3.25
for i, (tag, val, sub, col) in enumerate(gauges):
    col_idx = i % 2
    row_idx = i // 2
    gx = 5.8 + col_idx * (gw + 0.233)
    gy = 1.6 + row_idx * 1.35
    add_card(s9, gx, gy, gw, 1.25, bg_color=CARD_BG, border_color=CARD_BORDER)
    tb_g = s9.shapes.add_textbox(Inches(gx + 0.15), Inches(gy + 0.12), Inches(gw - 0.3), Inches(1.0))
    tf_g = tb_g.text_frame
    tf_g.word_wrap = True
    tf_g.margin_left = tf_g.margin_top = tf_g.margin_right = tf_g.margin_bottom = 0
    p0 = tf_g.paragraphs[0]
    p0.text = tag
    p0.font.size = Pt(8.5)
    p0.font.bold = True
    p0.font.color.rgb = CYAN
    p0.alignment = PP_ALIGN.CENTER
    p1 = tf_g.add_paragraph()
    p1.text = val
    p1.font.size = Pt(14)
    p1.font.bold = True
    p1.font.color.rgb = col
    p1.alignment = PP_ALIGN.CENTER
    p1.space_before = Pt(1)
    p2 = tf_g.add_paragraph()
    p2.text = sub
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = TEXT_MUTED
    p2.alignment = PP_ALIGN.CENTER

# Bottom: 4 Sequence Cards
seq = [
    ("1. Detect anomaly", "<10ms polling on voltage, current & ground."),
    ("2. Safe EVSE response", "<100ms pilot trip disengages main contactor safely."),
    ("3. Notify", "Instant alerts delivered to driver and host devices."),
    ("4. Inspect if needed", "Remote diagnostics logs flagged for technical review.")
]
sq_w = 2.8
for i, (tit, desc) in enumerate(seq):
    sx = 0.8 + i * (sq_w + 0.177)
    add_card(s9, sx, 4.45, sq_w, 2.25, bg_color=CARD_BG, border_color=CARD_BORDER)
    tb_sq = s9.shapes.add_textbox(Inches(sx + 0.15), Inches(4.6), Inches(sq_w - 0.3), Inches(1.9))
    tf_sq = tb_sq.text_frame
    tf_sq.word_wrap = True
    tf_sq.margin_left = tf_sq.margin_top = tf_sq.margin_right = tf_sq.margin_bottom = 0
    p0 = tf_sq.paragraphs[0]
    p0.text = tit
    p0.font.size = Pt(11.5)
    p0.font.bold = True
    p0.font.color.rgb = CYAN if i == 1 else TEXT_WHITE
    p0.space_after = Pt(4)
    p1 = tf_sq.add_paragraph()
    p1.text = desc
    p1.font.size = Pt(10)
    p1.font.color.rgb = TEXT_LIGHT
add_footer(s9, 9)


# ==============================================================================
# SLIDE 10: 09 : PLATFORM
# ==============================================================================
s10 = prs.slides.add_slide(blank_layout)
set_slide_background(s10)
add_header(s10, "09 : PLATFORM", "The ecosystem", "Unifying demand, intelligence, and distributed electrical supply")

# 3 Structural Cards
cols = [
    ("[ DEMAND LAYER ]", "EV OWNER", "Convenient charging", [
        "- Urban apartment valet scheduling",
        "- Highway corridor route reservations",
        "- Transparent fixed per-kWh tariffs",
        "- Real-time battery status tracking"
    ], "Zero charging chore for urban owners."),
    ("[ CORE ENGINE ]", "PLATFORM", "Booking - Security - Payment", [
        "- Dynamic slot booking optimization",
        "- Cryptographic JWT token auth broker",
        "- Automated payment clearing & escrow",
        "- Computer vision condition comparator"
    ], "Coordinates identity, hardware, and billing."),
    ("[ SUPPLY LAYER ]", "CHARGER PROVIDERS", "Hubs - Homes - Hotels", [
        "- Hubs: Owned automated vertical towers",
        "- Homes: Residential private driveways",
        "- Businesses: Commercial hospitality",
        "- Hardware telemetry & lock relays"
    ], "Turns idle electrical assets into revenue.")
]
cw = 3.75
for i, (layer, tit, sub, bullets, foot) in enumerate(cols):
    cx = 0.8 + i * (cw + 0.24)
    add_card(s10, cx, 1.6, cw, 5.25, bg_color=CARD_BG, border_color=BORDER_ACTIVE if i == 1 else CARD_BORDER)
    tb_c = s10.shapes.add_textbox(Inches(cx + 0.2), Inches(1.8), Inches(cw - 0.4), Inches(4.8))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True
    tf_c.margin_left = tf_c.margin_top = tf_c.margin_right = tf_c.margin_bottom = 0
    p0 = tf_c.paragraphs[0]
    p0.text = layer
    p0.font.size = Pt(9.5)
    p0.font.bold = True
    p0.font.color.rgb = CYAN
    p1 = tf_c.add_paragraph()
    p1.text = tit
    p1.font.size = Pt(16)
    p1.font.bold = True
    p1.font.color.rgb = TEXT_WHITE
    p1.space_before = Pt(2)
    p2 = tf_c.add_paragraph()
    p2.text = sub
    p2.font.size = Pt(11)
    p2.font.bold = True
    p2.font.color.rgb = CYAN_LIGHT if i == 1 else TEXT_LIGHT
    p2.space_after = Pt(8)
    for b in bullets:
        pb = tf_c.add_paragraph()
        pb.text = b
        pb.font.size = Pt(10)
        pb.font.color.rgb = TEXT_LIGHT
        pb.space_after = Pt(4)
    pf = tf_c.add_paragraph()
    pf.text = foot
    pf.font.size = Pt(10)
    pf.font.color.rgb = CYAN if i == 1 else TEXT_MUTED
    pf.space_before = Pt(14)
add_footer(s10, 10)


# ==============================================================================
# SLIDE 11: 10 : BUSINESS
# ==============================================================================
s11 = prs.slides.add_slide(blank_layout)
set_slide_background(s11)
add_header(s11, "10 : BUSINESS", "Scale without building every charger", "Asset-right hybrid capital strategy balancing high-margin anchors with network scaling")

# Top 3 Equation Cards
eq_cards = [
    ("OWNED HUBS", "Controlled Urban Operations", [
        "- Vertical parking land efficiency",
        "- Premium white-glove valet service",
        "- High volume and predictability",
        "- Solar and battery margin capture"
    ], CYAN),
    ("+ DISTRIBUTED ASSETS", "Verified Unused Chargers", [
        "- Zero civil construction CapEx",
        "- Rapid geographic corridor coverage",
        "- Automated self-service operations",
        "- Host revenue-share alignment"
    ], TEXT_LIGHT),
    ("= CHARGING NETWORK", "Maximum Capital Efficiency", [
        "More access with less infrastructure duplication. Expands across highway miles without building billions in single-use charger concrete."
    ], CYAN_LIGHT)
]
for i, (tag, tit, items, tag_col) in enumerate(eq_cards):
    ex = 0.8 + i * (cw + 0.24)
    add_card(s11, ex, 1.6, cw, 2.7, bg_color=CARD_BG, border_color=BORDER_ACTIVE if i == 2 else CARD_BORDER)
    tb_e = s11.shapes.add_textbox(Inches(ex + 0.2), Inches(1.75), Inches(cw - 0.4), Inches(2.3))
    tf_e = tb_e.text_frame
    tf_e.word_wrap = True
    tf_e.margin_left = tf_e.margin_top = tf_e.margin_right = tf_e.margin_bottom = 0
    p0 = tf_e.paragraphs[0]
    p0.text = tag
    p0.font.size = Pt(9.5)
    p0.font.bold = True
    p0.font.color.rgb = tag_col
    p1 = tf_e.add_paragraph()
    p1.text = tit
    p1.font.size = Pt(13)
    p1.font.bold = True
    p1.font.color.rgb = TEXT_WHITE
    p1.space_before = Pt(2)
    p1.space_after = Pt(4)
    for it in items:
        pi = tf_e.add_paragraph()
        pi.text = it
        pi.font.size = Pt(10)
        pi.font.color.rgb = TEXT_LIGHT
        pi.space_after = Pt(2)

# Bottom: 7 Diversified Revenue Streams Grid
add_card(s11, 0.8, 4.5, 11.733, 2.35, bg_color=CARD_BG, border_color=CARD_BORDER)
tb_rev_hdr = s11.shapes.add_textbox(Inches(1.0), Inches(4.65), Inches(11.333), Inches(0.35))
tf_rev_hdr = tb_rev_hdr.text_frame
tf_rev_hdr.margin_left = tf_rev_hdr.margin_top = tf_rev_hdr.margin_right = tf_rev_hdr.margin_bottom = 0
pr = tf_rev_hdr.paragraphs[0]
pr.text = "7 DIVERSIFIED REVENUE STREAMS"
pr.font.size = Pt(10)
pr.font.bold = True
pr.font.color.rgb = CYAN

rev_streams = [
    ("Charging Fee", "Per-kWh markup"),
    ("Valet Fee", "Pickup service"),
    ("Platform Fee", "Clearing fee"),
    ("Host Cut", "Marketplace split"),
    ("Membership", "Monthly tier"),
    ("Fleet Plans", "B2B contracts"),
    ("Premium", "Care & detail")
]
rw = 1.55
for i, (tit, sub) in enumerate(rev_streams):
    rx = 0.95 + i * (rw + 0.1)
    add_card(s11, rx, 5.15, rw, 1.45, bg_color=BG_COLOR, border_color=CARD_BORDER)
    tb_r = s11.shapes.add_textbox(Inches(rx + 0.1), Inches(5.35), Inches(rw - 0.2), Inches(1.05))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True
    tf_r.margin_left = tf_r.margin_top = tf_r.margin_right = tf_r.margin_bottom = 0
    p0 = tf_r.paragraphs[0]
    p0.text = tit
    p0.font.size = Pt(10.5)
    p0.font.bold = True
    p0.font.color.rgb = TEXT_WHITE
    p0.alignment = PP_ALIGN.CENTER
    p1 = tf_r.add_paragraph()
    p1.text = sub
    p1.font.size = Pt(9)
    p1.font.color.rgb = TEXT_MUTED
    p1.alignment = PP_ALIGN.CENTER
    p1.space_before = Pt(4)
add_footer(s11, 11)


# ==============================================================================
# SLIDE 12: 11 : ROADMAP
# ==============================================================================
s12 = prs.slides.add_slide(blank_layout)
set_slide_background(s12)
add_header(s12, "11 : ROADMAP", "Pilot -> network -> scale", "Strategic 5-phase growth pathway to nationwide market leadership")

road_phases = [
    ("01", "Pilot", "Charging hub + valet in one city", "Validate valet turnaround logistics, app booking, and subscriber retention in flagship metro.", "Phase 1", True),
    ("02", "Connect", "Verified residential / commercial", "Onboard first 250 private host chargers with external smart lock controllers and CCTV kits.", "Phase 2", False),
    ("03", "Self-service", "Secure highway charging", "Deploy automated IoT smart locks along key highway transit corridors for unassisted travel.", "Phase 3", False),
    ("04", "Expand", "Geographic network growth", "Multi-city launch with 10 urban hubs and 2,500+ host locations across high-adoption states.", "Phase 4", False),
    ("05", "Integrate", "Fleets, hotels, parking", "Commercial fleet partnerships, hospitality chains, and utility grid peak-shaving demand response.", "Phase 5", False)
]
pw = 2.25
for i, (num, tit, sub, desc, foot, is_act) in enumerate(road_phases):
    px = 0.8 + i * (pw + 0.12)
    add_card(s12, px, 1.6, pw, 5.25, bg_color=CARD_BG, border_color=BORDER_ACTIVE if is_act else CARD_BORDER)
    tb_rd = s12.shapes.add_textbox(Inches(px + 0.15), Inches(1.8), Inches(pw - 0.3), Inches(4.8))
    tf_rd = tb_rd.text_frame
    tf_rd.word_wrap = True
    tf_rd.margin_left = tf_rd.margin_top = tf_rd.margin_right = tf_rd.margin_bottom = 0
    p0 = tf_rd.paragraphs[0]
    p0.text = num
    p0.font.size = Pt(20)
    p0.font.bold = True
    p0.font.color.rgb = CYAN if is_act else TEXT_DIM
    p1 = tf_rd.add_paragraph()
    p1.text = tit
    p1.font.size = Pt(14)
    p1.font.bold = True
    p1.font.color.rgb = TEXT_WHITE
    p1.space_before = Pt(2)
    p2 = tf_rd.add_paragraph()
    p2.text = sub
    p2.font.size = Pt(9.5)
    p2.font.bold = True
    p2.font.color.rgb = CYAN_LIGHT if is_act else TEXT_LIGHT
    p2.space_before = Pt(1)
    p2.space_after = Pt(8)
    p3 = tf_rd.add_paragraph()
    p3.text = desc
    p3.font.size = Pt(9.5)
    p3.font.color.rgb = TEXT_LIGHT
    p3.space_after = Pt(16)
    pf = tf_rd.add_paragraph()
    pf.text = foot
    pf.font.size = Pt(9)
    pf.font.color.rgb = TEXT_MUTED
add_footer(s12, 12)


# ==============================================================================
# SLIDE 13: 12 : CONCLUSION
# ==============================================================================
s13 = prs.slides.add_slide(blank_layout)
set_slide_background(s13)

# Executive Header
tb_c_hdr = s13.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(1.2))
tf_c_hdr = tb_c_hdr.text_frame
tf_c_hdr.word_wrap = True
tf_c_hdr.margin_left = tf_c_hdr.margin_top = tf_c_hdr.margin_right = tf_c_hdr.margin_bottom = 0
p0 = tf_c_hdr.paragraphs[0]
p0.text = "EXECUTIVE SUMMARY"
p0.font.size = Pt(9.5)
p0.font.bold = True
p0.font.color.rgb = CYAN
p0.alignment = PP_ALIGN.CENTER
p1 = tf_c_hdr.add_paragraph()
p1.text = "CHARGE WITHOUT THE CHASE."
p1.font.size = Pt(26)
p1.font.bold = True
p1.font.color.rgb = TEXT_WHITE
p1.alignment = PP_ALIGN.CENTER
p1.space_before = Pt(1)
p2 = tf_c_hdr.add_paragraph()
p2.text = "From apartment charging to highway mobility - one secure EV charging network."
p2.font.size = Pt(11.5)
p2.font.color.rgb = CYAN_LIGHT
p2.alignment = PP_ALIGN.CENTER
p2.space_before = Pt(2)

# Middle Row: Left Photo Card + Right 4 Pillars in 2x2 Grid
add_photo_with_badge(
    s13, 'images/highway_scenic_ev.jpg',
    left=0.8, top=1.75, width=5.2, height=2.85,
    badge_title="Corridor Freedom",
    badge_sub="Connected highway route travel with verified host chargers."
)

# Right: 4 Pillars in 2x2 Grid
pillars_c = [
    ("CONVENIENCE", "Eliminates charging chores for apartment residents via scheduled valet service."),
    ("SECURITY", "3-zone perimeter, CCTV event logs, smart hardware lock, and evidence verification."),
    ("UTILIZATION", "Monetizes idle private chargers, bridging highway range gaps with zero CapEx."),
    ("SCALABILITY", "Owned hubs combined with asset-light host network scales rapidly.")
]
pil_w2 = 3.05
pil_h2 = 1.35
for i, (tit, desc) in enumerate(pillars_c):
    c_idx = i % 2
    r_idx = i // 2
    cx = 6.25 + c_idx * (pil_w2 + 0.233)
    cy = 1.75 + r_idx * (pil_h2 + 0.15)
    add_card(s13, cx, cy, pil_w2, pil_h2, bg_color=CARD_BG, border_color=CARD_BORDER)
    tb_cp = s13.shapes.add_textbox(Inches(cx + 0.15), Inches(cy + 0.12), Inches(pil_w2 - 0.3), Inches(pil_h2 - 0.24))
    tf_cp = tb_cp.text_frame
    tf_cp.word_wrap = True
    tf_cp.margin_left = tf_cp.margin_top = tf_cp.margin_right = tf_cp.margin_bottom = 0
    p0 = tf_cp.paragraphs[0]
    p0.text = tit
    p0.font.size = Pt(10.5)
    p0.font.bold = True
    p0.font.color.rgb = TEXT_WHITE
    p1 = tf_cp.add_paragraph()
    p1.text = desc
    p1.font.size = Pt(9.5)
    p1.font.color.rgb = TEXT_MUTED
    p1.space_before = Pt(2)

# Bottom Takeaway Box
add_card(s13, 0.8, 4.8, 11.733, 1.85, bg_color=CARD_BG, border_color=CARD_BORDER)
tb_concl = s13.shapes.add_textbox(Inches(1.0), Inches(4.95), Inches(11.333), Inches(1.55))
tf_concl = tb_concl.text_frame
tf_concl.word_wrap = True
tf_concl.margin_left = tf_concl.margin_top = tf_concl.margin_right = tf_concl.margin_bottom = 0
p0 = tf_concl.paragraphs[0]
p0.text = "EXECUTIVE CONCLUSION"
p0.font.size = Pt(10)
p0.font.bold = True
p0.font.color.rgb = CYAN
p1 = tf_concl.add_paragraph()
p1.text = "By integrating company-owned high-density vertical hubs with a verified distributed host network, the platform resolves both urban residential charging barriers and long-distance highway travel anxiety without capital duplication."
p1.font.size = Pt(10.5)
p1.font.color.rgb = TEXT_LIGHT
p1.space_before = Pt(2)
p2 = tf_concl.add_paragraph()
p2.text = "Charging infrastructure transitions from an inconvenient stationary destination into a seamless, coordinated network service: delivering high utilization for hosts and effortless convenience for EV owners."
p2.font.size = Pt(10)
p2.font.color.rgb = TEXT_MUTED
p2.space_before = Pt(3)

# Subnote
tb_prop = s13.shapes.add_textbox(Inches(0.8), Inches(6.75), Inches(11.733), Inches(0.3))
tf_prop = tb_prop.text_frame
tf_prop.word_wrap = True
tf_prop.margin_left = tf_prop.margin_top = tf_prop.margin_right = tf_prop.margin_bottom = 0
pp = tf_prop.paragraphs[0]
pp.text = "Proposed concept - technology, hardware and operational details should be validated during pilot implementation."
pp.font.size = Pt(9)
pp.font.color.rgb = TEXT_DIM
pp.alignment = PP_ALIGN.CENTER
add_footer(s13, 13)

# Save PowerPoint File
out_file = 'EV_Charging_Network_Presentation.pptx'
prs.save(out_file)
print(f"Successfully generated {out_file} with 13 widescreen slides and authentic photographic visuals!")
