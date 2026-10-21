import os
import pptx
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image

prs = pptx.Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# Theme Colors (Matching web deck exactly)
BG_COLOR = RGBColor(2, 6, 23)        # #020617 (slate-950)
CARD_BG = RGBColor(15, 23, 42)       # #0f172a (slate-900)
CARD_BORDER = RGBColor(30, 41, 59)   # #1e293b (slate-800)
CARD_ACTIVE = RGBColor(8, 51, 68)    # #083344 (cyan-950)
BORDER_ACTIVE = RGBColor(6, 182, 212)# #06b6d4 (cyan-500)

CYAN = RGBColor(34, 211, 238)        # #22d3ee (cyan-400)
CYAN_LIGHT = RGBColor(103, 232, 249) # #67e8f9 (cyan-300)
EMERALD = RGBColor(52, 211, 153)     # #34d399 (emerald-400)
EMERALD_BG = RGBColor(6, 78, 59)     # #064e3b (emerald-900)
AMBER = RGBColor(251, 191, 36)       # #fbbf24 (amber-400)
AMBER_BG = RGBColor(69, 26, 3)       # #451a03 (amber-950)
ROSE = RGBColor(248, 113, 113)       # #f87171 (rose-400)
ROSE_BG = RGBColor(76, 5, 25)        # #4c0519 (rose-950)

TEXT_WHITE = RGBColor(255, 255, 255) # white
TEXT_LIGHT = RGBColor(203, 213, 225) # slate-300
TEXT_MUTED = RGBColor(148, 163, 184) # slate-400
TEXT_DIM = RGBColor(100, 116, 139)   # slate-500
FOOTER_TEXT = RGBColor(71, 85, 105)  # slate-600

CROPPED_DIR = 'images/pptx_cropped'
os.makedirs(CROPPED_DIR, exist_ok=True)

def get_cropped_image(image_path, target_w_inch, target_h_inch):
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

def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=CARD_BORDER, border_width=1.0):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(border_width)
    return card

def add_button(slide, left, top, width, height, tag, label, is_active=False, active_col=CYAN, active_bg=CARD_ACTIVE, target_slide=None):
    bg = active_bg if is_active else CARD_BG
    br = active_col if is_active else CARD_BORDER
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = bg
    card.line.color.rgb = br
    card.line.width = Pt(1.5 if is_active else 1.0)
    
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.1)
    tf.margin_right = Inches(0.1)
    tf.margin_top = Inches(0.08)
    tf.margin_bottom = Inches(0.05)
    
    p0 = tf.paragraphs[0]
    p0.text = tag
    p0.font.size = Pt(8.5)
    p0.font.bold = True
    p0.font.color.rgb = active_col if is_active else TEXT_MUTED
    
    p1 = tf.add_paragraph()
    p1.text = label
    p1.font.size = Pt(10)
    p1.font.bold = True
    p1.font.color.rgb = TEXT_WHITE if is_active else TEXT_LIGHT
    
    if target_slide is not None:
        card.click_action.target_slide = target_slide
    return card

def add_photo_with_badge(slide, img_path, left, top, width, height, badge_title=None, badge_sub=None, badge_color=CYAN):
    cropped_path = get_cropped_image(img_path, width, height)
    slide.shapes.add_picture(cropped_path, Inches(left), Inches(top), width=Inches(width), height=Inches(height))
    
    frame = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    frame.fill.background()
    frame.line.color.rgb = CARD_BORDER
    frame.line.width = Pt(1.5)

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

def add_footer_nav(slide, current, total=13, prev_slide=None, next_slide=None, home_slide=None):
    # Status text
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(5.0), Inches(0.3))
    tf = tb.text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = f"EV CHARGING NETWORK : Slide {current:02d} of {total:02d}"
    p.font.size = Pt(9)
    p.font.color.rgb = FOOTER_TEXT
    
    # Interactive navigation pills on bottom right
    rx = 9.8
    if prev_slide is not None:
        btn_prev = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(rx), Inches(7.02), Inches(0.95), Inches(0.3))
        btn_prev.fill.solid()
        btn_prev.fill.fore_color.rgb = CARD_BG
        btn_prev.line.color.rgb = CARD_BORDER
        btn_prev.text_frame.text = "< Prev"
        btn_prev.text_frame.paragraphs[0].font.size = Pt(8.5)
        btn_prev.text_frame.paragraphs[0].font.color.rgb = TEXT_LIGHT
        btn_prev.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        btn_prev.click_action.target_slide = prev_slide
        rx += 1.05

    if next_slide is not None:
        btn_next = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(rx), Inches(7.02), Inches(0.95), Inches(0.3))
        btn_next.fill.solid()
        btn_next.fill.fore_color.rgb = CARD_BG
        btn_next.line.color.rgb = CARD_BORDER
        btn_next.text_frame.text = "Next >"
        btn_next.text_frame.paragraphs[0].font.size = Pt(8.5)
        btn_next.text_frame.paragraphs[0].font.color.rgb = CYAN
        btn_next.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        btn_next.click_action.target_slide = next_slide
        rx += 1.05

    if home_slide is not None:
        btn_home = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(rx), Inches(7.02), Inches(0.65), Inches(0.3))
        btn_home.fill.solid()
        btn_home.fill.fore_color.rgb = CARD_BG
        btn_home.line.color.rgb = CARD_BORDER
        btn_home.text_frame.text = "Home"
        btn_home.text_frame.paragraphs[0].font.size = Pt(8.5)
        btn_home.text_frame.paragraphs[0].font.color.rgb = TEXT_MUTED
        btn_home.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        btn_home.click_action.target_slide = home_slide

print("Creating blank slides dictionary...")
slides = {}
# Slide keys:
# 1, 2, 3
# 4_1, 4_2, 4_3, 4_4, 4_5 (Urban 5 steps)
# 5_1, 5_2, 5_3, 5_4, 5_5 (Highway 5 phases)
# 6_1, 6_2, 6_3 (Trust 3 zones)
# 7_1, 7_2 (Access 2 modes)
# 8_1, 8_2 (Verification 2 states)
# 9_1, 9_2, 9_3 (Monitoring 3 telemetry states)
# 10, 11, 12, 13

keys_order = [
    '1', '2', '3',
    '4_1', '4_2', '4_3', '4_4', '4_5',
    '5_1', '5_2', '5_3', '5_4', '5_5',
    '6_1', '6_2', '6_3',
    '7_1', '7_2',
    '8_1', '8_2',
    '9_1', '9_2', '9_3',
    '10', '11', '12', '13'
]

for k in keys_order:
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    slides[k] = s

# ==============================================================================
# SLIDE 1: TITLE SLIDE (WITH INTERACTIVE JUMP BUTTONS)
# ==============================================================================
s1 = slides['1']
tb1 = s1.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(6.5), Inches(2.2))
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
add_card(s1, 0.8, 2.9, 6.4, 0.7, bg_color=CARD_BG, border_color=CARD_BORDER)
tb_q = s1.shapes.add_textbox(Inches(1.0), Inches(3.02), Inches(6.0), Inches(0.45))
tf_q = tb_q.text_frame
tf_q.word_wrap = True
tf_q.margin_left = tf_q.margin_top = tf_q.margin_right = tf_q.margin_bottom = 0
pq = tf_q.paragraphs[0]
pq.text = '"YOU DON\'T NEED TO FIND A CHARGER. WE BRING CHARGING TO YOU."'
pq.font.size = Pt(11)
pq.font.bold = True
pq.font.color.rgb = CYAN_LIGHT

# Interactive Clickable Model Cards
# Model 1 Button (Jumps to 4_1)
btn_m1 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(3.8), Inches(3.1), Inches(2.1))
btn_m1.fill.solid()
btn_m1.fill.fore_color.rgb = CARD_BG
btn_m1.line.color.rgb = CYAN
btn_m1.line.width = Pt(1.5)
btn_m1.click_action.target_slide = slides['4_1']
tf_m1 = btn_m1.text_frame
tf_m1.word_wrap = True
tf_m1.margin_left = Inches(0.15)
tf_m1.margin_right = Inches(0.15)
tf_m1.margin_top = Inches(0.15)
p = tf_m1.paragraphs[0]
p.text = "MODEL 1 : URBAN [CLICK TO VIEW]"
p.font.size = Pt(9)
p.font.bold = True
p.font.color.rgb = CYAN
p_tit = tf_m1.add_paragraph()
p_tit.text = "Charging Hub + Valet"
p_tit.font.size = Pt(13)
p_tit.font.bold = True
p_tit.font.color.rgb = TEXT_WHITE
p_tit.space_before = Pt(2)
p_tit.space_after = Pt(4)
p_desc = tf_m1.add_paragraph()
p_desc.text = "Vertical parking towers with dedicated chargers & solar. Scheduled evening pickup with morning return."
p_desc.font.size = Pt(9.5)
p_desc.font.color.rgb = TEXT_LIGHT

# Model 2 Button (Jumps to 5_1)
btn_m2 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.1), Inches(3.8), Inches(3.1), Inches(2.1))
btn_m2.fill.solid()
btn_m2.fill.fore_color.rgb = CARD_BG
btn_m2.line.color.rgb = CARD_BORDER
btn_m2.line.width = Pt(1.5)
btn_m2.click_action.target_slide = slides['5_1']
tf_m2 = btn_m2.text_frame
tf_m2.word_wrap = True
tf_m2.margin_left = Inches(0.15)
tf_m2.margin_right = Inches(0.15)
tf_m2.margin_top = Inches(0.15)
p = tf_m2.paragraphs[0]
p.text = "MODEL 2 : HIGHWAY [CLICK TO VIEW]"
p.font.size = Pt(9)
p.font.bold = True
p.font.color.rgb = TEXT_MUTED
p_tit = tf_m2.add_paragraph()
p_tit.text = "Secure Distributed"
p_tit.font.size = Pt(13)
p_tit.font.bold = True
p_tit.font.color.rgb = TEXT_WHITE
p_tit.space_before = Pt(2)
p_tit.space_after = Pt(4)
p_desc = tf_m2.add_paragraph()
p_desc.text = "Verified host network with smart lock & virtual security. Monetizes private residential and business chargers."
p_desc.font.size = Pt(9.5)
p_desc.font.color.rgb = TEXT_LIGHT

# Interaction Helper Box
add_card(s1, 0.8, 6.05, 6.4, 0.6, bg_color=CARD_BG, border_color=CARD_BORDER)
tb_h = s1.shapes.add_textbox(Inches(1.0), Inches(6.15), Inches(6.0), Inches(0.4))
tf_h = tb_h.text_frame
tf_h.margin_left = tf_h.margin_top = tf_h.margin_right = tf_h.margin_bottom = 0
ph = tf_h.paragraphs[0]
ph.text = "Interactive Deck: In Slide Show mode (F5), click on cards and buttons to watch demonstrations update."
ph.font.size = Pt(9.5)
ph.font.color.rgb = CYAN_LIGHT

# Right Hero Photo with Badge
add_photo_with_badge(
    s1, 'images/ev_charging_hero.jpg',
    left=7.5, top=0.8, width=5.033, height=5.85,
    badge_title="Infrastructure Vision",
    badge_sub="Coordinated automated charging for urban & highway mobility."
)
add_footer_nav(s1, 1, 13, next_slide=slides['2'])


# ==============================================================================
# SLIDE 2: 01 : PROBLEM
# ==============================================================================
s2 = slides['2']
add_header(s2, "01 : PROBLEM", "The problem in one picture", "Empirical friction in urban EV ownership and long-distance travel")

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

add_photo_with_badge(
    s2, 'images/public_charger_queue.jpg',
    left=7.5, top=1.6, width=5.033, height=4.25,
    badge_title="Highway Range Gap",
    badge_sub="Long-distance corridors have charging voids while private chargers sit idle."
)

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
add_footer_nav(s2, 2, 13, prev_slide=slides['1'], next_slide=slides['3'], home_slide=slides['1'])


# ==============================================================================
# SLIDE 3: 02 : SOLUTION
# ==============================================================================
s3 = slides['3']
add_header(s3, "02 : SOLUTION", "One platform. Two charging models.", "Unifying high-density company-owned hubs with an asset-light distributed network")

# Model 1 Card (Clickable to jump to 4_1)
c1 = add_card(s3, 0.8, 1.6, 5.7, 5.25, bg_color=CARD_BG, border_color=CYAN, border_width=1.5)
c1.click_action.target_slide = slides['4_1']
add_photo_with_badge(s3, 'images/vertical_parking_tower.jpg', left=1.0, top=1.8, width=5.3, height=1.9, badge_title="Company-Owned Infrastructure [Click to inspect]")
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
p_foot1.text = "Valet Concierge: Scheduled evening pickup and 7:00 AM delivery to owner bay. (Click card to inspect)"
p_foot1.font.size = Pt(9.5)
p_foot1.font.color.rgb = CYAN
p_foot1.space_before = Pt(4)

# Model 2 Card (Clickable to jump to 5_1)
c2 = add_card(s3, 6.833, 1.6, 5.7, 5.25, bg_color=CARD_BG, border_color=CARD_BORDER, border_width=1.5)
c2.click_action.target_slide = slides['5_1']
add_photo_with_badge(s3, 'images/residential_driveway_charger.jpg', left=7.033, top=1.8, width=5.3, height=1.9, badge_title="Asset-Light Host Hardware [Click to inspect]", badge_color=TEXT_MUTED)
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
p_foot2.text = "Distributed Scaling: Expands coverage across highways with zero construction CapEx. (Click card to inspect)"
p_foot2.font.size = Pt(9.5)
p_foot2.font.color.rgb = CYAN
p_foot2.space_before = Pt(4)
add_footer_nav(s3, 3, 13, prev_slide=slides['2'], next_slide=slides['4_1'], home_slide=slides['1'])


# ==============================================================================
# SLIDE 4: URBAN LIFECYCLE (5 INTERACTIVE STEPS: 4_1 to 4_5)
# ==============================================================================
urban_step_data = [
    ("STAGE 01 : BOOKING", "Owner Schedules Pickup Window via Mobile App",
     "Owner specifies target state of charge (80% or 100%) and selects 9:00 PM pickup with 7:00 AM delivery. Cloud platform reserves dedicated vertical hub stall and dispatches concierge valet.",
     "Status: Reserved  |  Target SoC: 80%  |  Route: Optimal Stacking Level 3"),
    ("STAGE 02 : VERIFIED PICKUP", "Concierge Valet Dispatched to Residence",
     "Trained, insured valet arrives at apartment garage during scheduled window. Conducts digital 360-degree walkaround scan, scans vehicle RFID tag, and drives EV to the hub.",
     "Valet ID: #V-882 (Verified)  |  Inspection: Complete  |  Transit ETA: 12 min"),
    ("STAGE 03 : TOWER STACKING", "Automated Car Lift Positions Vehicle in Bay",
     "Vehicle arrives at centralized hub and is positioned onto automated turntable. Mechanical car lift stacks vehicle vertically across multi-tier steel structure, tripling density.",
     "Lift Mechanism: Automated Pallet  |  Bay: Stall 24-B  |  Safety Latch: Engaged"),
    ("STAGE 04 : CHARGE & STORAGE", "Dedicated 22kW AC Port Delivers Overnight Charge",
     "Dedicated 22kW port charges battery at manufacturer-recommended rate. Rooftop solar generation and on-site BESS buffer grid peaks to capture green power savings.",
     "Charge Rate: 22 kW AC  |  Battery Level: 42% -> 80%  |  Peak Shaving: Active"),
    ("STAGE 05 : MORNING RETURN", "Vehicle Returned Fully Charged by 7:00 AM",
     "Valet retrieves car from vertical tower at 6:30 AM and parks it in owner's assigned stall before 7:00 AM. Driver receives mobile notification with session evidence log.",
     "Delivery Status: Completed 6:45 AM  |  Final SoC: 80% (310 mi)  |  Valet Signed: Yes")
]

u_step_keys = ['4_1', '4_2', '4_3', '4_4', '4_5']
step_labels = [
    ("01 : BOOK", "Schedule"),
    ("02 : PICKUP", "Verified Valet"),
    ("03 : TOWER", "Stack Parking"),
    ("04 : CHARGE", "Solar + 22kW"),
    ("05 : RETURN", "Morning Drop")
]

for s_idx, k in enumerate(u_step_keys):
    cur_s = slides[k]
    add_header(cur_s, "03 : URBAN", "Model 1 - Charging Hub + Valet", "Interactive Lifecycle Demonstration: Click steps to inspect operations")
    
    # 5 Interactive Stepper Buttons
    btn_w = 2.25
    for b_idx, (tag, lbl) in enumerate(step_labels):
        bx = 0.8 + b_idx * (btn_w + 0.12)
        is_act = (b_idx == s_idx)
        add_button(
            cur_s, bx, 1.55, btn_w, 0.7,
            tag, lbl, is_active=is_act, active_col=CYAN,
            target_slide=slides[u_step_keys[b_idx]]
        )
    
    # Photo on Left
    add_photo_with_badge(
        cur_s, 'images/vertical_parking_tower.jpg',
        left=0.8, top=2.45, width=4.8, height=2.6,
        badge_title="Tower Stacking System",
        badge_sub="Automated car lift positions vehicle in dedicated stall."
    )
    
    # Dynamic Operation Details on Right
    st_tag, st_tit, st_desc, st_met = urban_step_data[s_idx]
    add_card(cur_s, 5.8, 2.45, 6.733, 2.6, bg_color=CARD_BG, border_color=CARD_BORDER)
    tb_op = cur_s.shapes.add_textbox(Inches(6.0), Inches(2.6), Inches(6.333), Inches(2.3))
    tf_op = tb_op.text_frame
    tf_op.word_wrap = True
    tf_op.margin_left = tf_op.margin_top = tf_op.margin_right = tf_op.margin_bottom = 0
    p = tf_op.paragraphs[0]
    p.text = st_tag
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = CYAN
    p_tit = tf_op.add_paragraph()
    p_tit.text = st_tit
    p_tit.font.size = Pt(13.5)
    p_tit.font.bold = True
    p_tit.font.color.rgb = TEXT_WHITE
    p_tit.space_before = Pt(2)
    p_tit.space_after = Pt(4)
    p_desc = tf_op.add_paragraph()
    p_desc.text = st_desc
    p_desc.font.size = Pt(10.5)
    p_desc.font.color.rgb = TEXT_LIGHT
    p_desc.space_after = Pt(8)

    p_met = tf_op.add_paragraph()
    p_met.text = st_met
    p_met.font.size = Pt(9.5)
    p_met.font.bold = True
    p_met.font.color.rgb = CYAN_LIGHT

    # 3 Pillars on Bottom
    pillars = [
        ("VERTICAL PARKING:", "3x land density through multi-tier automated stacking."),
        ("DEDICATED CHARGERS:", "Each parking stall wired with dedicated 22kW AC port."),
        ("SOLAR POTENTIAL:", "Rooftop generation captures clean power for peak shaving.")
    ]
    pil_w = 3.8
    for i, (title, text) in enumerate(pillars):
        px = 0.8 + i * (pil_w + 0.16)
        add_card(cur_s, px, 5.25, pil_w, 1.45, bg_color=CARD_BG, border_color=CARD_BORDER)
        tb_p = cur_s.shapes.add_textbox(Inches(px + 0.2), Inches(5.4), Inches(pil_w - 0.4), Inches(1.1))
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
    add_footer_nav(cur_s, 4, 13, prev_slide=slides['3'], next_slide=slides['5_1'], home_slide=slides['1'])


# ==============================================================================
# SLIDE 5: HIGHWAY WORKFLOW (5 INTERACTIVE PHASES: 5_1 to 5_5)
# ==============================================================================
highway_phase_data = [
    ("PHASE 01 : LISTING", "Host Publishes Availability Window & Tariff",
     "Homeowner configures idle hours (e.g., 9:00 AM - 5:00 PM). Smart controller links to platform with zero manual host intervention. Platform coordinates identity, payment escrow, and charger solenoid unlocking.",
     "Hardware: 240V / 32A  |  Availability: Mon-Fri 8 hrs  |  Rate: Tariff + Margin"),
    ("PHASE 02 : ROUTE RESERVATION", "Highway Driver Discovers & Books En-Route Bay",
     "Long-distance traveler identifies verified host along transit corridor via navigation app. Pre-reserves 45-minute slot. Platform holds charge fee in escrow.",
     "Reservation ID: #HW-491  |  Arrival Window: 1:30 PM  |  Escrow Deposit: $12.50 Held"),
    ("PHASE 03 : CRYPTOGRAPHIC AUTH", "Driver Arrives & Initiates Bluetooth Unlock",
     "Driver arrives at host driveway. Mobile app connects to charger lock controller via Bluetooth Low Energy (BLE) or NFC. Signed digital key verifies authorization.",
     "Auth Method: BLE Token 256-bit  |  Holster Latch: Unlocked  |  Proximity: 1.2m"),
    ("PHASE 04 : AUTONOMOUS CHARGE", "Self-Service Plug-in & Telemetry Streaming",
     "Driver plugs vehicle in without disturbing host. Smart controller engages main contactor, monitors ground leakage, and streams 1 Hz telemetry to mobile app.",
     "Power: 7.4 kW  |  Delivered: 14.8 kWh  |  GFCI: Normal  |  Session: In Progress"),
    ("PHASE 05 : AUTO-LATCH & ESCROW", "Session Concluded, Cable Re-Latched, Payout Cleared",
     "Driver unplugs, re-holsters connector into smart dock where solenoid lock re-engages. Optical inspection validates site condition; escrow clears automatically.",
     "Holster Latch: Locked  |  Departure Verified: Yes  |  Host Payout: $9.80 Cleared")
]

h_step_keys = ['5_1', '5_2', '5_3', '5_4', '5_5']
h_step_labels = [
    ("01 : LIST", "Host Node"),
    ("02 : BOOK", "Route Reserve"),
    ("03 : AUTH", "BLE Token"),
    ("04 : CHARGE", "Self-Plug"),
    ("05 : LOCK", "Auto-Latch")
]

for s_idx, k in enumerate(h_step_keys):
    cur_s = slides[k]
    add_header(cur_s, "04 : HIGHWAY", "Model 2 - Secure Distributed Charging", "Turn unused charging capacity into verified network capacity.")
    
    for b_idx, (tag, lbl) in enumerate(h_step_labels):
        bx = 0.8 + b_idx * (btn_w + 0.12)
        is_act = (b_idx == s_idx)
        add_button(
            cur_s, bx, 1.55, btn_w, 0.7,
            tag, lbl, is_active=is_act, active_col=CYAN,
            target_slide=slides[h_step_keys[b_idx]]
        )

    add_photo_with_badge(
        cur_s, 'images/residential_driveway_charger.jpg',
        left=0.8, top=2.45, width=4.8, height=2.6,
        badge_title="Verified Host Hardware",
        badge_sub="Residential Level 2 charger unlocked for road trip stopovers."
    )

    st_tag, st_tit, st_desc, st_met = highway_phase_data[s_idx]
    add_card(cur_s, 5.8, 2.45, 6.733, 2.6, bg_color=CARD_BG, border_color=CARD_BORDER)
    tb_hp = cur_s.shapes.add_textbox(Inches(6.0), Inches(2.6), Inches(6.333), Inches(2.3))
    tf_hp = tb_hp.text_frame
    tf_hp.word_wrap = True
    tf_hp.margin_left = tf_hp.margin_top = tf_hp.margin_right = tf_hp.margin_bottom = 0
    p = tf_hp.paragraphs[0]
    p.text = st_tag
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = CYAN
    p_tit = tf_hp.add_paragraph()
    p_tit.text = st_tit
    p_tit.font.size = Pt(13.5)
    p_tit.font.bold = True
    p_tit.font.color.rgb = TEXT_WHITE
    p_tit.space_before = Pt(2)
    p_tit.space_after = Pt(4)
    p_desc = tf_hp.add_paragraph()
    p_desc.text = st_desc
    p_desc.font.size = Pt(10.5)
    p_desc.font.color.rgb = TEXT_LIGHT
    p_desc.space_after = Pt(8)

    p_met = tf_hp.add_paragraph()
    p_met.text = st_met
    p_met.font.size = Pt(9.5)
    p_met.font.bold = True
    p_met.font.color.rgb = CYAN_LIGHT

    add_card(cur_s, 0.8, 5.25, 11.733, 1.45, bg_color=CARD_BG, border_color=CARD_BORDER)
    tb_crit = cur_s.shapes.add_textbox(Inches(1.0), Inches(5.4), Inches(11.333), Inches(1.1))
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
    add_footer_nav(cur_s, 5, 13, prev_slide=slides['4_1'], next_slide=slides['6_1'], home_slide=slides['1'])


# ==============================================================================
# SLIDE 6: TRUST & SECURITY (3 INTERACTIVE ZONES: 6_1 to 6_3)
# ==============================================================================
trust_zone_data = [
    ("ACTIVE ZONE : GREEN (SAFE)", "Driveway Staging Bay & Charger Port Area",
     "Vehicle parked in bay. Charging cable connected. No unauthorized proximity detected. Host family perimeter is untouched. Optical AI verifies human presence is restricted exclusively to vehicle charging port.",
     "Perimeter Status: Normal  |  Audio Alert: Inactive  |  Solenoid Lock: Engaged",
     "CAM 01 : GREEN SAFE ZONE  |  REC 1080P", "Vehicle in Bay : Normal Charging Allowed", EMERALD),
    ("ACTIVE ZONE : YELLOW (WARNING BUFFER)", "3-Meter Buffer Zone Around Driveway Stall",
     "Pedestrian motion detected near driveway boundary. External smart speaker emits friendly audio guidance: 'Please remain near your vehicle during charging.' Strobe pulse alerts host app silently.",
     "Perimeter Status: Buffer Alert  |  Audio Alert: Active (Warning)  |  Solenoid Lock: Engaged",
     "CAM 01 : YELLOW BUFFER ZONE  |  WARNING 3M", "Driver Outside Stall : Polite Audio Guidance Emitted", AMBER),
    ("ACTIVE ZONE : RED (RESTRICTED RESIDENCE)", "Private Home Entryway & Backyard Perimeter",
     "Unauthorized presence detected near private residence doors. System immediately trips power contactor, triggers strobe alarm, and alerts host and local security dispatch.",
     "Perimeter Status: CRITICAL BREACH  |  EVSE Contactor: DISENGAGED  |  Alert: Security Dispatched",
     "CAM 01 : RED RESTRICTED  |  BREACH ALARM", "Boundary Crossed : Contactor Tripped & Alarm Sounded", ROSE)
]

z_keys = ['6_1', '6_2', '6_3']
z_buttons = [
    ("ZONE 1 : GREEN", "Safe Authorized Stall", EMERALD, EMERALD_BG),
    ("ZONE 2 : YELLOW", "Warning Buffer (3m)", AMBER, AMBER_BG),
    ("ZONE 3 : RED", "Restricted Residence", ROSE, ROSE_BG)
]
z_w = 3.8

for s_idx, k in enumerate(z_keys):
    cur_s = slides[k]
    add_header(cur_s, "05 : TRUST", "Security: a charger is not an open door", "Interactive Virtual Perimeter: Click zones to test CCTV monitoring states")
    
    for b_idx, (tag, lbl, col, col_bg) in enumerate(z_buttons):
        zx = 0.8 + b_idx * (z_w + 0.16)
        is_act = (b_idx == s_idx)
        add_button(
            cur_s, zx, 1.55, z_w, 0.7,
            tag, lbl, is_active=is_act, active_col=col, active_bg=col_bg,
            target_slide=slides[z_keys[b_idx]]
        )

    st_tag, st_tit, st_desc, st_met, hud_tag, hud_desc, hud_col = trust_zone_data[s_idx]
    add_photo_with_badge(
        cur_s, 'images/cctv_security_driveway.jpg',
        left=0.8, top=2.45, width=5.5, height=2.6,
        badge_title=hud_tag,
        badge_sub=hud_desc,
        badge_color=hud_col
    )

    add_card(cur_s, 6.5, 2.45, 6.033, 2.6, bg_color=CARD_BG, border_color=CARD_BORDER)
    tb_sec = cur_s.shapes.add_textbox(Inches(6.7), Inches(2.6), Inches(5.633), Inches(2.3))
    tf_sec = tb_sec.text_frame
    tf_sec.word_wrap = True
    tf_sec.margin_left = tf_sec.margin_top = tf_sec.margin_right = tf_sec.margin_bottom = 0
    p = tf_sec.paragraphs[0]
    p.text = st_tag
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = hud_col
    p_tit = tf_sec.add_paragraph()
    p_tit.text = st_tit
    p_tit.font.size = Pt(13)
    p_tit.font.bold = True
    p_tit.font.color.rgb = TEXT_WHITE
    p_tit.space_before = Pt(2)
    p_tit.space_after = Pt(4)
    p_desc = tf_sec.add_paragraph()
    p_desc.text = st_desc
    p_desc.font.size = Pt(10)
    p_desc.font.color.rgb = TEXT_LIGHT
    p_desc.space_after = Pt(8)

    p_met = tf_sec.add_paragraph()
    p_met.text = st_met
    p_met.font.size = Pt(9.5)
    p_met.font.bold = True
    p_met.font.color.rgb = hud_col

    add_card(cur_s, 0.8, 5.25, 11.733, 1.45, bg_color=CARD_BG, border_color=CARD_BORDER)
    tb_secbot = cur_s.shapes.add_textbox(Inches(1.0), Inches(5.4), Inches(11.333), Inches(1.1))
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
    add_footer_nav(cur_s, 6, 13, prev_slide=slides['5_1'], next_slide=slides['7_1'], home_slide=slides['1'])


# ==============================================================================
# SLIDE 7: ACCESS CONTROL (2 INTERACTIVE MODES: 7_1, 7_2)
# ==============================================================================
access_mode_data = [
    ("FIRMWARE MODE : OWNER MODE", "Owner NFC / Smartphone Bypass",
     "Owner registers master keycard or smartphone. Instant proximity tap releases holster latch. No billing fees or reservations required for personal electricity usage. System switches automatically to Guest Mode when reservation begins.",
     "Authorization: Master Key  |  Billing: Direct Home  |  Latch: Instant Tap",
     "Owner Master NFC Bypass: Instant latch release for household residents."),
    ("FIRMWARE MODE : TRAVELLER GUEST MODE", "Time-Bound Cryptographic Reservation Token",
     "Transient traveller verifies via platform mobile app. Cloud broker generates 60-minute signed JWT token over BLE. Power flow metered for automated billing and host payout.",
     "Authorization: JWT 256-Bit  |  Billing: Platform Escrow  |  Latch: Booked Window",
     "Traveller Cryptographic Token: Ephemeral BLE session handshake.")
]

acc_keys = ['7_1', '7_2']
acc_buttons = [
    ("MODE 01 : OWNER", "Owner Authenticates -> Uses -> Locks"),
    ("MODE 02 : TRAVELLER", "Temporary Token Authorization -> Locks")
]
m_w = 5.75

for s_idx, k in enumerate(acc_keys):
    cur_s = slides[k]
    add_header(cur_s, "06 : ACCESS", "Smart charger control", "Hardware-level access security and dual-mode separation")
    
    for b_idx, (tag, lbl) in enumerate(acc_buttons):
        mx = 0.8 + b_idx * (m_w + 0.233)
        is_act = (b_idx == s_idx)
        add_button(
            cur_s, mx, 1.55, m_w, 0.7,
            tag, lbl, is_active=is_act, active_col=CYAN,
            target_slide=slides[acc_keys[b_idx]]
        )

    st_tag, st_tit, st_desc, st_met, photo_sub = access_mode_data[s_idx]
    add_photo_with_badge(
        cur_s, 'images/smart_phone_nfc_charger.jpg',
        left=0.8, top=2.45, width=4.8, height=2.6,
        badge_title="NFC & BLE Tap Interface",
        badge_sub=photo_sub
    )

    add_card(cur_s, 5.8, 2.45, 6.733, 2.6, bg_color=CARD_BG, border_color=CARD_BORDER)
    tb_acc = cur_s.shapes.add_textbox(Inches(6.0), Inches(2.6), Inches(6.333), Inches(2.3))
    tf_acc = tb_acc.text_frame
    tf_acc.word_wrap = True
    tf_acc.margin_left = tf_acc.margin_top = tf_acc.margin_right = tf_acc.margin_bottom = 0
    p = tf_acc.paragraphs[0]
    p.text = st_tag
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = CYAN
    p_tit = tf_acc.add_paragraph()
    p_tit.text = st_tit
    p_tit.font.size = Pt(13.5)
    p_tit.font.bold = True
    p_tit.font.color.rgb = TEXT_WHITE
    p_tit.space_before = Pt(2)
    p_tit.space_after = Pt(4)
    p_desc = tf_acc.add_paragraph()
    p_desc.text = st_desc
    p_desc.font.size = Pt(10.5)
    p_desc.font.color.rgb = TEXT_LIGHT
    p_desc.space_after = Pt(8)

    p_met = tf_acc.add_paragraph()
    p_met.text = st_met
    p_met.font.size = Pt(9.5)
    p_met.font.bold = True
    p_met.font.color.rgb = CYAN_LIGHT

    add_card(cur_s, 0.8, 5.25, 11.733, 1.45, bg_color=CARD_BG, border_color=CARD_BORDER)
    tb_accbot = cur_s.shapes.add_textbox(Inches(1.0), Inches(5.4), Inches(11.333), Inches(1.1))
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
    add_footer_nav(cur_s, 7, 13, prev_slide=slides['6_1'], next_slide=slides['8_1'], home_slide=slides['1'])


# ==============================================================================
# SLIDE 8: EVIDENCE & VERIFICATION (2 INTERACTIVE STATES: 8_1, 8_2)
# ==============================================================================
v_states = [
    ("Pre-Session Baseline Scanned & Standby",
     "Algorithms await post-session photo capture to execute 16-point feature diff across baseline and departure imagery.",
     "[ Click to Run CV Diff Analysis ]", CYAN, False),
    ("16-Point Computer Vision Diff Complete",
     "Comparison verified: zero scratches, zero cable sheath abrasions, zero pin oxidation. Host equipment protected. Escrow payment released.",
     "[ Reset Scan ]", EMERALD, True)
]

for s_idx, k in enumerate(['8_1', '8_2']):
    cur_s = slides[k]
    add_header(cur_s, "07 : VERIFICATION", "Every session leaves evidence", "Tripartite photographic inspection workbench & computer vision comparison")
    
    st_tit, st_desc, btn_lbl, status_col, is_done = v_states[s_idx]
    
    badge_t = "Port Latch & Contact Verification" if not is_done else "Verification Complete (100% Passed)"
    badge_s = "Automated high-resolution macro visual evidence logging." if not is_done else "Zero cosmetic scratches or contact wear detected."
    add_photo_with_badge(
        cur_s, 'images/ev_port_inspection.jpg',
        left=0.8, top=1.6, width=4.8, height=3.45,
        badge_title=badge_t,
        badge_sub=badge_s,
        badge_color=status_col
    )

    steps_v = [
        ("[ 01 : BEFORE ] Pre-Session Baseline Scan", "Vehicle port flap, connector pins, cable sheath, and driveway baseline recorded prior to charging."),
        ("[ 02 : CHARGING ] Continuous Session Logging", "1 Hz telemetry stream + CCTV perimeter motion clip recording throughout the active charging session."),
        ("[ 03 : AFTER ] Post-Session Diff", "Repeat matching photos + computer vision condition comparison to verify zero damage or cable strain.")
    ]
    vy = 1.6
    for title, text in steps_v:
        add_card(cur_s, 5.8, vy, 6.733, 1.05, bg_color=CARD_BG, border_color=CARD_BORDER)
        tb_v = cur_s.shapes.add_textbox(Inches(6.0), Inches(vy + 0.1), Inches(6.333), Inches(0.85))
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

    # Bottom Workbench Card with Clickable Trigger
    add_card(cur_s, 0.8, 5.25, 11.733, 1.45, bg_color=CARD_BG, border_color=CARD_BORDER)
    tb_vbot = cur_s.shapes.add_textbox(Inches(1.0), Inches(5.4), Inches(8.0), Inches(1.1))
    tf_vbot = tb_vbot.text_frame
    tf_vbot.word_wrap = True
    tf_vbot.margin_left = tf_vbot.margin_top = tf_vbot.margin_right = tf_vbot.margin_bottom = 0
    p0 = tf_vbot.paragraphs[0]
    p0.text = "AUTOMATED AI DIFFERENCE VERIFICATION"
    p0.font.size = Pt(9.5)
    p0.font.bold = True
    p0.font.color.rgb = status_col
    p1 = tf_vbot.add_paragraph()
    p1.text = st_tit
    p1.font.size = Pt(12)
    p1.font.bold = True
    p1.font.color.rgb = TEXT_WHITE
    p1.space_before = Pt(1)
    p1.space_after = Pt(2)
    p2 = tf_vbot.add_paragraph()
    p2.text = st_desc
    p2.font.size = Pt(10)
    p2.font.color.rgb = TEXT_LIGHT

    # Clickable Test Button
    target_k = '8_2' if not is_done else '8_1'
    btn_scan = cur_s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.2), Inches(5.55), Inches(3.1), Inches(0.85))
    btn_scan.fill.solid()
    btn_scan.fill.fore_color.rgb = CARD_ACTIVE if not is_done else CARD_BG
    btn_scan.line.color.rgb = status_col
    btn_scan.line.width = Pt(1.5)
    btn_scan.click_action.target_slide = slides[target_k]
    tf_bs = btn_scan.text_frame
    tf_bs.word_wrap = True
    tf_bs.margin_left = tf_bs.margin_right = Inches(0.1)
    tf_bs.margin_top = Inches(0.18)
    p_b0 = tf_bs.paragraphs[0]
    p_b0.text = btn_lbl
    p_b0.font.size = Pt(10)
    p_b0.font.bold = True
    p_b0.font.color.rgb = status_col
    p_b0.alignment = PP_ALIGN.CENTER
    add_footer_nav(cur_s, 8, 13, prev_slide=slides['7_1'], next_slide=slides['9_1'], home_slide=slides['1'])


# ==============================================================================
# SLIDE 9: LIVE MONITORING (3 INTERACTIVE TELEMETRY STATES: 9_1, 9_2, 9_3)
# ==============================================================================
telemetry_states = [
    ([
        ("TEMPERATURE", "32.4 deg C", "Thermocouple (Safe)", CYAN_LIGHT),
        ("VOLT / CURRENT", "230V / 32A", "Phase Health: Nominal", TEXT_WHITE),
        ("POWER DELIVERY", "7.36 kW", "Class 1 Meter Output", CYAN_LIGHT),
        ("FAULT CODE", "0x00 NORMAL", "GFCI State: Active Protected", EMERALD)
    ], "Normal charging: continuous 1 Hz telemetry stream monitoring voltage, current, and temperature."),
    ([
        ("TEMPERATURE", "72.8 deg C", "THERMAL WARNING EXCEEDED", ROSE),
        ("VOLT / CURRENT", "230V / 16A", "Amperage Throttled 50%", AMBER),
        ("POWER DELIVERY", "3.68 kW", "Protective Derating Active", AMBER),
        ("FAULT CODE", "0x11 TEMP_WARN", "Pilot Trip Warning Emitted", AMBER)
    ], "Thermal protection triggered: pilot signal automatically throttles amperage from 32A down to 16A to cool connector pins."),
    ([
        ("TEMPERATURE", "28.1 deg C", "System Disengaged", TEXT_MUTED),
        ("VOLT / CURRENT", "0V / 0A", "Contactor Disconnected (<10ms)", ROSE),
        ("POWER DELIVERY", "0.00 kW", "Power Flow Terminated", ROSE),
        ("FAULT CODE", "0xFA GFCI_TRIP", "Ground Leakage >5mA Tripped", ROSE)
    ], "Immediate sub-10ms safety trip: ground fault leakage detected. Hardware isolation contactor opened immediately; emergency push alert dispatched.")
]

t_buttons = [
    ("SIM 01 : NORMAL", "Normal 7.4 kW"),
    ("SIM 02 : THERMAL", "Thermal Rise (72 deg C)"),
    ("SIM 03 : GFCI TRIP", "Ground Fault Trip GFCI")
]
sim_w = 3.8

for s_idx, k in enumerate(['9_1', '9_2', '9_3']):
    cur_s = slides[k]
    add_header(cur_s, "08 : MONITORING", "Protect the charger and the vehicle", "Real-time sensor telemetry and automated fault isolation controls")
    
    # Simulation Injector Buttons
    for b_idx, (tag, lbl) in enumerate(t_buttons):
        bx = 0.8 + b_idx * (sim_w + 0.16)
        is_act = (b_idx == s_idx)
        act_col = CYAN if b_idx == 0 else (AMBER if b_idx == 1 else ROSE)
        act_bg = CARD_ACTIVE if b_idx == 0 else (AMBER_BG if b_idx == 1 else ROSE_BG)
        add_button(
            cur_s, bx, 1.55, sim_w, 0.7,
            tag, lbl, is_active=is_act, active_col=act_col, active_bg=act_bg,
            target_slide=slides[f'9_{b_idx+1}']
        )

    # Photo on Left
    add_photo_with_badge(
        cur_s, 'images/ev_telemetry_dashboard.jpg',
        left=0.8, top=2.45, width=4.8, height=2.6,
        badge_title="Live Telemetry Supervision",
        badge_sub="Continuous monitoring protects charger hardware and vehicle battery."
    )

    # Gauges
    gauges_data, note_text = telemetry_states[s_idx]
    gw = 3.25
    for i, (tag, val, sub, col) in enumerate(gauges_data):
        col_idx = i % 2
        row_idx = i // 2
        gx = 5.8 + col_idx * (gw + 0.233)
        gy = 2.45 + row_idx * 1.35
        add_card(cur_s, gx, gy, gw, 1.25, bg_color=CARD_BG, border_color=CARD_BORDER)
        tb_g = cur_s.shapes.add_textbox(Inches(gx + 0.15), Inches(gy + 0.12), Inches(gw - 0.3), Inches(1.0))
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

    # 4 Response Sequence Cards
    seq = [
        ("1. Detect anomaly", "<10ms polling on voltage, current & ground."),
        ("2. Safe EVSE response", "<100ms pilot trip disengages main contactor safely."),
        ("3. Notify", "Instant alerts delivered to driver and host devices."),
        ("4. Inspect if needed", "Remote diagnostics logs flagged for technical review.")
    ]
    sq_w = 2.8
    for i, (tit, desc) in enumerate(seq):
        sx = 0.8 + i * (sq_w + 0.177)
        add_card(cur_s, sx, 5.25, sq_w, 1.45, bg_color=CARD_BG, border_color=CARD_BORDER)
        tb_sq = cur_s.shapes.add_textbox(Inches(sx + 0.15), Inches(5.4), Inches(sq_w - 0.3), Inches(1.1))
        tf_sq = tb_sq.text_frame
        tf_sq.word_wrap = True
        tf_sq.margin_left = tf_sq.margin_top = tf_sq.margin_right = tf_sq.margin_bottom = 0
        p0 = tf_sq.paragraphs[0]
        p0.text = tit
        p0.font.size = Pt(11)
        p0.font.bold = True
        p0.font.color.rgb = CYAN if i == 1 else TEXT_WHITE
        p0.space_after = Pt(2)
        p1 = tf_sq.add_paragraph()
        p1.text = desc
        p1.font.size = Pt(9.5)
        p1.font.color.rgb = TEXT_LIGHT
    add_footer_nav(cur_s, 9, 13, prev_slide=slides['8_1'], next_slide=slides['10'], home_slide=slides['1'])


# ==============================================================================
# SLIDE 10: 09 : PLATFORM
# ==============================================================================
s10 = slides['10']
add_header(s10, "09 : PLATFORM", "The ecosystem", "Unifying demand, intelligence, and distributed electrical supply")

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
add_footer_nav(s10, 10, 13, prev_slide=slides['9_1'], next_slide=slides['11'], home_slide=slides['1'])


# ==============================================================================
# SLIDE 11: 10 : BUSINESS
# ==============================================================================
s11 = slides['11']
add_header(s11, "10 : BUSINESS", "Scale without building every charger", "Asset-right hybrid capital strategy balancing high-margin anchors with network scaling")

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
add_footer_nav(s11, 11, 13, prev_slide=slides['10'], next_slide=slides['12'], home_slide=slides['1'])


# ==============================================================================
# SLIDE 12: 11 : ROADMAP
# ==============================================================================
s12 = slides['12']
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
add_footer_nav(s12, 12, 13, prev_slide=slides['11'], next_slide=slides['13'], home_slide=slides['1'])


# ==============================================================================
# SLIDE 13: 12 : CONCLUSION
# ==============================================================================
s13 = slides['13']
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

add_photo_with_badge(
    s13, 'images/highway_scenic_ev.jpg',
    left=0.8, top=1.75, width=5.2, height=2.85,
    badge_title="Corridor Freedom",
    badge_sub="Connected highway route travel with verified host chargers."
)

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

# Restart Button on Slide 13
btn_rst = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.666), Inches(6.8), Inches(2.0), Inches(0.4))
btn_rst.fill.solid()
btn_rst.fill.fore_color.rgb = CARD_ACTIVE
btn_rst.line.color.rgb = CYAN
btn_rst.line.width = Pt(1.5)
btn_rst.text_frame.text = "Restart Deck (Slide 1)"
btn_rst.text_frame.paragraphs[0].font.size = Pt(9.5)
btn_rst.text_frame.paragraphs[0].font.bold = True
btn_rst.text_frame.paragraphs[0].font.color.rgb = CYAN
btn_rst.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
btn_rst.click_action.target_slide = slides['1']

add_footer_nav(s13, 13, 13, prev_slide=slides['12'], home_slide=slides['1'])

out_file = 'EV_Charging_Network_Presentation.pptx'
prs.save(out_file)
print(f"Successfully generated interactive presentation {out_file} with {len(prs.slides)} slides!")
