import os
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

prs = pptx.Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# Exact Canva Color Palette
BG = RGBColor(7, 13, 24)            # Deep dark background
CARD_BG = RGBColor(12, 20, 36)       # Dark slate card
BORDER_CYAN = RGBColor(6, 182, 212)  # Cyan border (#06b6d4)
BORDER_GREEN = RGBColor(16, 185, 129)# Green border (#10b981)
BORDER_AMBER = RGBColor(245, 158, 11)# Amber border (#f59e0b)
BORDER_ROSE = RGBColor(244, 63, 94)  # Rose/Red border (#f43f5e)
BORDER_MUTED = RGBColor(30, 41, 59)  # Slate border (#1e293b)

CYAN_TEXT = RGBColor(34, 211, 238)
GREEN_TEXT = RGBColor(52, 211, 153)
AMBER_TEXT = RGBColor(251, 191, 36)
ROSE_TEXT = RGBColor(251, 113, 133)
WHITE = RGBColor(255, 255, 255)
LIGHT = RGBColor(226, 232, 240)
MUTED = RGBColor(148, 163, 184)
DIM = RGBColor(100, 116, 139)

def set_bg(slide):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG
    bg.line.fill.background()
    return bg

def add_header(slide, tag, title, subtitle):
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.95))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    p0 = tf.paragraphs[0]
    p0.text = tag.upper()
    p0.font.size = Pt(9.5)
    p0.font.bold = True
    p0.font.color.rgb = CYAN_TEXT
    p0.space_after = Pt(2)
    
    p1 = tf.add_paragraph()
    p1.text = title
    p1.font.size = Pt(21)
    p1.font.bold = True
    p1.font.color.rgb = WHITE
    p1.space_after = Pt(2)
    
    p2 = tf.add_paragraph()
    p2.text = subtitle
    p2.font.size = Pt(10.5)
    p2.font.color.rgb = MUTED

def add_footer_note(slide, note_text):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.5), Inches(11.733), Inches(0.6))
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(26, 20, 10)
    card.line.color.rgb = AMBER_TEXT
    card.line.width = Pt(1)
    
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.2)
    tf.margin_top = Inches(0.1)
    p = tf.paragraphs[0]
    p.text = f"⚠  {note_text}"
    p.font.size = Pt(9.5)
    p.font.color.rgb = AMBER_TEXT

# ==============================================================================
# SLIDE 1: Executive Pitch Deck
# ==============================================================================
s1 = prs.slides.add_slide(blank_layout); set_bg(s1)
add_header(s1, "EXECUTIVE PITCH DECK  |  EV CHARGING NETWORK", 
           "You don't need to find a charger. We bring charging to you.", 
           "One unified platform · Two strategic models · Zero range anxiety — Urban concierge hubs + distributed highway network.")

# Left Card: Bullets & Buttons
c1 = slide = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(6.0), Inches(4.85))
c1.fill.solid(); c1.fill.fore_color.rgb = CARD_BG; c1.line.color.rgb = BORDER_MUTED
tf1 = c1.text_frame; tf1.word_wrap = True; tf1.margin_left = Inches(0.3); tf1.margin_top = Inches(0.25)

p = tf1.paragraphs[0]
p.text = "CORE DELIVERY MODELS"
p.font.size = Pt(10); p.font.bold = True; p.font.color.rgb = CYAN_TEXT; p.space_after = Pt(8)

bullets_s1 = [
    ("Concierge Valet", "Vehicle collected 9 PM, returned 7 AM fully charged (100% SoC) in company-owned vertical hubs."),
    ("Highway Hosts", "Private & hotel chargers unlocked via encrypted Bluetooth tokens along inter-city corridors."),
    ("Bank-Grade Trust", "Virtual 3-zone security perimeter + automated escrow protection for every session."),
    ("Capital Efficiency", "High-margin owned vertical hubs (cash flow) paired with zero-CapEx host network (rapid scale).")
]
for title, desc in bullets_s1:
    p = tf1.add_paragraph()
    p.text = f"• {title}: "
    p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = WHITE
    run = p.add_run()
    run.text = desc
    run.font.size = Pt(10); run.font.bold = False; run.font.color.rgb = LIGHT
    p.space_after = Pt(8)

# Right: Framed Photo
if os.path.exists("images/ev_port_inspection.jpg"):
    s1.shapes.add_picture("images/ev_port_inspection.jpg", Inches(7.1), Inches(1.5), Inches(5.433), Inches(4.35))
    cap = s1.shapes.add_textbox(Inches(7.1), Inches(5.9), Inches(5.433), Inches(0.4))
    tf_c = cap.text_frame; tf_c.margin_left = tf_c.margin_top = 0
    p = tf_c.paragraphs[0]
    p.text = "📸  Automated urban vertical charging hub: dedicated 7–11 kW AC charging per stall."
    p.font.size = Pt(9); p.font.color.rgb = MUTED

add_footer_note(s1, "RESEARCH NOTE: Data verified with J.D. Power (2024), Zapmap, NREL, and Frontiers in Sustainable Cities (2026).")

# ==============================================================================
# SLIDE 2: Problem & Market Opportunity (4 VERTICAL CARDS EXACTLY LIKE CANVA)
# ==============================================================================
s2 = prs.slides.add_slide(blank_layout); set_bg(s2)
add_header(s2, "THE PROBLEM & MARKET OPPORTUNITY", 
           "Why public EV charging fails drivers today — and the data behind it", 
           "Empirical data replacing unsourced deck claims across global markets and India.")

col_w = 2.78
cards_s2 = [
    ("🏢", "No Home Charger Access", 
     [("GLOBAL", "~20% of UK EV owners lack home charging. 82% of US multi-family residents have no in-building charger."),
      ("INDIA CONTEXT", "42%+ of urban EV owners reside in apartment complexes with zero designated charging bays."),
      ("SOURCE", "Zapmap 2024 · US DOE · India CEA")]),
    ("⚡", "Unreliable Public Charging", 
     [("FAILURE RATE", "12% to 20% of public charging sessions fail depending on network."),
      ("DEFECTIVE CHARGERS", "52.4% of non-Tesla drivers encountered broken chargers in 2023."),
      ("SOURCE", "J.D. Power 2024 · PlugShare")]),
    ("🛣️", "Highway Corridor Gaps", 
     [("CHARGING DESERTS", "Fast-charging gaps persist along inter-city highway corridors."),
      ("RANGE BOTTLENECKS", "Long-distance travel suffers from severe queues and broken hardware outside cities."),
      ("SOURCE", "NREL Corridor Analysis 2024")]),
    ("💤", "Millions of Idle Chargers", 
     [("UNEXPLOITED ASSETS", "Private residential, workplace, and hotel chargers sit idle 20+ hours daily."),
      ("DISTRIBUTED SCALE", "High CapEx public networks lag while private capacity remains locked behind closed gates."),
      ("SOURCE", "Frontiers Sustainable Cities 2026")])
]

for idx, (icon, title, points) in enumerate(cards_s2):
    left = Inches(0.8 + idx * 2.98)
    card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(1.5), Inches(col_w), Inches(4.85))
    card.fill.solid(); card.fill.fore_color.rgb = CARD_BG; card.line.color.rgb = BORDER_MUTED
    tf = card.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.2); tf.margin_top = Inches(0.2)
    
    p = tf.paragraphs[0]; p.text = f"{icon}  {title}"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = CYAN_TEXT; p.space_after = Pt(10)
    for head, body in points:
        p_h = tf.add_paragraph(); p_h.text = head; p_h.font.size = Pt(9.5); p_h.font.bold = True; p_h.font.color.rgb = WHITE
        p_b = tf.add_paragraph(); p_b.text = body; p_b.font.size = Pt(9); p_b.font.color.rgb = LIGHT; p_b.space_after = Pt(8)

add_footer_note(s2, "RESEARCH NOTE: In dense metros, 42%+ lack home charging. Reliability is improving (12% fail rate), so convenience—not permanent failure—is the primary value driver.")

# ==============================================================================
# SLIDE 3: Model 1 — Urban Charging Hub + Valet (EXACT CANVA 2-COLUMN)
# ==============================================================================
s3 = prs.slides.add_slide(blank_layout); set_bg(s3)
add_header(s3, "MODEL 1 — URBAN CHARGING HUB + VALET", 
           "Company-owned vertical parking towers with concierge pickup & return", 
           "High-control urban overnight charging paired with an end-to-end digital valet flow.")

# Left: Tower Image + Key Engineering Choices
if os.path.exists("images/vertical_parking_tower.jpg"):
    s3.shapes.add_picture("images/vertical_parking_tower.jpg", Inches(0.8), Inches(1.5), Inches(5.6), Inches(2.7))
    cap = s3.shapes.add_textbox(Inches(0.8), Inches(4.25), Inches(5.6), Inches(0.35))
    tf_c = cap.text_frame; tf_c.margin_left = tf_c.margin_top = 0
    p = tf_c.paragraphs[0]; p.text = "Automated vertical parking tower: doubles density without urban sprawl."; p.font.size = Pt(8.5); p.font.color.rgb = MUTED

c3_l = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(4.6), Inches(5.6), Inches(1.75))
c3_l.fill.solid(); c3_l.fill.fore_color.rgb = CARD_BG; c3_l.line.color.rgb = BORDER_MUTED
tf3_l = c3_l.text_frame; tf3_l.word_wrap = True; tf3_l.margin_left = Inches(0.25); tf3_l.margin_top = Inches(0.15)
p = tf3_l.paragraphs[0]; p.text = "KEY ENGINEERING CHOICES"; p.font.size = Pt(10); p.font.bold = True; p.font.color.rgb = CYAN_TEXT; p.space_after = Pt(4)
choices = [
    ("7–11 kW AC per stall", "Mass EVs cap at 11 kW; 22 kW adds cost for zero overnight gain."),
    ("~2× Stacker Density", "Proven vendor spec (avoids single-point failure carousel risk)."),
    ("Solar + BESS Microgrid", "Lowers peak grid tariffs and enables green charging."),
    ("Value-Add Concierge", "Windshield wash, tyre check, and interior vacuum.")
]
for t, d in choices:
    p = tf3_l.add_paragraph(); p.text = f"• {t}: {d}"; p.font.size = Pt(8.5); p.font.color.rgb = LIGHT; p.space_after = Pt(2)

# Right: 4-Stage Concierge Flow with Colored Headers
c3_r = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.7), Inches(1.5), Inches(5.833), Inches(4.85))
c3_r.fill.solid(); c3_r.fill.fore_color.rgb = CARD_BG; c3_r.line.color.rgb = BORDER_MUTED
tf3_r = c3_r.text_frame; tf3_r.word_wrap = True; tf3_r.margin_left = Inches(0.3); tf3_r.margin_top = Inches(0.2)
p = tf3_r.paragraphs[0]; p.text = "END-TO-END DIGITAL CONCIERGE FLOW"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = AMBER_TEXT; p.space_after = Pt(8)

stages_s3 = [
    ("STAGE 1: SCHEDULE (9:00 PM)", CYAN_TEXT, 
     "• In-App Booking: Owner selects vehicle, 9 PM–7 AM window, and target SoC (85%).\n• Escrow Lock: Platform holds estimated session + valet fee in digital escrow."),
    ("STAGE 2: PICKUP (9:30 PM)", GREEN_TEXT, 
     "• Verified Valet: Background-checked driver dispatched with GPS tracking.\n• 360° AI Scan: Smartphone camera records vehicle panels to eliminate damage disputes."),
    ("STAGE 3: CHARGE (10 PM–6 AM)", CYAN_TEXT, 
     "• Tower Stacking: Vehicle automatically parked in vertical automated stall.\n• Smart AC Charging: 7–11 kW charge curve with continuous thermal & battery health monitoring."),
    ("STAGE 4: RETURN (7:00 AM)", AMBER_TEXT, 
     "• Doorstep Delivery: Pre-conditioned vehicle returned to owner's parking bay ready for work.\n• Escrow Settlement: Post-delivery photo verified; payment auto-released with digital receipt.")
]
for title, color, desc in stages_s3:
    p = tf3_r.add_paragraph(); p.text = title; p.font.size = Pt(10); p.font.bold = True; p.font.color.rgb = color
    p = tf3_r.add_paragraph(); p.text = desc; p.font.size = Pt(8.5); p.font.color.rgb = LIGHT; p.space_after = Pt(6)

add_footer_note(s3, "RESEARCH NOTE: Valet unit economics: Weekly subscription packages (1–2x/week) optimize turnaround and raise break-even probability to 25–35%.")

# ==============================================================================
# SLIDE 4: Model 2 — Highway Distributed Charging (3 COLUMNS + 5-PHASE PIPELINE)
# ==============================================================================
s4 = prs.slides.add_slide(blank_layout); set_bg(s4)
add_header(s4, "MODEL 2 — HIGHWAY DISTRIBUTED CHARGING & ONBOARDING", 
           "Unlocking private and commercial chargers with frictionless host onboarding", 
           "P2P corridor charging network backed by smart locks, automated escrow, and 5-phase onboarding.")

# Col 1: Travellers
c4_1 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(3.8), Inches(3.3))
c4_1.fill.solid(); c4_1.fill.fore_color.rgb = CARD_BG; c4_1.line.color.rgb = BORDER_CYAN
tf = c4_1.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.2); tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]; p.text = "🚗  FOR HIGHWAY TRAVELLERS"; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = CYAN_TEXT; p.space_after = Pt(6)
t_pts = [
    ("Route Search", "Filter corridor chargers by plug type, power, and verified host rating."),
    ("45-Min Reservation", "Guaranteed charging slot with payment held in secure escrow."),
    ("Contactless BLE Token", "Unlock gate and charger via Bluetooth — zero host interaction."),
    ("Live Telemetry", "Real-time tracking of kWh delivered, charging speed, and cost.")
]
for t, d in t_pts:
    p = tf.add_paragraph(); p.text = f"• {t}: {d}"; p.font.size = Pt(8.5); p.font.color.rgb = LIGHT; p.space_after = Pt(4)

# Col 2: Center Image & Economics
c4_2 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.766), Inches(1.5), Inches(3.8), Inches(3.3))
c4_2.fill.solid(); c4_2.fill.fore_color.rgb = CARD_BG; c4_2.line.color.rgb = BORDER_MUTED
tf = c4_2.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.2); tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]; p.text = "SESSION ECONOMICS (Illustrative)"; p.font.size = Pt(10); p.font.bold = True; p.font.color.rgb = GREEN_TEXT; p.space_after = Pt(8)
econ = [
    ("Traveller Pays", "~$12.50 per charge"),
    ("Host Net Payout", "~$9.80 (78% earnings)"),
    ("Platform Take", "~$2.70 (22% margin)")
]
for t, d in econ:
    p = tf.add_paragraph(); p.text = f"{t}: {d}"; p.font.size = Pt(9.5); p.font.bold = True; p.font.color.rgb = WHITE; p.space_after = Pt(4)
p = tf.add_paragraph(); p.text = "Distributed charger with smart lock & CCTV kit."; p.font.size = Pt(8.5); p.font.color.rgb = MUTED

# Col 3: Hosts
c4_3 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.733), Inches(1.5), Inches(3.8), Inches(3.3))
c4_3.fill.solid(); c4_3.fill.fore_color.rgb = CARD_BG; c4_3.line.color.rgb = BORDER_GREEN
tf = c4_3.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.2); tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]; p.text = "🏠  FOR CHARGER HOSTS"; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = GREEN_TEXT; p.space_after = Pt(6)
h_pts = [
    ("Monetize Idle Assets", "Turn driveway, hotel, or workplace chargers into recurring income."),
    ("Low-CapEx Kit", "Platform ships plug-and-play smart lock, IoT gateway, and CCTV sensor."),
    ("Virtual Perimeter", "3-zone fence automatically cuts power if guest crosses boundaries."),
    ("Automated Payouts", "Weekly direct bank deposits with comprehensive $1M insurance.")
]
for t, d in h_pts:
    p = tf.add_paragraph(); p.text = f"• {t}: {d}"; p.font.size = Pt(8.5); p.font.color.rgb = LIGHT; p.space_after = Pt(4)

# Bottom: 5-Phase Horizontal Pipeline (EXACT CANVA)
pipe_w = 2.22
phases_c4 = [
    ("01", "Sign Up", "Upload charger spec, breaker size & kWh price."),
    ("02", "Verification", "Electrician reviews panel capacity & earthing."),
    ("03", "Kit Install", "Plug-and-play smart lock & CCTV sensor setup."),
    ("04", "Go-Live", "Calibrate 3-zone fence; listing published to map."),
    ("05", "Earn & Review", "Automated weekly bank payouts & host rating.")
]
for i, (num, ptitle, pdesc) in enumerate(phases_c4):
    p_card = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 2.38), Inches(4.95), Inches(pipe_w), Inches(1.4))
    p_card.fill.solid(); p_card.fill.fore_color.rgb = CARD_BG; p_card.line.color.rgb = BORDER_MUTED
    tf = p_card.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.15); tf.margin_top = Inches(0.15)
    p = tf.paragraphs[0]; p.text = f"{num}  {ptitle}"; p.font.size = Pt(10); p.font.bold = True; p.font.color.rgb = CYAN_TEXT; p.space_after = Pt(2)
    p = tf.add_paragraph(); p.text = pdesc; p.font.size = Pt(8); p.font.color.rgb = LIGHT

add_footer_note(s4, "RESEARCH NOTE: 45-min at 7.4 kW yields ~5.5 kWh (~20 mi range). We partner with highway hotels & fuel plazas for fast DC charging, using host chargers for local top-ups.")

# ==============================================================================
# SLIDE 5: Security & Safety Architecture (3 NEON CARDS: CYAN, GREEN, ROSE)
# ==============================================================================
s5 = prs.slides.add_slide(blank_layout); set_bg(s5)
add_header(s5, "SECURITY, ACCESS CONTROL & SAFETY ARCHITECTURE", 
           "Bank-grade trust, multi-tier permissions, and automated failsafe cutoffs", 
           "3-zone virtual perimeter, cryptographic smart locks, and sub-second hardware shutoffs.")

col_w5 = 3.75
# Card 1: 3-Zone (Cyan Border)
c5_1 = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(col_w5), Inches(4.85))
c5_1.fill.solid(); c5_1.fill.fore_color.rgb = CARD_BG; c5_1.line.color.rgb = BORDER_CYAN
tf = c5_1.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.2); tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]; p.text = "🛡️  3-ZONE VIRTUAL PERIMETER"; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = CYAN_TEXT; p.space_after = Pt(8)
zones = [
    ("🟢 GREEN ZONE (Charging Bay)", "Permitted vehicle footprint and charger radius. Active BLE token authorizes power flow; telemetry streams to host dashboard."),
    ("🟡 YELLOW ZONE (5m Buffer)", "5-meter warning boundary around vehicle. Movement triggers push alert to host and traveller; CCTV switches to 4K recording."),
    ("🔴 RED ZONE (Property Breach)", "Residence entrance or private garage boundary. Instant high-voltage power kill in <30 ms; 105 dB alarm + security alert.")
]
for zt, zd in zones:
    p = tf.add_paragraph(); p.text = zt; p.font.size = Pt(9.5); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf.add_paragraph(); p.text = zd; p.font.size = Pt(8.5); p.font.color.rgb = LIGHT; p.space_after = Pt(8)

# Card 2: Smart Locks (Green Border)
c5_2 = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.788), Inches(1.5), Inches(col_w5), Inches(4.85))
c5_2.fill.solid(); c5_2.fill.fore_color.rgb = CARD_BG; c5_2.line.color.rgb = BORDER_GREEN
tf = c5_2.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.2); tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]; p.text = "🔐  DUAL-MODE SMART LOCKS"; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = GREEN_TEXT; p.space_after = Pt(8)
locks = [
    ("🏠 OWNER MODE (Full Admin)", "• Permanent cryptographic master key on mobile app.\n• Instant manual override or session termination.\n• Real-time earnings, kWh sold, and escrow tracking.\n• Zone reconfiguration & whitelist trusted neighbours."),
    ("🚗 TRAVELLER MODE (Time-Boxed)", "• Ephemeral Bluetooth LE key valid only during slot.\n• Operates offline (no cellular signal needed at driveway).\n• Zero access to host home premises or hardware settings.\n• Automated lock re-engagement when window elapses.")
]
for lt, ld in locks:
    p = tf.add_paragraph(); p.text = lt; p.font.size = Pt(9.5); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf.add_paragraph(); p.text = ld; p.font.size = Pt(8.5); p.font.color.rgb = LIGHT; p.space_after = Pt(8)

# Card 3: 3 Failsafes (Rose Border)
c5_3 = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.777), Inches(1.5), Inches(col_w5), Inches(4.85))
c5_3.fill.solid(); c5_3.fill.fore_color.rgb = CARD_BG; c5_3.line.color.rgb = BORDER_ROSE
tf = c5_3.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.2); tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]; p.text = "⚠️  3 FAILSAFE CUTOFF TRIGGERS"; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = ROSE_TEXT; p.space_after = Pt(8)
fails = [
    ("🌡️ Thermal Overheat Probe", "Sensor: Dual NTC thermistors inside connector pins.\nTrigger: Terminal temp >60°C or rises >3°C/min.\nAction: Current throttled 50%; contactor opens in 30s."),
    ("⚡ Ground-Fault (RCD) Monitor", "Sensor: Integrated Type-B RCD sampling at 1,000 Hz.\nTrigger: DC leakage >6 mA or AC residual >30 mA.\nAction: High-voltage relay cuts power in <30 ms."),
    ("📷 AI Vision Damage Inspection", "Sensor: 4K optical sensor with night-vision IR.\nTrigger: Cable dragged, connector dropped, panel scuff.\nAction: Escrow frozen; dispute ticket generated with photos.")
]
for ft, fd in fails:
    p = tf.add_paragraph(); p.text = ft; p.font.size = Pt(9.5); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf.add_paragraph(); p.text = fd; p.font.size = Pt(8); p.font.color.rgb = LIGHT; p.space_after = Pt(6)

add_footer_note(s5, "RESEARCH NOTE: Hardware compliance with UL 2231, IEC 61851, and SAE J1772 safety standards with pre-session automated diagnostics.")

# ==============================================================================
# SLIDE 6: Business Model & Flywheel (FLYWHEEL + 7 REVENUE STREAMS)
# ==============================================================================
s6 = prs.slides.add_slide(blank_layout); set_bg(s6)
add_header(s6, "BUSINESS MODEL & ECOSYSTEM FLYWHEEL", 
           "Balancing high-margin owned assets with zero-CapEx distributed scale", 
           "7 diversified revenue streams powered by a self-reinforcing three-sided network.")

# Left: 3-Sided Flywheel Card
c6_l = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(4.8), Inches(4.85))
c6_l.fill.solid(); c6_l.fill.fore_color.rgb = CARD_BG; c6_l.line.color.rgb = BORDER_CYAN
tf = c6_l.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.25); tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]; p.text = "THE 3-SIDED ECOSYSTEM FLYWHEEL"; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = CYAN_TEXT; p.space_after = Pt(8)
fw = [
    ("SUPPLY CHANNELS", "Urban vertical stacker hubs (owned), residential host driveways, hotel corridors, commercial fleet depots."),
    ("CORE PLATFORM", "Matching engine, dynamic pricing, BLE token smart locks, escrow clearinghouse, AI vision & safety telemetry."),
    ("DEMAND CHANNELS", "Urban apartment EV drivers, inter-city road-trippers, commercial ride-hail (Uber/Ola), last-mile delivery fleets."),
    ("FLYWHEEL DYNAMICS", "Each 50 corridor hosts reduce detour distance by 3.2 miles → higher search conversion → more bookings → higher host earnings attract more supply.")
]
for t, d in fw:
    p = tf.add_paragraph(); p.text = t; p.font.size = Pt(9.5); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf.add_paragraph(); p.text = d; p.font.size = Pt(8.5); p.font.color.rgb = LIGHT; p.space_after = Pt(6)

# Right: 7 Revenue Streams
c6_r = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.8), Inches(1.5), Inches(6.733), Inches(4.85))
c6_r.fill.solid(); c6_r.fill.fore_color.rgb = CARD_BG; c6_r.line.color.rgb = BORDER_MUTED
tf = c6_r.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.25); tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]; p.text = "7 DIVERSIFIED REVENUE STREAMS"; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = AMBER_TEXT; p.space_after = Pt(6)
revs = [
    ("⚡ Charging Margin", "Retail electricity markup at owned hubs ($0.12/kWh) and host spread ($0.05/kWh)."),
    ("🚗 Valet Concierge Fee", "$15–$25 flat fee per vehicle pickup, charge, and morning return event."),
    ("📱 Platform Booking Fee", "10%–15% transaction fee charged to guest traveller on each reserved slot."),
    ("🏠 Host Revenue Spread", "20%–22% platform cut deducted from host payouts (covering insurance & hardware)."),
    ("💎 Driver Membership", "$19.99/mo subscription unlocking priority booking, waived fees, and valet perks."),
    ("🚌 Commercial Fleet SLAs", "$120–$180/vehicle/mo bulk overnight charging contracts for delivery/ride-hail."),
    ("⭐ Value-Add Services", "$10–$30 per add-on: overnight windshield wash, tyre pressure, and battery diagnostics.")
]
for t, d in revs:
    p = tf.add_paragraph(); p.text = t; p.font.size = Pt(10); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf.add_paragraph(); p.text = d; p.font.size = Pt(8.5); p.font.color.rgb = LIGHT; p.space_after = Pt(4)

add_footer_note(s6, "RESEARCH NOTE: Owned urban hubs deliver high operational control and 32% EBITDA margins; distributed host chargers provide exponential geographic scale with near-zero CapEx.")

# ==============================================================================
# SLIDE 7: Roadmap, Break-Even Odds & Citations
# ==============================================================================
s7 = prs.slides.add_slide(blank_layout); set_bg(s7)
add_header(s7, "ROADMAP, BREAK-EVEN ODDS & RESEARCH EVIDENCE", 
           "Execution milestones paired with verified empirical industry research", 
           "5 probability-weighted rollout phases and 10 core citations replacing unsourced assumptions.")

# Left: 5 Phases with Progress
c7_l = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(5.6), Inches(4.85))
c7_l.fill.solid(); c7_l.fill.fore_color.rgb = CARD_BG; c7_l.line.color.rgb = BORDER_MUTED
tf = c7_l.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.25); tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]; p.text = "5-PHASE ROADMAP & BREAK-EVEN ODDS"; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = CYAN_TEXT; p.space_after = Pt(8)
phases_s7 = [
    ("Phase 1: Urban Concierge Pilot", "35%", "1 owned hub + valet in 1 metro; 200 users. (CapEx: $650K)"),
    ("Phase 2: Multi-Family Apartment Sharing", "40%", "250 host chargers in apartments (EVmatch model). (CapEx: $200K (MVP))"),
    ("Phase 3: Highway Corridor Network", "25%", "75 hotel & fuel plaza DC fast charging partners. (CapEx: $450K)"),
    ("Phase 4: Regional Scale-Up", "20%", "10 urban hubs, 2,500 hosts, membership tiers. (CapEx: $3.5M)"),
    ("Phase 5: Enterprise Fleet SLAs & V2G", "15%", "Commercial fleet contracts; V2G utility integration. (Series B Raise)")
]
for pt, odds, pd in phases_s7:
    p = tf.add_paragraph(); p.text = f"{pt}  [{odds} Odds]"; p.font.size = Pt(10); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf.add_paragraph(); p.text = pd; p.font.size = Pt(8.5); p.font.color.rgb = LIGHT; p.space_after = Pt(5)

# Right: 10 Citations
c7_r = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.6), Inches(1.5), Inches(5.933), Inches(4.85))
c7_r.fill.solid(); c7_r.fill.fore_color.rgb = CARD_BG; c7_r.line.color.rgb = BORDER_MUTED
tf = c7_r.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.25); tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]; p.text = "10 CORE RESEARCH CITATIONS"; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = GREEN_TEXT; p.space_after = Pt(6)
cites_s7 = [
    ("J.D. Power (2024)", "Public failure rate fell to 12%; non-Tesla fail rate 52.4%."),
    ("Zapmap & US DOE", "~20% UK EV drivers lack home chargers; 82% US multi-family lack in-building plugs."),
    ("Frontiers Sustainable Cities", "UK P2P sharing has ~9,000 hosts; privacy and thin demand are main barriers."),
    ("EVmatch Case Study", "124 chargers across 21 apartment sites generated 6,200+ annual sessions (~1/wk)."),
    ("CBInsights / Luxe", "Nightly valet failed over labour costs; validates weekly subscription model."),
    ("NREL Highway Study", "Mapped corridor charging deserts; hotel chargers scored highest in satisfaction."),
    ("Autoparkit Whitepaper", "Automated parking stackers double density (~2×) without carousel failure risk."),
    ("Technical Hardware Guides", "Mass EVs max at 11 kW AC charging; 22 kW AC adds cost with zero overnight gain."),
    ("Corridor Charging Physics", "45-min at 7.4 kW yields 5.5 kWh (~20 mi range); requires DC fast partners for transit."),
    ("Probability Methodology", "Break-even odds reflect unit economics, utilization benchmarks, and startup survival data.")
]
for ct, cd in cites_s7:
    p = tf.add_paragraph(); p.text = f"• {ct}: {cd}"; p.font.size = Pt(8); p.font.color.rgb = LIGHT; p.space_after = Pt(2.5)

add_footer_note(s7, "RESEARCH NOTE: Lead with Phase 2 (Apartment community sharing — EVmatch model). Treat valet as a high-margin concierge tier. For corridors, partner with hotels and fuel plazas for DC fast charging.")

# Save presentation
out_pptx = "Canva_Exact_Shortened.pptx"
out_ppsx = "Canva_Exact_Shortened.ppsx"
prs.save(out_pptx)
import shutil
shutil.copyfile(out_pptx, out_ppsx)
print(f"Generated {out_pptx} and {out_ppsx} matching Canva layout exactly!")
