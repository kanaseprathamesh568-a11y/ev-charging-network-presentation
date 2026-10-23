import os
import pptx
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

prs = pptx.Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# Theme Colors
BG_DARK = RGBColor(2, 6, 23)        # #020617
CARD_BG = RGBColor(15, 23, 42)      # #0f172a
CARD_BORDER = RGBColor(30, 41, 59)  # #1e293b
CYAN = RGBColor(34, 211, 238)       # #22d3ee
CYAN_BRIGHT = RGBColor(6, 182, 212) # #06b6d4
EMERALD = RGBColor(52, 211, 153)    # #34d399
AMBER = RGBColor(251, 191, 36)      # #fbbf24
ROSE = RGBColor(248, 113, 113)      # #f87171
WHITE = RGBColor(255, 255, 255)
LIGHT = RGBColor(226, 232, 240)     # slate-200
MUTED = RGBColor(148, 163, 184)     # slate-400
DIM = RGBColor(100, 116, 139)       # slate-500

def set_bg(slide):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG_DARK
    bg.line.fill.background()
    return bg

def add_header(slide, tag, title):
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.9))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    p0 = tf.paragraphs[0]
    p0.text = f"{tag.upper()}"
    p0.font.size = Pt(9.5)
    p0.font.bold = True
    p0.font.color.rgb = CYAN
    p0.space_after = Pt(2)
    
    p1 = tf.add_paragraph()
    p1.text = title
    p1.font.size = Pt(22)
    p1.font.bold = True
    p1.font.color.rgb = WHITE

def add_footer(slide, current_idx, total=11):
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(7.1), Inches(11.7), Inches(0.3))
    tf = tb.text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = f"Har Ghar Charging  |  Team: VanGuards  |  Slide {current_idx:02d} of {total:02d}"
    p.font.size = Pt(8.5)
    p.font.color.rgb = DIM

def make_card(slide, left, top, width, height, bg_col=CARD_BG, border_col=CARD_BORDER):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = bg_col
    card.line.color.rgb = border_col
    return card

# ==============================================================================
# SLIDE 1: Title & Problem Statement
# ==============================================================================
s1 = prs.slides.add_slide(blank_layout); set_bg(s1)
add_header(s1, "TEAM: VANGUARDS  |  01 / 11", "Har Ghar Charging: Charge. Anytime. Anywhere.")
c1 = make_card(s1, 0.8, 1.5, 6.0, 5.3)
tf1 = c1.text_frame; tf1.word_wrap = True; tf1.margin_left = Inches(0.3); tf1.margin_top = Inches(0.3)
p = tf1.paragraphs[0]; p.text = "THE CORE PROBLEM"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = AMBER; p.space_after = Pt(8)
p = tf1.add_paragraph(); p.text = "Getting society approvals, separate meters, and dedicated EV parking in apartments is nearly impossible."; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = WHITE; p.space_after = Pt(6)
p = tf1.add_paragraph(); p.text = "Urban EV owners are forced onto overcrowded, unreliable public fast chargers for routine daily charging."; p.font.size = Pt(11.5); p.font.color.rgb = LIGHT; p.space_after = Pt(14)

points = [
    ("🚫 Society Approvals", "Lengthy RWA disputes & arbitrary bans on private charger installations."),
    ("⚡ Grid & Meter Caps", "Disallowed separate meters and building transformer capacity limits."),
    ("🅿️ Parking Scarcity", "No dedicated EV bays in 80%+ of urban residential multi-family housing."),
    ("⏳ Public Inconvenience", "Long charging queues, broken plugs, and severe daily time loss.")
]
for title, desc in points:
    p = tf1.add_paragraph()
    p.text = f"{title}: {desc}"
    p.font.size = Pt(10)
    p.font.color.rgb = LIGHT
    p.space_after = Pt(5)

if os.path.exists("images/public_charger_queue.jpg"):
    s1.shapes.add_picture("images/public_charger_queue.jpg", Inches(7.1), Inches(1.5), Inches(5.4), Inches(5.3))
add_footer(s1, 1)

# ==============================================================================
# SLIDE 2: Executive Pitch / Core Value Proposition
# ==============================================================================
s2 = prs.slides.add_slide(blank_layout); set_bg(s2)
add_header(s2, "EXECUTIVE PITCH  |  CORE PROPOSITION", "You don't need to find a charger. We bring charging to you.")
c2 = make_card(s2, 0.8, 1.5, 6.0, 5.3)
tf2 = c2.text_frame; tf2.word_wrap = True; tf2.margin_left = Inches(0.3); tf2.margin_top = Inches(0.3)
p = tf2.paragraphs[0]; p.text = "ONE PLATFORM · TWO SYNERGISTIC MODELS · ZERO RANGE ANXIETY"; p.font.size = Pt(10); p.font.bold = True; p.font.color.rgb = CYAN; p.space_after = Pt(12)

bullets_s2 = [
    ("🌙 Overnight Concierge Valet", "Vehicle collected at 9 PM, returned at 7 AM fully charged (100% SoC) from owned vertical charging towers."),
    ("🛣️ Distributed Highway Hosts", "Private driveway & hotel chargers unlocked along corridors via cryptographic BLE tokens."),
    ("🔒 Bank-Grade Trust Layer", "3-zone perimeter protection, 4K AI damage audit, and automated digital escrow settlement."),
    ("📈 Capital-Efficient Scale", "High-margin owned vertical hubs generate cash flow; zero-CapEx host network provides exponential scale.")
]
for title, desc in bullets_s2:
    p = tf2.add_paragraph(); p.text = title; p.font.size = Pt(11.5); p.font.bold = True; p.font.color.rgb = WHITE
    pd = tf2.add_paragraph(); pd.text = desc; pd.font.size = Pt(10); pd.font.color.rgb = LIGHT; pd.space_after = Pt(8)

p = tf2.add_paragraph(); p.text = "Verified benchmarks: J.D. Power (2024), Zapmap, NREL & India CEA."; p.font.size = Pt(8.5); p.font.color.rgb = MUTED

if os.path.exists("images/ev_port_inspection.jpg"):
    s2.shapes.add_picture("images/ev_port_inspection.jpg", Inches(7.1), Inches(1.5), Inches(5.4), Inches(5.3))
add_footer(s2, 2)

# ==============================================================================
# SLIDE 3: The Problem & Market Opportunity (Data Behind It)
# ==============================================================================
s3 = prs.slides.add_slide(blank_layout); set_bg(s3)
add_header(s3, "MARKET OPPORTUNITY & EMPIRICAL DATA", "Why Public EV Charging Fails Drivers Today")

stat_boxes = [
    ("42%+", "Urban Apts Lack Bays", "India Metros: Over 42% of EV owners live in apartments with zero charging infrastructure.\n\nUS: 82% multi-family residents lack in-building plugs.\n(Source: India CEA / US DOE)", 0.8),
    ("12-20%", "Session Failure Rate", "High public charging failure rate on non-Tesla networks.\n\n52.4% of drivers reported defective or offline chargers in 2023.\n(Source: J.D. Power / PlugShare)", 3.8),
    ("Corridor", "Highway Charging Gaps", "Fast-charging deserts outside city rings.\n\nLong holiday queues & broken hardware make inter-city road trips stressful.\n(Source: NREL Corridor Study)", 6.8),
    ("20+ Hrs", "Idle Private Capacity", "Millions of residential, hotel & workplace chargers sit empty 20+ hours daily.\n\nHigh-CapEx public networks lag while private assets remain locked.\n(Source: Frontiers Sustainable Cities)", 9.8)
]

for stat, title, body, left in stat_boxes:
    cb = make_card(s3, left, 1.5, 2.75, 4.3)
    tf = cb.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.2); tf.margin_top = Inches(0.2)
    p = tf.paragraphs[0]; p.text = stat; p.font.size = Pt(24); p.font.bold = True; p.font.color.rgb = CYAN; p.space_after = Pt(2)
    p = tf.add_paragraph(); p.text = title; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = WHITE; p.space_after = Pt(8)
    p = tf.add_paragraph(); p.text = body; p.font.size = Pt(9.5); p.font.color.rgb = LIGHT

note_c = make_card(s3, 0.8, 6.0, 11.75, 0.8)
tf_n = note_c.text_frame; tf_n.word_wrap = True; tf_n.margin_left = Inches(0.3); tf_n.margin_top = Inches(0.15)
p = tf_n.paragraphs[0]
p.text = "💡 Key Takeaway: Public infrastructure alone cannot scale fast enough. Unlocking idle private assets and overnight concierge charging solves daily friction with near-zero public CapEx."
p.font.size = Pt(10.5); p.font.color.rgb = AMBER
add_footer(s3, 3)

# ==============================================================================
# SLIDE 4: Model 1 — Urban Charging Hub + Valet
# ==============================================================================
s4 = prs.slides.add_slide(blank_layout); set_bg(s4)
add_header(s4, "MODEL 1 — URBAN CHARGING HUB + VALET", "Automated Vertical Parking Towers with Concierge Pickup")

# Left: Tower specs & image
if os.path.exists("images/vertical_parking_tower.jpg"):
    s4.shapes.add_picture("images/vertical_parking_tower.jpg", Inches(0.8), Inches(1.5), Inches(5.6), Inches(2.8))

c4_l = make_card(s4, 0.8, 4.45, 5.6, 2.35)
tf4_l = c4_l.text_frame; tf4_l.word_wrap = True; tf4_l.margin_left = Inches(0.25); tf4_l.margin_top = Inches(0.2)
p = tf4_l.paragraphs[0]; p.text = "KEY ENGINEERING CHOICES"; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = CYAN; p.space_after = Pt(4)
eng_points = [
    ("• 7–11 kW AC per Stall", "Matches onboard mass-EV chargers; avoids costly 22 kW hardware with zero overnight benefit."),
    ("• 2× Stacker Density", "Automated puzzle/stacker doubles parking capacity without requiring extra land."),
    ("• Solar + BESS Storage", "Integrated microgrid buffers peak tariffs and guarantees renewable power."),
    ("• Valet Subscription", "Offered 1–2× weekly to optimize fleet turnaround and break-even economics.")
]
for title, desc in eng_points:
    p = tf4_l.add_paragraph()
    p.text = f"{title}: {desc}"
    p.font.size = Pt(9); p.font.color.rgb = LIGHT; p.space_after = Pt(2)

# Right: 4-Stage Flow
c4_r = make_card(s4, 6.7, 1.5, 5.85, 5.3)
tf4_r = c4_r.text_frame; tf4_r.word_wrap = True; tf4_r.margin_left = Inches(0.3); tf4_r.margin_top = Inches(0.3)
p = tf4_r.paragraphs[0]; p.text = "END-TO-END DIGITAL CONCIERGE FLOW"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = AMBER; p.space_after = Pt(10)

stages = [
    ("STAGE 1: SCHEDULE (9:00 PM)", "Owner books 9 PM–7 AM slot with target SoC (85%); session fee locked in digital escrow."),
    ("STAGE 2: PICKUP (9:30 PM)", "GPS-dispatched verified valet executes 360° AI smartphone scan to record panel conditions."),
    ("STAGE 3: CHARGE (10 PM–6 AM)", "Vehicle stacked automatically into vertical bay with continuous thermal & battery health monitoring."),
    ("STAGE 4: RETURN (7:00 AM)", "Pre-conditioned EV delivered to owner's slot; payment released post-delivery verification.")
]
for title, desc in stages:
    p = tf4_r.add_paragraph(); p.text = title; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf4_r.add_paragraph(); p.text = desc; p.font.size = Pt(9.5); p.font.color.rgb = LIGHT; p.space_after = Pt(8)
add_footer(s4, 4)

# ==============================================================================
# SLIDE 5: Model 2 — Highway Distributed Charging
# ==============================================================================
s5 = prs.slides.add_slide(blank_layout); set_bg(s5)
add_header(s5, "MODEL 2 — HIGHWAY DISTRIBUTED CHARGING", "Unlocking Private & Commercial Chargers with Host Onboarding")

# Travellers card
c5_1 = make_card(s5, 0.8, 1.5, 3.8, 3.6)
tf5_1 = c5_1.text_frame; tf5_1.word_wrap = True; tf5_1.margin_left = Inches(0.2); tf5_1.margin_top = Inches(0.2)
p = tf5_1.paragraphs[0]; p.text = "🚗 FOR HIGHWAY TRAVELLERS"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = CYAN; p.space_after = Pt(6)
t_pts = [
    ("• Smart Corridor Search", "Filter by plug type, charging speed, and verified host rating."),
    ("• 45-Min Reservation", "Guaranteed slot held in escrow; zero queuing anxiety."),
    ("• Contactless BLE Token", "Unlock gate and charger from phone—zero host contact needed."),
    ("• Live Telemetry", "Track real-time charging speed, kWh delivered, and cost in-app.")
]
for t, d in t_pts:
    p = tf5_1.add_paragraph(); p.text = f"{t}: {d}"; p.font.size = Pt(9); p.font.color.rgb = LIGHT; p.space_after = Pt(4)

# Economics in Center
c5_2 = make_card(s5, 4.8, 1.5, 3.7, 3.6)
tf5_2 = c5_2.text_frame; tf5_2.word_wrap = True; tf5_2.margin_left = Inches(0.2); tf5_2.margin_top = Inches(0.2)
p = tf5_2.paragraphs[0]; p.text = "📊 SESSION UNIT ECONOMICS"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = EMERALD; p.space_after = Pt(8)
econ = [
    ("Avg Session Price", "~$12.50 per charge"),
    ("Host Net Payout", "~$9.80 (78% earnings)"),
    ("Platform Margin", "~$2.70 (22% take rate)"),
    ("Asset Payback", "< 3 months for host kit")
]
for t, d in econ:
    p = tf5_2.add_paragraph(); p.text = t; p.font.size = Pt(9.5); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf5_2.add_paragraph(); p.text = d; p.font.size = Pt(11.5); p.font.color.rgb = CYAN; p.space_after = Pt(6)

# Hosts card
c5_3 = make_card(s5, 8.7, 1.5, 3.8, 3.6)
tf5_3 = c5_3.text_frame; tf5_3.word_wrap = True; tf5_3.margin_left = Inches(0.2); tf5_3.margin_top = Inches(0.2)
p = tf5_3.paragraphs[0]; p.text = "🏡 FOR CHARGER HOSTS"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = AMBER; p.space_after = Pt(6)
h_pts = [
    ("• Monetize Downtime", "Convert idle driveway or hotel chargers into steady monthly income."),
    ("• Low-CapEx Kit", "Platform provides plug-and-play smart lock, IoT gateway, and camera."),
    ("• Virtual Perimeter", "3-zone boundary cuts power instantly if guest crosses limits."),
    ("• Direct Weekly Payouts", "Automatic bank deposits backed by comprehensive $1M insurance.")
]
for t, d in h_pts:
    p = tf5_3.add_paragraph(); p.text = f"{t}: {d}"; p.font.size = Pt(9); p.font.color.rgb = LIGHT; p.space_after = Pt(4)

# 5-Phase Onboarding Bar
c5_bar = make_card(s5, 0.8, 5.3, 11.7, 1.5)
tf5_b = c5_bar.text_frame; tf5_b.word_wrap = True; tf5_b.margin_left = Inches(0.25); tf5_b.margin_top = Inches(0.15)
p = tf5_b.paragraphs[0]; p.text = "5-PHASE HOST ONBOARDING PIPELINE"; p.font.size = Pt(10); p.font.bold = True; p.font.color.rgb = WHITE; p.space_after = Pt(4)
p = tf5_b.add_paragraph()
p.text = "01 Sign Up (Upload specs & rates)  ➔  02 Verification (Electrician inspection)  ➔  03 Kit Install (Smart lock & sensor)  ➔  04 Go-Live (Perimeter test & publish)  ➔  05 Earn (Automated payouts & reviews)"
p.font.size = Pt(10.5); p.font.color.rgb = CYAN
add_footer(s5, 5)

# ==============================================================================
# SLIDE 6: Security, Access Control & Safety Architecture
# ==============================================================================
s6 = prs.slides.add_slide(blank_layout); set_bg(s6)
add_header(s6, "SECURITY & SAFETY ARCHITECTURE", "Bank-Grade Trust, Access Control, and Automated Cutoffs")

col_w = 3.75
# Col 1: 3-Zone Perimeter
c6_1 = make_card(s6, 0.8, 1.5, col_w, 4.4)
tf = c6_1.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.2); tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]; p.text = "🛡️ 3-ZONE VIRTUAL PERIMETER"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = CYAN; p.space_after = Pt(8)
z_pts = [
    ("🟢 Green Zone (Bay)", "Permitted vehicle footprint. Active BLE token unlocks charging flow; telemetry streams to host."),
    ("🟡 Yellow Zone (5m Buffer)", "Warning boundary. Triggers immediate push notification and escalates CCTV to 4K recording."),
    ("🔴 Red Zone (Property)", "Private home entrance. Instant <30ms high-voltage cutoff + 105 dB local alarm & security alert.")
]
for t, d in z_pts:
    p = tf.add_paragraph(); p.text = t; p.font.size = Pt(10); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf.add_paragraph(); p.text = d; p.font.size = Pt(8.5); p.font.color.rgb = LIGHT; p.space_after = Pt(5)

# Col 2: Smart Locks
c6_2 = make_card(s6, 4.77, 1.5, col_w, 4.4)
tf = c6_2.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.2); tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]; p.text = "🔐 DUAL-MODE SMART LOCKS"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = EMERALD; p.space_after = Pt(8)
l_pts = [
    ("👑 Host Mode (Full Admin)", "Permanent cryptographic master key. Real-time earnings, manual emergency override, and trusted neighbor whitelist."),
    ("🚗 Traveller Mode (Time-Boxed)", "Ephemeral Bluetooth LE token valid strictly for booked slot. Operates offline; zero access to host residence or electrical settings."),
    ("🔒 Auto Re-Lock", "Contactor automatically trips and gate re-engages the second booking expires.")
]
for t, d in l_pts:
    p = tf.add_paragraph(); p.text = t; p.font.size = Pt(10); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf.add_paragraph(); p.text = d; p.font.size = Pt(8.5); p.font.color.rgb = LIGHT; p.space_after = Pt(5)

# Col 3: Failsafes
c6_3 = make_card(s6, 8.75, 1.5, col_w, 4.4)
tf = c6_3.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.2); tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]; p.text = "⚡ 3 FAILSAFE CUTOFF TRIGGERS"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = ROSE; p.space_after = Pt(8)
f_pts = [
    ("🔥 Thermal Overheat Probe", "Dual NTC thermistors in connector. Throttles at >60°C; disconnects completely within 30s."),
    ("⚡ Ground-Fault RCD Monitor", "Integrated Type-B RCD sampling at 1,000 Hz. Cuts high-voltage power in <30ms upon leakage."),
    ("📷 AI Vision Damage Inspection", "4K sensor detects dragged cables or vehicle panel scuffs; auto-freezes escrow for dispute.")
]
for t, d in f_pts:
    p = tf.add_paragraph(); p.text = t; p.font.size = Pt(10); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf.add_paragraph(); p.text = d; p.font.size = Pt(8.5); p.font.color.rgb = LIGHT; p.space_after = Pt(5)

# Standards footer
c6_b = make_card(s6, 0.8, 6.05, 11.7, 0.75)
tf = c6_b.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.3); tf.margin_top = Inches(0.15)
p = tf.paragraphs[0]
p.text = "Standards Compliance: Hardware adheres strictly to UL 2231, IEC 61851, and SAE J1772 with pre-session automated self-diagnostics."
p.font.size = Pt(9.5); p.font.color.rgb = MUTED
add_footer(s6, 6)

# ==============================================================================
# SLIDE 7: Business Model & Ecosystem Flywheel
# ==============================================================================
s7 = prs.slides.add_slide(blank_layout); set_bg(s7)
add_header(s7, "BUSINESS MODEL & ECOSYSTEM FLYWHEEL", "High-Margin Owned Hubs Balanced with Zero-CapEx Scale")

# Left: 3-Sided Flywheel
c7_l = make_card(s7, 0.8, 1.5, 4.8, 5.3)
tf = c7_l.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.25); tf.margin_top = Inches(0.25)
p = tf.paragraphs[0]; p.text = "THE 3-SIDED NETWORK FLYWHEEL"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = CYAN; p.space_after = Pt(10)
fw_pts = [
    ("1. Supply Channels", "Urban vertical towers (owned) + residential driveway hosts + hotel corridor partners."),
    ("2. Platform Core", "Matching engine, dynamic pricing, BLE smart locks, and automated escrow clearinghouse."),
    ("3. Demand Channels", "Urban apartment EV drivers, inter-city road trippers, and commercial delivery fleets."),
    ("⚡ Flywheel Dynamic", "Every 50 new corridor hosts cuts detour distance by 3.2 miles → drives higher conversion → generates more host earnings → attracts new host supply.")
]
for t, d in fw_pts:
    p = tf.add_paragraph(); p.text = t; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf.add_paragraph(); p.text = d; p.font.size = Pt(9); p.font.color.rgb = LIGHT; p.space_after = Pt(8)

# Right: 7 Revenue Streams
c7_r = make_card(s7, 5.8, 1.5, 6.7, 5.3)
tf = c7_r.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.25); tf.margin_top = Inches(0.25)
p = tf.paragraphs[0]; p.text = "7 DIVERSIFIED REVENUE STREAMS"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = AMBER; p.space_after = Pt(8)

rev_streams = [
    ("⚡ Charging Margin", "$0.12/kWh retail markup at owned hubs; $0.05/kWh spread on host networks."),
    ("🚗 Valet Concierge Fee", "$15–$25 flat fee per vehicle pickup, charge, and morning return."),
    ("📱 Platform Booking Fee", "10%–15% transaction fee paid by guest traveller per reserved slot."),
    ("🏡 Host Revenue Take", "20%–22% platform cut deducted from host payouts (covers insurance & hardware)."),
    ("💎 Driver Membership", "$19.99/month subscription unlocking priority booking, waived fees, and valet perks."),
    ("🚐 Commercial Fleet SLAs", "$120–$180/vehicle/month bulk overnight charging contracts for ride-hail/logistics."),
    ("⭐ Value-Add Detailing", "$10–$30 per add-on: overnight windshield wash, tire check & battery diagnostics.")
]
for t, d in rev_streams:
    p = tf.add_paragraph()
    p.text = f"{t}: {d}"
    p.font.size = Pt(9.5); p.font.color.rgb = LIGHT; p.space_after = Pt(4)
add_footer(s7, 7)

# ==============================================================================
# SLIDE 8: Roadmap & Execution Evidence
# ==============================================================================
s8 = prs.slides.add_slide(blank_layout); set_bg(s8)
add_header(s8, "ROADMAP & RESEARCH EVIDENCE", "5-Phase Execution Plan Backed by Industry Benchmarks")

# Left: 5 Phases
c8_l = make_card(s8, 0.8, 1.5, 5.6, 5.3)
tf = c8_l.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.25); tf.margin_top = Inches(0.25)
p = tf.paragraphs[0]; p.text = "5-PHASE ROLLOUT ROADMAP"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = CYAN; p.space_after = Pt(8)

phases = [
    ("Phase 1: Urban Concierge Pilot", "1 owned hub + valet service in 1 metro; 200 users ($650K CapEx, 35% break-even odds)."),
    ("Phase 2: Multi-Family Apartment Sharing", "250 host chargers in residential societies; EVmatch model ($200K CapEx, 40% odds)."),
    ("Phase 3: Highway Corridor Network", "75 hotel & highway plaza DC fast-charging partners ($450K CapEx, 25% odds)."),
    ("Phase 4: Regional Scale-Up", "10 urban hubs, 2,500 hosts, recurring driver membership tiers ($3.5M CapEx, 20% odds)."),
    ("Phase 5: Enterprise Fleet SLAs & V2G", "Commercial fleet bulk contracts and utility grid balancing (Series B expansion).")
]
for t, d in phases:
    p = tf.add_paragraph(); p.text = t; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf.add_paragraph(); p.text = d; p.font.size = Pt(9); p.font.color.rgb = LIGHT; p.space_after = Pt(6)

# Right: Evidence Citations
c8_r = make_card(s8, 6.6, 1.5, 5.9, 5.3)
tf = c8_r.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.25); tf.margin_top = Inches(0.25)
p = tf.paragraphs[0]; p.text = "VERIFIED RESEARCH CITATIONS"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = EMERALD; p.space_after = Pt(8)

cites = [
    ("J.D. Power (2024)", "52.4% non-Tesla failure experience; validates urgent need for managed reliability."),
    ("Zapmap & US DOE", "82% multi-family residents lack in-building chargers; primary market barrier."),
    ("EVmatch Case Study", "124 apartment chargers produced 6,200+ annual sessions; confirms sharing feasibility."),
    ("NREL Corridor Study", "Hospitality-located corridor chargers achieve highest dwell-time user satisfaction."),
    ("Autoparkit Study", "Automated vertical stackers achieve ~2× parking density with zero land expansion.")
]
for t, d in cites:
    p = tf.add_paragraph(); p.text = t; p.font.size = Pt(10); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf.add_paragraph(); p.text = d; p.font.size = Pt(9); p.font.color.rgb = LIGHT; p.space_after = Pt(6)
add_footer(s8, 8)

# ==============================================================================
# SLIDE 9: Competitive Landscape
# ==============================================================================
s9 = prs.slides.add_slide(blank_layout); set_bg(s9)
add_header(s9, "COMPETITIVE LANDSCAPE", "Connecting Fragmented Pieces into One Complete Network")

c9_m = make_card(s9, 0.8, 1.5, 5.8, 5.3)
tf = c9_m.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.25); tf.margin_top = Inches(0.25)
p = tf.paragraphs[0]; p.text = "MARKET POSITIONING MATRIX"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = CYAN; p.space_after = Pt(8)
comp_matrix = [
    ("EZAR EV", "Valet charging only (High CapEx, single-point model, no P2P sharing)."),
    ("ChargeKr & GigerEv", "P2P listing only (No valet, no hardware access controls, unmonitored)."),
    ("BharatREVOLT & ValEV", "Partner venue charging (No residential multi-family penetration)."),
    ("Har Ghar Charging (Ours)", "Unified platform: Company-operated hubs + P2P host sharing + bank-grade BLE hardware security + digital escrow.")
]
for t, d in comp_matrix:
    p = tf.add_paragraph(); p.text = t; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf.add_paragraph(); p.text = d; p.font.size = Pt(9.5); p.font.color.rgb = LIGHT; p.space_after = Pt(8)

c9_d = make_card(s9, 6.8, 1.5, 5.7, 5.3)
tf = c9_d.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.25); tf.margin_top = Inches(0.25)
p = tf.paragraphs[0]; p.text = "KEY STRATEGIC DIFFERENTIATORS"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = AMBER; p.space_after = Pt(10)
diffs = [
    ("✓ Dual Complementary Delivery", "Urban overnight valet solves apartment friction; highway P2P solves road trips."),
    ("✓ Bank-Grade Physical Security", "3-zone perimeter monitoring and sub-second cutoffs eliminate homeowner liability."),
    ("✓ AI Vehicle Inspection", "360° smartphone body audit eliminates damage disputes before and after valet."),
    ("✓ Balanced Capital Structure", "Owned vertical hubs generate solid margins; host network drives rapid zero-CapEx scale.")
]
for t, d in diffs:
    p = tf.add_paragraph(); p.text = t; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf.add_paragraph(); p.text = d; p.font.size = Pt(9); p.font.color.rgb = LIGHT; p.space_after = Pt(8)
add_footer(s9, 9)

# ==============================================================================
# SLIDE 10: Pilot Execution with ₹50,000
# ==============================================================================
s10 = prs.slides.add_slide(blank_layout); set_bg(s10)
add_header(s10, "PILOT EXECUTION PLAN  |  ₹50,000 BUDGET", "Real-World Proof of Concept & Unit Economics Validation")

# Budget
c10_b = make_card(s10, 0.8, 1.5, 3.8, 5.3)
tf = c10_b.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.2); tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]; p.text = "BUDGET BREAKDOWN (₹50,000)"; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = CYAN; p.space_after = Pt(6)
budget = [
    ("₹12,000", "Smart access controller prototype"),
    ("₹10,000", "Contingency reserve"),
    ("₹8,000", "Web app & booking MVP"),
    ("₹5,000", "CCTV & sensor monitoring"),
    ("₹5,000", "Partner charger setup & wiring check"),
    ("₹5,000", "Electrical & safety consultation"),
    ("₹5,000", "Pilot user acquisition / local ads")
]
for amt, item in budget:
    p = tf.add_paragraph(); p.text = f"{amt} — {item}"; p.font.size = Pt(9.5); p.font.color.rgb = LIGHT; p.space_after = Pt(4)

# 30-Day Plan
c10_p = make_card(s10, 4.8, 1.5, 4.2, 5.3)
tf = c10_p.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.2); tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]; p.text = "30-DAY EXECUTION TIMELINE"; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = AMBER; p.space_after = Pt(6)
wks = [
    ("Week 1: Validate & Partner", "Survey 20–30 apartment EV owners; sign 1–2 test charger hosts."),
    ("Week 2: Build MVP", "Deploy mobile booking flow, slot scheduling, and escrow integration."),
    ("Week 3: Setup Hardware", "Install smart BLE controller and calibrate 3-zone virtual perimeter."),
    ("Week 4: Execute Pilot", "Run 5–10 real paid sessions; collect feedback and measure telemetry.")
]
for w, d in wks:
    p = tf.add_paragraph(); p.text = w; p.font.size = Pt(10); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf.add_paragraph(); p.text = d; p.font.size = Pt(8.5); p.font.color.rgb = LIGHT; p.space_after = Pt(6)

# Expected Outcome
c10_o = make_card(s10, 9.2, 1.5, 3.3, 5.3)
tf = c10_o.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.2); tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]; p.text = "EXPECTED OUTCOMES"; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = EMERALD; p.space_after = Pt(8)
outs = [
    ("5–10 Paid Sessions", "Proof of actual customer willingness to pay."),
    ("Unit Economics", "Validate margins, host splits, and operational costs."),
    ("Safety Verification", "Flawless performance of BLE locks & perimeter cutoffs."),
    ("Traction Metrics", "Hard data to unlock seed capital & apartment MOUs.")
]
for o, d in outs:
    p = tf.add_paragraph(); p.text = f"• {o}"; p.font.size = Pt(10); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf.add_paragraph(); p.text = d; p.font.size = Pt(8.5); p.font.color.rgb = LIGHT; p.space_after = Pt(6)
add_footer(s10, 10)

# ==============================================================================
# SLIDE 11: Conclusion
# ==============================================================================
s11 = prs.slides.add_slide(blank_layout); set_bg(s11)
add_header(s11, "CONCLUSION  |  THE FUTURE OF CHARGING", "Charge Without the Chase: One Secure EV Charging Network")

c11_l = make_card(s11, 0.8, 1.5, 5.8, 4.4)
tf = c11_l.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.25); tf.margin_top = Inches(0.25)
p = tf.paragraphs[0]; p.text = "CORE STRATEGIC PILLARS"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = CYAN; p.space_after = Pt(8)
pillars = [
    ("1. Solves Real Daily Friction", "Eliminates the apartment charging gap with overnight valet concierge."),
    ("2. Establishes Absolute Trust", "3-zone perimeter protection, 4K AI audits, and automated digital escrow."),
    ("3. Leverages Existing Assets", "Mobilizes idle private & hotel chargers without waiting for grid upgrades."),
    ("4. Accelerates Green Adoption", "Removes range anxiety and promotes sustainable urban mobility.")
]
for t, d in pillars:
    p = tf.add_paragraph(); p.text = t; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf.add_paragraph(); p.text = d; p.font.size = Pt(9); p.font.color.rgb = LIGHT; p.space_after = Pt(6)

if os.path.exists("images/ev_charging_hero.jpg"):
    s11.shapes.add_picture("images/ev_charging_hero.jpg", Inches(6.8), Inches(1.5), Inches(5.7), Inches(4.4))

c11_b = make_card(s11, 0.8, 6.05, 11.7, 0.75)
tf = c11_b.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.3); tf.margin_top = Inches(0.15)
p = tf.paragraphs[0]
p.text = "FROM APARTMENT PARKING TO HIGHWAY TRANSIT — HAR GHAR CHARGING: CHARGE. ANYTIME. ANYWHERE."
p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = AMBER
add_footer(s11, 11)

# Save presentation
pptx_path = "Har_Ghar_Charging_Shortened.pptx"
ppsx_path = "Har_Ghar_Charging_Shortened.ppsx"
prs.save(pptx_path)
import shutil
shutil.copyfile(pptx_path, ppsx_path)
print(f"Generated {pptx_path} and {ppsx_path} successfully!")
