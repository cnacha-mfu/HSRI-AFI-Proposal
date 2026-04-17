"""Regenerate Figure 1 (System Architecture) and Figure 2 (Patient Journey)."""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch
import matplotlib as mpl

# Use TH Sarabun New for Thai character support
mpl.rcParams['font.family'] = 'TH Sarabun New'
mpl.rcParams['font.size'] = 9

# ─────────────────────────────────────────────────────────────────────────────
# Shared helpers
# ─────────────────────────────────────────────────────────────────────────────

def rounded_box(ax, x, y, w, h, text, bg, fg='#1a1a2e',
                fontsize=9, bold_title=None, pad=0.15, radius=0.08):
    """Draw a rounded rectangle with centred text."""
    box = FancyBboxPatch((x, y), w, h,
                         boxstyle=f"round,pad={pad}",
                         linewidth=1.2, edgecolor='#555',
                         facecolor=bg, zorder=3)
    ax.add_patch(box)
    cx, cy = x + w / 2, y + h / 2
    if bold_title:
        ax.text(cx, cy + 0.13, bold_title, ha='center', va='center',
                fontsize=fontsize, fontweight='bold', color=fg, zorder=4,
                wrap=True)
        ax.text(cx, cy - 0.16, text, ha='center', va='center',
                fontsize=fontsize - 1.5, color=fg, zorder=4,
                wrap=True, linespacing=1.3)
    else:
        ax.text(cx, cy, text, ha='center', va='center',
                fontsize=fontsize, color=fg, zorder=4,
                multialignment='center', linespacing=1.35)
    return (x + w / 2, y + h)   # top-centre anchor


def arrow(ax, x1, y1, x2, y2, color='#444', lw=1.4):
    ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle='->', color=color,
                                lw=lw, connectionstyle='arc3,rad=0'))


def dashed_arrow(ax, x1, y1, x2, y2, color='#888', lw=1.2):
    ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle='->', color=color, lw=lw,
                                linestyle='dashed',
                                connectionstyle='arc3,rad=0'))


def side_label(ax, x, y, h, text, color):
    ax.text(x, y + h / 2, text, ha='center', va='center',
            fontsize=8, color=color, rotation=90, fontweight='bold',
            bbox=dict(boxstyle='round,pad=0.2', fc=color + '22',
                      ec='none', zorder=1))


# ─────────────────────────────────────────────────────────────────────────────
# FIGURE 1 – System Architecture (5-layer pipeline)
# ─────────────────────────────────────────────────────────────────────────────

def fig1():
    FIG_W, FIG_H = 13, 11
    fig, ax = plt.subplots(figsize=(FIG_W, FIG_H))
    ax.set_xlim(0, FIG_W)
    ax.set_ylim(0, FIG_H)
    ax.axis('off')
    fig.patch.set_facecolor('#f7f7f7')

    # ── colour palette ──────────────────────────────────────────────────────
    C_INPUT  = '#cce5ff'
    C_CORE   = '#d4edda'
    C_NER    = '#e2d9f3'
    C_RISK   = '#fff3cd'
    C_OUT    = '#fde2e2'
    C_INTEG  = '#d6eaf8'
    BG_BAND  = '#eef2f7'

    # ── layer bands (background) ────────────────────────────────────────────
    bands = [
        (9.6, 10.6, '#dce8f5', 'Input'),
        (7.8,  9.2, '#ddf0e4', 'AI core'),
        (6.0,  7.4, '#ede6f8', ''),
        (4.2,  5.6, '#fef6d8', 'Output'),
        (2.0,  3.6, '#d6eaf8', 'Integration'),
    ]
    for y0, y1, fc, lbl in bands:
        bg = FancyBboxPatch((0.5, y0), FIG_W - 1.0, y1 - y0,
                            boxstyle='round,pad=0.05', linewidth=0.5,
                            edgecolor='#ccc', facecolor=fc, zorder=1)
        ax.add_patch(bg)
        if lbl:
            ax.text(0.85, (y0 + y1) / 2, lbl, ha='center', va='center',
                    fontsize=8, color='#444', rotation=90, fontweight='bold')

    # ── Layer 1: Input (4 boxes) ─────────────────────────────────────────────
    IY = 9.65; IH = 0.7
    boxes_in = [
        (1.0,  IY, 2.2, IH, 'Voice input\nPatient / CHW',       C_INPUT),
        (3.5,  IY, 2.2, IH, 'Text / form\nStructured Q&A',      C_INPUT),
        (6.0,  IY, 2.2, IH, 'Vital signs\nBP · temp · SpO2',    C_INPUT),
        (8.5,  IY, 2.2, IH, 'Medical history\nFrom HIS/EMR',    C_INPUT),
    ]
    in_bottoms = []
    for x, y, w, h, txt, c in boxes_in:
        rounded_box(ax, x, y, w, h, txt, c, fontsize=8.5)
        in_bottoms.append((x + w / 2, y))  # bottom-centre for downward arrow

    # ── Layer 2: AI core (3 boxes) ───────────────────────────────────────────
    CY = 8.0; CH = 0.7
    boxes_core = [
        (1.0, CY, 3.0, CH,
         'Thai Medical ASR\nWhisper fine-tuned · WER ≤10%', C_CORE),
        (4.4, CY, 2.5, CH,
         'Preprocessor\nText clean + normalise',             C_CORE),
        (7.3, CY, 3.4, CH,
         'HIS connector\nHL7 FHIR / API',                    C_CORE),
    ]
    core_tops = []
    core_bots = []
    for x, y, w, h, txt, c in boxes_core:
        rounded_box(ax, x, y, w, h, txt, c, fontsize=8.5)
        core_tops.append((x + w / 2, y + h))
        core_bots.append((x + w / 2, y))

    # arrows: input → core
    # Voice → ASR
    arrow(ax, in_bottoms[0][0], in_bottoms[0][1],
              core_tops[0][0],  core_tops[0][1])
    # Text → Preprocessor
    arrow(ax, in_bottoms[1][0], in_bottoms[1][1],
              core_tops[1][0],  core_tops[1][1])
    # Vital signs → Preprocessor (angled)
    arrow(ax, in_bottoms[2][0], in_bottoms[2][1],
              core_tops[1][0],  core_tops[1][1])
    # Medical history → HIS connector
    arrow(ax, in_bottoms[3][0], in_bottoms[3][1],
              core_tops[2][0],  core_tops[2][1])

    # ── Layer 2b: NER + Slot filling (wide box) ───────────────────────────────
    NY = 6.2; NH = 0.75
    NX = 1.0; NW = FIG_W - 2.0
    rounded_box(ax, NX, NY, NW, NH,
                'NER + Slot filling\n'
                'Symptoms · exposure · duration · medications · BP · temp · SpO2\n'
                'GLiNER zero-shot + regex rules',
                C_NER, fontsize=8.5)
    ner_top = (NX + NW / 2, NY + NH)
    ner_bot = (NX + NW / 2, NY)

    # arrows: core → NER
    for bx, by in core_bots:
        arrow(ax, bx, by, bx, NY + NH)

    # ── Layer 3: AI risk model ─────────────────────────────────────────────
    RY = 4.4; RH = 0.75
    RX = 1.0; RW = FIG_W - 2.0
    rounded_box(ax, RX, RY, RW, RH,
                'AI risk model\n'
                'Bayesian prior (CCRU prevalence data) + LLM reasoning layer\n'
                'Diseases: scrub typhus · dengue · leptospirosis · bacterial sepsis',
                C_RISK, fontsize=8.5)
    risk_top = (RX + RW / 2, RY + RH)
    risk_bot = (RX + RW / 2, RY)

    arrow(ax, ner_bot[0], ner_bot[1], risk_top[0], risk_top[1])

    # ── Layer 4: Output (3 boxes) ─────────────────────────────────────────
    OY = 2.3; OH = 0.85
    BW = (FIG_W - 2.6) / 3   # equal widths
    boxes_out = [
        (1.0,         OY, BW, OH,
         'Triage recommendation\nManage / refer / urgent\n+ rationale in Thai', C_OUT),
        (1.0 + BW + 0.3, OY, BW, OH,
         'Draft SOAP note\nS / O / A auto-filled\nDoctor approves P section',   C_OUT),
        (1.0 + 2*(BW + 0.3), OY, BW, OH,
         'AMR signal\nAntibiotic stewardship\nflag + guidance',                 C_OUT),
    ]
    out_tops = []
    out_bots = []
    for x, y, w, h, txt, c in boxes_out:
        rounded_box(ax, x, y, w, h, txt, c, fontsize=8.5)
        out_tops.append((x + w / 2, y + h))
        out_bots.append((x + w / 2, y))

    # risk → each output box
    for ox, oy in out_tops:
        arrow(ax, risk_bot[0], risk_bot[1], ox, oy)

    # ── Layer 5: Integration (2 boxes) ────────────────────────────────────
    IgY = 0.5; IgH = 0.8
    IgW1 = 5.5; IgW2 = 3.5
    IgX1 = 1.0
    IgX2 = FIG_W - 1.0 - IgW2
    rounded_box(ax, IgX1, IgY, IgW1, IgH,
                'Clinician review & approve\nDoctor confirms before saving to record',
                C_INTEG, fontsize=8.5)
    rounded_box(ax, IgX2, IgY, IgW2, IgH,
                'HIS / EMR\nFHIR R4 export',
                C_INTEG, fontsize=8.5)

    clin_top  = (IgX1 + IgW1 / 2, IgY + IgH)
    clin_right = (IgX1 + IgW1, IgY + IgH / 2)
    his_top   = (IgX2 + IgW2 / 2, IgY + IgH)
    his_left  = (IgX2, IgY + IgH / 2)

    # output boxes → clinician
    for ox, oy in out_bots:
        arrow(ax, ox, oy, clin_top[0], clin_top[1])

    # clinician → HIS (horizontal)
    arrow(ax, clin_right[0], clin_right[1], his_left[0], his_left[1])

    # ── Title ──────────────────────────────────────────────────────────────
    ax.text(FIG_W / 2, FIG_H - 0.25,
            u'\u0e23\u0e39\u0e1b\u0e17\u0e35\u0e48 3-1  System Architecture \u2014 AI Acute Febrile Illness Screening',
            ha='center', va='top', fontsize=11, fontweight='bold', color='#222')

    plt.tight_layout(pad=0.3)
    out = r'G:\My Drive\Research\MORU\word\media\image1.png'
    fig.savefig(out, dpi=150, bbox_inches='tight', facecolor=fig.get_facecolor())
    plt.close(fig)
    print('Saved Figure 1')


# ─────────────────────────────────────────────────────────────────────────────
# FIGURE 2 – Patient Journey (swimlane flowchart)
# ─────────────────────────────────────────────────────────────────────────────

def fig2():
    FIG_W, FIG_H = 13, 15
    fig, ax = plt.subplots(figsize=(FIG_W, FIG_H))
    ax.set_xlim(0, FIG_W)
    ax.set_ylim(0, FIG_H)
    ax.axis('off')
    fig.patch.set_facecolor('#f7f7f7')

    # Lane definitions (x_start, width, header_color, bg_color, label)
    lanes = [
        (0.3,  2.8, '#2e7d7d', '#d9f0f0', 'Patient / CHW'),
        (3.3,  3.4, '#5b6abf', '#dde2f5', 'AI system'),
        (6.9,  3.0, '#8b5e3c', '#f5ead9', 'Clinician'),
        (10.1, 2.6, '#4a4a6a', '#e8e8f0', 'HIS / System'),
    ]

    HEADER_H = 0.55
    LANE_Y_TOP = FIG_H - 0.4

    # Draw lane backgrounds and headers
    for lx, lw, hc, bc, lbl in lanes:
        # background
        bg = FancyBboxPatch((lx, 0.3), lw, LANE_Y_TOP - 0.3,
                            boxstyle='square,pad=0', linewidth=0.8,
                            edgecolor='#aaa', facecolor=bc, zorder=0)
        ax.add_patch(bg)
        # header
        hdr = FancyBboxPatch((lx, LANE_Y_TOP - HEADER_H), lw, HEADER_H,
                             boxstyle='square,pad=0', linewidth=0,
                             facecolor=hc, zorder=1)
        ax.add_patch(hdr)
        ax.text(lx + lw / 2, LANE_Y_TOP - HEADER_H / 2, lbl,
                ha='center', va='center', fontsize=9.5,
                fontweight='bold', color='white', zorder=2)

    # ── Helper: lane centre x ──────────────────────────────────────────────
    def lcx(lane_idx):
        lx, lw = lanes[lane_idx][0], lanes[lane_idx][1]
        return lx + lw / 2

    # ── Box dimensions ─────────────────────────────────────────────────────
    BW   = 2.3   # standard box width
    BH   = 0.65  # standard box height
    DH   = 0.55  # diamond height
    DW   = 1.8   # diamond width

    def box(lane, cy, txt, bg=None, bh=BH):
        """Draw box centred at (lcx(lane), cy); return top and bottom centres."""
        if bg is None:
            bg = lanes[lane][3]
        darker = '#ffffff'
        lx = lcx(lane) - BW / 2
        ly = cy - bh / 2
        rounded_box(ax, lx, ly, BW, bh, txt, bg,
                    fontsize=8.5, fg='#1a1a2e')
        return (lcx(lane), cy + bh / 2), (lcx(lane), cy - bh / 2)

    def diamond(lane, cy, txt):
        """Draw a decision diamond; return top, bottom, left, right centres."""
        cx = lcx(lane)
        pts = [(cx, cy + DH / 2),
               (cx + DW / 2, cy),
               (cx, cy - DH / 2),
               (cx - DW / 2, cy)]
        poly = plt.Polygon(pts, closed=True, facecolor='#fffacd',
                           edgecolor='#888', linewidth=1.2, zorder=3)
        ax.add_patch(poly)
        ax.text(cx, cy, txt, ha='center', va='center',
                fontsize=8.5, zorder=4, color='#222')
        return ((cx, cy + DH / 2),   # top
                (cx, cy - DH / 2),   # bottom
                (cx - DW / 2, cy),   # left
                (cx + DW / 2, cy))   # right

    # ── Place nodes (cy = vertical centre, lane 0-3) ──────────────────────
    # Row positions (y) — enough vertical spacing
    Y = [13.8, 12.9, 12.0, 11.0, 9.9, 8.8, 7.9, 6.5, 5.5, 4.4, 3.3, 2.2, 1.2]

    # Lane 0 – Patient / CHW
    t_arrives, b_arrives  = box(0, Y[0],  'Patient arrives\nFever ≥38°C',          '#b2dfdb')
    t_opens,   b_opens    = box(0, Y[1],  'CHW opens app\nLINE OA / web',          '#b2dfdb')
    t_answers, b_answers  = box(0, Y[2],  'Answers questions\nBy voice (Thai)\nor text input', '#b2dfdb')
    t_manage,  b_manage   = box(0, Y[7],  'Manage locally\nHome care advice',       '#b2dfdb')

    # Lane 1 – AI system
    t_pull,    b_pull     = box(1, Y[1],  'Pull HIS record\nFHIR lookup',           '#c5cae9')
    t_asr,     b_asr      = box(1, Y[3],  'ASR → NER\nExtract symptoms,\nexposure, duration', '#c5cae9', bh=0.8)
    t_risk,    b_risk      = box(1, Y[5],  'Risk model\nBayesian + LLM\nProbability per Dx',  '#c5cae9', bh=0.8)
    t_draft,   b_draft    = box(1, Y[9],  'Draft SOAP\n+ AMR flag\nS/O/A auto-filled',       '#c5cae9', bh=0.8)

    # Diamond for severity (lane 1)
    d_sev = diamond(1, Y[6], 'Severity\nlevel?')
    d_top, d_bot, d_left, d_right = d_sev

    # Lane 2 – Clinician
    t_refer,   b_refer    = box(2, Y[7],  'Refer to hospital\n+ referral note',     '#f0d9c0')
    t_review,  b_review   = box(2, Y[9],  'Review & edit\nComplete Plan (P)\nOne-click approve', '#f0d9c0', bh=0.8)

    # Lane 3 – HIS / System
    t_his_r,   b_his_r    = box(3, Y[1],  'HIS response\nDx history, meds',         '#d1c4e9')
    t_save,    b_save     = box(3, Y[9],  'Save to HIS\nFHIR R4 POST\nAudit log entry', '#d1c4e9', bh=0.8)

    # Bottom bar: Outcome tracking
    BX, BY, BW2, BH2 = 1.0, 0.38, FIG_W - 2.0, 0.6
    rounded_box(ax, BX, BY, BW2, BH2,
                'Outcome tracking & follow-up alert   |   Day 3 / Day 7 auto-ping to CHW',
                '#b3e5fc', fg='#01579b', fontsize=9)
    track_top = (BX + BW2 / 2, BY + BH2)

    # ── Arrows ────────────────────────────────────────────────────────────
    # Patient arrives → CHW opens app
    arrow(ax, b_arrives[0], b_arrives[1], t_opens[0], t_opens[1])
    # CHW opens → Pull HIS (horizontal)
    arrow(ax, lcx(0) + BW/2, Y[1], lcx(1) - BW/2, Y[1])
    # Pull HIS → HIS response (horizontal)
    arrow(ax, lcx(1) + BW/2, Y[1], lcx(3) - BW/2, Y[1])
    # HIS response → Pull HIS (dashed return)
    dashed_arrow(ax, lcx(3) - BW/2, Y[1] - 0.15, lcx(1) + BW/2, Y[1] - 0.15, color='#999')
    # CHW opens → answers
    arrow(ax, b_opens[0], b_opens[1], t_answers[0], t_answers[1])
    # Answers → ASR/NER (vertical then cross-lane)
    arrow(ax, b_answers[0], b_answers[1], t_asr[0], t_asr[1])
    # ASR → Risk model
    arrow(ax, b_asr[0], b_asr[1], t_risk[0], t_risk[1])
    # Risk → Diamond top
    arrow(ax, b_risk[0], b_risk[1], d_top[0], d_top[1])
    # Diamond left → Manage locally (Low)
    arrow(ax, d_left[0], d_left[1], lcx(0) + BW/2, Y[7])
    ax.text(d_left[0] - 0.1, d_left[1] + 0.08, 'Low', fontsize=8, color='#555', ha='right')
    # Diamond right → Refer to hospital (High)
    arrow(ax, d_right[0], d_right[1], lcx(2) - BW/2, Y[7])
    ax.text(d_right[0] + 0.1, d_right[1] + 0.08, 'High', fontsize=8, color='#555', ha='left')
    # Diamond bottom → Draft SOAP
    arrow(ax, d_bot[0], d_bot[1], t_draft[0], t_draft[1])
    # Draft SOAP → Review & edit (horizontal)
    arrow(ax, lcx(1) + BW/2, Y[9], lcx(2) - BW/2, Y[9])
    # Review & edit → Save to HIS (horizontal)
    arrow(ax, lcx(2) + BW/2, Y[9], lcx(3) - BW/2, Y[9])
    # Save to HIS → Outcome tracking (bottom bar)
    arrow(ax, lcx(3), b_save[1], lcx(3), track_top[1])
    # Draft SOAP → Outcome tracking
    arrow(ax, lcx(1), b_draft[1], lcx(1), track_top[1])

    # ── Title ────────────────────────────────────────────────────────────
    ax.text(FIG_W / 2, FIG_H - 0.15,
            'รูปที่ 3-2  Patient Journey — AI Acute Febrile Illness Screening',
            ha='center', va='top', fontsize=11, fontweight='bold', color='#222')

    plt.tight_layout(pad=0.3)
    out = r'G:\My Drive\Research\MORU\word\media\image2.png'
    fig.savefig(out, dpi=150, bbox_inches='tight', facecolor=fig.get_facecolor())
    plt.close(fig)
    print('Saved Figure 2')


if __name__ == '__main__':
    fig1()
    fig2()
    print('Done.')
