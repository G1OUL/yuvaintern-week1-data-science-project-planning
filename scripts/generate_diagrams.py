import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np
import os

# Dynamically resolve project directory
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
output_dir = os.path.join(BASE_DIR, "diagrams")
os.makedirs(output_dir, exist_ok=True)

# Set common style
plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
plt.rcParams['font.family'] = 'sans-serif'

# ==========================================
# DIAGRAM 1: Data Science Lifecycle Flowchart
# ==========================================
fig, ax = plt.subplots(figsize=(12, 7), dpi=300)
ax.set_xlim(0, 12)
ax.set_ylim(0, 7)
ax.axis('off')

# Stage boxes definition: 8 stages arranged in two rows or circular/serpentine flow
stages = [
    ("1. Business\nUnderstanding", 1.5, 5.5, "#1F4E79"),
    ("2. Data\nCollection", 4.5, 5.5, "#203764"),
    ("3. Data Quality\n& Cleaning", 7.5, 5.5, "#2F5597"),
    ("4. Exploratory\nData Analysis", 10.5, 5.5, "#385723"),
    ("5. Feature\nEngineering", 10.5, 2.2, "#806000"),
    ("6. Machine Learning\nModeling", 7.5, 2.2, "#C65911"),
    ("7. Model\nEvaluation", 4.5, 2.2, "#A61C1C"),
    ("8. Communication\n& Monitoring", 1.5, 2.2, "#548235"),
]

box_w, box_h = 2.2, 1.2

for name, x, y, color in stages:
    rect = patches.FancyBboxPatch(
        (x - box_w/2, y - box_h/2), box_w, box_h,
        boxstyle="round,pad=0.1,rounding_size=0.2",
        linewidth=1.5, edgecolor="#2C3E50", facecolor=color, alpha=0.9
    )
    ax.add_patch(rect)
    ax.text(x, y, name, color="white", fontsize=10, weight="bold", ha="center", va="center")

# Forward arrows
# Row 1: 1 -> 2 -> 3 -> 4
for i in range(3):
    x1 = stages[i][1] + box_w/2
    y1 = stages[i][2]
    x2 = stages[i+1][1] - box_w/2
    y2 = stages[i+1][2]
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle="-|>", lw=2, color="#2C3E50", mutation_scale=15))

# Transition 4 -> 5 (down)
ax.annotate("", xy=(stages[4][1], stages[4][2] + box_h/2),
            xytext=(stages[3][1], stages[3][2] - box_h/2),
            arrowprops=dict(arrowstyle="-|>", lw=2, color="#2C3E50", mutation_scale=15))

# Row 2 (right to left): 5 -> 6 -> 7 -> 8
for i in range(4, 7):
    x1 = stages[i][1] - box_w/2
    y1 = stages[i][2]
    x2 = stages[i+1][1] + box_w/2
    y2 = stages[i+1][2]
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle="-|>", lw=2, color="#2C3E50", mutation_scale=15))

# Feedback loops
# 1. From Evaluation (7) back to Modeling (6) / Feature Engineering (5) / Cleaning (3)
ax.annotate("Model Refinement &\nFeature Iteration Loop",
            xy=(stages[4][1], stages[4][2] - box_h/2),
            xytext=(stages[6][1], stages[6][2] - box_h/2),
            arrowprops=dict(arrowstyle="-|>", lw=1.8, color="#C00000", linestyle="dashed",
                            connectionstyle="arc3,rad=0.35", mutation_scale=14),
            fontsize=9, color="#900C3F", weight="bold", ha="center", va="top")

# 2. Feedback loop from EDA (4) back to Data Cleaning (3)
ax.annotate("Data Cleansing & Schema Audit Loop",
            xy=(stages[2][1], stages[2][2] + box_h/2),
            xytext=(stages[3][1], stages[3][2] + box_h/2),
            arrowprops=dict(arrowstyle="-|>", lw=1.8, color="#0070C0", linestyle="dashed",
                            connectionstyle="arc3,rad=-0.4", mutation_scale=14),
            fontsize=9, color="#004D80", weight="bold", ha="center", va="bottom")

# 3. Long feedback loop from Evaluation (7) back to Business Understanding (1)
ax.annotate("Business Metric & KPI Realignment",
            xy=(stages[0][1], stages[0][2] - box_h/2),
            xytext=(stages[6][1] - 0.5, stages[6][2] + box_h/2),
            arrowprops=dict(arrowstyle="-|>", lw=1.8, color="#7030A0", linestyle="dotted",
                            connectionstyle="arc3,rad=0.25", mutation_scale=14),
            fontsize=9, color="#5B2C6F", weight="bold", ha="center", va="center")

# Title and subtitle banner
ax.text(6.0, 6.7, "END-TO-END DATA SCIENCE LIFECYCLE FLOWCHART",
        fontsize=14, weight="heavy", color="#1B365D", ha="center")
ax.text(6.0, 6.35, "Customer Churn Prediction: Iterative Exploration, Modeling, and Realignment Loops",
        fontsize=10, style="italic", color="#555555", ha="center")

plt.tight_layout()
diag1_path = os.path.join(output_dir, "diagram1_lifecycle_flowchart.png")
plt.savefig(diag1_path, dpi=300, bbox_inches='tight')
plt.close()
print("Saved Diagram 1 to", diag1_path)

# ==========================================
# DIAGRAM 2: Conceptual Project Architecture
# ==========================================
fig, ax = plt.subplots(figsize=(13, 7.5), dpi=300)
ax.set_xlim(0, 13)
ax.set_ylim(0, 8)
ax.axis('off')

# Title
ax.text(6.5, 7.6, "CONCEPTUAL PROJECT ARCHITECTURE & DATA FLOW",
        fontsize=14, weight="heavy", color="#1B365D", ha="center")
ax.text(6.5, 7.25, "Structured Flow from Ingestion, Governance, Feature Pipeline, to Retention Decisioning",
        fontsize=10, style="italic", color="#555555", ha="center")

# Pillars / Layers
layers = [
    ("Approved Hypothetical\nData Sources\n(CRM, Billing, Usage)", 1.8, 4.5, 2.2, 2.6, "#2E4053"),
    ("Ingestion & Raw Staging\n(Data Provenance &\nAudit Logs)", 4.6, 4.5, 2.2, 2.6, "#2874A6"),
    ("Data Validation &\nPreparation Pipeline\n(Missing, Scaling, Encode)", 7.4, 4.5, 2.2, 2.6, "#148F77"),
    ("Feature Store &\nModel Experiments\n(Baselines, LogReg, RF)", 10.2, 4.5, 2.2, 2.6, "#B9770E"),
]

for title, x, y, w, h, col in layers:
    rect = patches.FancyBboxPatch(
        (x - w/2, y - h/2), w, h,
        boxstyle="round,pad=0.1,rounding_size=0.15",
        linewidth=1.5, edgecolor="#1B2631", facecolor=col, alpha=0.9
    )
    ax.add_patch(rect)
    ax.text(x, y + 0.3, title, color="white", fontsize=9.5, weight="bold", ha="center", va="center")

# Output stage at the bottom
output_rect1 = patches.FancyBboxPatch(
    (5.0 - 2.0, 1.3 - 0.7), 4.0, 1.3,
    boxstyle="round,pad=0.1,rounding_size=0.15",
    linewidth=1.5, edgecolor="#1B2631", facecolor="#922B21", alpha=0.9
)
ax.add_patch(output_rect1)
ax.text(5.0, 1.3, "Model Evaluation & Review\n(ROC-AUC, Precision, Recall, Subgroups)",
        color="white", fontsize=9.5, weight="bold", ha="center", va="center")

output_rect2 = patches.FancyBboxPatch(
    (9.8 - 1.8, 1.3 - 0.7), 3.8, 1.3,
    boxstyle="round,pad=0.1,rounding_size=0.15",
    linewidth=1.5, edgecolor="#1B2631", facecolor="#1E8449", alpha=0.9
)
ax.add_patch(output_rect2)
ax.text(9.8, 1.3, "Actionable Business Recommendations\n(Retention Campaigns & Intervention Capacity)",
        color="white", fontsize=9.5, weight="bold", ha="center", va="center")

# Connecting arrows across layers
for i in range(3):
    x1 = layers[i][1] + layers[i][3]/2
    y1 = layers[i][2]
    x2 = layers[i+1][1] - layers[i+1][3]/2
    y2 = layers[i+1][2]
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle="-|>", lw=2.2, color="#1B365D", mutation_scale=16))

# Layer 4 down to Evaluation
ax.annotate("", xy=(5.0 + 1.2, 1.3 + 0.65), xytext=(10.2 - 0.5, 4.5 - 1.3),
            arrowprops=dict(arrowstyle="-|>", lw=2.2, color="#1B365D", mutation_scale=16))

# Evaluation to Business Recommendations
ax.annotate("", xy=(9.8 - 1.8, 1.3), xytext=(5.0 + 2.0, 1.3),
            arrowprops=dict(arrowstyle="-|>", lw=2.2, color="#1B365D", mutation_scale=16))

# Governance & Privacy banners across the top and side
gov_box = patches.FancyBboxPatch(
    (0.7, 6.2), 11.6, 0.65,
    boxstyle="square,pad=0.05",
    linewidth=1, edgecolor="#C0392B", facecolor="#FDEDEC", linestyle="--"
)
ax.add_patch(gov_box)
ax.text(6.5, 6.52, "[GOVERNANCE & PRIVACY LAYER]: Data Minimization • Strict Access Control • Anonymized Customer Identifiers • No PII Leakage",
        fontsize=9, weight="bold", color="#922B21", ha="center", va="center")

leakage_box = patches.FancyBboxPatch(
    (0.7, 2.45), 11.6, 0.6,
    boxstyle="square,pad=0.05",
    linewidth=1, edgecolor="#27AE60", facecolor="#EAFAF1", linestyle="--"
)
ax.add_patch(leakage_box)
ax.text(6.5, 2.75, "[INTEGRITY & PROVENANCE LAYER]: Immutable Raw Data • Preprocessors Fitted on Train Only • Temporal Split Horizon Audit",
        fontsize=9, weight="bold", color="#196F3D", ha="center", va="center")

plt.tight_layout()
diag2_path = os.path.join(output_dir, "diagram2_conceptual_architecture.png")
plt.savefig(diag2_path, dpi=300, bbox_inches='tight')
plt.close()
print("Saved Diagram 2 to", diag2_path)

# ==========================================
# DIAGRAM 3: 33-Hour Project Timeline
# ==========================================
fig, ax = plt.subplots(figsize=(12, 6.5), dpi=300)

activities = [
    "1. Problem framing & success criteria",
    "2. Data-source & privacy planning",
    "3. Data quality & preparation strategy",
    "4. Exploratory analysis & feature planning",
    "5. Baseline & model strategy",
    "6. Evaluation & responsible-use planning",
    "7. Deliverables, risks, & documentation",
    "8. Final review & presentation"
]
hours = [3, 3, 5, 4, 5, 4, 5, 4]
colors = ["#1F4E79", "#2E75B6", "#008080", "#2E8B57", "#D9822B", "#C0392B", "#7D3C98", "#34495E"]

y_pos = np.arange(len(activities))

# Horizontal bar chart
bars = ax.barh(y_pos, hours, align='center', color=colors, edgecolor='#1B2631', height=0.62, alpha=0.92)

# Value labels on bars
for bar, h in zip(bars, hours):
    ax.text(bar.get_width() + 0.12, bar.get_y() + bar.get_height()/2,
            f"{h} Hours", va='center', ha='left', fontsize=10, weight='bold', color='#1B365D')

ax.set_yticks(y_pos)
ax.set_yticklabels(activities, fontsize=10.5, weight='semibold', color='#2C3E50')
ax.invert_yaxis()  # labels read top-to-bottom
ax.set_xlabel('Allocated Effort (Hours)', fontsize=11, weight='bold', color='#1B365D')
ax.set_xlim(0, 6.5)

# Styling
ax.grid(axis='x', linestyle='--', alpha=0.5, color='#95A5A6')
ax.set_axisbelow(True)

# Title & Summary callout
ax.set_title("PROJECT WORKSTREAM ALLOCATION (TOTAL: 33 HOURS)\nWeek 1: Data Science Project Planning & Strategy Design",
             fontsize=13, weight='heavy', color='#1B365D', pad=18)

# Callout box for total effort
total_hours = sum(hours)
callout_text = f"Total Planned Effort: {total_hours} Hours\nAllocation Range: 30 - 35 Hours Target Met"
ax.text(4.2, 7.2, callout_text, fontsize=10, weight='bold', color='#154360',
        bbox=dict(boxstyle='round,pad=0.6', facecolor='#EBF5FB', edgecolor='#2980B9', lw=1.5))

plt.tight_layout()
diag3_path = os.path.join(output_dir, "diagram3_project_timeline.png")
plt.savefig(diag3_path, dpi=300, bbox_inches='tight')
plt.close()
print("Saved Diagram 3 to", diag3_path)
