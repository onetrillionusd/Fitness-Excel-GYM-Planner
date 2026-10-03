from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, numbers
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta


def style_header(ws, row=1, start_col=1, end_col=None, color="1F4E78"):
    end_col = ws.max_column if end_col is None else end_col
    for cell in ws.iter_cols(min_row=row, max_row=row, min_col=start_col, max_col=end_col):
        for c in cell:
            c.font = Font(bold=True, color="FFFFFF", size=11)
            c.fill = PatternFill("solid", fgColor=color)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = Border(
                left=Side(style="medium", color="1F4E78"),
                right=Side(style="medium", color="1F4E78"),
                top=Side(style="medium", color="1F4E78"),
                bottom=Side(style="medium", color="1F4E78"),
            )


def set_col_widths(ws, widths):
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def add_border(ws, min_row, max_row, min_col, max_col, color="D9D9D9"):
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = Border(
                left=Side(style="thin", color=color),
                right=Side(style="thin", color=color),
                top=Side(style="thin", color=color),
                bottom=Side(style="thin", color=color),
            )


def build_premium_dashboard(ws):
    ws.title = "Dashboard"
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    
    # Title
    ws.merge_cells("A1:F1")
    ws["A1"] = "💪 STRENGTH & SHRED ELITE PROGRAM"
    ws["A1"].font = Font(size=20, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="C00000")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

    # KPI Cards Section
    ws["A3"] = "WEEKLY PERFORMANCE METRICS"
    ws["A3"].font = Font(size=13, bold=True, color="FFFFFF")
    ws["A3"].fill = PatternFill("solid", fgColor="1F4E78")
    ws.merge_cells("A3:F3")

    # Card 1: Sessions Completed
    ws["A5"] = "Sessions Completed"
    ws["A5"].font = Font(bold=True, color="FFFFFF", size=10)
    ws["A5"].fill = PatternFill("solid", fgColor="4472C4")
    ws.merge_cells("A5:B5")
    ws["A6"] = "=COUNTIF('Weekly Tracker'!B:B,TRUE)"
    ws["A6"].font = Font(size=16, bold=True, color="4472C4")
    ws["A6"].fill = PatternFill("solid", fgColor="E7F0F7")
    ws.merge_cells("A6:B6")

    # Card 2: Weekly Volume
    ws["D5"] = "Total Volume (kg)"
    ws["D5"].font = Font(bold=True, color="FFFFFF", size=10)
    ws["D5"].fill = PatternFill("solid", fgColor="70AD47")
    ws.merge_cells("D5:E5")
    ws["D6"] = "=IFERROR(SUMIF('Workout Log'!A:A,">0",'Workout Log'!H:H),0)"
    ws["D6"].font = Font(size=16, bold=True, color="70AD47")
    ws["D6"].fill = PatternFill("solid", fgColor="E8F5E0")
    ws.merge_cells("D6:E6")

    # Card 3: Avg Weight
    ws["A8"] = "Current Bodyweight"
    ws["A8"].font = Font(bold=True, color="FFFFFF", size=10)
    ws["A8"].fill = PatternFill("solid", fgColor="FF6B35")
    ws.merge_cells("A8:B8")
    ws["A9"] = "=IFERROR(INDEX('Progress Tracker'!B:B,MATCH(MAX('Progress Tracker'!A:A),'Progress Tracker'!A:A,0)),\"-\")"
    ws["A9"].font = Font(size=16, bold=True, color="FF6B35")
    ws["A9"].fill = PatternFill("solid", fgColor="FFF0E8")
    ws.merge_cells("A8:B8")
    ws.merge_cells("A9:B9")

    # Card 4: Fat Loss Target
    ws["D8"] = "Waist Target"
    ws["D8"].font = Font(bold=True, color="FFFFFF", size=10)
    ws["D8"].fill = PatternFill("solid", fgColor="7030A0")
    ws.merge_cells("D8:E8")
    ws["D9"] = "80-82 cm"
    ws["D9"].font = Font(size=16, bold=True, color="7030A0")
    ws["D9"].fill = PatternFill("solid", fgColor="F3E8F5")
    ws.merge_cells("D9:E9")

    # Progress Section
    ws["A12"] = "PROGRESS SNAPSHOT"
    ws["A12"].font = Font(size=13, bold=True, color="FFFFFF")
    ws["A12"].fill = PatternFill("solid", fgColor="1F4E78")
    ws.merge_cells("A12:F12")

    progress_data = [
        ["Metric", "Current", "Target", "Progress", "Status"],
        ["Bench Press", "95 kg", "105 kg", "=95/105", "In Progress"],
        ["Waist Circumference", "84 cm", "80 cm", "=(84-80)/84", "On Track"],
        ["Bodyweight", "78.7 kg", "76 kg", "=(78.7-76)/78.7", "On Track"],
    ]
    for idx, row_data in enumerate(progress_data, 13):
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=idx, column=col_idx, value=value)
            if idx == 13:
                cell.font = Font(bold=True, color="FFFFFF", size=10)
                cell.fill = PatternFill("solid", fgColor="366092")
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.alignment = Alignment(horizontal="left")
            cell.border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )

    # Format progress percentages
    ws["E14"].number_format = "0%"
    ws["E15"].number_format = "0%"
    ws["E16"].number_format = "0%"

    # Coaching Tips
    ws["A18"] = "🎯 THIS WEEK'S FOCUS"
    ws["A18"].font = Font(size=12, bold=True, color="FFFFFF")
    ws["A18"].fill = PatternFill("solid", fgColor="70AD47")
    ws.merge_cells("A18:F18")

    tips = [
        "✓ Complete all 3 sessions (Monday/Tuesday/Thursday)",
        "✓ Maintain 1.8g protein per kg bodyweight",
        "✓ Stay in 300-500 kcal deficit",
        "✓ Log all meals in Nutrition Tracker sheet",
        "✓ Sleep minimum 7 hours nightly",
    ]
    for idx, tip in enumerate(tips, 19):
        ws[f"A{idx}"] = tip
        ws[f"A{idx}"].font = Font(size=10, color="1F4E78")
        ws.merge_cells(f"A{idx}:F{idx}")

    set_col_widths(ws, {"A": 18, "B": 14, "C": 14, "D": 18, "E": 14, "F": 14})


def build_weekly_tracker_premium(ws):
    ws.title = "Weekly Tracker"
    ws.freeze_panes = "A2"
    
    ws.merge_cells("A1:J1")
    ws["A1"] = "WEEKLY TRAINING TRACKER"
    ws["A1"].font = Font(size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="1F4E78")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 25

    headers = [
        "Week Starting",
        "Mon ✓/✗",
        "Mon Workout",
        "Tue ✓/✗",
        "Tue Workout",
        "Thu ✓/✗",
        "Thu Workout",
        "Sat ✓/✗",
        "Sat Workout",
        "Weekly Notes",
    ]
    ws.append(headers)
    style_header(ws, row=2, start_col=1, end_col=10, color="1F4E78")

    sample_rows = [
        ["2026-10-05", "✓", "Push", "✓", "Pull", "✓", "Legs", "✗", "", "Great week - all sessions completed"],
        ["2026-10-12", "✓", "Push", "✓", "Pull", "✓", "Legs", "✓", "30min cardio", "Added Saturday - good energy"],
        ["2026-10-19", "✓", "Push", "✓", "Pull", "✓", "Legs", "✗", "", "Solid consistency"],
    ]
    for row in sample_rows:
        ws.append(row)

    add_border(ws, 3, ws.max_row, 1, 10)

    set_col_widths(ws, {"A": 15, "B": 10, "C": 18, "D": 10, "E": 18, "F": 10, "G": 18, "H": 10, "I": 18, "J": 32})


def build_workout_log_premium(ws):
    ws.title = "Workout Log"
    ws.freeze_panes = "A2"

    ws.merge_cells("A1:K1")
    ws["A1"] = "DETAILED WORKOUT LOG - STRENGTH & BODYBUILDING"
    ws["A1"].font = Font(size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="C00000")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 25

    headers = [
        "Date",
        "Week",
        "Session",
        "Exercise",
        "Muscle",
        "Sets",
        "Reps",
        "Weight (kg)",
        "Rest (sec)",
        "RPE",
        "Notes",
    ]
    ws.append(headers)
    style_header(ws, row=2, color="C00000")

    sample_rows = [
        ["2026-10-05", "1", "PUSH", "Bench Press", "Chest", "4", "5-6", "95", "180", "8", "Strong - add 2.5kg next week"],
        ["2026-10-05", "1", "PUSH", "Incline DB Press", "Chest", "3", "8-10", "35", "120", "7", "Good form"],
        ["2026-10-05", "1", "PUSH", "Shoulder Press", "Shoulders", "3", "6-8", "65", "120", "7", "Controlled"],
        ["2026-10-05", "1", "PUSH", "Tricep Dips", "Triceps", "3", "6-8", "BW+15", "120", "8", "Weighted"],
        ["2026-10-06", "1", "PULL", "Deadlift", "Back", "4", "3-5", "140", "240", "9", "Excellent grip"],
        ["2026-10-06", "1", "PULL", "Barbell Row", "Back", "4", "6-8", "115", "180", "8", "Explosive"],
        ["2026-10-07", "1", "LEGS", "Back Squat", "Legs", "4", "5-6", "120", "180", "8", "Deep reps"],
        ["2026-10-07", "1", "LEGS", "RDL", "Hamstrings", "3", "8-10", "100", "120", "7", "Good stretch"],
    ]
    for row in sample_rows:
        ws.append(row)

    add_border(ws, 3, ws.max_row, 1, 11)

    # Add data validation for Session type
    from openpyxl.worksheet.datavalidation import DataValidation
    dv = DataValidation(type="list", formula1='"PUSH,PULL,LEGS,CONDITIONING"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(f"C3:C1000")

    set_col_widths(ws, {"A": 12, "B": 8, "C": 14, "D": 22, "E": 14, "F": 8, "G": 10, "H": 12, "I": 12, "J": 10, "K": 28})


def build_nutrition_premium(ws):
    ws.title = "Nutrition Tracker"
    ws.freeze_panes = "A2"

    ws.merge_cells("A1:J1")
    ws["A1"] = "🥗 DAILY NUTRITION & MACRO TRACKER"
    ws["A1"].font = Font(size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="70AD47")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 25

    headers = [
        "Date",
        "Meal",
        "Food Item",
        "Quantity",
        "Protein (g)",
        "Carbs (g)",
        "Fat (g)",
        "Calories",
        "Water (L)",
        "Notes",
    ]
    ws.append(headers)
    style_header(ws, row=2, color="70AD47")

    sample_rows = [
        ["2026-10-05", "Breakfast", "Eggs (3) + Oatmeal", "3 eggs + 50g", "18", "25", "8", "240", "0.5", "Solid carbs"],
        ["2026-10-05", "Snack 1", "Greek Yogurt + Berries", "150g + 100g", "20", "12", "2", "150", "0.3", "Protein boost"],
        ["2026-10-05", "Lunch", "Chicken + Rice + Veg", "200g + 80g + 150g", "48", "50", "6", "450", "0.7", "Post-workout"],
        ["2026-10-05", "Snack 2", "Protein Shake", "1 shake (30g)", "30", "35", "2", "280", "0.3", "Pre-sleep"],
        ["2026-10-05", "Dinner", "Salmon + Broccoli", "180g + 200g", "35", "15", "12", "360", "0.7", "Omega-3 rich"],
        ["", "DAILY TOTAL", "", "", "=SUM(E3:E7)", "=SUM(F3:F7)", "=SUM(G3:G7)", "=SUM(H3:H7)", "=SUM(I3:I7)", ""],
    ]
    for row_idx, row in enumerate(sample_rows, 3):
        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            if row_idx == 8:  # Total row
                cell.font = Font(bold=True, color="FFFFFF", size=11)
                cell.fill = PatternFill("solid", fgColor="366092")
            cell.border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )

    # Macro breakdown
    ws["A10"] = "DAILY MACRO TARGETS & RECOMMENDATIONS"
    ws["A10"].font = Font(size=12, bold=True, color="FFFFFF")
    ws["A10"].fill = PatternFill("solid", fgColor="70AD47")
    ws.merge_cells("A10:J10")

    macro_recs = [
        ["Macro", "Daily Target", "Current", "% of Calories", "Status"],
        ["Protein (g)", "170", "=E8", "=(E8*4)/(H8)*100", ""],
        ["Carbs (g)", "280", "=F8", "=(F8*4)/(H8)*100", ""],
        ["Fats (g)", "70", "=G8", "=(G8*9)/(H8)*100", ""],
        ["Total Calories", "2100-2150", "=H8", "100%", ""],
    ]

    for idx, row_data in enumerate(macro_recs, 11):
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=idx, column=col_idx, value=value)
            if idx == 11:
                cell.font = Font(bold=True, color="FFFFFF", size=10)
                cell.fill = PatternFill("solid", fgColor="366092")
            else:
                cell.border = Border(left=Side(style="thin", color="D9D9D9"),
                                    right=Side(style="thin", color="D9D9D9"),
                                    top=Side(style="thin", color="D9D9D9"),
                                    bottom=Side(style="thin", color="D9D9D9"))
            cell.alignment = Alignment(horizontal="center")

    # Format percentages
    ws["D13"].number_format = "0.0%"
    ws["D14"].number_format = "0.0%"
    ws["D15"].number_format = "0.0%"

    # Fat Loss Guidelines
    ws["A18"] = "💡 FAT LOSS NUTRITION GUIDELINES"
    ws["A18"].font = Font(size=11, bold=True, color="FFFFFF")
    ws["A18"].fill = PatternFill("solid", fgColor="70AD47")
    ws.merge_cells("A18:J18")

    guidelines = [
        "✓ Protein: 1.6-2.2g/kg bodyweight (preserves muscle during deficit)",
        "✓ Caloric Deficit: 300-500 kcal below maintenance for steady 0.5-1 kg/week loss",
        "✓ Best Carbs: Oats, rice, sweet potatoes, whole wheat, fruits, beans",
        "✓ Best Proteins: Chicken, turkey, fish, eggs, Greek yogurt, cottage cheese",
        "✓ Healthy Fats: Olive oil, avocado, nuts, salmon, chia seeds",
        "✓ Vegetables: Unlimited - low calorie, high satiety (broccoli, spinach, peppers)",
        "✓ Hydration: 3-4 liters daily + more on training days",
        "✓ Meal Prep: Plan meals in advance for consistency & portion control",
    ]
    for idx, guideline in enumerate(guidelines, 19):
        ws[f"A{idx}"] = guideline
        ws[f"A{idx}"].font = Font(size=9, color="1F4E78")
        ws.merge_cells(f"A{idx}:J{idx}")

    set_col_widths(ws, {"A": 16, "B": 12, "C": 22, "D": 16, "E": 12, "F": 12, "G": 12, "H": 12, "I": 10, "J": 20})


def build_progress_premium(ws):
    ws.title = "Progress Tracker"
    ws.freeze_panes = "A2"

    ws.merge_cells("A1:J1")
    ws["A1"] = "📊 COMPREHENSIVE PROGRESS TRACKING"
    ws["A1"].font = Font(size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="4472C4")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 25

    headers = [
        "Date",
        "Weight (kg)",
        "Waist (cm)",
        "Chest (cm)",
        "Hips (cm)",
        "Arm (cm)",
        "Thigh (cm)",
        "Bench (kg)",
        "Squat (kg)",
        "Notes",
    ]
    ws.append(headers)
    style_header(ws, row=2, color="4472C4")

    sample_rows = [
        ["2026-09-01", "80.5", "86", "101", "96", "32", "58", "85", "110", "Baseline - Starting"],
        ["2026-09-15", "79.8", "85", "100", "95", "32", "58", "90", "115", "Good progress - strength up"],
        ["2026-10-05", "78.7", "84", "99", "94", "31.5", "57.5", "95", "120", "Excellent - shred working"],
    ]
    for row in sample_rows:
        ws.append(row)

    add_border(ws, 3, ws.max_row, 1, 10)

    # Progress Summary
    ws["A6"] = "PROGRESS SUMMARY"
    ws["A6"].font = Font(size=12, bold=True, color="FFFFFF")
    ws["A6"].fill = PatternFill("solid", fgColor="4472C4")
    ws.merge_cells("A6:J6")

    summary_data = [
        ["Metric", "Start", "Current", "Change", "Goal", "Progress %"],
        ["Bodyweight (kg)", "80.5", "=B5", "=B8-B7", "76", "=(B7-76)/(B7-76)*100"],
        ["Waist (cm)", "86", "=C5", "=C8-C7", "80", "=(C8-80)/(C7-80)*100"],
        ["Bench (kg)", "85", "=H5", "=H8-H7", "105", "=(H8-105)/(H7-105)*100"],
        ["Squat (kg)", "110", "=I5", "=I8-I7", "140", "=(I8-140)/(I7-140)*100"],
    ]

    for idx, row_data in enumerate(summary_data, 7):
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=idx, column=col_idx, value=value)
            if idx == 7:
                cell.font = Font(bold=True, color="FFFFFF", size=10)
                cell.fill = PatternFill("solid", fgColor="366092")
            else:
                cell.border = Border(left=Side(style="thin", color="D9D9D9"),
                                    right=Side(style="thin", color="D9D9D9"),
                                    top=Side(style="thin", color="D9D9D9"),
                                    bottom=Side(style="thin", color="D9D9D9"))

    set_col_widths(ws, {"A": 14, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12, "H": 12, "I": 12, "J": 20})


def build_program_guide(ws):
    ws.title = "Program Guide"
    
    ws.merge_cells("A1:E1")
    ws["A1"] = "STRENGTH & SHRED - 12 WEEK PROGRAM"
    ws["A1"].font = Font(size=15, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="C00000")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 30

    sections = [
        ("MONDAY - PUSH DAY", "4472C4", [
            ["Exercise", "Sets x Reps", "Rest", "Notes"],
            ["Bench Press", "4 x 5-6", "180s", "Main strength builder"],
            ["Incline DB Press", "3 x 8-10", "120s", "Upper chest hypertrophy"],
            ["Shoulder Press", "3 x 6-8", "120s", "Shoulder strength"],
            ["Tricep Dips", "3 x 6-8", "120s", "Tricep strength"],
            ["Lateral Raises", "3 x 12-15", "60s", "Shoulder isolation"],
        ]),
        ("TUESDAY - PULL DAY", "70AD47", [
            ["Exercise", "Sets x Reps", "Rest", "Notes"],
            ["Deadlift", "4 x 3-5", "240s", "Main compound lift"],
            ["Barbell Row", "4 x 6-8", "180s", "Back strength"],
            ["Pull-ups", "3 x 6-10", "120s", "Lat development"],
            ["Barbell Curl", "3 x 8-10", "90s", "Bicep hypertrophy"],
            ["Face Pulls", "3 x 12-15", "60s", "Rear delt & shoulder health"],
        ]),
        ("THURSDAY - LEGS DAY", "FF6B35", [
            ["Exercise", "Sets x Reps", "Rest", "Notes"],
            ["Back Squat", "4 x 5-6", "180s", "Leg strength"],
            ["RDL", "3 x 8-10", "120s", "Hamstring & lower back"],
            ["Leg Press", "3 x 8-10", "120s", "Leg hypertrophy"],
            ["Leg Curl", "3 x 10-12", "90s", "Hamstring isolation"],
            ["Calf Raises", "3 x 12-15", "60s", "Calf development"],
        ]),
    ]

    current_row = 3
    for section_title, color, exercises in sections:
        ws.merge_cells(f"A{current_row}:E{current_row}")
        title_cell = ws[f"A{current_row}"]
        title_cell.value = section_title
        title_cell.font = Font(size=12, bold=True, color="FFFFFF")
        title_cell.fill = PatternFill("solid", fgColor=color)
        title_cell.alignment = Alignment(horizontal="center")
        current_row += 1

        for ex_idx, exercise_row in enumerate(exercises):
            for col_idx, value in enumerate(exercise_row, 1):
                cell = ws.cell(row=current_row, column=col_idx, value=value)
                if ex_idx == 0:  # Header row
                    cell.font = Font(bold=True, color="FFFFFF", size=10)
                    cell.fill = PatternFill("solid", fgColor=color)
                else:
                    cell.border = Border(left=Side(style="thin", color="D9D9D9"),
                                        right=Side(style="thin", color="D9D9D9"),
                                        top=Side(style="thin", color="D9D9D9"),
                                        bottom=Side(style="thin", color="D9D9D9"))
            current_row += 1

        current_row += 1

    # Saturday Optional
    ws.merge_cells(f"A{current_row}:E{current_row}")
    title_cell = ws[f"A{current_row}"]
    title_cell.value = "SATURDAY - OPTIONAL CONDITIONING"
    title_cell.font = Font(size=12, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill("solid", fgColor="7030A0")
    title_cell.alignment = Alignment(horizontal="center")
    current_row += 1

    conditioning = [
        ["Type", "Duration", "Intensity", "Benefit"],
        ["Steady State Cardio", "20-30 min", "60-70% max HR", "Fat loss + recovery"],
        ["HIIT Training", "15-20 min", "80-90% max HR", "Max fat burn efficiently"],
        ["Active Recovery", "30-45 min", "Low intensity", "Blood flow & mobility"],
    ]

    for cond_idx, cond_row in enumerate(conditioning):
        for col_idx, value in enumerate(cond_row, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=value)
            if cond_idx == 0:
                cell.font = Font(bold=True, color="FFFFFF", size=10)
                cell.fill = PatternFill("solid", fgColor="7030A0")
            else:
                cell.border = Border(left=Side(style="thin", color="D9D9D9"),
                                    right=Side(style="thin", color="D9D9D9"),
                                    top=Side(style="thin", color="D9D9D9"),
                                    bottom=Side(style="thin", color="D9D9D9"))
        current_row += 1

    set_col_widths(ws, {"A": 22, "B": 18, "C": 12, "D": 28, "E": 15})


def build_settings_premium(ws):
    ws.title = "Settings"
    
    ws.merge_cells("A1:D1")
    ws["A1"] = "⚙️ PROGRAM SETTINGS & PERSONAL INFO"
    ws["A1"].font = Font(size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="1F4E78")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 25

    settings = [
        ["Setting", "Value", "Description"],
        ["Program Goal", "Strength + Shred", "Build muscle & lose fat simultaneously"],
        ["Training Frequency", "3 core + 1 optional", "Mon/Tue/Thu mandatory, Sat optional"],
        ["Program Duration", "12 weeks", "Complete cycle: 3 phases of 4 weeks"],
        ["Current Bodyweight", "78.7 kg", "As of last measurement"],
        ["Target Bodyweight", "76 kg", "Goal weight (fat loss)"],
        ["Target Waist", "80 cm", "Goal circumference"],
        ["Daily Protein", "170 g", "1.6-2.2g per kg bodyweight"],
        ["Daily Calories", "2100-2150", "300-500 kcal deficit"],
        ["Maintenance Calories", "2600", "Estimated baseline"],
        ["Sleep Goal", "7-9 hours", "Essential for recovery & hormones"],
        ["Water Intake", "3-4 liters", "Daily hydration target"],
    ]

    for idx, row_data in enumerate(settings, 2):
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=idx, column=col_idx, value=value)
            if idx == 2:
                cell.font = Font(bold=True, color="FFFFFF", size=10)
                cell.fill = PatternFill("solid", fgColor="366092")
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.border = Border(left=Side(style="thin", color="D9D9D9"),
                                    right=Side(style="thin", color="D9D9D9"),
                                    top=Side(style="thin", color="D9D9D9"),
                                    bottom=Side(style="thin", color="D9D9D9"))

    current_row = len(settings) + 3

    ws[f"A{current_row}"] = "KEY SUCCESS FACTORS"
    ws[f"A{current_row}"].font = Font(size=12, bold=True, color="FFFFFF")
    ws[f"A{current_row}"].fill = PatternFill("solid", fgColor="70AD47")
    ws.merge_cells(f"A{current_row}:D{current_row}")
    current_row += 1

    factors = [
        "1️⃣ Consistency: Never miss scheduled sessions - your body adapts to routine",
        "2️⃣ Progressive Overload: Add weight or reps each week - this drives adaptation",
        "3️⃣ Nutrition Discipline: Log everything - even small extras add calories",
        "4️⃣ Recovery: Sleep 7-9 hours - this is when muscle is built",
        "5️⃣ Hydration: Drink water constantly - impacts performance & recovery",
        "6️⃣ Tracking: Record all workouts & measurements - data drives decisions",
    ]

    for factor in factors:
        ws[f"A{current_row}"] = factor
        ws[f"A{current_row}"].font = Font(size=10, color="1F4E78")
        ws.merge_cells(f"A{current_row}:D{current_row}")
        current_row += 1

    set_col_widths(ws, {"A": 25, "B": 18, "C": 35, "D": 20})


def main():
    wb = Workbook()
    
    # Build all premium sheets
    build_premium_dashboard(wb.active)
    build_weekly_tracker_premium(wb.create_sheet())
    build_workout_log_premium(wb.create_sheet())
    build_nutrition_premium(wb.create_sheet())
    build_progress_premium(wb.create_sheet())
    build_program_guide(wb.create_sheet())
    build_settings_premium(wb.create_sheet())

    filename = "Fitness_GYM_Planner_PREMIUM.xlsx"
    wb.save(filename)
    print(f"✅ Premium Workbook Created: {filename}")
    print(f"📊 Program: 3 core sessions + 1 optional conditioning day")
    print(f"💪 Goals: Strength training + Bodybuilding + Fat Loss Shredding")
    print(f"🥗 Includes: Nutrition tracker with macro calculations & fat-loss guidance")
    print(f"\n🎯 Ready to use - Open in Excel and start tracking!")


if __name__ == "__main__":
    main()
