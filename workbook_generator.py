from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import BarChart, LineChart, Reference
from datetime import datetime


def style_header(ws, row=1, start_col=1, end_col=None):
    end_col = ws.max_column if end_col is None else end_col
    for cell in ws.iter_cols(min_row=row, max_row=row, min_col=start_col, max_col=end_col):
        for c in cell:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F4E78")
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )


def set_col_widths(ws, widths):
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def build_dashboard(ws):
    ws.title = "Dashboard"
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False

    ws["A1"] = "STRENGTH & SHRED FITNESS PLANNER"
    ws["A1"].font = Font(size=18, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="C00000")

    ws["A3"] = "Weekly Overview"
    ws["A3"].font = Font(size=12, bold=True)

    cards = [
        ("A5", "Sessions This Week", "=COUNTIF('Weekly Tracker'!B:B,TRUE)"),
        ("D5", "Avg Strength Load", "=ROUND(AVERAGE('Workout Log'!H:H),0)"),
        ("A8", "Current Weight", "=IFERROR(INDEX('Progress Tracker'!B:B,MATCH(MAX('Progress Tracker'!A:A),'Progress Tracker'!A:A,0)), \"-\")"),
        ("D8", "Current Goal", "=IFERROR(INDEX('Goals'!B:B,1),\"Strength + Fat Loss\")"),
    ]

    for cell, label, formula in cards:
        ws[cell] = label
        ws[cell].font = Font(bold=True, color="FFFFFF")
        ws[cell].fill = PatternFill("solid", fgColor="FF6B35")
        ws[cell].alignment = Alignment(horizontal="center")
        ws[cell].border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )

        target = ws.cell(row=int(cell[1:]) + 1, column=ord(cell[0]) - 64)
        target.value = formula
        target.font = Font(bold=True, size=12, color="C00000")
        target.fill = PatternFill("solid", fgColor="FFF2CC")

    ws["A12"] = "Nutrition Rules"
    ws["A12"].font = Font(size=12, bold=True)

    rules = [
        "Protein: 1.6-2.2g per kg bodyweight",
        "Calories: 300-500 kcal below maintenance",
        "Water: 3-4 liters daily",
        "Training: 3 sessions minimum + optional 4th",
        "Sleep: 7-9 hours nightly",
        "Consistency: Track workouts + food daily",
    ]
    for idx, rule in enumerate(rules, 14):
        ws[f"A{idx}"] = rule
        ws[f"A{idx}"].font = Font(size=10)

    ws["D12"] = "Fat Loss Recommendations"
    ws["D12"].font = Font(size=12, bold=True)
    recs = [
        "Eat lean protein at each meal",
        "Use rice, potatoes, oats for carbs",
        "Add vegetables to every meal",
        "Keep liquid calories low",
        "Track all meals and snacks",
        "Prioritize recovery and sleep",
    ]
    for idx, rec in enumerate(recs, 14):
        ws[f"D{idx}"] = rec
        ws[f"D{idx}"].font = Font(size=10)

    ws["A30"] = "Program Focus"
    ws["A30"].font = Font(size=12, bold=True)
    ws["A31"] = "- Strength training for muscle retention"
    ws["A32"] = "- Fat-loss phase with calorie control"
    ws["A33"] = "- High protein and recovery"
    ws["A34"] = "- Maximum consistency, minimum burnout"

    set_col_widths(ws, {"A": 30, "B": 18, "C": 18, "D": 30, "E": 18, "F": 18, "G": 18, "H": 18})


def build_weekly_tracker(ws):
    ws.title = "Weekly Tracker"
    ws.freeze_panes = "A2"
    headers = [
        "Week Starting",
        "Mon Session",
        "Mon Workout",
        "Tue Session",
        "Tue Workout",
        "Thu Session",
        "Thu Workout",
        "Sat Optional",
        "Sat Workout",
        "Notes",
    ]
    ws.append(headers)
    style_header(ws)

    sample_rows = [
        ["2026-10-05", "YES", "Push", "YES", "Pull", "YES", "Legs", "NO", "", "Progressive overload"],
        ["2026-10-12", "YES", "Push", "YES", "Pull", "YES", "Legs", "YES", "Conditioning", "Strong week"],
    ]
    for row in sample_rows:
        ws.append(row)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=10):
        for cell in row:
            cell.border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )

    set_col_widths(ws, {"A": 14, "B": 12, "C": 24, "D": 12, "E": 24, "F": 12, "G": 24, "H": 15, "I": 24, "J": 28})


def build_workout_log(ws):
    ws.title = "Workout Log"
    ws.freeze_panes = "A2"
    headers = [
        "Date",
        "Week",
        "Session",
        "Exercise",
        "Muscle Group",
        "Sets",
        "Reps",
        "Weight (kg)",
        "Rest (sec)",
        "RPE",
        "Notes",
    ]
    ws.append(headers)
    style_header(ws)

    sample_rows = [
        ["2026-10-05", "1", "Push", "Bench Press", "Chest", "4", "5-6", "90", "180", "8", "Strong"],
        ["2026-10-05", "1", "Push", "Incline DB Press", "Upper Chest", "3", "8-10", "35", "120", "7", "Good form"],
        ["2026-10-06", "1", "Pull", "Deadlift", "Back", "4", "3-5", "140", "240", "9", "Heavy"],
        ["2026-10-06", "1", "Pull", "Barbell Row", "Back", "4", "6-8", "110", "180", "8", "Controlled"],
        ["2026-10-07", "1", "Legs", "Back Squat", "Legs", "4", "5-6", "120", "180", "8", "Hard but solid"],
    ]
    for row in sample_rows:
        ws.append(row)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=11):
        for cell in row:
            cell.border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )

    set_col_widths(ws, {"A": 12, "B": 8, "C": 12, "D": 24, "E": 16, "F": 8, "G": 10, "H": 12, "I": 12, "J": 10, "K": 24})


def build_nutrition_tracker(ws):
    ws.title = "Nutrition Tracker"
    ws.freeze_panes = "A2"
    ws["A1"] = "FOOD HABITS & FAT LOSS TRACKER"
    ws["A1"].font = Font(size=15, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="70AD47")

    headers = [
        "Date",
        "Meal",
        "Foods / Meal",
        "Quantity",
        "Protein (g)",
        "Carbs (g)",
        "Fat (g)",
        "Calories",
        "Hydration (L)",
        "Notes",
    ]
    ws.append(headers)
    style_header(ws)

    sample_rows = [
        ["2026-10-05", "Breakfast", "Eggs + Oatmeal", "3 eggs + 50g oats", "18", "25", "8", "240", "0.5", "Good start"],
        ["2026-10-05", "Lunch", "Chicken + Rice + veg", "200g chicken + 75g rice", "45", "45", "6", "410", "0.7", "Post-workout"],
        ["2026-10-05", "Snack", "Greek yogurt + berries", "150g + 100g", "20", "12", "2", "150", "0.3", "Protein hit"],
        ["2026-10-05", "Dinner", "Salmon + broccoli", "180g salmon + 200g veg", "35", "15", "12", "360", "0.7", "High quality fats"],
        ["2026-10-05", "Daily Total", "", "", "=SUM(E3:E6)", "=SUM(F3:F6)", "=SUM(G3:G6)", "=SUM(H3:H6)", "=SUM(I3:I6)", ""],
    ]
    for row in sample_rows:
        ws.append(row)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=10):
        for cell in row:
            cell.border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )

    ws["A15"] = "Fat Loss Nutrition Recommendations"
    ws["A15"].font = Font(size=12, bold=True, color="FFFFFF")
    ws["A15"].fill = PatternFill("solid", fgColor="70AD47")

    recs = [
        "Protein: 1.6-2.2g/kg bodyweight every day",
        "Calorie deficit: 300-500 kcal under maintenance",
        "Best carbs: oats, rice, potatoes, fruit, beans",
        "Best fats: avocado, nuts, olive oil, salmon",
        "Vegetables: 2-3 servings every meal",
        "Water: 3-4 liters daily, more on training days",
        "Avoid liquid calories and mindless snacking",
        "Plan meals ahead to improve consistency",
        "Prioritize whole foods over processed foods",
        "Focus on strength training to preserve muscle during fat loss",
    ]
    for idx, rec in enumerate(recs, 16):
        ws[f"A{idx}"] = rec
        ws[f"A{idx}"].font = Font(size=10)

    set_col_widths(ws, {"A": 18, "B": 12, "C": 24, "D": 18, "E": 12, "F": 12, "G": 12, "H": 12, "I": 12, "J": 24})


def build_progress_tracker(ws):
    ws.title = "Progress Tracker"
    ws.freeze_panes = "A2"
    headers = [
        "Date",
        "Bodyweight (kg)",
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
    style_header(ws)

    sample_rows = [
        ["2026-09-01", "80.5", "86", "101", "96", "32", "58", "85", "110", "Baseline"],
        ["2026-09-15", "79.8", "85", "100", "95", "32", "58", "90", "115", "Strong progress"],
        ["2026-10-05", "78.7", "84", "99", "94", "31.5", "57.5", "95", "120", "Strength + shred on track"],
    ]
    for row in sample_rows:
        ws.append(row)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=10):
        for cell in row:
            cell.border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )

    set_col_widths(ws, {"A": 12, "B": 14, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12, "H": 12, "I": 12, "J": 26})


def build_goals(ws):
    ws.title = "Goals"
    ws.freeze_panes = "A2"
    headers = ["Goal", "Target", "Current", "Deadline", "Status", "Notes"]
    ws.append(headers)
    style_header(ws)

    sample_rows = [
        ["Increase bench press", "105 kg", "95 kg", "12 weeks", "In progress", "Add 5 kg every 2 weeks"],
        ["Drop waist size", "80 cm", "84 cm", "12 weeks", "In progress", "Maintain deficit and train consistently"],
        ["Strength and lean muscle retention", "Keep muscle while losing fat", "On track", "12 weeks", "In progress", "Priority: protein + progressive overload"],
    ]
    for row in sample_rows:
        ws.append(row)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=6):
        for cell in row:
            cell.border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )

    set_col_widths(ws, {"A": 26, "B": 16, "C": 16, "D": 16, "E": 16, "F": 28})


def build_program_structure(ws):
    ws.title = "Program Structure"
    ws["A1"] = "3-DAY CORE PROGRAM + OPTIONAL 4TH DAY"
    ws["A1"].font = Font(size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="C00000")

    sessions = [
        ["Day", "Focus", "Main Exercises", "Optional Add-ons", "Goal"],
        ["Monday", "Push", "Bench Press, Incline DB Press, Shoulder Press, Triceps",
         "Cable fly, lateral raises", "Build chest/shoulders & upper-body strength"],
        ["Tuesday", "Pull", "Deadlift, Barbell Rows, Pull-up/Lat Pulldown, Bicep Curl",
         "Face pulls, rear delt work", "Back development & pulling power"],
        ["Thursday", "Legs", "Back Squat, RDL, Leg Press, Leg Curl",
         "Calf raises, leg extension", "Lower body strength & muscular legs"],
        ["Saturday (Optional)", "Conditioning / Recovery",
         "Treadmill walk, bike, intervals, mobility work",
         "30-45 min low to moderate cardio", "Accelerate fat loss with less fatigue"],
    ]

    for row in sessions:
        ws.append(row)

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=5):
        for cell in row:
            cell.border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    style_header(ws, row=1)

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 32
    ws.column_dimensions["D"].width = 28
    ws.column_dimensions["E"].width = 30


def build_settings(ws):
    ws.title = "Settings"
    ws["A1"] = "PERSONAL SETTINGS"
    ws["A1"].font = Font(size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="1F4E78")

    settings = [
        ["Setting", "Value"],
        ["Program Goal", "Strength + Shred / Fat Loss"],
        ["Session Frequency", "3 core sessions + 1 optional"],
        ["Target Weight", "75-78 kg"],
        ["Target Waist", "80-82 cm"],
        ["Daily Protein", "160-175 g"],
        ["Daily Calories", "2100-2150 kcal"],
        ["Water Intake", "3-4 L"],
        ["Sleep Goal", "7-9 hours"],
        ["Training Days", "Mon / Tue / Thu + Optional Sat"],
    ]
    for row in settings:
        ws.append(row)

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=2):
        for cell in row:
            cell.border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )
    style_header(ws, row=1, start_col=1, end_col=2)
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 28


def main():
    wb = Workbook()
    build_dashboard(wb.active)
    build_weekly_tracker(wb.create_sheet())
    build_workout_log(wb.create_sheet())
    build_nutrition_tracker(wb.create_sheet())
    build_progress_tracker(wb.create_sheet())
    build_goals(wb.create_sheet())
    build_program_structure(wb.create_sheet())
    build_settings(wb.create_sheet())

    filename = "Fitness_GYM_Planner.xlsx"
    wb.save(filename)
    print(f"Workbook created: {filename}")
    print("Program: 3 core sessions + 1 optional conditioning day")
    print("Goal: Strength + body building + fat loss shredding")
    print("Included: Nutrition tracker with fat-loss food habits recommendations")


if __name__ == "__main__":
    main()
