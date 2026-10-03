from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
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

    ws["A1"] = "FITNESS GYM PLANNER"
    ws["A1"].font = Font(size=18, bold=True, color="1F1F1F")

    ws["A3"] = "Overview"
    ws["A3"].font = Font(size=12, bold=True)

    cards = [
        ("A5", "Total Workouts", "=COUNTA('Workout Log'!A2:A1000)"),
        ("D5", "Avg Weekly Volume", "=ROUND(AVERAGE('Workout Log'!G2:G1000),0)"),
        ("A8", "Current Goal", "=LOOKUP(2,1/('Goals'!A2:A1000<>""), 'Goals'!B2:B1000)"),
        ("D8", "Bodyweight Trend", "=ROUND(AVERAGE('Measurements'!E2:E1000),1)"),
    ]

    for cell, label, formula in cards:
        ws[cell] = label
        ws[cell].font = Font(bold=True, color="FFFFFF")
        ws[cell].fill = PatternFill("solid", fgColor="4472C4")
        ws[cell].alignment = Alignment(horizontal="center")
        ws[cell].border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )

        target = ws.cell(row=int(cell[1:]) + 1, column=ord(cell[0]) - 64)
        target.value = formula
        target.font = Font(bold=True, size=12)
        target.fill = PatternFill("solid", fgColor="D9EAF7")

    ws["A12"] = "Weekly Progress Summary"
    ws["A12"].font = Font(size=12, bold=True)

    chart = BarChart()
    chart.type = "bar"
    chart.style = 10
    chart.title = "Workout Frequency"
    chart.y_axis.title = "Sessions"
    chart.x_axis.title = "Week"

    data = Reference(ws, min_col=2, min_row=14, max_col=2, max_row=18)
    cats = Reference(ws, min_col=1, min_row=15, max_row=18)
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(cats)
    chart.height = 7
    chart.width = 12
    ws.add_chart(chart, "A14")

    ws["H12"] = "Performance Trend"
    ws["H12"].font = Font(size=12, bold=True)
    line_chart = LineChart()
    line_chart.title = "Strength Trend"
    line_chart.style = 13
    line_chart.y_axis.title = "Load (kg)"
    line_chart.x_axis.title = "Date"
    data2 = Reference(ws, min_col=8, min_row=13, max_col=8, max_row=18)
    cats2 = Reference(ws, min_col=7, min_row=14, max_row=18)
    line_chart.add_data(data2, titles_from_data=False)
    line_chart.set_categories(cats2)
    line_chart.height = 7
    line_chart.width = 12
    ws.add_chart(line_chart, "H14")

    ws["A20"] = "Planner Notes"
    ws["A20"].font = Font(size=12, bold=True)
    ws["A21"] = "- Track consistency & recovery"
    ws["A22"] = "- Prioritize progressive overload"
    ws["A23"] = "- Review body measurements every 4 weeks"

    set_col_widths(ws, {"A": 18, "B": 18, "C": 18, "D": 18, "E": 18, "F": 18, "G": 18, "H": 18, "I": 18, "J": 18})


def build_weekly_plan(ws):
    ws.title = "Weekly Plan"
    ws.freeze_panes = "A2"
    headers = [
        "Week",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
        "Sunday",
        "Priority",
        "Notes",
    ]
    ws.append(headers)
    style_header(ws)

    for i in range(1, 9):
        row = [f"Wk {i}", "Push", "Pull", "Legs", "Upper", "Conditioning", "Recovery", "Rest", "Main Goal", "Focus on form and progression"]
        ws.append(row)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=10):
        for cell in row:
            cell.border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )

    set_col_widths(ws, {"A": 12, "B": 18, "C": 18, "D": 18, "E": 18, "F": 18, "G": 18, "H": 18, "I": 18, "J": 28})


def build_workout_log(ws):
    ws.title = "Workout Log"
    ws.freeze_panes = "A2"
    headers = [
        "Date",
        "Week",
        "Workout Type",
        "Exercise",
        "Muscle Group",
        "Sets",
        "Reps",
        "Weight (kg)",
        "RPE",
        "Notes",
    ]
    ws.append(headers)
    style_header(ws)

    sample_rows = [
        ["2026-10-01", "Wk 1", "Strength", "Bench Press", "Chest", "5", "5", "80", "8", "Controlled tempo"],
        ["2026-10-03", "Wk 1", "Strength", "Back Squat", "Legs", "5", "5", "100", "8", "Strong brace"],
        ["2026-10-05", "Wk 1", "Cardio", "Treadmill Run", "Conditioning", "1", "20 min", "0", "7", "Steady pace"],
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

    set_col_widths(ws, {"A": 14, "B": 10, "C": 18, "D": 22, "E": 18, "F": 10, "G": 10, "H": 14, "I": 10, "J": 28})


def build_exercise_library(ws):
    ws.title = "Exercise Library"
    ws.freeze_panes = "A2"
    headers = [
        "Exercise",
        "Category",
        "Primary Muscle",
        "Secondary Muscle",
        "Equipment",
        "Instructions",
        "Target Reps",
        "Target Sets",
    ]
    ws.append(headers)
    style_header(ws)

    sample_rows = [
        ["Bench Press", "Strength", "Chest", "Shoulders, Triceps", "Barbell", "Lower to chest and press upward", "5-8", "4-5"],
        ["Deadlift", "Strength", "Back", "Hamstrings, Glutes", "Barbell", "Brace core and drive from floor", "3-6", "3-5"],
        ["Squat", "Strength", "Legs", "Glutes, Core", "Barbell", "Depth controlled with neutral spine", "5-8", "4-5"],
        ["Pull-Up", "Bodyweight", "Back", "Biceps", "Pull-up Bar", "Use controlled range of motion", "6-10", "3-4"],
    ]
    for row in sample_rows:
        ws.append(row)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=8):
        for cell in row:
            cell.border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    set_col_widths(ws, {"A": 20, "B": 18, "C": 18, "D": 18, "E": 18, "F": 30, "G": 12, "H": 12})


def build_progress(ws):
    ws.title = "Progress"
    ws.freeze_panes = "A2"
    headers = ["Date", "Bodyweight (kg)", "Bench (kg)", "Squat (kg)", "Deadlift (kg)", "Conditioning Time", "Notes"]
    ws.append(headers)
    style_header(ws)

    sample_rows = [
        ["2026-09-01", "76.5", "80", "100", "120", "20 min", "Base start"],
        ["2026-09-15", "77.0", "85", "105", "125", "22 min", "Strength improving"],
        ["2026-10-01", "77.5", "90", "110", "130", "25 min", "Target progress"],
    ]
    for row in sample_rows:
        ws.append(row)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=7):
        for cell in row:
            cell.border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )

    set_col_widths(ws, {"A": 14, "B": 16, "C": 14, "D": 14, "E": 14, "F": 18, "G": 30})


def build_measurements(ws):
    ws.title = "Measurements"
    ws.freeze_panes = "A2"
    headers = [
        "Date",
        "Weight (kg)",
        "Chest (cm)",
        "Waist (cm)",
        "Hips (cm)",
        "Arm (cm)",
        "Thigh (cm)",
        "Notes",
    ]
    ws.append(headers)
    style_header(ws)

    sample_rows = [
        ["2026-09-01", "76.5", "96", "82", "95", "31", "55", "Starting point"],
        ["2026-10-01", "77.5", "97", "80", "94", "32", "56", "Improved waist"],
    ]
    for row in sample_rows:
        ws.append(row)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=8):
        for cell in row:
            cell.border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )

    set_col_widths(ws, {"A": 14, "B": 14, "C": 14, "D": 14, "E": 14, "F": 12, "G": 12, "H": 28})


def build_goals(ws):
    ws.title = "Goals"
    ws.freeze_panes = "A2"
    headers = ["Goal ID", "Goal Description", "Target Date", "Category", "Target Value", "Current Status", "Progress %", "Notes"]
    ws.append(headers)
    style_header(ws)

    sample_rows = [
        ["G1", "Increase bench press strength", "2026-12-31", "Strength", "110 kg", "In progress", "60%", "Add 5 kg every 2 weeks"],
        ["G2", "Reduce waist circumference", "2026-12-31", "Body Composition", "78 cm", "In progress", "70%", "Maintain calories and cardio"],
        ["G3", "Complete 3 weekly sessions", "2026-11-30", "Consistency", "12 sessions", "In progress", "80%", "Stay on schedule"],
    ]
    for row in sample_rows:
        ws.append(row)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=8):
        for cell in row:
            cell.border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )

    set_col_widths(ws, {"A": 12, "B": 28, "C": 14, "D": 16, "E": 16, "F": 18, "G": 12, "H": 28})


def build_notes(ws):
    ws.title = "Notes"
    ws["A1"] = "Gym Coach / Follow-Up Notes"
    ws["A1"].font = Font(size=16, bold=True)
    ws["A3"] = "Session Feedback"
    ws["A3"].font = Font(size=12, bold=True)
    ws["A5"] = "- Focus on technique and progression"
    ws["A6"] = "- Monitor recovery and soreness"
    ws["A7"] = "- Add conditioning twice weekly"
    ws["A9"] = "Coach Recommendation"
    ws["A9"].font = Font(size=12, bold=True)
    ws["A11"] = "- Maintain 7+ hours sleep"
    ws["A12"] = "- Hydrate consistently"
    ws["A13"] = "- Keep weekly logs updated"

    for cell in ["A1", "A3", "A9"]:
        ws[cell].fill = PatternFill("solid", fgColor="D9EAF7")

    ws.column_dimensions["A"].width = 60


def build_settings(ws):
    ws.title = "Settings"
    ws["A1"] = "Gym Planner Settings"
    ws["A1"].font = Font(size=16, bold=True)
    settings = [
        ["Setting", "Value"],
        ["Trainer Name", "Your Name"],
        ["Gym Name", "Your Gym"],
        ["Program Start Date", datetime.today().strftime("%Y-%m-%d")],
        ["Goal Type", "Muscle Gain / Fat Loss / Strength"],
        ["Preferred Units", "kg / cm"],
        ["Session Frequency", "3-5 per week"],
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
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 30


def main():
    wb = Workbook()
    build_dashboard(wb.active)
    build_weekly_plan(wb.create_sheet())
    build_workout_log(wb.create_sheet())
    build_exercise_library(wb.create_sheet())
    build_progress(wb.create_sheet())
    build_measurements(wb.create_sheet())
    build_goals(wb.create_sheet())
    build_notes(wb.create_sheet())
    build_settings(wb.create_sheet())

    filename = "Fitness_GYM_Planner.xlsx"
    wb.save(filename)
    print(f"Workbook created: {filename}")


if __name__ == "__main__":
    main()
