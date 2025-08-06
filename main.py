import pandas as pd
import xlrd
import requests
import re
import json

sheet_id = "1qKt-qplbttz9mG79q_v7w6FS3ae50o9_"

sheets_data = {}

# https://docs.google.com/spreadsheets/d/1qKt-qplbttz9mG79q_v7w6FS3ae50o9_/export?format=xls&gid=1003107230

url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xls"
try:
    r = requests.get(url, allow_redirects=True)
    open("sched.xls", 'wb').write(r.content)
except Exception as e:
    print(f"Failed to download: {e}")

book = xlrd.open_workbook("sched.xls")

print(f"Sheets found: {book.sheet_names()}")

sh = book.sheet_by_index(1)

ncols = sh.ncols
nrows = sh.nrows

raw_group_names: list[str] = [str(cell.value).strip() for cell in sh.row(0) if cell.value]

seen = set()
group_names: list[str] = []
for g in raw_group_names:
    if g and g not in seen:
        seen.add(g)
        group_names.append(g)

data: dict[str, list] = {
    "groupNames": group_names,
    "groups": []
}

for gi, i in enumerate(range(2, ncols, 6)):
    num_idx = i - 1
    a_idx = i
    b_idx = i + 3

    if b_idx >= ncols:
        break

    if gi < len(group_names):
        group_name = group_names[gi]
    else:
        group_name = sh.cell(0, a_idx).value or f"group_{gi}"

    num_col = sh.col(num_idx)

    slots_flat: list[dict] = []

    num_start = 3  # row index where lesson numbering starts

    # Walk down the column in steps of 6 rows (each slot block)
    for j in range(num_start, len(num_col), 6):
        # Ensure row exists
        if j >= nrows:
            break

        num_cell = num_col[j]
        if num_cell.value == "":
            continue

        # Determine lessons for this slot by inspecting the six-row block
        lesson_top_idx = j - 1  # first cell of the block
        lesson_fourth_idx = j + 2  # 4th cell (0-based offset 3)

        def val(row_i: int, col_i: int) -> str:
            if 0 <= row_i < nrows and 0 <= col_i < ncols:
                return str(sh.cell(row_i, col_i).value).strip()
            return ""
        
        def assemble_lesson(row_i: int, subj_col: int):
            subj = val(row_i, subj_col)
            if subj == "":
                return None

            # Two possible vertical layouts:
            # 1) Subject row height == 1 (alternating weeks):
            #    teacher -> +1, room -> +2
            # 2) Subject row height == 3 (full-block merge):
            #    the merged subject spans +0,+1,+2 rows; real teacher starts at +2 / +3 below top

            teacher = val(row_i + 1, subj_col)
            room = val(row_i + 2, subj_col)

            if teacher == "" or teacher == subj or room == subj:
                teacher = val(row_i + 2, subj_col)
                room = val(row_i + 4, subj_col)

            return {"subject": subj, "teacher": teacher, "room": room}

        v_label = val(lesson_top_idx, a_idx)
        is_split = v_label.upper() in ("A", "А")

        if is_split:
            lessons_A: list[dict] = []
            lessons_B: list[dict] = []

            for row_idx in (lesson_top_idx, lesson_fourth_idx):
                lesson_A = assemble_lesson(row_idx, a_idx + 1)
                lesson_B = assemble_lesson(row_idx, a_idx + 3)

                if lesson_A:
                    lessons_A.append(lesson_A)
                if lesson_B:
                    lessons_B.append(lesson_B)

            slot = {
                "number": num_cell.value,
                "split": True,
                "A": lessons_A,  # may contain 0,1,2 lessons (empty/full/alternating)
                "B": lessons_B,
            }
        else:
            lessons: list[dict] = []

            for row_idx in (lesson_top_idx, lesson_fourth_idx):
                lesson = assemble_lesson(row_idx, a_idx)
                if lesson:
                    if lesson not in lessons:
                        lessons.append(lesson)

            slot = {
                "number": num_cell.value,
                "split": False,
                "lessons": lessons,  # len 0 -> empty, 1 -> weekly, 2 -> alternating
            }

        slots_flat.append(slot)

    day_chunks = [slots_flat[k:k + 8] for k in range(0, len(slots_flat), 8)]

    data["groups"].append({
        "name": group_name,
        "days": day_chunks  # expect 6 elements, each with up to 8 slots
    })

json_file = json.dumps(data, indent=2, ensure_ascii=False)
with open("sched.json", "w", encoding="utf-8") as outfile:
    outfile.write(json_file)

print(f"Parsed {len(data['groups'])} groups and wrote to sched.json")