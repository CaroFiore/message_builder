import openpyxl as xl
import datetime as dt
import pyperclip as clipboard

TAKENROOSTER = "takenrooster.xlsx"
TODAY = dt.datetime(2026,11,16)

def create_message(today):
    xlfile = xl.load_workbook(TAKENROOSTER)
    sheet = xlfile.active

    # Idk why but I don't want to bother with it. Anyway you get a list of tuples
    _data = list(sheet.iter_rows(values_only=True))
    data = []
    for item in _data:
        data.append(item)

    headers = data[0]
    matches = [dict(zip(headers, row)) for row in data[1:]]

    # filter matches in the upcoming week and add them to the task list
    tasks = []

    for match in matches:
        if match["datum"] < today + dt.timedelta(7) and match["datum"] > today:
            task = {}
            # used for grouping
            task["datum"] = match["datum"] # datetime

            task["dag"] = match["datum"].strftime("%A")# string
            task["start"] = match["van"].strftime("%H:%M") # datetime
            if match["tot"]:
                task["eind"] = match["tot"].strftime("%H:%M") # datetime
            else:
                task["eind"] = ""

            task["thuisteam"] = match["NVC"]
            task["taaksoort"] = match["taak"]

            if match["team"]:
                task["ingedeeld"] = match["team"]
            elif match["naam"]:
                task["ingedeeld"] = match["naam"]
            else:
                task["ingedeeld"] = "(Nog) niet ingedeeld, oeps.."

            task["opmerking"] = match["opmerking"]

            tasks.append(task)

    if not tasks:
        print("No tasks found, quitting..")
        return []

    # Sort these days by name
    day_order = {
        "Monday": 0,
        "Tuesday": 1,
        "Wednesday": 2,
        "Thursday": 3,
        "Friday": 4,
        "Saturday": 5,
        "Sunday": 6
    }   

    date_groups = set(task["dag"] for task in tasks)

    date_groups = sorted(date_groups, key = lambda day: day_order[day])

    day_translate = {
        "Monday" : "Maandag",
        "Tuesday" : "Dinsdag",
        "Wednesday" : "Woensdag",
        "Thursday" : "Donderdag",
        "Friday" : "Vrijdag",
        "Saturday" : "Zaterdag",
        "Sunday" : "Zondag"
    }


    # Group by day. Don't bother comprehending. Point is that we get a list like this:
    # day1 : [tasks], day2 : [tasks]
    day_schedule = {}

    for day in date_groups:
        day_list = []
        for task in tasks:
            if task["dag"] == day:
                day_list.append(task)
        day_schedule[day] = day_list
    print(day_schedule)


    # Okay, lets make a message now shall we
    lines = []
    lines.append(f"Goeiemorgen, dit is een automatisch bericht! Het is vandaag {today.strftime('%d-%m-%Y')}.")
    
    for day in day_schedule:
        lines.append("-" * 50)
        lines.append(f"{day_translate[day]}:")

        for task in day_schedule[day]:
            team = f" bij {task['thuisteam'].lower()}" if task["thuisteam"] else ""

            if task["taaksoort"] == "zaalwacht":
                task["taaksoort"] = "Zaalwacht " + task["opmerking"]
                task["start"] = ""

            lines.append(
                f"{task['start']} "
                f"{task['ingedeeld']} heeft taak {task['taaksoort']}"
                f"{team}"
            )

    lines.append("-" * 50)
    lines.append("Klopt dit totaal niet? Stuur een berichtje naar Pascal!")

    return lines

if __name__ == "__main__":
    message_to_copy = ""
    for line in create_message(TODAY):
        message_to_copy += line + "\n"

    print(message_to_copy)
    clipboard.copy(message_to_copy)