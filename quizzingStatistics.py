# For ods import
from collections import OrderedDict
import pyexcel_ods3 as pe

# Get files
from os import listdir
from os.path import isfile, join

from functools import partial
from sys import exit
from colorama import init, Fore, Style


def filter_index(elem, divs, index):
    new_divs = []
    for div in divs:
        if len(div) <= index + 1 and elem[0:index] == div[0:index]:
            new_divs.append(div)
    return new_divs


def sort_list_order(elem, sub_divs):
    index = 0
    possible_divs = sub_divs

    while len(possible_divs) > 0:
        index = index + 1
        possible_divs = filter_index(elem, possible_divs, index)
    div_index = get_div_index(elem, sub_divs) * 100000
    index = index - 1

    try:
        return div_index + int(elem[index:])
    except:
        if len(elem) == index:
            return div_index - 1
        elif len(elem) == index + 1:
            if elem[index] == 'W' or elem[index] == 'X' or elem[index] == 'Y' or elem[index] == 'Z':
                return div_index + ord(elem[index])
            else:
                return div_index + ord(elem[index]) + 255
        else:
            return div_index + ord(elem[index] * 250) + elem[index + 1] + 255


def get_all_divs():
    divs = open("stats-settings.txt", "r").read().split('|')[0].split(',')
    main_divs = []
    sub_divs = []
    connected_divs = []
    for div in divs:
        connected_divs.append([div.split(':')[0], div.split(':')[1].split('-')])
        main_divs.append(connected_divs[-1][0])
        for elem in connected_divs[-1][1]:
            sub_divs.append(elem)

    return main_divs, sub_divs, connected_divs


def get_parsed_data(main_divs, sub_divs, connected_divs):
    folder_path = open("stats-settings.txt", "r").read().split('|')[1]
    print("Using settings path to find parsed files: " + folder_path)
    while True:
        try:
            if folder_path != 'MANUAL_MANUAL':
                file_names = [f for f in listdir(folder_path) if isfile(join(folder_path, f)) and f[-4:] == '.ods']
                file_paths = [folder_path + '/' + f for f in listdir(folder_path) if
                              isfile(join(folder_path, f)) and f[-4:] == '.ods']
        except Exception as e:
            print("Path invalid, please enter path manually.")
        else:
            if folder_path != 'MANUAL_MANUAL':
                print(Fore.CYAN + Style.BRIGHT + str(
                    len(file_paths)) + " files found." + Style.RESET_ALL + ' Is this ok? (y/n)')
                ok = input()
            else:
                ok = 'n'
            if ok == 'n':
                folder_path = input('Please give path to parsed quizzes: (example: Quizzes/Day1/Parsed)\n')
            else:
                break

    print("\nLoading, please wait...")
    success = False
    while not success:
        try:
            loaded_files = [pe.get_data(f) for f in file_paths]
            success = True
        except:
            print("Loading the files failed, one of them is invalid, you may have one of them open which is creating "
                  "an invalid temporary file.")
            input("Press enter to try again.\n")
            file_names = [f for f in listdir(folder_path) if isfile(join(folder_path, f)) and f[-4:] == '.ods']
            file_paths = [folder_path + '/' + f for f in listdir(folder_path) if
                          isfile(join(folder_path, f)) and f[-4:] == '.ods']
    print("Done loading.\n")

    print("Verifying validity of data, please wait...")
    fresh_data = OrderedDict()
    for name, file in zip(file_names, loaded_files):
        if 'Sheet1' in file:
            if file['Sheet1'] != [[]] and file['Sheet1'] != {}:
                if get_div(file['Sheet1'][1][1], connected_divs) in sub_divs:
                    fresh_data.update({file['Sheet1'][1][1]: file['Sheet1']})
                else:
                    print(Fore.RED + "ERROR -> File [" + name + "] loading failed, has an invalid division." + Style.RESET_ALL)
            else:
                print(Fore.RED + "ERROR -> File [" + name + "] loading failed, it is empty." + Style.RESET_ALL)
        else:
            print(Fore.RED + "ERROR -> File [" + name + "] loading failed, 'Sheet1' does not exist." + Style.RESET_ALL)
    print("Done verifying.\n")
    print("Sorting, please wait...")

    all_divs = OrderedDict()
    for div in sub_divs:
        all_divs.update({div: []})

    for div in all_divs:
        for key in fresh_data.keys():
            if get_div(key, connected_divs) == div:
                all_divs[div].append(key)

    for div in all_divs:
        all_divs[div].sort(key=partial(sort_list_order, sub_divs=sub_divs))

    data = OrderedDict()
    for div in all_divs:
        for key in all_divs[div]:
            data.update({key: fresh_data[key]})

    print("Done sorting.\n")

    return data


def get_div(quiz, div_connections):
    for index in range(len(div_connections)):
        for sub_div in div_connections[index][1]:
            if quiz[0:len(sub_div)] == sub_div:
                return div_connections[index][0]
    if quiz[0] == "W":
        return "W"
    if quiz[0] == "X":
        return "X"
    if quiz[0] == "Y":
        return "Y"
    if quiz[0] == "Z":
        return "Z"
    return "NULL"


def get_div_index(quiz, sub_divs):
    for index in range(len(sub_divs)):
        if quiz[0:len(sub_divs[index])] == sub_divs[index]:
            return index
    return -1


def get_team_index(draw_table, team_name):
    i = 0
    for row in draw_table:
        if row[0] == team_name:
            return i
        i = i + 1
    return -1


def do_draw():
    print("Doing draw.\n")

    div_data = get_all_divs()
    main_divs = div_data[0]
    sub_divs = div_data[1]
    connected_divs = div_data[2]

    data = get_parsed_data(main_divs, sub_divs, connected_divs)

    if input("Would you like to continue? (y/n):\n") == 'n':
        print("\nRestarting.\n")
        return

    print("Working.\n")
    print("Creating Draw data.\n")
    draw = OrderedDict()

    for div in main_divs:
        draw.update({div + " Div": [
            ["", "---Place---", "---Score---", "---Points---", "---Errors---", "---Place---", "---Score---",
             "---Points---", "---Errors---", "---Place---", "---Score---", "---Points---", "---Errors---",
             "Place", "Score", "Points", "Errors", "Place", "Score", "Points", "Errors", "Place", "Score",
             "Points", "Errors", "Place", "Score", "Points", "Errors", "Place", "Score", "Points", "Errors",
             "Place", "Score", "Points", "Errors", "Place", "Score", "Points", "Errors", "Place", "Score", "Points",
             "Errors", "Place", "Score", "Points", "Errors", "Place", "Score", "Points", "Errors", "Place", "Score",
             "Points", "Errors", "Place", "Score", "Points", "Errors", "Place", "Score", "Points", "Errors", "Place",
             "Score", "Points", "Errors", "Place", "Score", "Points", "Errors", "Place", "Score", "Points", "Errors",
             "Place", "Score", "Points", "Errors"]]})

        print("Moving " + div + " Div.")
        count = 0
        for key in data.keys():
            quiz = data[key][1][1]
            if get_div(quiz, connected_divs) == div and quiz[-1] != 'W' and quiz[-1] != 'X' and quiz[-1] != 'Y' and \
                    quiz[-1] != 'Z':
                for i in range(3):
                    team_name = data[key][i + 1][0]
                    index = get_team_index(draw[div + " Div"], team_name)
                    if index == -1:
                        draw[div + " Div"].append([team_name])
                    for j in range(4):
                        draw[div + " Div"][index].append(data[key][i + 1][j + 2])
                count = count + 1
        print(Fore.CYAN + (Style.BRIGHT if count != 0 else '') + str(
            count) + ' Entries moved in ' + div + ' Div\n' + Style.RESET_ALL)

    draw.update({"WXYZ": [
        ["", "---Place---", "---Score---", "---Points---", "---Errors---", "---Place---", "---Score---",
         "---Points---", "---Errors---", "---Place---", "---Score---", "---Points---", "---Errors---",
         "Place", "Score", "Points", "Errors", "Place", "Score", "Points", "Errors", "Place", "Score",
         "Points", "Errors", "Place", "Score", "Points", "Errors", "Place", "Score", "Points", "Errors",
         "Place", "Score", "Points", "Errors", "Place", "Score", "Points", "Errors", "Place", "Score", "Points",
         "Errors", "Place", "Score", "Points", "Errors", "Place", "Score", "Points", "Errors", "Place", "Score",
         "Points", "Errors", "Place", "Score", "Points", "Errors", "Place", "Score", "Points", "Errors", "Place",
         "Score", "Points", "Errors", "Place", "Score", "Points", "Errors", "Place", "Score", "Points", "Errors",
         "Place", "Score", "Points", "Errors"]]})

    print("Moving WXYZ.")
    count = 0
    for key in data.keys():
        quiz = data[key][1][1]
        if quiz[-1] == 'W' or quiz[-1] == 'X' or quiz[-1] == 'Y' or quiz[-1] == 'Z':
            for i in range(3):
                team_name = data[key][i + 1][0]
                index = get_team_index(draw["WXYZ"], team_name)
                if index == -1:
                    draw["WXYZ"].append([team_name])
                for j in range(4):
                    draw["WXYZ"][index].append(data[key][i + 1][j + 2])
            count = count + 1
    print(Fore.CYAN + (Style.BRIGHT if count != 0 else '') + str(count) + ' Entries moved in WXYZ\n' + Style.RESET_ALL)

    draw_name = input("How would you like to save the draw file (no .ods extension)?\n")

    while True:
        try:
            print("Saving.\n")
            pe.save_data(draw_name + ".ods", draw)
            break
        except:
            print(Fore.RED + "Saving failed, you may have a file of the same name open." + Style.RESET_ALL)
            input("Press enter to try again.\n")

    print(Fore.GREEN + "Saving succeeded.\n" + Style.RESET_ALL)
    print("Returning to main.\n")


def do_summary():
    print("Doing Summary.\n")

    div_data = get_all_divs()
    main_divs = div_data[0]
    sub_divs = div_data[1]
    connected_divs = div_data[2]

    data = get_parsed_data(main_divs, sub_divs, connected_divs)

    if input("Would you like to continue? (y/n):\n") == 'n':
        print("\nRestarting.\n")
        return

    print("Working.\n")
    print("Creating Summary data.\n")
    summary = OrderedDict()
    for div in main_divs:
        summary.update({div + " Div Team": [["Team", "Quiz", "Place", "Score", "Points", "Errors", "Quiz No"]]})
        summary.update({div + " Div Quizzer": [["Quizzer", "Team", "Quiz", "Points", "Errors", "Jumps", "Quiz No"]]})

    all_count = 0
    for div in main_divs:
        print("Moving " + div + " Div Team.")
        count = 0
        for key in data.keys():
            quiz = data[key][1][1]
            if get_div(quiz, connected_divs) == div:
                for i in range(3):
                    summary[div + ' Div Team'].append(data[key][i + 1][0:6])
                summary[div + ' Div Team'][-3].append(summary[div + ' Div Team'][-3][1])
                count = count + 1
        print(Fore.CYAN + (Style.BRIGHT if count != 0 else '') + str(
            count) + ' Entries moved in ' + div + ' Div Team\n' + Style.RESET_ALL)
        all_count = all_count + count

        print("Moving " + div + " Div Quizzer.")
        count = 0
        for key in data.keys():
            quiz = data[key][1][1]
            if get_div(quiz, connected_divs) == div:
                for i in range(15):
                    summary[div + ' Div Quizzer'].append(data[key][i + 6][0:6])
                summary[div + ' Div Quizzer'][-15].append(summary[div + ' Div Quizzer'][-15][2])
                count = count + 1
        print(Fore.CYAN + (Style.BRIGHT if count != 0 else '') + str(
            count) + ' Entries moved in ' + div + ' Div Quizzer\n' + Style.RESET_ALL)

    summary_name = input("How would you like to save the summary file (no .ods extension)?\n")

    while True:
        try:
            print("Saving.\n")
            pe.save_data(summary_name + ".ods", summary)
            break
        except:
            print(Fore.RED + "Saving failed, you may have a file of the same name open." + Style.RESET_ALL)
            input("Press enter to try again.\n")

    print(Fore.GREEN + "Saving succeeded.\n" + Style.RESET_ALL)
    print("Returning to main.\n")


def create_settings():
    print('\n\n(1/2) What divisions are in these quizzes?\n'
                 'First put the main category, colon, then sub-categories separated by hyphens. With separate entries separated by comas.\n'
                 '(example: A:A-K, B:B-M, C:C-D) ' + Fore.GREEN + '(or (d)efault):' + Style.RESET_ALL)
    divs = input()
    divs = divs.replace(' ', '')
    if divs == 'd':
        divs = 'A:A-K,B:B-M,C:C-D'
        print('Default selected: ' + divs + '\n')

    print('(2/2) Please give path to parsed quizzes: (example: Quizzes/Day1/Parsed)' + Fore.GREEN + ' (or (m)anual):' + Style.RESET_ALL)
    folder_path = input()
    if folder_path == 'm':
        folder_path = 'MANUAL_MANUAL'
        print('Manual selected.\n')
    file = open("stats-settings.txt", "w")
    file.write(divs + '|' + folder_path)
    file.close()
    print("Settings saved successfully.\n\n")


def main():
    print(Fore.YELLOW + Style.BRIGHT + "This program was made by Timothy Letkeman for Canadian Midwest Quizzing, 2019.")
    print("If you need to contact Timothy Letkeman: tiggy02@gmail.com.\n" + Style.RESET_ALL)
    if not isfile("stats-settings.txt"):
        print("This computer has no settings, you must create them.\n")
        create_settings()

    choice = 'start'
    while True:
        choice = input('Select an option: (s)ummary, (d)raw, (r)eset settings, or (q)uit?\n')
        if choice != 's' and choice != 'd' and choice != 'q' and choice != 'r':
            print('Invalid option.')
        if choice == 's':
            do_summary()
        if choice == 'd':
            do_draw()
        if choice == 'r':
            create_settings()
        if choice == 'q':
            exit()


init()
if __name__ == '__main__':
    main()
