#!/usr/bin/env python3

###
## Call: python3 main.py advisory.xlsx class_1.xlsx class_2.xlsx ... class_n.xlsx outfile.xlsx
###

import sys, os
import excel_io

MAX_STUDENTS = 28

###
## Calculate computer assignments
###
def calc_assignments(lists):
    # Remove students that are only in advisory
    adv_only = []
    for name in lists[0]:
        if not any(name in lis for lis in lists[1:]):
            adv_only.append(name)
            lists[0].remove(name)
    adv_class = lists[0]
    lists[0] = [""] * 28

    # Iterate through classes row by row
    # If student is in advisory, place them in the same row as they are in class
    # If there is already an advisory student in that row, shift down the class with fewer students
    for i in range(MAX_STUDENTS):
        num_pushes = 0
        repeat = -1
        for j, lis in enumerate(lists[1:]):
            if len(lis) > i and lis[i] != "":
                if lis[i] in adv_class and lists[0][i] == "":
                    lists[0][i] = lis[i]
                    adv_class.remove(lis[i])
                    repeat = j + 1
                elif lis[i] in adv_class and lists[0][i] != "":
                    if len(lists[repeat]) < len(lis):
                        # Insert new name at i
                        lists[0].insert(i + num_pushes, lis[i])
                        lists[0].pop
                        adv_class.remove(lis[i])
                        # Shift previous class down
                        for k in range(num_pushes + 1):
                            lists[repeat].insert(i, "")
                        # Update tracking variables
                        repeat = j + 1
                        num_pushes += 1
                    else:
                        # Insert new name at i + num_pushes
                        lists[0][i + num_pushes + 1] = lis[i]
                        adv_class.remove(lis[i])
                        # Shift current class down
                        for k in range(num_pushes + 1):
                            lis.insert(i, "")
                        # Update tracking variables
                        num_pushes += 1

    # Insert advisory-only students as low as possible
    for i in range(MAX_STUDENTS - len(lists[0])):
        lists[0].append("")
    for i in range(len(lists[0]) - MAX_STUDENTS):
        lists[0].pop()
    for name in adv_only[::-1]:
        for i, elem in enumerate(lists[0][::-1]):
            if elem == "":
                lists[0][27 - i] = name
                break

    return lists

###
## Main
###
if __name__ == "__main__":
    # Check arguments
    if (len(sys.argv) < 4):
        sys.exit("Wrong arguments!")

    # Extract class lists from Excel files
    class_lists = []
    class_names = []
    for file in sys.argv[1:-1]:
        class_names.append(file[:-11])
        class_lists.append(excel_io.read_class_lists(os.path.relpath("rosters\\" + file)))

    # Calculate computer assignments
    assignments = calc_assignments(class_lists)

    # Write computer assignments to output Excel file
    excel_io.write_computer_assignments(sys.argv[-1], assignments, class_names)