#!/usr/bin/env python3

###
## Call: python3 main.py advisory.xlsx class_1.xlsx class_2.xlsx ... class_n.xlsx outfile.xlsx
###

import sys, os
import excel_io

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

    # Make students in class have same number as they have in advisory
    for i, lis in enumerate(lists[1:]):
        for j, name in enumerate(lis):
            if name in lists[0]:
                for k in range(lists[0].index(name) - j):
                    lis.insert(j, "")

    # Make students in advisory have same number as they have in class
    for i, name in enumerate(lists[0]):
        for j, lis in enumerate(lists[1:]):
            if name != "" and name in lis:
                for k in range(lis.index(name) - i):
                    lists[0].insert(i, "")

    # Fix students that are too low in advisory
    for i, name in enumerate(lists[0]):
        for j, lis in enumerate(lists[1:]):
            if name != "" and name in lis:
                if i != lis.index(name) and lists[0][lis.index(name)] == "":
                    lists[0][lis.index(name)] = name
                    lists[0][i] = ""

    # Insert advisory-only students as low as possible
    for i in range(28 - len(lists[0])):
        lists[0].append("")
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
    if (len(sys.argv) < 3):
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