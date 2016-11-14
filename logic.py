import sys
import os
import string
import xlwt
from inn_stems import STEMS as INN_STEMS


def get_stem_conflicts(names_list, ignore, file_name="stem_conflicts"):
    # Create dictionary with each line as a key and each value as a list of stem conflicts
    names_dict = {}
    for line in names_list:
        names_dict[line] = find_stems(strip_names(line), ignore)

    build_doc(names_list, names_dict, file_name)


def strip_names(line):
    # Gets stripped name on each line
    full_line = line.replace("[", "(").replace("]", ")").replace("*", "(").split("(")
    names_on_line = full_line[0].split("/")

    # Returns a list of each name on the line
    return [x.strip(string.whitespace).lower() for x in names_on_line if x.strip(string.whitespace)]


def find_stems(names_list, ignore):
    # Find stems for each name on the line
    conflicts = []

    # checks each name on the line and returns a list of any/all conflicts found
    for name in names_list:
        for infix in INN_STEMS["Infix"]:
            #fix = infix.strip("-")
            if infix in name and infix not in ignore:
                hit = "-" + infix + "-"
                defin = INN_STEMS["All"][hit]
                if (" (" + defin + ")") not in  " ".join(conflicts) and "USAN" not in defin:
                    conflicts.append(hit + " (" + defin + ")")

        for prefix in INN_STEMS["Prefix"]:
            #fix = prefix.strip("-")
            if name[:len(prefix)] == prefix and prefix not in ignore:
                hit = prefix + "-"
                defin = INN_STEMS["All"][hit]
                if (" (" + defin + ")") not in  " ".join(conflicts) and "USAN" not in defin:
                    conflicts.append(hit + " (" + defin + ")")

        for suffix in INN_STEMS["Suffix"]:
            #fix = suffix.strip("-")
            if name[-len(suffix):] == suffix and suffix not in ignore:
                hit = "-" + suffix
                defin = INN_STEMS["All"][hit]
                if (" (" + defin + ")") not in " ".join(conflicts) and "USAN" not in defin:
                    conflicts.append(hit + " (" + defin + ")")

    return conflicts


def build_doc(names_list, names_dict, file_name):
    # Builds an excel doc containing each line of names with all
    # associated conflicts in separate cells

    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Names and Conflicts", cell_overwrite_ok=True)


    h_style = xlwt.easyxf('font: name Verdana, bold True; align: vert centre, \
        horiz centre; borders: left thin, top thin, bottom thin, right thin')
    t_style = xlwt.easyxf('font: name Verdana, bold False; align: vert centre, \
        wrap on; borders: left thin, top thin, bottom thin, right thin')
    # Build columns 0 and 1; The first column(0) has the names and rationale that are pasted in
    # Second column(1) shows just the names that were screened, for PMs to easily double check as needed
    sheet.write(0, 0, "Name (rationale)", style=h_style)
    sheet.write(0, 1, "Names screened", style=h_style)
    sheet.col(0).width = 60 * 256
    sheet.col(1).width = 18 * 256
    sheet.row(0).height = 1000

    # Add empty items to each list of conflicts so there is an equal number of total "conflicts"  for every name
    # This ensures that all cells are written to and bordered properly (mostly to make it look pretty look "pretty")

    max_cols = max([len(names_dict[x]) for x in names_dict])
    for x in names_dict:
        if len(names_dict[x]) < max_cols:
            for i in range(max_cols - len(names_dict[x])):
                names_dict[x].append(" ")

    # Write names to each row in both the first and second columns (as discussed above)
    for x in range(1, len(names_list)+1):
        # Write the names as they were pasted in
        sheet.write(x, 0, names_list[x-1], t_style)

        # Write the names stripped of rationale/notes/pronunciation
        names = "\n".join(strip_names(names_list[x-1])).title()
        sheet.write(x, 1, names, t_style )

        # Format the height for each row
        sheet.row(x).height_mismatch = True
        sheet.row(x).height = 1000

        total_cols = 1

        # Write each conflict to cells in the same row as the associated names
        for conflict in names_dict[names_list[x-1]]:
            sheet.write(x, total_cols+1, conflict, t_style)
            sheet.write(0, total_cols+1, "INN", h_style)
            sheet.col(total_cols).width = 18 * 256
            total_cols += 1



    workbook.save(file_name)


    return # Nothing to return
