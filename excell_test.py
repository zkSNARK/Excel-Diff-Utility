"""A utility to diff 2 Excel files and report the cells which have changed.

This utility is for people who for some reason need to make save points
to an excel file and report the differences between save points into the
console.

Why this functionality isn't built into Excel, I don't know.  But I do know
that it was easy to make... so what's up Microsoft?

Expects an Excel file to be under version control in a local repo. If you
do not have an excel file under version control (what normal human would?),
you can create a repo (on a linux or mac computer) with ...
    --make_git

Pass the file you want to diff with ...
    --file <filename>


Does not handle credential input or other git configuration elements.

Author : Chris Goebel

License : Anyone can do anything with this file.  Submit PR's and I'll add
them if they are reasonable.
"""

import argparse
import os
import datetime
import subprocess
import shutil
import git

import pandas as pd
import numpy as np

from xlsxwriter.utility import xl_rowcol_to_cell


def is_git_directory(path='.'):
    return subprocess.call(['git', '-C', path, 'status'], stderr=subprocess.STDOUT, stdout=open(os.devnull, 'w')) == 0


def handle_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("-g", "--make_git", help="create local git repo if none exists", action="store_true")
    parser.add_argument("-f", "--file", help="The Excel file to track.", action="store_true")
    args = parser.parse_args()
    return args.make_git, args.file


def tmp_name(excel_file):
    f = excel_file.split('.')
    return f[0] + "_tmp." + f[1]


def copy_file_to_tmp(excel_file):
    shutil.copyfile(excel_file, tmp_name(excel_file))


def reset_file_in_git(commit_id: str, excel_file: str):
    process = subprocess.Popen(["git", "checkout", commit_id, excel_file], stdout=subprocess.PIPE)
    output = process.communicate()[0]


def diff_old_and_new_file(file1, file2):
    """Compare 2 excel files and report differences.

    Uses the general idea from https://stackoverflow.com/a/52763561/6840486
    and adds handling for Excel work sheets.
    """

    xlsx_template = pd.ExcelFile(file1)
    xlsx_test = pd.ExcelFile(file2)

    for sheet in xlsx_template.sheet_names:
        template = pd.read_excel(xlsx_template, sheet, na_values=np.nan, header=None)
        test_sheet = pd.read_excel(xlsx_test, sheet, na_values=np.nan, header=None)

        rt, ct = template.shape  # row, col of template
        rtest, ctest = test_sheet.shape  # row, col of test

        df = pd.DataFrame(columns=['Cell_Location', 'Previous_Value', 'Current_Value'])

        for rowNo in range(max(rt, rtest)):
            for colNo in range(max(ct, ctest)):
                # Fetching the template value at a cell
                try:
                    template_val = template.iloc[rowNo, colNo]
                except:
                    template_val = np.nan

                # Fetching the testsheet value at a cell
                try:
                    test_sheet_val = test_sheet.iloc[rowNo, colNo]
                except:
                    test_sheet_val = np.nan

                # Comparing the values
                if str(template_val) != str(test_sheet_val):
                    cell = xl_rowcol_to_cell(rowNo, colNo)
                    dfTemp = pd.DataFrame([[cell, template_val, test_sheet_val]],
                                          columns=['Cell_Location', 'Previous_Value', 'Current_Value'])
                    df = df.append(dfTemp)

        if not df.index.empty:
            print("CHANGES DETECTED IN SHEET : " + sheet)
            print(df)
            print()


def copy_tmp_back(excel_file):
    shutil.copyfile(tmp_name(excel_file), excel_file)


def delete_file(file):
    os.remove(file)


def commit_changes(excel_file):
    process = subprocess.Popen(["git", "add", excel_file], stdout=subprocess.PIPE)
    output = process.communicate()[0]
    process = subprocess.Popen(["git", "commit", "-m", f"File auto-updated {datetime.datetime.now()}"],
                               stdout=subprocess.PIPE)
    output = process.communicate()[0]


def changes_detected(file):
    repo = git.Repo(".")
    changed = [item.a_path for item in repo.index.diff(None)]
    return file in changed


def changes_detected_from_commit(file, commit):
    changed = [item.a_path for item in commit.diff(None)]
    return file in changed


def create_repo_here():
    repo = git.Repo.init(".")
    return repo


def select_commit_for_diff(repo: git.Repo):
    l = []
    print()
    print("Select from available checkpoints.")
    print("----------------------------------")
    for i, c in enumerate(repo.iter_commits()):
        print(f"  [{i}] Save point : {c}: {c.authored_datetime}")
        l.append(c)

    print()
    user_select = int(input("Which checkpoint would you like to diff against (enter for most recent)? "))
    return l[user_select]


def main():
    make_git, excel_file = handle_args()
    if make_git:
        if is_git_directory():
            repo = git.Repo(".")
            print("Git repo creation requested through flags, but repo already exists in current directory.")
        else:
            repo = create_repo_here()
            if excel_file:
                repo.index.add([excel_file])
                repo.index.commit("initial commit")
                print("Git repo created in current directory and added file: " + excel_file + ".")
            else:
                print("Created git repo in current directory.")
    else:
        repo = git.Repo(".")

    if excel_file:
        print("requested to run on file : " + excel_file)
    else:
        print("Using default file : test.xlsx")
        excel_file = "test.xlsx"

    while True:
        commit = select_commit_for_diff(repo)
        changed = changes_detected_from_commit(excel_file, commit)

        if changed:
            print("Changes detected.")
            copy_file_to_tmp(excel_file)

            # select commit for diff and reset file in git to that commit
            reset_file_in_git(str(commit), excel_file)

            diff_old_and_new_file(excel_file, tmp_name(excel_file))

            copy_tmp_back(excel_file)
            delete_file(tmp_name(excel_file))
            commit_changes(excel_file)
        else:
            print("no changes detected")


if __name__ == "__main__":
    main()
