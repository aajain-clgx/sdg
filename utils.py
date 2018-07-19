"""Utility for SDG project"""

import re
from collections import defaultdict


def get_column(matrix, colnumber):
    """Utility function to extract column from a matrix containing entire worksheet"""
    return [a[colnumber] for a in matrix]

def build_target_list(colval):
    """Generates a list of targets from a cell"""

    nonempty = (a.strip() for a in colval.splitlines() if a.strip() != '')
    tlist = [a for a in nonempty if not(a[0].isalpha() or a[0] == '?')]
    tlist.sort()
    return tlist

def get_target_list(worksheet, row, col):
    colval = worksheet[row][col]
    return build_target_list(colval)

def validate_target_format(wks, colnumber):
    """Helper function to validate the syntax of a target specified in 
        Direct or Indirect Target column
        The format should be like digit.digit or digit.alpha
    """
    syntax = re.compile("^\d{1,2}\.([0-9]{1,2}|[a-z])$")
    mismatches = defaultdict(list)

    for n,x in enumerate(wks):
        target_list = get_target_list(wks, n, colnumber)
        invalid_list = [a for a in target_list if syntax.match(a) is None]
        if len(invalid_list) > 0:
            mismatches[n+1].extend(invalid_list)

    return mismatches


def finddups(wks, colnumber):

    unique_dict = defaultdict(set)
    mismatches = defaultdict(list)

    for n, item in enumerate(wks):
        # Ignore first line since it is title
        x = item[colnumber].split()[0]
        cid = x[:-1] if x[-1] == "." else x
        if item[colnumber] not in unique_dict[cid]:
            mismatches[cid].append([item[colnumber], n+1])
        unique_dict[cid].add(item[colnumber])

    remove = [item for item in mismatches if len(mismatches[item]) > 1]
    mismatches = { k:mismatches[k] for k in mismatches if k in remove}

    return mismatches
