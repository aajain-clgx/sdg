"""Utility for SDG project"""

import re


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
