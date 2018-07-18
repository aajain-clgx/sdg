#!/usr/bin/env python3

import os
import sys
import argparse
from oauth2client.service_account import ServiceAccountCredentials
try:
    import gspread
except:
    print("Please install gspread module")
    print("pip3 install gspread")
    sys.exit(1)
try:
    from prettytable import PrettyTable
except:
    print("Please install prettytable module")
    print("pip3 install prettytable")
    sys.exit(1)

import similarity
from collections import defaultdict
import traceback
import pprint

CLIENT_CREDENTIALS = "/home/bear/.google/sdg_id.json"
GOOGLE_CREDENTIALS = None


def initialize_google():
    """Initialize Google API using client credentials
       Store client oauth json credentials file in 
       CLIENT_CREDENTIALS folder
       To setup oauth: http://gspread.readthedocs.io/en/latest/oauth2.html
    """
    
    # Validate credentials client file
    if os.path.exists(CLIENT_CREDENTIALS):
        scope = [ 'https://spreadsheets.google.com/feeds',
                  'https://www.googleapis.com/auth/drive']

        cred =  ServiceAccountCredentials.from_json_keyfile_name(CLIENT_CREDENTIALS, scope)
        global GOOGLE_CREDENTIALS
        GOOGLE_CREDENTIALS = gspread.authorize(cred)
        
    else:
        raise Exception("Please create Google OAUTH client and copy it in '{}'".format(CLIENT_CREDENTIALS))
    

def open_spreadsheet(spreadsheet_name):
    """Open spreadsheet by name and return"""

    global GOOGLE_CREDENTIALS
    return GOOGLE_CREDENTIALS.open(spreadsheet_name)


def validate(spreadsheet, args):
    """Perform various cross validations on worksheets 
        in a spreadsheet and generate output a report"""
 
    def find_similar_text(col, args):
        """ Uses similarity.py module to find similar text
            * We remove all duplicates first
            * We store row location of all indicator text
            * We apply NLP similarity on deduped list
            * In fast mode, we use sorted list and start comparing from the location
              of the first string.
            * In deep mode, we search every sentence with other
        """

        col_line_dict = defaultdict(list)
        for n,x in enumerate(col):
            col_line_dict[x].append(n+1)

        colunique = list(set(col))
        colunique.sort()

        similar_lines = []

        for n1, val1 in enumerate(colunique):
            if val1 == '':
                continue
            for n2, val2 in enumerate(colunique[n1:]):
                if val2 == '':
                    continue
                if val1 == val2:
                    continue
                similar_val = similarity.cosine_sim(val1, val2)
                if similar_val > 0.7:
                    similar_lines = [col_line_dict[val1], val1, col_line_dict[val2], val2]
                    print("\n Row: {}\n'{}'\n----\n Row: {}\n'{}'\nSimilarity Score = {:.3f}\n\n\n ====".format(
                            similar_lines[0], similar_lines[1],
                            similar_lines[2], similar_lines[3], similar_val))
                else:
                    if not args.deeply_similar:
                        break


    def validate_sdg_compass_metrics_sheet(wks):
        """Perform validate on 'SDG Compass Metrics sheet
           * Finds duplicate SDG Goals
           * Finds duplicate SDG Target
        """
    
        def finddups(col, categoryname):
            unique_dict = defaultdict(set)
            mismatches = defaultdict(list)

            for n, item in enumerate(col):
                # Ignore first line since it is title
                if n < 1:
                    continue
                x = item.split()[0]
                cid = x[:-1] if x[-1] == "." else x
                if item not in unique_dict[cid]:
                    mismatches[cid].append([item, n+1])
                unique_dict[cid].add(item)

            remove = [item for item in mismatches if len(mismatches[item]) > 1]
            mismatches = { k:mismatches[k] for k in mismatches if k in remove}
        
            if len(mismatches) > 0:
                print("\n\n Test for duplicate values for {} found: Failed".format(categoryname))
                print("\n ------------")
                print("\n {} ID -> list of (Mismatched Values, Mismatched Rows)".format(categoryname))
                print("\n ------------")
                pprint.pprint(mismatches)
            else:
                print("\n\n Test for duplicate values for {} found: Passed".format(categoryname))
                

        goals = wks.col_values(1)
        finddups(goals, "SDG Goal")

        targets = wks.col_values(2)
        finddups(targets, "SDG Target")

                    
    def validate_bia_sdg_mapping_sheet(wks):
        """Performs validation on 'BIA to SDG mapping' sheet
           Validates that we have identical DirectTargets and IndirectTargets
           for rows with same Concept Code.

           Assumptions:
            * First two rows are ignored by the script since they contain column title. 
              If you add additional title row, script will need to be updated.
            * DirectTargets/IndirectTargets columns: If the line begins with alphabet, 
              it is ignored by script. If it begins with a numerical, it is used for
              building graph mapping
        """
       
        def crossvalidate(concepts_column, colnumber):
            """Helper function for validate_bia_sdg_mapping_sheet
               Common code to validate either DirectTarget or IndirectTarget
            """

            concepts_dict = {}
            targets = wks.col_values(colnumber)
            mismatches = defaultdict(list)
                
            # Google sheets cleanup           
            difflen = len(targets) - len(concepts_column)
            if difflen < 0:
                targets.extend([''] * abs(difflen))
                            
            for n,x in enumerate(concepts_column):
                # Ignore first two lines since they contain titles
                if n < 2:
                    continue
                if x in concepts_dict:
                    if concepts_dict[x][0] != targets[n]:
                        if x not in mismatches:
                            mismatches[x].append(concepts_dict[x][1])
                            mismatches[x].append(n+1)
                else:
                    concepts_dict[x] = [targets[n], n+1]

            return mismatches


        # Read Concepts, Direct_Targets, Indirect_Targets column
        concepts_column = wks.col_values(1)
        direct_mismatches = crossvalidate(concepts_column, 33)
        indirect_mismatches = crossvalidate(concepts_column, 34)

        if len(direct_mismatches) == 0:
            print("\n Test for mismatched DirectTarget text found (for same ConceptCode): Passed")
        else:
            print("\n Test for mismatched DirectTarget text found (for same ConceptCode): Failed")
            table = PrettyTable(["Concept Code", "Mismatched Text Rows"])
            table.border = True

            for x in direct_mismatches:
                table.add_row([x, ",".join((str(y) for y in direct_mismatches[x]))])
            print(table)

        if len(indirect_mismatches) == 0:
            print("\n Test for mismatched IndirectTarget text found (for same ConceptCode): Passed")
        else:
            print("\n Test for mismatched IndirectTarget text found (for same ConceptCode): Failed")
            table = PrettyTable(["Concept Code", "Mismatched Text Rows"])
            table.border = True
            for x in indirect_mismatches:
                table.add_row([x, ",".join((str(y) for y in indirect_mismatches[x]))])
            print(table)


    worksheet_name = "BIA to SDG mapping"
    wks = spreadsheet.worksheet(worksheet_name)
    print('\n\n{0:-^60}\n'.format('Validate worksheet: {}'.format(worksheet_name)))
    validate_bia_sdg_mapping_sheet(wks)

    worksheet_name = "SDG Compass Metrics"
    print('\n\n{0:-^60}\n'.format('Validate worksheet: {}'.format(worksheet_name)))
    wks = spreadsheet.worksheet(worksheet_name)
    validate_sdg_compass_metrics_sheet(wks)
    
    print('\n\n{0:-^60}\n'.format('Finding similar Indicator value candidates'))
    metricwks = spreadsheet.worksheet("SDG Compass Metrics")
    find_similar_text(metricwks.col_values(6), args)


def main():
 
    # Define inputs to scripts
   
    parser = argparse.ArgumentParser(
        description="Tool to manipulate SDG spreadsheet such as validation, build graph links")

    parser.add_argument(
        "action",
        action = "store",
        choices = ["validate"],
        help = "Pass one of these action for script to perform"
    )

    parser.add_argument(
        "--sheet",
        action = "store",
        default = "BIA V5 SDG Alignment",
        help = "Name of Google spreadsheet to open"
    )
        
    parser.add_argument(
        "--deeply-similar",
        type=bool,
        default = False,
        help = "Perform exhaustive similarity search between text. Warning: Takes long time"
    )
    
    args = parser.parse_args()

    try:

        # Initialize Google and open Spreadsheet        
        initialize_google()
        ssheet = open_spreadsheet(args.sheet)
        
        if args.action == "validate":
            validate(ssheet, args)
     
        # Test to read DirectTarget column
        wks = ssheet.worksheet("BIA to SDG mapping")
        targets = [a.strip() for a in wks.acell("AH47").value.splitlines() if not(a.strip()[0].isalpha())]
        

    except Exception as ex:
        traceback.print_exc()
        print("Error while processing script: {}".format(ex))
        
 


if __name__ == "__main__":
  main()
