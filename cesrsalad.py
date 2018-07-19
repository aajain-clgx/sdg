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
import utils
from collections import defaultdict
import traceback
import pprint

CLIENT_CREDENTIALS = "~/.google/sdg_id.json"
GOOGLE_CREDENTIALS = None


def initialize_google():
    """Initialize Google API using client credentials
       Store client oauth json credentials file in 
       CLIENT_CREDENTIALS folder
       To setup oauth: http://gspread.readthedocs.io/en/latest/oauth2.html
    """
    
    # Validate credentials client file
    client_cred_file = os.path.expanduser(CLIENT_CREDENTIALS)
    if os.path.exists(client_cred_file):
        scope = [ 'https://spreadsheets.google.com/feeds',
                  'https://www.googleapis.com/auth/drive']

        cred =  ServiceAccountCredentials.from_json_keyfile_name(client_cred_file, scope)
        global GOOGLE_CREDENTIALS
        GOOGLE_CREDENTIALS = gspread.authorize(cred)
        
    else:
        raise Exception("Please create Google OAUTH client and copy it in '{}'".format(CLIENT_CREDENTIALS))
    

def open_spreadsheet(spreadsheet_name):
    """Open spreadsheet by name and return"""

    global GOOGLE_CREDENTIALS
    return GOOGLE_CREDENTIALS.open(spreadsheet_name)


def validate(worksheets, args):
    """Perform various cross validations on worksheets 
        in a spreadsheet and generate output a report"""
 
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
       
        def validate_target_format(colnumber, coltitle):
            """Helper function to validate the syntax of a target specified in 
               Direct or Indirect Target column
               The format should be like digit.digit or digit.alpha
            """

            mismatches = utils.validate_target_format(wks, colnumber)

            if len(mismatches) == 0:
                print("\n Test invalid syntax for {} Targets: Passed".format(coltitle))
            else:
                print("\n Test invalid syntax for {} Targets: Failed".format(coltitle))
                table = PrettyTable(["Row Number", "Invalid Systax {} Targets".format(coltitle)])
                table.border = True

                for x in mismatches:
                    table.add_row([x, ",".join("'{}'".format(str(y)) for y in mismatches[x])])
                print(table)


        def crossvalidate_with_concept_code(colnumber, coltitle):
            """Helper function for validate_bia_sdg_mapping_sheet
               Common code to validate either DirectTarget or IndirectTarget
            """

            concepts_dict = {}
            mismatches = defaultdict(list)
                            
            for n,x in enumerate(wks):
                # Ignore first two lines since they contain titles
                target_list = utils.get_target_list(wks,n, colnumber)
                concept_code = x[0]
                if concept_code in concepts_dict:
                    if concepts_dict[concept_code][0] != target_list:
                        if concept_code not in mismatches:
                            if concepts_dict[concept_code] not in mismatches[concept_code]:
                                mismatches[concept_code].append(concepts_dict[concept_code])
                            mismatches[concept_code].append([target_list, n+1])
                else:
                    concepts_dict[concept_code] = [target_list, n+1]

            if len(mismatches) == 0:
                print("\n Test for mismatched {} Targets found (for same ConceptCode): Passed".format(coltitle))
            else:
                print("\n Test for mismatched {} Target text found (for same ConceptCode): Failed".format(coltitle))
                table = PrettyTable(["Concept Code", "Mismatched {} Targets".format(coltitle), "Mismatched Row Number"])
                table.border = True

                for x in mismatches:
                    for items in mismatches[x]:
                        table.add_row([x, ",".join((str(y) for y in items[0])), items[1]])
                print(table)

        # Validate Targets Syntax
        validate_target_format(32, "Direct")
        validate_target_format(33, "Indirect")        

        # Perform Validation Direct_Targets, Indirect_Targets column against Concept Code
        crossvalidate_with_concept_code(32, "Direct")
        crossvalidate_with_concept_code(33, "Indirect")


    def find_similar_text(wks, args):
        """ Uses similarity.py module to find similar text
            * We remove all duplicates first
            * We store row location of all indicator text
            * We apply NLP similarity on deduped list
            * In fast mode, we use sorted list and start comparing from the location
              of the first string.
            * In deep mode, we search every sentence with other
        """
        col = utils.get_column(wks, 5)
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
    
        def finddups(colnumber, categoryname):

            mismatches = utils.finddups(wks, colnumber)
        
            if len(mismatches) > 0:
                print("\n\n Test for duplicate values for {} found: Failed".format(categoryname))
                print("\n ------------")
                print("\n {} ID -> list of (Mismatched Values, Mismatched Rows)".format(categoryname))
                print("\n ------------")
                pprint.pprint(mismatches)
            else:
                print("\n\n Test for duplicate values for {} found: Passed".format(categoryname))
                

        finddups(0, "SDG Goal")
        finddups(1, "SDG Target")

     
    # Build source dictionary for Targets and                
    worksheet_name = "BIA to SDG mapping"
    wks = worksheets[worksheet_name]
    print('\n\n{0:-^60}\n'.format('Validate worksheet: {}'.format(worksheet_name)))
    validate_bia_sdg_mapping_sheet(wks)

    worksheet_name = "SDG Compass Metrics"
    print('\n\n{0:-^60}\n'.format('Validate worksheet: {}'.format(worksheet_name)))
    wks = worksheets[worksheet_name]
    validate_sdg_compass_metrics_sheet(wks)
    
    print('\n\n{0:-^60}\n'.format('Finding similar Indicator value candidates'))
    metricwks = worksheets["SDG Compass Metrics"]
    find_similar_text(metricwks, args)


def buildgraph(sheet):
    """Build a python graph representation of data"""

    def graph(): return defaultdict(graph)
    def graph_to_dict(graph): return {k: graph_to_dict(t[k]) for k in t}

    wks = sheet.worksheet("BIA to SDG mapping")
    all_values = wks.get_all_values()
    
    concept_graph = graph()
    
    for n, rows in enumerate(all_values):
        concept_dict[rows[1]]
    

def download_and_remove_title(sheet):
    """
        Download entire worksheets data so we minimize calls to Google API
        This is to prevent getting dinged from Google for excessive API usage
    """
    # Setup how many rows are occupied by title
    # For processing, we igore these rows

    wks_ignore_title_count = {}
    wks_ignore_title_count["BIA to SDG mapping"] = 2
    wks_ignore_title_count["BIA to SDG Target Mapping"] = 2
    wks_ignore_title_count["SDG Goal"] = 2
    wks_ignore_title_count["SDG Compass Metrics"] = 1

    worksheet_dict = {}
    for wks in sheet.worksheets():
        wks_data = wks.get_all_values()
        if wks.title in wks_ignore_title_count:
            worksheet_dict[wks.title] = wks_data[wks_ignore_title_count[wks.title]:]
        else:
            worksheet_dict[wks.title] = wks_data

    return worksheet_dict


def main():
 
    # Define inputs to scripts
   
    parser = argparse.ArgumentParser(
        description="Tool to manipulate SDG spreadsheet such as validation, build graph links and sync sheets")

    parser.add_argument(
        "action",
        action = "store",
        choices = ["validate", "graph"],
        help = "Pass one of these action for script to perform"
    )

    parser.add_argument(
        "--live",
        action = "store_true",
        default = False,
        help = "If specified, use spreadsheet version shared by everyone"
    )
        
    parser.add_argument(
        "--deeply-similar",
        action = "store_true",
        default = False,
        help = "Perform exhaustive similarity search between text. Warning: Takes long time"
    )
    
    args = parser.parse_args()

    try:

        # Initialize Google and open Spreadsheet        
        initialize_google()
        sheet_name = "BIA V5 SDG Alignment as of 05-2018 WORKING DRAFT.xlsx" \
                     if args.live else "BIA V5 SDG Alignment"
        ssheet = open_spreadsheet(sheet_name)
        print("Using sheet: {}".format(sheet_name))

        # Cache GoogleSheet data to avoid bumping into API limits
        worksheet_dict = download_and_remove_title(ssheet)
     
        if args.action == "validate":
            validate(worksheet_dict, args)

    except Exception as ex:
        traceback.print_exc()
        print("Error while processing script: {}".format(ex))
        

if __name__ == "__main__":
  main()
