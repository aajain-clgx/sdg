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
try:
    import xlsxwriter
except:
    print("Please install xlsxwriter module")
    print("pip3 install xlsxwriter")
    sys.exit(1)

import similarity
import utils
from collections import defaultdict, OrderedDict
import traceback
import pprint
import datetime

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


def validate(worksheets, args, report, ignore_title_count):
    """Perform various cross validations on worksheets 
        in a spreadsheet and generate output a report"""
 
    def validate_bia_sdg_mapping_sheet(wks, reportsheet_dict, title_count, wks_target_mapping):
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
       
        def validate_target_format(colnumber, coltitle, rd, title_count):
            """Helper function to validate the syntax of a target specified in 
               Direct or Indirect Target column
               The format should be like digit.digit or digit.alpha
            """

            mismatches = utils.validate_target_format(wks, colnumber)
            
            if len(mismatches) == 0:
                print("\n Test invalid syntax for {} Targets: Passed. Good Job!".format(coltitle))
                rd["sheet"].write(rd["row"], 0, "Test invalid syntax for {} Targets".format(coltitle))
                rd["sheet"].write(rd["row"], 1, "Passed. Good Job!", rd["green"])
                rd["row"]+=1
                rd["sheet"].write(rd["row"], 0, None)
                rd["row"]+=1
            else:
                print("\n Test invalid syntax for {} Targets: Failed".format(coltitle))
                rd["sheet"].write(rd["row"], 0, "Test invalid syntax for {} Targets".format(coltitle))
                rd["sheet"].write(rd["row"], 1, "Failed", rd["red"])
                rd["row"]+=1
                rd["sheet"].write(rd["row"], 0, None)
                rd["row"]+=1

                table = PrettyTable(["Row Number", "Invalid Syntax {} Targets".format(coltitle)])
                table.border = True
                rd["sheet"].write_row(rd["row"], 0, tuple(["Row Number", "Invalid Syntax {} Targets".format(coltitle)]), rd["bold"])
                rd["row"]+=1

                for x in mismatches:
                    # Fix Row count by adding number of rows used for title
                    table.add_row([x+title_count, ",".join("'{}'".format(str(y)) for y in mismatches[x])])
                    rd["sheet"].write_row(rd["row"], 0, tuple([x+title_count, ",".join("'{}'".format(str(y)) for y in mismatches[x])]))
                    rd["row"]+=1
                rd["sheet"].write(rd["row"], 0, None)
                rd["row"]+=1
                rd["sheet"].write(rd["row"], 0, None)
                rd["row"]+=1
                print(table)

        def crossvalidate_with_concept_code(colnumber, coltitle, rd, title_count):
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
                print("\n Test for mismatched {} Targets found (for same ConceptCode): Passed. Good Job!".format(coltitle))
                rd["sheet"].write(rd["row"], 0, "Test for mismatched {} Targets found (for same ConceptCode)".format(coltitle))
                rd["sheet"].write(rd["row"], 1, "Passed. Good Job!", rd["green"])
                rd["row"]+=1

            else:
                print("\n Test for mismatched {} Target text found (for same ConceptCode): Failed".format(coltitle))
                table = PrettyTable(["Concept Code", "Mismatched {} Targets".format(coltitle), "Mismatched Row Number"])
                table.border = True
                rd["sheet"].write(rd["row"], 0, "Test for mismatched {} Targets found (for same ConceptCode)".format(coltitle))
                rd["sheet"].write(rd["row"], 1, "Failed", rd["red"])
                rd["row"]+=1
                rd["sheet"].write(rd["row"], 0, None)
                rd["row"]+=1
                rd["sheet"].write_row(rd["row"], 0, tuple(["Concept Code", "Mismatched {} Targets".format(coltitle), "Mismatched Row Number"]), rd["bold"])
                rd["row"]+=1

                for x in mismatches:
                    for items in mismatches[x]:
                        # Fix row count by adding number of rows used for title
                        items[1]+=title_count
                        table.add_row([x, ",".join((str(y) for y in items[0])), items[1]])
                        rd["sheet"].write_row(rd["row"], 0, tuple([x, ",".join((str(y) for y in items[0])), items[1]]))
                        rd["row"]+=1
                print(table)
                rd["sheet"].write(rd["row"], 0, None)
                rd["row"]+=1
                rd["sheet"].write(rd["row"], 0, None)
                rd["row"]+=1

        def missing_concept_code_from_target_sheet(rd):

            concept_code = set(utils.get_column(wks, 0))
            concept_code_from_target_sheet = set(utils.get_column(wks_target_mapping, 0))

            missing_in_sheet_1 = concept_code_from_target_sheet - concept_code
            missing_in_sheet_2 = concept_code - concept_code_from_target_sheet


            if len(missing_in_sheet_1)  ==  0:
                print("\n Test for Concept Code missing from 'BIA to SDG Target Mapping' worksheet: Passed. Good Job!")
                rd["sheet"].write(rd["row"], 0, "Test for Concept Code missing from 'BIA to SDG Target Mapping' worksheet")
                rd["sheet"].write(rd["row"], 1, "Passed. Good Job!", rd["green"])
                rd["row"]+=1
                rd["sheet"].write(rd["row"], 0, None)
                rd["row"]+=1
            else:
                print("\nTest for Concept Code missing from 'BIA to SDG Target Mapping' worksheet: Failed")
                rd["sheet"].write(rd["row"], 0, "Test for Concept Code missing from 'BIA to SDG Target Mapping' worksheet")
                rd["sheet"].write(rd["row"], 1, "Failed", rd["red"])
                rd["row"]+=1
                rd["sheet"].write(rd["row"], 0, None)
                rd["row"]+=1

                table = PrettyTable(["Concept Codes"])
                table.border = True
                rd["sheet"].write_row(rd["row"], 0, tuple(["Concept Code"]), rd["bold"])
                rd["row"]+=1

                for x in missing_in_sheet_1:
                    # Fix Row count by adding number of rows used for title
                    table.add_row([x])
                    rd["sheet"].write_row(rd["row"], 0, tuple([x]))
                    rd["row"]+=1
                rd["sheet"].write(rd["row"], 0, None)
                rd["row"]+=1
                rd["sheet"].write(rd["row"], 0, None)
                rd["row"]+=1
                print(table)


            if len(missing_in_sheet_2)  ==  0:
                print("\n Test for Concept Code from this sheet but missing in 'BIA to SDG Target Mapping' worksheet: Passed. Good Job!")
                rd["sheet"].write(rd["row"], 0, "Test for Concept Code from this sheet but missing in 'BIA to SDG Target Mapping' worksheet")
                rd["sheet"].write(rd["row"], 1, "Passed. Good Job!", rd["green"])
                rd["row"]+=1
                rd["sheet"].write(rd["row"], 0, None)
                rd["row"]+=1
            else:
                print("\nTest for Concept Code from this sheet but missing in 'BIA to SDG Target Mapping' worksheet: Failed")
                rd["sheet"].write(rd["row"], 0, "Test for Concept Code from this sheet but missing in 'BIA to SDG Target Mapping' worksheet")
                rd["sheet"].write(rd["row"], 1, "Failed", rd["red"])
                rd["row"]+=1
                rd["sheet"].write(rd["row"], 0, None)
                rd["row"]+=1

                table = PrettyTable(["Concept Codes"])
                table.border = True
                rd["sheet"].write_row(rd["row"], 0, tuple(["Concept Code"]), rd["bold"])
                rd["row"]+=1

                for x in missing_in_sheet_2:
                    # Fix Row count by adding number of rows used for title
                    table.add_row([x])
                    rd["sheet"].write_row(rd["row"], 0, tuple([x]))
                    rd["row"]+=1
                rd["sheet"].write(rd["row"], 0, None)
                rd["row"]+=1
                rd["sheet"].write(rd["row"], 0, None)
                rd["row"]+=1
                print(table)

       
        # Create Validation Report Worksheet
        rwks = report.add_worksheet("BIA to SDG mapping")
        bold = report.add_format({'bold': True})
        bg_green = report.add_format({'bold': True, 'bg_color': "#c5eac6"})
        bg_red = report.add_format({'bold': True, 'bg_color': "#efc6c6"})
        report_dict = {"sheet": rwks, "row": 3, "bold": bold, "green": bg_green, "red": bg_red}
        report_dict["sheet"].write(0, 0, "Test Status")
        report_dict["sheet"].write(1, 0, None)
        report_dict["sheet"].write(2, 0, None)    

        # Validate Targets Syntax
        validate_target_format(33, "Direct", report_dict, title_count)
        validate_target_format(34, "Indirect", report_dict, title_count)

        # Perform Validation Direct_Targets, Indirect_Targets column against Concept Code
        crossvalidate_with_concept_code(33, "Direct", report_dict, title_count)
        crossvalidate_with_concept_code(34, "Indirect", report_dict, title_count)
        
        # Missing concept codes from target sheet
        missing_concept_code_from_target_sheet(report_dict)


    def find_similar_text(wks, args, rd, title_count):
        """ Uses similarity.py module to find similar text
            * We remove all duplicates first
            * We store row location of all indicator text
            * We apply NLP similarity on deduped list
            * In fast mode, we use sorted list and start comparing from the location
              of the first string.
            * In deep mode, we search every sentence with other
        """
        col = utils.get_column(wks, 5)
        qacol = utils.get_column(wks, 10)
        col_line_dict = defaultdict(list)
        for n,x in enumerate(col):
            col_line_dict[x].append(n+1+title_count)

        colunique = list(set(col))
        colunique.sort()

        all_similar_lines = []
        header_written = False

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
                    
                    # Filter by column "Indicator QA Status". If a value exists remove it from
                    # the list
                        
                    qacoltest = []
                    qacoltest.extend(col_line_dict[val1])
                    qacoltest.extend(col_line_dict[val2])
                    #for x in qacoltest:
                    #    print(qacol[x-1-title_count])
                    allqacol_filled = all([qacol[x-1-title_count] in ('Complete', 'Needs Review') for x in qacoltest])
                    #print(allqacol_filled)
                    
                    # If this flag is set, do not filter by QA column
                    if args.all_similar:
                        allqacol_filled = False                    

                    if not(allqacol_filled):
                        all_similar_lines.append(similar_lines)
                        #print(qacoltest)
                        #print("ZZZZ===")
                        #print(similar_lines)
                        #print("=====ZZZZ")

                    if len(all_similar_lines) > 0 and not header_written:
                        rd["sheet"].write(rd["row"], 0, "Test for similar indicators values")
                        rd["sheet"].write(rd["row"], 1, "Failed", rd["red"])
                        rd["row"]+=1
                        rd["sheet"].write(rd["row"], 0, None)
                        rd["row"]+=1
                        rd["sheet"].write_row(rd["row"], 0, tuple(["Rows 1", "Similar Text 1", 
                                        "Rows 2", "Similar Text 2", "Similarity Score"]), rd["bold"])
                        rd["row"]+=1
                        header_written = True
                    if not(allqacol_filled):
                        rd["sheet"].write_row(rd["row"], 0, tuple([
                                ','.join((str(s) for s in similar_lines[0])), "'{}'".format(similar_lines[1]), 
                                ','.join((str(s) for s in similar_lines[2])), "'{}'".format(similar_lines[3]),
                                '{:.3f}'.format(similar_val)]))
                        rd["row"]+=1

                        print("\n Row: {}\n'{}'\n----\n Row: {}\n'{}'\nSimilarity Score = {:.3f}\n\n\n ====".format(
                            similar_lines[0], similar_lines[1],
                            similar_lines[2], similar_lines[3], similar_val))
                else:
                    if not args.deeply_similar:
                        break

                if len(all_similar_lines) == 0:
                    print("\n\n Test for similar indicators values: Passed. Good Job!")
                    rd["sheet"].write(rd["row"], 0, "Test for similar indicators values")
                    rd["sheet"].write(rd["row"], 1, "Passed. Good Job!", rd["green"])
                    rd["row"]+=1


    def validate_sdg_compass_indicators_sheet(wks, report, title_count):
        """Perform validate on 'SDG Compass Indicator sheet
           * Finds duplicate SDG Goals
           * Finds duplicate SDG Target
        """
    
        def finddups(colnumber, categoryname, rd, title_count):

            mismatches = utils.finddups(wks, colnumber)
        
            if len(mismatches) > 0:
                print("\n\n Test for duplicate values for {} found: Failed".format(categoryname))
                print("\n ------------")
                print("\n {} ID -> list of (Mismatched Values, Mismatched Rows)".format(categoryname))
                print("\n ------------")
                # Fix row count
                for k in mismatches:
                    for j in mismatches[k]:
                        j[1] += title_count
                pprint.pprint(mismatches)
                rd["sheet"].write(rd["row"], 0, "Test for duplicate values for {} found".format(categoryname))
                rd["sheet"].write(rd["row"], 1, "Failed", rd["red"])
                rd["row"]+=1
                rd["sheet"].write(rd["row"], 0, None)
                rd["row"]+=1
                rd["sheet"].write_row(rd["row"], 0, tuple(["{} ID".format(categoryname), "Mismatched Value", "Mismatched Row Number"]), rd["bold"])
                rd["row"]+=1
                for k in mismatches:
                    for j in mismatches[k]:
                        rd["sheet"].write_row(rd["row"], 0, tuple([k, "'{}'".format(j[0]), j[1]]))
                        rd["row"]+=1
                rd["sheet"].write(rd["row"], 0, None)
                rd["row"]+=1
                rd["sheet"].write(rd["row"], 0, None)
                rd["row"]+=1
            else:
                print("\n\n Test for duplicate values for {} found: Passed. Good Job!".format(categoryname))
                rd["sheet"].write(rd["row"], 0, "Test for duplicate values for {} found".format(categoryname))
                rd["sheet"].write(rd["row"], 1, "Passed. Good Job!", rd["green"])
                rd["row"]+=1

                
        # Create Reporting worksheet
        rwks = report.add_worksheet("SDG Compass Indicators")
        bold = report.add_format({'bold': True})
        bg_green = report.add_format({'bold': True, 'bg_color': "#c5eac6"})
        bg_red = report.add_format({'bold': True, 'bg_color': "#efc6c6"})
        report_dict = {"sheet": rwks, "row": 3, "bold": bold, "green": bg_green, "red": bg_red}
        report_dict["sheet"].write(0, 0, "Test Status")
        report_dict["sheet"].write(1, 0, None)
        report_dict["sheet"].write(2, 0, None)    
        
        # Perform Validation
        finddups(1, "SDG Goal", report_dict, title_count)
        finddups(2, "SDG Target", report_dict, title_count)
        
        # Perform similarity checks for Indicator column
        print('\n\n{0:-^60}\n'.format('Finding similar Indicator value candidates'))
        find_similar_text(wks, args, report_dict, title_count)

    
    def validate_sdg_target(wks, report, title_count):
        
        # Create Reporting worksheet
        rwks = report.add_worksheet("SDG Targets")
        bold = report.add_format({'bold': True})
        bg_green = report.add_format({'bold': True, 'bg_color': "#c5eac6"})
        bg_red = report.add_format({'bold': True, 'bg_color': "#efc6c6"})
        report_dict = {"sheet": rwks, "row": 3, "bold": bold, "green": bg_green, "red": bg_red}
        report_dict["sheet"].write(0, 0, "Test Status")
        report_dict["sheet"].write(1, 0, None)
        report_dict["sheet"].write(2, 0, None)    
        
        # Perform Validation
        utils.build_finddups_report(wks, 1, "SDG Goals", report_dict, title_count)
        utils.build_finddups_report(wks, 2, "SDG Target", report_dict, title_count)


    def build_business_to_indicator_map(wks, report, title_count):

        # Create Reporting worksheet
        rwks = report.add_worksheet("Business Theme Mapping")
        bold = report.add_format({'bold': True})
        bg_green = report.add_format({'bold': True, 'bg_color': "#c5eac6"})
        bg_red = report.add_format({'bold': True, 'bg_color': "#efc6c6"})
        rd = {"sheet": rwks, "row": 3, "bold": bold, "green": bg_green, "red": bg_red}
        rd["sheet"].write(0, 0, "Mapping of Business Themes to CESR Indicator, SDG Targets")
        rd["sheet"].write(1, 0, None)
        rd["sheet"].write(2, 0, None)

        business_theme_map = {}

        for n, row in enumerate(wks):
            theme = row[4]
            cesr_indicator = row[5]
            target = row[2]
            
            if theme in business_theme_map:
                if cesr_indicator in business_theme_map[theme]:
                    business_theme_map[theme][cesr_indicator].add(target)
                else:
                    business_theme_map[theme][cesr_indicator] = {target}
            else:
                business_theme_map[theme] = {}
                business_theme_map[theme][cesr_indicator] = {target}
                
        
        rd["sheet"].write_row(rd["row"], 0, tuple(["Business Theme", "CESR Indicator", "SDG Target"]), rd["bold"])
        rd["row"]+=1
        sorttheme_map = OrderedDict(sorted(business_theme_map.items(), key=lambda t: t[0]))
        for theme in sorttheme_map:
            for indicator in business_theme_map[theme]:
                targets = list(business_theme_map[theme][indicator])
                targets.sort()
                for val in targets:
                    rd["sheet"].write_row(rd["row"], 0, tuple(["{}".format(theme), "{}".format(indicator), "{}".format(val)]))
                    rd["row"]+=1
        rd["sheet"].write(rd["row"], 0, None)
        rd["row"]+=1
        rd["sheet"].write(rd["row"], 0, None)
        rd["row"]+=1

     
    # Build source dictionary for Targets and                
    worksheet_name = "BIA to SDG mapping"
    wks = worksheets[worksheet_name]
    wks_target_mapping = worksheets["BIA to SDG Target Mapping"]
    title_count = ignore_title_count[worksheet_name]
    print('\n\n{0:-^60}\n'.format('Validate worksheet: {}'.format(worksheet_name)))
    validate_bia_sdg_mapping_sheet(wks, report, title_count, wks_target_mapping)

    worksheet_name = "SDG Compass Indicators"
    print('\n\n{0:-^60}\n'.format('Validate worksheet: {}'.format(worksheet_name)))
    wks = worksheets[worksheet_name]
    title_count = ignore_title_count[worksheet_name]
    validate_sdg_compass_indicators_sheet(wks, report, title_count)

    worksheet_name = "SDG Targets"
    print('\n\n{0:-^60}\n'.format('Validate worksheet: {}'.format(worksheet_name)))
    wks = worksheets[worksheet_name]
    title_count = ignore_title_count[worksheet_name]
    validate_sdg_target(wks, report, title_count)


    worksheet_name = "SDG Compass Indicators"
    print('\n\n{0:-^60}\n'.format('Build Business Theme -> (Indicator, Target) Map in worksheet: {}'.format(worksheet_name)))
    wks = worksheets[worksheet_name]
    title_count = ignore_title_count[worksheet_name]
    build_business_to_indicator_map(wks, report, title_count)


def sync(writesheet, worksheets, title_count, direct_column):
    """Build a python graph representation of data"""

    def graph(): return defaultdict(graph)
    def graph_to_dict(graph): return {k: graph_to_dict(t[k]) for k in t}


    def build_target_map():

        target_map = {}
        wks = worksheets["SDG Targets"]
        
        target_col = utils.get_column(wks, 2)
        for cell in target_col:
            strsplit = cell.split()
            target_map[strsplit[0].strip()] = cell
        
        return target_map

    def sync_table(target_map):

        update_cells = []
        target_text_map = build_target_map()
        colnumber = 33 #if direct_column else 34
        valid_target_dict = utils.get_valid_target_map(worksheets["BIA to SDG mapping"], colnumber)
        write_row_num = title_count
        for row in worksheets["BIA to SDG Target Mapping"]:
            
            writetarget = row[0]
        
            targets = []
            if writetarget in valid_target_dict:
                targets = valid_target_dict[writetarget]
            padding = [''] * (20 - len(targets))
            targets.extend(padding)
            
            write_row_num += 1
        
            for n in range(len(targets)):
                
                if targets[n] == '':
                    update_cells.append(gspread.Cell(write_row_num, 12 + n, ''))
                else:
                    update_cells.append(gspread.Cell(write_row_num, 12 + n, target_text_map[targets[n]]))
                    
        # Update all cells
        writesheet.update_cells(update_cells)


    tmap = build_target_map()
    sync_table(tmap)
    

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
    wks_ignore_title_count["SDG Targets"] = 1
    wks_ignore_title_count["SDG Compass Indicators"] = 1

    worksheet_dict = {}
    for wks in sheet.worksheets():
        wks_data = wks.get_all_values()
        if wks.title in wks_ignore_title_count:
            worksheet_dict[wks.title] = wks_data[wks_ignore_title_count[wks.title]:]
        else:
            worksheet_dict[wks.title] = wks_data

    return worksheet_dict, wks_ignore_title_count


def main():
 
    # Define inputs to scripts
   
    parser = argparse.ArgumentParser(
        description="Tool to manipulate SDG spreadsheet such as validation, build graph links and sync sheets")

    parser.add_argument(
        "action",
        action = "store",
        choices = ["validate", "sync"],
        help = "Pass one of these action for script to perform"
    )

    parser.add_argument(
        "--all-similar",
        action = "store_true",
        default = False,
        help = "If specified, similar indicator test does not filter based on QA column"
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
        sheet_name = "BIA V5 SDG Alignment as of 05-2018 WORKING DRAFT.xlsx"
        ssheet = open_spreadsheet(sheet_name)
        print("Using sheet: {}".format(sheet_name))

        # Cache GoogleSheet data to avoid bumping into API limits
        worksheet_dict, ignore_title_count = download_and_remove_title(ssheet)
     
        if args.action == "validate":
            # Create validation report
            validation_report = xlsxwriter.Workbook("Validation Report - {}.xlsx".format(datetime.date.today().isoformat()))
            validate(worksheet_dict, args, validation_report, ignore_title_count)
            validation_report.close()
        elif args.action == "sync":
            sync(ssheet.worksheet("BIA to SDG Target Mapping"), worksheet_dict, ignore_title_count["BIA to SDG Target Mapping"], True)
    

    except Exception as ex:
        traceback.print_exc()
        print("Error while processing script: {}".format(ex))
        

if __name__ == "__main__":
  main()
