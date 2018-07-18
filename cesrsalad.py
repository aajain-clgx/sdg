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
    import spacy
except:
    print("Please install spacy module")
    print("pip3 install spacy")
    sys.exit(1)
from collections import defaultdict
import traceback
from sklearn.feature_extraction.text import TfidfVectorizer

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
    """Perform various cross validations and output a report"""
 
    def find_similar_text(col):
        nlp = spacy.load('en')
        similaritydict = {}
        count = 0

        vect = TfidfVectorizer(min_df=1)
        tfidf = vect.fit_transform(col)
        result = tfidf * tfidf.T
        print(result)

        for n1, val1 in enumerate(col):
        
            if n1 < 50:
                continue;
            count+=1
            l = []
            d1 = nlp(val1)
    
            for n2, val2 in enumerate(col):
                if n1 == n2:
                    continue
                d2 = nlp(val2)
                similarity = d1.similarity(d2)
                if similarity > 0.89:
                    l.append([ n2, val2, similarity])
            
            #l.sort(key=lambda x: x[4])
            similaritydict[n1] = l
            if count > 3:
                break
            

        x = PrettyTable(["Indicator Description Row", "Indicator Description Value", "Similar Row", "Similar Row Value", "Similarity Value"])
        for k in similaritydict:
            similar_lines = similaritydict[k]
            for line in similar_lines:
                x.add_row([k, col[k], line[0], line[1], line[2]])
        print(x)  
                    
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

    #print('\n\n{0:-^60}\n'.format('Validate similarity lines'))
    #metricwks = spreadsheet.worksheet("SDG Compass Metrics")
    #find_similar_text(metricwks.col_values(6))    

        

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
