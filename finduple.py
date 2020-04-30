import pandas as pd
import numpy as np
import openpyxl.workbook
import re
import sys
import getopt

# there are no questions with more or less possible answers than this
MAX_ANSWERS = 6
MIN_ANSWERS = 3
re_illegal_chars = re.compile('[\n\r\t]')
# input Excel file name
ifile = 'qst.xlsx'
# output Excel file name
ofile = 'out.xlsx'

# --- These logical vars determine performed processing ---
# Check for illegal characters in the answers. These characters are defined in  RegEx re_illegal_chars
do_check_illegal_chr = False
# Check for duplicate questions. The result is printed to console
do_check_duplicates = False
# Delete duplicates from main DataFrame
do_delete_duplicates = False
# Split questions to separate DataFrame by BRIEFTEXT and store each in separate sheet of Excel file
# If split questions is False main DataFrame will be stored to a single sheet of an Excel file
do_split_by_brieftext = True


# When parsing input Excel file each retrieved question with it's answers is stored to an instance
# of this class. These instances are stored to a list named questions
class Question:
    def __init__(self, nmb_arg=-1, qid_arg=-1, qst_arg='', ans_args={}, bt_arg=''):
        self.__nmb = nmb_arg
        self.__qid = qid_arg
        self.__question = qst_arg
        self.__answers = ans_args
        self.__brieftext = bt_arg

    def get_nmb(self):
        return self.__nmb
    nmb = property(get_nmb)

    def get_qid(self):
        return self.__qid
    qid = property(get_qid)

    def get_question(self):
        return self.__question
    question = property(get_question)

    def get_answers(self):
        return self.__answers
    answers = property(get_answers)

    def get_brieftext(self):
        return self.__brieftext
    brieftext = property(get_brieftext)

    # Check the answers for illegal characters, empty text, duplicate text
    def check_answers_ic(self):
        result = 'OK'
        for a_txt in self.__answers:
            if re_illegal_chars.search(a_txt) is not None:
                # print('ANS:', self.__answers)
                result = 'FAIL: Illegal characters found in answers'
                break
        return result


# print the Question to console
def print_question(qst):
    print('NMB={0}\tQID={1}\t{2}'.format(qst.nmb, qst.qid, qst.brieftext))
    print('Question: ' + qst.question)
    n = 1
    for ans in qst.answers:
        print('   {0}: {1} \t\tSCR:{2}'.format(n, ans, qst.answers[ans]))
        n += 1
    print()
    return


# Read Excel file and put data in to a DataFrame
# only the first Sheet is processed
xl = pd.ExcelFile(ifile)
print('Parsing file {0}...'.format(ifile))
print('There are {0} sheets:'.format(len(xl.sheet_names)))
print(xl.sheet_names)
print('Reading data from ' + xl.sheet_names[0])
data_frame = xl.parse(xl.sheet_names[0])
print('Rows in DataFrame:', len(data_frame.index))

# Parse questions from DataFrame and put them in a list of Questions objects
questions = []
# retrieve all BRIEFTEXT values and store to set used later to split questions to separate DataFrames.
bt_set = set()
row = 0
for qst in data_frame['QID']:
    # check if QID column is a number. In this case the row contains a question text
    if not pd.isna(qst):
        # print('NMB=', df1['NMB'][row])
        nmb = int(data_frame['NMB'][row])
        # print('QID=', df1['QID'][row])
        qid = int(data_frame['QID'][row])
        # print('QUESTION: ', df1['QUESTION/ANSWER'][row])
        bt = data_frame['BRIEFTEXT'][row]
        bt_set.add(bt)
        qst_text = data_frame['QUESTION/ANSWER'][row]
        qst_text = qst_text.strip(" \n\r?:.")
        # retrieve the answers
        a = 1
        answers = {}
        while a < MAX_ANSWERS:
            if (row+a) >= len(data_frame.index):
                break
            if pd.isna(data_frame['SCR'][row + a]):
                break
            answer_text = data_frame['QUESTION/ANSWER'][row + a]
            answer_text = answer_text.strip(" \n\r?:.")
            answer_score = data_frame['SCR'][row + a]
            # check for duplicate answers
            if answer_text in answers:
                print('WARNING! Duplicate answer in question NMB={0} QID={1}'.format(nmb, qid))
            answers[answer_text] = int(answer_score)
            a += 1
        q = Question(nmb, qid, qst_text, answers, bt)
        questions.append(q)
    row += 1
print('Parsed {0} questions in the file.'.format(len(questions)))
print('BRIEFTEXT:')
print(bt_set)


# ----------- CHECKS ------------
if do_check_illegal_chr:
    for qst in questions:
        fail = False
        # CHECK: for illegal characters
        if qst.check_answers_ic() != 'OK':
            print(qst.check_answers_ic())
            fail = True
        # CHECK: for too many answers
        if len(qst.answers) == MAX_ANSWERS:
            print('WARNING: Answers number is at MAX!')
            fail = True
        # Check: for min answers
        if len(qst.answers) < MIN_ANSWERS:
            print('WARNING: Answers number is below MIN!')
            fail = True
        if fail:
            print_question(qst)
            print()

# CHECK: for duplicates
# dictionary of all duplicated questions NMB:QID
if do_check_duplicates:
    duplicates = {}
    dup_counter = 0
    for qst in questions:
        # Skip if already detected
        if qst.nmb in duplicates:
            continue
        # duplicates of the current question
        dup = {}
        for qst2 in questions:
            # Skip self check
            if qst.nmb == qst2.nmb:
                continue
            q1_txt = qst.question.lower()
            q2_txt = qst2.question.lower()
            q1_bt = qst.brieftext
            q2_bt = qst.brieftext
            if (q1_bt + q1_txt) == (q2_bt + q2_txt):
                dup_counter += 1
                duplicates[qst.nmb] = qst.qid
                duplicates[qst2.nmb] = qst2.qid
                dup[qst.nmb] = qst.qid
                dup[qst2.nmb] = qst2.qid
        if len(dup) != 0:
            print('DUP {0}: {1}'.format(dup_counter-1, dup))

# split all questions from questions[] into separate DataFrames.
# save each DataFrame to separate Sheet in an Excel file named ofile
# each sheet is named after BRIEFTEXT field
writer = pd.ExcelWriter(ofile)
if do_split_by_brieftext:
    for bt in bt_set:
        row = 0
        df = pd.DataFrame({'QID': '', 'BRIEFTEXT': '', 'QUESTION/ANSWER': '', 'SCR': ''}, index=[0])
        for qst in questions:
            if qst.brieftext != bt:
                continue
            df.loc[row] = [qst.qid, bt, qst.question, np.nan]
            print('row={0}'.format(row))
            row += 1
        sheet = bt.replace('/', '-')
        df.to_excel(writer, sheet)
# save main DataFrame to the output file
else:
    data_frame.to_excel(writer)
writer.save()
