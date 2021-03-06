import pandas as pd
import numpy as np
import re
import string
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# -------------- INPUT PARAMETERS : Might be adjusted --------------
# There are no questions with more or less possible answers than this
MAX_ANSWERS = 6
MIN_ANSWERS = 3

# Threshold of similarity for string comparision with compare_text() function.
# Above it strings are considered close enough. Value is between 0 and 100.
# 100 - completely identical
# 99 - completely identical with punctuation symbols removed and case ignored
# Value below 99 will use similarity search and not exact match. This might take longer time on bigger files
# Set to 99 for fast work and set lower value for similarity check
SIMILARITY_THRESHOLD = 65

re_illegal_chars = re.compile('[\n\r\t]')
# input Excel file name
ifile = 'qst.xlsx'
# output Excel file name
ofile = 'tmp.xlsx'

# ---- These logical vars determine performed processing ----
# Check for illegal characters in the answers. These characters are defined in  RegEx re_illegal_chars
do_check_illegal_chr = False
# Check for duplicate questions. The result is printed to console
do_check_duplicates = False
# Delete excess duplicate questions in dialog mode
do_delete_duplicates = True
# Delete 4-digit QID questions - these questions will be skipped while parsing input file
do_delete_4d_qid = True
# Split questions to separate DataFrame by BRIEFTEXT and store each in separate sheet of Excel file
# If split questions is False main DataFrame will be stored to a single sheet of an Excel file
do_split_by_brieftext = False
# -------------- END OF INPUT PARAMETERS --------------


# When parsing input Excel file each retrieved question with it's answers is stored to an instance
# of this class. These instances are stored to a list named questions
class Question:
    def __init__(self, nmb_arg=-1, qid_arg=-1, qst_arg='', ans_args={}, bt_arg='', status_arg='', docId_arg=-1):
        self.__nmb = nmb_arg
        self.__qid = qid_arg
        self.__question = qst_arg
        self.__answers = ans_args
        self.__brieftext = bt_arg
        self.__status = status_arg
        self.__docId = docId_arg

    def get_nmb(self):
        return self.__nmb
    nmb = property(get_nmb)

    def get_qid(self):
        return self.__qid
    qid = property(get_qid)

    def get_docId(self):
        return self.__docId
    docId = property(get_docId)

    def get_question(self):
        return self.__question
    question = property(get_question)

    def get_answers(self):
        return self.__answers
    answers = property(get_answers)

    def get_brieftext(self):
        return self.__brieftext
    brieftext = property(get_brieftext)

    def get_status(self):
        return self.__status
    status = property(get_status)

    # Check the answers for illegal characters, TODO: empty text, duplicate text
    def check_answers_ic(self):
        result = 'OK'
        for a_txt in self.__answers:
            if re_illegal_chars.search(a_txt) is not None:
                # print('ANS:', self.__answers)
                result = 'FAIL: Illegal characters found in answers'
                break
        return result

    # convert instance of Question to string
    def __str__(self):
        str_result = 'NMB={0}\tQID={1}\t{2}\tDocID={3}\tStatus={4}\r\n'\
            .format(self.__nmb, self.__qid, self.__brieftext, self.__docId, self.__status)
        str_result += 'Въпрос: {0}\r\n'.format(self.__question)
        ans_index = 1
        for answer in self.__answers:
            str_result += '   {0}: {1} \t\tSCR:{2}\r\n'.format(ans_index, answer, self.__answers[answer])
            ans_index += 1
        return str_result

    # represent instance of Question as string
    def __repr__(self):
        str_result = 'NMB={0}\tQID={1}\t{2}\tDocID={3}\tStatus={4}\r\n'\
            .format(self.__nmb, self.__qid, self.__brieftext, self.__docId, self.__status)
        str_result += 'Въпрос: {0}\r\n'.format(self.__question)
        ans_index = 1
        for answer in self.__answers:
            str_result += '   {0}: {1} \t\tSCR:{2}\r\n'.format(ans_index, answer, self.__answers[answer])
            ans_index += 1
        return str_result


# Returns the level of similarity from 0 to 100.
# If strings are exactly same returned value is 100
# Returns 99 when strings match with case ignored and punctuation symbols removed
# Lower values are result of similarity comparision
def compare_text(t1='', t2=''):
    if t1 == t2:
        return 100

    # remove punctuation symbols, convert to lowercase, compare
    text1 = ''.join([word for word in t1 if word not in string.punctuation])
    text1 = text1.lower()
    text2 = ''.join([word for word in t2 if word not in string.punctuation])
    text2 = text2.lower()
    if text1 == text2:
        return 99

    # speed up comparision if similarity check is not required
    if SIMILARITY_THRESHOLD >= 99:
        return 0

    vectorizer = TfidfVectorizer()
    matrix = vectorizer.fit_transform([text1, text2])
    cs = cosine_similarity(matrix[0:1], matrix)
    scs = str(cs[0, 1])  # convert to string and remove the first character
    fcs = float(scs) * 100  # convert to float
    ics = int(fcs)  # convert to integer
    return ics


# Read Excel file and put data into a DataFrame
# only the first Sheet is processed
xl = pd.ExcelFile(ifile)
print('Parsing file {0}...'.format(ifile))
print('There are {0} sheets:'.format(len(xl.sheet_names)))
print(xl.sheet_names)
print('Reading data from ' + xl.sheet_names[0])
input_data_frame = xl.parse(xl.sheet_names[0])
print('Rows in DataFrame:', len(input_data_frame.index))
# check if there is a second sheet named "Класификатор" and parse it
qualifier = {}
row = 0
if len(xl.sheet_names) >= 2:
    if xl.sheet_names[1] == 'Класификатор':
        qualifier_data_frame = xl.parse(xl.sheet_names[1])
        print("Прочетен е 'Класификатор'")
        for qlf in qualifier_data_frame['DocID']:
            qlf_txt = qualifier_data_frame['Документен класификатор'][row]
            row += 1
            # print('DocID={0}  Text:{1}'.format(int(qlf), qlf_txt))
            qualifier[int(qlf)] = qlf_txt
else:
    print("WARNING: No second sheet named 'Класификатор' found!")

# Parse questions from DataFrame and put them in a list of Questions objects
questions = []
# retrieve all BRIEFTEXT values and store to set used later to split questions to separate DataFrames.
bt_set = set()
row = 0
skipped = 0
for qst in input_data_frame['QID']:
    # check if QID column is a number. In this case the row contains a question text
    if not pd.isna(qst):
        # print('NMB=', df1['NMB'][row])
        nmb = int(input_data_frame['NMB'][row])
        # print('QID=', df1['QID'][row])
        qid = int(input_data_frame['QID'][row])
        # print('QUESTION: ', df1['QUESTION/ANSWER'][row])
        status = input_data_frame['Status'][row]
        docId = int(input_data_frame['DocID'][row])
        bt = input_data_frame['BRIEFTEXT'][row]
        bt = bt.strip('.')
        bt_set.add(bt)
        qst_text = input_data_frame['QUESTION/ANSWER'][row]
        qst_text = qst_text.strip(" \n\r?:.")
        # retrieve the answers
        a = 1
        answers = {}
        while a < MAX_ANSWERS:
            if (row+a) >= len(input_data_frame.index):
                break
            if pd.isna(input_data_frame['SCR'][row + a]):
                break
            answer_text = input_data_frame['QUESTION/ANSWER'][row + a]
            answer_text = answer_text.strip(" \n\r?:.")
            answer_score = input_data_frame['SCR'][row + a]
            # check for duplicate answers
            if answer_text in answers:
                print('WARNING! Duplicate answer in question NMB={0} QID={1}'.format(nmb, qid))
            answers[answer_text] = int(answer_score)
            a += 1
        # exclude 4-digit QID questions
        if do_delete_4d_qid:
            if qid < 10000:
                skipped += 1
                row += 1
                continue
        # CHECK: for too many answers
        if len(answers) == MAX_ANSWERS:
            print('WARNING: Answers number is at MAX in question NMB={0} QID={1}'.format(nmb, qid))
            fail = True
        # Check: for min answers
        if len(answers) < MIN_ANSWERS:
            print('WARNING: Answers number is below MIN in question NMB={0} QID={1}'.format(nmb, qid))
            fail = True
        q = Question(nmb, qid, qst_text, answers, bt, status, docId)
        questions.append(q)
    row += 1
print('Parsed {0} questions in the file. Skipped {1} with 4-digit QIDs'.format(len(questions), skipped))
print('BRIEFTEXT categories found:')
print(bt_set)
# sort questions[] by QID
questions.sort(key=lambda x: x.qid)

# ----------- CHECKS ------------
if do_check_illegal_chr:
    for qst in questions:
        # CHECK: for illegal characters
        if qst.check_answers_ic() != 'OK':
            print(qst.check_answers_ic())
            print(qst)
            print()

# CHECK: for duplicates, output to console
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
            if compare_text(q1_bt + ' ' + q1_txt, q2_bt + ' ' + q2_txt) >= SIMILARITY_THRESHOLD:
                dup_counter += 1
                duplicates[qst.nmb] = qst.qid
                duplicates[qst2.nmb] = qst2.qid
                dup[qst.nmb] = qst.qid
                dup[qst2.nmb] = qst2.qid
        if len(dup) != 0:
            print('DUP {0}: {1}'.format(dup_counter-1, dup))

# Delete excess duplicate questions in dialog mode
to_delete_qid = []
if do_delete_duplicates:
    duplicates = []
    dup_counter = 0
    end_deletion = False
    print('\r\n\r\n------------------------------------------------')
    print('Изберете въпроси които да бъдат изтрити. Изборът става с въвеждане на индексите им разделени с интервал.')
    print('Въведете END за приключване на работа преди да е достигнат края на списъка с дублирани въпроси.')
    input('Потвърдете с Enter.')
    for qst in questions:
        # Skip if already detected
        if qst.nmb in duplicates:
            continue
        # duplicates of the current question
        dup = {}
        dup_index = 0
        for qst2 in questions:
            # Skip self check
            if qst.nmb == qst2.nmb:
                continue
            # if qst.question.lower() == qst2.question.lower():
            q1_txt = qst.question.lower()
            q2_txt = qst2.question.lower()
            q1_bt = qst.brieftext
            q2_bt = qst.brieftext
            if compare_text(q1_bt + ' ' + q1_txt, q2_bt + ' ' + q2_txt) >= SIMILARITY_THRESHOLD:
                if qst.nmb not in duplicates:
                    duplicates.append(qst.nmb)
                    dup[dup_index] = qst
                    dup_index += 1
                if qst2.nmb not in duplicates:
                    duplicates.append(qst2.nmb)
                    dup[dup_index] = qst2
                    dup_index += 1
        if len(dup) != 0:
            dup_counter += 1
            print('\r\n\r\n------------------------------------------------')
            print('Count={0}'.format(dup_counter))
            n = 0
            for n in range(len(dup)):
                print('ВЪПРОС {0}:'.format(n+1))
                if dup[n].docId in qualifier:
                    print(qualifier[dup[n].docId])
                print(dup[n])
            # ^([1-9]\s)+[1-9]?$|^[1-9]{1}$|^$
            # - RegEx matches single digits 1 to 9 separate by space or only one single digit or an empty string
            re_check_user_input = re.compile('^([1-{0}]\\s)+[1-{0}]?$|^[1-{0}]'.format(n + 1) + '{1}$|^$')
            while True:
                user_selection = input('Изберете за изтриване, END за край:')
                if user_selection.strip().upper() == 'END':
                    end_deletion = True
                    break
                if re_check_user_input.match(user_selection):
                    break
                else:
                    print('ERROR: Invalid user input. Try again!')
            if end_deletion:
                break
            if user_selection != '':
                selections = user_selection.split()
                for selected in selections:
                    qst = dup[int(selected)-1]
                    print('Delete QID={0}'.format(qst.qid))
                    to_delete_qid.append(qst.qid)
            else:
                print('Nothing to delete.')
# split all questions from questions[] into separate DataFrames.
# save each DataFrame to separate Sheet in an Excel file named ofile
# each sheet is named after BRIEFTEXT field
writer = pd.ExcelWriter(ofile)
print('QID to delete:')
print(to_delete_qid)
if do_split_by_brieftext:
    print('Writing to output file ' + ofile + ' on separate sheets. Working...', end='')
    for bt in bt_set:
        row = 0
        df = pd.DataFrame({'NMB': '', 'QID': '', 'BRIEFTEXT': '',
                           'QUESTION/ANSWER': '', 'SCR': '', 'DocID': '', 'Status': ''}, index=[0])
        valid_qst_counter = 0
        for qst in questions:
            if qst.brieftext != bt:
                continue
            if qst.qid in to_delete_qid:
                continue
            valid_qst_counter += 1
            # adding a row with question
            df.loc[row] = [qst.nmb, qst.qid, bt, qst.question, np.nan, qst.docId, qst.status]
            row += 1
            # adding rows with answers
            for ans in qst.answers:
                df.loc[row] = [np.nan, np.nan, np.nan, ans, qst.answers[ans], np.nan, np.nan]
                row += 1
        # do not add empty sheet
        if valid_qst_counter != 0:
            # slash sign is illegal in sheet names
            sheet = bt.replace('/', '-')
            df.to_excel(writer, sheet)
            print('.', end='')
# save to the output file on single sheet
else:
    # export all questions to single DataFrame
    print('Writing to output file ' + ofile + ' on single sheet. Working...', end='')
    row = 0
    df = pd.DataFrame({'NMB': '', 'QID': '', 'BRIEFTEXT': '',
                       'QUESTION/ANSWER': '', 'SCR': '', 'DocID': '', 'Status': ''}, index=[0])
    for qst in questions:
        if qst.qid in to_delete_qid:
            print('Skip saving QID={0}'.format(qst.qid))
            continue
        # adding a row with question
        df.loc[row] = [qst.nmb, qst.qid, qst.brieftext, qst.question, np.nan, qst.docId, qst.status]
        row += 1
        # adding rows with answers
        for ans in qst.answers:
            df.loc[row] = [np.nan, np.nan, np.nan, ans, qst.answers[ans], np.nan, np.nan]
            row += 1
    df.to_excel(writer, 'Sheet')

while True:
    try:
        writer.save()
        break
    except:
        print()
        print('ERROR: Failed writing to {0}. Maybe open by other program.'.format(ofile))
        input('To try again press Enter. ')

print('\r\nDONE!')
