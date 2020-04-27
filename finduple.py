import pandas as pd
import re

# there are no questions with more possible answers than this
MAX_ANSWERS = 6
re_illegal_chars = re.compile('[\n\r\t]')

class Question:
    def __init__(self, nmb_arg=-1, qid_arg=-1, qst_arg='', ans_arg={}):
        self.__nmb = nmb_arg
        self.__qid = qid_arg
        self.__question = qst_arg
        self.__answers = ans_arg

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

    # Check the answers for illegal characters, empty text, duplicate text
    def check_answers_ic(self):
        result = 'OK'
        for a_txt in self.__answers:
            if re_illegal_chars.search(a_txt) is not None:
                #print('ANS:', self.__answers)
                result = 'FAIL: Illegal charracters found in answers'
                break
        return result


# Read Excel file and put data in to a DataFrame
xl = pd.ExcelFile('qst.xlsx')

# print(xl.sheet_names)
df1 = xl.parse(xl.sheet_names[0])
print('Rows in DataFrame:', len(df1.index))

# Parse questions from DataFrame and put them in a list of Questions objects
questions = []
row = 0
for qst in df1['QID']:
    # check if QID column is a number. In this case the row contains a question text
    if not pd.isna(qst):
        # print('NMB=', df1['NMB'][row])
        nmb = int(df1['NMB'][row])
        # print('QID=', df1['QID'][row])
        qid = int(df1['QID'][row])
        # print('QUESTION: ', df1['QUESTION/ANSWER'][row])
        qst_text = df1['QUESTION/ANSWER'][row]
        qst_text = qst_text.strip(" \n\r?:.")
        # retrieve the answers
        a = 1
        answers = {}
        while a < MAX_ANSWERS:
            if (row+a) >= len(df1.index):
                break
            if pd.isna(df1['SCR'][row+a]):
                break
            answer_text = df1['QUESTION/ANSWER'][row+a]
            answer_text = answer_text.strip(" \n\r?:.")
            answer_score = df1['SCR'][row+a]
            # check for duplicate answers
            if answer_text in answers:
                print('WARNING! Duplicate answer in question NMB={0} QID={1}'.format(nmb, qid))
            answers[answer_text]=int(answer_score)
            a += 1
        q = Question(nmb, qid, qst_text, answers)
        questions.append(q)
    row += 1
print('Parsed {0} questions in the file.\n\n'.format(len(questions)))

# CHECK: for illegal characters
for qst in questions:
    if qst.check_answers_ic() != 'OK':
        print(qst.nmb, ' ', qst.qid, ' ', qst.question)
        print(qst.check_answers_ic())
        print()

# CHECK: for duplicates
# dictionary of all duplicated questions NMB:QID
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
        if q1_txt == q2_txt:
            dup_counter += 1
            duplicates[qst.nmb] = qst.qid
            duplicates[qst2.nmb] = qst2.qid
            dup[qst.nmb] = qst.qid
            dup[qst2.nmb] = qst2.qid
    if len(dup) != 0:
        print('DUP {0}: {1}'.format(dup_counter-1, dup))