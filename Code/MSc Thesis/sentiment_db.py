from __future__ import print_function

import xlwt
import xlrd

from nltk.classify import DecisionTreeClassifier

from nltk.classify import NaiveBayesClassifier
from nltk.classify import MaxentClassifier
from nltk.classify import PositiveNaiveBayesClassifier
from nltk.classify import WekaClassifier
from nltk.classify import SklearnClassifier

import mysql.connector
import random
import re

from xlwt import Workbook
from xlrd import open_workbook
from random import shuffle
from statistics import mean


import logging
import numpy as np
from optparse import OptionParser
import sys
###################################
reload(sys)  
sys.setdefaultencoding('utf8')

#################################################


from time import time
import matplotlib.pyplot as plt

from sklearn.datasets import fetch_20newsgroups
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.feature_extraction.text import HashingVectorizer
from sklearn.feature_selection import SelectKBest, chi2
from sklearn.linear_model import RidgeClassifier
from sklearn.pipeline import Pipeline
from sklearn.svm import LinearSVC
from sklearn.linear_model import SGDClassifier
from sklearn.linear_model import Perceptron
from sklearn.linear_model import PassiveAggressiveClassifier
from sklearn.naive_bayes import BernoulliNB, MultinomialNB
from sklearn.neighbors import KNeighborsClassifier
from sklearn.neighbors import NearestCentroid
from sklearn.ensemble import RandomForestClassifier
from sklearn.ensemble import GradientBoostingClassifier
from sklearn.utils.extmath import density
from sklearn import metrics
from sklearn.ensemble import AdaBoostClassifier
from sklearn.ensemble import GradientBoostingRegressor
from sklearn.tree import DecisionTreeClassifier

from sklearn.svm import SVC
#from sklearn.ensemble import VotingClassifier

class SentimentData:
	def __init__(self):
		self.data = []
		self.target = []

def replace_all(text, dic):
    for i, j in dic.iteritems():
        text = text.replace(i, j)
    return text




emodict={

"%-("	:  "Negative",
"%-)"	:  "Positive",
"(-:"	:  "Positive",
"(:"	:  "Positive",
"(^ ^)"	:  "Positive",
"(^-^)"	:  "Positive",
"(^.^)"	:  "Positive",
"(^_^)"	:  "Positive",
"(o:"	:  "Positive",
"(o;"	:  "Neutral",
")-:"	:  "Negative",
"):"	:  "Negative",
")o:"	:  "Negative",
"*)"	:  "Neutral",
"*\o/*"	:  "Positive",
"--^--@"	:  "Positive",
"0:)"	:  "Positive",
"38*"	:  "Negative",
"8)"	:  "Positive",
"8-)"	:  "Neutral",
"8-0"	:  "Negative",
"8/"	:  "Negative",
#"8\"	:  "Negative",
"8c"	:  "Negative",
":#"	:  "Negative",
":'("	:  "Negative",
":'-("	:  "Negative",
":("	:  "Negative",
":)"	:  "Positive",
":*("	:  "Negative",
":,("	:  "Negative",
":-&"	:  "Negative",
":-("	:  "Negative",
":-(o)"	:  "Negative",
":-)"	:  "Positive",
":-*"	:  "Positive",
":-*"	:  "Positive",
":-/"	:  "Negative",
":-/"	:  "Neutral",
":-D"	:  "Positive",
":-O"	:  "Neutral",
":-P"	:  "Positive",
":-S"	:  "Negative",
#":-\"	:  "Negative",
#":-\"	:  "Neutral",
":-|"	:  "Negative",
":-}"	:  "Positive",
":/"	:  "Negative",
":0->-<|:"	:  "Neutral",
":3"	:  "Positive",
":9"	:  "Positive",
":D"	:  "Positive",
":E"	:  "Negative",
":F"	:  "Negative",
":O"	:  "Negative",
":P"	:  "Positive",
":P"	:  "Positive",
":S"	:  "Negative",
":X"	:  "Positive",
":["	:  "Negative",
":["	:  "Negative",
#":\"	:  "Negative",
":]"	:  "Positive",
":_("	:  "Negative",
":b)"	:  "Positive",
":l"	:  "Neutral",
":o("	:  "Negative",
":o)"	:  "Positive",
":p"	:  "Positive",
":s"	:  "Negative",
"0:|"	:  "Negative",
":|"	:  "Neutral",
":p"	:  "Positive",
":("	:  "Negative",
";)"	:  "Neutral",
";^)"	:  "Positive",
";o)"	:  "Neutral",
"</3-1"	:  "Negative",
"<3"	:  "Positive",
"<:}"	:  "Neutral",
"<o<"	:  "Negative",
">/"	:  "Negative",
">:("	:  "Negative",
">:)"	:  "Positive",
">:D"	:  "Positive",
">:L"	:  "Negative",
">:O"	:  "Negative",
">=D"	:  "Positive",
">["	:  "Negative",
#">\"	:  "Negative",
">o>"	:  "Negative",
"@}->--"	:  "Positive",
"B("	:  "Negative",
"Bc"	:  "Negative",
"D:"	:  "Negative",
"X("	:  "Negative",
"X("	:  "Negative",
"X-("	:  "Negative",
"XD"	:  "Positive",
"XD"	:  "Positive",
"XO"	:  "Negative",
"XP"	:  "Negative",
"XP"	:  "Positive",
"^_^"	:  "Positive",
"^o)"	:  "Negative",
"x3?"	:  "Positive",
"xD"	:  "Positive",
"xP"	:  "Negative",
"|8C"	:  "Negative",
"|8c"	:  "Negative",
"|D"	:  "Positive",
"}:)"	:  "Positive",



}

contractions_dict = { 
"ain't": "am not",
"aren't": "are not",
"can't": "cannot",
"can't've": "cannot have",
"'cause": "because",
"could've": "could have",
"couldn't": "could not",
"couldn't've": "could not have",
"didn't": "did not",
"doesn't": "does not",
"don't": "do not",
"hadn't": "had not",
"hadn't've": "had not have",
"hasn't": "has not",
"haven't": "have not",
"he'd": "he would",
"he'd've": "he would have",
"he'll": "he will",
"he'll've": "he will have",
"he's": "he is",
"how'd": "how did",
"how'd'y": "how do you",
"how'll": "how will",
"how's": "how is",
"i'd": "i would",
"i'd've": "i would have",
"i'll": "i will",
"i'll've": "i will have",
"i'm": "i am",
"i've": "i have",
"isn't": "is not",
"it'd": "it would",
"it'd've": "it would have",
"it'll": "it will",
"it'll've": "it will have",
"it's": "it is",
"let's": "let us",
"ma'am": "madam",
"mayn't": "may not",
"might've": "might have",
"mightn't": "might not",
"mightn't've": "might not have",
"must've": "must have",
"mustn't": "must not",
"mustn't've": "must not have",
"needn't": "need not",
"needn't've": "need not have",
"o'clock": "of the clock",
"oughtn't": "ought not",
"oughtn't've": "ought not have",
"shan't": "shall not",
"sha'n't": "shall not",
"shan't've": "shall not have",
"she'd": "she would",
"she'd've": "she would have",
"she'll": "she will",
"she'll've": "she will have",
"she's": "she has",
"should've": "should have",
"shouldn't": "should not",
"shouldn't've": "should not have",
"so've": "so have",
"so's": "so is",
"that'd": "that would",
"that'd've": "that would have",
"that's": "that is",
"there'd": "there would",
"there'd've": "there would have",
"there's": "there is",
"they'd": "they would",
"they'd've": "they would have",
"they'll": "they will",
"they'll've": "they will have",
"they're": "they are",
"they've": "they have",
"to've": "to have",
"wasn't": "was not",
"we'd": "we would",
"we'd've": "we would have",
"we'll": "we will",
"we'll've": "we will have",
"we're": "we are",
"we've": "we have",
"weren't": "were not",
"what'll": "what will",
"what'll've": "what will have",
"what're": "what are",
"what's": "what is",
"what've": "what have",
"when's": "when is",
"when've": "when have",
"where'd": "where did",
"where's": "where is",
"where've": "where have",
"who'll": "who will",
"who'll've": "who will have",
"who's": "who is",
"who've": "who have",
"why's": "why is",
"why've": "why have",
"will've": "will have",
"won't": "will not",
"won't've": "will not have",
"would've": "would have",
"wouldn't": "would not",
"wouldn't've": "would not have",
"y'all": "you all",
"y'all'd": "you all would",
"y'all'd've": "you all would have",
"y'all're": "you all are",
"y'all've": "you all have",
"you'd": "you would",
"you'd've": "you would have",
"you'll": "you will",
"you'll've": "you will have",
"you're": "you are",
"you've": "you have"
}

contractions_regex = re.compile('(%s)' % '|'.join(contractions_dict.keys()))

def expand_contractions(s, contractions_dict=contractions_dict):
     def replace(match):
         return contractions_dict[match.group(0)]
     return contractions_regex.sub(replace, s.lower())


wb = open_workbook("a.xlsx")
s = wb.sheet_by_index(0)

#wb2 = open_workbook("rand.xls")
#s2 = wb2.sheet_by_index(0)



#wb1 = Workbook()
#Sheet1 = wb1.add_sheet('Sheet1')


# Display progress logs on stdout
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s %(levelname)s %(message)s')


# parse commandline arguments
op = OptionParser()
op.add_option("--report",
              action="store_true", dest="print_report",
              help="Print a detailed classification report.")
op.add_option("--chi2_select",
              action="store", type="int", dest="select_chi2",
              help="Select some number of features using a chi-squared test")
op.add_option("--confusion_matrix",
              action="store_true", dest="print_cm",
              help="Print the confusion matrix.")
op.add_option("--top10",
              action="store_true", dest="print_top10",
              help="Print ten most discriminative terms per class"
                   " for every classifier.")
op.add_option("--all_categories",
              action="store_true", dest="all_categories",
              help="Whether to use all categories or not.")
op.add_option("--use_hashing",
              action="store_true",
              help="Use a hashing vectorizer.")
op.add_option("--n_features",
              action="store", type=int, default=2 ** 16,
              help="n_features when using the hashing vectorizer.")
op.add_option("--filtered",
              action="store_true",
              help="Remove newsgroup information that is easily overfit: "
                   "headers, signatures, and quoting.")

(opts, args) = op.parse_args()
if len(args) > 0:
    op.error("this script takes no arguments.")
    sys.exit(1)

#print(__doc__)
#op.print_help()
#print()


###############################################################################
# Loading Project data from database



data_train=SentimentData()

data_train.target=[i for i in range(750)]
data_train.data=[i for i in range(750)]


all_comments=[]
all_ratings=[]

for cell_num in range(0,750):
	comments=expand_contractions(s.cell(cell_num,0).value)
        comments=replace_all(comments,emodict)
	all_comments.append(comments)
	all_ratings.append(s.cell(cell_num,1).value)

print("Started building classifier")

random_list=[i for i in range(750)] 
shuffle(random_list)

for i in range(0,750):
	index=random_list[i]
	index=int(index)
        
	data_train.data[i]=all_comments[index]
	data_train.target[i]=all_ratings[index]


t0 = time()
if opts.use_hashing:
	vectorizer = HashingVectorizer(stop_words='english', non_negative=True, n_features=opts.n_features)
    	X_train = vectorizer.transform(data_train.data)
else:
	vectorizer = TfidfVectorizer(sublinear_tf=True, max_df=0.5, stop_words='english')
    	X_train = vectorizer.fit_transform(data_train.data).toarray()

y_train=data_train.target[0:750]

if opts.use_hashing:
	feature_names = None
else:
	feature_names = vectorizer.get_feature_names()

if opts.select_chi2:
	t0 = time()
	ch2 = SelectKBest(chi2, k=opts.select_chi2)
	X_train = ch2.fit_transform(X_train, y_train)
	X_test = ch2.transform(X_test)
	if feature_names:
	# keep selected feature names
		feature_names = [feature_names[i] for i
                 	in ch2.get_support(indices=True)]

if feature_names:
	feature_names = np.asarray(feature_names)

#Are we using GradientBoostingClassifier?
classifier=GradientBoostingClassifier()

    	
# Training classifier
classifier.fit(X_train, y_train)
print("Finished training classifier")


print("Fetching comments from database..")

comments_text=[]
comment_id=[]
db=mysql.connector.connect(user='root', password='',
                              host='localhost',
                              database='Thesis')


#db = MySQLdb.connect(host="localhost",    # your host, usually localhost
#                    user="root",         # your username
#                    passwd="",  # your password
#                    db="Sentiment")        # name of the data base

# you must create a Cursor object. It will let
#  you execute all the queries you need
cur = db.cursor()

# Use all the SQL you like
cur.execute("SELECT * FROM inline_comments")


# print all the first cell of all the rows

count1=0

for row in cur.fetchall():
    	comments_text.append(replace_all(expand_contractions(row[10]),emodict).decode('utf-8','ignore').encode("utf-8"))
	comment_id.append(row[0])
        count1=count1+1
        #if count1==5000:
		#break	
    


comments_count=len(comments_text)

print("Found {} comments".format(comments_count))

for x in range(0,comments_count):
	
	X_test = vectorizer.transform([comments_text[x]]).toarray()
       
	pred = classifier.predict(X_test)

	#print(comments_text[x])	
	#print(pred)
	if x%1000 == 0:
		print("Finished classifying: {}".format(x))

	cur.execute("""UPDATE inline_comments SET sentiment_score=%s WHERE comment_id=%s""",(int(pred[0]),comment_id[x])) 

	
db.commit()
db.close()




###################################################################################
