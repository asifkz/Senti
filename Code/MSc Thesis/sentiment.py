from __future__ import print_function

import xlwt
import xlrd

from nltk.classify import DecisionTreeClassifier

from nltk.classify import NaiveBayesClassifier
from nltk.classify import MaxentClassifier
from nltk.classify import PositiveNaiveBayesClassifier
from nltk.classify import WekaClassifier
from nltk.classify import SklearnClassifier


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

mystop_words=[
'i', 'me', 'my', 'myself', 'we', 'our', 'ours', 'ourselves', 'you', 'your', 'yours',
'yourself', 'yourselves', 'he', 'him', 'his', 'himself', 'she', 'her', 'hers',
'herself', 'it', 'its', 'itself', 'they', 'them', 'their', 'theirs', 'themselves',
 'this', 'that', 'these', 'those', 'am', 'is', 'are', 'was', 'were', 'be', 'been', 'being', 
'have', 'has', 'had', 'having', 'do', 'does', 'did', 'doing', 'a', 'an', 'the', 
'and',  'if', 'or', 'as', 'until', 'while', 'of', 'at', 'by', 'for',   'between', 'into',
'through', 'during', 'to', 'from', 'in', 'out', 'on', 'off', 'then', 'once', 'here',
 'there',  'all', 'any', 'both', 'each', 'few', 'more',
 'other', 'some', 'such',  'than', 'too', 'very', 's', 't', 'can', 'will',  'don', 'should', 'now'
]

emodict={

"%-("	:  "NegativeSentiment",
"%-)"	:  "PositiveSentiment",
"(-:"	:  "PositiveSentiment",
"(:"	:  "PositiveSentiment",
"(^ ^)"	:  "PositiveSentiment",
"(^-^)"	:  "PositiveSentiment",
"(^.^)"	:  "PositiveSentiment",
"(^_^)"	:  "PositiveSentiment",
"(o:"	:  "PositiveSentiment",
"(o;"	:  "NeutralSentiment",
")-:"	:  "NegativeSentiment",
"):"	:  "NegativeSentiment",
")o:"	:  "NegativeSentiment",
"*)"	:  "NeutralSentiment",
"*\o/*"	:  "PositiveSentiment",
"--^--@":  "PositiveSentiment",
"0:)"	:  "PositiveSentiment",
"38*"	:  "NegativeSentiment",
"8)"	:  "PositiveSentiment",
"8-)"	:  "NeutralSentiment",
"8-0"	:  "NegativeSentiment",
"8/"	:  "NegativeSentiment",
#"8\"	:  "NegativeSentiment",
"8c"	:  "NegativeSentiment",
":#"	:  "NegativeSentiment",
":'("	:  "NegativeSentiment",
":'-("	:  "NegativeSentiment",
":("	:  "NegativeSentiment",
":)"	:  "PositiveSentiment",
":*("	:  "NegativeSentiment",
":,("	:  "NegativeSentiment",
":-&"	:  "NegativeSentiment",
":-("	:  "NegativeSentiment",
":-(o)"	:  "NegativeSentiment",
":-)"	:  "PositiveSentiment",
":-*"	:  "PositiveSentiment",
":-*"	:  "PositiveSentiment",
":-/"	:  "NegativeSentiment",
":-/"	:  "NeutralSentiment",
":-D"	:  "PositiveSentiment",
":-O"	:  "NeutralSentiment",
":-P"	:  "PositiveSentiment",
":-S"	:  "NegativeSentiment",
#":-\"	:  "NegativeSentiment",
#":-\"	:  "NeutralSentiment",
":-|"	:  "NegativeSentiment",
":-}"	:  "PositiveSentiment",
":/"	:  "NegativeSentiment",
":0->-<|:"	:  "NeutralSentiment",
":3"	:  "PositiveSentiment",
":9"	:  "PositiveSentiment",
":D"	:  "PositiveSentiment",
":E"	:  "NegativeSentiment",
":F"	:  "NegativeSentiment",
":O"	:  "NegativeSentiment",
":P"	:  "PositiveSentiment",
":P"	:  "PositiveSentiment",
":S"	:  "NegativeSentiment",
":X"	:  "PositiveSentiment",
":["	:  "NegativeSentiment",
":["	:  "NegativeSentiment",
#":\"	:  "NegativeSentiment",
":]"	:  "PositiveSentiment",
":_("	:  "NegativeSentiment",
":b)"	:  "PositiveSentiment",
":l"	:  "NeutralSentiment",
":o("	:  "NegativeSentiment",
":o)"	:  "PositiveSentiment",
":p"	:  "PositiveSentiment",
":s"	:  "NegativeSentiment",
"0:|"	:  "NegativeSentiment",
":|"	:  "NeutralSentiment",
":p"	:  "PositiveSentiment",
":("	:  "NegativeSentiment",
";)"	:  "NeutralSentiment",
";^)"	:  "PositiveSentiment",
";o)"	:  "NeutralSentiment",
"</3-1"	:  "NegativeSentiment",
"<3"	:  "PositiveSentiment",
"<:}"	:  "NeutralSentiment",
"<o<"	:  "NegativeSentiment",
">/"	:  "NegativeSentiment",
">:("	:  "NegativeSentiment",
">:)"	:  "PositiveSentiment",
">:D"	:  "PositiveSentiment",
">:L"	:  "NegativeSentiment",
">:O"	:  "NegativeSentiment",
">=D"	:  "PositiveSentiment",
">["	:  "NegativeSentiment",
#">\"	:  "NegativeSentiment",
">o>"	:  "NegativeSentiment",
"@}->--":  "PositiveSentiment",
"B("	:  "NegativeSentiment",
"Bc"	:  "NegativeSentiment",
"D:"	:  "NegativeSentiment",
"X("	:  "NegativeSentiment",
"X("	:  "NegativeSentiment",
"X-("	:  "NegativeSentiment",
"XD"	:  "PositiveSentiment",
"XD"	:  "PositiveSentiment",
"XO"	:  "NegativeSentiment",
"XP"	:  "NegativeSentiment",
"XP"	:  "PositiveSentiment",
"^_^"	:  "PositiveSentiment",
"^o)"	:  "NegativeSentiment",
"x3?"	:  "PositiveSentiment",
"xD"	:  "PositiveSentiment",
"xP"	:  "NegativeSentiment",
"|8C"	:  "NegativeSentiment",
"|8c"	:  "NegativeSentiment",
"|D"	:  "PositiveSentiment",
"}:)"	:  "PositiveSentiment",



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
# Load some categories from the training set

data_train=SentimentData()
data_test=SentimentData()


data_train.target=[i for i in range(750)]
data_train.data=[i for i in range(750)]

data_test.target=[i for i in range(100)]
data_test.data=[i for i in range(100)]

all_comments=[]
all_ratings=[]

for cell_num in range(0,750):
	comments=expand_contractions(s.cell(cell_num,0).value)
        comments=replace_all(comments,emodict)
	all_comments.append(comments)
	all_ratings.append(s.cell(cell_num,1).value)

results = []

for k in range (0,100):

    #print("Run# {}".format(k))
   
    r2=0
    m2=1000

    random_list=[i for i in range(750)] 
    shuffle(random_list)
    #print (random_list)

    train_data=[]
    test_data=[]
    x1=0
    y1=750
    for i in range(x1,y1):
	index=random_list[i]
            #print (index)
	index=int(index)
        
	data_train.data[i]=all_comments[index]
	data_train.target[i]=all_ratings[index]

            #train_data.append((xx,yy))
			

        #if i>=100:

            #index=random_list[i]
            #print (index)

            #xx=(s.cell(index,0).value)
            #yy=(s.cell(index,1).value)

	    #data_test.data[i-100]=xx	
 	    #data_test.target[i-100]=yy
            #test_data.append((xx,yy))
        #print (xx)
        #print (yy)
        #x=nltk.sentiment.util.demo_liu_hu_lexicon(xx, plot=False)

        #print(x)

    
        #Sheet1.write(i-x1,0,s.cell(i,0).value)
        #Sheet1.write(i-x1,1,x)
    #wb1.save('ar1.xls')

    #print (len(test_data))
    #print (len(train_data))

    def size_mb(docs):
    	return sum(len(s.encode('utf-8')) for s in docs) / 1e6

    #data_train_size_mb = size_mb(data_train.data)
    #data_test_size_mb = size_mb(data_test.data)

    #print("%d documents - %0.3fMB (training set)" % (
    	#len(data_train.data), data_train_size_mb))
    #print("%d documents - %0.3fMB (test set)" % (
    	#len(data_test.data), data_test_size_mb))
    #print("%d categories" % len(categories))
    #print()

    # split a training set and a test set
    #y_train, y_test = data_train.target, data_test.target

    #print("Extracting features from the training data using a sparse vectorizer")
    t0 = time()
    if opts.use_hashing:
	vectorizer = HashingVectorizer(stop_words=mystop_words, non_Negative=True, n_features=opts.n_features)
    	X_train = vectorizer.transform(data_train.data)
    else:
	vectorizer = TfidfVectorizer(sublinear_tf=True, max_df=0.5, stop_words=mystop_words)
    	X_train = vectorizer.fit_transform(data_train.data).toarray()


   

    X_test=X_train[650:750]	

    #print(len(X_test))
	
    X_train=X_train[0:650]	

    #print(len(X_train))	

    y_train=data_train.target[0:650]
    y_test = data_train.target[650:750]	










    #print("done in %fs at %0.3fMB/s" % (duration, data_train_size_mb / duration))
    #print("n_samples: %d, n_features: %d" % X_train.shape)
    #print()

    #print("Extracting features from the test data using the same vectorizer")
    t0 = time()
    #X_test = vectorizer.transform(data_test.data)
    #X_test = vectorizer.fit_transform(data_test.data).toarray()
    #vectorizer = TfidfVectorizer(sublinear_tf=True, max_df=0.5,
                                 #stop_words='english')
    #X_test = vectorizer.fit_transform(data_test.data).toarray()	
    #duration = time() - t0
    #print("done in %fs at %0.3fMB/s" % (duration, data_test_size_mb / duration))
    #print("n_samples: %d, n_features: %d" % X_test.shape)
    #print()

    # mapping from integer feature name to original token string
    if opts.use_hashing:
    	feature_names = None
    else:
    	feature_names = vectorizer.get_feature_names()

    if opts.select_chi2:
    	#print("Extracting %d best features by a chi-squared test" %
          	#opts.select_chi2)
        t0 = time()
        ch2 = SelectKBest(chi2, k=opts.select_chi2)
        X_train = ch2.fit_transform(X_train, y_train)
        X_test = ch2.transform(X_test)
        if feature_names:
        	# keep selected feature names
        	feature_names = [feature_names[i] for i
                         	in ch2.get_support(indices=True)]
        #print("done in %fs" % (time() - t0))
        #print()

    if feature_names:
    	feature_names = np.asarray(feature_names)


    def trim(s):
    	"""Trim string to fit on terminal (assuming 80-column display)"""
    	return s if len(s) <= 80 else s[:77] + "..."



###############################################################################
# Benchmark classifiers
    #print ("hello")	
    def benchmark(clf):
    	#print("hello")
    	#print('_' * 80)
    	#print("Training: ")
    	#print(clf)
    	t0 = time()

#####################################################################################

    	

	clf.fit(X_train, y_train)



##################################################################################
    	train_time = time() - t0
    	#print("train time: %0.3fs" % train_time)

    	t0 = time()
    	pred = clf.predict(X_test)
	#print(y_test)
	#print (pred)
    	test_time = time() - t0
    	#print("test time:  %0.3fs" % test_time)

    	score = metrics.accuracy_score(y_test, pred)
	#if(clf=="Ridge Classifier"):
    	#print("Accuracy: %0.3f" % score)
        print("%0.3f" % score)
    	if hasattr(clf, 'coef_'):
        	#print("dimensionality: %d" % clf.coef_.shape[1])
        	#print("density: %f" % density(clf.coef_))

        	if opts.print_top10 and feature_names is not None:
            		#print("top 10 keywords per class:")
            		for i, category in enumerate(categories):
                		top10 = np.argsort(clf.coef_[i])[-10:]
                		#print(trim("%s: %s"
                     			# % (category, " ".join(feature_names[top10]))))
        		#print()

    	#if opts.print_report:
        	#print("classification report:")
        	#print(metrics.classification_report(y_test, pred,
                                            #target_names=categories))

    	#if opts.print_cm:
        	#print("confusion matrix:")
        	#print(metrics.confusion_matrix(y_test, pred))

    	#print()
    	clf_descr = str(clf).split('(')[0]
    	return clf_descr, score, train_time, test_time


   
    #print("hello1")	


    #for clf, name in (
        #(RidgeClassifier(tol=1e-2, solver="lsqr"), "Ridge Classifier"),
        #(Perceptron(n_iter=50), "Perceptron"),
        #(PassiveAggressiveClassifier(n_iter=50), "Passive-Aggressive"),
        #(KNeighborsClassifier(n_neighbors=3), "kNN")):
        #(RandomForestClassifier(n_estimators=100), "Random forest")):
    	#print('=' * 80)
    	#print(name)
    	#results.append(benchmark(clf))



    #for penalty in ["l2", "l1"]:
    	#print('=' * 80)
    	#print("%s penalty" % penalty.upper())


    #results.append(benchmark(RidgeClassifier(tol=1e-2, solver="lsqr")))
    #results.append(benchmark(Perceptron(n_iter=50)))	
    #results.append(benchmark(PassiveAggressiveClassifier(n_iter=50)))
    #results.append(benchmark(KNeighborsClassifier(n_neighbors=3)))
    #results.append(benchmark(KNeighborsClassifier(n_neighbors=3)))
    #for penalty in ["l2", "l1"]:	
    	#results.append(benchmark(LinearSVC(loss='l2', penalty=penalty,dual=False, tol=1e-3)))


	
	


    #for penalty in ["l2", "l1"]:
    	#print('=' * 80)
    	#print("%s penalty" % penalty.upper())
    	# Train Liblinear model
    	#results.append(benchmark(LinearSVC(loss='l2', penalty=penalty,
                                            #dual=False, tol=1e-3)))

    	# Train SGD model
    	#results.append(benchmark(SGDClassifier(alpha=.0001, n_iter=50,
                                           #penalty=penalty)))

	# Train SGD with Elastic Net penalty
    #print('=' * 80)
    #print("Elastic-Net penalty")
    #results.append(benchmark(SGDClassifier(alpha=.0001, n_iter=50, penalty="elasticnet")))

    # Train NearestCentroid without threshold
    #print('=' * 80)
    #print("NearestCentroid (aka Rocchio classifier)")
    #results.append(benchmark(NearestCentroid()))

    # Train sparse Naive Bayes classifiers
    #print('=' * 80)
    #print("Naive Bayes")
    #results.append(benchmark(MultinomialNB(alpha=.01)))
    #results.append(benchmark(BernoulliNB(alpha=.01)))

    #print('=' * 80)
    #print("LinearSVC with L1-based feature selection")
    # The smaller C, the stronger the regularization.
    # The more regularization, the more sparsity.
    #results.append(benchmark(Pipeline([
		  #('feature_selection', LinearSVC(penalty="l1", dual=False, tol=1e-3)),
		  #('classification', LinearSVC())
		#])))

    # make some plots

    #indices = np.arange(len(results))

    #results = [[x[i] for x in results] for i in range(4)]

    #clf_names, score, training_time, test_time = results
    #training_time = np.array(training_time) / np.max(training_time)
    #test_time = np.array(test_time) / np.max(test_time)

    #plt.figure(figsize=(12, 8))
    #plt.title("Score")
    #plt.barh(indices, score, .2, label="score", color='r')
    #plt.barh(indices + .3, training_time, .2, label="training time", color='g')
    #plt.barh(indices + .6, test_time, .2, label="test time", color='b')
    #plt.yticks(())
    #plt.legend(loc='best')
    #plt.subplots_adjust(left=.25)
    #plt.subplots_adjust(top=.95)
    #plt.subplots_adjust(bottom=.05)





###############################
	
	
    #print (X_train[0])
    #X_train=X_train.toarray()

    #new_list = [[b for _,b in sub] for sub in X_train]
    #X_train=new_list	
    
    #y_train=y_train.todense()
    #X_test=X_test.toarray()
    #y_test=y_test.toarray()


    results.append(benchmark(GradientBoostingClassifier())[1])
    #results.append(benchmark(RandomForestClassifier())[1])
    #results.append(benchmark(DecisionTreeClassifier())[1])
    #results.append(benchmark(AdaBoostClassifier())[1])
    #results.append(benchmark(GradientBoostingRegressor()))
    #results.append(benchmark(SVC(kernel='rbf', probability=True))[1])

    #clf1 = DecisionTreeClassifier(max_depth=4)
    #clf2 = KNeighborsClassifier(n_neighbors=7)
    #clf3 = SVC(kernel='rbf', probability=True)
    #eclf = VotingClassifier(estimators=[('dt', clf1), ('knn', clf2), ('svc', clf3)], voting='soft', weights=[2,1,2])

    #clf1 = clf1.fit(X,y)
    #clf2 = clf2.fit(X,y)
    #clf3 = clf3.fit(X,y)
    #eclf = eclf.fit(X,y)
	

##########################

print("-------------------------")
print("Average accuracy: {}".format(mean(results)))
print("-------------------------")


    #for i, c in zip(indices, clf_names):
    	#plt.text(-.3, i, c)

    #plt.show()
	
