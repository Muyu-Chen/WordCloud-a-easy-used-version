# -*- coding: UTF-8 -*-
from wordcloud import WordCloud
#import matplotlib.pyplot as plt
import numpy as np
#from PIL import Image
import jieba
from jieba.analyse import *
import pandas as pd
import numpy as np

'''
It's same to 'README.md', but there just an English version
program target: 
1. read a determined column of form (xlsx) with sentence
2. change the format to 'array'
3. delete the sapce (Chinses only) and line breaks in the 'array' 
4. (Chinese only) depart diffierent words(if you donot use this, you can delete those codes)
5. delete the stopping words (if you use English, you should find a stopping words list by yourself,
I just provide some simple words like 'a', 'an', 'the'.
6. draw a wordscloud by 'wordcloud' and save it as a file 'WordsCloudOutPut.jpg' 

other notes:
1. YOU HAVE TO CHANGE 'stopwordlist.txt' AND DELETE 'dictionary.txt' IF YOU DO NOT USE CHINESE !
2. if there are anything wrong with the path, please change all relative path with absolute path.
3. you should creat a folder named 'Output' before running this program.
3. if this program helps you, please give me a star!  
                                ! PLEASE !
'''

strF='./data/data.xlsx' #The location of your form file.
filename = "./Output/Words.txt"#The sentences after the process of deletting the sapce (Chinses only) and line breaks.
outfilename = "./Output/WordsCutOut.txt"#The words after deletting stopping words. Also, it's a Intermediate file.
font=r'C:\Windows\Fonts\Deng.ttf' #you can choose a font you perfer in this folder
#if you do not use chinese, you should change the font (like 'C:\Windows\Fonts\times.ttf')
#beacause you maybe don't have this font in your computer.

#changing the format to 'array'
def ToArray(str1,cols): 
    #'str1' means the path of the file, and 'cols' means the column which you want to use
    #Attention: column 'A' in EXCEL equals 0)
    FunctionRead=pd.read_excel(str1,usecols=cols,sheet_name='Sheet1')
    FunctionArray=np.array(FunctionRead.stack())
    print('column-' + str(cols[0]+1) + ' is finished')
    return FunctionArray

#creating stopping words list
jieba.load_userdict("./data/dictionary.txt") #Chinese only,but you cannot delete the code, you can just use a null file
#dictionary.txt means which words(in chinese) you want to remain. it not work in other language but chinese.

#reading the stopping words list
def stopwordslist():
    stopwords = [line.strip() for line in open('./data/stopwordlist.txt',encoding='UTF-8').readlines()]
    return stopwords

# depart the sencences(Chinese only) and delete stopping words
def seg_depart(sentence):
    sentence_depart = jieba.cut(sentence.strip())
    #if you use other languages, use the code below and delete the code before this sentence
    #sentence_depart = sentence
    stopwords = stopwordslist()
    outstr = ''
    for word in sentence_depart:
        if word not in stopwords:
            if word != '\t':
                outstr += word
                outstr += " "
    return outstr

WordsList=[]
CommentArray=ToArray(strF,[4])
#'[4]' means the column you choose to read
CommentList=CommentArray.tolist()

#deletting the line breaks in the 'array' 
for x in CommentArray:
    WordsList.append(x.replace("\n"," "))

#writing the sentence into file
Com=open('./Output/Words.txt','w',encoding='UTF-8')
for line in WordsList:
    Com.write(str(line)+'\n')
Com.close

#writing the final products into file
outputs = open(outfilename, 'w', encoding='UTF-8')
for line in WordsList:
    line_seg = seg_depart(line)
    outputs.write(line_seg + '\n')
outputs.close()

#getting the final data
with open(outfilename,'r', encoding='utf-8')as f:
    text=f.read()
#setting the parameters of wordcloud
wc=WordCloud(
    scale=4,
    font_path=font,
    max_words=200,
    margin=2,
    background_color='white',
    max_font_size=300,
    min_font_size=2,
    collocations=False,
    width=1600,height=1200
)
#making the wordcloud
wc.generate(text)
wc.to_file('WordsCloud.jpg')
#writing the frequency of words into file
FrequWrite = open('./Output/WordsFrequency.txt','w', encoding='utf-8')
for keyword, weight in extract_tags(text, topK=80,withWeight=True,allowPOS=()):
   print('%s %s' % (keyword, weight),file=FrequWrite)
