#Displaying frequencies of words which is equal to or greater than 0.001 in descending order 
#in an excel file
#Iona Amanda Dsouza - 121100235
#4th March 2022 

import docx
import re
import os
import pyexcel 

def analyze(docfile):
    #reading the docx file
    doc = docx.Document(docfile)
    d = {}
    word_count = 0
    result = {}
    
    for para in doc.paragraphs:
        data = para.text       
        if data:
            line = data.strip(" ").split(" ")
            for words in line:
                word =re.sub('[^A-Za-z0-9]+', '', words)
                word = word.lower()
                word_count = word_count + 1
                
                #checking if the word in there in the dictionary
                if word in d.keys():
                    d[word] = d[word]+1
                else:
                    d[word] = 1 
    #calculating word frequency                
    for key,value in d.items(): 
        word_freq = d[key] / word_count              
        if(word_freq >= 0.001): 
            result.update({key:word_freq})
     
    #arranging the dictionary in descending order according to the frequencies  
    result =dict(sorted(result.items(),key = lambda x:x[1],reverse = True))
    result = list(result.items())
    #print(result)

    #adding data to excel file
    docname = os.path.basename(docfile)
    file_name = os.path.splitext(docname)[0] + "_word_stats.xlsx"
    pyexcel.save_as(array = result,dest_file_name = file_name,sheet_name = "Word Frequency Stats")
    
 
#analyze("pride_and_prejudice.docx")    