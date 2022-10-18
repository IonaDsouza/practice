import docx
import pyexcel 
import re
import os
import string

def analyze(docfile):
    doc = docx.Document(docfile)
    d= {}
    freq = 0
    count =0
    tot = 0
    result ={}
    #doc_name = os.path.basename(docfile)
    #excel_file_name = os.path.splitext(doc_name)[0]+"_word_stats.xlsx"
    #punctuation = ".?:,\"';()\."
    #punctuation = "\".?:,\';!._-()"
    #punctuation = ".!"\',\"?:)(_\"\".-;'-"
    
    space =" "
    for para in doc.paragraphs:
        data = para.text
        #data = data.replace("-"," ")
        if data:
            temp = data.strip(" ").split(" ")
            for i in temp:
                #new = i.strip(string.punctuation).lower()
                #new = re.sub('[%s]' % re.escape(string.punctuation),space, i)
                new =re.sub(r'[^A-Za-z0-9]+', '', i)
                new = new.lower()
                #print(new)
                count = count+1

                if new in d.keys():
                    d[new] = d[new]+1
                else:
                        d[new] = 1 

    # dc = d.values()
    # count = sum(dc)

    for key,value in d.items(): 
        freq = d[key]/count 
    # for k,v in d.items():             
        if(freq>=0.001):
            tot = tot+1  
            result.update({key:freq})
    #result =sorted(result.items(),key = lambda item:item[1],reverse = True)
    #result = list(result.items())
    print(result)
    #print(count)
    print(tot)




    # doc_name = os.path.basename(docfile)
    # file_name = os.path.splitext(doc_name)[0]+"wordstats.xlsx"
    # pyexcel.save_as(array = result,dest_file_name=file_name,sheet_name ="Word stat")

 
analyze("pride_and_prejudice.docx")    