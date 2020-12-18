# https://radimrehurek.com/gensim/models/keyedvectors.html
import gensim.downloader as api
import numpy as np
import xlwt
from xlwt import Workbook
import xlrd
from gensim.models.keyedvectors import KeyedVectors


# The precondition is that word is indeed in word_vector. Check it before using this function
# Input are the word embedding vectors and a word to be analyzed. The output is the projection on the gender axis.
def getGenderStereotype(word_vectors,word):
    # Obtain the word vectors
    v_word = word_vectors[word]
    v_he = word_vectors['he']
    v_she = word_vectors['she']
    v_gender = v_he - v_she
    # Calculate the scalar projection of 'word' on the gender axis.
    gender_metrics = np.dot(v_word,v_gender) / np.linalg.norm(v_gender)
    return gender_metrics.item()
        
# Input are the occupation list, a word_vector embedding.
# Output a excel file writing the metrics of each occupation.
def createOccupationExcelFile(occupation_list,model,word_vectors):
    wb = Workbook()
    # Create the work sheet
    occu_gen_score =  wb.add_sheet(model[:10])
    row=1
    # Add the first row as the field name
    occu_gen_score.write(0,0,"Occupation")
    occu_gen_score.write(0,1,"Metrics_1")
    # For each occupation name
    for occupation in occupation_list:
        # skip to the next one if the occupation is not in the word vector
        if not (occupation in word_vectors):
            print(occupation," is not in the embedding.")
            continue
        # otherwise write the occupation and go to the next row.
        else:
            occu_gen_score.write(row,0,str(occupation))
            occu_gen_score.write(row,1,getGenderStereotype(word_vectors,occupation))
            row +=1
    # Write the number of valid occupations
    occu_gen_score.write(0,5,"total items")
    occu_gen_score.write(1,5,row-1)
    wb.save(model+".xls")
    return

def getListFromText(filename):
    file = open(filename, "r")
    list = []
    for line in file:
        list.append(line.strip().lower())
    file.close()
    return list

# The excel file should contain exactly two columns where the first one is occupation and the second one is the score
def getDictFromExcel(filename,column):
    workbook = xlrd.open_workbook(filename)
    sheet = workbook.sheet_by_index(0)
    num = int(sheet.cell_value(1,5))
    scoreDict = {}
    for row in range(1,num+1):
        scoreDict[sheet.cell_value(row,0).lower()]=sheet.cell_value(row,column)
    return scoreDict


def captureBiasedOccupation(n,filename):
    occup_gender = getDictFromExcel(filename,1)
    man = sorted(occup_gender, key=occup_gender.get, reverse=True)[:n]
    woman = sorted(occup_gender, key=occup_gender.get, reverse=False)[:n]
    return man,woman


def addNewEmb(modelName):
    print("Load ", modelName, " begin:")
    word_vectors = api.load(modelName)
    print("Load model succesfully \n")
    occupation_list = getListFromText("occupation_name.txt")
    createOccupationExcelFile(occupation_list,modelName,word_vectors)
    
    
## Get embedding models
#addNewEmb("fasttext-wiki-news-subwords-300")
#addNewEmb("glove-wiki-gigaword-100")
#addNewEmb("glove-wiki-gigaword-300")
#addNewEmb("glove-twitter-100")
#addNewEmb("glove-twitter-200")
#
#print("Load google news begin:")
#word_vectors =KeyedVectors.load_word2vec_format(
#            '~/Desktop/HC/Junior1/CS325_Linguistis/GroupProject/GoogleNews-vectors-negative300.bin', binary=True
#        )
#print("Load model succesfully \n")
occupation_list = getListFromText("occupation_name.txt")
#createOccupationExcelFile(occupation_list,"GoogleNews",word_vectors)

# Combine Excel files together
wiki_news = getDictFromExcel("fasttext-wiki-news-subwords-300.xls",1)
wiki100 = getDictFromExcel("glove-wiki-gigaword-100.xls",1)
wiki300 = getDictFromExcel("glove-wiki-gigaword-300.xls",1)
twitter100 = getDictFromExcel("glove-twitter-100.xls",1)
twitter200 = getDictFromExcel("glove-twitter-200.xls",1)
GoogleNews = getDictFromExcel("GoogleNews.xls",1)
model_lists = [wiki_news,wiki100,wiki300,twitter100,twitter200,GoogleNews]
model_names = ["wiki_news","wiki100","wiki300","twitter100","twitter200","GoogleNews"]
wb = Workbook()

# Create the work sheet
combinedResult =  wb.add_sheet("Combined")
row=1
# Add the first row as the field name
combinedResult.write(0,0,"Occupation")
for i in range(0,len(model_lists)):
    combinedResult.write(0,i+1,str(model_names[i]))
# For each occupation name
for occupation in occupation_list:
    # skip to the next one if the occupation is not in the word vector
    # Write occupation
    combinedResult.write(row,0,str(occupation))
    # Write corresponding bias
    for i in range(0,len(model_lists)):
        model = model_lists[i]
        column = i+1
        if not (occupation in model.keys()):
            print(occupation," is not in the embedding.")
            continue
        # otherwise write the occupation and go to the next row.
        else:
            combinedResult.write(row,column,model[occupation])
    row +=1
# Write the number of valid occupations
combinedResult.write(0,7,"total items")
combinedResult.write(1,7,row-1)
wb.save("combined.xls")
