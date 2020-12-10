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
        
# Input are the occupation list, a word_vector embedding, and a dictionary indicating the impression score of each occupation.
# Output a excel file writing the metrics and impression score of each occupation.
def createOccupationExcelFile(occupation_list,model,word_vectors,occup_imp):
    wb = Workbook()
    # Create the work sheet
    occu_gen_score =  wb.add_sheet(model)
    row=1
    # Add the first row as the field name
    occu_gen_score.write(0,0,"Occupation")
    occu_gen_score.write(0,1,"Metrics_1")
    occu_gen_score.write(0,2,"Impression Score")
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
            if not (occupation in occup_imp.keys()):
                print(occupation, " has not impression score yet, update it in the occup_imp xls file")
            else:
                occu_gen_score.write(row,2,occup_imp[occupation])
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
    print(num)
    occup_imp = {}
    for row in range(1,num+1):
        occup_imp[sheet.cell_value(row,0).lower()]=sheet.cell_value(row,column)
    return occup_imp


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
    occup_imp = getDictFromExcel("occup_imp.xls",1)
    createOccupationExcelFile(occupation_list,modelName,word_vectors,occup_imp)
#
#addNewEmb("glove-wiki-gigaword-100")
#addNewEmb("glove-wiki-gigaword-300")
#addNewEmb("glove-twitter-100")
#
#print("Load google news begin:")
#word_vectors =KeyedVectors.load_word2vec_format(
#            '~/Desktop/HC/Junior1/CS325_Linguistis/GroupProject/GoogleNews-vectors-negative300.bin', binary=True
#        )
#print("Load model succesfully \n")
occupation_list = getListFromText("occupation_name.txt")
occup_imp = getDictFromExcel("occup_imp.xls",1)
#createOccupationExcelFile(occupation_list,"GoogleNews",word_vectors,occup_imp)

# Combine four Excel files together
wiki100 = getDictFromExcel("glove-wiki-gigaword-100.xls",1)
wiki300 = getDictFromExcel("glove-wiki-gigaword-300.xls",1)
twitter100 = getDictFromExcel("glove-twitter-100.xls",1)
GoogleNews = getDictFromExcel("GoogleNews.xls",1)
model_lists = [wiki100,wiki300,twitter100,GoogleNews]

wb = Workbook()
# Create the work sheet
combinedResult =  wb.add_sheet("Combined")
row=1
# Add the first row as the field name
combinedResult.write(0,0,"Occupation")
combinedResult.write(0,1,"wiki100")
combinedResult.write(0,2,"wiki300")
combinedResult.write(0,3,"twitter100")
combinedResult.write(0,4,"GoogleNews")
combinedResult.write(0,6,"impression")

# For each occupation name
for occupation in occupation_list:
    # skip to the next one if the occupation is not in the word vector
    #Write occupation
    combinedResult.write(row,0,str(occupation))
    # Write impression score
    if not (occupation in occup_imp.keys()):
        print(occupation, " has not impression score yet, update it in the occup_imp xls file")
    else:
        combinedResult.write(row,6,occup_imp[occupation])
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
combinedResult.write(0,5,"total items")
combinedResult.write(1,5,row-1)
wb.save("combined.xls")


## test analogies
#result = word_vectors.most_similar(positive=['woman', 'king'], negative=['man'])
#most_similar_key, similarity = result[0]  # look at the first match
#print(f"{most_similar_key}: {similarity:.4f}")
#
#result = word_vectors.most_similar(positive=['woman', 'intelligent'], negative=['man'])
#most_similar_key, similarity = result[0]  # look at the first match
#print(f"{most_similar_key}: {similarity:.4f}")
#
#result = word_vectors.most_similar(positive=['caucasian', 'lazy'], negative=['hispanic'])
#most_similar_key, similarity = result[0]  # look at the first match
#print(f"{most_similar_key}: {similarity:.4f}")
#
#result = word_vectors.most_similar(positive=['caucasian', 'housekeeper'], negative=['hispanic'])
#most_similar_key, similarity = result[0]  # look at the first match
#print(f"{most_similar_key}: {similarity:.4f}")

#
## test similarity between two words
#similarity = word_vectors.similarity('lawyer', 'man')
#print(similarity)
#
#similarity = word_vectors.similarity('lawyer', 'woman')
#print(similarity)
