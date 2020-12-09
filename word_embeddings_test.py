# https://radimrehurek.com/gensim/models/keyedvectors.html
import gensim.downloader as api
import numpy as np
import xlwt
from xlwt import Workbook
import xlrd


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
def createOccupationExcelFile(occupation_list,word_vectors,occu_imp_Dict):
    wb = Workbook()
    # Create the work sheet
    occu_gen_score =  wb.add_sheet("Occupation_Gender_Score")
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
    wb.save("Occupation.xls")

def getListFromText(filename):
    file = open(filename, "r")
    list = []
    for line in file:
        list.append(line.strip())
    file.close()
    return list

# The excel file should contain exactly two columns where the first one is occupation and the second one is the score
def getDictFromExcel(filename):
    workbook = xlrd.open_workbook(filename)
    sheet = workbook.sheet_by_index(0)
    num = int(sheet.cell_value(1,5))
    print(num)
    occup_imp = {}
    for row in range(1,num+1):
        occup_imp[sheet.cell_value(row,0)]=sheet.cell_value(row,1)
    return occup_imp


## load the embeddings
print("Load glove wiki begin:")
word_vectors = api.load("glove-wiki-gigaword-100")
print("Load model succesfully \n")
## word_vectors = api.load("glove-twitter-200")
occupation_list = getListFromText("occupation_name.txt")
occup_imp = getDictFromExcel("occup_imp.xls")
createOccupationExcelFile(occupation_list,word_vectors,occup_imp)





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
