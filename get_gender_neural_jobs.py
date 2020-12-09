import pandas as pd

df = pd.read_excel (r'cpsaat11_reformatted.xlsx')
print (df)

gender_neutral_jobs_list = ''

for index, row in df.iterrows():
    value = row['Women']
    job = row['Occupation']
    if type(value) == float and value > 0:
        if value >= 40.0 and value <= 60.0:
            data = job + ': ' + str(value)
            gender_neutral_jobs_list += data
            gender_neutral_jobs_list += '\n'


print(gender_neutral_jobs_list)

text_file = open("gender_neutral_jobs.txt", "w")
text_file.write(gender_neutral_jobs_list)
text_file.close()