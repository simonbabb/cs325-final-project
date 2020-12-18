# cs325-final-project

Freqeuntly used website:
Download embedding model: http://vectors.nlpl.eu/repository/#
Scalar projection: https://www.ck12.org/book/ck-12-college-precalculus/section/9.6/


Paper:
Quantify stereotype: https://arxiv.org/pdf/1606.06121.pdf
100 year change: https://www.pnas.org/content/pnas/115/16/E3635.full.pdf


The most important file is get_raw_data.py which automatically generates the combined.xls which contains the metrics for 6 different models. 

cpsaat related files are orgiinal occupation list.
get_gender_neutral_jobs.py are used to automatically obtain a list of gender netrual jobs with 40~60% neutrality. This list is then processed to obtain a occupation_name.txt where most of them are in the embedding.

To run the program there are more than 5 gigabytes models needed to be downloaded. The code get_raw_data.py and combined.xls are important output.
