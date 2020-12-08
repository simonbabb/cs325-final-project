# from gensim.test.utils import lee_corpus_list
# from gensim.models import Word2Vec
# model = Word2Vec(lee_corpus_list, vector_size=24, epochs=100)
# word_vectors = model.wv

# https://radimrehurek.com/gensim/models/keyedvectors.html
import gensim.downloader as api

# load the embeddings
word_vectors = api.load("glove-wiki-gigaword-100")
# word_vectors = api.load("glove-twitter-200")

# test analogies
result = word_vectors.most_similar(positive=['woman', 'king'], negative=['man'])
most_similar_key, similarity = result[0]  # look at the first match
print(f"{most_similar_key}: {similarity:.4f}")

result = word_vectors.most_similar(positive=['woman', 'intelligent'], negative=['man'])
most_similar_key, similarity = result[0]  # look at the first match
print(f"{most_similar_key}: {similarity:.4f}")

result = word_vectors.most_similar(positive=['caucasian', 'lazy'], negative=['hispanic'])
most_similar_key, similarity = result[0]  # look at the first match
print(f"{most_similar_key}: {similarity:.4f}")

result = word_vectors.most_similar(positive=['caucasian', 'housekeeper'], negative=['hispanic'])
most_similar_key, similarity = result[0]  # look at the first match
print(f"{most_similar_key}: {similarity:.4f}")


# test similarity between two words
similarity = word_vectors.similarity('lawyer', 'man')
print(similarity)

similarity = word_vectors.similarity('lawyer', 'woman')
print(similarity)

similarity = word_vectors.similarity('lawyer', 'he')
print(similarity)

similarity = word_vectors.similarity('lawyer', 'she')
print(similarity)



similarity = word_vectors.similarity('bossy', 'man')
print(similarity)

similarity = word_vectors.similarity('bossy', 'woman')
print(similarity)

similarity = word_vectors.similarity('bossy', 'he')
print(similarity)

similarity = word_vectors.similarity('bossy', 'she')
print(similarity)



similarity = word_vectors.similarity('executive', 'man')
print(similarity)

similarity = word_vectors.similarity('executive', 'woman')
print(similarity)

similarity = word_vectors.similarity('executive', 'he')
print(similarity)

similarity = word_vectors.similarity('executive', 'she')
print(similarity)


