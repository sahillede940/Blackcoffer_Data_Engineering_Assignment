import string
import os
import requests
from bs4 import BeautifulSoup
import pandas as pd
from nltk.tokenize import sent_tokenize, word_tokenize
import nltk
from nltk.corpus import cmudict
from nltk.corpus import stopwords as nltk_stopwords

nltk.download('punkt')
nltk.download('cmudict')
nltk.download('stopwords')

input_sheet = "Input"
output_sheet = "Output"
stopwords = []
positive_words = []
negative_words = []

# Read the excel sheet
df = pd.read_excel(input_sheet+".xlsx")


def read_excel():
    df = pd.read_excel(output_sheet+".xlsx")
    print(df)

# Function to extract the content of the articles

# check if Articles dir is present or not
if not os.path.exists("Articles"):
    os.mkdir("Articles")

def extract_content():
    for index, row in df.iterrows():
        # if file already exists, skip the extraction
        if os.path.exists("Articles/" + str(row["URL_ID"]) + ".txt"):
            print("Article " + str(row["URL_ID"]) + " already exists")
            continue

        page = requests.get(row["URL"])
        soup = BeautifulSoup(page.content, "html.parser")
        with open("Articles/" + str(row["URL_ID"]) + ".txt", "w") as file:
            # add title to the file
            title = soup.select_one(
                "h1.entry-title") or soup.select_one("tdb-title-text")
            if title:
                file.write(title.get_text())
                file.write("\n")
            # add content to the file
            file.write("\n")
            content = soup.select_one(
                "div.td-post-content") or soup.select_one("div.tdb-block-inner")
            if content:
                file.write(content.get_text())

            print("Article " + str(row["URL_ID"]) + " extracted successfully")
            break

# get the stopwords from the files


def read_stopwords(stopwords=[]):
    with open("StopWords/StopWords_Auditor.txt", "r") as file:
        stopwords += file.read().split("\n")
    with open("StopWords/StopWords_Currencies.txt", "r", encoding="ISO-8859-1") as file:
        stopwords += [line.split("|")[0] for line in file.read().split("\n")]
    with open("StopWords/StopWords_DatesandNumbers.txt", "r") as file:
        stopwords += file.read().split("\n")
    with open("StopWords/StopWords_Generic.txt", "r") as file:
        stopwords += file.read().split("\n")
    with open("StopWords/StopWords_GenericLong.txt", "r") as file:
        stopwords += file.read().split("\n")
    with open("StopWords/StopWords_Geographic.txt", "r") as file:
        stopwords += file.read().split("\n")
    with open("StopWords/StopWords_Names.txt", "r") as file:
        stopwords += file.read().split("\n")
    # remove duplicates
    stopwords = list(set(stopwords))


def read_master_dict(positive_words=[], negative_words=[]):
    with open("MasterDictionary/positive-words.txt", "r") as file:
        positive_words += file.read().split("\n")
        positive_words = [
            word for word in positive_words if word not in stopwords]

    with open("MasterDictionary/negative-words.txt", "r", encoding="ISO-8859-1") as file:
        negative_words += file.read().split("\n")
        negative_words = [
            word for word in negative_words if word not in stopwords]


def sentiment_analysis():
    for index, row in df.iterrows():
        with open("Articles/" + str(row["URL_ID"]) + ".txt", "r") as file:
            article = file.read()
            tokens = nltk.word_tokenize(article)
            positive = 0
            negative = 0
            for word in tokens:
                if word in positive_words:
                    positive += 1
                elif word in negative_words:
                    negative += 1
            # save into the excel sheet
            df.at[index, "POSITIVE SCORE"] = positive
            df.at[index, "NEGATIVE SCORE"] = negative
            if positive + negative == 0:
                df.at[index, "POLARITY SCORE"] = 0
                df.at[index, "SUBJECTIVITY SCORE"] = 0
            else:
                df.at[index, "POLARITY SCORE"] = (
                    positive-negative)/(positive+negative) + 0.000001
                df.at[index, "SUBJECTIVITY SCORE"] = (
                    positive+negative)/(len(tokens)) + 0.000001
            print("Sentiment Analysis for Article " +
                  str(row["URL_ID"]) + " completed")


# Analysis of Readability
def readability_analysis():
    for index, row in df.iterrows():
        with open("Articles/" + str(row["URL_ID"]) + ".txt", "r") as file:
            article = file.read()

            # calculate the average sentence length
            sentences = sent_tokenize(article)
            words = word_tokenize(article)
            if len(sentences) == 0:
                avg_sentence_length = 0
            else:
                avg_sentence_length = len(words)/len(sentences)

            # calculate the percentage of complex words
            complex_words = 0
            for word in words:
                if len(word) > 6:
                    complex_words += 1
            if len(words) > 0:
                percentage_complex_words = complex_words/len(words)
            else:
                percentage_complex_words = 0

            # calculate the fog index
            fog_index = 0.4 * (avg_sentence_length +
                               percentage_complex_words)

            # save into the excel sheet
            df.at[index, "AVG SENTENCE LENGTH"] = avg_sentence_length
            df.at[index, "PERCENTAGE OF COMPLEX WORDS"] = percentage_complex_words
            df.at[index, "FOG INDEX"] = fog_index

            print("Readability Analysis for Article " +
                  str(row["URL_ID"]) + " completed")


# Average Number of Words Per Sentence
def avg_words_per_sentence():
    for index, row in df.iterrows():
        with open("Articles/" + str(row["URL_ID"]) + ".txt", "r") as file:
            article = file.read()
            sentences = sent_tokenize(article)
            words = word_tokenize(article)
            if len(sentences) > 0:
                avg_words_per_sentence = len(words)/len(sentences)
            else:
                avg_words_per_sentence = 0
            # save into the excel sheet
            df.at[index, "AVG NUMBER OF WORDS PER SENTENCE"] = avg_words_per_sentence
            print("avg_words_per_sentence for Article " +
                  str(row["URL_ID"]) + " completed")

# Complex Word Count


def complex_word_count():
    d = cmudict.dict()  # Load the CMU Pronouncing Dictionary

    def count_syllables(word):
        try:
            # Some words have multiple pronunciations; choose the first
            pronunciation = d[word.lower()][0]
            return len([p for p in pronunciation if p[-1].isdigit()])
        except KeyError:
            # Word not found in dictionary
            return 0  # Or consider alternative methods mentioned below
    for index, row in df.iterrows():
        with open("Articles/" + str(row["URL_ID"]) + ".txt", "r") as file:
            article = file.read()
            words = word_tokenize(article)
            complex_words = [
                word for word in words if count_syllables(word) > 2]
            # save into the excel sheet
            df.at[index, "COMPLEX WORD COUNT"] = len(complex_words)
            print("Complex word count for Article " +
                  str(row["URL_ID"]) + " completed")


# WORD COUNT
def word_count():

    for index, row in df.iterrows():
        with open("Articles/" + str(row["URL_ID"]) + ".txt", "r") as file:
            # removing the stop words (using stopwords class of nltk package).
            # removing any punctuations like ? ! , . from the word before counting.
            article = file.read()
            words = word_tokenize(article)
            words = [
                word for word in words if word not in nltk_stopwords.words('english')]
            # remove punctuations
            for i in range(len(words)):
                words[i] = words[i].translate(
                    str.maketrans('', '', string.punctuation))
            # remove empty strings
            words = [word for word in words if word != ""]
            # save into the excel sheet
            df.at[index, "WORD COUNT"] = len(words)
            print("Word count for Article " +
                  str(row["URL_ID"]) + " completed")


# Syllable Count Per Word
def syllable_count_per_word():
    d = cmudict.dict()  # Load the CMU Pronouncing Dictionary

    def count_syllables(word):
        try:
            # Some words have multiple pronunciations; choose the first
            pronunciation = d[word.lower()][0]
            return len([p for p in pronunciation if p[-1].isdigit()])
        except KeyError:
            # Word not found in dictionary
            return 0  # Or consider alternative methods mentioned below
    for index, row in df.iterrows():
        with open("Articles/" + str(row["URL_ID"]) + ".txt", "r") as file:
            article = file.read()
            words = word_tokenize(article)
            syllable_count = [count_syllables(word) for word in words]
            # save into the excel sheet
            if len(words) > 0:
                syllable_per_word = sum(syllable_count)/len(words)
            else:
                syllable_per_word = 0
            df.at[index, "SYLLABLE PER WORD"] = syllable_per_word
            print("Syllable per word for Article " +
                  str(row["URL_ID"]) + " completed")

# Personal Pronouns


def personal_pronouns():
    personal_pronouns = ["i", "we", "my", "ours", "us"]
    for index, row in df.iterrows():
        with open("Articles/" + str(row["URL_ID"]) + ".txt", "r") as file:
            article = file.read()
            # filter out word "US" from the article
            words = word_tokenize(article)
            for i in range(len(words)):
                if words[i] == "US":
                    # for excluding the word "US" from the list of personal pronouns
                    words[i] = "s"
            personal_pronouns_count = [
                word for word in words if word.lower() in personal_pronouns]

            # save into the excel sheet
            df.at[index, "PERSONAL PRONOUNS"] = len(personal_pronouns_count)
            print("Personal pronoun count for Article " +
                  str(row["URL_ID"]) + " completed")


# Average Word Length
def avg_word_length():
    for index, row in df.iterrows():
        with open("Articles/" + str(row["URL_ID"]) + ".txt", "r") as file:
            article = file.read()
            words = word_tokenize(article)
            if len(words) > 0:
                avg_word_length = sum([len(word) for word in words])/len(words)
            else:
                avg_word_length = 0
            # save into the excel sheet
            df.at[index, "AVG WORD LENGTH"] = avg_word_length
            print("Average word length count for Article " +
                  str(row["URL_ID"]) + " completed")


extract_content()
read_stopwords(stopwords)
read_master_dict(positive_words, negative_words)
sentiment_analysis()
readability_analysis()
avg_words_per_sentence()
complex_word_count()
word_count()
syllable_count_per_word()
personal_pronouns()
avg_word_length()

df.to_excel(output_sheet+".xlsx", index=False)

print("You can now open file named Output.xlsx")
