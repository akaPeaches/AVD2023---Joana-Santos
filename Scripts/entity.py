import spacy
from collections import Counter
from openpyxl import Workbook

# Load the Portuguese model
nlp = spacy.load('pt_core_news_lg')

# Open the file and read the text
with open('C:/Users/jucas/Desktop/git.pi/AVD2023---Joana-Santos/Scripts/Camilo-A_Morgada_de_Romariz.md', 'r', encoding='utf-8') as file:
    text = file.read().lower()  # Lowercase the text

# Process the text
doc = nlp(text)

# Extract names, locations, and organizations, and count their occurrences
names = [entity.text.lower() for entity in doc.ents if entity.label_ == 'PER']
locs = [entity.text.lower() for entity in doc.ents if entity.label_ == 'LOC']
orgs = [entity.text.lower() for entity in doc.ents if entity.label_ == 'ORG']
prepositions = [token.text.lower() for token in doc if token.pos_ == 'ADP']
adjectives = [token.text.lower() for token in doc if token.pos_ == 'ADJ']
verbs = [(token.text.lower(), token.lemma_.lower()) for token in doc if token.pos_ == 'VERB']  # Lowercase the verb form and the lemma
verb_lemmas = [token.lemma_.lower() for token in doc if token.pos_ == 'VERB']  # Lowercase the lemmas
nouns = [token.text.lower() for token in doc if token.pos_ == 'NOUN']
stopwords = [token.text.lower() for token in doc if token.is_stop]

# Rest of the code remains the same...


name_counts = Counter(names)
loc_counts = Counter(locs)
org_counts = Counter(orgs)
prep_counts = Counter(prepositions)
adj_counts = Counter(adjectives)
verb_counts = Counter(verbs)
verb_lemma_counts = Counter(verb_lemmas)  # New counter for verb lemmas
noun_counts = Counter(nouns)
stopword_counts = Counter(stopwords)

sorted_name_counts = sorted(name_counts.items(), key=lambda x: x[1], reverse=True)
sorted_loc_counts = sorted(loc_counts.items(), key=lambda x: x[1], reverse=True)
sorted_org_counts = sorted(org_counts.items(), key=lambda x: x[1], reverse=True)
sorted_prep_counts = sorted(prep_counts.items(), key=lambda x: x[1], reverse=True)
sorted_adj_counts = sorted(adj_counts.items(), key=lambda x: x[1], reverse=True)
sorted_verb_counts = sorted(verb_counts.items(), key=lambda x: x[1], reverse=True)
sorted_verb_lemma_counts = sorted(verb_lemma_counts.items(), key=lambda x: x[1], reverse=True)  # Sort the lemma counts
sorted_noun_counts = sorted(noun_counts.items(), key=lambda x: x[1], reverse=True)
sorted_stopword_counts = sorted(stopword_counts.items(), key=lambda x: x[1], reverse=True)

wb = Workbook()
ws = wb.active

ws.append(['Nomes', 'Count', 'Localizações', 'Count', 'Organizações', 'Count', 'Prepositions', 'Count', 'Adjectives', 'Count', 'Verbs', 'Lemmas', 'Count', 'Verb Lemmas', 'Count', 'Nouns', 'Count', 'Stopwords', 'Count'])

max_len = max(len(sorted_name_counts), len(sorted_loc_counts), len(sorted_org_counts), len(sorted_prep_counts), len(sorted_adj_counts), len(sorted_verb_counts), len(sorted_verb_lemma_counts), len(sorted_noun_counts), len(sorted_stopword_counts))

for i in range(max_len):
    row = []

    if i < len(sorted_name_counts):
        row.extend(sorted_name_counts[i])
    else:
        row.extend(['', ''])
        
    if i < len(sorted_loc_counts):
        row.extend(sorted_loc_counts[i])
    else:
        row.extend(['', ''])

    if i < len(sorted_org_counts):
        row.extend(sorted_org_counts[i])
    else:
        row.extend(['', ''])

    if i < len(sorted_prep_counts):
        row.extend(sorted_prep_counts[i])
    else:
        row.extend(['', ''])

    if i < len(sorted_adj_counts):
        row.extend(sorted_adj_counts[i])
    else:
        row.extend(['', ''])

    if i < len(sorted_verb_counts):
        verb, lemma, count = sorted_verb_counts[i][0][0], sorted_verb_counts[i][0][1], sorted_verb_counts[i][1]
        row.extend([verb, lemma, count])
    else:
        row.extend(['', '', ''])

    if i < len(sorted_verb_lemma_counts):
        row.extend(sorted_verb_lemma_counts[i])
    else:
        row.extend(['', ''])

    if i < len(sorted_noun_counts):
        row.extend(sorted_noun_counts[i])
    else:
        row.extend(['', ''])

    if i < len(sorted_stopword_counts):
        row.extend(sorted_stopword_counts[i])
    else:
        row.extend(['', ''])

    ws.append(row)

wb.save('A_morgada_de_Romariz_entities.xlsx')

