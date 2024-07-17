import os
import pickle
from natasha import Segmenter, NewsEmbedding, NewsMorphTagger, MorphVocab, Doc

# Load term dictionaries
term_dict_files = ['arts_terms', 'bio_terms', 'cs_terms', 'math_terms', 'mus_terms', 'phys_terms', 'rus_terms']
term_dicts = {}

for term_file in term_dict_files:
    with open(os.path.join(os.getcwd(), 'termdicts', term_file), 'rb') as f:
        term_dicts[term_file.split('_')[0]] = pickle.load(f)

# Natasha NLP setup
segmenter = Segmenter()
emb = NewsEmbedding()
morph_tagger = NewsMorphTagger(emb)
morph_vocab = MorphVocab()

class NLP_analyzer:
    def __init__(self, text):
        self.text = text

    def lemmatize(self):
        doc = Doc(self.text)
        doc.segment(segmenter)
        doc.tag_morph(morph_tagger)
        for token in doc.tokens:
            token.lemmatize(morph_vocab)
        return [token.lemma for token in doc.tokens]

def TERM_EXTRACTOR(termdict, lemmas):
    termdict_lemmas = termdict.keys()
    terms = []  # идентифицированные термины
    ignoring_for_text = set()  # для хранения информации о токенах, которые определены как термины

    max_len = 6
    for length in range(max_len, 0, -1):
        for i in range(len(lemmas) - length + 1):
            if any(j in ignoring_for_text for j in range(i, i + length)):
                continue
            collocation_candidate = " ".join(lemmas[i:i + length])
            if collocation_candidate in termdict_lemmas:
                terms.append((collocation_candidate, termdict[collocation_candidate]))
                ignoring_for_text.update(range(i, i + length))

    return terms

def get_textbooks_list():
    textbooks_dir = 'textbooks'
    textbooks_list = [f for f in os.listdir(textbooks_dir) if os.path.isfile(os.path.join(textbooks_dir, f))]
    lemmas_lists = []
    term_occurrences = {}

    for textbook in textbooks_list:
        with open(os.path.join(textbooks_dir, textbook), 'r', encoding='utf-8') as f:
            text = f.read()
            lemmas = NLP_analyzer(text).lemmatize()
        lemmas_lists.append(lemmas)
    
    for lemmas in lemmas_lists:
        for term_key, termdict in term_dicts.items():
            extracted_terms = TERM_EXTRACTOR(termdict, lemmas)
            for term, term_form in extracted_terms:
                if term_form not in term_occurrences:
                    term_occurrences[term_form] = set()
                term_occurrences[term_form].add(term_key)
    
    # Convert sets back to lists for the final output and filter terms that appear in at least two dictionaries
    filtered_term_occurrences = {term: list(dicts) for term, dicts in term_occurrences.items() if len(dicts) >= 2}
    
    return filtered_term_occurrences

# Generate and print the list of terms with their occurrences in term dictionaries
terms_dict = get_textbooks_list()
for term, dicts in terms_dict.items():
    print(f"Term: {term}, Found in dictionaries: {dicts}")
