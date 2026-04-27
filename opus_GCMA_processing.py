import sys
sys.path.insert(1, r'C:\Users\pracownik\Documents\IBL-PAN-Python')
sys.path.insert(1, r'C:\Users\pracownik\Documents\GCMA')
import pandas as pd
from my_functions import gsheet_to_df, gdoc_to_str
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor


#%% def

def get_letter_content(letter_id):
    # letter_id = letters_ids[0]
    url = id_transcription_url.get(letter_id)
    try:
       
        if pd.notna(url):
            docs_id = max(url.split('/'),key=len)
            text = gdoc_to_str(docs_id)
        else: text = None
        letters_content.append({letter_id:text})
    except: print(url)
    
#%% download
df_metadata = gsheet_to_df('1Yi6m0vEJndcmApy-y9vlsMiLayUaYJ48aGlHJN-69RM', 'Letter manifestations')

id_transcription_url = {k:v for k,v in dict(zip(df_metadata['letter_manifestation_ID'], df_metadata['transcription'])).items() if pd.notna(v)}
letters_ids = list(id_transcription_url.keys())

letters_content = []
with ThreadPoolExecutor() as excecutor:
    list(tqdm(excecutor.map(get_letter_content, letters_ids),total=len(letters_ids)))   
letters_content = [{k:v.strip() if isinstance(v,str) else None for k,v in e.items()} for e in letters_content]

letters_content_dict = {k: v for item in letters_content for k, v in item.items()}

df_metadata['full text'] = df_metadata['letter_manifestation_ID'].apply(lambda x: letters_content_dict.get(x))
#%%
dict_final = {}

for i, row in tqdm(df_metadata.iterrows(), total=df_metadata.shape[0]):
    # i = 0
    # row = df.loc[i]
    if pd.notna(row['full text']):
        
        # document_type = row['document type']
        document_type = 'letter'
        date = row['date_of_letter']
        author = row['author_name']
        recipient = row['recipient_name']
        keywords = row['keywords_manual']
        abstract = row['abstract']
        full_text = row['full text']
        
        gcma_prompt = f"""You are a historian specializing in early modern diplomatic correspondence, particularly the interplay between secular rulers, ecclesiastical authorities, and ambassadors. You will use the Grammar of Conspiracy Matrix Analysis (GCMA) to analyze the following primary source, paying close attention to both the linguistic markers of conspiracy and the historical–diplomatic context (ecclesiastical politics, negotiated agreements, and codes of conduct in ambassadorial dispatches).
    
        This manual consists of the following elements:
        Characteristics of the source
        Fact-checking classification
        Paranoia classification
        Markers of conspiratorial language
        JSON output format
        Instructions with steps to take
    
        Characteristics of the source:
        •	document type: {document_type}
        •	date/place:  {date}
        •	author: {author}
        •	recipient: {recipient}
        •	keywords: {keywords}
        •	abstract: {abstract}
        •	full text: {full_text}
    
        Fact-checking Classification:
        Classify each extracted excerpt into one of three categories based on the following definitions:
        	•	Observed Conspiracies
        Definition: Excerpts describing documented covert actions or secret negotiations, with precise historical evidence (dates, decrees, testimony).
        Markers: Factual specificity (clear references to official documents, treaties, amounts of money, or times and places), cross-verification (mentions of multiple independent sources, e.g., “as confirmed by the Venetian Senate” or “matching the Spanish ambassador’s dispatch”), neutral tone (factual statements without “might,” “could,” or emotional qualifiers).
        Early modern example: An ambassador states that on March 12, the Pope wrote a letter to King X, offering financial support for a rebellion, citing an archived receipt or sworn witness.
        	•	Rumored Conspiracies
        Definition: Excerpts indicating partial but inconclusive evidence of conspiratorial acts, informed primarily by secondhand sources or unverified gossip.
        Markers: Hedge language (for example, “it is believed,” “we have heard,” “common talk says…”), partial specificity with vague details, and diplomatic caveats indicating a need for corroboration. This category may also contain elements of paranoia, as the language used may exhibit heightened suspicion or anxiety regarding hidden motives.
        Early modern example: A letter referencing an alleged plot by the Spanish ambassador “heard from reliable sources,” but only one letter or a single witness supports it, with no corroboration from official channels.
        	•	Inferred Conspiracies
        Definition: Excerpts in which the author posits hidden motives or covert designs based on personal deduction or circumstantial evidence, rather than direct statements from external sources.
        Markers: Interpretive language (for example, “surely,” “must be,” “cannot conceive otherwise”), conditional or subjunctive constructions (such as “if they delay… it must be…,” “were it not for…”), and a mildly charged emotional tone (suggestive adjectives such as “covert,” “insidious”). Paranoid elements are often present here as the writer may overemphasize the secrecy or nefarious intent behind actions.
        Early modern example: An ambassador concluding that a delay in negotiations implies France’s secret alignment with the Habsburg cause, even though no official statement backs it.
        Paranoia classification
        Classify each excerpt, answering the question of whether paranoia is present by assigning a “true” or “false” value, and then explain your reasoning.
        Definition
        Excerpts characterized by heightened emotional or irrational language, reflecting an obsessive belief in unstoppable conspiracies. The text shows extreme ideological bias.
        Markers to Look For
        	•	Exaggerated Rhetoric: “Every corner is watched,” “No one can be trusted!”
        	•	Emphatic, Emotional Tone: Repeated exclamations, doomsday imagery, rhetorical questions (“How could they do otherwise but conspire?!”).
        	•	Syntactic Intensifiers: Many ellipses, incomplete sentences, or breathless punctuation, e.g., “—indeed, they have … —we are undone!”
        Early Modern Example
        An ambassador’s letter describing “eyes in every corner” and insisting “the apocalypse is nigh unless we eradicate these papists,” with no measured analysis.
        Markers of Conspiratorial Language:
        For the purposes of this analysis, the following markers must be documented and included in the JSON output:
        	1.	Modal Verbs: Words such as “might,” “could,” “must,” and “should” when used to imply uncertainty or obligation.
        	2.	Evaluative Language: Adjectives and adverbs that convey judgment (for example, “clandestine,” “insidious,” “exceeding”) or sarcasm.
        	3.	Syntax and Conditional Constructions: Phrases and sentence structures that indicate hidden motives (for example, “if…then,” “were it not for…”).
        	4.	Discourse Markers: Expressions that signal speculation or hidden intent (for example, “it is believed,” “one might think,” “so that it was rather to be thought”).
        	5.	Punctuation: Specific punctuation choices (such as ellipsis and exclamation point) that suggest conspiratorial thinking.
        Paranoia Note: When classifying conspiracies, be aware that language indicative of both rumored and inferred conspiracies can contain elements of paranoia. This may be revealed through anxious, suspicious, or hyperbolic language. Such markers of paranoia should be noted in your analysis as contributing factors to the overall conspiratorial tone.
        JSON output format:
        Return your findings as a JSON file that includes:
        	•	text_source: The full text of the source letter.
        	•	metadata: All metadata provided (document type, author, recipient, date/place, keywords, abstract).
        	•	markers_of_conspiratorial_language: A key listing the markers defined above, along with example terms or constructions for each.
        	•	analysis_overview: An overview that includes a list of segments and a count of the total excerpts identified for each category (observed, rumored, inferred).
        	•	excerpts: A list of extracted excerpts, where each entry includes:
        	•	The minimal contiguous excerpt.
        	•	The assigned Fact-checking classification (observed, rumored, inferred).
        	•	A brief explanation for the chosen Fact-checking classification, including any indications of paranoia if applicable.
        	•	The assigned Paranoia classification (true, false).
        	•	A brief explanation for the chosen Paranoia classification.
        	•	A list of people mentioned in the excerpt.
        	•	The segment of the text where the excerpt was found with the full name as in “analysis_overview.”
        Follow this JSON structure exactly:
        Follow this JSON structure exactly:
    {{
      "text_source": "",
      "metadata": {{
        "document_type": "",
        "author": "",
        "recipient": "",
        "date_place": "",
        "keywords": "",
        "abstract": ""
      }},
      "markers_of_conspiratorial_language": {{
        "modal_verbs": [],
        "evaluative_language": [],
        "syntax_and_conditional_constructions": [],
        "discourse_markers": [],
        "punctuation": []
      }},
      "analysis_overview": {{
        "segments": [],
        "no_of_excerpts": {{
          "observed": 0,
          "rumored": 0,
          "inferred": 0
        }}
      }},
      "excerpts": [
        {{
          "excerpt": "",
          "gcma_classification": "",
          "explanation_of_chosen_gcma_type": "",
          "paranoia_detected": false,
          "explanation_of_paranoia_detection": "",
          "people_mentioned_in_excerpt": [],
          "segment": ""
        }}
      ]
    }}
        Instructions:
        	1.	Segmentation: Break the letter into thematic segments containing conspiratorial material, clearly identifying each.
        	2.	Extraction: Identify textual references signaling conspiratorial activities—covert actions, secret negotiations, or hidden motives—explicitly demonstrating deception, manipulation, or hidden intent beyond ordinary diplomatic maneuvering. Extract the smallest contiguous excerpts (as many as you can find; write all the excerpts you calculated for “no_of_excerpts”) with all necessary linguistic and contextual markers.
        	3.	Analyzing Conspiratorial Language: Identify linguistic markers as described above, clearly situating them within the broader historical–diplomatic context. Explicitly highlight any expressions or stylistic features indicative of paranoia, noting their impact on meaning and intent.
        	4.	Classification: Classify each excerpt as instructed using both classifiers: Fact-checking classification and Paranoia classification.
        	5.	Generate a JSON file: Generate a JSON file following the structure provided above. Use “UTF-8” encoding for the JSON file. I want to save the file on my computer as “GCMA_analysis.json”. I want to download the file by clicking on a link. Provide a downloadable link for me.
        """
        
        dict_final.update({row['letter_manifestation_ID']: gcma_prompt})

for k, v in dict_final.items():
    with open(f"data/GCMA/{k}.txt", "w", encoding="utf-8") as file:
        file.write(v)

#%%
