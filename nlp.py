import pandas as pd
import spacy

excel_file_path = 'f1.xlsx'
df = pd.read_excel(excel_file_path, engine='openpyxl')

print(df.head())

nlp = spacy.load('en_core_web_sm')

def extract_entities(question):
    doc = nlp(question)
    entities = {ent.label_: ent.text for ent in doc.ents}
    print(f"Extracted entities: {entities}")
    return entities

def map_question_to_query(question, df):
    entities = extract_entities(question)
  
    df['Name'] = df['Name'].str.lower()
    
    if 'PERSON' in entities:
        name = entities['PERSON'].lower()
        if name in df['Name'].values:
            if 'department' in question.lower():
                result = df.loc[df['Name'] == name, 'Department']
                if not result.empty:
                    return f"The department of {entities['PERSON']} is {result.values[0]}"
            elif 'email' in question.lower():
                result = df.loc[df['Name'] == name, 'Email']
                if not result.empty:
                    return f"The email of {entities['PERSON']} is {result.values[0]}"
            elif 'contact' in question.lower() or 'phone' in question.lower():
                result = df.loc[df['Name'] == name, 'Contact']
                if not result.empty:
                    return f"The contact number of {entities['PERSON']} is {result.values[0]}"
            elif 'course' in question.lower():
                result = df.loc[df['Name'] == name, 'Course']
                if not result.empty:
                    return f"{entities['PERSON']} is enrolled in the {result.values[0]} course"
            elif 'college' in question.lower():
                result = df.loc[df['Name'] == name, 'College']
                if not result.empty:
                    return f"{entities['PERSON']} is studying at {result.values[0]}"
            elif 'birth' in question.lower() or 'dob' in question.lower():
                result = df.loc[df['Name'] == name, 'Date of birth']
                if not result.empty:
                    return f"{entities['PERSON']}'s date of birth is {result.values[0]}"
    else:
        potential_name = []
        for token in nlp(question):
            if token.pos_ == 'PROPN':
                potential_name.append(token.text.lower())
        
        if potential_name:
            name = ' '.join(potential_name)
            if name in df['Name'].values:
                if 'department' in question.lower():
                    result = df.loc[df['Name'] == name, 'Department']
                    if not result.empty:
                        return f"The department of {name.title()} is {result.values[0]}"
                elif 'email' in question.lower():
                    result = df.loc[df['Name'] == name, 'Email']
                    if not result.empty:
                        return f"The email of {name.title()} is {result.values[0]}"
                elif 'contact' in question.lower() or 'phone' in question.lower():
                    result = df.loc[df['Name'] == name, 'Contact']
                    if not result.empty:
                        return f"The contact number of {name.title()} is {result.values[0]}"
                elif 'course' in question.lower():
                    result = df.loc[df['Name'] == name, 'Course']
                    if not result.empty:
                        return f"{name.title()} is enrolled in the {result.values[0]} course"
                elif 'college' in question.lower():
                    result = df.loc[df['Name'] == name, 'College']
                    if not result.empty:
                        return f"{name.title()} is studying at {result.values[0]}"
                elif 'birth' in question.lower() or 'dob' in question.lower():
                    result = df.loc[df['Name'] == name, 'Date of birth']
                    if not result.empty:
                        return f"{name.title()}'s date of birth is {result.values[0]}"
    
    return "I'm sorry, I couldn't understand the question or find the information."

# Continuously prompt the user for questions
while True:
    user_question = input("Please enter your question (or type 'exit' to quit): ")
    if user_question.lower() == 'exit':
        break
    answer = map_question_to_query(user_question, df)
    print(answer)
