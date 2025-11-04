import pandas as pd
from langdetect import detect
from googletrans import Translator

# Load Excel
df = pd.read_csv("school_student_sample.csv")

translator = Translator()

# Function to check and translate
def translate_if_marathi(text):
    try:
        if isinstance(text, str):  # Only process text
            lang = detect(text)
            if lang == "mr":  # Marathi language code
                return translator.translate(text, src="mr", dest="en").text
        return text
    except:
        return text

# Apply translation to all cells
df_translated = df.applymap(translate_if_marathi)

# Save new Excel
df_translated.to_csv("output.csv", index=False)

print("✅ Translation completed. New file saved as output.xlsx")
