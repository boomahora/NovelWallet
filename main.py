import pandas as pd
import time
from wordwrangling import *

start_time = time.time()

# Paths
doc_path = r'C:\Repos\WordReplacement\input\NW - Final.docx'
word_list_path = r'C:\Repos\WordReplacement\input\wordlist_new.xlsx'
modified_doc_path = r'C:\Repos\WordReplacement\input\modified_document.docx'
output_excel_path = r'C:\Repos\WordReplacement\output\word_page_mapping.xlsx'
output_doc_path = r'C:\Repos\WordReplacement\output\clean test.docx'

# Load word list
df_word_list = pd.read_excel(word_list_path, header=None)
words = df_word_list.iloc[:, 0].astype(str).tolist()

# Modify original document and extract old chapter ranges
old_chapter_ranges, pdf_path = modify_original_word(doc_path, modified_doc_path, words)
print("Modified words")
print("--- %s seconds ---" % (time.time() - start_time))

# Extract words from the modified PDF
extracted_words = extract_words_from_pdf(pdf_path)
print("Extracted words")
print("--- %s seconds ---" % (time.time() - start_time))

# Post-process the docx to remove identifiers and set font
remove_identifiers_from_docx(modified_doc_path, output_doc_path)
set_font_for_docx(output_doc_path)
print("New word-file")
print("--- %s seconds ---" % (time.time() - start_time))
# Extract new chapter ranges from the cleaned PDF
clean_pdf_path = output_doc_path.replace('.docx', '.pdf')
new_chapter_ranges = fix_chapter_dict(read_pdf_pages(clean_pdf_path))
print("New chapter ranges")
print("--- %s seconds ---" % (time.time() - start_time))
# Map extracted words to chapters using old and new chapter ranges
word_chapter_mappings = map_extracted_words(extracted_words, words, old_chapter_ranges, new_chapter_ranges)
print("Word mappings")
print("--- %s seconds ---" % (time.time() - start_time))

# Save the results to an Excel file
df = pd.DataFrame(word_chapter_mappings, columns=["Word", "Page Number", "Chapter"]).sort_values(by=['Word', 'Chapter', 'Page Number'])
df.to_excel(output_excel_path.replace('.xlsx', '_with_chapters.xlsx'), index=False)

print("--- %s seconds ---" % (time.time() - start_time))
