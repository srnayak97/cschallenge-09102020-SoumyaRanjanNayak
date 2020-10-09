from spellchecker import SpellChecker
import re
from termcolor import colored
import pandas as pd 

#Enter the file name and path
path=r'E:\\GitHub\\spell_corrections\\'
file_name='Assignment_Sampledata.txt'


spell = SpellChecker()
file1 = open(path+file_name, 'r', encoding="utf8") 
Lines = file1.readlines()

class text_validation:
    def check_error(Lines):
        count=1
        word_list=[]
        for line in Lines: 
    
            res1 = re.sub(r'[^\w\s]', '', line)
            res = re.sub('[0-9]', '', res1)
    
        
            try:
                misspelled = spell.unknown(res.split())
                
                for word in misspelled:
                    word_list.append((count,word,spell.correction(word)))
                count=count+1
            
            except Exception as e:
                print(e)
         
        return word_list
    
    
    def creat_excel(answer):    
        df = pd.DataFrame(columns=['Line', 'word', 'correct_word'])
        for count, word, correct_word in answer:
            df = df.append({'Line': count, 'word': word, 'correct_word': correct_word}, ignore_index=True)
            
        writer = pd.ExcelWriter('E:\\GitHub\\spell_corrections\\result.xlsx', engine='xlsxwriter')
        df.to_excel(writer, sheet_name="Sheet1", index=False)
        
        worksheet = writer.sheets['Sheet1']
        workbook = writer.book
        cell_format = workbook.add_format()
        cell_format.set_bold()
        cell_format.set_font_color('green')
        
        worksheet.set_column('C:C', None, cell_format)
        
        writer.close()

          
answer=text_validation.check_error(Lines)
text_validation.creat_excel(answer)
file1.close()
    