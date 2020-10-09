from spellchecker import SpellChecker
import re
from termcolor import colored

#Enter the file name and path
path=r'E:\\'
file_name='Assignment_Sampledata.txt'


spell = SpellChecker()
file1 = open(path+file_name, 'r', encoding="utf8") 
Lines = file1.readlines()

class text_validation:
    def check_error(Lines):
        count=1
        for line in Lines: 
    
            res1 = re.sub(r'[^\w\s]', '', line)
            res = re.sub('[0-9]', '', res1)
    
        
            try:
                misspelled = spell.unknown(res.split())
                word_list=[]
                for word in misspelled:
                    print("Line",count)
                    print(word, ':', colored(spell.correction(word), 'green', attrs=['bold']))
                    word_list.append({word:colored(spell.correction(word), 'green', attrs=['bold'])})
                count=count+1
            except Exception as e:
                print(e)
        return word_list
            
print(text_validation.check_error(Lines))