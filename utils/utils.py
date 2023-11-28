import os
import olefile
import re

def hwp_to_txt(hwp_file):
    f = olefile.OleFileIO(hwp_file)  
    encoded_text = f.openstream('PrvText').read() 
    decoded_text = encoded_text.decode('utf-16')  
    
    with open(f"{hwp_file[:-4]}.txt", 'w', encoding='utf-8') as f:
        f.write(decoded_text)

    return decoded_text