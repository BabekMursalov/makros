
Bu proqram, göstərilən bir Word sənədini oxuyaraq içərisindəki sualları və cavabları Excel faylının içərisinə köçürən bir Python skriptidir.

Əvvəlcədən, docx və openpyxl kitabxanalarını proqramın başında çağırırıq.

python
Copy code
from docx import Document
from openpyxl import Workbook
word_to_excel adlı bir funksiya təyin edirik və bu funksiya iki parametr qəbul edir: word_file və excel_file. word_file dəyişəni, oxunacaq Word faylının adını, excel_file isə yaradılacaq Excel faylının adını təmsil edir.

python
Copy code
def word_to_excel(word_file, excel_file):
Funksiya daxilində Word faylını açmaq üçün Document sinifindən istifadə edirik. Verilən word_file adı ilə fayl açılır və doc dəyişəninə mənimsədilir.

python
Copy code
doc = Document(word_file)
Daha sonra, openpyxl kitabxanasından istifadə edərək yeni bir Excel faylı yaradırıq. Workbook sinifindən istifadə edərək iş kitabı yaradırıq və wb dəyişəninə mənimsədilir.

python
Copy code
wb = Workbook()
Yaradılan Excel faylındakı varsayılan iş hərəkətini əldə etmək üçün active xüsusiyyətindən istifadə edərək sheet dəyişəninə mənimsədirik.

python
Copy code
sheet = wb.active
Daha sonra, Word sənədindəki sualları və şıqları Excel faylına köçürmək üçün bir dövr istifadə edirik. Bu üçün doc.paragraphs üzərində dolaşırıq.

python
Copy code
for paragraph in doc.paragraphs:
Hər bir paragrafın mətnini əldə edirik və başda və sonda olan boşluqları təmizləyirik.

python
Copy code
text = paragraph.text.strip()
Əgər paragraf bir sualla bitirsə (sual işarəsi ilə bitirsə), əvvəlki sualı və şıqları Excel faylına əlavə edirik. Bunun üçün sheet.append() metodundan istifadə edərək bir sətir yaradırıq. Sual mətni A sütununa, şıqlar isə B, C, D və E sütunlarına yerləşdirilir.

python
Copy code
if text.endswith('?'):
    if question
