from datetime import datetime
import pandas as pd
from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.pdfmetrics import registerFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
from sqlalchemy import create_engine
import subprocess

class Atest():
    def __init__(self, atest_path, docs_path, db_path, document_id=None):
        self.atest_path = atest_path
        self.db_path = db_path
        if document_id == None:
            # Get leatest file from documents directory
            files = docs_path.glob('*.xls')
            latest_file = max(files, key=lambda x: x.stat().st_ctime)
            self.document_id = latest_file.name.split('.')[0].replace('_','/')
        else:
            self.document_id = document_id
            
        #FS document reading and data extraction
        fs_df = pd.read_excel(str(docs_path/self.document_id.replace('/','_')) + '.xls')
        self.assortment = fs_df \
            .iloc[:,[1,4]] \
            .dropna(how='any')[1:] \
            .rename(columns={'Unnamed: 1':'Asortyment', 'Unnamed: 4': 'Ilość'})

        # sort assortment by quantity in ascending order
        self.assortment['Ilość'] = self.assortment['Ilość'].apply(lambda x: int(''.join(x.split()[:-1])))
        self.assortment = self.assortment.sort_values(by='Ilość', ascending=False)
        self.assortment['Ilość'] = self.assortment['Ilość'].apply(lambda x: str(x) + ' kg')
        
        self.assortment['Asortyment'] = self.assortment['Asortyment'].replace(regex=r'Otręby pszenne.*', value='Otręby pszenne')
        self.client = fs_df.iloc[5,11].split('\n')[0].strip()

    def draw_table(self, data):
        style = TableStyle([
            ('VALIGN',      (0,0),  (-1,-1),    'MIDDLE'),
            ('FONT',        (0,0),  (-1,-1),    'Arial'),
            ('ALIGN',       (0,0),  (-1,-1),    'CENTER'),
            ('INNERGRID',   (0,1),  (-1,-3),    0.5,            colors.grey),
            ('BOX',         (0,0),  (-1,-1),    2,              colors.black),
            ('LINEABOVE',   (0,1),  (-1,1),     1,              colors.black),
            ('LINEABOVE',   (0,2),  (-1,2),     2,              colors.black),
            ('LINEABOVE',   (1,3),  (-1,3),     1,              colors.black),
            ('LINEABOVE',   (0,-2), (-1,-2),    2,              colors.black),
            ('LINEAFTER',   (0,1),  (0,-3),     1,              colors.black),
            ('LINEAFTER',   (1,1),  (1,-3),     1,              colors.black),
            ('LINEAFTER',   (-3,0), (-3,1),     1,              colors.black),
            ('LINEAFTER',   (-2,2), (-2,-3),    1,              colors.black),
            ('SPAN',        (0,0),  (-3,0)),
            ('SPAN',        (-2,0), (-1,0)),
            ('SPAN',        (1,1),  (-1,1)),
            ('SPAN',        (2,-1), (5,-1)),
            ('SPAN',        (0,-2),  (4,-2)),
            ('SPAN',        (-5,-2),  (-1,-2)),
            ('FONT',        (0,0),  (1,0),      'Arial-bd',     32),
            ('FONT',        (-2,0), (-1,0),     'Arial-bd',     16),
            ('FONT',        (1,1),  (-1,1),     'Arial-bi',  14),
            ('FONT',        (-1,3),  (-1,-3),   'Arial-it',     10),
            ('FONT',        (0,-2), (-1,-2),    'Arial-it',     11),
            ('FONT',        (0,-1), (-1,-1),    'Arial-it',     16),
            ('FONT',        (2,-1), (6,-1),     'Arial-bi',  16),
        ])  
        return Table(data, style=style)

    def make_pdf(self):
        current_date = datetime.today().strftime('%d.%m.%Yr')

        db_path = Path('./resources/params/params.db')
        engine = create_engine(f'sqlite:///{db_path}')
        param_df = pd.read_sql('params', con=engine, index_col='index')
        param_df.iloc[2] = param_df.iloc[2].apply(lambda x: str(x).replace('.', ','))
        param_df.iloc[3] = param_df.iloc[3].apply(lambda x: str(x).replace('.', ','))
        param_df.iloc[4] = param_df.iloc[4].apply(lambda x: str(x).replace('.', ','))

        registerFont(TTFont('Arial', 'arial.ttf'))
        registerFont(TTFont('Arial-it', 'ariali.ttf'))
        registerFont(TTFont('Arial-bd', 'arialbd.ttf'))
        registerFont(TTFont('Arial-bi', 'arialbi.ttf'))
        stylesheet = getSampleStyleSheet()
        doc = SimpleDocTemplate(str(self.atest_path / f'Atest {self.document_id.replace("/","_")}.pdf'), pagesize=A4, bottomMargin=0.5*inch)

        data = [
            [f'\n{" "*3}ATEST JAKOŚCIOWY{" "*3}\n'] + ['']*7 + [self.document_id,''],
            ['Odbiorca',f'\n{self.client}\n']] +\
            [[param_df.iloc[i,0], param_df.iloc[i,1]] +\
            [param_df[assort][i] for assort in self.assortment['Asortyment']] +\
            ['']*(7-len(self.assortment)) + [param_df.iloc[i,-1]] for i in range(9)] +\
            [['Ilość','-'] + self.assortment['Ilość'].tolist(),
            ['Data i podpis\n\n'] + ['']*4 + ['Pieczątka\n\n'],
            [current_date,'','Tadeusz Krajewski']]
        t = self.draw_table(data)

        p = Paragraph('''
            <para fontname=Arial-it>
            Badanie wykonano metodą bliskiej podczerwieni przy pomocy urządzenia 
            <font face=Arial-bi>INFRAMATIC 8600</font>
            </para>''', stylesheet["BodyText"])
        elements = [t,p]

        # open and save the pdf
        try:
            doc.build(elements)
            subprocess.Popen([self.atest_path / f"Atest {self.document_id.replace('/','_')}.pdf"], shell=True)
            return True
        except IOError:
            print('Plik został już otworzony!')
            return False

if __name__ == '__main__':
    atest_path = Path('./resources/atesty/')
    docs_path = Path('./resources/faktury/')
    db_path = Path('./resources/params/data_export.xlsx')
    atest = Atest(atest_path, docs_path, db_path)
    atest.make_pdf()