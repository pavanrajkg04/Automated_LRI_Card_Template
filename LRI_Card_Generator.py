# -*- coding: utf-8 -*-
"""
Created on Sat May 18 15:40:40 2024

@author: Pavan Raj K G
"""


from spire.doc import *
from spire.doc.common import *
from docx import Document as dc
from spire.doc import Document
import pandas as pd 
import os 


'''----------------------------------------------------------------------------- Data Manuplution --------------------------------------------------------------'''

file_paths = [
    "LRIDatasets\ARA_4c04_B_TableToExcel.xls",
    "LRIDatasets\ARA_4c04_Cu_TableToExcel.xls",
    "LRIDatasets\ARA_4c04_EC_TableToExcel.xls",
    "LRIDatasets\ARA_4c04_Fe_TableToExcel.xls",
    "LRIDatasets\ARA_4c04_K2O_TableToExcel.xls",
    "LRIDatasets\ARA_4c04_Mn_TableToExcel.xls",
    "LRIDatasets\ARA_4c04_N_TableToExcel.xls",
    "LRIDatasets\ARA_4c04_OC_TableToExcel.xls",
    "LRIDatasets\ARA_4c04_P2O5_TableToExcel.xls",
    "LRIDatasets\ARA_4c04_pH_TableToExcel.xls",
    "LRIDatasets\ARA_4c04_S_TableToExcel.xls",
    "LRIDatasets\ARA_4c04_Soil_Phase_Union.xls",
    "LRIDatasets\ARA_4c04_Zn_TableToExcel.xls"
]

dfs = [pd.read_excel(file_path, engine='xlrd') for file_path in file_paths]

hissa_details = pd.read_excel("LRIDatasets\hissa_details_4c2c4c04.xlsx")
hissa_details.rename(columns={'BhoomiVillageName': 'Village'}, inplace=True)
hissa_details.rename(columns={'survey_no': 'DXF_TEXT'}, inplace=True)
hissa_details['DXF_TEXT'] = pd.to_numeric(hissa_details['DXF_TEXT'], errors='coerce')
merged_df = pd.DataFrame(hissa_details)


for df in dfs:
    # Convert 'DXF_TEXT' column to numeric if needed
    df['DXF_TEXT'] = pd.to_numeric(df['DXF_TEXT'], errors='coerce')
    #df['DXF_TEXT'] = df['DXF_TEXT'].astype(int)
    print(df)
    df.head()

for df in dfs:
    merged_df = pd.merge(df, hissa_details, on=['Village', 'DXF_TEXT'], how="inner")


'''
B = pd.read_excel(r"LRIDatasets\ARA_4c04_B_TableToExcel.xls", engine='xlrd')
Cu = pd.read_excel(r"LRIDatasets\ARA_4c04_Cu_TableToExcel.xls", engine='xlrd')
Ec = pd.read_excel(r"LRIDatasets\ARA_4c04_EC_TableToExcel.xls", engine='xlrd')
Fe = pd.read_excel(r"LRIDatasets\ARA_4c04_Fe_TableToExcel.xls", engine='xlrd')
K2o = pd.read_excel(r"LRIDatasets\ARA_4c04_K2O_TableToExcel.xls", engine='xlrd')
Mn = pd.read_excel(r"LRIDatasets\ARA_4c04_Mn_TableToExcel.xls", engine='xlrd')
N = pd.read_excel(r"LRIDatasets\ARA_4c04_N_TableToExcel.xls", engine='xlrd')
OC = pd.read_excel(r"LRIDatasets\ARA_4c04_OC_TableToExcel.xls", engine='xlrd')
P2o5 = pd.read_excel(r"LRIDatasets\ARA_4c04_P2O5_TableToExcel.xls", engine='xlrd')
pH = pd.read_excel(r"LRIDatasets\ARA_4c04_pH_TableToExcel.xls", engine='xlrd')
S = pd.read_excel(r"LRIDatasets\ARA_4c04_S_TableToExcel.xls", engine='xlrd')
Soilphase = pd.read_excel(r"LRIDatasets\ARA_4c04_Soil_Phase_Union.xls", engine='xlrd')
Zn = pd.read_excel(r"LRIDatasets\ARA_4c04_Zn_TableToExcel.xls", engine='xlrd')

hissa_details = pd.read_excel(r"LRIDatasets\hissa_details_4c2c4c04.xlsx")

hissa_details.rename(columns={'BhoomiVillageName': 'Village'}, inplace=True)
hissa_details.rename(columns={'survey_no': 'DXF_TEXT'}, inplace=True)
hissa_details['DXF_TEXT'].dtype
hissa_details['DXF_TEXT'] = pd.to_numeric(hissa_details['DXF_TEXT'], errors='coerce')

B['DXF_TEXT'] = pd.to_numeric(B['DXF_TEXT'], errors='coerce')
B['DXF_TEXT'] = B['DXF_TEXT'].astype(int)
B.head()

merged_df = pd.merge(B, hissa_details, on=['Village', 'DXF_TEXT'], how="inner")

'''


'''---------------------------------------------------------------------- Data Fetching -----------------------------------------------------------------------'''

#details
Farmer_name = "ರಂಗನಾಥ ದೇಶಪಾಂಡೆ ಬಿನ್ ಬಿ.ಬಿ.ದೇಶಪಾಂಡೆ"
Gender = "ಗಂಡು"
MWS = "ಗುಮ್ಮನಾಯಕನಪಾಳ್ಯ (4C3D7v05)"
Address = "ವಸಂತಪುರ  ಗ್ರಾಮ, ಬಾಗೇಪಲ್ಲಿ ತಾ||,  ಚಿಕ್ಕಬಳ್ಳಾಪುರ ಜಿ||"
soil_year = "2023"
Survey_hissa = "56 _ 3"
Area = "-"
Annual_rain_Fall = "835"

soil_Dept = "abcd"
Soil = "abcdeewrewrewrew"




'''-------------------------------------------------------------------- Documentation ----------------------------------------------------------------------'''
# Load the existing document
doc = Document()
doc.LoadFromFile('Dummy.docx')



# Insert the first textbox and set its wrapping style
textBox1 = doc.Sections[0].Paragraphs[0].AppendTextBox(width=350, height=472)
textBox1.Format.TextWrappingStyle = TextWrappingStyle.Inline
textBox1.Format.HorizontalOrigin = HorizontalOrigin.LeftMarginArea
textBox1.Format.HorizontalPosition = 10
textBox1.Format.Background = Color.get_Black()

# Insert the second textbox and set its wrapping style and position
textBox2 = doc.Sections[0].Paragraphs[0].AppendTextBox(width=350, height=472)
textBox2.Format.TextWrappingStyle = TextWrappingStyle.Inline
textBox2.Format.HorizontalOrigin = HorizontalOrigin.RightMarginArea
textBox2.Format.HorizontalPosition = -50

# Insert text into the first textbox
para1 = textBox1.Body.AddParagraph()
para1.AppendHTML('''
    <div>
    <h4 style=" color: blue; text-align: center">ಭೂ-ಸಂಪನ್ಮೂಲ ಮಾಹಿತಿ ಕಾರ್ಡ್‌ಗಳನ್ನು ಬಳಸಲು ಉಪಯುಕ್ತ ಸಲಹೆಗಳು </h4>
    <ul>
       <li style="font-size: 5.5pt; text-align: justify">
           <strong>ಮಣ್ಣಿನ ಆಳ:</strong> ಬೆಳೆಗಳ ಬೇರಿನ ಬೆಳವಣಿಗೆಗೆ ಮಣ್ಣಿನ ಆಳವು  ಅತ್ಯಂತ ಮುಖ್ಯವಾಗಿರುವುದರಿಂದ ಕಡಿಮೆ ಆಳದ ಮಣ್ಣಿನಲ್ಲಿ ಅಲ್ಪಾವಧಿ ಹಾಗೂ ಕಡಿಮೆ ಆಳಕ್ಕೆ ಬೇರುಗಳನ್ನು ಬಿಡುವ ಬೆಳೆಗಳನ್ನು ಬೆಳೆಯುವುದು. ಬಹುವಾರ್ಷಿಕ ತೋಟಗಾರಿಕೆ ಬೆಳೆ ಬೆಳೆಯಲು ಆಳವಾದ ಗುಂಡಿ ತೆಗೆದು ಉತ್ತಮ ಮಣ್ಣನ್ನು ಹೊರಗಡೆಯಿಂದ ತಂದು ತುಂಬಿಸುವುದು ಸೂಕ್ತ.
       </li>
       <li style="font-size: 5.5pt; text-align: justify">
           <strong>ಮಣ್ಣಿನ ಕಣಗಾತ್ರ:</strong> ಜೇಡಿ ಮಣ್ಣಿಗೆ ಮರಳು ಅಥವಾ ಮೂಲ ಬಂಡೆಯಿಂದ (Parent material) ಶಿಥಿಲಗೊಂಡಿರುವ ಮಣ್ಣನ್ನು ಬೆರೆಸುವುದು. ಮರುಳು ಮಣ್ಣಿಗೆ ಕೆರೆಯ ಹೂಳು ಅಥವಾ ಕಪ್ಪು ಜೇಡಿ ಮಣ್ಣನ್ನು ಬೆರೆಸುವುದು ಸೂಕ್ತ.
       </li>
       <li style="font-size: 5.5pt; text-align: justify">
           <strong>ಮಣ್ಣಿನ ಗರಸು:</strong> ಗರಸು ಶೇ. 35 ಕ್ಕೂ ಅಧಿಕವಿದ್ದ ಮಣ್ಣಿಗೆ ಕೆರೆಯ ಹೂಳು ಅಥವಾ ಕಪ್ಪು ಜೇಡಿ ಮಣ್ಣನ್ನು ಸರಿಯಾದ ಪ್ರಮಾಣದಲ್ಲಿ ಬೆರೆಸುವುದರಿಂದ ಮಣ್ಣಿನ ಪರಿಮಾಣವನ್ನು (volume of soil) ಸುಧಾರಿಸುವುದರ ಜೊತೆಗೆ ನೀರು-ಗಾಳಿ ಮತ್ತು ಪೋಷಕಾಂಶಗಳ ಲಭ್ಯತೆಯ ಪ್ರಮಾಣವನ್ನು ಹೆಚ್ಚಿಸಬಹುದು.
       </li>
       <li style="font-size: 5.5pt; text-align: justify">
           <strong>ಮಣ್ಣಿನ ಇಳಿಜಾರು:</strong> ಮಣ್ಣಿನ ಇಳಿಜಾರಿಗೆ ಅನುಗುಣವಾಗಿ ಶಿಫಾರಸ್ಸು ಮಾಡಿದ ಸಂರಕ್ಷಣಾ ಕ್ರಮಗಳಾದ ಕಂದಕ ಬದು, ಸಮಪಾತಳಿ ಬದು, ವಾರಾಡಿ ಬದು ನಿರ್ಮಿಸುವುದು. ಇಳಿಜಾರಿಗೆ ಅಡ್ಡಲಾಗಿ ಬಿತ್ತನೆ ಮಾಡುವುದು ಹಾಗೂ ಈಗಾಗಲೇ ಹೊಲದಲ್ಲಿ ಇರುವ ಬದುಗಳ ಸುಸ್ಥಿತಿಯನ್ನು ಕಾಪಾಡುವುದು. ಈ ಕ್ರಮಗಳನ್ನು ಪ್ರತಿ ವರ್ಷವೂ ಕೈಗೊಳ್ಳುವುದು ಸೂಕ್ತ.
       </li>
       <li style="font-size: 5.5pt; text-align: justify">
           <strong>ಮಣ್ಣಿನ ಸವಕಳಿ ತಡೆಗಟ್ಟುವುದು:</strong> ಬದುಗಳ ನಿರ್ಮಾಣ, ಭೂ-ಮೇಲ್ಮೈಯನ್ನು ಸಮತಟ್ಟು ಮಾಡುವುದು, ಇಳಿಜಾರಿಗೆ ಅಡ್ಡಲಾಗಿ ಬಿತ್ತನೆ, ಕಂದಕ ಬದು ನಿರ್ಮಾಣ, ಹೊದಿಕೆ ಬೆಳೆ ಹಾಗೂ ಸಾವಯವ ಹೊದಿಕೆ ಹಾಕುವುದರಿಂದ ಸವಕಳಿಯ ಪ್ರಮಾಣವನ್ನು ಕಡಿಮೆ ಮಾಡಬಹುದು, ಸದಾಕಾಲ ಮಣ್ಣಿನ ಹೊದಿಕೆ ಇರುವಂತೆ ನೋಡಿಕೊಳ್ಳುವುದು ಸೂಕ್ತ. ಬದುಗಳ ಸುಸ್ಥಿತಿಯನ್ನು ಪ್ರತಿ ವರ್ಷವೂ ಕಾಪಾಡುವುದು.
       </li>
       <li style="font-size: 5.5pt; text-align: justify">
           <strong>ನೀರಿನ ಲಭ್ಯತೆಯ ಸಾಮರ್ಥ್ಯವನ್ನು ಸುಧಾರಿಸುವುದು:</strong> ಅಧಿಕ ಪ್ರಮಾಣದಲ್ಲಿ ಸಾವಯವ ಗೊಬ್ಬರಗಳ ಬಳಕೆ, ಮಣ್ಣು-ನೀರಿನ ಸಂರಕ್ಷಣೆ ಹಾಗೂ ಮರುಳು ಮಣ್ಣಿಗೆ ಕಪ್ಪು ಜೇಡಿ ಮಣ್ಣು ಬೆರೆಸುವುದು, ಇತ್ಯಾದಿ ಕ್ರಮಗಳನ್ನು ಅಳವಡಿಸುವುದು.
       </li>
       <li style="font-size: 5.5pt; text-align: justify">
           <strong>ಮಣ್ಣು ಮತ್ತು ನೀರಿನ ಸಂರಕ್ಷಣಾ ಯೋಜನೆ:</strong> ಶಿಫಾರಸ್ಸು ಮಾಡಿದ ಮಣ್ಣು ಹಾಗೂ ನೀರಿನ ಸಂರಕ್ಷಣೆ ಮತ್ತು ಒಳ ಚರಂಡಿ ಉಪಚಾರ ಯೋಜನೆಗಳನ್ನು ಅಳವಡಿಸುವುದು, ಹಸಿರೆಲೆ ಗೊಬ್ಬರ ಅಥವಾ ಸಾವಯವ ಹೊದಿಕೆ ಮಾಡುವುದರ ಜೊತೆಗೆ ಇತ್ಯಾದಿ ಶಿಫಾರಸ್ಸು ಮಾಡಿದ ಕ್ರಮಗಳನ್ನು ಅಳವಡಿಸಬೇಕು ಹಾಗೂ ಮಣ್ಣು ಮತ್ತು ನೀರಿನ ಸಂರಕ್ಷಣಾ ಕ್ರಮಗಳನ್ನು ಪ್ರತಿ ವರ್ಷವೂ ಕೈಗೊಳ್ಳಬೇಕು.
       </li>
       <li style="font-size: 5.5pt; text-align: justify">
           <strong>ಮಣ್ಣಿನಲ್ಲಿ ಸಾವಯವ ಇಂಗಾಲ:</strong> 0.5% ಕ್ಕಿಂತ ಅಧಿಕವಿರುವಂತೆ ಸಾಕಷ್ಟು ಪ್ರಮಾಣದಲ್ಲಿ ಸಾವಯವ ಗೊಬ್ಬರಗಳನ್ನು ಒದಗಿಸಬೇಕು ಹಾಗೂ ಪ್ರತಿ ವರ್ಷವೂ ಒದಗಿಸುವುದು ಅವಶ್ಯಕ.
       </li>
       <li style="font-size: 5.5pt; text-align: justify">
           <strong>ಮಣ್ಣಿನ ರಸಸಾರ:</strong> 6.5 ಕ್ಕಿಂತ ಕಡಿಮೆ ಇದ್ದರೆ ಸುಟ್ಟ ಸುಣ್ಣವನ್ನು ಭೂಮಿಗೆ ನಿರ್ದಿಷ್ಟ ಶಿಫಾರಸ್ಸಿನಂತೆ ಬೆರೆಸಬೇಕು. ಮಣ್ಣಿನ ರಸಸಾರ 8.5ಕ್ಕಿಂತ ಹೆಚ್ಚಿದ್ದರೆ ಶಿಫಾರಸ್ಸು ಮಾಡಿದ ಪ್ರಮಾಣದಲ್ಲಿ ಜಿಪ್ಸ್‌ಮ್‌ ಅನ್ನು ಮಣ್ಣಿಗೆ ಸೇರಿಸಬೇಕು. 2 ವರ್ಷದ ನಂತರ ಪುನಃ ಪರೀಕ್ಷಿಸಿ ಅಗತ್ಯಾನುಸಾರ ಸೂಕ್ತ ಬದಲಾವಣೆ ಮಾಡಬೇಕು.
       </li>
       <li style="font-size: 5.5pt; text-align: justify">
             <strong>ಪ್ರಧಾನ ಪೋಷಕಾಂಶಗಳ ಪ್ರಮಾಣ(N:P:K)</strong> ಕಡಿಮೆ ಇದ್ದಲ್ಲಿ. ಯಾವುದೇ ಒಂದು ಬೆಳೆಗೆ ನಿರ್ದಿಷ್ಟವಾಗಿ ಶಿಫಾರಸ್ಸು ಮಾಡಿದ ಪ್ರಮಾಣಕ್ಕಿಂತ, ಶೇಕಡಾ 25 ರಷ್ಟು ಹೆಚ್ಚಿನ ಪ್ರಮಾಣದಲ್ಲಿ ಕೊಡಬೇಕು. ಅಂದರೆ ಉದಾಹರಣೆಗೆ ಮೆಕ್ಕೆ ಜೋಳಕ್ಕೆ  ಹೆಕ್ಟೇರ್‌ಗೆ 100 ಕೆ.ಜಿ  ಸಾರಜನಕ ಶಿಫಾರಸ್ಸು ಮಾಡಿದರೆ, ಸಾರಜನಕ ಕಡಿಮೆ   ಇರುವ   ಭೂಮಿಯಲ್ಲಿ 125 ಕೆ ಜಿ  ಹೆಕ್ಟೇರ್‌ಗೆ     ಹಾಕಬೇಕಾಗುತ್ತದೆ. ಅದೆರೀತಿ 	ಪ್ರಧಾನ ಪೋಷಕಾಂಶಗಳು ಮಣ್ಣಿನಲ್ಲಿ ಹೆಚ್ಚು   ಇದ್ದರೆ,   ಶೇಕಡಾ 25ರಷ್ಟು ಕಡಿಮೆ   ಹಾಕಬೇಕು.   ಅಂದರೆ   100 ಕೆ.ಜಿ ಶಿಫಾರಸ್ಸು  ಮಾಡಿದ್ದರೆ  75ಕೆ.ಜಿ ಒದಗಿಸಿದರೆ ಸಾಕು,  ಇದೇ  ಕ್ರಮವನ್ನು ರಂಜಕ ಹಾಗೂ ಪೊಟ್ಯಾಷ್ ಗೂ ಅನುಸರಿಸಬೇಕು. 

       </li>
       <li style="font-size: 5.5pt; text-align: justify">
                 ಹೆಚ್ಚಿನ ಮಾಹಿತಿಗಳನ್ನು sujala3lri.karnataka.gov.in ಜಾಲತಾಣದಲ್ಲಿ ಪಡೆಯಬಹುದು
       </li>
    </ul>
 
        <p style="border: 2px solid #000; text-align: center; height:auto; background-color: #F9D4F1;font-size: 7pt;"><strong>ರೈತ ಸಹಾಯವಾಣಿ ಕೇಂದ್ರಗಳು:- ರೈತ ಸಹಾಯವಾಣಿ - 1800-425-3553, ವಾಣಿಜ್ಯ ಮಿತ್ರ - 92433 45433, ತೋಟಗಾರಿಕೆ ಸಹಾಯವಾಣಿ - 1800-425-7910, ಮತ್ತು, ಕಷಿ ಮಾರಾಟಿ ಬಾಂಕಿ - 1800-425-1552</strong></p>
    </div>
''')

# Insert text into the second textbox
para2 = textBox2.Body.AddParagraph()
para2.Format.LineSpacing = 0
para2.line_spacing = 0
para2.AppendHTML(f'''
    <h5 style="text-align: center">ಕರ್ನಾಟಕ ಸರ್ಕಾರ</h5>
    <table style="width:100%">
        <tr>
            <th><img src="wordbank.png" style="height:1.5cm ; width: 2cm;"/></th>
            <th><img src="wdd.png" style="height:1.5cm ; width: 1.6cm;"/></th>
            <th><img src="gok.png" style="height:1.5cm ; width: 1.6cm;"/></th>
            <th><img src="nbss.png" style="height:1.5cm ; width: 1.6cm;"/></th>
            <th><img src="icar.png" style="height:1.5cm ; width: 1.6cm;"/></th>
        </tr>
    </table>
    <p style="text-align: center; font-size:8pt;">
       <strong> REWARD </br>
        ಜಲಾನಯನ ಅಭಿವೃದ್ಧಿ ಇಲಾಖೆ</br>
        ಮತ್ತು ರಾಷ್ಟ್ರೀಯ ಮಣ್ಣು ಸರ್ವೇಕ್ಷಣಾ ಮತ್ತು ಭೂ ಬಳಕೆ ನಿಯೋಜನೆ ಸಂಸ್ಥೆ</br>
        ಪ್ರಾದೇಶಿಕ ಕೇಂದ್ರ , ಹೆಬ್ಬಾಳ  ಬೆಂಗಳೂರು-560024</br>
        ಸಂಪರ್ಕಿಸಿ: ಇ-ಮೇಲ್: nbssgis@gmail.com </strong></br>
    </p>
    <p style="text-align: center"><img src="lricrd.png" style="height:1cm ; width:7cm;"/> </br></p>
    <div style="align:center;">
    <table style="width: 100%; border: 0.5px solid black; border-collapse: collapse; font-size: 6.5pt; line-height: 0;">
        <tr style="height:0.95pt;line-height: 0;">
            <td style="margin:0;padding: 0; text-align: left;height:0.95pt;line-height: 0;"><strong>ರೈತರ ಹೆಸರು</strong></td>
            <td style="margin:0;padding: 0; text-align: left;height:0.95pt;line-height: 0;">{Farmer_name}</td>
        </tr>
        <tr>
            <td style="padding: 0; text-align: left;"><strong>ಲಿಂಗ (ಗಂಡು / ಹೆಣ್ಣು)</strong></td>
            <td style="padding: 0; text-align: left;">{Gender}</td>
        </tr>
        <tr>
            <td style="padding: 0; text-align: left;"><strong>ಕಿರುಜಲಾನಯನ ಪ್ರದೇಶದ ಹೆಸರು</strong></td>
            <td style="padding: 0; text-align: left;">{MWS}</td>
        </tr>
        <tr>
            <td style="padding: 0; text-align: left;"><strong>ವಿಳಾಸ</strong></td>
            <td style="padding: 0; text-align: left;">{Address}</td>
        </tr>
        <tr>
            <td style="padding: 0; text-align: left;"><strong>ಮಣ್ಣಿನ ಮಾದರಿ  ವರ್ಷ</strong></td>
            <td style="padding: 0; text-align: left;">{soil_year}</td>
        </tr>
        <tr>
            <td style="padding: 0; text-align: left;"><strong>ಸರ್ವೇ ಸಂಖ್ಯೆ / ಹಿಸ್ಸಾ ಸಂಖ್ಯೆ</strong></td>
            <td style="padding: 0; text-align: left;">{Survey_hissa}</td>
        </tr>
        <tr>
            <td style="padding: 0; text-align: left;"><strong>ಕ್ಷೇತ್ರದ ವಿಸ್ತೀರ್ಣ(ಎಕರೆ / ಗುಂಟೆ)</strong></td>
            <td style="padding: 0; text-align: left;">{Area}</td>
        </tr>
        <tr>
            <td style="padding: 0; text-align: left;"><strong>ವಾರ್ಷಿಕ ಸರಾಸರಿ ಮಳೆ (ಮಿ.ಮೀ.)</strong></td>
            <td style="padding: 0; text-align: left;">{Annual_rain_Fall}</td>
        </tr>
        <tr>
            <td colspan="2" style="padding: 0; text-align: left;"><strong>*ಸೂಚನೆ: ಸರ್ವೇ ನಂಬರ್‌ನ ಒಟ್ಟು ವಿಸ್ತೀರ್ಣ</strong></td>
        </tr>
    </table>
    </br>
    </div>
    
    <table style="width: 100%; border: 0.5px solid black; border-collapse: collapse; font-size: 6.5pt; line-height: 0;">
        <tr style="height: 1px;">
            <td colspan="2" style="padding: 0; text-align: center;"><strong>ಭೂಮೇಲ್ಮೈ ಲಕ್ಷಣ ಮತ್ತು ಮಣ್ಣಿನ ಗುಣಧರ್ಮಗಳ ವಿವರಗಳು</strong></td>
        </tr>
        <tr style="height: 10px;">
            <td style="padding: 0; text-align: left; line-height: 1;"><strong>ಮಣ್ಣಿನ ಆಳ</strong></td>
            <td style="padding: 0; text-align: left; line-height: 1;">{soil_Dept}</td>
        </tr>
        <tr style="height: 10px;">
            <td style="padding: 0; text-align: left; line-height: 1;"><strong>ಮಣ್ಣಿನ ಕಣಗಾತ್ರ</strong></td>
            <td style="padding: 0; text-align: left; line-height: 1;">{Soil} </td>
        </tr>
        <tr style="height: 10px;">
            <td style="padding: 0; text-align: left; line-height: 1;"><strong>ಮಣ್ಣಿನ ಗರಸಿನ ಪ್ರಮಾಣ (ಶೇ)</strong></td>
            <td style="padding: 0; text-align: left; line-height: 1;">{Soil}</td>
        </tr>
        <tr style="height: 10px;">
            <td style="padding: 0; text-align: left; line-height: 1;"><strong>ಮಣ್ಣಿನ ಇಳಿಜಾರು (ಶೇ)</strong></td>
            <td style="padding: 0; text-align: left; line-height: 1;">{Soil}</td>
        </tr>
        <tr style="height: 10px;">
            <td style="padding: 0; text-align: left; line-height: 1;"><strong>ಮಣ್ಣಿನ ಸವಕಳಿ</strong></td>
            <td style="padding: 0; text-align: left; line-height: 1;">{Soil}</td>
        </tr>
        <tr style="height: 10px;">
            <td style="padding: 0; text-align: left; line-height: 1;"><strong>ಭೂ ಸಾಮರ್ಥ್ಯ </strong></td>
            <td style="padding: 0; text-align: left; line-height: 1;">{Soil}</td>
        </tr>
        <tr style="height: 10px;">
            <td style="padding: 0; text-align: left; line-height: 1;"><strong>ಮಣ್ಣಿನಲ್ಲಿ ನೀರು ಹಿಡಿದಿಟ್ಟುಕೊಳ್ಳುವ ಸಾಮರ್ಥ್ಯ</strong></td>
            <td style="padding: 0; text-align: left; line-height: 1;">{Soil}</td>
        </tr>
        <tr style="height: 10px;">
            <td style="padding: 0; text-align: left; line-height: 1;"><strong>ಮಣ್ಣು ಮತ್ತು ನೀರಿನ ಸಂರಕ್ಷಣಾ  ಯೋಜನೆ</strong></td>
            <td style="padding: 0; text-align: left; line-height: 1;">{Soil}</td>
        </tr>
        <tr style="height: 10px;">
            <td style="padding: 0; text-align: left; line-height: 1;"><strong>ಸಾಂಪ್ರದಾಯಿಕ ಮಣ್ಣಿನ ಹೆಸರು</strong></td>
            <td style="padding: 0; text-align: left; line-height: 1;">{Soil}</td>
        </tr>
        
    </table>
''')

#page 2
new_section = doc.AddSection()
paragraph = new_section.AddParagraph()
# Insert the first textbox and set its wrapping style
textBox3 = paragraph.AppendTextBox(width=350, height=472)
textBox3.Format.TextWrappingStyle = TextWrappingStyle.Inline
textBox3.Format.HorizontalOrigin = HorizontalOrigin.LeftMarginArea
textBox3.Format.HorizontalPosition = 10

# Insert the second textbox and set its wrapping style and position
textBox4 = paragraph.AppendTextBox(width=350, height=472)
textBox4.Format.TextWrappingStyle = TextWrappingStyle.Inline
textBox4.Format.HorizontalOrigin = HorizontalOrigin.RightMarginArea
textBox4.Format.HorizontalPosition = -50


# Insert text into the second textbox
para3 = textBox3.Body.AddParagraph()
para3.AppendHTML('''
                <table style="width: 100%; border: 1px solid black; border-collapse: collapse; font-size: 7pt;">
                    <tr style="height: 10px;">
                        <td colspan="2"style="padding: 0; text-align: left; line-height: 1;"><strong>ಪ್ರಯೋಗಾಲಯದ ಹೆಸರು ಮತ್ತು ವಿಳಾಸ:</strong></td>
                        <td colspan="3"style="padding: 0; text-align: left; line-height: 1;">ರಾಷ್ಟ್ರೀಯ ಮಣ್ಣು ಸರ್ವೇಕ್ಷಣಾ ಮತ್ತು ಭೂ ಬಳಕೆ ನಿಯೋಜನೆ ಸಂಸ್ಥೆ ಪ್ರಾದೇಶಿಕ ಕೇಂದ್ರ, ಹೆಬ್ಬಾಳ ಬೆಂಗಳೂರು-560024</td>
                    </tr>
                    <tr style="height: 1px;">
                        <td colspan="5" style="padding: 0; text-align: center;"><strong>ಮಣ್ಣು ಪರೀಕ್ಷಾ ವರದಿ (Soil Test Results)</strong></td>
                    </tr>

                    <tr style="height: 10px;padding: 0;">
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ಕ್ರ.ಸಂ</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ನಿಯಾತಾಂಕ (Parameter)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ಪರೀಕ್ಷೆ ಮೌಲ್ಯ (Test Value)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ಘಟಕ  ((Unit)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ಷರಾ (Remarks)</strong></td>
                    </tr>
                    <tr style="height: 10px;">
                        <td style="padding: 0; text-align: center; line-height: 1;">01</td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ರಸಸಾರ (pH)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{test_Value_pH}</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">-</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{Remarks_pH}</td>
                    </tr>
                    <tr style="height: 10px;">
                        <td style="padding: 0; text-align: center; line-height: 1;">02</td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ವಿದ್ಯುತ್ ವಾಹಕತೆ (EC)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{test_Value_EC}</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">ಡೆ.ಸೈ./ಮೀ.</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{Remarks_EC}</td>
                    </tr>
                    <tr style="height: 10px;">
                        <td style="padding: 0; text-align: center; line-height: 1;">03</td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ಸಾವಯವ ಇಂಗಾಲ   (OC)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{test_Value_OC}</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">ಶೇಕಡಾ</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{Remarks_OC}</td>
                    </tr>
                    <tr style="height: 10px;">
                        <td style="padding: 0; text-align: center; line-height: 1;">04</td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ಲಭ್ಯ ಸಾರಜನಕ  (N)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{test_Value_N}</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">ಕಿ.ಗ್ರಾಂ/ಹೆ</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{Remarks_N}</td>
                    </tr>
                    <tr style="height: 10px;">
                        <td style="padding: 0; text-align: center; line-height: 1;">05</td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ಲಭ್ಯ ರಂಜಕ (P2O5)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{test_Value_p2o5}</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">ಕಿ.ಗ್ರಾಂ/ಹೆ</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{Remarks_p2o5}</td>
                    </tr>
                    <tr style="height: 10px;">
                        <td style="padding: 0; text-align: center; line-height: 1;">06</td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ಲಭ್ಯ  ಪೊಟ್ಯಾಶ್ (K2O)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{test_Value_pH==k2o}</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">ಕಿ.ಗ್ರಾಂ/ಹೆ</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{Remarks_k2o}</td>
                    </tr>
                    <tr style="height: 10px;">
                        <td style="padding: 0; text-align: center; line-height: 1;">07</td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ಲಭ್ಯ ಗಂಧಕ (S)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{test_Value_S}</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">ಪಿ.ಪಿ.ಎಂ</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{Remarks_S}</td>
                    </tr>
                    <tr style="height: 10px;">
                        <td style="padding: 0; text-align: center; line-height: 1;">08</td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ಲಭ್ಯ ಸತು (Zn)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{test_Value_zn}</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">ಪಿ.ಪಿ.ಎಂ</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{Remarks_zn}</td>
                    </tr>
                    <tr style="height: 10px;">
                        <td style="padding: 0; text-align: center; line-height: 1;">09</td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ಲಭ್ಯ ಬೋರಾನ್  (B)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{test_Value_B}</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">ಪಿ.ಪಿ.ಎಂ</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{Remarks_B}</td>
                    </tr>
                    
                    <tr style="height: 10px;">
                        <td style="padding: 0; text-align: center; line-height: 1;">10</td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ಲಭ್ಯ ಕಬ್ಬಿಣ  (Fe)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{test_Value_fe}</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">ಪಿ.ಪಿ.ಎಂ</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{Remarks_fe}</td>
                    </tr>
                    <tr style="height: 10px;">
                        <td style="padding: 0; text-align: center; line-height: 1;">11</td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ಲಭ್ಯ ಮ್ಯಾಂಗನೀಸ್ (Mn)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{test_Value_mn}</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">ಪಿ.ಪಿ.ಎಂ</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{Remarks_mn}</td>
                    </tr>
                    <tr style="height: 10px;">
                        <td style="padding: 0; text-align: center; line-height: 1;">12</td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ಲಭ್ಯ ತಾಮ್ರ (Cu)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{test_Value_cu}</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">ಪಿ.ಪಿ.ಎಂ</td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{Remarks_cu}</td>
                    </tr>
                    <tr style="height: 1px;">
                        <td colspan="5" style="padding: 0; text-align: center;"><strong>ಸೂಚನೆ: ಪ್ರಸ್ತುತ ಮಣ್ಣಿನ ಫಲವತ್ತತೆ ಸ್ಥಿತಿಯನ್ನು ಜಿ.ಪಿ.ಎಸ್ ಆಧಾರಿತ ಪ್ರತಿ 320 ಮೀಟರ್‍ಗಳುಳ್ಳ ದತ್ತಾಂಶದಿಂದ ಅಂದಾಜಿಸಲಾಗಿದೆ ಹಾಗೂ ಇದು ಮುಂದಿನ ಮೂರು ವರ್ಷಕ್ಕೆ ಅನ್ವಯಿಸುತ್ತದೆ. ಸಂಪೂರ್ಣ ವಿವರಗಳಿಗಾಗಿ ದಯವಿಟ್ಟು ಕಿರು ಜಲಾನಯನ ಪ್ರದೇಶದ  ಎಲ್.ಆರ್‍.ಐ. ವರದಿ ಅಥವಾ ಅಟ್ಲಾಸ್ ಗಳನ್ನು ಪರಿಶೀಲಿಸಿ.</strong></td>
                    </tr>
                    <tr style="height: 1px;">
                        <td colspan="5" style="padding: 0; text-align: center;"><img src="colourcode.png" style="height:0.32cm;width:12cm;"/></strong></td>
                    </tr>
                    
                    
                </table>
                </br>
                <table style="width: 100%; border: 1px solid black; border-collapse: collapse; font-size: 7pt;">
                    
                    <tr style="height: 1px;">
                        <td colspan="4" style="padding: 0; text-align: center;"><strong>ದ್ವಿತೀಯ ಮತ್ತು ಲಘು ಪೋಷಕಾಂಶಗಳ ಕೊರತೆ ಇರುವ ಮಣ್ಣಿಗೆ ಶಿಫಾರಸ್ಸು</strong></td>
                    </tr>

                    <tr style="height: 10px;">
                        <th style="padding: 0; text-align: center; line-height: 1;"><strong>ಕ್ರ.ಸಂ</strong></th>
                        <th style="padding: 0; text-align: center; line-height: 1;"><strong>ನಿಯಾತಾಂಕ (Parameter)</strong></th>
                        <th style="padding: 0; text-align: center; line-height: 1;"><strong>ಗೊಬ್ಬರ</strong></th>
                        <th rowspan="7" style="padding: 0; text-align: center; line-height: 1;font-size: 6pt; "><strong>ಲಘು ಪೋಷಕಾಂಶಗಳ ಬಳಕೆಯ ಶಿಫಾರಸ್ಸಿನ ಪ್ರಮಾಣವು ಬೆಳೆಯಿಂದ ಬೆಳೆಗೆ ಭಿನ್ನವಾಗಿರುತ್ತದೆ, ಹತ್ತಿರದ ರೈತ ಸಂಪರ್ಕ ಕೇಂದ್ರ  ಅಥವಾ ಕೃಷಿ ವಿಜ್ಞಾನ ಕೇಂದ್ರ ವಿಜ್ಞಾನಿಗಳೊಂದಿಗೆ ಸಮಾಲೋಚಿಸಿ ಬಳಕೆಯ ಪ್ರಮಾಣವನ್ನು ನಿರ್ಧರಿಸುವುದು ಸೂಕ್ತ.</strong></td>
                    </tr>
                    <tr style="height: 10px;">
                        <td style="padding: 0; text-align: center; line-height: 1;">01</td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ಗಂಧಕ (S)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{Value_sxzcxzcxcczxcx}</td>
                        
                    </tr>
                    <tr style="height: 10px;">
                        <td style="padding: 0; text-align: center; line-height: 1;">02</td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ಬೋರಾನ್ (B)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{Value_b}</td>

                    </tr>
                    <tr style="height: 10px;">
                        <td style="padding: 0; text-align: center; line-height: 1;">03</td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ಸತು (Zn)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{Value_zn}</td>

                    </tr>
                    <tr style="height: 10px;">
                        <td style="padding: 0; text-align: center; line-height: 1;">04</td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ಕಬ್ಬಿಣ (Fe)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{Value_fe}</td>

                    </tr>
                    <tr style="height: 10px;">
                        <td style="padding: 0; text-align: center; line-height: 1;">05</td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ಮ್ಯಾಂಗನೀಸ್ (Mn)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{Value_mn}</td>

                    </tr>
                    <tr style="height: 10px;">
                        <td style="padding: 0; text-align: center; line-height: 1;">06</td>
                        <td style="padding: 0; text-align: center; line-height: 1;"><strong>ತಾಮ್ರ (Cu)</strong></td>
                        <td style="padding: 0; text-align: center; line-height: 1;">{Value_cu}</td>

                    </tr>
                   
                </table>
   
''')

# Insert text into the second textbox
para4 = textBox4.Body.AddParagraph()
para4.AppendHTML(''' 
                      <table style="width: 100%; border: 1px solid black; border-collapse: collapse; font-size: 6pt;";>
                          <tr>
                            <th>ಪೋಷಕಾಂಶಗಳು</th>
                            <th>ಅತಿ ಕಡಿಮೆ ಫಲವತ್ತತೆ</th>
                            <th>ಕಡಿಮೆ ಫಲವತ್ತತೆ</th>
                            <th>ಮಧ್ಯಮ ಫಲವತ್ತತೆ</th>
                            <th>ಅಧಿಕ ಫಲವತ್ತತೆ</th>
                            <th>ಅತ್ಯಧಿಕ ಫಲವತ್ತತೆ</th>
                          </tr>
                          <tr>
                            <td>ಸಾವಯವ ಇಂಗಾಲ (ಶೇ)</td>
                            <td>< 0.25</td>
                            <td>0.25 - 0.5</td>
                            <td>0.5 - 0.75</td>
                            <td>0.75 - 1.00</td>
                            <td>> 1.00</td>
                          </tr>
                          <tr>
                            <td>ಲಭ್ಯ ಸಾರಜನಕ (ಕಿ.ಗ್ರಾಂ/ಹೆ)</td>
                            <td>< 140</td>
                            <td>140 - 280</td>
                            <td>280 - 560</td>
                            <td>560 - 700</td>
                            <td>> 700</td>
                          </tr>
                          <tr>
                            <td>ಲಭ್ಯ ರಂಜಕ (ಕಿ.ಗ್ರಾಂ/ಹೆ)</td>
                            <td>< 11.5</td>
                            <td>11.5 - 23</td>
                            <td>23 - 57</td>
                            <td>57 - 91</td>
                            <td>> 91</td>
                          </tr>
                          <tr>
                            <td>ಲಭ್ಯ ಪೊಟ್ಯಾಶ್ (ಕಿ.ಗ್ರಾಂ/ಹೆ)</td>
                            <td>< 72</td>
                            <td>72 - 145</td>
                            <td>145 - 337</td>
                            <td>337 - 675</td>
                            <td>> 675</td>
                          </tr>
                          <tr>
                            <td>ಲಭ್ಯ ಗಂಧಕ (ಪಿ.ಪಿ.ಎಂ)</td>
                            <td>-</td>
                            <td>< 10</td>
                            <td>10 - 20</td>
                            <td>> 20</td>
                            <td>-</td>
                          </tr>
                          <tr>
                            <td>ಲಭ್ಯ ಸತು (ಪಿ.ಪಿ.ಎಂ)</td>
                            <td>-</td>
                            <td>< 0.6</td>
                            <td>> 0.6</td>
                            <td>-</td>
                            <td>-</td>
                          </tr>
                          <tr>
                            <td>ಲಭ್ಯ ಕಬ್ಬಿಣ (ಪಿ.ಪಿ.ಎಂ)</td>
                            <td>-</td>
                            <td>< 4.5</td>
                            <td>> 4.5</td>
                            <td>-</td>
                            <td>-</td>
                          </tr>
                          <tr>
                            <td>ಲಭ್ಯ ತಾಮ್ರ (ಪಿ.ಪಿ.ಎಂ)</td>
                            <td>-</td>
                            <td>< 0.2</td>
                            <td>> 0.2</td>
                            <td>-</td>
                            <td>-</td>
                          </tr>
                          <tr>
                            <td>ಲಭ್ಯ ಮ್ಯಾಂಗನೀಸ್ (ಪಿ.ಪಿ.ಎಂ)</td>
                            <td>-</td>
                            <td>< 1.0</td>
                            <td>> 1.0</td>
                            <td>-</td>
                            <td>-</td>
                          </tr>
                          <tr>
                            <td>ಲಭ್ಯ ಬೋರಾನ್ (ಪಿ.ಪಿ.ಎಂ)</td>
                            <td>-</td>
                            <td>< 0.5</td>
                            <td>0.5 - 1.0</td>
                            <td>> 1.0</td>
                            <td>-</td>
                          </tr>
                    </table>
                    </br>
                    <table style="width: 100%; border: 1px solid black; border-collapse: collapse; font-size: 6pt;">
                        <tr>
                            <td colspan="4" style="font-size:8pt;text-align:center"><strong>ಭೂ ಸಂಪನ್ಮೂಲ ಮಾಹಿತಿ ಆಧಾರದ ಮೇಲೆ ಸೂಚಿತ ಬೆಳೆ ಯೋಜನೆ</strong></td>
                        </tr>
                        <tr>
                          <th style="padding: 0; text-align: center; line-height: 1;"><strong>ಸೂಕ್ತತೆ</strong></th>
                          <th style="padding: 0; text-align: center; line-height: 1;"><strong>ಸೂಕ್ತವಾದ ಬೆಳೆಗಳು</strong></th>
                          <th style="padding: 0; text-align: center; line-height: 1;"><strong>ಮಿತಿಗಳು</strong></th>
                          <th style="padding: 0; text-align: center; line-height: 1;"><strong>ಸೂಚಿಸಲಾದ ನಿರ್ವಹಣಾ ಪದ್ಧತಿಗಳು</strong></th>
                        <tr>
                        <tr>
                            <td>ಹೆಚ್ಚು ಸೂಕ್ತ </td>
                            <td>- </td>
                            <td>- </td>
                            <td rowspan=7>ಇಳಿಜಾರು >10 ಇದ್ದಲ್ಲಿ  ತೋಟಗಾರಿಕೆ ಬೆಳೆಗಳಿಗೆ ಜಗತಿ ಕಟ್ಟೆ ಮತ್ತು ಬಾಹ್ಯರೇಖಾ ಕಂದಕಗಳನ್ನು ಅಳವಡಿಸುವುದು ಹಾಗೂ ಕ್ಷೇತ್ರಿಯ ಬೆಳೆಗಳಿಗೆ ಸಮಪಾತಳಿ ಬದು ಮತ್ತು ಇಳಿಜಾರಿಗೆ ಅಡ್ಡಲಾಗಿ ಉಳುಮೆ ಮಾಡುವುದು</td>
                        </tr>
                        
                        <tr>
                            <td>ಸಾಧಾರಣ ಸೂಕ್ತ </td>
                            <td>- </td>
                            <td>- </td>
                        </tr>
                        <tr>
                            <td>ಅಲ್ಪ ಸೂಕ್ತ</td>
                            <td>ಬೀಟ್ರೂಟ್, ಅವರೆ, ಸೇವಂತಿಗೆ, ಚೆಂಡುಹೂವು, ಈರುಳ್ಳಿ, ಟೊಮ್ಯಾಟೊ, ಬದನೆ, ಅಲಸಂದೆ, ಶೇಂಗಾ, ಮೆಕ್ಕೆಜೋಳ, ಭೀಮ ಬಿದಿರು, ಹೂಕೋಸು, ರಾಗಿ</td>
                            <td>ಇಳಿಜಾರು </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>ತೇಗದ ಮರ, ಸಿಲ್ವರ್‍ ಓಕ್, ಹೆಬ್ಬೇವು</td>
                            <td>ಬೇರಿನ ಬೆಳವಣಿಗೆ ತಡೆಯುವುವಿಕೆ </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>ಸೀಬೆ, ಪರಂಗಿ, ಸೂರ್ಯಕಾಂತಿ, ತೊಗರಿ</td>
                            <td>ಬೇರಿನ ಬೆಳವಣಿಗೆ ತಡೆಯುವುವಿಕೆ , ಇಳಿಜಾರು </td>
                        </tr>
                        <tr>
                            <td>ಪ್ರಸ್ತುತ ಸೂಕ್ತವಲ್ಲ</td>
                            <td>ತಗ್ಗು ಪ್ರದೇಶದ ಭತ್ತ</td>
                            <td>- </td>
                        </tr>
                        <tr>
                            <td> </td>
                            <td>ಮಾವು</td>
                            <td>ಬೇರಿನ ಬೆಳವಣಿಗೆ ತಡೆಯುವುವಿಕೆ  </td>
                        </tr>
                        <tr>
                            <td colspan=4>ಸೂಚನೆ: ತೋಟಗಾರಿಕಾ ಬೆಳೆಗಳ ಸೂಕ್ತತೆ  ನೀರಾವರಿಯ ಲಭ್ಯತೆಯನ್ನು ಅವಲಂಭಿಸಿದೆ  </td>
                        </tr>
                             
                    </table>
                    <h5>ವಿತರಣೆ ಮಾಡಿದ ತಿಂಗಳು ಮತ್ತು ವರ್ಷ :     ಮೇ- 2024</h5>
                 
    
''')






'''--------------------------------------------------------------------- Data Convertion & Save ------------------------------------------------'''
# Save the document
doc.SaveToFile("output/AddTextBox.docx", FileFormat.Docx)

doc1 = dc("output/AddTextBox.docx")

for para in doc1.paragraphs:
    if para.text == "Evaluation Warning: The document was created with Spire.Doc for Python.":
        para.text = ""
    print(para.text)
    
doc1.save("output/LRI_Card.docx")




#convert('output/LRI_Card.docx', 'output/LRI_Card.pdf')
#pdfkit.from_file('output/AddTextBox.docx', 'output.pdf', options=options)
