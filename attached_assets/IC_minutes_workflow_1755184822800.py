#%%
import os
from dotenv import load_dotenv
load_dotenv()
import IC_functions
from IPython.display import display, Markdown
#%%
api_key=os.getenv("GOOGLE_API_KEY")
analyzer=IC_functions.ICMemoAnalyzer(api_key,model_name="gemini-2.5-flash-lite")
#
pdf_path = "To_DO_IC_Minutes/Numeris - Investment Committee IC II - Apr.2025.pdf"
os.path.exists(pdf_path)
#%%
results = analyzer.analyze_pdf(pdf_path)
#%%
ic_minutes= analyzer.create_IC_minutes(results)
#%%
analyzer.output_docx(ic_minutes,path="IC_minutes_Numeris.docx")
#%%
