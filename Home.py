import streamlit as st
from calendar import month_abbr
from datetime import datetime
import oracledb
import cx_Oracle
import sys

from PIL import Image
import google.generativeai as genai
from dotenv import load_dotenv
import io  # Import io module for byte conversion

import pandas as pd 
import os
from langchain_experimental.agents import create_csv_agent
from langchain.llms import OpenAI 
from langchain.document_loaders.csv_loader import CSVLoader

from langchain.retrievers import WikipediaRetriever
import json
from langchain_community.document_loaders import Docx2txtLoader
from langchain.indexes import VectorstoreIndexCreator


from unstructured.partition.pptx import partition_pptx
from langchain_community.document_loaders import UnstructuredPowerPointLoader
from langchain.chains import create_extraction_chain
import pptx
from langchain_openai import ChatOpenAI
import matplotlib.pyplot as plt

import os
import pinecone
from langchain.chains import RetrievalQA
from langchain.embeddings import OpenAIEmbeddings
from langchain.llms import OpenAI
from langchain.vectorstores import Pinecone
from langchain.chains import RetrievalQAWithSourcesChain
import streamlit as st

# LangChain data loaders imports
from langchain_community.document_loaders import PyPDFLoader
from langchain_community.document_loaders import Docx2txtLoader
from langchain_community.document_loaders import TextLoader
from langchain_community.document_loaders import CSVLoader
from langchain_community.document_loaders import UnstructuredPowerPointLoader
from langchain_community.document_loaders import UnstructuredExcelLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_community.vectorstores import Pinecone as LangPinecone
from langchain_openai import OpenAIEmbeddings
import cx_Oracle
import pandas as pd
import oracledb

load_dotenv()



st.set_page_config(layout='wide')

# OpenAI Api credentials


genai.configure(api_key=os.environ["api_key"])

model=genai.GenerativeModel('gemini-pro-vision')

### Conect with Headcount ATP Database

# Check if Oracle Client has already been initialized
if not cx_Oracle.clientversion():
    try:
        
        cx_Oracle.init_oracle_client(config_dir=r"D:\Langchain_Applications\Resources\Wallet_vectordb")
    except Exception as err:
        st.error(f"Error initializing Oracle client: {err}")
        sys.exit(1)

# path for excel file
xls_file = r"D:\Langchain_Applications\Resources\Report.xlsx"
output_csv = r"D:\Langchain_Applications\Resources\Report.csv"

# Path to the image file
image_path = r"D:\Langchain_Applications\Resources\currentYearActuals.png"

# path for document folder
wordFolderPath = r'D:\Langchain_Applications\Resources\wordDocFolder'

# path for pptx file
pptxFile = UnstructuredPowerPointLoader(r"D:\Langchain_Applications\Resources\Trends_and_Charts _2022.pptx")

# Read the XLS file using pandas and openpyxl as the engine
data = pd.read_excel(xls_file, engine='openpyxl')

# Save data as a csv file
data.to_csv(output_csv, index=False)

# Initialize OpenAI agent
openai_agent = create_csv_agent(OpenAI(temperature=0), output_csv, verbose=True)

# Streamlit application
if "visibility" not in st.session_state:
        st.session_state.visibility = "visible"
        st.session_state.disabled = False
col1, col2 = st.columns(2)

with col1:
# Add your local image with left margin
    st.sidebar.image("./images/company_logo.jpg")
    st.markdown(
        """
        <style>
        div[data-testid="stImage"] {
            float: left;
            margin-left: 1%;
            margin-top: -3%;
            width: 8%;
            position: fixed;

            z-index:1;

        }
        </style>
        """,
        unsafe_allow_html=True,
    )  

with col2:
    pass

st.markdown("""
<style>
.big-font {
    font-size:20px;
    font-weight:bold;
    float:left;
    margin-left: -2%;
    margin-top: -5%;           
}

</style>
""", unsafe_allow_html=True)

st.markdown('<p class="big-font">PROCUREMENT KPI DASHBOARD</p>', unsafe_allow_html=True)
    
    
# Get Graphical Dashboard for oraganization word document Data

def getMonthChart(folder_path, year, month, orgName):
    # Load multiple Word documents
    word_files = [os.path.join(folder_path, file) for file in os.listdir(folder_path) if file.endswith('.docx')]

    loaders = []
    for word_file in word_files:
        loader = Docx2txtLoader(word_file)
        loaders.append(loader)

    # Create index from loaders
    index = VectorstoreIndexCreator().from_loaders(loaders)
    
    # prompt written for get some data variables
    prompt = """
Extract monthly procurement cost as 
   - previous year
   - target cost
   - actual cost

give number withought comma
Output in JSON format:
{
   "previous_year": previous_year,
   "target_cost": target_cost,
   "actual_cost": actual_cost
}
"""
# Query to get data for selected oraganization for month and year
    query = f"Extract procurement costs for {orgName} in the month of {month}, including previous years, target, and actual for {month} {year}  "

    # Concatenate prompt with the main query
    main_query = query + prompt

    # Query the index
    responses = index.query(main_query)
    print(responses)

    # Parse the response into a dictionary
    try:
        response_dict = json.loads(responses)
        # Access values from the dictionary
        previous_year = response_dict.get("previous_year")
        target_cost = response_dict.get("target_cost")
        actual_cost = response_dict.get("actual_cost")

        # Print the values
        print("Previous Year:", previous_year)
        print("Target Cost:", target_cost)
        print("Actual Cost:", actual_cost)

        # Display the values using Streamlit
        monthProcurementData = {
            "name": ["Previous year Cost", "Target Cost", "Actual Cost"],
            "Procurement Cost": [previous_year, target_cost, actual_cost]
        }
    
        chart_data = pd.DataFrame(monthProcurementData)
        chart_data.set_index("name", inplace=True)
        #st.bar_chart(chart_data)

    except json.JSONDecodeError as e:
        print("Error parsing response:", e)

    return monthProcurementData


## Get metadata of the organization from wikipedia
def wikipediaMetadata(org):
   retriever = WikipediaRetriever()
   docs = retriever.get_relevant_documents(query=org)
   r=docs[0].metadata  # meta-information of the Document
   return r['summary']

# Get organization name from ATP Database
def getOrganizationName():
    # Database connection credintials for 'Headcount' ATP database from oracle cloud
    connection = oracledb.connect(
            config_dir=r"D:\Langchain_Applications\Resources\Wallet_vectordb",
            user=os.environ["USER"],
            password=os.environ["PASSWORD"],
            dsn="vectordb_low",
            wallet_location=r"D:\Langchain_Applications\Resources\Wallet_vectordb",
            wallet_password=os.environ["WALLET_PASSWORD"]
        )
    cursor = connection.cursor()
    # Execute query to fetch organization names
    query = "SELECT organization_name FROM ADMIN.organization"
    cursor.execute(query)
    organizations = cursor.fetchall()
    return organizations

# Prompt written for extract perticular data variables from data
format_template="""
Extract the  following  value

	Cost of Procurement	- 
	Cost Savings Per Month	-
	Total Worth of Inventory - 
	Procurement ROI - 
    Purchase Order Cycle Time  -
    Supplier Lead Time -
    Number of Suppliers -
   

    Output in json format:
    "Cost_of_Procurement": Cost_of_Procurement,
    "Cost_Savings_Per_Month": Cost_Savings_Per_Month,
    "Total_Worth_of_Inventory": Total_Worth_of_Inventory,
    "Procurement_ROI": Procurement_ROI,
    "Purchase_Order_Cycle_Time": Purchase_Order_Cycle_Time,
    "Supplier_Lead_Time": Supplier_Lead_Time,
    "Number_of_Suppliers": Number_of_Suppliers

"""

@st.cache_data

# function written for get projected procurement data from excle files
def getProcurementCost(year, month):
    query = f"extract the all data for coloum for {month} {year} {format_template} "
    res = openai_agent.run(query)
    try:

        # Convert the generated text to a JSON object
        json_output = json.loads(res)
    
        cost_of_procurement = json_output["Cost_of_Procurement"]
        cost_savings_per_month = json_output["Cost_Savings_Per_Month"]
        total_worth_of_inventory = json_output["Total_Worth_of_Inventory"]
        procurement_roi = json_output["Procurement_ROI"]
        purchase_order_cycle_time = json_output["Purchase_Order_Cycle_Time"]
        supplier_lead_time = json_output["Supplier_Lead_Time"]
        number_of_suppliers = json_output["Number_of_Suppliers"]

        # Create a dictionary containing the variables
        data = {
        "Metrics": ["Cost_of_Procurement", "Cost_Savings_Per_Month", "Total_Worth_of_Inventory", 
                "Procurement_ROI", "Purchase_Order_Cycle_Time", "Supplier_Lead_Time", 
                "Number_of_Suppliers"],
        "Values": [cost_of_procurement, cost_savings_per_month, total_worth_of_inventory, 
               procurement_roi, purchase_order_cycle_time, supplier_lead_time, 
               number_of_suppliers]
        }

        # Create a DataFrame
        df = pd.DataFrame(data)
        return df
    except:
        return res

## function written for get Actual procurement data from Image file
@st.cache_data
def getActualProcurementCost(year, month):
    input_query = f"Sumarise the coloum for {month} {year} {format_template} "

    # Get image parts from image path
    image_data = input_image_details(image_path)

    # Get response from the model
    response = get_gemini_response(input_prompt, image_data, input_query)

    try:

        # Convert the generated text to a JSON object
        json_output = json.loads(response)
    
        cost_of_procurement = json_output["Cost_of_Procurement"]
        cost_savings_per_month = json_output["Cost_Savings_Per_Month"]
        total_worth_of_inventory = json_output["Total_Worth_of_Inventory"]
        procurement_roi = json_output["Procurement_ROI"]
        purchase_order_cycle_time = json_output["Purchase_Order_Cycle_Time"]
        supplier_lead_time = json_output["Supplier_Lead_Time"]
        number_of_suppliers = json_output["Number_of_Suppliers"]

        # Create a dictionary containing the variables
        data = {
        "Metrics": ["Cost_of_Procurement", "Cost_Savings_Per_Month", "Total_Worth_of_Inventory", 
                "Procurement_ROI", "Purchase_Order_Cycle_Time", "Supplier_Lead_Time", 
                "Number_of_Suppliers"],
        "Values": [cost_of_procurement, cost_savings_per_month, total_worth_of_inventory, 
               procurement_roi, purchase_order_cycle_time, supplier_lead_time, 
               number_of_suppliers]
        }

        # Create a DataFrame
        df = pd.DataFrame([data])
        return df
    except:
       return response

def input_image_details(image_path):
    # Open the image file using PIL and convert it to bytes
    with open(image_path, 'rb') as f:
        bytes_data = f.read()

    # Create image parts dictionary
    image_parts = [{
        "mime_type": "image/png",  # Change the MIME type as per your image type
        "data": bytes_data
    }]

    return image_parts

def get_gemini_response(input_prompt, image_parts, prompt):
    # Generate response using the model
    response = model.generate_content([input_prompt, image_parts[0], prompt])
    return response.text

# Input prompt to get customized data
input_prompt = """
You are an expert in understanding Image document. We will upload an image as an excel and you will 
have to answer any questions based on the uploaded excel image and extract the data in json output.

Output in json format:
    "Cost_of_Procurement": Cost_of_Procurement,
    "Cost_Savings_Per_Month": Cost_Savings_Per_Month,
    "Total_Worth_of_Inventory": Total_Worth_of_Inventory,
    "Procurement_ROI": Procurement_ROI,
    "Purchase_Order_Cycle_Time": Purchase_Order_Cycle_Time,
    "Supplier_Lead_Time": Supplier_Lead_Time,
    "Number_of_Suppliers": Number_of_Suppliers

"""

# function written to Draw graphs on the data of pptx file
def pptxGraph(loader,period):

    data = loader.load()

    schema ={
  "type": " string",
  "properties": {
    "month": {"type": "string"},
    "target": {"type": "integer"},
    "actual": {"type": "integer"},
    "previous_year": {"type": "integer"},
    
  },
  "required": ["month", "target","actual","previous_year"]
  }
    
    # Run chain
    llm = ChatOpenAI(temperature=0, model="gpt-3.5-turbo")
    chain = create_extraction_chain(schema, llm)
    r=chain.invoke(data)
    print(r['text'])
    print(r)
    
    # logics to customize charts for quarterly, monthaly and half yearly
    if period == 'Quarterly':
        if quarterly == 'Q1':
            start = 0
            end = 3
        elif quarterly == 'Q2':   
            start = 3
            end = 6
        else:
            start = 0
            end = 6     
    else: 
        start = 0
        end = 6   
        
    monthList=[]
    for i in range(start, end):
       monthList.append(r['text'][i]["month"])   
    

    targetList=[]
    for i in range(start, end):
       targetList.append(r['text'][i]["target"])   
    

    actualList=[]
    for i in range(start, end):
       actualList.append((r['text'][i]["actual"]))
    

    prev_yearsList=[]
    for i in range(start, end):
       prev_yearsList.append(r['text'][i]["previous_year"])
    
    
    
    prev_yeardict = dict(zip(monthList, prev_yearsList))
    # my_dict = {
    #     'Month': monthList,
    #     'Previous year cost':prev_yearsList
    # }
    # prev_yeardict = pd.DataFrame(my_dict)

    # Create a line plot
    plt.plot(monthList, targetList, label='Target')
    plt.plot(monthList, actualList, label='Actual')
    plt.plot(monthList, prev_yearsList, label='Previous year')

    # Add labels and title
    plt.xlabel('Month')
    plt.ylabel('Value')
    plt.title('Monthly Target, Actual, and Previous Year')

    # Add legend
    plt.legend()

    return plt, prev_yeardict

# To get data from word documents 
def getMonthChart(folder_path, year, month, orgName):
    # Load multiple Word documents from folder
    word_files = [os.path.join(folder_path, file) for file in os.listdir(folder_path) if file.endswith('.docx')]

    loaders = []
    for word_file in word_files:
        loader = Docx2txtLoader(word_file)
        loaders.append(loader)

    # Create index from loaders
    index = VectorstoreIndexCreator().from_loaders(loaders)

    prompt = """
Extract monthly procurement cost as 
   - previous year
   - target cost
   - actual cost

give number withought comma
Output in JSON format:
{
   "previous_year": previous_year,
   "target_cost": target_cost,
   "actual_cost": actual_cost
}
"""
    # subquery to fetch data as per our need
    query = f"Extract procurement costs for {orgName} in the month of {month}, including previous years, target, and actual for {month} {year}  "

    # Concatenate prompt with the main query
    main_query = query + prompt

    # Query the index
    responses = index.query(main_query)
    print(responses)

    # Parse the response into a dictionary
    try:
        response_dict = json.loads(responses)
        # Access values from the dictionary
        previous_year = response_dict.get("previous_year")
        target_cost = response_dict.get("target_cost")
        actual_cost = response_dict.get("actual_cost")

        # Print the values
        print("Previous Year:", previous_year)
        print("Target Cost:", target_cost)
        print("Actual Cost:", actual_cost)

        # Display the values using Streamlit
        monthProcurementData = {
            "name": ["Previous year Cost", "Target Cost", "Actual Cost"],
            "Procurement Cost": [previous_year, target_cost, actual_cost]
        }
    
        chart_data = pd.DataFrame(monthProcurementData)
        chart_data.set_index("name", inplace=True)
        #st.bar_chart(chart_data)

    except json.JSONDecodeError as e:
        print("Error parsing response:", e)

    return monthProcurementData

on = st.sidebar.toggle('Activate Chatbot')

### Streamlit application  ###################################################

## To Access a Dashboard UI
if not on:
    
    # Store the initial value of widgets in session state
    if "visibility" not in st.session_state:
        st.session_state.visibility = "visible"
        st.session_state.disabled = False
    col1, col2, col3 = st.columns([0.25, 0.15, 0.6])


    with col1:

        # Get supplier name from Headcount database
        #organizationName=getOrganizationName()
        organizationName=[]
        for org in getOrganizationName():
            organizationName.append(org[0])
        
        option = st.selectbox(
            'Select the supplier',
            organizationName)


    # Get the current year
    this_year = datetime.now().year
    var = datetime.now().month

    with col2:
        report_year = st.selectbox(
        "Year",
        options=range(this_year, this_year - 5, -1),  # range from this_year to this_year - 5
        index=None,  # No default selection
        format_func=lambda x: str(x),  # Display integers as strings
        placeholder="Year..."
        )

        # Display the selected year
        #st.text(f'Selected Year: {report_year}') 

    with col3:
        period = st.radio(
            "",
            ('Monthly', 'Quarterly', 'HalfYearly', 'Yearly'),
            horizontal=True, 
            label_visibility="hidden"
        )
        # Get the current month
        this_month = datetime.now().month



    if "visibility" not in st.session_state:
        st.session_state.visibility = "visible"
        st.session_state.disabled = False
    col1, col2, col3 = st.columns([0.6, 0.25, 0.25])


    with col1:
        pass

    with col2:

        if period == 'Monthly':
            
            # Create a dropdown for choosing a month
            report_month_str = st.selectbox(
                '',
                month_abbr[1:],  # excluding the first empty string in month_abbr
                index=this_month - 1,  # default index set to the current month
                key="report_month_str"  # setting a key for this element
            )
        
            # Convert the selected month abbreviation to its corresponding number
            report_month = month_abbr[this_month - 1]
            
            # Display the selected month
            #st.text(f'Selected Month: {report_month_str}')
            var = report_month_str



        elif period == 'Quarterly':
            quarterly = st.selectbox(
            '',
            ('Q1', 'Q2', 'Q3', 'Q4'))

            #st.write('You selected:', quarterly)
            if quarterly=='Q1':
                var =f' average First Quarter'
            elif quarterly=='Q2':
                var =f' average Second Quarter'
            elif quarterly=='Q3':
                var =f' average Third Quarter'
            elif quarterly=='Q4':
                var =f' average Fourth Quarter'
    

        elif period == 'HalfYearly':
            half = st.selectbox(
            '',
            ('1HY', '2HY'))

            #st.write('You selected:', half)
            if half =='1HY':
                var =f' average first half of year'
            elif half =='2HY':
                var =f' average Second half of year'

        
        else:
            #st.text(f'Selected Year: {report_year}') 
            var= 'average year'


    with col3:
        '',
        '',
        submit = st.button("Submit")



    st.subheader(f'', divider='rainbow')


    # Inject custom CSS to set the width of the sidebar
    st.markdown(
        """
        <style>
            section[data-testid="stSidebar"] {
                width: 330px !important; # Set the width to your desired value
                font-size: 10px;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


    #st.sidebar.subheader(f"Graphical Analysis of Procurement Data: {option}")
    
    # To run script after providing data variable to dashboard
    if submit:
        st.markdown(
        f"<div style='border:1px solid #e6e6e6; padding:10px; border-radius:5px; font-size: 15px;'>{wikipediaMetadata(option)}</div>",
        unsafe_allow_html=True
        )

        st.subheader(f'', divider='rainbow')

        st.sidebar.markdown(
        f'<h3 style="font-size: 13px; font-weight: bold; float: left; margin-top: 5%;">Graphical Analysis of Procurement Data: {option}</h3>', 
        unsafe_allow_html=True
        )

        if "visibility" not in st.session_state:
            st.session_state.visibility = "visible"
            st.session_state.disabled = False,   
        col1, col2 = st.columns(2)
        
        # call to the pptx function
        graphpoints, prev_yeardict = pptxGraph(pptxFile, period)
        
        with col1:
            #st.subheader(f'<p style="font-family: Arial; color: black; font-size: 25px;">The projected cost of Procurement for {var} {report_year} </p>', divider='rainbow')
            st.write(f'<p style="font-size: 15px; font-weight: bold;">The projected cost of Procurement for {var} {report_year}</p>', unsafe_allow_html=True)
            st.subheader(f'', divider='rainbow')
            
            # To provied Projected procurement data to table
            res = getProcurementCost(str(report_year), str(var))
            st.write(res)
            
            st.write(f'<p style="font-size: 15px; font-weight: bold;">The Actual cost of previous year </p>', unsafe_allow_html=True)
            st.subheader(f'', divider='rainbow')
            
        
            #st.write(prev_yeardict)
            #var= 'Feb'
            #report_year= str(2022)
            
            # To provied preavious month data to table
            val = var +" "+ str(report_year)

            if val in prev_yeardict:  
                specific_pair = {val: prev_yeardict[val]}
                st.table(specific_pair)
                print(specific_pair)  
            else:
                st.table(prev_yeardict)
               
               
    
        with col2:
            #st.subheader(f'The Actual cost of Procurement for {var} {report_year} ', divider='rainbow')
            st.write(f'<p style="font-size: 15px; font-weight: bold;">The Actual cost of Procurement for {var} {report_year}</p>', unsafe_allow_html=True)
            st.subheader(f'', divider='rainbow')
            
            # To provied Projected procurement data to table
            Actual_res= getActualProcurementCost(str(report_year), str(var)) 
            st.write(Actual_res)

            st.subheader(f'', divider='rainbow')
            
            monthProcurementData=getMonthChart(wordFolderPath, report_year, var, option)
            chart_data = pd.DataFrame(monthProcurementData)
            chart_data.set_index("name", inplace=True)
            #st.sidebar.bar_chart(chart_data)
            #st.sidebar.line_chart(chart_data)
            
            st.sidebar.write("")  # Add an empty space using st.sidebar.write()
            st.sidebar.write("")  # Add an empty space using st.sidebar.write()
            
            # To plot graph to dashboard sidebar
            if period == 'Monthly':
                st.sidebar.bar_chart(chart_data)
            else: 
                st.sidebar.pyplot(graphpoints)  
            
            

        
############################ Add natural language application  #####################################     

# search from vector database
def query_pinecone(user_query):
    
    # initialize pinecone database index
    index = pc.Index(os.environ["PINECONE_INDEX"])
    vector_store = Pinecone(index, embeddings.embed_query, "text")
    # Created chain to get respone on data using natural language withought source
    qa = RetrievalQA.from_chain_type(llm=OpenAI(temperature=0), chain_type="stuff", retriever=vector_store.as_retriever())
    try:
        # Created chain to get respone on data using natural language with source
        qa_with_sources = RetrievalQAWithSourcesChain.from_chain_type(
            llm=OpenAI(temperature=0),
            chain_type="stuff",
            retriever=vector_store.as_retriever()
        )
        sources = qa_with_sources(user_query)
    except  Exception as e:
        sources = {'sources': "For this information, I did not find any specific source file"}
        print("An error occurred during data processing:", e)
        
    
    answer = qa(user_query)
    
    return answer, sources

### Switch to the natural language chatbot UI
if on:
    
    # created embedings using openai
    embeddings = OpenAIEmbeddings()
    
    # create connection to connect with pinecone database
    pc=pinecone.Pinecone(Pinecone_api_key=os.environ["PINECONE_API_KEY"], environment=os.environ["PINECONE_ENV"])
    
    user_query=st.text_input("ask question !!!")
    submit=st.button("submit", key="button1")
    
    if submit:
        st.subheader(f'', divider='rainbow')
        
        answer, sources=query_pinecone(user_query)
        #st.write(sources)
        st.write(answer['result'])
        
        # st.subheader(f'source of information', divider='rainbow')
        
        original_title = '<p style="font-family:Courier; color:green; font-size: 18px; font-weight: italic;">Source of information</p>'
        st.markdown(original_title, unsafe_allow_html=True)
        
        # To provied a  source file of data
        #st.write(sources)
        st.write(sources['sources'])
