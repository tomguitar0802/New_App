import time
t0 = time.time()
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

@st.cache
def df_to_xlsx(df):
    output = BytesIO()
    df.to_excel(output, index=True, sheet_name="Sheet1")
    processed_data=output.getvalue()
    return processed_data

Path=st.sidebar.file_uploader('Excel')

if Path is not None:
    sheet_names=pd.ExcelFile(Path).sheet_names
    
    #for extraction
    df_extract=pd.read_excel(Path,sheet_name=sheet_names[0])
    BG_starts=df_extract["BG_Start"].to_list()
    BG_finishes=df_extract["BG_Finish"].to_list()
    Peak_starts=df_extract["Peak_Start"].to_list()
    Peak_finishes=df_extract["Peak_Finish"].to_list()
    
    #VOC_info
    df_BG=pd.read_excel(Path,sheet_name=sheet_names[1])
    mz=df_BG["Exact mass"].to_list()
    df_BG=df_BG.set_index("VOC species").T
    df_Peak=pd.read_excel(Path,sheet_name=sheet_names[1])
    df_Peak=df_Peak.set_index("VOC species").T  
    
    #Raw data
    df_data=pd.read_excel(Path,sheet_name=sheet_names[2])
    df_data=df_data.drop(columns="Time")
    df_data=df_data.set_index("Data point")
    
    #BG_mean
    n=0
    for BG_start,BG_finish in zip(BG_starts,BG_finishes):
        n=n+1
        Peak_n="Peak "+str(n)
        df_BG.loc[Peak_n]=df_data.iloc[BG_start:BG_finish+1].mean()
        
    #Peak_mean
    n=0
    for Peak_start,Peak_finish in zip(Peak_starts,Peak_finishes):
        n=n+1
        Peak_n="Peak "+str(n)
        df_Peak.loc[Peak_n]=df_data.iloc[Peak_start:Peak_finish+1].mean()
    
    #subtract
    df=df_Peak-df_BG
    df=df.drop("Exact mass")
    if st.sidebar.checkbox("negative values to zero",True):
        df=df.where(df>0,0)
    st.dataframe(df.T)
    
    #mass spectre
    options=[]
    for i in range(len(BG_starts)):
        options.append("Peak "+str(i+1))
    option=st.sidebar.selectbox("mass spectre for?",options)
    
    fig,ax=plt.subplots()
    ax.ticklabel_format(useOffset=False,useMathText=True)
    plt.bar(x=mz,
            height=df.loc[option].to_list())
    ax.set_xlabel("m/z")
    ax.set_ylabel("Intensity")
    st.pyplot(fig)
    
    #time
    st.write('Elapsed time[s] =', str(float(time.time() - t0)))
    
    #download
    df_xlsx = df_to_xlsx(df.T)
    st.download_button(label="Download Excel", data=df_xlsx, file_name=Path.name.replace(".xlsx","")+"_processed.xlsx")
