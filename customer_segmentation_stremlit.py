import numpy as np
import pandas as pd
import streamlit as st
import re
import pytz
from sklearn.cluster import KMeans
from dateutil.relativedelta import relativedelta 
from datetime import datetime, timedelta, timezone
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
import xlsxwriter
from io import BytesIO


def take_out_threshold(col):
    area_temp=[]
    threshold_value=[]
    for i in range(5):
        area_temp=np.str(col[i:])
        split_list = list(filter(None, re.split(r'[(,!#$?%^\]]', area_temp)))
        print(split_list[1],split_list[2])
        threshold_value.append([float(split_list[1]),float(split_list[2])])
    return threshold_value



#左邊sidebar部份
st.sidebar.markdown("""<font size="6"><b>誰是你的VIP?</b></font>""", unsafe_allow_html=True)
st.sidebar.markdown("""<font size="2">1.若沒有樣本檔，請先下載,<br>&nbsp;&nbsp;並編輯貼上自己的資料。</font>""", unsafe_allow_html=True)
st.sidebar.download_button("下載excel sample",
     data="sample.xlsx",
     file_name='sample.xlsx'
     #,mime='text/csv'
     )
st.sidebar.text("  ")

uploaded_file = st.sidebar.file_uploader("2.上傳編輯完的csv", type="xlsx")
st.sidebar.markdown("""<font size="2">3.上傳結束後，稍等1~2分鐘,<br>&nbsp;&nbsp;右邊會開始呈現分類結果。</font><br><br><br>""", unsafe_allow_html=True)


#右邊的版面格式
st.markdown("""<font size="3"><b>消費者RFM分析&nbsp;:&nbsp;<br>依消費者R(Recency)、F(Frequency)、M(Monetary)頻率來判斷，並使用機器學習K-means將其分群。</b></font>""", unsafe_allow_html=True)
st.markdown("""<font size="2">注意&nbsp;:&nbsp;主要分析一年內的資料，超過一年的資料將全被歸類為同一區塊。</font>""", unsafe_allow_html=True)
st.image("https://bnextmedia.s3.hicloud.net.tw/image/album/2020-07/img-1594266434-18060@900.jpg", caption="圖一 : RFM模型", width=None)
st.markdown("""<font size="3">●「重要價值客戶」:<br>為重要VIP客戶，營收主要來源，可多做行銷活動宣傳發送，主動連繫。<br>
        如果一家公司「重要價值」的客戶不多，其他都是價值很低的「一般保持」客戶，表示客戶結構很不健康，無法承受客戶流失的風險。<br><br>
        ●「重要保持客戶」:<br>是指最近一次的消費時間離現在較久，但消費頻率和金額都很高的客戶，企業要主動介紹新趨勢等資訊，和他們保持聯繫。<br><br>
        ●「重要發展客戶」:<br>是最近一次消費時間較近、消費金額高，但頻率不高、忠誠度不高的潛力客戶。為新客戶，企業必須嚴格檢視每一次服務體驗，是否讓客戶非常滿意<br><br>
        ●「重要挽留客戶」:<br>則是最近一次消費時間較遠、消費頻率不高，但消費金額高的用戶。企業要主動釐清久未光顧消費的原因，避免失去這群客戶。</font>"""
        ,unsafe_allow_html=True)

st.markdown("""<font size="1"><font color="gray"><br><br>圖片和文字來源:<br>
<a href="https://www.managertoday.com.tw/articles/view/60050">這個連結模型怎麼用？將客戶價值分 8 種，挖出你的「黃金級」顧客</a><br>
<a href="https://jerrywangtc.blog/rfm-case-introduce/">RFM模型案例分析指標介紹</a><br>
<a href="https://ithelp.ithome.com.tw/articles/10216203">一服見效的 AI 應用系列 第 2 篇</a><br>
</font></font>""",unsafe_allow_html=True)

st.markdown("""---""")



#如果有上傳檔案，即開始執行報表分析
df = pd.read_excel(uploaded_file,skiprows=2)
if df:
    #df = pd.read_excel(uploaded_file,skiprows=2)
    st.markdown("""<font size="3">●原始檔檢視</font>""",unsafe_allow_html=True)
    df['InvoiceDate'] = pd.to_datetime(df['InvoiceDate'])
    st.dataframe(df,width=1000,height=300)


