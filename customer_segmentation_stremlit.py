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
     data="https://github.com/chiahsun09/stremlit_RFM-model/raw/main/sample.xlsx",
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


@st.cache(allow_output_mutation=True)
def load_data():
    df=pd.read_excel(uploaded_file,skiprows=2)
    return df


#如果有上傳檔案，即開始執行報表分析
if uploaded_file is not None:
    df=load_data()
    st.write(df)
    st.markdown("""<font size="3">●原始檔檢視</font>""",unsafe_allow_html=True)
    df['InvoiceDate'] = pd.to_datetime(df['InvoiceDate'])
    st.dataframe(df,width=1000,height=300)


    #新增一個date欄位，轉成年月
    df['InvoiceDate'] = pd.to_datetime(df['InvoiceDate'])
    df['date']=df.InvoiceDate.astype(np.str).str.slice(0,8).str.replace('-','')
    df['date'] = pd.to_numeric(df['date'], errors='coerce')
  

        
    #轉換台灣時區和判斷距離3、6、9、12月以前的日期
    tz = pytz.timezone('Asia/Taipei')
    datetime_now = datetime.now(tz)

    three_month_ago  = datetime_now - relativedelta(months=+3)
    six_month_ago   = datetime_now - relativedelta(months=+6)
    nine_month_ago  = datetime_now - relativedelta(months=+9)
    twelve_month_ago = datetime_now - relativedelta(months=+12)

    three_month_ago=int(three_month_ago.strftime('%Y%m'))
    six_month_ago=int(six_month_ago.strftime('%Y%m'))
    nine_month_ago=int(nine_month_ago.strftime('%Y%m'))
    twelve_month_ago=int(twelve_month_ago.strftime('%Y%m'))


    #將消費日期分區###################################################################################
    def f(row):
        if row['date'] > three_month_ago:
            val = 5
        elif row['date'] <= three_month_ago and row['date'] > six_month_ago:
            val = 4
        elif row['date'] <= six_month_ago and row['date'] > nine_month_ago:
            val = 3
        elif row['date'] <= nine_month_ago and row['date'] > twelve_month_ago: 
            val = 2
        else:
            val=1
            
        return val

    df_recency_1=df[['CustomerID','date']].drop_duplicates()
    df_recency_1['Recency_Flag'] = df_recency_1.apply(f, axis=1)
    df_recency = df_recency_1.groupby('CustomerID',as_index=False)['Recency_Flag'].max()
    st.markdown("""---""")
    st.markdown("""<font size="3">●以消費日期近→遠區分</font>""",unsafe_allow_html=True)
    st.dataframe(df_recency,width=1000,height=300)
    

    #消費日期分區bar_chart
    st.bar_chart(df_recency.sort_values(by='Recency_Flag',ascending=True).Recency_Flag.value_counts())
    st.markdown("""<font size="2">說明&nbsp;:&nbsp;<br>
        得分&nbsp;1&nbsp;:&nbsp;{}&nbsp; / &nbsp; 
        得分&nbsp;2&nbsp;:&nbsp;{} &nbsp; / &nbsp;
        得分&nbsp;3&nbsp;:&nbsp;{}&nbsp; / &nbsp;
        &nbsp;得分&nbsp;4&nbsp;:&nbsp;{} &nbsp;/ &nbsp;
        得分&nbsp;5&nbsp;:&nbsp;{}
        </font>""".format("超過12個月","9-12個月內","6-9個月內","3-6個月內","3個月內"),unsafe_allow_html=True)



    ###計算購買頻率(Frequency)############################################################################
    Cust_freq=df[['InvoiceNo','CustomerID']].drop_duplicates()
    #Calculating the count of unique purchase for each customer
    Cust_freq_count=Cust_freq.groupby(['CustomerID'])['InvoiceNo'].aggregate('count').\
    reset_index().sort_values('InvoiceNo', ascending=False, axis=0)
    

    #InvoiceNo分5個區段
    unique_invoice=Cust_freq_count[['InvoiceNo']].drop_duplicates()
    unique_invoice['Freqency_Band'] = pd.cut(unique_invoice['InvoiceNo'], 5)
    unique_invoice=unique_invoice[['Freqency_Band']].drop_duplicates()
    #st.dataframe(unique_invoice)


    #分好的區段，將界限值取出 
    freqency_area=take_out_threshold(unique_invoice['Freqency_Band'])
       

    def f2(row):
        if row['InvoiceNo'] < freqency_area[4][1]:
            val = 1
        elif row['InvoiceNo'] >= freqency_area[3][0] and row['InvoiceNo']<freqency_area[3][1]:
            val = 2
        elif row['InvoiceNo'] >= freqency_area[2][0] and row['InvoiceNo']<freqency_area[2][1]:
            val = 3
        elif row['InvoiceNo'] >= freqency_area[1][0] and row['InvoiceNo']<freqency_area[1][1]:
            val = 4
        else: 
            val = 5
    
        return val
    

    #消費頻率分類處理    
    Cust_freq_count['Freq_Flag'] = Cust_freq_count.apply(f2, axis=1)
    st.markdown("""---""")
    st.markdown("""<font size="3">●以消費頻率多→少區分<br>&nbsp;&nbsp;得分在5~1之間，消費頻率越高得分越高</font>""",unsafe_allow_html=True)
    st.dataframe(Cust_freq_count,width=1000,height=300)
    st.bar_chart(Cust_freq_count.sort_values(by='Freq_Flag').Freq_Flag.value_counts())
    st.markdown("""<font size="2">說明&nbsp;:&nbsp;<br>
        得分&nbsp;1&nbsp;:&nbsp;小於{:.0f}&nbsp; / &nbsp;得分&nbsp;2&nbsp;:&nbsp;介於{:.0f}-{:.0f}&nbsp;/ &nbsp;
        得分&nbsp;3&nbsp;:&nbsp;介於{:.0f}-{:.0f}&nbsp; / &nbsp;得分&nbsp;4&nbsp;:&nbsp;介於{:.0f}-{:.0f}&nbsp; / &nbsp;
        得分&nbsp;5&nbsp;:&nbsp;大於{:.0f}</font>
        """.format(freqency_area[4][1],freqency_area[3][0],freqency_area[3][1],freqency_area[2][0],
        freqency_area[2][1],freqency_area[1][0],freqency_area[1][1],freqency_area[1][1]),unsafe_allow_html=True)   
       


    #計算購買金額(Monetary)#############################################################################
    #Calculating the Sum of total monetary purchase for each customer
    Cust_monetary = df.groupby(['CustomerID'])['Total_Price'].aggregate('sum').\
    reset_index().sort_values('Total_Price', ascending=False)
    
    #splitting Total price in 5 parts
    unique_price=Cust_monetary[['Total_Price']].drop_duplicates()
    unique_price=unique_price[unique_price['Total_Price'] > 0]
    unique_price['monetary_Band'] = pd.qcut(unique_price['Total_Price'], 5)
    unique_price=unique_price[['monetary_Band']].drop_duplicates()
    

    #分好的區段，將界限值取出
    Money_band=take_out_threshold(unique_price['monetary_Band'])


    #界定5~1級別    
    def f3(row):
        if row['Total_Price'] <= Money_band[4][1]:
            val = 1
        elif row['Total_Price'] >= Money_band[3][0] and row['Total_Price'] < Money_band[3][1]:
            val = 2
        elif row['Total_Price'] >= Money_band[2][0] and row['Total_Price'] < Money_band[2][1]:
            val = 3
        elif row['Total_Price'] >= Money_band[1][0] and row['Total_Price'] < Money_band[1][1]:
            val = 4    
        else:
            val = 5
    
        return val
        

    Cust_monetary['Monetary_Flag'] = Cust_monetary.apply(f3, axis=1)
    st.markdown("""---""")
    st.markdown("""●<font size="3">以消費金額多→少區分<br>&nbsp;&nbsp;得分在5~1之間，消費越高得分越高</font>""",unsafe_allow_html=True)
    st.dataframe(Cust_monetary,1000,300)
    st.bar_chart(Cust_monetary.sort_values(by='Monetary_Flag').Monetary_Flag.value_counts())
    st.markdown("""<font size="2">說明&nbsp;:&nbsp;<br>
        得分&nbsp;1&nbsp;:&nbsp;小於{:.2f}&nbsp; / &nbsp;得分&nbsp;2&nbsp;:&nbsp;介於{:.2f}-{:.2f}&nbsp;/ &nbsp;
        得分&nbsp;3&nbsp;:&nbsp;介於{:.2f}-{:.2f}&nbsp; / &nbsp;得分&nbsp;4&nbsp;:&nbsp;介於{:.2f}-{:.2f}&nbsp; / &nbsp;
        得分&nbsp;5&nbsp;:&nbsp;大於{:.2f}</font>
        """.format(Money_band[4][1],Money_band[3][0],Money_band[3][1],Money_band[2][0],Money_band[2][1],
        Money_band[1][0],Money_band[1][1],Money_band[1][1]),unsafe_allow_html=True)           


    #合併 RFM 欄位#############################################################################
    Cust_All=pd.merge(df_recency,Cust_freq_count[['CustomerID','Freq_Flag']], on=['CustomerID'],how='left')
    Cust_All=pd.merge(Cust_All,Cust_monetary[['CustomerID','Monetary_Flag']], on=['CustomerID'],how='left')
    Cust_All['total_score'] = Cust_All[["Recency_Flag","Freq_Flag","Monetary_Flag"]].apply(lambda x: x.sum(), axis=1)
    #Cust_All=Cust_All.drop(columns=["df_recency_map"])
    Cust_All.sort_values('total_score', axis=0, ascending=False, inplace=False)
    st.markdown("""---""")
    #st.markdown("""<font size="3">●合併RFM欄位</font>""",unsafe_allow_html=True)
    #st.dataframe(Cust_All.tail(10),1000,300)


    #使用 K means 分群#############################################################################
    kmeans = KMeans(n_clusters=4, init='k-means++', random_state=0)
    Cust_All_2=Cust_All.drop(Cust_All.columns[0], axis=1)
    Cust_All['clusters'] = kmeans.fit_predict(Cust_All_2)
    st.markdown("""<font size="3">●合併RFM欄位，使用K-means分群</font>""",unsafe_allow_html=True)
    st.dataframe(Cust_All,1000,300)


    #結論部份，sum數值相對最高，即是VIP族群(Cluster)########################################################
    Cluster_summary=round(pd.DataFrame(kmeans.cluster_centers_),2)
    Cluster_summary['sum'] = Cluster_summary.apply(lambda x: x.sum(), axis=1)
    vip=Cluster_summary.index[Cluster_summary['sum'] ==max(Cluster_summary['sum'])][0]
    st.markdown("""---""")
    st.markdown("""<font size="3"><b>●結論:</b><br>sum分數最高，即為最重要價值VIP族群。
    被分類為<font color="red"> {0:2d} </font>的群體是VIP。""".format(vip),unsafe_allow_html=True)
    st.markdown("""<font size="3">使用k-means將客戶分為4類，客戶重要性依sum分數高低來判斷。</font>""",unsafe_allow_html=True)
    st.dataframe(Cluster_summary.applymap(lambda x: '%.2f'%x))
   
    #3D圖
    #st.set_option('deprecation.showPyplotGlobalUse', False)
    fig, ax = plt.subplots()
    colors=['purple', 'blue', 'green', 'gold']
    #fig = plt.figure()
    fig.set_size_inches(12, 8)
    ax = fig.add_subplot(111, projection='3d')
    for i in range(kmeans.n_clusters):
        df_cluster=Cust_All[Cust_All['clusters']==i]
        ax.scatter(df_cluster['Recency_Flag'], df_cluster['Monetary_Flag'],df_cluster['Freq_Flag'],s=50,label='Cluster'+str(i), c=colors[i])
        plt.legend()
    #畫中心點    
    ax.scatter(kmeans.cluster_centers_[:,0],kmeans.cluster_centers_[:,1],kmeans.cluster_centers_[:,2],s=100,marker='^', c='red', alpha=0.7, label='Centroids')
    st.pyplot(fig)
    


    #其他建議事項
    st.image("https://miro.medium.com/max/788/0*RfgVR31iUo1RkbHK.jpg", caption="圖二 : RFM實施方法", width=None)
    st.markdown("""<font size="3"><br>以實際情況考量，若分類太多，做行銷活動也不容易。假定整份顧客名單消費金額都夠高，皆屬於值得經營:<br><br>
    ●日期較近，頻率較高 : 「重要價值客戶」，即VIP客戶，可多做專屬行銷活動宣傳發送。<br>
    ●日期較近，頻率較低 : 「重要發展客戶」，即新進顧客，促銷組合推薦，注意客戶體驗，發送歡迎信。<br>
    ●日期較遠，頻率較高 : 「重要保持客戶」，這些顧客為忠誠顧客，要時時主動連繫，介紹新資訊。<br>
    ●日期較遠，頻率較低 : 「重要挽留客戶」，可發送關懷信挽留，並釐清頻率降低原因，不然可能會失去這個顧客。<br><br>
    再依k-means分類重要性程度，代入分類，做後續行銷活動。</font><br><br><br>""",unsafe_allow_html=True)

    st.markdown("""<font size="3"><b>可再加強部份:</b><br>
    依產業別，要選用不同的因素加權評分，例如線上遊戲業者，可能需把在線時間考量進去評分;
    航空業者不單考慮自家的會員積分，另外可把顧客是否有其他家會員積分級別，即代表是經常飛行的人，這樣在分析頻率才會比較準確。<br><br>
    電商若是平均販售較平價的產品，顧客汰換率比較高，可以考慮把頻率(F)加權評分;相反的如果是高價商品，金額大小(M)就會比較重要。<br><br>
    加權評分需要熟知產業知識才能處理，若分析結果不符預期，要和其相關部門再互相溝通確認。<br><br>
    此次分析方法認為RFM三個因子都一樣重要，故沒有特別做加權評分。後續若認為消費金額(M)很重要，可自行加權計算(如乘以1.5倍)。</font><br><br>""",unsafe_allow_html=True)


    #處理完檔案下載連結
    #csv檔
    @st.cache
    def convert_df(df):
    #IMPORTANT: Cache the conversion to prevent computation on every rerun
        return df.to_csv().encode('utf-8')

    csv = convert_df(Cust_All)
    st.download_button(
    label="📥下載分類結果csv",
    data=convert_df(Cust_All),
    file_name='result.csv',
    mime='text/csv',
    )
    
    
    #存成excel檔
    @st.cache
    def to_excel(df):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        format1 = workbook.add_format({'num_format': '0.00'}) 
        worksheet.set_column('A:A', None, format1)  
        writer.save()
        processed_data = output.getvalue()
        return processed_data

    df_xlsx = to_excel(Cust_All)
    st.download_button(label='📥下載分類結果excel',
            data=df_xlsx ,
            file_name= 'result.xlsx')



