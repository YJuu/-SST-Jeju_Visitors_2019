import numpy as np
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
from matplotlib import font_manager,rc

file_path = 'C:/Users/user/Jeju_visitor.xlsx'
font_path = 'C:/Users/user/AppData/Local/Microsoft/Windows/Fonts/dovemayo.otf'
font_name = font_manager.FontProperties(fname=font_path).get_name()
matplotlib.rc('font',family=font_name)

month = [x for x in range(1,13)]
month_sheet = {1:'1월',2:'2월',3:'3월',4:'4월',5:'5월',6:'6월',7:'7월',8:'8월',9:'9월',10:'10월',11:'11월',12:'12월'}
visitor_sum_19 = []
visitor_sum_18 = []
visitor_asia_19 = []
visitor_asia_18 = []

visitor_jp_19 = []

visitor_8_country = []
visitor_8_visitor = []

fig = plt.figure()
ax1 = fig.add_subplot(1,3,1)
ax2 = fig.add_subplot(1,3,2)
ax3 = fig.add_subplot(1,3,3)

idx = np.arange(len(month))
xlabel = ['1월', '2월', '3월', '4월', '5월','6월','7월','8월','9월','10월','11월','12월']

for i in range(1,13):
    temp = pd.read_excel(file_path,sheet_name=month_sheet[i],usecols=[1,3,4])
    visitor_sum_19.append(temp.iloc[3,1])
    visitor_sum_18.append(temp.iloc[3,2])
    visitor_asia_19.append(temp.iloc[5,1])
    visitor_asia_18.append(temp.iloc[5,2])
    
    visitor_jp_19.append(temp.iloc[7,1])
    
    if(i == 8):
        for x in range(0,9):
            t = 7 + x*2
            visitor_8_country.append(temp.iloc[t,0])
            visitor_8_visitor.append(temp.iloc[t,1])
 
#subplot 1. 월간 제주도 방문 외국인

width = 0.3
offset = 0.17
ax1.set_title('월간 제주도 방문 외국인')
ax1.set_xticks(idx)
ax1.set_xticklabels(xlabel)
ax1.bar(idx-offset, visitor_sum_18,color='lightskyblue',width=width, label = '2018 방문객')
ax1.bar(idx+offset, visitor_sum_19,color='lightpink',width=width, label = '2019 방문객')
ax1.bar(idx-offset, visitor_asia_18,color='pink',width=width,label = '2018 아시아 방문객')
ax1.bar(idx+offset, visitor_asia_19,color='powderblue',width=width,label = '2019 아시아 방문객' )
ax1.legend(prop={'size':8})


#subplot 2. 제주 관광통계 - 일본

ax2.set_title("제주 관광통계 - 일본")
ax2.plot(xlabel, visitor_jp_19)

#subplot 3. 국가별 8월 방문자율

exp=(0,0.1,0,0,0,0,0,0,0)
ax3.pie(visitor_8_visitor, labels=visitor_8_country, explode=exp, autopct="%0.2f%%", startangle=90, counterclock=False, shadow = True, textprops = {'fontsize':7})
ax3.set_rcParms['font.size'] = 10
ax3.set_title("아시아국가별 8월 제주도 방문자율")
ax3.legend(labels=visitor_8_country, loc='upper right', prop={'size':6})

plt.show()