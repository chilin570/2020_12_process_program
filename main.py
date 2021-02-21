
import warnings
warnings.filterwarnings("ignore")
import pandas as pd
import numpy as np 
filename = input("請輸入欲分析資料位置> ").replace(" ", "")
df = pd.read_excel(filename)

far_associate_correct_ans = input("請輸入遠距聯想測驗標準答案> ").replace(" ", "")
if far_associate_correct_ans == "":
    far_associate_correct_ans = "遠距聯想測驗標準答案.xlsx"
    

筷子 = input("請輸入竹筷子常模> ").replace(" ", "")
if 筷子 == "":
    筷子 = "新編竹筷子常模.xlsx"
吸管 = input("請輸入吸管常模> ").replace(" ", "")
if 吸管 == "":
    吸管 = "不尋常用途－吸管常模.xlsx"

寶特瓶 = input("請輸入寶特瓶常模> ").replace(" ", "")
if 寶特瓶 == "":
    寶特瓶 = "不尋常用途－寶特瓶常模.xlsx"

norm_df = {}
norm_df["竹筷子"] = pd.read_excel(筷子)
norm_df["吸管"] = pd.read_excel(吸管)
norm_df["寶特瓶"] = pd.read_excel(寶特瓶)

標準輸出 = input("請輸入平台計分輸出格式> ").replace(" ", "")
if 標準輸出 == "":
    標準輸出 = "./平台計分輸出格式.xlsx"

xls = pd.ExcelFile(標準輸出)

out_df = pd.read_excel(xls, "基本格式")

interact_df = pd.read_excel(xls, "互動歷程分析")


time_setting = input("請輸入共變包含的時間（單位為秒，預設為60）> ").replace(" ", "")
if time_setting == "":
    time_setting = 60
else:
    time_setting = int(time_setting)

already_score = input("有沒有先前輸出的評分完畢的資料？\n如果有，將能節省大量運算時間> ").replace(" ","")
if already_score == "":
    #先淨化資料
    df["ID"] = df["UserName"].astype(str) + "-" + df["Single/Double Mode"].astype(str) + "-" + df["Quiz Class"].astype(str) + "-" + df["Quiz #"].astype(str)#"].astype(str)
    df["IDT"] = df["ID"] + df["Time"]
    from datetime import datetime
    df["TimeStamp"] = df.Time.apply(datetime.strptime, args=('%Y/%m/%d, %H:%M:%S',))
    df = df.sort_values("TimeStamp")
    df_test = df.copy()
    print("正在移除空值...完成")
    df = df[df.Ans.isnull() == False]

    print("正在移除重複的資料...", end = "")
    for ID in df.ID.unique():
        target = df[df.ID == ID]
        df.drop(target.index[:-1], inplace = True)
    #df.to_excel("移除過無效與重複資料.xlsx")



    import pandas as pd
    print("完成")
    #df = pd.read_excel("移除過無效與重複資料.xlsx")
    #清掉多餘欄位
    for col in ["Unnamed: 7","Unnamed: 8","Unnamed: 9","Unnamed: 10"]:
        try:
            df = df.drop(columns=col)
        except:
            pass
    correct_df = pd.read_excel("遠距聯想測驗標準答案.xlsx")

    ans = []
    for _, row in df.iterrows():
        if row["Quiz Class"] in [4,5,6]:
            #是遠距聯想測驗
            version = row["Quiz Class"] - 3
            quizNumber = row["Quiz #"] + 1
            target = correct_df[correct_df["版本"] == version][correct_df["題號"] == quizNumber]
            target = list(target.iterrows())[0][1]
            if len(target.答案) == 0:
                ans.append(9)
            elif row.Ans in target.答案.replace(" ",""):
                ans.append(1)
            else:
                ans.append(0)

        else:
            ans.append(None)
    df["associate_score"] = ans
    import jieba
    def cut(text):
        for n in "、，。（）().,":
            text = text.replace(n, " ")
        outList = list(jieba.cut(text))
        outList = [i for i in outList if i != " "]
        return  ','.join(outList)
    for key in norm_df.keys():
        try:
            norm_df[key]["cut"] = norm_df[key].反應項目.apply(cut)
        except:
            norm_df[key]["cut"] = norm_df[key].反應項.apply(cut)

    def similiar(a_list, b_list):
        a_list = set(a_list)
        b_list = set(b_list)
        score = 0
        times = 0
        for norm_part in a_list:
            for ans_part in b_list:
                if norm_part == ans_part:
                    score += 1
            times += 1
        score = score / times
        return score
    print("正在計算成績...", end = "")
    all_類別 = []
    all_獨創力 = []
    all_check = []
    all_highest_score = []
    for _, row in df.iterrows():
        if row["Quiz Class"] in [1,2,3]:
            #是創造力測驗
            out_類別 = None
            out_獨創力 = None
            out_check = None
            highest_score = -1

            row_ans_cut_list = cut(row.Ans).split(",")

            quiz_class = row["Quiz Class"]
            if quiz_class == 1:
                mode = "寶特瓶"
            elif quiz_class == 2:
                mode = "吸管"
            elif quiz_class == 3:
                mode = "竹筷子"
            for _, norm_row in norm_df[mode].iterrows():
                norm_row_cut_list = norm_row.cut.split(",")
                similiar_score = similiar(norm_row_cut_list, row_ans_cut_list)

                if similiar_score > highest_score:
                    highest_score = similiar_score
                    out_類別 = norm_row.類別
                    out_獨創力 = norm_row.獨創力
                    out_check = None#norm_row.check

            all_highest_score.append(highest_score)
            all_類別.append(out_類別)
            all_獨創力.append(out_獨創力)
            all_check.append(out_check)

        else:
            all_類別.append(None)
            all_獨創力.append(None)
            all_check.append(None)
            all_highest_score.append(None)
    df["創造力_類別"] = all_類別
    df["創造力_獨創力"] = all_獨創力
    df["創造力_check"] = all_check

    df["創造力_score"] = all_highest_score
    #df.to_csv("評分完畢的資料.csv")


    import datetime
    now_time = datetime.datetime.now().strftime("%Y_%m_%d_%H%M%S")
    df.to_excel(f"評分完畢的資料_{now_time}.xlsx")
    import os
    current_path = os.path.abspath('.')
    print(f"評分完畢，輸出暫存檔 {current_path}/評分完畢的資料_{now_time}.xlsx")
else:
    df = pd.read_excel(already_score)

import pandas as pd
import numpy as np
print("正在計算並輸出成績")

associate_scores = {}
valid_answer = {}
creative_scores = {}
creative_classes = {}
single_double_mode = {}
associate_answers = {}



for _, row in df.iterrows():
    
    times_count = valid_answer.setdefault(row.UserName, {"竹筷子":0, "寶特瓶": 0, "吸管": 0})
    valid_answer[row.UserName] = times_count

    score = creative_scores.setdefault(row.UserName, {"竹筷子":0, "寶特瓶": 0, "吸管": 0})
    creative_scores[row.UserName] = score

    classes = creative_classes.setdefault(row.UserName, {"竹筷子": set(), "寶特瓶": set(), "吸管": set()})
    creative_classes[row.UserName] = classes

    mode_data = single_double_mode.setdefault(row.UserName, {"CR01": -1, "CR02": -1, "CC01": -1, "竹筷子":-1, "寶特瓶": -1, "吸管": -1})
    single_double_mode[row.UserName] = mode_data

    score_data = associate_scores.setdefault(row.UserName, {"CR01": 0, "CR02": 0, "CC01": 0})
    associate_scores[row.UserName] = score_data

    record = associate_answers.setdefault(row.UserName, {"CR01": {}, "CR02": {}, "CC01": {}})
    associate_answers[row.UserName] = record

    
    if row["Quiz Class"] in [4,5,6]:
        
        
        
        mode_data = single_double_mode.setdefault(row.UserName, {"CR01": -1, "CR02": -1, "CC01": -1, "竹筷子":-1, "寶特瓶": -1, "吸管": -1})
        quiz_class = row["Quiz Class"]
        if quiz_class == 4:
            mode = "CR01"
        elif quiz_class == 5:
            mode = "CR02"
        elif quiz_class == 6:
            mode = "CC01"
        mode_data[mode] = row["Single/Double Mode"]
        single_double_mode[row.UserName] = mode_data
        
        
        if row.associate_score < 4:
            score_data = associate_scores.setdefault(row.UserName, {"CR01": 0, "CR02": 0, "CC01": 0})
            score_data[mode] += row.associate_score
            associate_scores[row.UserName] = score_data
        
        
        record = associate_answers.setdefault(row.UserName, {"CR01": {}, "CR02": {}, "CC01": {}})
        record[mode][row["Quiz #"]] = row["associate_score"]
        associate_answers[row.UserName] = record

    elif row["Quiz Class"] in [1,2,3] and row.Ans:
        #單雙
        
        quiz_class = row["Quiz Class"]
        if quiz_class == 1:
            mode = "寶特瓶"
        elif quiz_class == 2:
            mode = "吸管"
        elif quiz_class == 3:
            mode = "竹筷子"
            
        mode_data = single_double_mode.setdefault(row.UserName, {"CR01": -1, "CR02": -1, "CC01": -1, "竹筷子":-1, "寶特瓶": -1, "吸管": -1})
        mode_data[mode] = row["Single/Double Mode"]
        single_double_mode[row.UserName] = mode_data
        
        #創造力總分
        score = creative_scores.setdefault(row.UserName, {"竹筷子":0, "寶特瓶": 0, "吸管": 0})
        score[mode] += row["創造力_獨創力"]
        creative_scores[row.UserName] = score
        
        #類別總數
        classes = creative_classes.setdefault(row.UserName, {"竹筷子": set(), "寶特瓶": set(), "吸管": set()})
        classes[mode].add(row["創造力_類別"])
        creative_classes[row.UserName] = classes
        
        
        
        #有效作答次數
        if len(row.Ans.replace(" ","")) > 0:
            times_count = valid_answer.setdefault(row.UserName, {"竹筷子":0, "寶特瓶": 0, "吸管": 0})
            times_count[mode] += 1
            valid_answer[row.UserName] = times_count
            
        

creative_classes_amount = {}
for key in creative_classes.keys():
    creative_classes_amount[key] = len(creative_classes[key])




#out_df
#template
template = out_df.copy()
def get_template_row():

    template["編號"] = [0]

    _, template_row = list(template.iterrows())[0]

    template_row = template_row.copy()
    return template_row
try:
    out_df = out_df.drop(index=0)
except:
    pass
for user_id in creative_classes.keys():
    out_row = get_template_row()
    out_row['編號'] = user_id
    out_row['CR01mode（單／雙）'] = single_double_mode[user_id]["CR01"]
    out_row['CR02mode（單／雙）'] = single_double_mode[user_id]["CR02"]
    out_row['CC01mode（單／雙）'] = single_double_mode[user_id]["CC01"]
    out_row['吸管聯想-執行模式（單／雙）'] = single_double_mode[user_id]["吸管"]
    out_row['寶特瓶聯想-執行模式（單／雙）'] = single_double_mode[user_id]["寶特瓶"]
    out_row['竹筷子聯想-執行模式（單／雙）'] = single_double_mode[user_id]["竹筷子"]
    out_row['CR01總分'] = associate_scores[user_id]["CR01"]
    out_row['CR02總分'] = associate_scores[user_id]["CR02"]
    out_row['CC01總分'] = associate_scores[user_id]["CC01"]


    for mode in ["吸管", "寶特瓶", "竹筷子"]:
        out_row[f"{mode}聯想-流暢性"] = valid_answer[user_id][mode]
        out_row[f"{mode}聯想-變通性"] = len(creative_classes[user_id][mode])
        out_row[f"{mode}聯想-獨創性"] = creative_scores[user_id][mode]



    for class_name in ["CR01", "CR02", "CC01"]:
        cr01 = associate_answers[user_id][class_name]
        for ans_num in cr01.keys():
            col_key = f"{class_name}_{str(ans_num+1).zfill(2)}"
            out_row[col_key] = associate_answers[user_id][class_name][ans_num]

    out_df = out_df.append(out_row)
out_df = out_df.sort_values("編號")

import os
import datetime
now_time = datetime.datetime.now().strftime("%Y_%m_%d_%H%M%S")
os.mkdir(str(now_time))
out_df.to_excel(f"{now_time}/output.xlsx")
def convert_time_to_int(inputs):
    user_ans_time = datetime.datetime.strptime(inputs, "%Y/%m/%d, %H:%M:%S")
    return user_ans_time.timestamp()
df["time_int"] = df.Time.apply(convert_time_to_int)
double_partners = []
classical_cooccurence_dict = {}
classical_cooccurence_time_dict = {}
more_creative_dict = {}



go_flow_dict = {}
go_flow_correct_dict = {}
stick_self_dict = {}
stick_self_correct_dict = {}
 
import datetime


print("正在處理共變資料")
for _, row in df.iterrows():
    
    userName = row.UserName
    user_ans_time = row.time_int
    if userName %2:
        partner = userName - 1
    else:
        partner = userName + 1
        
    #classical_cooccurence = classical_cooccurence_dict.setdefault(partner,  {"竹筷子":0, "寶特瓶": 0, "吸管": 0})
    #classical_cooccurence_dict[partner] = classical_cooccurence
    #classical_cooccurence_time = classical_cooccurence_time_dict.setdefault(partner,  {"竹筷子":0, "寶特瓶": 0, "吸管": 0})
    #classical_cooccurence_time_dict[partner] = classical_cooccurence_time
    #partner_more_creative = more_creative_dict.setdefault(partner,  {"竹筷子":0, "寶特瓶": 0, "吸管": 0})
    #more_creative_dict[partner] = partner_more_creative

    single_double = row["Single/Double Mode"]
    if single_double == 2:
        double_partners.append(partner)
        
    
    
    if row["Quiz Class"] in [4,5,6]:

        quiz_class = row["Quiz Class"]
        if quiz_class == 4:
            mode = "CR01"
        elif quiz_class == 5:
            mode = "CR02"
        elif quiz_class == 6:
            mode = "CC01"
        single_double = row["Single/Double Mode"]
        
        if single_double == 2:
            partner_ans = df[(df.UserName == partner) &
               (df["Quiz Class"] == row["Quiz Class"]) &
               (df["Quiz #"] == row["Quiz #"]) & 
               (df["time_int"] > row["time_int"])]
            partner_ans = list(partner_ans.iterrows())
            if len(partner_ans) > 0: #夥伴有在之後作答該題目
                partner_ans = partner_ans[0][1]
                partner_go_flow = go_flow_dict.setdefault(partner, {"CR01":0, "CR02":0, "CC01":0})
                partner_go_flow_correct = go_flow_correct_dict.setdefault(partner, {"CR01":0, "CR02":0, "CC01":0})
                partner_stick_self = stick_self_dict.setdefault(partner, {"CR01":0, "CR02":0, "CC01":0})
                partner_stick_self_correct = stick_self_correct_dict.setdefault(partner, {"CR01":0, "CR02":0, "CC01":0})
                if row.Ans in partner_ans.Ans:
                    partner_go_flow[mode] += 1
                    if partner_ans.associate_score == 1:
                        partner_go_flow_correct[mode] += 1
                else:
                    partner_stick_self[mode] += 1
                    if partner_ans.associate_score == 1:
                        partner_stick_self_correct[mode] += 1
        
        

    elif row["Quiz Class"] in [1,2,3] and row.Ans:
        #單雙
        
        quiz_class = row["Quiz Class"]
        if quiz_class == 1:
            mode = "寶特瓶"
        elif quiz_class == 2:
            mode = "吸管"
        elif quiz_class == 3:
            mode = "竹筷子"
            
        
        single_double = row["Single/Double Mode"]
        
        
        #類別共變
        if single_double == 2:
            user_df = df[(userName == df.UserName) &
               (df["Quiz Class"] == row["Quiz Class"])&
               (df["Quiz #"] == row["Quiz #"]) &
               (df.time_int > user_ans_time) &
               (df.time_int < user_ans_time + time_setting)]

            partner_df = df[(partner == df.UserName) &
                           (df["Quiz Class"] == row["Quiz Class"])&
                           (df["Quiz #"] == row["Quiz #"]) &
                           (df.time_int > user_ans_time) &
                           (df.time_int < user_ans_time + time_setting)]


            #類別共變

            for __, partner_ans in partner_df.iterrows():
                classical_cooccurence_time = classical_cooccurence_time_dict.setdefault(partner, {"CR01": 0, "CR02": 0, "CC01": 0, "竹筷子":0, "寶特瓶": 0, "吸管": 0})
                classical_cooccurence_time[mode] += 1
                classical_cooccurence_time_dict[partner] = classical_cooccurence_time

                classical_cooccurence = classical_cooccurence_dict.setdefault(partner, {"CR01": 0, "CR02": 0, "CC01": 0, "竹筷子":0, "寶特瓶": 0, "吸管": 0})
                if partner_ans.創造力_類別 == row.創造力_類別:
                    classical_cooccurence[mode] += 1
                    classical_cooccurence_dict[partner] = classical_cooccurence
                else:
                    pass
                #獨創共變
                partner_more_creative = more_creative_dict.setdefault(partner,  {"竹筷子":0, "寶特瓶": 0, "吸管": 0})
                partner_more_creative[mode] += (partner_ans.創造力_獨創力 - row.創造力_獨創力)
                more_creative_dict[partner]= partner_more_creative


        else:
            pass
        
        
        if single_double == 2:
            pass
        else:
            pass
        

interact_template = interact_df.copy()
def get_interact_template_row():

    interact_template["編號"] = [0]

    _, template_row = list(interact_template.iterrows())[0]

    template_row = template_row.copy()
    return template_row
for user_id in set(double_partners):
    out_row = get_interact_template_row()
    out_row["編號"] = int(user_id)
    for key in ['CR01', 'CR02', 'CC01', '竹筷子', '寶特瓶', '吸管']:
        try:
            if key in ["寶特瓶", "吸管"]:
                if classical_cooccurence_time_dict[user_id][key] != 0:
                    out_row[f"{key}類別共變{time_setting}s"] = classical_cooccurence_dict[user_id][key] / classical_cooccurence_time_dict[user_id][key]
                    out_row[f"{key}獨創共變{time_setting}s"] = more_creative_dict[user_id][key] / classical_cooccurence_time_dict[user_id][key]
                else:
                    out_row[f"{key}類別共變{time_setting}s"] = None
                    out_row[f"{key}獨創共變{time_setting}s"] = None
        except:
            pass
        try:
            if key in ["CR01", "CR02", "CC01"]:
                out_row[f"{key}隨波逐流反應數"] = go_flow_dict[user_id][key]
                out_row[f"{key}隨波逐流答對數"] = go_flow_correct_dict[user_id][key]
                out_row[f"{key}堅持己見反應數"] = stick_self_dict[user_id][key]
                out_row[f"{key}堅持己見答對數"] = stick_self_correct_dict[user_id][key]

            for class_name in ["CR01", "CR02", "CC01"]:
                cr01 = associate_answers[user_id][class_name]
                for ans_num in cr01.keys():
                    col_key = f"{class_name}_{str(ans_num+1).zfill(2)}"
                    out_row[col_key] = int(associate_answers[user_id][class_name][ans_num])
        except:
            pass
    interact_df = interact_df.append(out_row)
interact_df = interact_df[interact_df["編號"].isnull()==False]
interact_df.to_excel(f"{now_time}/共變.xlsx")

print(f"輸出完成，輸出檔案在 {os.path.abspath(str(now_time))} 資料夾中")