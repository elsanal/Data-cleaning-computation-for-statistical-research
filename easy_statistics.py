import numpy as np
import pandas as pd
import statistics
import math
import time


class statisticHelper:

    def readExcel(file_to_read, sheet_name=0, header=0, colums_offset=None):
        """ Read an excel file that contain, and split the data at each empty row
           companies are separated from each other by empty row. 

        Returns:
            a list that contains the splitted data
        """
        print("start reading")
        origin_file = pd.read_excel(io=file_to_read,sheet_name=sheet_name,header=header)
        if colums_offset is None:
            list_container = np.split(origin_file, origin_file[origin_file.isnull().all(1)].index)
        else:
            origin_file = origin_file.iloc[:, colums_offset:]
            list_container = np.split(origin_file, origin_file[origin_file.isnull().all(1)].index)
            list_container = [List for List in list_container if len(List) > 3] 
        print('end reading')    
        return list_container, origin_file
        
        
    def compute_log(file_to_read, path_to_save, output_filename, sheet_name=0, header=0, colums_offset=None, nb_items=100000):
        data={'':''}
        count = 0
        file = pd.DataFrame(columns=['Log1', 'Log2', 'Log3'])
        companies, origin_file = statisticHelper.readExcel(file_to_read=file_to_read, sheet_name=sheet_name, header=header, colums_offset=colums_offset)
        print("Computing Log...")
        for k, comp in enumerate(companies):
            if k==nb_items:
                break
            if k==0:
                company=comp
            else:
                company = comp[2:]
            actif = 0
            for i in range(0, len(company)):
                if company['Company '].iloc[i] !='#':
                    try :
                        log1 = 0
                        log2 = 0
                        log3 = 0

                        if (actif>=1):
                            day_cur = float(company['closing price'].iloc[i])
                            day_prev= float(company['closing price'].iloc[i-1])
                            if ((day_cur)<=0 or (day_prev)<=0):
                                log1=0
                            else:
                                log1 = math.log(float(company['closing price'].iloc[i])/float(company['closing price'].iloc[i-1]))
                        if (actif>=2):
                            day_prev_2 = float(company['closing price'].iloc[i-2])
                            if (day_prev_2)<=0 or (day_prev)<=0:
                                log2=0
                            else:
                                log2 = math.log(float(company['closing price'].iloc[i-1])/float(company['closing price'].iloc[i-2]))
                        if (actif>=3):
                            day_prev_3 = float(company['closing price'].iloc[i-3])
                            if (day_prev_2)<=0 or (day_prev_3)<=0:
                                log3=0
                            else:
                                log3 = math.log(float(company['closing price'].iloc[i-2])/float(company['closing price'].iloc[i-3]))
                        item = {'Log1':log1, 'Log2':log2, 'Log3':log3}
                        nw = pd.DataFrame(data=item, index=[''])
                        file = pd.concat([file, nw], axis=0)
                    except:
                        item = {'Log1':0, 'Log2':0, 'Log3':0}
                        nw = pd.DataFrame(data=item, index=[''])
                        file = pd.concat([file, nw], axis=0)
                        count = count+1
                    actif = actif+1
                else:
                    actif = 0
                    item = {'Log1':'#', 'Log2':'#', 'Log3':'#'}
                    nw = pd.DataFrame(data=item, index=[''])
                    file = pd.concat([file, nw], axis=0)         
            item = {'Log1':'', 'Log2':'', 'Log3':''}        
            nw = pd.DataFrame(data=item, index=[''])
            file = pd.concat([file, nw], axis=0)
            file = pd.concat([file, nw], axis=0)
        print(f'Error : {count}')
        file.to_excel(path_to_save + 'log.xlsx')
        _, log = statisticHelper.readExcel(path_to_save + 'log.xlsx', sheet_name=sheet_name, header=header, colums_offset=1)
        file = pd.concat([origin_file, log], axis=1)
        file.to_excel(path_to_save + output_filename + '.xlsx')
        print("Log computed")    
        
   
   # Calculate RSI
    def compute_gain_loss(file_to_read, path_to_save, output_filename, sheet_name=0, header=0, colums_offset=None, nb_items=100000):
        gain = []
        loss = []
        label = []
        companies, origin_file = statisticHelper.readExcel(file_to_read=file_to_read, sheet_name=sheet_name, header=header, colums_offset=colums_offset)
        for k, comp in enumerate(companies):
            if k==nb_items:
                break
            if k==0:
                company=comp
            else:
                company = comp
                actif = 0

            for i in range(0,len(company)):
                try:
                    if(company['Log1'].iloc[i]!='#'):
                        if float(company['Log1'].iloc[i])>=0:
                            gain.append(float(company['Log1'].iloc[i]))
                            loss.append(0)
                            label.append(1)
                        elif float(company['Log1'].iloc[i])<0:
                            gain.append(0)
                            loss.append(float(company['Log1'].iloc[i]))
                            label.append(0)
                        else:
                            gain.append('')
                            loss.append('')
                            label.append('')
                    else:
                        gain.append('#')
                        loss.append('#')
                        label.append('#')
                except:
                    gain.append(0)
                    loss.append(0)
                    label.append(0)
            
        gain_f = pd.DataFrame(data=gain, columns=['Gain'])
        loss_f = pd.DataFrame(data=loss, columns=['Loss'])
        label_f = pd.DataFrame(data=label, columns=['label'])
        file = pd.concat([origin_file, gain_f], axis=1)
        file = pd.concat([file, loss_f], axis=1)
        file = pd.concat([file, label_f], axis=1)
        file.to_excel(path_to_save + output_filename + '.xlsx')
        print("Gain & loss done!")   
        
    
    def compute_avg_gain_loss(file_to_read, path_to_save, output_filename, sheet_name=0, header=0, colums_offset=None, nb_items=100000):
        MIN = 0
        MAX = 10
        avg_gain = []
        avg_loss = []
        Avg, origin_file = statisticHelper.readExcel(file_to_read=file_to_read, sheet_name=sheet_name, header=header, colums_offset=colums_offset)
        empti = [' ']*10
        empti1 = [' ']*10
        for j, comp in enumerate(Avg):
            if j==nb_items:
                break
            if j==0:
                company = comp
                avg_gain = empti + avg_gain 
                avg_loss = empti + avg_loss

            else:
                company = comp[2:]
                avg_gain = avg_gain + ['']*2
                avg_loss = avg_loss + ['']*2
                avg_gain = avg_gain + empti
                avg_loss = avg_loss + empti
            actif = 0
            MIN = 0
            MAX = 10
            i = 0
        # gain and loss computation
            while i < (len(company)-10):
                try:
                    if(company['label'].iloc[MAX]!='#'):
                        avg_G = sum(company['Gain'][MIN:MAX])/len(company['Gain'][MIN:MAX])*1.0
                        avg_gain.append(avg_G)
                        avg_L = sum(company['Loss'][MIN:MAX])/len(company['Gain'][MIN:MAX])*1.0
                        avg_loss.append(avg_L)
                        MAX= MAX+1
                        i = i+1
                    elif company['label'].iloc[MAX]=='#':
                        # print('curent index: ', i, 'length comp : ',len(company))
                        avg_gain.append('#')
                        avg_loss.append('#')
                        i = i+11
                        MIN = i
                        MAX = MIN+10  
                        if len(company) - MAX <=11:
                            avg_gain = avg_gain + [' ']*(len(company)-i)
                            avg_loss = avg_loss + [' ']*(len(company)-i)
                        else:
                            avg_gain = avg_gain + empti1
                            avg_loss = avg_loss + empti1

                    else:
                        avg_gain.append('&')
                        avg_loss.append('&')                   
                except Exception as e:
                    print(e)
        gain_avg_f = pd.DataFrame(data=avg_gain, columns=['Avg. Gain'])
        loss_avg_f = pd.DataFrame(data=avg_loss, columns=['Avg. Loss'])
        file = pd.concat([origin_file, gain_avg_f], axis=1)
        file = pd.concat([file, loss_avg_f], axis=1)
        file.to_excel(path_to_save + output_filename + '.xlsx')
        print("Average gain & loss done!") 
        
        
    def compute_rsi(file_to_read, path_to_save, output_filename, sheet_name=0, header=0, colums_offset=None):
        _, origin = statisticHelper.readExcel(file_to_read=file_to_read, sheet_name=sheet_name, header=header, colums_offset=colums_offset)
        
        RSI = []
        RS = []
        for i in range(0, len(origin)):
            try:
                if origin['Avg. Gain'].iloc[i] == ' ':
                    RSI.append(' ')
                    RS.append(' ')
                elif origin['Avg. Gain'].iloc[i] == '#':
                    RSI.append('#')
                    RS.append('#')
                elif origin['Avg. Loss'].iloc[i] == 0:
                    RSI.append(0)
                    RS.append(0)
                else:
                    rs = origin['Avg. Gain'].iloc[i]*1.0/origin['Avg. Loss'].iloc[i]
                    rs = abs(rs)
                    rsi = 100 - (100/(1+rs))
                    RS.append(rs)
                    RSI.append(rsi)
            except Exception as e:
                print('curent index: ', i)
                print(e)
        rs_f = pd.DataFrame(data=RS, columns=['RS'])
        rsi_file = pd.DataFrame(data=RSI, columns=['RSI'])

        file = pd.concat([origin, rs_f], axis=1)
        file = pd.concat([file, rsi_file], axis=1,)
        file.to_excel(path_to_save + output_filename + '.xlsx')
        print("RSI done!")
        
    def compute_SMA(offset, file_to_read, path_to_save, output_filename, sheet_name=0, header=0, colums_offset=None, nb_items=100000):
        companies, origin_file = statisticHelper.readExcel(file_to_read=file_to_read, sheet_name=sheet_name, header=header, colums_offset=colums_offset,)
        SMA = []
        for k, comp in enumerate(companies):
            curr_close = 0
            startIndex = 0
            if k == nb_items:
                break 
            if k==0:
                company=comp
                SMA = SMA + ['']*offset
            else:
                company = comp[2:]
                margin = offset
                for j in range(0, offset):
                    if company.iloc[j,:1]['Company ']=='#':
                        margin = j
                        break     
                SMA = SMA + ['']*(margin+2)
            for i in range(len(company)):
                if(company.iloc[i,:1]['Company ']!='#'):
                    try:
                        curr_close = curr_close + company['closing price'].iloc[i]
                        startIndex = startIndex + 1
                        if startIndex>offset:
                            sma = curr_close/(startIndex)
                            SMA.append(sma)
    #                         print('SMA : ',sma)
                    except Exception as e:
                        print(e)
                elif company.iloc[i,:1]['Company ']=='#':
                    SMA.append('#')
                    curr_close = 0
                    if (len(company) - i-1) < offset:
                        SMA = SMA + ['']*(len(company) - i-1)
                    else:
                        SMA = SMA + ['']*offset
                    startIndex = 0
        
        file = pd.DataFrame(data=SMA, columns=['SMA'+str(offset)])
        file = pd.concat([origin_file, file], axis=1)
        file.to_excel(path_to_save + output_filename + '.xlsx')
        print("SMA{offset} done!".format(offset=offset))    
        

    def compute_EMA(offset, file_to_read, path_to_save, output_filename, sheet_name=0, header=0, colums_offset=None, nb_items=100000):
        companies, origin_file = statisticHelper.readExcel(file_to_read=file_to_read, sheet_name=sheet_name, header=header, colums_offset=colums_offset)
        EMA = []
        for k, comp in enumerate(companies):
            curr_close = 0
            startIndex = 0
            if k == nb_items:
                break 
            if k==0:
                company=comp
                EMA = EMA + ['']*offset
            else:
                company = comp[2:]
                margin = offset
                for j in range(0, offset):
                    if company.iloc[j,:1]['Company ']=='#':
                        margin = j
                        break    
                EMA = EMA + ['']*(margin+2)  
            for i in range(len(company)):
                if(company.iloc[i,:1]['Company ']!='#'):                
                    try: 
                        startIndex = startIndex + 1
                        if startIndex==offset+1:
                            ema = float(company['SMA'+str(12)].iloc[i])   
                            EMA.append(ema)
                        elif startIndex>offset:
                            curr_close = float(company['closing price'].iloc[i])
                            prev_ema = EMA[-1]
                            ema = (curr_close - prev_ema) * (2/(startIndex+1)) + prev_ema
                            EMA.append(ema)
                    except Exception as e:
                        print('index : ',i, "SMA : ", company['SMA'+str(12)].iloc[i], "close : ", company['closing price'].iloc[i])
                        print('length EMA : ', len(EMA))
                        print(e)
                elif company.iloc[i,:1]['Company ']=='#':
                    EMA.append('#')
                    curr_close = 0
                    if (len(company) - i-1) < offset:
                        EMA = EMA + ['']*(len(company) - i-1)
                    else:
                        EMA = EMA + ['']*offset
                    startIndex = 0
            
        file = pd.DataFrame(data=EMA, columns=['EMA' + str(offset)])
        file = pd.concat([origin_file, file], axis=1)
        file.to_excel(path_to_save + output_filename + '.xlsx')
        print("EMA{offset} done!".format(offset=offset)) 
        
        
    def compute_MACD(first_ema, second_ema, file_to_read, path_to_save, output_filename, sheet_name=0, header=0, colums_offset=None, nb_items=100000):
        companies, origin_file = statisticHelper.readExcel(file_to_read=file_to_read, sheet_name=sheet_name, header=header, colums_offset=colums_offset)
        MACD = []
        PPO = []
        offset=second_ema
        for k, comp in enumerate(companies):
            curr_close = 0
            startIndex = 0
            if k == nb_items:
                break 
            if k==0:
                company=comp
                MACD = MACD + ['']*offset
                PPO = PPO + ['']*offset
            else:
                company = comp[2:]
                margin = offset
                for j in range(0, offset):
                    if company.iloc[j,:1]['Company ']=='#':
                        margin = j
                        break    
                MACD = MACD + ['']*(margin+2)
                PPO = PPO + ['']*(margin+2)           
            for i in range(len(company)):
                if(company.iloc[i,:1]['Company ']!='#'): 
                    try: 
                        
                        if startIndex < offset:
                            pass
                        else :
                            ema12 = company['EMA'+str(first_ema)].iloc[i]
                            ema26 = company['EMA'+str(second_ema)].iloc[i]
                            macd = ema12 - ema26
                            MACD.append(macd)
                            if ema26 == 0:
                                PPO.append(0)
                            else:    
                                ppo = ((macd)*1.0/ema26)*100
                                PPO.append(ppo)
                        startIndex = startIndex + 1
                    except Exception as e:
                        print(e)
                elif company.iloc[i,:1]['Company ']=='#':
                    MACD.append('#')
                    PPO.append('#')
                    if (len(company) - i-1) < offset:
                        MACD = MACD + ['']*(len(company) - i-1)
                        PPO = PPO + ['']*(len(company) - i-1)
                    else:
                        MACD = MACD + ['']*offset
                        PPO = PPO + ['']*offset
                    startIndex = 0

        macd_file = pd.DataFrame(data=MACD, columns=['MACD'])
        ppo_file = pd.DataFrame(data=PPO, columns=['PPO'])
        file = pd.concat([macd_file, ppo_file], axis=1,)
        file = pd.concat([origin_file, file], axis=1,)
        file.to_excel(path_to_save + output_filename + '.xlsx')
        print('MACD done') 
    
    
    def compute_ROC(file_to_read, path_to_save, output_filename, sheet_name=0, header=0, colums_offset=None, nb_items=100000):
        companies, origin_file = statisticHelper.readExcel(file_to_read=file_to_read, sheet_name=sheet_name, header=header, colums_offset=colums_offset)
        ROC = []
        offset = 9
        for k, comp in enumerate(companies):
            curr_close = 0
            startIndex = 0
            START = 0
            END = 9
            if k == nb_items:
                break 
            if k==0:
                company=comp
                ROC = ROC + ['']*offset
            else:
                company = comp[2:]
                ROC = ROC + ['']*(offset+2)
            
            for i in range(len(company)):
                if(company.iloc[i,:1]['Company ']!='#'):                
                    try:     
                        if(startIndex>=offset) :
                            closeUP = company['closing price'].iloc[START]
                            closeDOWN = company['closing price'].iloc[i]
                            if closeUP == 0:
                                roc = 0
                            else:
                                roc = ((closeDOWN - closeUP)/(closeUP*1.0))*100
                            ROC.append(roc)    
                            START = START + 1 
                        startIndex = startIndex + 1
                    except Exception as e:
                        print("Index : ", i)
                        print(e)
                elif company.iloc[i,:1]['Company ']=='#':
                    ROC.append('#')
                    START = i+1 
                    if (len(company) - i-1) < offset:
                        ROC = ROC + ['']*(len(company) - i-1)
                    else:
                        ROC = ROC + ['']*offset
                    startIndex = 0          

        roc_file = pd.DataFrame(data=ROC, columns=['ROC'])
        file = pd.concat([origin_file, roc_file], axis=1,)
        file.to_excel(path_to_save + output_filename + '.xlsx')  
        print('ROC done! :)')
    
    # calculate the return
    def compute_year_return(company,startIndex):
        yearReturn = [0]
        RetAvgList = []
        ClosePrice = []
        closingList = []
        
        for i in range(startIndex,len(company)-1):
            # values of closing price
            x1 = company.iloc[i,2]
            x2 = company.iloc[i+1,2]  
                
            # list of closing price
            closingList.append(x1)
            if(company.iloc[i,:]["Company "]=='#'):
                yearReturn = []
                closingList = []
                

            elif((company.iloc[i+1,:]["Company "]=='#') or (i == len(company)-2)):
                ## avg, var, std with return values
                try:
                    yearAvg = statistics.mean(yearReturn[1:])
                    var = statistics.variance(yearReturn[1:], yearAvg)*252
                    std = statistics.stdev(yearReturn[1:])*math.sqrt(252)
                    RetAvgList.append({"Average":yearAvg,"Variance":var,"Standard":std})
                except:
                    print("Error1")
                
                ## avg, var, std with closing price values
                try:
                    cl_yearAvg = statistics.mean(closingList)
                    cl_var = statistics.variance(closingList, yearAvg)*252
                    cl_std = statistics.stdev(closingList)*math.sqrt(252)
                    ClosePrice.append({"closing Average":cl_yearAvg," closing Variance":cl_var," closing Standard":cl_std})
    #                 annualReturn.append(0)
                except:
                    print("Error2")
            else :    
                if(x1==0):
                    ri=0
                else:
                    try:
                        ri = (x2 - x1) /x1
                    except:
                        print("Error in line : ",company.iloc[3:4,0:1])
                yearReturn.append(ri)
        return RetAvgList,ClosePrice 
    
    
    def compute_return(file_to_read, path_to_save, output_filename, sheet_name=0, header=0, colums_offset=None, nb_items=None):
        RetAvgList = []
        data={'':''}
        file = pd.DataFrame( data = data,columns=['Company'])
        companies, origin_file = statisticHelper.readExcel(file_to_read=file_to_read, sheet_name=sheet_name, header=header, colums_offset=colums_offset)
        if nb_items is not None:
            limit = nb_items
        else :
            limit = len(companies)    
        for i in range(0, limit):
            RetAvgList.clear()
            print("Company No. ", i)
            if(i==0):
                RetAvgList,ClosePrice = statisticHelper.compute_year_return(companies[0],0)
                start = 0
                print("Number of year : ", len(RetAvgList))
                for indx in range(len(RetAvgList)):
    #                 print(indx)
                    item1 = ClosePrice[indx]
                    item2 = RetAvgList[indx]
                    items = item1.copy()
                    items.update(item2)
    #                 print(items)
                    nw = pd.DataFrame(data=items, index=[''])
    #                 print("Year : ",indx)

                    for k in range(start, len(companies[0])):
                        if(companies[0].iloc[k,:]["Company "]!='#'):
                            file = pd.concat([file,nw], axis=0)
                        else:
                            start = k+1
                            space = {" ":"#"}
                            sp = pd.DataFrame(data=space, index=[''])
                            file = pd.concat([file,sp], axis=0)
                            break    
                                                
                space = {" ":""}
                sp = pd.DataFrame(data=space, index=[''])
                file = pd.concat([file,sp], axis=0)
                file = pd.concat([file,sp], axis=0)
                
            else :
                
                try :
                    RetAvgList,ClosePrice = statisticHelper.compute_year_return(companies[i].iloc[2:,:],2)
                    start = 2
                    print("Number of year : ", len(RetAvgList))
                    for indx in range(len(RetAvgList)):
                        item1 = ClosePrice[indx]
                        item2 = RetAvgList[indx]
                        items = item1.copy()
                        items.update(item2)
                        nw = pd.DataFrame(data=items, index=[''])
    #                     print("year : ",indx)
                        for k in range(start, len(companies[i])):
                            if(companies[i].iloc[k,:]["Company "]!='#'):
                                file = pd.concat([file,nw], axis=0)
                            else:
                                start = k+1
                                space = {" ":"#"}
                                sp = pd.DataFrame(data=space, index=[''])
                                file = pd.concat([file,sp], axis=0)
                                break
                except:
                    print("Error")
                space = {" ":""}
                sp = pd.DataFrame(data=space, index=[''])
                file = pd.concat([file,sp], axis=0)
                file = pd.concat([file,sp], axis=0)
            
        file.to_excel(path_to_save + 'return.xlsx')  
        _, origin_file = statisticHelper.readExcel(file_to_read=path_to_save+'return.xlsx', sheet_name=0, header=0, colums_offset=3)
        file = origin_file
        file.to_excel(path_to_save + output_filename +'.xlsx') 
        
        
    def merge(output_path, output_filename, path1, path2, sheet_name1=0, sheet_name2=0, header1=0, header2=0,):
        file1 = pd.read_excel(path1,sheet_name=sheet_name1,header=header1)
        file2 = pd.read_excel(path2,sheet_name=sheet_name2,header=header2)
        print('merging...')
        file = pd.concat([file1, file2.iloc[:,1:]], axis=1)
        file.to_excel('{output_path}{output_filename}.xlsx'.format(output_path=output_path,output_filename=output_filename))
        print("Merging finished.")                