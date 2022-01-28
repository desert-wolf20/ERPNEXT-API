from frappeclient import FrappeClient
import json
import time

def create_sales_invoice():
    print("CREATING SALES Invoice ...")
    #conn = FrappeClient("HOST-Name", "username", "Password")

    import pandas as pd

    df = pd.read_excel('May 2020 cash -Jeddah.xlsx',dtype=(str,int))
    pd.set_option('display.max_rows', 10)
    print(df)
    finalList = []
    finalDict = {}
    grouped = df.groupby(level=0)
    for key, value in grouped:
    
    
        dictionary = {}
    
        j = grouped.get_group(key).reset_index()
        dictionary['doctype'] = j.at[0, 'doctype']
        dictionary['posting_date'] = j.at[0, 'posting_date']
        
        dictionary['customer'] = j.at[0, 'customer']
        dictionary['customer_group'] = j.at[0, 'customer_group']
        
        dictionary['channel'] = j.at[0, 'channel']
        dictionary['naming_series'] = j.at[0, 'naming_series']
        
        dictionary['status'] = j.at[0, 'status']
        dictionary['set_posting_time'] = j.at[0, 'set_posting_time']
        
        dictionary['set_warehouse'] = j.at[0, 'set_warehouse']
        dictionary['due_date'] = j.at[0, 'due_date']
        
        dictionary['update_stock'] = j.at[0, 'update_stock']
    
        dictList = []
        dictList1 = []
        dictList2 = []

        anotherDict1 = {}
        anotherDict15 = {}
        anotherDict2 = {} # items  taxes
        
        finallistitems = []
        ### 330 ###
        anotherDict301={}
        anotherDict302={}
        dictList301 = []
        dictList302 = []
        for i in j.index:
            if (j.at[i, 'qty330']) != '0':
                anotherDict301['doctype'] = j.at[i, 'doctype1']
                anotherDict301['qty'] = j.at[i, 'qty330']
                
                anotherDict301['delivered_qty'] = j.at[i, 'qty330']
                anotherDict301['item_code'] = j.at[i, 'item_code330']
        
                anotherDict301['rate'] = j.at[i, 'rate330']
                anotherDict301['uom'] = j.at[i, 'uom']
                
                anotherDict301['cost_center'] = j.at[i, 'cost_center']
        
                dictList301.append(anotherDict301)
            
        
        for i in j.index:
            
            if (j.at[i, 'qty330foc']) != '0':
                anotherDict302['doctype'] = j.at[i, 'doctype1']
                anotherDict302['qty'] = j.at[i, 'qty330foc']
                
                anotherDict302['delivered_qty'] = j.at[i, 'qty330foc']
                anotherDict302['item_code'] = j.at[i, 'item_code330']
        
                anotherDict302['rate'] = j.at[i, 'ratefoc']
                anotherDict302['uom'] = j.at[i, 'uom']
                
                anotherDict302['cost_center'] = j.at[i, 'cost_center']
        
                dictList302.append(anotherDict302)
        
        if (j.at[i, 'qty330foc']) != '0' and (j.at[i, 'qty330']) != '0': 
            finallistitems.extend(dictList301 + dictList302)
        elif (j.at[i, 'qty330']) != '0':
            finallistitems.extend(dictList301)
        elif (j.at[i, 'qty330foc']) != '0':
            finallistitems.extend(dictList302)
        ### 330f ###
        anotherDict301f={}
        anotherDict302f={}
        dictList301f = []
        dictList302f = []
        for i in j.index:
            if (j.at[i, 'qty330f']) != '0':
                anotherDict301f['doctype'] = j.at[i, 'doctype1']
                anotherDict301f['qty'] = j.at[i, 'qty330f']
                
                anotherDict301f['delivered_qty'] = j.at[i, 'qty330f']
                anotherDict301f['item_code'] = j.at[i, 'item_code330f']
        
                anotherDict301f['rate'] = j.at[i, 'rate330f']
                anotherDict301f['uom'] = j.at[i, 'uom']
                
                anotherDict301f['cost_center'] = j.at[i, 'cost_center']
        
                dictList301f.append(anotherDict301f)
            
        
        for i in j.index:
            
            if (j.at[i, 'qty330ffoc']) != '0':
                anotherDict302f['doctype'] = j.at[i, 'doctype1']
                anotherDict302f['qty'] = j.at[i, 'qty330ffoc']
                
                anotherDict302f['delivered_qty'] = j.at[i, 'qty330ffoc']
                anotherDict302f['item_code'] = j.at[i, 'item_code330f']
        
                anotherDict302f['rate'] = j.at[i, 'ratefoc']
                anotherDict302f['uom'] = j.at[i, 'uom']
                
                anotherDict302f['cost_center'] = j.at[i, 'cost_center']
        
                dictList302f.append(anotherDict302f)
        
        if (j.at[i, 'qty330ffoc']) != '0' and (j.at[i, 'qty330f']) != '0': 
            finallistitems.extend(dictList301f + dictList302f)
        elif (j.at[i, 'qty330f']) != '0':
            finallistitems.extend(dictList301f)
        elif (j.at[i, 'qty330ffoc']) != '0':
            finallistitems.extend(dictList302f)

        ### 500 ###
        anotherDict501={}
        anotherDict502={}
        dictList501 = []
        dictList502 = []
        for i in j.index:
            if (j.at[i, 'qty500']) != '0':
                anotherDict501['doctype'] = j.at[i, 'doctype1']
                anotherDict501['qty'] = j.at[i, 'qty500']
                
                anotherDict501['delivered_qty'] = j.at[i, 'qty500']
                anotherDict501['item_code'] = j.at[i, 'item_code500']
        
                anotherDict501['rate'] = j.at[i, 'rate500']
                anotherDict501['uom'] = j.at[i, 'uom']
                
                anotherDict501['cost_center'] = j.at[i, 'cost_center']
        
                dictList501.append(anotherDict501)
            
        
        for i in j.index:
            
            if (j.at[i, 'qty500foc']) != '0':
                anotherDict502['doctype'] = j.at[i, 'doctype1']
                anotherDict502['qty'] = j.at[i, 'qty500foc']
                
                anotherDict502['delivered_qty'] = j.at[i, 'qty500foc']
                anotherDict502['item_code'] = j.at[i, 'item_code500']
        
                anotherDict502['rate'] = j.at[i, 'ratefoc']
                anotherDict502['uom'] = j.at[i, 'uom']
                
                anotherDict502['cost_center'] = j.at[i, 'cost_center']
        
                dictList502.append(anotherDict502)
        if (j.at[i, 'qty500foc']) != '0' and (j.at[i, 'qty500']) != '0': 
            finallistitems.extend(dictList501 + dictList502)
        elif (j.at[i, 'qty500']) != '0':
            finallistitems.extend(dictList501)
        elif (j.at[i, 'qty500foc']) != '0':
            finallistitems.extend(dictList502)

        
        ### 200 RE ###
        anotherDict200={}
        anotherDict201={}
        dictList201 = []
        dictList202 = []
        for i in j.index:
            if (j.at[i, 'qty200']) != '0':
                anotherDict200['doctype'] = j.at[i, 'doctype1']
                anotherDict200['qty'] = j.at[i, 'qty200']
                
                anotherDict200['delivered_qty'] = j.at[i, 'qty200']
                anotherDict200['item_code'] = j.at[i, 'item_code200']
        
                anotherDict200['rate'] = j.at[i, 'rate200']
                anotherDict200['uom'] = j.at[i, 'uom']
                
                anotherDict200['cost_center'] = j.at[i, 'cost_center']
        
                dictList201.append(anotherDict200)
            
        
        for i in j.index:
            
            if (j.at[i, 'qty200foc']) != '0':
                anotherDict201['doctype'] = j.at[i, 'doctype1']
                anotherDict201['qty'] = j.at[i, 'qty200foc']
                
                anotherDict201['delivered_qty'] = j.at[i, 'qty200foc']
                anotherDict201['item_code'] = j.at[i, 'item_code200']
        
                anotherDict201['rate'] = j.at[i, 'ratefoc']
                anotherDict201['uom'] = j.at[i, 'uom']
                
                anotherDict201['cost_center'] = j.at[i, 'cost_center']
        
                dictList202.append(anotherDict201)
        if (j.at[i, 'qty200foc']) != '0' and (j.at[i, 'qty200']) != '0': 
            finallistitems.extend(dictList202 + dictList201)
        elif (j.at[i, 'qty200']) != '0':
            finallistitems.extend(dictList201)
        elif (j.at[i, 'qty200foc']) != '0':
            finallistitems.extend(dictList202)

        ### 200 Frid ###
        anotherDict201shf={}
        anotherDict202shf={}
        dictList201shf = []
        dictList202shf = []
        
        for i in j.index:
            if (j.at[i, 'qty200f']) != '0':
                anotherDict201shf['doctype'] = j.at[i, 'doctype1']
                anotherDict201shf['qty'] = j.at[i, 'qty200f']
                
                anotherDict201shf['delivered_qty'] = j.at[i, 'qty200f']
                anotherDict201shf['item_code'] = j.at[i, 'item_code200f']
        
                anotherDict201shf['rate'] = j.at[i, 'rate200f']
                anotherDict201shf['uom'] = j.at[i, 'uom']
                
                anotherDict201shf['cost_center'] = j.at[i, 'cost_center']
        
                dictList201shf.append(anotherDict201shf)
            
        
        for i in j.index:
            
            if (j.at[i, 'qty200ffoc']) != '0':
                anotherDict202shf['doctype'] = j.at[i, 'doctype1']
                anotherDict202shf['qty'] = j.at[i, 'qty200ffoc']
                
                anotherDict202shf['delivered_qty'] = j.at[i, 'qty200ffoc']
                anotherDict202shf['item_code'] = j.at[i, 'item_code200f']
        
                anotherDict202shf['rate'] = j.at[i, 'ratefoc']
                anotherDict202shf['uom'] = j.at[i, 'uom']
                
                anotherDict202shf['cost_center'] = j.at[i, 'cost_center']
        
                dictList202shf.append(anotherDict202shf)
                
        if (j.at[i, 'qty200ffoc']) != '0' and (j.at[i, 'qty200f']) != '0': 
            finallistitems.extend(dictList201shf + dictList202shf)
        elif (j.at[i, 'qty200f']) != '0':
            finallistitems.extend(dictList201shf)
        elif (j.at[i, 'qty200ffoc']) != '0':
            finallistitems.extend(dictList202shf)
            
        dictionary['items'] = finallistitems
        ###############################################
        for i in j.index:
    
            anotherDict2['charge_type'] = j.at[i, 'charge_type']
            anotherDict2['account_head'] = j.at[i, 'account_head']
            
            anotherDict2['rate'] = j.at[i, 'rateVAT']
            anotherDict2['description'] = j.at[i, 'description']
    
    
            dictList2.append(anotherDict2)
    
        dictionary['taxes'] = dictList2
    
    

        finalList.append(dictionary)

    
    counter=1
    for i in finalList:
        if 1 == 1:
            
            print(counter)
            
            print(i)

            conn = FrappeClient("Hostname", "Username", "Password")
            docc=conn.insert(i)
            print("done")
            time.sleep(1)
            counter += 1
if __name__=="__main__":

    create_sales_invoice()
