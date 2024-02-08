#!/usr/bin/env python
# coding: utf-8

# In[1]:


import numpy as np
import requests
import pandas as pd
import matplotlib.pyplot as plt
import time
import xlwings as xw

# Initial row index
current_row = 5

while True:
    try:
        url = "https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY"

        headers = {
            "accept-encoding": ": gzip, deflate, br",
            "accept-language": "en-US,en;q=0.9",
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36"
        }

        session = requests.Session()
        data = session.get(url, headers=headers).json()["records"]["data"]
        ocdata = []
        for i in data:
            for j, k in i.items():
                if j == "CE" or j == "PE":
                    info = k
                    info["instrument Type"] = j
                    ocdata.append(info)
        df = pd.DataFrame(ocdata)

        # Extract rows where 'ce' is equal to desired_name
        rows_with_desired_name_ce = df[df['instrument Type'] == 'CE'].copy()
        if not rows_with_desired_name_ce.empty:
            print("\nRows with name '{}':".format('CE'))
        else:
            print("\nRows with name '{}' not found.".format('CE'))

        # Extract rows where 'pe' is equal to desired_name
        rows_with_desired_name_pe = df[df['instrument Type'] == 'PE'].copy()
        if not rows_with_desired_name_pe.empty:
            print("\nRows with name '{}':".format('PE'))
        else:
            print("\nRows with name '{}' not found.".format('PE'))

        ###################################################################################################
        rows_with_desired_name_ce['1'] = rows_with_desired_name_ce['bidQty'] * rows_with_desired_name_ce[
            'bidprice']
        rows_with_desired_name_ce['2'] = rows_with_desired_name_ce['askQty'] * rows_with_desired_name_ce[
            'askPrice']
        rows_with_desired_name_pe['3'] = rows_with_desired_name_pe['bidQty'] * rows_with_desired_name_pe[
            'bidprice']
        rows_with_desired_name_pe['4'] = rows_with_desired_name_pe['askQty'] * rows_with_desired_name_pe[
            'askPrice']

        call_buy_sum = rows_with_desired_name_ce['1'].sum()
        call_sell_sum = rows_with_desired_name_ce['2'].sum()
        put_buy_sum = rows_with_desired_name_pe['3'].sum()
        put_sell_sum = rows_with_desired_name_pe['4'].sum()

        final = [call_buy_sum, call_sell_sum, put_buy_sum, put_sell_sum]
        sub_call_result = pd.DataFrame({'': final})

        labels = sub_call_result .index
        values = sub_call_result.values.flatten()  
        # Plotting the bar chart
        plt.bar(labels, values, color='black')
        plt.xlabel('Categories')
        plt.ylabel('Values')
        plt.title('booking price difference in ask-bid quantity for call and put')
        
        print('total:')
        print(final)
       # Display the plot
        plt.show()

        
        

        rows_with_desired_name_ce['5'] = rows_with_desired_name_ce['bidQty']
        rows_with_desired_name_ce['6'] = rows_with_desired_name_ce['askQty']
        rows_with_desired_name_pe['7'] = rows_with_desired_name_pe['bidQty']
        rows_with_desired_name_pe['8'] = rows_with_desired_name_pe['askQty']

        call_bid = rows_with_desired_name_ce['5'].sum()
        call_ask = rows_with_desired_name_ce['6'].sum()
        put_bid = rows_with_desired_name_pe['7'].sum()
        put_ask = rows_with_desired_name_pe['8'].sum()
        qty = [call_bid, call_ask, put_bid, put_ask]

        # Display the plot
        print(qty)
        
        result = pd.DataFrame({'': qty})

        labels = result.index
        values = result.values.flatten()
       # Plotting the bar chart
        plt.bar(labels, values, color='red')

    # Adding labels and title
        plt.xlabel('Categories')
        plt.ylabel('Values')
        plt.title('all quantity')

    # Display the plot

        plt.show()

        TOTAL1 = rows_with_desired_name_pe['openInterest'].sum()
        TOTAL2 = rows_with_desired_name_ce['openInterest'].sum()

        pcr = TOTAL1 / TOTAL2
        rounded_number_pcr = round(pcr, 4)

        TOTAL3 = rows_with_desired_name_pe['totalTradedVolume'].sum()
        TOTAL4 = rows_with_desired_name_ce['totalTradedVolume'].sum()

        volume = TOTAL3 / TOTAL4
        rounded_number_volume = round(volume, 4)
        desired_value = df['underlyingValue'].iloc[1]
        
        call_iv = rows_with_desired_name_ce['impliedVolatility'].sum()
        put_iv = rows_with_desired_name_pe['impliedVolatility'].sum()

        print('pcr:')
        print(rounded_number_pcr)
        print('volume:')
        print(rounded_number_volume)
######################################################################################################        
       
       
        
        identifiers_1=['21500']
        identifiers_2=['21600']
        identifiers_3=['21700']
        identifiers_4=['21800']
        identifiers_5=['21900']
        identifiers_6=['22000']
        identifiers_7=['22100']
        identifiers_8=['22200']
        identifiers_9=['21950']
        identifiers_10=['22300']
        
        
        oi_1c = rows_with_desired_name_ce[rows_with_desired_name_ce['strikePrice'].astype(str).isin(identifiers_1)]
        oi_1p = rows_with_desired_name_pe[rows_with_desired_name_pe['strikePrice'].astype(str).isin(identifiers_1)]
        oi_2c = rows_with_desired_name_ce[rows_with_desired_name_ce['strikePrice'].astype(str).isin(identifiers_2)]
        oi_2p = rows_with_desired_name_pe[rows_with_desired_name_pe['strikePrice'].astype(str).isin(identifiers_2)]
        oi_3c = rows_with_desired_name_ce[rows_with_desired_name_ce['strikePrice'].astype(str).isin(identifiers_3)]
        oi_3p = rows_with_desired_name_pe[rows_with_desired_name_pe['strikePrice'].astype(str).isin(identifiers_3)]
        oi_4c = rows_with_desired_name_ce[rows_with_desired_name_ce['strikePrice'].astype(str).isin(identifiers_4)]
        oi_4p = rows_with_desired_name_pe[rows_with_desired_name_pe['strikePrice'].astype(str).isin(identifiers_4)]
        
        oi_5c = rows_with_desired_name_ce[rows_with_desired_name_ce['strikePrice'].astype(str).isin(identifiers_5)]
        oi_5p = rows_with_desired_name_pe[rows_with_desired_name_pe['strikePrice'].astype(str).isin(identifiers_5)]
        oi_6c = rows_with_desired_name_ce[rows_with_desired_name_ce['strikePrice'].astype(str).isin(identifiers_6)]
        oi_6p = rows_with_desired_name_pe[rows_with_desired_name_pe['strikePrice'].astype(str).isin(identifiers_6)]
        oi_7c = rows_with_desired_name_ce[rows_with_desired_name_ce['strikePrice'].astype(str).isin(identifiers_7)]
        oi_7p = rows_with_desired_name_pe[rows_with_desired_name_pe['strikePrice'].astype(str).isin(identifiers_7)]
        oi_8c = rows_with_desired_name_ce[rows_with_desired_name_ce['strikePrice'].astype(str).isin(identifiers_8)]
        oi_8p = rows_with_desired_name_pe[rows_with_desired_name_pe['strikePrice'].astype(str).isin(identifiers_8)]
        oi_9c = rows_with_desired_name_ce[rows_with_desired_name_ce['strikePrice'].astype(str).isin(identifiers_9)]
        oi_9p = rows_with_desired_name_pe[rows_with_desired_name_pe['strikePrice'].astype(str).isin(identifiers_9)]
        oi_10c = rows_with_desired_name_ce[rows_with_desired_name_ce['strikePrice'].astype(str).isin(identifiers_10)]
        oi_10p = rows_with_desired_name_pe[rows_with_desired_name_pe['strikePrice'].astype(str).isin(identifiers_10)]
        
        
        
        
        
        
        
        c_oi_1 = oi_1c['openInterest'].sum()
        p_oi_1 = oi_1p['openInterest'].sum()
        c_oi_2 = oi_2c['openInterest'].sum()
        p_oi_2 = oi_2p['openInterest'].sum()
        c_oi_3 = oi_3c['openInterest'].sum()
        p_oi_3 = oi_3p['openInterest'].sum()
        c_oi_4 = oi_4c['openInterest'].sum()
        p_oi_4 = oi_4p['openInterest'].sum()
        c_oi_5 = oi_5c['openInterest'].sum()
        p_oi_5 = oi_5p['openInterest'].sum()
        c_oi_6 = oi_6c['openInterest'].sum()
        p_oi_6 = oi_6p['openInterest'].sum()
        c_oi_7 = oi_7c['openInterest'].sum()
        p_oi_7 = oi_7p['openInterest'].sum()
        c_oi_8 = oi_8c['openInterest'].sum()
        p_oi_8 = oi_8p['openInterest'].sum()
        c_oi_9 = oi_9c['openInterest'].sum()
        p_oi_9 = oi_9p['openInterest'].sum()
        c_oi_10 = oi_10c['openInterest'].sum()
        p_oi_10 = oi_10p['openInterest'].sum()
        
        c_vol_1 = oi_1c['totalTradedVolume'].sum()
        p_vol_1 = oi_1p['totalTradedVolume'].sum()
        c_vol_2 =  oi_2c['totalTradedVolume'].sum()
        p_vol_2 = oi_2p['totalTradedVolume'].sum()
        c_vol_3 =  oi_3c['totalTradedVolume'].sum()
        p_vol_3 = oi_3p['totalTradedVolume'].sum()
        c_vol_4 = oi_4c['totalTradedVolume'].sum()
        p_vol_4 = oi_4p['totalTradedVolume'].sum()
        
        c_vol_5 = oi_5c['totalTradedVolume'].sum()
        p_vol_5 = oi_5p['totalTradedVolume'].sum()
        c_vol_6 = oi_6c['totalTradedVolume'].sum()
        p_vol_6 = oi_6p['totalTradedVolume'].sum()
        c_vol_7 = oi_7c['totalTradedVolume'].sum()
        p_vol_7 = oi_7p['totalTradedVolume'].sum()
        c_vol_8 = oi_8c['totalTradedVolume'].sum()
        p_vol_8 = oi_8p['totalTradedVolume'].sum()
        c_vol_9 = oi_9c['totalTradedVolume'].sum()
        p_vol_9 = oi_9p['totalTradedVolume'].sum()
        c_vol_10 = oi_10c['totalTradedVolume'].sum()
        p_vol_10 = oi_10p['totalTradedVolume'].sum()
        
        oi_data =[{'value': c_oi_1, 'strikePrice': oi_1c['strikePrice'].iloc[0]},
                  {'value': c_oi_2, 'strikePrice': oi_2c['strikePrice'].iloc[0]},
                  {'value': c_oi_3, 'strikePrice': oi_3c['strikePrice'].iloc[0]},
                  {'value': c_oi_4, 'strikePrice': oi_4c['strikePrice'].iloc[0]},
                  {'value': c_oi_5, 'strikePrice': oi_5c['strikePrice'].iloc[0]},
                  {'value': c_oi_6, 'strikePrice': oi_6c['strikePrice'].iloc[0]},
                  {'value': c_oi_7, 'strikePrice': oi_7c['strikePrice'].iloc[0]},
                  {'value': c_oi_8, 'strikePrice': oi_8c['strikePrice'].iloc[0]},
                  {'value': c_oi_9, 'strikePrice': oi_9c['strikePrice'].iloc[0]},
                  {'value': c_oi_10, 'strikePrice': oi_10c['strikePrice'].iloc[0]}]
        
        # Find the dictionary with the highest value
        max_oi_data = max(oi_data, key=lambda x: x['value'])

        print("Highest open interest value:", max_oi_data['value'])
        print("Corresponding strike price:", max_oi_data['strikePrice'])
        
        
        # Sort the list of dictionaries by value in descending order
        sorted_oi_data = sorted(oi_data, key=lambda x: x['value'], reverse=True)

# Find the dictionary with the second-highest value
        second_max_oi_data = sorted_oi_data[1] if len(sorted_oi_data) > 1 else None

        if second_max_oi_data:
            print("Second-highest open interest value:", second_max_oi_data['value'])
            print("Corresponding strike price:", second_max_oi_data['strikePrice'])
        else:
            print("There is no second-highest value.")
            
            
        oi_data_1 = [{'value': p_oi_1, 'strikePrice': oi_1p['strikePrice'].iloc[0]},
                     {'value': p_oi_2, 'strikePrice': oi_2p['strikePrice'].iloc[0]},
                     {'value': p_oi_3, 'strikePrice': oi_3p['strikePrice'].iloc[0]},
                     {'value': p_oi_4, 'strikePrice': oi_4p['strikePrice'].iloc[0]},
                     {'value': p_oi_5, 'strikePrice': oi_5p['strikePrice'].iloc[0]},
                     {'value': p_oi_6, 'strikePrice': oi_6p['strikePrice'].iloc[0]},
                     {'value': p_oi_7, 'strikePrice': oi_7p['strikePrice'].iloc[0]},
                     {'value': p_oi_8, 'strikePrice': oi_8p['strikePrice'].iloc[0]},
                     {'value': p_oi_9, 'strikePrice': oi_9p['strikePrice'].iloc[0]},
                     {'value': p_oi_10, 'strikePrice': oi_10p['strikePrice'].iloc[0]},]

# Find the dictionary with the highest value
        max_oi_data_1 = max(oi_data_1, key=lambda x: x['value'])

        print("Highest open interest value_p:", max_oi_data_1['value'])
        print("Corresponding strike price_p:", max_oi_data_1['strikePrice'])

# Sort the list of dictionaries by value in descending order
        sorted_oi_data_1 = sorted(oi_data_1, key=lambda x: x['value'], reverse=True)

# Find the dictionary with the second-highest value
        second_max_oi_data_1 = sorted_oi_data_1[1] if len(sorted_oi_data_1) > 1 else None

        if second_max_oi_data_1:
            print("Second-highest open interest value_p:", second_max_oi_data_1['value'])
            print("Corresponding strike price_p:", second_max_oi_data_1['strikePrice'])
        else:
            print("There is no second-highest value.")
            
            
###################################################################################################################
        vol_data = [{'value': c_vol_1, 'strikePrice': oi_1c['strikePrice'].iloc[0]},
                    {'value': c_vol_2, 'strikePrice': oi_2c['strikePrice'].iloc[0]},
                    {'value': c_vol_3, 'strikePrice': oi_3c['strikePrice'].iloc[0]},
                    {'value': c_vol_4, 'strikePrice': oi_4c['strikePrice'].iloc[0]},
                    {'value': c_vol_5, 'strikePrice': oi_5c['strikePrice'].iloc[0]},
                    {'value': c_vol_6, 'strikePrice': oi_6c['strikePrice'].iloc[0]},
                    {'value': c_vol_7, 'strikePrice': oi_7c['strikePrice'].iloc[0]},
                    {'value': c_vol_8, 'strikePrice': oi_8c['strikePrice'].iloc[0]},
                    {'value': c_vol_9, 'strikePrice': oi_9c['strikePrice'].iloc[0]},
                    {'value': c_vol_10, 'strikePrice': oi_10c['strikePrice'].iloc[0]},]
        
        # Find the dictionary with the highest value
        max_vol_data = max(vol_data, key=lambda x: x['value'])

        print("Highest volume value:", max_vol_data['value'])
        print("Corresponding strike price:", max_vol_data['strikePrice'])
        
        
        # Sort the list of dictionaries by value in descending order
        sorted_vol_data = sorted(vol_data, key=lambda x: x['value'], reverse=True)

# Find the dictionary with the second-highest value
        second_max_vol_data = sorted_vol_data[1] if len(sorted_vol_data) > 1 else None

        if second_max_vol_data:
            print("Second-highest volume value:", second_max_vol_data['value'])
            print("Corresponding strike price:", second_max_vol_data['strikePrice'])
        else:
            print("There is no second-highest value.")
            
      # Find the dictionary with the third-highest value
        third_max_vol_data = sorted_vol_data[2] if len(sorted_vol_data) > 2 else None

        if third_max_vol_data:
            print("Third-highest volume value:", third_max_vol_data['value'])
            print("Corresponding strike price:", third_max_vol_data['strikePrice'])
        else:
            print("There is no third-highest value.")   
            
            
            
        vol_data_1 = [{'value': p_vol_1, 'strikePrice': oi_1p['strikePrice'].iloc[0]},
                     {'value': p_vol_2, 'strikePrice': oi_2p['strikePrice'].iloc[0]},
                     {'value': p_vol_3, 'strikePrice': oi_3p['strikePrice'].iloc[0]},
                     {'value': p_vol_4, 'strikePrice': oi_4p['strikePrice'].iloc[0]},
                     {'value': p_vol_5, 'strikePrice': oi_5p['strikePrice'].iloc[0]},
                     {'value': p_vol_6, 'strikePrice': oi_6p['strikePrice'].iloc[0]},
                     {'value': p_vol_7, 'strikePrice': oi_7p['strikePrice'].iloc[0]},
                     {'value': p_vol_8, 'strikePrice': oi_8p['strikePrice'].iloc[0]},
                     {'value': p_vol_9, 'strikePrice': oi_9p['strikePrice'].iloc[0]},
                     {'value': p_vol_10, 'strikePrice': oi_10p['strikePrice'].iloc[0]},]

# Find the dictionary with the highest value
        max_vol_data_1 = max(vol_data_1, key=lambda x: x['value'])

        print("Highest volume value_p:", max_vol_data_1['value'])
        print("Corresponding strike price_p:", max_vol_data_1['strikePrice'])

# Sort the list of dictionaries by value in descending order
        sorted_vol_data_1 = sorted(vol_data_1, key=lambda x: x['value'], reverse=True)

# Find the dictionary with the second-highest value
        second_max_vol_data_1 = sorted_vol_data_1[1] if len(sorted_vol_data_1) > 1 else None

        if second_max_vol_data_1:
            print("Second-highest volume value_p:", second_max_vol_data_1['value'])
            print("Corresponding strike price_p:", second_max_vol_data_1['strikePrice'])
        else:
            print("There is no second-highest value.")
            
        # Find the dictionary with the third-highest value
        third_max_vol_data_1 = sorted_vol_data_1[2] if len(sorted_vol_data_1) > 2 else None

        if third_max_vol_data_1:
            print("Third-highest volume value_p:", third_max_vol_data_1['value'])
            print("Corresponding strike price_p:", third_max_vol_data_1['strikePrice'])
        else:
            print("There is no third-highest value_p.")
    
        

 ####################################################################################################
        wb = xw.Book('optionchain.xlsx')
        st = wb.sheets('MY')
        
        # Write values to the current row
        st.range(f'A{current_row}').value = rounded_number_pcr
        st.range(f'B{current_row}').value = rounded_number_volume
        st.range(f'C{current_row}').value = desired_value
       
        
      
       
        
        st.range(f'D{current_row}').value =  max_oi_data['strikePrice'] 
        st.range(f'E{current_row}').value =   np.round(max_oi_data['value']*50/100000, 2)
        st.range(f'F{current_row}').value =   second_max_oi_data['strikePrice']
        st.range(f'G{current_row}').value =    np.round(second_max_oi_data['value']*50/100000, 2)
        
        st.range(f'I{current_row}').value =   max_oi_data_1['strikePrice']
        st.range(f'J{current_row}').value =    np.round(max_oi_data_1['value']*50/100000, 2)
        st.range(f'K{current_row}').value =   second_max_oi_data_1['strikePrice']
        st.range(f'L{current_row}').value =    np.round(second_max_oi_data_1['value']*50/100000, 2)
        
        
        st.range(f'N{current_row}').value =  max_vol_data['strikePrice']
        st.range(f'O{current_row}').value =   np.round(max_vol_data['value']*50/10000000, 2)
        st.range(f'P{current_row}').value =   second_max_vol_data['strikePrice']
        st.range(f'Q{current_row}').value =    np.round(second_max_vol_data['value']*50/10000000, 2)
        st.range(f'R{current_row}').value =  third_max_vol_data['strikePrice']
        st.range(f'S{current_row}').value =   np.round(third_max_vol_data['value']*50/10000000, 2)
        
        
        st.range(f'U{current_row}').value =   max_vol_data_1['strikePrice']
        st.range(f'V{current_row}').value =    np.round(max_vol_data_1['value']*50/10000000, 2)
        st.range(f'W{current_row}').value =   second_max_vol_data_1['strikePrice']
        st.range(f'X{current_row}').value =    np.round(second_max_vol_data_1['value']*50/10000000, 2)
        st.range(f'Y{current_row}').value =   third_max_vol_data_1['strikePrice']
        st.range(f'Z{current_row}').value =    np.round(third_max_vol_data_1['value']*50/10000000, 2)
        #st.range(f'AA{current_row}').value =   np.round(max_vol_data['value']*50/10000000, 2)/np.round(second_max_vol_data['value']*50/10000000, 2)
        #st.range(f'AB{current_row}').value =   np.round(max_vol_data['value']*50/10000000, 2)/np.round(third_max_vol_data['value']*50/10000000, 2)
        #st.range(f'AC{current_row}').value =   np.round(max_vol_data_1['value']*50/10000000, 2)/np.round(second_max_vol_data_1['value']*50/10000000, 2)
        #st.range(f'AD{current_row}').value =   np.round(max_vol_data_1['value']*50/10000000, 2)/ np.round(third_max_vol_data_1['value']*50/10000000, 2)
         
        st.range(f'AA{current_row}').value =  np.round(call_bid*50/1000000, 2)
        st.range(f'AB{current_row}').value =  np.round(call_ask*50/1000000, 2)
        st.range(f'AC{current_row}').value =  np.round(put_bid*50/1000000, 2)
        st.range(f'AD{current_row}').value =  np.round(put_ask*50/1000000, 2)
        
        
        #st.range(f'AG{current_row}').value = np.round(max_vol_data['value']*50/10000000, 2)/np.round(second_max_vol_data['value']*50/10000000, 2)
        #st.range(f'AH{current_row}').value =  np.round(max_vol_data_1['value']*50/10000000, 2)/np.round(second_max_vol_data_1['value']*50/10000000, 2)
        #st.range(f'AI{current_row}').value = np.round(max_vol_data['value']*50/10000000, 2)+np.round(second_max_vol_data['value']*50/10000000, 2)/np.round(max_vol_data_1['value']*50/10000000, 2)/np.round(second_max_vol_data_1['value']*50/10000000, 2)
        

        # Increment the row index for the next iteration
        current_row += 1

        # Sleep for 5 seconds before the next iteration
        time.sleep(60)

    except Exception as e:
        print(f"An error occurred: {e}")
        # You might want to add more specific exception handling based on your requirements
        # Stop the loop or handle the error accordingly
        break


# In[ ]:





# In[ ]:




