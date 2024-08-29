import pandas as pd
from datetime import datetime, timedelta
from functions.grouping import grouping

def counting(file_data, date, save_grouping, saved_as, is_grouped):
    if is_grouped == 0:
        df_awb = grouping(file_data=file_data, date=date, save_grouping=save_grouping, saved_as=saved_as)
    else:
        df_awb = pd.read_excel(file_data)
    

    all_data = []

    customer_name = ["SHOPEE", "TIKTOK", "TOKOPEDIA", "LAZADA", "ALL SHIPMENT"]

    # Count the data
    for name in range(0, 5):
        # Data for every sheet
        data1 = []
        for i in range(0, 8):
            # Lists for H+0 - H+7 data
            data2 = []
            for idx in range(1, 3):
                # Count data for SHOPEE, TIKTOK, TOKOPEDIA, LAZADA
                if name <= 3:
                    total_cnote = len(df_awb[(df_awb['RING AREA'] == f'RING {idx}') &
                                            (df_awb['ENTRY AWB'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y")) & (df_awb['CUSTOMER'] == customer_name[name])])
                    # Disabled temporarily
                    # cnote_cancel = len(df_cancel[(df_cancel['Transaction date'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y")) & (df_cancel['RING'] == f'RING {idx}') & (df_cancel['CUSTOMER'] == customer_name[name])])
                    cnote_unreceiving = len(df_awb[(df_awb['STATUS POD'] == "UNRECEIVING") & (df_awb['RING AREA'] == f'RING {idx}') & (
                        df_awb['ENTRY AWB'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y")) & (df_awb['CUSTOMER'] == customer_name[name])])
                    cnote_unmanifest = len(df_awb[(df_awb['STATUS POD'] == "UNMANIFEST") & (df_awb['RING AREA'] == f'RING {idx}') & (
                        df_awb['ENTRY AWB'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y")) & (df_awb['CUSTOMER'] == customer_name[name])])
                    cnote_unappv_om = len(df_awb[(df_awb['MANIFEST APPROVED'] == 'N') & (df_awb['RING AREA'] == f'RING {idx}') & (df_awb['STATUS MANIFEST 2'] == "BUTUH MANIFEST") & (
                        df_awb['ENTRY AWB'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y")) & (df_awb['CUSTOMER'] == customer_name[name])])
                    cnote_unsmu = len(df_awb[(df_awb['SM NO'] == '-') & (df_awb['RING AREA'] == f'RING {idx}') & (df_awb['STATUS MANIFEST 2'] == "BUTUH MANIFEST") & (
                        df_awb['ENTRY AWB'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y")) & (df_awb['CUSTOMER'] == customer_name[name])])
                    # Disabled temporarily
                    # final_connote = total_cnote - cnote_cancel
                # Count data for ALL SHIPMENT
                else:
                    total_cnote = len(df_awb[(df_awb['RING AREA'] == f'RING {idx}') &
                                            (df_awb['ENTRY AWB'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y"))])
                    # Disabled temporarily
                    # cnote_cancel = len(df_cancel[(df_cancel['Transaction date']== (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y")) & (df_cancel['RING'] == f'RING {idx}')])
                    cnote_unreceiving = len(df_awb[(df_awb['STATUS POD'] == "UNRECEIVING") & (df_awb['RING AREA'] == f'RING {idx}') & (
                        df_awb['ENTRY AWB'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y"))])
                    cnote_unmanifest = len(df_awb[(df_awb['STATUS POD'] == "UNMANIFEST") & (df_awb['RING AREA'] == f'RING {idx}') & (
                        df_awb['ENTRY AWB'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y"))])
                    cnote_unappv_om = len(df_awb[(df_awb['MANIFEST APPROVED'] == 'N') & (df_awb['RING AREA'] == f'RING {idx}') & (df_awb['STATUS MANIFEST 2'] == "BUTUH MANIFEST") & (
                        df_awb['ENTRY AWB'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y"))])
                    cnote_unsmu = len(df_awb[(df_awb['SM NO'] == '-') & (df_awb['RING AREA'] == f'RING {idx}') & (df_awb['STATUS MANIFEST 2'] == "BUTUH MANIFEST") & (
                        df_awb['ENTRY AWB'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y"))])
                    # Disabled temporarily
                    # final_connote = total_cnote - cnote_cancel
                data2.append([0, cnote_unreceiving, cnote_unmanifest, cnote_unappv_om, cnote_unsmu,
                            total_cnote, 0])
            data1.append(data2)
        all_data.append(data1)
    
    # Return the result
    return all_data