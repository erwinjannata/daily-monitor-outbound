import re
import pandas as pd
import xlwings as xl
from tkinter.messagebox import showinfo
from user_id import user
from customer_id import customer
from datetime import datetime, timedelta


def grouping(file_data, date, save_grouping, saved_as):
    # Read data source
    data_awb = pd.read_excel(file_data, sheet_name='AWB')
    data_cancel = pd.read_excel(file_data, sheet_name='AWB CANCEL')
    df_awb = pd.DataFrame(data_awb)
    df_cancel = pd.DataFrame(data_cancel)

    # Read current date
    today = datetime.now().strftime("%#m/%#d/%Y")

    # Create new dataframe of filtered AWB data
    cabang = []
    status_receiving = []
    tgl_receiving = []
    status_manifest = []
    tgl_manifest = []
    tgl_entry_awb = []
    type_kiriman = []
    status_manifest_2 = []
    grouping_service = []
    status_pod = []
    date_runsheet = []
    jam_entry = []
    am_pm = []
    ring_area = []
    customer_grouping = []

    # Grouping Process
    for index in range(0, df_awb.shape[0]):
        # ---- CABANG
        cabang.append(user[str(df_awb['CNOTE USER ID'][index])])

        # ---- STATUS RECEIVING
        if not df_awb['RECEIVING NO'][index] == '-':
            status_receiving.append("RECEIVING")
        else:
            status_receiving.append("UNRECEIVING")

        # ---- TGL RECEIVING
        tgl_receiving.append(df_awb['RECEIVING DATE']
                             [index].strftime("%#m/%#d/%Y") or "")

        # ---- STATUS MANIFEST
        if not df_awb['MANIFEST OUTB '][index] == '-':
            status_manifest.append("MANIFESTED")
        else:
            status_manifest.append("UNMANIFEST")

        # ---- TGL MANIFEST
        try:
            tgl_manifest.append(df_awb['MANIFEST DATE']
                                [index].strftime("%#m/%#d/%Y"))
        except:
            tgl_manifest.append("")

        # ---- ENTRY AWB
        try:
            tgl_entry_awb.append(df_awb['CNOTE DATE']
                                 [index].strftime("%#m/%#d/%Y"))
        except:
            tgl_entry_awb.append("")

        # ---- TYPE KIRIMAN
        if df_awb['SERVICE'][index] == "P2P":
            type_kiriman.append('ROKET')
        else:
            type_kiriman.append(df_awb['Shipment Type'][index])

        # ---- STATUS MANIFEST
        if df_awb['Shipment Type'][index] == 'DOMESTIC':
            status_manifest_2.append('BUTUH MANIFEST')
        elif df_awb['Shipment Type'][index] == 'INTERCITY':
            status_manifest_2.append('BUTUH MANIFEST')
        else:
            status_manifest_2.append('NON MANIFEST')

        # ---- GROUPING SERVICE
        if "CTCJTR" in df_awb['SERVICE'][index]:
            grouping_service.append("CTC JTR")
        elif "INTL" in df_awb['SERVICE'][index]:
            grouping_service.append("INTERNASIONAL")
        elif "CML" in df_awb['SERVICE'][index]:
            grouping_service.append("CML")
        elif "CTC" in df_awb['SERVICE'][index]:
            grouping_service.append("CTC EXPRESS")
        elif "P2P" in df_awb['SERVICE'][index]:
            grouping_service.append("P2P")
        elif "P2P" in df_awb['SERVICE'][index]:
            grouping_service.append("P2P")
        elif "JTR" in df_awb['SERVICE'][index]:
            grouping_service.append("JTR")
        elif "TRC" in df_awb['SERVICE'][index]:
            grouping_service.append("JTR")
        elif "SPS" in df_awb['SERVICE'][index] or "YES" in df_awb['SERVICE'][index]:
            grouping_service.append("EXPRESS")
        elif "REG" in df_awb['SERVICE'][index] or "OKE" in df_awb['SERVICE'][index]:
            grouping_service.append("EXPRESS")
        else:
            grouping_service.append("BLANK")

        # ---- STATUS POD
        if df_awb['RECEIVING NO'][index] == '-':
            status_pod.append("UNRECEIVING")
        elif df_awb['Shipment Type'][index] == "INTERNASIONAL":
            status_pod.append("")
        elif df_awb['Shipment Type'][index] == "INTRACITY" and (df_awb['MANIFEST INB NO'][index] == '-' and df_awb['RUNSHEET NO'][index] == '-'):
            status_pod.append("UNRUNSHEET")
        elif (df_awb['Shipment Type'][index] == "INTERCITY" or df_awb['Shipment Type'][index] == "DOMESTIC") and df_awb['MANIFEST OUTB '][index] == '-':
            status_pod.append("UNMANIFEST")
        elif (df_awb['Shipment Type'][index] == "INTERCITY" or df_awb['Shipment Type'][index] == "DOMESTIC") and df_awb['MANIFEST INB NO'][index] == '-':
            status_pod.append("UNINBOUND")
        elif df_awb['RUNSHEET NO'][index] == '-':
            status_pod.append("UNRUNSHEET")
        elif (df_awb['RUNSHEET NO'][index] != '-' and df_awb['POD STATUS'][index] == '-') and (df_awb['RUNSHEET DATE'][index] != '-' and datetime.strptime(str(df_awb['RUNSHEET DATE'][index]), '%Y-%d-%m %H:%M:%S').strftime("%#m/%#d/%Y") == today):
            status_pod.append("ONDELIVERY")
        elif df_awb['RUNSHEET NO'][index] != '-' and df_awb['POD STATUS'][index] == '-':
            status_pod.append("OPEN")
        elif df_awb['POD STATUS'][index] == 'CR1':
            status_pod.append("RETURN")
        elif df_awb['POD STATUS'][index] == 'DP5':
            status_pod.append("PROBLEM")
        elif df_awb['POD STATUS'][index] == 'CL1':
            status_pod.append("CLOSE ORIGIN")
        elif df_awb['POD STATUS'][index] == 'CL2' or df_awb['POD STATUS'][index] == 'CL4':
            status_pod.append("CLOSE DESTINATION")
        elif (df_awb['POD STATUS'][index] == 'D25' or df_awb['POD STATUS'][index] == 'D26') or df_awb['POD STATUS'][index] == 'D37':
            status_pod.append("BREACH")
        elif (df_awb['POD STATUS'][index] == 'R24' or df_awb['POD STATUS'][index] == 'R25') or (df_awb['POD STATUS'][index] == 'R26' or df_awb['POD STATUS'][index] == 'R37'):
            status_pod.append("BREACH RETURN")
        elif df_awb['POD STATUS'][index] == 'UF':
            status_pod.append("IRREGULARITY")
        elif str(df_awb['POD STATUS'][index])[:2] == "CR":
            status_pod.append("CUSTOMER REQUEST")
        elif str(df_awb['POD STATUS'][index])[:2] == "PS":
            status_pod.append("IRREGULARITY")
        elif str(df_awb['POD STATUS'][index])[:1] == "U":
            status_pod.append("UNDEL")
        elif str(df_awb['POD STATUS'][index])[:1] == "D":
            status_pod.append("SUCCESS")
        elif str(df_awb['POD STATUS'][index])[:1] == "R":
            status_pod.append("SUCCESS RETURN")
        else:
            status_pod.append("")

        # ---- DATE RUNSHEET
        try:
            date_runsheet.append(df_awb['RUNSHEET DATE']
                                 [index].strftime("%#m/%#d/%Y"))
        except:
            date_runsheet.append("")

        # ---- JAM ENTRY
        if df_awb['CNOTE DATE'][index].strftime("%H:%M") > "18:00":
            jam_entry.append("NON")
        else:
            jam_entry.append("AM-PM")

        # ---- AM-PM
        if df_awb['CNOTE USER ID'][index] == 'RAHMAUL':
            am_pm.append("RETAIL")
        else:
            am_pm.append("CORPORATE")

        # ---- RING AREA
        try:
            if user[df_awb['CNOTE USER ID'][index]] == 'MATARAM':
                ring_area.append("RING 1")
            else:
                ring_area.append("RING 2")
        except:
            ring_area.append("")

        # ---- CUSTOMER
        try:
            customer_grouping.append(
                customer[str(df_awb['CUST NO'][index])])
        except:
            customer_grouping.append("")

    # Add the new column to dataframe
    df_awb.loc[:, "CABANG"] = cabang
    df_awb.loc[:, "STATUS RECEIVING"] = status_receiving
    df_awb.loc[:, "TGL RECEIVING"] = tgl_receiving
    df_awb.loc[:, "STATUS MANIFEST"] = status_manifest
    df_awb.loc[:, "TGL MANIFEST"] = tgl_manifest
    df_awb.loc[:, "ENTRY AWB"] = tgl_entry_awb
    df_awb.loc[:, "TYPE KIRIMAN"] = type_kiriman
    df_awb.loc[:, "STATUS MANIFEST 2"] = status_manifest_2
    df_awb.loc[:, "GRUPING SERVICE"] = grouping_service
    df_awb.loc[:, "STATUS POD"] = status_pod
    df_awb.loc[:, "DATE RUNSHEET"] = date_runsheet
    df_awb.loc[:, "JAM ENTRY"] = jam_entry
    df_awb.loc[:, "AM PM"] = am_pm
    df_awb.loc[:, "RING AREA"] = ring_area
    df_awb.loc[:, "CUSTOMER"] = customer_grouping

    # Create new dataframe of filtered AWB Cancel data
    cabang_cancel = []
    ring_area_cancel = []
    customer_grouping_cancel = []

    for index in range(0, df_cancel.shape[0]):
        # ----
        cabang_cancel.append(user[str(df_cancel['Cnote User'][index])])

        # ----
        try:
            if user[df_cancel['Cnote User'][index]] == 'MATARAM':
                ring_area_cancel.append("RING 1")
            else:
                ring_area_cancel.append("RING 2")
        except:
            ring_area.append("")

        # ----
        try:
            customer_grouping_cancel.append(
                customer[str(df_cancel['Agent id'][index])])
        except:
            customer_grouping_cancel.append("")

    # Add new column to dataframe
    df_cancel.loc[:, "CABANG"] = cabang_cancel
    df_cancel.loc[:, "RING"] = ring_area_cancel
    df_cancel.loc[:, "CUSTOMER"] = customer_grouping_cancel

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

                    cnote_cancel = len(df_cancel[(df_cancel['Transaction date']
                                                  == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y")) & (df_cancel['RING'] == f'RING {idx}') & (df_cancel['CUSTOMER'] == customer_name[name])])

                    cnote_unreceiving = len(df_awb[(df_awb['STATUS POD'] == "UNRECEIVING") & (df_awb['RING AREA'] == f'RING {idx}') & (
                        df_awb['ENTRY AWB'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y")) & (df_awb['CUSTOMER'] == customer_name[name])])

                    cnote_unmanifest = len(df_awb[(df_awb['STATUS POD'] == "UNMANIFEST") & (df_awb['RING AREA'] == f'RING {idx}') & (
                        df_awb['ENTRY AWB'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y")) & (df_awb['CUSTOMER'] == customer_name[name])])

                    cnote_unappv_om = len(df_awb[(df_awb['MANIFEST APPROVED'] == 'N') & (df_awb['RING AREA'] == f'RING {idx}') & (df_awb['STATUS MANIFEST 2'] == "BUTUH MANIFEST") & (
                        df_awb['ENTRY AWB'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y")) & (df_awb['CUSTOMER'] == customer_name[name])])

                    cnote_unsmu = len(df_awb[(df_awb['SM NO'] == '-') & (df_awb['RING AREA'] == f'RING {idx}') & (df_awb['STATUS MANIFEST 2'] == "BUTUH MANIFEST") & (
                        df_awb['ENTRY AWB'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y")) & (df_awb['CUSTOMER'] == customer_name[name])])

                    final_connote = total_cnote - cnote_cancel
                # Count data for ALL SHIPMENT
                else:
                    total_cnote = len(df_awb[(df_awb['RING AREA'] == f'RING {idx}') &
                                             (df_awb['ENTRY AWB'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y"))])

                    cnote_cancel = len(df_cancel[(df_cancel['Transaction date']
                                                  == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y")) & (df_cancel['RING'] == f'RING {idx}')])

                    cnote_unreceiving = len(df_awb[(df_awb['STATUS POD'] == "UNRECEIVING") & (df_awb['RING AREA'] == f'RING {idx}') & (
                        df_awb['ENTRY AWB'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y"))])

                    cnote_unmanifest = len(df_awb[(df_awb['STATUS POD'] == "UNMANIFEST") & (df_awb['RING AREA'] == f'RING {idx}') & (
                        df_awb['ENTRY AWB'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y"))])

                    cnote_unappv_om = len(df_awb[(df_awb['MANIFEST APPROVED'] == 'N') & (df_awb['RING AREA'] == f'RING {idx}') & (df_awb['STATUS MANIFEST 2'] == "BUTUH MANIFEST") & (
                        df_awb['ENTRY AWB'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y"))])

                    cnote_unsmu = len(df_awb[(df_awb['SM NO'] == '-') & (df_awb['RING AREA'] == f'RING {idx}') & (df_awb['STATUS MANIFEST 2'] == "BUTUH MANIFEST") & (
                        df_awb['ENTRY AWB'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=i)).strftime("%#m/%#d/%Y"))])

                    final_connote = total_cnote - cnote_cancel

                data2.append([cnote_cancel, cnote_unreceiving, cnote_unmanifest, cnote_unappv_om, cnote_unsmu,
                              total_cnote, final_connote])
            data1.append(data2)
        all_data.append(data1)

    # Append the filtered dataframe into existing excel file
    if save_grouping == True:
        with pd.ExcelWriter(saved_as, mode='a', if_sheet_exists='replace') as writer:
            df_awb.to_excel(writer, sheet_name='AWB', index=False)
            df_cancel.to_excel(writer, sheet_name='AWB CANCEL', index=False)

        showinfo(title="Message",
                 message=f"Proses selesai")

    return all_data
    # --------


def daily_monitor(file_data, file_report, date, saved_as, over_month):
    datas = grouping(file_data=file_data, date=date,
                     save_grouping=False, saved_as="")

    # Connect to book
    app = xl.App(visible=False)
    target_book = xl.Book(file_report)

    tanggal = int(date.split('/')[1])
    real_date = tanggal

    try:
        for item in range(0, 5):
            target_sheet = target_book.sheets[item]

            # Read worksheet spesification
            global max_row, merged_row, merged_col
            max_row = int(re.findall(
                r'\d+', (target_sheet.range("B4").end("down").address))[0])
            merged_row = target_sheet.range("A4").merge_area.count
            merged_col = target_sheet.range("C1").merge_area.count

            # Read condition if writing to previous month report
            if over_month == 1:
                cell_row = max_row + ((real_date - 1) * merged_row)
            else:
                cell_row = (real_date * merged_row)

            # Fill data for H+0 - H+7
            for i in range(0, 8):
                if cell_row >= 3 and cell_row < max_row:
                    for idx in range(0, 2):
                        # Cancel
                        target_sheet[(cell_row + idx), 2 + (merged_col * i)
                                     ].value = datas[item][i][idx][0]

                        # UN-RCC
                        target_sheet[(cell_row + idx), 3 + (merged_col * i)
                                     ].value = datas[item][i][idx][1]

                        # UN-OM
                        target_sheet[(cell_row + idx), 4 + (merged_col * i)
                                     ].value = datas[item][i][idx][2]

                        # UN-APPV OM
                        target_sheet[(cell_row + idx), 5 + (merged_col * i)
                                     ].value = datas[item][i][idx][3]

                        # UN-SMU
                        target_sheet[(cell_row + idx), 6 + (merged_col * i)
                                     ].value = datas[item][i][idx][4]

                        # TOTAL CONNOTE
                        target_sheet[(cell_row + idx), 7 + (merged_col * i)
                                     ].value = datas[item][i][idx][5]

                        # TOTAL CONNNOTE - CANCEL
                        target_sheet[(cell_row + idx), 8 + (merged_col * i)
                                     ].value = datas[item][i][idx][6]

                        # % CANCEL
                        target_sheet[(cell_row + idx), 9 + (merged_col * i)
                                     ].value = (datas[item][i][idx][0] / datas[item][i][idx][5]) if datas[item][i][idx][5] != 0 else 0

                        # % UN-RCC
                        target_sheet[(cell_row + idx), 10 + (merged_col * i)
                                     ].value = (datas[item][i][idx][1] / datas[item][i][idx][6]) if datas[item][i][idx][6] != 0 else 0

                        # % UN-OM
                        target_sheet[(cell_row + idx), 11 + (merged_col * i)
                                     ].value = (datas[item][i][idx][2] / datas[item][i][idx][6]) if datas[item][i][idx][6] != 0 else 0

                        # % UN-APPV OM
                        target_sheet[(cell_row + idx), 12 + (merged_col * i)
                                     ].value = (datas[item][i][idx][3] / datas[item][i][idx][6]) if datas[item][i][idx][6] != 0 else 0

                        # % UN-SMU
                        target_sheet[(cell_row + idx), 13 + (merged_col * i)
                                     ].value = (datas[item][i][idx][4] / datas[item][i][idx][6]) if datas[item][i][idx][6] != 0 else 0

                    # TOTAL PER RING
                    for idx in range(0, 7):
                        target_sheet[(cell_row + 2), (idx + 2) + (merged_col * i)
                                     ].value = datas[item][i][0][idx] + datas[item][i][1][idx]

                    # % TOTAL PER RING
                    # % CANCEL
                    target_sheet[(cell_row + 2), 9 + (merged_col * i)
                                 ].value = (target_sheet[(cell_row + 2), 2 + (merged_col * i)].value / target_sheet[(cell_row + 2), 7 + (merged_col * i)].value) if target_sheet[(cell_row + 2), 7 + (merged_col * i)].value != 0 else 0

                    # % UN-RCC
                    target_sheet[(cell_row + 2), 10 + (merged_col * i)
                                 ].value = (target_sheet[(cell_row + 2), 3 + (merged_col * i)].value / target_sheet[(cell_row + 2), 8 + (merged_col * i)].value) if target_sheet[(cell_row + 2), 8 + (merged_col * i)].value != 0 else 0

                    # % UN-OM
                    target_sheet[(cell_row + 2), 11 + (merged_col * i)
                                 ].value = (target_sheet[(cell_row + 2), 4 + (merged_col * i)].value / target_sheet[(cell_row + 2), 8 + (merged_col * i)].value) if target_sheet[(cell_row + 2), 8 + (merged_col * i)].value != 0 else 0

                    # % UN-APPV OM
                    target_sheet[(cell_row + 2), 12 + (merged_col * i)
                                 ].value = (target_sheet[(cell_row + 2), 5 + (merged_col * i)].value / target_sheet[(cell_row + 2), 8 + (merged_col * i)].value) if target_sheet[(cell_row + 2), 8 + (merged_col * i)].value != 0 else 0

                    # % UN-SMU
                    target_sheet[(cell_row + 2), 13 + (merged_col * i)
                                 ].value = (target_sheet[(cell_row + 2), 6 + (merged_col * i)].value / target_sheet[(cell_row + 2), 8 + (merged_col * i)].value) if target_sheet[(cell_row + 2), 8 + (merged_col * i)].value != 0 else 0

                    cell_row -= merged_row
                else:
                    cell_row -= merged_row
                    continue

        target_book.save(saved_as)
        target_book.close()
        showinfo(title="Message",
                 message=f"Proses selesai")
        app.quit()
    except Exception as e:
        target_book.close()
        app.quit()
        showinfo(title="Message",
                 message="Program mengalami masalah, silahkan hubungi tim IT")
        print(e)
