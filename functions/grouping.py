import gc
import sys
import pandas as pd
from pathlib import Path
from tkinter.messagebox import showinfo
from datetime import datetime

def grouping(file_data, date, save_grouping, saved_as):
    # Read data source and refs
    try:
        working_dir = Path.cwd()
        data_awb = pd.read_excel(file_data)
        # data_cancel = pd.read_excel(file_data, sheet_name='AWB CANCEL')

        df_awb = pd.DataFrame(data_awb)
        # df_cancel = pd.DataFrame(data_cancel)
        df_user = pd.read_excel(
            rf'{working_dir}\REFS.xlsx', sheet_name='USERS')
        df_customer = pd.read_excel(
            rf'{working_dir}\REFS.xlsx', sheet_name='CUSTOMER')

        user = {}
        customer = {}

        for i in range(0, df_user.shape[0]):
            user[str(df_user['USER ID'][i])] = df_user["CABANG"][i]

        for i in range(0, df_customer.shape[0]):
            customer[str(df_customer['CUST_ID_2'][i])
                     ] = df_customer['BIG_GROUPING_CUST_2'][i]
    except Exception as e:
        showinfo(title="Error", message=e)
        sys.exit()

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
        try:
            # ---- CABANG
            cabang.append(user[str(df_awb['CNOTE USER ID'][index])])

            # ---- STATUS RECEIVING
            if not df_awb['RECEIVING NO'][index] == '-':
                status_receiving.append("RECEIVING")
            else:
                status_receiving.append("UNRECEIVING")

            # ---- TGL RECEIVING
            if not df_awb['RECEIVING DATE'][index] == "-":
                tgl_receiving.append(
                    df_awb['RECEIVING DATE'][index].strftime("%#m/%#d/%Y"))
            else:
                tgl_receiving.append("")

            # ---- STATUS MANIFEST
            if not df_awb['MANIFEST OUTB '][index] == '-':
                status_manifest.append("MANIFESTED")
            else:
                status_manifest.append("UNMANIFEST")

            # ---- TGL MANIFEST
            if not df_awb['MANIFEST DATE'][index] == "-":
                tgl_manifest.append(
                    df_awb['MANIFEST DATE'][index].strftime("%#m/%#d/%Y"))
            else:
                tgl_manifest.append("")

            # ---- ENTRY AWB
            if not df_awb['CNOTE DATE'][index] == "-":
                tgl_entry_awb.append(df_awb['CNOTE DATE']
                                     [index].strftime("%#m/%#d/%Y"))
            else:
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
            elif (df_awb['RUNSHEET NO'][index] != '-' and df_awb['POD STATUS'][index] == '-') and ((df_awb['RUNSHEET DATE'][index] != '-' and isinstance(df_awb['RUNSHEET DATE'][index], str) == False) and datetime.strptime((str(df_awb['RUNSHEET DATE'][index])), '%Y-%d-%m %H:%M:%S').strftime("%#m/%#d/%Y") == today):
                status_pod.append("ONDELIVERY")
            elif (df_awb['RUNSHEET NO'][index] != '-' and df_awb['POD STATUS'][index] == '-') and ((df_awb['RUNSHEET DATE'][index] != '-' and isinstance(df_awb['RUNSHEET DATE'][index], str) == True) and datetime.strptime((str(df_awb['RUNSHEET DATE'][index])), '%d/%m/%Y %H:%M:%S').strftime("%#m/%#d/%Y") == today):
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
            elif df_awb['POD STATUS'][index] == 'OS':
                status_pod.append("UNDEL")
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
            if not df_awb['RUNSHEET DATE'][index] == "-" and isinstance(df_awb['RUNSHEET DATE'][index], str) == False:
                date_runsheet.append(df_awb['RUNSHEET DATE']
                                     [index].strftime("%#d/%#m/%Y"))
            elif not df_awb['RUNSHEET DATE'][index] == "-" and isinstance(df_awb['RUNSHEET DATE'][index], str) == True:
                date_runsheet.append(datetime.strptime(str(df_awb['RUNSHEET DATE']
                                     [index]), '%d/%m/%Y %H:%M:%S').strftime("%#m/%#d/%Y"))
            else:
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
        except Exception as e:
            showinfo(title="Error", message=e)
            sys.exit()

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

    # # Create new dataframe of filtered AWB Cancel data
    # cabang_cancel = []
    # ring_area_cancel = []
    # customer_grouping_cancel = []

    # for index in range(0, df_cancel.shape[0]):
    #     try:
    #         # ----
    #         cabang_cancel.append(user[str(df_cancel['Cnote User'][index])])

    #         # ----
    #         try:
    #             if user[df_cancel['Cnote User'][index]] == 'MATARAM':
    #                 ring_area_cancel.append("RING 1")
    #             else:
    #                 ring_area_cancel.append("RING 2")
    #         except:
    #             ring_area.append("")

    #         # ----
    #         try:
    #             customer_grouping_cancel.append(
    #                 customer[str(df_cancel['Agent id'][index])])
    #         except:
    #             customer_grouping_cancel.append("")
    #     except Exception as e:
    #         showinfo(title="Error", message=e)
    #         sys.exit()

    # # Add new column to dataframe
    # df_cancel.loc[:, "CABANG"] = cabang_cancel
    # df_cancel.loc[:, "RING"] = ring_area_cancel
    # df_cancel.loc[:, "CUSTOMER"] = customer_grouping_cancel

    # Write dataframe into existing excel file
    if save_grouping == True:
        with pd.ExcelWriter(saved_as, mode='w', engine='xlsxwriter', ) as writer:
            df_awb.to_excel(writer, sheet_name='AWB', index=False)
            # df_cancel.to_excel(writer, sheet_name='AWB CANCEL', index=False)

        del df_awb
        gc.collect()
        showinfo(title="Message",
                 message=f"Proses selesai")
    # Return the grouping result
    else:
        return df_awb
    # --------