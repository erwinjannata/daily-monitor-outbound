import re
import xlwings as xl
from tkinter.messagebox import showinfo
from functions.counting import counting

def combining(file_data, file_report, date, saved_as, over_month, is_grouped):
    datas = counting(file_data=file_data, date=date, save_grouping=False, saved_as=saved_as, is_grouped=is_grouped)

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
                                     ].value = ""

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
                                     ].formula = f"={target_sheet[(cell_row + idx), 7 + (merged_col * i)].get_address().replace('$', '')}-{target_sheet[(cell_row + idx), 2 + (merged_col * i)].get_address().replace('$', '')}"

                        # % CANCEL
                        # target_sheet[(cell_row + idx), 9 + (merged_col * i)
                        #              ].value = (datas[item][i][idx][0] / datas[item][i][idx][5]) if datas[item][i][idx][5] != 0 else 0
                        target_sheet[(cell_row + idx), 9 + (merged_col * i)
                                     ].formula = f"=IFERROR({target_sheet[(cell_row + idx), 2 + (merged_col * i)].get_address().replace('$', '')}/{target_sheet[(cell_row + idx), 7 + (merged_col * i)].get_address().replace('$', '')}, 0)"

                        # % UN-RCC
                        # target_sheet[(cell_row + idx), 10 + (merged_col * i)
                        #              ].value = (datas[item][i][idx][1] / datas[item][i][idx][6]) if datas[item][i][idx][6] != 0 else 0
                        target_sheet[(cell_row + idx), 10 + (merged_col * i)
                                     ].formula = f"=IFERROR({target_sheet[(cell_row + idx), 3 + (merged_col * i)].get_address().replace('$', '')}/{target_sheet[(cell_row + idx), 8 + (merged_col * i)].get_address().replace('$', '')}, 0)"

                        # % UN-OM
                        # target_sheet[(cell_row + idx), 11 + (merged_col * i)
                        #              ].value = (datas[item][i][idx][2] / datas[item][i][idx][6]) if datas[item][i][idx][6] != 0 else 0
                        target_sheet[(cell_row + idx), 11 + (merged_col * i)
                                     ].formula = f"=IFERROR({target_sheet[(cell_row + idx), 4 + (merged_col * i)].get_address().replace('$', '')}/{target_sheet[(cell_row + idx), 8 + (merged_col * i)].get_address().replace('$', '')}, 0)"

                        # % UN-APPV OM
                        # target_sheet[(cell_row + idx), 12 + (merged_col * i)
                        #              ].value = (datas[item][i][idx][3] / datas[item][i][idx][6]) if datas[item][i][idx][6] != 0 else 0
                        target_sheet[(cell_row + idx), 12 + (merged_col * i)
                                     ].formula = f"=IFERROR({target_sheet[(cell_row + idx), 5 + (merged_col * i)].get_address().replace('$', '')}/{target_sheet[(cell_row + idx), 8 + (merged_col * i)].get_address().replace('$', '')}, 0)"

                        # % UN-SMU
                        # target_sheet[(cell_row + idx), 13 + (merged_col * i)
                        #              ].value = (datas[item][i][idx][4] / datas[item][i][idx][6]) if datas[item][i][idx][6] != 0 else 0
                        target_sheet[(cell_row + idx), 13 + (merged_col * i)
                                     ].formula = f"=IFERROR({target_sheet[(cell_row + idx), 6 + (merged_col * i)].get_address().replace('$', '')}/{target_sheet[(cell_row + idx), 8 + (merged_col * i)].get_address().replace('$', '')}, 0)"

                    # TOTAL PER RING
                    for idx in range(0, 7):
                        target_sheet[(cell_row + 2), (idx + 2) + (merged_col * i)
                                     ].value = datas[item][i][0][idx] + datas[item][i][1][idx]

                    # TOTAL CANCEL
                    target_sheet[(cell_row + 2), 2 + (merged_col * i)
                                 ].formula = f"=SUM({target_sheet[(cell_row + 0), 2 + (merged_col * i)].get_address().replace('$', '')}:{target_sheet[(cell_row + 1), 2 + (merged_col * i)].get_address().replace('$', '')})"
                    
                    # TOTAL FINAL (TOTAL CNOTE - CANCEL)
                    target_sheet[(cell_row + 2), 8 + (merged_col * i)
                                 ].formula = f"=SUM({target_sheet[(cell_row + 0), 8 + (merged_col * i)].get_address().replace('$', '')}:{target_sheet[(cell_row + 1), 8 + (merged_col * i)].get_address().replace('$', '')})"
                    
                    # % TOTAL PER RING
                    # % CANCEL
                    # target_sheet[(cell_row + 2), 9 + (merged_col * i)
                    #              ].value = (target_sheet[(cell_row + 2), 2 + (merged_col * i)].value / target_sheet[(cell_row + 2), 7 + (merged_col * i)].value) if target_sheet[(cell_row + 2), 7 + (merged_col * i)].value != 0 else 0
                    target_sheet[(cell_row + 2), 9 + (merged_col * i)
                                 ].formula = f"=IFERROR({target_sheet[(cell_row + 2), 2 + (merged_col * i)].get_address().replace('$', '')}/{target_sheet[(cell_row + 2), 7 + (merged_col * i)].get_address().replace('$', '')}, 0)"

                    # % UN-RCC
                    # target_sheet[(cell_row + 2), 10 + (merged_col * i)
                    #              ].value = (target_sheet[(cell_row + 2), 3 + (merged_col * i)].value / target_sheet[(cell_row + 2), 8 + (merged_col * i)].value) if target_sheet[(cell_row + 2), 8 + (merged_col * i)].value != 0 else 0
                    target_sheet[(cell_row + 2), 10 + (merged_col * i)
                                 ].formula = f"=IFERROR({target_sheet[(cell_row + 2), 3 + (merged_col * i)].get_address().replace('$', '')}/{target_sheet[(cell_row + 2), 8 + (merged_col * i)].get_address().replace('$', '')}, 0)"

                    # % UN-OM
                    # target_sheet[(cell_row + 2), 11 + (merged_col * i)
                    #              ].value = (target_sheet[(cell_row + 2), 4 + (merged_col * i)].value / target_sheet[(cell_row + 2), 8 + (merged_col * i)].value) if target_sheet[(cell_row + 2), 8 + (merged_col * i)].value != 0 else 0
                    target_sheet[(cell_row + 2), 11 + (merged_col * i)
                                 ].formula = f"=IFERROR({target_sheet[(cell_row + 2), 4 + (merged_col * i)].get_address().replace('$', '')}/{target_sheet[(cell_row + 2), 8 + (merged_col * i)].get_address().replace('$', '')}, 0)"

                    # % UN-APPV OM
                    # target_sheet[(cell_row + 2), 12 + (merged_col * i)
                    #              ].value = (target_sheet[(cell_row + 2), 5 + (merged_col * i)].value / target_sheet[(cell_row + 2), 8 + (merged_col * i)].value) if target_sheet[(cell_row + 2), 8 + (merged_col * i)].value != 0 else 0
                    target_sheet[(cell_row + 2), 12 + (merged_col * i)
                                 ].formula = f"=IFERROR({target_sheet[(cell_row + 2), 5 + (merged_col * i)].get_address().replace('$', '')}/{target_sheet[(cell_row + 2), 8 + (merged_col * i)].get_address().replace('$', '')}, 0)"

                    # % UN-SMU
                    # target_sheet[(cell_row + 2), 13 + (merged_col * i)
                    #              ].value = (target_sheet[(cell_row + 2), 6 + (merged_col * i)].value / target_sheet[(cell_row + 2), 8 + (merged_col * i)].value) if target_sheet[(cell_row + 2), 8 + (merged_col * i)].value != 0 else 0
                    target_sheet[(cell_row + 2), 13 + (merged_col * i)
                                 ].formula = f"=IFERROR({target_sheet[(cell_row + 2), 6 + (merged_col * i)].get_address().replace('$', '')}/{target_sheet[(cell_row + 2), 8 + (merged_col * i)].get_address().replace('$', '')}, 0)"

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
