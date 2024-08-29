import xlwings as xl

app = xl.App(visible=False)

book = xl.Book(r'C:\ERWIN\tes.xlsx')
sheet = book.sheets[0]

dest1 = sheet[0,0].get_address().replace('$', '')
dest2 = sheet[0,1].get_address().replace('$', '')
sheet[0,2].formula2 = f"""=IFERROR({dest1}/{dest2}, 0)"""
print(dest1)
print(dest2)

book.save()
book.close()
app.quit()