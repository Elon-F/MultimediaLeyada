import os
import pandas as pd

PATH = r"\\nas\it\computerlist_new"
DEST = r"results.xlsx"

if __name__ == '__main__':
    l = []
    for class_name in os.listdir(PATH):
        item = os.path.join(PATH, class_name)
        if os.path.isdir(item):
            for pc_name in os.listdir(item):
                if "Final" in pc_name:  # don't touch the "Final" files
                    continue
                pc = os.path.join(item, pc_name)
                f = open(pc, encoding="utf-16")
                contents = f.read().split("\n")
                # print(pc)
                l.append({"ClassName": class_name, "PcName": os.path.splitext(pc_name)[0], "Serial": contents[1].rstrip(), "Vendor": " ".join(contents[3].rstrip().split())})

    df = pd.DataFrame(l, columns={"ClassName", "PcName", "Serial", "Vendor"})
    dfs = {"computers": df}
    # df.to_excel(os.path.join(PATH, DEST), index=False)
    writer = pd.ExcelWriter(os.path.join(PATH, DEST), engine='xlsxwriter')
    for sheet_name, df in dfs.items():  # loop through `dict` of dataframes
        df.to_excel(writer, sheet_name=sheet_name, index=False)  # send df to writer
        worksheet = writer.sheets[sheet_name]  # pull worksheet object
        for idx, col in enumerate(df):  # loop through all columns
            series = df[col]
            max_len = max((
                series.astype(str).map(len).max(),  # len of largest item
                len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
            worksheet.set_column(idx, idx, max_len)  # set column width
    writer.save()
