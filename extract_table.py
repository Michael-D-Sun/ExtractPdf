import pandas as pd
import pdfplumber


def parse_pdf(file_path):
    with pdfplumber.open(file_path) as pdf:
        tables = []
        for page in pdf.pages:
            table = page.extract_table()
            tables.append(table)
    return tables


def convert_to_excel(tables, output_file):
    with pd.ExcelWriter(output_file, mode='a', if_sheet_exists='overlay') as writer:
        for i, table in enumerate(tables):
            # print(table)
            df = pd.DataFrame(table[1:], columns=table[0])
            df.to_excel(writer, sheet_name=f"Data", index=False, startrow=writer.book["Data"].max_row)
            # writer.save()


def extract_table_from_pdf(file_path, output_file):
    print("extract_table_from_pdf")
    tables = parse_pdf(file_path)
    convert_to_excel(tables, output_file)


if __name__ == "__main__":
    extract_table_from_pdf("02 Schedules - CN for ASEAN_cn.pdf", "output.xlsx")


