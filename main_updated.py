#########################Code for extracting college and stream details
# from PyPDF2 import PdfReader
# import re
# import pandas as pd

 
# pdf_path = r"C:\Users\ketan\Downloads\MH Cap 1 Cut Off.pdf"
# pdf_reader = PdfReader(pdf_path)

# pattern = r'(\d{4})\s-\s([^\n,]+)'
# pattern2 = r'(\d{9})\s-\s([^\n,]+)'
# pattern3 = r'Status:\s+(.*)'
# pattern4 = r'(\d{9}F)\s-\s([^\n,]+)'

# code=0
# count=0
# count1=0
# list=[]

# for i in range(len(pdf_reader.pages)):
#     page_number = i
#     if 0 <= page_number < len(pdf_reader.pages):
#         page = pdf_reader.pages[page_number]
#         string = page.extract_text()
#         matches = re.search(pattern, string)
        
#         if matches:
#             college_code = matches.group(1)
#             college_name=matches.group(2)
#             if(college_code!=code):
#                 count=count+1
#                 code=college_code  
#         matching = re.search(pattern3, string)
        
#         if matching:

#                 status = matching.group(1)
#         matches1 = re.findall(pattern4, string)           
#         matches2 = re.findall(pattern2, string)
#         if matches2:
#             for match2 in matches2:
#                 stream_code = match2[0]
#                 stream_name = match2[1]   
#                 list.append({"serial_no":count, "college_code": college_code,"college_name":college_name,"stream_name": stream_name,"status":status})
#         if matches1:    
#             for match2 in matches1:
#                 stream_code = match2[0]
#                 stream_name = match2[1]    
#                 list.append({"serial_no":count, "college_code": college_code,"college_name":college_name,"stream_name": stream_name,"status":status})
#                 # print(list)
#                 count1=count1+1
#                 print(count1,"pagenumber",page_number,college_code)w
                
                
#     else:
#         print(f"Page {page_number} is out of range.")
        

# df = pd.DataFrame(list)
# excel_file_path = r"C:\Users\ketan\Downloads\extract.xlsx"
# df.to_excel(excel_file_path, index=False)
# print(f"Data saved to {excel_file_path}")  


###################################################


#########################Code for extracting table data
# import tabula
# import PyPDF2
# import pandas as pd
# import re

# pattern_stream_list = r'(\d{9})\s-\s([^\n,]+)'
# pattern = r'\((.*?)\)'  # percentile
# pattern4 = r'(\d{9}F)\s-\s([^\n,]+)'

# pdf_directory = r"C:\Users\ketan\Downloads\MH Cap 1 Cut Off.pdf"
# pdf_reader = PyPDF2.PdfReader(pdf_directory)
# text = ''
# list_pages = []
# result = []  # Move the result list outside of the page loop

# def get_table_data(final_data, tables):
#     data = {}
#     stream = list(final_data.keys())
#     stream = stream[0]
#     value = list(final_data.values())
#     list_value = value[0]
#     table_list = []

#     for i, table in enumerate(tables):
#         i = str(i + 1)
#         df = pd.DataFrame(table)

#         if i in list_value:
#             for _, row in df.iterrows():
#                 for column, value in row.items():
#                     value = re.search(pattern, str(value))
#                     if value:
#                         value = value.group(1)
#                         table_list.append({column: value})

#     data[stream] = table_list
#     return data

# def stream_table_number_extraction(text, start_stream, next_stream, table_num):
#     start_pattern = re.escape(start_stream)
#     end_pattern = re.escape(next_stream)
    
#     start_match = re.search(start_pattern, text)
#     end_match = re.search(end_pattern, text)
#     index_number = []

#     if start_match and end_match:
#         start_index = start_match.start()
#         end_index = end_match.end()
#         extracted_data = text[start_index:end_index]
#         table_list = re.findall("Stage", extracted_data)

#         for j in table_list:
#             table_num = table_num + 1
#             index_number.append(str(table_num))

#     final_data = {start_stream: index_number}
#     return final_data, table_num

# for i in range(len(pdf_reader.pages)):
#     # i=16
#     # print("Page No ", i + 1)
#     page = pdf_reader.pages[i]
#     text = page.extract_text()
#     table_number = 0
#     matches2 = re.findall(pattern_stream_list, text)
#     matches3 = re.findall(pattern4, text)
#     if matches2:
#         stream_name_list = [j[1] for j in matches2]
#         print("true", i + 1)
#     elif matches3:
#         stream_name_list = [j[1] for j in matches3]
#         print(i + 1, stream_name_list)
#     else:
#         stream_name_list = False

#     if stream_name_list:
#         tables = tabula.read_pdf(pdf_directory, pages=str(i + 1))
#         print('stage count between the stream pending---->')

#         for j, start_stream in enumerate(stream_name_list):
#             try:
#                 next_stream = stream_name_list[j + 1]
#             except:
#                 next_stream = "Legends"
#             final_data, table_number = stream_table_number_extraction(text, start_stream, next_stream, table_number)
#             dict1 = get_table_data(final_data, tables)
#             result.append(dict1)
#     else:
#         # Extract the table and add it to the last inserted dictionary
#         tables = tabula.read_pdf(pdf_directory, pages=str(i + 1))
#         if result:
#             last_dict = result[-1]  # Get the last inserted dictionary
#             last_stream = list(last_dict.keys())[0]  # Get the stream name from the last dictionary
#             table_data = get_table_data({last_stream: [str(table_number + 1)]}, tables)
#             last_dict[last_stream].extend(table_data[last_stream])
#             print("trueeeeeeeeee")

# # ... (rest of the code) ...
# flattened_data = []


# for entry in result:
#     for stream, values in entry.items():
#         row = {"stream": stream}
#         for item in values:
#             row.update(item)
#         flattened_data.append(row)

# df = pd.DataFrame(flattened_data)


# excel_file_path = r"C:\Users\ketan\Downloads\output.xlsx"
# df.to_excel(excel_file_path, index=False)

# print("no stream in: ",list_pages)

##################################################################



# import pandas as pd


# input_file1_path = r"C:\Users\ketan\Downloads\collegedata.xlsx"  # Replace with the actual file path
# input_file2_path = r"C:\Users\ketan\Downloads\output.xlsx" # Replace with the actual file path

# df1 = pd.read_excel(input_file1_path)
# df2 = pd.read_excel(input_file2_path)

# if "stream" in df2.columns:
#     df2.drop(columns=["stream"], inplace=True)
# merged_df = pd.concat([df1, df2], axis=1)
# merged_output_path = r"C:\Users\ketan\Downloads\merged_output.xlsx"  
# merged_df.to_excel(merged_output_path, index=False)
# print("Columns from the second Excel file added to the first Excel file.")

###############################################3







import time
import re
import pandas as pd
import tabula
from PyPDF2 import PdfReader


pattern = r'(\d{4})\s-\s([^\n,]+)'
pattern2 = r'(\d{9})\s-\s([^\n,]+)'
pattern3 = r'Status:\s+(.*)'
pattern4 = r'(\d{9}F)\s-\s([^\n,]+)'
pattern5 = r'\((.*?)\)'           # percentile


def extract_college_stream_details(pdf_reader):
    code=0
    count=0
    count1=0
    list=[]
    
    for i in range(len(pdf_reader.pages)):
        page_number = i
        print(page_number)
        if 0 <= page_number < len(pdf_reader.pages):
            page = pdf_reader.pages[page_number]
            string = page.extract_text()
            matches = re.search(pattern, string)
            
            if matches:
                college_code = matches.group(1)
                college_name=matches.group(2)
                if(college_code!=code):
                    count=count+1
                    code=college_code  
            matching = re.search(pattern3, string)
            
            if matching:
                    status = matching.group(1)
            matches1 = re.findall(pattern4, string)           
            matches2 = re.findall(pattern2, string)
            if matches2:
                for match2 in matches2:
                    stream_code = match2[0]
                    stream_name = match2[1]   
                    list.append({"serial_no":count, "college_code": college_code,"college_name":college_name,"stream_name": stream_name,"status":status})
            if matches1:    
                for match2 in matches1:
                    stream_code = match2[0]
                    stream_name = match2[1]    
                    list.append({"serial_no":count, "college_code": college_code,"college_name":college_name,"stream_name": stream_name,"status":status})
                    # print(list)
                    count1=count1+1
                    print(count1,"pagenumber",page_number,college_code)          
        else:
            print(f"Page {page_number} is out of range.")
            
    df1 = pd.DataFrame(list)
    return df1

def extract_table_data(pdf_directory,pdf_reader):
    pattern_stream_list = r'(\d{9})\s-\s([^\n,]+)'
    text = ''
    list_pages = []
    result = []  

    def get_table_data(final_data, tables):
        data = {}
        stream = list(final_data.keys())
        stream = stream[0]
        value = list(final_data.values())
        list_value = value[0]
        table_list = []

        for i, table in enumerate(tables):
            i = str(i + 1)
            df = pd.DataFrame(table)

            if i in list_value:
                for _, row in df.iterrows():
                    for column, value in row.items():
                        value = re.search(pattern5, str(value))
                        if value:
                            value = value.group(1)
                            table_list.append({column: value})

        data[stream] = table_list
        return data

    def stream_table_number_extraction(text, start_stream, next_stream, table_num):
        start_pattern = re.escape(start_stream)
        end_pattern = re.escape(next_stream)
        
        start_match = re.search(start_pattern, text)
        end_match = re.search(end_pattern, text)
        index_number = []

        if start_match and end_match:
            start_index = start_match.start()
            end_index = end_match.end()
            extracted_data = text[start_index:end_index]
            table_list = re.findall("Stage", extracted_data)

            for j in table_list:
                table_num = table_num + 1
                index_number.append(str(table_num))

        final_data = {start_stream: index_number}
        return final_data, table_num

    for i in range(len(pdf_reader.pages)):
        # i=16
        # print("Page No ", i + 1)
        page = pdf_reader.pages[i]
        text = page.extract_text()
        table_number = 0
        matches2 = re.findall(pattern_stream_list, text)
        matches3 = re.findall(pattern4, text)
        if matches2:
            stream_name_list = [j[1] for j in matches2]
            print("true", i + 1)
        elif matches3:
            stream_name_list = [j[1] for j in matches3]
            print(i + 1, stream_name_list)
        else:
            stream_name_list = False

        if stream_name_list:
            tables = tabula.read_pdf(pdf_directory, pages=str(i + 1))
            print('stage count between the stream pending---->')

            for j, start_stream in enumerate(stream_name_list):
                try:
                    next_stream = stream_name_list[j + 1]
                except:
                    next_stream = "Legends"
                final_data, table_number = stream_table_number_extraction(text, start_stream, next_stream, table_number)
                dict1 = get_table_data(final_data, tables)
                result.append(dict1)
        else:
            # Extract the table and add it to the last inserted dictionary
            tables = tabula.read_pdf(pdf_directory, pages=str(i + 1))
            if result:
                last_dict = result[-1]  # Get the last inserted dictionary
                last_stream = list(last_dict.keys())[0]  # Get the stream name from the last dictionary
                table_data = get_table_data({last_stream: [str(table_number + 1)]}, tables)
                last_dict[last_stream].extend(table_data[last_stream])
                print("trueeeeeeeeee")

    flattened_data = []
    for entry in result:
        for stream, values in entry.items():
            row = {"stream": stream}
            for item in values:
                row.update(item)
            flattened_data.append(row)

    print("no stream in: ",list_pages)
    df2 = pd.DataFrame(flattened_data)
    return df2

def merge_excel_files(df1,df2,filename):
    if(len(df1)==len(df2)):
        if "stream" in df2.columns:
            df2.drop(columns=["stream"], inplace=True)
        merged_df = pd.concat([df1, df2], axis=1)
        merged_output_path = f"C:\\Users\\Acc User\\Downloads\\{filename}.xlsx"
        merged_df.to_excel(merged_output_path, index=False)
        print("Columns from the second Excel file added to the first Excel file.")
    return merged_df

time.sleep(1)
pdf_path = r"C:\Users\Acc User\Downloads\2021ENGG_CAP1_CutOff.pdf"
splitted_list=pdf_path.split("\\")
splitted_filename=splitted_list[-1].split(".")
filename=splitted_filename[-2]
print(filename)

pdf_reader = PdfReader(pdf_path)
df1=extract_college_stream_details(pdf_reader)
print(len(df1))

time.sleep(1)
df2=extract_table_data(pdf_path,pdf_reader)
print(len(df2))

time.sleep(1)
merged_df=merge_excel_files(df1, df2,filename)
print(len(merged_df)," done")



















