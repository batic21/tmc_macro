import os
import pandas as pd
result_macroname =[]
result_filename = []
result_line_number = []
result_csvfile = []
import xlsxwriter

M_TRUE = 1
M_FALSE = 0
dfoundidx = 0
dup_found = []

def ty_read_csv_file(fpath):
      return   pd.read_csv(fpath)
  
def check_duplicate(val_word):
    idx = 0
    try:
      idx = result_macroname.index(val_word)
      retval = M_TRUE
    except:
      retval = M_FALSE
 
    return retval, idx     
    
def process_file(fname):
    df =  ty_read_csv_file(fname)     
    m_row_count = len(df)
    progress = 0
    progress_div = float(100/m_row_count)
    prg_crt=0
        
    temp_duplicate_macro = 0

    
    for r in range(m_row_count):
       if prg_crt == 500:
            print(fname + ": " + str(int(progress)) + "%", end="\r")
            prg_crt = 0
       else:
            prg_crt = prg_crt + 1
       progress = progress + progress_div
      # print(str(r + 2)+": " + df.loc[r][0] + " : " + str(df.loc[r][1])  + " : " + str(df.loc[r][2]))
       
       dfound, dindex = check_duplicate(df.loc[r][0])
       
       if dfound == M_TRUE:   
           #print("======duplicate found  index = " + str(dindex) + "found index" + str(dfoundidx))    
           #print(result_filename) 
           result_csvfile[dindex].append([fname])
           result_filename[dindex].append(os.path.basename(str(df.loc[r][1])))
           result_line_number[dindex].append(df.loc[r][2])                       
           
           #result report
           try:
               dup_found.index(dindex) 
           except:
               dup_found.append(dindex)  
               
           temp_duplicate_macro = temp_duplicate_macro + 1                               
       else:
           result_macroname.append(str(df.loc[r][0]))
           result_csvfile.append([fname])
           tidx = result_macroname.index(str(df.loc[r][0]))
           result_filename.append([os.path.basename(str(df.loc[r][1]))])
           result_line_number.append([df.loc[r][2]])
        
       
       
    #time.sleep(0.01)   
    print(fname + ": 100%  Total Macro: "+ str(m_row_count) + "  Duplicate Macro: "+ str(temp_duplicate_macro), end="\r")
    print("")


def macro_export(mname, mfile, mline):
    workbook = xlsxwriter.Workbook('result.xlsx')
    worksheet = workbook.add_worksheet()
    text_format = workbook.add_format({'text_wrap': True})
    worksheet.set_column('A:A', 50)
    worksheet.set_column('B:B', 50)
    worksheet.set_column('C:C', 50)
    worksheet.set_column('D:D', 30)
    
   # print(mline)
    worksheet.write('A1', "MACRO NAME")
    worksheet.write('B1', "CSV Filename")
    worksheet.write('C1', "Code Filenames")
    worksheet.write('D1', "Line of Code")
    worksheet.write('E1', "Duplicate Macro Count")
    
    ex_len = len(mname)
    
    for e in range(ex_len):
        worksheet.write('A' + str(e+2), str(mname[e]), text_format) 
        
        tstr1 = str(result_csvfile[e])
        tstr = tstr1.replace('[','')   
        tstr2 = tstr.replace(']','')
        #worksheet.write('B' + str(e+2),tstr2, text_format) # str(result_csvfile[e]))
        worksheet.write('B' + str(e+2), str(result_csvfile[e]), text_format)
        worksheet.write('C' + str(e+2), str(mfile[e]), text_format)
        worksheet.write('D' + str(e+2), str(mline[e]),text_format)
        worksheet.write('E' + str(e+2), str(len(mline[e])),text_format)
        
    workbook.close()

if __name__ == '__main__':
    
    for x in os.listdir():
      if x.endswith(".csv"):
        # Prints only text file present in My Folder
        process_file(str(x)) 
    
    macro_export(result_macroname,result_filename,result_line_number)
    
    
    #process_file("Adc_Internal.c.macros.csv") 
    #process_file("Adc_Data.c.macros.csv")    
    #print(dup_found)
    
   # for d in range(len(dup_found)):
   #     rdx = dup_found[d]
        #print(result_macroname[rdx] + " | " + str(result_filename[rdx])) # + " | "  + result_line_number[d] )
        #print(result_macroname[rdx] + " | " + "count:" + str(len(result_filename[rdx])-1)  + "  " +  str(result_filename[rdx])) # + " | "  + result_line_number[d] )
   #    print(result_macroname[rdx] + " | " + "Duplicate:" + str(len(result_filename[rdx])-1) )# + "  " +  str(result_filename[rdx])) # + " | "  + result_line_number[d] )
    
    print("============================================")    
    print("\nDuplicate count:" + str(len(dup_found)))    
    print("Total macro: " + str(len(result_macroname)))
    print("============================================")    
    print("see output result on file: result.xlsx")
    
       
