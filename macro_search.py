import os
import pandas as pd
import xlsxwriter
result_macroname =[]
result_filename = []
result_line_number = []
result_csvfile = []


M_TRUE = 1
M_FALSE = 0
dfoundidx = 0
dup_found = []
report_summmary = []


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
    prg_crt=0        
    temp_duplicate_macro = 0
    result_macroname.clear()
    result_filename.clear()
    result_line_number.clear()
    result_csvfile.clear()
    dup_found.clear()
    #dfoundidx = 0
    
    if m_row_count > 0:
        progress_div = float(100/m_row_count)

         
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
               #result_csvfile[dindex].append([fname])
               result_filename[dindex].append(os.path.basename(str(df.loc[r][1])))
               result_line_number[dindex].append(df.loc[r][2])                       
           
               #result report
               try:
                   dup_found.index(dindex) #ignore
                   #print("ignore")
               except:
                   dup_found.append(dindex)  #add duplicate
                   temp_duplicate_macro = temp_duplicate_macro + 1
                                                         
           else:
               result_macroname.append(str(df.loc[r][0]))
               #result_csvfile.append([fname])
               #tidx = result_macroname.index(str(df.loc[r][0]))
               result_filename.append([os.path.basename(str(df.loc[r][1]))])
               result_line_number.append([df.loc[r][2]])
    else:
            result_macroname.append('N/A')
            result_filename.append('N/A')
            result_line_number.append(None)
       
       
    #time.sleep(0.01)   
    print(fname + ": 100%  Total Macro: "+ str(m_row_count) + "  Duplicate Macro: "+ str(temp_duplicate_macro), end="\r")
    print("")


def macro_export(mname, mfile, mline, excel_fname):
    workbook = xlsxwriter.Workbook('result_' + excel_fname + '.xlsx')
    worksheet = workbook.add_worksheet()
    text_format = workbook.add_format({'text_wrap': True})
    summary_macrofound = M_FALSE
    worksheet.set_column('A:A', 50)
    worksheet.set_column('B:B', 50)
    worksheet.set_column('C:C', 50)
    worksheet.set_column('D:D', 30)
    worksheet.set_column('E:E', 30)
    
   # print(mline)
    worksheet.write('A1', "MACRO NAME")
    #worksheet.write('B1', "CSV Filename")
    worksheet.write('B1', "Code Filenames")
    worksheet.write('C1', "Line of Code")
    worksheet.write('D1', "Duplicate Macro Count (found:" + str(len(dup_found)) +")")
    
    ex_len = len(mname)
    #print(mline[0])
    
    for e in range(ex_len):
        worksheet.write('A' + str(e+2), str(mname[e]), text_format) 
        worksheet.write('B' + str(e+2), str(mfile[e]), text_format)
        worksheet.write('C' + str(e+2), str(mline[e]), text_format)
        try:
           worksheet.write('D' + str(e+2), str(len(mline[e])),text_format)
           summary_macrofound = M_TRUE
        except:
           worksheet.write('D' + str(e+2), str('None'),text_format)
           summary_macrofound = M_FALSE   

        try:
            if len(mline[e]) > 1:
                 worksheet.write('E' + str(e+2), "DUPLICATE FOUND",text_format)
            else:
                 worksheet.write('E' + str(e+2), "NO DUPLICATE",text_format)
        except:
            worksheet.write('E' + str(e+2), "NO DUPLICATE",text_format)


    worksheet.write('A' + str(e+4), "TOTAL NUMBER of MACRO", text_format)
    worksheet.write('B' + str(e+4),str(ex_len), text_format)   
    worksheet.write('A' + str(e+5), "TOTAL NUMBER of DUPLICATE MACRO", text_format)
    worksheet.write('B' + str(e+5),str(len(dup_found)), text_format)

    if summary_macrofound == M_TRUE:
        report_summmary.append([excel_fname,str(ex_len),str(len(dup_found))])
    else:
        report_summmary.append([excel_fname,str(0),str(0)])

    workbook.close()

def macro_export_summary(): 
    workbook = xlsxwriter.Workbook('macro_export_summary.xlsx')
    worksheet = workbook.add_worksheet()
    text_format = workbook.add_format({'text_wrap': True})
    worksheet.set_column('A:A', 50)
    worksheet.set_column('B:B', 50)
    worksheet.set_column('C:C', 50)

    worksheet.write('A1', "CSV NAME")
    worksheet.write('B1', "MACRO COUNT")
    worksheet.write('C1', "DUPLICATE Count")

    ex_len = len(report_summmary)  

    for e in range(ex_len):
        worksheet.write('A' + str(e+2), str(report_summmary[e][0]), text_format)
        worksheet.write('B' + str(e+2), str(report_summmary[e][1]), text_format)
        worksheet.write('C' + str(e+2), str(report_summmary[e][2]), text_format)
    workbook.close()    


if __name__ == '__main__':
    
    for x in os.listdir():
      if x.endswith(".csv"):
        # Prints only text file present in My Folder
        process_file(str(x))
        macro_export(result_macroname,result_filename,result_line_number,str(x)) 
    macro_export_summary()
    
    
    
    
    #process_file("Adc_Internal.c.macros.csv") 
    #process_file("Adc_Data.c.macros.csv")    
    #print(dup_found)
    
   # for d in range(len(dup_found)):
   #     rdx = dup_found[d]
        #print(result_macroname[rdx] + " | " + str(result_filename[rdx])) # + " | "  + result_line_number[d] )
        #print(result_macroname[rdx] + " | " + "count:" + str(len(result_filename[rdx])-1)  + "  " +  str(result_filename[rdx])) # + " | "  + result_line_number[d] )
   #    print(result_macroname[rdx] + " | " + "Duplicate:" + str(len(result_filename[rdx])-1) )# + "  " +  str(result_filename[rdx])) # + " | "  + result_line_number[d] )
    
   # print("============================================")    
   # print("\nDuplicate count:" + str(len(dup_found)))    
   # print("Total macro: " + str(len(result_macroname)))
   # print("============================================")    
    print("see output result on file: result.xlsx")
    input("\n\nPress enter to continue")
    
       
