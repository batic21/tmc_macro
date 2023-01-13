import pandas as pd
excel_path ="QTtools.xlsm"

#V3 A to V2 A
#V3 B to V2 B
#V3 C to V2 C
#V3 D to V2 D
#V3 E to V2 E
#V3 G to V2 F
#V3 I to V2 G
#V3 J to V2 H
#V3 K to V2 I
#V3 L to V2 J

v3_A = 0
v3_B = 1
v3_C = 2
v3_D = 3
v3_E = 4
v3_G = 6
v3_I = 8
v3_J = 9
v3_K = 10
v3_L = 11

v2_A = 0
v2_B = 1
v2_C = 2
v2_D = 3
v2_E = 4
v2_F = 5
v2_G = 6
v2_H = 7
v2_I = 8
v2_J = 9




def ty_read_excel_file(fpath):
      return   pd.read_excel(fpath, sheet_name=['V3','V2'])

if __name__ == '__main__':
    print("Filename:" + excel_path)
    df =  ty_read_excel_file(excel_path)     
    fV3 = df.get('V3')
    fV2 = df.get('V2')
    v3_count = fV3[fV3.columns[0]].count()
    v2_count = fV2[fV2.columns[0]].count()
    comp = 0

    #print(fV3)

    for r in range(v3_count):
       print(fV3.loc[r][v3_A] + " - " + str(r))
       
       for e in range(v2_count):
          #V3 A to V2 A
          if fV3.loc[r][v3_A] == fV2.loc[e][v2_A]:
            comp = comp + 1
            #V3 B to V2 B
            if fV3.loc[r][v3_B] == fV2.loc[e][v2_B]:
                comp = comp + 1
                #V3 C to V2 C
                if fV3.loc[r][v3_C] == fV2.loc[e][v3_C]:
                    comp = comp + 1             
                    #V3 D to V2 D
                    if fV3.loc[r][v3_D] == fV2.loc[e][v2_D]:
                        comp = comp + 1
                        #V3 E to V2 E
                        if fV3.loc[r][v3_E] == fV2.loc[e][v2_E]:
                            comp = comp + 1
                            #V3 G to V2 F
                            if fV3.loc[r][v3_G] == fV2.loc[e][v2_F]:
                                comp = comp + 1
                                #V3 I to V2 G
                                if fV3.loc[r][v3_I] == fV2.loc[e][v2_G]:
                                    comp = comp + 1                       
                                    #V3 J to V2 H
                                    if fV3.loc[r][v3_J] == fV2.loc[e][v2_H]:
                                        comp = comp + 1                     
                                        #V3 K to V2 I
                                        if fV3.loc[r][v3_K] == fV2.loc[e][v2_I]:
                                            comp = comp + 1                            
                                            #V3 L to V2 J
                                            if fV3.loc[r][v3_L] == fV2.loc[e][v2_J]:
                                                print("matched: v3:" + str (r) + " V2:" + str (e))
                                                break              
          if comp > 2:
              print("Reached " + str(comp))  
          comp = 0
    #print(fV3.iloc[1]['File name'])

    print("V3 count:" + str(v3_count))
    print("V2 Count:" + str(v2_count))

    #compare

    #V3 A to V2 A