def makeComiFcn(IT_number,motor_name):
    
    import pandas as pd
    from numpy import isnan
    import sys

    geometry_excel = pd.ExcelFile("//techdata/eng machine data/Innovation/Simulation/Magnetic FEA/" + motor_name +\
                                  "/Geometry Dimensions/" + motor_name + " Geometry Definition.xlsx")

    comi_file = open('//techdata/eng machine data/Innovation/Simulation/Magnetic FEA/' + motor_name +\
         '/Geometry Dimensions/dimensions_' + motor_name + '_IT' + str(IT_number) + '.comi','w')


    nb_sheets = len(geometry_excel.sheet_names)
    IT_check = True

    for i in range(nb_sheets):
        sheet = geometry_excel.parse(i,None)
        if pd.MultiIndex.tolist(sheet.iloc[:,0]).count(IT_number) == 0:
            IT_check = False
            sys.exit('Error : IT Number is not in the list')

    # Check the longest variable name, for a nice presentation of the comi file
    max_length_name = 0

    for i in range(nb_sheets):
        sheet = geometry_excel.parse(i,None)
        nb_variable = len(sheet.iloc[1,:])
        for j in range(nb_variable):
            if len(str(sheet.iloc[1,j]))>max_length_name:
                max_length_name = len(str(sheet.iloc[1,j]))

    # Write the comi file
    for k in range(nb_sheets):
        comi_file.write('/---------------------------------\n')
        comi_file.write('/ ' + geometry_excel.sheet_names[k] + '\n')
        comi_file.write('/---------------------------------\n')

        sheet = geometry_excel.parse(k,None)
        n = len(sheet.iloc[0,:])
        write = 0
        for i in range(n):
            if sheet.iloc[0,i]==sheet.iloc[0,i]:
                write = 1
                if sheet.iloc[0,i] != 'Notes':
                    comi_file.write('\n/ ' + str(sheet.iloc[0,i]) + '\n')
                else:
                    comi_file.write('\n/ Notes:\n')
                        
            if write == 1:
                if sheet.iloc[IT_number+1,i] != sheet.iloc[IT_number+1,i]:
                    if str(sheet.iloc[1,i]) == 'Notes':
                        comi_file.write('/ NaN\n')
						
                elif str(sheet.iloc[1,i]) == 'Notes':
                    comi_file.write('/ ' + sheet.iloc[IT_number+1,i])
                                     
                elif str(sheet.iloc[1,i][0]) == '#':
                    comi_file.write('$CONS    ' + str(sheet.iloc[1,i]))
                    for l in range(max_length_name-len(str(sheet.iloc[1,i]))+3):
                        comi_file.write(' ')
                    comi_file.write(str(sheet.iloc[IT_number+1,i]) + '\n')
					
                else:              
                    comi_file.write('$STRING  ' + str(sheet.iloc[1,i]))
                    for l in range(max_length_name-len(str(sheet.iloc[1,i]))+3):
                        comi_file.write(' ')
                    comi_file.write("'" + str(sheet.iloc[IT_number+1,i]) + "'\n")
                                                          
        comi_file.write('\n\n')
        
    comi_file.close()
