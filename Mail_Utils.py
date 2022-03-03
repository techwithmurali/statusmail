import os,openpyxl
import datetime



def conv_Excel_Dict(ExcelName,SheetName,p_col_list):
    print('start of convExcelDict  - {} {}'.format(ExcelName,SheetName))
    wb = openpyxl.load_workbook(ExcelName,data_only=True)
    sheet = wb.get_sheet_by_name(SheetName)
    lst_row = []
    dict_sheet = {}
    l_rownum = 0
    for row in sheet.rows:
        l_rownum = l_rownum + 1
        l_col_cnt = 0
        for cell in row:
            l_col_cnt = l_col_cnt + 1
            if l_col_cnt in p_col_list:
                
                if cell.value is None:
                    cell.value = ' '
                if isinstance(cell.value, datetime.date):
                    cell.value = cell.value.strftime("%d %B, %Y")
                lst_row.append(cell.value)
        dict_sheet[l_rownum] = lst_row
        lst_row = []
    print('returning from convexceldict ')
    #print(dict_sheet)
    return dict_sheet

def build_header(p_dict_out,p_header):
    l_html_str = '<span style = "color:black">Hi,<br> {} <br> <br> \n'.format(p_header)
    p_dict_out['Header'] = l_html_str

def build_footer(p_dict_out,p_name,Footer_text):
    
    l_html_str = '<span style = "color:black"> <br> <b> {} </b> <br><br>\n'.format(Footer_text)
    l_html_str = l_html_str + '<span style = "color:black">Thanks and Regards<br> {} <br><br>\n'.format(p_name)
    p_dict_out['Footer'] = l_html_str

def write_dict_htlmfile(p_dict,htmlFile):
    fp = open(htmlFile,'w')
    for key,value in p_dict.items():
        fp.write(value + '\n')
    fp.close()
          
def Conv_Dict_HTMLDict(p_dict,Title,p_dict_out,l_dict_sheet_desc):
    print('inside Conv_Dict_HTMLDict sheet - {}'.format(Title))
    #print(p_dict)
    if len(p_dict.keys()) == 1:
        print('return here since leng is 1')
        return 
    #print('************ $$$$$$$$$$$$$$$$$$$$$$$$$$  ***************')
    l_html_str = '<span style = "color:black"><b>{}</b><br>\n'.format(l_dict_sheet_desc.get(Title,Title))
    l_html_str = l_html_str + '<br>'
    l_html_str = l_html_str + '<html><table border = 1>\n'
    #----------------------------------------------------------------------------------
    for key,values in p_dict.items():
        if len(values) == 0:
            #print(key)
            print('continue next iteration since len values is 0')
            continue
            
        if str(key)=='1':
            #print('first row')
            l_col_str = ''
            #l_cntr = 0
            for value in values:
                
                #if l_cntr != (len(values) - 1):
                l_col_str = l_col_str + '<td><b><span style="color:black" >{}</b></td>\n'.format(value)
                #l_cntr = l_cntr + 1
            #print(l_col_str)
            l_html_str = l_html_str + '<tr style="background-color:red">{}</tr>\n'.format(l_col_str)
        else:
            l_col_str = ''
            '''if values[-1] == 'Z':
                l_cntr = 0 
                for value in values:
                    
                    if l_cntr != (len(values) - 1):
                        l_col_str = l_col_str + '<td align = "left"><b><span style = "color:black">{}</b></td>\n'.format(value)
                    l_cntr = l_cntr + 1
            else:'''
            #l_cntr = 0 

            for value in values:
                
                #if l_cntr != (len(values) - 1):
                l_col_str = l_col_str + '<td align = "left"><span style = "color:black">{}</td>\n'.format(value)
                    #l_cntr = l_cntr + 1
        #print('else part ' + l_col_str)
            l_html_str = l_html_str +  '<tr style="background-color:white">{}</tr>\n'.format(l_col_str)
    #----------------------------------------------------------------------------------
    l_html_str = l_html_str + '</table></html><br>\n'
    p_dict_out[Title] = l_html_str


def conv_Excel_html(p_dict_details): 
    print('starting preparing html file')  
       
    ExcelName     = p_dict_details.get('ExcelName','')
    ExcelPath     = p_dict_details.get('ExcelPath','')
    Work_dir_path = p_dict_details.get('Work_dir_path','')
    MsgBody     = p_dict_details.get('MsgBody','')
    p_sheetList = p_dict_details.get('SheetList','')
    ToList      = p_dict_details.get('ToList','')
    sender      = p_dict_details.get('Sender','')
    p_cols_reqd = p_dict_details.get('ColsReqd','')
    Footer_text = p_dict_details.get('Footer_text','')
    l_dict_sheet_desc = p_dict_details.get('dictSheetDesc','')
    daily_status_html   = os.path.join(ExcelPath,'{}.html'.format(os.path.splitext(ExcelName)[0]))

    deflist = [1,2,3,4]
    l_html_dict = {}
    build_header(l_html_dict,MsgBody)
    wb = openpyxl.load_workbook(os.path.join(ExcelPath,ExcelName),data_only=True)
    for sheet in wb.sheetnames:
        if sheet in p_sheetList:
            l_Col_list = p_cols_reqd.get(sheet,deflist)
            dict_sheet = conv_Excel_Dict(os.path.join(ExcelPath,ExcelName),sheet,l_Col_list)
            Conv_Dict_HTMLDict(dict_sheet,sheet,l_html_dict,l_dict_sheet_desc)
    build_footer(l_html_dict,'Muralidharan R',Footer_text)
    write_dict_htlmfile(l_html_dict,daily_status_html)
    print('Completed preparing html file')
    

               
if __name__ == "__main__":
    pass
    #schedule_job()
