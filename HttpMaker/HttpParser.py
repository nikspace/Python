# coding=utf8
import os
import urllib
import urllib2
import cookielib
import time
from StringIO import StringIO
import gzip
from HtmlControler import HtmlControler
from nt import lstat
import ExcelControler
import ConfigManager


'''
Http with cookie :should declare in the header part,just like below content
'''
# 获取一个保存cookie的对象
cj = cookielib.LWPCookieJar();
# 将一个保存cookie对象，和一个HTTP的cookie的处理器绑定
cookie_support = urllib2.HTTPCookieProcessor(cj);
# 创建一个opener，将保存了cookie的http处理器，还有设置一个handler用于处理http的URL的打开
opener = urllib2.build_opener(cookie_support, urllib2.HTTPHandler);
# 将包含了cookie、http处理器、http的handler的资源和urllib2对象板顶在一起
urllib2.install_opener(opener);

# 设置办公网代理，非公司网路可以注释掉此行
#proxy = urllib2.ProxyHandler({'http': 'web-proxy.oa.com:8080'});
#opener = urllib2.build_opener(proxy,cookie_support);

opener = urllib2.build_opener(cookie_support);
urllib2.install_opener(opener); 


def get_timestap_ms():
    second = time.time();
    #print str(float(second)*1000);
    return str(second*1000);

# global var
GET_URL_Login = 'http://221.181.71.161:8081/fluxWebTms/Login'
POST_URL_Login = 'http://221.181.71.161:8081/fluxWebTms/Login';
#no needs for login
GET_URL_ServiceList = 'http://221.181.71.161:8081/fluxWebTms/ServiceList?id=2001&name=servicelistpath&opr=1003&type=1000&etc='+get_timestap_ms()+'&qs=1s';
GET_URL_Welocome = 'http://221.181.71.161:8081/fluxWebTms/welcome2System.jsp?qs=1s';
POST_URL_GetTaskActiveFlag = 'http://221.181.71.161:8081/fluxWebTms/setSysAction.getTaskActiveFlag.action';
POST_URL_QueryBBS= 'http://221.181.71.161:8081/fluxWebTms/setSysAction.queryBBS.action';
POST_URL_SetSysAcion= 'http://221.181.71.161:8081/fluxWebTms/setSysAction.getUserData.action';

# Query Action
POST_URL_QueryGrid = 'http://221.181.71.161:8081/fluxWebTms/basCommonAction.jspTagRecentQueryGrid.action?queryFunctionId=NT_ANJI_TSS01_Q&gridFunctionId=NT_ANJI_TSS01_L&pageFlag=&orderByStr=%20order%20by%20BALANCER_NAME%20asc%20,TRANS_CORP_NAME%20asc%20,LOAD_NUMBER%20asc%20&whereFields=UDF55&noWhereFields=&otherWhere=&objectId=&where_flag=N'
# Query Click
POST_URL_QueryClick_Prefix = 'http://221.181.71.161:8081/fluxWebTms/basCommonAction.jspTagQueryGridResultSet.action?queryFunctionId=&gridFunctionId=NT_ANJI_TSS01_D&pageFlag=&otherWhere=%20and%20VIN=%27'
POST_URL_QueryClick_Suppix = '%27&orderByStr=%20order%20by%20SEND_TIME%20desc%20,SN,MAIN_ID,VIN'
URL_EXIT = 'http://221.181.71.161:8081/fluxWebTms/Exit?id=2001&name=exitservice&opr=1002&type=1001&opr2=1002' ;
'''
        result cols info:
        0: 序号
        1: 验证结果
        2: 校验未通过原因
        3: 标准距离
        4: 扫描距离
        6: 扫描操作时间
        9: VIN码
        10: 操作名称
        24: TSS省
        25: TSS市
        44: 运输公司 
        '''
RESULT_DICT = {0: '序号',1: '验证结果',2: '校验未通过原因',3: '标准距离',4: '扫描距离',6: '扫描操作时间',9: 'VIN码',10: '操作名称',24: 'TSS省',25: 'TSS市',44: '运输公司 '}
RESULT_DICT_CSV = {9:'VIN',10:'操作名称',24:'当前位置省份',25:'当前位置城市 ',6:'实际到达时间',1:'校验结果',2:'校验未通过原因'}

OPERATE_TYPE_DICT={'01':'装车','03':'在途','05':'已交车'}


RESULT_LIST = []#查询有效结果list


RECORD_LIST=[] #需要写入CSV的记录LIST with line str list

CURRENT_VIN = '' #当前操作的VIN


def decode(encodeStr):
    buf = StringIO( encodeStr)
    f = gzip.GzipFile(fileobj=buf)
    data = f.read()
    return data;

def login():     
    headers = {
        'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Encoding':'gzip, deflate, sdch',
        'Accept-Language':'zh-CN,zh;q=0.8',
        'Cache-Control':'max-age=0',
        'Cookie':'JSESSIONID=B429E49C0C427D90292E8125848D719C; webfxtab_NT_ANJI_REP10_P_tabPane=1',
        'Host':'221.181.71.161:8081',
        'Proxy-Connection':'keep-alive',
        'Upgrade-Insecure-Requests':'1',
        'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/44.0.2403.125 Safari/537.36'
    }
    req = urllib2.Request(GET_URL_Login,headers=headers);
    #print req;
    result = opener.open(req)
    #print decode(result.read())
    #print result.info();
    #for index, cookie in enumerate(cj):
    #    print '[',index, ']',cookie;
    conf = ConfigManager.ConfigManager(r'c:\http_maker_conf.ini')
    user = conf.get('http_maker', 'user')
    pwd = conf.get('http_maker','password')
    if user == '' or pwd == '':
        print 'conf file is not found!'
        exit(1)
    postdata = {}
    postdata['opr'] = 1001
    postdata['userId'] = user
    postdata['pw'] = pwd
    postdata['theme'] = 'Workplac1H'
    postdata['language'] = 'cn'
    postdata['htmlColor'] = 'HtmlColor1'
    postdata['htmlDbType'] = 'SEC'
    postdata = urllib.urlencode(postdata)
    #print postdata
    #headers = {'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/44.0.2403.125 Safari/537.36'}
    headers = {
        'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Encoding':'gzip, deflate, sdch',
        'Accept-Language':'zh-CN,zh;q=0.8',
        'Cache-Control':'max-age=0',
        'Content-Type':'application/x-www-form-urlencoded',
        'Host':'221.181.71.161:8081',
        'Origin':'http://221.181.71.161:8081',
        'Referer':'http://221.181.71.161:8081/fluxWebTms/Login',        
        'Proxy-Connection':'keep-alive',
        'Upgrade-Insecure-Requests':'1',
        'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/44.0.2403.125 Safari/537.36'
    }
       
    req = urllib2.Request (
        url=POST_URL_Login,
        data=postdata,
        headers=headers
    )
    result = urllib2.urlopen(req)
    return decode(result.read())
    
    
    

def query_grid(VIN):
    data = {
        'commonObject.UDF01':'',
        'commonObject.UDF03':'OUTBOUND_TIME',
        'commonObject.UDF04':'2016-06-07 00:00:00',
        'commonObject.UDF06':'2016-06-07 23:59:59',
        'commonObject.UDF07':'F,N',
        'commonObject.UDF08':'',
        'commonObject.UDF35':'',
        'commonObject.UDF36':'',
        'commonObject.UDF10':'',
        'commonObject.UDF10NAME':'',
        'commonObject.UDF45':'',
        'commonObject.UDF46':'',
        'commonObject.UDF135':'',
        'commonObject.UDF136':'',
        'commonObject.UDF110':'',
        'commonObject.UDF110NAME':'',
        'commonObject.UDF50':'',
        'commonObject.UDF51':'',
        'commonObject.UDF56':'',
        'commonObject.UDF55': VIN,
        'commonObject.UDF57':'',
        'commonObject.UDF200':'',
        '_':''
                
    }
    
    postdata = urllib.urlencode(data)
    #print postdata
    headers = {
        'Accept':'text/javascript, text/html, application/xml, text/xml, */*',
        'Accept-Encoding':'gzip, deflate',
        'Accept-Language':'zh-CN,zh;q=0.8',
        'Content-type':'application/x-www-form-urlencoded; charset=UTF-8',
        'Host':'221.181.71.161:8081',
        'Origin':'http://221.181.71.161:8081',
        'Proxy-Connection':'keep-alive',
        'Referer':'http://221.181.71.161:8081/fluxWebTms/baseInitJspAction.toCommonJspPage.action?mc=NT_ANJI_TSS01&FMTN=Y&qs=1s',
        'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/44.0.2403.125 Safari/537.36',
        'X-Prototype-Version':'1.6.0',
        'X-Requested-With':'XMLHttpRequest'
    }
    req = urllib2.Request (
        url=POST_URL_QueryGrid,
        data=postdata,
        headers=headers
    )
    result = urllib2.urlopen(req)
    text = result.read()
    return decode(text)

def query_click(VIN):
    data = {
        'queryFunctionId':'',
        'gridFunctionId':'NT_ANJI_TSS01_D',
        'pageFlag':'',
        'otherWhere':' and VIN="LSVXZ65N5G2082315%27"',
        'orderByStr':' order by SEND_TIME desc ,SN,MAIN_ID,VIN'
    }
    postdata = urllib.urlencode(data)
    #print postdata
    headers = {
        'Accept':'text/javascript, text/html, application/xml, text/xml, */*',
        'Accept-Encoding':'gzip, deflate',
        'Accept-Language':'zh-CN,zh;q=0.8',
        'Content-type':'application/x-www-form-urlencoded; charset=UTF-8',
        'Host':'221.181.71.161:8081',
        'Origin':'http://221.181.71.161:8081',
        'Proxy-Connection':'keep-alive',
        'Referer':'http://221.181.71.161:8081/fluxWebTms/baseInitJspAction.toCommonJspPage.action?mc=NT_ANJI_TSS01&FMTN=Y&qs=1s',
        'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/44.0.2403.125 Safari/537.36',
        'X-Prototype-Version':'1.6.0',
        'X-Requested-With':'XMLHttpRequest'
    }
    req = urllib2.Request (
        url=POST_URL_QueryClick_Prefix + VIN + POST_URL_QueryClick_Suppix,
        data=postdata,
        headers=headers
    )
    result = urllib2.urlopen(req)
    text = result.read()
    return decode(text)

def validate_login(data):
    '''
    <title>FLUX.OMS(Web2.0)</title>
    '''
    hc = HtmlControler()
    text = hc.get_title_text(data)
    print text
    if text == 'FLUX.OMS(Web2.0)' :
        return True
    return False    

def validate_grid(data):
    if data.endswith("'pageCount':'0','totalCount':'0'}") == True :
        print 'validate_grid: response is empty'
        return False
    if data.find(';[##########]') == -1:
        print 'validate_grid: validate_grid failed'
        return False
    return True
def validate_click(data):
    if data.find(';[##########]') == -1:
        print 'validate_click: validate_click failed'
        return False
    return True

def parse_grid(data):
    '''
    @return: True if parse ok
    @return: False if pares nok
    data is empty: [];[##########]{'pageIndex':'1','pageSize':'20','pageCount':'0','totalCount':'0'}
    data is normal:[["1","LSVXJ45L9G2021947","斯柯达","Y16008834191","YA51183","SD16060409C4","不分段","安亭库","毕节","安亭库","毕节","贵州省","毕节","重庆博宇","安吉运输","2016-06-05 13:24:03.0","","","","2016-06-05 15:00:00","2016-06-05 13:27:50.0","2016-06-05 13:27:50.0","","公路","","","","","3","2","?","","","","","","80884884"]];
    [##########]{'pageIndex':'1','pageSize':'20','pageCount':'1','totalCount':'1'}
    '''
    if data.endswith("'pageCount':'0','totalCount':'0'}") == True :
        print 'responsed is empty'
        return False
    else:
        index = data.find(';[##########]')
        data_str = data[0:index - 1]
        
        data_list = list(data_str)
        if len(data_list) == 1:
            print 'data_list[0]: '+ data_list[0]
            print 'data_list: ' + data_list
            print '运输公司: ' + data_list[13]
            
            return True
        else:
            return False
            
    return False

def write_csv_file(path,data_list):
    '''
    @return: 
    '''
    try:
        # 打开一个文件
        # 重命名文件test1.txt到test2.txt。
        os.rename( path, path+'bak' )
        os.remove(path)
        fo = open(path, "wb")
        
        # write header
        line = ''
        for key in RESULT_DICT.keys():
            line += RESULT_DICT[key] + ','
        line = line[0:len(line) - 1] + '\n'
        fo.write( line)     
        #write content      
        for lst in data_list:
            line = ''
            for key in RESULT_DICT.keys():
                line += str(lst[key]) + ','
            line = line[0:len(line) - 1] + '\n'
            fo.write( line)
        # 关闭打开的文件
        fo.close()
    except:
        print '[write_csv_file] Error occur'
        raise
def get_valid_result_lst(result_list):
    return 
#yyyy-mm-dd
def get_today_date_str():
    return str(time.strftime('%Y%m%d',time.localtime(time.time())))   
#yyyymmddHHMMSS
def get_today_datetime_str():
    return str(time.strftime('%Y%m%d%H%M%S',time.localtime(time.time())))    

def write_result_global(data_list):
    '''
    @summary: write result to file with append mode
    RESULT_DICT_CSV = {9:'VIN',10:'操作名称',24:'当前位置省份',25:'当前位置城市 ',6:'实际到达时间',1:'校验结果',2:'校验未通过原因'}
    OPERATE_TYPE_DICT={'01':'装车','03':'在途','05':'已交车'}
    '''
    try:
        global CURRENT_VIN 
        global RESULT_LIST
        global RECORD_LIST
        #tmp lst
        lst = []
        #transport company
        trans_list = []
        finish_list_ok = [] #未交车
        finish_list_nok = [] #未交车 校验未通过
        ontheway_list = [] #在途
        #if there is 05 code  
        for l in data_list:
            
            #重庆博宇        
            if str(l[44]) == '重庆博宇':
                trans_list.append(l)
                #已交车
                if str(l[10])[0:2] == '05':
                    if str(l[1]) == '校验通过':
                        l[25] = '已交车'
                        #l[9] = l[6] #当前市 = 操作时间
                        finish_list_ok.append(l)
                    else:
                        l[25] = '已交车 ' #校验未通过原因
                        finish_list_nok.append(l)
                #当日在途    
                if str(l[10])[0:2] == '03':
                     if str(l[6])[0:10] == get_today_date_str():
                         ontheway_list.append(l)
                         
        if len(trans_list) > 0: 
            if len(finish_list_ok) >= 1:
                lst = finish_list_ok[-1]
                print '[write_csv_result] '+ CURRENT_VIN + ': 已交车'
            elif len(finish_list_nok) >= 1:
                lst = finish_list_nok[-1]
                print '[write_csv_result] '+ CURRENT_VIN + ': 已交车,校验未通过'
            elif len(ontheway_list) >= 1:
                lst = ontheway_list[0]
                print '[write_csv_result] '+ CURRENT_VIN + ': 在途 省['+lst[24]+']市['+lst[25]+']'
            
        
        if len(lst) == 0:
            print '[write_csv_result] '+ CURRENT_VIN + ': 无符合要求结果数据'
            #校验结果    校验未通过原因    实际到达时间    VIN    操作名称    当前位置省份    当前位置城市 
            line =  '校验通过,'','',' + CURRENT_VIN + ',03在途定位,'',''\n'
        else:
            #add to global 
            RESULT_LIST.append(lst)
            # compose to line
            line = ''
            for key in RESULT_DICT_CSV.keys():
                line += str(lst[key]) + ','
            line = line[0:len(line) - 1] + '\n'
        RECORD_LIST.append(line)    
        
        print '[write_result_global] --------------------OVERVIEW Start--------------------------------'
        print 'All Record Count        :' + str(len(RECORD_LIST))
        print 'Availiable Result Count :' + str(len(RESULT_LIST))
        print '[write_result_global] --------------------OVERVIEW End----------------------------------'
    except:
        print '[write_csv_result]: Error occur!'
        raise    
    
def write_csv_result(path):
    '''
    @summary: write result to file with append mode
    RESULT_DICT_CSV = {9:'VIN',10:'操作名称',24:'当前位置省份',25:'当前位置城市 ',6:'实际到达时间',1:'校验结果',2:'校验未通过原因'}
    OPERATE_TYPE_DICT={'01':'装车','03':'在途','05':'已交车'}
    '''
    try:
        
        # 打开一个文件
        # 重命名文件test1.txt到test2.txt
        fo = open(path, "a")
        # write header
        line = ''       
        for key in RESULT_DICT_CSV.keys():
            line += RESULT_DICT_CSV[key] + ','
        line = line[0:len(line) - 1] + '\n'
        fo.write( line)
         
        #write line content
        for l in RECORD_LIST:
            fo.write(str(l))
        # 关闭打开的文件
        fo.close() 
    except:
        print '[write_csv_result]: Error occur!'
        raise    
def parse_click(result):
    '''
    @return: True if parse ok
    @return: False if pares nok
    data is empty: [];[##########]{'pageIndex':'1','pageSize':'20','pageCount':'0','totalCount':'0'}
    data is normal:[["1","校验通过","","0","           0.0","A06574","2016-06-07 09:51:24.0","263001","重庆博宇","LSVXJ45L9G2021947","03在途定位（无板车身份卡）","","","","","","","武汉云申","2016-06-07 09:51:30.0","A6759333","","","","","江苏省","苏州市","昆山市","","江苏省苏州市昆山市","","2016-06-07 09:51:24.0","31303400","12113508","6759333","0","","Y16008834191","安亭库","01","毕节","55170","贵州省","毕节","SD16060409C4","重庆博宇","安吉运输","YA51183","安亭库","01","毕节","55170","贵州省","毕节","公路","公路","6759333LSVXJ45L9G2021947","A46875341"],
    ["2","校验通过","","0","           0.0","A06574","2016-06-06 20:48:55.0","263001","重庆博宇","LSVXJ45L9G2021947","03在途定位（无板车身份卡）","","","","","","","武汉云申","2016-06-06 20:49:08.0","A6757359","","","","","上海市","上海市","嘉定区","","上海市上海市嘉定区","","2016-06-06 20:48:55.0","31296890","12115874","6757359","0","","Y16008834191","安亭库","01","毕节","55170","贵州省","毕节","SD16060409C4","重庆博宇","安吉运输","YA51183","安亭库","01","毕节","55170","贵州省","毕节","公路","公路","6757359LSVXJ45L9G2021947","A46856001"],
    ["3","校验通过","","50","           0.4","A06809","2016-06-05 13:24:03.0","263001","重庆博宇","LSVXJ45L9G2021947","01装车（无板车身份卡）","安亭库","50000163","","","","","景德镇恒通","2016-06-05 13:24:11.0","A6752343","","","","","上海市","上海市","嘉定区","嘉定区","上海市上海市嘉定区","上海市上海市嘉定区","2016-06-05 13:24:03.0","31294479","12117052","6752343","1","","Y16008834191","安亭库","01","毕节","55170","贵州省","毕节","SD16060409C4","重庆博宇","安吉运输","YA51183","安亭库","01","毕节","55170","贵州省","毕节","公路","公路","6752343LSVXJ45L9G2021947","A46797742"]];
    [##########]{'pageIndex':'1','pageSize':'20','pageCount':'1','totalCount':'3'}
   
    '''
    
    if result.endswith("'pageCount':'0','totalCount':'0'}") == True :
        # currently never enter here
        print '[parse_click]: response is empty'
        return False
    else:
        index = result.find(';[##########]')
        result_str = result[0:index]
        #print result_str
        result_list = list(eval(result_str))
        result_list_len = len(result_list)
        
        print '[parse_click] result_list_len = ' + str(result_list_len)
        '''
        result cols info:
        0: 序号
        1: 验证结果
        2: 校验未通过原因
        3: 标准距离
        4: 扫描距离
        6: 扫描操作时间  str[0:18]
        9: VIN码
        10: 操作名称
        24: TSS省
        25: TSS市
        44: 运输公司 
        '''
        write_result_global(result_list)
        
        
    return True
    

FROM_FILE_PATH = 'D:\\' + get_today_date_str()+'.xlsx'
TO_FILE_PATH = 'D:\\' + get_today_date_str()+'_update.xlsx'

def auto_work():
    #  excel read to query_data
    #query_datas = {'LSVXJ45L9G2021947','LSVXZ65N5G2082315','LSJZ14C32GS048033'}
    #query_datas = {'LSJA24U64GS039658'}
    #query_datas = {'LSJA24U60GS033954'}
    global CURRENT_VIN
    global FROM_FILE_PATH
    global TO_FILE_PATH
    query_datas = ExcelControler.read_vin_list(FROM_FILE_PATH)
    data = login()
    if validate_login(data) == False:
        print 'login failed'
        return False
    else:
        print 'login success'
    
    for VIN in query_datas:
        CURRENT_VIN = VIN
        result = query_grid(VIN)
        if validate_grid(result) == False :
            print 'query_grid ' + VIN +' failed'
            continue
        else:
            print 'query_graid ' + VIN + ' success'
        
        #no need to parse grid result 
        #------------
        # click
        result = query_click(VIN)
        if validate_click(result) == False :
            print 'query_click ' + VIN +' failed'
            continue
        else:
            print 'query_click ' + VIN + ' success'
        
        #parse click result
        if parse_click(result) == False:
            print '[auto_work] failed'
        else:
            print '[auto_work] success'
        
    #write original result data
    #write_csv_file(r'D:\test.csv',result_list)
    
    ExcelControler.excel_update(FROM_FILE_PATH,TO_FILE_PATH,RESULT_LIST)
    write_csv_result(r'D:\result_'+ get_today_datetime_str()+ '.csv')        

if __name__ == "__main__":
    '''
    @testcase
    '''
    auto_work()
    
