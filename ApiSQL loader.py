# 
# -*- coding: utf-8 -*-
"""
@author: Christian Rivera

This program does the following
1. loads an excel file from a gui prompt
2. reads the rows of the excel file into an api call to avention
3. returns the avention api call as a list of dictionaries
4. compares the returned list to the initial excel file.
5. creates a new excel file showing which companies Avention found and which it did not
6. load the Avention API data into MRP's sql database in the db, "ListData", 
        and table, "dbo.AVENTION_CUSTOM_MATCHED
"""


#Creates the GUI for the user to pick his excel file
from Tkinter import Tk
from tkFileDialog import askopenfilename
Tk().withdraw()
filename = askopenfilename()
##############################################





### load excel file into a list of dictionaries named Data
import xlrd
workbook = xlrd.open_workbook(filename)
workbook = xlrd.open_workbook(filename, on_demand = True)
worksheet = workbook.sheet_by_index(0)
first_row = [] # The row where we stock the name of the column
for col in range(worksheet.ncols):
    first_row.append( worksheet.cell_value(0,col) )
# transform the workbook to a list of dictionaries
data =[]
for row in range(1, worksheet.nrows):
    elm = {}
    for col in range(worksheet.ncols):
        elm[first_row[col]]=worksheet.cell_value(row,col)
    data.append(elm)
################################################3


#convert List of Dictionaries from Unicode to utf-8C
holder = []
for mydict in data:
     holder.append( {unicode(k).encode("utf-8"): unicode(v).encode("utf-8") for k,v in mydict.iteritems()} )
print "unicode?"
print holder[0]
data = holder
############################################3

#### only use company, country, customid, superfluous input columns break the program

for i in data:
    try:
        #i['PostalCode'] = str(int(float(i['PostalCode'])))
        i['CustomID'] = str(int(float(i['CustomID'])))
    except:
        pass

#############################################
print "and this far"


## split 'Data' into a list of a list of dictionaries called Chunks where each
# chunk is 1000 rows of the Data list of dictionaries.  The api call can only handle 1000 at a time
chunks = [data[y:y+1000] for y in xrange(0,len(data),1000)]

#'finalList' is where all the returned rows are placed as a list of dictionaries
finalList = []

###############################################

print "chunks"

#calling the data from  api
import requests
import json
for i in chunks:
    url = #api call string goes here
    api = requests.get(url, data = json.dumps(i))
    js = json.loads(api.content)
    
    dict_list = js[0]["Information"]
    print "Number of companies entered is %s" % len(i)
    print "Number of companies api returned is %s" % len(dict_list)
    
    for x in dict_list:
        finalList.append(x)

print len(data)
print len(finalList)
#####################################################



#convert List of Dictionaries from Unicode to utf-8C

holder = []
for mydict in finalList:
     holder.append( {unicode(k).encode("utf-8"): unicode(v).encode("utf-8") for k,v in mydict.iteritems()} )
print "unicode?"
finalList = holder
#####################################################

#eliminate duplicates from finallist
seen = set()
result = []
mrpID = set()
for d in finalList:
    h = d['keyid']
    j = d['customid']
    mrpID.add(j)
    if h not in seen:
        result.append(d)
        seen.add(h)
        
        
finalList = result
print len(finalList)
print "final list"
print finalList[1]
####################################33

####################################33


#mark which companies were found valid in api and which werent

filtered = []
for x in data:
    h = x['CustomID']
    if h in mrpID:
        x['match'] = 'True'
    else:
        x["match"] = 'False'
        filtered.append(x)
#######################################################         



#create unicode lists to write excel file
holder = []
for i in data:
    unidict = {k.decode('utf-8'): v.decode('utf-8') for k, v in i.items()}
    holder.append(unidict)
#print holder[0]
data = holder
print "got this far baby"

holder = []
for i in finalList:
    unidict = {k.decode('utf-8'): v.decode('utf-8') for k, v in i.items()}
    holder.append(unidict)
#print holder[0]
finalList = holder

holder = []
for i in filtered:
    unidict = {k.decode('utf-8'): v.decode('utf-8') for k, v in i.items()}
    holder.append(unidict)
#print holder[0]
filtered = holder


####################################################3


integerList = ['customid','keyid','matchconfidence','postalcode','employees',
               'primarynaic','primaryussic','primaryuk2007sic','primaryanzsic','primarynaics2012',
               'primaryisicrev4sic','primarynacerev2sic','yearfounded','monthfounded','dayfounded',
               'creditnumericscore','ultimateparentkeyid','parentkeyid','CustomID','PostalCode']

floatList = ['salesusd','salesgbp','saleseur','sales','assetsusd','assetsgbp','assetseur','assets',
             'totalliabilities','totalliabilitiesusd','totalliabilitiesgbp','totalliabilitieseur',
             'longtermdebt']



import xlsxwriter
wb = xlsxwriter.Workbook("Accounts Report Demo output.xlsx")
number_format = wb.add_format({'num_format':'0'})
ws=wb.add_worksheet("raw data") 

first_row=0
ordered_list = ['']
ordered_list = data[0].keys()
for header in ordered_list:
    col=ordered_list.index(header) # we are keeping order.
    ws.write(first_row,col,header) # we have written first row which is the header of worksheet also.

row=1
for player in data:
    for _key,_value in player.items():
        col=ordered_list.index(_key)
        if _key in integerList:
            try:
                ws.write_number(row,col,float(unicode(_value).encode('utf-8')),number_format)
            except:
                pass
        elif _key in floatList:
            try:
                ws.write_number(row,col,float(unicode(_value).encode('utf-8')),number_format)
            except:
                pass
        else:
            ws.write(row,col,_value)
    row+=1 #enter the next row
    
 
wx=wb.add_worksheet("matched data")

ordered_list = [u'customid',u'matchconfidence',u'keyid',u'companyname',u'address1',u'address2',u'address3',u'city',u'stateorprovinceabbrev',
                u'postalcode',u'county',u'countryname', u'countryiso2',u'phone',u'fax',u'primaryurl',u'employees',u'industrygroupname',u'industrysectorname',
                u'os2010industryname',u'primarynaic',u'primarynaicdesc', u'primaryussic',u'primaryussicdesc', u'primaryuk2007sic',u'primaryuk2007sicdesc',
                u'primaryanzsic',u'primaryanzsicdesc',u'primarynaics2012',u'primarynaics2012desc',u'primaryisicrev4sic',u'primaryisicrev4sicdesc',u'primarynacerev2sic',u'primarynacerev2sicdesc',u'currencyiso3',
                u'currencyname',u'salesusd',u'salesgbp',u'saleseur',u'sales',u'assetsusd',u'assetsgbp',u'assetseur',u'assets',u'ownershiptype',u'entitytype',
                u'businessdescription',u'parentkeyid',u'parentname',u'ultimateparentkeyid',u'ultimateparentname',u'tickerexchange',u'tickersymbol',u'abinumber',u'regno',
                u'creditrating',u'creditnumericscore',u'creditlimit',u'creditflag',u'sales1yeargrowth',u'totalassets1yrgrowth',u'netincome1yrgrowth',
                u'operatingmargin',u'workingcapital',u'currentassets',u'fixedassets',u'currentliabilities',u'totalliabilities',u'totalliabilitiesusd',u'totalliabilitiesgbp',
                u'totalliabilitieseur',u'longtermdebt',u'yearfounded',u'monthfounded',u'dayfounded']

print "HERE!"





first_row=0
#ordered_list = finalList[0].keys()
for header in ordered_list:
    col=ordered_list.index(header) # we are keeping order.
    wx.write(first_row,col,header) # we have written first row which is the header of worksheet also.

row=1
for player in finalList:
    for _key,_value in player.items():
        col=ordered_list.index(_key)
        if _key in integerList:
            try:
                wx.write_number(row,col,float(unicode(_value).encode('utf-8')),number_format)
            except:
                pass
        elif _key in floatList:
            try:
                wx.write_number(row,col,float(unicode(_value).encode('utf-8')),number_format)
            except:
                pass
        else:
            wx.write(row,col,_value)
    row+=1 #enter the next row   
    
wz=wb.add_worksheet("no match data")
first_row=0
ordered_list = filtered[0].keys()
for header in ordered_list:
    col=ordered_list.index(header) # we are keeping order.
    wz.write(first_row,col,header) # we have written first row which is the header of worksheet also.

row=1
for player in filtered:
    for _key,_value in player.items():
        col=ordered_list.index(_key)
        if _key in integerList:
            try:
                wz.write_number(row,col,float(unicode(_value).encode('utf-8')),number_format)
            except:
                pass
        elif _key in floatList:
            try:
                wz.write_number(row,col,float(unicode(_value).encode('utf-8')),number_format)
            except:
                pass
        else:
            wz.write(row,col,_value)
    row+=1 #enter the next row   
    
wb.close()

# send results to Excel sheet showing which companies were verified and which werent.

#####################################################################

 # set up sql connection.  
import pypyodbc
connection = pypyodbc.connect()#sql string goes here
cursor = connection.cursor();
########################################

# pull all keyID's from sql table
exist = cursor.execute("select KeyID from API_CUSTOM_MATCHED")
existing = set()
for x in exist:
    existing.add(x['KeyID'])

#make sure new additions aren't duplicates of anything currently on the table
result = []
for d in finalList:
    h = d['keyid']
    if h not in existing:
        result.append(d)

finalList = result
##################################

#run a for loop to insert each dictionary as a row
x = 0
y = 0
for i in finalList:
    try:
        sqlCommand = ("""insert into API_CUSTOM_MATCHED
    (MRP_ROW_ID,
    CompanyName,
    Address1,
    Address2,
    Address3,
    City,
    County,
    StateOrProvinceAbbrev,
    PostalCode,
    CountryName,
    Phone,
    PrimaryURL,
    MatchConfidence,
    KeyID,
    countryiso2,
    employees,
    entitytype,
    industrygroupname,
    industrysectorname,
    ownershiptype,
    parentkeyid,
    parentname,
    ultimateparentkeyid,
    ultimateparentname,
    primarynaic,
    primarynaicdesc,
    primaryussic,
    primaryussicdesc,
    salesusd,
    assetsusd,
    yearfounded)
    values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""")
        values = [ i['customid'],i['companyname'],i['address1'],i['address2'],i['address3'],i['city'],i['county'],i['stateorprovinceabbrev'],i['postalcode'],i['countryname'],i['phone'],i['primaryurl'], i['matchconfidence'],i['keyid'],
              i['countryiso2'], i['employees'],i['entitytype'],i['industrygroupname'],
                i['industrysectorname'],(i['ownershiptype']),(i['parentkeyid']), i['parentname'], (i['ultimateparentkeyid']),i['ultimateparentname'],i['primarynaic'],i['primarynaicdesc'],
                i['primaryussic'],i['primaryussicdesc'],(i['salesusd']), (i['assetsusd']),i['yearfounded']]
        cursor.execute(sqlCommand, values)
        print x
        x = x + 1
    except:

        y = y+1
        #print "rejected %s" % y

#commit all the new rows    
cursor.commit()

#doubly make sure all things are commited
connection.commit

# close the connection.
connection.close()
