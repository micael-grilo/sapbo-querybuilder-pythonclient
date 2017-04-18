import requests
from bs4 import BeautifulSoup
import xlwt

##################### Login DATA
server = "SERVER_IP"    #Server IP Address
user = "USERNAME"	#Server Administrator Username
passw = "PASSWORD"	#Server Administrator Password 
s = requests.session()
######################

def login():
	url="http://"+server+":8080/AdminTools/querybuilder/logon?framework="
	logindata= {'aps':server,
			'usr':user,
			'pwd':passw,
			'aut':'secEnterprise',
			'main_page':'ie.jsp'}
	s.post(url, data=logindata)


def reportdata(report_id, rowi,sheet):
	url = "http://"+server+":8080/AdminTools/querybuilder/query.jsp"
	data = {"sqlStmt":"SELECT SI_CUID, SI_ID, SI_NAME, SI_UPDATE_TS, SI_CREATION_TIME, SI_OWNER, SI_AUTHOR FROM CI_INFOOBJECTS WHERE SI_ID="+str(report_id),
			"SUBMIT":"Submit Query",
			"main_page":"query.jsp"}
	soup = BeautifulSoup(s.post(url, data=data).text)
	try: 
                table = soup.find("tr", { "class" : "header" }).parent
        except AttributeError:
                return
	r = 1
	for row in table.findAll("tr"):
		c=rowi
		cells = row.findAll("td")
		if c == 0:
			for cell in cells:
				sheet.write(c, r, cell.text)
				c+=1
		else:
			if len(cells)>1:
				sheet.write(c,r,cells[1].text)
			c+=1
		r+=1

def ListReportsByUniversedata(universe_id, sheet, i):
	url = "http://"+server+":8080/AdminTools/querybuilder/query.jsp"
	data = {"sqlStmt":"SELECT SI_CUID, SI_NAME, SI_OWNER, SI_CREATION_TIME, SI_UPDATE_TS, SI_WEBI FROM CI_APPOBJECTS WHERE SI_KIND='Universe' AND SI_ID = "+str(universe_id),
			"SUBMIT":"Submit Query",
			"main_page":"query.jsp"}
	soup = BeautifulSoup(s.post(url, data=data).text)
	table = soup.find("tr", { "class" : "header" }).parent
	uni_name = table.findAll("tr")[1].findAll("td")[1].text
	print uni_name[:30]
	data = table.find("table", {"class" : "basic"})
	for row in data.findAll("tr"):
		cells = row.findAll("td")
		if cells[0].text != "SI_TOTAL":
                        sheet.write((1 if i==0 else i),0, uni_name)
                        sheet.write((1 if i==0 else i),1, universe_id)
                        reportdata(cells[1].text,i,sheet)
                        i = [i+1,i+2][i==0]
	return i
        

def ListAllUniverses():
        Universes = []
	url = "http://"+server+":8080/AdminTools/querybuilder/query.jsp"
	data = {"sqlStmt":"SELECT si_id FROM CI_AppObjects WHERE si_kind = 'Universe'",
			"SUBMIT":"Submit Query",
			"main_page":"query.jsp"}
	soup = BeautifulSoup(s.post(url, data=data).text)
	table = soup.findAll("tr", {"class" : "header" })
	for each in table:
                Universes.append(int(each.parent.findAll("tr")[1].findAll("td")[1].text))
	return Universes

def ListReportsWithNoUniverse():
        Reports = []
        url = "http://"+server+":8080/AdminTools/querybuilder/query.jsp"
	data = {"sqlStmt":"SELECT SI_ID FROM CI_Infoobjects WHERE si_kind = 'WebI' AND SI_INSTANCE=0 and SI_UNIVERSE.SI_TOTAL=0",
			"SUBMIT":"Submit Query",
			"main_page":"query.jsp"}
	soup = BeautifulSoup(s.post(url, data=data).text)
	table = soup.findAll("tr", {"class" : "header" })
	for each in table:
                Reports.append(int(each.parent.findAll("tr")[1].findAll("td")[1].text))
	return Reports
        
			
def main():
        print "Doing some magic...."
	login()
	book = xlwt.Workbook(encoding="utf-8")
        Universes = ListAllUniverses()
        print "Found "+str(len(Universes))+" Universes!"
        sheet = book.add_sheet("Reports By Universe")
        i = 0
	for Universe in Universes:
		i = ListReportsByUniversedata(Universe,sheet,i)
	print "Listing Reports with no Universe!"
	Reports = ListReportsWithNoUniverse()
	i+=1
	for report in Reports:
                reportdata(int(report), i,sheet)
                i+=1
	book.save("Universe_data.xls")
	print "Magic Happens in File > Universe_data.xls!"

main()


