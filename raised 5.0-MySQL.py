import xlsxwriter, time, os.path, shutil, smtplib, signature
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import timedelta, date, datetime
import MySQLdb

DB_connected = False
try:
    db = MySQLdb.connect(host="___",    # your host, usually localhost
                     user="___",         # your username
                     passwd="",  # your password
                     db="___")        # name of the data base
    cur = db.cursor()
    DB_connected = True
except (Exception) as error:
    print (error)
    
LogFile=open("RaisedLog.txt","a")

listacod=["Country","Batch","NrAdj","Adjustment","Contract","Country","AmountEur","EUR","AmountCurr","Curr","Reason","Date"]#lista cu tipurile de date de pe fiecare linie din fisierul de interpretat


def daterange(date1, date2):
    for n in range(int ((date2 - date1).days)+1):
        yield date1 + timedelta(n)


def trimitere_email(zi_din_an,luna_data_nerulata,an_data_nerulata):
    fromaddr = "___"
    toaddr = "___"
    toccaddr = "___"
    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['CC'] = toccaddr
    msg['Subject'] = "Ajustari trimise Customer Service - luna "+str(luna_data_nerulata)+"/"+str(an_data_nerulata)
    body = "Fisierul cu ajustarile trimise a sosit si este disponibil in atasament. <BR><BR><BR> Acest mail a fost generat automat, va rugam nu dati reply."
    msg.attach(MIMEText(body, 'html'))
    filename = "Raised_"+str(zi_din_an)+"_Luna"+str(luna_data_nerulata)+".xlsx"
    attachment = open(os.path.join(str(os.getcwd()),filename), "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msg.attach(part)
    server = smtplib.SMTP_SSL('___', 465)
    server.login("___", "___")
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr.split(",") + toccaddr.split(","), text)
    server.quit()


def procedura_main(zi_din_an,luna_data_nerulata,an_data_nerulata):
    Dictionar={}
    VerificareContracte=[]
    workbook=xlsxwriter.Workbook("Raised_"+str(zi_din_an)+"_Luna"+str(luna_data_nerulata)+".xlsx")
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:J', 10)
    worksheet.set_column('K:K', 30)
    worksheet.set_column('L:L', 10)
    formattitlu = workbook.add_format({'align': 'center','bold': True})
    formatcontinut = workbook.add_format({'align': 'center'})
    formatcontinut2 = workbook.add_format({'align': 'left'})
    formatcontinut3 = workbook.add_format({'align': 'left','color': 'red'})
    
    #creare lista in baza fisierului
    for f in RawText:
        if "\n" in f:
            f=f.replace("\n","").replace("\\"," ")#curatare de "\n" de la sfarsitul fiecarei linii din lista initiala 'lista'
            Country1=f[0:2]
            Batch=f[2:8]
            NrAdj=f[8:12]
            Adjustment=f[14:21]
            Contract=f[21:31]
            ContractPrezent=0
            if Contract in VerificareContracte:
                ContractPrezent=1
            VerificareContracte.append(Contract)
            Country2=f[31:33]
            AmountEur=f[33:43].lstrip("0")            
            if AmountEur.isdigit():
                AmountEur=str(int(int(AmountEur)/100))+"."+str(int(AmountEur)%100).rjust(2,"0")
            EUR=f[43:46]
            AmountCurr=f[46:56].lstrip("0")
            if AmountCurr.isdigit():
                AmountCurr=str(int(int(AmountCurr)/100))+"."+str(int(AmountCurr)%100).rjust(2,"0")
            Curr=f[56:59]
            Reason=f[59:89]
            Date=f[103:120]
            if " " in Date:
                Date=Date.replace(" ","")
            NewElement=[Country1,Batch,NrAdj,Adjustment,Contract,Country2,AmountEur,EUR,AmountCurr,Curr,Reason,Date,ContractPrezent]
            Dictionar[len(Dictionar)]=NewElement
    RawText.close()

    L=0
    for e in listacod:
        worksheet.write(0,L,e,formattitlu)#creare cap de tabel in Excel
        L+=1
        
    K=1
    for e in Dictionar:
        L=0#reprezinta coloana pe care scriem datele
        worksheet.write(K,L,Dictionar[e][0],formatcontinut)
        worksheet.write(K,L+1,Dictionar[e][1],formatcontinut)
        worksheet.write(K,L+2,int(Dictionar[e][2]),formatcontinut)
        worksheet.write(K,L+3,Dictionar[e][3],formatcontinut)
        if Dictionar[e][12]==1:
            worksheet.write(K,L+4,Dictionar[e][4],formatcontinut3)
        else:
            worksheet.write(K,L+4,Dictionar[e][4],formatcontinut)
        worksheet.write(K,L+5,Dictionar[e][5],formatcontinut)
        worksheet.write(K,L+6,float(Dictionar[e][6]),formatcontinut)
        worksheet.write(K,L+7,Dictionar[e][7],formatcontinut)
        worksheet.write(K,L+8,float(Dictionar[e][8]),formatcontinut)
        worksheet.write(K,L+9,Dictionar[e][9],formatcontinut)
        worksheet.write(K,L+10,Dictionar[e][10],formatcontinut2)
        worksheet.write(K,L+11,Dictionar[e][11],formatcontinut)
        K=K+1#reprezinta randul pe care scriem datele

        query="INSERT INTO raised VALUES(default,'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}')".format(Dictionar[e][0], Dictionar[e][1],Dictionar[e][2],Dictionar[e][3],Dictionar[e][4],Dictionar[e][5],Dictionar[e][6],Dictionar[e][7],Dictionar[e][8],Dictionar[e][9],Dictionar[e][10],Dictionar[e][11])
        try:
            cur.execute(query)
            print (query)
            db.commit()
            for row in cur.fetchall():
                print row[0]
        except(Exception) as error:
            print (error)
            
    workbook.close()
    LogFile.write(str(time.ctime())+": S-a terminat procesarea fisierului LBROMG"+zi_din_an+".txt\n")
    print (str(time.ctime())+": S-a terminat procesarea fisierului LBROMG"+zi_din_an+".txt\n")
    if os.path.isdir("\\\\__country ftp\\ro\\"+str(an_data_nerulata)):
        trimitere_email(zi_din_an,luna_data_nerulata,an_data_nerulata)
        shutil.move("Raised_"+str(zi_din_an)+"_Luna"+str(luna_data_nerulata)+".xlsx",os.path.join("\\\\__country ftp\\ro\\"+str(an_data_nerulata),"Raised_"+str(zi_din_an)+"_Luna"+str(luna_data_nerulata)+".xlsx"))
    else:
        trimitere_email()
        os.mkdir("\\\\__country ftp\\ro\\"+str(anul_curent))
        shutil.move("Raised_"+str(zi_din_an)+"_Luna"+str(luna_data_nerulata)+".xlsx",os.path.join("\\\\__country ftp\\ro\\"+str(an_data_nerulata),"Raised_"+str(zi_din_an)+"_Luna"+str(luna_data_nerulata)+".xlsx"))
    

if DB_connected == True:
    
    #Daca cumva scriptul nu a rulat intr-o zi/cateva zile la rand, cu prima ocazie cand ruleaza verifica in baza de date ce zile nu au rulat si insereaza default False pe ele
    #preluare lista de zile din baza de date
    query = "SELECT zi_calendaristica FROM zile_rulate_raised"
    lista_zile_DB=[]
    try:
        cur.execute(query)
        db.commit()
        for row in cur.fetchall():
            lista_zile_DB.append(row[0])
    except(Exception) as error:
        print (error)

    for zi in daterange(date(2020, int('01'), int('01')), date(time.localtime().tm_year, time.localtime().tm_mon, time.localtime().tm_mday)):
        if str(zi) not in lista_zile_DB:
            zi_din_an=str(time.strptime(str(zi),"%Y-%m-%d").tm_yday).rjust(3,"0")
            query = "INSERT INTO zile_rulate_raised VALUES (default, '{0}','{1}','false',default)".format(zi_din_an,str(zi))
            try:
                cur.execute(query)
                print (query)
                db.commit()
            except(Exception) as error:
                print (error)

    #se preiau toate inregistrarile din tabele unde avem False la rulat
    query = "SELECT * FROM zile_rulate_raised WHERE rulat='false';"
    try:
        cur.execute(query)
        print (query)
        db.commit()
        for row in cur.fetchall():
            print (row[2])
            data_nerulata=row[2]
            zi_formatata=datetime.strftime(datetime.strptime(data_nerulata,"%Y-%m-%d"),"%d%b%y")
            zi_din_an=str(time.strptime(data_nerulata,"%Y-%m-%d").tm_yday).rjust(3,"0")
            an_data_nerulata=data_nerulata[0:4]
            luna_data_nerulata=data_nerulata[5:7]
            Fisier="\\\\__clona_server\\country ftp\\ro\\LBROMG"+zi_din_an+".TXT"

            if os.path.isfile(Fisier):
                query = "UPDATE zile_rulate_raised SET rulat='true' WHERE zi_din_an='{0}'".format(zi_din_an)
                try:
                    cur.execute(query)
                    print (query)
                    db.commit()
                except(Exception) as error:
                    print (error)
                anul_crearii_fisierului=time.localtime(os.path.getmtime(Fisier)).tm_year
                if an_data_nerulata==str(anul_crearii_fisierului):
                    try:
                        RawText=open(Fisier,"r")
                    except:
                        LogFile.write(str(time.ctime())+": Fisierul"+zi_din_an+".txt nu se poate deschide\n")
                        print (str(time.ctime())+": Fisierul"+zi_din_an+".txt nu se poate deschide\n")
                    else:
                        procedura_main(zi_din_an,luna_data_nerulata,an_data_nerulata)
                else:
                    LogFile.write(str(time.ctime())+": Anul crearii fisierului LBROMG"+zi_din_an+".txt ("+str(anul_crearii_fisierului)+") nu este anul curent.\n")
                    print (str(time.ctime())+": Anul crearii fisierului LBROMG"+zi_din_an+".txt ("+str(anul_crearii_fisierului)+") nu este anul curent.\n")
            else:
                LogFile.write(str(time.ctime())+": Fisierul LBROMG"+zi_din_an+".txt nu exista.\n")
                print (str(time.ctime())+": Fisierul LBROMG"+zi_din_an+".txt nu exista.\n")
                query = "UPDATE zile_rulate_raised SET rulat='true' WHERE zi_din_an='{0}'".format(zi_din_an)
                try:
                    cur.execute(query)
                    print (query)
                    db.commit()
                except(Exception) as error:
                    print (error)

    except(Exception) as error:
        print (error)
       
    db.close()

LogFile.close()
