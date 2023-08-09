#!/usr/bin/env python3
import nmap3
import sqlite3 
from sqlite3 import Error
import xlsxwriter 
# creates an Excel file
wrkbk = xlsxwriter.Workbook('data.xlsx') 
# creates a page in Excel
wrksht = wrkbk.add_worksheet()  
# scans the results of nmap
nmap = nmap3.Nmap()  
# creates a database table to connect
connect = sqlite3.connect('database.db')
# creates a table to list information including a primary key
connect.execute('''create table if not exists scan (
    id INTEGER primary key autoincrement,
    name text not null,
    ip text not null,
    protocol text not null,
    service text not null, 
    port_id not null);''')  
# creates a tables that connects foreign key to primary key table
connect.execute('''create table if not exists port (
    id INTEGER primary key autoincrement,
    reason text not null,
    stats text not null,
    runtime text not null,
    task_results text not null,
    host_id integer not null, 
    FOREIGN KEY (host_id) REFERENCES scan(id));''') 
# saves the tables after committing 
connect.commit() 
 
class Nmap: 
    def __init__(self, targetname) -> None: 
        self.targetname = targetname 
        self.hostname=""
        self.protocol=""
        self.service=""
        self.port_id=""
        self.ip="" 
        self.ipname=""
        self.reason="" 
        self.stats="" 
        self.runtime="" 
        self.task_results="" 
        self.host_id=""
    def PortScanner(self):
        results = nmap.scan_top_ports(self.targetname) 
        ii=0
        port_id=1
        for i in results:
            print(i)
            if (ii==0):
                self.ip = i
                hostname=results[self.ip]['hostname'][0]['name']
                protocol=results[self.ip]['ports'][0]['protocol']
                service=results[self.ip]['ports'][0]['service']['name']
                port_id=results[self.ip]['ports'][0]['portid']
                self.hostname=hostname
                self.protocol=protocol
                self.service=service 
                self.port_id=port_id
                print(self.hostname)
                print(self.protocol)
                print(self.service) 
                print(self.port_id)
                ii+=1 
                
        print(results) 

    def IpScanner(self): 
        total = nmap.scan_top_ports(self.targetname) 
        ii=0 
        host_id=1
        for i in total:
            print(i) 
            if (ii==0):
                self.task_results = i 
                reason=total[self.task_results]['ports'][0]['reason'] 
                stats=total['stats']['args'] 
                runtime=total['runtime']['timestr'] 
                task_results=total['task_results'][2]['task'] 
                self.reason=reason
                self.stats=stats
                self.runtime=runtime 
                self.task_results=task_results
                self.host_id=host_id
                print(self.reason) 
                print(self.stats) 
                print(self.runtime) 
                print(self.task_results) 
                print(self.host_id)
                ii+=1
                
        print(total)     

    # def NmapScanner(self):
    #     results2 = nmap.nmap_version_detection(self.targetname) 
    #     ii=0 
    #     port_id=1
    #     for i in results2:
    #         print(i)
    #         if (ii==0):
    #             self.ip = i
    #             hostname=results2[self.ip]['hostname'][0]['name']
    #             protocol=results2[self.ip]['ports'][0]['protocol']
    #             service=results2[self.ip]['ports'][0]['service']['name'] 
    #             port_id=results2[self.ip]['ports'][0]['portid']
    #             self.hostname=hostname
    #             self.protocol=protocol
    #             self.service=service
    #             self.port_id=port_id
    #             print(self.hostname)
    #             print(self.protocol)
    #             print(self.service) 
    #             print(self.port_id)
    #             ii+=1 
    #     print(results2)  
        
    # def DataScanner(self):
    #     total2 = nmap.nmap_version_detection(self.targetname) 
    #     ii=0
    #     host_id=1
    #     for i in total2:
    #         print(i)
    #         if (ii==0):
    #             self.task_results = i
    #             reason=total2[self.task_results]['ports'][0]['reason']
    #             stats=total2['stats']['args']
    #             runtime=total2['runtime']['timestr'] 
    #             task_results=total2['task_results'][2]['task']
    #             self.reason=reason
    #             self.stats=stats
    #             self.runtime=runtime 
    #             self.task_results=task_results
    #             self.host_id=host_id
    #             print(self.reason)
    #             print(self.stats)
    #             print(self.runtime) 
    #             print(self.task_results)
    #             print(self.host_id)
    #             ii+=1 
    #     print(total2)  
        
    def CRUD(self): 
        try:
            connect.execute('''insert into scan (name, ip, protocol, service, 
port_id)
                values (?, ?, ?, ?, ?)''',
                (self.hostname, self.ip, self.protocol, self.service, 
self.port_id)) 
            connect.execute('''insert into port (reason, stats, runtime, 
task_results, host_id)
                values (?, ?, ?, ?, ?)''',
                (self.reason, self.stats, self.runtime, self.task_results, 
self.host_id))  
        except: 
            print("Fail to insert data values in database tables") 
        finally:
            print("")
        
        wrksht.write('A1', "Hostname")  
        wrksht.write('B1', "IP Address")  
        wrksht.write('C1', "Protocol")  
        wrksht.write('D1', "Service")  
        wrksht.write('E1', "Reason") 
        wrksht.write('F1', "Stats")  
        wrksht.write('G1', "Runtime") 
        wrksht.write('H1', "Task Results") 
        if self.hostname=="google.com":
            wrksht.write('A2', self.hostname)  
            wrksht.write('B2', self.ip) 
            wrksht.write('C2', self.protocol) 
            wrksht.write('D2', self.service) 
            wrksht.write('E2', self.reason) 
            wrksht.write('F2', self.stats) 
            wrksht.write('G2', self.runtime) 
            wrksht.write('H2', self.task_results) 
        elif self.hostname=="amazon.com":
            wrksht.write('A3', self.hostname)  
            wrksht.write('B3', self.ip)  
            wrksht.write('C3', self.protocol)  
            wrksht.write('D3', self.service)  
            wrksht.write('E3', self.reason) 
            wrksht.write('F3', self.stats) 
            wrksht.write('G3', self.runtime) 
            wrksht.write('H3', self.task_results) 
        elif self.hostname=="microsoft.com":
            wrksht.write('A4', self.hostname)  
            wrksht.write('B4', self.ip)  
            wrksht.write('C4', self.protocol)  
            wrksht.write('D4', self.service) 
            wrksht.write('E4', self.reason) 
            wrksht.write('F4', self.stats) 
            wrksht.write('G4', self.runtime) 
            wrksht.write('H4', self.task_results) 
       
        elif self.hostname=="netflix.com":
            wrksht.write('A5', self.hostname)  
            wrksht.write('B5', self.ip)  
            wrksht.write('C5', self.protocol)  
            wrksht.write('D5', self.service)  
            wrksht.write('E5', self.reason) 
            wrksht.write('F5', self.stats)  
            wrksht.write('G5', self.runtime) 
            wrksht.write('H5', self.task_results) 
        elif self.hostname=="apple.com":
            wrksht.write('A6', self.hostname)  
            wrksht.write('B6', self.ip)  
            wrksht.write('C6', self.protocol)  
            wrksht.write('D6', self.service)  
            wrksht.write('E6', self.reason) 
            wrksht.write('F6', self.stats)  
            wrksht.write('G6', self.runtime) 
            wrksht.write('H6', self.task_results) 
        
            
        connect.commit()  
        
    
# prints the Google results from each method 
google = Nmap("google.com") 
google.PortScanner() 
google.IpScanner() 
# google.NmapScanner() 
# google.DataScanner()
google.CRUD() 
# prints the Amazon results from each method
amazon = Nmap("amazon.com") 
amazon.PortScanner() 
amazon.IpScanner() 
# amazon.NmapScanner()
# amazon.DataScanner()
amazon.CRUD()   
# prints the Microsoft results from each method
microsoft = Nmap("microsoft.com") 
microsoft.PortScanner() 
microsoft.IpScanner() 
# microsoft.NmapScanner() 
# microsoft.DataScanner()
microsoft.CRUD() 
# prints the Netflix results from each method
netflix = Nmap("netflix.com") 
netflix.PortScanner() 
netflix.IpScanner() 
# netflix.NmapScanner()
# netflix.DataScanner() 
netflix.CRUD()  
# prints the Apple results from each method 
apple = Nmap("apple.com") 
apple.PortScanner() 
apple.IpScanner()
# apple.NmapScanner() 
# apple.DataScanner() 
apple.CRUD() 
# closes the connections from Excel
wrkbk.close()     
# closes the connections from database table
connect.close()   