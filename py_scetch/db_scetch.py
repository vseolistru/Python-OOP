import sqlite3
import logging
conn = sqlite3.connect("new_db.db") 
cursor = conn.cursor()
cursor.execute("""CREATE TABLE IF NOT EXISTS "table_link" (
	"requests",
        "position_id",  
	"link_ardess"
     )""")
conn.commit()
conn.close()
word1=['word1','word2','word3','word4']
word=['word6','word7','word8']
request1=word[0]
request=word[1]
position=(45)
position1=(42)
conn = sqlite3.connect("new_db.db") 
cursor = conn.cursor()
cursor.execute("UPDATE table_link SET requests='"+str(request)+"', position_id='"+str(position)+"' WHERE id=1") 
cursor.execute("UPDATE table_link SET requests='"+str(request1)+"', position_id='"+str(position1)+"' WHERE id=2")                
conn.commit()  
conn.close()

list_requests=[]
links_position=[]
conn = sqlite3.connect("from_db.db") 
cursor = conn.cursor()
sql= "SELECT position_id FROM new_request" ##запросы имя ряда, имя таблицы
cursor.execute(sql)
row=cursor.fetchall()
conn.close() 
for i in row:
  links_position.append(i)
conn = sqlite3.connect("from_db.db") 
cursor = conn.cursor()
sql= "SELECT requests_db FROM new_request" ##запросы имя ряда, имя таблицы
logging.basicConfig(filename="sample.log",format='%(message)s--%(asctime)s;', level=logging.INFO)
cursor.execute(sql)
row=cursor.fetchall()
conn.close() 
for i in row:
  list_requests.append(i)
res=[dict(zip(suba, subb)) for (suba, subb) in zip(list_requests, links_position)]
for i in res:
  logging.info(i)       
print('ok')
logging.info('                '+'sample.log-finished')








