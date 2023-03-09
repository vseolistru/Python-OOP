import sqlite3

table1 = 'sqlite_sequence'
table2 = 'test_table'

word1=['word1','word2','word4','word5']
word=['word6','word7','word8']


class Sql_driver: 

  def __init__(self, word, table):
    self.word = word
    self.table  = table
   
  def create ():
    conn = sqlite3.connect("test_db.db") 
    cursor = conn.cursor()
    cursor.execute("""CREATE TABLE IF NOT EXISTS "test_table" (
	"id" INTEGER PRIMARY KEY AUTOINCREMENT,
        "test_data" #TEXT 
	
       )""")
    conn.commit()
    conn.close()

  def update():
    conn = sqlite3.connect("test_db.db") 
    cursor = conn.cursor()
    cursor.execute("UPDATE test_table SET test_data='"+str(i)+"' WHERE id>=1")                
    conn.commit()  
    conn.close()

  def select():
    list_gen = []
    conn = sqlite3.connect("test_db.db") 
    cursor = conn.cursor()
    cursor.execute("SELECT test_data FROM test_table")
    row = yield cursor.fetchall()
    return row
    conn.close() 


  def insert(self, new_data):
    conn = sqlite3.connect("test_db.db") 
    cursor = conn.cursor()
    for i in new_data:
      cursor.execute("INSERT INTO '"+str(table2)+"' (test_data) VALUES ('"+str(i)+"') ") 
    print('success_insert')
    conn.commit()  
    conn.close() 
  
  def delete(self, table):
    conn = sqlite3.connect("test_db.db") 
    cursor = conn.cursor()
    cursor.execute("DELETE  FROM '"+str(table)+"'")             
    conn.commit()  
    conn.close()

Sql_driver.delete(word, table2)
Sql_driver.delete(word,table1)

Sql_driver.insert(table2, word)
res =list( Sql_driver.select())



print (res)

     

  
    

Sql_driver.delete(word,table1)

Sql_driver.delete(word, table2)

Sql_driver.insert(table2,word1)
res = Sql_driver.select()
print (res)






