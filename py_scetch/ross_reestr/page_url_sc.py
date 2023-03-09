import requests
import re


def search_ft_pages_links(url):
  r = requests.get(url)
  resp = r.text
  links_to_test = []
  re_links = re.findall('<a href=\".*.\/["]', resp)
  for i in re_links:
    a_replace =''.join(i).replace('<a href="','').replace('"','').replace(url[8:],'').replace('https:/','').replace('http:/','')[1:]
    link_string = url+a_replace
    links_to_test.append(link_string)
  return links_to_test
  
  
  
def search_pages_to_request(links_to_test):
  links_to_test = links_to_test
  list_url_test_result = []
  for i in (links_to_test):
       try:
         r = requests.get (str(i))
         r_code = r.status_code  
         result = i + " " + str(r_code)         
       except:    
         code = '502'
         result =i+" " +code   
       list_url_test_result.append(result)
  return list_url_test_result
  
   

url = 'https://www.sgryadki.ru/'
links_to_test = search_ft_pages_links(url)


new_result = search_pages_to_request(links_to_test)
for i in new_result:
  print(i)






