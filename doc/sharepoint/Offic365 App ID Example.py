import sys

from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

from office365.sharepoint.listitems.caml.caml_query import CamlQuery  
from flask import Flask, redirect,  url_for
from office365.sharepoint.client_context import ClientContext

app = Flask(__name__)
 

@app.route("/")
def index(): 

   
   site_url = 'https://macysinc.sharepoint.com/sites/OSO/'
   app_principal = {
        'client_id': '42e8f0a6-aa09-4833-bdfa-6e27f60df7b0',
        'client_secret': 'D8iBldyvflj2iQnEOQY04/O84+JTAqzTqWpTmm6EZu0='        
   }
       
   context_auth = AuthenticationContext(url=site_url)
   context_auth.acquire_token_for_app(client_id=app_principal['client_id'], client_secret=app_principal['client_secret'])
       
   ctx = ClientContext(site_url, context_auth)
   web = ctx.web
   ctx.load(web)
   ctx.execute_query()
   print("Web site title: {0}".format(web.properties['Title']))

   list_title = "StressTest"

   target_folder = ctx.web.lists.get_by_title(list_title).root_folder
   target_list   = ctx.web.lists.get_by_title(list_title)
   

   SP_DOC_LIBRARY ='StressTest'


   caml_query = CamlQuery()
   caml_query.ViewXml = '''<View Scope="RecursiveAll"><Query><Where><Eq><FieldRef Name='FileRef' /><Value Type='Text'>/sites/OSO/StressTest/01001399/AXYZ123</Value></Eq></Where></Query></View>'''

   #caml_query =CamlQuery.parse(qry_text)
   caml_query.FolderServerRelativeUrl = SP_DOC_LIBRARY 
    
   # 3 Retrieve list items based on the CAML query 
   oList = ctx.web.lists.get_by_title(SP_DOC_LIBRARY) 
   items = oList.get_items(caml_query) 
   ctx.execute_query()

   # 5. Loop through all list items
   for item in items: 
     item.set_property('InspectionID', 'Yeah')
     item.update()
     ctx.execute_query()
     print('File downloaded :{0}'.format(item.get_property('Title')))


   return ('return') 

if __name__ == '__main__':
    app.run('0.0.0.0', debug=True)
