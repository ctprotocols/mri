# -*- coding: utf-8 -*-
"""
Created on Wed Jul 28 14:54:42 2021

@author: EastmanE
"""

# def update_site():
import pandas as pd
import datetime as dt
import os


#Website updates
#date and time as string for recordkeeping  



folder_new = r'\\ESSNFS01\CT_Protocols\new' 
folder_3T = r'C:\Users\EastmanE\Box\Imaging (BOX Share)\Medical Physics\MRI Protocol Website\Protocols\3 T'
folder_1T = r'C:\Users\EastmanE\Box\Imaging (BOX Share)\Medical Physics\MRI Protocol Website\Protocols\1.5 T'



updatedate = dt.datetime.today().strftime('%m/%d/%Y %H:%M')

mripath = r'\\ESSNFS01\CT_Protocols\mri_testing_update.html'

message = """<html>
    <head>
        <link rel="stylesheet" href="mystyle.css">
    </head>

    <div class = row1>
        <div id="main" >
            <div class="sidenav">
            <img src ="images/brand-content-logo.png" width = "200px" alt="logo" style="vertical-align:middle; padding-left: 0px; padding-top: 20px;">
            
            <h1 id = "title"> MRI Protocols </h1>

            <h2 style="font-size:20px"> 1.5 T </h2>
            """
              
bodyorder = ['Neuro', 'Body', 'Upper Extremity', 'Lower Extremity']  

neuro = ['Brain', 'Brain_Pituitary', 'Neck', 'C-Spine', 'T-Spine', 'L-Spine', 'Complete Spine', 'Pelvis_Neuro', 'Sacrum _ Coccyx _ SI Joints', 'Brachial Plexus']
body = ['Chest', 'Breasts', 'Abdomen', 'Pelvis', 'MRA Extremity']
upper = ['Spine Survey (rheumatology)', 'Sternum', 'Sternoclavicual Joints', 'Chest MSK', 'Clavicle', 'Shoulder', 'Scapula', 'Humerus', 'Elbow', 'Forearm', 'Wrist', 'Hand', 'Finger(s)']
lower = ['Pelvis MSK', 'Hip', 'Hip Bilat', 'Femur Left', 'Femur Bilat', 'Knee_bad', 'Tib-Fib' , 'Tib-Fib Bilat', 'Ankle', 'Foot Toes']

bodylists = [neuro, body, upper, lower]
          
for idx, subfold_3T in enumerate(bodyorder):
    message = message + """<details> <summary id = "heading1">""" + subfold_3T + """</summary>"""
    thislist = bodylists[idx]
    
    for file_3T in thislist:
        file3Tog = file_3T
        file_3T = file_3T + '.xlsx'
        message =message + """<details> <summary id = "heading2" >""" + file_3T.split('.')[0] +""" </summary><ul style="list-style-type:disc;"> """
        
        list_3T = list(pd.read_excel(os.path.join(folder_1T, subfold_3T, file_3T), sheet_name = None).keys())
        # list_3T.sort()
        for tab_3T in list_3T:
            path_newhtml = """//ESSNFS01/CT_Protocols/new/1.5T/""" + file3Tog + ' ' + tab_3T +'.html'
            displayname = file_3T.split('.')[0] + ' ' + tab_3T
            f = open(path_newhtml, 'w') 
            message_new = """<html> <style>
table, th, td {
  background-color: #F0F8FF;
  color: black;
    border: 1px solid black;
  border-collapse: collapse;
  font-family: Helvetica;
  font-size: 14px;
}
table.center {
    text-align: center;
}
</style><table class = "center"> <tr>"""
            df = pd.read_excel(os.path.join(folder_1T, subfold_3T, file_3T), sheet_name = tab_3T, usecols="A:N")
            for col in df.columns:
                message_new = message_new + '<th>' + col + '</th>'
            message_new += '</tr>'
            
            for row in df.index:
                message_new += '<tr>'
                for col in df.columns:
                    if str(df.loc[row, col]) == 'nan':
                        df.loc[row, col] = ''
                    message_new = message_new + '<td>' + str(df.loc[row, col]) + '</td>'
                message_new += '</tr>'
            message_new += '</table>'
            
            f.write(message_new)
            
            f.close()

            
            message = message + """<li id = '"""+ displayname + "'><a href = '"""+ path_newhtml+ """' target="iframe_a" onClick="myFunction('"""+ displayname + """');">""" + tab_3T +""" </a></li>"""

        message += """</ul></details>"""
    message += """</details>"""            
    
    
    
    
     
message+= """<h2 style="font-size:20px"> 3 T </h2>
            """
              
bodyorder = ['Neuro', 'Body', 'Upper Extremity', 'Lower Extremity']  

neuro = ['Brain', 'Pituitary', 'Neck', 'C-Spine', 'T-Spine', 'L-Spine', 'Complete Spine', 'Pelvis_Neuro', 'Brachial Plexus']
body = ['Chest', 'Breast', 'Abdomen', 'Pelvis']
upper = ['MRA Extremity', 'Spine Survey (rheumatology)', 'Sternum', 'SC Joints', 'Chest_MSK', 'Clavical', 'Shoulder', 'Scapula', 'Humerus', 'Elbow', 'Wrist', 'Hand', 'Finger(s)']
lower = ['Pelvis_MSK', 'Sacrum_Coccyx_SI Joints', 'Hip', 'Hip Bi_Lateral', 'Femur', 'Femur Bilat.', 'Knee', 'Tib-Fib' , 'Tib-Fib Bilat.', 'Ankle', 'Foot Toe', 'Whole Body Screening']

bodylists = [neuro, body, upper, lower]
          
for idx, subfold_3T in enumerate(bodyorder):
    message = message + """<details> <summary id = "heading1">""" + subfold_3T + """</summary>"""
    thislist = bodylists[idx]
    
    for file_3T in thislist:
        file3Tog = file_3T
        file_3T = file_3T + '.xlsx'
        message =message + """<details> <summary id = "heading2" >""" + file_3T.split('.')[0] +""" </summary><ul style="list-style-type:disc;"> """
        
        list_3T = list(pd.read_excel(os.path.join(folder_3T, subfold_3T, file_3T), sheet_name = None).keys())
        # list_3T.sort()
        for tab_3T in list_3T:
            path_newhtml = """//ESSNFS01/CT_Protocols/new/3T/""" + file3Tog + ' ' + tab_3T +'.html'
            displayname = file_3T.split('.')[0] + ' ' + tab_3T
            f = open(path_newhtml, 'w')
            message_new = """<html> <style>
table, th, td {
  background-color: #F0F8FF;
  color: black;
    border: 1px solid black;
  border-collapse: collapse;
  font-family: Helvetica;
  font-size: 14px;
}
table.center {
    text-align: left;
}
</style><table class = "center"> <tr>"""
            df = pd.read_excel(os.path.join(folder_3T, subfold_3T, file_3T), sheet_name = tab_3T, usecols="A:N")
            for col in df.columns:
                message_new = message_new + '<th>' + col + '</th>'
            message_new += '</tr>'
            
            for row in df.index:
                message_new += '<tr>'
                for col in df.columns:
                    if str(df.loc[row, col]) == 'nan':
                        df.loc[row, col] = ''
                    message_new = message_new + '<td>' + str(df.loc[row, col]) + '</td>'
                message_new += '</tr>'
            message_new += '</table>'
            
            f.write(message_new)
            
            f.close()

            
            message = message + """<li id = '"""+ displayname + "'><a href = '"""+ path_newhtml+ """' target="iframe_a" onClick="myFunction('"""+ displayname + """');">""" + tab_3T +""" </a></li>"""

        message += """</ul></details>"""
    message += """</details>"""            

message = message + "<p>Updated " + updatedate + """</p>"""

message +=            """
        </div>
        </div>
    <div class="column2">
        <div id="main" >
            <h1 id="demo" >Select a protocol.</h1>
            <iframe frameBorder = "0" name="iframe_a" scrolling="no" style="position:relative; height:1200px; width: 1200px; top: 0px; padding-left: 400px; overflow:hidden;"></iframe>
        </div>

    </div>
<script>

  
  function myFunction(myBtn) {
    document.getElementById("demo").innerHTML =  myBtn;  
    }
</script>"""


message += """ </html>"""

f = open(mripath, 'w')

f.write(message)

f.close()
print('Website updated. Program is complete.')



