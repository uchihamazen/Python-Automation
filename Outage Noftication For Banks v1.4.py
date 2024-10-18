from email import message
import win32com.client as client
import pandas as pd
import datetime 
def service_name (x):
    if x == 1:
        return "Transaction Authorization"
    elif x == 2:
        return "Jira"
    elif x == 3:
        return "Citrix"
    elif x == 4:
        return "FTP"
    elif x == 5:
        return "3D Secure"
    elif x == 6:
        return str(input("Enter impacted service:\n"))
def zero(x):
    if x<10:
        return '{:02}'.format(x)
    else:
        return min


xls = pd.ExcelFile(r"\\Server_IP\IT Service Desk\IT Service Desk - Documents\Updated Contact Sheet.xlsx")
df1 = pd.read_excel(xls, 'Banks')
df2 = pd.read_excel(xls, 'Banks')
df3 = pd.read_excel(xls, 'Banks')
to = "itservicedesk"
cc =  ""
banks = int(input ("Enter spcifec banks: 1-TWO 2-FEP 3-All 4-3D Secure Banks 5-Specific Banks\n"))
service = int(input ("Enter Service Name: 1-Transaction Authorization 2-Jira 3-Citrix 4-FTP 5-3D Secure 6-Add it manually\n"))
impacted_service = service_name(service)
ticket = str(input ("Enter Ticket Number:\n"))
hour = int(input ("Enter the hour:\n"))
min = int(input ("Enter the min:\n"))
min_lead = zero(min)
subject = "NI Service Interruption Notification" + " " + str(impacted_service) + " " +  str(ticket)
secure_3d_subject = "NI Service Interruption Notification" + " " + "3D Secure" + " " +  str(ticket)
date ='{dt.day}/{dt.month}/{dt.year}'.format(dt = datetime.datetime.now(datetime.timezone.utc)) 
#///////////////////////////////////////////////////////////////
df1 = df1.loc[df1['Issue'] == "two"]
df2 = df2.loc[df2['Issue'] == "fep"]
df3 = df3.loc[df3['3dsecure'] == "3d"]
df1_final = df1.filter(['Email'])
df2_final = df2.filter(['Email'])
df3_final = df3.filter(['Email'])
df1_index1 = df1_final.iloc[:490]
df1_index2 = df1_final.iloc[491:]
df3_3d1 = df3_final.iloc[:416]
df3_3d2 = df3_final.iloc[417:]
bank_email1 = df1_index1.to_string(index = False, header = False) #TWO BCC 1
bank_email2 = df1_index2.to_string(index = False, header = False) #TWO BCC 2
bank_email3 = df2_final.to_string(index = False, header = False) #FEP BCC
bank_email_3d1 = df3_3d1.to_string(index = False, header = False) #3D Secure BCC 1
bank_email_3d2 = df3_3d2.to_string(index = False, header = False) #3D Secure BCC 2
#////////////////////////////////////////////////////////////




if banks == 5:
    bank = str(input ("Enter bank name:\n"))
    df1 = (df1.loc[df1['Code'] == bank ])
    df1_final = df1.filter(['Email']) 
    df3 = None
    array = [bank]
    while True:
        add = int(input("Do you want to add another bank ? 1- Yes   2-No\n"))
        if add == 1:
            df2 = pd.read_excel(xls, 'Banks')
            bank = str(input ("Enter bank name:\n"))
            df2 = (df2.loc[df2['Code'] == bank ])
            df2 = df2.filter(['Email'])
            df3 = pd.concat([df1_final,df2])
            df1_final = df3
            array.append(bank)
            affected_banks = [i.upper() for i in array]
            print("Affected Banks:" + " " + str(affected_banks))
            

        else:
             break
    bank_email4 = df3.to_string(index = False, header = False)




             


outlook = client.Dispatch("Outlook.Application")

message1 = outlook.CreateItem(0)
message2 = outlook.CreateItem(0)
message3 = outlook.CreateItem(0)
message4 = outlook.CreateItem(0)
message5 = outlook.CreateItem(0)
message6 = outlook.CreateItem(0)
message7 = outlook.CreateItem(0)
if banks == 1:
    message1.Display()
    message2.Display()
    message7.BCC = bank_email3
    message7.To = to
    message7.CC= cc
    message7.Subject = "NI Service Interruption Notification" + " " + "International Transactions (OFFUS)" + " " +  str(ticket)
    message7.Display()
elif banks == 2:
    message3.Display()
elif banks == 3:
    message1.Display()
    message2.Display()
    message3.Display()
elif banks == 4:
    message5.Display()
    message6.Display()
    message5.BCC = bank_email_3d1 
    message5.subject = secure_3d_subject
    message5.To = to 
    message5.CC = cc
    message6.BCC = bank_email_3d2
    message6.subject = secure_3d_subject
    message6.To = to 
    message6.CC = cc
elif banks == 5:
    message4.Display()
    message4.BCC = bank_email4
    message4.subject = subject
message1.To = to
message2.To = to
message3.To = to
message4.To = to
message1.CC = cc
message2.CC = cc
message3.CC = cc
message4.CC = cc
message1.BCC = bank_email1
message2.BCC = bank_email2
message3.BCC = bank_email3
message1.subject = subject
message2.subject = subject
message3.subject = subject



html_body= """<p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;'><span style="font-size:16px;color:#002060;">Dear Valued Client</span></p>
<p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;'><span style="font-size:16px;color:black;">&nbsp;</span></p>
<p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;'><span style="font-size:16px;color:#002060;">Kindly note that we are experiencing service interruption that may impact your ({}),</span><span style="color:black;">&nbsp;</span><span style="font-size:16px;color:#002060;">NI engineers are currently engaged in addressing reported issue. The next update will be in 1 hour or as the status changes.</span></p>
<p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;'><span style="font-size:16px;color:#002060;">&nbsp;</span></p>
<p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;'><span style="font-size:16px;color:#002060;">Appreciate if you can check from your end and confirm if there is any service degradation was detected and in case impact confirmed, please open a Jira Ticket, and reference the below mentioned ({})</span></p>
<p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;'><span style='font-size:15px;font-family:"Times New Roman",serif;color:black;'>&nbsp;</span></p>
<table style="width:508.5pt;border:solid windowtext 1.0pt;">
    <tbody>
        <tr>
            <td colspan="2" style="width:505.5pt;border:solid windowtext 1.0pt;background:#FF7B77;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;text-align:center;line-height:11.55pt;'><strong><span style='font-size:15px;font-family:"Calibri",sans-serif;color:white;'>INCIDENT DETAILS&nbsp;</span></strong><span style="font-size:15px;color:black;">&nbsp;</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:.2in;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">Incident Number&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;padding:  .75pt .75pt .75pt .75pt;height:.2in;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;text-indent:-.25in;line-height:11.55pt;background:white;'><span style="font-size:15px;color:#172B4D;">&nbsp; &nbsp; &nbsp; &nbsp;</span><span style="color:blue;text-decoration:underline;"><a href="https://jiraafrica.network.global/browse/{}"><span style="font-size:15px;">{}</span></a></span><span style="font-size:15px;color:#172B4D;">&nbsp;</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:15.65pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">Company&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;padding:  .75pt .75pt .75pt .75pt;height:15.65pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">Network International (NI)</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">Start Date / Time&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;padding:  .75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">{} &ndash; {}:{} GMT&nbsp;</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">Issue Originator&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;padding:  .75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">NI</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">Summary&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;padding:  .75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">Service Interruption on {}</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">Notes</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;padding:  .75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-right:0in;margin-left:.25in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;margin-top:0in;margin-bottom:8.0pt;text-indent:-.25in;line-height:105%;'><span style="font-size:13px;line-height:105%;font-family:Symbol;">&middot;</span><span style='font-size:9px;line-height:105%;font-family:"Times New Roman",serif;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span><span style="font-size:15px;line-height:105%;color:black;">NI Engineers are currently engaged in addressing reported issue. The next update will be in 1 hour or as the status changes</span><span style="font-size:15px;line-height:105%;color:#201F1E;">.</span></p>
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:#201F1E;">&nbsp;</span></p>
                <p style='margin-right:0in;margin-left:.25in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;margin-top:0in;margin-bottom:8.0pt;text-indent:-.25in;line-height:105%;'><span style="font-size:13px;line-height:105%;font-family:Symbol;">&middot;</span><span style='font-size:9px;line-height:105%;font-family:"Times New Roman",serif;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span><span style="font-size:15px;line-height:105%;color:black;">Bank team to Check and confirm status from their end&nbsp;</span></p>
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">&nbsp;</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">Priority&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;background:  white;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">Medium</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">Status&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;background:  white;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">Open</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">Impact&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;background:  white;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">Under Validation</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">Assigned Group</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;padding:  .75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">IT Service Desk</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">Reported by&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;padding:  .75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;margin:0in;line-height:11.55pt;'><span style="font-size:15px;color:black;">IT Service Desk</span></p>
            </td>
        </tr>
    </tbody>
</table>"""

message1.HTMLBody = html_body.format(impacted_service,ticket,ticket,ticket,date,hour,min_lead,impacted_service)
message2.HTMLBody = html_body.format(impacted_service,ticket,ticket,ticket,date,hour,min_lead,impacted_service)
message3.HTMLBody = html_body.format(impacted_service,ticket,ticket,ticket,date,hour,min_lead,impacted_service)
message4.HTMLBody = html_body.format(impacted_service,ticket,ticket,ticket,date,hour,min_lead,impacted_service)
message5.HTMLBody = html_body.format(impacted_service,ticket,ticket,ticket,date,hour,min_lead,impacted_service)
message6.HTMLBody = html_body.format(impacted_service,ticket,ticket,ticket,date,hour,min_lead,impacted_service)
message7.HTMLBody = html_body.format("International Transactions (OFFUS)",ticket,ticket,ticket,date,hour,min_lead,"International Transactions (OFFUS)")

# Copyright (c) [2022] [Mazen Yasser]
# LinkedIn: [https://www.linkedin.com/in/mazen-yasser-998b10217]

# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"),
# to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
# and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

# 1. The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
# 2. Modifications to the Software are strictly prohibited.
# 3. Redistributions of the Software must retain the above copyright notice, this list of conditions and the following disclaimer.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
# DEALINGS IN THE SOFTWARE.

# Prohibition on Modification:
# Redistribution or modification of the Software, in whole or in part, is strictly prohibited without explicit written permission from the author.
