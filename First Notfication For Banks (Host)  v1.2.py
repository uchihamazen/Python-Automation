from email import header
import win32com.client as client
import pandas as pd 
import datetime

def zero(x):
    if x<10:
       return '{:02}'.format(x)
    else:
        return min

xls = pd.ExcelFile(r"\\192.168.17.10\IT Service Desk\IT Service Desk - Documents\Updated Contact Sheet.xlsx")
df1 = pd.read_excel(xls, 'Banks')
df2 = pd.read_excel(xls, 'Codes')  
df3 = pd.read_excel(xls, 'Links')
bank = str(input ("Enter Bank Code:\n"))
while True:
    filter = df2.filter(['Code'])
    dropna = filter.dropna()
    find = bank in dropna.values
    if  find == False:
       print("Warning: You have entered a wrong bank, please try again !")
       bank = str(input ("Enter Bank Code:\n")) 
    else:
        break  
host_choose = int(input ("Select Host: 1-TCI 2-TIC 3-ATM 4-IFlex 5-123 Switch 6-GIM 7-GHIPSS 8-Issuer Arksys 9-IFlex 10-Fawry 11-mVisa 12-FIB Gateway 13-AAIB H2H POS Premium\n"))
incident = int(input ("Select Issue : 1-Discconection  2-Timeout & TRXs Failure 3-BS Offilne  4-Signed Off 5-Instability 6-System Error 7-Payment Not Connected\n"))
date ='{dt.day}/{dt.month}/{dt.year}'.format(dt = datetime.datetime.now(datetime.timezone.utc)) 
hour = int(input ("Enter the hour:\n"))
min = int(input ("Enter the min:\n"))
min_lead = zero(min)
df1 = (df1.loc[df1['Code'] == bank ])
df1_final = df1.filter(['Email'])
bank_email = df1_final.to_string(index = False, header = False)
bank_index = df2[df2['Code'] == bank].index[0]
bank_name = df2.iloc[bank_index]['Name']
code_name = df2.iloc[bank_index]['Project']
df3 = (df3.loc[df3['Codes'] == bank ])
df3_final = df3.filter(['Link'])
links = df3_final.to_string(index = False, header = False)


if host_choose == 1:
     x = "TCI"
elif host_choose == 2:
    x = "TIC"
elif host_choose == 3:
    x = "ATM"
elif host_choose == 4:
    x = "IFlex"
elif host_choose == 5:
    x = "123 Switch"
elif host_choose == 6:
    x = "GIM"
elif host_choose == 7:
    x = "GHIPSS"
elif host_choose == 8:
    x = "Issuer Arksys"
elif host_choose == 9:
    x= "IFlex"
elif host_choose == 10:
    x= "Fawry"
elif host_choose == 11:
    x= "mVisa"
elif host_choose == 12:
    x= "FIB Gateway"
elif host_choose == 13:
    x= "AAIB H2H POS Premium"
if incident == 1:
     y = "Disconnection"
elif incident == 2:
    y = "Timeout & TRXs Failure"
elif incident == 3:
    y = "BS Offline"
elif incident == 4:
    y = "Signed Off"
elif incident == 5:
    y = "Instability"
elif incident == 5:
    y = "System Error"
elif incident == 6:
    y = "TRXs failure"
elif incident == 7:
    y = "Payment Not Connected"

    
print("Your Jira Description:"+ " " + str(code_name)+ " " + "-" + " " + str(bank_name) + " " + "-" + " " + str(x) + " " + "Host Interface" + " " + str(y))
jira = str(input ("Enter Jira Ticket\n"))


outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)
message.Display()
message.To = bank_email
message.CC = "ITservicedesk.dl@network.global;cse.dl@network.global;rd.africa@emp-group.com;onlinesupport.Africa@network.global;Network.Africa@network.global;Mohamed.Hafez@network.global;possupport.Africa@network.global"
message.Subject = str(code_name)+ " " + "-" + " " + str(bank_name) + " " + "-" + " " + str(x) + " " + "Host Interface" + " " + str(y) + " " + str(jira)
html_body= """<p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:normal;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">Dear Valued Client,</span></p>
<p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:normal;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">&nbsp;</span></p>
<p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:normal;font-size:15px;font-family:"Calibri",sans-serif;text-indent:.5in;'><span style="color:black;">This is NIA IT-Service Desk team notifying you of an anomaly affecting your host connectivity&nbsp;with the following details:</span></p>
<p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:normal;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-family:"Times New Roman",serif;color:black;'>&nbsp;</span></p>
<table style="width:508.5pt;border:solid windowtext 1.0pt;">
    <tbody>
        <tr>
            <td colspan="2" style="width:505.5pt;border:solid windowtext 1.0pt;background:#FF7B77;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;text-align:center;'><strong><span style="color:white;">INCIDENT DETAILS&nbsp;</span></strong><span style="color:black;">&nbsp;</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:.2in;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">Incident Number&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;padding:  .75pt .75pt .75pt .75pt;height:.2in;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:  11.55pt;font-size:15px;font-family:"Calibri",sans-serif;text-indent:-.25in;background:white;'><span style="color:#172B4D;">&nbsp; &nbsp; &nbsp; &nbsp;</span><a href="https://jiraafrica.network.global/browse/{}">{}</a></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:15.65pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">Company&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;padding:  .75pt .75pt .75pt .75pt;height:15.65pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">{}</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">Start Date / Time&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;padding:  .75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">{} &ndash; {}:{} GMT&nbsp;</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">Issue Originator&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;padding:  .75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">Client</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">Summary&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;padding:  .75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">{} Host Interface {}</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">Issue Description/Notes&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;padding:  .75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">Next update will be in 30 Minute&nbsp;</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">Priority&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;padding:  .75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">Medium&nbsp;</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">Status&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;padding:  .75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">Open</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">Impact&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;background:  white;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">Limited &ndash; Transaction Authorization impact</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:44.3pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">Network Connection status&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;padding:  .75pt .75pt .75pt .75pt;height:44.3pt;">
               {}
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">Reported by &nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;padding:  .75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">Service Desk</span></p>
            </td>
        </tr>
        <tr>
            <td style="width:137.25pt;border:solid windowtext 1.0pt;background:  #E5E4E2;padding:.75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:black;">Additional Info&nbsp;</span></p>
            </td>
            <td style="width:366.75pt;border:solid windowtext 1.0pt;padding:  .75pt .75pt .75pt .75pt;height:13.05pt;">
                <p style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;line-height:11.55pt;font-size:15px;font-family:"Calibri",sans-serif;'>Please restart the host from your end&nbsp;</p>
            </td>
        </tr>
    </tbody>
</table>
"""

message.HTMLBody = html_body.format(jira,jira,bank_name,date,hour,min_lead,x,y,links)


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

