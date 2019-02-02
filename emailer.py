"""
Searches through Outlook e-mails for keywords and displays the keywords plus the previous sentence and next sentence for context.
Results are buffered in order to display in emailsearchGUI

searchstring: keyword
foldertolookin: which folder range in Outlook to look in, by default, only inbox will be searched
desiredsender: user can specify a string to filter by sender name/sender e-mail (should search both)
attachmentsonly: whether search is only for e-mails that contain attachments
getduplicates: long e-mail chains bring up the same keyword results as replies. Turning this off helps reduce redundant results
desireddate: how far back to look. I have found that I always look up to present day so creating a second cutoff date has not been necessary
"""


from win32com.client import Dispatch
import io
from datetime import datetime, timedelta

outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

def findstringinemail(searchstring, foldertolookin=None, desiredsender = "", attachmentsonly = False ,getallduplicates = True, desireddate = (datetime.now() + timedelta(days=-30) ).strftime("%y-%m-%d %H:%M")):
    buffer = io.StringIO()
    isuniquesubject = []
    global total
    
    total=0
    outlook = Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    #default to look in inbox folder
    if foldertolookin==None:
        foldertolookin = Dispatch("Outlook.Application").GetNamespace("MAPI").GetDefaultFolder("6")
    elif len(foldertolookin)==5:
        foldertolookin=namespace.Folders.Item(foldertolookin[0]).Folders[(foldertolookin[1])].Folders[(foldertolookin[2])].Folders[(foldertolookin[3])].Folders[(foldertolookin[4])]  
    elif len(foldertolookin)==4:
        foldertolookin=namespace.Folders.Item(foldertolookin[0]).Folders[(foldertolookin[1])].Folders[(foldertolookin[2])].Folders[(foldertolookin[3])]    
    elif len(foldertolookin)==3:
        foldertolookin=namespace.Folders.Item(foldertolookin[0]).Folders[(foldertolookin[1])].Folders[(foldertolookin[2])]
    elif len(foldertolookin)==2:
        foldertolookin=namespace.Folders.Item(foldertolookin[0]).Folders[(foldertolookin[1])]
    elif len(foldertolookin)==1:
        foldertolookin=namespace.Folders.Item(foldertolookin[0])
    print ("[][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][]\n"+
        "+++++++++++++++++++++++++++++++++++++NOW SEARCHING IN: ", foldertolookin, "+++++++++++++++++++++++++++++++++++++++ \n" + 
    "[][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][]" , file = buffer)
    messages = foldertolookin.Items
    for msg in reversed(foldertolookin.Items):
        try:
            dateofmessage = msg.SentOn.strftime("%y-%m-%d %H:%M")
        except:
            print("THERE WAS AN EXCEPTION - DATE", msg.Subject, file = buffer)
            dateofmessage = datetime.now().strftime("%y-%m-%d %H:%M") 
        if desireddate < dateofmessage:
            if msg.Class==43:
                if msg.SenderEmailType=='EX':
                    sendersemail = msg.Sender.GetExchangeUser().PrimarySmtpAddress
                else:
                    sendersemail = msg.SenderEmailAddress
            try:
                msg.Subject
            except:
                print("THERE WAS AN EXCEPTION - SUBJECT ", file = buffer)
                print (sendersemail, file = buffer)
                
            if "Undeliverable" in msg.Subject:
                msg = messages.GetNext ()


            try:
                if hasattr(msg, "SenderName"):
                    name = msg.SenderName
            except:
                print("THERE WAS AN EXCEPTION - NAME", file = buffer)
                print (msg.Subject, file = buffer)   
                print (sendersemail, file = buffer) 
                name = "UNKNOWN"
                
            if searchstring in msg.Subject or  searchstring in msg.body :
                print (msg.Subject, dateofmessage)
                attachments = []
                for att in msg.Attachments:
                    attachments.append(str(att.Filename))
                if (desiredsender in sendersemail or desiredsender in name):
                    
                    total += 1
                    
                    print("message number: ", total, file = buffer)
                    if searchstring in msg.Subject :
                        
                        if msg.Subject not in isuniquesubject or getallduplicates ==True:
                            if attachmentsonly == False or attachments:
                                print("IN SUBJECT: ", dateofmessage,"        ", sendersemail, " " , name ,"    " , msg.Subject,"\n", attachments, file = buffer)
                                print("----------------------------------------------------------------------------------------------------------", file = buffer)
                                
                                isuniquesubject.append(str(msg.Subject))
                    
                    if searchstring in msg.body:
    
                        if msg.Subject not in isuniquesubject or getallduplicates ==True:
                            if attachmentsonly == False or attachments:
                #arrange each message by text line, and into a list of such lines     
                                print("IN TEXT: ", dateofmessage,"        ", sendersemail, " " , name ,"    "  , msg.Subject, "\n", attachments, file = buffer)   
                            
                                isuniquesubject.append(str(msg.Subject))
                                messageline = msg.body.splitlines()
                                messageline = [x for x in messageline if x.strip()]
                        
                        
                                for x in range(len(messageline)):
                        
                                    #set message limit so that long reply chains arent captured
                                    if searchstring in messageline[x]:
                                        #want to get the previous and next sentence for context of the search string, if they exist
                                        try:
                                            print(messageline[x-1], file = buffer)
                                        except:
                                            pass

                                        print(messageline[x], file = buffer)
                                        
                                        try:
                                            print(messageline[x+1], file = buffer)
                                        except:
                                            pass
        
                                print("----------------------------------------------------------------------------------------------------------", file = buffer)


    print("TOTAL NUMBER OF MESSAGES IN THIS FOLDER: ", total, file = buffer)
    return buffer
    


if __name__ == '__main__': 
    '''
    testt=getemailstructure(["Inbox","CAQ", "FAI","Hanley", "Ming", "JP"])
    
    saveresults('valuefile' ,testt)
    
    text1 = open('valuefile.txt', 'r')
    foldertolookin = text1.read()
    print ("here", foldertolookin)
    text1.close()'''

    searchstring = " "
    '''
    foldertolookin = [[1, 23, 0], [1, 23, 1], [1, 25], [1, 26], [1, 27]]
    if foldertolookin != None :
        for index in range(len(foldertolookin)):
            date = datetime.strptime("2018-06-23", '%Y-%m-%d').strftime('%y-%m-%d')
            test = findstringinemail(searchstring, foldertolookin[index],"Mine",False, False, date)
            output = test.getvalue()
            print(output)'''



