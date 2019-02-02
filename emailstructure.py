
from win32com.client import Dispatch

def getrootlist():
    outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
    for i in range(50):
        try:
            foldernames = outlook.GetDefaultFolder(i).name

            print(i, foldernames)
        except:
            pass

def getemailstructure(desiredfolder):
    indexed = []
    outlook = Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    counter = namespace.Folders.count
    for x in range (1,counter+1):
        print(x)
        test0 = namespace.Folders.Item(x)
        print(test0)
        if not any(keyfolders in str(test0) for keyfolders in ("Mt. ", "Public")):
            counter1= namespace.Folders.Item(x).Folders.Count
            print(counter1)
            
            
            for z in range (0,counter1):
                test = namespace.Folders.Item(x).Folders[z]
                print("in: ", test0,"subfolder: ", test)
                counter2= namespace.Folders.Item(x).Folders[z].Folders.Count
                if any((keyfolders in str(test) and len(keyfolders) == len(str(test))) for keyfolders in (desiredfolder)):
                    print ("got the desired folder at index: ",x, z)
                    indexed.append([x, z])
                
                
                for k in range (0,counter2):
                    test1 = namespace.Folders.Item(x).Folders[z].Folders[k]
                    print("in: ", test0 ,"subfolder: ", test,"subsubfolder: ", test1)
                    counter3= namespace.Folders.Item(x).Folders[z].Folders[k].Folders.Count
                    if any((keyfolders in str(test1) and len(keyfolders) == len(str(test1))) for keyfolders in (desiredfolder)):
                        print ("got the desired folder at index: ",x, z, k)
                        indexed.append([x, z, k])
                        
                        for j in range (0,counter3):
                            test2 = namespace.Folders.Item(x).Folders[z].Folders[k].Folders[j]
                            print("in: ", test0 ,"subfolder: ", test,"subsubfolder: ", test1, "4th level: ", test2)
                            counter4= namespace.Folders.Item(x).Folders[z].Folders[k].Folders[j].Folders.Count
                            if any((keyfolders in str(test2) and len(keyfolders) == len(str(test2))) for keyfolders in (desiredfolder)):
                                    print ("got the desired folder at index: ",x, z, k, j)
                                    indexed.append([x, z, k, j])
                                    
                            for m in range (0,counter4):
                                test3 = namespace.Folders.Item(x).Folders[z].Folders[k].Folders[j].Folders[m]
                                print("in: ", test0 ,"subfolder: ", test,"subsubfolder: ", test1, "4th level: ", test2, "5th level: ", test3)
                                counter4= namespace.Folders.Item(x).Folders[z].Folders[k].Folders[j].Folders[m].Folders.Count
                                if any((keyfolders in str(test3) and len(keyfolders) == len(str(test3))) for keyfolders in (desiredfolder)):
                                    print ("got the desired folder at index: ",x, z, k, j, m)
                                    indexed.append([x, z, k, j, m])

    return indexed

def saveresults(filename, values):
    f = open(filename + ".txt", "w")
    f.write(str(values))
    f.close()
    
    
if __name__ == '__main__':
    testt=getemailstructure(["CAQ","TEST1", "TESTEST"])
    print(testt)
    saveresults('valuefile' ,testt)

    print("done")