
#import programs from python libraries
import xlwt
import pdfquery
import csv
import re

#user input is faster than reading the pdf page numbers in the document. Possible future improvment
pages = raw_input('Please enter the number of pages in the document:    ')

#convert user input to integer
pages = int(pages)

#Path to pdf file for PDFQuery access. PDFQuery is the program that pulls in the data from the pdf
pdf = pdfquery.PDFQuery('D:\New Storage\Coding\Python Projects\Iso Pull\Lack.pdf')

#load pdf to active for PDFQuery
pdf.load(range(0,1))

#cycle through page numbers
for pagenumber in range(0,pages):

    #create a string sub to avoid messiness in the pdf.pq page number callout
    pagesub = 'LTPage[page_index="%s"]' % pagenumber

    #find text in boxes. boxes are inches*72. Lower left corner of box to upper right
    #Also, keep in mind coordinates of BOM and Iso number may need tweaking due to coordinate find

    Item = pdf.pq(pagesub + ' :in_bbox("947.52,379.44,960.48,750.16")').text()
    QTY = pdf.pq(pagesub + ' :in_bbox("960.48,379.44,987.12,750.16")').text()
    Size = pdf.pq(pagesub + ' :in_bbox("987.12,379.44,1020.24,750.16")').text()
    Sch_Minwall = pdf.pq(pagesub + ' :in_bbox("1020.24,379.44,1059.12,750.16")').text()
    Description2 = pdf.pq(pagesub + ' :in_bbox("1059.12,379.44,1203.84,750.16")').text()
    Description2 = Description2[:len(Description2)/2]
    test = Description2


    #splits the text into list removing blank spaces
    Item = Item.split()
    QTY = QTY.split()
    Size = Size.split()
    Sch_Minwall = Sch_Minwall.split()


    #List of all delimiters in pipe codes. Needed for description
    delimiters = [' Pipe ,' , ' Hvy Hex Nut ,' , ' 45 Deg Elbow ,' , ' 90 Deg Elbow ,' , ' Flange Adapter ,' , ' MJ Adapter ,' , ' Tee ,' , ' Concentric Reducer ,' , 
' Backing Ring ,' , ' Std Blt & 2HvyHex Nt ,' , ' Non Metal Flat Gskt ,' , ' Blind Flg. ,' , ' Gate Valve Underground ,' , ' Wye ,' , ' Bell End Flange ,' , ' Insulating Gskt Set ,' , ' Flg. Adapter ,' , ' CAP ,' , 
' Reducing Tee ,' , ' Red. Male Adapter ,' , ' 90 Deg. Red. Elbow ,' , ' Flange Connecting Pc ,' , ' Bulkhead Union ,' , ' Female Connector ,' , ' 22.5 Deg Elbow ,' , ' Coupling ,' , ' Con. Swage ,' , ' Increaser C.I. ,' , 
' Reducer C.I. ,' , ' Reducing Y & 1/8Bend ,' , ' 45 Deg Red. Lateral ,' , ' Con. Reducer ,' , ' Check Valve Pistn Lft ,' , ' Check Valve Dual Pl LP ,' , ' Check Valve Swing ,' , ' Ball Valve LP FB ,' , ' Flexible Coupling ,' , ' Red. Coupling ,' , 
' 45 Deg LR Elbow ,' , ' 45 Deg SR Elbow ,' , ' 90 Deg LR Elbow ,' , ' Flgd Pipe ,' , ' Bell and Spigot Pipe ,' , ' 22.5 Deg. Elbow ,' , ' 45 Deg. Elbow ,' , ' 90 Deg. Elbow ,' , ' Equal Tee ,' , ' Red. Tee ,' , 
' Reducer Concentric ,' , ' Reducer Eccentric ,' , ' Reducing WYE ,' , ' Dielectric Union ,' , ' 11.25 Deg. Elbow ,' , ' Gate Valve LP ,' , ' Globe Valve LP ,' , ' 45 Deg Eq Lat Tee ,' , ' 45 Deg. 3D Elbow ,' , ' 45 Deg. Eq Lat Tee ,' , 
' 45 Deg. LR Elbow ,' , ' 45 Deg. SR Elbow ,' , ' 90 Deg. 3D Elbow ,' , ' 90 Deg. Elbow Asym ,' , ' 90 Deg. LR Elbow ,' , ' Adapter Female ,' , ' Adapter Flange Fem. ,' , ' Adapter Flange Male ,' , ' Adapter Male ,' , ' Eccentric Reducer ,' , 
' Cap Screw ,' , ' SO Flg. - Ring Type ,' , ' Nipple ,' , ' Cplg. ,' , ' Ecc. Swage ,' , ' Plug Hex Head ,' , ' Sp. Wound Gskt ,' , ' Thd. Flg. ,' , ' 1/4 Bend Long Sweep ,' , ' One-Eigth Bend ,' , 
' Sanitary Tee,' , ' Y & One Eighth Bend ,' , ' Y Branch,' , ' End Termination ,' , ' Std Blt ,' , ' Hex. Head Bush ,' , ' Elbolet ,' , ' Nipolet ,' , ' Threadolet ,' , ' 45 Deg Latrolet ,' , 
' 90 Deg. Elbow 3D ,' , ' 90 Deg. SR Elbow ,' , ' Coupling Cam Lock ,' , ' Eq. Tee ,' , ' Female Adapter ,' , ' Male Adapter ,' , ' Sockolet ,' , ' Weldolet ,' , ' Ecc. Reducer ,' , ' Lateral ,' , 
' 45 Deg Lateral ,' , ' Flatolet ,' , ' Male Connector ,' , ' Plug Round Head ,' , ' Reducing Insert ,' , ' WN Flg. ,' , ' SW Flg. ,' , ' WN Ori Flg 0.5"" Thd ,' , ' Gate Valve ,' , ' Globe Valve ,' , 
' Butterfly Valve Wafer Type ,' , ' Butterfly Valve Lug Type ,' , ' Butterfly Valve ,' , ' 45 Deg. Elbow 3D ,' , ' Union ,' , ' WN Ori Flg 0.75"" Thd ,' , ' End Plate ,' , ' Pipet ,' , ' Plug Square Head ,' , ' Transition Nipple ,' , 
' Rigid Coupling ,' , ' PressFit Coupling ,' , ' Weld Adapter ,' , ' Seal Ring ,' , ' Ball Vlv LP FB ,' , ' Ball Vlv FB ,' , ' Butterfly Vlv ,' , ' Adapter Nipple ,' , ' 11.25 Deg Elbow ,' , ' Cross ,' , 
' Transition Fitting ,' , ' Ball Valve ,' , ' Male Adapater ,' , ' Reducer Bushing ,' , ' Van Stone Flg. ,' , ' T-U Ball Check Valve ,' , ' Bfly Vlv Lug Type ,' , ' 45 Deg Red Lateral ,' , ' Branch Saddle ,' , ' SW Flg. - ,' , 
' Swage Nipple ,' , ' 90 Deg SR Elbow ,' , ' True Y ,' , ' 45 Deg Street Elbow ,' , ' 90 Deg Street Elbow ,' , ' Check Valve Tilting Disc ,' , ' WN OriFlg0.75""SW tap ,' , ' Check Valve Swing LP ,' , ' Gate Valve SP ,' , ' Globe Valve SP ,' , 
' Check Valve Swing SP ,' , ' 45 Deg Eq Lat Custom ,' , ' 45Deg Red Lat Custom ,' , ' Fitting Reducer ,' , ' Machine Bolt ,' , ' Check Valve Y-Pattern Swing Disc ,' , ' Check Valve Guided Piston ,' , ' Dieletric Union ,' , ' Check Valve Spring ,' , ' 90 Deg Elbow,' , 
' Tube ,' , ' Weld Connector ,' , ' Bulkhead Redng Union ,' , ' Reducing Union ,' , ' 45 Deg. Y ,' , ' Lateral Wye ,' , ' Reducing 45 Deg. Y ,' , ' 30 Deg. Elbow ,' , ' 5.625 Deg. Elbow ,' , ' 90 Deg. Y ,' , 
' Eq. Cross ,' , ' 45 Dg Mitr 2 Ct 2.5D ,' , ' 90 Dg 5pc Elbow 1.0D ,' , ' 90 Dg 5pc Elbow 1.5D ,' , ' 90 Dg 5pc Elbow 2.5D ,' , ' Butterfly Valve DblFlg Short Pt ,'
                ]

    #used to cycle through list appends. May be able to make this code cleaner in future, but
    #the code is suitable for now since it is not very intensive
    appendcounter = 0

    #cycle through the values in the delimiter list to split the string into a list for excel pasting
    for value in delimiters:
        if value in Description2:
            
            #on first iteration
            if appendcounter == 0:
                front, mid, back = Description2.partition(value)
                Description = [front]
                Description2 = mid+back
                appendcounter = 1
                continue
            else:
                front, mid, back = Description2.partition(value)

            #keep the end of the split string for the next iteration. Add the front of the split string
            #to the primary list
            Description2 = mid+back
            Description.append(front)
                
            appendcounter += 1

    #if the last delimiter, add the end of the split string to the primary list
    Description.append(Description2)
    #reset appendcounter just in case
    appendcounter = 0

    #first time through, create the workbook
    if pagenumber == 0:
        overallrowcounter = 1
        #open excel book
        testbook = xlwt.Workbook()
        sh = testbook.add_sheet("Sheet1")

        #create column headers
        sh.write(0,0,'Iso')
        sh.write(0,1,'Item')
        sh.write(0,2,'QTY')
        sh.write(0,3,'Size')
        sh.write(0,4,'Schedule/Min Wall')
        sh.write(0,5,'Description')
 
    #Fill in rows for each column from individuall lists for BOM
    for i in range(0,len(Item)/2):
        sh.write(overallrowcounter,0,Iso)
        sh.write(overallrowcounter,1,Item[i])
        sh.write(overallrowcounter,2,QTY[i])
        sh.write(overallrowcounter,3,Size[i])
        sh.write(overallrowcounter,4,Sch_Minwall[i])
        sh.write(overallrowcounter,5,Description[i])
        #keeps track of current row no matter number of pages
        overallrowcounter += 1

    #save the changes in the workbook
    testbook.save("testbook.xls")
