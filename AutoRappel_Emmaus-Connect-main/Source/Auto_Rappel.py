# coding: utf-8

from ipaddress import IPv4Address  # for your IP address
from pyairmore.request import AirmoreSession  # to create an AirmoreSession
from pyairmore.services.messaging import MessagingService  # to send messages
from openpyxl import load_workbook
from tkinter import *
import tkinter.font as font
from tkinter.filedialog import *
from functools import partial
import time

def quit(): #fonction quitter pour fermer l'app
    root.destroy()

def connexion(entree): #function de connexion au service airmore et fonction secondaire

    def Clear(): #efface tout les widgets présent sur la page ne laissant que le fond d'écran et le menu déroulant
        for widget in fenetre.winfo_children():
            widget.destroy()

    def myfunction(event): #foncion event de la barre déroulant pour l'affichage des différents contacts
        canvas3.configure(scrollregion=canvas3.bbox("all"),width=800,height=650)

    def menu():# definition du menu : choix permCo parcours ou rdv
        Clear()
        bt1=Button(fenetre, text="Permanence connecté", command= Permco, bg  = "darkorange2")
        bt1['font'] = font.Font(size=20)
        bt1.pack(pady = 20)
        bt2=Button(fenetre, text="Parcours d'initiation", command= parcours, bg  = "darkgoldenrod1")
        bt2['font'] = font.Font(size=20)
        bt2.pack(pady = 20)
        bt3=Button(fenetre, text="Rendez-vous", command= RDV, bg  = "cadetblue2")
        bt3['font'] = font.Font(size=20)
        bt3.pack(pady = 20)
        bt4=Button(fenetre, text="PUB", command= RDV, bg  = "crimson")
        bt4['font'] = font.Font(size=20)
        bt4.pack(pady = 20)

    def Permco(): #fonction appeller par le bouton permco du menu : gestion de tout les contacts permco
        label = Label(fenetre, text="Selectionné un fichier de contact (.xlsx)", bg="cyan3")
        label['font'] = font.Font(size=20)
        label.place(x=550, y=400, anchor="center")
        fenetre.update()
        filepath = askopenfilename(title="Ouvrir le fichier de rappelle",filetypes=[('xlsx file','.xlsx'),('all files','.*')])
        print(filepath)
        wb = load_workbook(filename = filepath)
        sheet_ranges = wb['Rappels pour participants lende']
        print(sheet_ranges['H18'].value)
        label.destroy()
        root.geometry('1500x850')
        fenetre.width = 1500

        def myfunction(event): #foncion event de la barre déroulant pour l'affichage des différents contacts
            canvas3.configure(scrollregion=canvas3.bbox("all"),width=1450,height=650)

        def checknum(row): #check de la validité du numéro de téléphone
            nb = 0
            num = str(sheet_ranges['I'+ str(row_1)].value)
            for x in num :
                if x in "0123456789":
                    if nb == 0 and x != "0" :
                        return(False)
                    else :
                        nb += 1
                if nb > 10 :
                    return(False)
            if nb == 10 :
                return(True)
            else :
                return(False)

        def check(row): #basculement de létat du contact lorsqu'il est checker pour faire suivre l'envoie des sms
            print("execute")
            if sheet_ranges['Z'+ str(row)].value == None:
                sheet_ranges['Z'+ str(row)].value = True
            elif sheet_ranges['Z'+ str(row)].value == True:
                sheet_ranges['Z'+ str(row)].value = False
            else :
                sheet_ranges['Z'+ str(row)].value = False

        def sendSMS(): #envoi des sms aux contacts checker
            row_1 = 1
            row_2 = 0
            while sheet_ranges['C'+ str(row_1)].value != "Permanence Connectée" and sheet_ranges['B'+ str(row_1)].value != "Total":
                row_1 += 1

            print(sheet_ranges['B'+ str(row_1)].value)
            print(row_1)
            if sheet_ranges['B'+ str(row_1)].value == "Total" :
                print("pas de permco")
                Clear()
                label = Label(fenetre, text="Pas de permco", bg="cyan3")
                label['font'] = font.Font(size=20)
                label.place(x=750, y=400, anchor="center")
                fenetre.update()
                time.sleep(1)
                menu()
                return()

            else :
                row_2 = row_1 +1
                print(sheet_ranges['C'+ str(row_2)].value)
                while sheet_ranges['C'+ str(row_2)].value == None and sheet_ranges['B'+ str(row_2)].value != "Total":
                    print("ok")
                    row_2 += 1
            print(row_2)
            while row_1 != row_2 :
                print(sheet_ranges['Z'+ str(row_1)].value)
                if sheet_ranges['Z'+ str(row_1)].value == True :
                    service.send_message(str(sheet_ranges['I'+ str(row_1)].value), "Bonjour, L'équipe d'Emmaus Connect vous rappelle votre Permanence Connectée le "+str( sheet_ranges['F'+ str(row_1)].value)[:-5]+ " à "+ str( sheet_ranges['F'+ str(row_1)].value)[-5:] +" au 104 route de Vienne, 69008 Lyon. \n\n Merci de nous prévenir en cas d'abscence en appelant le 01 80 05 98 80. \n\n L'équipe Emmaus Connect")

                row_1 += 1
            menu()

        def valider(row, entre): #validation des changement du numéro de tel
            print(entre.get())
            sheet_ranges['I'+ str(row)] =  entre.get()
            wb.save(filename = filepath)
            Permco()
        def modif(row): #page de modif du numéro de téléphone lorsque incorrect
            Clear()
            label = Label(fenetre, text = "Modifier Numéro :", bg ="darkgoldenrod1")
            label['font'] = font.Font(size=20)
            label.place(x=10, y=10)
            vallue = StringVar()
            vallue.set(str(sheet_ranges['I'+ str(row)].value))
            entree = Entry(fenetre, textvariable=vallue,bd = 10, width=30)
            entree['font'] = font.Font(size=20)
            entree.place(x=10, y=100)
            bt1=Button(fenetre, text="ANNULER", command= Permco, bg = "cyan3")
            bt1['font'] = font.Font(size=20)
            bt1.place(x=10, y=180)
            bt2=Button(fenetre, text="VALIDER", command= partial(valider, row, entree), bg = "cyan3")
            bt2['font'] = font.Font(size=20)
            bt2.place(x=250, y=180)




        Clear()
        myframe=Frame(fenetre,relief=GROOVE,width=1480,height=700,bd=1, bg= "cyan3")
        myframe.place(x=10,y=50)

        canvas3=Canvas(myframe, bg ='cyan3')
        frame=Frame(canvas3, bg ='cyan3')
        myscrollbar=Scrollbar(myframe,orient="vertical",command=canvas3.yview)
        canvas3.configure(yscrollcommand=myscrollbar.set)

        myscrollbar.pack(side="right",fill="y")
        canvas3.pack(side="left")
        canvas3.create_window((0,0),window=frame,anchor='nw')
        frame.bind("<Configure>",myfunction)
        row_1 = 1
        row_2 = 0
        i = 1

        Label(frame,text="Nom", font = ("Comic Sans MS", 20, "bold"), bg ="darkgoldenrod1").grid(row=0,column=1, padx = 10)
        Label(frame,text="Numéro", font = ("Comic Sans MS", 20, "bold"), bg ="darkgoldenrod1").grid(row=0,column=2, padx = 10)
        Label(frame,text="Date/Heure", font = ("Comic Sans MS", 20, "bold"), bg ="darkgoldenrod1").grid(row=0,column=4, padx = 10)
        Label(frame,text="Statut", font = ("Comic Sans MS", 20, "bold"), bg ="darkgoldenrod1").grid(row=0,column=5, padx = 10)
        while sheet_ranges['C'+ str(row_1)].value != "Permanence Connectée" and sheet_ranges['B'+ str(row_1)].value != "Total":
            row_1 += 1

        print(sheet_ranges['B'+ str(row_1)].value)
        print(row_1)
        if sheet_ranges['B'+ str(row_1)].value == "Total" :
            print("pas de permco")
            Clear()
            label = Label(fenetre, text="Pas de permco", bg="cyan3")
            label['font'] = font.Font(size=20)
            label.place(x=750, y=400, anchor="center")
            fenetre.update()
            time.sleep(1)
            menu()
            return()

        else :
            row_2 = row_1 +1
            print(sheet_ranges['C'+ str(row_2)].value)
            while sheet_ranges['C'+ str(row_2)].value == None and sheet_ranges['B'+ str(row_2)].value != "Total":
                print("ok")
                row_2 += 1
        print(row_2)

        while row_1 != row_2 :

            Label(frame,text= str(i), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3", highlightbackground="cyan3").grid(row=i,column=0, padx = 10)
            Label(frame,text=str(sheet_ranges['H'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3").grid(row=i,column=1, padx = 10)

            Label(frame,text=str(sheet_ranges['F'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3").grid(row=i,column=4, padx = 10)
            Label(frame,text=str(sheet_ranges['K'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3").grid(row=i,column=5, padx = 10)
            if checknum(row_1) == True :
                Label(frame,text=str(sheet_ranges['I'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3").grid(row=i,column=2, padx = 10)
                if sheet_ranges['Z'+ str(row_1)].value == True:
                   btn= Checkbutton(frame, text='Rappel',  font = ("Comic Sans MS", 15, "bold"), command=partial(check , row_1), bg ='cyan3')
                   btn.select()
                   btn.grid(row=i,column=6, padx = 10)
                else :
                    Checkbutton(frame, text='Rappel',  font = ("Comic Sans MS", 15, "bold"), command=partial(check , row_1), bg ='cyan3').grid(row=i,column=6, padx = 10)
            else :
                Label(frame,text=str(sheet_ranges['I'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg="red").grid(row=i,column=2, padx = 10)
                photto = PhotoImage(file='./img/pencils.gif')
                imgButton = Button(frame, image=photto, command=partial(modif, row_1))
                imgButton.image = photto
                imgButton.grid(row=i,column=3)
            i +=1
            row_1 += 1

        bt=Button(fenetre, text="validé", command = sendSMS)
        bt['font'] = font.Font(size=20)
        bt.place(x=750, y=750, anchor="center")

        ################################# Fin Fonction PERMCO ##############################################################

    def Pub():#fonction appeller par le bouton Pub du menu : gestion de tout les contacts Pub
        label = Label(fenetre, text="Selectionné un fichier de contact (.xlsx)", bg="cyan3")
        label['font'] = font.Font(size=20)
        label.place(x=550, y=400, anchor="center")
        fenetre.update()
        filepath = askopenfilename(title="Ouvrir le fichier de rappelle",filetypes=[('xlsx file','.xlsx'),('all files','.*')])
        print(filepath)
        wb = load_workbook(filename = filepath)
        sheet_ranges = wb['']
        print(sheet_ranges['H18'].value)
        label.destroy()
        root.geometry('1500x850')
        fenetre.width = 1500

        def myfunction(event): #foncion event de la barre déroulant pour l'affichage des différents contacts
            canvas3.configure(scrollregion=canvas3.bbox("all"),width=1450,height=650)

        def checknum(row): #check de la validité du numéro de téléphone
            nb = 0
            num = str(sheet_ranges['I'+ str(row_1)].value)
            for x in num :
                if x in "0123456789":
                    if nb == 0 and x != "0" :
                        return(False)
                    else :
                        nb += 1
                if nb > 10 :
                    return(False)
            if nb == 10 :
                return(True)
            else :
                return(False)

        def check(row): #basculement de létat du contact lorsqu'il est checker pour faire suivre l'envoie des sms
            print("execute")
            if sheet_ranges['Z'+ str(row)].value == None:
                sheet_ranges['Z'+ str(row)].value = True
            elif sheet_ranges['Z'+ str(row)].value == True:
                sheet_ranges['Z'+ str(row)].value = False
            else :
                sheet_ranges['Z'+ str(row)].value = False

        def sendSMS(): #envoi des sms aux contacts checker
            row_1 = 1
            row_2 = 0
            while sheet_ranges['C'+ str(row_1)].value != "Permanence Connectée" and sheet_ranges['B'+ str(row_1)].value != "Total":
                row_1 += 1

            print(sheet_ranges['B'+ str(row_1)].value)
            print(row_1)
            if sheet_ranges['B'+ str(row_1)].value == "Total" :
                print("pas de Pub")
                Clear()
                label = Label(fenetre, text="Pas de Pub", bg="cyan3")
                label['font'] = font.Font(size=20)
                label.place(x=750, y=400, anchor="center")
                fenetre.update()
                time.sleep(1)
                menu()
                return()

            else :
                row_2 = row_1 +1
                print(sheet_ranges['C'+ str(row_2)].value)
                while sheet_ranges['C'+ str(row_2)].value == None and sheet_ranges['B'+ str(row_2)].value != "Total":
                    print("ok")
                    row_2 += 1
            print(row_2)
            while row_1 != row_2 :
                print(sheet_ranges['Z'+ str(row_1)].value)
                if sheet_ranges['Z'+ str(row_1)].value == True :
                    service.send_message(str(sheet_ranges['I'+ str(row_1)].value), "Bonjour, L'équipe d'Emmaus Connect vous rappelle votre Permanence Connectée le "+str( sheet_ranges['F'+ str(row_1)].value)[:-5]+ " à "+ str( sheet_ranges['F'+ str(row_1)].value)[-5:] +" au 104 route de Vienne, 69008 Lyon. \n\n Merci de nous prévenir en cas d'abscence en appelant le 01 80 05 98 80. \n\n L'équipe Emmaus Connect")

                row_1 += 1
            menu()

        def valider(row, entre): #validation des changement du numéro de tel
            print(entre.get())
            sheet_ranges['I'+ str(row)] =  entre.get()
            wb.save(filename = filepath)
            Pub()
        def modif(row): #page de modif du numéro de téléphone lorsque incorrect
            Clear()
            label = Label(fenetre, text = "Modifier Numéro :", bg ="darkgoldenrod1")
            label['font'] = font.Font(size=20)
            label.place(x=10, y=10)
            vallue = StringVar()
            vallue.set(str(sheet_ranges['I'+ str(row)].value))
            entree = Entry(fenetre, textvariable=vallue,bd = 10, width=30)
            entree['font'] = font.Font(size=20)
            entree.place(x=10, y=100)
            bt1=Button(fenetre, text="ANNULER", command= Pub, bg = "cyan3")
            bt1['font'] = font.Font(size=20)
            bt1.place(x=10, y=180)
            bt2=Button(fenetre, text="VALIDER", command= partial(valider, row, entree), bg = "cyan3")
            bt2['font'] = font.Font(size=20)
            bt2.place(x=250, y=180)




        Clear()
        myframe=Frame(fenetre,relief=GROOVE,width=1480,height=700,bd=1, bg= "cyan3")
        myframe.place(x=10,y=50)

        canvas3=Canvas(myframe, bg ='cyan3')
        frame=Frame(canvas3, bg ='cyan3')
        myscrollbar=Scrollbar(myframe,orient="vertical",command=canvas3.yview)
        canvas3.configure(yscrollcommand=myscrollbar.set)

        myscrollbar.pack(side="right",fill="y")
        canvas3.pack(side="left")
        canvas3.create_window((0,0),window=frame,anchor='nw')
        frame.bind("<Configure>",myfunction)
        row_1 = 1
        row_2 = 0
        i = 1

        Label(frame,text="Nom", font = ("Comic Sans MS", 20, "bold"), bg ="darkgoldenrod1").grid(row=0,column=1, padx = 10)
        Label(frame,text="Numéro", font = ("Comic Sans MS", 20, "bold"), bg ="darkgoldenrod1").grid(row=0,column=2, padx = 10)
        Label(frame,text="Date/Heure", font = ("Comic Sans MS", 20, "bold"), bg ="darkgoldenrod1").grid(row=0,column=4, padx = 10)
        Label(frame,text="Statut", font = ("Comic Sans MS", 20, "bold"), bg ="darkgoldenrod1").grid(row=0,column=5, padx = 10)
        while sheet_ranges['C'+ str(row_1)].value != "Permanence Connectée" and sheet_ranges['B'+ str(row_1)].value != "Total":
            row_1 += 1

        print(sheet_ranges['B'+ str(row_1)].value)
        print(row_1)
        if sheet_ranges['B'+ str(row_1)].value == "Total" :
            print("pas de Pub")
            Clear()
            label = Label(fenetre, text="Pas de Pub", bg="cyan3")
            label['font'] = font.Font(size=20)
            label.place(x=750, y=400, anchor="center")
            fenetre.update()
            time.sleep(1)
            menu()
            return()

        else :
            row_2 = row_1 +1
            print(sheet_ranges['C'+ str(row_2)].value)
            while sheet_ranges['C'+ str(row_2)].value == None and sheet_ranges['B'+ str(row_2)].value != "Total":
                print("ok")
                row_2 += 1
        print(row_2)

        while row_1 != row_2 :

            Label(frame,text= str(i), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3", highlightbackground="cyan3").grid(row=i,column=0, padx = 10)
            Label(frame,text=str(sheet_ranges['H'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3").grid(row=i,column=1, padx = 10)

            Label(frame,text=str(sheet_ranges['F'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3").grid(row=i,column=4, padx = 10)
            Label(frame,text=str(sheet_ranges['K'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3").grid(row=i,column=5, padx = 10)
            if checknum(row_1) == True :
                Label(frame,text=str(sheet_ranges['I'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3").grid(row=i,column=2, padx = 10)
                if sheet_ranges['Z'+ str(row_1)].value == True:
                    Checkbutton(frame, text='Rappel',  font = ("Comic Sans MS", 15, "bold"), command=partial(check , row_1), bg ='cyan3').grid(row=i,column=6, padx = 10).select()
                else :
                    Checkbutton(frame, text='Rappel',  font = ("Comic Sans MS", 15, "bold"), command=partial(check , row_1), bg ='cyan3').grid(row=i,column=6, padx = 10)
            else :
                Label(frame,text=str(sheet_ranges['I'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg="red").grid(row=i,column=2, padx = 10)
                photto = PhotoImage(file='./img/pencils.gif')
                imgButton = Button(frame, image=photto, command=partial(modif, row_1))
                imgButton.image = photto
                imgButton.grid(row=i,column=3)
            i +=1
            row_1 += 1

        bt=Button(fenetre, text="validé", command = sendSMS)
        bt['font'] = font.Font(size=20)
        bt.place(x=750, y=750, anchor="center")

        ################################# Fin Fonction Pub ##############################################################

    def RDV(): #fonction appeller par le bouton RDV du menu : gestion de tout les contacts RDV
        label = Label(fenetre, text="Selectionné un fichier de contact (.xlsx)", bg="cyan3")
        label['font'] = font.Font(size=20)
        label.place(x=550, y=400, anchor="center")
        fenetre.update()
        filepath = askopenfilename(title="Ouvrir le fichier de rappelle",filetypes=[('xlsx file','.xlsx'),('all files','.*')])
        print(filepath)
        wb = load_workbook(filename = filepath)
        sheet_ranges = wb['Rappels pour participants lende']
        print(sheet_ranges['H18'].value)
        label.destroy()
        root.geometry('1500x850')
        fenetre.width = 1500

        def myfunction(event): #foncion event de la barre déroulant pour l'affichage des différents contacts
            canvas3.configure(scrollregion=canvas3.bbox("all"),width=1450,height=650)

        def checknum(row): #check de la validité du numéro de téléphone
            nb = 0
            num = str(sheet_ranges['I'+ str(row_1)].value)
            for x in num :
                if x in "0123456789":
                    if nb == 0 and x != "0" :
                        return(False)
                    else :
                        nb += 1
                if nb > 10 :
                    return(False)
            if nb == 10 :
                return(True)
            else :
                return(False)

        def check(row): #basculement de létat du contact lorsqu'il est checker pour faire suivre l'envoie des sms
            print("execute")
            if sheet_ranges['Z'+ str(row)].value == None:
                sheet_ranges['Z'+ str(row)].value = True
            elif sheet_ranges['Z'+ str(row)].value == True:
                sheet_ranges['Z'+ str(row)].value = False
            else :
                sheet_ranges['Z'+ str(row)].value = False

        def sendSMS(): #envoi des sms aux contacts checker
            row_1 = 1
            row_2 = 0
            while sheet_ranges['C'+ str(row_1)].value != "RDV bénéficiaire" and sheet_ranges['B'+ str(row_1)].value != "Total":
                row_1 += 1

            print(sheet_ranges['B'+ str(row_1)].value)
            print(row_1)
            if sheet_ranges['B'+ str(row_1)].value == "Total" :
                print("pas de RDV")
                Clear()
                label = Label(fenetre, text="Pas de RDV", bg="cyan3")
                label['font'] = font.Font(size=20)
                label.place(x=750, y=400, anchor="center")
                fenetre.update()
                time.sleep(1)
                menu()
                return()

            else :
                row_2 = row_1 +1
                print(sheet_ranges['C'+ str(row_2)].value)
                while sheet_ranges['C'+ str(row_2)].value == None and sheet_ranges['B'+ str(row_2)].value != "Total":
                    print("ok")
                    row_2 += 1
            print(row_2)
            while row_1 != row_2 :
                print(sheet_ranges['Z'+ str(row_1)].value)
                if sheet_ranges['Z'+ str(row_1)].value == True :
                    service.send_message(str(sheet_ranges['I'+ str(row_1)].value), "Bonjour, L'équipe d'Emmaus Connect vous rappelle votre rendez-vous de "+str(sheet_ranges['E'+ str(row_1)].value) + " le "+str( sheet_ranges['F'+ str(row_1)].value)[:-5]+ " à "+ str( sheet_ranges['F'+ str(row_1)].value)[-5:] +" au 104 route de Vienne, 69008 Lyon. \n\n Merci de nous prévenir en cas d'abscence en appelant le 01 80 05 98 80. \n\n L'équipe Emmaus Connect")

                row_1 += 1
            menu()

        def valider(row, entre): #validation des changement du numéro de tel
            print(entre.get())
            sheet_ranges['I'+ str(row)] =  entre.get()
            wb.save(filename = filepath)
            RDV()
        def modif(row): #page de modif du numéro de téléphone lorsque incorrect
            Clear()
            label = Label(fenetre, text = "Modifier Numéro :", bg ="darkgoldenrod1")
            label['font'] = font.Font(size=20)
            label.place(x=10, y=10)
            vallue = StringVar()
            vallue.set(str(sheet_ranges['I'+ str(row)].value))
            entree = Entry(fenetre, textvariable=vallue,bd = 10, width=30)
            entree['font'] = font.Font(size=20)
            entree.place(x=10, y=100)
            bt1=Button(fenetre, text="ANNULER", command= RDV, bg = "cyan3")
            bt1['font'] = font.Font(size=20)
            bt1.place(x=10, y=180)
            bt2=Button(fenetre, text="VALIDER", command= partial(valider, row, entree), bg = "cyan3")
            bt2['font'] = font.Font(size=20)
            bt2.place(x=250, y=180)




        Clear()
        myframe=Frame(fenetre,relief=GROOVE,width=1480,height=700,bd=1, bg= "cyan3")
        myframe.place(x=10,y=50)

        canvas3=Canvas(myframe, bg ='cyan3')
        frame=Frame(canvas3, bg ='cyan3')
        myscrollbar=Scrollbar(myframe,orient="vertical",command=canvas3.yview)
        canvas3.configure(yscrollcommand=myscrollbar.set)

        myscrollbar.pack(side="right",fill="y")
        canvas3.pack(side="left")
        canvas3.create_window((0,0),window=frame,anchor='nw')
        frame.bind("<Configure>",myfunction)
        row_1 = 1
        row_2 = 0
        i = 1

        Label(frame,text="Nom", font = ("Comic Sans MS", 20, "bold"), bg ="darkgoldenrod1").grid(row=0,column=1, padx = 10)
        Label(frame,text="Numéro", font = ("Comic Sans MS", 20, "bold"), bg ="darkgoldenrod1").grid(row=0,column=2, padx = 10)
        Label(frame,text="Date/Heure", font = ("Comic Sans MS", 20, "bold"), bg ="darkgoldenrod1").grid(row=0,column=4, padx = 10)
        Label(frame,text="Statut", font = ("Comic Sans MS", 20, "bold"), bg ="darkgoldenrod1").grid(row=0,column=5, padx = 10)
        Label(frame,text="Type RDV", font = ("Comic Sans MS", 20, "bold"), bg ="darkgoldenrod1").grid(row=0,column=6, padx = 10)
        while sheet_ranges['C'+ str(row_1)].value != "RDV bénéficiaire" and sheet_ranges['B'+ str(row_1)].value != "Total":
            row_1 += 1

        print(sheet_ranges['B'+ str(row_1)].value)
        print(row_1)
        if sheet_ranges['B'+ str(row_1)].value == "Total" :
            print("pas de RDV")
            Clear()
            label = Label(fenetre, text="Pas de RDV", bg="cyan3")
            label['font'] = font.Font(size=20)
            label.place(x=750, y=400, anchor="center")
            fenetre.update()
            time.sleep(1)
            menu()
            return()

        else :
            row_2 = row_1 +1
            print(sheet_ranges['C'+ str(row_2)].value)
            while sheet_ranges['C'+ str(row_2)].value == None and sheet_ranges['B'+ str(row_2)].value != "Total":
                print("ok")
                row_2 += 1
        print(row_2)

        while row_1 != row_2 :

            Label(frame,text= str(i), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3", highlightbackground="cyan3").grid(row=i,column=0, padx = 10)
            Label(frame,text=str(sheet_ranges['H'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3").grid(row=i,column=1, padx = 10)

            Label(frame,text=str(sheet_ranges['F'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3").grid(row=i,column=4, padx = 10)
            Label(frame,text=str(sheet_ranges['K'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3").grid(row=i,column=5, padx = 10)
            Label(frame,text=str(sheet_ranges['E'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3").grid(row=i,column=6, padx = 10)
            if checknum(row_1) == True :
                Label(frame,text=str(sheet_ranges['I'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3").grid(row=i,column=2, padx = 10)
                Checkbutton(frame, text='Rappel',  font = ("Comic Sans MS", 15, "bold"), command=partial(check , row_1), bg ='cyan3').grid(row=i,column=7, padx = 10)
            else :
                Label(frame,text=str(sheet_ranges['I'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg="red").grid(row=i,column=2, padx = 10)
                photto = PhotoImage(file='./img/pencils.gif')
                imgButton = Button(frame, image=photto, command=partial(modif, row_1))
                imgButton.image = photto
                imgButton.grid(row=i,column=3)
            i +=1
            row_1 += 1

        bt=Button(fenetre, text="validé", command = sendSMS)
        bt['font'] = font.Font(size=20)
        bt.place(x=750, y=750, anchor="center")

        ################################# Fin Fonction RDV ##############################################################

    def parcours(): #fonction appeller par le bouton parcours du menu : gestion de tout les contacts parcours
        label = Label(fenetre, text="Selectionné un fichier de contact (.xlsx)", bg="cyan3")
        label['font'] = font.Font(size=20)
        label.place(x=550, y=400, anchor="center")
        fenetre.update()
        filepath = askopenfilename(title="Ouvrir le fichier de rappelle",filetypes=[('xlsx file','.xlsx'),('all files','.*')])
        print(filepath)
        wb = load_workbook(filename = filepath)
        sheet_ranges = wb['Rappels pour participants lende']
        print(sheet_ranges['H18'].value)
        label.destroy()
        root.geometry('1500x850')
        fenetre.width = 1500

        def myfunction(event): #foncion event de la barre déroulant pour l'affichage des différents contacts
            canvas3.configure(scrollregion=canvas3.bbox("all"),width=1450,height=650)

        def checknum(row): #check de la validité du numéro de téléphone
            nb = 0
            num = str(sheet_ranges['I'+ str(row_1)].value)
            for x in num :
                if x in "0123456789":
                    if nb == 0 and x != "0" :
                        return(False)
                    else :
                        nb += 1
                if nb > 10 :
                    return(False)
            if nb == 10 :
                return(True)
            else :
                return(False)

        def check(row): #basculement de létat du contact lorsqu'il est checker pour faire suivre l'envoie des sms
            print("execute")
            if sheet_ranges['Z'+ str(row)].value == None:
                sheet_ranges['Z'+ str(row)].value = True
            elif sheet_ranges['Z'+ str(row)].value == True:
                sheet_ranges['Z'+ str(row)].value = False
            else :
                sheet_ranges['Z'+ str(row)].value = False

        def sendSMS(): #envoi des sms aux contacts checker
            row_1 = 1
            row_2 = 0
            while ("Parcours d'initiation" not in str(sheet_ranges['C'+ str(row_1)].value)) and (str(sheet_ranges['B'+ str(row_1)].value) != "Total") :
                row_1 += 1

            print(sheet_ranges['B'+ str(row_1)].value)
            print(row_1)
            if sheet_ranges['B'+ str(row_1)].value == "Total" :
                print("pas de parcours")
                Clear()
                label = Label(fenetre, text="Pas de parcours", bg="cyan3")
                label['font'] = font.Font(size=20)
                label.place(x=750, y=400, anchor="center")
                fenetre.update()
                time.sleep(1)
                menu()
                return()

            else :
                row_2 = row_1 +1
                print(sheet_ranges['C'+ str(row_2)].value)
                while sheet_ranges['C'+ str(row_2)].value == None and sheet_ranges['B'+ str(row_2)].value != "Total":
                    print("ok")
                    row_2 += 1
            print(row_2)
            while row_1 != row_2 :
                print(sheet_ranges['Z'+ str(row_1)].value)
                if sheet_ranges['Z'+ str(row_1)].value == True :
                    service.send_message(str(sheet_ranges['I'+ str(row_1)].value), "Bonjour, L'équipe d'Emmaus Connect vous rappelle votre Parcours d'Initiation le "+str( sheet_ranges['F'+ str(row_1)].value)[:-5]+ " à "+ str( sheet_ranges['F'+ str(row_1)].value)[-5:] +" au 104 route de Vienne, 69008 Lyon. \n\n Merci de nous prévenir en cas d'abscence en appelant le 01 80 05 98 80. \n\n L'équipe Emmaus Connect")

                row_1 += 1
            menu()

        def valider(row, entre): #validation des changement du numéro de tel
            print(entre.get())
            sheet_ranges['I'+ str(row)] =  entre.get()
            wb.save(filename = filepath)
            parcours()
        def modif(row): #page de modif du numéro de téléphone lorsque incorrect
            Clear()
            label = Label(fenetre, text = "Modifier Numéro :", bg ="darkgoldenrod1")
            label['font'] = font.Font(size=20)
            label.place(x=10, y=10)
            vallue = StringVar()
            vallue.set(str(sheet_ranges['I'+ str(row)].value))
            entree = Entry(fenetre, textvariable=vallue,bd = 10, width=30)
            entree['font'] = font.Font(size=20)
            entree.place(x=10, y=100)
            bt1=Button(fenetre, text="ANNULER", command= parcours, bg = "cyan3")
            bt1['font'] = font.Font(size=20)
            bt1.place(x=10, y=180)
            bt2=Button(fenetre, text="VALIDER", command= partial(valider, row, entree), bg = "cyan3")
            bt2['font'] = font.Font(size=20)
            bt2.place(x=250, y=180)




        Clear()
        myframe=Frame(fenetre,relief=GROOVE,width=1480,height=700,bd=1, bg= "cyan3")
        myframe.place(x=10,y=50)

        canvas3=Canvas(myframe, bg ='cyan3')
        frame=Frame(canvas3, bg ='cyan3')
        myscrollbar=Scrollbar(myframe,orient="vertical",command=canvas3.yview)
        canvas3.configure(yscrollcommand=myscrollbar.set)

        myscrollbar.pack(side="right",fill="y")
        canvas3.pack(side="left")
        canvas3.create_window((0,0),window=frame,anchor='nw')
        frame.bind("<Configure>",myfunction)
        row_1 = 1
        row_2 = 0
        i = 1

        Label(frame,text="Nom", font = ("Comic Sans MS", 20, "bold"), bg ="darkgoldenrod1").grid(row=0,column=1, padx = 10)
        Label(frame,text="Numéro", font = ("Comic Sans MS", 20, "bold"), bg ="darkgoldenrod1").grid(row=0,column=2, padx = 10)
        Label(frame,text="Date/Heure", font = ("Comic Sans MS", 20, "bold"), bg ="darkgoldenrod1").grid(row=0,column=4, padx = 10)
        Label(frame,text="Statut", font = ("Comic Sans MS", 20, "bold"), bg ="darkgoldenrod1").grid(row=0,column=5, padx = 10)
        while ("Parcours d'initiation" not in str(sheet_ranges['C'+ str(row_1)].value)) and (str(sheet_ranges['B'+ str(row_1)].value) != "Total"):
            row_1 += 1

        print(sheet_ranges['B'+ str(row_1)].value)
        print(row_1)
        if sheet_ranges['B'+ str(row_1)].value == "Total" :
            print("pas de parcours")
            Clear()
            label = Label(fenetre, text="Pas de parcours", bg="cyan3")
            label['font'] = font.Font(size=20)
            label.place(x=750, y=400, anchor="center")
            fenetre.update()
            time.sleep(1)
            menu()
            return()

        else :
            row_2 = row_1 +1
            print(sheet_ranges['C'+ str(row_2)].value)
            while sheet_ranges['C'+ str(row_2)].value == None and sheet_ranges['B'+ str(row_2)].value != "Total":
                print("ok")
                row_2 += 1
        print(row_2)

        while row_1 != row_2 :

            Label(frame,text= str(i), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3", highlightbackground="cyan3").grid(row=i,column=0, padx = 10)
            Label(frame,text=str(sheet_ranges['H'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3").grid(row=i,column=1, padx = 10)

            Label(frame,text=str(sheet_ranges['F'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3").grid(row=i,column=4, padx = 10)
            Label(frame,text=str(sheet_ranges['K'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3").grid(row=i,column=5, padx = 10)
            if checknum(row_1) == True :
                Label(frame,text=str(sheet_ranges['I'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg ="cyan3").grid(row=i,column=2, padx = 10)
                Checkbutton(frame, text='Rappel',  font = ("Comic Sans MS", 15, "bold"), command=partial(check , row_1), bg ='cyan3').grid(row=i,column=6, padx = 10)
            else :
                Label(frame,text=str(sheet_ranges['I'+ str(row_1)].value), font = ("Comic Sans MS", 15, "bold"), bg="red").grid(row=i,column=2, padx = 10)
                photto = PhotoImage(file='./img/pencils.gif')
                imgButton = Button(frame, image=photto, command=partial(modif, row_1))
                imgButton.image = photto
                imgButton.grid(row=i,column=3)
            i +=1
            row_1 += 1

        bt=Button(fenetre, text="validé", command = sendSMS)
        bt['font'] = font.Font(size=20)
        bt.place(x=750, y=750, anchor="center")

        ################################# Fin Fonction parcours ##############################################################

    ip = IPv4Address(entree.get()[:-5])
    print(ip)
    Clear()

    label = Label(fenetre, text="En attente de connexion", bg="cyan3")
    label['font'] = font.Font(size=20)
    label.place(x=550, y=400, anchor="center")
    print("ok")
    fenetre.update()
    session = AirmoreSession(ip)
    if not session.is_server_running:
        print("Attente")
        lable.destroy()
        label = Label(fenetre, text="En attente de connexion", bg="cyan3")
        label['font'] = font.Font(size=20)
        label.place(x=550, y=400, anchor="center")
        print("ok")
        fenetre.update()

    was_accepted = session.request_authorization()
    print(was_accepted)
    if was_accepted == False :
        label.destroy()
        label = Label(fenetre, text="Non Connecté", bg="cyan3")
        label['font'] = font.Font(size=20)
        label.place(x=550, y=400, anchor="center")
        fenetre.update()
        time.sleep(1)
        ip_func()
    else :
        label.destroy()
        label = Label(fenetre, text="Connecté", bg="cyan3")
        label['font'] = font.Font(size=20)
        label.place(x=550, y=400, anchor="center")
        fenetre.update()
        label.destroy()
        menuDeroulant1.add_command(label='Menu', command = menu)
        menu()
        service = MessagingService(session)

     ################################# Fin Fonction Connexion ##############################################################
def ip_func():
    photo2 = PhotoImage(file="./img/Airmore_2.gif")
    photo = PhotoImage(file="./img/Airmore_1.gif")
    canvas = Canvas(fenetre,width=750, height=400)
    canvas.create_image(0, 0, anchor=NW, image=photo)
    canvas.image = photo
    canvas.grid(row=0,column=0, rowspan=4)
    canvas2 = Canvas(fenetre,width=750, height=400)
    canvas2.create_image(0, 0, anchor=NW, image=photo2)
    canvas2.image = photo2
    canvas2.grid(row=5,column=0, rowspan=4)

    value = StringVar()
    value.set("192.168.169.53:2333")
    f = font.Font(size=20)
    entree = Entry(fenetre, textvariable=value, x=100, bd=4)
    entree['font'] = f
    entree.grid(row=1,column=1)

    bouton=Button(fenetre, text="valider", command= partial(connexion, entree))
    bouton['font'] = f
    bouton.grid(row=2 ,column=1)

 ################################# Boucle principal et premiere fenêtre ##############################################################
root = Tk()
root.geometry('1100x850')
zoneMenu = Frame(root, borderwidth=3, bg='darkgoldenrod1')
zoneMenu.pack()
menuFichier = Menubutton(zoneMenu, text='Menu', width='20', borderwidth=2, bg='darkgoldenrod1', activebackground='cyan3',relief = RAISED)
menuFichier.grid(row=0,column=0)
menuDeroulant1 = Menu(menuFichier)
menuDeroulant1.add_command(label='Quitter', command = quit)
menuFichier.configure(menu=menuDeroulant1)
bg = PhotoImage(file = "./img/BG.gif")

fenetre = Canvas(root, width = 1100, height = 800)
fenetre.pack(fill = "both", expand = True)
fenetre.create_image( 0, 0, image = bg,
                     anchor = "nw")

ip_func()

fenetre.mainloop()


################################################################################################# FIN Prog #################################################################################