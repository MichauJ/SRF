import numpy as np
import pandas as pd
import os
import winpath
import openpyxl
#wersja 1.3:
#dodano funkcje size_chart()
print("Wersja 1.3\nProgram wyciąga rezultaty z pliku csv i tworzy plik .xlsx gotowy do wysłania Partnerowi.")
wybor = int(input("Co chcesz sprawdzić? [wpisz cyfrę i potwierdź enter]\n1. Wyniki retaggingu\n2. Feedback odnośnie sustainability request\n3. Podsumowanie requestu o retagging\n"))

def sustainability_results():
    if os.path.exists(winpath.get_desktop()+"\\Sustainability request feedback") == False:
        os.mkdir(winpath.get_desktop()+"\\Sustainability request feedback")
        print("")
        print("Utworzono folder źródłowy na pulpicie o nazwie Sustainability request feedback.")
    input("Pobierz plik z https://drive.google.com/drive/folders/1rplbkGNyTk2CGbbI4M8SN11OXdj6PWH0 (z arkusza REPORT) jako csv i umieść go w folderze Sustainability request feedback. Wciśnij Enter aby kontynuować.")
    csv = winpath.get_desktop()+"\\Sustainability request feedback\\REPORT.csv"
    #1. zmiana nazwy pliku csv na odpowiedni
    folder = os.listdir(winpath.get_desktop()+"\\Sustainability request feedback")
    for i in folder:
        if i[-4:] == ".csv":
            os.rename(winpath.get_desktop()+"\\Sustainability request feedback\\"+i,csv)
    #2. Odczyt csv
    df = pd.read_csv(csv,low_memory = False,skiprows=10,usecols=['Config SKU','Merchant name','Certification','Primary Issue'],header=0)
    #3. Filtrowanie po nazwie partnera lub SKUs
    wybor1 = int(input("Filtrowanie (wpisz tylko cyfrę i enter): po nazwie partnera - 1, po liście SKU - 2\n"))
    if wybor1 == 1:
        partner = input("Podaj nazwę partnera (możesz podać kilka wersji, oddzielone przecinkami): ")
        partner = list(partner.split(","))
        result = df[df['Merchant name'].isin(partner)] 
        result.to_excel(winpath.get_desktop()+"\\Sustainability request feedback\\"+'Sustainability request feedback'+'.xlsx', index=False)
    if wybor1 ==2:
        #stworzenie pliku SKU.txt
        file = winpath.get_desktop()+"\\Sustainability request feedback\\SKU.txt"
        if os.path.exists(file) == False:
            open(file, mode='w').close()
        input("Skopiuj listę SKU (oddzielone enterami) do pliku 'SKU.txt' w folderze Sustainability request feedback i wciśnij enter aby kontynuować")
        SKU = []
        with open(file) as f:
            SKUs = f.readlines()
        for i in SKUs:
            SKU.append(i.replace("\n", ""))
        result = df[df['Config SKU'].isin(SKU)]
        result.to_excel(winpath.get_desktop()+"\\Sustainability request feedback\\"+'Sustainability request feedback'+'.xlsx', index=False)
        print("Uwaga! SKUs mogły zostać dodane przez różnych partnerów jednocześnie.\nPrzed wysłaniem sprawdź czy w pliku dla partnera nie ma nazw innych partnerów.")        
    print(result)
    print("W folderze źródłowym utworzono plik .xlsx zawierający powyższe wyniki.")
    
def retagging():
    if os.path.exists(winpath.get_desktop()+"\\Retagging results") == False:
        os.mkdir(winpath.get_desktop()+"\\Retagging results")
        print("")
        print("Utworzono folder źródłowy na pulpicie o nazwie Retagging results.")
    input("Pobierz wyniki dla odpowiedniego tygodnia z https://docs.google.com/spreadsheets/d/1v1dJCgGd47F2nCSBrfvyfbzIxPd_KKKxlx-MLgn6hzE/edit#gid=1460715287 jako plik csv i umieść go w folderze Retagging results.\nJeśli w folderze znajduje się inny plik csv, usuń go i zostaw tylko jeden.\nWciśnij Enter aby kontynuować.")
    csv = winpath.get_desktop()+"\\Retagging results\\REPORT.csv"
    #1. zmiana nazwy pliku csv na odpowiedni
    folder = os.listdir(winpath.get_desktop()+"\\Retagging results")
    for i in folder:
        if i[-4:] == ".csv":
            os.rename(winpath.get_desktop()+"\\Retagging results\\"+i,csv)
    #2. Odczyt csv        
    df = pd.read_csv(csv,low_memory = False,usecols=['Config SKU','Requested Season','Partner name','Feedback'],header=0)
    #3. Filtrowanie po nazwie partnera lub SKUs
    wybor1 = int(input("Filtrowanie (wpisz tylko cyfrę i enter): po nazwie partnera - 1, po liście SKU - 2\n"))
    if wybor1 == 1:
        partner = input("Podaj nazwę partnera (możesz podać kilka wersji, oddzielone przecinkami): ")
        partner = list(partner.split(","))
        result = df[df['Partner name'].isin(partner)].replace({"Excluded by ALM: ":"","Excluded by AMS: ":"",";":"","Excluded by ALM ":"","Excluded by AMS ":""},regex=True)    
        result.to_excel(winpath.get_desktop()+"\\Retagging results\\"+partner[0]+' Retagging request results'+'.xlsx', index=False)
    if wybor1 ==2:
        #stworzenie pliku SKU.txt
        file = winpath.get_desktop()+"\\Retagging results\\SKU.txt"
        if os.path.exists(file) == False:
            open(file, mode='w').close()
        input("Skopiuj listę SKU (oddzielone enterami) do pliku 'SKU.txt' w folderze Retagging results i wciśnij enter aby kontynuować")
        SKU = []
        with open(file) as f:
            SKUs = f.readlines()
        for i in SKUs:
            SKU.append(i.replace("\n", ""))
        result = df[df['Config SKU'].isin(SKU)]
        result = result.replace({"Excluded by ALM: ":"","Excluded by AMS: ":"",";":"","Excluded by ALM ":"","Excluded by AMS ":""},regex=True)
        result.to_excel(winpath.get_desktop()+"\\Retagging results\\"+'Retagging request results'+'.xlsx', index=False)
        print("Uwaga! SKUs mogły zostać dodane przez różnych partnerów jednocześnie.\nPrzed wysłaniem sprawdź czy w pliku dla partnera nie ma nazw innych partnerów.")
    print(result)
    print("W folderze źródłowym utworzono plik .xlsx zawierający powyższe wyniki.")
    
def retagging_request():
    if os.path.exists(winpath.get_desktop()+"\\Retagging results") == False:
        os.mkdir(winpath.get_desktop()+"\\Retagging results")
        print("")
        print("Utworzono folder źródłowy na pulpicie o nazwie Retagging results.")
    input("Pobierz wyniki dla odpowiedniego tygodnia z https://docs.google.com/spreadsheets/d/1BUDEEDykTzN1bRKUr9_0cEor6ob8i6mddTrlpC1rBPQ/edit#gid=1930906818 jako plik csv i umieść go w folderze Retagging results.\nJeśli w folderze znajduje się inny plik csv, usuń go i zostaw tylko jeden.\nWciśnij Enter aby kontynuować.")
    csv = winpath.get_desktop()+"\\Retagging results\\REPORT.csv"
    #1. zmiana nazwy pliku csv na odpowiedni
    folder = os.listdir(winpath.get_desktop()+"\\Retagging results")
    for i in folder:
        if i[-4:] == ".csv":
            os.rename(winpath.get_desktop()+"\\Retagging results\\"+i,csv)
    #2. Odczyt csv        
    df = pd.read_csv(csv,low_memory = False,usecols=['CW#','Zalando config SKU (zalando_article_variant)','Requested Season','Partner name'],header=0,skiprows=2)
    #3. Filtrowanie po nazwie partnera lub SKUs
    wybor1 = int(input("Filtrowanie (wpisz tylko cyfrę i enter): po nazwie partnera - 1, po liście SKU - 2\n"))
    if wybor1 == 1:
        partner = input("Podaj nazwę partnera (możesz podać kilka wersji, oddzielone przecinkami): ")
        partner = list(partner.split(","))
        result = df[df['Partner name'].isin(partner)]
        result = result[['Zalando config SKU (zalando_article_variant)','Requested Season']]
        result.to_excel(winpath.get_desktop()+"\\Retagging results\\"+partner[0]+' Retagging request summary'+'.xlsx', index=False)
    if wybor1 ==2:
        #stworzenie pliku SKU.txt
        file = winpath.get_desktop()+"\\Retagging results\\SKU.txt"
        if os.path.exists(file) == False:
            open(file, mode='w').close()
        input("Skopiuj listę SKU (oddzielone enterami) do pliku 'SKU.txt' w folderze Retagging results i wciśnij enter aby kontynuować")
        SKU = []
        with open(file) as f:
            SKUs = f.readlines()
        for i in SKUs:
            SKU.append(i.replace("\n", ""))
        #df['Requested']=df['Zalando config SKU (zalando_article_variant)'].map(lambda x: 'yes' if x in SKU == True else 'no')
        #result = df[df['Zalando config SKU (zalando_article_variant)'].isin(SKU)]
        #result.to_excel(winpath.get_desktop()+"\\Retagging results\\"+'Retagging request results'+'.xlsx', index=False)
        SKU_in = []
        for i in SKU:
            if i in list(df['Zalando config SKU (zalando_article_variant)']):
                SKU_in.append('yes')
            else:
                SKU_in.append('no')
        #result = pd.DataFrame({'SKU':SKU,'Present':list(lambda x: 'yes' if x in df[df['Zalando config SKU (zalando_article_variant)']] else 'no')}).shift()[1:]
        result = pd.DataFrame(zip(SKU,SKU_in),columns = ('SKU','Retagging requested'))
        result.to_excel(winpath.get_desktop()+"\\Retagging results\\"+'Retagging request summary'+'.xlsx', index=False)
    print(result)
    print("W folderze źródłowym utworzono plik .xlsx zawierający powyższe wyniki.")    

def size_chart():
    if os.path.exists(winpath.get_desktop()+"\\Finding size chart") == False:
        os.mkdir(winpath.get_desktop()+"\\Finding size chart")
        print("")
        print("Utworzono folder źródłowy na pulpicie o nazwie Finding size chart.")
    input("Pobierz plik size chart library z https://sites.google.com/a/zalando.de/sizing-business-integration/size-charts jako csv i umieść go w folderze Finding size chart. Wciśnij Enter aby kontynuować.")
    csv = winpath.get_desktop()+"\\Finding size chart\\SC LIBRARY.csv"
    #1. zmiana nazwy pliku csv na odpowiedni
    folder = os.listdir(winpath.get_desktop()+"\\Finding size chart")
    for i in folder:
        if i[-4:] == ".csv":
            os.rename(winpath.get_desktop()+"\\Finding size chart\\"+i,csv)
    df = pd.read_csv(csv,low_memory = False,skiprows=3,header=0,dtype=str)
    #tworzenie pomocniczego pliku gdzie trzeba wpisać pożądane rozmiary
    #cg1 = input("Podaj kategorię z dostępnych: Accesories,Beauty,Home,Shoes,Sports,Textile,Textile - Lingerie & Beachwear: ")
    cg1 = 'Shoes'
    #cg2 = input("Podaj gender z dostępnych: Female,Kids,Unisex,Male: ")
    cg2 = 'Female'
    bc = input("Podaj brand code i/lub 000 dla standardowych SC: ")
    if os.path.exists(winpath.get_desktop()+"\\Finding size chart\\Required sizes.xlsx") == False:
       sizes = pd.DataFrame([],columns = ['Supplier Size','EU','UK','US','FR','IT','Brand code','CG1','CG2'],dtype=str)
       sizes.to_excel(winpath.get_desktop()+"\\Finding size chart\\Required sizes.xlsx",index = False)
       print("Utworzono plik 'Required sizes' w folderze źródłowym.")
    input("Uzupełnij plik 'Required sizes' w folderze źródłowym szukanymi rozmiarami i wciśnij enter aby kontynuować")
    searched_sizes = pd.read_excel(winpath.get_desktop()+"\\Finding size chart\\Required sizes.xlsx",dtype=str).replace({'NaN':'-'},regex=True)
    result = df[df['BRAND_CODE'].isin(list(searched_sizes['Brand code']))
    &df['CG1'].isin(list(searched_sizes['CG1']))
    &df['CG2'].isin(list(searched_sizes['CG2']))
    &df['EU'].isin(list(searched_sizes['EU']))
    &df['UK'].isin(list(searched_sizes['UK']))
    &df['FR'].isin(list(searched_sizes['FR']))
    &df['IT'].isin(list(searched_sizes['IT']))
    &df['Supplier Size'].isin(list(searched_sizes['Supplier Size']))]
    print(list(searched_sizes['Brand code']),list(searched_sizes['CG1']),list(searched_sizes['CG2']))
    print("Znaleziono size charty spełniające kryteria: ",list(set(result['Size Chart Code'])))
    result.to_excel(winpath.get_desktop()+"\\Finding size chart\\Found size charts.xlsx",index = False)
   
   
    print(result)
    result.to_excel(winpath.get_desktop()+"\\Finding size chart\\Found sizes.xlsx",index = False)



start = 'c'        
while start =='c':            
    if wybor == 2:  
        try:
            sustainability_results()
        except FileExistsError:
            print("Wystąpił błąd - usuń z folderu wszystkie pliki csv poza właściwym raportem i uruchom ponownie.")
        except FileNotFoundError:
            print("W folderze źródłowym nie ma żadnego pliku csv. Pobierz plik z https://drive.google.com/drive/folders/1rplbkGNyTk2CGbbI4M8SN11OXdj6PWH0 i umieść go w folderze Sustainability request feedback.")
        except Exception as error:
            print("Wystąpił błąd:", type(error).__name__, "–", error)
            print("Skopiuj otrzymany błąd do https://docs.google.com/spreadsheets/d/1CaVxhDBdZUEmGofsQE_0yM2PSbKJleC-oAh0tVbRFvo/edit#gid=0")
        finally:
            pass
    if wybor == 1:
        try:
            retagging()
        except FileExistsError:
            print("Wystąpił błąd - usuń z folderu wszystkie pliki csv poza właściwym raportem i uruchom ponownie.")
        except FileNotFoundError:
            print("Pobierz wyniki dla odpowiedniego tygodnia z https://docs.google.com/spreadsheets/d/1v1dJCgGd47F2nCSBrfvyfbzIxPd_KKKxlx-MLgn6hzE/edit#gid=1460715287 jako plik csv i umieść go w folderze Retagging results.")
        except Exception as error:
            print("Wystąpił błąd:", type(error).__name__, "–", error)
            print("Skopiuj otrzymany błąd do https://docs.google.com/spreadsheets/d/1CaVxhDBdZUEmGofsQE_0yM2PSbKJleC-oAh0tVbRFvo/edit#gid=0")
        finally:
            pass
    if wybor ==3:
        try:
            retagging_request()
        except FileExistsError:
            print("Wystąpił błąd - usuń z folderu wszystkie pliki csv poza właściwym raportem i uruchom ponownie.")
        except FileNotFoundError:
            print("Pobierz wyniki dla odpowiedniego tygodnia z https://docs.google.com/spreadsheets/d/1BUDEEDykTzN1bRKUr9_0cEor6ob8i6mddTrlpC1rBPQ/edit#gid=1930906818 jako plik csv i umieść go w folderze Retagging results.")
        except Exception as error:
            print("Wystąpił błąd:", type(error).__name__, "–", error)
            print("Skopiuj otrzymany błąd do https://docs.google.com/spreadsheets/d/1CaVxhDBdZUEmGofsQE_0yM2PSbKJleC-oAh0tVbRFvo/edit#gid=0")
        finally:
            pass
    start = input("Zamknij lub wpisz c aby zacząć od początku:  ")
    if start == 'c':
        wybor = int(input("Co chcesz sprawdzić? [wpisz cyfrę i potwierdź enter]\n1. Wyniki retaggingu\n2. Feedback odnośnie sustainability request\n3. Podsumowanie requestu o retagging\n"))    
