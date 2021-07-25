# -*- coding: utf-8 -*-

"""
Created on Tue Nov 26 16:30:41 2019

@author: Hisham
"""

import pandas as pd
from pandas import Series
import pathlib
from tabulate import tabulate
import os


class rota():
    def DayJuggler(self, df, days, Mreq, Ereq, Nreq, Mcode, Ecode, Ncode, NOcode, WOcode, nil, Dreq, roster):
        ##########################################################
        ##     This function juggles different count of days    ##
        ##     to find out if 29 or 30 or 31 or 28 days gives   ##
        ##     a desired output matching requirements           ##
        ##########################################################

        # reset tweekid
        tweek = int

        #    reseting switches

        summary_switch = False
        match_switch = True
        filesave_switch = False
        pdf_file = False
        if (len(df.columns.values) <= 32):
            for tweek in range(3, 10):
                if (tweek < 10):
                    print('Tweek: ', tweek)
                    df = roster.copy(deep=True)

                    df.columns = df.columns.astype(str)
                    df.sort_values(by=['EMP.ID'], inplace=True)
                    df.reset_index(inplace=True)
                    del df['index']  # deletes the column created of old index after reset_index()
                    NameList = list(df['NAMES'])
                    Emplist = list(df['EMP.ID'])
                    Slno = list(df['S.NO'])
                    df = df.drop(columns=['NAMES'])
                    df = df.drop(columns=['EMP.ID'])
                    df = df.drop(columns=['S.NO'])
                    df['temp'] = float("NaN")  # creates an extra temp days for testing
                    days = list(df.columns.values)
                    staffs = range(len(df))
                    df = self.NilMaker(self, df, tweek, nil)
                    df = self.RosterMaker(self, df, tweek, Mcode, Ecode, Ncode, NOcode, WOcode, nil)
                    del df['temp']  # deletes the temporary day from roster before saving
                    days = list(df.columns.values)
                    df = self.CodeCounter(self, df, summary_switch, tweek, Mcode, Ecode, Ncode, NOcode, WOcode, nil)
                    df.insert(loc=0, column='S.NO', value=Slno)
                    df, df2, Mlist = self.Rsummary(self, df, tweek, Mcode, Ecode, Ncode, NOcode, WOcode, nil, Nreq,
                                                   Mreq, Ereq)
                    print("Dayjuggled")
                    print(tabulate(df2, headers='keys', tablefmt='psql', showindex=False))

                    print("Nreq:", Nreq, "Dreq:", Dreq)

                    if ((df2.iloc[2][2] >= Nreq) and (df2.iloc[0][2] >= Dreq)) or pdf_file == True:

                        match_switch = True
                        print("Match found")

                        summary_switch = True
                    else:
                        match_switch = False
                        summary_switch = False

                    if match_switch == True and filesave_switch == False:
                        for i in range(6):
                            Emplist.append(int('0'))
                            NameList.append(int('0'))
                        df.insert(loc=1, column='EMP.ID', value=Emplist)
                        df.insert(loc=2, column='NAMES', value=NameList)

                        df = self.Csummary(self, df, Mcode, Ecode, Ncode, NOcode, WOcode, days)

                        df = self.HoursCalculator(self, df)

                        df.to_excel("rosterupdate.xls", index=None)
                        print("New file saved!")
                        filesave_switch = True
                        self.ExcelDesigner(self, df, df2, tweek, days, Mreq, Ereq, Nreq, Mcode, Ecode, Ncode, NOcode,
                                           WOcode, nil)

                        break
            if filesave_switch == False:
                print("requirement match not found!")
            return df, df2

    def NewStaffCleaner(self, NOT_df, tweek, Mcode, Ecode, Ncode, NOcode, WOcode, nil):
        if (NOT_df.empty == False):
            New_NameList = list(NOT_df['NAMES'])
            New_Emplist = list(NOT_df['EMP.ID'])
            New_Slno = list(NOT_df['S.NO'])
            del NOT_df['S.NO']
            del NOT_df['EMP.ID']
            del NOT_df['NAMES']
            NOT_df = NOT_df.reset_index()
            del NOT_df['index']
            NOT_df = self.NilMaker(self, NOT_df, tweek, nil)
            NOT_df = self.RosterMaker(self, NOT_df, tweek, Mcode, Ecode, Ncode, NOcode, WOcode, nil)

        else:
            New_NameList = []
            New_Emplist = []
            New_Slno = []
            del NOT_df['S.NO']
            del NOT_df['EMP.ID']
            del NOT_df['NAMES']

        # NOT_df,df2,Mlist=Rsummary(NOT_df)

        return NOT_df, New_NameList, New_Emplist, New_Slno

    def staffactuator(self, df, pdf, previous_roster, roster):
        ## campares the dataframes in case of continuum for staff list changes.
        ## filters out staff list from previous roster.
        ## creates another df of new staff
        pdf = previous_roster.copy(deep=True)
        pdf.columns = pdf.columns.astype(str)
        x = len(pdf)
        # pdf.drop([x - 1, x - 2, x - 3, x - 4, x - 5, x - 6], inplace=True)  # deleting rows at the bottom

        ndf = roster.copy(deep=True)

        ndf.columns = ndf.columns.astype(str)

        nlen = len(ndf)

        plen = len(pdf)
        ndf.sort_values(by=['EMP.ID'], inplace=True)
        ndf.reset_index(inplace=True)
        del ndf['index']  # deleting index column created after reset_index()
        pdf.sort_values(by=['EMP.ID'], inplace=True)
        pdf.reset_index(inplace=True)
        del pdf['index']
        pEMP = pdf['EMP.ID']
        nEMP = ndf['EMP.ID']
        NOT_df = False
        change = False
        if (nlen > plen):
            NOT_df = ndf[~ndf['EMP.ID'].isin(pEMP)]
            ndf = ndf[ndf['EMP.ID'].isin(pEMP)]
            change = True

        elif (plen > nlen):
            pdf = pdf[pdf['EMP.ID'].isin(nEMP)]
            NOT_df = ndf[~ndf['EMP.ID'].isin(pEMP)]
            ndf = ndf[~ndf['EMP.ID'].isin(list(NOT_df['EMP.ID']))]
            change = True
        elif (plen == nlen):
            NOT_df = ndf[~ndf['EMP.ID'].isin(nEMP)]

        if change == True:
            pdf = pdf[pdf['EMP.ID'].isin(nEMP)]

        NameList = list(ndf['NAMES'])
        Emplist = list(ndf['EMP.ID'])
        Slno = list(ndf['S.NO'])
        del ndf['NAMES']
        del ndf['EMP.ID']
        del ndf['S.NO']
        days = list(ndf.columns.values)
        staffs = range(len(ndf))

        return ndf, pdf, NOT_df, NameList, Emplist, Slno, days, staffs

    def ReqReader(self, req_location, tweek_location):
        ## This function reads the requirement from the requirement sheet.
        ## This function also reads the tweekID from previous file in case
        ## if it detects previous file
        Mreq = 0
        Ereq = 0
        Nreq = 0
        req = pd.read_excel("roster.xls", sheet_name="Requirement")
        print("requirement read")
        if (req_location != 0):  # filters 'others' class of designations'
            Mreq = req.iloc[0][req_location]
            Ereq = req.iloc[1][req_location]
            Nreq = req.iloc[2][req_location]

        if Mreq > Ereq:
            Dreq = Mreq
        elif Ereq > Mreq:
            Dreq = Ereq
        else:
            Dreq = Mreq

        file = pathlib.Path("previous_roster.xls")
        if file.exists():
            tweek = 6
            if (tweek_location != 0):
                pdf = pd.read_excel("previous_roster.xls", sheet_name="Requirement")
                tweek = pdf.iloc[3][tweek_location]

        else:
            tweek = 0
        return Nreq, Mreq, Dreq, Ereq, tweek

    def newCLmaker(self, df3, tweek, Mcode, Ecode, Ncode, NOcode, WOcode, nil, roster):
        ##this function activates for continuum pathway
        ## reads the planned leave and assigns them
        df4 = roster.copy(deep=True)

        df4.columns = df4.columns.astype(str)
        df4.sort_values(by=['EMP.ID'], inplace=True)
        df4.reset_index(inplace=True)
        del df4['index']

        del df4['S.NO']
        del df4['EMP.ID']
        del df4['NAMES']
        staffs = range(len(df4))
        days = list(df4.columns.values)

        for staff in staffs:
            for day in days:
                if (type(df4.loc[staff, day]) == str):
                    if (df3.loc[staff, day] == Ecode or df3.loc[staff, day] == Mcode
                            or df3.loc[staff, day] == Ncode):
                        df3.loc[staff, day] = df4.loc[staff, day]
        return df3

    def ShiftReversor(self, pdf, Mcode, Ecode, Ncode, NOcode, WOcode):
        ##This function reverses the shift codes to numbers in case of creating continuum
        ## gets called into action when a previous file is detected
        ## cleans the data to bare bones and has algorithm to reverse engineer roster code to
        ##numbers it used to be.

        # df=pd.read_excel("roster.xls", skiprows=1)
        # df=pd.read_excel("previous_roster.xls",sheet_name="Duplicate", skiprows=1)
        del pdf["S.NO"]
        del pdf["NAMES"]
        del pdf["EMP.ID"]
        del pdf["Morning"]
        del pdf["Evening"]
        del pdf["Night"]
        del pdf["NightOff"]
        del pdf["WeekOff"]
        del pdf["Hours/Mn"]
        x = len(pdf)
        # print(x)
        # pdf.drop([x-1,x-2,x-3,x-4,x-5, x-6],inplace= True) #deleting rows at the bottom

        # print(pdf)

        # creating a list to add as new column in df3

        staffs = range(len(pdf))
        days = list(pdf.columns.values)
        # print(days)
        npdf = pd.DataFrame(columns=days)
        # print(npdf)
        plist = []
        for staff in staffs:
            plist = Series.tolist(pdf.iloc[staff])
            # This substitutes all '1' with '10' in list 'a' and places result in list 'c':

            plist = list(map(lambda p: p.replace(Ncode, "4"), plist))
            check = 0
            count = 0
            Monthlyweekoff = 0
            for p in plist:
                if p == WOcode:
                    Monthlyweekoff = Monthlyweekoff + 1
            # print(Monthlyweekoff)
            for i, p in enumerate(plist):
                if (p == "4"):
                    plist[i] = 4
                    check = 1
                elif (check == 1 and p == Ecode):
                    plist[i] = 1
                elif (check == 1 and p == Mcode):
                    plist[i] = 2
                    count = count + 1
                    if (count == 6):
                        count = 0
                        check = 0
                elif (check == 0 and (p == Ecode or p == Mcode)):
                    plist[i] = 3

                # elif(p=='50'):
                #   plist[i]=plist[i-1]
                # elif(p!="A4" and p!="M6" and p!="4" and p!='50'):
                #  plist[i]=plist[i-1]

            # print(plist)
            plist = pd.Series(plist, index=days)
            # pdf.iloc[staff]=plist
            npdf = npdf.append(plist, ignore_index=True)
        # print(npdf)
        return npdf, plist

    # print(npdf.head(40))
    # df3=npdf
    # print(df3.head(40))

    def NewShiftMaker(self, npdf, plist, tweek, df3, Mcode, Ecode, Ncode, NOcode, WOcode, nil):
        ##This function contains the algorithm to create the continuum for each staff

        daycount = 0
        staffs = range(len(df3))
        days = list(df3.columns.values)
        prlist = []
        c = 0

        for staff in staffs:
            prlist = []
            plist = Series.tolist(df3.iloc[staff])

            for r in reversed(plist):
                prlist.append(r)
            reversed(plist)

            # if the last column is not 1,2,3,4, wocode or nocode, the program can't function
            # hence in exception cases, pulling out values from second last and third last column
            # as long as it is expected value (first  if)

            # if(prlist[0]!=1 or prlist[0]!=2 or prlist[0]!=3
            #   or prlist[0]!=4 or prlist[0]!= WOcode
            #  or prlist[0]!=NOcode):
            #      prlist[0]=prlist[0+1]

            # print(prlist)

            if (prlist[0] == 1):  # 1
                daycount = 1
                shift_code = prlist[0]
                if ((prlist[1] == 1)):  # 2
                    shift_code = prlist[1]
                    daycount = daycount + 1
                    if (prlist[1] == 1 and prlist[1] == prlist[2]):  # 3
                        shift_code = prlist[1]
                        daycount = daycount + 1
                        if (prlist[2] == 1 and prlist[2] == prlist[3]):  # 4
                            shift_code = prlist[2]
                            daycount = daycount + 1
                            if (prlist[3] == 1 and prlist[3] == prlist[4]):  # 5
                                shift_code = prlist[3]
                                daycount = daycount + 1
                                if (prlist[4] == 1 and prlist[4] == prlist[5]):  # 6
                                    shift_code = prlist[4]
                                    daycount = daycount + 1
            elif (prlist[0] == 2):  # 1
                daycount = 1
                shift_code = prlist[0]
                if ((prlist[1] == 2)):  # 2
                    shift_code = prlist[1]
                    daycount = daycount + 1
                    if (prlist[1] == 2 and prlist[1] == prlist[2]):  # 3
                        shift_code = prlist[1]
                        daycount = daycount + 1
                        if (prlist[2] == 2 and prlist[2] == prlist[3]):  # 4
                            shift_code = prlist[2]
                            daycount = daycount + 1
                            if (prlist[3] == 2 and prlist[3] == prlist[4]):  # 5
                                shift_code = prlist[3]
                                daycount = daycount + 1
                                if (prlist[4] == 2 and prlist[4] == prlist[5]):  # 6
                                    shift_code = prlist[4]
                                    daycount = daycount + 1


            elif (prlist[0] == WOcode and prlist[1] == 1):
                daycount = 0
                shift_code = 2
            elif (prlist[0] == WOcode and prlist[1] == 2):
                daycount = 0
                # print("hi")
                shift_code = 3
            elif (prlist[0] == WOcode and prlist[1] == 3):
                daycount = 0
                shift_code = 4
            elif (prlist[0] == NOcode and prlist[1] == 4):
                daycount = 8
                shift_code = 4
            elif (prlist[0] == WOcode and prlist[1] == NOcode):
                daycount = 0
                shift_code = 1



            elif (prlist[0] == 4):  # 1
                daycount = 1
                shift_code = prlist[0]
                if ((prlist[1] == 4)):  # 2
                    shift_code = prlist[1]
                    daycount = daycount + 1
                    if (prlist[1] == 4 and prlist[1] == prlist[2]):  # 3
                        shift_code = prlist[1]
                        daycount = daycount + 1
                        if (prlist[2] == 4 and prlist[2] == prlist[3]):  # 4
                            shift_code = prlist[2]
                            daycount = daycount + 1
                            if (prlist[3] == 4 and prlist[3] == prlist[4]):  # 5
                                shift_code = prlist[3]
                                daycount = daycount + 1
                                if (prlist[4] == 4 and prlist[4] == prlist[5]):  # 6
                                    shift_code = prlist[4]
                                    daycount = daycount + 1
                                    if (prlist[5] == 4 and prlist[5] == prlist[6]):  # 7
                                        shift_code = prlist[5]
                                        daycount = daycount + 1
                                        if (prlist[6] != 4):  # 8
                                            shift_code = NOcode
                                            daycount = 0




            elif (prlist[0] == 3 and tweek == 9):
                tweek = tweek
                if (prlist[0] == 3):  # 1
                    daycount = 1
                    shift_code = prlist[0]

                    if ((prlist[1] == 3)):  # 2
                        shift_code = prlist[1]
                        daycount = daycount + 1

                        if (prlist[2] == 3):  # 3
                            shift_code = prlist[2]
                            daycount = daycount + 1

                            if (prlist[3] == 3):  # 4
                                shift_code = prlist[3]
                                daycount = daycount + 1

                                if (prlist[4] == 3):  # 5
                                    shift_code = prlist[4]
                                    daycount = daycount + 1
                                    if (prlist[5] == 3):  # 6
                                        shift_code = prlist[5]
                                        daycount = daycount + 1
                                        if (prlist[5] == 3 and prlist[5] == prlist[6]):  # 7
                                            shift_code = prlist[5]
                                            daycount = daycount + 1
                                            if (prlist[6] == 3 and prlist[6] == prlist[7]):  # 8
                                                shift_code = prlist[6]
                                                daycount = daycount + 1
                                                if (prlist[7] == 3 and prlist[7] == prlist[8]):  # 9
                                                    shift_code = prlist[7]
                                                    daycount = daycount + 1
                                                    if (prlist[8] == 3 and prlist[8] == prlist[9]):  # 10
                                                        shift_code = prlist[8]
                                                        daycount = daycount + 1
                                                        if (prlist[8] == 3 and prlist[8] == prlist[9]):  # 11
                                                            shift_code = 4
                                                            daycount = 0
            elif (prlist[0] == 3 and tweek == 8):
                tweek = tweek
                if (prlist[0] == 3):  # 1
                    daycount = 1
                    shift_code = prlist[0]

                    if ((prlist[1] == 3)):  # 2
                        shift_code = prlist[1]
                        daycount = daycount + 1

                        if (prlist[2] == 3):  # 3
                            shift_code = prlist[2]
                            daycount = daycount + 1

                            if (prlist[3] == 3):  # 4
                                shift_code = prlist[3]
                                daycount = daycount + 1

                                if (prlist[4] == 3):  # 5
                                    shift_code = prlist[4]
                                    daycount = daycount + 1
                                    if (prlist[5] == 3):  # 6
                                        shift_code = prlist[5]
                                        daycount = daycount + 1
                                        if (prlist[5] == 3 and prlist[5] == prlist[6]):  # 7
                                            shift_code = prlist[5]
                                            daycount = daycount + 1
                                            if (prlist[6] == 3 and prlist[6] == prlist[7]):  # 8
                                                shift_code = prlist[6]
                                                daycount = daycount + 1
                                                if (prlist[7] == 3 and prlist[7] == prlist[8]):  # 9
                                                    shift_code = prlist[7]
                                                    daycount = daycount + 1
                                                    if (prlist[8] == 3 and prlist[8] == prlist[9]):  # 10
                                                        shift_code = 4
                                                        daycount = 0


            elif (prlist[0] == 3 and tweek == 7):
                tweek = tweek
                if (prlist[0] == 3):  # 1
                    daycount = 1
                    shift_code = prlist[0]

                    if ((prlist[1] == 3)):  # 2
                        shift_code = prlist[1]
                        daycount = daycount + 1

                        if (prlist[2] == 3):  # 3
                            shift_code = prlist[2]
                            daycount = daycount + 1

                            if (prlist[3] == 3):  # 4
                                shift_code = prlist[3]
                                daycount = daycount + 1

                                if (prlist[4] == 3):  # 5
                                    shift_code = prlist[4]
                                    daycount = daycount + 1
                                    if (prlist[5] == 3):  # 6
                                        shift_code = prlist[5]
                                        daycount = daycount + 1
                                        if (prlist[5] == 3 and prlist[5] == prlist[6]):  # 7
                                            shift_code = prlist[5]
                                            daycount = daycount + 1
                                            if (prlist[6] == 3 and prlist[6] == prlist[7]):  # 8
                                                shift_code = prlist[6]
                                                daycount = daycount + 1
                                                if (prlist[7] == 3 and prlist[7] == prlist[8]):  # 9
                                                    shift_code = 4
                                                    daycount = 0
            elif (prlist[0] == 3 and tweek == 6):

                tweek = tweek
                if (prlist[0] == 3):  # 1
                    daycount = 1
                    shift_code = prlist[0]

                    if ((prlist[1] == 3)):  # 2
                        shift_code = prlist[1]
                        daycount = daycount + 1

                        if (prlist[2] == 3):  # 3
                            shift_code = prlist[2]
                            daycount = daycount + 1

                            if (prlist[3] == 3):  # 4
                                shift_code = prlist[3]
                                daycount = daycount + 1

                                if (prlist[4] == 3):  # 5
                                    shift_code = prlist[4]
                                    daycount = daycount + 1
                                    if (prlist[5] == 3):  # 6
                                        shift_code = prlist[5]
                                        daycount = daycount + 1
                                        if (prlist[5] == 3 and prlist[5] == prlist[6]):  # 7
                                            shift_code = prlist[5]
                                            daycount = daycount + 1
                                            if (prlist[6] == 3 and prlist[6] == prlist[7]):  # 8
                                                shift_code = 4
                                                daycount = 0
            elif (prlist[0] == 3 and tweek == 5):
                tweek = tweek
                if (prlist[0] == 3):  # 1
                    daycount = 1
                    shift_code = prlist[0]

                    if ((prlist[1] == 3)):  # 2
                        shift_code = prlist[1]
                        daycount = daycount + 1

                        if (prlist[2] == 3):  # 3
                            shift_code = prlist[2]
                            daycount = daycount + 1

                            if (prlist[3] == 3):  # 4
                                shift_code = prlist[3]
                                daycount = daycount + 1

                                if (prlist[4] == 3):  # 5
                                    shift_code = prlist[4]
                                    daycount = daycount + 1
                                    if (prlist[5] == 3):  # 6
                                        shift_code = prlist[5]
                                        daycount = daycount + 1
                                        if (prlist[5] == 3 and prlist[5] == prlist[6]):  # 7
                                            shift_code = 4
                                            daycount = 0


            elif (prlist[0] == 3 and tweek == 4):
                tweek = tweek
                if (prlist[0] == 3):  # 1
                    daycount = 1
                    shift_code = prlist[0]

                    if ((prlist[1] == 3)):  # 2
                        shift_code = prlist[1]
                        daycount = daycount + 1

                        if (prlist[2] == 3):  # 3
                            shift_code = prlist[2]
                            daycount = daycount + 1

                            if (prlist[3] == 3):  # 4
                                shift_code = prlist[3]
                                daycount = daycount + 1

                                if (prlist[4] == 3):  # 5
                                    shift_code = prlist[4]
                                    daycount = daycount + 1
                                    if (prlist[5] == 3):  # 6
                                        shift_code = 4
                                        daycount = 0

            elif (prlist[0] == 3 and tweek == 3):
                tweek = tweek
                if (prlist[0] == 3):  # 1
                    daycount = 1
                    shift_code = prlist[0]

                    if ((prlist[1] == 3)):  # 2
                        shift_code = prlist[1]
                        daycount = daycount + 1

                        if (prlist[2] == 3):  # 3
                            shift_code = prlist[2]
                            daycount = daycount + 1

                            if (prlist[3] == 3):  # 4
                                shift_code = prlist[3]
                                daycount = daycount + 1

                                if (prlist[4] == 3):  # 5
                                    shift_code = 4
                                    daycount = 0

            elif (prlist[0] == 3 and tweek == 2):
                tweek = tweek
                if (prlist[0] == 3):  # 1
                    daycount = 1
                    shift_code = prlist[0]
                    if ((prlist[1] == 3)):  # 2
                        shift_code = prlist[1]
                        daycount = daycount + 1

                        if (prlist[2] == 3):  # 3
                            shift_code = prlist[2]
                            daycount = daycount + 1
                            if (prlist[3] == 3):  # 4
                                shift_code = 4
                                daycount = 0


            elif (prlist[0] == 3 and tweek == 1):
                tweek = tweek
                if (prlist[0] == 3):  # 1
                    daycount = 1
                    shift_code = prlist[0]

                    if ((prlist[1] == 3)):  # 2
                        shift_code = prlist[1]
                        daycount = daycount + 1

                        if (prlist[2] == 3):  # 3
                            shift_code = prlist[2]
                            daycount = daycount + 1

            elif (prlist[0] == 3 and tweek == 0):
                if (prlist[0] == 3):  # 1
                    daycount = 1
                    shift_code = prlist[0]

                    if ((prlist[1] == 3)):  # 2
                        shift_code = 4
                        daycount = 0

            for day in days:
                if (shift_code == 4):
                    if (daycount <= 6):
                        # if(df3.loc[staff,day]==nil):
                        df3.loc[staff, day] = shift_code
                        daycount = daycount + 1

                    elif (daycount == 7):
                        df3.loc[staff, day] = NOcode

                        daycount = daycount + 1
                    elif (daycount == 8):
                        df3.loc[staff, day] = WOcode
                        daycount = daycount + 1
                        shift_code = 1
                        daycount = 0

                elif (shift_code == 1 or shift_code == 2):
                    if (daycount <= 5):
                        # print(daycount)
                        # if(df3.loc[staff,day]==nil):
                        df3.loc[staff, day] = shift_code
                        daycount = daycount + 1


                    elif (daycount == 6):
                        # print("endofshift1")
                        # print(daycount)
                        df3.loc[staff, day] = WOcode
                        daycount = 0

                        # print("endofshiftw0")
                        if (shift_code == 1):
                            shift_code = 2
                        elif (shift_code == 2):
                            shift_code = 3
                elif (shift_code == 3):
                    if (daycount < tweek):
                        # if(df3.loc[staff,day]==nil):
                        df3.loc[staff, day] = shift_code
                        daycount = daycount + 1
                    else:
                        # if(df3.loc[staff,day]==nil):
                        df3.loc[staff, day] = WOcode
                        daycount = 0
                        shift_code = 4
            daycount = 0
            shift_code = 0
        return df3

    # print(df3.head(30))
    def newcodemaker(self, blankdf, df3, tweek, Mcode, Ecode, Ncode, NOcode, WOcode, nil):
        ## This function makes the receiving df to 31 days...
        ## creates continuum of 31 days regardless of input number of days of month.
        ## other than that it just assigned roster codes to number in each cell.

        endcode = 0
        l = []

        for item in range(len(df3)):
            l.append('50')
        # print(l)
        month = 0
        if (len(blankdf.columns.values) == 28):
            blankdf["29"] = l
            blankdf["30"] = l
            blankdf["31"] = l
            month = 1
        elif (len(blankdf.columns.values) == 29):

            blankdf["30"] = l
            blankdf["31"] = l
            month = 2
        elif (len(blankdf.columns.values) == 30):
            blankdf["31"] = l
            month = 3
            # print(df.shape())
        daycount = 0
        l = 0

        staffs = range(len(blankdf))
        days = list(blankdf.columns.values)

        for staff in staffs:
            if (l == 0):
                l = 1
            elif (l == 1):
                l = 0
            daycount = 0
            c = 0

            for day in days:
                if (blankdf.loc[staff, day] == 1):

                    df3.loc[staff, day] = Ecode

                    endcode = 1
                    daycount = daycount + 1
                elif (blankdf.loc[staff, day] == 2):
                    df3.loc[staff, day] = Mcode
                    daycount = daycount + 1
                    endcode = 2
                elif (blankdf.loc[staff, day] == 4):
                    df3.loc[staff, day] = Ncode
                    daycount = daycount + 1
                    endcode = 4
                    c = 1
                elif (blankdf.loc[staff, day] == 3):
                    if (l == 1):
                        df3.loc[staff, day] = Ecode
                        daycount = daycount + 1
                        endcode = 4
                    elif (l == 0):
                        df3.loc[staff, day] = Mcode
                        daycount = daycount + 1
                        endcode = 4
                elif (blankdf.loc[staff, day] == WOcode):
                    daycount = 0
                    df3.loc[staff, day] = WOcode
                elif (blankdf.loc[staff, day] == NOcode):
                    df3.loc[staff, day] = NOcode
                    daycount = 8 + tweek + 2
                    endcode = 4
                elif (blankdf.loc[staff, day] == '50'):
                    # print(daycount)
                    if (daycount == (7 + tweek + 2) and endcode == 4):
                        df3.loc[staff, day] = NOcode
                        daycount = daycount + 1
                        endcode = 4
                    elif (daycount == (8 + tweek + 2) and endcode == 4):
                        df3.loc[staff, day] = WOcode
                        daycount = 0
                        endcode = 1
                    elif (daycount == 0 and endcode == 1):
                        df3.loc[staff, day] = Ecode
                        daycount = daycount + 1
                    elif ((daycount >= tweek + 2 and daycount < (7 + tweek + 2)) and endcode == 4):
                        df3.loc[staff, day] = Ncode
                        daycount = daycount + 1
                    elif ((daycount > 0 and daycount < tweek + 2) and endcode == 4):
                        if (l == 0):
                            df3.loc[staff, day] = Mcode
                            daycount = daycount + 1
                        elif (l == 1):
                            df3.loc[staff, day] = Ecode
                            daycount = daycount + 1
                    elif (daycount == 0 and endcode == 4):
                        df3.loc[staff, day] = Ecode
                        daycount = daycount + 1
                        endcode = 1
                    elif (daycount <= 5 and endcode == 2):
                        df3.loc[staff, day] = Mcode
                        daycount = daycount + 1
                    elif (daycount == 6 and endcode == 2):
                        df3.loc[staff, day] = WOcode
                        daycount = 1
                        endcode = 4
                    elif ((daycount > 0 and daycount <= 5) and endcode == 1):
                        df3.loc[staff, day] = Ecode
                        daycount = daycount + 1
                    elif (daycount == 6 and endcode == 1):
                        df3.loc[staff, day] = WOcode
                        daycount = 1
                        endcode = 2
                    elif (daycount == 0 and endcode == 1):
                        df3.loc[staff, day] = Ecode
                        daycount = 1
                        endcode = 2

        df3 = df3.dropna()

        newdf = df3

        return newdf

    def DataActuator(self, df3, roster):
        ## this function deletes columns of days to create desired number of days.

        currentdf = roster.copy(deep=True)
        currentdf.columns = currentdf.columns.astype(str)
        del currentdf["S.NO"]
        del currentdf["EMP.ID"]
        del currentdf["NAMES"]
        newreq = len(currentdf.columns.values)
        newx = len(df3.columns.values)

        if (newreq == 30):
            # print("loc of day31",df3.columns.get_loc("Day31"))
            df3 = df3.drop(df3.columns[30], axis=1)

            return df3
        elif (newreq == 29):

            df3 = df3.drop(df3.columns[30], axis=1)
            df3 = df3.drop(df3.columns[29], axis=1)

            return df3
        elif (newreq == 28):
            df3 = df3.drop(df3.columns[30], axis=1)
            df3 = df3.drop(df3.columns[29], axis=1)
            df3 = df3.drop(df3.columns[28], axis=1)

            return df3
        elif (newreq == 31):
            return df3

        else:
            return [print("Roster days out of scope!")]

    def HoursCalculator(self, df):
        ## this functions creates a column for hours worked for each staff.
        ## partially obsolete function

        # print('Calculating Hours....')
        staffs = range(len(df))
        days = list(df.columns.values)
        size = len(days)
        HoursList = []
        for staff in staffs:
            # print(df.iloc[0][3])
            # print(df.iloc[staff][size+3])
            mhours = (df.iloc[staff][size - 4]) * 7
            ehours = (df.iloc[staff][size - 3]) * 7
            nhours = (df.iloc[staff][size - 2]) * 10
            total_hours = (int(((mhours + ehours + nhours) / len(days)) * 30))
            HoursList.append(total_hours)
        # HoursList.extend([0,0,0,0,0,0])
        # print(len(HoursList))
        df['Hours/Mn'] = HoursList

        # print('Hours Calculated!')
        return df

    def NilMaker(self, df, tweek, nil):
        ## Only called if there is no previous month input
        ## for creating a new roster with no input of previous roster
        ## this function identifies all occupied cells to avoid assigning roster

        staffs = range(len(df))
        days = list(df.columns.values)
        i = 0
        # print('Identifying Planned Leaves.....')

        for staff in staffs:
            for day in days:
                if (type(df.loc[staff, day]) != str):
                    df.loc[staff, day] = nil
        # print('Planned leaves identified!')
        return df

    def Frequency(self, List):
        ## this function works to mainly display frequency of each shift...

        counter = 0
        num = List[0]

        for i in List:
            curr_frequency = List.count(i)
            if (curr_frequency > counter):
                counter = curr_frequency
                num = i
                Freq = List.count(i)

        return Freq

    def most_frequent(self, List):
        ## This function calculates the mode for summary dataframe/sheet

        counter = 0
        num = List[0]

        for i in List:
            curr_frequency = List.count(i)
            if (curr_frequency > counter):
                counter = curr_frequency
                num = i

        return num

    def codemaker(self, df, Mcode, Ecode, Ncode, NOcode, WOcode, nil):
        ##THis is function to assign shift codes for numbers determined in previous function rostermaker
        ## only gets activated when no previous file is given as input

        l = 0
        staffs = range(len(df))
        days = list(df.columns.values)
        for staff in staffs:
            if (l == 1):
                l = 0
            elif (l == 0):
                l = 1
            for day in days:
                if (df.loc[staff, day] == 1):
                    df.loc[staff, day] = Ecode
                elif (df.loc[staff, day] == 2):
                    df.loc[staff, day] = Mcode
                elif (df.loc[staff, day] == 4):
                    df.loc[staff, day] = Ncode

                elif (df.loc[staff, day] == 3):

                    if (l == 1):
                        df.loc[staff, day] = Ecode

                    else:
                        df.loc[staff, day] = Mcode

        return df

    def RosterMaker(self, df, tweek, Mcode, Ecode, Ncode, NOcode, WOcode, nil):
        ##This function works to create roster as numbers
        ##this function only gets activated when no previous file is given as input
        # print('Creating Roster.....')
        daycount = 0
        staffs = range(len(df))
        days = list(df.columns.values)
        shift_code = 1
        # iterating through dataframe
        for staff in staffs:
            for day in days:
                if (shift_code == 4):
                    if (daycount <= 6):
                        if (df.loc[staff, day] == nil):
                            df.loc[staff, day] = shift_code
                        daycount = daycount + 1
                    elif (daycount == 7):
                        df.loc[staff, day] = NOcode

                        daycount = daycount + 1
                    elif (daycount == 8):
                        df.loc[staff, day] = WOcode

                        daycount = daycount + 1
                        shift_code = 1
                        daycount = 0

                elif (shift_code == 1 or shift_code == 2):
                    if (daycount <= 5):
                        if (df.loc[staff, day] == nil):
                            df.loc[staff, day] = shift_code
                        daycount = daycount + 1
                    else:
                        df.loc[staff, day] = WOcode
                        daycount = 0
                        if (shift_code == 1):
                            shift_code = 2
                        elif (shift_code == 2):
                            shift_code = 3
                elif (shift_code == 3):
                    if (daycount < tweek):
                        if (df.loc[staff, day] == nil):
                            df.loc[staff, day] = shift_code
                        daycount = daycount + 1
                    else:
                        if (df.loc[staff, day] == nil):
                            df.loc[staff, day] = WOcode
                        daycount = 0
                        shift_code = 4

        df = self.codemaker(self, df, Mcode, Ecode, Ncode, NOcode, WOcode, nil)
        # print('Roster Prepared!')
        return df

    def CodeCounter(self, df, summary_switch, tweek, Mcode, Ecode, Ncode, NOcode, WOcode, nil):
        ##this function is for the console based view of each table as summary

        staffs = range(len(df))
        days = list(df.columns.values)
        # assigns alphanum codes
        M_count = 0  # chcker for morning count
        E_count = 0
        N_count = 0
        B_count = 0  # checker for weekoff count
        NO_count = 0
        for staff in staffs:  # iterates through staff, ie, rows.
            for day in days:  # iterates through days, ie columns
                if (df.loc[staff, day] == Ecode):  # checks if code is for morning
                    M_count = M_count + 1  # steps up morning code checker
                if (df.loc[staff, day] == Mcode):  # checks if code is for Evening
                    E_count = E_count + 1  # steps up evening code checker
                if (df.loc[staff, day] == Ncode):  # checks if code is for Night
                    N_count = N_count + 1  # steps up night code checker
                if (df.loc[staff, day] == WOcode):  # checks if code is for weekoff
                    B_count = B_count + 1  # steps up weekoff code checker
                if (df.loc[staff, day] == NOcode):
                    NO_count = NO_count + 1

        if (summary_switch == True):
            print("Number of staff Morning:", M_count)  # prints morning staff count
            print("Number of Staff Evening", E_count)  # prints evening staff count
            print("Number of staff Night:", N_count)  # prints night staff count
            print("Number of Week off:   ", B_count)  # prints weekoff staff count
            print("Number of Night off:  ", NO_count)  # prints weekoff staff count
            print("-------------------------------------")
        return df

    def Rsummary(self, df, tweek, Mcode, Ecode, Ncode, NOcode, WOcode, nil, Nreq, Mreq, Ereq):
        # creates summary of people in shifts for each day
        # partially obsolete after dynamic encoding to excel
        staffs = range(len(df))
        days = list(df.columns.values)

        # counters for each shift
        Mcount = 0
        Ecount = 0
        Ncount = 0
        Bcount = 0
        NOcount = 0
        LPcount = 0

        # list for each counts
        Mlist = []
        Elist = []
        Nlist = []
        Blist = []
        NOlist = []
        LPlist = []

        for day in days:  # iterating each day
            for staff in staffs:  # iterateing each day each staff
                if (df.loc[staff, day] == Mcode):  # checking if cell is monring staff
                    Mcount = Mcount + 1  # if yes, step up morning counter for the day
                if (df.loc[staff, day] == Ecode):  # checking if cell is Evening staff
                    Ecount = Ecount + 1  # if yes, step up evening counter for the day
                if (df.loc[staff, day] == Ncode):  # checking if cell is Night staff
                    Ncount = Ncount + 1  # if yes, step up night counter for the day
                if (df.loc[staff, day] == WOcode):  # checking if cell is Off staff
                    Bcount = Bcount + 1  # if yes, step up off counter for the day
                if (df.loc[staff, day] == NOcode):
                    NOcount = NOcount + 1
            Mlist.append(Mcount)  # add count of morning staff each day
            Elist.append(Ecount)  # add count of Evening staff each day
            Nlist.append(Ncount)  # add count of night staff each day
            Blist.append(Bcount)  # add count of off staff each day
            NOlist.append(NOcount)
            LPcount = (Mcount + Ecount + Ncount) - (Nreq + (Mreq) + (Ereq))
            LPlist.append(LPcount)
            Mcount = 0  # reseting morning counter after iterating each day
            Ecount = 0  # reseting evening counter after iterating each day
            Ncount = 0  # reseting night counter after iterating each day
            Bcount = 0  # reseting off counter after iterating each day
            NOcount = 0

        Mlist.pop(0)
        Elist.pop(0)
        Nlist.pop(0)
        NOlist.pop(0)
        Blist.pop(0)
        LPlist.pop(0)
        totlist = [["Morning", max(Mlist), min(Mlist), self.most_frequent(self, Mlist), self.Frequency(self, Mlist)],
                   ["Evening", max(Elist), min(Elist), self.most_frequent(self, Elist), self.Frequency(self, Elist)],
                   ["Night", max(Nlist), min(Nlist), self.most_frequent(self, Nlist), self.Frequency(self, Nlist)]]
        collist = list(df.columns.values)  # creating a new list of colums, for making index of series
        Mlist.insert(0, "Morning")  # inserting 'Morning' as first item so it can be added to df
        Mseries = pd.Series(Mlist, index=collist)  # creating a series with same index as df
        df = df.append(Mseries, ignore_index=True)  # adding series to df
        Elist.insert(0, "Evening")  # inserting 'evening' as first item so it can be added to df
        Eseries = pd.Series(Elist, index=collist)  # creating a series with same index as df
        df = df.append(Eseries, ignore_index=True)  # adding series to df
        Nlist.insert(0, "Night")  # inserting 'night' as first item so it can be added to df
        Nseries = pd.Series(Nlist, index=collist)  # creating a series with same index as df
        df = df.append(Nseries, ignore_index=True)  # adding series to df
        Blist.insert(0, "WeekOff")  # inserting 'off' as first item so it can be added to df
        Bseries = pd.Series(Blist, index=collist)  # creating a series with same index as df
        df = df.append(Bseries, ignore_index=True)  # adding series to df
        NOlist.insert(0, "NightOff")  # inserting 'Morning' as first item so it can be added to df
        NOseries = pd.Series(NOlist, index=collist)
        df = df.append(NOseries, ignore_index=True)
        LPlist.insert(0, "LP available")
        LPSeries = pd.Series(LPlist, index=collist)
        df = df.append(LPSeries, ignore_index=True)

        df2 = pd.DataFrame(totlist, columns=['Shift', 'Max', "Min", "Mode", "Freq."])

        return df, df2, Mlist

    def Csummary(self, df, Mcode, Ecode, Ncode, NOcode, WOcode, days):
        # creates summary as columns for each staff
        # only makes static summary... partially obsolete function
        days = list(df.columns.values)
        CMcount = 0  # staff wise counter for morning
        CEcount = 0  # staff wise counter for evening
        CNcount = 0  # staff wise counter for night
        CBcount = 0  # staff wise count for off
        CNOcount = 0  # staff wise count of night off
        CMlist = []  # list for morning duty of each staff
        CElist = []  # list for evening duty of each staff
        CNlist = []  # list for night duty of each staff
        CBlist = []  # list for week off duty of each staff
        CNOlist = []  # list for night off duty of each staff
        staffs = range(len(df))
        for staff in staffs:  # iterating through each staff
            for day in days:  # iterating thgouth each day of each staff
                if (df.loc[staff, day] == Mcode):  # checking if duty code is morning
                    CMcount = CMcount + 1  # increasing morning counter
                if (df.loc[staff, day] == Ecode):  # checking if duty code is eveing
                    CEcount = CEcount + 1  # increasing morning counter
                if (df.loc[staff, day] == Ncode):  # checking if duty code is night
                    CNcount = CNcount + 1  # increasing morning counter
                if (df.loc[staff, day] == WOcode):  # checking if duty code is off
                    CBcount = CBcount + 1  # increasing morning counter
                if (df.loc[staff, day] == NOcode):
                    CNOcount = CNOcount + 1
            CMlist.append(CMcount)  # adding each staff count of morning to list
            CMcount = 0  # reseting morning counter

            CElist.append(CEcount)  # adding each staff count of evening to list
            CEcount = 0  # reseting evening counter

            CNlist.append(CNcount)  # adding each staff count of night to list
            CNcount = 0  # reseting night counter

            CBlist.append(CBcount)  # adding each staff count of off to list
            CBcount = 0  # reseting off counter

            CNOlist.append(CNOcount)
            CNOcount = 0

        df["Morning"] = CMlist  # adding morning column
        df["Evening"] = CElist  # adding evening column
        df["Night"] = CNlist  # adding night column
        df["NightOff"] = CNOlist
        df["WeekOff"] = CBlist  # adding off column
        return df

    def ColorCoder(df):
        # obsolete function.... used for hardcoding colors to cells

        df = df.style.applymap(
            lambda x: 'background-color : pink' if x == 'N6' else '' or 'background-color : yellow' if x == 'NO' else ''
                                                                                                                      or 'background-color : orange' if x == 'WO' else '')
        return df

    def ExcelDesigner(self, df, df2, tweek, days, Mreq, Ereq, Nreq, Mcode, Ecode, Ncode, NOcode, WOcode, nil):
        ##This function handles all the design element of df when transfering to excel...
        ##should usually be the last function to be called

        reqfile = pathlib.Path("testreq.xls")
        if reqfile.exists():
            req = pd.read_excel("testreq.xls", sheet_name="Requirement")
            os.remove("testreq.xls")

        else:
            req = pd.read_excel("roster.xls", sheet_name="Requirement")
        readme = pd.read_excel("roster.xls", sheet_name="ReadMe")
        df = pd.read_excel('rosterupdate.xlsx')

        writer = pd.ExcelWriter('rosterupdate.xlsx', engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Update', startrow=1, index=None)
        df2.to_excel(writer, sheet_name='Summary', startrow=0, index=None)
        # tweek_row = {'Shift':'TweekID', 'Count':tweek}

        req.to_excel(writer, sheet_name='Requirement', startrow=0, index=None)

        ################################################################################
        ## converting column names back to int format

        col_names = list(df.columns.values)
        # print(col_names)
        for i in range(len(col_names)):
            for j in range(32):
                if col_names[i] == str(j):
                    col_names[i] = int(col_names[i])
        # print('new col_names',col_names)
        df.columns = col_names
        # print('columns names: ', df.columns.values)

        ###############################################################################

        df.to_excel(writer, sheet_name='Duplicate', startrow=1, index=None)
        readme.to_excel(writer, sheet_name="ReadMe", startrow=0, index=None)
        # print(df)

        # df=ColorCoder(df)
        workbook = writer.book
        worksheet = writer.sheets['Update']

        # adding equations to excel as each column summary
        Mlendf = len(df) - 3
        Elendf = len(df) - 2
        Nlendf = len(df) - 1
        WOlendf = len(df)
        NOlendf = len(df) + 1
        LPlendf = len(df) + 2
        staffcounter = len(df) - 4  # this minus value should be 1+ than minus value of mlendf

        i = 0
        listrows = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R',
                    'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF',
                    'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS']

        M_formula_area = "%s%s" % (listrows[i], str(Mlendf))
        M_formula_equation = '=COUNTIF(%s%s%s%s%s%s )' % (listrows[i], "2:", listrows[i], staffcounter, ',', Ecode)
        for i in range(len(days) + 3):
            M_formula_area = "%s%s" % (listrows[i], str(Mlendf))
            M_formula_equation = '=COUNTIF(%s%s%s%s%s%s%s )' % (
                listrows[i], "2:", listrows[i], staffcounter, ',"', Mcode, '"')
            worksheet.write_formula(M_formula_area, M_formula_equation)
            E_formula_area = "%s%s" % (listrows[i], str(Elendf))
            E_formula_equation = '=COUNTIF(%s%s%s%s%s%s%s)' % (
                listrows[i], "2:", listrows[i], staffcounter, ',"', Ecode, '"')
            worksheet.write_formula(E_formula_area, E_formula_equation)
            N_formula_area = "%s%s" % (listrows[i], str(Nlendf))
            N_formula_equation = '=COUNTIF(%s%s%s%s%s%s%s)' % (
                listrows[i], "2:", listrows[i], staffcounter, ',"', Ncode, '"')
            worksheet.write_formula(N_formula_area, N_formula_equation)
            WO_formula_area = "%s%s" % (listrows[i], str(WOlendf))
            WO_formula_equation = '=COUNTIF(%s%s%s%s%s%s%s)' % (
                listrows[i], "2:", listrows[i], staffcounter, ',"', WOcode, '"')
            worksheet.write_formula(WO_formula_area, WO_formula_equation)
            NO_formula_area = "%s%s" % (listrows[i], str(NOlendf))
            NO_formula_equation = '=COUNTIF(%s%s%s%s%s%s%s)' % (
                listrows[i], "2:", listrows[i], staffcounter, ',"', NOcode, '"')
            worksheet.write_formula(NO_formula_area, NO_formula_equation)
            LP_formula_area = "%s%s" % (listrows[i], str(LPlendf))
            LP_formula_equation = "=sum(%s%s%s%s%s)-(%s+%s+%s)" % (
                listrows[i], LPlendf - 5, ":", listrows[i], LPlendf - 3, Nreq, Mreq, Ereq)
            worksheet.write_formula(LP_formula_area, LP_formula_equation)
            i = i + 1

        # adding equations at the end as row summary
        j = 3
        staffnum = 3
        if len(days) == 31:
            Mcol = "AI"
            Ecol = "AJ"
            Ncol = "AK"
            NOcol = "AL"
            WOcol = "AM"
            Hcol = "AN"
            cellend = "AH"

        elif len(days) == 30:
            Mcol = "AH"
            Ecol = "AI"
            Ncol = "AJ"
            NOcol = "AK"
            WOcol = "AL"
            Hcol = "AM"
            cellend = "AG"

        elif len(days) == 29:
            Mcol = "AG"
            Ecol = "AH"
            Ncol = "AI"
            NOcol = "AJ"
            WOcol = "AK"
            Hcol = "AL"
            cellend = "AF"

        elif len(days) == 28:
            Mcol = "AF"
            Ecol = "AG"
            Ncol = "AH"
            NOcol = "AI"
            WOcol = "AJ"
            Hcol = "AK"
            cellend = "AE"

        else:
            print("Days not matching!")
            print(len(days))
            print(days)

        MH_formula_area = "%s%s" % (Mcol, j)
        MH_formula_equation = "=countif(%s%s%s%s%s%s)" % ("D", j, ":", cellend, j, ',"M6"')
        for j in range(len(df) - 2):
            if (j >= 3):
                MH_formula_area = "%s%s" % (Mcol, j)
                MH_formula_equation = "=countif(%s%s%s%s%s%s%s%s)" % ("D", j, ":", cellend, j, ',"', Mcode, '"')
                worksheet.write_formula(MH_formula_area, MH_formula_equation)
                EH_formula_area = "%s%s" % (Ecol, j)
                EH_formula_equation = "=countif(%s%s%s%s%s%s%s%s)" % ("D", j, ":", cellend, j, ',"', Ecode, '"')
                worksheet.write_formula(EH_formula_area, EH_formula_equation)
                NH_formula_area = "%s%s" % (Ncol, j)
                NH_formula_equation = "=countif(%s%s%s%s%s%s%s%s)" % ("D", j, ":", cellend, j, ',"', Ncode, '"')
                worksheet.write_formula(NH_formula_area, NH_formula_equation)
                NOH_formula_area = "%s%s" % (NOcol, j)
                NOH_formula_equation = "=countif(%s%s%s%s%s%s%s%s)" % ("D", j, ":", cellend, j, ',"', NOcode, '"')
                worksheet.write_formula(NOH_formula_area, NOH_formula_equation)
                WOH_formula_area = "%s%s" % (WOcol, j)
                WOH_formula_equation = "=countif(%s%s%s%s%s%s%s%s)" % ("D", j, ":", cellend, j, ',"', WOcode, '"')
                worksheet.write_formula(WOH_formula_area, WOH_formula_equation)
                HH_formula_area = "%s%s" % (Hcol, j)
                HH_formula_equation = "=%s%s*7+%s%s*7+%s%s*10" % (Mcol, j, Ecol, j, Ncol, j)
                worksheet.write_formula(HH_formula_area, HH_formula_equation)

        Nvalue = '="%s"' % (Ncode)
        NOvalue = '="%s"' % (NOcode)
        WOvalue = '="%s"' % (WOcode)
        format1 = workbook.add_format({'bg_color': '#FFC7CE',
                                       'font_color': '#9C0006'})

        # Light yellow fill with dark yellow text.
        format2 = workbook.add_format({'bg_color': '#FFEB9C',
                                       'font_color': '#9C6500'})

        # Green fill with dark green text.
        format3 = workbook.add_format({'bg_color': '#C6EFCE',
                                       'font_color': '#006100'})

        worksheet.conditional_format('D3:XF200', {'type': 'cell',
                                                  'criteria': '=',
                                                  'value': Nvalue,
                                                  'format': format1
                                                  })
        worksheet.conditional_format('D3:XF200', {'type': 'cell',
                                                  'criteria': '=',
                                                  'value': NOvalue,
                                                  'format': format2
                                                  })
        worksheet.conditional_format('D3:XF200', {'type': 'cell',
                                                  'criteria': '=',
                                                  'value': WOvalue,
                                                  'format': format3
                                                  })

        ## freezing pane at d3
        worksheet.freeze_panes('D3')
        writer.save()

    def main(self, roster, previous_roster, req_location, tweek_location):

        df = roster.copy(deep=True)
        df.columns = df.columns.astype(str)

        print("File Detected!")

        NameList = list(df['NAMES'])
        Emplist = list(df['EMP.ID'])
        Slno = list(df['S.NO'])

        df = df.drop(columns=['NAMES'])
        df = df.drop(columns=['EMP.ID'])
        df = df.drop(columns=['S.NO'])
        rows = len(df)
        days = list((df.columns.values))
        # print(days)
        print("Number of staff: ", rows)
        print("Number of days: ", len(days))

        Nreq, Mreq, Dreq, Ereq, ptweek = self.ReqReader(self, req_location, tweek_location)
        print('this is the "ptweek"', ptweek)
        print('Nreq', Nreq)
        tweek = int
        looper = '1'
        pexit = ''
        Ecode = 'A4'
        Mcode = 'M6'
        Ncode = 'N6'
        NOcode = 'NO'
        WOcode = 'WO'
        nil = 'Nil'
        printer_switch = 1  # a switch to avoid multiple prints
        printer2_switch = 1  # a switch to avoid multiple prints
        printer3_switch = 1  # a switch to avoid multiple prints
        summary_switch = True
        elifticket = 0
        Error_switch = False
        filesave_switch = False
        pdf_file = False
        for tweek in range(2, 8):

            # saving names into a list and deleting names column so we get 2d array of just dates and roster
            if (tweek < 10):

                df = roster.copy(deep=True)

                df.columns = df.columns.astype(str)
                df.sort_values(by=['EMP.ID'], inplace=True)
                df.reset_index(inplace=True)
                del df['index']  # deletes the column created of old index after reset_index()
                NameList = list(df['NAMES'])
                Emplist = list(df['EMP.ID'])
                Slno = list(df['S.NO'])
                df = df.drop(columns=['NAMES'])
                df = df.drop(columns=['EMP.ID'])
                df = df.drop(columns=['S.NO'])

                days = list(df.columns.values)
                staffs = range(len(df))

                df = self.NilMaker(self, df, tweek, nil)
                # calling function to create roster
                df = self.RosterMaker(self, df, tweek, Mcode, Ecode, Ncode, NOcode, WOcode, nil)
                pdf = previous_roster.copy(deep=True)

                if pdf.empty == False:
                    pdf.columns = pdf.columns.astype(str)

                    print("previous file found")
                    tweek = ptweek
                    pdf_file = True

                    ###########################################################
                    # Checking length of two files
                    df, pdf, NOT_df, NameList, Emplist, Slno, days, staffs = self.staffactuator(self, df, pdf,
                                                                                                previous_roster, roster)
                    # print(pdf)

                    #########################################################

                    NOT_df, New_NameList, New_Emplist, New_Slno = self.NewStaffCleaner(self, NOT_df, tweek, Mcode,
                                                                                       Ecode, Ncode, NOcode, WOcode,
                                                                                       nil)

                    npdf, plist = self.ShiftReversor(self, pdf, Mcode, Ecode, Ncode, NOcode, WOcode)

                    df3 = npdf.copy(deep=True)

                    df3 = self.NewShiftMaker(self, npdf, plist, tweek, df3, Mcode, Ecode, Ncode, NOcode, WOcode, nil)
                    blankdf = df.copy(deep=True)

                    ##in next function, df3 is assigned as df, and df is assigned as blankdf, confusing, pls note
                    df3 = self.newcodemaker(self, df3, blankdf, tweek, Mcode, Ecode, Ncode, NOcode, WOcode, nil)

                    df3 = self.DataActuator(self, df3, roster)

                    df3 = self.newCLmaker(self, df3, tweek, Mcode, Ecode, Ncode, NOcode, WOcode, nil, roster)

                    df = df3.copy(deep=True)

                    df = df.append(NOT_df, ignore_index=True, sort=False)
                    Slno = Slno + New_Slno  # combining new and old slno
                    Emplist = Emplist + New_Emplist
                    NameList = NameList + New_NameList

                df = self.CodeCounter(self, df, summary_switch, tweek, Mcode, Ecode, Ncode, NOcode, WOcode, nil)

                # print(df)

                # inserting again the nameslist previously deleted into the first column

                df.insert(loc=0, column='S.NO', value=Slno)
                # print(df)

                # creating row summary
                df, df2, Mlist = self.Rsummary(self, df, tweek, Mcode, Ecode, Ncode, NOcode, WOcode, nil, Nreq, Mreq,
                                               Ereq)

                print("Tweek:", tweek)
                print(tabulate(df2, headers='keys', tablefmt='psql', showindex=False))
                print("Nreq:", Nreq, "Dreq:", Dreq)
                if ((df2.iloc[2][2] >= Nreq) and (df2.iloc[0][2] >= Dreq)) or pdf_file == True:

                    match_switch = True
                    print("Match found")
                    # print(tweek)
                    summary_switch = True
                else:
                    match_switch = False
                    summary_switch = False

                if match_switch == True and filesave_switch == False:
                    for i in range(6):
                        Emplist.append('')
                        NameList.append('')
                    df.insert(loc=1, column='EMP.ID', value=Emplist)
                    df.insert(loc=2, column='NAMES', value=NameList)

                    df = self.Csummary(self, df, Mcode, Ecode, Ncode, NOcode, WOcode, days)

                    df = self.HoursCalculator(self, df)
                    df.to_excel("rosterupdate.xlsx", index=None)
                    filesave_switch = True

                    self.ExcelDesigner(self, df, df2, tweek, days, Mreq, Ereq, Nreq, Mcode, Ecode, Ncode, NOcode,
                                       WOcode, nil)
                    break

        if filesave_switch == False and pdf_file == False:
            df, df2 = self.DayJuggler(self, df, days, Mreq, Ereq, Nreq, Mcode, Ecode, Ncode, NOcode, WOcode, nil, Dreq,
                                      roster)

            print("Try changing staff requirment in Requirement sheet")
        print("Press any key to exit")

        return df, tweek, df2, days
        # input (pexit)


class preprocessor:

    def DataLoader():
        # this function loads the data into flow

        roster = pd.read_excel("roster.xls", skiprows=1)
        file = pathlib.Path("previous_roster.xls")
        if file.exists():
            previous_roster = pd.read_excel("previous_roster.xls", sheet_name="Duplicate", skiprows=1)
        else:
            previous_roster = pd.DataFrame(columns=roster.columns.values)

        return roster, previous_roster

    def seperator(roster, previous_roster):
        # this function superates the roster into different rosters based on designnation

        TL = ['TL', 'tl']  # teamlead
        SN = ['S', 's']  # Senior nurse
        JR = ['J', 'j']
        other = ['c', 'C', 'O', "o"]

        TL_roster = roster[roster['S.NO'].isin(TL)]
        SN_roster = roster[roster['S.NO'].isin(SN)]
        JR_roster = roster[roster['S.NO'].isin(JR)]
        Other_roster = roster[roster['S.NO'].isin(other)]

        if previous_roster.empty == False:
            PTL_roster = previous_roster[previous_roster['S.NO'].isin(TL)]
            PSN_roster = previous_roster[previous_roster['S.NO'].isin(SN)]
            PJR_roster = previous_roster[previous_roster['S.NO'].isin(JR)]
            POther_roster = previous_roster[previous_roster['S.NO'].isin(other)]
        else:
            PTL_roster = pd.DataFrame(columns=TL_roster.columns.values)
            PSN_roster = pd.DataFrame(columns=SN_roster.columns.values)
            PJR_roster = pd.DataFrame(columns=JR_roster.columns.values)
            POther_roster = pd.DataFrame(columns=Other_roster.columns.values)

        return TL_roster, SN_roster, JR_roster, Other_roster, PTL_roster, PSN_roster, PJR_roster, POther_roster

    def Total_Summary(TL_summary, SN_summary, JR_summary, Other_summary, indexlist):

        # from the summary returns for roster of each design, a total summary in created.

        Tot_summary = TL_summary.append(SN_summary, ignore_index=False, sort=False)
        Tot_summary = Tot_summary.append(JR_summary, ignore_index=False, sort=False)
        Tot_summary = Tot_summary.append(Other_summary, ignore_index=False, sort=False)
        Tot_summary.insert(loc=0, column='Desig', value=indexlist)

        return Tot_summary

    def ReqWriter(TL_tweek, SN_tweek, JR_tweek):

        # writes the new tweek ids into req df, to be written to requirement sheet later
        req = pd.read_excel("roster.xls", sheet_name="Requirement")
        req.iloc[3, 1] = TL_tweek
        req.iloc[3, 2] = SN_tweek
        req.iloc[3, 3] = JR_tweek
        Mreq = req.iloc[0][1] + req.iloc[0][2] + req.iloc[0][3]
        Ereq = req.iloc[1][1] + req.iloc[1][2] + req.iloc[1][3]
        Nreq = req.iloc[2][1] + req.iloc[2][2] + req.iloc[2][3]
        req.to_excel('testreq.xls', sheet_name='Requirement', index=None)
        return req, Mreq, Ereq, Nreq

    def Total_Roster(TL_roster, SN_roster, JR_roster, Other_roster):
        # this function clubs all rosters together
        TL_roster = TL_roster.loc[TL_roster['S.NO'] != 'Morning']
        TL_roster = TL_roster.loc[TL_roster['S.NO'] != 'Evening']
        TL_roster = TL_roster.loc[TL_roster['S.NO'] != 'Night']
        TL_roster = TL_roster.loc[TL_roster['S.NO'] != 'WeekOff']
        TL_roster = TL_roster.loc[TL_roster['S.NO'] != 'NightOff']
        TL_roster = TL_roster.loc[TL_roster['S.NO'] != 'LP available']

        SN_roster = SN_roster.loc[SN_roster['S.NO'] != 'Morning']
        SN_roster = SN_roster.loc[SN_roster['S.NO'] != 'Evening']
        SN_roster = SN_roster.loc[SN_roster['S.NO'] != 'Night']
        SN_roster = SN_roster.loc[SN_roster['S.NO'] != 'WeekOff']
        SN_roster = SN_roster.loc[SN_roster['S.NO'] != 'NightOff']
        SN_roster = SN_roster.loc[SN_roster['S.NO'] != 'LP available']

        JR_roster = JR_roster.loc[JR_roster['S.NO'] != 'Morning']
        JR_roster = JR_roster.loc[JR_roster['S.NO'] != 'Evening']
        JR_roster = JR_roster.loc[JR_roster['S.NO'] != 'Night']
        JR_roster = JR_roster.loc[JR_roster['S.NO'] != 'WeekOff']
        JR_roster = JR_roster.loc[JR_roster['S.NO'] != 'NightOff']
        JR_roster = JR_roster.loc[JR_roster['S.NO'] != 'LP available']

        Total_Roster = TL_roster.append(SN_roster, ignore_index=True, sort=False)
        Total_Roster = Total_Roster.append(JR_roster, ignore_index=True, sort=False)
        Total_Roster = Total_Roster.append(Other_roster, ignore_index=True, sort=False)

        return Total_Roster

    def main():

        print("__Scripting By Hisham__")

        print("Create a 'roster.xls file' days in the top row (2) and names of staff in first ")
        print("column in the same folder as the script. The Script will take care of the rest.\n")
        print("Detecting file.....")

        # declaring expected location of requirment
        TL_req = 1
        SN_req = 2
        JR_req = 3

        # declaring expected location of tweekids
        TL_tweekid = 1
        SN_tweekid = 2
        JR_tweekid = 3

        # initialting roster codes within class preprocessor

        Mcode = 'M6'
        Ecode = 'A4'
        Ncode = 'N6'
        NOcode = 'NO'
        WOcode = 'WO'
        nil = 'Nil'
        pexit = ''

        roster, previous_roster = preprocessor.DataLoader()
        TL_roster, SN_roster, JR_roster, Other_roster, PTL_roster, PSN_roster, PJR_roster, POther_roster = preprocessor.seperator(
            roster, previous_roster)

        TL_roster, TL_tweek, TL_summary, days = rota.main(rota, TL_roster, PTL_roster, TL_req, TL_tweekid)
        SN_roster, SN_tweek, SN_summary, days = rota.main(rota, SN_roster, PSN_roster, SN_req, SN_tweekid)
        JR_roster, JR_tweek, JR_summary, days = rota.main(rota, JR_roster, PJR_roster, JR_req, JR_tweekid)
        Other_roster, Other_tweek, Other_summary, days = rota.main(rota, Other_roster, POther_roster, 0, 0)

        print("printing all tweeks")
        print(TL_tweek, SN_tweek, JR_tweek)
        # print(Other_summary)

        indexlist = ['Team Leads', 'Team Leads', 'Team Leads', 'Seniors', 'Seniors', 'Seniors',
                     'Juniors', 'Juniors', 'Juniors', 'Others', 'Others', 'Others']
        Tot_summary = preprocessor.Total_Summary(TL_summary, SN_summary, JR_summary, Other_summary, indexlist)
        req, Mreq, Ereq, Nreq = preprocessor.ReqWriter(TL_tweek, SN_tweek, JR_tweek)
        Tot_roster = preprocessor.Total_Roster(TL_roster, SN_roster, JR_roster, Other_roster)

        Tot_roster.to_excel('rosterupdate.xlsx', index=None)
        rota.ExcelDesigner(rota, Tot_roster, Tot_summary, 1222, days, Mreq, Ereq, Nreq, Mcode,
                           Ecode, Ncode, NOcode, WOcode, nil)
        print("Press any key to exit")
        #input(pexit)


#preprocessor.main()




