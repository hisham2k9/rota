from django.shortcuts import render
from django.http import HttpResponse
import pandas as pd
import os
import datetime
import pathlib
from django.core.files import File
from django.utils.encoding import smart_str
from .rota import *




# Create your views here.

def rotahome(request):
    x= 'BMH'
    return render(request, 'rota/rotahome.html', {'intro':x})

def home2(request):
    x="Please choose the Application"

    return render(request, 'rota/home2.html', {'select': x})
def download(request):

    excel_file1 = request.FILES["roster"]if 'roster' in request.FILES else False
    excel_file2 = request.FILES["previous_roster"]if 'previous_roster' in request.FILES else False
    print("excel file 1:", excel_file1)
    print("excel file 2:", excel_file2)
    if excel_file1 != False and excel_file2==False:
        pfile = pathlib.Path("previous_roster.xls")
        if pfile.exists():
            os.remove("previous_roster.xls")

        roster = pd.ExcelFile(excel_file1)
        df = pd.read_excel(roster, sheet_name='Update', skiprows=1)
        req = pd.read_excel(roster, sheet_name='Requirement')
        sum = pd.read_excel(roster, sheet_name='Summary')
        re = pd.read_excel(roster, sheet_name='ReadMe')
        dup = pd.read_excel(roster, sheet_name='Duplicate')
        roster = pd.ExcelWriter('roster.xls', engine='xlsxwriter')
        df.to_excel(roster, sheet_name='Update', startrow=1, index=None)
        sum.to_excel(roster, sheet_name='Summary', index=None)
        req.to_excel(roster, sheet_name="Requirement", index=None)
        dup.to_excel(roster, sheet_name='Duplicate', startrow=1, index=None)
        re.to_excel(roster, sheet_name="ReadMe", index=None)
        roster.save()

        print(rota)
        self = rota
        
        preprocessor.main()

        try:
            os.remove('roster.xls')
        except:
            pass
        path_to_file = os.path.realpath("rosterupdate.xlsx")
        with open(path_to_file, 'rb') as excel:
            file = excel.read()

        response = HttpResponse(file, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = 'attachment; filename=New_roster_{date}.xlsx'.format(
            date=datetime.datetime.now().strftime('%d_%m_%Y'))
        return response

    elif excel_file1==False and (excel_file2 !=False or excel_file2==False):
        return render(request, 'rota/download.html', {'result': 'No empty roster detected'})

    elif excel_file2 != False and excel_file1 !=False:
        print('hello')
        roster = pd.ExcelFile(excel_file1)
        df = pd.read_excel(roster, sheet_name='Update', skiprows=1)
        req = pd.read_excel(roster, sheet_name='Requirement')
        sum = pd.read_excel(roster, sheet_name='Summary')
        re = pd.read_excel(roster, sheet_name='ReadMe')
        dup = pd.read_excel(roster, sheet_name='Duplicate')
        roster = pd.ExcelWriter('roster.xls', engine='xlsxwriter')
        df.to_excel(roster, sheet_name='Update', startrow=1, index=None)
        sum.to_excel(roster, sheet_name='Summary', index=None)
        req.to_excel(roster, sheet_name="Requirement", index=None)
        dup.to_excel(roster, sheet_name='Duplicate', startrow=1, index=None)
        re.to_excel(roster, sheet_name="ReadMe", index=None)
        roster.save()

        previous_roster=pd.ExcelFile(excel_file2)

        pdf1=pd.read_excel(previous_roster, sheet_name='Update', skiprows=1)
        psum=pd.read_excel(previous_roster, sheet_name='Summary')
        preq = pd.read_excel(previous_roster, sheet_name='Requirement')
        pdf=pd.read_excel(previous_roster, sheet_name='Duplicate', skiprows=1)
        readme=pd.read_excel(previous_roster,sheet_name='ReadMe')

        previous_roster=pd.ExcelWriter('previous_roster.xls', engine='xlsxwriter')
        pdf1.to_excel(previous_roster, sheet_name='Update', startrow=1, index= None)
        psum.to_excel(previous_roster,sheet_name='Summary', index=None)
        preq.to_excel(previous_roster, sheet_name='Requirement', index=None)
        pdf.to_excel(previous_roster, sheet_name= 'Duplicate', startrow=1, index=None)
        readme.to_excel(previous_roster, sheet_name='ReadMe', index=None)
        previous_roster.save()
        self = rota
        preprocessor.main()
        path_to_file = os.path.realpath("rosterupdate.xlsx")
        with open(path_to_file, 'rb') as excel:
            file = excel.read()

        try:
            os.remove("roster.xls")
            os.remove("previous_roster.xls")
        except:
            pass

        response = HttpResponse(file, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = 'attachment; filename=New_roster_{date}.xlsx'.format(
            date=datetime.datetime.now().strftime('%d_%m_%Y'))
        return response

    elif excel_file2 !=False and excel_file1==False:
        return render(request, 'rota/download.html', {'result': 'No empty roster detected'})
    else:
        return render(request, 'rota/download.html', {'result': 'No file detected'})



def template(request):
    print('hellow template')
    path_to_file = os.path.realpath("Template.xlsx")
    with open(path_to_file, 'rb') as excel:
        file = excel.read()
    response = HttpResponse(file, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename=Template.xlsx'
    return response









