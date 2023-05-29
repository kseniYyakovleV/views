from django.shortcuts import render
from django.http import HttpResponse
from django.http import FileResponse
from os.path import abspath
from django.shortcuts import render
from rest_framework.response import Response
from rest_framework import generics
from .models import Spare_parts as sp

# Create your views here.
def home_page(request):
    "home page"
    print(request.build_absolute_uri())
    if request.method == "POST":
        return render(request, 'home.html', {"new_item_text": request.POST["item_text"]})
    return render(request, "home.html")

def get_file(request):
    response = FileResponse(open(abspath("SW_Repin_mart_2023.xlsx"), "rb"))
    return response

def load_excel_file_exe(request):
    with open (abspath("Load_Excel_File.exe"), "rb")as my_application:
        response = HttpResponse(my_application.read(), headers = {
            "Content-Type": "application/vnd.microsoft.portable-executable",
            "Content-Disposition": "attachment; filename = Load_Excel_File.exe"})
        return response

def load_excel_file(request):
    with open(abspath("SW_Repin_mart_2023.xlsx"), "rb") as file:
        my_data = file.read()
        response = HttpResponse(my_data, headers = {
            "Content-Type": "application/vnd.ms-excel",
            "Content-Disposition": "attachment; filename = 19_04_2023.xlsx"})
        return response
    
def load_apk_file(request):
    with open(abspath("game.apk"), "rb") as file:
        data = file.read()
        response = HttpResponse(data, headers = {
            "Content-Type": "application/vnd.android.package-archive",
            "Content-Disposition": "attachment; filename = game.apk"})
        return response
    
def load_image(request):
    image_id = request.GET["image"]
    print(image_id)
    with open(abspath("lists/images/"+image_id+".png"), "rb") as image:
        data = image.read()
        response = HttpResponse(data, headers = {
            "Content-Type": "image/png",
            "Content-Disposition": "attachment; filename = "+image_id+".png"})
        return response
    

def show_image(request):
    image_id = request.GET["image"]
    response = FileResponse(open(abspath("lists/images/"+image_id+".png"), "rb"))
    return response
    
class Items_list(generics.ListAPIView):
    def get(self, request):
        return Response(sp.get_all())
    


class One_item(generics.GenericAPIView):
    def get(self, request):
        id = request.GET["id"]
        item = sp.objects.filter(id = id)[0]
        return Response(item.get_full_info())
    


def change_items_count(request):
    if request.method=="GET":
        item_id = request.GET["id"]
        count_difference = request.GET["difference"]
        print(item_id, count_difference)
        print(sp.get_all())
        return HttpResponse("Yes!")

    