from django.urls import path

from . import views

urlpatterns=[
    path('home2', views.home2, name='home'),
    path('rota', views.rotahome, name = 'rotahome'),
    path('download', views.download, name = 'download')
]