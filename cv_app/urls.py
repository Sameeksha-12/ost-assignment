from django.urls import path
from . import views

urlpatterns = [
    path('', views.upload_files, name='upload_files'),
    path('download_excel/', views.download_excel, name='download_excel'),
]
