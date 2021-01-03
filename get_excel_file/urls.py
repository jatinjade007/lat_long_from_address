from django.urls import path

from . import views

app_name = "get_excel_file"

urlpatterns = [
    path('', views.index, name='index'),
]
