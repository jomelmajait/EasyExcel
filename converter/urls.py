from django.urls import path
from . import views

urlpatterns = [
    path('', views.upload_view, name='upload'),
    path('convert/', views.convert_image, name='convert_image'),
]