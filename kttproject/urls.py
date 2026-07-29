"""
URL configuration for kttproject project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/6.0/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.urls import path, include 
from kttapp import views
from Inpayment import views
from Innonpayment import views
from Out import views
from Transhipment import views
from Coo import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('kttapp.urls')),
    path('inpayment/', include('Inpayment.urls')),
    path('innonpayment/', include('Innonpayment.urls')),
    path('out/', include('Out.urls')),
    path('transhipment/', include('Transhipment.urls')),
    path('coo/', include('Coo.urls')),
]
