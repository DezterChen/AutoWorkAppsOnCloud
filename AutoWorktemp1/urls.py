# -*- coding: utf-8 -*-
"""
Created on Sat Jun 23 21:28:50 2018

@author: Dezter
"""
from django.urls import path
#from django.conf.urls import url
import AutoWorktemp1.views as views

urlpatterns = [
    path('hello/', views.hello, name='hello'),
    path('query/', views.query, name='query'),   
    #path('', views.index, name='index'),
    path('<int:question_id>/', views.detail, name='detail'),
    # ex: /polls/5/results/
    path('<int:question_id>/results/', views.results, name='results'),
    # ex: /polls/5/vote/
    path('<int:question_id>/vote/', views.vote, name='vote'),
    path('index/', views.index, name='index'),    
    path('add/',views.add, name='add'),
    path('', views.index2, name='index2'),
    path('chart_data', views.chart_data, name='chart_data')
    
]
