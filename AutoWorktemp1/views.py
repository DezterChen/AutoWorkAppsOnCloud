
from django.http import HttpResponse, JsonResponse
from django.template import RequestContext
from django.shortcuts import render, render_to_response
import sys
"""
def index(requet):
    return HttpResponse("Hello, world. You're at the temp1 index.")
"""
def hello(request):

    return render(request, 'AutoWorktemp1/hello.html', {
        'data': 'BOM Generator',
    })

"""
def hello(request):
    if request.method == 'GET':
        return render(request, 'AutoWorktemp1/hello.html', {
            'data': 'Hello Django ',
        })
    elif request.method == 'POST':
        sys.path.append('D:/Python/ForWork/BomGenerator')
        import SubTool_BomGenerator as SB_BG
        py_obj = SB_BG.ST_BomGenerator()
        return render(request, 'AutoWorktemp1/output.html', {'output': py_obj})
"""

def query(request):
    sys.path.append('F:/python/ForWork/AutoWorkAPPonCloud/AutoWorktemp1/AutoWorkApps')
    import SubTool_BomGenerator as SB_BG
    py_obj = SB_BG.ST_BomGenerator()
    return JsonResponse(py_obj, safe=False)


"""
def BomGenerator(request):
    sys.path.append('D:/Python/ForWork/BomGenerator')
    import SubTool_BomGenerator as SB_BG
    SB_BG.ST_BomGenerator()
    return render(request, 'AutoWorktemp1/hello.html')
"""    
def detail(request, question_id):
    return HttpResponse("You're looking at question %s." % question_id)

def results(request, question_id):
    response = "You're looking at the results of question %s."
    return HttpResponse(response % question_id)

def vote(request, question_id):
    return HttpResponse("You're voting on question %s." % question_id)

def index(request):
    return render(request, 'AutoWorktemp1/index.html')
    
def add(request):
    a = request.GET['a']
    b = request.GET['b']
    a = int(a)
    b = int(b)
    result = str(a+b)
    return JsonResponse(result, safe=False)


def index2(request):
# get context of request from client
    context = RequestContext(request)
    # construct dictionary to pass template + context
    context_dict = {'buildingName': 'The Building',
                    'boldmessage': 'Put a message here'}

#render and return to client
    return render_to_response('AutoWorktemp1/index2.html', context_dict, context)   

def chart_data(request):

    if (request.method == 'POST'):
        dataX = [0,10,20,30,40,50,60]
        dataY = [25.0,24.2,25,24.0,24.5,25.1,25.5]
    
        response = {"x": dataX,
                    "y": dataY}
    
    return JsonResponse(response)


