from django.shortcuts import render
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt

import json

@csrf_exempt
def uploadImgs(request):
    if request.method == 'POST':
        data = {}
        segments_list = []
        segments_1 = {}
        segments_1["coordinates"] = [0,0,100,50]
        segments_1["type"] = "Text"
        segments_1["text"] = "Test Text 1"
        segments_list.append(segments_1)
        segments_2 = {}
        segments_2["coordinates"] = [80,0,200,50]
        segments_2["type"] = "Text"
        segments_2["text"] = "Test Text 2"
        segments_list.append(segments_2)
        data["segments"] = segments_list
        data["name"] = "TestPng"
        return JsonResponse(data, safe=False)
    else:
        return JsonResponse({'message': 'Invalid request method'})