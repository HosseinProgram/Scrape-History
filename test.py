import requests
from pprintpp import pprint
headers={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36 Edg/109.0.1518.61","Accept-Language":"en-US,en;q=0.9,fa;q=0.8","Connection": "keep-alive","Cookie": "ASP.NET_SessionId=kislljlalcplvzmn2q2ycni0"}

r=requests.get('http://cdn.tsetmc.com/api/BestLimits/46348559193224090/20230215',headers=headers)
pprint(r.json())