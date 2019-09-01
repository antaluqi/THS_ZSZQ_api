import ZSZQ
import pprint,datetime,random
from retrying import retry

api=ZSZQ.API()
r=api.get_history_entrusts('2018-01-01','2018-01-05')
pprint.pprint(r)

'''  retrying '''
'''
def run(s):
    return not (s%7==0)


@retry(wait_fixed=1000,stop_max_attempt_number=3,retry_on_result=run,wrap_exception=True)
def make_trouble():
    print(datetime.datetime.now().second)
    return datetime.datetime.now().second


#修饰器

def ret(func):
  r=1
  x=10
  while r%5!=0:
      x=x*10
      r = func(x)
      print(r)
      continue
  print('out')

@ret
def aa(x=100):
    r=int(random.random() * x)
    return r

aa


'''
