
import ZSZQ
import pprint,datetime,random

api=ZSZQ.API()
r=api.buy('601398',5.4,200)
#r=api.cancel_entrust(10)
pprint.pprint(r)


