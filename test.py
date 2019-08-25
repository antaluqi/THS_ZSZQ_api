import ZSZQ

api=ZSZQ.API()

r=api.get_history_entrusts('2018-07-03','2018-08-01')
print(r)