import ZSZQ
import pprint

api=ZSZQ.API()

r=api.get_history_entrusts('2018-02-03','2018-02-10')
pprint.pprint(r)