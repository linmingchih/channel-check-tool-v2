import sys
from pyaedt import Edb
import json
from collections import defaultdict

edb_path = sys.argv[1]
edb_version = sys.argv[2]

# edb_path = 'data/Galileo_G87173_204_applied.aedb'
json_path = edb_path.replace('.aedb', '.json')
edb = Edb(edb_path, edbversion=edb_version)


#%%
info = {}
info['component'] = defaultdict(list)
for component_name, component in edb.components.components.items():
    for pin_name, pin in component.pins.items(): 
        info['component'][component_name].append((pin_name, pin.net_name)) 

#%%
info['diff'] = {}
for differential_pair_name, differential_pair in edb.differential_pairs.items.items(): 
    pos = differential_pair.positive_net.name
    neg = differential_pair.negative_net.name
    info['diff'][differential_pair_name] = (pos, neg)
    
with open(json_path, 'w') as f:
    json.dump(info, f, indent=3)
    
edb.close_edb()