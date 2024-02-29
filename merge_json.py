import json
data = []
for i in range(1, 27):
    with open(f'{i}.json', 'r') as file:
        links = json.load(file)
        data += links
print(len(data))
with open('merge.json', 'w') as file:
    json.dump(data, file)