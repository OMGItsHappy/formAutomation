import json

data = json.load(open("test.json"))

print(type(data["questionItem"]['question']['required']))