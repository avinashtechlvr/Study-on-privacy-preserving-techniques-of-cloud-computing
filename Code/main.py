import pandas as pd
from cryptography.fernet import Fernet
import json
import pymongo


############## -- Functions -- ##############
def read_excel_json(file_name):
    try:
        df = pd.read_excel(file_name)
        df = df.where(pd.notnull(df), None)
        d = df.to_dict(orient='records')
        return d
    except Exception as e:
        print("Error in reading excel:", str(e))


def get_headers(data):
    return list(data.keys())


def encrypt(s):
    try:
        key = Fernet.generate_key()
        ciphertext = s.encode()
        f = Fernet(key)
        ciphertext = f.encrypt(ciphertext)
        return ciphertext, key
    except Exception as e:
        print("Error in encrypting:", str(e))


def decrypt(s, key):
    try:
        f = Fernet(key)
        cleartext = f.decrypt(s)
        cleartext = cleartext.decode()
        return cleartext
    except Exception as e:
        print("Error in decrypting:", str(e))


def read_json(file_name):
    try:
        with open(file_name, 'r') as f:
            data = json.load(f)
        return data
    except Exception as e:
        print("Error in reading json:", str(e))


def append_json(file_name, data):
    try:
        with open(file_name, 'w') as f:
            json.dump(data, f)
    except Exception as e:
        print("Error in writing to json:", str(e))


def create_encrypted_dict(data):
    headers = get_headers(data)
    print(headers)
    key_json = {}
    encrypt_json = {}
    for j in headers:
        encrypted_data, key = encrypt(str(data[j]))
        key_json[j] = key.decode("utf-8")
        encrypt_json[j] = encrypted_data.decode("utf-8")  # convert bytes to string
    return encrypt_json, key_json


def decrypted_dic(data, key):
    headers = get_headers(data)
    decrypt_json = {}
    for j in headers:
        # convert string to bytes
        if j != "_id":
            decrypted_data = decrypt(data[j].encode("utf-8"), key[j].encode("utf-8"))
            decrypt_json[j] = decrypted_data
        decrypt_json["_id"] = data["_id"]
    return decrypt_json


# add function to add to mongo db
def add_to_mongo(data):
    try:
        client = pymongo.MongoClient(
            'mongodb+srv://miniproject:miniproject123@sandbox.sqigw.mongodb.net/myFirstDatabase?retryWrites=true&w=majority')
        db = client.ehr
        collection = db["ehr"]
        collection.insert_one(data)
    except Exception as e:
        print("Error in adding to mongo:", str(e))


def create_id():
    try:
        with open('id.txt', 'r') as f:
            id = f.read()
        # increment id by 1
        id = int(id) + 1
        # write incremented id to id.txt
        with open('id.txt', 'w') as f:
            f.write(str(id))
        return id
    except Exception as e:
        print("Error in reading id:", str(e))


def read_id(jsonfile):
    try:
        with open(jsonfile, 'r') as f:
            data = json.load(f)
        id = data["_id"]
        return id
    except Exception as e:
        print("Error in reading id:", str(e))


def json_toexcel(data, filename):
    try:
        if "xlsx" not in filename:
            filename = filename + ".xlsx"
        df = pd.DataFrame(data, index=[0])

        df.to_excel(filename, index=False)
    except Exception as e:
        print("Error in converting to excel:", str(e))


def id_in_database(id):
    client = pymongo.MongoClient(
        'Here enter your mongo db connection string')
    db = client.ehr
    collection = db["ehr"]
    finding = collection.find_one({"_id": int(id)})
    return finding


############## -- Functions -- ##############

# Taking Input File
data = read_excel_json('test.xlsx')
# encrypting data
encrypt_json, key_json = create_encrypted_dict(data[0])
# generating id
id = create_id()
encrypt_json["_id"] = id
key_json["_id"] = id

# creating file paths
filename = data[0]['First Name'] + data[0]['Last Name'] + ".json"
keyfname = data[0]['First Name'] + data[0]['Last Name'] + "key.json"

# printing encrypted data and keys
print(encrypt_json)
print(key_json)
# creating json files
append_json(filename, encrypt_json)
append_json(keyfname, key_json)
# Decrypting data
print("Decrypted data after encryption:", decrypted_dic(read_json(filename), read_json(keyfname)))

# adding to mongodb
add_to_mongo(encrypt_json)

# fetching details from mongodb
data = id_in_database(id)
# data aftter downloading from mongodb
print("Downloaded from mongo db", data)
print("After Downloading data from mongo db and then decrypting")

# Decrypting downloaded data
dec_data = decrypted_dic(data, read_json(keyfname))
print(dec_data)
# deleting id as excel file does not have id
del dec_data["_id"]
# creating path for excel file
filename = dec_data["First Name"] + dec_data["Last Name"] + ".xlsx"
# converting to excel file
json_toexcel(dec_data, filename)
