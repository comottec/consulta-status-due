import hashlib
 
# Declaring Password
password = '07028528'
# adding 5gz as password
salt = "ofi-consulta-due"
 
# Adding salt at the last of the password
dataBase_password = password+salt+"."+salt[::-1]+password[::-1]
print(dataBase_password)

# Encoding the password
hashed = hashlib.md5(dataBase_password.encode())
 
# Printing the Hash
print(hashed.hexdigest())
