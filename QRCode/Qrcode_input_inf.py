import qrcode
value = 1

def Validate_filename (name):
    val = name.count(".png")
    if val < 1:
        print("Debes colocar la extensión del archivo .png")
        return 0
    else: return 1

# Validate url web name
def Validate_Value_Url(url):
    url1 = url[0:7]
    url2 = url[0:8]
    if url1 == "http://" or url2 == "https://":
        value = 3
    else:
        value = 1
        print("Escriba de forma correcta la dirección web, http:// o https://")
    return value

while value < 2:
    # Enter the url page
    data = input("Enter the url site for create qrcode: \n")
    #Validate_Value_Url(data)
    if Validate_Value_Url(data) > 2:
        break

# Outup file name
while value < 5:
    filename = input("Enter the filename: \n")
    if Validate_filename (filename) > 0:
        break

# Generate QR Code
img = qrcode.make(data)
# Save img to file
img.save(filename)