import qrcode
value = 1

def Validate_filename (name):
    # Validate .png at the final of name.
    if name.find(".png", len(name)-4, len(name)) < 0:
        print("Debes colocar la extensión del archivo .png")
        return 0
    else: return 1

# Validate url web name
def Validate_Value_Url(url):
    if url.find("http://", 0, 7) > -1 or url.find("https://", 0, 8) > -1:
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