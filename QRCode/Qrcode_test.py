import qrcode
# Example to web page
data = "https://www.google.com.co"
# Outup file name
filename = "web.png"
# Generate QR Code
img = qrcode.make(data)
# Save img to file
img.save(filename)