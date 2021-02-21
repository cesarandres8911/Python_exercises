import re

# Precompile the patterns
regexes = [
    re.compile(p)
    for p in ['this', 'that']
]
text = 'Does this text match the pattern?'

print('Text: {!r}\n'.format(text))

for regex in regexes:
    print('Seeking "{}" ->'.format(regex.pattern),
          end=' ')

    if regex.search(text):
        print('match!')
    else:
        print('no match')


def remove_urls (vTEXT):
       vTEXT = re.sub(r'(<https|<http)?:\/\/(\w|\.|\/|\?|\=|\&|\%)*\b\>', '', vTEXT, flags=re.MULTILINE)
       return(vTEXT)

remove_ur = "<https://hcm19.sapsf.com/ui/uicore/img/email_footer_left.png> <https://hcm19.sapsf.com/companyLogoServlet/?companyId=ciprodecos> <https://hcm19.sapsf.com/ui/uicore/img/email_footer_right.png> Cordial saludo Para los fines pertinentes, informamos que el Sr. (sra) TEOFILO PEREA GELVIS, identificado (a) con Cédula de Ciudadanía No. 77182453, labora en la Empresa C.I PRODECO S.A(1000) hasta el día 24/09/2020, desempeñando como último cargo el de OPERADOR DE CAMION 789 CAL(693), en el departamento de Operaciones. Atentamente GESTIÓN HUMANA - MINA CALENTURITAS <https://hcm19.sapsf.com/ui/uicore/img/email_footer_center.jpg>"
#print( remove_urls("this is a test https://sdfs.sdfsdf.com/sdfsdf/sdfsdf/sd/sdfsdfs?bob=%20tree&jef=man lets see this too https://sdfsdf.fdf.com/sdf/f end"))
print (remove_urls(remove_ur))