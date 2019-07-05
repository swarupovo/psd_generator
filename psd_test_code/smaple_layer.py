from psd_tools import PSDImage

psd = PSDImage.open('Industrial_visit_certificate.psd')
psd.compose().save('Industrial_visit_certificate.png')

# print(dir(psd))
a=psd.name
# print(a)
# 

print(psd.descendants())
for layer in psd.descendants():
    print(layer)

# for layer in psd:
#     # print(dir(layer))
#     print(layer.name)