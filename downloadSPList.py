from shareplum import Site
from requests_ntlm import HttpNtlmAuth

cred = HttpNtlmAuth('ccholon', '3.Ch0l0n.k1ds')
site = Site('http://tng/SharedServices/AR', auth=cred)

list = site.List('Document_Circulation').get_list_items()
for item in list:
    




