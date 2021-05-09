from pyDes import *
import base64
import hashlib
import re


def Decrypt(cipherText):
    my_key = "Asdf#NBtyeUKKBBCCkkLL"
    encoded_key = my_key.encode("utf-8")
    m = hashlib.md5()
    m.update(encoded_key)
    digest_key = m.digest()
    decoded_data = base64.b64decode(cipherText)
    k = triple_des(digest_key, ECB)
    k = triple_des(digest_key, ECB, b"\0\0\0\0\0\0\0\0", pad=None, padmode=PAD_PKCS5)
    my_out = k.decrypt(decoded_data)
    my_out = my_out.decode("utf-8")
    res = re.sub(r'[^\x00-\x7F]+','', my_out)
    return res

