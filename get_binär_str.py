import base64

#erstellen des binärstings von xlsx datei

xlsx = open('vorlage.xltx', 'rb')
encoded = base64.b64encode(xlsx.read())
xlsx.close()

# 2. Base64-String in lesbare Zeilen stückeln
def chunk_string(s, length=140):
    return [s[i:i+length] for i in range(0, len(s), length)]


# 3. Format für Copy & Paste als gültiger Python-Bytestring
def format_for_python_bytes_literal(b64_str, chunk_size=140):
    chunks = chunk_string(b64_str, chunk_size)
    return 'xls_base64 = (\n' + '\n'.join([f"    b'{chunk}'" for chunk in chunks]) + '\n)'


# 4. base64-Str von bytes zu str (ASCII-Zeichen)
base64_str = encoded.decode('ascii')

# 5. Ausgabe für Copy & Paste im Code
print(format_for_python_bytes_literal(base64_str))
