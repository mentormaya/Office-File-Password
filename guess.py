import msoffcrypto, sys

office_file_types = ['docx', 'xlsx', 'pptx']

file_type = sys.argv[1].split('.')[-1]

if file_type in office_file_types:
  with open(sys.argv[1], 'rb') as f:
    try:
      file = msoffcrypto.OfficeFile(f)
      file.load_key(password=sys.argv[2], verify_password=True)
      print("password verified")
    except:
      print("invalid password")
else:
  print("Invalid file type!")