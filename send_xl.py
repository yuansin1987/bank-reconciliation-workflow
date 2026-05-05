import urllib.request, urllib.parse, json, uuid

APP_ID = 'cli_a951642e8878dbd2'
APP_SECRET = 'PY2xTH4USYS4Ew4oUXkKScrLnQ22qnJq'
UID = 'ou_57883f64459bea7339fc6c78d094d0e1'

data = json.dumps({'app_id': APP_ID, 'app_secret': APP_SECRET}).encode()
req = urllib.request.Request('https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal', data=data, headers={'Content-Type': 'application/json'})
token = json.loads(urllib.request.urlopen(req).read())['tenant_access_token']
print('token ok')

boundary = 'U' + str(uuid.uuid4()).replace('-','')
xlsx = r'C:\Users\wdxyh\Desktop\银行对账与发票核对_2026年03月.xlsx'
fname = '银行对账与发票核对_2026年03月.xlsx'
with open(xlsx, 'rb') as f:
    file_data = f.read()

hdr = ('--' + boundary + '\r\nContent-Disposition: form-data; name="file_type"\r\n\r\nxlsx\r\n--' + boundary + '\r\nContent-Disposition: form-data; name="file_name"\r\n\r\n' + fname + '\r\n--' + boundary + '\r\nContent-Disposition: form-data; name="file"; filename="' + fname + '"\r\nContent-Type: application/octet-stream\r\n\r\n').encode()

body = hdr + file_data + ('\r\n--' + boundary + '--\r\n').encode()

req = urllib.request.Request('https://open.feishu.cn/open-apis/im/v1/files', data=body, headers={'Authorization': 'Bearer ' + token, 'Content-Type': 'multipart/form-data; boundary=' + boundary})
resp = json.loads(urllib.request.urlopen(req).read())
print('upload:', resp.get('code'), resp.get('data', {}).get('file_key', ''))

if resp.get('code') == 0:
    fk = resp['data']['file_key']
    msg = json.dumps({'receive_id': UID, 'msg_type': 'file', 'content': json.dumps({'file_key': fk})}).encode()
    req2 = urllib.request.Request('https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=open_id', data=msg, headers={'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json'})
    r2 = json.loads(urllib.request.urlopen(req2).read())
    print('send:', r2.get('code'))
