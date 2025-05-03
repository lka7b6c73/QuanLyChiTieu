import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

SCOPE = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

CREDS_INFO = {
    "type": "service_account",
    "project_id": "graceful-wall-458715-d4",
    "private_key_id": "0647d0954f1399739113bddcd3aca408f7d26c4a",
    "private_key": """-----BEGIN PRIVATE KEY-----
MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQDt7OZgG0czcKsS
126ju1hBda9wEYhsI1YjSrjva0wC2KJ2VhI8dAkzPLYE46i1oANEiy19k0vDtHlz
PuD31KymIIk70LzlQrF31wKT6FArnQTXnwLIuiwJSHlkE4mSw42o+I7gUoVVyEpO
sAUQQjbUTEbk5EKhbHWNiNwg38swCE8Ui07F73+EIFEck8GINK+kS62HRFy6hf5B
He4eUD3X7EZ6iM2W2Kl3BsjxT/eOhgAM7khPvAPcjyfGs3+aFyOgWwVI10v2L1LK
H+1Rrv5LEsLbJqeCZRpesUQ9eWlChWDHTmjwbk/W6mBPN7iH5QnXiRHM1l8Y9LtN
jCTsJ2tfAgMBAAECggEAbOvHmZbTVTrOroz+gylooWsRekIkjLDupbB3EnAx95ma
YodLyO4sKg8wmvNoEGHlLRN0K2lrxXfRI8/U6NAti84YVVBdsa0DFsRRU0oGrbiv
91A904vbThKmNAj3sb8hkonbytksXlWXowdVAhj2BeG3kODomgLPnBBcCha56833
/8C5g8N5Tm2CSxdAawMU8uJSkR3TVS/h0SplCuOXhktxxrS6zRMD35cf1lWpBHbu
CPFXMVz3QbpJjOOFbommIGF1B2pIoabYVbKs6vcOk9Xd2z4ml/hPl+RvSfx1cPug
EV2mGtSIBCtEPieP/+569MeLr0m5AJgGYGvifvnoTQKBgQD5Biksy0zR2syCac4r
Fk7npMShZSodBuk8l6UbQ6olwzZea+SD2AstT909hwN1o3Mrky+/edqojDmEXB/Z
YiGNT0b8IAXtMjeJV/9F8Qr7PVUzjXPY0Y3pR2yxQVi5V3BKAy+QXudIStLt0m/D
jXRpEf3WElnmrB19FPxYq1iNWwKBgQD0lyWE2uyPzn8+v3SHFpTampR8l/xjrM9b
8s7cTkpQS3b02b97Qe3ntcW0epRBW/++FTFRhnhYPYuZvSNGtDPADK0AOP0gBGOw
yCqJshFU3qDJQHs7RqGv+k/qrxAtW1u65lJYSPGMO2iyEr1DpiwEpeaQAseW0eaj
8oPjFEBlTQKBgDoj0xcvO/c+80J7e3QZ2EBC+tmHqgZu56OK8DRmXuJEKnxvCkIx
/aINpGTKEee0Sp5g1eQJiiCR5JffflwdiiHY4YTZ4ShY8hhx4BqvRVYVIsBPhT2Y
514Qm/Lig9l57bCJ/9/gGHWLaqW13Pw2GXiyh6YkULMHejcK4EO/dYRDAoGBAMXl
7YN7JuCMMcRxYDVU/geJ/w54Ysu7LOO2p8z0w4gIJy2haR6vj3BlUJvgekgPyfLg
btEJviGXWnZ/5CT29NA08V02vwfCxW5p5pEWtJM4tltfgtXJJlWecD7wemRYgPXO
GV45Vsch4moCY4Ry1TbFjLGErjzHhJeuRyvaXOgZAoGBAOXUuERD7VKWkbU6l+kd
Y1Nf86idIQucGfsxgMXTtaERSdP81IR7pzpKeINFbwYEXDm6ndP6unKdK7d3vlCf
nxnbZuKDbLXm+pIi7x9zVW1b0rI4Ava7R00J1CufuvIAr9SBg7oLdoy+zakOtXik
4P8b7gLlDrX/3tw7sQoniHGE
-----END PRIVATE KEY-----""",
    "client_email": "quanlychitieubot@graceful-wall-458715-d4.iam.gserviceaccount.com",
    "client_id": "100136608456499102671",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/quanlychitieubot%40graceful-wall-458715-d4.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
}

CREDS = Credentials.from_service_account_info(CREDS_INFO, scopes=SCOPE)
gc = gspread.authorize(CREDS)
sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1wfZ65pua8kd4EHJHcXLGByr7JI6fiibc7lT406wcgLA/edit?usp=sharing")
worksheet = sh.sheet1

st.title("Quản lý chi tiêu của Khỉ và Cứt")

boxi = st.selectbox("Chọn loại chi tiêu:", ["Ăn uống", "Chi phí", "Mua đồ ăn", "Mua đồ dùng", "Tiền đi chơi", "Khác"])
mota = st.text_input("Mô tả:")
gia = st.number_input("Giá (nghìn đồng):", min_value=0, step=1)

if st.button("Xác nhận"):
    if mota and gia:
        now = datetime.now()
        thang = now.month
        ngay = now.day
        thoigian = now.strftime("%H:%M:%S")
        worksheet.append_row([thang, ngay, thoigian, mota, boxi, gia])
        st.success(f"Đã ghi: {thang}/{ngay} {thoigian} - {mota} - {boxi} - {gia} nghìn đồng")
    else:
        st.error("Vui lòng nhập đầy đủ thông tin.")
