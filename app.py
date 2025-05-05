import streamlit as st
import pandas as pd
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
import google.auth.transport.requests
import requests
import os

def get_credentials_from_query():
    client_config = {
        "web": {
            "client_id": st.secrets["GOOGLE_CLIENT_ID"],
            "client_secret": st.secrets["GOOGLE_CLIENT_SECRET"],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
            "redirect_uris": ["https://slides-generator-cgoprblddggy2thr2fb4oi.streamlit.app/"]
        }
    }

    flow = Flow.from_client_config(
        client_config,
        scopes=[
            "https://www.googleapis.com/auth/presentations",
            "https://www.googleapis.com/auth/drive.file"
        ],
        redirect_uri="https://slides-generator-cgoprblddggy2thr2fb4oi.streamlit.app/"
    )

    query_params = st.experimental_get_query_params()
    if "code" not in query_params:
        auth_url, _ = flow.authorization_url(prompt='consent')
        st.markdown(f"[こちらをクリックしてGoogleにログイン]({auth_url})")
        st.stop()
    else:
        code = query_params["code"][0]
        flow.fetch_token(code=code)
        return flow.credentials

def generate_slide(creds, df):
    slides = build("slides", "v1", credentials=creds)
    presentation = slides.presentations().create(body={"title": "CSVから自動生成"}).execute()
    pres_id = presentation["presentationId"]
    slide_id = presentation["slides"][0]["objectId"]

    requests = []
    for i, row in df.iterrows():
        requests.append({
            "createShape": {
                "objectId": f"text_{i}",
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "transform": {
                        "scaleX": 1, "scaleY": 1,
                        "translateX": 50,
                        "translateY": 50 + i * 50,
                        "unit": "PT"
                    },
                    "size": {
                        "height": {"magnitude": 30, "unit": "PT"},
                        "width": {"magnitude": 500, "unit": "PT"},
                    }
                }
            }
        })
        requests.append({
            "insertText": {
                "objectId": f"text_{i}",
                "insertionIndex": 0,
                "text": ", ".join([str(v) for v in row.values])
            }
        })

    slides.presentations().batchUpdate(
        presentationId=pres_id, body={"requests": requests}
    ).execute()

    return f"https://docs.google.com/presentation/d/{pres_id}/edit"

st.title("📊 CSV → Google Slides ヒートマップ生成ツール（自動認証）")

uploaded = st.file_uploader("CSVファイルをアップロードしてください", type="csv")
if uploaded:
    df = pd.read_csv(uploaded)
    st.write("プレビュー：", df.head())

    creds = get_credentials_from_query()
    if creds:
        if st.button("スライドを生成"):
            url = generate_slide(creds, df)
            st.success(f"スライドを生成しました！ [開く]({url})")
