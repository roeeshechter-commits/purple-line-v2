import streamlit as st
import pandas as pd
import anthropic
import json
import io
import base64
import os
from datetime import datetime
from pdf2image import convert_from_bytes
import re

# ==========================================
# 1. הגדרות מערכת ויישור לימין (RTL)
# ==========================================
st.set_page_config(page_title="מערכת בקרת דוחות - הקו הסגול", layout="wide")

# הזרקת קוד CSS ליישור האתר לימין (עברית)
st.markdown("""
    <style>
    body, .stApp, .stMarkdown, h1, h2, h3, h4, h5, h6, p, div, label {
        direction: rtl;
        text-align: right;
    }
    .stTabs [data-baseweb="tab-list"] {
        justify-content: flex-start;
        direction: rtl;
    }
    .stButton>button {
        float: right;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 מערכת בקרת דוחות לוגיסטיים - הקו הסגול")

# משיכת המפתח
try:
    api_key = st.secrets["ANTHROPIC_API_KEY"]
    client = anthropic.Anthropic(api_key=api_key)
except KeyError:
    st.error("🚨 שגיאה: לא נמצא מפתח API ב-Secrets.")
    st.stop()

# ==========================================
# 2. פונקציות עיבוד
# ==========================================
def pdf_to_base64_images(pdf_bytes):
    images = convert_from_bytes(pdf_bytes)
    base64_images = []
    for img in images:
        buffer = io.BytesIO()
        img.save(buffer, format="JPEG")
        img_str = base64.b64encode(buffer.getvalue()).decode("utf-8")
        base64_images.append(img_str)
    return base64_images

def analyze_document_with_claude(base64_image):
    prompt = """
    Analyze this logistics document. Extract data into strict JSON:
    {
      "delivery_note_number": "number or null",
      "absorption_cert_number": "number or null",
      "reference_number": "number or null",
      "document_type": "Delivery or Absorption"
    }
    Return ONLY JSON.
    """
    try:
        response = client.messages.create(
            model="claude-3-5-sonnet-20241022",
            max_tokens=500,
            temperature=0,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": base64_image}},
                        {"type": "text", "text": prompt}
                    ]
                }
            ]
        )
        result_text = response.content[0].text.strip()
        # שליפת ה-JSON גם אם קלוד הוסיף טקסט מיותר
        match = re.search(r'\{.*\}', result_text, re.DOTALL)
        if match:
            return json.loads(match.group(0))
        return {"error": "Failed to parse JSON"}
    except Exception as e:
        return {"error": str(e)}

def save_to_history(df):
    """שמירת התוצאות לקובץ היסטוריה מקומי"""
    history_file = "history_log.csv"
    df["תאריך בדיקה"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    if os.path.exists(history_file):
        df_existing = pd.read_csv(history_file)
        df_combined = pd.concat([df_existing, df], ignore_index=True)
        df_combined.to_csv(history_file, index=False, encoding='utf-8-sig')
    else:
        df.to_csv(history_file, index=False, encoding='utf-8-sig')

# ==========================================
# 3. ממשק משתמש - טאבים
# ==========================================
tab_main, tab_history = st.tabs(["🔍 בדיקה חדשה", "🕒 היסטוריית בדיקות"])

with tab_main:
    col1, col2 = st.columns(2)
    with col1:
        excel_file = st.file_uploader("1. העלה קובץ חפורת מרכזי (Excel)", type=["xlsx", "xls", "csv"])
    with col2:
        pdf_files = st.file_uploader("2. העלה דוחות סרוקים (PDF)", type=["pdf"], accept_multiple_files=True)

    if st.button("🚀 התחל בדיקת התאמות", type="primary"):
        if not excel_file or not pdf_files:
            st.warning("נא להעלות את כל הקבצים הנדרשים לפני תחילת הבדיקה.")
        else:
            if excel_file.name.endswith('.csv'):
                df_master = pd.read_csv(excel_file)
            else:
                df_master = pd.read_excel(excel_file)
                
            df_master = df_master.astype(str)
            results = []
            
            my_bar = st.progress(0, text="מפענח סריקות...")
            total_files = len(pdf_files)
            
            for i, pdf in enumerate(pdf_files):
                pdf_bytes = pdf.read()
                images_base64 = pdf_to_base64_images(pdf_bytes)
                
                for page_num, img_b64 in enumerate(images_base64):
                    data = analyze_document_with_claude(img_b64)
                    
                    if "error" in data:
                        results.append({
                            "קובץ": pdf.name, "עמוד": page_num + 1, "סוג תעודה": "שגיאה",
                            "מספר תעודה": "-", "אסמכתא (מקור)": "-", "סטטוס": "❌ שגיאת פענוח AI",
                            "הערות": data["error"]
                        })
                        continue
                        
                    doc_type = data.get("document_type", "Unknown")
                    deliv_num = str(data.get("delivery_note_number", ""))
                    absorp_num = str(data.get("absorption_cert_number", ""))
                    ref_num = str(data.get("reference_number", ""))
                    
                    status = "✅ תקין"
                    notes = []
                    
                    if doc_type == "Absorption" and (ref_num == "None" or ref_num == "" or ref_num == "null"):
                        status = "❌ חריגה"
                        notes.append("חסרה אסמכתא")
                    
                    search_term = deliv_num if doc_type == "Delivery" else absorp_num
                    if search_term and search_term not in ["None", "", "null"]:
                        found_in_excel = df_master.apply(lambda row: row.astype(str).str.contains(search_term).any(), axis=1).any()
                        if not found_in_excel:
                            status = "❌ חריגה"
                            notes.append(f"לא נמצא באקסל")
                            
                    results.append({
                        "קובץ": pdf.name, "עמוד": page_num + 1,
                        "סוג תעודה": "משלוח" if doc_type == "Delivery" else "קליטה",
                        "מספר תעודה": deliv_num if doc_type == "Delivery" else absorp_num,
                        "אסמכתא (מקור)": ref_num if ref_num != "None" else "",
                        "סטטוס": status,
                        "הערות": ", ".join(notes)
                    })
                        
                my_bar.progress((i + 1) / total_files, text=f"עובד על קובץ {i+1} מתוך {total_files}...")
                
            my_bar.empty()
            
            if results:
                df_results = pd.DataFrame(results)
                st.success("הבדיקה הסתיימה! להלן התוצאות:")
                st.dataframe(df_results, use_container_width=True)
                
                # שמירה להיסטוריה
                save_to_history(df_results)
                
                csv = df_results.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 הורד דוח פערים (CSV)", data=csv, file_name='validation_report.csv', mime='text/csv')
            else:
                st.warning("לא נמצאו נתונים לפענוח.")

with tab_history:
    st.subheader("היסטוריית בדיקות קודמות")
    st.info("כאן נשמרים כל הפלטים של הבדיקות שביצעת (הנתונים נשמרים זמנית על השרת).")
    if os.path.exists("history_log.csv"):
        df_hist = pd.read_csv("history_log.csv")
        st.dataframe(df_hist, use_container_width=True)
        
        csv_hist = df_hist.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 הורד היסטוריה מלאה", data=csv_hist, file_name='full_history.csv', mime='text/csv')
    else:
        st.write("עדיין לא בוצעו בדיקות.")
