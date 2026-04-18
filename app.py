import streamlit as st
import pandas as pd
import anthropic
import json
import io
import base64
from pdf2image import convert_from_bytes

# ==========================================
# 1. הגדרות מערכת וחיבור אוטומטי ל-API
# ==========================================
st.set_page_config(page_title="מערכת בקרת דוחות לוגיסטיים", layout="wide")

st.title("📊 מערכת בקרת דוחות לוגיסטיים - הקו הסגול")
st.markdown("העלה את קובץ החפורת המרכזי (Excel) ואת הסריקות של תעודות המשלוח והקליטה (PDF).")

# משיכת המפתח אוטומטית מ-Streamlit Secrets
try:
    api_key = st.secrets["ANTHROPIC_API_KEY"]
    client = anthropic.Anthropic(api_key=api_key)
except KeyError:
    st.error("🚨 שגיאה: לא נמצא מפתח API. אנא ודא שהגדרת ANTHROPIC_API_KEY ב-Secrets של סטרימליט.")
    st.stop()

# ==========================================
# 2. פונקציות עיבוד PDF ו-AI
# ==========================================
def pdf_to_base64_images(pdf_bytes):
    """מקבל קובץ PDF, ממיר לעמודים ומחזיר רשימה של תמונות בפורמט Base64 שקלוד יכול לקרוא"""
    images = convert_from_bytes(pdf_bytes)
    base64_images = []
    for img in images:
        buffer = io.BytesIO()
        img.save(buffer, format="JPEG")
        img_str = base64.b64encode(buffer.getvalue()).decode("utf-8")
        base64_images.append(img_str)
    return base64_images

def analyze_document_with_claude(base64_image):
    """שולח עמוד סרוק לקלוד כדי לחלץ את מספרי התעודות והאסמכתאות"""
    prompt = """
    Analyze this logistics document (Delivery Note or Absorption Certificate).
    Extract the data into a strict JSON format with the following keys exactly:
    - "delivery_note_number": The number of the delivery note (if present, else null)
    - "absorption_cert_number": The number of the absorption/weighing certificate (if present, else null)
    - "reference_number": Any external reference number, 'Assmachta', or source document number mentioned in the text (if present, else null)
    - "document_type": Either "Delivery", "Absorption", or "Unknown"
    Return ONLY valid JSON. No Markdown formatting, no other text.
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
        # ניקוי הפלט למקרה שקלוד הוסיף מילים
        result_text = response.content[0].text.strip()
        if result_text.startswith("```json"):
            result_text = result_text[7:-3]
        return json.loads(result_text)
    except Exception as e:
        return {"error": str(e)}

# ==========================================
# 3. ממשק משתמש (העלאת קבצים ובדיקה)
# ==========================================
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. קובץ חפורת מרכזי (Excel)")
    excel_file = st.file_uploader("העלה את הטבלה המרכזית", type=["xlsx", "xls", "csv"])

with col2:
    st.subheader("2. דוחות סרוקים (PDF)")
    pdf_files = st.file_uploader("העלה תעודות משלוח וקליטה", type=["pdf"], accept_multiple_files=True)

if st.button("🚀 התחל בדיקת התאמות", type="primary"):
    if not excel_file or not pdf_files:
        st.warning("נא להעלות את כל הקבצים הנדרשים לפני תחילת הבדיקה.")
    else:
        # טעינת קובץ האקסל לזיכרון
        if excel_file.name.endswith('.csv'):
            df_master = pd.read_csv(excel_file)
        else:
            df_master = pd.read_excel(excel_file)
            
        # המרת כל הנתונים באקסל למחרוזות כדי למנוע בעיות השוואה (למשל מספר 123 מול טקסט "123")
        df_master = df_master.astype(str)
        
        results = []
        progress_text = "מפענח סריקות (זה עשוי לקחת קצת זמן)..."
        my_bar = st.progress(0, text=progress_text)
        
        total_files = len(pdf_files)
        
        # מעבר על כל קבצי ה-PDF שהועלו
        for i, pdf in enumerate(pdf_files):
            pdf_bytes = pdf.read()
            images_base64 = pdf_to_base64_images(pdf_bytes)
            
            # בדיקת כל עמוד ב-PDF
            for page_num, img_b64 in enumerate(images_base64):
                data = analyze_document_with_claude(img_b64)
                
                # אם חזר JSON תקין, נבצע את הבדיקות הלוגיות
                if "error" not in data:
                    doc_type = data.get("document_type", "Unknown")
                    deliv_num = str(data.get("delivery_note_number", ""))
                    absorp_num = str(data.get("absorption_cert_number", ""))
                    ref_num = str(data.get("reference_number", ""))
                    
                    status = "✅ תקין"
                    notes = []
                    
                    # בדיקה 1: אם זו תעודת קליטה, האם יש אסמכתא?
                    if doc_type == "Absorption" and (ref_num == "None" or ref_num == ""):
                        status = "❌ חריגה"
                        notes.append("חסרה אסמכתא בתעודת הקליטה")
                    
                    # בדיקה 2: חיפוש בקובץ החפורת
                    search_term = deliv_num if doc_type == "Delivery" else absorp_num
                    if search_term and search_term != "None":
                        # חיפוש גמיש בכל העמודות באקסל
                        found_in_excel = df_master.apply(lambda row: row.astype(str).str.contains(search_term).any(), axis=1).any()
                        if not found_in_excel:
                            status = "❌ חריגה"
                            notes.append(f"תעודה {search_term} לא נמצאה בקובץ האקסל")
                            
                    results.append({
                        "קובץ": pdf.name,
                        "עמוד": page_num + 1,
                        "סוג תעודה": doc_type,
                        "מספר תעודה": deliv_num if doc_type == "Delivery" else absorp_num,
                        "אסמכתא (מקור)": ref_num,
                        "סטטוס": status,
                        "הערות": ", ".join(notes)
                    })
                    
            my_bar.progress((i + 1) / total_files, text=f"עובד על קובץ {i+1} מתוך {total_files}...")
            
        my_bar.empty()
        st.success("הבדיקה הסתיימה!")
        
        # תצוגת התוצאות
        if results:
            df_results = pd.DataFrame(results)
            st.dataframe(df_results, use_container_width=True)
            
            # כפתור הורדה
            csv = df_results.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 הורד דוח פערים (CSV)",
                data=csv,
                file_name='validation_report.csv',
                mime='text/csv',
            )
