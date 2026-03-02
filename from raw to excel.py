import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="סורק נתונים", layout="wide")
st.title("🔍 סורק נתוני ALS")

# כפתור העלאה פשוט
als_files = st.file_uploader("העלה קבצי ALS לבדיקה", type=["xlsx"], accept_multiple_files=True)

if als_files:
    all_records = []
    
    for f in als_files:
        try:
            # קריאת הקובץ
            df_raw = pd.read_excel(f, sheet_name=None) # קורא את כל הגיליונות
            
            for sheet_name, df in df_raw.items():
                st.write(f"סורק גיליון: {sheet_name} בקובץ {f.name}")
                # הפיכה לטקסט כדי שיהיה קל לחפש
                all_records.append(df)
                
            st.success(f"סיימתי לסרוק את {f.name}")
        except Exception as e:
            st.error(f"שגיאה בקובץ {f.name}: {e}")

    if all_records:
        # איחוד הכל לקובץ אחד פשוט
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for i, df in enumerate(all_records):
                df.to_excel(writer, sheet_name=f"Data_{i}", index=False)
        
        st.markdown("---")
        st.subheader("אם הסריקה הצליחה, הכפתור למטה יוריד קובץ עם הנתונים הגולמיים:")
        st.download_button("הורד נתונים גולמיים", output.getvalue(), "raw_data_check.xlsx")
