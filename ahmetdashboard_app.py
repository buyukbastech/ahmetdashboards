import pandas as pd
import streamlit as st
import plotly.express as px
from io import BytesIO
from fpdf import FPDF
from datetime import datetime
import base64
import tempfile
import matplotlib.pyplot as plt

st.set_page_config(page_title="Zabbix Alarm Görselleştirme", layout="wide")
st.title("📊 Zabbix Alarm Verisi Görselleştirici")

uploaded_file = st.file_uploader("📁 Excel dosyasını yükleyin", type=["xlsx"])

# Font dosyası ayarı
FONT_PATH = "DejaVuSans.ttf"
@st.cache_data
def download_font():
    import requests
    url = "https://github.com/dejavu-fonts/dejavu-fonts/raw/version_2_37/ttf/DejaVuSans.ttf"
    r = requests.get(url)
    with open(FONT_PATH, "wb") as f:
        f.write(r.content)

if not st.session_state.get("font_downloaded", False):
    download_font()
    st.session_state["font_downloaded"] = True

def find_best_match_column(df, candidates):
    """
    DataFrame sütunlarından candidates listesindeki en yakın isimli sütunu bulur.
    Yoksa None döner.
    """
    cols_lower = {col.lower(): col for col in df.columns}
    candidates_lower = [c.lower() for c in candidates]

    # Doğrudan eşleşme arar
    for cand in candidates_lower:
        if cand in cols_lower:
            return cols_lower[cand]

    # Yakın eşleşme arama (örneğin arada boşluk, _ farkı vs için)
    for cand in candidates_lower:
        for col_lower, col_orig in cols_lower.items():
            if cand.replace(" ", "") == col_lower.replace(" ", ""):
                return col_orig
    return None

def generate_mail_preview(df, col_ekip, col_kisim):
    previews = []
    grouped = df.groupby([col_ekip, col_kisim])
    for (ekip, kisim), group in grouped:
        konu = f"[Uyarı] {ekip} - {kisim} için kritik alarm bildirimi"
        mesaj = f"Merhaba {ekip} ekibi,\n\n" \
                f"Aşağıda {kisim} kısmında tespit edilen kritik uyarılar listelenmiştir:\n\n"
        for idx, row in group.iterrows():
            mesaj += f"- Problem: {row.get('Problem', 'Bilinmiyor')}\n" \
                     f"  Başlangıç: {row.get('Time', 'Bilinmiyor')}\n" \
                     f"  Durum: {row.get('Status', 'Bilinmiyor')}\n" \
                     f"  Süre (dk): {row.get('Duration', 'Bilinmiyor')}\n\n"
        mesaj += "Lütfen en kısa sürede kontrol edip geri dönüş sağlayınız.\n\nSaygılarımızla,\nZabbix Alarm Takip Sistemi"
        previews.append({col_ekip: ekip, col_kisim: kisim, 'Mail Konusu': konu, 'Mail İçeriği': mesaj})
    return pd.DataFrame(previews)

def filter_columns_for_manual_selection(columns, allowed_keywords):
    """
    columns içinden allowed_keywords listesinde geçen sütunları filtreler.
    Eğer filtrelenmiş liste boşsa orijinal tüm sütun listesini döner.
    """
    filtered = [col for col in columns if any(k.lower() in col.lower() for k in allowed_keywords)]
    return filtered if filtered else columns

def highlight_duration(row):
    """
    DataFrame stil fonksiyonu.
    Duration 60 dk üzerindeyse satırı kırmızı zemin beyaz yazı yapar.
    """
    try:
        duration = float(row['Duration'])
        if duration > 60:
            return ['background-color: red; color: white'] * len(row)
        else:
            return [''] * len(row)
    except:
        return [''] * len(row)

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names

    selected_sheets = st.multiselect("🧲 Görselleştirilecek sayfaları seçin", sheet_names, default=sheet_names)

    df_list = []
    styled_dfs = {}
    for sheet in selected_sheets:
        df = pd.read_excel(uploaded_file, sheet_name=sheet)
        df['Sayfa'] = sheet

        # Duration sütunu varsa, numerik hale getir
        if 'Duration' in df.columns:
            df['Duration'] = pd.to_numeric(df['Duration'], errors='coerce')

        # Stil uygulayıp sakla (Streamlit'de kullanmak için)
        styled = df.style.apply(highlight_duration, axis=1)
        styled_dfs[sheet] = styled

        df_list.append(df)

    all_data = pd.concat(df_list, ignore_index=True)
    all_data['Status'] = all_data['Status'].astype(str)

    # Sütun isimleri listesi, küçük harfe çevrilmiş
    all_cols_lower = [c.lower() for c in all_data.columns]

    # Önceden aranan sütun isimleri
    possible_ekip_names = ['sorumlu ekip', 'sorumlu_ekip', 'team', 'owner', 'responsible team']
    possible_kisim_names = ['kısım', 'kisim', 'bölüm', 'bolum', 'section', 'part']

    # Sütunları bulmaya çalışma
    col_ekip = find_best_match_column(all_data, possible_ekip_names)
    col_kisim = find_best_match_column(all_data, possible_kisim_names)

    # Manuel seçim için filtrelenecek sütun anahtar kelimeleri (örnek)
    ekip_allowed_keywords = ['sorgu', 'departman', 'ekip', 'birim', 'team', 'department']
    kisim_allowed_keywords = ['kısım', 'kisim', 'bölüm', 'bolum', 'section', 'part']

    # Eğer bulunamadıysa dropdown ile manuel seçim yaptır
    if col_ekip is None:
        st.warning("📌 'Sorumlu Ekip' sütunu bulunamadı, lütfen manuel seçin.")
        filtered_cols_ekip = filter_columns_for_manual_selection(all_data.columns.tolist(), ekip_allowed_keywords)
        col_ekip = st.selectbox("Sorumlu Ekip sütunu seçiniz", filtered_cols_ekip)

    if col_kisim is None:
        st.warning("📌 'Kısım' sütunu bulunamadı, lütfen manuel seçin.")
        filtered_cols_kisim = filter_columns_for_manual_selection(all_data.columns.tolist(), kisim_allowed_keywords)
        col_kisim = st.selectbox("Kısım sütunu seçiniz", filtered_cols_kisim)

    # Zaman filtreleme
    if 'Time' in all_data.columns:
        all_data['Time'] = pd.to_datetime(all_data['Time'], errors='coerce')
        min_date, max_date = all_data['Time'].min(), all_data['Time'].max()
        date_range = st.date_input("Tarih Aralığı Seçin", [min_date, max_date])
        if len(date_range) == 2:
            all_data = all_data[(all_data['Time'] >= pd.to_datetime(date_range[0])) & (all_data['Time'] <= pd.to_datetime(date_range[1]))]

    def group_status(status):
        try:
            if str(status).lower() == 'resolved':
                return 'Resolved'
            else:
                return 'Problem'
        except:
            return 'Unknown'

    all_data['StatusGroup'] = all_data['Status'].apply(group_status)

    color_map = {'Resolved': 'green', 'Problem': 'red', 'Unknown': 'gray'}

    selected_status = st.multiselect("Filtrelenecek Status değerleri", all_data['Status'].unique(), default=all_data['Status'].unique())
    filtered_data = all_data[all_data['Status'].isin(selected_status)]

    st.markdown(f"*Filtrelenmiş kayıt sayısı: {len(filtered_data)}*")

    st.header("📊 Grafik Türü Seçimi")
    chart_type = st.selectbox("Grafik Türü", ['Bar Grafik', 'Cizgi Grafik', 'Pasta Grafik'])

    chart_fig = None
    chart_title = ""

    if chart_type == 'Pasta Grafik':
        pie_data = filtered_data['StatusGroup'].value_counts().reset_index()
        pie_data.columns = ['StatusGroup', 'Sayi']

        fig = px.pie(pie_data, names='StatusGroup', values='Sayi', title="Alarmların Statü Bazlı Dağılımı",
                     color='StatusGroup', color_discrete_map=color_map, hole=0.3)
        st.plotly_chart(fig, use_container_width=True)
        chart_fig = fig
        chart_title = "Status Pasta Grafiği"

    else:
        columns = filtered_data.columns.tolist()
        x_col = st.selectbox("X Ekseni Kolonu", columns, index=0)
        y_col = st.selectbox("Y Ekseni Kolonu", columns, index=1 if len(columns) > 1 else 0)

        if chart_type == 'Bar Grafik':
            fig = px.bar(filtered_data, x=x_col, y=y_col, color='StatusGroup', color_discrete_map=color_map,
                         title=f"{x_col} vs {y_col} (Bar Grafik)")
        else:
            fig = px.line(filtered_data, x=x_col, y=y_col, color='StatusGroup', color_discrete_map=color_map,
                          title=f"{x_col} vs {y_col} (Çizgi Grafik)")

        st.plotly_chart(fig, use_container_width=True)
        chart_fig = fig
        chart_title = f"{x_col} vs {y_col}"

    # Saatlik dağılım grafiği
    if 'Time' in filtered_data.columns:
        st.subheader("🕒 Saatlik Alarm Dağılımı")
        hourly = filtered_data.copy()
        hourly['Saat'] = hourly['Time'].dt.hour
        hourly_count = hourly.groupby(['Saat', 'StatusGroup']).size().reset_index(name='Sayac')
        fig2 = px.bar(hourly_count, x='Saat', y='Sayac', color='StatusGroup', barmode='group',
                      color_discrete_map=color_map, title="Saatlik Alarm Sayısı Dağılımı")
        st.plotly_chart(fig2, use_container_width=True)

    # Veri tablosunu stil ile göster (Duration 60dk üzeri kırmızı)
    st.header("📋 Filtrelenmiş Veri Tablosu")
    def highlight_row(row):
        try:
            if float(row.get('Duration', 0)) > 60:
                return ['background-color: red; color: white'] * len(row)
        except:
            pass
        return [''] * len(row)

    st.dataframe(filtered_data.style.apply(highlight_row, axis=1), height=400)

    # Mail önizleme oluşturma ve gösterme
    if st.button("📧 Mail Önizlemeleri Oluştur"):
        mail_preview_df = generate_mail_preview(filtered_data, col_ekip, col_kisim)
        st.subheader("📧 Mail Önizlemeleri")
        st.dataframe(mail_preview_df)

    # Excel dosyasını stil ile indirilebilir hale getirme
    def to_excel_styled(df_dict):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for sheet_name, df_styled in df_dict.items():
                # Orijinal df çıkar
                df = df_styled.data if hasattr(df_styled, 'data') else df_styled
                # Filtre ve stil uygulama
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]
                # Filtre uygula
                worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)

                # Duration kolonu var ise 60dk üzerini kırmızı yap
                if 'Duration' in df.columns:
                    dur_idx = df.columns.get_loc('Duration')
                    red_format = workbook.add_format({'bg_color': '#FF0000', 'font_color': '#FFFFFF'})
                    for row_num, dur_val in enumerate(df['Duration'], start=1):
                        try:
                            if float(dur_val) > 60:
                                worksheet.set_row(row_num, None, red_format)
                        except:
                            pass
            writer.save()
            processed_data = output.getvalue()
        return processed_data

    excel_styled_data = to_excel_styled({sheet: styled_dfs[sheet] for sheet in styled_dfs})

    b64 = base64.b64encode(excel_styled_data).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="Zabbix_Alarm_Stil.xlsx">⬇️ Stil uygulanmış Excel dosyasını indir</a>'
    st.markdown(href, unsafe_allow_html=True)

    # Ayrı Grafikler sayfası için basit genel özet grafik (Streamlit içinde)
    st.header("📈 Tüm Sayfaların Özet Grafikleri")

    summary = all_data.groupby(['Sayfa', 'StatusGroup']).size().reset_index(name='Count')
    fig_summary = px.bar(summary, x='Sayfa', y='Count', color='StatusGroup', barmode='group',
                         color_discrete_map=color_map, title="Sayfa Bazlı Alarm Durumu Dağılımı")
    st.plotly_chart(fig_summary, use_container_width=True)