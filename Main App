import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import os
import smtplib
from email.message import EmailMessage

# CONFIGURATION
st.set_page_config(page_title="Targeting Performance Tool", layout="wide")
st.title("🎯 Targeting Performance Dashboard")

EMAIL = os.getenv("EMAIL_ADDRESS", "automationbot.oms@gmail.com")
PASSWORD = os.getenv("EMAIL_PASSWORD", "ktxpnxchzwuuhvpd")
AM_EMAILS = {
    "Shira": "shirab@onlinemediasolutions.com",
}

# UPLOAD SECTION
with st.sidebar:
    uploaded_file = st.file_uploader("Upload 7-Day Excel Report", type=["xlsx"])
    send_email_flag = st.checkbox("Enable email preview & sending")
    test_mode = st.checkbox("🧪 Enable Test Mode for AMs")
    st.markdown("---")
    st.caption("Last upload is cached until replaced.")

if 'last_df' not in st.session_state:
    st.session_state['last_df'] = None

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.session_state['last_df'] = df
elif st.session_state['last_df'] is not None:
    df = st.session_state['last_df']
else:
    st.info("Please upload a report to begin.")
    st.stop()

df['Date'] = pd.to_datetime(df['Date'])
group_cols = ['AM', 'Publisher', 'Domain', 'Country', 'Device']

# TARGETING LOGIC
am_groups = {}
for keys, group in df.groupby(group_cols):
    am = keys[0]
    if am not in am_groups:
        am_groups[am] = []

    group = group.sort_values("Date")
    group['3DayImps'] = group['Pub Impressions'].rolling(3).sum()

    acceptable = (group['3DayImps'] > 1000).any()
    good = (group['Pub Impressions'] > 5000).any()

    if not acceptable:
        continue

    best_day = group['Pub Impressions'].max()
    worst_day = group['Pub Impressions'].min()

    tag = ""
    if best_day > 5000 and (group['Pub Impressions'] < 500).any():
        tag = "🔴"
    elif (group['Pub Impressions'] < 500).any() and best_day > 5000:
        tag = "🟢"

    am_groups[am].append({
        'Publisher': keys[1],
        'Domain': keys[2],
        'Country': keys[3],
        'Device': keys[4],
        'Good': good,
        'Tag': tag,
        'Best': best_day,
        'Worst': worst_day,
        'Daily Data': group
    })

st.success("Data processed successfully.")

# DECLINING TARGETINGS SUMMARY
st.subheader("📉 Top Declining Targetings")

def calculate_declines(combos, metric):
    declines = []
    for combo in combos:
        data = combo["Daily Data"]
        if data[metric].max() and data[metric].min():
            drop_value = data[metric].max() - data[metric].min()
            drop_pct = (1 - data[metric].min() / data[metric].max()) * 100
            declines.append({
                "Publisher": combo["Publisher"],
                "Domain": combo["Domain"],
                "Device": combo["Device"],
                "Drop Value": round(drop_value, 2),
                "Drop %": round(drop_pct, 1),
                "Metric": metric
            })
    return sorted(declines, key=lambda x: -x["Drop Value"])[:15]

for am, combos in am_groups.items():
    st.markdown(f"### 🔻 Decline Summary for **{am}**")
    for metric in ["Pub Impressions", "Revenue", "GP$"]:
        top_declines = calculate_declines(combos, metric)
        if top_declines:
            st.markdown(f"**Top Drops by {metric}:**")
            st.dataframe(pd.DataFrame(top_declines))

# TREND CHARTS PER AM
st.subheader("📈 AM Trend Charts")
for am, combos in am_groups.items():
    st.markdown(f"### {am}")
    for item in combos:
        fig, ax = plt.subplots(figsize=(3, 1))
        daily_data = item['Daily Data']
        ax.plot(daily_data['Date'], daily_data['Pub Impressions'], marker='o', linestyle='-')
        ax.set_title(f"{item['Publisher']} / {item['Domain']} / {item['Device']}", fontsize=8)
        ax.set_xticks([])
        ax.set_yticks([])
        ax.grid(True, linestyle='--', linewidth=0.5)
        st.pyplot(fig)

# SMART FLAGS
st.subheader("🚩 Smart Targeting Flags")
flag_summary = {"📉 Decline": 0, "⚠️ Volatile": 0, "🔥 Growth": 0}
for am in am_groups:
    for item in am_groups[am]:
        imp = item['Daily Data']['Pub Impressions']
        if imp.iloc[-1] < imp.iloc[0]:
            flag_summary["📉 Decline"] += 1
        elif imp.std() > imp.mean() * 0.5:
            flag_summary["⚠️ Volatile"] += 1
        elif imp.iloc[-1] > imp.iloc[0]:
            flag_summary["🔥 Growth"] += 1
st.write(flag_summary)

# DOWNLOAD REPORTS
st.markdown("### 📥 Download Reports")
for am, combos in am_groups.items():
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        for i, combo in enumerate(combos):
            combo['Daily Data'].to_excel(writer, sheet_name=f"{combo['Publisher'][:10]}_{i}"[:31], index=False)
        writer.save()
    st.download_button(f"Download for {am}", buffer.getvalue(), file_name=f"{am}_report.xlsx")

# EMAIL SENDING SECTION
if send_email_flag:
    st.markdown("---")
    st.header("📬 Email Results to Selected AMs")
    sender_email = st.text_input("Sender Email")
    selected_ams = [am for am in am_groups if st.checkbox(f"Send to {am}", key=f"email_{am}")]

    if st.button("Send Emails"):
        for am in selected_ams:
            recipient = AM_EMAILS.get(am)
            if not recipient:
                st.warning(f"No email address for {am}")
                continue

            msg = EmailMessage()
            msg["Subject"] = f"Your HB Tag Report – {am}"
            msg["From"] = sender_email
            msg["To"] = recipient

            html_summary = f"""
            <html>
                <body>
                    <p>Hi {am}!</p>
                    <p>Here is the performance of your HB tags in the past week:</p>
                    <ul>
            """
            for combo in am_groups[am]:
                html_summary += f"<li><b>{combo['Publisher']} / {combo['Domain']} / {combo['Device']}</b><br>"
                html_summary += f"Best Day: {combo['Best']} | Worst Day: {combo['Worst']} | Tag: {combo['Tag']}</li><br>"
            html_summary += """
                    </ul>
                    <p>Best regards,<br>The automation bot</p>
                </body>
            </html>
            """

            msg.set_content(
                f"Hi {am}!\n\nHere is the performance of your HB tags in the past week.\n\nSee attachment for full report."
            )
            msg.add_alternative(html_summary, subtype='html')

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                for i, combo in enumerate(am_groups[am]):
                    combo['Daily Data'].to_excel(writer, sheet_name=f"{combo['Publisher'][:10]}_{i}"[:31], index=False)
                writer.save()
            msg.add_attachment(
                buffer.getvalue(),
                maintype='application',
                subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                filename=f"{am}_report.xlsx"
            )

            try:
                with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                    smtp.login(EMAIL, PASSWORD)
                    smtp.send_message(msg)
                    st.success(f"Email sent to {am} ({recipient})")
            except Exception as e:
                st.error(f"Failed to email {am}: {e}")
