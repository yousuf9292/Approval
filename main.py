import glob
import io
from email import encoders
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
from datetime import date
import smtplib
import os

today=date.today()
onFirstDataRendered = JsCode("""
parmas.api.onFilterChanged();
}
""")




# if 'disabled' not in st.session_state:
#    st.session_state.disabled = True
# def disable():
#    st.session_state.disabled = False

username=st.text_input("Enter Username")
password=st.text_input("Enter Password",type='password')
button=st.button("Upload")
# send_email_to=st.text_input('Enter EmailAddress: ',disabled=True)

if ((username==st.secrets.get('user1') and password==st.secrets.get('password1'))
        or (username==st.secrets.get('user2') and password==st.secrets.get('password2'))):
    st.toast("Login Successfully")
    uploaded_file = st.file_uploader("Choose a file",key="uploader")
    if uploaded_file is not None:
        dataframe_original = pd.read_excel(uploaded_file)
        dataframe_original['DocDate'] = dataframe_original['DocDate'].astype('datetime64[ns]')
        dataframe_original['DocDueDate'] = dataframe_original['DocDate'].astype('datetime64[ns]')
        gb = GridOptionsBuilder.from_dataframe(dataframe_original, onFirstDataRendered=onFirstDataRendered)
        gb.configure_default_column(editable=True, filter=True)
        gb.configure_selection('multiple', use_checkbox=True, groupSelectsChildren=True,
                               groupSelectsFiltered=True)
        gridOptions = gb.build()
        grid_response = AgGrid(dataframe_original, gridOptions, allow_unsafe_jscode=True)
        out_df = grid_response["data"]
        selected_df = grid_response.selected_rows
        try:
            selected_df_to_download=dataframe_original[~dataframe_original['#'].isin(selected_df['#'])]
        except:
            print("Select Data")
        def to_excel(selected_df_to_download) -> bytes:
            output = io.BytesIO()
            writer = pd.ExcelWriter(output, engine="xlsxwriter")
            selected_df_to_download.to_excel(writer, sheet_name="Sheet1", index=False)
            writer.close()
            processed_data = output.getvalue()
            return processed_data
        send_button = st.button("Send Email")
        if send_button:
            try:
                sender = st.secrets.get('email_from')
                recipient = st.secrets.get('email_to')
                multipart = MIMEMultipart()
                multipart["From"] = sender
                multipart["To"] = recipient
                attachment = MIMEApplication(to_excel(selected_df))
                attachment["Content-Disposition"] = 'attachment; filename=" {}"'.format(f"{today}.xlsx")
                multipart.attach(attachment)
                server = smtplib.SMTP("mail.group-ge.com", 587)
                server.starttls()
                server.login(sender, st.secrets.get('email_pass'))
                server.sendmail(sender, recipient, multipart.as_string())
                server.quit()
            except:
                print("__")
        try:
            button = st.download_button(
            "Download as excel",
            data=to_excel(selected_df_to_download),
            file_name=f'{today}.xlsx',
            mime="application/vnd.ms-excel",
            )
        except:
            st.error("Select Data")


elif username=='' and password=='':
    st.toast("Kindly Enter Data")
elif ((username!=st.secrets.get('user1') or password!=st.secrets.get('password1'))
      and (username!=st.secrets.get('user2') or password!=st.secrets.get('password2'))):
    st.toast("Invalid Credentials")



# if send_button:
    # list_of_files = glob.glob(f'C:\\Users\\{os.getlogin()}\\Downloads\\*')
    # latest_file = max(list_of_files, key=os.path.getctime)
    # sender = st.secrets.get('email_from')
    # recipient = st.secrets.get('email_to')
    #
    # message = MIMEMultipart()
    # filename = latest_file
    #
    # with open(filename, "rb") as attachment:
    #
    #     # Add file as application/octet-stream
    #     # Email client can usually download this automatically as attachment
    #     part = MIMEBase("application", "octet-stream")
    #     part.set_payload(attachment.read())
    # # Encode file in ASCII characters to send by email
    # encoders.encode_base64(part)
    # # Add header as key/value pair to attachment part
    # part.add_header(
    #     "Content-Disposition",
    #     f"attachment; filename= {filename.split('\\')[-1]}",
    # )
    # message.attach(part)
    #
    # smtp = smtplib.SMTP("mail.group-ge.com", port=587)
    # smtp.starttls()
    # smtp.login(sender,st.secrets.get('email_pass'))
    # smtp.sendmail(sender, recipient, message.as_string())
    # smtp.quit()

