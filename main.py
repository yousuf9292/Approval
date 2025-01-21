import glob
import io
import pythoncom
import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
from datetime import date
import os
import win32com.client as win32

today=date.today()
onFirstDataRendered = JsCode("""
parmas.api.onFilterChanged();
}
""")


#if 'disabled' not in st.session_state:
#    st.session_state.disabled = True
#def disable():
#    st.session_state.disabled = False
username=st.text_input("Enter Username")
password=st.text_input("Enter Password",type='password')
button=st.button("Upload")
send_email_to=st.text_input('Enter EmailAddress: ',value='yousufsyed900@gmail.com',disabled=True)
send_button=st.button("Send Email")
if (username=='yousuf' and password=='abc') or (username=='yousuf1' and password=='abcs'):
    st.toast("Login Successfully")
    uploaded_file = st.file_uploader("Choose a file",key="uploader")
    if uploaded_file is not None:
        dataframe = pd.read_excel(uploaded_file)
        dataframe['DocDate'] = dataframe['DocDate'].astype('datetime64[ns]')
        dataframe['DocDueDate'] = dataframe['DocDate'].astype('datetime64[ns]')
        gb = GridOptionsBuilder.from_dataframe(dataframe, onFirstDataRendered=onFirstDataRendered)
        gb.configure_default_column(editable=True, filter=True)
        gb.configure_selection('multiple', use_checkbox=True, groupSelectsChildren=True,
                               groupSelectsFiltered=True)
        gridOptions = gb.build()
        grid_response = AgGrid(dataframe, gridOptions, allow_unsafe_jscode=True)
        out_df = grid_response["data"]
        selected_df = grid_response.selected_rows
        def to_excel(selected_df) -> bytes:
            output = io.BytesIO()
            writer = pd.ExcelWriter(output, engine="xlsxwriter")
            selected_df.to_excel(writer, sheet_name="Sheet1", index=False)
            writer.close()
            processed_data = output.getvalue()
            return processed_data
        try:
            button = st.download_button(
                "Download as excel",
                data=to_excel(selected_df),
                file_name=f'{today}.xlsx',
                mime="application/vnd.ms-excel",
            )
        except:
            st.error("Select Data")


elif username=='' and password=='':
    st.toast("Kindly Enter Data")
elif (username!='yousuf' or password!='abc') and (username!='yousuf1' or password!='abcs'):
    st.toast("Invalid Credentials")



if send_button:
    print("inside")
    olApp = win32.Dispatch('Outlook.Application',pythoncom.CoInitialize())
    olNS = olApp.GetNameSpace('MAPI')

    # construct the email item object
    mailItem = olApp.CreateItem(0)
    mailItem.To = send_email_to
    list_of_files = glob.glob('C:\\Users\\SAP User 1\\Downloads\\*')
    latest_file = max(list_of_files, key=os.path.getctime)
    print(latest_file)
    mailItem.Attachments.Add(latest_file)
    mailItem.Send()
