import pandas as pd
from io import BytesIO

def generate_all_invoices(dump_file, rate_file, invoice_month):

    team_leads = [
        "Pathuri, Natraj",
        "Karimilla, Venugopal",
        "Lingineni, Srinivas"
    ]

    digital_team_leads = [
        "Nigam, Amit",
        "Harwell, Ansley",
        "Parvathaneni, Basava Dharmatej",
        "Jose, Jubin"
    ]

    frapi_team_leads = [
        "Navari, Sachithananda",
        "Bagora, Pranav"
    ]

    exclude_projects = [
        "Contractor Time Off/Holiday"
    ]

    rate_df = pd.read_excel(rate_file)

    # ================= DATA =================

    df = pd.read_excel(dump_file, sheet_name="Export")
    df['WORK_DATE'] = pd.to_datetime(df['WORK_DATE'])
    df = df[df['WORK_DATE'].dt.strftime('%Y-%m') == invoice_month]
    df = df[df['TEAM_LEAD'].isin(team_leads)]
    df = df[~df['PROJECT_NAME'].isin(exclude_projects)]
    df = df.merge(rate_df, on='RESOURCE', how='left')
    df['Bill Amt'] = df['Rate'] * df['HOURS']

    hrs_pivot = pd.pivot_table(df,index=['RESOURCE'],values=['HOURS','Bill Amt'],aggfunc='sum',fill_value=0).reset_index()
    hrs_pivot.columns = ['RESOURCE', 'Total Hrs', 'Total Bill Amt']

    total_row = pd.DataFrame({
        'RESOURCE':['Grand Total'],
        'Total Hrs':[hrs_pivot['Total Hrs'].sum()],
        'Total Bill Amt':[hrs_pivot['Total Bill Amt'].sum()]
    })

    hrs_pivot = pd.concat([hrs_pivot,total_row],ignore_index=True)

    pm_pivot = pd.pivot_table(df,index=['PHASE_FINANCIAL_CATEGORY','INITIATIVE_KEY'],values='Bill Amt',aggfunc='sum',fill_value=0).reset_index()
    pm_pivot.columns=['PHASE_FINANCIAL_CATEGORY','INITIATIVE_KEY','Total Bill Amt']

    category_totals = pm_pivot.groupby('PHASE_FINANCIAL_CATEGORY')['Total Bill Amt'].sum().reset_index()
    category_totals['INITIATIVE_KEY']='Subtotal'

    grand_total=pd.DataFrame({
        'PHASE_FINANCIAL_CATEGORY':['Grand Total'],
        'INITIATIVE_KEY':[''],
        'Total Bill Amt':[pm_pivot['Total Bill Amt'].sum()]
    })

    pm_pivot_final=pd.concat([pm_pivot,category_totals,grand_total],ignore_index=True)

    data_output=BytesIO()
    with pd.ExcelWriter(data_output,engine='openpyxl') as writer:
        df.to_excel(writer,sheet_name="Updated Tempo",index=False)
        hrs_pivot.to_excel(writer,sheet_name="hrs",index=False)
        pm_pivot_final.to_excel(writer,sheet_name="PM",index=False)
    data_output.seek(0)

    # ================= DIGITAL =================

    digital_df = pd.read_excel(dump_file, sheet_name="Export")
    digital_df['WORK_DATE'] = pd.to_datetime(digital_df['WORK_DATE'])
    digital_df = digital_df[digital_df['WORK_DATE'].dt.strftime('%Y-%m') == invoice_month]
    digital_df = digital_df[digital_df['TEAM_LEAD'].isin(digital_team_leads)]
    digital_df = digital_df[~digital_df['PROJECT_NAME'].isin(exclude_projects)]

    digital_hrs = pd.pivot_table(digital_df,index=['RESOURCE'],values='HOURS',aggfunc='sum',fill_value=0).reset_index()
    digital_hrs.columns=['RESOURCE','Total Hours']

    digital_invoice = digital_hrs.merge(rate_df,on='RESOURCE',how='left')
    digital_invoice['Amount']=digital_invoice['Total Hours']*digital_invoice['Rate']

    # 🔴 FORCE REMOVE TEAM
    digital_invoice = digital_invoice.drop(columns=['TEAM'], errors='ignore')

    digital_total=pd.DataFrame({
        'RESOURCE':['Grand Total'],
        'Total Hours':[digital_invoice['Total Hours'].sum()],
        'Rate':[''],
        'Amount':[digital_invoice['Amount'].sum()]
    })

    digital_invoice_final=pd.concat([digital_invoice,digital_total],ignore_index=True)

    digital_output=BytesIO()
    with pd.ExcelWriter(digital_output,engine='openpyxl') as writer:
        digital_df.to_excel(writer,sheet_name="Updated Tempo",index=False)
        digital_hrs.to_excel(writer,sheet_name="Resource Hours",index=False)
        digital_invoice_final.to_excel(writer,sheet_name="Digital_Invoice",index=False)
    digital_output.seek(0)

    # ================= FRAPI =================

    frapi_df = pd.read_excel(dump_file, sheet_name="Export")
    frapi_df['WORK_DATE'] = pd.to_datetime(frapi_df['WORK_DATE'])
    frapi_df = frapi_df[frapi_df['WORK_DATE'].dt.strftime('%Y-%m') == invoice_month]
    frapi_df = frapi_df[frapi_df['TEAM_LEAD'].isin(frapi_team_leads)]
    frapi_df = frapi_df[~frapi_df['PROJECT_NAME'].isin(exclude_projects)]

    frapi_hrs = pd.pivot_table(frapi_df,index=['RESOURCE'],values='HOURS',aggfunc='sum',fill_value=0).reset_index()
    frapi_hrs.columns=['RESOURCE','Total Hours']

    frapi_invoice = frapi_hrs.merge(rate_df,on='RESOURCE',how='left')
    frapi_invoice['Amount']=frapi_invoice['Total Hours']*frapi_invoice['Rate']

    # 🔴 FORCE REMOVE TEAM
    frapi_invoice = frapi_invoice.drop(columns=['TEAM'], errors='ignore')

    frapi_total=pd.DataFrame({
        'RESOURCE':['Grand Total'],
        'Total Hours':[frapi_invoice['Total Hours'].sum()],
        'Rate':[''],
        'Amount':[frapi_invoice['Amount'].sum()]
    })

    frapi_invoice_final=pd.concat([frapi_invoice,frapi_total],ignore_index=True)

    frapi_output=BytesIO()
    with pd.ExcelWriter(frapi_output,engine='openpyxl') as writer:
        frapi_df.to_excel(writer,sheet_name="Updated Tempo",index=False)
        frapi_hrs.to_excel(writer,sheet_name="Resource Hours",index=False)
        frapi_invoice_final.to_excel(writer,sheet_name="Frapi_Invoice",index=False)
    frapi_output.seek(0)

    return data_output,digital_output,frapi_output
