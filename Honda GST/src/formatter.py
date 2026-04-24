def format_table(df):
    rows = ""

    for _, row in df.iterrows():
        rows += f"""
        <tr>
            <td>{row['Dealer code']}</td>
            <td>{row['Particulars']}</td>
            <td>{row['Invoice Date']}</td>
            <td>{row['Invoice']}</td>
            <td style="text-align:right;">{row['Taxable']}</td>
            <td style="text-align:right;">{row['IGST']}</td>
            <td style="text-align:right;">{row['CGST']}</td>
            <td style="text-align:right;">{row['SGST']}</td>
        </tr>
        """

    return f"""
    <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse; font-family:Arial, sans-serif; font-size:13px;">
        <tr style="background:#D9EAF7; font-weight:bold;">
            <th>Dealer Code</th>
            <th>Particulars</th>
            <th>Invoice Date</th>
            <th>Invoice</th>
            <th>Taxable</th>
            <th>IGST</th>
            <th>CGST</th>
            <th>SGST</th>
        </tr>
        {rows}
    </table>
    """