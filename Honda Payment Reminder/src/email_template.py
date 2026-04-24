def generate_email_html(row, invoice_date, payment_date):
    return f"""
<html>
<body style="font-family: Arial, sans-serif; font-size:14px; color:#000;">

<p>Dear Sir,</p>

<p><b>Greeting from Garantie!!</b></p>

<p>
I would like to inform you that, the EW Plus claim payment cleared on date <b>{payment_date}</b>.
</p>

<p>
Kindly check the payment in your dealership account and revert us via on same mail.
</p>

<p>
<span style="background-color:#ff0000; color:#000; font-weight:bold; padding:2px 6px;">
Attention
</span>
<span style="background-color:yellow;">
 kindly share the hard copy of invoice by courier to our below mention address with the Signature and Stamp.
</span>
</p>

<p>
Angle Paisa Consultancy Services Pvt Ltd (Garantie)<br>
B-7, Second Floor, Sector-1, Noida UP, India-201301
</p>

<p><b>Below are the complete payment details for your reference…..</b></p>

<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse; width:500px;">
    <tr>
        <td><b>Dealer code</b></td>
        <td>{row['Dealer Code']}</td>
    </tr>
    <tr>
        <td><b>Frame no</b></td>
        <td>{row['Frame no']}</td>
    </tr>
    <tr>
        <td><b>Invoice no</b></td>
        <td>{row['Invoice no']}</td>
    </tr>
    <tr>
        <td><b>Invoice date</b></td>
        <td>{invoice_date}</td>
    </tr>
    <tr>
        <td><b>Account Name</b></td>
        <td>{row['Account Name']}</td>
    </tr>
    <tr>
        <td><b>Bank Name</b></td>
        <td>{row['Bank name']}</td>
    </tr>
    <tr>
        <td><b>Account Number</b></td>
        <td>{row['Account number']}</td>
    </tr>
    <tr>
        <td><b>IFSC Code</b></td>
        <td>{row['IFSC Code']}</td>
    </tr>
    <tr>
        <td><b>Payment date</b></td>
        <td>{payment_date}</td>
    </tr>
    <tr>
        <td><b>Reference No.</b></td>
        <td>{row['Refrence no.']}</td>
    </tr>
    <tr>
        <td><b>Amount</b></td>
        <td>{row['Claim Amount']}</td>
    </tr>
</table>

<br>

<p>
Regards,<br><br>
Vikas Choudhary<br>
Team Claims<br>
Garantie<br>
M- 8527955526
</p>

</body>
</html>
"""