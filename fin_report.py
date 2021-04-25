""" Module to compute different types of transactions,generate report and
email it to the user"""

#!/usr/bin/env python3

#import modules
import smtplib
import argparse
import sys
from pprint import pprint
from email.message import EmailMessage
import filehandling as fh


#Argument parser
def get_arguments():
    parser = argparse.ArgumentParser(description='Automation to generate finance'
                                    ' report and email it.')
    parser.add_argument('-f', '--file', required=True,help='Monthly statement'
                        ' report file location (xlsx type only)')
    parser.add_argument('-se','--sender-email',required=True,help='Senders email address')
    parser.add_argument('-re','--reciever-email',required=False, 
                        help='Reciever email address (Default:Sender email)')
    parser.add_argument('-ps','--password',required=True,
    help='Senders password(put them in single quotes(\'\')) ')
    
    parser.parse_args(args=None if sys.argv[1:] else ['--help'])
    args = parser.parse_args()
    if not args.reciever_email:
        args.reciever_email = args.sender_email
    return args
#Main Function 
def main():
    args=get_arguments()
    if '.pdf' in args.file or '.xlsx' in args.file:
        msg=compute(args.file)
    else:
        print('Incorrect File format')
        exit(1)
    send_mail(args.sender_email,args.password,args.reciever_email,msg)
    print('Email sent')
#Function to compute transactions 
def compute(path):
    transactions = fh.call_file(path)
    creadited = 0
    debited = 0
    upi_sent =0
    upi_recived = 0
    month = transactions[-1]['date'].strftime('%B')
    opening_balance = transactions[0]['balance'] - transactions[0]['amount']
    closing_balance = transactions[-1]['balance'] 
    for each_entry in transactions:
        if each_entry['type'] == 'Debited':
            debited = debited + each_entry['amount']
            if 'UPI' in each_entry['info']:
                upi_sent = upi_sent + each_entry['amount']
        elif each_entry['type'] == 'Cridited':
            creadited = creadited + each_entry['amount']
            if 'UPI' in each_entry['info']:
                upi_recived = upi_recived + each_entry['amount']
    k='-'*120
    report= '''\n\nHi User,\n\nReport of {mnt} month:\n {ln}\n
    Account opening balance: ₹{op_bal}\n
    Account closing balance: ₹{cl_bal}\n
    UPI spent: ₹{u_st}\n
    UPI recived: ₹{r_st}\n
    Total amount spent in {mnt} : ₹{dbt}\n
    Total amount recived in {mnt} : ₹{crd}\n {ln}'''.format(mnt=month,ln=k,
    op_bal=opening_balance,cl_bal=closing_balance,u_st=upi_sent,r_st=upi_recived,
    dbt=debited,crd=creadited)
    print(report)
    return report

#Function tto send mail 
def send_mail(sender_email,passward,receiver_email,message):
    msg = EmailMessage()
    msg.set_content(message)
    msg['Subject'] = 'Monthly Report'
    msg['From'] = sender_email
    msg['To'] = receiver_email
    # creates SMTP session
    session = smtplib.SMTP('smtp.gmail.com', 587)
    session.ehlo()
    session.starttls()
    session.ehlo()
    try:
        # Authentication
        session.login(sender_email, passward)
        # sending the mail
        session.send_message(msg)
        # terminating the session
    except:
        print('Something went wrong. Maybe incorrect email address!!!')
        exit(1)
    session.quit()

if __name__ == "__main__":
    main()