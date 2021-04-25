# Finance-automation

Create an account on pdftable_api and use the api key to convert pdf to excel.\n 
Add the api key in filehandling.py
Email can only be sent and recived from gmail (Sender has to disable 2-setep-authentication)
Install pdftable_api
pip install https://github.com/pdftables/python-pdftables-api/archive/master.tar.gz
or
pip3 install https://github.com/pdftables/python-pdftables-api/archive/master.tar.gz

Install openpyxl
pip insatll openpyxl
or
pip3 install openxl


Usage:
./fin_report.py -f <file address> -se <sender email address> [-re <reciever email address>] -ps <senders password> 
  or 
  python3 fin_report.py -f <file address> -se <sender email address> [-re <reciever email address>] -ps <senders password>
