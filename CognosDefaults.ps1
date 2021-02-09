#These variables will override/set anything you specify on the command line or in the CognosDownload.ps1 file.
#Only uncomment lines you want to be static ALL the time.

#Only set this if no other users in your domain need the script
#$username = '0000username'

#Your DSN for your district.
#$espdsn = 'schoolsms' #Recommend to set this

#Override password file path.
#$passwordfile = 'c:\scripts\mysavedpassword.txt'

#If you always want to save to the same path.
#$savepath = 'c:\scripts\files'

#eFinance DSN
#$efpdsn = 'schoolfms' #Recommend to set this

#eFinance User
#$efpuser = 'efinanceusername'

#Email Configuration so you don't have to put it on the command line. Still need to specify -SendMail on command line.
#$mailfrom = 'from@yourdomain.com"
#$mailto="technology@yourdomain.com"
#$smtpserver="smtp-relay.gmail.com"
#$smtpport="587"
#$smtppasswordfile="C:\Scripts\emailpw.txt" #change to a file path for email server password not needed if you use smtp-relay and auth your public IPs

#Example for Multiuser Environment. You can specify any variable from above in the switch statement.
#No Default needed as it will be specified at the command line, above, or default in CognosDownload.ps1
#switch($username){
    # '0401cmillsap' {
    #     $efpuser = '0403cmillsap'
    #     $passwordfile = 'c:\scripts\0403cmillsap-password.txt'
    #     $SendMail = $True
    # }
    # '0402cweber' {
    #     $efpuser = $username;
    #     $passwordfile = 'c:\scripts\0402cweber-password.txt';
    #     $savepath = "c:\scripts\ImportFiles"
    # }
#     'SSOusername' { $efpuser =''; $passwordfile = 'c:\scripts\importfiles\scripts\userpw1.txt' }
#     'SSOusername2' { $efpuser = ''; $passwordfile = 'c:\scripts\importfiles\scripts\userpw1.txt' }
#}