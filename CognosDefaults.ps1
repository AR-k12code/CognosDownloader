#These variables will override/set anything you specify on the command line or in the CognosDownload.ps1 file.
#Only uncomment lines you want to be static ALL the time.

#$username = '0000username' #Only set this if no other users in your domain need the script
#$passwordfile = 'c:\scripts\mysavedpassword.txt'
#$espdsn = 'schoolsms' #Recommend to set this
#$savepath = 'c:\scripts\files' #Only use this if your 
#$efpdsn = 'schoolfms' #Recommend to set this
#$efpuser = 'efinanceusername'
#$mailfrom = 'from@yourdomain.com" #--- VARIABLE --- change for your email from address
#$mailto="technology@yourdomain.com" #--- VARIABLE --- change for your email server
#$smtpserver="smtp-relay.gmail.com" #--- VARIABLE --- change for your email to address
#$smtpport="587"
#$smtppasswordfile="C:\Scripts\emailpw.txt" #--- VARIABLE --- change to a file path for email server password not needed if you use smtp-relay and auth your public IPs

#Example Multiuser
#switch($username){
#     'SSOusername' { $efpuser = ''}
#     'SSOusername' {$efpuser = 'EfinUsername'}
#     'SSOusername' {$efpuser =''}
#     'SSOusername' {$efpuser = ''}
#     default {$efpuser = ''}
# }
# switch($username){
#     'SSOusername' { $passwordfile = 'c:\scripts\importfiles\scripts\userpw1.txt'}
#     'SSOusername' {$passwordfile = 'C:\Scripts\importfiles\scripts\userpw2.txt'}
#     'SSOusername' {$passwordfile ='C:\Scripts\importfiles\scripts\userpw3.txt'}
#     'SSOusername' {$passwordfile = 'C:\Scripts\importfiles\scripts\userpw4.txt'}
#     default {$passwordfile = 'C:\Scripts\importfiles\scripts\apscnpw.txt'}
# }
