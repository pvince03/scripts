Send-MailMessage -Credential $credential ` 
-useSSL ` 
-smtpServer 'email-smtp.us-west-2.amazonaws.com' ` 
-port 587 ` 
-from 'patrick.vince@hyland.com' ` 
-to 'patrick.vince@hyland.com' ` 
-subject ` 
'test from my laptop' ` 
-body 'Email via Amazon SES SMTP Endpoint' 