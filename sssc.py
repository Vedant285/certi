#read line by line from a csv file
with open('./Data/Participant List.csv', 'r') as f:
    list_participate = f.readlines()
    list_participate = [x.strip().split(',') for x in list_participate]
    print(list_participate)


list_participate = list_participate[1:]
print(list_participate)


import win32com.client as win32
import os

def send_email(name,email):
  outlook = win32.Dispatch('outlook.application')
  mail = outlook.CreateItem(0)
  mail.To = email
  mail.Subject = 'Azure Ai Fundamentals'
  mail.Body = 'Thank you for your participation in the Azure Ai Fundamentals conducted by MLSA. Please find attached your certificate.'
  name = name
  ambassador = "Vedant Shukla"
  event = "Azure Ai Fundamentals"
  mail.HTMLBody = f'''
  <html><head></head><body>
    
      
    <table id="m_30972499125270521headerModule" width="600" cellspacing="0" cellpadding="0" border="0" align="center" style="
        max-width: 600px;
        width: 100%;
        background: rgb(248, 249, 250);
        --darkreader-inline-bgimage: initial;
        --darkreader-inline-bgcolor: #1b1e1f;
      " data-darkreader-inline-bgimage="" data-darkreader-inline-bgcolor="">
      <tbody>
        <tr>
          <td align="center" valign="top" style="padding: 32px 24px 22px 24px" class="m_30972499125270521header" id="m_30972499125270521header-logos">
            <table width="600" cellspacing="0" cellpadding="0" border="0" align="center" style="max-width: 600px; width: 100%">
              <tbody>
                <tr>
                  <td width="300" align="left" valign="top" style="
                      font-family: 'Google Sans', 'Roboto', Helvetica, Arial,
                        sans-serif;
                      font-size: 14px;
                      vertical-align: middle;
                    ">
                    <img src="https://ci4.googleusercontent.com/proxy/dplLFJdqtQ1_dzznV_S7BoxjtL_gnz1n_SAKc4_LRSWbATYDGOASSh0QV3mMpkyGumPQK0pM4n6P_7d2QegGQQX22KWurdz6EW02J8ns_xM0zeT_1Gulx0DGcSKtg93tNTVONXfZvSsa3hw-15cKHbfk7YIbqF4pwm0Jdte76cvIA3PsjLYp0snOimR-p5T7ueNIguzy9uqY=s0-d-e1-ft#https://images.ecomm.microsoft.com/cdn/mediahandler/azure-emails-templates/production/shared/images/templates/shared/microsoft-2x.png" alt="MLSA" class="m_30972499125270521logo m_30972499125270521no-arrow CToWUd" width="75%" height="100%" data-bit="iit">
                  </td>
                  <p style="
                                      padding: 5px;
                                      margin: 5px;
                                      overflow-wrap: normal;
                                      text-align: right;
                                      line-height: 1;
                                      font-size: 18px;
                                      font-family: 'Segoe UI Semibold',
                                        SegoeUISemibold, 'Segoe UI', SegoeUI,
                                        Roboto, 'Helvetica Neue', Arial,
                                        sans-serif;
                                      font-weight: 600;
                                      color: rgb(0, 120, 215);
                                      --darkreader-inline-color: #4bb2ff;
                                    " data-darkreader-inline-color="">
                                    <span class="il">Learn</span>
                                    <span class="il">Student</span>
                                    <span class="il">Ambassadors</span>
                                  </p>
                </tr>
              </tbody>
            </table>
          </td>
        </tr>
      </tbody>
    </table>
    <table id="m_30972499125270521heroModuleb29c98b4-7be1-44cf-b196-0ff1e11e3bbb" width="600" cellspacing="0" cellpadding="0" border="0" align="center" style="max-width: 600px; width: 100%">
      <tbody>
        <tr>
          <td align="center" class="m_30972499125270521hero-wrap" valign="top" style="
              padding: 0px;
              border-bottom: 1px solid rgb(215, 215, 215);
              --darkreader-inline-border-bottom: #3b4043;
            " data-darkreader-inline-border-bottom="">
            <table border="0" cellspacing="0" cellpadding="0">
              <tbody>
                <tr>
                  <td class="m_30972499125270521hero-image-mobile" id="m_30972499125270521emMobileHeroImageb29c98b4-7be1-44cf-b196-0ff1e11e3bbb" colspan="2" style="display: none">
                    <a target="_blank"><img class="m_30972499125270521no-arrow CToWUd" style="
                          width: 100%;
                          display: block;
                          margin: 0px;
                          border: 0px;
                          --darkreader-inline-border-top: initial;
                          --darkreader-inline-border-right: initial;
                          --darkreader-inline-border-bottom: initial;
                          --darkreader-inline-border-left: initial;
                        " src="https://www.gstatic.com/devrel-devsite/prod/veced1430215d0a0c094ac0570f79b1e47a9902cf6b60d19f36522e018b212f9e/cloud/images/social-icon-google-cloud-1200-630.png" alt="Google Cloud Logo" data-bit="iit" data-darkreader-inline-border-top="" data-darkreader-inline-border-right="" data-darkreader-inline-border-bottom="" data-darkreader-inline-border-left=""></a>
                  </td>
                </tr>

                <tr>
                  <td width="50%" style="vertical-align: middle" class="m_30972499125270521width100">
                    <table cellspacing="0" cellpadding="0" border="0" align="center" style="max-width: 600px; width: 100%">
                      <tbody>
                        <tr>
                          <td class="m_30972499125270521hero-text" style="
                              padding: 20px 10px 5px 40px;
                              text-align: left;
                              vertical-align: middle;
                            ">
                            <table cellspacing="0" cellpadding="0" border="0" align="center" style="max-width: 600px; width: 100%">
                              <tbody>
                                <tr>
                                  <td>
                                    <div id="m_30972499125270521hero-eyebrowb29c98b4-7be1-44cf-b196-0ff1e11e3bbb">
                                      <table cellspacing="0" cellpadding="0" border="0" align="center" style="max-width: 600px; width: 100%">
                                        <tbody>
                                          <tr>
                                            
                                          </tr>
                                        </tbody>
                                      </table>
                                    </div>
                                  </td>
                                </tr>
                                <tr>
                                  <td id="m_30972499125270521hero-titleb29c98b4-7be1-44cf-b196-0ff1e11e3bbb" style="
                                      color: rgb(60, 64, 67);
                                      font-family: 'Google Sans', Roboto,
                                        Helvetica, Arial, sans-serif;
                                      font-size: 32px;
                                      font-weight: bold;
                                      line-height: 40px;
                                      margin: 0px;
                                      padding: 0px 0px 20px;
                                      text-align: left;
                                      vertical-align: top;
                                      overflow-wrap: break-word;
                                      --darkreader-inline-color: #c0bab2;
                                      border-collapse: collapse !important;
                                    " data-darkreader-inline-color="">
                                    <div>
                                    Thank you for your participation in the
                                    {event}!
                                  </div>
                                  </td>
                                </tr>
                              </tbody>
                            </table>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </td>
                </tr>
                <tr>
                  <td width="45%" style="width: 240px; vertical-align: center" class="m_30972499125270521hom">
                    <div id="m_30972499125270521emDesktopHeroImageb29c98b4-7be1-44cf-b196-0ff1e11e3bbb">
                      <a target="_blank"><img class="m_30972499125270521no-arrow CToWUd" src="https://ci5.googleusercontent.com/proxy/MgsznD8aqCOrPjJs_4IXTrSgseeUmmGV6CbaNPgcZggxzDkdRv0e0fjxyJib-pVklnQtoLeM1vgwrvN30gamB0ferwg_-W-DhasJ3XRsuTWO6ay4o7fLAc-50sQ_iV7xsnvi_sCih6Cw=s0-d-e1-ft#https://imaginemedia.blob.core.windows.net/content/MLSA_Badge_GENERIC-d4ac807c84a4.png" width="100%" height="100%" alt="" style="
                            display: block;
                            border: 0px;
                            margin: 0px;
                            --darkreader-inline-border-top: initial;
                            --darkreader-inline-border-right: initial;
                            --darkreader-inline-border-bottom: initial;
                            --darkreader-inline-border-left: initial;
                          " data-bit="iit" data-darkreader-inline-border-top="" data-darkreader-inline-border-right="" data-darkreader-inline-border-bottom="" data-darkreader-inline-border-left=""></a>
                    </div>
                  </td>
                </tr>
              </tbody>
            </table>
          </td>
        </tr>
      </tbody>
    </table>
    <table id="m_30972499125270521greetingModulede787cf3-2bd8-4e67-95ab-7819270408fe" width="600" cellspacing="0" cellpadding="0" border="0" align="center" style="max-width: 600px; width: 100%">
      <tbody>
        <tr>
          <td align="center" valign="top" style="padding: 26px 40px 13px 40px" class="m_30972499125270521inner-container">
            <table width="600" cellspacing="0" cellpadding="0" border="0" align="center" style="max-width: 600px; width: 100%">
              <tbody>
                <tr>
                  <td style="
                      color: rgb(60, 64, 67);
                      font-family: 'Google Sans', Helvetica, Arial, sans-serif;
                      font-size: 16px;
                      line-height: 24px;
                      margin: 0px;
                      padding: 0px;
                      text-align: left;
                      font-weight: bold;
                      --darkreader-inline-color: #c0bab2;
                    " data-darkreader-inline-color="">
                    <div id="m_30972499125270521greetingContentde787cf3-2bd8-4e67-95ab-7819270408fe" style="
                        color: rgb(60, 64, 67);
                        font-family: 'Google Sans', Helvetica, Arial,
                          sans-serif;
                        font-size: 18px;
                        line-height: 24px;
                        margin: 0px;
                        padding: 0px;
                        text-align: left;
                        font-weight: bold;
                        word-break: break-word;
                        --darkreader-inline-color: #c0bab2;
                      " data-darkreader-inline-color="">
                      <div>Dear {name},</div>
                    </div>
                  </td>
                </tr>
              </tbody>
            </table>
          </td>
        </tr>
      </tbody>
    </table>
    <table id="m_30972499125270521textModule186afddd5-4aa5-479e-8405-026934e0891e" width="600" cellspacing="0" cellpadding="0" border="0" align="center" style="max-width: 600px; width: 100%">
      <tbody>
        <tr>
          <td align="center" valign="top" style="padding: 5px 40px 13px 40px" class="m_30972499125270521inner-container">
            <table width="600" cellspacing="0" cellpadding="0" border="0" align="center" style="max-width: 600px; width: 100%">
              <tbody>
                <tr>
                  <td style="
                      color: rgb(60, 64, 67);
                      font-family: 'Google Sans', Helvetica, Arial, sans-serif;
                      font-size: 16px;
                      font-weight: 400;
                      line-height: 24px;
                      margin: 0px;
                      padding: 0px;
                      text-align: left;
                      word-break: break-word;
                      --darkreader-inline-color: #c0bab2;
                    " data-darkreader-inline-color="">
                    <div id="m_30972499125270521textContent186afddd5-4aa5-479e-8405-026934e0891e">
                      <div>
                        Thank you so much for your interest in being a part of
                        the
                        <a style="
                            text-decoration: none;
                            color: rgb(66, 133, 244);
                            font-weight: normal;
                            --darkreader-inline-color: #4ba0f4;
                          " data-darkreader-inline-color=""><strong>{event}</strong></a>
                        program.
                        <br>
                        We appreciate your time and effort in completing the
                        program.
                        <br>
                        This email contains your certificate of participation.
                      </div>
                    </div>
                  </td>
                </tr>
              </tbody>
            </table>
          </td>
        </tr>
      </tbody>
    </table>

    <table id="m_30972499125270521textModule25f253c04-e2ce-4bd4-81e1-dd82e05bce0c" width="600" cellspacing="0" cellpadding="0" border="0" align="center" style="max-width: 600px; width: 100%">
      <tbody>
        <tr>
          <td align="center" valign="top" style="padding: 16px 40px 13px 40px" class="m_30972499125270521inner-container">
            <table width="600" cellspacing="0" cellpadding="0" border="0" align="center" style="max-width: 600px; width: 100%">
              <tbody>
                <tr>
                  <td style="
                      color: rgb(60, 64, 67);
                      font-family: 'Google Sans', Helvetica, Arial, sans-serif;
                      font-size: 16px;
                      font-weight: 400;
                      line-height: 24px;
                      margin: 0px;
                      padding: 0px;
                      text-align: left;
                      word-break: break-word;
                      --darkreader-inline-color: #c0bab2;
                    " data-darkreader-inline-color="">
                    <div id="m_30972499125270521textContent25f253c04-e2ce-4bd4-81e1-dd82e05bce0c">
                      <div>
                        Please feel free to reply back to this email should
                        you have any questions.
                      </div>
                    </div>
                  </td>
                </tr>
              </tbody>
            </table>
          </td>
        </tr>
      </tbody>
    </table>

    <table id="m_30972499125270521textModule25f253c04-e2ce-4bd4-81e1-dd82e05bce0c" width="600" cellspacing="0" cellpadding="0" border="0" align="center" style="max-width: 600px; width: 100%">
      <tbody>
        <tr>
          <td align="center" valign="top" style="padding: 8px 40px 13px 40px" class="m_30972499125270521inner-container">
            <table width="600" cellspacing="0" cellpadding="0" border="0" align="center" style="max-width: 600px; width: 100%">
              <tbody>
                <tr>
                  <td style="
                      color: rgb(60, 64, 67);
                      font-family: 'Google Sans', Helvetica, Arial, sans-serif;
                      font-size: 16px;
                      font-weight: 400;
                      line-height: 24px;
                      margin: 0px;
                      padding: 0px;
                      text-align: left;
                      word-break: break-word;
                      --darkreader-inline-color: #c0bab2;
                    " data-darkreader-inline-color="">
                    <div id="m_30972499125270521textContent25f253c04-e2ce-4bd4-81e1-dd82e05bce0c">
                      <div>
                        Looking forward to seeing you at further programs,<br>
                        <div style="padding-top: 4px"></div>
                        <strong>
                        {ambassador},
                      </strong>
                      <strong>Microsoft Learn Student Ambassadors.</strong>
                        
                        
                        
                        
                      </div>
                    </div>
                  </td>
                </tr>
              </tbody>
            </table>
          </td>
        </tr>
      </tbody>
    </table>



  </body></html>
  '''

  # To attach a file to the email (optional):
  # Relative path is not working for some reason so I have used the absolute path to the ./Output/PDF/ folder
  attachment  = r"C:\Users\Sunit Dwivedi\Documents\InstantCDDVD\gssoc\MLSA-Certificate-Generator_Email-Sender\Output\PDF\\" + name + '.pdf' 
  mail.Attachments.Add(attachment)

  mail.Send()
  




import win32com.client as win32
import os
from pdf2image import convert_from_path

# Constants
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"


def send_event_cert_email(name,email):
    	
		outlook = win32.Dispatch('outlook.application')
		mail = outlook.CreateItem(0)
		mail.To = email
		event = "Azure Ai Fundamentals Ambassador Challenge" #editable
		mail.Subject = 'Thank you for completing the ' + event + ' [MLSA]' # editable

		# Comment the unrequired theme
		template_theme = "image005.png" #dark
		# template_theme = "image004.png" #light

		attachment  = os.getcwd()+"\Output\PDF\\" + name + '.pdf'

		# Store Pdf with convert_from_path function
		images = convert_from_path(attachment,poppler_path= os.getcwd()+ r"\poppler-23.05.0\Library\bin")
		images[0].save('Attendee_cert.jpg', 'JPEG')
		image_cert = os.getcwd()+"\Attendee_cert.jpg"

		mail.Attachments.Add(attachment)
		attachment_cert_img = mail.Attachments.Add(image_cert)
		attachment_cert_img.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, 'image1')

		name = name # Of the respective attendee
		ambassador = "Vedant Shukla" # Of the respective Ambassador (editable)
		ambassadorFirstname = ambassador.split()[0]
		ambassadorLastname = ambassador.split()[1]
		ambassadorContact = "" # +countryCode phoneNumber (editable)
		ambassadorTwitter = "https://www.twitter.com/" # rahulvijan1" # editable
		ambassadorLinkedIn = "https://www.linkedin.com/in/vedant5" # rahulv24" # editable
		ambassadorFacebook = "https://www.facebook.com/" # rahul.vijan.24" # editable
		mail.HTMLBody =  '''
			<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40">
			<head>
			<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
			<meta name="Generator" content="Microsoft Word 15 (filtered medium)">
			<!--[if !mso]><style>v\:* {behavior:url(#default#VML);}
			o\:* {behavior:url(#default#VML);}
			w\:* {behavior:url(#default#VML);}
			.shape {behavior:url(#default#VML);}
			</style><![endif]--><style><!--
			/* Font Definitions */
			@font-face
				{font-family:Mangal;
				panose-1:0 0 4 0 0 0 0 0 0 0;}
			@font-face
				{font-family:"Cambria Math";
				panose-1:2 4 5 3 5 4 6 3 2 4;}
			@font-face
				{font-family:Calibri;
				panose-1:2 15 5 2 2 2 4 3 2 4;}
			@font-face
				{font-family:"Segoe UI";
				panose-1:2 11 5 2 4 2 4 2 2 3;}
			@font-face
				{font-family:"Segoe UI Semibold";
				panose-1:2 11 7 2 4 2 4 2 2 3;}
			@font-face
				{font-family:"Segoe Pro Semibold";}
			/* Style Definitions */
			p.MsoNormal, li.MsoNormal, div.MsoNormal
				{margin:0cm;
				font-size:11.0pt;
				font-family:"Calibri",sans-serif;}
			a:link, span.MsoHyperlink
				{mso-style-priority:99;
				color:#0563C1;
				text-decoration:underline;}
			p.MsoListParagraph, li.MsoListParagraph, div.MsoListParagraph
				{mso-style-priority:34;
				margin-top:0cm;
				margin-right:0cm;
				margin-bottom:0cm;
				margin-left:36.0pt;
				font-size:11.0pt;
				font-family:"Calibri",sans-serif;}
			p.Bodycopy, li.Bodycopy, div.Bodycopy
				{mso-style-name:"Body copy";
				mso-style-priority:1;
				margin:0cm;
				line-height:12.0pt;
				font-size:9.0pt;
				font-family:"Segoe UI",sans-serif;
				color:#505050;}
			span.EmailStyle22
				{mso-style-type:personal-compose;
				font-family:"Segoe UI",sans-serif;
				color:windowtext;}
			.MsoChpDefault
				{mso-style-type:export-only;
				font-size:10.0pt;
				font-family:"Calibri",sans-serif;}
			@page WordSection1
				{size:612.0pt 792.0pt;
				margin:72.0pt 72.0pt 72.0pt 72.0pt;}
			div.WordSection1
				{page:WordSection1;}
			/* List Definitions */
			@list l0
				{mso-list-id:22288891;
				mso-list-template-ids:-362500670;}
			@list l0:level1
				{mso-level-tab-stop:36.0pt;
				mso-level-number-position:left;
				text-indent:-18.0pt;}
			@list l0:level2
				{mso-level-tab-stop:72.0pt;
				mso-level-number-position:left;
				text-indent:-18.0pt;}
			@list l0:level3
				{mso-level-tab-stop:108.0pt;
				mso-level-number-position:left;
				text-indent:-18.0pt;}
			@list l0:level4
				{mso-level-tab-stop:144.0pt;
				mso-level-number-position:left;
				text-indent:-18.0pt;}
			@list l0:level5
				{mso-level-tab-stop:180.0pt;
				mso-level-number-position:left;
				text-indent:-18.0pt;}
			@list l0:level6
				{mso-level-tab-stop:216.0pt;
				mso-level-number-position:left;
				text-indent:-18.0pt;}
			@list l0:level7
				{mso-level-tab-stop:252.0pt;
				mso-level-number-position:left;
				text-indent:-18.0pt;}
			@list l0:level8
				{mso-level-tab-stop:288.0pt;
				mso-level-number-position:left;
				text-indent:-18.0pt;}
			@list l0:level9
				{mso-level-tab-stop:324.0pt;
				mso-level-number-position:left;
				text-indent:-18.0pt;}
			@list l1
				{mso-list-id:178593379;
				mso-list-template-ids:200835744;}
			ol
				{margin-bottom:0cm;}
			ul
				{margin-bottom:0cm;}
			--></style><!--[if gte mso 9]><xml>
			<o:shapedefaults v:ext="edit" spidmax="1026" />
			</xml><![endif]--><!--[if gte mso 9]><xml>
			<o:shapelayout v:ext="edit">
			<o:idmap v:ext="edit" data="1" />
			</o:shapelayout></xml><![endif]-->
			</head>
			<body lang="EN-IN" link="#0563C1" vlink="#954F72" style="word-wrap:break-word">
			<div class="WordSection1">
			<table class="MsoNormalTable" border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100.0%;background:white">
			<tbody>
			<tr style="height:888.3pt">
			<td style="padding:0cm 0cm 15.0pt 0cm;height:888.3pt">
			<div align="center">
			<table class="MsoNormalTable" border="0" cellspacing="0" cellpadding="0" style="max-width:496.25pt">
			<tbody>
			<tr>
			<td style="border:solid #E3E3E3 1.0pt;padding:0cm 0cm 0cm 0cm">
			<table class="MsoNormalTable" border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100.0%">
			<tbody>
			<tr>
			<td style="padding:11.25pt 15.0pt 18.0pt 17.25pt">
			<table class="MsoNormalTable" border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100.0%">
			<tbody>
			<tr>
			<td valign="top" style="padding:0cm 0cm 0cm 0cm">
			<p class="MsoNormal" style="mso-line-height-alt:.75pt"><span style="font-size:1.0pt"><img width="201" height="38" style="width:2.0937in;height:.3958in" id="Picture_x0020_2" src="https://raw.githubusercontent.com/rv2442/MLSA-Certificate-Generator_Email-Sender/master/Data/image002.png"></span><span style="font-size:1.0pt"><o:p></o:p></span></p>
			</td>
			<td width="10" style="width:7.5pt;padding:0cm 0cm 0cm 0cm"></td>
			<td valign="top" style="padding:0cm 0cm 0cm 0cm">
			<p class="MsoNormal" align="right" style="text-align:right;line-height:21.0pt"><span style="font-size:15.0pt;font-family:&quot;Segoe UI&quot;,sans-serif;color:#505050">&nbsp;<o:p></o:p></span></p>
			</td>
			</tr>
			</tbody>
			</table>
			</td>
			</tr>
			<tr>
			<td style="padding:0cm 0cm 0cm 0cm">
			<p class="MsoNormal" align="center" style="text-align:center"><a href="https://studentambassadors.microsoft.com/"><span style="color:#5C2D91;text-decoration:none"><img border="0" width="659" height="220" style="width:6.8645in;height:2.2916in" id="Picture_x0020_1" src="https://raw.githubusercontent.com/rv2442/MLSA-Certificate-Generator_Email-Sender/master/Data/'''+template_theme+'''"></span></a><o:p></o:p></p>
			</td>
			</tr>
			<tr>
			<td style="padding:12.75pt 16.5pt 18.75pt 16.5pt">
			<table class="MsoNormalTable" border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100.0%">
			<tbody>
			<tr>
			<td style="padding:0cm 0cm 12.0pt 0cm">
			<p class="MsoNormal" align="center" style="margin-top:6.0pt;text-align:center;line-height:22.0pt;mso-line-height-rule:exactly">
			<span style="font-size:20.0pt;font-family:&quot;Segoe UI Semibold&quot;,sans-serif;color:#2F2F2F">Thank you for your participation in the '''+event+''' program!<o:p></o:p></span></p>
			</td>
			</tr>

			<tr>
			<td style="padding:0cm 15.0pt 18.75pt 15.0pt"> 

				<!--
					Add your Message below this
				-->
			<p class="MsoNormal" align="center" style="text-align:left"><span style="font-size:12.0pt;font-family:&quot;Segoe UI&quot;,sans-serif;color:#505050">Dear '''+name+''',<o:p></o:p></span></p>

			<p class="MsoNormal" align="center" style="text-align:left"><span style="font-size:12.0pt;font-family:&quot;Segoe UI&quot;,sans-serif;color:#505050">We appreciate your time and effort in completing the program.
	This email contains your certificate of participation.<o:p></o:p></span></p>
			
			<p class="MsoNormal" align="center" style="text-align:left"><span style="font-size:12.0pt;font-family:&quot;Segoe UI&quot;,sans-serif;color:#505050">
	Please feel free to reply back to this email should you have any questions. Looking forward to seeing you at future events.</span></p>

			</td>
			</tr> 
			</tbody>
			</table>
			</div>
			

			
			<div style="text-align: center;">
			<img src="cid:image1" width="300"/>
			</div> <br><br><br><br>

			</td>
			</tr>

			<tr>
			<td style="padding:0cm 0cm 0cm 0cm;padding-bottom:22px">
			<table class="MsoTableGrid" border="0" cellspacing="0" cellpadding="0" width="378" style="width:10.0cm;border-collapse:collapse">
			<tbody>
			<tr style="page-break-inside:avoid;height:45.0pt">
			<td width="126" valign="top" style="width:94.5pt;padding:0cm 0cm 0cm 0cm;height:45.0pt;padding-left:17px">
			<p class="Bodycopy"><a href="https://studentambassadors.microsoft.com/"><span style="color:#505050;text-decoration:none"><img border="0" width="109" height="95" style="width:1.1354in;height:.9895in" id="Picture_x0020_14" src="https://raw.githubusercontent.com/rv2442/MLSA-Certificate-Generator_Email-Sender/master/Data/image008.png"></span></a><o:p></o:p></p>
			</td>
			<td width="252" style="width:189.0pt;padding:0cm 0cm 0cm 0cm;height:45.0pt">
			<p class="Bodycopy"><span style="font-family:&quot;Segoe Pro Semibold&quot;,sans-serif">'''+ambassadorFirstname+''' '''+ambassadorLastname+'''<o:p></o:p></span></p>
			<p class="Bodycopy">Microsoft<span style="font-family:&quot;Segoe UI Semibold&quot;,sans-serif">
			</span>Learn<span style="font-family:&quot;Segoe UI Semibold&quot;,sans-serif"> </span>Student<span style="font-family:&quot;Segoe UI Semibold&quot;,sans-serif">
			</span>Ambassador<span style="font-family:&quot;Segoe UI Semibold&quot;,sans-serif"><o:p></o:p></span></p>
			<p class="Bodycopy">Mobile: '''+ambassadorContact+'''<o:p></o:p></p>
			<p class="Bodycopy">'''+ambassadorFirstname+'''.'''+ambassadorLastname+'''@studentambassadors.com<o:p></o:p></p>
			<p class="Bodycopy"><o:p>&nbsp;</o:p></p>
			<p class="Bodycopy"><a href="'''+ambassadorTwitter+'''"><span style="color:#505050;mso-fareast-language:JA;text-decoration:none"><img border="0" width="22" height="22" style="width:.2395in;height:.2395in" id="Picture_x0020_13" src="https://raw.githubusercontent.com/rv2442/MLSA-Certificate-Generator_Email-Sender/master/Data/image010.png"></span></a>&nbsp;&nbsp;
			<a href="'''+ambassadorFacebook+'''"><span style="color:#505050;mso-fareast-language:JA;text-decoration:none"><img border="0" width="23" height="23" style="width:.2395in;height:.2395in" id="Picture_x0020_12" src="https://raw.githubusercontent.com/rv2442/MLSA-Certificate-Generator_Email-Sender/master/Data/image012.png"></span></a>&nbsp;&nbsp;
			<a href="'''+ambassadorLinkedIn+'''"><span style="color:#505050;mso-fareast-language:JA;text-decoration:none"><img border="0" width="23" height="23" style="width:.2395in;height:.2395in" id="Picture_x0020_11" src="https://raw.githubusercontent.com/rv2442/MLSA-Certificate-Generator_Email-Sender/master/Data/image014.png"></span></a>&nbsp;&nbsp;<o:p></o:p></p>
			
			</td>
			</tr>
			</tbody>
			</table>
			</td>
			</tr>
			</tbody>
			</table>
			</td>
			</tr>
			</tbody>
			</table>
			</td>
			</tr>
			</tbody>
			</table>
			</div>
			</td>
			</tr>
			</tbody>
			</table>
			</div>
			</body>
			</html>
			'''
		mail.Send()
		os.remove(image_cert)


import time

for participant in list_participate:
    name = participant[0]
    email = participant[1]
    print(name,email)
    try:
        send_event_cert_email(name,email)
        print("done \n")
        time.sleep(2)
    except: 
        print("Error \n")  
