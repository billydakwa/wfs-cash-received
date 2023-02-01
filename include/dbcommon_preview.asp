<%
response.Charset="##@BUILDER.strCharset s##"

suggestAllContent=true

Session.LCID = ##@BUILDER.strLCID##
session.codepage=##@BUILDER.nCodepage##

strConnection = "##@BUILDER.strODBCString s##"
strLeftWrapper="##@BUILDER.cLeftWrap s##"
strRightWrapper="##@BUILDER.cRightWrap s##"
bSubqueriesSupported=true

cFrom 					= "##@BUILDER.strFromEmail s##"
cSmtpServer 			= "##@BUILDER.strSMTPServer##"
cSmtpPort 				= "##@BUILDER.strSMTPPort##"
cSMTPUser				= "##@BUILDER.strSMTPUser##"
cSMTPPassword			= "##@BUILDER.strSMTPPassword##"

##if @BUILDER.bDynamicPermissions##
	gPermissionsRefreshTime=0
	gPermissionsRead=false
##endif##

%>
