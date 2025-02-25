#include "include\gastro.h"
#define KSeFErrorCode_OK 0 && brak błędu
#define KSeFErrorCode_Unknown 1001 && nieokreślony błąd


#if .f.
	** https://epuap.gov.pl/wps/portal/strefa-urzednika/inne-systemy/crwde/!ut/p/a1/jVJtcpswFLxKL9BKQFLDT0ISBo_j1rHrBP_RKCA7CiAxQkKFm3R6nfZeBfwFHjsNwzAjvd19u-8BVuAZrBgu6QZLyhlO2_PqK4K2HfiuB8fQuzehewVv3OnSMeAMNoCwARiOeWWYczj27ckNdL0HZ7l4GFnwG_wYH1543P_yn8DqXUhg7ADvWdwqXPYQNiZHyHTMwHdsOPFdcwTdmWXPru2ZAU0LPJWUaDBvlXxBY7SmqSTCe8VsQ2IQSqHIoZZQIhQjCSq4kFxjRjEI4aEsq7yjC4x0LRBTGREgnC8eg6nfgnCLshpLewJfrwsi-xIdHUV1hXAiFU5Z1YOnuJC940m7GEuMcvUCwlt3cXei2PPTq4iB3qb5pDSjMtqHP9eK4VrjfqohF4QGPMbprgqSkuiksHW102pvMx6TXsOEpypjuBu0esNRg_psnMl0DN0Wt51I_Mh1cItaaNPVtKBzYAquvZRGCYmXOFX9joz87A8jF6S8MGuWoJjLqlnSYLmRpCVZ4JfBhDuffYIS9Z9fMdfkU87__j6nP1z-scE-HIq4YnKf7qCgq0JTItPmpyQ7qV25eOUahLrmoprQQn55K3KQZz-eIf2ejcvp9QffzT8GQkOy/dl5/d5/L2dBISEvZ0FBIS9nQSEh/pw/Z7_292IG980LGA270AQ38Q58Q1023/act/id=0/545747626635/=/#Z7_292IG980LGA270AQ38Q58Q1023
	** #include "vfp2c.h"
	clear all
	LoadEnvironment()
	local ;
		iddok as long,;
		result as integer,;
		nrKSeF as string,;
		oxml as MSXML2,;
		upo as object

	set step on
	local KSeF_sql as KSeFSQL of "prgs\common\ksefapi.prg"
	KSeF_sql = newobject("KSeFSQL","prgs\common\ksefapi.prg")

	local KSeF_api as KSeFApi of "prgs\common\ksefapi.prg"
	KSeF_api = newobject("KSeFApi","prgs\common\ksefapi.prg")
	if (vartype(KSeF_api)!="O")
		return
	endif
	with KSeF_api as KSeFApi of prgs\common\KSeFApi.prg
		local _oe as exception
		_oe = null
		try
			debugout .GetTimeZone()

			debugout "2023-05-16T14:43:25.246Z"
			debugout transform(.GetUtcTimeString("2023-05-16T14:43:25.246Z"),"@YL"))
			debugout transform(.GetUtcTimeString("2023-05-16T14:43:25.246Z"),"@YS"))
			debugout transform(.GetUtcTimeString("2023-05-16T14:43:25.246Z"),"@D"))


			bindevent(KSeF_api,"OnSendInvoice",KSeF_sql,"OnSaveResponse")
			bindevent(KSeF_api,"OnSendInvoice",KSeF_sql,"OnThrowIfException",1)

			bindevent(KSeF_api,"OnSaveResponse",KSeF_sql,"OnSaveResponse")

			bindevent(KSeF_api,"OnInvoiceQuery",KSeF_sql,"OnSaveResponse")
			bindevent(KSeF_api,"OnInvoiceQuery",KSeF_sql,"OnThrowIfException",1)

			bindevent(KSeF_api,"OnSessionStatus",KSeF_sql,"OnSaveResponse")
			bindevent(KSeF_api,"OnSessionStatus",KSeF_sql,"OnThrowIfException",1)

			bindevent(KSeF_api,"OnCommonStatus",KSeF_sql,"OnSaveResponse")
			bindevent(KSeF_api,"OnCommonStatus",KSeF_sql,"OnThrowIfException",1)

			bindevent(KSeF_api,"OnValidateInvoice",KSeF_sql,"OnValidateInvoice")
			bindevent(KSeF_api,"OnValidateInvoice",KSeF_sql,"OnThrowIfException",1)

			bindevent(KSeF_api,"OnGetInvoice",KSeF_sql,"OnGetInvoice")
			bindevent(KSeF_api,"OnGetInvoice",KSeF_sql,"OnThrowIfException",1) && to drugie

			bindevent(KSeF_api,"OnSaveInvoiceStatus",KSeF_sql,"OnSaveResponse")
			bindevent(KSeF_api,"OnSaveInvoiceStatus",KSeF_sql,"OnThrowIfException")
			*bindevent(KSeF_api,"OnSaveInvoiceStatus",KSeF_sql,"OnSaveInvoiceStatus",1) &&to na końcu

			bindevent(KSeF_api,"OnConnect",KSeF_sql,"OnConnect",1)
			bindevent(KSeF_api,"OnConnect",KSeF_sql,"OnThrowIfException")

			bindevent(KSeF_api,"OnHtmlView",KSeF_sql,"OnHtmlView",1)
			bindevent(KSeF_api,"OnHtmlView",KSeF_sql,"OnThrowIfException")

			bindevent(KSeF_api,"OnOnlineInvoiceGet",KSeF_sql,"OnSaveResponse")
			**bindevent(KSeF_api,"OnOnlineInvoiceGet",KSeF_sql,"OnThrowIfException",1) && to drugie
			bindevent(KSeF_api,"OnSaveResponse",KSeF_sql,"OnSaveResponse")

			.SqlGetConfiguration()
			.Content.SelectSingleNode("./Configuration/UserAgent").text = "GastroPOS"
			encrypt_conf = .Content.xml && CryptAPI(strconv(.Content.xml,16),_screen.gastro.hash,0)
			.LoadConfiguration(encrypt_conf)

			*.connect()

			*.SQLiteQuery("select * from session","ses")
			.SQLiteQuery("select * from communication","cm")

			#if .f. && pobranie zakresu faktury
				set step on
				local reksef as stroing
				refksef = "1894466533-20230522-E64751-514C0B-C8"
				=.InvoicesDownloadDetail("2023-04-20", "2023-05-25", 2,"2023/05/0001",refksef)
				.ThrowIfException("Błąd pobierania zakresu faktur")
				.SaveBodyResponseToFile("InvoicesDownloadRangeDate_DetailZakup.xml")

				=.OnlineInvoiceGet(refksef)
				.ThrowIfException("Błąd pobierania faktury")
				.SaveBodyResponseToFile("OnlineInvoiceGet_faktua.xml")

				.InvoiceXMLtoHTML(.ElementResponse(""))

				exit
				.CommonSessionStatus("20230519-SE-5A825268CA-9C08415299-6F")
				exit
				= .InvoicesDownloadRangeDate("2023-04-20", "2023-05-25", 1,30,0)
				.ThrowIfException("Błąd pobierania zakresu faktur")
				.SaveBodyResponseToFile("InvoicesDownloadRangeDate_30-0.xml")

				= .InvoicesDownloadRangeDate("2023-04-20", "2023-05-25", 1,30,1)
				.ThrowIfException("Błąd pobierania zakresu faktur")
				.SaveBodyResponseToFile("InvoicesDownloadRangeDate30-1.xml")


				exit
			#endif

			#if .f.
				.SetIDSaleDocument("650650000028922")
				raiseevent(KSeF_sql,"OnGetInvoice",KSeF_api)
				.InvoiceXMLtoHTML(.ElementResponse(""))
				.SaveBodyResponseToFile(transform(.IDSaleDocument)+".html")
				exit
			#endif

			#if .f.
				.SetIDSaleDocument("650650000028922")
				.SendInvoiceXML()
				.InvoiceStatusFromResponse()
			#endif
			
			#if (.f.)
				.SessionStatus("20230516-SE-6F09671EF5-B06A8521F5-77")
				** sesje zamknięte
				** 16.05.2023

				.CommonSessionStatus("20230516-SE-E7C44D5531-CDC1CD8ADB-F2")
				.CommonSessionStatus("20230516-SE-D0F34CB834-84A00B25D7-BC")
				.CommonSessionStatus("20230516-SE-11CE90DABA-FA2ECC2050-8D")
				.CommonSessionStatus("20230516-SE-24D7C06E81-632F0F3BAC-D4")
				.CommonSessionStatus("20230516-SE-EEEBC5F6FC-BCAF417D3A-05")
				** 15.05.2023
				.CommonSessionStatus("20230515-SE-4BFB8A2357-31F577B273-3D")
				.CommonSessionStatus("20230515-SE-D53C2A86E4-84BD508379-93")
				exit
				.InvoiceStatusFromReferenceNumber("20230516-EE-C3304809C9-26EDCEFAFB-A3")
				exit

				raiseevent(KSeF_sql,"OnGetInvoice",KSeF_api)
				.InvoiceXMLtoHTML(.ElementResponse(""))
				.SendInvoiceXML()

				if (KSeFErrorCode_OK == result)
					nrKSeF = "1147581153-20230515-67B49B-B49FA3-87"
					.InvoiceStatusFromReferenceNumber(nrKSeF)
					.InvoiceStatusFromReferenceNumber("20230512-EE-B3D5A993BA-4DC981EA40-C7")
					nrKSeF = "1147581153-20230510-580D33-BC04BC-CE"
					if (KSeFErrorCode_OK == .OnlineInvoiceGet(nrKSeF))
						.InvoiceXMLtoHTML(.ElementResponse("")) &&UTF-8 -> DBCS
					endif

					nrKSeF = "20230421-SE-D0EF904AA0-9D476365E8-7E"

					result = .SessionStatus(nrKSeF)
					result = .CommonSessionStatus(nrKSeF)

					nrKSeF = "20230510-SE-61376F07E6-8893D1BEF3-1D"  && tu są faktury
					result = .SessionStatus(nrKSeF)
					result = .CommonSessionStatus(nrKSeF)

					if (KSeFErrorCode_OK == result)
						upo = .ElementResponse("/upo")
						if (!isnull(upo) and !empty(upo))
							local oupo as string &&base64
							oupo = strconv(upo,14)
							oxml = .loadxml(oupo)
							debugout("Decode in client")
							debugout(replicate("_",100))
							debugout(oxml.xml)
						endif
					endif


					result = .InvoicesDownloadRangeDate("2023-04-20", "2023-05-25", 1)
					if (result==KSeFErrorCode_OK)
						.SaveBodyResponseToFile("InvoicesDownloadRangeDate.xml")
					endif

				endif
			#endif
		catch to _oe
			ZLogEx(_oe)
			_screen.gastro.Showexception(_oe)
		finally
		endtry
	endwith
	? "---- full list"
	list dlls
	KSeF_sql = null
	KSeF_api = null
	release KSeF_sql
	release KSeF_api
	release all
	clear memory
	clear program
	*clear all

	list dlls &&[TO PRINTER [PROMPT] | TO FILE FileName [ADDITIVE]]   [NOCONSOLE]

	return

procedure LoadEnvironment()
	modi project c:\svn\trunk\gastro\gastrow.pjx nowait noshow
	public pgastro as integer
	public ManagerEnv as object

	ManagerEnv = newobject("CManagerEnvironment","prgs\CManagerEnvironment.prg")
	with ManagerEnv
		.ApplicationName = "GASTRO"
		.FileLog = ".\gastro.log"
	endwith
	restore from WYSZKOWIANKA.MEM additive

	with _screen.gastro.connection
		.odbc = alltrim(X_SQLSERWER)
		.userid = alltrim(X_SQLUSER)
		.password = alltrim(X_SQLPASSWORD)
	endwith

	_oe=null
	try
		_screen.gastro.ReloadVariablesFromSql()
	catch to _oe
		ZLogEx(_oe)
		_screen.gastro.DefaultPublicVars()
		with _screen.gastro.connection
			.odbc = alltrim(X_SQLSERWER)
			.userid = alltrim(X_SQLUSER)
			.password = alltrim(X_SQLPASSWORD)
		endwith
	endtry
endproc

#endif


define class KSeFDB as session
	#define virtual
	_memberdata = ;
		'<VFPData>'+;
		[<memberdata display="PosID" name="posid" type="property" />]+;
		;
		'<memberdata type="method" display="Init" name="init"/>'+;
		'<memberdata type="method" display="Destroy" name="destroy"/>'+;
		'<memberdata type="method" display="CheckExported" name="checkexported"/>'+;
		'<memberdata type="method" display="GetExchangedDocument" name="getexchangeddocument"/>'+;
		'<memberdata type="method" display="LoadConfiguration" name="loadconfiguration"/>'+;
		'<memberdata type="method" display="SetCompany" name="setcompany"/>'+;
		'<memberdata type="method" display="SetPOS" name="setpos"/>'+;
		'<memberdata type="method" display="Context" name="context"/>'+;
		'<memberdata type="method" display="ToWebPage" name="towebpage"/>'+;
		'<memberdata type="method" display="OpenPage" name="openpage"/>'+;
		'<memberdata type="method" display="GetXmlCursor" name="GetXmlCursor"/>'+;
		'<memberdata type="method" display="GetAuthentication" name="GetAuthentication"/>'+;		
		'<memberdata type="method" display="ShowInvoiceStatus" name="showinvoicestatus" />'+;
		'<memberdata type="method" display="CreateInvoiceSettings" name="createinvoicesettings" />'+;
		;
		'<memberdata type="method" display="OnUPOView" name="onupoview" />'+;
		'<memberdata type="method" display="OnHtmlView" name="onhtmlview" />'+;
		'<memberdata type="method" display="OnThrowIfException" name="onthrowifexception" />'+;
		'<memberdata type="method" display="OnConnect" name="onconnect" />'+;
		'<memberdata type="method" display="OnValidateInvoice" name="onvalidateinvoice" />'+;
		'<memberdata type="method" display="OnSaveResponse" name="onsaveresponse" />'+;
		'<memberdata type="method" display="OnGetInvoice" name="ongetinvoice" />'+;
		'<memberdata type="method" display="OnSaveInvoiceStatus" name="onsaveinvoicestatus" />'+;
		'<memberdata type="method" display="OnRollbackExport" name="onrollbackexport" />'+;
		'<memberdata type="method" display="OnSendInvoice" name="onsendinvoice" />'+;
		'<memberdata type="method" display="OnOnlineInvoiceGet" name="ononlineinvoiceget" />'+;
		'<memberdata type="method" display="OnInvoiceQuery" name="oninvoicequery" />'+;
		'<memberdata type="method" display="OnSessionStatus" name="onsessionstatus" />'+;
		'<memberdata type="method" display="OnCommonStatus" name="oncommonstatus" />'+;
		'<memberdata type="method" display="OnValidDatabase" name="oncommonstatus" />'+;
		'</VFPData>'

	upo = ""
	Priority = 0
	EnumPriority = null
	PosID = 0
	CompanyID = 0
	MailerConfiguration = null
	MailerConfigurationCrypt = null
	company = null &&załadowane dane firmy
	AuthenticationUser = null &&struct type,identifier
	
	protected upo as string
	protected save_id_session as integer

	procedure OnGetExportedDocumentList(oFilter as object, _crs_str as string@) as alias &&virtual
	procedure ShowInvoiceStatus(InvoiceSettings as object,KSeF_api as "KSeFApi" of "prgs\common\ksefapi.prg") &&virtual
	procedure OnRollbackExport(ksef as "KSeFApi" of "prgs\common\ksefapi.prg") as object &&virtual
	procedure CheckExported(_iddocument as integer, _datasession as integer) as string &&virtual
	procedure GetExchangedDocument(_iddocument as integer) as string &&virtual
	procedure OnHtmlView(ksef as "KSeFApi" of "prgs\common\ksefapi.prg") as void &&virtual
	procedure OnUPOView(ksef as "KSeFApi" of "prgs\common\ksefapi.prg") as void &&virtual
	procedure OnSaveInvoiceStatus(ksef as "KSeFApi" of "prgs\common\ksefapi.prg") as void &&virtual
	procedure OnGetInvoice(ksef as "KSeFApi" of "prgs\common\ksefapi.prg") as void &&virtual
	procedure OnConnect(ksef as "KSeFApi" of "prgs\common\ksefapi.prg") &&virtual
	procedure OnSaveResponse(ksef as "KSeFApi" of "prgs\common\ksefapi.prg") &&virtual
	procedure LoadConfiguration(oconf as MSXML2) as object
	procedure OnThrowIfException(ksef as "KSeFApi" of "prgs\common\ksefapi.prg") &&virtual
	procedure TerminateUnfinishedSessions(ksef as "KSeFApi" of "prgs\common\ksefapi.prg") &&virtual	
	
	*********************************************************************************************************************
	* Utworzenie pustego obiektu autentykacji
	* Klasy dziedziczące jedynie go uzupełnią
	*********************************************************************************************************************	
	procedure GetAuthentication() as XMLNode
		local Authentication as string
		local result as object
		result = GetXMLObject()	
		text to Authentication noshow textmerge pretext 15
				<Authentication>
					<Context>
						<Type>onip</Type>
						<Identifier/>
					</Context>
					<User>
						<Type/>
						<Identifier/>
					</User>
				</Authentication>
		ENDTEXT	
		with result as MSXML2
			.preservewhitespace = .f.
			.loadxml(Authentication)
			with (.ParseError)
				if (.ErrorCode!=0)
					set step on
					debugout .ErrorCode,.srcText,.reason
					=ThrowException("Błąd tworzenia obiektu autentykacji: "+chr(13)+.srcText,KSeFErrorCode_Unknown)
				endif
			endwith
		endwith
		return result			
	endproc


		****************************************************
		* Procedur bezpołączeniowa z KSeF
		* Działa na bazie danych i generuje kolekcję plików
		* HTML z audytami danych
		****************************************************
	procedure OnValidDatabase()

	procedure init()
		set safety off
		set console off
		set point to  '.'
		set multilocks on
		set exact on
		set reproces to 1
		set autosave on
		set null off
		set nulldisplay to ''
		try
			set enginebehavior 70
		catch
		endtry
		set century on
		set deleted on
		set date to ansi
		set sysmenu off
		set talk off
		set confirm on
		set currency right

		set currency to "zł"
		set refresh to 1
		set escape off
		set notify off
		set typeahead to 0
		set cpdialog off
		set seconds off
		set hours to 24

		this.AuthenticationUser = newobject("empty")
		addproperty(this.AuthenticationUser,"type","pesel")
		addproperty(this.AuthenticationUser,"identifier","")
		
		this.EnumPriority = newobject("Empty")
		addproperty(this.EnumPriority,"POS",1)
		addproperty(this.EnumPriority,"Company",2)
		this.save_id_session = this.datasessionid
	endproc

	procedure destroy()
		if (this.save_id_session!=this.datasessionid)
			set datasession to (this.save_id_session)
		endif
		close databases
		if (this.save_id_session!=this.datasessionid)
			set datasession to (this.datasessionid)
		endif
	endproc

	procedure SetPOS(_PosID as integer)
		with this
			.PosID = nvl(_PosID,0)
			.Priority = .EnumPriority.POS
		endwith
	endproc

	procedure SetCompany(_CompanyID as integer)
		with this
			.CompanyID = nvl(_CompanyID,0)
			.Priority = .EnumPriority.company
		endwith
	endproc

	procedure Context()
		with this
			if (.Priority==.EnumPriority.company)
				return "COMPANY:"+transform(.CompanyID)
			endif
			if (.Priority==.EnumPriority.POS)
				return "POS:"+transform(.PosID)
			endif
		endwith
		return "COMPANY:"+transform(.CompanyID)
	endproc

	hidden procedure ToWebPage(_href as string)
		try
			declare integer ShellExecute in shell32;
				integer hwnd,;
				string lpOperation,;
				string lpFile,;
				string lpParameters,;
				string lpDirectory,;
				integer nShowCmd
			ShellExecute(0,"open", _href,null, "", 3)
		catch to _oe
			_screen.gastro.Showexception(_oe)
		endtry
	endproc

	procedure OpenPage()
		local _oe as exception

		if (type("send_mail")=="C")
			if (send_mail=="SoftechSendMail")
				if (type("email")=="C" and type("mail_file")=="C")
					if (vartype(this.MailerConfiguration)!="C" or empty(this.MailerConfiguration))
						_screen.gastro.ShowMessage("Nie skonfigurowano mailera dla firmy")
						return
					endif
					send_mail = send_mail+".exe"
					if (!file(send_mail))
						_screen.gastro.ShowMessage("Brak oprogramowania do wysyłki email"+chr(13)+chr(13)+send_mail)
						return
					endif
					*local _cred as string,oxml as object
					_oe = null
					try
						local cResult as string
						cResult=[]
						*_screen.gastro.crypt.ExportSessionKey(_screen.gastro.hash,@cResult)

						*ssml_Initialize()
						*ssml_SetSessionKey(strconv(cResult,13))
						ssml_SetCredential(this.MailerConfigurationCrypt)
						ssml_SetMailFileDefinition(mail_file)
						if (ssml_Send()==1)
							_screen.gastro.ShowMessage("Poniższa informacja została wysłana na adres"+chr(13)+chr(13)+email)
						else
							_screen.gastro.ShowMessage("Problem z wysyłką maila, sprawdź konfigurację mailera",20,255)
						endif
					catch to _oe
						ZLogEx(_oe)
						_screen.gastro.Showexception(_oe)
					endtry
					return
				endif
			endif
			return
		endif

		if (type("_url")=="C")
			if (_url=="show_upo" and !empty(this.upo))
				_oe = null
				try
					KSeF_api =_screen.gastro.ksef
					with KSeF_api as KSeFApi of "prgs\common\ksefapi.prg"
						bindevent(KSeF_api,"OnUPOView",this,"OnUPOView",1)
						bindevent(KSeF_api,"OnUPOView",this,"OnThrowIfException")
						.UPOToHTML(this.upo)

					endwith
				catch to _oe
					ZLogEx(_oe)
				finally
					try
						unbindevents(KSeF_api,"OnUPOView",this,"OnUPOView")
						unbindevents(KSeF_api,"OnUPOView",this,"OnThrowIfException")
					catch to _oe
						ZLogEx(_oe)
					endtry
				endtry
			endif

			if (_url=="export_to_file")
				if (type("_file")=="C")
					strtofile(strconv(this.tag,14),_file)
					messagebox("Zapisano plik: "+_file,48,"Uwaga")
					return
				endif
			endif
			this.ToWebPage(_url)
		endif
	endproc

	procedure GetXmlCursor(_cursor as string)
		if (!used(_cursor) or reccount(_cursor)=0)
			return null
		endif

		local oxml as MSXML2
		local cOutput as string
		try
			go top in (_cursor)
			cOutput=""
			cursortoxml(_cursor, "cOutput",1,0,0,"")
			oxml = GetXMLObject()
			with oxml as MSXML2
				.preservewhitespace = .f.
				.loadxml(cOutput)
				with (.ParseError)
					if (.ErrorCode!=0)
						set step on
						debugout .ErrorCode,.srcText,.reason
						=ThrowException("Błąd deserializacji XML: "+chr(13)+.srcText,KSeFErrorCode_Unknown)
					endif
				endwith
			endwith
			use in (_cursor)
		catch to loException
			ZLogEx(loException)
		endtry
		return oxml
	endproc

	procedure CreateInvoiceSettings()
		local contractor as empty
		contractor = newobject("empty")
		addproperty(contractor,"IDDokumentu","")
		addproperty(contractor,"Nazwa_1","")
		addproperty(contractor,"Nazwa_2","")
		addproperty(contractor,"Nazwa_3","")
		addproperty(contractor,"Adres_1","")
		addproperty(contractor,"Adres_2","")
		addproperty(contractor,"NIP","")
		addproperty(contractor,"email","")
		return contractor
	endproc
enddefine

***********************************************************************
* Author.....: Piotr Kuliński (c) piotr.kulinski@gmail.com
* Date.......: 2023.05.16 09:06:31
* Comment....: Klasa odbiorcy zdarzeń KSEF.
*              Jej zadaniem jest zapis i odczyt danych z SQL.
***********************************************************************
define class KSeFSQL as KSeFDB of "prgs\common\ksefapi.prg"
	#define override

	_memberdata = ;
		'<VFPData>'+;
		'</VFPData>'

	procedure OnThrowIfException(ksef as "KSeFApi" of "prgs\common\ksefapi.prg") &&override
		ksef.ThrowIfException("Błąd podczas wymiany danych z systemem KSeF (v."+transform(ksef.ApiVersion)+")")
	endproc

	procedure OnConnect(ksef as "KSeFApi" of "prgs\common\ksefapi.prg") &&override
		ksef.sessionReferenceNumber = ksef.XmlGetValue(ksef.Content,"/ApiResponse/session","")
		* to sesja do kontroli
	endproc

	procedure OnValidateInvoice(ksef as "KSeFApi" of "prgs\common\ksefapi.prg" ) &&override
		ksef.ValidateInvoice()
	endproc

	*/**
	* Pobranie informacji o wyeksportowanych dokumentach
	* oFilter = createobject("empty")
	* addproperty(oFilter,"Data",[2023-10-01])
	* addproperty(oFilter,"IDKasy",0)
	* addproperty(oFilter,"Mask",6)
	*/
	procedure OnGetExportedDocumentList(oFilter as object,_crs_str as string@) as alias &&override
		local ;
			_oe as exception,;
			SqlCommand as string,;
			_CurrentCursorName as string

		if (oFilter.SessionID!=-1)
			_session = this.datasessionid
			set datasession to (oFilter.SessionID)
		endif
		_CurrentCursorName= iif(empty(_crs_str),[OnGetExportedDocumentList],_crs_str)
		_oe = null
		try
			text to SqlCommand noshow textmerge pretext 15
				exec [dbo].[GetKSeFSendingInvoice]
					@odDnia='<<oFilter.Data>>',
					@kasaID=<<oFilter.IDKasy>>,
					@filter=<<oFilter.Mask>>,
					@resultFormater=<<oFilter.ResultFormater>>
			ENDTEXT
			=SqlCommandExecuteException(_screen.gastro.ConnectionHandle,@SqlCommand,_CurrentCursorName)
		catch to _oe
			ZLogEx(_oe)
		finally
			if(!used(_CurrentCursorName))
				_CurrentCursorName=[]
			endif
			if (oFilter.SessionID!=-1)
				this.datasessionid = _session
				set datasession to (_session)
			endif
		endtry
		return _CurrentCursorName
	endproc

	procedure OnSaveResponse(ksef as "KSeFApi" of "prgs\common\ksefapi.prg") &&override
		local ;
			_oe as exception,;
			SqlCommand as string,;
			_CurrentCursorName as string
		_CurrentCursorName= "response"
		_oe = null
		try

			***********************************************************************************
			* procedura zapisująca response
			* Szczegóły w SQL, w zależności od typu odpowiedzi procedura może oznaczać
			* dokumenty jako wyeksportowane, przypisywać im UPO i inne informacje
			***********************************************************************************
			text to SqlCommand noshow textmerge pretext 15
				exec [dbo].[KSeF_OnSaveResponse] @response = '<<sqltekst(ksef.Content.xml)>>',@KodSystemowy=<<ksef.ApiVersion>>
			ENDTEXT
			=SqlCommandExecuteException(_screen.gastro.ConnectionHandle,@SqlCommand,_CurrentCursorName)
		catch to _oe
			ZLogEx(_oe)
		finally
			=CrsClose(_CurrentCursorName)
		endtry
	endproc

	procedure OnGetInvoice(ksef as "KSeFApi" of "prgs\common\ksefapi.prg") as void &&override
		local ;
			SqlCommand as string,;
			_value as string,;
			_length as integer,;
			CurrentCursorName as string,;
			_ActualCursorName as string,;
			_value as string,;
			_exp as string,;
			_oe as exception

		_oe = null
		try
			_CurrentCursorName= "invoice"
			text to SqlCommand noshow textmerge pretext 15
            	exec [dbo].[ConvertToInvoiceKSef]
					 @id=<<transform(ksef.IDSaleDocument)>>
					,@kodSystemowy=<<transform(ksef.APIVersion)>>
			ENDTEXT
			=SqlCommandExecuteException(_screen.gastro.ConnectionHandle,@SqlCommand,_CurrentCursorName)
			store "" to _value,_exp
			I=0
			_ActualCursorName = _CurrentCursorName
			do while (used(_ActualCursorName))
				select (_ActualCursorName)
				scan
					_exp = curval(field(1))
					if ("C"==vartype(_exp))
						&& to bardzo dziwne, czasem pobrany XML ma w dowolnym miejscu NULL-a
						_value = _value + strtran(_exp,chr(0),"")
					endif
				endscan
				use
				I = I + 1
				_ActualCursorName = _CurrentCursorName+transform(I)
			enddo
		catch to _oe
			ZLogEx(_oe)
		endtry

		local _response as string

		if (!isnull(_oe))
			&& opakowanie w standardowy ApiResponse
			text to _response noshow textmerge pretext 7
				<?xml version="1.0" encoding="utf-8"?>
				<ApiResponse type="xml">
					 <session><<ksef.sessionReferenceNumber>></session>
					 <exception>
						 <exceptionDetailList>
							 <ExceptionDetail>
							    <exceptionDescription>Błąd pobieranie dokumentu <<transform(ksef.IDSaleDocument)>> z SQL
							    <<_oe.Message>></exceptionDescription>
							    <exceptionCode><<_oe.errorno>></exceptionCode>
							 </ExceptionDetail>
						 </exceptionDetailList>
						 <serviceCtx><<_oe.name>></serviceCtx>
						 <serviceName>OnGetInvoice</serviceName>
	  				</exception>
				</ApiResponse>
			ENDTEXT

			ksef.Content = ksef.loadxml(_response)
			ksef.ThrowIfException("Błąd pobierania dokumentu do eksportu KSeF")
		endif


		text to _response noshow textmerge pretext 15
			<?xml version="1.0" encoding="utf-8"?>
			<ApiResponse type="xml">
				<session><<ksef.sessionReferenceNumber>></session>
				<response><<_value>></response>
			</ApiResponse>
		ENDTEXT
		ksef.Content = ksef.loadxml(_response)
	endproc

	************************************************************
	* deprecate
	* Nie ma potrzeby już podłączania tego zdarzenia, jeśli wcześniej
	* jest wykonany OnSaveResponse, ponieważ procedura SQL podpięta w OnSaveResponse
	* automatycznie powiązanie statusu z wysłaną fakturą i ją odpowiednio
	* oznaczy w bazie
	************************************************************
	procedure OnSaveInvoiceStatus(ksef as "KSeFApi" of "prgs\common\ksefapi.prg") as void &&override

		local ;
			_oe as exception,;
			SqlCommand as string,;
			_CurrentCursorName as string,;
			_value as string,;
			uniqueXmlKsef as string,;
			typ as string,;
			resp as MSXML2,;
			fk as MSXML2,;
			invoiceNumber as string,;
			ksefReferenceNumber as string,;
			acquisitionTimestamp as string,;
			upo as string

		_CurrentCursorName= "invoice"
		_oe = null
		try
			store "" to invoiceNumber ,uniqueXmlKsef,ksefReferenceNumber,acquisitionTimestamp
			resp = ksef.Content.SelectSingleNode("/ApiResponse/response")
			typ = lower(ksef.XmlGetValue(ksef.Content,"/ApiResponse/@type",""))
			if (typ=="invoicestatus")
				fk = resp.SelectSingleNode("./"+typ)
				if (!isnull(fk))
					invoiceNumber = ksef.XmlGetValue(fk,"./invoiceNumber","")
					ksefReferenceNumber = ksef.XmlGetValue(fk,"./ksefReferenceNumber","")
					acquisitionTimestamp = ksef.XmlGetValue(fk,"./acquisitionTimestamp","")
				else
					set step on
					&& problem, ta odpowiedź nie ma takiego elementu w odpowiedź
					&& exception
				endif
				upo = ksef.XmlGetValue(resp,"./upo","")

				text to uniqueXmlKsef noshow textmerge pretext 15
				<KSeF>
					<session><<ksef.sessionReferenceNumber>></session>
  					<invoiceNumber><<invoiceNumber>></invoiceNumber>
  					<ksefReferenceNumber><<ksefReferenceNumber>></ksefReferenceNumber>
  					<acquisitionTimestamp><<acquisitionTimestamp>></acquisitionTimestamp>
  					<upo><<upo>></upo>
				</KSeF>
				ENDTEXT
			endif
			text to SqlCommand noshow textmerge pretext 15
				declare @result int;
				exec @result = dbo.CDNWSHandShakingIU
					@cTable='NaglowkiDokumentowSprzedazy',
					@vDokId=<<transform(ksef.IDSaleDocument)>>,
					@UpdateType='S',
					@direction=2,
					@SysID='KSeF',
					@AdditionalInformation='<<sqltekst(uniqueXmlKsef)>>';
            	select @result [result]
			ENDTEXT
			=SqlCommandExecuteException(_screen.gastro.ConnectionHandle,@SqlCommand,_CurrentCursorName)
			_value = invoice.result
		catch to _oe
			_value=-1
			ZLogEx(_oe)
		finally
			=CrsClose(_CurrentCursorName)
		endtry
	endproc

	procedure OnUPOView(ksef as "KSeFApi" of "prgs\common\ksefapi.prg") as void &&override
		local ;
			_view as form,;
			_page as string,;
			_oe as exception,;
			oThis as session
		_oe = null
		try
			wait clear
			oThis = this
			ksef.SaveBodyResponseToFile(transform(int(seconds()))+".html")
			this.tag = ksef.tag

			_page = fullpath(ksef.ContentFileName())
			if (file(_page))
				_view=newobject("ksefwebview","prgs\common\ksef_webview.prg")
				with _view
					.titleinfo.caption = "Podgląd Urzędowego Poświadczenia Odbioru (UPO)"
					.BrowsePage.oAction = oThis
					.BrowsePage.cSourceURL = "file://"+_page
					.show(1)
				endwith
			endif
		catch to _oe
			ZLogEx(_oe)
			try
				declare integer ShellExecute in shell32;
					integer hwnd,;
					string lpOperation,;
					string lpFile,;
					string lpParameters,;
					string lpDirectory,;
					integer nShowCmd
				ShellExecute(0,"open", _page ,null, "", 3)
			catch to _oe
				_screen.gastro.Showexception(_oe)
			endtry
		finally
			if ("O"==vartype(_view))
				_view.release()
				release _view
				_view=null
			endif
		endtry
	endproc

	procedure OnHtmlView(ksef as "KSeFApi" of "prgs\common\ksefapi.prg") as void &&override
		local ;
			_view as form,;
			_page as string,;
			_oe as exception
		_oe = null
		try
			wait clear
			ksef.SaveBodyResponseToFile(transform(int(seconds()))+".html")
			_page = fullpath(ksef.ContentFileName())
			if (file(_page))
				_view=newobject("ksefwebview","prgs\common\ksef_webview.prg")
				with _view
					.titleinfo.caption = "Podgląd faktury ustrukturyzowanej w oparciu o wzór z Krajowego System e-Faktur"
					.BrowsePage.cSourceURL = "file://"+_page
					.show(1)
				endwith
			endif
		catch to _oe
			ZLogEx(_oe)
			try
				declare integer ShellExecute in shell32;
					integer hwnd,;
					string lpOperation,;
					string lpFile,;
					string lpParameters,;
					string lpDirectory,;
					integer nShowCmd
				ShellExecute(0,"open", _page ,null, "", 3)
			catch to _oe
				_screen.gastro.Showexception(_oe)
			endtry
		finally
			if ("O"==vartype(_view))
				_view.release()
				release _view
				_view=null
			endif
		endtry
	endproc



	************************************************************
	* Procedura zamyka wszystkie zainicjowane przez program sesje,
	* które nie mają znacznika zamknięcia
	* Dla takich sesji będzie można spróbować pobrać UPO
	* generuje zdarzenia:
	*	OnProgressMessage
	* 	OnSaveResponse
	************************************************************
	procedure TerminateUnfinishedSessions(ksef as "KSeFApi" of "prgs\common\ksefapi.prg") as object &&override
		local SqlCommand as string
		local _oe as exception
		local _cursor as string
		_cursor = [exported]
		_oe = null
		try
			raiseevent(ksef,"OnProgressMessage","Zamykam otwarte sesje" )
			text to SqlCommand noshow textmerge pretext 15
				exec [dbo].[CommunicationTable]
					@filter=1,
					@resultFormater=0
			ENDTEXT
			=SqlCommandExecuteException(_screen.gastro.ConnectionHandle,@SqlCommand,_cursor)
			select (_cursor)
			scan
				ksef.SessionTerminateSpecified(&_cursor..session,&_cursor..token)
			endscan
		catch to _oe
			ZLogEx(_oe)
		finally
			=CrsClose(_cursor)
		endtry
		return isnull(_oe)
	endproc


	************************************************************
	* Procedura zbliżona do CheckExported, z tą różnicą
	* że zwraca query z rozkodowanymi informacjami KSeF z additional
	************************************************************
	procedure GetExchangedDocument(_iddocument as integer) as string &&override
		if (empty(_iddocument))
			return []
		endif
		** kontrola czy dokument był eksportowany
		local SqlCommand as string
		local _oe as exception
		local _cursor as string
		_cursor = [exported]
		_oe = null
		try
			text to SqlCommand noshow textmerge pretext 15
				select * from [dbo].[GetExchangedDocument](<<_iddocument>>,default)
			ENDTEXT
			=SqlCommandExecuteException(_screen.gastro.ConnectionHandle,@SqlCommand,_cursor)
		catch to _oe
			ZLogEx(_oe)
		endtry
		return iif(isnull(_oe) and used(_cursor),_cursor,[])
	endproc

	************************************************************
	* Zwraca kursor z informacjami o wyeksportowanym dokumencie
	************************************************************
	procedure CheckExported(_iddocument as integer, _datasession as integer) as string &&override
		if (empty(_iddocument))
			return []
		endif
		** kontrola czy dokument był eksportowany
		local SqlCommand as string
		local _oe as exception
		local _cursor as string
		local _used as logical
		*local _cur_session as Integer
		*_cur_session = this.datasessionid
		_datasession = evl(_datasession,this.save_id_session)
		_cursor = [exported]
		_oe = null
		try
			if (_datasession!=this.save_id_session)
				set datasession to (_datasession)
			endif
			text to SqlCommand noshow textmerge pretext 15
				exec [dbo].[CheckExported] @SeekID=<<_iddocument>>,@version=2
			ENDTEXT
			=SqlCommandExecuteException(_screen.gastro.ConnectionHandle,@SqlCommand,_cursor)
			_used = used(_cursor)
		catch to _oe
			ZLogEx(_oe)
		finally
			if (_datasession!=this.save_id_session)
				set datasession to (this.save_id_session)
			endif
		endtry
		return iif(isnull(_oe) and _used,_cursor,[])
	endproc

	procedure OnRollbackExport(ksef as "KSeFApi" of "prgs\common\ksefapi.prg") as object &&override
		local _oe as exception
		local _crs as exception
		local _docid as int
		_oe = null
		try
			** zdarzenia wywoływane z innej sesji, nie operują na sesji w CheckExported
			** pomimo, że to ten sam obiekt
			_crs = this.CheckExported(ksef.IDSaleDocument,set("datasession"))
			if (empty(_crs))
				exit
			endif
			_docid = alltrim(transform(&_crs..id))
			text to SqlCommand noshow textmerge pretext 15
				exec [dbo].[UnlockExchangeDocuments]
				@documents_ids='<Dokumenty><doc id="<<_docid>>"/></Dokumenty>'
			ENDTEXT
			=SqlCommandExecuteException(_screen.gastro.ConnectionHandle,@SqlCommand,"SqlResult")
		catch to _oe
			ZLogEx(_oe)
			assert .f. message "Błąd podczas wycofywania eksportu"
		finally
			=CrsClose(_crs)
		endtry
	endproc

	************************************************************************
	* Aktualizacja podanej konfiguracji w kontekście frmy
	************************************************************************
	procedure LoadConfiguration(oconf as MSXML2) as object &&override
		local _oe as exception
		_oe = null
		try
			if (this.Priority = this.EnumPriority.POS)
				text to SqlCommand noshow textmerge pretext PRETEXT_VFP
					exec [dbo].[Firmy_RD] @flgType=4,@ID=<<nvl(this.PosID,1)>>
				ENDTEXT
			else
				text to SqlCommand noshow textmerge pretext PRETEXT_VFP
					exec [dbo].[Firmy_RD] @flgType=1,@ID=<<nvl(this.CompanyID,1)>>
				ENDTEXT
			endif
			=SqlCommandExecuteException(_screen.gastro.ConnectionHandle,@SqlCommand,"firma")
			if (reccount([firma])==0)
				ThrowException("Brak konfiguracji KSeF dla firmy nr: "+transform(nvl(this.CompanyID,1)))
			endif
			*!*	 select firma
			*!*	 if (this.CompanyID==0)
			*!*	 	scan for (alltrim(nvl(firma.ksef_token,''))!='')
			*!*	 		exit
			*!*	 	endscan
			*!*	 endif
			*!*	 if (eof([firma]))
			*!*	 	ThrowException("Brak poprawnej konfiguracji KSeF, przypisanej do firm(y)")
			*!*	 endif
						
			with oconf
				.SelectSingleNode("/Configuration/NIP").text = alltrim(nvl(firma.Nip,''))
			   	.SelectSingleNode("/Configuration/APIToken").text = alltrim(nvl(firma.ksef_token,''))
			endwith

			this.MailerConfigurationCrypt = firma.MailConfigurationEncrypted
			this.MailerConfiguration = CryptAPi(strconv(this.MailerConfigurationCrypt,16),_screen.gastro.hash,0)

			this.company = this.GetXmlCursor("firma")
		catch to _oe
			ZLogEx(_oe)
			throw _oe
		finally
			=CrsClose([firma])
		endtry
	endproc


	*****************************************************************************************************
	* Pobranie autentykacji dla wybranego użytkownika
	* Nowa metoda dla KSeF v 2.0
	*****************************************************************************************************
	* W pierwszej kolejności jest wybierany kontekst firmy, czy to na podstawie numeru POS-a (PosId) czy 
	* też na podstawie już wskazanej firmy (CompanyID), ma to na celu ustalenie kontekstu fimy.
	*****************************************************************************************************
	* Dalej jest sprawdzany identyfikator użytkownika, który ma na celu ustalenie kontekstu użytkownika.
	*****************************************************************************************************
	procedure GetAuthentication() as XMLNode
		local _oe as exception
		local result as object
		_oe = null
		try		
			result = dodefault()
			
			if (this.Priority = this.EnumPriority.POS)
				text to SqlCommand noshow textmerge pretext PRETEXT_VFP
					exec [dbo].[Firmy_RD] @flgType=4,@ID=<<nvl(this.PosID,1)>>
				ENDTEXT
			else
				text to SqlCommand noshow textmerge pretext PRETEXT_VFP
					exec [dbo].[Firmy_RD] @flgType=1,@ID=<<nvl(this.CompanyID,1)>>
				ENDTEXT
			endif
			=SqlCommandExecuteException(_screen.gastro.ConnectionHandle,@SqlCommand,"firma")
			if (reccount([firma])==0)
				ThrowException("Brak konfiguracji KSeF dla firmy nr: "+transform(nvl(this.CompanyID,1)))
			endif
					
			result.selectSingleNode("/Authentication/Context/Identifier").text= alltrim(firma.nip)			

			local ctx_user as object			
			ctx_user = result.selectSingleNode("/Authentication/User")			
			** kompatybilność wsteczna, api token zazwyczaj firma będzie już miałą wpisany
			if (!(alltrim(nvl(firma.ksef_token,''))==''))
				ctx_user.selectSingleNode("Type").text= 'apitoken'
				ctx_user.selectSingleNode("Identifier").text= alltrim(firma.ksef_token)
			endif

			this.MailerConfigurationCrypt = firma.MailConfigurationEncrypted
			this.MailerConfiguration = CryptAPi(strconv(this.MailerConfigurationCrypt,16),_screen.gastro.hash,0)

			this.company = this.GetXmlCursor("firma")
		catch to _oe
			ZLogEx(_oe)
			throw _oe
		finally
			=CrsClose([firma])
		endtry
		return result
	endproc

	************************************************************************
	* Wyświetlenie status dokumentu.
	* Przekazujemy obiekt utworzony poprzez CreateInvoiceSettings
	* Uzupełniamy w nim numer faktury i dane kontrahenta, na ich
	* podstawie zostanie stworzona treść do wydruku i wysłania maila
	************************************************************************
	procedure ShowInvoiceStatus(InvoiceSettings as object,KSeF_api as KSeFApi of "prgs\common\ksefapi.prg") &&override
		local ;
			_oe as exception;
			,_message as string;
			,additional as string;
			,isConnect as logical;
			,sessionReferenceNumber as string;
			,_ksefReferenceNumber as string;
			,elementReferenceNumber as string;
			,acquisitionTimestamp as string;
			,invoiceNumber as string;
			,isConnect as logical;
			,exportinfo as string;
			,invoiceGross as decimal;
			,customerEmail as string;
			,customerName as string;
			,customerNip as string;
			,clink as string;
			,hash as string;
			,qrcode as string

		_oe = null
		try
			wait clear
			exportinfo = this.GetExchangedDocument(InvoiceSettings.IDDokumentu)
			if (!(exportinfo==""))
				if (reccount(exportinfo)>0)
					if (at('KSeF',&exportinfo..systems)>0 and empty(evl(&exportinfo..session,'')))
						ThrowException(;
							"Błąd krytyczny, dokument jest oznaczony jako wyeksportowany do systemu KSeF, "+;
							"ale nie ma informacji zwrotnych z systemu. Dokument należy sprawdzić w systemie KSeF, jeśli "+;
							"w nim istniej, należy pobrać";
							)
					endif
				else
					ThrowException("Dokument nie był wysłany do systemu KSeF")
				endif
			endif

			with KSeF_api as KSeFApi of prgs\common\KSeFApi.prg
				.SetIDSaleDocument(InvoiceSettings.IDDokumentu)
				this.upo = alltrim(nvl(&exportinfo..upo,''))
				sessionReferenceNumber = alltrim(nvl(&exportinfo..session,''))
				_ksefReferenceNumber = alltrim(nvl(&exportinfo..ksefReferenceNumber,''))
				elementReferenceNumber = alltrim(nvl(&exportinfo..elementReferenceNumber,''))
				acquisitionTimestamp = alltrim(nvl(&exportinfo..acquisitionTimestamp,''))
				invoiceNumber = alltrim(nvl(&exportinfo..invoiceNumber,''))
				invoiceGross = nvl(&exportinfo..invoiceGross,0.00)
				customerName = alltrim(nvl(&exportinfo..customerName,''))
				customerNip = alltrim(nvl(&exportinfo..customerNip,''))
				hash = alltrim(nvl(&exportinfo..hash,''))
				
				cLink = .GetVerificationLink(_ksefReferenceNumber)

				use in (exportinfo)

				bindevent(KSeF_api,"OnConnect",this,"OnConnect",1)
				bindevent(KSeF_api,"OnConnect",this,"OnThrowIfException")

				bindevent(KSeF_api,"OnCommonStatus",this,"OnSaveResponse")
				bindevent(KSeF_api,"OnCommonStatus",this,"OnThrowIfException",1)
				
				if (.sessionReferenceNumber=="")
					.connect(this)
				endif

				** faktura wysłana, jest odpowiedź, ale nie zarejestrowano numeru referencyjnego (opóźnienie systemu KSeF)
				** jeśli to ta sama sesja, możemy spróbować pobrać, jeśli nie to już nas ratuje tylko UPO
				** spróbujemy pobrać, bo w sumie aktualizując upo, to
				if (_ksefReferenceNumber=="" and .sessionReferenceNumber==sessionReferenceNumber and elementReferenceNumber!='')
					.InvoiceStatusFromReferenceNumber(elementReferenceNumber)
				endif

				if ([]!=elementReferenceNumber) 
					
				endif
				** do zmiany, DLL może pobrać status z lokalnej bazy i go zwrócić bez połaczenia z KSeF
				.CommonSessionStatus(sessionReferenceNumber) && zapisze i rozparsuje na SQL
				
				exportinfo = this.GetExchangedDocument(InvoiceSettings.IDDokumentu)
				if (!(exportinfo==""))
					this.upo = alltrim(nvl(&exportinfo..upo,''))
					_ksefReferenceNumber = alltrim(nvl(&exportinfo..ksefReferenceNumber,''))
					acquisitionTimestamp = alltrim(nvl(&exportinfo..acquisitionTimestamp,''))
					invoiceNumber = alltrim(nvl(&exportinfo..invoiceNumber,''))
					invoiceGross = nvl(&exportinfo..invoiceGross,0.00)
					customerName = alltrim(nvl(&exportinfo..customerName,''))
					customerNip = alltrim(nvl(&exportinfo..customerNip,''))
					customerEmail = alltrim(nvl(&exportinfo..customerEmail,''))
					use in (exportinfo)
				endif
				*endif
				local ocompany as MSXML2
				ocompany=null
				if (!isnull(this.company) and vartype(this.company)=="O")
					ocompany = this.company.SelectSingleNode("/VFPData/firma")
				endif

				local _fileName as string
				_fileName = sys(2023)+"\"+substr(sys(2015),3,10)
				local _mailfile as string
				_mailfile=fullpath(_fileName+".xml")

				local cHead as string

				text to cHead noshow textmerge pretext 15
				<html>
					<head><style>@media print{.noprint{display: none !important;height: 0;}}</style></head>
				<body>
					<font size="4">Faktura wystawiona w <b>K</b>rajowym <b>S</b>ystemie <b><font color="red">e</font>F</b>aktur</font>
					<br/><br/>
					<table>
					<tr><td>Wysłana do KSeF:</td><td><<transform(.GetUtcTimeString(acquisitionTimestamp),"@YL"))>></td></tr>
					<tr><td>W sesji:</td><td><<sessionReferenceNumber>></td></tr>
					<tr><td>Numer KSeF:</td><td><b><font size="5" color="blue"><<_ksefReferenceNumber>></font></b></td></tr>
				ENDTEXT
				
				qrcode = ""				
				if (""!=cLink)
					try
						.GetQRCode(clink ,200,200,"png")
						qrcode = .XmlGetValue(.Content,"/ApiResponse/response","")
					catch to ex1
					endtry
					
					if (!(alltrim(evl(qrcode,""))==""))
						text to cHead noshow textmerge pretext 15 ADDITIVE
						<tr><td></td><td><div><img src="data:image/png;base64,<<qrcode>>"alt="<<_ksefReferenceNumber>>" /></div></td></tr>
						ENDTEXT
						**<tr><td></td><td><<ksefReferenceNumber>></td>
					endif
				endif
				text to cHead noshow textmerge pretext 15 additive
					<tr><td>Numer faktury:</td><td><b><<invoiceNumber>></b></td></tr>
				ENDTEXT
				if (!isnull(customerName))
					text to cHead noshow textmerge pretext 15 additive
						<tr><td>NIP:</td><td><<iif(customerNip=="" or left(customerNip,4)=="ZAGR","<i>Brak identyfikatora</i>",customerNip)>></td></tr>
						<tr><td>Pełna nazwa nabywcy:</td><td><<customerName>></td></tr>
						<tr><td>Należność ogółem:</td><td><<invoiceGross>></td></tr>
					ENDTEXT
				endif

				text to _mail noshow textmerge pretext 15
							<<cHead>>
				ENDTEXT
				if (this.upo!="")
					text to _mail noshow textmerge pretext 15 additive
							<tr><td colspan="2"/></tr>
							<tr><td>Status UPO:</td><td><i><font color="green">Dokument posiada Urzędowe Poświadczenie Odbioru (UPO)</font></i></td></tr>
					ENDTEXT
				endif

				text to _message noshow textmerge pretext 3
						<<cHead>>
				ENDTEXT

				if (this.upo=="")
					text to _message noshow textmerge pretext 15 additive
					 <tr class="noprint"><td>Status UPO:</td><td>Dokument nie posiada jeszcze Urzędowego Poświadczenie Odbioru (UPO)</td></tr>
					ENDTEXT
				else
					text to _message noshow textmerge pretext 15 additive
					<tr><td colspan="2"></td></tr>
					<tr><td>Status UPO:</td><td><a href="vfps://runaction?openpage&_url=show_upo"><i><font color="green">Dokument posiada Urzędowe Poświadczenie Odbioru (UPO)</font></i><a></td></tr>
					ENDTEXT
				endif

				text to _message noshow textmerge pretext 15 additive
					</table>
					<br/><br/>
					<button onclick="window.print();" class="noprint">
						Wydrukuj
					</button>
				ENDTEXT


				if (!empty(customerEmail) and vartype(this.MailerConfiguration)=="C" and !empty(this.MailerConfiguration))
					text to _message noshow textmerge pretext 15 additive
							<p class="noprint">
								<br/><br/>
								<a href="vfps://runaction?openpage&send_mail=SoftechSendMail&email=<<customerEmail>>&mail_file=<<_mailfile>>">
									Wyślij to potwierdzenie mailem do klienta (<<customerEmail>>)
								<a>
							</p>
					ENDTEXT
				else
					text to _message noshow textmerge pretext 15 additive
							<p class="noprint">
								<br/><br/>
								<font color="blue"><i>Wysłanie potwierdzenia mailem nie jest możliwe, nie uzupełniono adresu eMail przy nabywcy</i></font>
							</p>
					ENDTEXT
				endif

				_cliptext=(_message)

				if (!isnull(ocompany) and vartype(ocompany)=="O")
					webUrl = strtran(KSeF_api.cUrl,"/api","/web")

					text to _mail noshow textmerge pretext 3 additive
							</table>
							<table>
							<tr><td>Sprzedawca</td><td/></tr>
							<tr><td/><td><<alltrim(oCompany.SelectSingleNode("./nazwa1").text)>></td></tr>
							<tr><td/><td><<alltrim(oCompany.SelectSingleNode("./nazwa2").text)>></td></tr>
							<tr><td/><td><<alltrim(oCompany.SelectSingleNode("./nazwa3").text)>></td></tr>
							<tr><td/><td><<alltrim(oCompany.SelectSingleNode("./adres1").text)>></td></tr>
							<tr><td/><td><<alltrim(oCompany.SelectSingleNode("./adres2").text)>></td></tr>
							<tr><td/><td><<alltrim(oCompany.SelectSingleNode("./nip").text)>></td></tr>
							</table>
							<br/>
							<a href="<<cLink>>">Zobacz fakturę wystawioną w Krajowym Systemie e-Faktur bez konieczności logowania</a><br>
							<a href="<<webUrl>>/anonymous-access">Jedną z form pobrania faktury, jest dostęp anonimowy w KSeF</a>
							<br/>
							<br/>--
							<br/>Z poważaniem
							<br/><<alltrim(oCompany.SelectSingleNode("./nazwa1").text)>>
						</body>
						</html>
					ENDTEXT
				endif

				if (vartype(this.MailerConfiguration)=="C" and !empty(this.MailerConfiguration))
					local oxml as MSXML2,omail as MSXML2, onode as MSXML2
					oxml = GetXMLDom(this.MailerConfiguration,.t.)
					omail = oxml.SelectSingleNode("/Configuration/Email")
					local xmlmail as string
					if (!empty(customerEmail))
						xmlmail = [<?xml version="1.0" encoding="utf-8" ?><Mail><Send>]
						onode = omail.SelectSingleNode("./From")
						if (!isnull(onode))
							xmlmail = xmlmail + [<From>]+onode.text+[</From>]
						endif
						onode = omail.SelectSingleNode("./BCC")
						if (!isnull(onode))
							xmlmail = xmlmail + [<Bcc>]+onode.text+[</Bcc>]
						endif
						onode = omail.SelectSingleNode("./CC")
						if (!isnull(onode))
							xmlmail = xmlmail + [<CC>]+onode.text+[</CC>]
						endif
						xmlmail = xmlmail + [<To>]+customerEmail+[</To>]
						xmlmail = xmlmail + [</Send>]
						xmlmail = xmlmail + [<Subject>Faktura: ]+invoiceNumber+[, Numer faktury KSEF: ]+_ksefReferenceNumber+[</Subject>]
						xmlmail = xmlmail + [<Body type="html">]
						xmlmail = xmlmail + "<![CDATA["
						xmlmail = xmlmail + strconv(_mail,9)
						xmlmail = xmlmail + "]]>"
						xmlmail = xmlmail + [</Body></Mail>]
						strtofile(xmlmail,_mailfile)
					endif
				endif
				_page=fullpath(_fileName+".html")
				strtofile(_message,_page)
			endwith

			raiseevent(KSeF_api,"OnProgressMessage","")

			local inf as form

			inf = newobject("_webform","libs\_webview.vcx","","file://"+_page)
			inf.icon="graphics\icon_222.ico"
			inf.caption="Status (KSeF)"
			inf.width=900
			inf.height=550
			inf.resize()
			inf.autocenter=.t.
			inf.oleWebBrowser.oAction = this
			inf.show(1)

		catch to _oe
			ZLogEx(_oe)
			_screen.gastro.ShowMessage(_oe.message)
		finally
			if (!empty(exportinfo))
				=CrsClose(exportinfo)
			endif
			if (isConnect)
				_oe = null
				try
					unbindevents(KSeF_api,"OnCommonStatus",this,"OnSaveResponse")
					unbindevents(KSeF_api,"OnCommonStatus",this,"OnThrowIfException")
					unbindevents(KSeF_api,"OnConnect",this,"OnConnect")
					unbindevents(KSeF_api,"OnConnect",this,"OnThrowIfException")
				catch to _oe
					ZLogEx(_oe)
				endtry
			endif
		endtry
	endproc


	****************************************************
	* Procedur bezpołączeniowa z KSeF
	* Działa na bazie danych i generuje kolekcję plików
	* HTML z audytami danych
	****************************************************
	procedure OnValidDatabase()
		local _oe as exception
		local _html_file as string
		_oe = null
		try
			text to SqlCommand noshow textmerge pretext 15
				exec [dbo].[ValidationDataKSEF] @module='firma'
				exec [dbo].[ValidationDataKSEF] @module='kontrahenci'
			ENDTEXT
			=SqlCommandExecuteException(_screen.gastro.ConnectionHandle,@SqlCommand,"valid_db")

			_html_file = "temp\ReportCompany.html"
			if (this.ReportCompany("valid_db",_html_file)>0)
				this.ShowValidReport("Raport firmy",_html_file)
			endif
			_html_file = "temp\ReportContractor.html"
			if (this.ReportContractor("valid_db1",_html_file)>0)
				this.ShowValidReport("Raport kontrahentów",_html_file)
			endif


		catch to _oe
			ZLogEx(_oe)
		finally
			=CrsClose([valid_db])
			=CrsClose([valid_db1])
			=CrsClose([valid_db2])
			=CrsClose([valid_db3])
		endtry
	endproc

	procedure ReportCompany(_tb as string, _html_file as string)
		local ReportFile as CFile of "prgs\common\cfile.prg"
		local ncount as integer
		local _html as string

		ncount = 0
		if (reccount(_tb)==0)
			return ncount
		endif

		ReportFile = newobject("CFile","common\CFile.prg","",_html_file)
		text to _html noshow textmerge pretext 15
				<html>
					<title>Report exchanged</title>
				<head>
				<style>
					th{color: white;background: green;}
					td{vertical-align:top;}

					@media print{.noprint{display: none !important;height: 0;}}
				</style></head>
				<body>
				<button onclick="window.print();" class="noprint"> Wydrukuj </button>
				<br/>
				<font size="5" color="red">Raport danych firmy do uzupełnienia</font>
				<table width="100%" border="1">
	            <thead>
	              <tr>
	                  <th width="30%" align="left">Nazwa firmy</th>
	                  <th width="10%" align="center">Bank</th>
	                  <th width="15%" align="center">Numer rachunku</th>
	                  <th width="15%" align="center">Konfiguracja KSeF</th>
	                  <th width="10%" align="center">Konfiguracja eMail</th>
	                  <th width="10%" align="center">Miejsce faktur</th>
	                  <th width="10%" align="center">NIP</th>
	              </tr>
	            </thead>
		ENDTEXT
		with ReportFile
			.create()
			.write(_html)
		endwith

		select (_tb)
		scan
			text to _html noshow textmerge pretext 7
			<tr>
				<td align="left"><<Firma>></td>
				<td align="center"><font color="<<iif(Bank=[OK],[green],[red])>>")><<Bank>></font></td>
				<td align="center"><font color="<<iif(NumerRachunku=[OK],[green],[red])>>")><<NumerRachunku>></font></td>
				<td align="center"><font color="<<iif(KonfiguracjaKSeF=[OK],[green],[red])>>")><<KonfiguracjaKSeF>></font></td>
				<td align="center"><font color="<<iif(KonfiguracjaEMail=[OK],[green],[red])>>")><<KonfiguracjaEMail>></font></td>
				<td align="center"><font color="<<iif(MiejsceFaktur=[OK],[green],[red])>>")><<MiejsceFaktur>></font></td>
				<td align="center"><font color="<<iif(NIP=[OK],[green],[red])>>")><<NIP>></font></td>
			</tr>
			ENDTEXT
			ReportFile.write(_html)
			ncount = ncount + 1
		endscan

		with ReportFile
			.write("</table></body></html>")
			.close()
		endwith
		return ncount
	endproc

	procedure ReportContractor(_tb as string, _html_file as string)
		local ReportFile as CFile of "prgs\common\cfile.prg"
		local ncount as integer
		local _html as string

		ncount = 0
		if (reccount(_tb)==0)
			return ncount
		endif

		ReportFile = newobject("CFile","common\CFile.prg","",_html_file)
		text to _html noshow textmerge pretext 15
				<html>
					<title>Report exchanged</title>
				<head>
				<style>
					th{color: white;background: green;}
					td{vertical-align:top;}

					@media print{.noprint{display: none !important;height: 0;}}
				</style></head>
				<body>
				<button onclick="window.print();" class="noprint"> Wydrukuj </button>
				<br/>
				<font size="5" color="red">Raport danych kontrahentów do uzupełnienia</font>
				<table width="100%" border="1">
	            <thead>
	              <tr>
	                  <th width="10%" align="left">IK klienta</th>
	                  <th width="30%" align="left">Nazwa</th>
	                  <th width="5%" align="center">NIP</th>
	                  <th width="15%" align="center">Kod pocztowy</th>
	                  <th width="10%" align="center">Poczta</th>
	                  <th width="15%" align="center">Ulica</th>
	                  <th width="10%" align="center">EMail</th>
	              </tr>
	            </thead>
		ENDTEXT

		with ReportFile
			.create()
			.write(_html)
		endwith

		select (_tb)
		scan
			text to _html noshow textmerge pretext 7
			<tr>
				<td align="left"><<IDKontrahenta>></td>
				<td align="left"><<Klient>></font></td>
				<td align="center"><font color="<<iif(NIP=[OK],[green],[red])>>")><<NIP>></font></td>
				<td align="center"><font color="<<iif(KodPocztowy=[OK],[green],[red])>>")><<KodPocztowy>></font></td>
				<td align="center"><font color="<<iif(Poczta=[OK],[green],[red])>>")><<Poczta>></font></td>
				<td align="center"><font color="<<iif(Ulica=[OK],[green],[red])>>")><<Ulica>></font></td>
				<td align="center"><font color="<<iif(email=[OK],[green],[blue])>>")><<email>></font></td>
			</tr>
			ENDTEXT
			ReportFile.write(_html)
			ncount = ncount + 1
		endscan

		with ReportFile
			.write("</table></body></html>")
			.close()
		endwith
		return ncount
	endproc

	procedure ShowValidReport(title_document as string, file_html as string)
		local i_help_page as string
		file_html = fullpath(file_html)
		if (file(file_html ))
			local inf as form
			inf = newobject("_webform","libs\_webview.vcx","","file://"+file_html )
			inf.icon = "graphics\icon_222.ico"
			inf.caption=title_document
			inf.width = _screen.width*0.80 &&750
			inf.height = _screen.height*0.80 &&600
			inf.resize()
			inf.autocenter=.t.
			inf.show(1)
		endif
	endproc

enddefine

define class KSeFDBF as KSeFSQL of "prgs\common\ksefapi.prg"
	#define override
	_memberdata = ;
		'<VFPData>'+;
		'</VFPData>'

	************************************************************
	* Procedura zbliżona do CheckExported, z tą różnicą
	* że zwracaquery z rozkodowanymi informacjami KSeF z additional
	************************************************************
	procedure GetExchangedDocument(_iddocument as integer) as string &&override
		if (_iddocument<0)
			=ThrowException( "Dokument został wyeksportowany do Szefa i czeka na synchronizację" )
		endif
		=ThrowException("Faktura została wyeksportowana i pobrana do szefa"+chr(13)+;
			"Operację usunięcia można przeprowadzić jedynie w szefie")
		return []
	endproc

	************************************************************
	* Zwraca kursor z informacjiami o wyeksportowanym dokumencie
	************************************************************
	procedure CheckExported(_iddocument as integer, _datasession as integer) as string &&override
		if (_iddocument<0)
			=ThrowException( "Dokument został wyeksportowany do Szefa i czeka na synchronizację" )
		endif
		=ThrowException("Faktura została wyeksportowana i pobrana do szefa"+chr(13)+;
			"Operację usunięcia można przeprowadzić jedynie w szefie")
		return []
	endproc

	*****************************************************************************************************
	* Pobranie autentykacji dla wybranego użytkownika
	* Nowa metoda dla KSeF v 2.0
	*****************************************************************************************************
	* W pierwszej kolejności jest wybierany kontekst firmy, czy to na podstawie numeru POS-a (PosId) czy 
	* też na podstawie już wskazanej firmy (CompanyID), ma to na celu ustalenie kontektu fimy.
	*****************************************************************************************************
	* Dalej jest sprawdzany identyfikator użytkownika, który ma na celu ustalenie konktektu użytkownika.
	*****************************************************************************************************
	procedure GetAuthentication() as XMLNode
		local _oe as exception
		local result as object
		local _nip as string, _token as string
		assert .f. message "Autentykacja"
		_oe = null
		try		
			result = dodefault()
			
			if  (this.PosID==m.numerkasy)
				_nip = 	m.nipwlasny
				_token = m.x_ksef_token_firma1
			else
				if  (this.PosID==m.kasadrugiejfirmy)
					_nip = 	m.nip2
					_token = m.x_ksef_token_firma2
				else
					store "" to _nip,_token
				endif
			endif
			*!*	 if (_token!="")
			*!*	 	_token = CryptAPi(strconv(_token,16),_screen.gastro.hash,0)
			*!*	 endif
					
			result.selectSingleNode("/Authentication/Context/Identifier").text= alltrim(firma.nip)			

			local ctx_user as object			
			ctx_user = result.selectSingleNode("/Authentication/User")			
			** kompatybilność wsteczna, api token zazwyczaj firma będzie już miałą wpisany
			if (_token!="")
				ctx_user.selectSingleNode("Type").text= 'apitoken'
				ctx_user.selectSingleNode("Identifier").text= alltrim(_token)
			endif

			*!*	 this.MailerConfigurationCrypt = firma.MailConfigurationEncrypted
			*!*	 this.MailerConfiguration = CryptAPi(strconv(this.MailerConfigurationCrypt,16),_screen.gastro.hash,0)

		catch to _oe
			ZLogEx(_oe)
			throw _oe
		endtry
		return result
	endproc

	************************************************************
	* Aktualizacja podanej konfiguracji w kontekście frmy
	************************************************************
	procedure LoadConfiguration(oconf as MSXML2) as object &&override
		local _oe as exception
		local _nip as string, _token as string

		_oe = null
		try
			if  (this.PosID==m.numerkasy)
				_nip = 	m.nipwlasny
				_token = m.x_ksef_token_firma1
			else
				if  (this.PosID==m.kasadrugiejfirmy)
					_nip = 	m.nip2
					_token = m.x_ksef_token_firma2
				else
					store "" to _nip,_token
				endif
			endif
			if (_token!="")
				_token = CryptAPi(strconv(_token,16),_screen.gastro.hash,0)
			endif

			with oconf
				.SelectSingleNode("/Configuration/NIP").text = alltrim(_nip)
				.SelectSingleNode("/Configuration/APIToken").text = alltrim(_token)				
			endwith
		catch to _oe
			ZLogEx(_oe)
			throw _oe
		finally
			=CrsClose([firma])
		endtry
	endproc

enddefine

define class KSeFApi as separator
	#define virtual
	#define DEFAULT_BUFFER_SIZE 16384

	#define API_PRODUCTION 1
	#define API_TEST 2
	#define API_DEMO 3

	#define true .t.
	#define false .f.

	hidden isload as logical
	hidden file_name as string

	RootDirectory = ""
	hidden DLLFile as string
	Content = null && as string
	oconf = null &&as MSXML2

	#if .f.
		lpEnumFunc =0
	#endif

	** kontekst POS-a/Firmy w obębie której pracujemy
	Context = ""

	CompanyID=0
	IDSaleDocument=0
	ApiVersion = 1
	cUrl=""

	** aktualne środowisko pracy
	APIEnvironment = API_TEST

	** dostępne środowiska pracy ustawiane w kontruktorze
	APIEnvironmentEnum = null

	** aktualny numer sesji
	sessionReferenceNumber = ""
	MailerConfiguration = null

	** Obiekt autentykacji przekazywany przy logowaniu do KSeF
	Authentication = null

	_memberdata = ;
		'<VFPData>'+;
		'<memberdata name="file_name_assign" type="method" display="file_name_assign"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="contentfilename" type="method" display="ContentFileName"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="localcontentresponse" type="method" display="LocalContentResponse"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="loadxml" type="method" display="LoadXml"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="init" type="method" display="Init"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="loadconfiguration" type="method" display="LoadConfiguration"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="exportcontenttoinvoices" type="method" display="ExportContentToInvoices"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="exportapicontenttofile" type="method" display="ExportAPIContentToFile"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="apicontenttofile" type="method" display="APIContentToFile"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="savecontenttofile" type="method" display="SaveContentToFile"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="GetLastResponseContent" type="method" display="GetLastResponseContent"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="loadmessageexception" type="method" display="LoadMessageException"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="getconfiguration" type="method" display="GetConfiguration"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="connect" type="method" display="Connect"   tooltiptext="x1" message="x1"/>'+;
		'<memberdata name="sessionterminate" type="method" display="SessionTerminate"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="sessionterminateall" type="method" display="SessionTerminateAll"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="sessionterminatespecified" type="method" display="SessionTerminateSpecified"  help="Zamknięcie sesji wskazanej numerem i tokenem"/>'+;
		'<memberdata name="sessionstatus" type="method" display="SessionStatus"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="commonsessionstatus" type="method" display="CommonSessionStatus"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="invoicesdownloadrangedate" type="method" display="InvoicesDownloadRangeDate"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="sendinvoicexml" type="method" display="SendInvoiceXML"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="invoicestatusfromreferencenumber" type="method" display="InvoiceStatusFromReferenceNumber"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="invoicestatusfromresponse" type="method" display="InvoiceStatusFromResponse"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="savebodyresponsetofile" type="method" display="SaveBodyResponseToFile"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="destroy" type="method" display="Destroy"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="error" type="method" display="Error"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="declare" type="method" display="Declare"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="sqlgetconfiguration" type="method" display="SqlGetConfiguration"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="content" type="member" display="Content"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="dllfile" type="member" display="DllFile"  help="Pobranie kolejnego kursora"/>'+;
		'<memberdata name="file_name" type="member" display="file_name"  helptext="x1"/>'+;
		'<memberdata name="idsaledocument" type="member" display="IDSaleDocument"  helptext="x1"/>'+;
		'<memberdata name="setidsaledocument" type="method" display="SetIDSaleDocument" />'+;
		'<memberdata name="elementresponse" type="method" display="ElementResponse" />'+;
		'<memberdata name="xmlgetnode" type="method" display="XmlGetNode" />'+;
		'<memberdata name="xmlgetvalue" type="method" display="XmlGetValue" />'+;
		'<memberdata name="xmlsetvalue" type="method" display="XmlSetValue" />'+;
		'<memberdata name="xmlupdatevalue" type="method" display="XmlUpdateValue" />'+;
		[<memberdata display="SessionReferenceNumber" name="sessionreferencenumber" type="method" />]+;
		[<memberdata display="ThrowIfException" name="throwifexception" type="method" />]+;
		[<memberdata display="UPOToHTML" name="upotohtml" type="method" />]+;
		[<memberdata display="ReLoadConfiguration" name="reloadconfiguration" type="method" />]+;
		;
		[<memberdata display="OnCommonStatus" name="oncommonstatus" type="method" />]+;
		[<memberdata display="OnConnect" name="onconnect" type="method" />]+;
		[<memberdata display="OnHtmlView" name="onhtmlview" type="method" />]+;
		[<memberdata display="OnInvoiceQuery" name="oninvoicequery" type="method" />]+;
		[<memberdata display="OnlineInvoiceGet" name="onlineinvoiceget" type="method" />]+;
		[<memberdata display="OnProgressMessage" name="onprogressmessage" type="method" />]+;
		[<memberdata display="OnSaveResponse" name="onsaveresponse" type="method"  />]+;
		[<memberdata display="OnStatus" name="onstatus" type="method" />]+;
		[<memberdata display="OnUPOView" name="onupoview" type="method" />]+;
		[<memberdata display="OnValidateInvoice" name="onvalidateinvoice" type="method" />]+;
		[<memberdata display="OnGetInvoice" name="ongetinvoice" type="method" />]+;
		[<memberdata display="OnSaveInvoiceStatus" name="onsaveinvoicestatus" type="method" />]+;
		[<memberdata display="OnSendInvoice" name="onsendinvoice" type="method" />]+;
		[<memberdata display="OnOnlineInvoiceGet" name="ononlineinvoiceget" type="method" />]+;
		[<memberdata display="OnSessionStatus" name="onsessionstatus" type="method" />]+;
		[<memberdata display="OnRollbackExport" name="onrollbackexport" type="method" helpfile="asdfsd"/>]+;
		;
		[<memberdata display="ApiVersion" name="apiversion" type="property" />]+;
		[<memberdata display="APIEnvironment" name="apienvironment" type="property" />]+;
		[<memberdata display="APIEnvironmentEnum" name="apienvironmentenum" type="property" />]+;
		[<memberdata display="PosID" name="posid" type="property"/>]+;
		'</VFPData>'

	procedure OnSendInvoice(sender as KSeF_api of "prgs\common\ksefapi.prg") && zdarzenie wywoływane po wysyłce faktury
	procedure OnOnlineInvoiceGet(sender as KSeF_api of "prgs\common\ksefapi.prg") && Zdarzenie wywoływane po pobraniu fakury
	procedure OnConnect(sender as KSeF_api of "prgs\common\ksefapi.prg") &&zdarzenie wywoływane po metodzie połączeniowej
	procedure OnSaveResponse(sender as KSeF_api of "prgs\common\ksefapi.prg") &&Zdarzenie ogólne zapisujące wowolny response do bazy
	procedure OnInvoiceQuery(sender as KSeF_api of "prgs\common\ksefapi.prg")
	procedure OnSessionStatus(sender as KSeF_api of "prgs\common\ksefapi.prg")
	procedure OnCommonStatus(sender as KSeF_api of "prgs\common\ksefapi.prg")
	procedure OnValidateInvoice(sender as KSeF_apif "prgs\common\ksefapi.prg")
	procedure OnGetInvoice(sender as KSeFApi of "prgs\common\ksefapi.prg") &&zdarzenie ma za zadanie odebranie dokumentu w postaci XML z systemu bazodanowego, określonego poprzez sender.IDSaleDocument
	procedure OnSaveInvoiceStatus(sender as KSeFApi of "prgs\common\ksefapi.prg") &&Zdarzenie obsługujące odpowiedź na pytanie o status wysłanej faktury, Numer referencyjny z wysyłki faktury
	procedure OnHtmlView(sender as KSeF_apif "prgs\common\ksefapi.prg")
	procedure OnUPOView(sender as KSeFApi of "prgs\common\ksefapi.prg") &&podgląd i eksport UPO
	procedure OnProgressMessage(_message as string) &&zdarzenie informacyjne

		**rollback wysyłki, na SQL zapisano już do dokumentów wyeksportowanych
	procedure OnRollbackExport(sender as KSeFApi of "prgs\common\ksefapi.prg")

		** kontrola połączenia
		** todo: Można spróbować podpiąć pod każdą metodę komuniacyjną, jeśi sesja upłynęłą
		** procedura zdarzeniowa powinna nawiązać połączenie i przejść dalej do realizacji procedury
	procedure OnCheckConnection(sender as KSeF_apif "prgs\common\ksefapi.prg")

		*************************************************************************************
		* Metoda konstruktora
		*************************************************************************************
	procedure init()
		with this as "KSeFAPi" of "ksefapi.prg"

			.APIEnvironmentEnum = newobject("empty")
			addproperty(.APIEnvironmentEnum,"production",API_PRODUCTION)
			addproperty(.APIEnvironmentEnum,"test",API_TEST)
			addproperty(.APIEnvironmentEnum,"demo",API_DEMO)

			.DLLFile = "ksefconnector.dll"
			try
				.declare()
				#if .f.
					set library to vfp2c32.fll additive
					#define VFP2C_INIT_CALLBACK	0x00000100
					INITVFP2C32(VFP2C_INIT_CALLBACK)
					this.lpEnumFunc = CreateCallbackFunc('ShowResponse','VOID','STRING',this)
					**KSeF_RegisterCallbackResponse(this.lpEnumFunc)
				#endif
				try
					local cResult as string
					cResult=""
					ssml_Initialize()
					_screen.gastro.crypt.ExportSessionKey(_screen.gastro.hash,@cResult)
					ssml_SetSessionKey(strconv(cResult,13))
				catch to ex
					ZLogEx(ex)
					_screen.gastro.ShowMessage("Błąd ładowania modułu do wysyłania email"+chr(13)+"Skontaktuj się z pomocą techniczną",30,255)
				endtry
			catch to _oe
				#if .f.
					if (this.lpEnumFunc>0)
						DestroyCallbackFunc(this.lpEnumFunc)
					endif
				#endif
				ZLogEx(_oe)
				.isload = .f.
				_screen.gastro.ShowMessage("Błąd ładowania modułu KSeF"+chr(13)+"Skontaktuj się z pomocą techniczną",30,255)
			endtry
			return .isload
		endwith
	endproc

	*************************************************************************************
	* Przeładowanie konfiguracji.
	* Jeśli jest taka potrzeba. Do biblioteki zostaje wysłana nowa konfiguracja
	* Jeśli biblioteka wykryje, że zmieniono konfigurację (ApiToken) to sprónuje
	* załadować token
	*************************************************************************************
	procedure ReLoadConfiguration(odb as [KSeFDB] of "prgs\common\ksefapi.prg") as void
		local isReload as logical
		local clientContext as string
		with this
			try
				clientContext = odb.Context()
				isReload = (clientContext!=.Context or isnull(odb.company))
				if (isReload)
					odb.LoadConfiguration(.oconf)
					.Context = clientContext
					.isload = (KSeFErrorCode_OK == KSeF_SetConfiguration(.oconf.xml))
				endif

				if (!.isload)
					=ThrowException("Błąd inicjacji biblioteki KSeF", 100, "")
				endif
			catch to _oe
				ZLogEx(_oe)
			endtry
		endwith
		return isReload
	endproc

	procedure LoadConfiguration(config as string) as void
		local ;
			Content as string

		with this
			try
				if (substr(config,1,1)=="<")
					Content = config
				else
					Content = filetostr(config)
				endif
				.oconf = .loadxml(Content)
				.RootDirectory = addbs(alltrim(.oconf.SelectSingleNode("./Configuration/DirectoryWorking/@response").value))
				.ApiVersion = val(.oconf.SelectSingleNode("./Configuration/APIVersion").text)
				.cUrl = lower(.oconf.SelectSingleNode("./Configuration/URL").text)
				do case
					case (at("ksef-test",.cUrl)>0)
						.APIEnvironment = .APIEnvironmentEnum.test
					case (at("ksef-demo",.cUrl)>0)
						.APIEnvironment = .APIEnvironmentEnum.demo
					otherwise
						.APIEnvironment = .APIEnvironmentEnum.production
				endcase
				.isload = (KSeFErrorCode_OK == KSeF_SetConfiguration(.oconf.xml))
				if (!.isload)
					nerror = GetLastError()
					debugout "GetLastError() ",GetLastError()
					=ThrowException("Błąd inicjacji biblioteki KSeF", 100, "")
				endif
				KSeF_SetPasswordCommunication(strconv(_screen.gastro.hash,13))
				**KSeF_QueueSendInvoiceTaskStart()
			catch to _oe
				ZLogEx(_oe)
				_message = _oe.message

				_oe = null
				try
					.GetLastResponseContent()

					if (!isnull(.Content.SelectSingleNode("/ApiResponse/exception")))
						_message = iif(isnull(_message) or empty(_message),"",_message+chr(13)+.LoadMessageException())
					endif
					*_screen.gastro.ShowMessage(	_message ,20,255)
				catch to _oe
					ZLogEx(_oe)
				finally
				endtry
				*"Błąd ładowania konfiguracji"+chr(13)+
				*"Sprawdź konfigurację oraz token przypisany do firmy"+chr(13)+
				_screen.gastro.ShowMessage(_message+chr(13)+chr(13)+;
					"Do momentu poprawienia konfiguracji i ponownego uruchomienia programu"+chr(13)+;
					"wysyłanie faktur do systemu KSeF nie będzie możliwe!",20,255)
			endtry
		endwith
	endproc

	*!*	 procedure SetToken(token as string) as void
	*!*	 	if ('O' == vartype(this.oconf))
	*!*	 		this.oconf.SelectSingleNode("/Configuration/APIToken").text = token
	*!*	 		this.LoadConfiguration(this.oconf.xml)
	*!*	 		return
	*!*	 	endif
	*!*	 	=ThrowException("Błąd przypisania token-a, możliwe, że nie załadowano konfiguracji KSeF")
	*!*	 endproc


	procedure SetUser(xml_user_authentication as object) as XMLNode	
		assert .f. message "Autentykacja SetUser"
		local Authentication as string		
		text to Authentication noshow textmerge pretext 15
				<Authentication>
					<Context><Type>onip</Type><Identifier/></Context>
					<<xml_user_authentication>>
				</Authentication>
		ENDTEXT			
		this.Authentication = GetXmlObject()		
		with this.Authentication as MSXML2
			.preservewhitespace = .f.
			.loadxml(Authentication )
			with (.ParseError)
				if (.ErrorCode!=0)
					set step on
					debugout .ErrorCode,.srcText,.reason
					=ThrowException("Błąd dekodowania danych autentykacji: "+chr(13)+.srcText,KSeFErrorCode_Unknown)
				endif
			endwith
		endwith
	endproc
	
	
	procedure SetIDSaleDocument(_value as int)
		this.IDSaleDocument = _value
	endproc

	procedure file_name_assign(_value as string)
		this.file_name = this.RootDirectory+_value
	endproc

	procedure ContentFileName()
		if (!empty(this.file_name))
			return this.file_name
		endif
		return ""
	endproc

	procedure LocalContentResponse()
		if (!isnull(this.Content))
			return this.Content
		endif
		return null
	endproc

	procedure XmlUpdateValue(xmlobject as MSXML2, xpath as string,_value as string)
		local el as object
		if (!isnull(xmlobject))
			el = xmlobject.SelectSingleNode(xpath)
			if (!isnull(el))
				el.text = transform(_value)
				return true
			endif
		endif
		return false
	endfunc

	procedure XmlCheckElement(xmlobject as MSXML2, xpath as string)
		local el as object
		if (!isnull(xmlobject))
			el = xmlobject.SelectSingleNode(xpath)
			if (!isnull(el))
				return .t.
			endif
		endif
		return .f.
	endfunc
	
	procedure XmlGetValue(xmlobject as MSXML2, xpath as string,_default as string)
		local el as object
		if (!isnull(xmlobject))
			el = xmlobject.SelectSingleNode(xpath)
			if (!isnull(el))
				return evl(el.text,_default)
			endif
		endif
		return _default
	endfunc

	procedure XmlGetOrCreateNode(xmlobject as MSXML2, xpath as string, _default as string) as object
		local el as object
		if (!isnull(xmlobject))
			el = xmlobject.SelectSingleNode(xpath)
			if (isnull(el))
				el = xmlobject.ownerdocument.createNode(1,xpath,"")
				el.text = _default
				xmlobject.appendChild(el)
			else 
				if (empty(el.text))
					el.text = _default
				endif
			endif
			return el
		endif
		return null
	endfunc
	
	procedure XmlGetNode(xmlobject as MSXML2, xpath as string) as object
		local el as object
		if (!isnull(xmlobject))
			el = xmlobject.SelectSingleNode(xpath)
			if (!isnull(el))
				return el
			endif
		endif
		return null
	endfunc

	procedure ElementResponse(xpath as string, xRoot as string) as string
		local _oe as exception
		local _value as string, XmlNode as MSXML2,respType as string

		store null to _value,_oe
		try
			respType = this.Content.SelectSingleNode("/ApiResponse/@type")
			if (vartype(xRoot)!="C")
				xRoot = "/ApiResponse/response"
			endif
			if (vartype(xpath)!="C")
				xpath= ""
			endif

			XmlNode = this.Content.SelectSingleNode(xRoot+xpath)
			if (!isnull(XmlNode))
				if (!inlist(respType.text,"String","InvoiceBody") and [O]==vartype(XmlNode.FirstChild))
					_value = XmlNode.FirstChild.xml
				else
					_value = XmlNode.text
				endif
			else
				debugout("Error XmlNode: "+xRoot+xpath)
			endif
		catch to _oe
			ZLogEx(_oe)
			_value = null
		endtry

		return _value
	endproc


	*********************************************************************
	* Załadowanie przekazanego XML do OXMDOM
	* Jeśli nie przekazujemy XML, wóczas metoda ładuje bieżący content
	* Metoda rzuca wyjątki
	*********************************************************************
	procedure loadxml(Content as string) as MSXML2
		if (empty(Content) or isnull(Content) or Content=="")
			ThrowException("Błąd deserializacji do XML - brak treści XML")
		endif
		local oxml as MSXML2
		oxml = GetXMLObject()
		with oxml as MSXML2
			.preservewhitespace = .f.
			.loadxml(Content)
			with (.ParseError)
				if (.ErrorCode!=0)
					set step on
					debugout .ErrorCode,.srcText,.reason
					=ThrowException("Błąd deserializacji XML: "+chr(13)+.srcText,KSeFErrorCode_Unknown)
				endif
			endwith
		endwith
		return oxml
	endproc

	***********************************************************************
	* Author.....: Piotr Kuliński (c) piotr.kulinski@gmail.com
	* Date.......: 2023.05.12 20:03:55
	* Comment....: Pobiera content z API i zapisuje do pliku
	*              Dodatkowo wpisuje jego zawartość do zmiennej
	*			   content_reponse
	***********************************************************************
	procedure ApiContentToFile(CFile) as void
		local length as integer
		this.file_name = CFile
		length = KSeF_ExportContentToFile(this.file_name)
		if (length > 0 and file(this.file_name))
			Content = filetostr(this.file_name)
			debugout(replicate("-",50))
			debugout(this.file_name)
			debugout("content size: "+transform(length))
			debugout(replicate("_",50))
			debugout(Content)
			debugout(replicate("-",50))
		else
			length = 0
		endif
		return length
	endproc

	***********************************************************************
	* Author.....: Piotr Kuliński (c) piotr.kulinski@gmail.com
	* Date.......: 2023.05.12 20:05:37
	* Comment....: Zrzuca aktualny content ustawiony/pobrany z api do pliku
	***********************************************************************
	procedure SaveContentToFile(CFile) as bool
		local length as integer
		length = len(this.Content.xml)
		if (length>0)
			this.file_name = CFile
			strtofile(this.Content.xml,this.file_name)
		endif
		return (length>0)
	endproc

	***********************************************************************
	* Author.....: Piotr Kuliński (c) piotr.kulinski@gmail.com
	* Date.......: 2023.05.12 20:05:37
	* Comment....: Zrzuca jedynie response z ApiResponse do pliku
	***********************************************************************
	procedure SaveBodyResponseToFile(CFile) as bool
		local length as integer
		length = len(this.Content.xml)
		if (length>0)
			this.file_name = CFile
			respType = this.Content.SelectSingleNode("/ApiResponse/@type")
			if (respType.text!="String")
				strtofile(this.Content.SelectSingleNode("/ApiResponse/response").FirstChild.xml,this.file_name)
			else
				strtofile(strconv(this.Content.SelectSingleNode("/ApiResponse/response").text,9),this.file_name)
			endif
		endif
		return (length>0)
	endproc

	************************************************************************************
	* Metoda wrappera wywołuje DLL w celu pobrania ostatnio zarejestrowanej odpowiedzi
	* Odpowiedź rejestrowana przez DLL w kontekście wywołania jakiejś metody.
	* Jeśli wywołujemy ponieranie zakresu faktur, to oczekumeny XML-a z listą faktur
	* (konkretnego typu podanego w atrybucie type) ew. oczekujemy exeption z API,
	* ale to możemy rozpownać po kodzie odpowiedzi z wywoływanej procedury.
	************************************************************************************
	procedure GetLastResponseContent() as integer {Ładujemy do bufora Content}
		with this
			.Content = .loadxml( .GetLastResponseContentString() )
			return len(.Content.xml)
		endwith
	endproc

	************************************************************************************
	* Metoda wrappera wywołuje DLL w celu pobrania ostatnio zarejestrowanej odpowiedzi
	* Odpowiedź rejestrowana przez DLL w kontekście wywołania jakiejś metody.
	* Jeśli wywołujemy pobieranie zakresu faktur, to oczekumeny XML-a z listą faktur
	* (konkretnego typu podanego w atrybucie type) ew. oczekujemy exeption z API,
	* ale to możemy rozpownać po kodzie odpowiedzi z wywoływanej procedury.
	************************************************************************************
	procedure GetLastResponseContentString() as integer {Ładujemy do bufora Content}
		local content_size as integer
		local Content as string

		Content = replicate(chr(0),DEFAULT_BUFFER_SIZE)
		content_size = KSeF_GetContent(@Content,DEFAULT_BUFFER_SIZE)
		if (content_size > DEFAULT_BUFFER_SIZE)
			Content = replicate(chr(0),content_size )
			content_size = KSeF_GetContent(@Content,content_size)
		endif
		if (content_size<=0)
			ThrowException("API nie zwróciło danych")
		endif
		debugout(replicate("-",50))
		debugout("content size: "+transform(content_size))
		debugout(replicate("-",50))
		debugout(Content)
		debugout(replicate("-",50))
		return Content
	endproc

	procedure LoadMessageException()
		local messageOut as string, ex as exception,_oe as exception
		messageOut =""
		_oe = null
		try
			ex = this.Content.SelectSingleNode("/ApiResponse/exception")
			if (!isnull(ex))
				messageOut = ex.SelectSingleNode("./serviceCtx").text+"("+ex.SelectSingleNode("./serviceName").text+")"+chr(13)
				XmlNode = ex.SelectSingleNode("./exceptionDetailList")
				if (!isnull(XmlNode ))
					for each node in XmlNode.childNodes
						if (node.nodeName=="ExceptionDetail")
							messageOut = messageOut +chr(13) +;
								"Kod błędu: "+this.XmlGetValue(node,"../ExceptionDetail/exceptionCode","5000")+chr(13)+;
								this.XmlGetValue(node,"../ExceptionDetail/exceptionDescription","nieokreślony błąd")
						endif
					endfor
				endif
			endif
		catch to _oe
			messageOut = _oe.message
			ZLogEx(_oe)
		endtry
		return messageOut
	endproc

	************************************************************************************
	* Pobranie konfiguracj aktualnie załadowanej do DLL
	************************************************************************************
	procedure GetConfiguration() as integer
		local ;
			_oe as exception,;
			result as integer
		_oe = null
		try
			result = KSeF_ExportConfiguration()
			this.GetLastResponseContent()
		catch to _oe
			result = KSeFErrorCode_Unknown
			ZLogEx(_oe)
		endtry
		return result
	endproc

	*****************************************************************************************************
	** Nawiązanie połączenia z KSeF
	** W pierwszej kolejności jest ustalana autentykacja na podstawie kontekstu bazy danych.
	** Jeśli w obiekcie ksef .Authentication nie jest ustalony identyfikator (nie wykonana SetUser
	** z odpowiednim kontekstem logowania), wóczas przyjmujemy że logujemy się w kontekscie apitoken
	** Metoda obiektu bazy danych (odb.GetAuthentication()) ustawia kontekst użytkownika na apitoken, o ile
	** udało sie jej ustalić firmę i token.
	**
	*****************************************************************************************************
	procedure connect(odb as [KSeFDB] of "prgs\common\ksefapi.prg") as void
		local result as integer
		local isReload as boolean
		local clientContext as string
		local auth as object
		local ctx_user as object
		local xpath as string
		local citem as String
		local item as object
		
		assert .f. message "KSeF connect"	
		with this
			raiseevent(this,"OnProgressMessage","Łączenie z Krajowym Systemem eFaktur")
			raiseevent(this,"OnCheckConnection",this)
			
			auth = odb.GetAuthentication()
			xpath = "/Authentication/Context/Identifier"
	
			.Authentication.SelectSingleNode(xpath).text = auth.selectSingleNode(xpath).text

			ctx_user = .Authentication.SelectSingleNode("/Authentication/User")
			item = .XmlGetOrCreateNode(ctx_user,"Type","apitoken")
			if (!isnull(item))
				if (alltrim(item.text)=="apitoken")
					item = .XmlGetOrCreateNode(ctx_user,"Identifier",auth.selectSingleNode("/Authentication/User/Identifier").text)
				endif
			else
				=ThrowException("Niewłaściwa autentykacja użytkownika")
			endif

			result = KSeF_LoadSession(.authentication.xml)
			.GetLastResponseContent()
			raiseevent(this,"OnSaveResponse",this)
			if (result!=KSeFErrorCode_OK)
				result = KSeF_Connect(.authentication.xml,true)
				.GetLastResponseContent()
				raiseevent(this,"OnSaveResponse",this)
				.ThrowIfException("Błąd połączenia z KSeF")
			endif
			.MailerConfiguration = odb.MailerConfiguration			
			raiseevent(this,"OnConnect",this)
		endwith
	endproc

	procedure SessionTerminate() as integer {0 - ok}
		local ;
			_oe as exception,;
			result as integer

		_oe = null
		try
			result = KSeF_SessionTerminate()
			this.GetLastResponseContent()
			raiseevent(this,"OnSaveResponse",this)
			this.sessionReferenceNumber=[]
		catch to _oe
			result = KSeFErrorCode_Unknown
			ZLogEx(_oe)
		endtry
		return result
	endproc

	************************************************************************************************************
	** Zamknięcie sesji wskazanej przez jej numer i token
	** w zasadzie wystarczył by token, ale w przypadku błędnej odpowiedzi w responsie nie pojawi się numer
	** sesji i trudno będzie to wiązać po stronie SQL
	************************************************************************************************************
	procedure SessionTerminateSpecified(_sessionReferenceNumber as string,_token as string) as integer {0 - ok}
		local ;
			_oe as exception,;
			result as integer

		_oe = null
		try
			raiseevent(this,"OnProgressMessage",_sessionReferenceNumber )
			result = KSeF_SessionTerminateSpecified(_sessionReferenceNumber,_token)
			this.GetLastResponseContent()
			raiseevent(this,"OnSaveResponse",this)
			if (this.sessionReferenceNumber==_sessionReferenceNumber)
				this.sessionReferenceNumber=[]
			endif
		catch to _oe
			result = KSeFErrorCode_Unknown
			ZLogEx(_oe)
		endtry
		return result
	endproc


	procedure SessionTerminateAll() as void
		local ;
			_oe as exception
		_oe = null
		try
			=KSeF_SessionTerminateAll()
			this.GetLastResponseContent()
			raiseevent(this,"OnSaveResponse",this)
			this.sessionReferenceNumber=[]
		catch to _oe
			ZLogEx(_oe)
		endtry
	endproc

	procedure SessionStatus(SessionRefNumber as string) as integer
		local result as integer
		raiseevent(this,"OnProgressMessage","Pobieranie statusu wysłanej faktury")
		result = KSeF_SessionStatus(iif(vartype(SessionRefNumber)!='C',null,SessionRefNumber))
		=this.GetLastResponseContent()
		raiseevent(this,"OnSessionStatus",this)
		raiseevent(this,"OnProgressMessage","")
		return result
	endproc

	procedure CommonSessionStatus(SessionRefNumber as string) as void
		raiseevent(this,"OnProgressMessage","Pobieranie statusu sesji i dokumentów UPO")
		=KSeF_CommonSessionStatus(iif(vartype(SessionRefNumber)!='C',null,SessionRefNumber))
		=this.GetLastResponseContent()
		raiseevent(this,"OnCommonStatus",this) && pod to zdarzenie podpinać również OnSaveResponse
		raiseevent(this,"OnProgressMessage","")
	endproc

	procedure InvoicesDownloadRangeDate(_from_date as string, _to_date as string, _subject as integer, _PageSize as integer, _PageOffset as integer) as void {0-są dane}
		if (vartype(_PageSize)!="N")
			_PageSize=100
		endif
		if (vartype(_PageOffset)!="N")
			_PageOffset=0
		endif
		raiseevent(this,"OnProgressMessage","Ładowanie dokumentów z systemu KSeF")
		=KSeF_InvoicesDownloadRangeDate(_from_date,_to_date,_subject,_PageSize,_PageOffset)
		this.GetLastResponseContent()
		raiseevent(this,"OnInvoiceQuery",this)
		raiseevent(this,"OnProgressMessage","")
	endproc

	procedure InvoicesDownloadIncrement(;
			_acquisitionTimestampThresholdFrom as string,;
			_acquisitionTimestampThresholdTo as string,;
			_subject as integer,;
			_PageSize as integer,;
			_PageOffset as integer;
			) as void
		if (vartype(_PageSize)!="N")
			_PageSize=100
		endif
		if (vartype(_PageOffset)!="N")
			_PageOffset=0
		endif
		raiseevent(this,"OnProgressMessage","Ładowanie dokumentów z systemu KSeF")
		=KSeF_InvoicesDownloadIncrement(_acquisitionTimestampThresholdFrom,_acquisitionTimestampThresholdTo,_subject,_PageSize,_PageOffset)
		=this.GetLastResponseContent()
		raiseevent(this,"OnInvoiceQuery",this)
		raiseevent(this,"OnProgressMessage","")
	endproc

	procedure InvoicesDownloadDetail(;
			_from_date as string,;
			_to_date as string,;
			_subject as integer,;
			_invoiceNumber as string,;
			_ksefReferenceNumber as string,;
			_PageSize as integer,;
			_PageOffset as integer;
			) as void

		if (vartype(_PageSize)!="N")
			_PageSize=100
		endif
		if (vartype(_PageOffset)!="N")
			_PageOffset=0
		endif
		raiseevent(this,"OnProgressMessage","Ładowanie dokumentów z systemu KSeF")
		=KSeF_InvoicesDownloadDetail(_from_date,_to_date,_subject,_invoiceNumber, _ksefReferenceNumber, _PageSize,_PageOffset)
		=this.GetLastResponseContent()
		raiseevent(this,"OnInvoiceQuery",this)
		raiseevent(this,"OnProgressMessage","")
	endproc

	procedure OnlineInvoiceGet(nrKSeF as string) as integer
		local result as integer
		result = KSeF_OnlineInvoiceGet(nrKSeF)
		this.GetLastResponseContent()
		raiseevent(this,"OnOnlineInvoiceGet",this)
		return result
	endproc

	procedure UPOToHTML(Content as string) as void
		=KSeF_ConvertUPOtoHTML(Content)
		=this.GetLastResponseContent()
		this.tag = Content
		raiseevent(this,"OnUPOView",this)
	endproc

	procedure InvoiceXMLtoHTML(Content as string) as void
		raiseevent(this,"OnProgressMessage","Generowanie podglądu dokumentu")
		=KSeF_InvoiceXMLtoHTML(Content,"")
		=this.GetLastResponseContent()
		raiseevent(this,"OnProgressMessage","")
		raiseevent(this,"OnHtmlView",this)
	endproc

	procedure InvoiceXMLtoHTMLwithKsefNumber(Content as string, ksef_number as string) as void
		raiseevent(this,"OnProgressMessage","Generowanie podglądu dokumentu")
		=KSeF_InvoiceXMLtoHTML(Content,ksef_number)
		=this.GetLastResponseContent()
		raiseevent(this,"OnProgressMessage","")
		raiseevent(this,"OnHtmlView",this)
	endproc

	procedure ValidateInvoice()
		raiseevent(this,"OnProgressMessage","Sprawdzanie poprawności dokumentu")
		= KSeF_ValidateInvoice(.ElementResponse(""))
		this.GetLastResponseContent()
		raiseevent(this,"OnProgressMessage","")
	endproc

	procedure ThrowIfException(_message as string)
		if (!isnull(this.Content.SelectSingleNode("/ApiResponse/exception")))
			_message = iif(isnull(_message) or empty(_message),"",_message+chr(13)+.LoadMessageException())
			=ThrowException(_message)
		endif
	endproc

	**************************************************************************************
	* Author.....: Piotr Kuliński (c) piotr.kulinski@gmail.com
	* Date.......: 2023.05.16 12:39:21
	* Comment....: Wysłanie fakury do KSeF
	**************************************************************************************
	* Wysyłamy fakturę xml zapisaną w elemencie /ApiResponse/response
	* Na wysyłkę dostajemy odrazu odpowiedź informującą o wysysłce z numerem referencyjnym
	* wysyłki, którym możemy się później posługiwać do bieżącego sprawdzania statusu
	* wysłanej faktury (metoda InvoiceStatusFromReferenceNumver).
	* Jeśli podpięliśmy obsługę zdarzeń OnValidateInvoice, to po poprawnym pobraniu dokumentu
	* z systemu bazodanowego nastąpi jej weryfikacja zgodności ze schematem KSeF.
	* Jeśli walidacja nie przejdzie rzucany jest wyjątek, w przeciwnym wypadku wysyłka jest
	* kontynuowana, a metoda KSeF_SendInvoiceXML zwraca obiekt OnlineInvoiceResponse lub/i
	* wyjątek i przerywa działanie.
	**************************************************************************************
	procedure SendInvoiceXML() as void {throw exception}
		local invoice_body as string
		with this
			raiseevent(this,"OnGetInvoice",this)
			invoice_body = .ElementResponse("")
			raiseevent(this,"OnValidateInvoice",this)
			=KSeF_SendInvoiceXML(invoice_body)
			.GetLastResponseContent()
			raiseevent(this,"OnSendInvoice",this)
		endwith
	endproc
	
	procedure GetVerificationLink(KSeFReferenceNumber as string) as void {throw exception}
		local link_verification as string
		with this
			=KSeF_GetVerificationLink(KSeFReferenceNumber,"")
			.GetLastResponseContent()
			link_verification = .XmlGetValue(.Content,"/ApiResponse/response","")
		endwith
		return link_verification
	endproc	

	procedure GetQRCode(toQr as string, width as integer,height as integer, format as dtring) as void {throw exception}
		local QRRequest as string
		with this
			text to QRRequest noshow textmerge pretext 15 &&3 additive
				<QRCode>
					<content width="<<width>>" height="<<height>>" format="<<format>>"><<toQr>></content>
				</QRCode>
			ENDTEXT

			=KSeF_GetQRCode(QRRequest)
			.GetLastResponseContent()
		endwith
	endproc

	procedure SendInvoiceAndCheck() as void {throw exception}
		local isPositiveResponse as logical
		local nTimeout as double
		local currentTime as double
		local elementReferenceNumber as string
		local processingDescription as string
		local processingCode as int
		local invoice_body as string

		with this as "KSeFApi" of "prg\common\ksefapi.prg"
			raiseevent(this,"OnProgressMessage","Wysyłka dokumentu do Krajowego Systemu eFaktur")
			raiseevent(this,"OnGetInvoice",this) && pobranie XML z fakturą
			invoice_body = .ElementResponse("")
			raiseevent(this,"OnValidateInvoice",this) && Walidacja poprawności

			=KSeF_SendInvoiceXML(invoice_body)
			.GetLastResponseContent()

			** OnlineInvoiceResponse -- do tego trzeba dodać IDDokumentu, ponieważ oznaczanie
			** na podstawie numeru faktury (w przypadku jeśli bedzie istniało kilka takich samych numerów,
			** a niełaściwy da się tego wykluczyć wówczas nie zostanie odszukany
			** właściwy numer faktury

			onode = this.XmlGetNode(this.Content,"/ApiResponse[@type='OnlineInvoiceResponse']/response")
			if (!isnull(onode))
				local XMLElement as object
				XMLElement = this.Content.createNode(1,"invoiceId","")
				XMLElement.text = transform(this.IDSaleDocument)
				onode.appendChild(XMLElement)
			endif

			raiseevent(this,"OnSendInvoice",this) && jeśli pod zdarzenie będzie podpięty również OnThrowIfException.. to przerwie tu przetwarzanie

			elementReferenceNumber = .ElementResponse("/elementReferenceNumber")
			currentTime = seconds()
			nTimeout = currentTime+30.000 && 30 sek
			do while (nTimeout > currentTime)

				=KSeF_InvoiceStatus(elementReferenceNumber)
				.GetLastResponseContent()

				* 200 - Zakończenie etapu archiwizacji danych faktury
				* 336 - Zakończenie etapu akceptacji faktury oraz generowania numeru KSeF
				processingCode = .ElementResponse("/processingCode")
				processingCode = iif(isnull(processingCode),0,val(processingCode))
				isPositiveResponse = (processingCode==200 or processingCode==336or processingCode==325)
				if (isPositiveResponse)
					exit
				endif
				processingDescription = .ElementResponse("/processingDescription")
				raiseevent(this,"OnProgressMessage","Oczekiwanie ("+ltrim(transform(nTimeout-currentTime,"999 sek"))+") na zakończenie etapu akceptacji faktury oraz generowania numeru KSeF")
				sleep(1000)
				currentTime = seconds()
			enddo

			if (isPositiveResponse)
				raiseevent(this,"OnProgressMessage","Zakończenie etapu akceptacji faktury oraz generowania numeru KSeF")
				raiseevent(this,"OnSaveResponse",this) && powinien być rollback jeśli !isPositiveResponse
			else
				raiseevent(this,"OnProgressMessage","Wycofanie eksportu dokumentu, brak potwierdzenia z KSeF")
				raiseevent(this,"OnRollbackExport",this)
				ThrowException("Wysyłka dokumentu do systemu KSeF nie powiodła się"+chr(13)+chr(13)+iif(!isnull(processingDescription),processingDescription,"Brak potwierdzenia dobioru") )
			endif
			raiseevent(this,"OnProgressMessage","")
		endwith
	endproc

	**************************************************************************************************
	* Author.....: Piotr Kuliński (c) piotr.kulinski@gmail.com
	* Date.......: 2023.05.16 14:02:07
	* Comment....: Odpowiedź typu InvoiceStatus, jeśli kod jest poprawny wówczas zawiera numera
	* referencyjny faktury nadany w KSeF
	**************************************************************************************************
	procedure InvoiceStatusFromReferenceNumber(elementReferenceNumber as string) as void
		=KSeF_InvoiceStatus(elementReferenceNumber)
		this.GetLastResponseContent()
		raiseevent(this,"OnSaveResponse",this)
	endproc

	***************************************************************************
	* Pobranie statusu wysłanej faktury na podstawie odpowiedzi
	* z wysyłki faktury
	***************************************************************************
	procedure InvoiceStatusFromResponse() as integer
		return this.InvoiceStatusFromReferenceNumber(this.ElementResponse("/elementReferenceNumber") )
	endproc

	procedure SQLiteQuery(qry as String, crl as string)
		local xml_result as string, result as integer
		with this
			result = KSeF_Query(qry)
			if (result==1)
				xml_result = .GetLastResponseContentString()
				if (len(xml_result)>0)
					strtofile(xml_result,"p:\"+transform(int(seconds()*1000))+'.xml')
					xmltocursor(xml_result,crl)
				endif				
			endif
		endwith
		return result
	endproc
	

	procedure QueuePutInvoice()
		local invoice_body as string
		with this
			raiseevent(this,"OnGetInvoice",this)
			invoice_body = .ElementResponse("")
			=KSeF_QueuePutInvoice(invoice_body)
		endwith
	endproc
	
	procedure destroy()
		local ;
			_oe as exception
		_oe = null
		try
			**KSeF_QueueSendInvoiceTaskStop()
			try
				#if .f.
					if (this.lpEnumFunc>0)
						DestroyCallbackFunc(this.lpEnumFunc)
					endif
				#endif
				**this.SessionTerminateAll()
				KSeF_Destroy()
			catch
			endtry

			try
				ssml_Destroy()
			catch
			endtry

			ReleaseDll(this.DLLFile)
			ReleaseDll("SoftechSendMailLibrary.dll")

		catch to _oe
			ZLogEx(_oe)
		endtry
	endproc

	*!*		procedure error(p1,p2,p3) as void helpstring "Globalna funkcja obsługi błędów"
	*!*			local ;
	*!*				ErrorCount as integer,;
	*!*				errormessage as string,;
	*!*				_oe as exception,;
	*!*				I as integer,;
	*!*				x as integer
	*!*			local ;
	*!*				array ErrorStruct[1] as Variant

	*!*			_oe=null
	*!*			try
	*!*				ZLog(transform(p1)+":"+transform(p2)+":"+transform(p3))
	*!*				ErrorCount=aerror(ErrorStruct)
	*!*				m.ErrorMessage=transform(p1)+'#'+transform(p2)+'#'+transform(p3)+chr(13)

	*!*				for I=1 to ErrorCount
	*!*					for x=1 to 7
	*!*						m.ErrorMessage = m.ErrorMessage + transform(x)+" - "+alltrim(transform(ErrorStruct[I,X]))+chr(13)
	*!*					endfor
	*!*				endfor
	*!*				debugout m.ErrorMessage, "Wyjątek w module KSeF Wrapper"
	*!*			catch to _oe
	*!*				debugout _oe.message,_oe.details
	*!*				ZLogEx(_oe)
	*!*			endtry
	*!*			set step on
	*!*			return isnull(_oe)
	*!*		 endproc


	***************************************************************************************************
	** Ważna rzecz, zwrócić uwagę na wielkość liter w nawie rzeczywistych metod z DLL,
	** Muszą być dokładnie takie same jak w wyeksportowanych metodach
	****************************************************************************************************
	protected procedure declare()
		local dll_sendmail as string
		dll_sendmail = "SoftechSendMailLibrary.dll"
		try
			declare long GetLastError in kernel32 as GetLastError
			declare integer SetPasswordCommunication in (this.DLLFile) as KSeF_SetPasswordCommunication string base64_password
			declare integer SetConfiguration in (this.DLLFile) as KSeF_SetConfiguration string xml_content
			declare integer SetConfigOnlyApi in (this.DLLFile) as KSeF_SetConfigOnlyApi string xml_content
			
			** deprecatetd
			** declare integer SessionInitialize in (this.DLLFile) as KSeF_SessionInitialize
			** declare integer SessionInitializeAndWaitForActivation in (this.DLLFile) as KSeF_SessionInitializeAndWaitForActivation			
			** declare integer LoadSession in (this.DLLFile) as KSeF_LoadSession
			
			declare integer Connect in (this.DLLFile) as KSeF_Connect  string authenctication, integer waitForAccept
			declare integer LoadSession in (this.DLLFile) as KSeF_LoadSession string authentication
			
			declare integer SessionTerminate in (this.DLLFile) as KSeF_SessionTerminate
			declare integer SessionTerminateAll in (this.DLLFile) as KSeF_SessionTerminateAll
			declare integer SessionTerminateSpecified in (this.DLLFile) as KSeF_SessionTerminateSpecified string sessionReferenceNumber, string token
			declare integer SessionStatus in (this.DLLFile) as KSeF_SessionStatus string refnumber
			declare integer SendInvoiceXML in (this.DLLFile) as KSeF_SendInvoiceXML string contents_xml
			declare integer InvoiceStatus in (this.DLLFile) as KSeF_InvoiceStatus string InvRefNumer
			declare integer ValidateXml in (this.DLLFile) as KSeF_ValidateInvoice string Content
			declare integer CommonSessionStatus in (this.DLLFile) as KSeF_CommonSessionStatus string refnumber
			declare integer OnlineInvoiceGet in (this.DLLFile) as KSeF_OnlineInvoiceGet string refnumber
			declare integer ExportContentToFile in (this.DLLFile) as KSeF_ExportContentToFile string fileName
			declare integer GetContent in (this.DLLFile) as KSeF_GetContent string @ Content, integer size_buffer
			declare integer ExportConfiguration in (this.DLLFile) as KSeF_ExportConfiguration

			declare integer InvoicesDownloadRangeDate in (this.DLLFile) as KSeF_InvoicesDownloadRangeDate ;
				string _from,;
				string _to,;
				integer _subject,;
				integer _PageSize,;
				integer _PageOffset

			declare integer InvoicesDownloadDetail in (this.DLLFile) as KSeF_InvoicesDownloadDetail ;
				string _from,;
				string _to,;
				integer _subject,;
				string _invoiceNumber, ;
				string _ksefReferenceNumber, ;
				integer _PageSize, ;
				integer _PageOffset

			declare integer InvoicesDownloadIncrement in (this.DLLFile) as KSeF_InvoicesDownloadIncrement ;
				string _acquisitionTimestampThresholdFrom,;
				string _acquisitionTimestampThresholdTo,;
				integer _subject,;
				integer _PageSize,;
				integer _PageOffset

			declare integer ConvertInvoiceXMLtoHTML in (this.DLLFile) as KSeF_InvoiceXMLtoHTML string _from, string ksef_number
			**dodane 2023-06-20 (wymagany parametr w konfiguracji, wskazujący xsl z translacją
			declare integer ConvertUPOtoHTML in (this.DLLFile) as KSeF_ConvertUPOtoHTML string base64_upo
			declare integer RegisterCallbackResponse in (this.DLLFile) as KSeF_RegisterCallbackResponse long ptrFun
			declare integer DestroyLibrary in (this.DLLFile) as KSeF_Destroy
			
			declare long GetQRCode in (this.DLLFile) as KSeF_GetQRCode string qr_definition
			declare long GetVerificationLink in (this.DLLFile) as KSeF_GetVerificationLink string KSeFReferenceNumber, string InvoiceSHA

			declare integer dbSqlQuery in (this.DLLFile) as KSeF_Query string txt_query
			declare integer QueuePutInvoice in (this.DLLFile) as KSeF_QueuePutInvoice string invoiceBody
			declare integer QueueSendInvoiceTaskStart in (this.DLLFile) as KSeF_QueueSendInvoiceTaskStart
			declare integer QueueSendInvoiceTaskStop in (this.DLLFile) as KSeF_QueueSendInvoiceTaskStop
			
			declare integer Sleep in win32api integer

			declare Initialize in (dll_sendmail) as ssml_Initialize
			declare destroy in (dll_sendmail) as ssml_Destroy
			declare SetSessionKey in (dll_sendmail) as ssml_SetSessionKey string base64_key
			declare SetCredential in (dll_sendmail) as ssml_SetCredential string hex_encrypt_credential
			declare SetMailFileDefinition in (dll_sendmail) as ssml_SetMailFileDefinition string xml_mail_definition
			declare integer Send in (dll_sendmail) as ssml_Send

			this.isload=.t.

		catch to _oe
			this.isload=.f.
			=ThrowException( "Błąd ładowania bibliotek KSeF "+_oe.message )
			*this.errorno=_oe.errorno
			*this.message=_oe.message+chr(13)+_oe.linecontents
		endtry

		return this.isload
	endproc

	procedure SqlGetConfiguration() as bool
		local ;
			SqlCommand as string,;
			_value as string,;
			_length as integer

		text to SqlCommand noshow textmerge pretext 15
            select convert(xml,tekstowa)
            from parametry
            where zmienna='x_ksef_configuration'
		ENDTEXT
		=SqlCommandExecuteException(_screen.gastro.ConnectionHandle,@SqlCommand,"ksef_configuration")
		select ksef_configuration
		_value = curval(field(1))
		_length  = len(_value)
		if (_length>0)
		    local encrypt_conf
  		    encrypt_conf = CryptAPI(strconv(_value ,16),_screen.gastro.hash,0)
  		    encrypt_conf ='<?xml version="1.0"?><Configuration xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema"><APIVersion>2</APIVersion><URL>'+;
  		    'https://ksef-test.mf.gov.pl/api</URL><APIToken>3144F194CE97D0BB51C22B651883CC29E59575DB9BB02CBC3E45AB8008F63C1F</APIToken><Authorization type="token">'+;
  		    '3144F194CE97D0BB51C22B651883CC29E59575DB9BB02CBC3E45AB8008F63C1F</Authorization><NIP>1147581153</NIP><UserAgent>GastroKlasyka</UserAgent>'+;
  		    '<DirectoryWorking root="ksef" request="" response="ksef\communication" result="ksef\communication" sessions="ksef" document="ksef\invoice"/>'+;
  		    "<PublicPEMKey>MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAuWosgHSpiRLadA0fQbzshi5TluliZfDsJujPlyYqp6A3qnzS3WmHxtwgO58uTbemQ1HCC2qwrMwuJqR6l8tgA4"+;
  		    "ilBMDbEEtkzgbjkJ6xoEqBptgxivP/ovOFYYoAnY6brZhXytCamSvjY9KI0g0McRk24pOueXT0cbb0tlwEEjVZ8NveQNKT2c1EEE2cjmW0XB3UlIBqNqiY2rWF86DcuFDTUy+KzSmTJTFvU/"+;
  		    "ENNyLTh5kkDOmB1SY1Zaw9/Q6+a4VJ0urKZPw+61jtzWmucp4CO2cfXg9qtF6cxFIrgfbtvLofGQg09Bh7Y6ZA5VfMRDVDYLjvHwDYUHg2dPIk0wIDAQAB</PublicPEMKey>"+;
  		    '<PublicPEMFile>p:\KSEF\publicKey_test.pem</PublicPEMFile><SchemasAndStyles xsl_invoice_to_html="file://c:\styluproszczony_vfa2.xsl"'+;
  		    ' xsl_upo_to_html="file://upo.xsl" xsd_invoice_validation="http://crd.gov.pl/wzor/2023/06/29/12648/schemat.xsd"/><TimeoutCommunication dll_put="00:00:30" dll_post="00:00:30"'+;
  		    ' dll_get="00:00:45" dll_wait_for_session="00:00:30" ksef_session_timeout="02:00:00"'+;
  		    ' ksef_refresh_session_befor_end="01:57:00.001"/><Logger enabled="true" limit_size_MB="1" file_name="{DirectoryWorking.root}\KSeF.log"/>'+;
  		    '<Encryption enabled="false" password="QmFyZHpvRLN1Z2llSGFzs29Eb1pha29kb3dhbmlh"/>'+;
  		    '</Configuration>'
			this.Content = this.loadxml(encrypt_conf) && &&strconv(_value,10)
		endif
		use in "ksef_configuration"
		return (_length>0)
	endproc

	procedure GetTimeZone()
		declare integer GetTimeZoneInformation in Win32API string @ TimeZoneStruct
		local lcTZ,lnDayLightSavings,lnOffset
		lcTZ = space(256)
		lnDayLightSavings = GetTimeZoneInformation(@lcTZ)
		lnOffset = this.CharToBin(substr(lcTZ,1,4),.t.)
		*** Subtract an hour if daylight savings is active
		if lnDayLightSavings = 2
			lnOffset = lnOffset - 60
		endif
		return abs(lnOffset)
	endproc

	function CharToBin(lcBinString,llSigned)
		local m.I, lnWord
		lnWord = 0
		for m.I = 1 to len(lcBinString)
			lnWord = lnWord + (asc(substr(lcBinString, m.I, 1)) * (2 ^ (8 * (m.I - 1))))
		endfor
		if llSigned and lnWord > 0x80000000
			lnWord = lnWord - 1 - 0xFFFFFFFF
		endif
		return lnWord
	endfunc

	function GetUtcTimeString(sTime)
		**transform(datetime(),"@YL")
		return this.GetUtcTime( ctot( evl(sTime,dtos(datetime())) ))
	endfunc

	function GetUtcTime(ltTime) as datetime
		if empty(ltTime)
			ltTime = datetime()
		endif
		*** Adjust the timezone offset
		return ltTime + (this.GetTimeZone() * 60)
	endfunc

enddefine

