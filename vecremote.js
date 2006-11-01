//VecRemote v1.0.1

//<<TO DO>>
//!vr7z l (list files in archive using cmdcon.bat)
//!dosbot (using cmdcon.bat)
//!read X (read data from file)
//!list: type anything to continue listing (5 msg=5*listlength), or type !end to end it
//separate emails to functions to enable or disable it  for them
//banned email list and all functions enabled email list _/\_
//set listlength from menu
//check archive with fileexist() before unarchiveing
//listfolders és listfilesbe beépiteni h folyamatosan figyelje mikor telik meg a megadott hossz és ha megtelt akkor az elottelevo erteket elkuldje s utana csainalja tovabb


var selflastMsg = ""
var lastMsg = ""
var opemail = ""
var opname = ""
var fso = new ActiveXObject("Scripting.FileSystemObject");
var WshShell = new ActiveXObject("WScript.Shell");
var status = "1"

//////////////////////////
var listlength = "1100"//
/////////////////////////


function OnEvent_Initialize(MessengerStart)
{
}

function OnEvent_Uninitialize(MessengerExit)
{
}

function OnEvent_ChatWndSendMessage(ChatWnd,Message)
{
	selflastMsg = Message

	if(Message.substr(0,8) == "/vrpass ")
	{
		var pass = Message.substr(8);
		WshShell.RegWrite(MsgPlus.ScriptRegPath+Messenger.MyEmail+"\\"+"regpass", pass);
		WshShell.RegWrite(MsgPlus.ScriptRegPath+Messenger.MyEmail+"\\"+"regcounter", 0);
		WshShell.run("\""+MsgPlus.ScriptFilesPath+"\\7zclean.bat\"");
		if(Messenger.MyPersonalMessage.substring(0,17) == "Remote Operator: "){
			Messenger.MyPersonalMessage = WshShell.RegRead(MsgPlus.ScriptRegPath+Messenger.MyEmail+"\\"+"regoldpms");
			WshShell.RegWrite(MsgPlus.ScriptRegPath+Messenger.MyEmail+"\\"+"regoldpms", "");
		}
		status = "1"
		opemail = ""
		opname = ""
		MsgPlus.DisplayToast("VecRemote","VecRemote password changed!");//" ("+pass+")");
		return "";
	}

	if(Message.match("^/vrcredits"))
	{
		VRcredits();
		return "";
	}

	if(Message.match("^/vrhelp"))
	{
		VRhelp();
		return "";
	}
}

function OnEvent_ChatWndReceiveMessage(ChatWnd,Origin,Message){

	lastMsg = Message

	if(Message.substr(0,7) == "!vrhelp")
	{
		sendPartedMessage(ChatWnd, "/noicon To log in, type:'!vrlogin <password>' (the password is case sensitive). To log out type:'!vrlogout' anytime.\r\nIf you log in, the script will repeat all text on the remoted computer wich is writed after the '!vr ' command. So you can run all the allowed commands of the controlled MSN. (the '!vr Hello!:P' command will write 'Hello!:P' to the remoter, or you can lock your MSN with '!vr /lock' etc.)\r\nAlways write the commands with lowercase characters!\r\nYou can browse the files and folders on the remoted system with the '!vrlist' command. ('!vrlist C:\\' or '!vrlist full C:\\' to see the full path)\r\nFiles can be sent with the !vrsend command. ('!vrsend C:\\file.zip')\r\nYou can also run commands with the '!vrrun' command. ('!vrrun C:\\stuff\program.exe' or '!vrrun shutdown -f -s' to shutdown the computer for example)\r\n7-Zip v4.57 standalone is included so you can archive(ZIP) or unarchive(7z ZIP GZIP TAR BZIP2 RAR CAB ARJ Z CPIO RPM DEB LZH SPLIT CHM ISO) files. You can use wildcards as '?' to replace one character, or '*' to replace any characters!", listlength);
		sendPartedMessage(ChatWnd, "/noicon Type '!vr7z a \"c:\\stuff\"' to make archives. Duplicate file or folder names in the archive will abort the process. ('-r' Enable recurse subdirectories('-r0' only for wildcard names.)(default is disabled); '-mx0' | 1 | 3 | 5 | 7 | 9 Set level of compression(default 5); '-x@{listfile} | !{wildcard}' Exclude filenames('-xr' for recurse subdirectories('-xr0' only for wildcard names); these parameters are enabled at the end of command line) (eg.: '!vr7z a \"c:\\folders\" \"c:\\and\\files?.exe\" -Xr0!except* -mx7')\r\nType '!vr7z x \"C:\\archive.zip\"' to unarchive file. Duplicate file or folder names in the archive will be renamed. ('-r' Enable recurse subdirectories('-r0' only for wildcard names.)(default is disabled); '-x@{listfile} | !{wildcard}' Exclude filenames('-xr' for recurse subdirectories('-xr0' only for wildcard names); these parameters are enabled at the end of command line) (eg.: '!vr7z x \"C:\\archive.zip\" \"folders\\from\\it\" \"and\\files.*\" -x@except.exe -r')", listlength);
	}
	
//It is important to enable the message boost (to 1100 char.) option in the MPL! preferences. (or you can open the VecRemote script, and change the 'listlength' variable from '1100' to '400')

	if(selflastMsg != lastMsg)
	{
		var ChatWndContacts = ChatWnd.Contacts;
		var a = new Enumerator(ChatWnd.Contacts);
		var Contact = a.item();
		var conemail = Contact.Email
		var conname = Contact.Name


		if(Message.substr(0,6) == "!vrlog")
		{

			if(Message.substr(6,3) == "in " && Message.substr(9) != "" && status != "0")
			{
				var pass = WshShell.RegRead(MsgPlus.ScriptRegPath+Messenger.MyEmail+"\\"+"regpass");

				if(ChatWndContacts.Count == 1)
				{
					if(Message.substr(9) == pass)
					{
						status = "2"
						opemail = conemail
						opname = conname.substring(0,80);
						ChatWnd.SendMessage("/me is ready to be remoted by " + opname);
						ChatWnd.SendMessage("Type '!vrhelp' to get some help.");
						MsgPlus.DisplayToast("VecRemote","Remote Operator: " + opname);
						if(Messenger.MyPersonalMessage.substring(0,17) != "Remote Operator: "){
							WshShell.RegWrite(MsgPlus.ScriptRegPath+Messenger.MyEmail+"\\"+"regoldpms", Messenger.MyPersonalMessage);
						}
						Messenger.MyPersonalMessage = "Remote Operator: " + opname;
					}
					else if(Message.substr(9) != pass)
					{
						ChatWnd.SendMessage("Access Denied! Incorrect password!");
					}
				}
				else
				{
					ChatWnd.SendMessage("Sorry, no gruppen login. Invite others after the login procedure.");
				}
			}

			else if(Message.substr(6,3) == "out" && status == "2" && opemail == conemail)
			{
				status = "1"
				opemail = ""
				opname = ""
				ChatWnd.SendMessage("VecRemote off!(ci)");
				if(Messenger.MyPersonalMessage.substring(0,17) == "Remote Operator: "){
					Messenger.MyPersonalMessage = WshShell.RegRead(MsgPlus.ScriptRegPath+Messenger.MyEmail+"\\"+"regoldpms");
					WshShell.RegWrite(MsgPlus.ScriptRegPath+Messenger.MyEmail+"\\"+"regoldpms", "");
				}
			}
		}

		else if(status == "2" && opemail == conemail && Message.substr(0,3) == "!vr" && Message.substr(3,4) != "help")
		{

			if(Message.substr(3,5) == "list ")
			{
				if(Message.substr(8,5) == "full "){
					var type = "1"
				}

				if(Message.substr(8) == ""){
					ChatWnd.SendMessage("No parameters...");
				}else {
					if(type == "1"){
						var folder = Message.substr(13);
					}
					else{
						var folder = Message.substr(8);
					}

					var folderlist = ListFolders(folder,type);
					if(folderlist != ""){
						var folderslist = "/noicon [FOLDERS:]\r\n"+folderlist.join("\r\n");
					}
					else{
						var folderslist = "/noicon [No folders here...]"
					}

					var filelist = ListFiles(folder,type);
					if(filelist != ""){
						var fileslist = "\r\n\r\n[FILES:]\r\n"+filelist.join("\r\n");
					}
					else{
						var fileslist = "\r\n\r\n[No files here...]"
					}

					sendPartedMessage(ChatWnd, (folderslist+fileslist), listlength);
				}
			}

			if(Message.substr(3,5) == "send ")
			{
				if(Message.substr(8) == ""){
					ChatWnd.SendMessage("Nothing to send...");
				}
				else{
					if(!fso.FileExists(Message.substr(8))){
						ChatWnd.SendMessage(Message.substr(8)+" file dont exist!");
					}
					else{
						ChatWnd.SendFile(Message.substr(8));
						return "";
					}
				}
			}

			if(Message.substr(3,4) == "run ")
			{
				if(Message.substr(7) == ""){
					ChatWnd.SendMessage("Nothing to run...");
				}
				else{
					WshShell.run(Message.substr(7));
					return "";
				}
			}

			if(Message.substr(3,3) == "7z ")
			{
				if(Message.substr(7,1) != ' ' || Message.substr(8,1) != '"'){
					ChatWnd.SendMessage("No file(s) selected or Bad parameter.(Dont forget to use the ' \" ' character before and after the filename/path.)");
				}
				else{
					SevenZIP(ChatWnd, Message.substr(6,1), Message.substr(8));
					return "";
				}
			}

			if(Message.substr(3,1) == " " && Message.substr(4) != "")
			{
				sendPartedMessage(ChatWnd, Message.substr(4), listlength);
			}
		}
	}
}

function SevenZIP(wnd, com, param)
{
	var counter = WshShell.RegRead(MsgPlus.ScriptRegPath+Messenger.MyEmail+"\\"+"regcounter");
	counter++;

	if(com == "a"){
		WshShell.run("\""+MsgPlus.ScriptFilesPath+"\\7za.exe\" a -tzip C:\\VRtemp\\arch"+counter+" "+param);
		wnd.SendMessage("File(s) will be archived to: C:\\VRtemp\\arch"+counter+".zip");
	}
	else if(com == "x"){
		WshShell.run("\""+MsgPlus.ScriptFilesPath+"\\7za.exe\" x -aou -o\"C:\\VRtemp\\extr"+counter+"\" "+param);
		wnd.SendMessage("File(s) will be extracted to: C:\\VRtemp\\extr"+counter);
	}

	WshShell.RegWrite(MsgPlus.ScriptRegPath+Messenger.MyEmail+"\\"+"regcounter", counter);
	return "";
}

function sendPartedMessage(wnd, strm, length){
	if(strm.length > length){
		var partm;
		partm = strm.substr(0,length);
		strm = strm.substr(partm.length);
		wnd.SendMessage(partm);
		sendPartedMessage(wnd, strm, length);
	}else{
		wnd.SendMessage(strm);
	}
}

function ListFolders(folder,type){
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var dirFolders = new Array();
	if(fso.FolderExists(folder)){
		f = fso.GetFolder(folder);
		folders = new Enumerator(f.SubFolders);
		if(folders.atEnd()) return false;
		for (var i = 0; !folders.atEnd(); folders.moveNext())
		{
			if(type == 1)
				dirFolders[i++] = folders.item();
			else {
				var temp = "" + folders.item();
				var index = temp.lastIndexOf("\\");
				dirFolders[i++] = temp.substr(index+1);
			}
		}
		return dirFolders;
	} else
		return false;
}

function ListFiles(folder,type){
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var dirFiles = new Array();
	if(fso.FolderExists(folder)){
		f = fso.GetFolder(folder);
		files = new Enumerator(f.Files);
		if(files.atEnd()) return false;
		for (var i = 0; !files.atEnd(); files.moveNext())
		{
			if(type == 1)
				dirFiles[i++] = files.item();
			else {
				var temp = "" + files.item();
				var index = temp.lastIndexOf("\\");
				dirFiles[i++] = temp.substr(index+1);
			}
		}
		return dirFiles;
	} else
		return false;
}

function VRcredits()
{
	CreditsWindow = MsgPlus.CreateWnd("VecRemote.xml", "CreditsWindow");
}

function VRhelp()
{
	HelpWindow = MsgPlus.CreateWnd("VecRemote.xml", "HelpWindow");
}

function OnEvent_MenuClicked(MenuItemId, Location, OriginWnd)
{
	if(MenuItemId == "VRpass"){
		if (OriginWnd.EditChangeAllowed == true){
			OriginWnd.EditText = "/vrpass ";
		}
	}

	if(MenuItemId == "VRhelp"){
		VRhelp();
	}

	if(MenuItemId == "VRcredits"){
		VRcredits();
	}
}