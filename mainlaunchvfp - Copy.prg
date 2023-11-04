Public lcMensajeMant, lcImgMant, lcRutaSis, lnTiempoMant,pcLocalApp
Public loShell  As Object
lnNewVer =0
lnCurVer=0
Set Step On
_Screen.Visible=.F.
lcRutaSis = Addbs(Justpath(Sys(16,1)))

lcRutaExe   = LeeIni(Curdir()+"eikonw.ini","Ruta","Cnf_Ruta")
lcExeRemote = LeeIni(Curdir()+"eikonw.ini","Ruta","Cnf_Ejecutable")
lcLocalApp  =lcRutaSis+"Eikonw.exe"

* Create a Shell.Application COM object
loShell = Createobject("Shell.Application")



llAcceso = LeeIni(Curdir()+"eikonw.ini","Ruta","Cnf_Acceso") = "S"
lnTiempoMant = Val(LeeIni(Curdir()+"eikonw.ini","Mantenimiento","Cnf_TiempoMant"))


lcExeRemoteApp=Addbs(lcRutaExe)+lcExeRemote
lcLocalFolder=Sys(5)+Curdir()
pcLocalApp=lcLocalApp

If llAcceso

	If File(lcExeRemoteApp,.T.)


		=Agetfileversion(aNewVer,lcExeRemoteApp)
		lnNewVer=aNewVer[4]

		=Agetfileversion(aCurVer,lcLocalApp)
		lnCurVer=aCurVer[4]



		If (lnNewVer > lnCurVer)

			Try
				&&Renombrar y Copiar
				lcMainApp=Justfname(lcLocalApp)

				If(IsExeRunning(lcMainApp))
					IsExeRunning(lcMainApp,.T.)
				Endif


				Set Step On
				lcTimeStamp=GetTimeStamp(Datetime(),1)
				lcTimeStamp=lcTimeStamp+".exe"
				Rename (Fullpath(lcLocalApp)) To (lcLocalApp+"_"+lcTimeStamp)




				* Set the source and destination file paths
				lcSourceFile =  lcExeRemoteApp  && Replace with your source file path
				lcDestFolder = lcLocalFolder  && Replace with your destination folder path

				* Get the source file's folder
				loSourceFolder = loShell.Namespace(Addbs(Justpath(lcSourceFile)))

				* Get the source file
				loSourceFile = loSourceFolder.ParseName(Justfname(lcSourceFile))

				* Get the destination folder
				loDestFolder = loShell.Namespace(lcDestFolder)
					
					
				* Copy the file to the destination folder
				loDestFolder.CopyHere(loSourceFile)




				bHaveError=.T.
				Do Case
					Case oException.ErrorNo=3
						Close Databases
						lcMessaError="El Archivo " + Upper(lcDbc) + "Esta abierto. Por Favor cerrar con el comando CLOSE DATABASE."+Chr(13)+Chr(10)+"Detalle del Error :"+Chr(10)

						lcErr=lcMessaError+lcErr
						Messagebox(lcErr, 16 , "Error en el Proceso de Carga.")

					Case oException.ErrorNo=1705
						lcMessaError="El Archivo " + Upper(lcDbc) + "Esta abierto. Por Favor cerrar con el comando CLOSE DATABASE."+Chr(13)+Chr(10)+"Detalle del Error :"+Chr(10)


						lcErr=lcMessaError+lcErr
						Messagebox(lcErr, 16 , "Error en el Proceso de Carga.")


					Case oException.ErrorNo=1429

						Messagebox("Anomalia al Intentar Cerrar la Aplicacion", 16 , "Anomila EikonLauncher")


						lcErr=[Program con Error : ] + Upper(oException.Procedure)+Chr(13)+;
							[Error: ] + Str(oException.ErrorNo) + Chr(13) + ;
							[Linea: ] + Str(oException.Lineno) + Chr(13) + ;
							[Mensaje: ] + oException.Message

						lcErr=lcMessaError+lcErr
						Messagebox(lcErr, 16 , "Error en el Proceso de Carga.")



					Otherwise
						lcErr  =[Program con Error : ] + Upper(oException.Procedure)+Chr(13)+;
							[Error: ] + Str(oException.ErrorNo) + Chr(13) + ;
							[Linea: ] + Str(oException.Lineno) + Chr(13) + ;
							[Mensaje: ] + oException.Message

						lcErr=lcMessaError+lcErr
						Messagebox(lcErr, 16 , "Error en el Proceso de Carga.")

				Endcase
			Finally

			Endtry



		Endif

	Endif

	IngresoSistema(Fullpath(lcLocalApp))
Else

	lcMensajeMant = LeeIni(lcRutaSis+"configlaunch.ini","Mantenimiento","MensajeMant")
	lcImgMant = LeeIni(lcRutaSis+"configlaunch.ini","Mantenimiento","ImgMant")

	FormMantenimiento()
Endif

*!*	Else
*!*		=Messagebox("No es Posible Acceder a la Aplicacion en este Momento [ " + Fullpath(lcExeRemote)+" ]",0+48,"EikonLauncher - Advertencia")
*!*		Return
*!*	Endif

*!*	Else

*!*		lcMensajeMant = LeeIni(lcRutaSis+"configlaunch.ini","Mantenimiento","MensajeMant")
*!*		lcImgMant = LeeIni(lcRutaSis+"configlaunch.ini","Mantenimiento","ImgMant")

*!*		FormMantenimiento()

*!*	Endif

*--------------------------------------------------------------------------------
*Nombre		         : LeeIni()
*Valor devuelto      : lcValor (Valor de la variable, si es que existe, sino retorna .NULL.)
*Autor				 : Juan Cueli
*Ejemplo			 : =LeeIni("C:MiArchivo.ini","Default","Puerto")
*Creación            : 11/02/2017
*----------------------------------------------------------------------------------
*Lista modificaciones:
*  Fecha		Autor	Cambio
*----------------------------------------------------------------------------------
*** <summary>
*** Lee un valor de un archivo INI. Si no existe el archivo, la sección o la entrada, retorna .NULL.
*** </summary>
*** <param name="pcArchivo">Nombre y ruta completa del archivo .INI</param>
*** <param name="pcSeccion">Sección del archivo .INI</param>
*** <param name="pcVariable">Variable del archivo .INI donde se guardara el valor</param>
*** <remarks></remarks>
Function LeeIni(pcArchivo,pcSeccion,pcVariable)
	Local lcValor, lnResultado, lnBufferSize, lcArchivo

	lcArchivo = pcArchivo

	Declare Integer GetPrivateProfileString ;
		IN WIN32API ;
		STRING cSeccion,;
		STRING cVariable,;
		STRING cDefecto,;
		STRING @aRetVal,;
		INTEGER nTam,;
		STRING cArchivo
	lnBufferSize = 255
	lcValor = Spac(lnBufferSize)
	lnResultado=GetPrivateProfileString(pcSeccion,pcVariable,"*NULL*",;
		@lcValor,lnBufferSize,lcArchivo)
	lcValor=Substr(lcValor,1,lnResultado)
	If lcValor="*NULL*"
		lcValor=.Null.
	Endif
	Return lcValor
Endfunc

*--------------------------------------------------------------------------------
*Nombre		         : FormMantenimiento()
*Valor devuelto      :
*Autor				 : Juan Cueli
*Ejemplo			 : FormMantenimiento()
*Creación            : 23/09/2019
*----------------------------------------------------------------------------------
*Lista modificaciones:
*  Fecha		Autor	Cambio
*----------------------------------------------------------------------------------
*** <summary>
*** Crea un formulario que le indica al usuario que el sistema está en mantenimiento y no puede acceder al mismo.
*** </summary>
*** <remarks></remarks>
Procedure FormMantenimiento
	Local loForm

	loForm = Createobject("Mantenimiento")

	loForm.Show()

	Read Events

	loForm = .Null.

	Release loForm

	Return

Define Class Mantenimiento As Form

	*-Propiedades de la ventana de Mantenimiento
	BackColor 	= Rgb(255, 255, 224)
	Caption  	= "Mantenimiento"
	AutoCenter 	= .T.
	Width 		= 490
	Height 		= 490
	BorderStyle = 1
	ControlBox	= .F.
	ShowWindow 	= 2

	*-Controles que tiene el formulario
	*-Label con mensaje de mantenimiento
	Add Object lblMant As Label With ;
		BackStyle  	= 0,;
		Caption   	= "", ;
		ForeColor 	= Rgb(255,0, 0), ;
		Left      	= 10, ;
		Top       	= 10, ;
		Width     	= 480,;
		Height	  	= 35,;
		wordwrap  	= .T.,;
		FontBold  	= .T.,;
		Alignment	= 2,;
		caption	  	= lcMensajeMant

	*-Imagen de la pantalla de mantenimiento
	Add Object imgMant As Image With ;
		picture 	= lcImgMant,;
		Left      	= 100, ;
		Top       	= 60, ;
		Width     	= 50,;
		Height	 	= 50

	*-Boton que cierra el formulario
	Add Object btncerrar As CommandButton With ;
		left 		= 335,;
		top 		= 400,;
		height 		= 27,;
		width 		= 135,;
		caption		= "Cerrar"

	*-Boton que le da acceso al sistema
	Add Object btnAcceder As CommandButton With ;
		left 		= 50,;
		top 		= 420,;
		height 		= 27,;
		width 		= 135,;
		caption		= "Accesar el sistema"

	*-Timer que indica el intervalo de tiempo en en que se va a verificar si se tiene o no acceso al sistema
	Add Object tmReintentar As Timer With ;
		interval = lnTiempoMant * 1000

	*-Boton para reintentar validar el acceso al sistema
	Add Object btnReintentar As CommandButton With ;
		left 		= 50,;
		top 		= 380,;
		height 		= 27,;
		width 		= 135,;
		caption		= "Volver a Intentar en "+Alltrim(Str(lnTiempoMant)),;
		nReintentar = (Thisform.tmReintentar.Interval)/1000

	*-Timer que reduce el tiempo para reintentar cada segundo
	Add Object tmCaptionBtnReintentar As Timer With ;
		interval = 1000

	*-Metodos y Procedimiento del formulario y los objetos
	*-Procedimiento que se corre cuando se cierra el formulario
	Procedure Destroy
		Clear Events
		Release lcMensajeMant, lcImgMant, lcRutaSis, lnTiempoMant
	Endproc

	*-Inicio del formulario
	Procedure Init
		Thisform.btnAcceder.Refresh
	Endproc

	*-Clic al boton cerrar
	Procedure btncerrar.Click
		Thisform.Release
	Endproc

	*-Refresh al boton acceder
	Procedure btnAcceder.Refresh
		llAcceso = LeeIni(Curdir()+"eikonw.ini","Ruta","Cnf_Acceso") = "S"

		This.Enabled = llAcceso
	Endproc

	*-Clic al boton acceder
	Procedure btnAcceder.Click
		IngresoSistema(pcLocalApp)
	Endproc

	*-Timer que validar el tiempo restante para reintentar acceder al sistema
	Procedure tmCaptionBtnReintentar.Timer
		If Thisform.btnReintentar.nReintentar > 0
			Thisform.btnReintentar.nReintentar = Thisform.btnReintentar.nReintentar - 1
		Else
			Thisform.btnReintentar.Click
		Endif

		Thisform.btnReintentar.Caption	= "Volver a Intentar en "+Alltrim(Str(Thisform.btnReintentar.nReintentar))
	Endproc

	*-Clic al boton reintentar
	Procedure btnReintentar.Click
		Thisform.btnReintentar.nReintentar = lnTiempoMant
		Thisform.btnAcceder.Refresh
	Endproc

Enddefine


*Endproc


*--------------------------------------------------------------------------------
*Nombre		         : IngresoSistema()
*Valor devuelto      :
*Autor				 : Juan Cueli
*Ejemplo			 : =IngresoSistema()
*Creación            : 25/09/2019
*----------------------------------------------------------------------------------
*Lista modificaciones:
*  Fecha		Autor	Cambio
*----------------------------------------------------------------------------------
*** <summary>
*** Lee un valor de un archivo INI. Si no existe el archivo, la sección o la entrada, retorna .NULL.
*** </summary>
*** <param name="pcArchivo">Nombre y ruta completa del archivo .INI</param>
*** <param name="pcSeccion">Sección del archivo .INI</param>
*** <param name="pcVariable">Variable del archivo .INI donde se guardara el valor</param>
*** <remarks></remarks>

Function IngresoSistema(pcMainApp)


	If File(pcMainApp)

		loShell.ShellExecute(pcMainApp)
		loShell= Null

	Else
		=Messagebox("No se Tiene Acceso a la Aplicacion [ " + pcMainApp +" ]",0+16,"EikonLauncher - Verificar con Soporte Tecnico")
	Endif

	Release lcMensajeMant, lcImgMant, lcRutaSis, lnTiempoMant

	Quit
Endfunc




Function IsExeRunning(tcName, tlTerminate)
	lcErr=''
	lcMessaError=''
	*Set Step On
	Try
		Local loLocator, loWMI, loProcesses, loProcess, llIsRunning
		loLocator 	= Createobject('WBEMScripting.SWBEMLocator')
		loWMI		= loLocator.ConnectServer()
		loWMI.Security_.ImpersonationLevel = 3  		&& Impersonate

		loProcesses	= loWMI.ExecQuery([SELECT * FROM Win32_Process WHERE Name = '] + tcName + ['])
		llIsRunning = .F.
		If loProcesses.Count > 0
			For Each loProcess In loProcesses
				llIsRunning = .T.
				If tlTerminate
					loProcess.Terminate(0)
				Endif
			Endfor
		Endif



	Catch To oException

		bHaveError=.T.
		Do Case
			Case oException.ErrorNo=3
				Close Databases
				lcMessaError="El Archivo " + Upper(lcDbc) + "Esta abierto. Por Favor cerrar con el comando CLOSE DATABASE."+Chr(13)+Chr(10)+"Detalle del Error :"+Chr(10)

				lcErr=lcMessaError+lcErr
				Messagebox(lcErr, 16 , "Error en el Proceso de Carga.")

			Case oException.ErrorNo=1705
				lcMessaError="El Archivo " + Upper(lcDbc) + "Esta abierto. Por Favor cerrar con el comando CLOSE DATABASE."+Chr(13)+Chr(10)+"Detalle del Error :"+Chr(10)


				lcErr=lcMessaError+lcErr
				Messagebox(lcErr, 16 , "Error en el Proceso de Carga.")


			Case oException.ErrorNo=1429

				Messagebox("Anomalia al Intentar Cerrar la Aplicacion", 16 , "Anomila EikonLauncher")


				lcErr=[Program con Error : ] + Upper(oException.Procedure)+Chr(13)+;
					[Error: ] + Str(oException.ErrorNo) + Chr(13) + ;
					[Linea: ] + Str(oException.Lineno) + Chr(13) + ;
					[Mensaje: ] + oException.Message

				lcErr=lcMessaError+lcErr
				Messagebox(lcErr, 16 , "Error en el Proceso de Carga.")



			Otherwise
				lcErr  =[Program con Error : ] + Upper(oException.Procedure)+Chr(13)+;
					[Error: ] + Str(oException.ErrorNo) + Chr(13) + ;
					[Linea: ] + Str(oException.Lineno) + Chr(13) + ;
					[Mensaje: ] + oException.Message

				lcErr=lcMessaError+lcErr
				Messagebox(lcErr, 16 , "Error en el Proceso de Carga.")

		Endcase
	Finally

	Endtry

	Return llIsRunning
Endfunc


Function GetTimeStamp(pDate,pHora,nType)
	*
	If Vartype(nType) <> "N" Then
		nType=1 &&Retorna Fecha
	Endif


	If Vartype(pDate)<>"D"
		*Messagebox("Error en la Funcion")
		*Return
		pDate=Date()
	Endif


	If Vartype(pHora)<> "T"
		*Messagebox("Error en la Funcion")
		*Return
		pHora=Datetime()
	Endif


	lcYear=Alltrim(Str(Year(pDate)))
	lcMonth=Alltrim(Str(Month(pDate)))
	lcDay=Alltrim(Str(Day(pDate)))

	lcHora=Alltrim(Str(Hour(Datetime())))
	lcMin=Alltrim(Str(Minute(Datetime())))
	lcSec=Alltrim(Str(Sec(Datetime())))



	If Len(lcMonth) = 1 Then
		lcMonth='0'+lcMonth
	Endif

	If Len(lcDay) = 1 Then
		lcDay='0'+lcDay
	Endif


	If Len(lcHora) = 1 Then
		lcHora='0'+lcHora
	Endif

	If Len(lcMin) = 1 Then
		lcMin='0'+lcMin
	Endif


	If Len(lcSec) = 1 Then
		lcSec='0'+lcSec
	Endif

	Do Case
		Case nType=1 &&Año-Mes-Dia
			*lcTimeBackup=lcYear+lcMonth+lcDay+'.txt'
			lcTimeBackup=lcYear+lcMonth+lcDay
			*lcTimeBackup=lcYear+lcMonth+lcDay+lcHora+lcSec


		Case nType=2
			lcTimeBackup=lcYear+lcMonth+lcDay+lcHora &&Año-Mes-Dia-Hora


		Case nType=3
			lcTimeBackup=lcYear+lcMonth+lcDay+lcHora &&Año-Mes-Dia-Hora


		Otherwise
			lcTimeBackup=lcYear+lcMonth+lcDay+lcHora+lcSec &&Año-Mes-Dia-Hora+Segun


	Endcase

	*lcTime=lcHora+lcMin+lcSec
	*lcTimeBackup=lcYear+lcMonth+lcDay+lcHora+'.txt'

	Return lcTimeBackup

Endfunc