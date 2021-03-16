//---------------------------------------------------------------------------
#include <windows.h>
#include <stdlib.h>
#include <stdio.h>
#include "UseDLL.h"

#pragma hdrstop

//---------------------------------------------------------------------------

#pragma argsused
int main(int argc, char* argv[])
{

/*  Variables will contain Values read from the device
    Variablen welche später die vom Gerät abgefragten Werte enthalten */
	__int16 Versionnumber=0;        //Enthält später die DLL Versionsnummer
	__int32 _IDNummer=0;	        //Enthält später die Geräte ID-Nummer
	__int16 _Addresse=0;	        //Enthält später die Geräte Addresse
	__int16 _error=0;		        //für GMHTransmit und sonstige Funktionen mit Fehlerrückgabe
	char text[20]="";		        //Enthält Später die Messart
	double AnzeigeWert=0;	        //Enthält Später den Anzeigewert

/*  Standard parameters for the GMHTransmit function
    Standardübergabe-Parameter für die Funktion GMHTransmit */
	__int16 _Priority =0;	        //für GMHTransmit
	double _Float =0;		        //für GMHTransmit
	__int32 _Int =0;		        //für GMHTransmit


/*  Load the DLL-File dynamically
    Die DLL-Datei dynamisch laden: */
	HMODULE h_COMPort_DLL; //HMODULE Struct für DLL-Speicher-Adresse

/*  Create a LPCSTRING Pointer dynamically,
    reserve Memory with the length of 'path + DLL-name' string
    Dynamisches Erzeugen eines LPCSTRING Pointers,
    Speicher der Grösse des Strings 'Pfad + DLL-Namen' reservieren */
	LPCTSTR lptest = (LPCTSTR)malloc(sizeof(nameof_COMPORTDLL));
	lptest = (LPCTSTR)nameof_COMPORTDLL;

	/*
		Optimal usage would be:
        Optimal wäre:
			LoadLibraryEx(DLL_Handle, ...)
			if (DLL_Handle)
				{
					Make a Pointer to the DLL Function
					Use DLL-Function
					FreeLibrary(DLL_Handle)
				}
	*/

    //Status
    printf("\nLoading: %s\n", lptest);


	h_COMPort_DLL = LoadLibraryEx(lptest,NULL,LOAD_WITH_ALTERED_SEARCH_PATH); //Load The DLL Library for using DLL-Functions
	if (h_COMPort_DLL) //If the Handle exitst, DLL File Could be loaded
		{
	/* Pointers to the DLL-Functions
	   Zeiger auf die DLL Funktionen */		
	GMH_OpenCom = (TGMH_OpenCom)GetProcAddress(h_COMPort_DLL, nameof_GMH_OpenCom);
	GMH_CloseCom = (TGMH_CloseCom)GetProcAddress(h_COMPort_DLL, nameof_GMH_CloseCom);
	GMH_GetVersionNumber = (TGMH_GetVersionNumber)GetProcAddress(h_COMPort_DLL, nameof_GMH_GetVersionNumber);
	GMH_Transmit = (TGMH_Transmit)GetProcAddress(h_COMPort_DLL, nameof_GMH_Transmit);
    GMH_GetMeasurement = (TGMH_GetMeasurement)GetProcAddress(h_COMPort_DLL, nameof_GMH_GetMeasurement);

/* end for loading a DLL dynamically <--
   DLL Dynamisch geladen <-- */

			//Now the internal DLL Functions can be used, as usual:
			_error = GMH_OpenCom(1);							//COM Port öffnen

            //Status -> Console
            if (_error < 0){printf("!Error opening COM-Port Errorcode %i\n", _error);}

			Versionnumber = GMH_GetVersionNumber();						//DLL-Versionsnummer lesen

            //Status -> Console
            if (_error < 0){printf("Error reading Versionnumber Errorcode %i\n", _error);}
            else {printf("Versionsnumber: %i\n", Versionnumber);}

			_error = GMH_Transmit(1, 0, &_Priority, &AnzeigeWert, &_Int);  //Anzeigewert Lesen (Wert in Float)

            //Status -> Console
            if (_error < 0){printf("Error reading Display Errorcode %i\n", _error);}
            else {printf("Anzeige: %f", AnzeigeWert);}

			_error = GMH_Transmit(1, 180, &_Priority, &_Float, &_Int);//Messart Lesen		(Wert in Int)
			_error = GMH_GetMeasurement(_Int, text);					//Messart decodieren

            //Status -> Console
            if (_error < 0){printf("Error reading Measurement Errorcode %i\n", _error);}
            else {printf("Measuring %s\n", text);}
			GMH_CloseCom();


			//Destroy Pointer to DLL
			FreeLibrary(h_COMPort_DLL);
		}
    else
        {
            //Status -> Console
            printf("\n\tError! DLL not existing!\n\t Is %s existing?\n\t Maybe you have to specify path and filename correctly\n\t Windows NT uses 'C:\\WINNT\\System32' instead of 'C:\\Windows\\System32'\n\n",lptest);
        }

    printf("\nPress Enter to close the application");
    getchar();
    return 0;
}
//---------------------------------------------------------------------------
 

