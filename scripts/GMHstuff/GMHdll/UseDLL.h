#ifndef USEDLL_H
#define USEDLL_H

	/* Name and Path to the DLL 
       Name und Pfad zur DLL */
	#define nameof_COMPORTDLL "C:\\Windows\\System32\\GMH3x32E.dll"
	//#define nameof_COMPORTDLL "C:\\WINNT\\System32\\GMH3x32E.dll"	

	/* Names of the functions as called in the DLL
	   Funktionsnamen genau wie sie in der DLL benannt wurden */
	#define nameof_GMH_OpenCom "GMH_OpenCom"
	#define nameof_GMH_CloseCom "GMH_CloseCom"
	#define nameof_GMH_Transmit "GMH_Transmit"
	#define nameof_GMH_GetVersionNumber "GMH_GetVersionNumber"
	#define nameof_GMH_GetMeasurement "GMH_GetMeasurement"


	/* Pointer typedefs corresponding to DLL-Function-types
	   Zeiger Typendefinitionen, entsprechend der DLL-Funktions-Typen */
	typedef __int16 (__stdcall *TGMH_OpenCom) (__int16);
	typedef __int16 (__stdcall *TGMH_CloseCom) (void);
	typedef __int16 (__stdcall *TGMH_Transmit) (__int16, __int16, __int16*, double*, __int32*);
	typedef __int16 (__stdcall *TGMH_GetVersionNumber) (void);
	typedef char (__stdcall *TGMH_GetMeasurement) (__int16, char*);


	/* Pointers to the DLL-Functions
	   Zeiger auf die DLL Funktionen */
	TGMH_OpenCom GMH_OpenCom;
	TGMH_CloseCom GMH_CloseCom;
	TGMH_Transmit GMH_Transmit;
	TGMH_GetVersionNumber GMH_GetVersionNumber;
	TGMH_GetMeasurement GMH_GetMeasurement;



#endif //USEDLL_H
