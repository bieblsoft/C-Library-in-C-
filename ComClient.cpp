/*
	.NET Systemprogrammierung

	COM Interoperabilität

	Programm "ComClient"

	Clientapplikation unter C++ für eine unter .NET verfasste Klassenbibliothek

	Die dll-Datei muss als Verweis in das C++ Projekt eingebunden werden.
	Es wird automatisch eine passende Header-Datei erzeugt (.tlh), die die
	Schnittstellenbeschreibung enthält.
	Die unter .NET erzeugte Klassenbibliothek wird als COM Klassenbibliothek erzeugt
	und muss in C++ über die import-Anweisung eingebunden werden.
	Entweder geschieht dies durch den Zusatz "raw_interfaces_only" oder man importiert 
	zusätzlich - falls vorhanden - die in der Bibliothek "mscorlib.tlb" enthaltenen
	Definitionen.

	Zweck des Programmes:
	Berechnung von Schuldtilgungen
	Der Benutzer kann über die Kommandozeile alle benötigten Parameter eingeben; ein Parameter
	darf dabei weggelassen werden, so dass automatisch einer der 4 Eckdaten automatisch
	berechnet wird.
*/

// ohne mscorlib.tlb
#import "..\Com Client and .NET Host\bin\debug\NETHost.tlb" raw_interfaces_only
using namespace NETHost;
#include <iostream>
using namespace std;

int main(int argc, char* argv[])
{
	// initialisiert die COM Library / dll im aktuellen Thread und legt
	// das Threadverfahren auf STA fest
	HRESULT hr = CoInitialize(NULL);

	ILoanPtr pILoan(__uuidof(Loan));	// Anlegen einer Instanz der Loan Klasse

	if (argc < 5) 
	{
	   cout << "Eingabe: ComClient Kredit Zinssatz Monate Rate" << endl;
	   cout << "         Einer der Werte muss 0 sein!" << endl;
	   return -1;
	}

	double openingBalance = atof(argv[1]);
	double rate = atof(argv[2])/100.0;
	short  term = atoi(argv[3]);
	double payment = atof(argv[4]);

	pILoan->put_Kreditbetrag(openingBalance);
	pILoan->put_Zinssatz(rate);
	pILoan->put_Monate(term);
	pILoan->put_Rate(payment);

	if (openingBalance == 0.00) 
	    pILoan->ComputeOpeningBalance(&openingBalance);
	if (rate == 0.00) pILoan->ComputeRate(&rate);
	if (term == 0) pILoan->ComputeTerm(&term);
	if (payment == 0.00) pILoan->ComputePayment(&payment);

	cout << "Kredit   = " << openingBalance << endl;
	cout << "Zinssatz = " << rate*100 << endl;
	cout << "Monate   = " << term << endl;
	cout << "Rate     = " << payment << endl;

	VARIANT_BOOL MorePmts;	// der Rückbagewert der .NET Methode wird hier als
						// zusätzlicher Parameter übergeben
	double Balance = 0.0;	// Die Parameter werden z.T. durch die Methode verändert !
	double Principal = 0.0;
	double Interest = 0.0;

	cout << "Nr.    Rate Tilgung  Zinsen  Kredit" << endl;
	cout << "--- ------- ------- ------- -------" << endl;

	Balance = openingBalance;

	for (short PmtNbr = 1; ; PmtNbr++) 
	{
	   pILoan->GetNextPmtDistribution(payment, &Balance, &Principal, &Interest, &MorePmts); 
	   if ( ! MorePmts )
		   break;

	   cout.width( 3 );
	   cout << PmtNbr;
	   cout.width ( 8 );
	   cout << payment;
	   cout.width ( 8 );
	   cout << Principal;
	   cout.width ( 8 );
	   cout << Interest;
	   cout.width ( 8 );
	   cout << Balance;
	   cout << endl;
	}

	// Deinitialisieren der COM Library
	CoUninitialize();
	return 0;
}
