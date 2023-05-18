// CasperWizardNET2.h

#pragma once

#include "StdAfx.h"

using namespace System;

#using "System.Drawing.dll"
#using "CasperWizardNET.dll"


namespace Wizard
{
	#include "Casper.h"

    // function prototypes
    typedef void (PASCAL *pCasperClearProducts)(void);
    typedef void (PASCAL *pCasperAddProduct)(char *, char *, char *, char *);
    typedef int (PASCAL *pCasperSetInputItems)(Wizard::CasperInputItem *, Wizard::CasperInputItem *,
        Wizard::CasperInputItem *, Wizard::CasperInputItem *, Wizard::CasperInputItem *, Wizard::CasperInputItem *,
        Wizard::CasperInputItem *, Wizard::CasperInputItem *, Wizard::CasperInputItem *, Wizard::CasperInputItem *);
    typedef int (PASCAL *pCasperWizard)(Wizard::CasperPurchase *, Wizard::CasperOptions *);
    typedef int (PASCAL *pCasperWizard2)(HWND, Wizard::CasperPurchase *, Wizard::CasperOptions *, int, int);
    typedef int (PASCAL *pCasperWizard3)(HWND, Wizard::CasperPurchase *, Wizard::CasperOptions2 *, int, int);
    typedef int (PASCAL *pCasperWizard4)(HWND, Wizard::CasperPurchase *, Wizard::CasperOptions2 *, int, int);
    typedef void (PASCAL *pCasperSetTaxInfo)(char *, char *, char *, char *, char *, char *);
    typedef int (PASCAL *pCasperSetStrings)(int, Wizard::CasperStrings *);
    typedef void (PASCAL *pCasperSetFieldDefaults)(Wizard::CasperFieldDefaults *);
    typedef int (PASCAL *pCasperSend)(Wizard::CasperPurchase *, Wizard::CasperFieldDefaults *, int, int);
    typedef int (PASCAL *pCasperLoadDefaultStrings)(int);
    typedef int (PASCAL *pCasperSetPersonalRequiredBits)(Wizard::CasperPersonalRequiredBits *);


    // the name of the supplied electronic purchase wizard DLL
    #define WIZARD_DLL_NAME _T("casper.dll")

    // library export names
    #define WIZARD_FN_CLEARPRODS  "CasperClearProducts"
    #define WIZARD_FN_ADDPROD     "CasperAddProduct"
    #define WIZARD_FN_INPUTITEM   "CasperSetInputItems"
    #define WIZARD_FN_WIZARD      "CasperWizard"
    #define WIZARD_FN_WIZARD2     "CasperWizard2"
    #define WIZARD_FN_WIZARD3     "CasperWizard3"
    #define WIZARD_FN_WIZARD4     "CasperWizard4"
    #define WIZARD_FN_SETTAX      "CasperSetTaxInfo"
    #define WIZARD_FN_SETSTRINGS  "CasperSetStrings"
    #define WIZARD_FN_SETFIELDS   "CasperSetFieldDefaults"
    #define WIZARD_FN_SEND        "CasperSend"
    #define WIZARD_FN_LOADDEFSTRINGS "CasperLoadDefaultStrings"
    #define WIZARD_FN_SETPERSONALBITS "CasperSetPersonalRequiredBits"
}

namespace CasperWizardNET
{
	public __gc class CasperWizard
	{
		public:
			CasperWizard ()
			{
				_hLibrary = NULL;
			}

			~CasperWizard ()
			{
				// check if library needs to be freed
				if (_hLibrary != NULL)
				{
					FreeLibrary (_hLibrary);
					_hLibrary = NULL;
				}
			}

			CWGeneralReturnCodesEnum ClearProducts ();
			CWGeneralReturnCodesEnum AddProduct (String *pstrSiteCode, String *pstrProductCode,
				String *pstrDescription, String *pstrPrice);
			CWReturnCodesEnum Wizard (CasperPurchase *pPurchase, CasperOptions *pOptions);
			CWReturnCodesEnum Wizard2 (long hParentWindow, CasperPurchase *pPurchase,
				CasperOptions *pOptions, bool bAllowAutoEmail, bool bAllowManualEmail);
			CWReturnCodesEnum Wizard3 (long hParentWindow, CasperPurchase *pPurchase,
				CasperOptions2 *pOptions, bool bAllowAutoEmail, bool bAllowManualEmail);
			CWReturnCodesEnum Wizard4 (long hParentWindow, CasperPurchase *pPurchase,
				CasperOptions2 *pOptions, bool bAllowAutoEmail, bool bAllowManualEmail);
			CWInputItemReturnCodesEnum SetInputItems (CasperInputItem *pItem1, CasperInputItem *pItem2,
				CasperInputItem *pItem3, CasperInputItem *pItem4, CasperInputItem *pItem5, CasperInputItem *pItem6,
				CasperInputItem *pItem7, CasperInputItem *pItem8, CasperInputItem *pItem9, CasperInputItem *pItem10);
			CWGeneralReturnCodesEnum SetTaxInfo (String *pstrStateAbbreviation,
				String *pstrStateDescription, String *pstrStateRate, String *pstrCountryName,
				String *pstrCountryDescription, String *pstrCountryRate);
			CWSetStringReturnCodesEnum SetStrings (CasperStrings *pStrings);
			CWGeneralReturnCodesEnum SetFieldDefaults (CasperFieldDefaults *pFields);
			int Send (CasperPurchase *pPurchase, CasperFieldDefaults *pFields,
				bool bAllowAutoEmail, bool bAllowManualEmail);
			CWInitReturnCodesEnum CasperWizard::Initialize (String *pstrCasperPath);


		private:
			// handle to dynamically loaded DLL
			HINSTANCE _hLibrary;

			// pointers to functions in library
			Wizard::pCasperClearProducts     _CasperClearProducts;
			Wizard::pCasperAddProduct        _CasperAddProduct;
			Wizard::pCasperSetInputItems     _CasperSetInputItems;
			Wizard::pCasperWizard            _CasperWizard;
			Wizard::pCasperWizard2           _CasperWizard2;
			Wizard::pCasperWizard3           _CasperWizard3;
			Wizard::pCasperWizard4           _CasperWizard4;
			Wizard::pCasperSetTaxInfo        _CasperSetTaxInfo;
			Wizard::pCasperSetStrings        _CasperSetStrings;
			Wizard::pCasperSetFieldDefaults  _CasperSetFieldDefaults;
			Wizard::pCasperSend              _CasperSend;

			// load COM object data into C Structure
			void CreateCasperPurchase (Wizard::CasperPurchase *pWizardPurchase, CasperPurchase *pPurchase);
			void CreateCasperOptions (Wizard::CasperOptions *pWizardOptions, CasperOptions *pOptions);
			void CreateCasperOptions2 (Wizard::CasperOptions2 *pWizardOptions, CasperOptions2 *pOptions);
			void CreateCasperInputItem (Wizard::CasperInputItem *pWizardItem, CasperInputItem *pItem);
			void CreateCasperStrings (Wizard::CasperStrings *pWizardStrings, CasperStrings *pStrings);
			void CreateCasperFieldDefaults (Wizard::CasperFieldDefaults *pWizardFields, CasperFieldDefaults *pFields);

			// copy results from C structure to COM object
			void ReturnCasperPurchase (Wizard::CasperPurchase *pWizardPurchase, CasperPurchase *pPurchase);

			// clean up COM object
			void DestroyCasperPurchase (Wizard::CasperPurchase *pWizardPurchase);
			void DestroyCasperOptions (Wizard::CasperOptions *pWizardOptions);
			void DestroyCasperOptions2 (Wizard::CasperOptions2 *pWizardOptions);
			void DestroyCasperInputItem (Wizard::CasperInputItem *pWizardItem);
			void DestroyCasperStrings (Wizard::CasperStrings *pWizardStrings);
			void DestroyCasperFieldDefaults (Wizard::CasperFieldDefaults *pWizardFields);

			// convert from CString to char *
			char *ToCharPointer (const CString &strValue);

		public:
			long LoadDefaultStrings(CWLanguageEnum language);
			CWReturnCodesEnum CasperSetPersonalRequiredBits(CasperPersonalRequiredBits *pBits);

		private:
			Wizard::pCasperLoadDefaultStrings        _CasperLoadDefaultStrings;
			Wizard::pCasperSetPersonalRequiredBits   _CasperSetPersonalRequiredBits;
			void CreateCasperPersonalRequiredBits (Wizard::CasperPersonalRequiredBits *pWizardPersonalRequiredBits,
				CasperPersonalRequiredBits *pPersonalRequiredBits);
			void DestroyCasperPersonalRequiredBits (Wizard::CasperPersonalRequiredBits *pWizardPersonalRequiredBits);
	};
}
