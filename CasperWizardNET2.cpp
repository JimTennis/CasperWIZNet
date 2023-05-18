// This is the main DLL file.

#include "CasperWizardNET2.h"

//#define NULL nullptr

namespace CasperWizardNET
{

CWGeneralReturnCodesEnum CasperWizard::ClearProducts ()
{
	CWGeneralReturnCodesEnum nResult = CWGeneralFail;

    // check if DLL is loaded
    if (_hLibrary != NULL)
    {
        // call the C function
        _CasperClearProducts ();

        // assign result
        nResult = CWGeneralSuccess;
    }

	return nResult;
}


CWGeneralReturnCodesEnum CasperWizard::AddProduct (String *pstrSiteCode, String *pstrProductCode,
	String *pstrDescription, String *pstrPrice)
{
	CWGeneralReturnCodesEnum nResult = CWGeneralFail;

    // check if DLL is loaded
    if (_hLibrary != NULL)
    {
        if (pstrSiteCode != NULL
            && pstrProductCode != NULL
            && pstrDescription != NULL
            && pstrPrice != NULL)
        {
            // convert to char pointers and call the C function
            CString strSiteCode = pstrSiteCode;
            CString strProductCode = pstrProductCode;
            CString strDescription = pstrDescription;
            CString strPrice = pstrPrice;
            _CasperAddProduct (strSiteCode.GetBuffer (0), strProductCode.GetBuffer (0),
                strDescription.GetBuffer (0),strPrice.GetBuffer (0));

            // assign result
            nResult = CWGeneralSuccess;
            strSiteCode.ReleaseBuffer ();
            strProductCode.ReleaseBuffer ();
            strDescription.ReleaseBuffer ();
            strPrice.ReleaseBuffer ();
        }
    }

	return nResult;
}


CWReturnCodesEnum CasperWizard::Wizard (CasperPurchase *pPurchase, CasperOptions *pOptions)
{
	CWReturnCodesEnum nResult = CWCOMObjectFail;

    try
    {
        // check if DLL is loaded
        if (_hLibrary == NULL)
        {
            nResult = CWCOMObjectNotLoaded;
            throw new long;
        }

        // check for NULL pointers
        if (pPurchase == NULL
            || pOptions == NULL)
        {
            nResult = CWCOMObjectNullPointer;
            throw new long;
        }

        // set default return code
        nResult = CWCOMObjectFail;

        // populate Wizard structures with data from COM objects
        Wizard::CasperPurchase WizardPurchase;
        Wizard::CasperOptions WizardOptions;
        CreateCasperPurchase (&WizardPurchase, pPurchase);
        CreateCasperOptions (&WizardOptions, pOptions);

		// call the Wizard
        nResult = (CWReturnCodesEnum) _CasperWizard (&WizardPurchase, &WizardOptions);

        // populate COM object with return data in Wizard structure
        ReturnCasperPurchase (&WizardPurchase, pPurchase);

        // clean up Wizard structures
        DestroyCasperPurchase (&WizardPurchase);
        DestroyCasperOptions (&WizardOptions);
    }
    catch (long *pnCode)
    {
        delete pnCode;
        pnCode = NULL;
    }

	return nResult;
}


CWReturnCodesEnum CasperWizard::Wizard2 (long hParentWindow, CasperPurchase *pPurchase,
    CasperOptions *pOptions, bool bAllowAutoEmail, bool bAllowManualEmail)
{
	CWReturnCodesEnum nResult = CWCOMObjectFail;

    try
    {
        // check if DLL is loaded
        if (_hLibrary == NULL)
        {
            nResult = CWCOMObjectNotLoaded;
            throw new long;
        }

        // check for NULL return variable
        if (pPurchase == NULL
            || pOptions == NULL)
        {
            nResult = CWCOMObjectNullPointer;
            throw new long;
        }

        // populate Wizard structures with data from COM objects
        Wizard::CasperPurchase WizardPurchase;
        Wizard::CasperOptions WizardOptions;
        CreateCasperPurchase (&WizardPurchase, pPurchase);
        CreateCasperOptions (&WizardOptions, pOptions);

        // call the Wizard
        nResult = (CWReturnCodesEnum) _CasperWizard2 ((HWND) hParentWindow, &WizardPurchase, &WizardOptions, bAllowAutoEmail,
            bAllowManualEmail);

        // populate COM object with return data in Wizard structure
        ReturnCasperPurchase (&WizardPurchase, pPurchase);

        // clean up Wizard structures
        DestroyCasperPurchase (&WizardPurchase);
        DestroyCasperOptions (&WizardOptions);
    }
    catch (long *pnCode)
    {
        delete pnCode;
        pnCode = NULL;
    }

	return nResult;
}


CWReturnCodesEnum CasperWizard::Wizard3 (long hParentWindow, CasperPurchase *pPurchase,
    CasperOptions2 *pOptions, bool bAllowAutoEmail, bool bAllowManualEmail)
{
	CWReturnCodesEnum nResult = CWCOMObjectFail;

    try
    {
        // check if DLL is loaded
        if (_hLibrary == NULL)
        {
            nResult = CWCOMObjectNotLoaded;
            throw new long;
        }

        // check for NULL return variable
        if (pPurchase == NULL
            || pOptions == NULL)
        {
            nResult = CWCOMObjectNullPointer;
            throw new long;
        }

        // populate Wizard structures with data from COM objects
        Wizard::CasperPurchase WizardPurchase;
        Wizard::CasperOptions2 WizardOptions;
        CreateCasperPurchase (&WizardPurchase, pPurchase);
        CreateCasperOptions2 (&WizardOptions, pOptions);

        // call the Wizard
        nResult = (CWReturnCodesEnum) _CasperWizard3 ((HWND) hParentWindow, &WizardPurchase, &WizardOptions, bAllowAutoEmail,
            bAllowManualEmail);

        // populate COM object with return data in Wizard structure
        ReturnCasperPurchase (&WizardPurchase, pPurchase);

        // clean up Wizard structures
        DestroyCasperPurchase (&WizardPurchase);
        DestroyCasperOptions2 (&WizardOptions);
    }
    catch (long *pnCode)
    {
        delete pnCode;
        pnCode = NULL;
    }

	return nResult;
}


CWReturnCodesEnum CasperWizard::Wizard4 (long hParentWindow, CasperPurchase *pPurchase,
    CasperOptions2 *pOptions, bool bAllowAutoEmail, bool bAllowManualEmail)
{
	CWReturnCodesEnum nResult = CWCOMObjectFail;

    try
    {
        // check if DLL is loaded
        if (_hLibrary == NULL)
        {
            nResult = CWCOMObjectNotLoaded;
            throw new long;
        }

        // check for NULL return variable
        if (pPurchase == NULL
            || pOptions == NULL)
        {
            nResult = CWCOMObjectNullPointer;
            throw new long;
        }

        // populate Wizard structures with data from COM objects
        Wizard::CasperPurchase WizardPurchase;
        Wizard::CasperOptions2 WizardOptions;
        CreateCasperPurchase (&WizardPurchase, pPurchase);
        CreateCasperOptions2 (&WizardOptions, pOptions);

        // call the Wizard
        nResult = (CWReturnCodesEnum) _CasperWizard4 ((HWND) hParentWindow, &WizardPurchase, &WizardOptions, bAllowAutoEmail,
            bAllowManualEmail);

        // populate COM object with return data in Wizard structure
        ReturnCasperPurchase (&WizardPurchase, pPurchase);

        // clean up Wizard structures
        DestroyCasperPurchase (&WizardPurchase);
        DestroyCasperOptions2 (&WizardOptions);
    }
    catch (long *pnCode)
    {
        delete pnCode;
        pnCode = NULL;
    }

	return nResult;
}


CWInputItemReturnCodesEnum CasperWizard::SetInputItems (CasperInputItem *pItem1, CasperInputItem *pItem2,
    CasperInputItem *pItem3, CasperInputItem *pItem4, CasperInputItem *pItem5, CasperInputItem *pItem6,
    CasperInputItem *pItem7, CasperInputItem *pItem8, CasperInputItem *pItem9, CasperInputItem *pItem10)
{
	CWInputItemReturnCodesEnum nResult = CWInputItemSuccess;

	Wizard::CasperInputItem WizardItem [CWGeneralEnum::CWMaxInputItems];
    Wizard::CasperInputItem *pWizardItem [CWGeneralEnum::CWMaxInputItems];
    long nIndex;
    HRESULT hResult = S_OK;


    try
    {
        // NULL pointers
        for (nIndex = 0;  nIndex < CWMaxInputItems;  nIndex++)
        {
            pWizardItem [nIndex] = NULL;
        }

        // check if DLL is loaded
        if (_hLibrary == NULL)
        {
            nResult = CWInputItemFail;
            throw new long;
        }

        // populate Wizard structures with data from COM objects
        CasperInputItem *pTempItem = NULL;
        for (nIndex = 0;  nIndex < CWMaxInputItems;  nIndex++)
        {
            switch (nIndex)
            {
                case 0:  pTempItem = pItem1;  break;
                case 1:  pTempItem = pItem2;  break;
                case 2:  pTempItem = pItem3;  break;
                case 3:  pTempItem = pItem4;  break;
                case 4:  pTempItem = pItem5;  break;
                case 5:  pTempItem = pItem6;  break;
                case 6:  pTempItem = pItem7;  break;
                case 7:  pTempItem = pItem8;  break;
                case 8:  pTempItem = pItem9;  break;
                case 9:  pTempItem = pItem10;  break;

                default:
                    nResult = CWInputItemFail;
                    throw new long;
                    break;
            }

            // check for NULL pointer
            if (pTempItem != NULL)
            {
                CreateCasperInputItem (&WizardItem [nIndex], pTempItem);
                pWizardItem [nIndex] = &WizardItem [nIndex];
            }
            else
            {
                pWizardItem [nIndex] = NULL;
            }
        }

        // call the Wizard function
        nResult = (CWInputItemReturnCodesEnum) _CasperSetInputItems (pWizardItem [0], pWizardItem [1], pWizardItem [2], pWizardItem [3],
            pWizardItem [4], pWizardItem [5], pWizardItem [6], pWizardItem [7], pWizardItem [8], pWizardItem [9]);
    }
    catch (long *pnCode)
    {
        delete pnCode;
        pnCode = NULL;
    }


    // clean up Wizard structures
    for (nIndex = 0;  nIndex < CWMaxInputItems;  nIndex++)
    {
        if (pWizardItem [nIndex] != NULL)
        {
            DestroyCasperInputItem (pWizardItem [nIndex]);
        }
    }

	return nResult;
}


CWGeneralReturnCodesEnum CasperWizard::SetTaxInfo (String *pstrStateAbbreviation,
	String *pstrStateDescription, String *pstrStateRate, String *pstrCountryName,
	String *pstrCountryDescription, String *pstrCountryRate)
{
	CWGeneralReturnCodesEnum nResult = CWGeneralFail;

    try
    {
        // check if DLL is loaded
        if (_hLibrary == NULL)
        {
            nResult = CWGeneralFail;
            throw new long;
        }

        // check for NULL pointers
        if (pstrStateAbbreviation != NULL
            && pstrStateDescription != NULL
            && pstrStateRate != NULL
            && pstrCountryName != NULL
            && pstrCountryDescription != NULL
            && pstrCountryRate != NULL)
        {
            // set default return code
            nResult = CWGeneralSuccess;

            // convert to char pointers and call function
            CString strStateAbbreviation = pstrStateAbbreviation;
            CString strStateDescription = pstrStateDescription;
            CString strStateRate = pstrStateRate;
            CString strCountryName = pstrCountryName;
            CString strCountryDescription = pstrCountryDescription;
            CString strCountryRate = pstrCountryRate;
            _CasperSetTaxInfo (strStateAbbreviation.GetBuffer (0), strStateDescription.GetBuffer (0),
                strStateRate.GetBuffer (0), strCountryName.GetBuffer (0), strCountryDescription.GetBuffer (0),
                strCountryRate.GetBuffer (0));

            strStateAbbreviation.ReleaseBuffer ();
            strStateDescription.ReleaseBuffer ();
            strStateRate.ReleaseBuffer ();
            strCountryName.ReleaseBuffer ();
            strCountryDescription.ReleaseBuffer ();
            strCountryRate.ReleaseBuffer ();
        }
    }
    catch (long *pnCode)
    {
        delete pnCode;
        pnCode = NULL;
    }

	return nResult;
}


CWSetStringReturnCodesEnum CasperWizard::SetStrings (CasperStrings *pStrings)
{
	CWSetStringReturnCodesEnum nResult = CWSetStringFail;

    try
    {
        // check if DLL is loaded
        if (_hLibrary == NULL)
        {
            throw new long;
        }

        // check for NULL pointer
        if (pStrings == NULL)
        {
            throw new long;
        }

        // populate Wizard structures with data from COM objects
        Wizard::CasperStrings WizardStrings;
        CreateCasperStrings (&WizardStrings, pStrings);

		// call the Wizard function
        nResult = (CWSetStringReturnCodesEnum) _CasperSetStrings (sizeof (Wizard::CasperStrings), &WizardStrings);

        // clean up Wizard structures
        DestroyCasperStrings (&WizardStrings);
    }
    catch (long *pnCode)
    {
        delete pnCode;
        pnCode = NULL;
    }

	return nResult;
}


CWGeneralReturnCodesEnum CasperWizard::SetFieldDefaults (CasperFieldDefaults *pFields)
{
	CWGeneralReturnCodesEnum nResult = CWGeneralFail;

    try
    {
        // check if DLL is loaded
        if (_hLibrary == NULL)
        {
            nResult = CWGeneralFail;
            throw new long;
        }

        // check for NULL pointer
        if (pFields == NULL)
        {
            nResult = CWGeneralFail;
            throw new long;
        }

        // populate Wizard structures with data from COM objects
        Wizard::CasperFieldDefaults WizardFields;
        CreateCasperFieldDefaults (&WizardFields, pFields);

		// call the Wizard function
        _CasperSetFieldDefaults (&WizardFields);

        // clean up Wizard structures
        DestroyCasperFieldDefaults (&WizardFields);
    }
    catch (long *pnCode)
    {
        delete pnCode;
        pnCode = NULL;
    }

	return nResult;
}


int CasperWizard::Send (CasperPurchase *pPurchase, CasperFieldDefaults *pFields,
    bool bAllowAutoEmail, bool bAllowManualEmail)
{
	int nResult = 0;

    try
    {
        // check if DLL is loaded
        if (_hLibrary == NULL)
        {
            throw new long;
        }

        // check for NULL pointers
        if (pPurchase == NULL
            || pFields == NULL)
        {
            throw new long;
        }

        // populate Wizard structures with data from COM objects
        Wizard::CasperPurchase WizardPurchase;
        Wizard::CasperFieldDefaults WizardFields;
        CreateCasperPurchase (&WizardPurchase, pPurchase);
        CreateCasperFieldDefaults (&WizardFields, pFields);

		// call the Wizard
        nResult = _CasperSend (&WizardPurchase, &WizardFields, bAllowAutoEmail, bAllowManualEmail);

        // populate COM object with return data in Wizard structure
        ReturnCasperPurchase (&WizardPurchase, pPurchase);

        // clean up Wizard structures
        DestroyCasperPurchase (&WizardPurchase);
        DestroyCasperFieldDefaults (&WizardFields);
    }
    catch (long *pnCode)
    {
        delete pnCode;
        pnCode = NULL;
    }

	return nResult;
}


void CasperWizard::CreateCasperPurchase (Wizard::CasperPurchase *pWizardPurchase,
	CasperPurchase *pPurchase)
{
    CString strTemp;
	char *pszTemp;

    try
    {
        // check for NULL pointer
        if (pWizardPurchase == NULL)
        {
            throw new long;
        }

        // NULL all pointers in structure
        pWizardPurchase->Currency = NULL;
        pWizardPurchase->Url = NULL;
        pWizardPurchase->Email = NULL;
        pWizardPurchase->Subject = NULL;

        // check for NULL pointer
        if (pPurchase == NULL)
        {
            throw new long;
        }

        // get currency description
        strTemp = pPurchase->get_strCurrency ();
        if (strTemp.GetLength () != 0)
        {
            pszTemp = new char [strTemp.GetLength () + 1];
            if (pszTemp == NULL)
            {
                throw new long;
            }
            _tcscpy (pszTemp, strTemp);
	        pWizardPurchase->Currency = pszTemp;
        }

        // get URL
        strTemp = pPurchase->get_strURL ();
        if (strTemp.GetLength () != 0)
        {
            pszTemp = new char [strTemp.GetLength () + 1];
            if (pszTemp == NULL)
            {
                throw new long;
            }
            _tcscpy (pszTemp, strTemp);
			pWizardPurchase->Url = pszTemp;
        }

        // clear Site Key
        for (long nIndex = 0;  nIndex < 32;  nIndex++)
        {
            pWizardPurchase->SiteKey [nIndex] = '\0';
        }

        // clear Error Message
        for (long nIndex2 = 0;  nIndex2 < 1000;  nIndex2++)
        {
            pWizardPurchase->ErrorMessage [nIndex2] = '\0';
        }

        // get Email
        strTemp = pPurchase->get_strEmailAddress ();
        if (strTemp.GetLength () != 0)
        {
            pszTemp = new char [strTemp.GetLength () + 1];
            if (pszTemp == NULL)
            {
                throw new long;
            }
            _tcscpy (pszTemp, strTemp);
	        pWizardPurchase->Email = pszTemp;
        }

        // get Subject
        strTemp = pPurchase->get_strEmailSubject ();
        if (strTemp.GetLength () != 0)
        {
            pszTemp = new char [strTemp.GetLength () + 1];
            if (pszTemp == NULL)
            {
                throw new long;
            }
            _tcscpy (pszTemp, strTemp);
	        pWizardPurchase->Subject = pszTemp;
        }
    }
    catch (long *pnCode)
    {
        delete pnCode;
        pnCode = NULL;
    }
}


void CasperWizard::CreateCasperOptions (Wizard::CasperOptions *pWizardOptions,
	CasperOptions *pOptions)
{
    CString strTemp;
    
    try
    {
        // check for NULL pointer
        if (pWizardOptions == NULL)
        {
            throw new long;
        }

        // NULL all pointers in structure
        pWizardOptions->DialogTitle = NULL;
        pWizardOptions->IntroTitle = NULL;
        pWizardOptions->IntroText = NULL;
        pWizardOptions->PayMethodText = NULL;
        pWizardOptions->FinalText = NULL;

        // check for NULL pointer
        if (pOptions == NULL)
        {
            throw new long;
        }

        // get dialog title
        strTemp = pOptions->get_strDialogTitle ();
	    pWizardOptions->DialogTitle = ToCharPointer (strTemp);

        // get hBitmap
		System::Drawing::Bitmap *pBitmap = static_cast <System::Drawing::Bitmap *> (pOptions->get_bitmap ());
		if (pBitmap != NULL)
		{
			pWizardOptions->WizardBitmap = (HBITMAP) pBitmap->GetHbitmap ().ToInt32 ();
		}

        // get intro title
        strTemp = pOptions->get_strIntroTitle ();
        pWizardOptions->IntroTitle = ToCharPointer (strTemp);

        // get intro text
        strTemp = pOptions->get_strIntroText ();
        pWizardOptions->IntroText = ToCharPointer (strTemp);

        // get pay method text
        strTemp = pOptions->get_strPayMethod ();
        pWizardOptions->PayMethodText = ToCharPointer (strTemp);

        // get final text
        strTemp = pOptions->get_strFinal ();
        pWizardOptions->FinalText = ToCharPointer (strTemp);

        // get Boolean values
        pWizardOptions->AcceptCreditCard = pOptions->get_bAcceptCreditCard ();
        pWizardOptions->AcceptSerialNumber = pOptions->get_bAcceptSerialNumber ();
        pWizardOptions->AcceptMasterCard = pOptions->get_bAcceptMasterCard ();
        pWizardOptions->AcceptVisa = pOptions->get_bAcceptVisa ();
        pWizardOptions->AcceptPayPal = pOptions->get_bAcceptPayPal ();
        pWizardOptions->AcceptAmericanExpress = pOptions->get_bAcceptAmericanExpress ();
        pWizardOptions->SkipPurchaseSummary = pOptions->get_bSkipPurchaseSummary ();
    }
    catch (long *pnCode)
    {
        delete pnCode;
        pnCode = NULL;
    }
}


void CasperWizard::CreateCasperOptions2 (Wizard::CasperOptions2 *pWizardOptions,
	CasperOptions2 *pOptions)
{
    // check for NULL pointers
    if (pWizardOptions != NULL
        && pOptions != NULL)
    {
        // get hBitmap
		System::Drawing::Bitmap *pBitmap = static_cast <System::Drawing::Bitmap *> (pOptions->get_hBitmap ());
		if (pBitmap != NULL)
		{
			pWizardOptions->WizardBitmap = (HBITMAP) pBitmap->GetHbitmap ().ToInt32 ();
		}

        // get Boolean values
        pWizardOptions->AcceptCreditCard = pOptions->get_bAcceptCreditCard ();
        pWizardOptions->AcceptSerialNumber = pOptions->get_bAcceptSerialNumber ();
        pWizardOptions->AcceptMasterCard = pOptions->get_bAcceptMasterCard ();
        pWizardOptions->AcceptVisa = pOptions->get_bAcceptVisa ();
        pWizardOptions->AcceptPayPal = pOptions->get_bAcceptPayPal ();
        pWizardOptions->AcceptAmericanExpress = pOptions->get_bAcceptAmericanExpress ();
        pWizardOptions->SkipPurchaseSummary = pOptions->get_bSkipPurchaseSummary ();
        pWizardOptions->Flags = pOptions->get_nFlags ();

        // set unused values
        pWizardOptions->Zero1 = 0;
        pWizardOptions->Zero2 = 0;
        pWizardOptions->Zero3 = 0;
    }
}


void CasperWizard::CreateCasperInputItem (Wizard::CasperInputItem *pWizardItem, 
	CasperInputItem *pItem)
{
    // check for NULL pointers
    if (pWizardItem != NULL
        && pItem!= NULL)
    {
        // get numerical values
        pWizardItem->Type = pItem->get_nType ();
        pWizardItem->BlankChar = pItem->get_cBlank ();

        // get string values
        CString strTemp;
        strTemp = pItem->get_strDescription ();
        _tcsncpy (pWizardItem->Description, (LPCTSTR) strTemp, CWMaxItemDescription);
        pWizardItem->Description [CWMaxItemDescription - 1] = '\0';

        strTemp = pItem->get_strMask ();
        _tcsncpy (pWizardItem->Mask, (LPCTSTR) strTemp, CWMaxItemMask);
        pWizardItem->Mask [CWMaxItemMask - 1] = '\0';
    }
}


void CasperWizard::CreateCasperStrings (Wizard::CasperStrings *pWizardStrings,
	CasperStrings *pStrings)
{
    CString strTemp;
    long nIndex;

    // structure to emulate CasperStrings structure
    struct structStrings
    {
		char *pszString [CWStringIDsEnum::CWSMax - 1];
    };
    structStrings sStrings;
    
    try
    {
        // check for NULL pointer
        if (pWizardStrings == NULL)
        {
            throw new long;
        }

        // NULL all pointers in structures
        for (nIndex = 0;  nIndex < CWStringIDsEnum::CWSMax - 1;  nIndex++)
        {
            sStrings.pszString [nIndex] = NULL;
            ((struct structStrings *) pWizardStrings)->pszString [nIndex] = NULL;
        }

        // check for NULL pointer
        if (pStrings == NULL)
        {
            throw new long;
        }

        // transfer strings to structure
        for (nIndex = 0;  nIndex < CWStringIDsEnum::CWSMax - 1;  nIndex++)
        {
            // get string
            strTemp = pStrings->strSetting ((CWStringIDsEnum) (nIndex + 1));
            ((struct structStrings *) pWizardStrings)->pszString [nIndex] = ToCharPointer (strTemp);
        }
    }
    catch (long *pnCode)
    {
        delete pnCode;
        pnCode = NULL;
    }
}


void CasperWizard::CreateCasperFieldDefaults (Wizard::CasperFieldDefaults *pWizardFields,
    CasperFieldDefaults *pFields)
{
    CString strTemp;
    long nIndex;

    try
    {
        // check for NULL pointer
        if (pWizardFields == NULL)
        {
            throw new long;
        }

        // NULL all pointers in structure
        pWizardFields->EmailAddress = NULL;
        pWizardFields->CreditCardNumber = NULL;
        pWizardFields->CreditCardName = NULL;
        pWizardFields->FirstName = NULL;
        pWizardFields->LastName = NULL;
        pWizardFields->Company = NULL;
        pWizardFields->Address1 = NULL;
        pWizardFields->Address2 = NULL;
        pWizardFields->City = NULL;
        pWizardFields->State = NULL;
        pWizardFields->PostalCode = NULL;
        pWizardFields->Country = NULL;
        pWizardFields->PhoneNumber = NULL;
        pWizardFields->FaxNumber = NULL;
        pWizardFields->SerialNumber = NULL;
		pWizardFields->CVN = NULL;

        for (nIndex = 0;  nIndex < CWMaxInputItems;  nIndex++)
        {
            pWizardFields->CustomInputs [nIndex] = NULL;
        }

        // check for NULL pointer
        if (pFields == NULL)
        {
            throw new long;
        }

        // get email address
        strTemp = pFields->get_strEmailAddress ();
        pWizardFields->EmailAddress = ToCharPointer (strTemp);

        // get credit card number
        strTemp = pFields->get_strCreditCardNumber ();
        pWizardFields->CreditCardNumber = ToCharPointer (strTemp);

        // get credit card name
        strTemp = pFields->get_strCreditCardName ();
        pWizardFields->CreditCardName = ToCharPointer (strTemp);

        // get first name
        strTemp = pFields->get_strFirstName ();
        pWizardFields->FirstName = ToCharPointer (strTemp);

        // get last name
        strTemp = pFields->get_strLastName ();
        pWizardFields->LastName = ToCharPointer (strTemp);

        // get company
        strTemp = pFields->get_strCompany ();
        pWizardFields->Company = ToCharPointer (strTemp);

        // get street address 1
        strTemp = pFields->get_strAddressLine1 ();
        pWizardFields->Address1 = ToCharPointer (strTemp);

        // get street address 2
        strTemp = pFields->get_strAddressLine2 ();
        pWizardFields->Address2 = ToCharPointer (strTemp);

        // get city
        strTemp = pFields->get_strCity ();
        pWizardFields->City = ToCharPointer (strTemp);

        // get state
        strTemp = pFields->get_strState ();
        pWizardFields->State = ToCharPointer (strTemp);

        // get postal code
        strTemp = pFields->get_strPostalCode ();
        pWizardFields->PostalCode = ToCharPointer (strTemp);

        // get country
        strTemp = pFields->get_strCountry ();
        pWizardFields->Country = ToCharPointer (strTemp);

        // get phone number
        strTemp = pFields->get_strPhone ();
        pWizardFields->PhoneNumber = ToCharPointer (strTemp);

        // get fax number
        strTemp = pFields->get_strFax ();
        pWizardFields->FaxNumber = ToCharPointer (strTemp);

        // get serial number
        strTemp = pFields->get_strSerialNumber ();
        pWizardFields->SerialNumber = ToCharPointer (strTemp);

        // get custom inputs
        for (nIndex = 0;  nIndex < CWMaxInputItems;  nIndex++)
        {
            strTemp = pFields->strCustomInput (nIndex);
            pWizardFields->CustomInputs [nIndex] = ToCharPointer (strTemp);
        }

        // set numerical values
        pWizardFields->PaymentMethod = pFields->get_nPaymentMethod ();
        pWizardFields->CreditCardType = pFields->get_nCreditCardType ();
        pWizardFields->CreditCardMonth = pFields->get_nCreditCardMonth ();
        pWizardFields->CreditCardYear = pFields->get_nCreditCardYear ();
        pWizardFields->SendMethod = pFields->get_nSendMethod ();
        pWizardFields->ProductIndex = pFields->get_nProductIndex ();

        // get CVV2 number
        strTemp = pFields->get_strCvv2();
        pWizardFields->CVN = ToCharPointer (strTemp);
	}
	catch (long *pnCode)
    {
        delete pnCode;
        pnCode = NULL;
    }
}


void CasperWizard::CreateCasperPersonalRequiredBits (Wizard::CasperPersonalRequiredBits *pWizardPersonalRequiredBits,
	CasperPersonalRequiredBits *pPersonalRequiredBits)
{
    CString strTemp;
    
    try
    {
        // check for NULL pointer
        if (pWizardPersonalRequiredBits == NULL)
        {
            throw new long;
        }

        // check for NULL pointer
        if (pPersonalRequiredBits == NULL)
        {
            throw new long;
        }

        // get integer values
        pWizardPersonalRequiredBits->FirstName = pPersonalRequiredBits->FirstName;
        pWizardPersonalRequiredBits->LastName = pPersonalRequiredBits->LastName;
        pWizardPersonalRequiredBits->Company = pPersonalRequiredBits->Company;
        pWizardPersonalRequiredBits->Address = pPersonalRequiredBits->Address;
        pWizardPersonalRequiredBits->AdditionalAddress = pPersonalRequiredBits->AdditionalAddress;
        pWizardPersonalRequiredBits->City = pPersonalRequiredBits->City;
        pWizardPersonalRequiredBits->PostalCode = pPersonalRequiredBits->PostalCode;
        pWizardPersonalRequiredBits->State = pPersonalRequiredBits->State;
        pWizardPersonalRequiredBits->Country = pPersonalRequiredBits->Country;
        pWizardPersonalRequiredBits->Phone = pPersonalRequiredBits->Phone;
        pWizardPersonalRequiredBits->Fax = pPersonalRequiredBits->Fax;
    }
    catch (long *pnCode)
    {
        delete pnCode;
        pnCode = NULL;
    }
}


void CasperWizard::ReturnCasperPurchase (Wizard::CasperPurchase *pWizardPurchase,
	CasperPurchase *pPurchase)
{
    // check for NULL pointers
    if (pWizardPurchase != NULL
        && pPurchase != NULL)
    {
        // copy Site Key
        pPurchase->strSiteKey = pWizardPurchase->SiteKey;

        // copy error message
        pPurchase->strErrorMessage = pWizardPurchase->ErrorMessage;
    }
}


void CasperWizard::DestroyCasperPurchase (Wizard::CasperPurchase *pWizardPurchase)
{
    // check for NULL pointer
    if (pWizardPurchase != NULL)
    {
        // deallocate memory
        if (pWizardPurchase->Currency != NULL)
        {
            delete [] pWizardPurchase->Currency;
            pWizardPurchase->Currency = NULL;
        }

        if (pWizardPurchase->Url != NULL)
        {
            delete [] pWizardPurchase->Url;
            pWizardPurchase->Url = NULL;
        }

        if (pWizardPurchase->Email != NULL)
        {
            delete [] pWizardPurchase->Email;
            pWizardPurchase->Email = NULL;
        }

        if (pWizardPurchase->Subject != NULL)
        {
            delete [] pWizardPurchase->Subject;
            pWizardPurchase->Subject = NULL;
        }
    }
}


void CasperWizard::DestroyCasperOptions (Wizard::CasperOptions *pWizardOptions)
{
    // check for NULL pointer
    if (pWizardOptions != NULL)
    {
        // deallocate memory
        if (pWizardOptions->DialogTitle != NULL)
        {
            delete [] pWizardOptions->DialogTitle;
            pWizardOptions->DialogTitle = NULL;
        }

        if (pWizardOptions->IntroTitle != NULL)
        {
            delete [] pWizardOptions->IntroTitle;
            pWizardOptions->IntroTitle = NULL;
        }

        if (pWizardOptions->IntroText != NULL)
        {
            delete [] pWizardOptions->IntroText;
            pWizardOptions->IntroText = NULL;
        }

        if (pWizardOptions->PayMethodText != NULL)
        {
            delete [] pWizardOptions->PayMethodText;
            pWizardOptions->PayMethodText = NULL;
        }

        if (pWizardOptions->FinalText != NULL)
        {
            delete [] pWizardOptions->FinalText;
            pWizardOptions->FinalText = NULL;
        }

		if (pWizardOptions->WizardBitmap != NULL)
		{
			DeleteObject (pWizardOptions->WizardBitmap);
			pWizardOptions->WizardBitmap = NULL;
		}
    }
}


void CasperWizard::DestroyCasperOptions2 (Wizard::CasperOptions2 *pWizardOptions)
{
    // check for NULL pointer
    if (pWizardOptions != NULL)
    {
		if (pWizardOptions->WizardBitmap != NULL)
		{
			DeleteObject (pWizardOptions->WizardBitmap);
			pWizardOptions->WizardBitmap = NULL;
		}
	}
}


void CasperWizard::DestroyCasperInputItem (Wizard::CasperInputItem *pWizardItem)
{
}


void CasperWizard::DestroyCasperStrings (Wizard::CasperStrings *pWizardStrings)
{
    // structure to emulate CasperStrings structure
    struct structStrings
    {
        char *pszString [CWSMax - 1];
    };

    // check for NULL pointer
    if (pWizardStrings != NULL)
    {
        // deallocate memory
        for (long nIndex = 0;  nIndex < CWSMax - 1;  nIndex++)
        {
            if (((struct structStrings *) pWizardStrings)->pszString [nIndex] != NULL)
            {
                delete [] ((struct structStrings *) pWizardStrings)->pszString [nIndex];
                ((struct structStrings *) pWizardStrings)->pszString [nIndex] = NULL;
            }
        }
    }
}


void CasperWizard::DestroyCasperFieldDefaults (Wizard::CasperFieldDefaults *pWizardFields)
{
    // check for NULL pointer
    if (pWizardFields != NULL)
    {
        // deallocate memory
        if (pWizardFields->EmailAddress != NULL)
        {
            delete [] pWizardFields->EmailAddress;
            pWizardFields->EmailAddress = NULL;
        }

        if (pWizardFields->CreditCardNumber != NULL)
        {
            delete [] pWizardFields->CreditCardNumber;
            pWizardFields->CreditCardNumber = NULL;
        }

        if (pWizardFields->CreditCardName != NULL)
        {
            delete [] pWizardFields->CreditCardNumber;
            pWizardFields->CreditCardNumber = NULL;
        }

        if (pWizardFields->FirstName != NULL)
        {
            delete [] pWizardFields->FirstName;
            pWizardFields->FirstName = NULL;
        }

        if (pWizardFields->LastName != NULL)
        {
            delete [] pWizardFields->LastName;
            pWizardFields->LastName = NULL;
        }

        if (pWizardFields->Company != NULL)
        {
            delete [] pWizardFields->Company;
            pWizardFields->Company = NULL;
        }

        if (pWizardFields->Address1 != NULL)
        {
            delete [] pWizardFields->Address1;
            pWizardFields->Address1 = NULL;
        }

        if (pWizardFields->Address2 != NULL)
        {
            delete [] pWizardFields->Address2;
            pWizardFields->Address2 = NULL;
        }

        if (pWizardFields->City != NULL)
        {
            delete [] pWizardFields->City;
            pWizardFields->City = NULL;
        }

        if (pWizardFields->State != NULL)
        {
            delete [] pWizardFields->State;
            pWizardFields->State = NULL;
        }

        if (pWizardFields->PostalCode != NULL)
        {
            delete [] pWizardFields->PostalCode;
            pWizardFields->PostalCode = NULL;
        }

        if (pWizardFields->Country != NULL)
        {
            delete [] pWizardFields->Country;
            pWizardFields->Country = NULL;
        }

        if (pWizardFields->PhoneNumber != NULL)
        {
            delete [] pWizardFields->PhoneNumber;
            pWizardFields->PhoneNumber = NULL;
        }

        if (pWizardFields->FaxNumber != NULL)
        {
            delete [] pWizardFields->FaxNumber;
            pWizardFields->FaxNumber = NULL;
        }

        if (pWizardFields->SerialNumber != NULL)
        {
            delete [] pWizardFields->SerialNumber;
            pWizardFields->SerialNumber = NULL;
        }

        if (pWizardFields->CVN != NULL)
        {
            delete [] pWizardFields->CVN;
            pWizardFields->CVN = NULL;
        }


        for (long nIndex = 0;  nIndex < CWMaxInputItems;  nIndex++)
        {
            if (pWizardFields->CustomInputs [nIndex] != NULL)
            {
                delete [] pWizardFields->CustomInputs [nIndex];
                pWizardFields->CustomInputs [nIndex] = NULL;
            }
        }
    }
}


void CasperWizard::DestroyCasperPersonalRequiredBits (Wizard::CasperPersonalRequiredBits *pWizardPersonalRequiredBits)
{
}


CWInitReturnCodesEnum CasperWizard::Initialize (String *pstrCasperPath)
{
	CWInitReturnCodesEnum nResult = CWInitFail;

    try
    {
        // determine path to DLL
        CString strPath;
        if (pstrCasperPath != NULL)
        {
            strPath = pstrCasperPath;
        }

        if (strPath.GetLength () == 0)
        {
            strPath = WIZARD_DLL_NAME;
        }

        // dynamically load the electronic purchase DLL
        _hLibrary = LoadLibrary (strPath);
        if (_hLibrary == NULL)
        {
            throw new long;
        }

        // retrieve the addresses of the functions we require
        _CasperClearProducts = reinterpret_cast <Wizard::pCasperClearProducts> (
            GetProcAddress (_hLibrary, WIZARD_FN_CLEARPRODS));

        _CasperAddProduct = reinterpret_cast <Wizard::pCasperAddProduct> (
            GetProcAddress (_hLibrary, WIZARD_FN_ADDPROD));

        _CasperSetInputItems = reinterpret_cast <Wizard::pCasperSetInputItems> (
            GetProcAddress (_hLibrary, WIZARD_FN_INPUTITEM));

        _CasperWizard = reinterpret_cast <Wizard::pCasperWizard> (
            GetProcAddress (_hLibrary, WIZARD_FN_WIZARD));

        _CasperWizard2 = reinterpret_cast <Wizard::pCasperWizard2> (
            GetProcAddress (_hLibrary, WIZARD_FN_WIZARD2));

        _CasperWizard3 = reinterpret_cast <Wizard::pCasperWizard3> (
            GetProcAddress (_hLibrary, WIZARD_FN_WIZARD3));

        _CasperWizard4 = reinterpret_cast <Wizard::pCasperWizard4> (
            GetProcAddress (_hLibrary, WIZARD_FN_WIZARD4));

        _CasperSetTaxInfo = reinterpret_cast <Wizard::pCasperSetTaxInfo> (
            GetProcAddress (_hLibrary, WIZARD_FN_SETTAX));

        _CasperSetStrings = reinterpret_cast <Wizard::pCasperSetStrings> (
            GetProcAddress (_hLibrary, WIZARD_FN_SETSTRINGS));

        _CasperSetFieldDefaults = reinterpret_cast <Wizard::pCasperSetFieldDefaults> (
            GetProcAddress (_hLibrary, WIZARD_FN_SETFIELDS));

        _CasperSend = reinterpret_cast <Wizard::pCasperSend> (
            GetProcAddress (_hLibrary, WIZARD_FN_SEND));

        _CasperLoadDefaultStrings = reinterpret_cast <Wizard::pCasperLoadDefaultStrings> (
            GetProcAddress (_hLibrary, WIZARD_FN_LOADDEFSTRINGS));

        _CasperSetPersonalRequiredBits = reinterpret_cast <Wizard::pCasperSetPersonalRequiredBits> (
            GetProcAddress (_hLibrary, WIZARD_FN_SETPERSONALBITS));

        if (_CasperClearProducts == NULL
            || _CasperAddProduct == NULL
            || _CasperWizard == NULL
            || _CasperWizard2 == NULL
            || _CasperWizard3 == NULL
            || _CasperWizard4 == NULL
            || _CasperSetTaxInfo == NULL
            || _CasperSetStrings == NULL
            || _CasperSetFieldDefaults == NULL
            || _CasperSend == NULL
            || _CasperLoadDefaultStrings == NULL
            || _CasperSetPersonalRequiredBits == NULL)
        {
            throw new long;
        }

        // set successful return code
        nResult = CWInitSuccess;
    }
    catch (long *pnCode)
    {
        delete pnCode;
        pnCode = NULL;

        // check if library needs to be freed
        if (_hLibrary != NULL)
        {
            FreeLibrary (_hLibrary);
            _hLibrary = NULL;
        }
    }

    return nResult;
}

// convert from CString to char *
char *CasperWizard::ToCharPointer (const CString &strValue)
{
	char *pszResult = NULL;

    if (strValue.GetLength () != 0)
    {
        pszResult = new char [strValue.GetLength () + 1];
        if (pszResult != NULL)
        {
	        _tcscpy (pszResult, (LPCTSTR) strValue);
		}
    }

	return pszResult;
}


long CasperWizard::LoadDefaultStrings(CWLanguageEnum language)
{
	return _CasperLoadDefaultStrings (language);
}


CWReturnCodesEnum CasperWizard::CasperSetPersonalRequiredBits(CasperPersonalRequiredBits *pBits)
{
	CWReturnCodesEnum nResult = CWReturnCodesEnum::CWCOMObjectFail;

    try
    {
        // check if DLL is loaded
        if (_hLibrary == NULL)
        {
            nResult = CWCOMObjectNotLoaded;
            throw new long;
        }

        // check for NULL pointers
        if (pBits == NULL)
        {
            nResult = CWCOMObjectNullPointer;
            throw new long;
        }

        // set default return code
        nResult = CWCOMObjectFail;

        // populate Wizard structures with data from COM objects
        Wizard::CasperPersonalRequiredBits WizardBits;
        CreateCasperPersonalRequiredBits (&WizardBits, pBits);

		// call the Wizard
		nResult = CWReturnCodesEnum::CWInternetSuccess;
		_CasperSetPersonalRequiredBits (&WizardBits);

        // clean up Wizard structures
        DestroyCasperPersonalRequiredBits (&WizardBits);
    }
    catch (long *pnCode)
    {
        delete pnCode;
        pnCode = NULL;
    }

	return nResult;
}

}  // namespace CasperWizardNET
