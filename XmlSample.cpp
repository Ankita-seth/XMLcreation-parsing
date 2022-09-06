//// XmlSample.cpp : This file contains the 'main' function. Program execution begins and ends there.
/*----------------------------------------------------------------------------------------------------------------------------------------------------------*/

//#include <iostream>
//#include<iostream>
//#include<atlbase.h>
//#include<comutil.h>
//#import<msxml6.dll>
//
//using namespace MSXML2;
//using namespace std;
//
//HRESULT EmployeeInput(MSXML2::IXMLDOMDocument2Ptr& docPtr, 
//                      MSXML2::IXMLDOMElementPtr& employeeElement, 
//                      CComBSTR bstrValue1, 
//                      CComBSTR bstrValue2, 
//                      CComBSTR bstrValue3, 
//                      CComVariant varValue1, 
//                      CComVariant varValue2)
//{
//    try
//    {
//        MSXML2::IXMLDOMElementPtr employeeElements = docPtr->createElement(L"employee");
//
//        employeeElements->raw_setAttribute(bstrValue1.m_str, varValue1);
//
//        employeeElements->raw_setAttribute(bstrValue2.m_str, varValue2);
//
//        employeeElements->put_text(bstrValue3.m_str);
//        employeeElement->appendChild(employeeElements);
//    }
//    catch (...)
//    {
//    }
//    return S_OK;
//}
//
//HRESULT CreateElement(MSXML2::IXMLDOMDocument2Ptr& docPtr, 
//                      MSXML2::IXMLDOMElementPtr& parentElement, 
//                      MSXML2::IXMLDOMElementPtr& newElement, 
//                      CComBSTR bstrElementName , 
//                      CComBSTR bstrText)
//{
//    try
//    {
//        if (bstrElementName.Length())
//        {
//            newElement = docPtr->createElement(bstrElementName.m_str);
//            if (newElement)
//            {
//                if(bstrText.Length())
//                    newElement->put_text(bstrText.m_str);
//
//                parentElement->appendChild(newElement);
//            }
//        }
//
//    }
//    catch (...)
//    {
//    }
//    return S_OK;
//}
//
//int main()
//{
//    try
//    {
//        //Initializing COM 
//        ::CoInitialize(NULL);
//
//
//        //Creating dom document 
//        MSXML2::IXMLDOMDocument2Ptr docPtr;
//        HRESULT hr = docPtr.CreateInstance(__uuidof(MSXML2::DOMDocument60));
//        if (hr != S_OK || !docPtr)
//        {
//            return 0;
//        }
//
//        //Loading dom document
//        CComBSTR bstrXML(L"<company></company>");
//        CComVariant varValue;
//        CComBSTR bstrValue;
//        VARIANT_BOOL vbRet = VARIANT_FALSE;
//        docPtr->raw_loadXML(bstrXML.m_str, &vbRet);
//        if (vbRet == VARIANT_TRUE)
//        {
//            //Crating elements
//            MSXML2::IXMLDOMElementPtr detailsElement;
//            CreateElement(docPtr, (MSXML2::IXMLDOMElementPtr&)docPtr->documentElement, detailsElement, L"details", L"");
//
//            MSXML2::IXMLDOMElementPtr nameElement;
//            CreateElement(docPtr, detailsElement, nameElement, L"name", L"Beehive Systems Ltd");
//
//            MSXML2::IXMLDOMElementPtr headOfficeElement;
//            CreateElement(docPtr, detailsElement, headOfficeElement, L"headoffice", L"Noida");
//
//
//            //Creating List
//            MSXML2::IXMLDOMElementPtr employeeElement = docPtr->createElement(L"employees");
//            docPtr->documentElement->appendChild(employeeElement);
//
//            EmployeeInput(docPtr, employeeElement, L"name", L"empid", L"New Delhi", L"Amit", L"0067");
//            EmployeeInput(docPtr, employeeElement, L"name", L"empid", L"Mumbai", L"Shyam", L"0068");
//            EmployeeInput(docPtr, employeeElement, L"name", L"empid", L"Banglore", L"Ankita", L"0069");
//            EmployeeInput(docPtr, employeeElement, L"name", L"empid", L"Heydrabad", L"Anup", L"0070");
//
//            docPtr->save(L"D:\\test.xml");
//
//        }
//        else
//        {
//            cout << "Fail to create xml.";
//        }
//
//    }
//    catch (...)
//    {
//    }
//
//    return 0;
//}

/*----------------------------------------------------------------------------------------------------------------------------------------------------------*/


//#include <iostream>
//#include<iostream>
//#include<atlbase.h>
//#include <comutil.h>
//#import <msxml6.dll>
//
//using namespace MSXML2;
//using namespace std;
//
//HRESULT StudentInput(MSXML2::IXMLDOMDocument2Ptr& docPtr,
//    MSXML2::IXMLDOMElementPtr& studentElement,
//    CComBSTR bstrValue1,
//    CComBSTR bstrValue2,
//    CComBSTR bstrValue3,
//    CComVariant varValue1,
//    CComVariant varValue2)
//{
//    try
//    {
//        MSXML2::IXMLDOMElementPtr studentElements = docPtr->createElement(L"student");
//
//        studentElements->raw_setAttribute(bstrValue1.m_str, varValue1);
//
//        studentElements->raw_setAttribute(bstrValue2.m_str, varValue2);
//
//        studentElements->put_text(bstrValue3.m_str);
//
//        studentElement->appendChild(studentElements);
//    }
//    catch (...)
//    {
//    }
//    return S_OK;
//}
//
//
//
//int main()
//{
//    try
//    {
//        //Initializing COM 
//        ::CoInitialize(NULL);
//
//
//        //Creating dom document 
//        MSXML2::IXMLDOMDocument2Ptr docPtr;
//        HRESULT hr = docPtr.CreateInstance(__uuidof(MSXML2::DOMDocument60));
//        if (hr != S_OK || !docPtr)
//        {
//            return 0;
//        }
//
//        //Loading dom document
//        CComBSTR bstrXML(L"<Students></Students>");
//        CComVariant varValue;
//        CComBSTR bstrValue;
//        VARIANT_BOOL vbRet = VARIANT_FALSE;
//        docPtr->raw_loadXML(bstrXML.m_str, &vbRet);
//        if (vbRet == VARIANT_TRUE)
//        {
//            //Creating List
//             MSXML2::IXMLDOMElementPtr studentElement = docPtr->documentElement;

//            StudentInput(docPtr, studentElement, L"name", L"roll", L"New Delhi", L"Amit", L"1");
//            StudentInput(docPtr, studentElement, L"name", L"roll", L"Mumbai", L"Shyam", L"2");
//            StudentInput(docPtr, studentElement, L"name", L"roll", L"Banglore", L"Ankita", L"3");
//            StudentInput(docPtr, studentElement, L"name", L"roll", L"Heydrabad", L"Anup", L"4");
//            StudentInput(docPtr, studentElement, L"name", L"roll", L"New Delhi", L"Amit", L"5");
//            StudentInput(docPtr, studentElement, L"name", L"roll", L"Mumbai", L"Shyam", L"6");
//            StudentInput(docPtr, studentElement, L"name", L"roll", L"Banglore", L"Ankita", L"7");
//            StudentInput(docPtr, studentElement, L"name", L"roll", L"Heydrabad", L"Anup", L"8");
//            StudentInput(docPtr, studentElement, L"name", L"roll", L"Banglore", L"Ankita", L"9");
//            StudentInput(docPtr, studentElement, L"name", L"roll", L"Heydrabad", L"Anup", L"10");
//
//            docPtr->save(L"D:\\test.xml");
//        }
//        else
//        {
//            cout << "Fail to create xml.";
//        }
//    }
//    catch (...)
//    {
//    }
//    return 0;
//}

     /*----------------------------------------------------------------------------------------------------------------------------------------------------------*/
//
//#include <iostream>
//#include <iostream>
//#include <atlbase.h>
//#include <comutil.h>
//#import <msxml6.dll>
//
//using namespace MSXML2;
//using namespace std;
//
//
//HRESULT paramInput(MSXML2::IXMLDOMDocument2Ptr& docPtr,
//    MSXML2::IXMLDOMElementPtr& paramElement,
//    CComBSTR bstrValue1,
//    CComBSTR bstrValue2,
//    CComBSTR bstrValue3,
//    CComBSTR bstrValue4,
//    CComBSTR bstrValue5,
//    CComVariant varValue1,
//    CComVariant varValue2,
//    CComVariant varValue3,
//    CComVariant varValue4)
//{
//    try
//    {
// 
//        MSXML2::IXMLDOMElementPtr paramElements = docPtr->createElement(L"param");
//
//        paramElements->raw_setAttribute(bstrValue1.m_str, varValue1);
//
//        paramElements->raw_setAttribute(bstrValue2.m_str, varValue2);
//
//        paramElements->raw_setAttribute(bstrValue3.m_str, varValue3);
//
//        paramElements->raw_setAttribute(bstrValue4.m_str, varValue4);
//
//        paramElements->put_text(bstrValue5.m_str);
//
//        paramElement->appendChild(paramElements);
//    }
//    catch (...)
//    {
//    }
//    return S_OK;
//}
//HRESULT CreateElement(MSXML2::IXMLDOMDocument2Ptr& docPtr,
//    MSXML2::IXMLDOMElementPtr& parentElement,
//    MSXML2::IXMLDOMElementPtr& newElement,
//    CComBSTR bstrElementName,
//    CComBSTR bstrText)
//{
//    try
//    {
//        if (bstrElementName.Length())
//        {
//            newElement = docPtr->createElement(bstrElementName.m_str);
//            if (newElement)
//            {
//                if (bstrText.Length())
//                    newElement->put_text(bstrText.m_str);
//
//                parentElement->appendChild(newElement);
//            }
//        }
//
//    }
//    catch (...)
//    {
//    }
//    return S_OK;
//}
//
//
//
//HRESULT setAttribute(
//    MSXML2::IXMLDOMElementPtr& objectElement,
//    CComBSTR bstrclsid,
//    CComBSTR bstrname,
//    CComBSTR bstrtype,
//    CComBSTR bstrid,
//    CComVariant varclsid,
//    CComVariant varname,
//    CComVariant vartype,
//    CComVariant varid
//)
//{
//    try
//    {
//        objectElement->raw_setAttribute(bstrclsid.m_str, varname);
//        objectElement->raw_setAttribute(bstrname.m_str, varname);
//        objectElement->raw_setAttribute(bstrtype.m_str, vartype);
//        objectElement->raw_setAttribute(bstrid.m_str, varid);
//    }
//    catch (...)
//    {
//    }
//    return S_OK;
//}
//
//int WriteXml()
//{
//    try
//    {
//        //Initializing COM 
//        ::CoInitialize(NULL);
//
//
//        //Creating dom document 
//        MSXML2::IXMLDOMDocument2Ptr docPtr;
//        HRESULT hr = docPtr.CreateInstance(__uuidof(MSXML2::DOMDocument60));
//        if (hr != S_OK || !docPtr)
//        {
//            return 0;
//        }
//
//        //Loading dom document
//        CComBSTR bstrXML(L"<objects></objects>");
//        CComVariant varValue;
//        CComBSTR bstrValue;
//        VARIANT_BOOL vbRet = VARIANT_FALSE;
//        docPtr->raw_loadXML(bstrXML.m_str, &vbRet);
//
//        if (vbRet == VARIANT_TRUE)
//        {
//            //Creating List
//            MSXML2::IXMLDOMElementPtr objectElement = docPtr->createElement(L"object");
//            docPtr->documentElement->appendChild(objectElement);
//
//            objectElement->setAttribute(L"id", L"{46AEFCE4-FBE7-4DD6-8C4C-7653CB988C2C}");
//            objectElement->setAttribute(L"type", L"OBJECT");
//            objectElement->setAttribute(L"name", L"Background");
//            objectElement->setAttribute(L"clsid", L"SceneObject");
//
//            MSXML2::IXMLDOMElementPtr paramElement = docPtr->createElement(L"params");
//            objectElement->appendChild(paramElement);
//
//            paramInput(docPtr, paramElement, L"numkey", L"interpolator", L"type", L"name", L"Background", L"-1", L"NULL", L"STRING", L"name");
//
//            paramInput(docPtr, paramElement, L"numkey", L"interpolator", L"type", L"name", L"", L"-1", L"NULL", L"TEXTURE", L"texturenodeid");
//
//            docPtr->save(L"D:\\test.xml");
//        }
//
//        else
//        {
//            cout << "Fail to create xml.";
//        }
//    }
//    catch (...)
//    {
//    }
//    return 0;
//}
//
//int ParseXml()
//{
//    try
//    {
//        //Initializing COM 
//        ::CoInitialize(NULL);
//
//        //Creating dom document 
//        MSXML2::IXMLDOMDocument2Ptr docPtr;
//        HRESULT hr = docPtr.CreateInstance(__uuidof(MSXML2::DOMDocument60));
//        if (hr != S_OK || !docPtr)
//        {
//            return 0;
//        }
//
//        //Loading dom document
//        CComVariant bstrXmlPath = L"d:\\test.xml";
//        CComVariant varValue;
//        CComBSTR bstrValue;
//        VARIANT_BOOL vbRet = VARIANT_FALSE;
//        hr = docPtr->raw_load(bstrXmlPath, &vbRet);
//        if (vbRet == VARIANT_TRUE)
//        {
//            MSXML2::IXMLDOMElementPtr eleObject;
//            MSXML2::IXMLDOMNodePtr nodeObject;
//            docPtr->documentElement->raw_selectSingleNode(CComBSTR(L"object").m_str, &nodeObject);
//            if (nodeObject)
//            {
//                eleObject = nodeObject;
//
//                bstrValue.Empty();
//                hr = eleObject->raw_getAttribute(CComBSTR("id"), &varValue);
//                bstrValue = varValue.bstrVal;
//
//                bstrValue.Empty();
//                hr = eleObject->raw_getAttribute(CComBSTR("name"), &varValue);
//                bstrValue = varValue.bstrVal;
//
//                bstrValue.Empty();
//                hr = eleObject->raw_getAttribute(CComBSTR("clsid"), &varValue);
//                bstrValue = varValue.bstrVal;
//
//                bstrValue.Empty();
//                hr = eleObject->raw_getAttribute(CComBSTR("type"), &varValue);
//                bstrValue = varValue.bstrVal;  
//            }       
//        }
//        else
//        {
//            OutputDebugString(L"failed to load xml.");
//        }
//    }
//    catch (...)
//    {
// 
//    }
//}
//
//int main()
//{
//    WriteXml();
//    ParseXml();
//}



//----------------------------------------------------------------------------------------------------------------------------------------------------------



#include <iostream>
#include<iostream>
#include<atlbase.h>
#include <comutil.h>
#import <msxml6.dll>

using namespace MSXML2;
using namespace std;

HRESULT setAttribute(
        MSXML2::IXMLDOMElementPtr& bookElement,
        CComBSTR bstrgenere,
        CComBSTR bstrpublicationDate,
        CComBSTR bstrisbn,
        CComVariant vargenere,
        CComVariant varpublicationDate,
        CComVariant varisbn
    )
    {
        try
        {
            bookElement->raw_setAttribute(bstrgenere.m_str, vargenere);
            bookElement->raw_setAttribute(bstrpublicationDate.m_str, varpublicationDate);
            bookElement->raw_setAttribute(bstrisbn.m_str, varisbn);
        }
        catch (...)
        {
        }
        return S_OK;
    }

HRESULT CreateElement(MSXML2::IXMLDOMDocument2Ptr& docPtr,
    MSXML2::IXMLDOMElementPtr& parentElement,
    MSXML2::IXMLDOMElementPtr& newElement,
    CComBSTR bstrElementName,
    CComBSTR bstrText)
{
    try
    {
        if (bstrElementName.Length())
        {
            newElement = docPtr->createElement(bstrElementName.m_str);
            if (newElement)
            {
                if (bstrText.Length())
                    newElement->put_text(bstrText.m_str);

                parentElement->appendChild(newElement);
            }
        }

    }
    catch (...)
    {
    }
    return S_OK;
}

int writeXML()
{
    try
    {
        //Initializing COM 
        ::CoInitialize(NULL);


        //Creating dom document 
        MSXML2::IXMLDOMDocument2Ptr docPtr;
        HRESULT hr = docPtr.CreateInstance(__uuidof(MSXML2::DOMDocument60));
        if (hr != S_OK || !docPtr)
        {
            return 0;
        }

        //Loading dom document
        CComBSTR bstrXML(L"<bookStore></bookStore>");
        CComVariant varValue;
        CComBSTR bstrValue;
        VARIANT_BOOL vbRet = VARIANT_FALSE;
        docPtr->raw_loadXML(bstrXML.m_str, &vbRet);
        if (vbRet == VARIANT_TRUE)
        {
            //Creating List
             //MSXML2::IXMLDOMElementPtr studentElement = docPtr->documentElement;
            MSXML2::IXMLDOMElementPtr bookElement = docPtr->createElement(L"book");
            docPtr->documentElement->appendChild(bookElement);

            bookElement->setAttribute(L"ISBN", L"1-861003-11");
            bookElement->setAttribute(L"publicationDate", L"1981");
            bookElement->setAttribute(L"genere", L"autobiography");

            MSXML2::IXMLDOMElementPtr titleElement;
            CreateElement(docPtr, bookElement, titleElement, L"title", L"The autibiography of Benjamin Franklin");

            MSXML2::IXMLDOMElementPtr authorElement;
            CreateElement(docPtr, (MSXML2::IXMLDOMElementPtr&)bookElement, authorElement, L"author", L"");

            MSXML2::IXMLDOMElementPtr FirstnameElement;
            CreateElement(docPtr, authorElement, FirstnameElement, L"Firstname", L"Benjamin");

            MSXML2::IXMLDOMElementPtr SecondnameElement;
            CreateElement(docPtr, authorElement, SecondnameElement, L"Secondname", L"Franklin");

            MSXML2::IXMLDOMElementPtr priceElement;
            CreateElement(docPtr, bookElement, priceElement, L"price", L"8.99");

            MSXML2::IXMLDOMElementPtr bookElement2 = docPtr->createElement(L"book");
            docPtr->documentElement->appendChild(bookElement2);

            bookElement2->setAttribute(L"ISBN", L"1-861004-12");
            bookElement2->setAttribute(L"publicationDate", L"1982");
            bookElement2->setAttribute(L"genere", L"tourism");

            MSXML2::IXMLDOMElementPtr titleElement2;
            CreateElement(docPtr, bookElement2, titleElement2, L"title", L"A Passage to India");

            MSXML2::IXMLDOMElementPtr authorElement2;
            CreateElement(docPtr, (MSXML2::IXMLDOMElementPtr&)bookElement2, authorElement2, L"author", L"");

            MSXML2::IXMLDOMElementPtr FirstnameElement2;
            CreateElement(docPtr, authorElement2, FirstnameElement2, L"Firstname", L"E.M.");

            MSXML2::IXMLDOMElementPtr SecondnameElement2;
            CreateElement(docPtr, authorElement2, SecondnameElement2, L"Secondname", L"Foster");

            MSXML2::IXMLDOMElementPtr priceElement2;
            CreateElement(docPtr, bookElement2, priceElement2, L"price", L"7.99");

            MSXML2::IXMLDOMElementPtr bookElement3 = docPtr->createElement(L"book");
            docPtr->documentElement->appendChild(bookElement3);

            bookElement3->setAttribute(L"ISBN", L"1-861005-13");
            bookElement3->setAttribute(L"publicationDate", L"1999");
            bookElement3->setAttribute(L"genere", L"finance");

            MSXML2::IXMLDOMElementPtr titleElement3;
            CreateElement(docPtr, bookElement3, titleElement3, L"title", L"Rich Dad Poor Dad");

            MSXML2::IXMLDOMElementPtr authorElement3;
            CreateElement(docPtr, (MSXML2::IXMLDOMElementPtr&)bookElement3, authorElement3, L"author", L"");

            MSXML2::IXMLDOMElementPtr FirstnameElement3;
            CreateElement(docPtr, authorElement3, FirstnameElement3, L"Firstname", L"Robert");

            MSXML2::IXMLDOMElementPtr SecondnameElement3;
            CreateElement(docPtr, authorElement3, SecondnameElement3, L"Secondname", L"Kiyosaki");

            MSXML2::IXMLDOMElementPtr priceElement3;
            CreateElement(docPtr, bookElement3, priceElement3, L"price", L"9.00");

            docPtr->save(L"D:\\test.xml");
        }
        else
        {
            cout << "Fail to create xml.";
        }
    }
    catch (...)
    {
    }
    return 0;
}

int ParseXml()
{
    try
    {
        //Initializing COM 
        ::CoInitialize(NULL);

        //Creating dom document 
        MSXML2::IXMLDOMDocument2Ptr docPtr;
        HRESULT hr = docPtr.CreateInstance(__uuidof(MSXML2::DOMDocument60));
        if (hr != S_OK || !docPtr)
        {
            return 0;
        }

        //Loading dom document
        CComVariant bstrXmlPath = L"d:\\test.xml";
        CComVariant varValue;
        CComBSTR bstrValue;
        VARIANT_BOOL vbRet = VARIANT_FALSE;
        hr = docPtr->raw_load(bstrXmlPath, &vbRet);
        if (vbRet == VARIANT_TRUE)
        {
            MSXML2::IXMLDOMElementPtr eleObject;
            MSXML2::IXMLDOMNodePtr nodeObject;
            docPtr->documentElement->raw_selectSingleNode(CComBSTR(L"book").m_str, &nodeObject);
            if (nodeObject)
            {
                eleObject = nodeObject;

                bstrValue.Empty();
                hr = eleObject->raw_getAttribute(CComBSTR("publicationDate"), &varValue);
                bstrValue = varValue.bstrVal;

                bstrValue.Empty();
                hr = eleObject->raw_getAttribute(CComBSTR("firstname"), &varValue);
                bstrValue = varValue.bstrVal;

                bstrValue.Empty();
                hr = eleObject->raw_getAttribute(CComBSTR("clsid"), &varValue);
                bstrValue = varValue.bstrVal;
 
            }       
        }
        else
        {
            OutputDebugString(L"failed to load xml.");
        }
    }
    catch (...)
    {
    }
}

int main()
{
    //writeXML();
    ParseXml();
    system("pause");
    return 0;
}